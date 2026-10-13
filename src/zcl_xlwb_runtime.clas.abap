"! <p class="shorttext synchronized" lang="en">XLWB Cloud: render form theo tên template</p>
"!
"! API duy nhất mà code nghiệp vụ cần gọi. Template lấy từ bảng ZXLWB_TMPL
"! (quản lý bằng app Fiori ZUI_XLWB_TMPL_O4), engine chọn tự động theo
"! cột ENGINE nên consumer không cần biết form dùng SpreadsheetML, XLSX hay DOCX.
"!
"! <pre>
"! DATA(ls_file) = zcl_xlwb_runtime=>render( iv_form_name = 'ORDER_FORM'
"!                                           ir_context   = REF #( ls_context ) ).
"! " ls_file-content   : xstring, đưa vào action RAP dạng base64
"! " ls_file-mime_type : điền vào mimetype của response
"! " ls_file-file_name : tên file gợi ý cho người dùng
"! </pre>
CLASS zcl_xlwb_runtime DEFINITION
  PUBLIC FINAL CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_file,
             content   TYPE xstring,
             mime_type TYPE string,
             file_name TYPE string,
             extension TYPE string,
             engine    TYPE string,
             source    TYPE string,
           END OF ty_file.

    CONSTANTS:
      c_engine_ssml TYPE string VALUE `SSML`,
      c_engine_xlsx TYPE string VALUE `XLSX`,
      c_engine_docx TYPE string VALUE `DOCX`,
      c_engine_asset TYPE string VALUE `ASST`.

    "! Render form từ template đang lưu trong bảng
    "! @parameter iv_form_name | key của template (ZXLWB_TMPL-FORM_NAME)
    "! @parameter ir_context   | REF TO structure chứa dữ liệu bind
    CLASS-METHODS render
      IMPORTING iv_form_name   TYPE string
                ir_context     TYPE REF TO data
      RETURNING VALUE(rs_file) TYPE ty_file
      RAISING   zcx_xlwb.

    "! Render trực tiếp từ nội dung template — dùng cho Preview (template chưa
    "! lưu) hoặc khi template đến từ nguồn khác bảng ZXLWB_TMPL.
    "! @parameter iv_engine | SSML, XLSX hoặc DOCX
    CLASS-METHODS render_with_template
      IMPORTING iv_template    TYPE xstring
                iv_engine      TYPE string
                ir_context     TYPE REF TO data
      RETURNING VALUE(rs_file) TYPE ty_file
      RAISING   zcx_xlwb.

"! Render ƯU TIÊN template trong bảng ZXLWB_TMPL; nếu form chưa được khai
    "! (hoặc bị đánh dấu không active) thì dùng template dự phòng do code dựng.
    "! Nhờ vậy người dùng có thể "chiếm quyền" một form bằng cách upload template
    "! của mình vào app quản lý template, KHÔNG cần sửa ABAP.
    "! @parameter iv_form_name         | key trong ZXLWB_TMPL, vd 'XLWB_EX05'
    "! @parameter iv_fallback_template | template do code dựng (xstring UTF-8)
    "! @parameter iv_fallback_engine   | engine của template dự phòng
    CLASS-METHODS render_prefer_table
      IMPORTING iv_form_name         TYPE string
                iv_fallback_template TYPE xstring
                iv_fallback_engine   TYPE string DEFAULT `SSML`
                ir_context           TYPE REF TO data
      RETURNING VALUE(rs_file)       TYPE ty_file
      RAISING   zcx_xlwb.

    "! Render bằng dữ liệu THẬT từ source class đã gán trên template
    "! (cột SOURCE_CLASS — class implement zif_xlwb_source).
    "! @parameter iv_keys | tham số chọn dạng JSON chuyển cho get_context
    CLASS-METHODS render_by_source
      IMPORTING iv_form_name   TYPE string
                iv_keys        TYPE string OPTIONAL
      RETURNING VALUE(rs_file) TYPE ty_file
      RAISING   zcx_xlwb.

    "! Instance source class của form (dùng lại được cho get_sample)
    CLASS-METHODS get_source_handler
      IMPORTING iv_form_name     TYPE string
      RETURNING VALUE(ro_source) TYPE REF TO zif_xlwb_source
      RAISING   zcx_xlwb.

    "! Form đã có template trong bảng chưa (active + có nội dung)
    CLASS-METHODS exists_in_table
      IMPORTING iv_form_name     TYPE string
      RETURNING VALUE(rv_exists) TYPE abap_bool.

    "! Danh sách template đang active (dùng cho value help / kiểm tra cấu hình)
    CLASS-METHODS list_active
      RETURNING VALUE(rt_forms) TYPE string_table.

    "! Xoá cache template trong session (gọi sau khi sửa template)
    CLASS-METHODS clear_cache.

    "! Chọn 1 ASSET (logo/ảnh, ENGINE=ASST) trong ZXLWB_TMPL -> base64 PNG.
    "! Dùng khi in để logo động: ls_ctx-logo = get_asset_b64( 'LOGO_CASLA' ).
    CLASS-METHODS get_asset_b64
      IMPORTING iv_name       TYPE string
      RETURNING VALUE(rv_b64) TYPE string
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_cache,
             form_name    TYPE c LENGTH 40,
             engine       TYPE c LENGTH 4,
             mime_type    TYPE c LENGTH 128,
             file_name    TYPE c LENGTH 128,
             source_class TYPE c LENGTH 30,
             template     TYPE xstring,
           END OF ty_cache,
           ty_caches TYPE SORTED TABLE OF ty_cache WITH UNIQUE KEY form_name.

    CLASS-DATA mt_cache TYPE ty_caches.

    CLASS-METHODS read_template
      IMPORTING iv_form_name    TYPE string
      RETURNING VALUE(rs_cache) TYPE ty_cache
      RAISING   zcx_xlwb.

    "! Render SSML rồi chuyển kết quả sang .xlsx (zcl_xlwb_ssml2xlsx).
    "! Chuyển lỗi thì trả bản SpreadsheetML gốc (.xls) kèm ghi chú trong SOURCE.
    CLASS-METHODS render_ssml
      IMPORTING iv_template_xml TYPE string
                ir_context      TYPE REF TO data
      RETURNING VALUE(rs_file)  TYPE ty_file
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_RUNTIME IMPLEMENTATION.


  METHOD render.
    DATA(ls_tmpl) = read_template( iv_form_name ).

    rs_file-engine    = ls_tmpl-engine.
    rs_file-mime_type = ls_tmpl-mime_type.
    rs_file-file_name = ls_tmpl-file_name.

    CASE ls_tmpl-engine.

      WHEN c_engine_xlsx.
        rs_file-extension = `xlsx`.
        rs_file-content   = NEW zcl_xlwb_xlsx( )->render(
                              iv_template = ls_tmpl-template
                              ir_context  = ir_context ).

      WHEN c_engine_ssml OR ``.
        " template SSML lưu dạng nhị phân UTF-8 -> chuyển về string cho engine
        DATA(lv_xml) = cl_abap_conv_codepage=>create_in( )->convert( ls_tmpl-template ).
        DATA(ls_ssml) = render_ssml( iv_template_xml = lv_xml ir_context = ir_context ).
        rs_file-content   = ls_ssml-content.
        rs_file-extension = ls_ssml-extension.
        DATA(lv_note) = ls_ssml-source.
        IF rs_file-extension = `xlsx`.
          " file cuối là OOXML -> mime/tên file của bảng (khai theo .xls) phải theo
          rs_file-mime_type = `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
          REPLACE FIRST OCCURRENCE OF PCRE '(?i)\.xls$' IN rs_file-file_name WITH `.xlsx`.
        ENDIF.

      WHEN c_engine_docx.
        rs_file-extension = `docx`.
        rs_file-content   = NEW zcl_xlwb_docx( )->render(
                              iv_template = ls_tmpl-template
                              ir_context  = ir_context ).
        IF rs_file-mime_type IS INITIAL.
          rs_file-mime_type = zcl_xlwb_docx=>c_mime_docx.
        ENDIF.

      WHEN c_engine_asset.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |'{ iv_form_name }' là ASSET (ảnh), không render trực tiếp — dùng get_asset_b64| ).

      WHEN OTHERS.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |Template '{ iv_form_name }' khai engine không hợp lệ: '{ ls_tmpl-engine }'| ).
    ENDCASE.

    IF rs_file-mime_type IS INITIAL.
      rs_file-mime_type = COND #( WHEN rs_file-extension = `xls`
        THEN `application/vnd.ms-excel`
        ELSE `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` ).
    ENDIF.

    IF rs_file-file_name IS INITIAL.
      rs_file-file_name = |{ iv_form_name }.{ rs_file-extension }|.
    ENDIF.

    rs_file-source = |ZXLWB_TMPL:{ to_upper( iv_form_name ) }| && lv_note.
  ENDMETHOD.


  METHOD get_asset_b64.
    DATA(lv_key) = to_upper( condense( iv_name ) ).
    SELECT SINGLE template FROM zxlwb_tmpl WHERE form_name = @lv_key INTO @DATA(lv_x).
    IF sy-subrc <> 0 OR lv_x IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Asset '{ iv_name }' chưa có trong ZXLWB_TMPL| ).
    ENDIF.
    rv_b64 = xco_cp=>xstring( lv_x )->as_string( xco_cp_binary=>text_encoding->base64 )->value.
  ENDMETHOD.


  METHOD render_ssml.
    DATA(lv_ssml) = NEW zcl_xlwb_engine( )->render( iv_template = iv_template_xml
                                                    ir_context  = ir_context ).
    rs_file-engine = c_engine_ssml.
    TRY.
        rs_file-content   = zcl_xlwb_ssml2xlsx=>convert( lv_ssml ).
        rs_file-extension = `xlsx`.
        rs_file-mime_type = `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
      CATCH zcx_xlwb INTO DATA(lx_conv).
        " an toàn: trả bản SpreadsheetML gốc thay vì chết; SOURCE mang ghi chú
        rs_file-content   = cl_abap_conv_codepage=>create_out( )->convert( lv_ssml ).
        rs_file-extension = `xls`.
        rs_file-mime_type = `application/vnd.ms-excel`.
        rs_file-source    = | [ssml2xlsx lỗi: { lx_conv->get_text( ) }]|.
    ENDTRY.
  ENDMETHOD.


  METHOD read_template.
    DATA(lv_key) = to_upper( condense( iv_form_name ) ).

    READ TABLE mt_cache INTO rs_cache WITH KEY form_name = lv_key.
    IF sy-subrc = 0.
      RETURN.
    ENDIF.

    SELECT SINGLE form_name, engine, mime_type, file_name, source_class, template, is_active
      FROM zxlwb_tmpl
      WHERE form_name = @lv_key
      INTO @DATA(ls_db).

    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Không tìm thấy template '{ iv_form_name }' trong ZXLWB_TMPL| ).
    ENDIF.

    IF ls_db-is_active IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Template '{ iv_form_name }' đang bị đánh dấu không active| ).
    ENDIF.

    IF ls_db-template IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Template '{ iv_form_name }' chưa có file nội dung| ).
    ENDIF.

    rs_cache = VALUE ty_cache( form_name    = ls_db-form_name
                               engine       = ls_db-engine
                               mime_type    = ls_db-mime_type
                               file_name    = ls_db-file_name
                               source_class = ls_db-source_class
                               template     = ls_db-template ).
    INSERT rs_cache INTO TABLE mt_cache.
  ENDMETHOD.


  METHOD get_source_handler.
    DATA(ls_tmpl) = read_template( iv_form_name ).

    IF ls_tmpl-source_class IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Template '{ iv_form_name }' chưa gán Source class (ZIF_XLWB_SOURCE)| ).
    ENDIF.

    DATA(lv_class) = to_upper( condense( CONV string( ls_tmpl-source_class ) ) ).
    DATA lo_obj TYPE REF TO object.
    TRY.
        CREATE OBJECT lo_obj TYPE (lv_class).
        ro_source = CAST zif_xlwb_source( lo_obj ).
      CATCH cx_sy_create_object_error cx_sy_move_cast_error.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |Source class '{ lv_class }' không tạo được instance | &&
                    |hoặc không implement ZIF_XLWB_SOURCE| ).
    ENDTRY.
  ENDMETHOD.


  METHOD render_by_source.
    DATA(lo_source) = get_source_handler( iv_form_name ).
    rs_file = render( iv_form_name = iv_form_name
                      ir_context   = lo_source->get_context( iv_keys ) ).
  ENDMETHOD.


  METHOD render_with_template.
    DATA(lv_engine) = to_upper( condense( iv_engine ) ).

    IF iv_template IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Nội dung template rỗng| ).
    ENDIF.

    CASE lv_engine.
      WHEN c_engine_xlsx.
        rs_file-engine    = c_engine_xlsx.
        rs_file-extension = `xlsx`.
        rs_file-mime_type = `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
        rs_file-content   = NEW zcl_xlwb_xlsx( )->render( iv_template = iv_template
                                                          ir_context  = ir_context ).
      WHEN c_engine_ssml OR ``.
        DATA(lv_xml) = cl_abap_conv_codepage=>create_in( )->convert( iv_template ).
        DATA(ls_ssml) = render_ssml( iv_template_xml = lv_xml ir_context = ir_context ).
        rs_file-engine    = ls_ssml-engine.
        rs_file-extension = ls_ssml-extension.
        rs_file-mime_type = ls_ssml-mime_type.
        rs_file-content   = ls_ssml-content.
        DATA(lv_note) = ls_ssml-source.
      WHEN c_engine_docx.
        rs_file-engine    = c_engine_docx.
        rs_file-extension = `docx`.
        rs_file-mime_type = zcl_xlwb_docx=>c_mime_docx.
        rs_file-content   = NEW zcl_xlwb_docx( )->render( iv_template = iv_template
                                                          ir_context  = ir_context ).

      WHEN OTHERS.
        RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Engine không hợp lệ: '{ iv_engine }'| ).
    ENDCASE.

    rs_file-file_name = |preview.{ rs_file-extension }|.
    rs_file-source    = |CODE| && lv_note.
  ENDMETHOD.


  METHOD render_prefer_table.
    IF exists_in_table( iv_form_name ) = abap_true.
      rs_file = render( iv_form_name = iv_form_name ir_context = ir_context ).
      RETURN.
    ENDIF.

    rs_file = render_with_template( iv_template = iv_fallback_template
                                    iv_engine   = iv_fallback_engine
                                    ir_context  = ir_context ).
    rs_file-file_name = |{ to_lower( iv_form_name ) }.{ rs_file-extension }|.
  ENDMETHOD.


  METHOD exists_in_table.
    DATA(lv_key) = to_upper( condense( iv_form_name ) ).

    READ TABLE mt_cache TRANSPORTING NO FIELDS WITH KEY form_name = lv_key.
    IF sy-subrc = 0.
      rv_exists = abap_true.
      RETURN.
    ENDIF.

    " template la LOB (rawstring) -> khong dung duoc trong WHERE, phai doc ra roi kiem tra
    SELECT SINGLE template
      FROM zxlwb_tmpl
      WHERE form_name = @lv_key
        AND is_active = @abap_true
      INTO @DATA(lv_template).

    rv_exists = xsdbool( sy-subrc = 0 AND lv_template IS NOT INITIAL ).
  ENDMETHOD.


  METHOD clear_cache.
    CLEAR mt_cache.
  ENDMETHOD.


  METHOD list_active.
    SELECT form_name FROM zxlwb_tmpl
      WHERE is_active = @abap_true
      ORDER BY form_name
      INTO TABLE @DATA(lt_names).

    LOOP AT lt_names INTO DATA(ls).
      APPEND CONV string( ls-form_name ) TO rt_forms.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
