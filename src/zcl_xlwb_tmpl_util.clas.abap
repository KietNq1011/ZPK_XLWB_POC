"! <p class="shorttext synchronized" lang="en">XLWB: template checks + export/import</p>
"!
"! Tiện ích cho app quản lý template: quét placeholder trong template,
"! kiểm tra cặp loop/điều kiện, đối chiếu path với sample JSON (validate
"! lúc SAVE — chặn lỗi trước khi render), và export/import template
"! dạng JSON để chuyển giữa các tenant (bảng Z không đi theo transport).
CLASS zcl_xlwb_tmpl_util DEFINITION
  PUBLIC FINAL CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_msg,
             severity TYPE c LENGTH 1,       " E = chặn save, W = cảnh báo
             text     TYPE string,
           END OF ty_msg,
           ty_msgs TYPE STANDARD TABLE OF ty_msg WITH EMPTY KEY.

    TYPES ty_frames TYPE STANDARD TABLE OF REF TO cl_abap_typedescr WITH EMPTY KEY.

    TYPES: BEGIN OF ty_export,
             form_name    TYPE c LENGTH 40,
             descr        TYPE c LENGTH 80,
             engine       TYPE c LENGTH 4,
             mime_type    TYPE c LENGTH 128,
             file_name    TYPE c LENGTH 128,
             source_class TYPE c LENGTH 30,
             sample_json  TYPE string,
             template_b64 TYPE string,
           END OF ty_export.

    "! Trả về mọi token {{...}} trong template (đã bỏ dấu ngoặc)
    CLASS-METHODS scan_placeholders
      IMPORTING iv_template      TYPE xstring
                iv_engine        TYPE string
      RETURNING VALUE(rt_tokens) TYPE string_table
      RAISING   zcx_xlwb.

    "! Kiểm tra template: cặp {{#}}/{{/}} khớp (E), sample JSON parse được (E),
    "! path có trong sample context (W). Bảng rỗng = sạch.
    CLASS-METHODS validate
      IMPORTING iv_template        TYPE xstring
                iv_engine          TYPE string
                iv_sample_json     TYPE string OPTIONAL
      RETURNING VALUE(rt_messages) TYPE ty_msgs.

    "! Đóng gói 1 template thành JSON (template nhị phân -> base64)
    CLASS-METHODS export_json
      IMPORTING is_tmpl        TYPE ty_export
                iv_template    TYPE xstring
      RETURNING VALUE(rv_json) TYPE string.

    "! Đọc JSON export ngược lại thành dữ liệu template
    CLASS-METHODS import_json
      IMPORTING iv_json            TYPE string
      EXPORTING es_tmpl            TYPE ty_export
                ev_template        TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    CLASS-METHODS to_text
      IMPORTING iv_template    TYPE xstring
                iv_engine      TYPE string
      RETURNING VALUE(rv_text) TYPE string
      RAISING   zcx_xlwb.

    CLASS-METHODS path_exists
      IMPORTING io_type          TYPE REF TO cl_abap_typedescr
                it_segments      TYPE string_table
                iv_from          TYPE i
      RETURNING VALUE(rv_exists) TYPE abap_bool.

    CLASS-METHODS collect_frames
      IMPORTING io_type   TYPE REF TO cl_abap_typedescr
      CHANGING  ct_frames TYPE ty_frames.

    CLASS-METHODS json_escape
      IMPORTING iv_text          TYPE string
      RETURNING VALUE(rv_escaped) TYPE string.

    CLASS-METHODS json_get_string
      IMPORTING ir_data         TYPE REF TO data
                iv_comp         TYPE string
      RETURNING VALUE(rv_value) TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_TMPL_UTIL IMPLEMENTATION.


  METHOD scan_placeholders.
    DATA(lv_text) = to_text( iv_template = iv_template iv_engine = iv_engine ).

    " (?-x): tắt extended mode PCRE — bẫy chuẩn của FIND PCRE pattern động
    FIND ALL OCCURRENCES OF PCRE '(?-x)\{\{([^{}]+)\}\}'
      IN lv_text RESULTS DATA(lt_results).

    LOOP AT lt_results INTO DATA(ls_res).
      READ TABLE ls_res-submatches INTO DATA(ls_sub) INDEX 1.
      IF sy-subrc = 0.
        DATA(lv_token) = condense( substring( val = lv_text
                                              off = ls_sub-offset
                                              len = ls_sub-length ) ).
        IF lv_token IS NOT INITIAL.
          APPEND lv_token TO rt_tokens.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD to_text.
    CASE to_upper( iv_engine ).

      WHEN 'DOCX' OR 'ASST'.
        " DOCX bind qua XML part; ASST là ảnh — không có placeholder {{...}}
        rv_text = ``.

      WHEN 'XLSX'.
        DATA(lo_zip) = NEW cl_abap_zip( ).
        lo_zip->load( EXPORTING zip = iv_template
                      EXCEPTIONS zip_parse_error = 1 OTHERS = 2 ).
        IF sy-subrc <> 0.
          RAISE EXCEPTION NEW zcx_xlwb( iv_text = |File không phải .xlsx (zip) hợp lệ| ).
        ENDIF.
        LOOP AT lo_zip->files INTO DATA(ls_file)
          WHERE name CP 'xl/*.xml' OR name CP 'xl/worksheets/*.xml'.
          DATA lv_part TYPE xstring.
          lo_zip->get( EXPORTING name = ls_file-name
                       IMPORTING content = lv_part
                       EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
          IF sy-subrc = 0 AND lv_part IS NOT INITIAL.
            rv_text = rv_text && cl_abap_conv_codepage=>create_in( )->convert( lv_part ).
          ENDIF.
        ENDLOOP.

      WHEN OTHERS.                                     " SSML: file là XML text UTF-8
        TRY.
            rv_text = cl_abap_conv_codepage=>create_in( )->convert( iv_template ).
          CATCH cx_root ##CATCH_ALL.
            RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Không đọc được template dạng text UTF-8| ).
        ENDTRY.
    ENDCASE.
  ENDMETHOD.


  METHOD validate.
    DATA lt_tokens TYPE string_table.

    TRY.
        lt_tokens = scan_placeholders( iv_template = iv_template iv_engine = iv_engine ).
      CATCH zcx_xlwb INTO DATA(lx_scan).
        APPEND VALUE #( severity = 'E' text = lx_scan->get_text( ) ) TO rt_messages.
        RETURN.
    ENDTRY.

    " 1) cặp mở/đóng {{#x}} {{?x}} {{^x}} ... {{/x}}
    DATA lt_stack TYPE string_table.
    LOOP AT lt_tokens INTO DATA(lv_token).
      DATA(lv_head) = lv_token(1).
      CASE lv_head.
        WHEN '#' OR '?' OR '^'.
          DATA(lv_path) = condense( substring( val = lv_token off = 1 ) ).
          IF strlen( lv_path ) > 0 AND substring( val = lv_path off = strlen( lv_path ) - 1 len = 1 ) = '>'.
            lv_path = substring( val = lv_path len = strlen( lv_path ) - 1 ).  " cell loop {{#x>}}
          ENDIF.
          APPEND to_upper( lv_path ) TO lt_stack.
        WHEN '/'.
          DATA(lv_close) = to_upper( condense( substring( val = lv_token off = 1 ) ) ).
          DATA(lv_top_idx) = lines( lt_stack ).
          IF lv_top_idx = 0.
            APPEND VALUE #( severity = 'E'
                            text = `Placeholder đóng '` && lv_close && `' không có mở tương ứng` )
              TO rt_messages.
          ELSE.
            READ TABLE lt_stack INDEX lv_top_idx INTO DATA(lv_top).
            IF lv_top <> lv_close.
              APPEND VALUE #( severity = 'E'
                              text = |Cặp placeholder lệch: mở '{ lv_top }' nhưng đóng '{ lv_close }'| )
                TO rt_messages.
            ENDIF.
            DELETE lt_stack INDEX lv_top_idx.
          ENDIF.
      ENDCASE.
    ENDLOOP.
    LOOP AT lt_stack INTO DATA(lv_open).
      APPEND VALUE #( severity = 'E'
                      text = `Placeholder mở '` && lv_open && `' không được đóng` )
        TO rt_messages.
    ENDLOOP.

    " 2) ô kiểu Number/DateTime chứa placeholder: engine render vẫn chạy
    "    (tự gán kiểu theo giá trị) nhưng mở TEMPLATE THÔ bằng Excel sẽ bị
    "    "The file is corrupt and cannot be opened" — bắt ngay lúc upload
    IF to_upper( iv_engine ) <> 'XLSX'.
      TRY.
          DATA(lv_raw_text) = to_text( iv_template = iv_template iv_engine = iv_engine ).
          FIND ALL OCCURRENCES OF PCRE
            '(?-x)ss:Type="(Number|DateTime)">\s*[^<]*\{\{'
            IN lv_raw_text MATCH COUNT DATA(lv_typed).
          IF lv_typed > 0.
            APPEND VALUE #( severity = 'W'
                            text = |{ lv_typed } ô kiểu Number/DateTime chứa placeholder — | &&
                                   |đổi ô về Type String (engine tự gán kiểu lúc render), | &&
                                   |nếu không Excel sẽ báo corrupt khi mở file template| )
              TO rt_messages.
          ENDIF.
        CATCH zcx_xlwb ##NO_HANDLER.
      ENDTRY.
    ENDIF.

    " 3) đối chiếu path với sample context (chỉ khi có sample JSON)
    IF iv_sample_json IS INITIAL.
      RETURN.
    ENDIF.

    DATA lr_ctx TYPE REF TO data.
    TRY.
        lr_ctx = zcl_xlwb_ctx_json=>parse( iv_sample_json ).
      CATCH zcx_xlwb INTO DATA(lx_json).
        APPEND VALUE #( severity = 'E' text = |Sample JSON lỗi: { lx_json->get_text( ) }| )
          TO rt_messages.
        RETURN.
    ENDTRY.

    " mọi frame có thể chứa path: context gốc + line type của mọi bảng lồng nhau
    DATA lt_frames TYPE ty_frames.
    DATA(lo_root) = cl_abap_typedescr=>describe_by_data_ref( lr_ctx ).
    APPEND lo_root TO lt_frames.
    collect_frames( EXPORTING io_type = lo_root CHANGING ct_frames = lt_frames ).

    LOOP AT lt_tokens INTO lv_token.
      lv_head = lv_token(1).
      DATA(lv_check) = ``.

      CASE lv_head.
        WHEN '*' OR '@'.
          CONTINUE.                                    " marker/biến hệ thống — bỏ qua
        WHEN '/'.
          CONTINUE.
        WHEN '#' OR '?' OR '^'.
          lv_check = condense( substring( val = lv_token off = 1 ) ).
          IF strlen( lv_check ) > 0 AND substring( val = lv_check off = strlen( lv_check ) - 1 len = 1 ) = '>'.
            lv_check = substring( val = lv_check len = strlen( lv_check ) - 1 ).
          ENDIF.
        WHEN OTHERS.
          lv_check = lv_token.
          " aggregate sum:/avg:/min:/max:/cnt:
          IF lv_check CS ':'.
            SPLIT lv_check AT ':' INTO DATA(lv_fn) DATA(lv_rest).
            IF to_upper( lv_fn ) = 'SUM' OR to_upper( lv_fn ) = 'AVG'
            OR to_upper( lv_fn ) = 'MIN' OR to_upper( lv_fn ) = 'MAX'
            OR to_upper( lv_fn ) = 'CNT'.
              lv_check = lv_rest.
            ELSE.
              CONTINUE.
            ENDIF.
          ENDIF.
      ENDCASE.

      IF lv_check IS INITIAL.
        CONTINUE.
      ENDIF.

      SPLIT to_upper( lv_check ) AT '.' INTO TABLE DATA(lt_segments).
      DATA(lv_found) = abap_false.
      LOOP AT lt_frames INTO DATA(lo_frame).
        IF path_exists( io_type = lo_frame it_segments = lt_segments iv_from = 1 ) = abap_true.
          lv_found = abap_true.
          EXIT.
        ENDIF.
      ENDLOOP.

      IF lv_found = abap_false.
        APPEND VALUE #( severity = 'W'
                        text = |Path '{ lv_check }' không có trong sample context| )
          TO rt_messages.
      ENDIF.
    ENDLOOP.

    SORT rt_messages BY severity text.
    DELETE ADJACENT DUPLICATES FROM rt_messages COMPARING severity text.
  ENDMETHOD.


  METHOD collect_frames.
    " gom line type của mọi bảng (mọi cấp) — mỗi line struct là 1 frame
    CASE io_type->kind.
      WHEN cl_abap_typedescr=>kind_struct.
        LOOP AT CAST cl_abap_structdescr( io_type )->components INTO DATA(ls_comp).
          DATA(lo_comp) = CAST cl_abap_structdescr( io_type )->get_component_type( ls_comp-name ).
          IF lo_comp->kind = cl_abap_typedescr=>kind_table
          OR lo_comp->kind = cl_abap_typedescr=>kind_struct.
            collect_frames( EXPORTING io_type = lo_comp CHANGING ct_frames = ct_frames ).
          ENDIF.
        ENDLOOP.
      WHEN cl_abap_typedescr=>kind_table.
        DATA(lo_line) = CAST cl_abap_tabledescr( io_type )->get_table_line_type( ).
        APPEND lo_line TO ct_frames.
        collect_frames( EXPORTING io_type = lo_line CHANGING ct_frames = ct_frames ).
    ENDCASE.
  ENDMETHOD.


  METHOD path_exists.
    READ TABLE it_segments INDEX iv_from INTO DATA(lv_seg).
    IF sy-subrc <> 0.
      rv_exists = abap_true.                          " hết segment = resolve xong
      RETURN.
    ENDIF.

    DATA(lo_type) = io_type.
    IF lo_type->kind = cl_abap_typedescr=>kind_table.
      lo_type = CAST cl_abap_tabledescr( lo_type )->get_table_line_type( ).
    ENDIF.

    IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
      rv_exists = abap_false.
      RETURN.
    ENDIF.

    DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
    IF NOT line_exists( lo_struct->components[ name = lv_seg ] ).
      rv_exists = abap_false.
      RETURN.
    ENDIF.

    rv_exists = path_exists( io_type     = lo_struct->get_component_type( lv_seg )
                             it_segments = it_segments
                             iv_from     = iv_from + 1 ).
  ENDMETHOD.


  METHOD export_json.
    DATA(lv_b64) = cl_web_http_utility=>encode_x_base64( iv_template ).

    rv_json =
      `{` && cl_abap_char_utilities=>newline &&
      `  "xlwb_export": 1,` && cl_abap_char_utilities=>newline &&
      `  "form_name": "` && json_escape( CONV #( is_tmpl-form_name ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "descr": "` && json_escape( CONV #( is_tmpl-descr ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "engine": "` && json_escape( CONV #( is_tmpl-engine ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "mime_type": "` && json_escape( CONV #( is_tmpl-mime_type ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "file_name": "` && json_escape( CONV #( is_tmpl-file_name ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "source_class": "` && json_escape( CONV #( is_tmpl-source_class ) ) && `",` && cl_abap_char_utilities=>newline &&
      `  "sample_json": "` && json_escape( is_tmpl-sample_json ) && `",` && cl_abap_char_utilities=>newline &&
      `  "template_b64": "` && lv_b64 && `"` && cl_abap_char_utilities=>newline &&
      `}`.
  ENDMETHOD.


  METHOD import_json.
    CLEAR: es_tmpl, ev_template.

    DATA(lr_data) = zcl_xlwb_ctx_json=>parse( iv_json ).

    es_tmpl-form_name    = json_get_string( ir_data = lr_data iv_comp = 'FORM_NAME' ).
    es_tmpl-descr        = json_get_string( ir_data = lr_data iv_comp = 'DESCR' ).
    es_tmpl-engine       = json_get_string( ir_data = lr_data iv_comp = 'ENGINE' ).
    es_tmpl-mime_type    = json_get_string( ir_data = lr_data iv_comp = 'MIME_TYPE' ).
    es_tmpl-file_name    = json_get_string( ir_data = lr_data iv_comp = 'FILE_NAME' ).
    es_tmpl-source_class = json_get_string( ir_data = lr_data iv_comp = 'SOURCE_CLASS' ).
    es_tmpl-sample_json  = json_get_string( ir_data = lr_data iv_comp = 'SAMPLE_JSON' ).

    IF es_tmpl-form_name IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |JSON import thiếu "form_name"| ).
    ENDIF.

    DATA(lv_b64) = json_get_string( ir_data = lr_data iv_comp = 'TEMPLATE_B64' ).
    IF lv_b64 IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |JSON import thiếu "template_b64"| ).
    ENDIF.
    ev_template = cl_web_http_utility=>decode_x_base64( lv_b64 ).
    IF ev_template IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |"template_b64" không giải mã được base64| ).
    ENDIF.
  ENDMETHOD.


  METHOD json_get_string.
    ASSIGN ir_data->* TO FIELD-SYMBOL(<ls_any>).
    ASSIGN COMPONENT iv_comp OF STRUCTURE <ls_any> TO FIELD-SYMBOL(<lv_val>).
    IF sy-subrc = 0.
      rv_value = |{ <lv_val> }|.
    ENDIF.
  ENDMETHOD.


  METHOD json_escape.
    rv_escaped = iv_text.
    REPLACE ALL OCCURRENCES OF '\' IN rv_escaped WITH '\\'.
    REPLACE ALL OCCURRENCES OF '"' IN rv_escaped WITH '\"'.
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf(1) IN rv_escaped WITH '\r'.
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN rv_escaped WITH '\n'.
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>horizontal_tab IN rv_escaped WITH '\t'.
  ENDMETHOD.
ENDCLASS.
