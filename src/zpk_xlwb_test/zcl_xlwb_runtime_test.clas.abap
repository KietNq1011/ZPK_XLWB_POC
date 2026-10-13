"! <p class="shorttext synchronized" lang="en">XLWB: test runtime + BO template</p>
"!
"! Bảng ZXLWB_TMPL được thay bằng test double (CL_OSQL_TEST_ENVIRONMENT) nên
"! test KHÔNG ghi dữ liệu thật và chạy được ở mức HARMLESS. Phần BO
"! Determination/validation của BO KHÔNG test ở đây: EML trên draft BO trong
"! ABAP Unit cần cấu hình test environment riêng — kiểm chứng bằng app Fiori.
CLASS zcl_xlwb_runtime_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    CONSTANTS c_form TYPE c LENGTH 40 VALUE 'ZZ_UNITTEST_TMPL'.

  PRIVATE SECTION.
    CLASS-DATA mo_osql TYPE REF TO if_osql_test_environment.

    CLASS-METHODS class_setup.
    CLASS-METHODS class_teardown.
    CLASS-METHODS ssml_template RETURNING VALUE(rv_xml) TYPE string.
    CLASS-METHODS fallback_template RETURNING VALUE(rv_x) TYPE xstring.

    "! Output giờ là .xlsx (zip) — nối text mọi part XML để assert nội dung;
    "! nếu là SpreadsheetML thô (fallback lỗi converter) thì convert utf-8.
    CLASS-METHODS out_text
      IMPORTING iv_content     TYPE xstring
      RETURNING VALUE(rv_text) TYPE string.

    METHODS teardown.
    METHODS insert_double
      IMPORTING iv_engine    TYPE c
                iv_active    TYPE abap_bool DEFAULT abap_true
                iv_template  TYPE xstring OPTIONAL.

    METHODS render_from_table   FOR TESTING.
    METHODS unknown_form        FOR TESTING.
    METHODS inactive_form       FOR TESTING.
    METHODS render_with_tmpl    FOR TESTING.
    METHODS prefer_uses_table   FOR TESTING.
    METHODS prefer_falls_back   FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_RUNTIME_TEST IMPLEMENTATION.


  METHOD class_setup.
    mo_osql = cl_osql_test_environment=>create( VALUE #( ( 'ZXLWB_TMPL' ) ) ).
  ENDMETHOD.


  METHOD class_teardown.
    mo_osql->destroy( ).
  ENDMETHOD.


  METHOD teardown.
    mo_osql->clear_doubles( ).
    zcl_xlwb_runtime=>clear_cache( ).
  ENDMETHOD.


  METHOD ssml_template.
    rv_xml =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Worksheet ss:Name="T"><Table>` &&
      `<Row><Cell><Data ss:Type="String">Hello {{customer}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#items}}{{matnr}}{{/items}}</Data></Cell></Row>` &&
      `</Table></Worksheet></Workbook>`.
  ENDMETHOD.


  METHOD fallback_template.
    rv_x = cl_abap_conv_codepage=>create_out( )->convert(
      `<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Worksheet ss:Name="F"><Table><Row><Cell>` &&
      `<Data ss:Type="String">FALLBACK {{customer}}</Data></Cell></Row></Table>` &&
      `</Worksheet></Workbook>` ).
  ENDMETHOD.


  METHOD out_text.
    DATA lv_magic TYPE x LENGTH 2.
    IF xstrlen( iv_content ) >= 2.
      lv_magic = iv_content.
    ENDIF.
    IF lv_magic <> CONV xstring( '504B' ).
      rv_text = cl_abap_conv_codepage=>create_in( )->convert( iv_content ).
      RETURN.
    ENDIF.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_content ).
    LOOP AT lo_zip->files INTO DATA(ls_f).
      lo_zip->get( EXPORTING name = ls_f-name
                   IMPORTING content = DATA(lv_raw)
                   EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
      IF sy-subrc = 0 AND lv_raw IS NOT INITIAL.
        rv_text &&= cl_abap_conv_codepage=>create_in( )->convert( lv_raw ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD insert_double.
    DATA lt_tmpl TYPE STANDARD TABLE OF zxlwb_tmpl WITH EMPTY KEY.

    DATA(lv_content) = COND xstring(
      WHEN iv_template IS SUPPLIED THEN iv_template
      ELSE cl_abap_conv_codepage=>create_out( )->convert( ssml_template( ) ) ).

    lt_tmpl = VALUE #( ( form_name = c_form
                         descr     = 'unit test'
                         engine    = iv_engine
                         mime_type = 'application/vnd.ms-excel'
                         file_name = 'unittest.xls'
                         is_active = iv_active
                         template  = lv_content ) ).
    mo_osql->insert_test_data( lt_tmpl ).
    zcl_xlwb_runtime=>clear_cache( ).
  ENDMETHOD.


  METHOD render_from_table.
    insert_double( iv_engine = 'SSML' ).

    TYPES: BEGIN OF ty_item,
             matnr TYPE string,
           END OF ty_item.
    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'ACME',
            items    TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( matnr = 'M-1' ) ( matnr = 'M-2' ) ).

    TRY.
        DATA(ls_file) = zcl_xlwb_runtime=>render( iv_form_name = CONV string( c_form )
                                                  ir_context   = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_out) = out_text( ls_file-content ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-extension exp = `xlsx`
                                        msg = `engine SSML -> output .xlsx (ssml2xlsx)` ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-file_name exp = `unittest.xlsx`
                                        msg = `file name bảng được đổi đuôi theo output` ).
    cl_abap_unit_assert=>assert_true( act = xsdbool( lv_out CS `Hello ACME` )
                                      msg = `render giá trị đơn` ).
    cl_abap_unit_assert=>assert_true( act = xsdbool( lv_out CS `M-2` )
                                      msg = `render vùng lặp` ).
  ENDMETHOD.


  METHOD unknown_form.
    DATA: BEGIN OF ls_ctx,
            dummy TYPE string VALUE 'x',
          END OF ls_ctx.

    TRY.
        zcl_xlwb_runtime=>render( iv_form_name = `ZZ_KHONG_TON_TAI_XYZ`
                                  ir_context   = REF #( ls_ctx ) ).
        cl_abap_unit_assert=>fail( msg = `phải raise zcx_xlwb khi không có template` ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>assert_true(
          act = xsdbool( lx->get_text( ) CS `ZZ_KHONG_TON_TAI_XYZ` )
          msg = `message phải nêu tên template` ).
    ENDTRY.
  ENDMETHOD.


  METHOD inactive_form.
    insert_double( iv_engine = 'SSML' iv_active = abap_false ).

    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'X',
          END OF ls_ctx.

    TRY.
        zcl_xlwb_runtime=>render( iv_form_name = CONV string( c_form )
                                  ir_context   = REF #( ls_ctx ) ).
        cl_abap_unit_assert=>fail( msg = `template inactive phải bị từ chối` ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>assert_true( act = xsdbool( lx->get_text( ) CS `active` )
                                          msg = `message phải nói template không active` ).
    ENDTRY.
  ENDMETHOD.


  METHOD render_with_tmpl.
    " Preview: render trực tiếp từ nội dung, không cần bản ghi trong bảng
    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'PREVIEW',
            items    TYPE string_table,
          END OF ls_ctx.

    TRY.
        DATA(ls_file) = zcl_xlwb_runtime=>render_with_template(
          iv_template = cl_abap_conv_codepage=>create_out( )->convert(
            `<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
            ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
            ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
            `<Worksheet ss:Name="P"><Table><Row><Cell>` &&
            `<Data ss:Type="String">Hi {{customer}}</Data></Cell></Row></Table>` &&
            `</Worksheet></Workbook>` )
          iv_engine   = `SSML`
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_out) = out_text( ls_file-content ).
    cl_abap_unit_assert=>assert_true( act = xsdbool( lv_out CS `Hi PREVIEW` )
                                      msg = `preview render` ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-file_name exp = `preview.xlsx`
                                        msg = `tên file preview` ).
  ENDMETHOD.


  METHOD prefer_uses_table.
    " Co ban ghi trong bang -> phai dung bang, KHONG dung template du phong
    insert_double( iv_engine = 'SSML' ).

    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'FROM-TABLE',
            items    TYPE string_table,
          END OF ls_ctx.

    TRY.
        DATA(ls_file) = zcl_xlwb_runtime=>render_prefer_table(
          iv_form_name         = CONV string( c_form )
          iv_fallback_template = fallback_template( )
          ir_context           = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_out) = out_text( ls_file-content ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-source
                                        exp = |ZXLWB_TMPL:{ c_form }|
                                        msg = `source phai chi ra bang (co note = converter loi)` ).
    cl_abap_unit_assert=>assert_true( act = xsdbool( lv_out CS `Hello FROM-TABLE` )
                                      msg = `phai render template cua bang` ).
    cl_abap_unit_assert=>assert_false( act = xsdbool( lv_out CS `FALLBACK` )
                                       msg = `khong duoc dung template du phong` ).
  ENDMETHOD.


  METHOD prefer_falls_back.
    " Bang trong (double da clear) -> phai quay ve template du phong
    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'FROM-CODE',
            items    TYPE string_table,
          END OF ls_ctx.

    TRY.
        DATA(ls_file) = zcl_xlwb_runtime=>render_prefer_table(
          iv_form_name         = CONV string( c_form )
          iv_fallback_template = fallback_template( )
          ir_context           = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_out) = out_text( ls_file-content ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-source exp = `CODE`
                                        msg = `source phai la CODE (co note = converter loi)` ).
    cl_abap_unit_assert=>assert_true( act = xsdbool( lv_out CS `FALLBACK FROM-CODE` )
                                      msg = `phai render template du phong` ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-file_name
                                        exp = to_lower( CONV string( c_form ) ) && `.xlsx`
                                        msg = `ten file suy ra tu form name` ).
  ENDMETHOD.
ENDCLASS.
