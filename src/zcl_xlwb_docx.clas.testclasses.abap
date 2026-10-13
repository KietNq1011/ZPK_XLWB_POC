*"* use this source file for your ABAP unit test classes
CLASS ltc_docx DEFINITION FINAL FOR TESTING DURATION SHORT RISK LEVEL HARMLESS.
  PRIVATE SECTION.
    TYPES: BEGIN OF ty_item,
             product TYPE string,
             qty     TYPE i,
           END OF ty_item,
           tt_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_ctx,
             name  TYPE string,
             items TYPE tt_items,
           END OF ty_ctx.

    DATA mo_cut TYPE REF TO zcl_xlwb_docx.

    METHODS setup.
    METHODS make_zip_entry
      IMPORTING io_zip     TYPE REF TO cl_abap_zip
                iv_name    TYPE string
                iv_content TYPE string.
    METHODS build_template
      RETURNING VALUE(rv_zip) TYPE xstring.

    METHODS render_swaps_part      FOR TESTING RAISING cx_static_check.
    METHODS prepare_from_ddic      FOR TESTING RAISING cx_static_check.
    METHODS render_escapes_xml     FOR TESTING RAISING cx_static_check.
    METHODS skips_builtin_ms_part  FOR TESTING RAISING cx_static_check.
    METHODS missing_part_raises    FOR TESTING.
ENDCLASS.


CLASS ltc_docx IMPLEMENTATION.

  METHOD setup.
    mo_cut = NEW #( ).
  ENDMETHOD.

  METHOD make_zip_entry.
    io_zip->add( name    = iv_name
                 content = cl_abap_conv_codepage=>create_out( )->convert( iv_content ) ).
  ENDMETHOD.

  METHOD build_template.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    make_zip_entry( io_zip = lo_zip iv_name = `word/document.xml`
                    iv_content = `<w:document/>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `customXml/itemProps1.xml`
                    iv_content = `<x/>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `customXml/item1.xml`
                    iv_content = `<?xml version="1.0" encoding="UTF-8"?><data xmlns="urn:test">old</data>` ).
    rv_zip = lo_zip->save( ).
  ENDMETHOD.

  METHOD render_swaps_part.
    DATA(ls_ctx) = VALUE ty_ctx(
      name  = `Công ty ABC`
      items = VALUE #( ( product = `SP1` qty = 2 )
                       ( product = `SP2` qty = 3 ) ) ).

    DATA(lv_out) = mo_cut->render( iv_template = build_template( )
                                   ir_context  = REF #( ls_ctx ) ).

    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( lv_out ).
    lo_zip->get( EXPORTING name = `customXml/item1.xml` IMPORTING content = DATA(lv_content) ).
    DATA(lv_xml) = cl_abap_conv_codepage=>create_in( )->convert( lv_content ).

    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<data xmlns="urn:test">*` ).
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<name>Công ty ABC</name>*` ).
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml
      exp = `*<items><row><product>SP1</product><qty>2</qty></row>` &&
            `<row><product>SP2</product><qty>3</qty></row></items>*` ).

    " document.xml phải giữ nguyên, không bị engine đụng vào
    lo_zip->get( EXPORTING name = `word/document.xml` IMPORTING content = DATA(lv_doc) ).
    cl_abap_unit_assert=>assert_equals(
      act = cl_abap_conv_codepage=>create_in( )->convert( lv_doc )
      exp = `<w:document/>` ).
  ENDMETHOD.

  METHOD prepare_from_ddic.
    " docx tối giản KHÔNG có Default xml trong content types + CHƯA có part nào
    DATA(lo_zip) = NEW cl_abap_zip( ).
    make_zip_entry( io_zip = lo_zip iv_name = `[Content_Types].xml`
      iv_content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
                   `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` &&
                   `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` &&
                   `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` &&
                   `</Types>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `word/document.xml`
                    iv_content = `<w:document/>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `word/_rels/document.xml.rels`
      iv_content = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` &&
                   `<Relationship Id="rId1" Type="x" Target="y"/></Relationships>` ).

    NEW zcl_xlwb_docx( )->prepare_docx(
      EXPORTING iv_docx        = lo_zip->save( )
                iv_spec        = |header=ZXLWB_TMPL; table=items:ZXLWB_TMPL|
      IMPORTING ev_docx        = DATA(lv_docx)
                ev_sample_json = DATA(lv_json) ).

    DATA(lo_res) = NEW cl_abap_zip( ).
    lo_res->load( lv_docx ).

    lo_res->get( EXPORTING name = `customXml/item1.xml` IMPORTING content = DATA(lv_part) ).
    DATA(lv_xml) = cl_abap_conv_codepage=>create_in( )->convert( lv_part ).
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<data xmlns="urn:zdocx:data">*` ).
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<form_name>form_name</form_name>*` ).
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<items><row>*` ).

    lo_res->get( EXPORTING name = `[Content_Types].xml` IMPORTING content = DATA(lv_ct) ).
    DATA(lv_ct_s) = cl_abap_conv_codepage=>create_in( )->convert( lv_ct ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_ct_s exp = `*Extension="xml"*` ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_ct_s exp = `*/customXml/itemProps1.xml*` ).

    lo_res->get( EXPORTING name = `word/_rels/document.xml.rels` IMPORTING content = DATA(lv_rels) ).
    cl_abap_unit_assert=>assert_char_cp(
      act = cl_abap_conv_codepage=>create_in( )->convert( lv_rels )
      exp = `*Id="rId2"*customXml/item1.xml*` ).

    cl_abap_unit_assert=>assert_char_cp( act = lv_json exp = `*"items":[*` ).
    cl_abap_unit_assert=>assert_char_cp( act = lv_json exp = `*"form_name":"form_name"*` ).
  ENDMETHOD.

  METHOD render_escapes_xml.
    DATA(ls_ctx) = VALUE ty_ctx( name = `A & <B>` ).

    DATA(lv_xml) = mo_cut->build_data_xml( iv_root_name = `data`
                                           iv_namespace = `urn:test`
                                           ir_context   = REF #( ls_ctx ) ).

    " & và < BẮT BUỘC escape; > tùy chọn theo chuẩn XML nên không assert
    cl_abap_unit_assert=>assert_char_cp(
      act = lv_xml exp = `*<name>A &amp; &lt;B*` ).
    IF lv_xml CS `<B>`.
      cl_abap_unit_assert=>fail( `Ký tự < trong dữ liệu chưa được escape` ).
    ENDIF.
  ENDMETHOD.

  METHOD skips_builtin_ms_part.
    " item1 = part built-in của Word, item2 = part data -> engine phải chọn item2
    DATA(lo_zip) = NEW cl_abap_zip( ).
    make_zip_entry( io_zip = lo_zip iv_name = `word/document.xml`
                    iv_content = `<w:document/>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `customXml/item1.xml`
                    iv_content = `<p xmlns="http://schemas.microsoft.com/office/2006/coverPageProps"/>` ).
    make_zip_entry( io_zip = lo_zip iv_name = `customXml/item2.xml`
                    iv_content = `<data xmlns="urn:test">old</data>` ).

    DATA(ls_ctx) = VALUE ty_ctx( name = `X` ).
    DATA(lv_out) = mo_cut->render( iv_template = lo_zip->save( )
                                   ir_context  = REF #( ls_ctx ) ).

    DATA(lo_res) = NEW cl_abap_zip( ).
    lo_res->load( lv_out ).

    lo_res->get( EXPORTING name = `customXml/item2.xml` IMPORTING content = DATA(lv_item2) ).
    cl_abap_unit_assert=>assert_char_cp(
      act = cl_abap_conv_codepage=>create_in( )->convert( lv_item2 )
      exp = `*<name>X</name>*` ).

    lo_res->get( EXPORTING name = `customXml/item1.xml` IMPORTING content = DATA(lv_item1) ).
    cl_abap_unit_assert=>assert_char_cp(
      act = cl_abap_conv_codepage=>create_in( )->convert( lv_item1 )
      exp = `*coverPageProps*` ).
  ENDMETHOD.

  METHOD missing_part_raises.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    make_zip_entry( io_zip = lo_zip iv_name = `word/document.xml`
                    iv_content = `<w:document/>` ).
    DATA ls_ctx TYPE ty_ctx.
    TRY.
        mo_cut->render( iv_template = lo_zip->save( )
                        ir_context  = REF #( ls_ctx ) ).
        cl_abap_unit_assert=>fail( `Phải raise ZCX_XLWB khi template không có custom XML part` ).
      CATCH zcx_xlwb.
    ENDTRY.
  ENDMETHOD.

ENDCLASS.

