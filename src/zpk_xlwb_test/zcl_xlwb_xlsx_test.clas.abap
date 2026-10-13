"! <p class="shorttext synchronized" lang="en">XLWB Cloud: XLSX engine unit tests</p>
"!
"! Template .xlsx dùng trong test được sinh bằng XCO XLSX write API để test
"! không phụ thuộc file ngoài. Trong dự án thật, template do user Save As từ
"! Excel và lưu trong bảng template.
CLASS zcl_xlwb_xlsx_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_item,
             matnr TYPE string,
             qty   TYPE p LENGTH 8 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY.

    "! 1x1 PNG trong suốt — chỗ này dự án thật đưa logo công ty vào
    CLASS-METHODS logo_base64 RETURNING VALUE(rv_b64) TYPE string.

  PRIVATE SECTION.
    METHODS build_template
      IMPORTING it_cells         TYPE string_table
      RETURNING VALUE(rv_xlsx)   TYPE xstring.
    METHODS cell_of
      IMPORTING iv_xlsx        TYPE xstring
                iv_col         TYPE string
                iv_row         TYPE i
      RETURNING VALUE(rv_text) TYPE string.
    METHODS part_exists
      IMPORTING iv_xlsx          TYPE xstring
                iv_name          TYPE string
      RETURNING VALUE(rv_exists) TYPE abap_bool.
    METHODS part_text
      IMPORTING iv_xlsx        TYPE xstring
                iv_name        TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    METHODS value_and_number FOR TESTING.
    METHODS row_loop         FOR TESTING.
    METHODS image_embedded   FOR TESTING.
    METHODS still_valid_xlsx FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_XLSX_TEST IMPLEMENTATION.


  METHOD logo_base64.
    rv_b64 = `iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==`.
  ENDMETHOD.


  METHOD build_template.
    " it_cells: mỗi dòng của bảng = 1 dòng Excel, cột A
    DATA(lo_wa) = xco_cp_xlsx=>document->empty( )->write_access( ).
    DATA(lo_ws) = lo_wa->get_workbook( )->worksheet->at_position( 1 ).
    lo_ws->set_name( `FORM` ).

    LOOP AT it_cells INTO DATA(lv_text).
      lo_ws->cursor(
        io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
        io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( sy-tabix )
        )->get_cell( )->value->write_from( lv_text ).
    ENDLOOP.

    rv_xlsx = lo_wa->get_file_content( ).
  ENDMETHOD.


  METHOD cell_of.
    DATA(lo_ra) = xco_cp_xlsx=>document->for_file_content( iv_xlsx )->read_access( ).
    DATA(lo_cell) = lo_ra->get_workbook( )->worksheet->at_position( 1 )->cursor(
      io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( iv_col )
      io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( iv_row )
      )->get_cell( ).
    IF lo_cell->has_value( ) = abap_false.
      RETURN.
    ENDIF.
    DATA lv_text TYPE string.
    lo_cell->get_value( )->write_to( REF #( lv_text ) ).
    rv_text = lv_text.
  ENDMETHOD.


  METHOD part_exists.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_xlsx ).
    LOOP AT lo_zip->files INTO DATA(ls_file).
      IF ls_file-name CS iv_name.
        rv_exists = abap_true.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD part_text.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_xlsx ).
    lo_zip->get( EXPORTING name            = iv_name
                 IMPORTING content         = DATA(lv_raw)
                 EXCEPTIONS zip_index_error = 1
                            OTHERS          = 2 ).
    IF sy-subrc = 0 AND lv_raw IS NOT INITIAL.
      rv_text = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8' )->convert( lv_raw ).
    ENDIF.
  ENDMETHOD.


  METHOD value_and_number.
    DATA(lv_tpl) = build_template( VALUE #(
      ( `Customer: {{customer}}` )
      ( `{{amount}}` ) ) ).

    DATA: BEGIN OF ls_ctx,
            customer TYPE string VALUE 'ACME Corp',
            amount   TYPE p LENGTH 8 DECIMALS 2 VALUE '1234.50',
          END OF ls_ctx.

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
    ENDTRY.

    DATA(lv_sheet) = part_text( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    DATA(lv_ss)    = part_text( iv_xlsx = lv_out iv_name = 'xl/sharedStrings.xml' ).

    cl_abap_unit_assert=>assert_equals(
      act = cell_of( iv_xlsx = lv_out iv_col = 'A' iv_row = 1 )
      exp = `Customer: ACME Corp`
      msg = |text substitution. sheet=[{ lv_sheet }] shared=[{ lv_ss }]| ).

    " ô toàn placeholder số → ghi dạng số, Excel tính toán được
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_sheet CS `<v>1234.5` )
      msg = |numeric cell must be written as <v>, sheet: { lv_sheet }| ).
  ENDMETHOD.


  METHOD row_loop.
    DATA(lv_tpl) = build_template( VALUE #(
      ( `ITEMS` )
      ( `{{#items}}{{matnr}} = {{qty}}{{/items}}` )
      ( `Total: {{sum:items.qty}}` ) ) ).

    DATA: BEGIN OF ls_ctx,
            items TYPE ty_items,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( matnr = 'SLAB-1' qty = '10.00' )
                            ( matnr = 'SLAB-2' qty = '20.00' )
                            ( matnr = 'SLAB-3' qty = '5.50' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
    ENDTRY.

    cl_abap_unit_assert=>assert_equals(
      act = cell_of( iv_xlsx = lv_out iv_col = 'A' iv_row = 2 )
      exp = `SLAB-1 = 10`
      msg = `first loop line` ).
    cl_abap_unit_assert=>assert_equals(
      act = cell_of( iv_xlsx = lv_out iv_col = 'A' iv_row = 4 )
      exp = `SLAB-3 = 5.5`
      msg = `third loop line — rows must be renumbered` ).
    cl_abap_unit_assert=>assert_equals(
      act = cell_of( iv_xlsx = lv_out iv_col = 'A' iv_row = 5 )
      exp = `Total: 35.5`
      msg = `footer moved down by 2 rows, aggregate computed` ).
  ENDMETHOD.


  METHOD image_embedded.
    DATA(lv_tpl) = build_template( VALUE #(
      ( `{{*image:logo}}` )
      ( `Report {{title}}` ) ) ).

    DATA: BEGIN OF ls_ctx,
            logo  TYPE string,
            title TYPE string VALUE 'Q3',
          END OF ls_ctx.
    ls_ctx-logo = logo_base64( ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
    ENDTRY.

    cl_abap_unit_assert=>assert_true(
      act = part_exists( iv_xlsx = lv_out iv_name = 'xl/media/' )
      msg = `PNG must be added to xl/media` ).
    cl_abap_unit_assert=>assert_true(
      act = part_exists( iv_xlsx = lv_out iv_name = 'xl/drawings/drawing1.xml' )
      msg = `drawing part must exist` ).

    DATA(lv_dw) = part_text( iv_xlsx = lv_out iv_name = 'xl/drawings/drawing1.xml' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_dw CS `twoCellAnchor` AND lv_dw CS `blip` )
      msg = `drawing must anchor a picture` ).

    DATA(lv_rels) = part_text( iv_xlsx = lv_out
                               iv_name = 'xl/worksheets/_rels/sheet1.xml.rels' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_rels CS `drawing1.xml` )
      msg = `sheet rels must point to the drawing` ).

    DATA(lv_types) = part_text( iv_xlsx = lv_out iv_name = '[Content_Types].xml' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_types CS `image/png` )
      msg = |png content type must be registered: { lv_types }| ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_types CS `drawing1.xml` )
      msg = `drawing override must be registered` ).

    DATA(lv_sheet) = part_text( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_sheet CS `<drawing` )
      msg = `worksheet must reference the drawing` ).

    " marker bị strip, nội dung khác vẫn render.
    " Đọc thẳng part XML: XCO read access không hỗ trợ worksheet có <drawing>.
    DATA(lv_ss) = part_text( iv_xlsx = lv_out iv_name = 'xl/sharedStrings.xml' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_ss CS `Report Q3` )
      msg = |other cells still rendered. shared=[{ lv_ss }]| ).
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_sheet CS `{{` )
      msg = `image marker must be stripped from the cell` ).
  ENDMETHOD.


  METHOD still_valid_xlsx.
    " file kết quả phải mở lại được bằng XCO (chứng minh OOXML còn hợp lệ)
    DATA(lv_tpl) = build_template( VALUE #( ( `{{a}}` ) ( `{{#t}}{{x}}{{/t}}` ) ) ).

    TYPES: BEGIN OF ty_line,
             x TYPE string,
           END OF ty_line.
    DATA: BEGIN OF ls_ctx,
            a TYPE string VALUE 'ok',
            t TYPE STANDARD TABLE OF ty_line WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-t = VALUE #( ( x = 'l1' ) ( x = 'l2' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
    ENDTRY.

    DATA(lo_ra) = xco_cp_xlsx=>document->for_file_content( lv_out )->read_access( ).
    cl_abap_unit_assert=>assert_true(
      act = lo_ra->get_workbook( )->worksheet->at_position( 1 )->exists( )
      msg = `rendered file must still be a readable xlsx` ).
    cl_abap_unit_assert=>assert_equals(
      act = cell_of( iv_xlsx = lv_out iv_col = 'A' iv_row = 3 )
      exp = `l2`
      msg = `second loop line` ).
  ENDMETHOD.
ENDCLASS.
