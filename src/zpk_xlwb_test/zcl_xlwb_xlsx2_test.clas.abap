"! <p class="shorttext synchronized" lang="en">XLWB: XLSX engine v2 tests (parity SSML)</p>
"!
"! Test 4 tính năng mới của engine OOXML: cell loop ngang, {{*break}},
"! sheet loop, {{*mergesame>}}.
CLASS zcl_xlwb_xlsx2_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    TYPES ty_row  TYPE string_table.
    TYPES ty_grid TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

    METHODS build_template
      IMPORTING it_grid        TYPE ty_grid
                iv_sheet_name  TYPE string DEFAULT `FORM`
      RETURNING VALUE(rv_xlsx) TYPE xstring.
    METHODS cell_of
      IMPORTING iv_xlsx        TYPE xstring
                iv_col         TYPE string
                iv_row         TYPE i
                iv_sheet       TYPE i DEFAULT 1
      RETURNING VALUE(rv_text) TYPE string.
    METHODS part_text
      IMPORTING iv_xlsx        TYPE xstring
                iv_name        TYPE string
      RETURNING VALUE(rv_text) TYPE string.
    METHODS assert_has
      IMPORTING iv_text TYPE string
                iv_sub  TYPE string
                iv_msg  TYPE string.

    METHODS cell_loop_columns  FOR TESTING.
    METHODS page_break_marker  FOR TESTING.
    METHODS sheet_loop         FOR TESTING.
    METHODS mergesame_across   FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_XLSX2_TEST IMPLEMENTATION.


  METHOD build_template.
    DATA(lo_wa) = xco_cp_xlsx=>document->empty( )->write_access( ).
    DATA(lo_ws) = lo_wa->get_workbook( )->worksheet->at_position( 1 ).
    lo_ws->set_name( iv_sheet_name ).

    DATA(lv_letters) = `ABCDEFGHIJ`.
    LOOP AT it_grid INTO DATA(lt_row).
      DATA(lv_r) = sy-tabix.
      LOOP AT lt_row INTO DATA(lv_text).
        DATA(lv_c) = substring( val = lv_letters off = sy-tabix - 1 len = 1 ).
        lo_ws->cursor(
          io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( lv_c )
          io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( lv_r )
          )->get_cell( )->value->write_from( lv_text ).
      ENDLOOP.
    ENDLOOP.

    rv_xlsx = lo_wa->get_file_content( ).
  ENDMETHOD.


  METHOD cell_of.
    DATA(lo_ra) = xco_cp_xlsx=>document->for_file_content( iv_xlsx )->read_access( ).
    DATA(lo_cell) = lo_ra->get_workbook( )->worksheet->at_position( iv_sheet )->cursor(
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


  METHOD part_text.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_xlsx ).
    lo_zip->get( EXPORTING name = iv_name
                 IMPORTING content = DATA(lv_raw)
                 EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    IF sy-subrc = 0 AND lv_raw IS NOT INITIAL.
      rv_text = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8' )->convert( lv_raw ).
    ENDIF.
  ENDMETHOD.


  METHOD assert_has.
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( iv_text CS iv_sub )
      msg = |{ iv_msg } — missing '{ iv_sub }' in: { substring( val = iv_text off = 0
               len = nmin( val1 = strlen( iv_text ) val2 = 1500 ) ) }| ).
  ENDMETHOD.


  METHOD cell_loop_columns.
    DATA(lv_tpl) = build_template( VALUE #(
      ( VALUE ty_row( ( `Product` ) ( `{{#cols>}}{{title}}` ) ( `END` ) ) ) ) ).

    TYPES: BEGIN OF ty_col, title TYPE string, END OF ty_col.
    DATA: BEGIN OF ls_ctx,
            cols TYPE STANDARD TABLE OF ty_col WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-cols = VALUE #( ( title = 'Jan' ) ( title = 'Feb' ) ( title = 'Mar' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl ir_context = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'B' iv_row = 1 )
                                        exp = `Jan` msg = `cot dau` ).
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'D' iv_row = 1 )
                                        exp = `Mar` msg = `cot cuoi` ).
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'E' iv_row = 1 )
                                        exp = `END` msg = `o sau vung lap phai dich sang phai` ).
  ENDMETHOD.


  METHOD page_break_marker.
    DATA(lv_tpl) = build_template( VALUE #(
      ( VALUE ty_row( ( `{{#labels}}{{*break}}{{name}}` ) ) )
      ( VALUE ty_row( ( `--{{/labels}}` ) ) ) ) ).

    TYPES: BEGIN OF ty_label, name TYPE string, END OF ty_label.
    DATA: BEGIN OF ls_ctx,
            labels TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-labels = VALUE #( ( name = 'L-A' ) ( name = 'L-B' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl ir_context = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_sheet) = part_text( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    " nhan 2: block 2 dong/nhan -> nhan thu 2 bat dau dong 3 -> brk id=2;
    " break truoc dong dau tien bi bo qua
    assert_has( iv_text = lv_sheet iv_sub = `<brk id="2"` iv_msg = `page break nhan 2` ).
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_sheet CS `<brk id="0"` )
      msg = `khong co break truoc dong dau` ).
    " text nam trong sharedStrings, khong nam trong sheet xml
    DATA(lv_sst) = part_text( iv_xlsx = lv_out iv_name = 'xl/sharedStrings.xml' ).
    assert_has( iv_text = lv_sst iv_sub = `L-B` iv_msg = `du lieu nhan 2` ).
  ENDMETHOD.


  METHOD sheet_loop.
    DATA(lv_tpl) = build_template(
      it_grid = VALUE #( ( VALUE ty_row( ( `TO: {{name}}` ) ) ) )
      iv_sheet_name = `{{#labels}}L{{@index}}` ).

    TYPES: BEGIN OF ty_label, name TYPE string, END OF ty_label.
    DATA: BEGIN OF ls_ctx,
            labels TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-labels = VALUE #( ( name = 'ACME' ) ( name = 'Globex' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl ir_context = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_wb) = part_text( iv_xlsx = lv_out iv_name = 'xl/workbook.xml' ).
    assert_has( iv_text = lv_wb iv_sub = `name="L1"` iv_msg = `sheet clone 1` ).
    assert_has( iv_text = lv_wb iv_sub = `name="L2"` iv_msg = `sheet clone 2` ).
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_wb CS `{{#labels}}` )
      msg = `sheet goc phai bi go khoi workbook` ).

    " du lieu tung sheet: sheet 1 = ACME, sheet 2 = Globex
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'A'
                                                       iv_row = 1 iv_sheet = 1 )
                                        exp = `TO: ACME` msg = `sheet 1` ).
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'A'
                                                       iv_row = 1 iv_sheet = 2 )
                                        exp = `TO: Globex` msg = `sheet 2` ).
  ENDMETHOD.


  METHOD mergesame_across.
    DATA(lv_tpl) = build_template( VALUE #(
      ( VALUE ty_row( ( `{{#items>}}{{*mergesame>:grp}}{{grp}}` ) ( `END` ) ) ) ) ).

    TYPES: BEGIN OF ty_item, grp TYPE string, END OF ty_item.
    DATA: BEGIN OF ls_ctx,
            items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( grp = 'A' ) ( grp = 'A' ) ( grp = 'B' ) ).

    TRY.
        DATA(lv_out) = NEW zcl_xlwb_xlsx( )->render(
          iv_template = lv_tpl ir_context = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_sheet) = part_text( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `<mergeCell ref="A1:B1"/>` iv_msg = `gop ngang nhom A` ).
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'C' iv_row = 1 )
                                        exp = `B` msg = `nhom B ngay sau vung gop` ).
    cl_abap_unit_assert=>assert_equals( act = cell_of( iv_xlsx = lv_out iv_col = 'D' iv_row = 1 )
                                        exp = `END` msg = `o sau vung lap dich dung cot` ).
  ENDMETHOD.
ENDCLASS.
