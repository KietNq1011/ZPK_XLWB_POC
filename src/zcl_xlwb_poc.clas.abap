CLASS zcl_xlwb_poc DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
  PRIVATE SECTION.
    TYPES: BEGIN OF ty_item,
             pos    TYPE i,
             matnr  TYPE c LENGTH 18,
             amount TYPE p LENGTH 8 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY.

    METHODS col IMPORTING iv_col        TYPE string
                RETURNING VALUE(ro_col) TYPE REF TO cl_xco_xlsx_coordinate.
    METHODS row IMPORTING iv_row        TYPE i
                RETURNING VALUE(ro_row) TYPE REF TO cl_xco_xlsx_coordinate.
    METHODS build_template RETURNING VALUE(rv_xlsx) TYPE xstring.

    METHODS roundtrip_template FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_POC IMPLEMENTATION.


  METHOD col.
    ro_col = xco_cp_xlsx=>coordinate->for_alphabetic_value( iv_col ).
  ENDMETHOD.


  METHOD row.
    ro_row = xco_cp_xlsx=>coordinate->for_numeric_value( iv_row ).
  ENDMETHOD.


  METHOD build_template.
    " "Template": bold title + merge A1:C1 + label B2 (simulates XCO-compatible template)
    DATA(lo_wa) = xco_cp_xlsx=>document->empty( )->write_access( ).
    DATA(lo_ws) = lo_wa->get_workbook( )->worksheet->at_position( 1 ).
    lo_ws->set_name( `FORM` ).

    DATA(lo_cell_a1) = lo_ws->cursor( io_column = col( `A` )
                                      io_row    = row( 1 ) )->get_cell( ).
    lo_cell_a1->value->write_from( `CASLA COMMERCIAL INVOICE` ).
    lo_cell_a1->apply_styles( VALUE #(
      ( xco_cp_xlsx=>style->font( )->set_bold( ) ) ) ).

    lo_ws->merge_cells( xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
        )->from_column( col( `A` ) )->from_row( row( 1 )
        )->to_column( col( `C` ) )->to_row( row( 1 ) )->get_pattern( ) ).

    lo_ws->cursor( io_column = col( `B` ) io_row = row( 2 )
      )->get_cell( )->value->write_from( `Proforma:` ).

    rv_xlsx = lo_wa->get_file_content( ).
  ENDMETHOD.


  METHOD roundtrip_template.
    DATA(lv_template) = build_template( ).
    cl_abap_unit_assert=>assert_not_initial( act = lv_template msg = `template empty` ).

    " ===== ROUND-TRIP: load existing template, write data into it =====
    DATA(lo_doc) = xco_cp_xlsx=>document->for_file_content( lv_template ).
    DATA(lo_wa)  = lo_doc->write_access( ).
    DATA(lo_ws)  = lo_wa->get_workbook( )->worksheet->at_position( 1 ).

    " single value (placeholder C2)
    lo_ws->cursor( io_column = col( `C` ) io_row = row( 2 )
      )->get_cell( )->value->write_from( `PI-2026-001` ).

    " item table A5:C7 via row stream
    DATA(lt_items) = VALUE ty_items(
      ( pos = 10 matnr = 'SLAB-001' amount = '1050.50' )
      ( pos = 20 matnr = 'SLAB-002' amount = '2200.00' )
      ( pos = 30 matnr = 'SLAB-003' amount = '315.25' ) ).
    lo_ws->select( xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
        )->from_column( col( `A` ) )->from_row( row( 5 )
        )->to_column( col( `C` ) )->to_row( row( 7 ) )->get_pattern( )
      )->row_stream( )->operation->write_from( REF #( lt_items ) )->execute( ).

    " dynamic sheet
    lo_wa->get_workbook( )->add_new_sheet( `DA` ).

    DATA(lv_out) = lo_wa->get_file_content( ).
    cl_abap_unit_assert=>assert_not_initial( act = lv_out msg = `output empty` ).

    " ===== READ BACK: everything the template had must survive =====
    DATA(lo_ra)  = xco_cp_xlsx=>document->for_file_content( lv_out )->read_access( ).
    DATA(lo_rws) = lo_ra->get_workbook( )->worksheet->at_position( 1 ).

    DATA lv_a1 TYPE string.
    lo_rws->cursor( io_column = col( `A` ) io_row = row( 1 )
      )->get_cell( )->get_value( )->write_to( REF #( lv_a1 ) ).
    cl_abap_unit_assert=>assert_equals( act = lv_a1 exp = `CASLA COMMERCIAL INVOICE`
      msg = `template title lost after round-trip` ).

    cl_abap_unit_assert=>assert_true(
      act = lo_rws->cursor( io_column = col( `A` ) io_row = row( 1 )
              )->get_cell( )->has_style( )
      msg = `A1 style lost after round-trip` ).

    cl_abap_unit_assert=>assert_true( act = lo_rws->has_merged_range( )
      msg = `merge A1:C1 lost after round-trip` ).

    DATA lv_c2 TYPE string.
    lo_rws->cursor( io_column = col( `C` ) io_row = row( 2 )
      )->get_cell( )->get_value( )->write_to( REF #( lv_c2 ) ).
    cl_abap_unit_assert=>assert_equals( act = lv_c2 exp = `PI-2026-001`
      msg = `single value write failed` ).

    DATA lv_b7 TYPE string.
    lo_rws->cursor( io_column = col( `B` ) io_row = row( 7 )
      )->get_cell( )->get_value( )->write_to( REF #( lv_b7 ) ).
    cl_abap_unit_assert=>assert_equals( act = lv_b7 exp = `SLAB-003`
      msg = `bulk table write failed` ).

    cl_abap_unit_assert=>assert_true(
      act = lo_ra->get_workbook( )->worksheet->for_name( `DA` )->exists( )
      msg = `dynamic sheet DA missing` ).
  ENDMETHOD.
ENDCLASS.
