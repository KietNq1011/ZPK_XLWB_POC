CLASS zcl_xlwb_zip_probe DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
  PRIVATE SECTION.
    METHODS zip_roundtrip FOR TESTING.
    METHODS base64_to_xstring FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_ZIP_PROBE IMPLEMENTATION.


  METHOD zip_roundtrip.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    DATA(lv_x) = cl_abap_conv_codepage=>create_out( )->convert( `<x>hello</x>` ).
    lo_zip->add( name = 'dir/a.xml' content = lv_x ).
    DATA(lv_arc) = lo_zip->save( ).
    cl_abap_unit_assert=>assert_not_initial( act = lv_arc msg = `zip save` ).

    DATA(lo_in) = NEW cl_abap_zip( ).
    lo_in->load( lv_arc ).
    lo_in->get( EXPORTING name = 'dir/a.xml' IMPORTING content = DATA(lv_back) ).
    cl_abap_unit_assert=>assert_equals(
      act = cl_abap_conv_codepage=>create_in( )->convert( lv_back )
      exp = `<x>hello</x>`
      msg = `zip load/get` ).
    cl_abap_unit_assert=>assert_equals( act = lines( lo_in->files ) exp = 1 msg = `files table` ).
  ENDMETHOD.


  METHOD base64_to_xstring.
    " 1x1 transparent PNG, pattern lay tu zcl_qm_export_logo
    DATA(lv_png) = xco_cp=>string(
      `iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==`
      )->as_xstring( xco_cp_binary=>text_encoding->base64 )->value.
    cl_abap_unit_assert=>assert_not_initial( act = lv_png msg = `base64 decode` ).
    DATA lv_magic TYPE x LENGTH 4.
    lv_magic = lv_png.
    cl_abap_unit_assert=>assert_equals(
      act = |{ lv_magic }| exp = `89504E47` msg = `PNG magic bytes` ).
  ENDMETHOD.
ENDCLASS.
