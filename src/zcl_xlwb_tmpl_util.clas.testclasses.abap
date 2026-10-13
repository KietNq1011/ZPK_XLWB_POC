CLASS ltcl_tmpl_util DEFINITION FINAL FOR TESTING
  DURATION SHORT RISK LEVEL HARMLESS.

  PRIVATE SECTION.
    METHODS ssml
      IMPORTING iv_body        TYPE string
      RETURNING VALUE(rv_bin) TYPE xstring.

    METHODS scan_finds_tokens     FOR TESTING RAISING zcx_xlwb.
    METHODS valid_template_clean  FOR TESTING.
    METHODS pair_mismatch_error   FOR TESTING.
    METHODS unknown_path_warning  FOR TESTING.
    METHODS bad_sample_json_error FOR TESTING.
    METHODS typed_placeholder_warn FOR TESTING.
    METHODS export_import_round   FOR TESTING RAISING zcx_xlwb.
ENDCLASS.


CLASS ltcl_tmpl_util IMPLEMENTATION.

  METHOD ssml.
    rv_bin = cl_abap_conv_codepage=>create_out( )->convert(
      `<Workbook><Worksheet><Table>` && iv_body && `</Table></Worksheet></Workbook>` ).
  ENDMETHOD.

  METHOD scan_finds_tokens.
    DATA(lt_tokens) = zcl_xlwb_tmpl_util=>scan_placeholders(
      iv_template = ssml( `<Cell>{{header.customer}}</Cell><Cell>{{#items}}{{name}}{{/items}}</Cell>` )
      iv_engine   = `SSML` ).

    cl_abap_unit_assert=>assert_equals( act = lines( lt_tokens ) exp = 4 ).
    cl_abap_unit_assert=>assert_table_contains( line = `header.customer` table = lt_tokens ).
    cl_abap_unit_assert=>assert_table_contains( line = `#items` table = lt_tokens ).
  ENDMETHOD.

  METHOD valid_template_clean.
    DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
      iv_template    = ssml( `<Cell>{{customer}}</Cell><Cell>{{#items}}{{name}}{{/items}}</Cell>` &&
                             `<Cell>{{sum:items.amount}}</Cell>` )
      iv_engine      = `SSML`
      iv_sample_json = `{ "customer": "ACME", "items": [ { "name": "A", "amount": 5 } ] }` ).

    cl_abap_unit_assert=>assert_initial( act = lt_msg ).
  ENDMETHOD.

  METHOD pair_mismatch_error.
    DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
      iv_template = ssml( `<Cell>{{#items}}{{name}}{{/rows}}</Cell>` )
      iv_engine   = `SSML` ).

    cl_abap_unit_assert=>assert_not_initial( act = lt_msg ).
    READ TABLE lt_msg INTO DATA(ls_msg) INDEX 1.
    cl_abap_unit_assert=>assert_equals( act = ls_msg-severity exp = 'E' ).
  ENDMETHOD.

  METHOD unknown_path_warning.
    DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
      iv_template    = ssml( `<Cell>{{customerx}}</Cell>` )
      iv_engine      = `SSML`
      iv_sample_json = `{ "customer": "ACME" }` ).

    cl_abap_unit_assert=>assert_equals( act = lines( lt_msg ) exp = 1 ).
    READ TABLE lt_msg INTO DATA(ls_msg) INDEX 1.
    cl_abap_unit_assert=>assert_equals( act = ls_msg-severity exp = 'W' ).
  ENDMETHOD.

  METHOD bad_sample_json_error.
    DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
      iv_template    = ssml( `<Cell>{{customer}}</Cell>` )
      iv_engine      = `SSML`
      iv_sample_json = `{ hong phai json` ).

    cl_abap_unit_assert=>assert_not_initial( act = lt_msg ).
    READ TABLE lt_msg INTO DATA(ls_msg) INDEX 1.
    cl_abap_unit_assert=>assert_equals( act = ls_msg-severity exp = 'E' ).
  ENDMETHOD.

  METHOD typed_placeholder_warn.
    " o Type Number chua placeholder: Excel bao "file is corrupt" khi mo
    " template tho -> validate phai canh bao (bug that PACKING_LIST_HQ 27/08)
    DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
      iv_template = ssml( `<Cell><Data ss:Type="Number">{{qty}}</Data></Cell>` )
      iv_engine   = `SSML` ).

    cl_abap_unit_assert=>assert_equals( act = lines( lt_msg ) exp = 1 ).
    READ TABLE lt_msg INTO DATA(ls_msg) INDEX 1.
    cl_abap_unit_assert=>assert_equals( act = ls_msg-severity exp = 'W' ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( ls_msg-text CS `Number` )
      msg = `message phai neu ro loai o` ).
  ENDMETHOD.

  METHOD export_import_round.
    DATA(lv_template) = ssml( `<Cell>{{msg}}</Cell>` ).

    DATA(lv_json) = zcl_xlwb_tmpl_util=>export_json(
      is_tmpl = VALUE #( form_name    = 'XLWB_UT_ROUND'
                         descr        = 'Test "quote" \ backslash'
                         engine       = 'SSML'
                         mime_type    = 'application/vnd.ms-excel'
                         file_name    = 'round.xls'
                         source_class = 'ZCL_XLWB_DEMO_FILES'
                         sample_json  = `{ "msg": "hi" }` )
      iv_template = lv_template ).

    zcl_xlwb_tmpl_util=>import_json(
      EXPORTING iv_json     = lv_json
      IMPORTING es_tmpl     = DATA(ls_back)
                ev_template = DATA(lv_template_back) ).

    cl_abap_unit_assert=>assert_equals( act = ls_back-form_name    exp = 'XLWB_UT_ROUND' ).
    cl_abap_unit_assert=>assert_equals( act = ls_back-engine       exp = 'SSML' ).
    cl_abap_unit_assert=>assert_equals( act = ls_back-source_class exp = 'ZCL_XLWB_DEMO_FILES' ).
    cl_abap_unit_assert=>assert_equals( act = ls_back-sample_json  exp = `{ "msg": "hi" }` ).
    cl_abap_unit_assert=>assert_equals( act = CONV string( ls_back-descr )
                                        exp = `Test "quote" \ backslash` ).
    cl_abap_unit_assert=>assert_equals( act = lv_template_back exp = lv_template ).
  ENDMETHOD.

ENDCLASS.

