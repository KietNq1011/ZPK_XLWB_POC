CLASS ltcl_ctx_json DEFINITION FINAL FOR TESTING
  DURATION SHORT RISK LEVEL HARMLESS.

  PRIVATE SECTION.
    METHODS get_comp
      IMPORTING ir_data         TYPE REF TO data
                iv_path         TYPE string
      RETURNING VALUE(rv_value) TYPE string.

    METHODS scalar_types        FOR TESTING RAISING zcx_xlwb.
    METHODS date_detection      FOR TESTING RAISING zcx_xlwb.
    METHODS nested_object       FOR TESTING RAISING zcx_xlwb.
    METHODS array_of_objects    FOR TESTING RAISING zcx_xlwb.
    METHODS array_of_scalars    FOR TESTING RAISING zcx_xlwb.
    METHODS escapes             FOR TESTING RAISING zcx_xlwb.
    METHODS unicode_escape      FOR TESTING RAISING zcx_xlwb.
    METHODS empty_containers    FOR TESTING RAISING zcx_xlwb.
    METHODS invalid_json_raises FOR TESTING.
    METHODS number_is_decfloat  FOR TESTING RAISING zcx_xlwb.
ENDCLASS.


CLASS ltcl_ctx_json IMPLEMENTATION.

  METHOD get_comp.
    " đọc giá trị theo path 'A.B' từ cây data động (chỉ dùng cho assert)
    DATA lr_cur TYPE REF TO data.
    lr_cur = ir_data.
    SPLIT to_upper( iv_path ) AT '.' INTO TABLE DATA(lt_seg).

    FIELD-SYMBOLS <ls_any> TYPE any.
    LOOP AT lt_seg INTO DATA(lv_seg).
      ASSIGN lr_cur->* TO <ls_any>.
      ASSIGN COMPONENT lv_seg OF STRUCTURE <ls_any> TO FIELD-SYMBOL(<lv_comp>).
      cl_abap_unit_assert=>assert_subrc( act = sy-subrc msg = |component { lv_seg }| ).
      GET REFERENCE OF <lv_comp> INTO lr_cur.
    ENDLOOP.

    ASSIGN lr_cur->* TO FIELD-SYMBOL(<lv_val>).
    rv_value = |{ <lv_val> }|.
  ENDMETHOD.

  METHOD scalar_types.
    DATA(lr) = zcl_xlwb_ctx_json=>parse(
      `{ "message": "Hello", "qty": 5, "ok": true, "off": false, "nix": null }` ).

    cl_abap_unit_assert=>assert_equals( act = get_comp( ir_data = lr iv_path = `message` )
                                        exp = `Hello` ).
    cl_abap_unit_assert=>assert_equals( act = get_comp( ir_data = lr iv_path = `qty` )
                                        exp = `5` ).
    cl_abap_unit_assert=>assert_equals( act = get_comp( ir_data = lr iv_path = `ok` )
                                        exp = `X` ).
    cl_abap_unit_assert=>assert_initial( act = get_comp( ir_data = lr iv_path = `off` ) ).
    cl_abap_unit_assert=>assert_initial( act = get_comp( ir_data = lr iv_path = `nix` ) ).
  ENDMETHOD.

  METHOD date_detection.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( `{ "shipdate": "2026-08-24", "text": "2026-99-99" }` ).

    ASSIGN lr->* TO FIELD-SYMBOL(<ls>).
    ASSIGN COMPONENT 'SHIPDATE' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv_date>).
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data( <lv_date> ).
    cl_abap_unit_assert=>assert_equals( act = lo_type->type_kind
                                        exp = cl_abap_typedescr=>typekind_date ).
    cl_abap_unit_assert=>assert_equals( act = |{ <lv_date> }| exp = `20260824` ).

    " tháng 99 không phải ngày -> giữ nguyên string
    ASSIGN COMPONENT 'TEXT' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv_text>).
    lo_type = cl_abap_typedescr=>describe_by_data( <lv_text> ).
    cl_abap_unit_assert=>assert_equals( act = lo_type->type_kind
                                        exp = cl_abap_typedescr=>typekind_string ).
  ENDMETHOD.

  METHOD nested_object.
    DATA(lr) = zcl_xlwb_ctx_json=>parse(
      `{ "header": { "customer": "ACME", "addr": { "city": "Hanoi" } } }` ).

    cl_abap_unit_assert=>assert_equals(
      act = get_comp( ir_data = lr iv_path = `header.customer` ) exp = `ACME` ).
    cl_abap_unit_assert=>assert_equals(
      act = get_comp( ir_data = lr iv_path = `header.addr.city` ) exp = `Hanoi` ).
  ENDMETHOD.

  METHOD array_of_objects.
    DATA(lr) = zcl_xlwb_ctx_json=>parse(
      `{ "items": [ { "pos": 10, "matnr": "SLAB-A" }, { "pos": 20, "matnr": "SLAB-B" } ] }` ).

    ASSIGN lr->* TO FIELD-SYMBOL(<ls>).
    ASSIGN COMPONENT 'ITEMS' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lt_items>).
    FIELD-SYMBOLS <lt_tab> TYPE ANY TABLE.
    ASSIGN <lt_items> TO <lt_tab>.
    cl_abap_unit_assert=>assert_equals( act = lines( <lt_tab> ) exp = 2 ).

    DATA lv_seen TYPE string.
    LOOP AT <lt_tab> ASSIGNING FIELD-SYMBOL(<ls_row>).
      ASSIGN COMPONENT 'MATNR' OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<lv_m>).
      lv_seen = lv_seen && <lv_m> && `;`.
    ENDLOOP.
    cl_abap_unit_assert=>assert_equals( act = lv_seen exp = `SLAB-A;SLAB-B;` ).
  ENDMETHOD.

  METHOD array_of_scalars.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( `{ "tags": [ "a", "b", "c" ] }` ).

    ASSIGN lr->* TO FIELD-SYMBOL(<ls>).
    ASSIGN COMPONENT 'TAGS' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv_any>).
    FIELD-SYMBOLS <lt_tab> TYPE ANY TABLE.
    ASSIGN <lv_any> TO <lt_tab>.
    cl_abap_unit_assert=>assert_equals( act = lines( <lt_tab> ) exp = 3 ).
  ENDMETHOD.

  METHOD escapes.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( `{ "t": "a\"b\\c\/d" }` ).
    cl_abap_unit_assert=>assert_equals(
      act = get_comp( ir_data = lr iv_path = `t` ) exp = `a"b\c/d` ).
  ENDMETHOD.

  METHOD unicode_escape.
    " backtick literal ABAP: backslash là ký tự thường nên chuỗi dưới đây
    " tới parser đúng dạng JSON escape \uXXXX
    DATA(lv_json) = `{ "u": "` && '\' && `u0041` && '\' && `u00e9" }`.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( lv_json ).
    cl_abap_unit_assert=>assert_equals(
      act = get_comp( ir_data = lr iv_path = `u` ) exp = `Aé` ).
  ENDMETHOD.

  METHOD empty_containers.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( `{ "obj": {}, "arr": [] }` ).

    ASSIGN lr->* TO FIELD-SYMBOL(<ls>).
    ASSIGN COMPONENT 'ARR' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv_any>).
    FIELD-SYMBOLS <lt_tab> TYPE ANY TABLE.
    ASSIGN <lv_any> TO <lt_tab>.
    cl_abap_unit_assert=>assert_equals( act = lines( <lt_tab> ) exp = 0 ).
  ENDMETHOD.

  METHOD invalid_json_raises.
    TRY.
        zcl_xlwb_ctx_json=>parse( `{ "a": ` ).
        cl_abap_unit_assert=>fail( `JSON cụt phải raise zcx_xlwb` ).
      CATCH zcx_xlwb ##NO_HANDLER.
    ENDTRY.

    TRY.
        zcl_xlwb_ctx_json=>parse( `hello` ).
        cl_abap_unit_assert=>fail( `chuỗi trần phải raise zcx_xlwb` ).
      CATCH zcx_xlwb ##NO_HANDLER.
    ENDTRY.
  ENDMETHOD.

  METHOD number_is_decfloat.
    DATA(lr) = zcl_xlwb_ctx_json=>parse( `{ "amount": 1327.55, "neg": -3, "exp": 1.5e2 }` ).

    ASSIGN lr->* TO FIELD-SYMBOL(<ls>).
    ASSIGN COMPONENT 'AMOUNT' OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv_a>).
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data( <lv_a> ).
    cl_abap_unit_assert=>assert_equals( act = lo_type->type_kind
                                        exp = cl_abap_typedescr=>typekind_decfloat34 ).
    cl_abap_unit_assert=>assert_equals( act = CONV decfloat34( <lv_a> )
                                        exp = CONV decfloat34( '1327.55' ) ).
    cl_abap_unit_assert=>assert_equals(
      act = get_comp( ir_data = lr iv_path = `exp` ) exp = |{ CONV decfloat34( 150 ) }| ).
  ENDMETHOD.

ENDCLASS.

