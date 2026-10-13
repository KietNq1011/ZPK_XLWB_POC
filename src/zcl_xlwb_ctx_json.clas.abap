"! <p class="shorttext synchronized" lang="en">XLWB: JSON payload -> context (REF TO DATA)</p>
"!
"! Chuyển payload JSON thành cây data động để đưa thẳng vào engine XLWB
"! (Preview trong app template — không cần structure ABAP khai sẵn).
"! Quy tắc kiểu: object → structure, array → standard table (row type theo
"! phần tử đầu), number → decfloat34, "YYYY-MM-DD" → TYPE d, true/false →
"! abap_bool, null → string rỗng. Engine nhận decfloat34 là Number và
"! TYPE d là DateTime nên number format của template có tác dụng.
CLASS zcl_xlwb_ctx_json DEFINITION
  PUBLIC FINAL CREATE PRIVATE.

  PUBLIC SECTION.
    "! Parse JSON thành REF TO DATA. JSON sai cú pháp → raise zcx_xlwb
    "! kèm vị trí lỗi.
    CLASS-METHODS parse
      IMPORTING iv_json        TYPE string
      RETURNING VALUE(rr_data) TYPE REF TO data
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_member,
             name TYPE string,
             ref  TYPE REF TO data,
           END OF ty_member,
           ty_members TYPE STANDARD TABLE OF ty_member WITH EMPTY KEY.

    DATA mv_json TYPE string.
    DATA mv_pos  TYPE i.
    DATA mv_len  TYPE i.

    METHODS parse_value
      RETURNING VALUE(rr_ref) TYPE REF TO data
      RAISING   zcx_xlwb.
    METHODS parse_object
      RETURNING VALUE(rr_ref) TYPE REF TO data
      RAISING   zcx_xlwb.
    METHODS parse_array
      RETURNING VALUE(rr_ref) TYPE REF TO data
      RAISING   zcx_xlwb.
    METHODS parse_string
      RETURNING VALUE(rv_text) TYPE string
      RAISING   zcx_xlwb.
    METHODS parse_number
      RETURNING VALUE(rr_ref) TYPE REF TO data
      RAISING   zcx_xlwb.
    METHODS expect_literal
      IMPORTING iv_literal TYPE string
      RAISING   zcx_xlwb.
    METHODS skip_ws.
    METHODS peek
      RETURNING VALUE(rv_char) TYPE string.
    METHODS fail
      IMPORTING iv_text TYPE string
      RAISING   zcx_xlwb.
    METHODS comp_name
      IMPORTING iv_key         TYPE string
      RETURNING VALUE(rv_name) TYPE string.
    METHODS build_struct
      IMPORTING it_members    TYPE ty_members
      RETURNING VALUE(rr_ref) TYPE REF TO data
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_CTX_JSON IMPLEMENTATION.


  METHOD parse.
    DATA(lo) = NEW zcl_xlwb_ctx_json( ).
    lo->mv_json = iv_json.
    lo->mv_len  = strlen( iv_json ).
    lo->mv_pos  = 0.

    lo->skip_ws( ).
    IF lo->mv_pos >= lo->mv_len.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |JSON rỗng — chưa có payload để preview| ).
    ENDIF.

    rr_data = lo->parse_value( ).

    lo->skip_ws( ).
    IF lo->mv_pos < lo->mv_len.
      lo->fail( |ký tự thừa sau khi JSON đã đóng| ).
    ENDIF.
  ENDMETHOD.


  METHOD parse_value.
    skip_ws( ).
    DATA(lv_c) = peek( ).

    CASE lv_c.
      WHEN '{'.
        rr_ref = parse_object( ).

      WHEN '['.
        rr_ref = parse_array( ).

      WHEN '"'.
        DATA(lv_str) = parse_string( ).
        " "YYYY-MM-DD" -> TYPE d để engine ghi kiểu DateTime
        IF strlen( lv_str ) = 10 AND lv_str+4(1) = '-' AND lv_str+7(1) = '-'
           AND lv_str(4) CO '0123456789' AND lv_str+5(2) CO '0123456789'
           AND lv_str+8(2) CO '0123456789'
           AND lv_str+5(2) BETWEEN '01' AND '12'
           AND lv_str+8(2) BETWEEN '01' AND '31'.
          DATA lv_date TYPE d.
          lv_date = lv_str(4) && lv_str+5(2) && lv_str+8(2).
          CREATE DATA rr_ref TYPE d.
          ASSIGN rr_ref->* TO FIELD-SYMBOL(<lv_d>).
          <lv_d> = lv_date.
        ELSE.
          CREATE DATA rr_ref TYPE string.
          ASSIGN rr_ref->* TO FIELD-SYMBOL(<lv_s>).
          <lv_s> = lv_str.
        ENDIF.

      WHEN 't'.
        expect_literal( `true` ).
        CREATE DATA rr_ref TYPE abap_bool.
        ASSIGN rr_ref->* TO FIELD-SYMBOL(<lv_t>).
        <lv_t> = abap_true.

      WHEN 'f'.
        expect_literal( `false` ).
        CREATE DATA rr_ref TYPE abap_bool.

      WHEN 'n'.
        expect_literal( `null` ).
        CREATE DATA rr_ref TYPE string.

      WHEN OTHERS.
        IF lv_c CO '-0123456789'.
          rr_ref = parse_number( ).
        ELSE.
          fail( |ký tự không hợp lệ '{ lv_c }'| ).
        ENDIF.
    ENDCASE.
  ENDMETHOD.


  METHOD parse_object.
    DATA lt_members TYPE ty_members.

    mv_pos = mv_pos + 1.                              " nuốt '{'
    skip_ws( ).

    IF peek( ) = '}'.
      mv_pos = mv_pos + 1.
    ELSE.
      DO.
        skip_ws( ).
        IF peek( ) <> '"'.
          fail( |thiếu tên thuộc tính (phải bắt đầu bằng ")| ).
        ENDIF.
        DATA(lv_key) = parse_string( ).
        skip_ws( ).
        IF peek( ) <> ':'.
          fail( |thiếu ':' sau tên '{ lv_key }'| ).
        ENDIF.
        mv_pos = mv_pos + 1.

        DATA(lv_name) = comp_name( lv_key ).
        IF NOT line_exists( lt_members[ name = lv_name ] ).
          APPEND VALUE #( name = lv_name ref = parse_value( ) ) TO lt_members.
        ELSE.
          parse_value( ).                             " key trùng sau sanitize -> giữ bản đầu
        ENDIF.

        skip_ws( ).
        CASE peek( ).
          WHEN ','.
            mv_pos = mv_pos + 1.
          WHEN '}'.
            mv_pos = mv_pos + 1.
            EXIT.
          WHEN OTHERS.
            fail( |thiếu ',' hoặc '\}' trong object| ).
        ENDCASE.
      ENDDO.
    ENDIF.

    rr_ref = build_struct( lt_members ).
  ENDMETHOD.


  METHOD build_struct.
    DATA lt_comp TYPE cl_abap_structdescr=>component_table.

    LOOP AT it_members INTO DATA(ls_m).
      APPEND VALUE #(
          name = ls_m-name
          type = CAST cl_abap_datadescr( cl_abap_typedescr=>describe_by_data_ref( ls_m-ref ) ) )
        TO lt_comp.
    ENDLOOP.

    IF lt_comp IS INITIAL.
      " object rỗng: structure vẫn phải có >=1 component
      APPEND VALUE #( name = 'DUMMY' type = cl_abap_elemdescr=>get_string( ) ) TO lt_comp.
    ENDIF.

    DATA(lo_struct) = cl_abap_structdescr=>create( lt_comp ).
    CREATE DATA rr_ref TYPE HANDLE lo_struct.
    ASSIGN rr_ref->* TO FIELD-SYMBOL(<ls_target>).

    LOOP AT it_members INTO ls_m.
      ASSIGN COMPONENT ls_m-name OF STRUCTURE <ls_target> TO FIELD-SYMBOL(<lv_comp>).
      ASSIGN ls_m-ref->* TO FIELD-SYMBOL(<lv_value>).
      <lv_comp> = <lv_value>.
    ENDLOOP.
  ENDMETHOD.


  METHOD parse_array.
    DATA lt_elems TYPE STANDARD TABLE OF REF TO data WITH EMPTY KEY.

    mv_pos = mv_pos + 1.                              " nuốt '['
    skip_ws( ).

    IF peek( ) = ']'.
      mv_pos = mv_pos + 1.
    ELSE.
      DO.
        APPEND parse_value( ) TO lt_elems.
        skip_ws( ).
        CASE peek( ).
          WHEN ','.
            mv_pos = mv_pos + 1.
            skip_ws( ).
          WHEN ']'.
            mv_pos = mv_pos + 1.
            EXIT.
          WHEN OTHERS.
            fail( |thiếu ',' hoặc ']' trong array| ).
        ENDCASE.
      ENDDO.
    ENDIF.

    " row type theo phần tử ĐẦU; phần tử sau khác cấu trúc -> bind theo tên,
    " component không khớp bị bỏ qua
    DATA lo_line TYPE REF TO cl_abap_datadescr.
    IF lt_elems IS INITIAL.
      lo_line = cl_abap_elemdescr=>get_string( ).
    ELSE.
      lo_line = CAST cl_abap_datadescr(
                  cl_abap_typedescr=>describe_by_data_ref( lt_elems[ 1 ] ) ).
    ENDIF.

    DATA(lo_tab) = cl_abap_tabledescr=>create( p_line_type = lo_line ).
    CREATE DATA rr_ref TYPE HANDLE lo_tab.
    ASSIGN rr_ref->* TO FIELD-SYMBOL(<lt_target>).

    DATA lr_row TYPE REF TO data.
    CREATE DATA lr_row TYPE HANDLE lo_line.
    ASSIGN lr_row->* TO FIELD-SYMBOL(<ls_row>).

    LOOP AT lt_elems INTO DATA(lr_elem).
      CLEAR <ls_row>.
      ASSIGN lr_elem->* TO FIELD-SYMBOL(<lv_elem>).

      IF lo_line->kind = cl_abap_typedescr=>kind_struct.
        DATA(lo_elem_type) = cl_abap_typedescr=>describe_by_data_ref( lr_elem ).
        IF lo_elem_type->kind = cl_abap_typedescr=>kind_struct.
          LOOP AT CAST cl_abap_structdescr( lo_elem_type )->components INTO DATA(ls_c).
            ASSIGN COMPONENT ls_c-name OF STRUCTURE <lv_elem> TO FIELD-SYMBOL(<lv_src>).
            ASSIGN COMPONENT ls_c-name OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<lv_dst>).
            IF sy-subrc = 0.
              TRY.
                  <lv_dst> = <lv_src>.
                CATCH cx_root ##CATCH_ALL.
                  " kiểu lệch giữa các dòng (vd null->string vs table) -> bỏ qua cột đó
              ENDTRY.
            ENDIF.
          ENDLOOP.
        ENDIF.
      ELSE.
        TRY.
            <ls_row> = <lv_elem>.
          CATCH cx_root ##CATCH_ALL.
            fail( |phần tử array không cùng kiểu với phần tử đầu| ).
        ENDTRY.
      ENDIF.

      INSERT <ls_row> INTO TABLE <lt_target>.
    ENDLOOP.
  ENDMETHOD.


  METHOD parse_string.
    mv_pos = mv_pos + 1.                              " nuốt '"'

    WHILE mv_pos < mv_len.
      DATA(lv_c) = substring( val = mv_json off = mv_pos len = 1 ).

      IF lv_c = '"'.
        mv_pos = mv_pos + 1.
        RETURN.
      ENDIF.

      IF lv_c <> '\'.
        rv_text = rv_text && lv_c.
        mv_pos = mv_pos + 1.
        CONTINUE.
      ENDIF.

      " escape sequence
      IF mv_pos + 1 >= mv_len.
        fail( |chuỗi kết thúc giữa escape| ).
      ENDIF.
      DATA(lv_e) = substring( val = mv_json off = mv_pos + 1 len = 1 ).
      mv_pos = mv_pos + 2.

      CASE lv_e.
        WHEN '"'.  rv_text = rv_text && '"'.
        WHEN '\'.  rv_text = rv_text && '\'.
        WHEN '/'.  rv_text = rv_text && '/'.
        WHEN 'n'.  rv_text = rv_text && cl_abap_char_utilities=>newline.
        WHEN 't'.  rv_text = rv_text && cl_abap_char_utilities=>horizontal_tab.
        WHEN 'r'.  rv_text = rv_text && cl_abap_char_utilities=>cr_lf(1).
        WHEN 'b' OR 'f'.
          rv_text = rv_text && ` `.
        WHEN 'u'.
          IF mv_pos + 4 > mv_len.
            fail( |\\u thiếu 4 ký tự hex| ).
          ENDIF.
          DATA(lv_hex) = to_upper( substring( val = mv_json off = mv_pos len = 4 ) ).
          mv_pos = mv_pos + 4.
          IF lv_hex CN '0123456789ABCDEF'.
            fail( |\\u{ lv_hex } không phải hex| ).
          ENDIF.
          DATA lv_x2 TYPE x LENGTH 2.
          lv_x2 = lv_hex.
          TRY.
              rv_text = rv_text &&
                cl_abap_conv_codepage=>create_in( codepage = `UTF-16BE`
                  )->convert( CONV xstring( lv_x2 ) ).
            CATCH cx_root ##CATCH_ALL.
              rv_text = rv_text && '?'.
          ENDTRY.
        WHEN OTHERS.
          fail( |escape '\\{ lv_e }' không hợp lệ| ).
      ENDCASE.
    ENDWHILE.

    fail( |chuỗi không được đóng bằng "| ).
  ENDMETHOD.


  METHOD parse_number.
    DATA(lv_start) = mv_pos.

    WHILE mv_pos < mv_len
      AND substring( val = mv_json off = mv_pos len = 1 ) CO '0123456789+-.eE'.
      mv_pos = mv_pos + 1.
    ENDWHILE.

    DATA(lv_token) = substring( val = mv_json off = lv_start len = mv_pos - lv_start ).

    DATA lv_num TYPE decfloat34.
    TRY.
        lv_num = lv_token.
      CATCH cx_sy_conversion_error.
        fail( |'{ lv_token }' không phải số| ).
    ENDTRY.

    CREATE DATA rr_ref TYPE decfloat34.
    ASSIGN rr_ref->* TO FIELD-SYMBOL(<lv_n>).
    <lv_n> = lv_num.
  ENDMETHOD.


  METHOD expect_literal.
    DATA(lv_l) = strlen( iv_literal ).
    IF mv_pos + lv_l > mv_len
    OR substring( val = mv_json off = mv_pos len = lv_l ) <> iv_literal.
      fail( |mong đợi '{ iv_literal }'| ).
    ENDIF.
    mv_pos = mv_pos + lv_l.
  ENDMETHOD.


  METHOD skip_ws.
    WHILE mv_pos < mv_len.
      DATA(lv_c) = substring( val = mv_json off = mv_pos len = 1 ).
      IF lv_c = ` ` OR lv_c = cl_abap_char_utilities=>horizontal_tab
      OR lv_c = cl_abap_char_utilities=>newline
      OR lv_c = cl_abap_char_utilities=>cr_lf(1).
        mv_pos = mv_pos + 1.
      ELSE.
        RETURN.
      ENDIF.
    ENDWHILE.
  ENDMETHOD.


  METHOD peek.
    IF mv_pos < mv_len.
      rv_char = substring( val = mv_json off = mv_pos len = 1 ).
    ENDIF.
  ENDMETHOD.


  METHOD fail.
    DATA(lv_from) = COND i( WHEN mv_pos > 30 THEN mv_pos - 30 ELSE 0 ).
    DATA(lv_ctx)  = substring( val = mv_json off = lv_from
                               len = COND #( WHEN mv_len - lv_from > 60
                                             THEN 60 ELSE mv_len - lv_from ) ).
    RAISE EXCEPTION NEW zcx_xlwb(
      iv_text = |JSON lỗi tại vị trí { mv_pos }: { iv_text } — gần "...{ lv_ctx }..."| ).
  ENDMETHOD.


  METHOD comp_name.
    DATA(lv_key) = to_upper( iv_key ).
    DATA(lv_len) = strlen( lv_key ).
    DATA(lv_i)   = 0.

    WHILE lv_i < lv_len AND strlen( rv_name ) < 30.
      DATA(lv_c) = substring( val = lv_key off = lv_i len = 1 ).
      rv_name = rv_name && COND string(
        WHEN lv_c CO 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_' THEN lv_c ELSE '_' ).
      lv_i = lv_i + 1.
    ENDWHILE.

    IF rv_name IS INITIAL.
      rv_name = 'X'.
    ELSEIF rv_name(1) CO '0123456789'.
      rv_name = 'X' && substring( val = rv_name off = 0
                                  len = COND #( WHEN strlen( rv_name ) > 29
                                                THEN 29 ELSE strlen( rv_name ) ) ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
