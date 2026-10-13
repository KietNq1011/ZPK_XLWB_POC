"! <p class="shorttext synchronized" lang="en">XLWB Cloud: context binding (dùng chung 2 engine)</p>
"!
"! Giữ context stack (frame của từng vòng lặp) và phân giải placeholder
"! thành giá trị. Cả engine SpreadsheetML và engine XLSX dùng class này,
"! nên cú pháp placeholder luôn giống nhau giữa hai định dạng.
CLASS zcl_xlwb_ctx DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_value,
             text       TYPE string,
             is_number  TYPE abap_bool,
             is_date    TYPE abap_bool,
             is_time    TYPE abap_bool,
             date_value TYPE d,
             time_value TYPE t,
           END OF ty_value.

    METHODS constructor
      IMPORTING ir_context TYPE REF TO data
      RAISING   zcx_xlwb.

    "! Kind: 'ROW' (vong lap dong) / 'CELL' (vong lap o ngang) / 'SHEET'.
    "! ir_table = bang dang lap - can cho group_span( ) de so sanh cac dong lan can.
    METHODS push
      IMPORTING ir_data  TYPE REF TO data
                iv_index TYPE i DEFAULT 0
                iv_count TYPE i DEFAULT 0
                iv_kind  TYPE string DEFAULT ''
                ir_table TYPE REF TO data OPTIONAL.
    METHODS pop.

    "! Gop o theo gia tri: trong vong lap iv_kind, xet cac dong/o LIEN TIEP co
    "! cung gia tri cua iv_path. Tra ve dong hien tai co phai dong dau nhom
    "! (ev_is_first) va nhom dai bao nhieu phan tu (ev_span >= 1).
    METHODS group_span
      IMPORTING iv_path     TYPE string
                iv_kind     TYPE string DEFAULT 'ROW'
      EXPORTING ev_is_first TYPE abap_bool
                ev_span     TYPE i
      RAISING   zcx_xlwb.

    "! So o cua marker dang "=N": N la so nguyen hoac path tro tren context
    METHODS span_of
      IMPORTING iv_spec       TYPE string
      RETURNING VALUE(rv_num) TYPE i
      RAISING   zcx_xlwb.

    "! Giá trị của 1 token (không gồm dấu ngoặc kép nhọn)
    METHODS token_value
      IMPORTING iv_token        TYPE string
      RETURNING VALUE(rs_value) TYPE ty_value
      RAISING   zcx_xlwb.

    "! Thay mọi {{...}} trong 1 đoạn text
    METHODS render_text
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_text) TYPE string
      RAISING   zcx_xlwb.

    METHODS resolve
      IMPORTING iv_path        TYPE string
                iv_from        TYPE string DEFAULT ''
      RETURNING VALUE(rr_data) TYPE REF TO data
      RAISING   zcx_xlwb.

    METHODS resolve_table
      IMPORTING iv_path        TYPE string
                iv_from        TYPE string DEFAULT ''
      RETURNING VALUE(rr_data) TYPE REF TO data
      RAISING   zcx_xlwb.

    METHODS is_truthy
      IMPORTING iv_path          TYPE string
      RETURNING VALUE(rv_truthy) TYPE abap_bool
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_frame,
             data  TYPE REF TO data,
             index TYPE i,
             count TYPE i,
             kind  TYPE string,
             table TYPE REF TO data,
           END OF ty_frame,
           ty_frames TYPE STANDARD TABLE OF ty_frame WITH EMPTY KEY.

    DATA mt_frames TYPE ty_frames.

    METHODS aggregate
      IMPORTING iv_func         TYPE string
                iv_spec         TYPE string
      RETURNING VALUE(rs_value) TYPE ty_value
      RAISING   zcx_xlwb.
    METHODS value_of
      IMPORTING ir_data         TYPE REF TO data
      RETURNING VALUE(rs_value) TYPE ty_value.
    "! Doc iv_path tren MOT dong bat ky (khong phai dong hien tai)
    METHODS value_in_row
      IMPORTING ir_row        TYPE REF TO data
                iv_path       TYPE string
      RETURNING VALUE(rv_key) TYPE string
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_CTX IMPLEMENTATION.


  METHOD constructor.
    IF ir_context IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Context reference is initial' ).
    ENDIF.
    APPEND VALUE ty_frame( data = ir_context ) TO mt_frames.
  ENDMETHOD.


  METHOD push.
    APPEND VALUE ty_frame( data  = ir_data
                           index = iv_index
                           count = iv_count
                           kind  = iv_kind
                           table = ir_table ) TO mt_frames.
  ENDMETHOD.


  METHOD pop.
    IF lines( mt_frames ) > 1.
      DELETE mt_frames INDEX lines( mt_frames ).
    ENDIF.
  ENDMETHOD.


  METHOD token_value.
    DATA(lv_token) = condense( iv_token ).
    IF lv_token IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |token_value called with empty token| ).
    ENDIF.

    IF substring( val = lv_token off = 0 len = 1 ) CA '#/?^*'.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Structural marker \{\{{ lv_token }\}\} at unexpected position | &&
                  |(missing opener, or block marker inside a processed area)| ).
    ENDIF.

    IF lv_token = '@index' OR lv_token = '@count'.
      DATA(lv_frame_idx) = lines( mt_frames ).
      WHILE lv_frame_idx > 0 AND mt_frames[ lv_frame_idx ]-count = 0.
        lv_frame_idx -= 1.
      ENDWHILE.
      IF lv_frame_idx = 0.
        RAISE EXCEPTION NEW zcx_xlwb( iv_text = |\{\{{ lv_token }\}\} used outside of a loop| ).
      ENDIF.
      rs_value-is_number = abap_true.
      rs_value-text = COND #( WHEN lv_token = '@index'
                              THEN |{ mt_frames[ lv_frame_idx ]-index }|
                              ELSE |{ mt_frames[ lv_frame_idx ]-count }| ).
      RETURN.
    ENDIF.

    DATA lv_func TYPE string.
    DATA lv_spec TYPE string.
    FIND PCRE '(?-x)^(sum|avg|min|max|cnt):(.+)$' IN lv_token SUBMATCHES lv_func lv_spec.
    IF sy-subrc = 0.
      rs_value = aggregate( iv_func = lv_func iv_spec = lv_spec ).
      RETURN.
    ENDIF.

    rs_value = value_of( resolve( iv_path = lv_token iv_from = 'value' ) ).
  ENDMETHOD.


  METHOD render_text.
    DATA(lv_rest) = iv_text.
    WHILE lv_rest IS NOT INITIAL.
      DATA lv_off TYPE i.
      DATA lv_len TYPE i.
      DATA lv_token TYPE string.
      FIND PCRE '(?-x)\{\{([^}]+)\}\}' IN lv_rest
        MATCH OFFSET lv_off MATCH LENGTH lv_len SUBMATCHES lv_token.
      IF sy-subrc <> 0.
        rv_text = rv_text && lv_rest.
        RETURN.
      ENDIF.
      rv_text = rv_text && substring( val = lv_rest off = 0 len = lv_off )
             && token_value( lv_token )-text.
      lv_rest = substring( val = lv_rest off = lv_off + lv_len ).
    ENDWHILE.
  ENDMETHOD.


  METHOD value_of.
    FIELD-SYMBOLS <lv_val> TYPE any.
    ASSIGN ir_data->* TO <lv_val>.
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data( <lv_val> ).

    CASE lo_type->type_kind.
      WHEN cl_abap_typedescr=>typekind_int  OR cl_abap_typedescr=>typekind_int1
        OR cl_abap_typedescr=>typekind_int2 OR cl_abap_typedescr=>typekind_int8
        OR cl_abap_typedescr=>typekind_packed OR cl_abap_typedescr=>typekind_float
        OR cl_abap_typedescr=>typekind_decfloat16 OR cl_abap_typedescr=>typekind_decfloat34.
        rs_value-is_number = abap_true.
        DATA lv_number TYPE decfloat34.
        lv_number = <lv_val>.
        rs_value-text = |{ lv_number NUMBER = RAW }|.
      WHEN cl_abap_typedescr=>typekind_date.
        rs_value-is_date    = abap_true.
        rs_value-date_value = <lv_val>.
        rs_value-text = COND #( WHEN rs_value-date_value IS INITIAL THEN ``
                                ELSE |{ rs_value-date_value DATE = ISO }| ).
      WHEN cl_abap_typedescr=>typekind_time.
        rs_value-is_time    = abap_true.
        rs_value-time_value = <lv_val>.
        rs_value-text = |{ rs_value-time_value TIME = ISO }|.
      WHEN OTHERS.
        rs_value-text = |{ <lv_val> }|.
    ENDCASE.
  ENDMETHOD.


  METHOD value_in_row.
    " Day mot frame tam cho dong can doc -> resolve( ) tim frame trong cung
    " truoc nen se lay dung dong nay, roi pop ngay.
    push( ir_data = ir_row ).
    TRY.
        rv_key = value_of( resolve( iv_path = iv_path iv_from = 'mergesame' ) )-text.
      CLEANUP.
        pop( ).
    ENDTRY.
    pop( ).
  ENDMETHOD.


  METHOD group_span.
    ev_is_first = abap_true.
    ev_span     = 1.

    " frame trong cung co dung kind VA co bang nguon
    DATA(lv_f) = lines( mt_frames ).
    WHILE lv_f > 0.
      IF mt_frames[ lv_f ]-kind = iv_kind AND mt_frames[ lv_f ]-table IS BOUND.
        EXIT.
      ENDIF.
      lv_f -= 1.
    ENDWHILE.
    IF lv_f = 0.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Marker gop o theo gia tri ('{ iv_path }') chi dung duoc trong | &&
                  |vong lap { iv_kind }| ).
    ENDIF.

    DATA(ls_frame) = mt_frames[ lv_f ].
    FIELD-SYMBOLS <lt_src> TYPE INDEX TABLE.
    ASSIGN ls_frame-table->* TO <lt_src>.

    DATA(lv_cur) = value_in_row( ir_row = ls_frame-data iv_path = iv_path ).

    " dong dau nhom? -> so voi dong ngay truoc
    IF ls_frame-index > 1.
      READ TABLE <lt_src> INDEX ls_frame-index - 1 ASSIGNING FIELD-SYMBOL(<ls_prev>).
      IF sy-subrc = 0
      AND value_in_row( ir_row = REF #( <ls_prev> ) iv_path = iv_path ) = lv_cur.
        ev_is_first = abap_false.
      ENDIF.
    ENDIF.

    " do dai nhom: dem xuoi tu dong hien tai
    DATA(lv_n) = ls_frame-index + 1.
    WHILE lv_n <= lines( <lt_src> ).
      READ TABLE <lt_src> INDEX lv_n ASSIGNING FIELD-SYMBOL(<ls_next>).
      IF sy-subrc <> 0
      OR value_in_row( ir_row = REF #( <ls_next> ) iv_path = iv_path ) <> lv_cur.
        EXIT.
      ENDIF.
      ev_span += 1.
      lv_n    += 1.
    ENDWHILE.
  ENDMETHOD.


  METHOD span_of.
    DATA(lv_spec) = condense( iv_spec ).

    " so nguyen viet truc tiep
    IF lv_spec CO '0123456789'.
      rv_num = CONV i( lv_spec ).
      RETURN.
    ENDIF.

    " nguoc lai: path tren context, phai ra so
    DATA(ls_val) = token_value( lv_spec ).
    IF ls_val-is_number = abap_false.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Span cua merge phai la so: '{ iv_spec }' cho ra '{ ls_val-text }'| ).
    ENDIF.
    rv_num = CONV i( CONV decfloat34( ls_val-text ) ).
  ENDMETHOD.


  METHOD resolve.
    SPLIT to_upper( iv_path ) AT '.' INTO TABLE DATA(lt_segs).
    DELETE lt_segs WHERE table_line IS INITIAL.
    IF lt_segs IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |Empty placeholder path (raw='{ iv_path }', from={ iv_from })| ).
    ENDIF.

    DATA(lv_frame) = lines( mt_frames ).
    WHILE lv_frame > 0.
      DATA(lr_cur) = mt_frames[ lv_frame ]-data.
      DATA(lv_ok) = abap_true.

      LOOP AT lt_segs INTO DATA(lv_seg).
        FIELD-SYMBOLS <ls_any> TYPE any.
        ASSIGN lr_cur->* TO <ls_any>.
        IF cl_abap_typedescr=>describe_by_data( <ls_any> )->kind <> cl_abap_typedescr=>kind_struct.
          lv_ok = abap_false.
          EXIT.
        ENDIF.
        ASSIGN COMPONENT lv_seg OF STRUCTURE <ls_any> TO FIELD-SYMBOL(<lv_comp>).
        IF sy-subrc <> 0.
          lv_ok = abap_false.
          EXIT.
        ENDIF.
        lr_cur = REF #( <lv_comp> ).
      ENDLOOP.

      IF lv_ok = abap_true.
        rr_data = lr_cur.
        RETURN.
      ENDIF.
      lv_frame -= 1.
    ENDWHILE.

    RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Placeholder path '{ iv_path }' not found in context| ).
  ENDMETHOD.


  METHOD resolve_table.
    rr_data = resolve( iv_path = iv_path iv_from = iv_from ).
    IF cl_abap_typedescr=>describe_by_data_ref( rr_data )->kind <> cl_abap_typedescr=>kind_table.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |'{ iv_path }' is not an internal table| ).
    ENDIF.
  ENDMETHOD.


  METHOD is_truthy.
    DATA(lr_data) = resolve( iv_path = iv_path iv_from = 'truthy' ).
    FIELD-SYMBOLS <lv_any> TYPE any.
    ASSIGN lr_data->* TO <lv_any>.
    IF cl_abap_typedescr=>describe_by_data( <lv_any> )->kind = cl_abap_typedescr=>kind_table.
      FIELD-SYMBOLS <lt_any> TYPE ANY TABLE.
      ASSIGN lr_data->* TO <lt_any>.
      rv_truthy = xsdbool( lines( <lt_any> ) > 0 ).
    ELSE.
      rv_truthy = xsdbool( <lv_any> IS NOT INITIAL ).
    ENDIF.
  ENDMETHOD.


  METHOD aggregate.
    DATA lv_tab_path TYPE string.
    DATA lv_comp     TYPE string.

    IF iv_func = 'cnt'.
      lv_tab_path = iv_spec.
    ELSE.
      DATA(lv_dot) = find( val = iv_spec sub = '.' occ = -1 ).
      IF lv_dot < 0.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |\{\{{ iv_func }:{ iv_spec }\}\}: component missing (expected table.component)| ).
      ENDIF.
      lv_tab_path = substring( val = iv_spec off = 0 len = lv_dot ).
      lv_comp     = to_upper( substring( val = iv_spec off = lv_dot + 1 ) ).
    ENDIF.

    FIELD-SYMBOLS <lt_tab> TYPE INDEX TABLE.
    DATA(lr_agg_tab) = resolve_table( iv_path = lv_tab_path iv_from = 'aggregate' ).
    ASSIGN lr_agg_tab->* TO <lt_tab>.

    rs_value-is_number = abap_true.
    IF iv_func = 'cnt'.
      rs_value-text = |{ lines( <lt_tab> ) }|.
      RETURN.
    ENDIF.

    DATA lv_sum TYPE decfloat34.
    DATA lv_min TYPE decfloat34.
    DATA lv_max TYPE decfloat34.
    DATA(lv_first) = abap_true.

    LOOP AT <lt_tab> ASSIGNING FIELD-SYMBOL(<ls_row>).
      ASSIGN COMPONENT lv_comp OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<lv_num>).
      IF sy-subrc <> 0.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |\{\{{ iv_func }:{ iv_spec }\}\}: component '{ lv_comp }' not found| ).
      ENDIF.
      DATA(lv_val) = CONV decfloat34( <lv_num> ).
      lv_sum += lv_val.
      IF lv_first = abap_true.
        lv_min = lv_val.
        lv_max = lv_val.
        lv_first = abap_false.
      ELSE.
        lv_min = nmin( val1 = lv_min val2 = lv_val ).
        lv_max = nmax( val1 = lv_max val2 = lv_val ).
      ENDIF.
    ENDLOOP.

    CASE iv_func.
      WHEN 'sum'. rs_value-text = |{ lv_sum NUMBER = RAW }|.
      WHEN 'min'. rs_value-text = |{ lv_min NUMBER = RAW }|.
      WHEN 'max'. rs_value-text = |{ lv_max NUMBER = RAW }|.
      WHEN 'avg'.
        DATA(lv_cnt) = lines( <lt_tab> ).
        rs_value-text = COND #( WHEN lv_cnt = 0 THEN `0`
                                ELSE |{ CONV decfloat34( lv_sum / lv_cnt ) NUMBER = RAW }| ).
    ENDCASE.
  ENDMETHOD.
ENDCLASS.
