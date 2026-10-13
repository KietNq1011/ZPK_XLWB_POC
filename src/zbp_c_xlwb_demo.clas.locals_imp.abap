CLASS lhc_example DEFINITION INHERITING FROM cl_abap_behavior_handler.

  PRIVATE SECTION.
    METHODS get_global_authorizations FOR GLOBAL AUTHORIZATION
      IMPORTING REQUEST requested_authorizations FOR Example RESULT result.

    METHODS regeneratefile FOR MODIFY
      IMPORTING keys FOR ACTION Example~regenerateFile RESULT result.

    METHODS regenerateall FOR MODIFY
      IMPORTING keys FOR ACTION Example~regenerateAll.

    "! Render lại file mẫu của một example (template ƯU TIÊN bảng ZXLWB_TMPL)
    "! rồi ghi vào bản ghi catalog qua EML internal update.
    METHODS regen_one
      IMPORTING iv_example_id   TYPE string
      RETURNING VALUE(rv_error) TYPE string.

ENDCLASS.


CLASS lhc_example IMPLEMENTATION.

  METHOD get_global_authorizations.
  ENDMETHOD.

  METHOD regen_one.
    TRY.
        DATA(ls_file) = zcl_xlwb_demo_files=>render( iv_example_id ).
      CATCH zcx_xlwb INTO DATA(lx).
        rv_error = lx->get_text( ).
        RETURN.
    ENDTRY.

    MODIFY ENTITIES OF zc_xlwb_demo IN LOCAL MODE
      ENTITY Example
        UPDATE FIELDS ( EngineType FileName MimeType Content )
        WITH VALUE #( ( ExampleId  = iv_example_id
                        EngineType = ls_file-engine
                        FileName   = ls_file-file_name
                        MimeType   = ls_file-mime_type
                        Content    = ls_file-content ) )
      FAILED DATA(ls_failed)
      REPORTED DATA(ls_reported).

    IF ls_failed-example IS NOT INITIAL.
      rv_error = |Không cập nhật được bản ghi { iv_example_id }|.
    ENDIF.
  ENDMETHOD.

  METHOD regeneratefile.
    LOOP AT keys INTO DATA(ls_key).
      DATA(lv_error) = regen_one( CONV #( ls_key-ExampleId ) ).

      IF lv_error IS NOT INITIAL.
        APPEND VALUE #( %tky = ls_key-%tky ) TO failed-example.
        APPEND VALUE #( %tky = ls_key-%tky
                        %msg = new_message_with_text(
                                 severity = if_abap_behv_message=>severity-error
                                 text     = lv_error ) )
          TO reported-example.
        CONTINUE.
      ENDIF.

      READ ENTITIES OF zc_xlwb_demo IN LOCAL MODE
        ENTITY Example
          ALL FIELDS WITH VALUE #( ( %tky = ls_key-%tky ) )
        RESULT DATA(lt_row).

      READ TABLE lt_row INTO DATA(ls_row) INDEX 1.
      IF sy-subrc = 0.
        APPEND VALUE #( %tky = ls_key-%tky %param = ls_row ) TO result.
        APPEND VALUE #( %tky = ls_key-%tky
                        %msg = new_message_with_text(
                                 severity = if_abap_behv_message=>severity-success
                                 text     = |Đã gen lại { ls_row-FileName } theo template hiện hành| ) )
          TO reported-example.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD regenerateall.
    SELECT example_id FROM zxlwb_demo ORDER BY example_id INTO TABLE @DATA(lt_ids).

    DATA(lv_ok)  = 0.
    DATA(lv_err) = 0.
    DATA(lv_first_error) = ``.

    LOOP AT lt_ids INTO DATA(ls_id).
      DATA(lv_error) = regen_one( CONV #( ls_id-example_id ) ).
      IF lv_error IS INITIAL.
        lv_ok += 1.
      ELSE.
        lv_err += 1.
        IF lv_first_error IS INITIAL.
          lv_first_error = |{ ls_id-example_id }: { lv_error }|.
        ENDIF.
      ENDIF.
    ENDLOOP.

    DATA(lv_text) = COND string(
      WHEN lv_err = 0 THEN |Đã gen lại { lv_ok } file mẫu theo template hiện hành|
      ELSE |Gen lại { lv_ok } file, { lv_err } lỗi — lỗi đầu: { lv_first_error }| ).

    READ TABLE keys INTO DATA(ls_key) INDEX 1.
    APPEND VALUE #( %cid = ls_key-%cid
                    %msg = new_message_with_text(
                             severity = COND #( WHEN lv_err = 0
                                                THEN if_abap_behv_message=>severity-success
                                                ELSE if_abap_behv_message=>severity-warning )
                             text     = lv_text ) )
      TO reported-example.
  ENDMETHOD.

ENDCLASS.

