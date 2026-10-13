CLASS lhc_part DEFINITION INHERITING FROM cl_abap_behavior_handler.

  PRIVATE SECTION.
    METHODS get_global_authorizations FOR GLOBAL AUTHORIZATION
      IMPORTING REQUEST requested_authorizations FOR Part RESULT result.

    METHODS setdefaults FOR DETERMINE ON MODIFY
      IMPORTING keys FOR Part~setDefaults.

    METHODS validateddic FOR VALIDATE ON SAVE
      IMPORTING keys FOR Part~validateDdic.

    METHODS generatexml FOR MODIFY
      IMPORTING keys FOR ACTION Part~generateXml RESULT result.

    METHODS detect_kind
      IMPORTING iv_ddic        TYPE string
      RETURNING VALUE(rv_kind) TYPE string.

ENDCLASS.


CLASS lhc_part IMPLEMENTATION.

  METHOD get_global_authorizations.
  ENDMETHOD.

  METHOD detect_kind.
    " '' = không tồn tại; STRUCTURE / TABLE TYPE / INVALID
    cl_abap_typedescr=>describe_by_name(
      EXPORTING p_name         = to_upper( condense( iv_ddic ) )
      RECEIVING p_descr_ref    = DATA(lo_type)
      EXCEPTIONS type_not_found = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      rv_kind = ``.
      RETURN.
    ENDIF.
    rv_kind = SWITCH #( lo_type->kind
                WHEN cl_abap_typedescr=>kind_struct THEN `STRUCTURE`
                WHEN cl_abap_typedescr=>kind_table  THEN `TABLE TYPE`
                ELSE `INVALID` ).
  ENDMETHOD.

  METHOD setdefaults.
    READ ENTITIES OF zr_xlwb_part IN LOCAL MODE
      ENTITY Part
        FIELDS ( PartName DdicName DdicKind RootName Namespace )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_part).

    LOOP AT lt_part ASSIGNING FIELD-SYMBOL(<ls>).
      DATA(ls_upd) = VALUE zr_xlwb_part( ).
      DATA(lv_touch) = abap_false.

      IF <ls>-RootName IS INITIAL.
        ls_upd-RootName = `data`.
        lv_touch = abap_true.
      ENDIF.

      IF <ls>-Namespace IS INITIAL.
        ls_upd-Namespace = `urn:zdocx:data`.
        lv_touch = abap_true.
      ENDIF.

      " tự nhận diện loại DDIC ngay khi nhập tên
      IF <ls>-DdicName IS NOT INITIAL.
        DATA(lv_kind) = detect_kind( CONV #( <ls>-DdicName ) ).
        IF lv_kind <> <ls>-DdicKind.
          ls_upd-DdicKind = lv_kind.
          lv_touch = abap_true.
        ENDIF.
      ENDIF.

      IF lv_touch = abap_false.
        CONTINUE.
      ENDIF.

      MODIFY ENTITIES OF zr_xlwb_part IN LOCAL MODE
        ENTITY Part
          UPDATE FIELDS ( DdicKind RootName Namespace )
          WITH VALUE #( ( %tky      = <ls>-%tky
                          DdicKind  = COND #( WHEN ls_upd-DdicKind IS NOT INITIAL
                                              THEN ls_upd-DdicKind ELSE <ls>-DdicKind )
                          RootName  = COND #( WHEN ls_upd-RootName IS NOT INITIAL
                                              THEN ls_upd-RootName ELSE <ls>-RootName )
                          Namespace = COND #( WHEN ls_upd-Namespace IS NOT INITIAL
                                              THEN ls_upd-Namespace ELSE <ls>-Namespace ) ) )
        REPORTED DATA(lt_rep_ignored).
    ENDLOOP.
  ENDMETHOD.

  METHOD validateddic.
    READ ENTITIES OF zr_xlwb_part IN LOCAL MODE
      ENTITY Part
        FIELDS ( PartName DdicName )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_part).

    LOOP AT lt_part INTO DATA(ls).
      DATA(lv_err) = ``.

      IF ls-DdicName IS INITIAL.
        lv_err = |Chưa nhập tên DDIC (structure / bảng DB / table type)|.
      ELSE.
        CASE detect_kind( CONV #( ls-DdicName ) ).
          WHEN ``.
            lv_err = |Không tìm thấy DDIC type '{ ls-DdicName }'|.
          WHEN `INVALID`.
            lv_err = |'{ ls-DdicName }' phải là structure, bảng DB hoặc table type|.
        ENDCASE.
      ENDIF.

      IF lv_err IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND VALUE #( %tky = ls-%tky ) TO failed-part.
      APPEND VALUE #( %tky = ls-%tky
                      %element-DdicName = if_abap_behv=>mk-on
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-part.
    ENDLOOP.
  ENDMETHOD.

  METHOD generatexml.
    READ ENTITIES OF zr_xlwb_part IN LOCAL MODE
      ENTITY Part
        FIELDS ( PartName DdicName RootName Namespace )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_part).

    LOOP AT keys INTO DATA(ls_key).
      READ TABLE lt_part INTO DATA(ls_part) WITH KEY %tky = ls_key-%tky.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_err) = ``.

      IF ls_part-DdicName IS INITIAL.
        lv_err = |Chưa nhập tên DDIC — nhập rồi bấm Generate lại|.
      ENDIF.

      IF lv_err IS INITIAL.
        TRY.
            NEW zcl_xlwb_docx( )->build_part_from_ddic(
              EXPORTING iv_ddic        = CONV #( ls_part-DdicName )
                        iv_root        = CONV #( ls_part-RootName )
                        iv_ns          = CONV #( ls_part-Namespace )
              IMPORTING ev_xml         = DATA(lv_xml)
                        ev_sample_json = DATA(lv_json)
                        ev_kind        = DATA(lv_kind) ).

            MODIFY ENTITIES OF zr_xlwb_part IN LOCAL MODE
              ENTITY Part
                UPDATE FIELDS ( DdicKind SampleJson PartMime PartFileName PartFile )
                WITH VALUE #( ( %tky         = ls_key-%tky
                                DdicKind     = lv_kind
                                SampleJson   = lv_json
                                PartMime     = `application/xml`
                                PartFileName = |{ to_lower( ls_part-PartName ) }_part.xml|
                                PartFile     = cl_abap_conv_codepage=>create_out( )->convert( lv_xml ) ) )
              REPORTED DATA(lt_rep_gen).

            READ ENTITIES OF zr_xlwb_part IN LOCAL MODE
              ENTITY Part
                ALL FIELDS WITH VALUE #( ( %tky = ls_key-%tky ) )
              RESULT DATA(lt_self).
            READ TABLE lt_self INTO DATA(ls_self) INDEX 1.
            IF sy-subrc = 0.
              APPEND VALUE #( %tky = ls_key-%tky %param = ls_self ) TO result.
            ENDIF.

            APPEND VALUE #( %tky = ls_key-%tky
                            %msg = new_message_with_text(
                                     severity = if_abap_behv_message=>severity-success
                                     text     = |Đã sinh XML part ({ lv_kind }) — tải file ở cột| &&
                                                | "XML Part file", add vào Word rồi map control| ) )
              TO reported-part.
            CONTINUE.
          CATCH zcx_xlwb INTO DATA(lx_gen).
            lv_err = lx_gen->get_text( ).
        ENDTRY.
      ENDIF.

      APPEND VALUE #( %tky = ls_key-%tky ) TO failed-part.
      APPEND VALUE #( %tky = ls_key-%tky
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-part.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.

