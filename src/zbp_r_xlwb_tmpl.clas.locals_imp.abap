CLASS lhc_template DEFINITION INHERITING FROM cl_abap_behavior_handler.

  PRIVATE SECTION.
    CONSTANTS:
      c_engine_ssml TYPE c LENGTH 4 VALUE 'SSML',
      c_engine_xlsx TYPE c LENGTH 4 VALUE 'XLSX',
      c_engine_docx TYPE c LENGTH 4 VALUE 'DOCX',
      c_engine_asset TYPE c LENGTH 4 VALUE 'ASST',
      c_mime_ssml   TYPE string VALUE `application/vnd.ms-excel`,
      c_mime_xlsx   TYPE string VALUE `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.

    METHODS get_global_authorizations FOR GLOBAL AUTHORIZATION
      IMPORTING REQUEST requested_authorizations FOR Template RESULT result.

    METHODS setdefaults FOR DETERMINE ON MODIFY
      IMPORTING keys FOR Template~setDefaults.

    METHODS validateengine FOR VALIDATE ON SAVE
      IMPORTING keys FOR Template~validateEngine.

    METHODS validatetemplate FOR VALIDATE ON SAVE
      IMPORTING keys FOR Template~validateTemplate.

    METHODS validatesource FOR VALIDATE ON SAVE
      IMPORTING keys FOR Template~validateSource.

    METHODS validateplaceholders FOR VALIDATE ON SAVE
      IMPORTING keys FOR Template~validatePlaceholders.

    METHODS previewrender FOR MODIFY
      IMPORTING keys FOR ACTION Template~previewRender RESULT result.

    METHODS preparedocx FOR MODIFY
      IMPORTING keys FOR ACTION Template~prepareDocx RESULT result.

    METHODS exporttemplate FOR MODIFY
      IMPORTING keys FOR ACTION Template~exportTemplate RESULT result.

    METHODS importtemplate FOR MODIFY
      IMPORTING keys FOR ACTION Template~importTemplate.

ENDCLASS.


CLASS lhc_template IMPLEMENTATION.

  METHOD get_global_authorizations.
  ENDMETHOD.

  METHOD setdefaults.
    " Điền MIME type và tên file theo engine nếu người dùng để trống.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Engine MimeType FileName IsActive Template )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl ASSIGNING FIELD-SYMBOL(<ls>).
      DATA(ls_upd) = VALUE zr_xlwb_tmpl( FormName = <ls>-FormName ).
      DATA(lv_touch) = abap_false.

      IF <ls>-Engine IS INITIAL.
        <ls>-Engine = c_engine_ssml.
      ENDIF.

      IF <ls>-MimeType IS INITIAL.
        ls_upd-MimeType = COND #( WHEN <ls>-Engine = c_engine_xlsx THEN c_mime_xlsx
                                  WHEN <ls>-Engine = c_engine_docx THEN zcl_xlwb_docx=>c_mime_docx
                                  WHEN <ls>-Engine = c_engine_asset THEN `image/png`
                                  ELSE c_mime_ssml ).
        lv_touch = abap_true.
      ENDIF.

      IF <ls>-FileName IS INITIAL AND <ls>-FormName IS NOT INITIAL.
        ls_upd-FileName = |{ <ls>-FormName }.| &&
                          COND #( WHEN <ls>-Engine = c_engine_xlsx THEN `xlsx`
                                  WHEN <ls>-Engine = c_engine_docx THEN `docx`
                                  WHEN <ls>-Engine = c_engine_asset THEN `png`
                                  ELSE `xls` ).
        lv_touch = abap_true.
      ENDIF.

      IF <ls>-IsActive IS INITIAL.
        ls_upd-IsActive = abap_true.
        lv_touch = abap_true.
      ENDIF.

      IF lv_touch = abap_false.
        CONTINUE.
      ENDIF.

      MODIFY ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
        ENTITY Template
          UPDATE FIELDS ( MimeType FileName IsActive )
          WITH VALUE #( ( FormName = ls_upd-FormName
                          MimeType = ls_upd-MimeType
                          FileName = ls_upd-FileName
                          IsActive = ls_upd-IsActive ) )
        REPORTED DATA(lt_rep_ignored).
    ENDLOOP.
  ENDMETHOD.

  METHOD validateengine.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Engine )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl INTO DATA(ls).
      IF ls-Engine = c_engine_ssml OR ls-Engine = c_engine_xlsx
         OR ls-Engine = c_engine_docx OR ls-Engine = c_engine_asset.
        CONTINUE.
      ENDIF.

      APPEND VALUE #( %tky = ls-%tky ) TO failed-template.
      APPEND VALUE #( %tky = ls-%tky
                      %element-Engine = if_abap_behv=>mk-on
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = |Engine phải là SSML, XLSX, DOCX hoặc ASST (đang là '{ ls-Engine }')| ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.

  METHOD validatetemplate.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Engine Template )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl INTO DATA(ls).
      DATA(lv_err) = ``.

      IF ls-Template IS INITIAL.
        lv_err = |Chưa tải file template lên|.
      ELSEIF ls-Engine = c_engine_xlsx OR ls-Engine = c_engine_docx.
        " .xlsx/.docx là zip: 2 byte đầu phải là 'PK' (x'504B')
        DATA lv_magic TYPE x LENGTH 2.
        lv_magic = ls-Template.
        IF lv_magic <> CONV xstring( '504B' ).
          lv_err = |Engine { ls-Engine } cần file zip (.| &&
                   COND string( WHEN ls-Engine = c_engine_docx THEN `docx` ELSE `xlsx` ) &&
                   |). File tải lên không phải zip.|.
        ENDIF.
      ELSEIF ls-Engine = c_engine_asset.
        " ASSET (logo/ảnh): PNG magic 89 50 4E 47
        DATA lv_png TYPE x LENGTH 4.
        lv_png = ls-Template.
        IF lv_png <> CONV xstring( '89504E47' ).
          lv_err = |Engine ASST cần file ảnh PNG (magic 89 50 4E 47 không khớp).|.
        ENDIF.
      ELSEIF ls-Engine = c_engine_ssml.
        DATA(lv_head) = ``.
        TRY.
            " convert( ) chi nhan SOURCE (khong co offset/length) -> phai cat xstring truoc.
            " Chi can 400 byte dau de nhan dien the <Workbook>, khong convert ca file.
            DATA(lv_len) = COND i( WHEN xstrlen( ls-Template ) > 400
                                   THEN 400 ELSE xstrlen( ls-Template ) ).
            DATA lv_part TYPE xstring.
            lv_part = ls-Template+0(lv_len).
            lv_head = cl_abap_conv_codepage=>create_in( )->convert( source = lv_part ).
          CATCH cx_root.
            lv_head = ``.
        ENDTRY.
        IF lv_head NS `<Workbook`.
          lv_err = |Engine SSML cần file XML Spreadsheet 2003 (thiếu thẻ <Workbook>)|.
        ENDIF.
      ENDIF.

      IF lv_err IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND VALUE #( %tky = ls-%tky ) TO failed-template.
      APPEND VALUE #( %tky = ls-%tky
                      %element-Template = if_abap_behv=>mk-on
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.

  METHOD validatesource.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName SourceClass )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl INTO DATA(ls).
      IF ls-SourceClass IS INITIAL.
        CONTINUE.
      ENDIF.

      DATA(lv_err) = ``.
      DATA(lv_class) = to_upper( condense( CONV string( ls-SourceClass ) ) ).
      DATA lo_descr TYPE REF TO cl_abap_typedescr.

      cl_abap_typedescr=>describe_by_name(
        EXPORTING p_name         = lv_class
        RECEIVING p_descr_ref    = lo_descr
        EXCEPTIONS type_not_found = 1 OTHERS = 2 ).

      IF sy-subrc <> 0.
        lv_err = |Class '{ lv_class }' không tồn tại|.
      ELSEIF lo_descr->kind <> cl_abap_typedescr=>kind_class.
        lv_err = |'{ lv_class }' không phải class|.
      ELSE.
        DATA(lo_class) = CAST cl_abap_classdescr( lo_descr ).
        IF NOT line_exists( lo_class->interfaces[ name = 'ZIF_XLWB_SOURCE' ] ).
          lv_err = |Class '{ lv_class }' chưa implement interface ZIF_XLWB_SOURCE|.
        ENDIF.
      ENDIF.

      IF lv_err IS INITIAL.
        CONTINUE.
      ENDIF.

      APPEND VALUE #( %tky = ls-%tky ) TO failed-template.
      APPEND VALUE #( %tky = ls-%tky
                      %element-SourceClass = if_abap_behv=>mk-on
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.

  METHOD validateplaceholders.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Engine Template SampleJson )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl INTO DATA(ls).
      IF ls-Template IS INITIAL.
        CONTINUE.                               " validateTemplate đã lo case này
      ENDIF.
      IF ls-Engine = c_engine_asset.
        CONTINUE.                               " ASSET là ảnh, không có placeholder
      ENDIF.

      DATA(lt_msg) = zcl_xlwb_tmpl_util=>validate(
                       iv_template    = ls-Template
                       iv_engine      = CONV #( ls-Engine )
                       iv_sample_json = ls-SampleJson ).

      LOOP AT lt_msg INTO DATA(ls_msg).
        IF ls_msg-severity = 'E'.
          APPEND VALUE #( %tky = ls-%tky ) TO failed-template.
        ENDIF.
        APPEND VALUE #( %tky = ls-%tky
                        %element-Template = if_abap_behv=>mk-on
                        %msg = new_message_with_text(
                                 severity = COND #( WHEN ls_msg-severity = 'E'
                                                    THEN if_abap_behv_message=>severity-error
                                                    ELSE if_abap_behv_message=>severity-warning )
                                 text     = ls_msg-text ) )
          TO reported-template.
      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.

  METHOD previewrender.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Engine Template MimeType FileName SampleJson )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT keys INTO DATA(ls_key).
      READ TABLE lt_tmpl INTO DATA(ls_tmpl) WITH KEY %tky = ls_key-%tky.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_json) = COND string( WHEN ls_key-%param-PayloadJson IS NOT INITIAL
                                   THEN ls_key-%param-PayloadJson
                                   ELSE ls_tmpl-SampleJson ).
      DATA(lv_err) = ``.

      IF ls_tmpl-Template IS INITIAL.
        lv_err = |Template chưa có file — upload trước khi preview|.
      ELSEIF lv_json IS INITIAL.
        lv_err = |Chưa có payload: nhập JSON vào dialog hoặc điền Sample JSON trên template|.
      ENDIF.

      IF lv_err IS INITIAL.
        TRY.
            DATA(lr_ctx)  = zcl_xlwb_ctx_json=>parse( lv_json ).
            DATA(ls_file) = zcl_xlwb_runtime=>render_with_template(
                              iv_template = ls_tmpl-Template
                              iv_engine   = CONV #( ls_tmpl-Engine )
                              ir_context  = lr_ctx ).

            " FE không hiển thị result abstract entity -> ghi file vào cột
            " PreviewFile (largeObject), người dùng tải qua link trên dòng
            MODIFY ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
              ENTITY Template
                UPDATE FIELDS ( PreviewName PreviewMime PreviewFile )
                WITH VALUE #( ( %tky        = ls_key-%tky
                                PreviewName = |{ ls_tmpl-FormName }_preview.{ ls_file-extension }|
                                PreviewMime = ls_file-mime_type
                                PreviewFile = ls_file-content ) )
              REPORTED DATA(lt_rep_prv).

            READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
              ENTITY Template
                ALL FIELDS WITH VALUE #( ( %tky = ls_key-%tky ) )
              RESULT DATA(lt_self).
            READ TABLE lt_self INTO DATA(ls_self) INDEX 1.
            IF sy-subrc = 0.
              APPEND VALUE #( %tky = ls_key-%tky %param = ls_self ) TO result.
            ENDIF.

            APPEND VALUE #( %tky = ls_key-%tky
                            %msg = new_message_with_text(
                                     severity = if_abap_behv_message=>severity-success
                                     text     = |Preview OK — tải file ở cột "Preview / export file"| ) )
              TO reported-template.
            CONTINUE.
          CATCH zcx_xlwb INTO DATA(lx).
            lv_err = lx->get_text( ).
        ENDTRY.
      ENDIF.

      APPEND VALUE #( %tky = ls_key-%tky ) TO failed-template.
      APPEND VALUE #( %tky = ls_key-%tky
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.

  METHOD preparedocx.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Template )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT keys INTO DATA(ls_key).
      READ TABLE lt_tmpl INTO DATA(ls_tmpl) WITH KEY %tky = ls_key-%tky.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_spec) = CONV string( ls_key-%param-PayloadJson ).
      DATA(lv_err) = ``.

      IF ls_tmpl-Template IS INITIAL.
        lv_err = |Template chưa có file — upload file Word (.docx) layout thô trước|.
      ELSEIF lv_spec IS INITIAL.
        lv_err = |Chưa có spec — ví dụ: header=ZSTRUCT_HD; table=items:ZSTRUCT_IT|.
      ENDIF.

      IF lv_err IS INITIAL.
        TRY.
            NEW zcl_xlwb_docx( )->prepare_docx(
              EXPORTING iv_docx        = ls_tmpl-Template
                        iv_spec        = lv_spec
              IMPORTING ev_docx        = DATA(lv_docx)
                        ev_sample_json = DATA(lv_sample) ).

            " File đã nhúng part -> cột PreviewFile để tải; SampleJson điền luôn
            " theo đúng cấu trúc DDIC để validate/preview dùng lại được.
            MODIFY ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
              ENTITY Template
                UPDATE FIELDS ( PreviewName PreviewMime PreviewFile SampleJson )
                WITH VALUE #( ( %tky        = ls_key-%tky
                                PreviewName = |{ ls_tmpl-FormName }_mapped.docx|
                                PreviewMime = zcl_xlwb_docx=>c_mime_docx
                                PreviewFile = lv_docx
                                SampleJson  = lv_sample ) )
              REPORTED DATA(lt_rep_prep).

            READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
              ENTITY Template
                ALL FIELDS WITH VALUE #( ( %tky = ls_key-%tky ) )
              RESULT DATA(lt_self_prep).
            READ TABLE lt_self_prep INTO DATA(ls_self_prep) INDEX 1.
            IF sy-subrc = 0.
              APPEND VALUE #( %tky = ls_key-%tky %param = ls_self_prep ) TO result.
            ENDIF.

            APPEND VALUE #( %tky = ls_key-%tky
                            %msg = new_message_with_text(
                                     severity = if_abap_behv_message=>severity-success
                                     text     = |Đã nhúng XML part — tải file ở cột "Preview /| &&
                                                | export file", map control trong Word rồi upload lại| ) )
              TO reported-template.
            CONTINUE.
          CATCH zcx_xlwb INTO DATA(lx_prep).
            lv_err = lx_prep->get_text( ).
        ENDTRY.
      ENDIF.

      APPEND VALUE #( %tky = ls_key-%tky ) TO failed-template.
      APPEND VALUE #( %tky = ls_key-%tky
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-error
                               text     = lv_err ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.


  METHOD exporttemplate.
    READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
      ENTITY Template
        FIELDS ( FormName Descr Engine MimeType FileName SourceClass SampleJson Template )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT keys INTO DATA(ls_key).
      READ TABLE lt_tmpl INTO DATA(ls_tmpl) WITH KEY %tky = ls_key-%tky.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      IF ls_tmpl-Template IS INITIAL.
        APPEND VALUE #( %tky = ls_key-%tky ) TO failed-template.
        APPEND VALUE #( %tky = ls_key-%tky
                        %msg = new_message_with_text(
                                 severity = if_abap_behv_message=>severity-error
                                 text     = |Template chưa có file — không có gì để export| ) )
          TO reported-template.
        CONTINUE.
      ENDIF.

      DATA(lv_json) = zcl_xlwb_tmpl_util=>export_json(
        is_tmpl     = VALUE #( form_name    = ls_tmpl-FormName
                               descr        = ls_tmpl-Descr
                               engine       = ls_tmpl-Engine
                               mime_type    = ls_tmpl-MimeType
                               file_name    = ls_tmpl-FileName
                               source_class = ls_tmpl-SourceClass
                               sample_json  = ls_tmpl-SampleJson )
        iv_template = ls_tmpl-Template ).

      MODIFY ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
        ENTITY Template
          UPDATE FIELDS ( PreviewName PreviewMime PreviewFile )
          WITH VALUE #( ( %tky        = ls_key-%tky
                          PreviewName = |{ ls_tmpl-FormName }.xlwbtmpl.json|
                          PreviewMime = `application/json`
                          PreviewFile = cl_abap_conv_codepage=>create_out( )->convert( lv_json ) ) )
        REPORTED DATA(lt_rep_exp).

      READ ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
        ENTITY Template
          ALL FIELDS WITH VALUE #( ( %tky = ls_key-%tky ) )
        RESULT DATA(lt_self).
      READ TABLE lt_self INTO DATA(ls_self) INDEX 1.
      IF sy-subrc = 0.
        APPEND VALUE #( %tky = ls_key-%tky %param = ls_self ) TO result.
      ENDIF.

      APPEND VALUE #( %tky = ls_key-%tky
                      %msg = new_message_with_text(
                               severity = if_abap_behv_message=>severity-success
                               text     = |Export OK — tải file JSON ở cột "Preview / export file"| ) )
        TO reported-template.
    ENDLOOP.
  ENDMETHOD.

  METHOD importtemplate.
    LOOP AT keys INTO DATA(ls_key).
      TRY.
          zcl_xlwb_tmpl_util=>import_json(
            EXPORTING iv_json     = ls_key-%param-PayloadJson
            IMPORTING es_tmpl     = DATA(ls_imp)
                      ev_template = DATA(lv_template) ).
        CATCH zcx_xlwb INTO DATA(lx).
          APPEND VALUE #( %cid = ls_key-%cid ) TO failed-template.
          APPEND VALUE #( %cid = ls_key-%cid
                          %msg = new_message_with_text(
                                   severity = if_abap_behv_message=>severity-error
                                   text     = lx->get_text( ) ) )
            TO reported-template.
          CONTINUE.
      ENDTRY.

      MODIFY ENTITIES OF zr_xlwb_tmpl IN LOCAL MODE
        ENTITY Template
          CREATE FIELDS ( FormName Descr Engine MimeType FileName IsActive
                          SourceClass SampleJson Template )
          WITH VALUE #( ( %cid        = ls_key-%cid
                          FormName    = ls_imp-form_name
                          Descr       = ls_imp-descr
                          Engine      = ls_imp-engine
                          MimeType    = ls_imp-mime_type
                          FileName    = ls_imp-file_name
                          IsActive    = abap_true
                          SourceClass = ls_imp-source_class
                          SampleJson  = ls_imp-sample_json
                          Template    = lv_template ) )
        MAPPED DATA(ls_mapped)
        FAILED DATA(ls_failed)
        REPORTED DATA(ls_reported).

      INSERT LINES OF ls_mapped-template INTO TABLE mapped-template.
      INSERT LINES OF ls_failed-template INTO TABLE failed-template.
      INSERT LINES OF ls_reported-template INTO TABLE reported-template.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.

