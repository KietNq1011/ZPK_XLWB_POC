"! <p class="shorttext synchronized" lang="en">XLWB: render form mẫu theo mã ví dụ</p>
"!
"! Một chỗ duy nhất biết "ví dụ EXnn dùng dữ liệu mẫu nào" — dùng chung cho
"! runner F9 ({@link zcl_xlwb_run}), app Fiori tra cứu ví dụ, và loader nạp
"! dữ liệu demo. Nhờ vậy dữ liệu mẫu không bị nhân bản ở nhiều nơi.
CLASS zcl_xlwb_demo_files DEFINITION
  PUBLIC FINAL CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_file,
             content   TYPE xstring,
             file_name TYPE string,
             mime_type TYPE string,
             engine    TYPE string,
           END OF ty_file.

    "! Render form mẫu của một ví dụ với dữ liệu mẫu sẵn có
    "! @parameter iv_example_id | EX00 EX01 EX02 EX02A EX03 EX04 EX05 EX08 EX11 EX13 EX14 EX16 EX17 EX20 EX21 EX22 — Word: DX00 DX01 DX02 DX05 DX20 DX21 — SSML logo: DS21
    CLASS-METHODS render
      IMPORTING iv_example_id  TYPE string
      RETURNING VALUE(rs_file) TYPE ty_file
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    "! Logo động từ asset LOGO_DEMO (ZXLWB_TMPL, ENGINE=ASST); rỗng nếu chưa có
    CLASS-METHODS get_logo_asset
      RETURNING VALUE(rv_b64) TYPE string.

    CONSTANTS:
      c_mime_ssml TYPE string VALUE `application/vnd.ms-excel`,
      c_mime_xlsx TYPE string VALUE `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
ENDCLASS.



CLASS ZCL_XLWB_DEMO_FILES IMPLEMENTATION.


  METHOD get_logo_asset.
    TRY.
        rv_b64 = zcl_xlwb_runtime=>get_asset_b64( `LOGO_DEMO` ).
      CATCH zcx_xlwb.
        CLEAR rv_b64.
    ENDTRY.
  ENDMETHOD.


  METHOD render.
    DATA(lv_id) = to_upper( condense( iv_example_id ) ).

    rs_file-engine    = `SSML`.
    rs_file-mime_type = c_mime_ssml.
    rs_file-file_name = |{ lv_id }.xls|.

    CASE lv_id.

      WHEN 'EX00'.
        rs_file-content = NEW zcl_xlwb_ex00_hello( )->get_file( ).

      WHEN 'EX01'.
        rs_file-content = NEW zcl_xlwb_ex01_label( )->get_file( VALUE #(
          to       = VALUE #( name = 'ACME Corp' street = '1 Main St' city = 'Chicago' )
          shipdate = cl_abap_context_info=>get_system_date( )
          weight   = '125.50' ) ).

      WHEN 'EX02'.
        rs_file-content = NEW zcl_xlwb_ex02_labels( )->get_file( VALUE #(
          ( name = 'ACME Corp'    city = 'Chicago' )
          ( name = 'Globex Ltd'   city = 'London' )
          ( name = 'Initech GmbH' city = 'Berlin' ) ) ).

      WHEN 'EX02A'.
        rs_file-content = NEW zcl_xlwb_ex02a_hlabels( )->get_file( VALUE #(
          ( name = 'ACME Corp'    city = 'Chicago' )
          ( name = 'Globex Ltd'   city = 'London' )
          ( name = 'Initech GmbH' city = 'Berlin' ) ) ).

      WHEN 'EX04'.
        rs_file-content = NEW zcl_xlwb_ex04_bundle( )->get_file( VALUE #(
          orderno  = 'PO-2026-0815'
          customer = 'ACME Corp'
          items    = VALUE #( ( matnr = 'SLAB-WHITE-3CM' qty = 20 amount = '2100.00' )
                              ( matnr = 'SLAB-GREY-2CM'  qty = 15 amount = '1327.50' ) )
          labels   = VALUE #( ( name = 'ACME Corp'  city = 'Chicago' )
                              ( name = 'Globex Ltd' city = 'London' ) ) ) ).

      WHEN 'EX03'.
        rs_file-content = NEW zcl_xlwb_ex03_sheets( )->get_file( VALUE #(
          ( name = 'ACME Corp'  city = 'Chicago' )
          ( name = 'Globex Ltd' city = 'London' ) ) ).

      WHEN 'EX05'.
        rs_file-content = NEW zcl_xlwb_ex05_order( )->get_file( VALUE #(
          orderno  = 'PO-2026-0815'
          customer = 'ACME Corp'
          orddate  = cl_abap_context_info=>get_system_date( )
          remark   = 'Rush order'
          items    = VALUE #(
            ( pos = 10 matnr = 'SLAB-WHITE-3CM' qty = 20 price = '105.00' amount = '2100.00' )
            ( pos = 20 matnr = 'SLAB-GREY-2CM'  qty = 15 price = '88.50'  amount = '1327.50' )
            ( pos = 30 matnr = 'SLAB-BLACK-3CM' qty = 5  price = '120.00' amount = '600.00' ) ) ) ).

      WHEN 'EX08G'.
        rs_file-content = NEW zcl_xlwb_ex08g_grouping( )->get_file( VALUE #(
          ( position = 'Manager' emps = VALUE #(
              ( name = 'Nguyen Van A' role = 'Sales manager' )
              ( name = 'Tran Thi B'   role = 'Plant manager' ) ) )
          ( position = 'Clerk' emps = VALUE #(
              ( name = 'Le Van C'     role = 'Warehouse clerk' ) ) ) ) ).

      WHEN 'EX09'.
        " kích thước NGẪU NHIÊN mỗi lần gen — đúng tinh thần 4.09
        rs_file-content = NEW zcl_xlwb_ex09_dyntable( )->get_file( ).

      WHEN 'EX08'.
        rs_file-content = NEW zcl_xlwb_ex08_dyncols( )->get_file( VALUE #(
          cols = VALUE #( ( title = 'Jan' ) ( title = 'Feb' ) ( title = 'Mar' ) ( title = 'Apr' ) )
          rows = VALUE #(
            ( rowname = 'White quartz' cells = VALUE #( ( val = 10 ) ( val = 20 ) ( val = 30 ) ( val = 40 ) ) )
            ( rowname = 'Grey quartz'  cells = VALUE #( ( val = 5 )  ( val = 6 )  ( val = 7 )  ( val = 8 ) ) ) ) ) ).

      WHEN 'EX11'.
        rs_file-content = NEW zcl_xlwb_ex11_multilevel( )->get_file( VALUE #(
          ( cityfrom = 'HANOI' cityto = 'SAIGON' conns = VALUE #(
              ( connid = 'VN213' carrier = 'VN' flights = VALUE #(
                  ( fldate = '20260801' capacity = 180 occupied = 150 )
                  ( fldate = '20260802' capacity = 180 occupied = 170 )
                  ( fldate = '20260803' capacity = 180 occupied = 160 ) ) )
              ( connid = 'VJ145' carrier = 'VJ' flights = VALUE #(
                  ( fldate = '20260801' capacity = 220 occupied = 200 ) ) ) ) )
          ( cityfrom = 'SAIGON' cityto = 'DANANG' conns = VALUE #(
              ( connid = 'VN311' carrier = 'VN' flights = VALUE #(
                  ( fldate = '20260805' capacity = 150 occupied = 120 ) ) ) ) ) ) ).

      WHEN 'EX13'.
        rs_file-content = NEW zcl_xlwb_ex13_gantt( )->get_file( VALUE #(
          monthname = 'August 2026'
          days  = VALUE #( ( label = '1' ) ( label = '2' ) ( label = '3' ) ( label = '4' ) ( label = '5' ) )
          weeks = VALUE #( ( label = 'W31' days = VALUE #( ( label = '1' ) ( label = '2' ) ( label = '3' ) ) )
                           ( label = 'W32' days = VALUE #( ( label = '4' ) ( label = '5' ) ) ) )
          tasks = VALUE #(
            ( phase = 'Design' task = 'Blueprint' duration = 2
              cells = VALUE #( ( mark = 'X' ) ( mark = 'X' ) ( mark = '' ) ( mark = '' ) ( mark = '' ) ) )
            ( phase = 'Build'  task = 'Cutting'   duration = 3
              cells = VALUE #( ( mark = '' ) ( mark = 'X' ) ( mark = 'X' ) ( mark = 'X' ) ( mark = '' ) ) ) ) ) ).

      WHEN 'EX14'.
        rs_file-content = NEW zcl_xlwb_ex14_tree( )->get_file( VALUE #(
          ( node_key = 'G1' text = 'Raw slabs' )
          ( node_key = 'I1' parent_key = 'G1' text = 'White quartz 3cm' qty = 120 )
          ( node_key = 'I2' parent_key = 'G1' text = 'Grey quartz 2cm'  qty = 80 )
          ( node_key = 'G2' text = 'Finished goods' )
          ( node_key = 'I3' parent_key = 'G2' text = 'Countertop A'     qty = 35 ) ) ).

      WHEN 'EX17'.
        rs_file-content = NEW zcl_xlwb_ex17_chart( )->get_file( VALUE #(
          ( name = 'Jan' revenue = '1200.00' )
          ( name = 'Feb' revenue = '1550.50' )
          ( name = 'Mar' revenue = '980.00' )
          ( name = 'Apr' revenue = '2010.25' ) ) ).

      WHEN 'EX16'.
        rs_file-content = NEW zcl_xlwb_ex16_advanced( )->get_file(
          iv_taxcode = '0101234567' iv_show_secret = abap_false ).

      WHEN 'EX20'.
        rs_file-engine    = `XLSX`.
        rs_file-mime_type = c_mime_xlsx.
        rs_file-file_name = |EX20.xlsx|.
        rs_file-content   = NEW zcl_xlwb_ex20_logo( )->get_file( VALUE #(
          logo     = COND #( LET lv_a = get_logo_asset( ) IN
                                WHEN lv_a IS NOT INITIAL THEN lv_a
                                ELSE zcl_xlwb_ex20_logo=>c_demo_logo )
          invno    = 'INV-2026-0042'
          customer = 'ACME Corp'
          items    = VALUE #( ( matnr = 'SLAB-WHITE' amount = '2100.00' )
                              ( matnr = 'SLAB-GREY'  amount = '1327.50' ) ) ) ).

      WHEN 'EX21'.
        " ƯU TIÊN template trong bảng (form PACKING_LIST_HQ) — upload bản
        " .xlsx với Engine = XLSX là file mẫu tự thành .xlsx thật.
        DATA(ls_ctx) = zcl_xlwb_ex21_packlist=>sample_context( ).
        DATA(ls_pl)  = zcl_xlwb_runtime=>render_prefer_table(
          iv_form_name         = `PACKING_LIST_HQ`
          iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert(
                                   zcl_xlwb_ex21_packlist=>get_template( ) )
          iv_fallback_engine   = `SSML`
          ir_context           = REF #( ls_ctx ) ).
        rs_file-engine    = ls_pl-engine.
        rs_file-mime_type = ls_pl-mime_type.
        rs_file-file_name = |PackingListHQ.{ ls_pl-extension }|.
        rs_file-content   = ls_pl-content.

      WHEN 'EX22'.
        rs_file-file_name = |EX22.xls|.
        rs_file-content   = NEW zcl_xlwb_src_order( )->get_file( ).

      WHEN 'EX23'.
        rs_file-file_name = |EX23.xls|.
        rs_file-content   = NEW zcl_xlwb_src_bcnctp( )->get_file( ).

      WHEN 'DX00' OR 'DX01' OR 'DX02' OR 'DX05' OR 'DX20'.
        " Ví dụ Word (DOCX): template + Sample JSON nằm trong bảng ZXLWB_TMPL —
        " render đúng đường production zcl_xlwb_runtime=>render (ENGINE = DOCX).
        DATA(lv_form) = SWITCH string( lv_id
                          WHEN 'DX00' THEN `DOCX_EX00_HELLO`
                          WHEN 'DX01' THEN `DOCX_EX01_LETTER`
                          WHEN 'DX02' THEN `DOCX_EX02_ITEMS`
                          WHEN 'DX20' THEN `DOCX_EX20_LABEL`
                          ELSE `DOCX_EX05_ORDER` ).
        SELECT SINGLE sample_json FROM zxlwb_tmpl
          WHERE form_name = @lv_form
          INTO @DATA(lv_sample_json).
        IF sy-subrc <> 0 OR lv_sample_json IS INITIAL.
          RAISE EXCEPTION NEW zcx_xlwb(
            iv_text = |Chưa có template '{ lv_form }' kèm Sample JSON trong ZXLWB_TMPL| ).
        ENDIF.
        DATA(ls_docx) = zcl_xlwb_runtime=>render(
                          iv_form_name = lv_form
                          ir_context   = zcl_xlwb_ctx_json=>parse( lv_sample_json ) ).
        rs_file-engine    = ls_docx-engine.
        rs_file-mime_type = ls_docx-mime_type.
        rs_file-file_name = ls_docx-file_name.
        rs_file-content   = ls_docx-content.

      WHEN 'DX21'.
        " Word letterhead: LOGO ĐỘNG select từ ZXLWB_TMPL (asset ENGINE=ASST) lúc in
        TYPES: BEGIN OF lty_company,
                 name     TYPE string,
                 address  TYPE string,
                 tax_code TYPE string,
               END OF lty_company,
               BEGIN OF lty_logo_ctx,
                 company  TYPE lty_company,
                 doc_no   TYPE string,
                 doc_date TYPE string,
                 recv     TYPE string,
                 body     TYPE string,
                 logo     TYPE string,
               END OF lty_logo_ctx.
        DATA(ls_lg) = VALUE lty_logo_ctx(
          company  = VALUE #( name = `CÔNG TY TNHH ABC` address = `KCN Đồng Văn, Hà Nam`
                              tax_code = `0700123456` )
          doc_no   = `BG-2026/0901`
          doc_date = `28.08.2026`
          recv     = `Công ty CP XYZ`
          body     = `Kính gửi Quý công ty, Chúng tôi trân trọng gửi báo giá đá thạch anh Q3/2026.` ).
        TRY.
            ls_lg-logo = zcl_xlwb_runtime=>get_asset_b64( `LOGO_DEMO` ).
          CATCH zcx_xlwb.
            ls_lg-logo = zcl_xlwb_ex20_logo=>c_demo_logo.   " fallback 1x1
        ENDTRY.
        DATA(ls_word_lg) = zcl_xlwb_runtime=>render(
                             iv_form_name = `DOCX_EX21_LOGO`
                             ir_context   = REF #( ls_lg ) ).
        rs_file-engine    = ls_word_lg-engine.
        rs_file-mime_type = ls_word_lg-mime_type.
        rs_file-file_name = ls_word_lg-file_name.
        rs_file-content   = ls_word_lg-content.

      WHEN 'DS21'.
        " SSML + marker ảnh: engine SpreadsheetML 2003 KHÔNG nhúng được ảnh.
        " Ví dụ này để ĐỐI CHIẾU với EX20 (XLSX có logo) và DX21 (Word có logo).
        DATA: BEGIN OF ls_ss,
                company  TYPE string,
                slogan   TYPE string,
                doc_no   TYPE string,
                customer TYPE string,
                logo     TYPE string,
              END OF ls_ss.
        ls_ss-company  = `CÔNG TY TNHH ABC`.
        ls_ss-slogan   = `Demo company`.
        ls_ss-doc_no   = `BG-2026/0901`.
        ls_ss-customer = `Công ty CP XYZ`.
        ls_ss-logo     = get_logo_asset( ).
        DATA(ls_dss) = zcl_xlwb_runtime=>render( iv_form_name = `XLWB_EX21_SSML`
                                                 ir_context   = REF #( ls_ss ) ).
        rs_file-engine    = ls_dss-engine.
        rs_file-mime_type = ls_dss-mime_type.
        rs_file-file_name = ls_dss-file_name.
        rs_file-content   = ls_dss-content.

      WHEN OTHERS.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |Mã ví dụ không hợp lệ: '{ iv_example_id }'| ).
    ENDCASE.




    " nội dung là zip (PK..) -> .xlsx thật (runtime đã convert): chỉnh tên + mime
    DATA lv_magic TYPE x LENGTH 2.
    IF xstrlen( rs_file-content ) >= 2.
      lv_magic = rs_file-content.
      IF lv_magic = CONV xstring( '504B' ) AND rs_file-engine <> `DOCX`.
        rs_file-mime_type = c_mime_xlsx.
        REPLACE FIRST OCCURRENCE OF PCRE '(?i)\.xls$' IN rs_file-file_name WITH `.xlsx`.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
