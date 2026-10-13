"! <p class="shorttext synchronized" lang="en">XLWB: test cho toàn bộ form mẫu</p>
"!
"! Các class form mẫu là class THƯỜNG (production code gọi được), nên test
"! của chúng gom về đây thay vì nằm trong từng class.
CLASS zcl_xlwb_examples_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    CLASS-METHODS as_text
      IMPORTING iv_file        TYPE xstring
      RETURNING VALUE(rv_text) TYPE string.
    METHODS assert_has
      IMPORTING iv_text TYPE string
                iv_sub  TYPE string
                iv_msg  TYPE string.

    METHODS ex00_hello       FOR TESTING.
    METHODS ex01_label       FOR TESTING.
    METHODS ex02_labels      FOR TESTING.
    METHODS ex03_sheets      FOR TESTING.
    METHODS ex05_order       FOR TESTING.
    METHODS ex08_dyncols     FOR TESTING.
    METHODS ex11_multilevel  FOR TESTING.
    METHODS ex13_gantt       FOR TESTING.
    METHODS ex16_advanced    FOR TESTING.
    METHODS ex20_logo        FOR TESTING.
    METHODS ex21_packlist    FOR TESTING.
    METHODS ex02a_hlabels    FOR TESTING.
    METHODS ex09_dyntable    FOR TESTING.
    METHODS ex08g_grouping   FOR TESTING.
    METHODS ex04_bundle      FOR TESTING.
    METHODS ex14_tree        FOR TESTING.
    METHODS ex17_chart       FOR TESTING.
    METHODS ex22_source      FOR TESTING.
    METHODS ex23_bcnctp      FOR TESTING.
    METHODS all_demo_files   FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_EXAMPLES_TEST IMPLEMENTATION.


  METHOD as_text.
    " output runtime giờ là .xlsx (zip): nối text mọi part XML để assert;
    " input không phải zip (SSML thô / part xml lẻ) thì convert utf-8 như cũ
    DATA lv_magic TYPE x LENGTH 2.
    IF xstrlen( iv_file ) >= 2.
      lv_magic = iv_file.
    ENDIF.
    IF lv_magic <> CONV xstring( '504B' ).
      rv_text = cl_abap_conv_codepage=>create_in( )->convert( iv_file ).
      RETURN.
    ENDIF.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_file ).
    LOOP AT lo_zip->files INTO DATA(ls_f).
      lo_zip->get( EXPORTING name = ls_f-name
                   IMPORTING content = DATA(lv_raw)
                   EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
      IF sy-subrc = 0 AND lv_raw IS NOT INITIAL.
        rv_text &&= cl_abap_conv_codepage=>create_in( )->convert( lv_raw ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD assert_has.
    cl_abap_unit_assert=>assert_true( act = xsdbool( iv_text CS iv_sub ) msg = iv_msg ).
  ENDMETHOD.


  METHOD ex00_hello.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex00_hello( )->get_file( ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Hello world!` iv_msg = `EX00 gia tri don` ).
  ENDMETHOD.


  METHOD ex01_label.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex01_label( )->get_file( VALUE #(
          to       = VALUE #( name = 'ACME Corp' street = '1 Main St' city = 'Chicago' )
          shipdate = '20260821'
          weight   = '125.50' ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `ACME Corp`               iv_msg = `EX01 path long nhau` ).
    assert_has( iv_text = lv_out iv_sub = `1 Main St, Chicago`      iv_msg = `EX01 text tron` ).
    " ngày ra serial Excel trong .xlsx (epoch 1899-12-30)
    DATA(lv_serial) = CONV d( '20260821' ) - CONV d( '18991230' ).
    assert_has( iv_text = lv_out iv_sub = |>{ lv_serial }<| iv_msg = `EX01 o kieu ngay -> serial` ).
  ENDMETHOD.


  METHOD ex02_labels.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex02_labels( )->get_file( VALUE #(
          ( name = 'ACME Corp'    city = 'Chicago' )
          ( name = 'Globex Ltd'   city = 'London' )
          ( name = 'Initech GmbH' city = 'Berlin' ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Label 1 / 3`  iv_msg = `EX02 @index/@count` ).
    assert_has( iv_text = lv_out iv_sub = `Initech GmbH` iv_msg = `EX02 label cuoi` ).
    assert_has( iv_text = lv_out iv_sub = `owBreaks`     iv_msg = `EX02 page break (SSML/xlsx)` ).
    assert_has( iv_text = lv_out iv_sub = `andscape`     iv_msg = `EX02 page setup giu nguyen` ).
  ENDMETHOD.


  METHOD ex03_sheets.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex03_sheets( )->get_file( VALUE #(
          ( name = 'ACME Corp'  city = 'Chicago' )
          ( name = 'Globex Ltd' city = 'London' ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Label 1`        iv_msg = `EX03 sheet 1` ).
    assert_has( iv_text = lv_out iv_sub = `Label 2`        iv_msg = `EX03 sheet 2` ).
    assert_has( iv_text = lv_out iv_sub = `TO: Globex Ltd` iv_msg = `EX03 du lieu tung sheet` ).
  ENDMETHOD.


  METHOD ex05_order.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex05_order( )->get_file( VALUE #(
          orderno  = 'PO-2026-0815'
          customer = 'ACME Corp'
          orddate  = '20260821'
          remark   = 'Rush order'
          items    = VALUE #(
            ( pos = 10 matnr = 'SLAB-WHITE-3CM' qty = 20 price = '105.00' amount = '2100.00' )
            ( pos = 20 matnr = 'SLAB-GREY-2CM'  qty = 15 price = '88.50'  amount = '1327.50' )
            ( pos = 30 matnr = 'SLAB-BLACK-3CM' qty = 5  price = '120.00' amount = '600.00' ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `PURCHASE ORDER PO-2026-0815` iv_msg = `EX05 header` ).
    assert_has( iv_text = lv_out iv_sub = `SLAB-GREY-2CM`   iv_msg = `EX05 dong item` ).
    assert_has( iv_text = lv_out iv_sub = `Total (3 items)` iv_msg = `EX05 cnt` ).
    assert_has( iv_text = lv_out iv_sub = `SUM(E$4:E6)` iv_msg = `EX05 cong thuc R1C1 -> A1` ).
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( lv_out CS `>4027.5` )
      msg = `EX05 sum:items.amount` ).
  ENDMETHOD.


  METHOD ex08_dyncols.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex08_dyncols( )->get_file( VALUE #(
          cols = VALUE #( ( title = 'Jan' ) ( title = 'Feb' ) ( title = 'Mar' ) ( title = 'Apr' ) )
          rows = VALUE #(
            ( rowname = 'White quartz' cells = VALUE #( ( val = 10 ) ( val = 20 ) ( val = 30 ) ( val = 40 ) ) )
            ( rowname = 'Grey quartz'  cells = VALUE #( ( val = 5 )  ( val = 6 )  ( val = 7 )  ( val = 8 ) ) ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Jan`          iv_msg = `EX08 cot dau` ).
    assert_has( iv_text = lv_out iv_sub = `Apr`          iv_msg = `EX08 cot cuoi (cell loop)` ).
    assert_has( iv_text = lv_out iv_sub = `Grey quartz`  iv_msg = `EX08 row loop` ).
  ENDMETHOD.


  METHOD ex11_multilevel.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex11_multilevel( )->get_file( VALUE #(
          ( cityfrom = 'HANOI' cityto = 'SAIGON' conns = VALUE #(
              ( connid = 'VN213' carrier = 'VN' flights = VALUE #(
                  ( fldate = '20260801' capacity = 180 occupied = 150 )
                  ( fldate = '20260802' capacity = 180 occupied = 170 )
                  ( fldate = '20260803' capacity = 180 occupied = 160 ) ) ) ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `HANOI - SAIGON`             iv_msg = `EX11 cap 1` ).
    assert_has( iv_text = lv_out iv_sub = `Subtotal VN213 (3 flights)` iv_msg = `EX11 subtotal theo cha` ).
    assert_has( iv_text = lv_out iv_sub = `>540<`                      iv_msg = `EX11 sum capacity` ).
    assert_has( iv_text = lv_out iv_sub = `mergeCell`                  iv_msg = `EX11 gop o dong` ).
  ENDMETHOD.


  METHOD ex13_gantt.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex13_gantt( )->get_file( VALUE #(
          monthname = 'August 2026'
          days  = VALUE #( ( label = '1' ) ( label = '2' ) ( label = '3' ) ( label = '4' ) ( label = '5' ) )
          weeks = VALUE #(
            ( label = 'W31' days = VALUE #( ( label = '1' ) ( label = '2' ) ( label = '3' ) ) )
            ( label = 'W32' days = VALUE #( ( label = '4' ) ( label = '5' ) ) ) )
          tasks = VALUE #(
            ( phase = 'Design' task = 'Blueprint' duration = 2
              cells = VALUE #( ( mark = 'X' ) ( mark = 'X' ) ( mark = '' ) ( mark = '' ) ( mark = '' ) ) ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `D1:H1`           iv_msg = `EX13 header thang gop 5 ngay` ).
    assert_has( iv_text = lv_out iv_sub = `W32`             iv_msg = `EX13 header tuan` ).
    assert_has( iv_text = lv_out iv_sub = `state="frozen"`  iv_msg = `EX13 freeze panes giu nguyen` ).
  ENDMETHOD.


  METHOD ex16_advanced.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex16_advanced( )->get_file(
          iv_taxcode = '0101234567' iv_show_secret = abap_false ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Draft,Released,Closed` iv_msg = `EX16 data validation dong` ).
    assert_has( iv_text = lv_out iv_sub = `(public copy)`         iv_msg = `EX16 khoi nghich dao` ).
    cl_abap_unit_assert=>assert_false( act = xsdbool( lv_out CS `CONFIDENTIAL` )
                                       msg = `EX16 khoi dieu kien bi bo` ).

    DATA(lt_chars) = zcl_xlwb_ex16_advanced=>to_chars( iv_text = 'ABC' iv_len = 10 ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_chars ) exp = 10
                                        msg = `EX16 matrix layout 10 o` ).
    cl_abap_unit_assert=>assert_initial( act = lt_chars[ 10 ]-c
                                         msg = `EX16 o cuoi rong` ).
  ENDMETHOD.


  METHOD ex20_logo.
    TRY.
        DATA(lv_file) = NEW zcl_xlwb_ex20_logo( )->get_file( VALUE #(
          logo     = zcl_xlwb_ex20_logo=>c_demo_logo
          invno    = 'INV-2026-0042'
          customer = 'ACME Corp'
          items    = VALUE #( ( matnr = 'SLAB-WHITE' amount = '2100.00' )
                              ( matnr = 'SLAB-GREY'  amount = '1327.50' ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( lv_file ).
    DATA(lv_media) = abap_false.
    DATA(lv_draw)  = abap_false.
    LOOP AT lo_zip->files INTO DATA(ls_f).
      IF ls_f-name CS 'xl/media/'.                lv_media = abap_true. ENDIF.
      IF ls_f-name CS 'xl/drawings/drawing1.xml'. lv_draw  = abap_true. ENDIF.
    ENDLOOP.
    cl_abap_unit_assert=>assert_true( act = lv_media msg = `EX20 logo vao xl/media` ).
    cl_abap_unit_assert=>assert_true( act = lv_draw  msg = `EX20 drawing part` ).

    lo_zip->get( EXPORTING name            = 'xl/sharedStrings.xml'
                 IMPORTING content         = DATA(lv_raw)
                 EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    DATA(lv_ss) = as_text( lv_raw ).
    assert_has( iv_text = lv_ss iv_sub = `INVOICE INV-2026-0042` iv_msg = `EX20 header` ).
    assert_has( iv_text = lv_ss iv_sub = `Total: 3427.5`         iv_msg = `EX20 aggregate` ).
  ENDMETHOD.


  METHOD ex21_packlist.
    DATA(ls_ctx) = zcl_xlwb_ex21_packlist=>sample_context( ).
    TRY.
        DATA(ls_file) = zcl_xlwb_runtime=>render_with_template(
          iv_template = cl_abap_conv_codepage=>create_out( )->convert(
                          zcl_xlwb_ex21_packlist=>get_template( ) )
          iv_engine   = `SSML`
          ir_context  = REF #( ls_ctx ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.

    DATA(lv_out) = as_text( ls_file-content ).
    assert_has( iv_text = lv_out iv_sub = `PACKING LIST`          iv_msg = `EX21 tieu de tu template` ).
    assert_has( iv_text = lv_out iv_sub = `CAS-2025-0645`         iv_msg = `EX21 so invoice` ).
    assert_has( iv_text = lv_out iv_sub = `TTR-1-NC20081645-2511` iv_msg = `EX21 item cuoi` ).
    assert_has( iv_text = lv_out iv_sub = `>6170<`                iv_msg = `EX21 dong TOTAL` ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-extension exp = `xlsx`
      msg = `EX21 output phai la .xlsx (ssml2xlsx)` ).
    cl_abap_unit_assert=>assert_equals( act = ls_file-source exp = `CODE`
      msg = |EX21 converter phai chay sach, source: { ls_file-source }| ).

    DATA(lv_rows) = 0.
    DATA(lv_rest) = lv_out.
    DATA(lv_pos) = find( val = lv_rest sub = `<row r=` ).
    WHILE lv_pos >= 0.
      lv_rows += 1.
      lv_rest = substring( val = lv_rest off = lv_pos + 7 ).
      lv_pos = find( val = lv_rest sub = `<row r=` ).
    ENDWHILE.
    cl_abap_unit_assert=>assert_equals( act = lv_rows exp = 28
      msg = |EX21 ky vong 28 dong, thuc te { lv_rows }| ).
  ENDMETHOD.


  METHOD ex02a_hlabels.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex02a_hlabels( )->get_file( VALUE #(
          ( name = 'ACME Corp'  city = 'Chicago' )
          ( name = 'Globex Ltd' city = 'London' ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Label 1` iv_msg = `EX02A tieu de nhan 1` ).
    assert_has( iv_text = lv_out iv_sub = `Label 2` iv_msg = `EX02A tieu de nhan 2` ).
    assert_has( iv_text = lv_out iv_sub = `Globex Ltd` iv_msg = `EX02A ten nhan 2` ).
    " marker cell loop khong duoc lot ra file ket qua
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_out CS `#labels` ) msg = `EX02A marker bi strip` ).
  ENDMETHOD.


  METHOD ex08g_grouping.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex08g_grouping( )->get_file( VALUE #(
          ( position = 'Manager' emps = VALUE #(
              ( name = 'Nguyen Van A' role = 'Sales manager' )
              ( name = 'Tran Thi B'   role = 'Plant manager' ) ) )
          ( position = 'Clerk' emps = VALUE #(
              ( name = 'Le Van C' role = 'Warehouse clerk' ) ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Manager (2 employees)` iv_msg = `EX08G summary + cnt` ).
    assert_has( iv_text = lv_out iv_sub = `outlineLevel="1"` iv_msg = `EX08G dong chi tiet co outline` ).
    assert_has( iv_text = lv_out iv_sub = `outlinePr summaryBelow="0"` iv_msg = `EX08G summary nam tren nhom` ).
    assert_has( iv_text = lv_out iv_sub = `outlineLevelRow="1"` iv_msg = `EX08G sheetFormatPr` ).
    assert_has( iv_text = lv_out iv_sub = `Le Van C` iv_msg = `EX08G nhom 2` ).
    " marker khong lot ra file
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_out CS `*group` ) msg = `EX08G marker bi strip` ).
  ENDMETHOD.


  METHOD ex09_dyntable.
    " kich thuoc co dinh de test on dinh; production dung random
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex09_dyntable( )->get_file(
                                  iv_rows = 4 iv_cols = 5 ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `DYNAMIC TABLE 4 x 5` iv_msg = `EX09 tieu de kich thuoc` ).
    assert_has( iv_text = lv_out iv_sub = `Col 5`    iv_msg = `EX09 header cot cuoi (cell loop)` ).
    assert_has( iv_text = lv_out iv_sub = `Row 4`    iv_msg = `EX09 dong cuoi (row loop)` ).
    assert_has( iv_text = lv_out iv_sub = `Row total` iv_msg = `EX09 cot tong dong` ).
    assert_has( iv_text = lv_out iv_sub = `Col total` iv_msg = `EX09 dong tong cot` ).
    " marker khong duoc lot ra file
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_out CS `#cells` ) msg = `EX09 marker bi strip` ).
  ENDMETHOD.


  METHOD ex04_bundle.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex04_bundle( )->get_file( VALUE #(
          orderno  = 'PO-2026-0815'
          customer = 'ACME Corp'
          items    = VALUE #( ( matnr = 'SLAB-WHITE-3CM' qty = 20 amount = '2100.00' )
                              ( matnr = 'SLAB-GREY-2CM'  qty = 15 amount = '1327.50' ) )
          labels   = VALUE #( ( name = 'ACME Corp'  city = 'Chicago' )
                              ( name = 'Globex Ltd' city = 'London' ) ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `name="Order"`  iv_msg = `EX04 sheet Order` ).
    assert_has( iv_text = lv_out iv_sub = `name="Labels"` iv_msg = `EX04 sheet Labels` ).
    assert_has( iv_text = lv_out iv_sub = `PO-2026-0815`     iv_msg = `EX04 orderno` ).
    assert_has( iv_text = lv_out iv_sub = `3427.5`           iv_msg = `EX04 sum amount` ).
    assert_has( iv_text = lv_out iv_sub = `Label 2 / 2`      iv_msg = `EX04 nhan thu 2` ).
  ENDMETHOD.


  METHOD ex14_tree.
    DATA(lt_nodes) = VALUE zcl_xlwb_ex14_tree=>ty_nodes(
      ( node_key = 'G1' text = 'Raw slabs' )
      ( node_key = 'I1' parent_key = 'G1' text = 'White quartz 3cm' qty = 120 )
      ( node_key = 'I2' parent_key = 'G1' text = 'Grey quartz 2cm'  qty = 80 )
      ( node_key = 'G2' text = 'Finished goods' )
      ( node_key = 'I3' parent_key = 'G2' text = 'Countertop A'     qty = 35 ) ).

    " helper chuyen flat -> nested
    DATA(lt_groups) = zcl_xlwb_ex14_tree=>to_nested( lt_nodes ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_groups ) exp = 2 ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_groups[ 1 ]-items ) exp = 2 ).

    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex14_tree( )->get_file( lt_nodes ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Raw slabs`       iv_msg = `EX14 nhom 1` ).
    assert_has( iv_text = lv_out iv_sub = `Subtotal Raw slabs (2 items)` iv_msg = `EX14 subtotal + cnt` ).
    assert_has( iv_text = lv_out iv_sub = `200`             iv_msg = `EX14 sum qty nhom 1` ).
    assert_has( iv_text = lv_out iv_sub = `Countertop A`    iv_msg = `EX14 item nhom 2` ).
  ENDMETHOD.


  METHOD ex17_chart.
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_ex17_chart( )->get_file( VALUE #(
          ( name = 'Jan' revenue = '1200.00' )
          ( name = 'Feb' revenue = '1550.50' ) ) ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `Jan`     iv_msg = `EX17 thang 1` ).
    assert_has( iv_text = lv_out iv_sub = `1550.5`  iv_msg = `EX17 so lieu` ).
  ENDMETHOD.


  METHOD ex22_source.
    DATA(lo_source) = NEW zcl_xlwb_src_order( ).

    " get_context doc duoc key tu JSON
    TRY.
        DATA(lr_ctx) = lo_source->zif_xlwb_source~get_context( `{ "orderno": "PO-UT-01" }` ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    ASSIGN lr_ctx->* TO FIELD-SYMBOL(<ls_order>).
    ASSIGN COMPONENT 'ORDERNO' OF STRUCTURE <ls_order> TO FIELD-SYMBOL(<lv_no>).
    cl_abap_unit_assert=>assert_equals( act = |{ <lv_no> }| exp = `PO-UT-01` ).

    " render_by_source can template trong bang -> o day chi test get_file (fallback)
    TRY.
        DATA(lv_out) = as_text( lo_source->get_file( ) ).
      CATCH zcx_xlwb INTO lx.
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `PO-2026-DEMO`   iv_msg = `EX22 sample orderno` ).
    assert_has( iv_text = lv_out iv_sub = `SLAB-WHITE-3CM` iv_msg = `EX22 item` ).
    assert_has( iv_text = lv_out iv_sub = `3427.5`         iv_msg = `EX22 sum amount` ).
  ENDMETHOD.


  METHOD ex23_bcnctp.
    " get_file dung sample tinh (khong dung DB) -> chay duoc moi client
    TRY.
        DATA(lv_out) = as_text( NEW zcl_xlwb_src_bcnctp( )->get_file( ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = lx->get_text( ) ).
        RETURN.
    ENDTRY.
    assert_has( iv_text = lv_out iv_sub = `BÁO CÁO NHU CẦU THÀNH PHẨM` iv_msg = `EX23 tieu de` ).
    assert_has( iv_text = lv_out iv_sub = `W35/2026`     iv_msg = `EX23 cot tuan dau` ).
    assert_has( iv_text = lv_out iv_sub = `CX W38/2026`  iv_msg = `EX23 cap cot cuoi (4 tuan)` ).
    assert_has( iv_text = lv_out iv_sub = `M50602014`    iv_msg = `EX23 dong PH4` ).
    assert_has( iv_text = lv_out iv_sub = `>65123<`      iv_msg = `EX23 so lieu qty` ).
    assert_has( iv_text = lv_out iv_sub = `state="frozen"` iv_msg = `EX23 freeze 6 cot + header` ).
    " sample 4 tuan = 8 nhan -> cot cuoi 14 = N; filter tren dong header 3
    assert_has( iv_text = lv_out iv_sub = `<autoFilter ref="A3:N3"/>` iv_msg = `EX23 auto filter header` ).
  ENDMETHOD.


  METHOD all_demo_files.
    " Moi vi du trong app Fiori phai render duoc file mau (nguon: ZCL_XLWB_DEMO_LOAD)
    DATA(lt_ids) = VALUE string_table(
      ( `EX00` ) ( `EX01` ) ( `EX02` ) ( `EX02A` ) ( `EX03` ) ( `EX04` )
      ( `EX05` ) ( `EX08` ) ( `EX08G` ) ( `EX09` ) ( `EX11` ) ( `EX13` ) ( `EX14` )
      ( `EX16` ) ( `EX17` ) ( `EX20` ) ( `EX21` ) ( `EX22` ) ( `EX23` ) ).

    LOOP AT lt_ids INTO DATA(lv_id).
      TRY.
          DATA(ls_file) = zcl_xlwb_demo_files=>render( lv_id ).
        CATCH zcx_xlwb INTO DATA(lx).
          cl_abap_unit_assert=>fail( msg = |{ lv_id }: { lx->get_text( ) }| ).
          CONTINUE.
      ENDTRY.
      cl_abap_unit_assert=>assert_not_initial( act = ls_file-content
                                               msg = |{ lv_id } file rong| ).
      cl_abap_unit_assert=>assert_not_initial( act = ls_file-file_name
                                               msg = |{ lv_id } thieu ten file| ).
      cl_abap_unit_assert=>assert_not_initial( act = ls_file-mime_type
                                               msg = |{ lv_id } thieu mime type| ).
    ENDLOOP.

    TRY.
        zcl_xlwb_demo_files=>render( `EX99` ).
        cl_abap_unit_assert=>fail( msg = `ma vi du la phai bi tu choi` ).
      CATCH zcx_xlwb.
    ENDTRY.
  ENDMETHOD.
ENDCLASS.
