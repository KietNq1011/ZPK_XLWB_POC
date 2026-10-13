"! <p class="shorttext synchronized" lang="en">XLWB Cloud: engine unit tests</p>
CLASS zcl_xlwb_engine_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
  PRIVATE SECTION.
    TYPES: BEGIN OF ty_item,
             matnr TYPE c LENGTH 18,
             qty   TYPE p LENGTH 8 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY.

    METHODS wrap
      IMPORTING iv_body       TYPE string
      RETURNING VALUE(rv_xml) TYPE string.
    METHODS render
      IMPORTING iv_body       TYPE string
                ir_context    TYPE REF TO data
      RETURNING VALUE(rv_out) TYPE string.
    METHODS assert_contains
      IMPORTING iv_out TYPE string
                iv_sub TYPE string
                iv_msg TYPE string.
    METHODS count_of
      IMPORTING iv_out          TYPE string
                iv_sub          TYPE string
      RETURNING VALUE(rv_count) TYPE i.

    METHODS pcre_probe          FOR TESTING.
    METHODS values_and_types    FOR TESTING.
    METHODS nested_path         FOR TESTING.
    METHODS row_loop_aggregates FOR TESTING.
    METHODS nested_loop         FOR TESTING.
    METHODS conditional_blocks  FOR TESTING.
    METHODS cell_loop_columns   FOR TESTING.
    METHODS sheet_loop          FOR TESTING.
    METHODS merge_and_break     FOR TESTING.
    METHODS formula_preserved   FOR TESTING.
    METHODS missing_field_error FOR TESTING.
    METHODS excel_row_index    FOR TESTING.
    METHODS merge_same_value   FOR TESTING.
    METHODS merge_span_number  FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_ENGINE_TEST IMPLEMENTATION.


  METHOD wrap.
    rv_xml =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      iv_body &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD render.
    TRY.
        rv_out = NEW zcl_xlwb_engine( )->render(
          iv_template = wrap( iv_body )
          ir_context  = ir_context ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = |zcx_xlwb: { lx->get_text( ) }| ).
    ENDTRY.
  ENDMETHOD.


  METHOD assert_contains.
    IF iv_out NS iv_sub.
      cl_abap_unit_assert=>fail(
        msg = |{ iv_msg } — missing '{ iv_sub }' in: { substring( val = iv_out off = 0 len = nmin( val1 = strlen( iv_out ) val2 = 1200 ) ) }| ).
    ENDIF.
  ENDMETHOD.


  METHOD count_of.
    DATA(lv_rest) = iv_out.
    DATA(lv_pos) = find( val = lv_rest sub = iv_sub ).
    WHILE lv_pos >= 0.
      rv_count += 1.
      lv_rest = substring( val = lv_rest off = lv_pos + strlen( iv_sub ) ).
      lv_pos = find( val = lv_rest sub = iv_sub ).
    ENDWHILE.
  ENDMETHOD.


  METHOD pcre_probe.
    DATA lv_tok TYPE string.
    FIND PCRE '\{\{([^}]+)\}\}' IN `Hello {{name}}` SUBMATCHES lv_tok.
    cl_abap_unit_assert=>assert_equals( act = sy-subrc exp = 0 msg = `find pcre subrc` ).
    cl_abap_unit_assert=>assert_equals( act = lv_tok exp = `name`
                                        msg = |submatch got '{ lv_tok }'| ).

    DATA lv_off TYPE i.
    DATA lv_len TYPE i.
    CLEAR lv_tok.
    FIND PCRE '\{\{([^}]+)\}\}' IN `Hello {{name}}`
      MATCH OFFSET lv_off MATCH LENGTH lv_len SUBMATCHES lv_tok.
    cl_abap_unit_assert=>assert_equals( act = lv_tok exp = `name`
                                        msg = |submatch+offset got '{ lv_tok }' off { lv_off } len { lv_len }| ).

    SPLIT to_upper( `header.customer` ) AT '.' INTO TABLE DATA(lt_segs).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_segs ) exp = 2
                                        msg = |split on expression gave { lines( lt_segs ) } lines| ).
  ENDMETHOD.


  METHOD values_and_types.
    DATA: BEGIN OF ls_ctx,
            name    TYPE string VALUE 'World',
            amount  TYPE p LENGTH 8 DECIMALS 2 VALUE '1050.50',
            docdate TYPE d VALUE '20260821',
          END OF ls_ctx.

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">Hello {{name}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{amount}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{docdate}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `Hello World`
                     iv_msg = `inline value substitution` ).
    assert_contains( iv_out = lv_out iv_sub = `>1050.5`
                     iv_msg = `numeric full-cell value` ).
    assert_contains( iv_out = lv_out iv_sub = `Number`
                     iv_msg = `ss:Type switched to Number` ).
    assert_contains( iv_out = lv_out iv_sub = `2026-08-21T00:00:00.000`
                     iv_msg = `date full-cell value` ).
    assert_contains( iv_out = lv_out iv_sub = `DateTime`
                     iv_msg = `ss:Type switched to DateTime` ).
  ENDMETHOD.


  METHOD nested_path.
    DATA: BEGIN OF ls_ctx,
            BEGIN OF header,
              customer TYPE string VALUE 'CASLA',
            END OF header,
          END OF ls_ctx.

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">To: {{header.customer}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `To: CASLA` iv_msg = `nested path` ).
  ENDMETHOD.


  METHOD row_loop_aggregates.
    DATA: BEGIN OF ls_ctx,
            items TYPE ty_items,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( matnr = 'SLAB-001' qty = '10.00' )
                            ( matnr = 'SLAB-002' qty = '20.50' )
                            ( matnr = 'SLAB-003' qty = '4.25' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">Pos</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{#items}}{{@index}}/{{@count}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{matnr}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{qty}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{/items}}</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{sum:items.qty}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{cnt:items}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{max:items.qty}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `SLAB-00` )
      exp = 3
      msg = `3 item rows expected` ).
    assert_contains( iv_out = lv_out iv_sub = `1/3` iv_msg = `@index/@count` ).
    assert_contains( iv_out = lv_out iv_sub = `3/3` iv_msg = `@index of last line` ).
    assert_contains( iv_out = lv_out iv_sub = `>34.75<` iv_msg = `sum aggregate` ).
    assert_contains( iv_out = lv_out iv_sub = `>3<` iv_msg = `cnt aggregate` ).
    assert_contains( iv_out = lv_out iv_sub = `>20.5` iv_msg = `max aggregate` ).
  ENDMETHOD.


  METHOD nested_loop.
    TYPES: BEGIN OF ty_grp,
             route TYPE string,
             items TYPE ty_items,
           END OF ty_grp.
    DATA: BEGIN OF ls_ctx,
            groups TYPE STANDARD TABLE OF ty_grp WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-groups = VALUE #(
      ( route = 'HAN-SGN' items = VALUE #( ( matnr = 'A1' qty = 1 ) ( matnr = 'A2' qty = 2 ) ) )
      ( route = 'SGN-DAD' items = VALUE #( ( matnr = 'B1' qty = 3 ) ) ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{#groups}}Route {{route}}</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{#items}}- {{matnr}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{/items}}</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">Subtotal {{route}}: {{sum:items.qty}}{{/groups}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `Route HAN-SGN` iv_msg = `group 1 header` ).
    assert_contains( iv_out = lv_out iv_sub = `- A2` iv_msg = `nested line` ).
    assert_contains( iv_out = lv_out iv_sub = `Subtotal HAN-SGN: 3` iv_msg = `subtotal group 1` ).
    assert_contains( iv_out = lv_out iv_sub = `Subtotal SGN-DAD: 3` iv_msg = `subtotal group 2` ).
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `Route ` )
      exp = 2
      msg = `2 groups expected` ).
  ENDMETHOD.


  METHOD conditional_blocks.
    DATA: BEGIN OF ls_ctx,
            show_remark TYPE abap_bool VALUE abap_true,
            items       TYPE ty_items,
          END OF ls_ctx.

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{?show_remark}}REMARK VISIBLE{{/show_remark}}</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{?items}}HAS ITEMS{{/items}}</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{^items}}NO ITEMS{{/items}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `REMARK VISIBLE` iv_msg = `? kept when true` ).
    IF lv_out CS `HAS ITEMS`.
      cl_abap_unit_assert=>fail( msg = `? block must be dropped for empty table` ).
    ENDIF.
    assert_contains( iv_out = lv_out iv_sub = `NO ITEMS` iv_msg = `^ kept when empty` ).
  ENDMETHOD.


  METHOD cell_loop_columns.
    TYPES: BEGIN OF ty_col,
             title TYPE string,
           END OF ty_col.
    DATA: BEGIN OF ls_ctx,
            cols TYPE STANDARD TABLE OF ty_col WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-cols = VALUE #( ( title = 'Mon' ) ( title = 'Tue' ) ( title = 'Wed' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">Day:</Data></Cell>` &&
                `<Cell ss:StyleID="hd"><Data ss:Type="String">{{#cols>}}{{title}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `Mon` iv_msg = `col 1` ).
    assert_contains( iv_out = lv_out iv_sub = `Wed` iv_msg = `col 3` ).
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `hd` )
      exp = 3
      msg = `template cell style must be cloned to every column` ).
  ENDMETHOD.


  METHOD sheet_loop.
    TYPES: BEGIN OF ty_sheet,
             label TYPE string,
           END OF ty_sheet.
    DATA: BEGIN OF ls_ctx,
            sheets TYPE STANDARD TABLE OF ty_sheet WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-sheets = VALUE #( ( label = 'North' ) ( label = 'South' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="{{#sheets}}Region {{@index}}"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{label}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `Region 1` iv_msg = `sheet 1 name` ).
    assert_contains( iv_out = lv_out iv_sub = `Region 2` iv_msg = `sheet 2 name` ).
    assert_contains( iv_out = lv_out iv_sub = `North` iv_msg = `sheet 1 content` ).
    assert_contains( iv_out = lv_out iv_sub = `South` iv_msg = `sheet 2 content` ).
  ENDMETHOD.


  METHOD merge_and_break.
    DATA: BEGIN OF ls_ctx,
            subs   TYPE ty_items,
            labels TYPE ty_items,
          END OF ls_ctx.
    ls_ctx-subs   = VALUE #( ( matnr = 'X' ) ( matnr = 'Y' ) ( matnr = 'Z' ) ).
    ls_ctx-labels = VALUE #( ( matnr = 'L1' ) ( matnr = 'L2' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{*mergedown:subs}}Group</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{#labels}}{{*break}}{{matnr}}{{/labels}}</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `MergeDown` iv_msg = `dynamic MergeDown` ).
    assert_contains( iv_out = lv_out iv_sub = `"2"` iv_msg = `MergeDown = lines-1` ).
    assert_contains( iv_out = lv_out iv_sub = `RowBreaks` iv_msg = `page breaks element` ).
    assert_contains( iv_out = lv_out iv_sub = `L2` iv_msg = `second label` ).
  ENDMETHOD.


  METHOD formula_preserved.
    DATA: BEGIN OF ls_ctx,
            dummy TYPE string VALUE 'x',
          END OF ls_ctx.

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell ss:Formula="=SUM(R[-2]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `SUM(R[-2]C:R[-1]C)` iv_msg = `formula untouched` ).
  ENDMETHOD.


  METHOD excel_row_index.
    " Excel sinh <Row ss:Index="N"> khi template co dong trong.
    " Engine phai thay neo do bang <Row/> rong, neu khong footer se nam sai cho
    " sau khi vung lap nhan ban.
    DATA: BEGIN OF ls_ctx,
            items TYPE ty_items,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( matnr = 'A1' ) ( matnr = 'A2' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{#items}}{{matnr}}{{/items}}</Data></Cell></Row>` &&
                `<Row ss:Index="5"><Cell><Data ss:Type="String">FOOTER</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    " 1 dong lap x 2 item + 3 dong rong thay cho neo Index=5 + 1 dong footer = 6
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `<Row` )
      exp = 6
      msg = |ky vong 6 dong sau khi normalize ss:Index| ).

    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_out CS `ss:Index` )
      msg = `ss:Index tren Row phai bi thay bang Row rong` ).

    " footer phai nam SAU item cuoi
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( find( val = lv_out sub = `FOOTER` ) > find( val = lv_out sub = `A2` ) )
      msg = `footer phai nam sau dong item cuoi` ).
  ENDMETHOD.


  METHOD missing_field_error.
    DATA: BEGIN OF ls_ctx,
            name TYPE string VALUE 'x',
          END OF ls_ctx.

    TRY.
        NEW zcl_xlwb_engine( )->render(
          iv_template = wrap( `<Worksheet ss:Name="S"><Table>` &&
                              `<Row><Cell><Data ss:Type="String">{{no_such_field}}</Data></Cell></Row>` &&
                              `</Table></Worksheet>` )
          ir_context  = REF #( ls_ctx ) ).
        cl_abap_unit_assert=>fail( msg = `zcx_xlwb expected for unknown path` ).
      CATCH zcx_xlwb INTO DATA(lx).
        IF lx->get_text( ) NS `no_such_field`.
          cl_abap_unit_assert=>fail( msg = `error text should name the missing path` ).
        ENDIF.
    ENDTRY.
  ENDMETHOD.


  METHOD merge_same_value.
    " {{*mergesame:cont}} gop DOC cac dong lien tiep cung so container.
    " Du lieu: A A B C C C  -> 3 nhom: span 2, 1, 3
    TYPES: BEGIN OF ty_line,
             cont  TYPE string,
             matnr TYPE string,
             qty   TYPE i,
           END OF ty_line.
    DATA: BEGIN OF ls_ctx,
            items TYPE STANDARD TABLE OF ty_line WITH EMPTY KEY,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( cont = 'A' matnr = 'M1' qty = 11 ) ( cont = 'A' matnr = 'M2' qty = 12 )
                            ( cont = 'B' matnr = 'M3' qty = 13 )
                            ( cont = 'C' matnr = 'M4' qty = 14 ) ( cont = 'C' matnr = 'M5' qty = 15 )
                            ( cont = 'C' matnr = 'M6' qty = 16 ) ).

    " co cot qty dung SAU cot mergesame de bat bug truot cot: o bi phu bi go
    " khoi dong thi o ke tiep phai duoc neo ss:Index, khong thi moi o phia sau
    " truot sang trai va rơi vao vung merge -> Excel bao file corrupt
    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row>` &&
                `<Cell><Data ss:Type="String">{{#items}}{{matnr}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{*mergesame:cont}}{{cont}}</Data></Cell>` &&
                `<Cell><Data ss:Type="String">{{qty}}{{/items}}</Data></Cell>` &&
                `</Row></Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    " moi dong item van con -> 6 dong
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `<Row` ) exp = 6
      msg = `gop o khong duoc lam mat dong` ).
    " nhom span 2 -> MergeDown="1"; nhom span 3 -> MergeDown="2"; nhom span 1 -> khong co
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `ss:MergeDown="1"` ) exp = 1
      msg = `nhom 2 dong -> MergeDown=1` ).
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `ss:MergeDown="2"` ) exp = 1
      msg = `nhom 3 dong -> MergeDown=2` ).
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `ss:MergeDown` ) exp = 2
      msg = `nhom 1 dong khong nhan MergeDown` ).
    " gia tri container chi con xuat hien o dong dau moi nhom: A, B, C = 3 lan
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `>C<` ) exp = 1
      msg = `o bi phu phai trong` ).
    " o bi phu phai bi GO HAN khoi dong (cell rong trong vung merge lam Excel
    " bao "file is corrupt"): 6 matnr + 3 dau nhom + 6 qty = 15
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `<Cell` ) exp = 15
      msg = `o bi merge phu khong duoc ton tai trong dong` ).
    " 3 dong bi phu (item 2, 5, 6): o qty ke sau o bi go phai duoc neo cot 3
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `ss:Index="3"` ) exp = 3
      msg = `o sau o bi go phai co ss:Index giu dung cot` ).
    " du 6 gia tri qty, khong mat cot nao
    assert_contains( iv_out = lv_out iv_sub = `>16<` iv_msg = `qty dong cuoi con nguyen` ).
    " cot khong gop thi khong bi anh huong
    assert_contains( iv_out = lv_out iv_sub = `M6` iv_msg = `cot khac van render du` ).
  ENDMETHOD.


  METHOD merge_span_number.
    " {{*mergeacross=N}} / {{*mergedown=N}}: N la so o duoc gop.
    " N viet truc tiep hoac lay tu context.
    DATA: BEGIN OF ls_ctx,
            cols  TYPE i VALUE 4,
            items TYPE ty_items,
          END OF ls_ctx.
    ls_ctx-items = VALUE #( ( matnr = 'A' ) ( matnr = 'B' ) ).

    DATA(lv_out) = render(
      iv_body = `<Worksheet ss:Name="S"><Table>` &&
                `<Row><Cell><Data ss:Type="String">{{*mergeacross=3}}Tieu de</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{*mergeacross=cols}}Theo context</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{*mergedown=2}}Doc</Data></Cell></Row>` &&
                `<Row><Cell><Data ss:Type="String">{{*mergeacross=1}}Khong gop</Data></Cell></Row>` &&
                `</Table></Worksheet>`
      ir_context = REF #( ls_ctx ) ).

    assert_contains( iv_out = lv_out iv_sub = `ss:MergeAcross="2"`
                     iv_msg = `=3 -> gop 3 o -> MergeAcross=2` ).
    assert_contains( iv_out = lv_out iv_sub = `ss:MergeAcross="3"`
                     iv_msg = `=cols (4) -> MergeAcross=3` ).
    assert_contains( iv_out = lv_out iv_sub = `ss:MergeDown="1"`
                     iv_msg = `=2 -> MergeDown=1` ).
    cl_abap_unit_assert=>assert_equals(
      act = count_of( iv_out = lv_out iv_sub = `ss:MergeAcross="0"` ) exp = 0
      msg = `=1 nghia la khong gop -> khong sinh thuoc tinh` ).
    assert_contains( iv_out = lv_out iv_sub = `Khong gop` iv_msg = `text van con` ).
  ENDMETHOD.
ENDCLASS.
