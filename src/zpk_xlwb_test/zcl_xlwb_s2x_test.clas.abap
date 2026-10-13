"! <p class="shorttext synchronized" lang="en">XLWB: ssml2xlsx converter tests</p>
CLASS zcl_xlwb_s2x_test DEFINITION
  PUBLIC FINAL CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS wrap
      IMPORTING iv_body       TYPE string
                iv_styles     TYPE string OPTIONAL
      RETURNING VALUE(rv_xml) TYPE string.
    METHODS conv
      IMPORTING iv_body        TYPE string
                iv_styles      TYPE string OPTIONAL
      RETURNING VALUE(rv_xlsx) TYPE xstring.
    METHODS part
      IMPORTING iv_xlsx        TYPE xstring
                iv_name        TYPE string
      RETURNING VALUE(rv_text) TYPE string.
    METHODS assert_has
      IMPORTING iv_text TYPE string
                iv_sub  TYPE string
                iv_msg  TYPE string.

    METHODS typed_values       FOR TESTING.
    METHODS styles_mapped      FOR TESTING.
    METHODS merges_exact       FOR TESTING.
    METHODS merge_border_fillers FOR TESTING.
    METHODS formula_r1c1_to_a1 FOR TESTING.
    METHODS multi_sheet        FOR TESTING.
    METHODS view_break_freeze  FOR TESTING.
    METHODS validation_list    FOR TESTING.
    METHODS still_valid_xlsx   FOR TESTING.
ENDCLASS.



CLASS ZCL_XLWB_S2X_TEST IMPLEMENTATION.


  METHOD wrap.
    rv_xml =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      COND string( WHEN iv_styles IS NOT INITIAL THEN |<Styles>{ iv_styles }</Styles>| ) &&
      iv_body &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD conv.
    TRY.
        rv_xlsx = zcl_xlwb_ssml2xlsx=>convert( wrap( iv_body = iv_body iv_styles = iv_styles ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        cl_abap_unit_assert=>fail( msg = |convert: { lx->get_text( ) }| ).
    ENDTRY.
  ENDMETHOD.


  METHOD part.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( iv_xlsx ).
    lo_zip->get( EXPORTING name = iv_name
                 IMPORTING content = DATA(lv_raw)
                 EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    IF sy-subrc = 0 AND lv_raw IS NOT INITIAL.
      rv_text = cl_abap_conv_codepage=>create_in( )->convert( lv_raw ).
    ENDIF.
  ENDMETHOD.


  METHOD assert_has.
    cl_abap_unit_assert=>assert_true(
      act = xsdbool( iv_text CS iv_sub )
      msg = |{ iv_msg } — missing '{ iv_sub }' in: { substring( val = iv_text off = 0
               len = nmin( val1 = strlen( iv_text ) val2 = 1500 ) ) }| ).
  ENDMETHOD.


  METHOD typed_values.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell><Data ss:Type="String">Hello world</Data></Cell>` &&
      `<Cell><Data ss:Type="Number">1234.5</Data></Cell>` &&
      `<Cell><Data ss:Type="DateTime">2026-08-21T00:00:00.000</Data></Cell></Row>` &&
      `</Table></Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    DATA(lv_sst)   = part( iv_xlsx = lv_out iv_name = 'xl/sharedStrings.xml' ).

    assert_has( iv_text = lv_sst iv_sub = `Hello world` iv_msg = `string vao sharedStrings` ).
    assert_has( iv_text = lv_sheet iv_sub = `t="s"` iv_msg = `cell string t=s` ).
    assert_has( iv_text = lv_sheet iv_sub = `<v>1234.5</v>` iv_msg = `number giu nguyen <v>` ).

    " 2026-08-21 -> serial ngay Excel (epoch 1899-12-30)
    DATA(lv_serial) = CONV d( '20260821' ) - CONV d( '18991230' ).
    assert_has( iv_text = lv_sheet iv_sub = |<v>{ lv_serial }</v>| iv_msg = `date -> serial` ).
    " o date khong style -> xf numFmt 14
    DATA(lv_styles) = part( iv_xlsx = lv_out iv_name = 'xl/styles.xml' ).
    assert_has( iv_text = lv_styles iv_sub = `numFmtId="14"` iv_msg = `numFmt ngay mac dinh` ).
  ENDMETHOD.


  METHOD styles_mapped.
    DATA(lv_out) = conv(
      iv_styles =
        `<Style ss:ID="hd"><Font ss:Bold="1" ss:Size="14" ss:Color="#FF0000" ss:FontName="Arial"/>` &&
        `<Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/>` &&
        `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/></Borders>` &&
        `<Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>` &&
        `<Style ss:ID="num"><NumberFormat ss:Format="#,##0.00 &quot;kg&quot;"/></Style>`
      iv_body =
        `<Worksheet ss:Name="S"><Table>` &&
        `<Column ss:AutoFitWidth="0" ss:Width="120"/>` &&
        `<Row ss:Height="22"><Cell ss:StyleID="hd"><Data ss:Type="String">Head</Data></Cell>` &&
        `<Cell ss:StyleID="num"><Data ss:Type="Number">12.5</Data></Cell></Row>` &&
        `<Row><Cell ss:StyleID="hd"/></Row>` &&
        `</Table></Worksheet>` ).

    DATA(lv_styles) = part( iv_xlsx = lv_out iv_name = 'xl/styles.xml' ).
    DATA(lv_sheet)  = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).

    assert_has( iv_text = lv_styles iv_sub = `<b/>` iv_msg = `font bold` ).
    assert_has( iv_text = lv_styles iv_sub = `rgb="FFFF0000"` iv_msg = `font color` ).
    assert_has( iv_text = lv_styles iv_sub = `val="Arial"` iv_msg = `font name` ).
    assert_has( iv_text = lv_styles iv_sub = `fgColor rgb="FFDDEBF7"` iv_msg = `fill` ).
    assert_has( iv_text = lv_styles iv_sub = `<bottom style="medium">` iv_msg = `border weight 2 -> medium` ).
    assert_has( iv_text = lv_styles iv_sub = `horizontal="center"` iv_msg = `alignment` ).
    assert_has( iv_text = lv_styles iv_sub = `wrapText="1"` iv_msg = `wrap text` ).
    assert_has( iv_text = lv_styles iv_sub = `formatCode="#,##0.00 &quot;kg&quot;"` iv_msg = `custom numFmt` ).
    assert_has( iv_text = lv_sheet iv_sub = `customWidth="1"` iv_msg = `do rong cot` ).
    assert_has( iv_text = lv_sheet iv_sub = `ht="22"` iv_msg = `chieu cao dong` ).
    " o rong co style van duoc giu (khung/border)
    assert_has( iv_text = lv_sheet iv_sub = `<c r="A2" s="` iv_msg = `o rong styled giu lai` ).
  ENDMETHOD.


  METHOD merges_exact.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell ss:MergeAcross="2"><Data ss:Type="String">span3</Data></Cell></Row>` &&
      `<Row><Cell ss:Index="2" ss:MergeDown="3"><Data ss:Type="String">down4</Data></Cell></Row>` &&
      `</Table></Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `<mergeCell ref="A1:C1"/>` iv_msg = `mergeacross` ).
    assert_has( iv_text = lv_sheet iv_sub = `<mergeCell ref="B2:B5"/>` iv_msg = `mergedown + ss:Index` ).
  ENDMETHOD.


  METHOD merge_border_fillers.
    " Vung merge co style (khung) phai co O DEM o MOI vi tri bi phu, cung xf —
    " chi ghi o neo thi Excel mat canh duoi + canh phai (khung vo).
    DATA(lv_out) = conv(
      iv_styles = `<Style ss:ID="box"><Borders>` &&
                  `<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
                  `<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
                  `<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
                  `<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
                  `</Borders></Style>`
      iv_body =
        `<Worksheet ss:Name="S"><Table>` &&
        `<Row><Cell ss:MergeAcross="2" ss:MergeDown="2" ss:StyleID="box">` &&
        `<Data ss:Type="String">BOX</Data></Cell></Row>` &&
        `</Table></Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `<mergeCell ref="A1:C3"/>` iv_msg = `vung merge 3x3` ).

    " xf cua o neo — moi o dem phai dung dung xf nay
    DATA lv_xf TYPE string.
    FIND PCRE '(?-x)<c r="A1" s="(\d+)"' IN lv_sheet SUBMATCHES lv_xf.
    cl_abap_unit_assert=>assert_subrc( act = sy-subrc msg = `phai co o neo A1 co style` ).
    cl_abap_unit_assert=>assert_differs( act = lv_xf exp = `0` msg = `o neo phai co xf khung` ).

    " du 9 o cua vung (A1..C3), 8 o dem la <c .../> rong cung xf
    LOOP AT VALUE string_table( ( `B1` ) ( `C1` ) ( `A2` ) ( `B2` ) ( `C2` )
                               ( `A3` ) ( `B3` ) ( `C3` ) ) INTO DATA(lv_cell).
      assert_has( iv_text = lv_sheet
                  iv_sub  = |<c r="{ lv_cell }" s="{ lv_xf }"/>|
                  iv_msg  = |o dem { lv_cell } phai co, cung xf voi o neo| ).
    ENDLOOP.

    " dong 2 va 3 phai ton tai (o dem tao ra dong moi)
    assert_has( iv_text = lv_sheet iv_sub = `<row r="3"` iv_msg = `dong 3 duoc tao cho o dem` ).

    " vung merge KHONG style thi khong sinh o dem (tranh phinh file)
    DATA(lv_plain) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell ss:MergeAcross="2"><Data ss:Type="String">x</Data></Cell></Row>` &&
      `</Table></Worksheet>` ).
    DATA(lv_plain_sheet) = part( iv_xlsx = lv_plain iv_name = 'xl/worksheets/sheet1.xml' ).
    cl_abap_unit_assert=>assert_false(
      act = xsdbool( lv_plain_sheet CS `<c r="B1"` )
      msg = `merge khong style thi khong can o dem` ).
  ENDMETHOD.


  METHOD formula_r1c1_to_a1.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell><Data ss:Type="Number">1</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="Number">2</Data></Cell></Row>` &&
      `<Row><Cell ss:Formula="=SUM(R1C:R[-1]C)"><Data ss:Type="Number">3</Data></Cell>` &&
      `<Cell ss:Formula="=RC[-1]*2"><Data ss:Type="Number">6</Data></Cell></Row>` &&
      `</Table></Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `<f>SUM(A$1:A2)</f>` iv_msg = `R1C1 tuyet doi + tuong doi` ).
    assert_has( iv_text = lv_sheet iv_sub = `<f>A3*2</f>` iv_msg = `RC[-1] cung dong` ).
    assert_has( iv_text = lv_sheet iv_sub = `<v>3</v>` iv_msg = `cached value giu lai` ).
  ENDMETHOD.


  METHOD multi_sheet.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="Order"><Table>` &&
      `<Row><Cell><Data ss:Type="String">o1</Data></Cell></Row></Table></Worksheet>` &&
      `<Worksheet ss:Name="Labels"><Table>` &&
      `<Row><Cell><Data ss:Type="String">l1</Data></Cell></Row></Table></Worksheet>` ).

    DATA(lv_wb) = part( iv_xlsx = lv_out iv_name = 'xl/workbook.xml' ).
    assert_has( iv_text = lv_wb iv_sub = `name="Order"`  iv_msg = `sheet 1` ).
    assert_has( iv_text = lv_wb iv_sub = `name="Labels"` iv_msg = `sheet 2` ).
    cl_abap_unit_assert=>assert_not_initial(
      act = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet2.xml' )
      msg = `part sheet2 phai ton tai` ).
  ENDMETHOD.


  METHOD view_break_freeze.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell><Data ss:Type="String">a</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">b</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">c</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">d</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">e</Data></Cell></Row>` &&
      `</Table>` &&
      `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<PageSetup><Layout x:Orientation="Landscape"/></PageSetup>` &&
      `<FreezePanes/><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane>` &&
      `<DoNotDisplayGridlines/>` &&
      `</WorksheetOptions>` &&
      `<PageBreaks xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<RowBreaks><RowBreak><Row>2</Row></RowBreak><RowBreak><Row>4</Row></RowBreak></RowBreaks>` &&
      `</PageBreaks>` &&
      `</Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `orientation="landscape"` iv_msg = `page setup` ).
    assert_has( iv_text = lv_sheet iv_sub = `showGridLines="0"` iv_msg = `gridlines` ).
    assert_has( iv_text = lv_sheet iv_sub = `xSplit="3"` iv_msg = `freeze xSplit` ).
    assert_has( iv_text = lv_sheet iv_sub = `topLeftCell="D1"` iv_msg = `freeze topLeftCell` ).
    assert_has( iv_text = lv_sheet iv_sub = `state="frozen"` iv_msg = `freeze state` ).
    assert_has( iv_text = lv_sheet iv_sub = `<brk id="2"` iv_msg = `row break 1` ).
    assert_has( iv_text = lv_sheet iv_sub = `<brk id="4"` iv_msg = `row break 2` ).
  ENDMETHOD.


  METHOD validation_list.
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell><Data ss:Type="String">x</Data></Cell></Row>` &&
      `</Table>` &&
      `<DataValidation xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<Range>R2C2</Range><Type>List</Type>` &&
      `<Value>&quot;Draft,Released,Closed&quot;</Value>` &&
      `</DataValidation>` &&
      `</Worksheet>` ).

    DATA(lv_sheet) = part( iv_xlsx = lv_out iv_name = 'xl/worksheets/sheet1.xml' ).
    assert_has( iv_text = lv_sheet iv_sub = `type="list"` iv_msg = `validation list` ).
    assert_has( iv_text = lv_sheet iv_sub = `sqref="B2"` iv_msg = `range R2C2 -> B2` ).
    assert_has( iv_text = lv_sheet iv_sub = `"Draft,Released,Closed"` iv_msg = `danh sach` ).
  ENDMETHOD.


  METHOD still_valid_xlsx.
    " file phai mo lai duoc bang XCO read access va doc dung gia tri o
    DATA(lv_out) = conv(
      `<Worksheet ss:Name="S"><Table>` &&
      `<Row><Cell><Data ss:Type="String">ok</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="Number">7</Data></Cell></Row>` &&
      `</Table></Worksheet>` ).

    DATA(lo_ra) = xco_cp_xlsx=>document->for_file_content( lv_out )->read_access( ).
    DATA(lo_ws) = lo_ra->get_workbook( )->worksheet->at_position( 1 ).
    cl_abap_unit_assert=>assert_true( act = lo_ws->exists( )
                                      msg = `xlsx phai doc lai duoc bang XCO` ).

    DATA lv_text TYPE string.
    lo_ws->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
                   io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
      )->get_cell( )->get_value( )->write_to( REF #( lv_text ) ).
    cl_abap_unit_assert=>assert_equals( act = lv_text exp = `ok` msg = `doc lai o A1` ).
  ENDMETHOD.
ENDCLASS.
