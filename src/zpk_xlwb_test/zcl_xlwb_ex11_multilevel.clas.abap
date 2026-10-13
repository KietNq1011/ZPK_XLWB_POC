"! <p class="shorttext synchronized" lang="en">XLWB ex 4.11/4.12: Multilevel list + dynamic merge</p>
"!
"! Danh sach 3 cap (Route -> Connection -> Flight) kem subtotal tung cap
"! va gop o doc cho cot cap cha. Tuong duong XLWB 4.11 (multilevel list),
"! 4.11a (grid + subtotals) va 4.12/4.12a (dynamic cell merging).
"! Bai hoc:
"! - LOOP LONG NHAU 3 cap: {{#routes}} -> {{#conns}} -> {{#flights}}
"! - trong loop con van doc duoc field cua loop cha (context stack:
"!   dong trong cung truoc, roi ra ngoai, cuoi cung la root)
"! - {{*mergedown:flights}} : ss:MergeDown = lines(flights)-1 -> o cap cha
"!   tu dong keo doc het so dong con (khong can biet truoc so dong)
"! - {{sum:flights.occupied}} tinh trong pham vi dong cha dang lap
"! - grouping rows: attribute x:RowLevel cua template duoc giu khi clone dong
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX11_MULTILEVEL`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex11_multilevel DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_flight,
             fldate   TYPE d,
             capacity TYPE i,
             occupied TYPE i,
           END OF ty_flight,
           ty_flights TYPE STANDARD TABLE OF ty_flight WITH EMPTY KEY,
           BEGIN OF ty_conn,
             connid  TYPE string,
             carrier TYPE string,
             flights TYPE ty_flights,
           END OF ty_conn,
           ty_conns TYPE STANDARD TABLE OF ty_conn WITH EMPTY KEY,
           BEGIN OF ty_route,
             cityfrom TYPE string,
             cityto   TYPE string,
             conns    TYPE ty_conns,
           END OF ty_route,
           ty_routes TYPE STANDARD TABLE OF ty_route WITH EMPTY KEY.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING it_routes      TYPE ty_routes
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX11_MULTILEVEL IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="lv1"><Font ss:Bold="1"/><Alignment ss:Vertical="Center"/></Style>` &&
      `<Style ss:ID="lv2"><Font ss:Italic="1"/></Style>` &&
      `<Style ss:ID="dt"><NumberFormat ss:Format="dd/mm/yyyy"/></Style>` &&
      `<Style ss:ID="sub"><Font ss:Bold="1"/><Borders>` &&
      `<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Flights">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="130"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="70"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="70"/>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Route</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Connection</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Date</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Capacity</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Occupied</Data></Cell></Row>` &&
      " cap 1: route
      `<Row><Cell ss:StyleID="lv1" ss:MergeAcross="4">` &&
      `<Data ss:Type="String">{{#routes}}{{cityfrom}} - {{cityto}}</Data></Cell></Row>` &&
      " cap 2: connection - o cot A gop doc het so dong flight cua chinh no
      `<Row><Cell ss:StyleID="lv2">` &&
      `<Data ss:Type="String">{{#conns}}{{*mergedown:flights}}{{carrier}} {{connid}}</Data></Cell>` &&
      " cap 3: flight
      `<Cell ss:StyleID="lv2"><Data ss:Type="String">{{#flights}}{{connid}}</Data></Cell>` &&
      `<Cell ss:StyleID="dt"><Data ss:Type="String">{{fldate}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{capacity}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{occupied}}{{/flights}}</Data></Cell></Row>` &&
      " subtotal cap 2
      `<Row><Cell ss:StyleID="sub" ss:MergeAcross="2">` &&
      `<Data ss:Type="String">Subtotal {{connid}} ({{cnt:flights}} flights)</Data></Cell>` &&
      `<Cell ss:StyleID="sub"><Data ss:Type="String">{{sum:flights.capacity}}</Data></Cell>` &&
      `<Cell ss:StyleID="sub"><Data ss:Type="String">{{sum:flights.occupied}}{{/conns}}</Data></Cell></Row>` &&
      " subtotal cap 1
      `<Row><Cell ss:StyleID="sub" ss:MergeAcross="4">` &&
      `<Data ss:Type="String">== Route {{cityfrom}} - {{cityto}}: {{cnt:conns}} connections{{/routes}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            routes TYPE ty_routes,
          END OF ls_context.
    ls_context-routes = it_routes.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX11_MULTILEVEL`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
