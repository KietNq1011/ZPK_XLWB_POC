"! <p class="shorttext synchronized" lang="en">XLWB ex 4.05/4.11a: Order form</p>
"!
"! Form chung tu kinh dien: Header / bang item / Footer.
"! Tuong duong XLWB 4.05 (Pattern Header-Line-Footer) + 4.11a (subtotal
"! vua bang GIA TRI tinh o server, vua bang CONG THUC Excel).
"! Bai hoc:
"! - vung lap 1 dong: {{#items}} va {{/items}} tren CUNG 1 dong
"! - cell nguyen placeholder so -> ss:Type Number, numfmt tien te cua template ap dung
"! - {{sum:items.amount}} / {{cnt:items}}: tong do ENGINE tinh (gia tri chet)
"! - ss:Formula R1C1 giu nguyen: =SUM(...) do EXCEL tinh lai khi mo file
"!   (dong ke tren so dong dong: R[-n]C khong dung duoc - dung cong thuc
"!   quet tu dong co dinh, o day tong toan cot: SUM(R2C:R[-1]C))
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX05_ORDER`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex05_order DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_item,
             pos    TYPE i,
             matnr  TYPE string,
             qty    TYPE p LENGTH 8 DECIMALS 0,
             price  TYPE p LENGTH 11 DECIMALS 2,
             amount TYPE p LENGTH 11 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_order,
             orderno  TYPE string,
             customer TYPE string,
             orddate  TYPE d,
             items    TYPE ty_items,
             remark   TYPE string,
           END OF ty_order.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING is_order       TYPE ty_order
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX05_ORDER IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="title"><Font ss:Bold="1" ss:Size="14"/></Style>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/>` &&
      `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `<Style ss:ID="cur"><NumberFormat ss:Format="#,##0.00"/></Style>` &&
      `<Style ss:ID="tot"><Font ss:Bold="1"/><NumberFormat ss:Format="#,##0.00"/>` &&
      `<Borders><Border ss:Position="Top" ss:LineStyle="Double" ss:Weight="3"/></Borders></Style>` &&
      `<Style ss:ID="dt"><NumberFormat ss:Format="dd/mm/yyyy"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Order">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="40"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="160"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="60"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="110"/>` &&
      `<Row><Cell ss:StyleID="title" ss:MergeAcross="4">` &&
      `<Data ss:Type="String">PURCHASE ORDER {{orderno}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">Customer:</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{customer}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">Date:</Data></Cell>` &&
      `<Cell ss:StyleID="dt"><Data ss:Type="String">{{orddate}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Pos</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Material</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Qty</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Price</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Amount</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#items}}{{pos}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{matnr}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{qty}}</Data></Cell>` &&
      `<Cell ss:StyleID="cur"><Data ss:Type="String">{{price}}</Data></Cell>` &&
      `<Cell ss:StyleID="cur"><Data ss:Type="String">{{amount}}{{/items}}</Data></Cell></Row>` &&
      `<Row><Cell ss:MergeAcross="1"><Data ss:Type="String">Total ({{cnt:items}} items), engine:</Data></Cell>` &&
      `<Cell><Data ss:Type="String"></Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.qty}}</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.amount}}</Data></Cell></Row>` &&
      `<Row><Cell ss:MergeAcross="2"><Data ss:Type="String">Total by Excel formula:</Data></Cell>` &&
      `<Cell><Data ss:Type="String"></Data></Cell>` &&
      `<Cell ss:StyleID="tot" ss:Formula="=SUM(R4C:R[-2]C)"><Data ss:Type="Number">0</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">Remark: {{remark}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_order.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX05_ORDER`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
