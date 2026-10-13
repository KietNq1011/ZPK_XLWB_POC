"! <p class="shorttext synchronized" lang="en">XLWB ex 4.04: workbook nhieu form</p>
"!
"! MOT lan render ra MOT workbook chua NHIEU form khac nhau — moi form
"! mot worksheet tinh (khac EX03 la sheet dong nhan ban theo du lieu).
"! Tuong duong XLWB 4.04 (save several forms into one workbook).
"! Bai hoc:
"! - template co nhieu <Worksheet>, moi sheet bind mot nhanh cua context
"! - context la structure GOM: phan order + phan labels
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX04_BUNDLE`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex04_bundle DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_item,
             matnr  TYPE string,
             qty    TYPE p LENGTH 8 DECIMALS 0,
             amount TYPE p LENGTH 11 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_label,
             name TYPE string,
             city TYPE string,
           END OF ty_label,
           ty_labels TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY,
           BEGIN OF ty_bundle,
             orderno  TYPE string,
             customer TYPE string,
             items    TYPE ty_items,
             labels   TYPE ty_labels,
           END OF ty_bundle.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING is_bundle      TYPE ty_bundle
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX04_BUNDLE IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="ti"><Font ss:Bold="1" ss:Size="14"/></Style>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="tot"><Font ss:Bold="1"/>` &&
      `<Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `<Style ss:ID="num"><NumberFormat ss:Format="#,##0.00"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Order">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="160"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="60"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Row><Cell ss:StyleID="ti" ss:MergeAcross="2">` &&
      `<Data ss:Type="String">ORDER {{orderno}} — {{customer}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Material</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Qty</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Amount</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#items}}{{matnr}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{qty}}</Data></Cell>` &&
      `<Cell ss:StyleID="num"><Data ss:Type="String">{{amount}}{{/items}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="tot"><Data ss:Type="String">TOTAL</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.qty}}</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.amount}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `<Worksheet ss:Name="Labels">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="220"/>` &&
      `<Row><Cell ss:StyleID="th">` &&
      `<Data ss:Type="String">{{#labels}}Label {{@index}} / {{@count}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{name}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{city}}{{/labels}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_bundle.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX04_BUNDLE`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
