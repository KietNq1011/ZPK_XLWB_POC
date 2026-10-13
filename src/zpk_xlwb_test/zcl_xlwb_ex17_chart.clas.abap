"! <p class="shorttext synchronized" lang="en">XLWB ex 4.17/4.18: chart trong template</p>
"!
"! Chart/dashboard: engine KHONG tao chart luc runtime; chart duoc VE SAN
"! trong template .xlsx (Excel > Insert > Chart tren vung du lieu) va engine
"! XLSX GIU NGUYEN khi bind du lieu vao vung do.
"! Quy trinh: ve bang du lieu + chart trong Excel -> Save As .xlsx ->
"! upload vao ZXLWB_TMPL voi Engine = XLSX (form XLWB_EX17_CHART).
"! Ban du phong trong code chi la bang du lieu (SSML) de vi du chay duoc
"! khi chua upload template that.
CLASS zcl_xlwb_ex17_chart DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_month,
             name    TYPE string,
             revenue TYPE p LENGTH 11 DECIMALS 2,
           END OF ty_month,
           ty_months TYPE STANDARD TABLE OF ty_month WITH EMPTY KEY.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING it_months      TYPE ty_months
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX17_CHART IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="num"><NumberFormat ss:Format="#,##0.00"/></Style>` &&
      `<Style ss:ID="note"><Font ss:Italic="1" ss:Color="#808080"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Data">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="110"/>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Month</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Revenue</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#months}}{{name}}</Data></Cell>` &&
      `<Cell ss:StyleID="num"><Data ss:Type="String">{{revenue}}{{/months}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="note" ss:MergeAcross="1"><Data ss:Type="String">` &&
      `Chart: ve trong template .xlsx (Insert - Chart tren vung du lieu) roi upload ` &&
      `voi Engine = XLSX — engine giu nguyen chart khi render.</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            months TYPE ty_months,
          END OF ls_context.
    ls_context-months = it_months.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX17_CHART`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
