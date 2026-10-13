"! <p class="shorttext synchronized" lang="en">XLWB ex 4.01: Shipping label</p>
"!
"! 1 nhan van chuyen: phan FROM tinh (nam san trong template), phan TO dong.
"! Tuong duong XLWB 4.01. Bai hoc:
"! - path long nhau {{to.name}} (structure trong structure)
"! - style/merge/do rong cot/chieu cao dong lay tu TEMPLATE, engine khong dung den
"! - cell nguyen placeholder kieu ngay -> ss:Type DateTime, numfmt cua template ap dung
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX01_LABEL`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex01_label DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_party,
             name   TYPE string,
             street TYPE string,
             city   TYPE string,
           END OF ty_party,
           BEGIN OF ty_label,
             to       TYPE ty_party,
             shipdate TYPE d,
             weight   TYPE p LENGTH 8 DECIMALS 2,
           END OF ty_label.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING is_label       TYPE ty_label
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX01_LABEL IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="title"><Font ss:Bold="1" ss:Size="14"/>` &&
      `<Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="lbl"><Font ss:Bold="1"/></Style>` &&
      `<Style ss:ID="dt"><NumberFormat ss:Format="dd/mm/yyyy"/></Style>` &&
      `<Style ss:ID="kg"><NumberFormat ss:Format="#,##0.00 &quot;kg&quot;"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Label">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="200"/>` &&
      `<Row ss:AutoFitHeight="0" ss:Height="24">` &&
      `<Cell ss:StyleID="title" ss:MergeAcross="1">` &&
      `<Data ss:Type="String">SHIPPING LABEL</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">FROM:</Data></Cell>` &&
      `<Cell><Data ss:Type="String">CASLA QUARTZ JSC, Ha Noi</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">TO:</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{to.name}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String"></Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{to.street}}, {{to.city}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">Ship date:</Data></Cell>` &&
      `<Cell ss:StyleID="dt"><Data ss:Type="String">{{shipdate}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">Weight:</Data></Cell>` &&
      `<Cell ss:StyleID="kg"><Data ss:Type="String">{{weight}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_label.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX01_LABEL`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
