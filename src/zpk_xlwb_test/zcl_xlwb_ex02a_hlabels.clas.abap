"! <p class="shorttext synchronized" lang="en">XLWB ex 4.02a: N labels NGANG (cell loop)</p>
"!
"! N nhan xep NGANG sang phai — moi nhan mot COT. Tuong duong XLWB 4.02a.
"! Bai hoc:
"! - {{#labels>}} tren NHIEU DONG: moi dong co cell loop rieng tren CUNG
"!   mot bang -> cac cot tu thang hang voi nhau
"! - style/do rong cua o mau di theo tung ban sao
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX02A_HLABELS`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex02a_hlabels DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_label,
             name TYPE string,
             city TYPE string,
           END OF ty_label,
           ty_labels TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING it_labels      TYPE ty_labels
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX02A_HLABELS IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="hd"><Font ss:Bold="1"/>` &&
      `<Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="lbl"><Font ss:Bold="1"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="HLabels">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="60"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="120"/>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">TO:</Data></Cell>` &&
      `<Cell ss:StyleID="hd"><Data ss:Type="String">{{#labels>}}Label {{@index}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">Name:</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{#labels>}}{{name}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">City:</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{#labels>}}{{city}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            labels TYPE ty_labels,
          END OF ls_context.
    ls_context-labels = it_labels.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX02A_HLABELS`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
