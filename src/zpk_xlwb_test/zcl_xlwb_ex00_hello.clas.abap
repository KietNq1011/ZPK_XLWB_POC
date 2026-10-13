"! <p class="shorttext synchronized" lang="en">XLWB ex 4.00: Hello World</p>
"!
"! Vi du don gian nhat: 1 gia tri don duoc bind vao 1 cell.
"! Tuong duong XLWB 4.00 "Hello World".
"! Bai hoc: cau truc template SpreadsheetML toi thieu + placeholder {{path}}.
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX00_HELLO`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex00_hello DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    "! Template - trong thuc te file nay do user thiet ke bang Excel
    "! (Save As "XML Spreadsheet 2003") va luu trong bang template
    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    "! Render ra file .xls (base64/download tu action RAP)
    METHODS get_file RETURNING VALUE(rv_file) TYPE xstring RAISING zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX00_HELLO IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Worksheet ss:Name="Hello">` &&
      `<Table>` &&
      `<Row><Cell><Data ss:Type="String">{{message}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            message TYPE string,
          END OF ls_context.
    ls_context-message = 'Hello world!'.

    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX00_HELLO`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
