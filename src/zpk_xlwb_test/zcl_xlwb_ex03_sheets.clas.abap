"! <p class="shorttext synchronized" lang="en">XLWB ex 4.03: 1 label / 1 worksheet</p>
"!
"! Moi nhan van chuyen tren 1 worksheet rieng (sheet dong).
"! Tuong duong XLWB 4.03 (va 4.15a - moi department 1 sheet).
"! Bai hoc: SHEET LOOP - ten worksheet chua {{#labels}}, phan con lai cua
"! ten la template va phai render ra ten DUY NHAT (dung {{@index}} hoac
"! 1 field khoa cua dong).
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX03_SHEETS`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex03_sheets DEFINITION
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



CLASS ZCL_XLWB_EX03_SHEETS IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Worksheet ss:Name="{{#labels}}Label {{@index}}">` &&
      `<Table>` &&
      `<Row><Cell><Data ss:Type="String">TO: {{name}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{city}}</Data></Cell></Row>` &&
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
                iv_form_name         = `XLWB_EX03_SHEETS`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
