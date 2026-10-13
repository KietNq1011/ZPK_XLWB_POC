"! <p class="shorttext synchronized" lang="en">XLWB ex 4.08a/4.09: Dynamic columns</p>
"!
"! Bang co so COT quyet dinh luc runtime (ma tran): moi dong context chua
"! 1 bang con cac o. Tuong duong XLWB 4.09 (dynamic table) + 4.08a
"! (an/hien cot: chi dua vao bang cot nhung cot can hien).
"! Bai hoc:
"! - {{#cols>}} : CELL LOOP - cell chua marker duoc nhan ban NGANG
"!   moi dong cua bang; style cua cell mau di theo tung ban sao
"! - ket hop row loop ({{#rows}}) + cell loop ({{#cells>}}) = ma tran
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX08_DYNCOLS`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex08_dyncols DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_cell,
             val TYPE p LENGTH 8 DECIMALS 2,
           END OF ty_cell,
           ty_cells TYPE STANDARD TABLE OF ty_cell WITH EMPTY KEY,
           BEGIN OF ty_col,
             title TYPE string,
           END OF ty_col,
           ty_cols TYPE STANDARD TABLE OF ty_col WITH EMPTY KEY,
           BEGIN OF ty_row,
             rowname TYPE string,
             cells   TYPE ty_cells,
           END OF ty_row,
           ty_rows TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY,
           BEGIN OF ty_matrix,
             cols TYPE ty_cols,
             rows TYPE ty_rows,
           END OF ty_matrix.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING is_matrix      TYPE ty_matrix
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX08_DYNCOLS IMPLEMENTATION.


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
      `</Styles>` &&
      `<Worksheet ss:Name="Matrix">` &&
      `<Table>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Product</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">{{#cols>}}{{title}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#rows}}{{rowname}}</Data></Cell>` &&
      `<Cell ss:StyleID="num"><Data ss:Type="String">{{#cells>}}{{val}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{/rows}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_matrix.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX08_DYNCOLS`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
