"! <p class="shorttext synchronized" lang="en">XLWB ex 4.08: row grouping (outline +/-)</p>
"!
"! Danh sach nhan vien theo vi tri — dong chi tiet GAP DUOC bang nut +/-
"! ben le trai Excel (outline), dung nhu vi du 4.08 cua tool goc.
"! Bai hoc:
"! - {{*group=N}}: dong thuoc nhom outline cap N (1..7); XML 2003 KHONG
"!   chua outline nen marker nay duoc converter ssml2xlsx chuyen thanh
"!   outlineLevel cua OOXML (summary row nam TREN nhom chi tiet)
"! - dong summary (vi tri) de binh thuong, cac dong chi tiet trong loop
"!   con mang marker -> Excel tu sinh nut +/- theo tung nhom
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX08G_GROUPING`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex08g_grouping DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_emp,
             name TYPE string,
             role TYPE string,
           END OF ty_emp,
           ty_emps TYPE STANDARD TABLE OF ty_emp WITH EMPTY KEY,
           BEGIN OF ty_position,
             position TYPE string,
             emps     TYPE ty_emps,
           END OF ty_position,
           ty_positions TYPE STANDARD TABLE OF ty_position WITH EMPTY KEY.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING it_positions   TYPE ty_positions
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX08G_GROUPING IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="ti"><Font ss:Bold="1" ss:Size="14"/></Style>` &&
      `<Style ss:ID="pos"><Font ss:Bold="1"/>` &&
      `<Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/>` &&
      `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Employees">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="180"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="140"/>` &&
      `<Row><Cell ss:StyleID="ti" ss:MergeAcross="1">` &&
      `<Data ss:Type="String">EMPLOYEE LIST</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="pos" ss:MergeAcross="1">` &&
      `<Data ss:Type="String">{{#positions}}{{position}} ({{cnt:emps}} employees)</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#emps}}{{*group=1}}    {{name}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{role}}{{/emps}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{/positions}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            positions TYPE ty_positions,
          END OF ls_context.
    ls_context-positions = it_positions.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX08G_GROUPING`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
