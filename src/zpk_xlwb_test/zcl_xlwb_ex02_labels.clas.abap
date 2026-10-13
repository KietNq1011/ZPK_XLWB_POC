"! <p class="shorttext synchronized" lang="en">XLWB ex 4.02/4.02b: N labels + page break</p>
"!
"! N nhan van chuyen noi tiep tu tren xuong, moi nhan in 1 trang rieng.
"! Tuong duong XLWB 4.02 (loop doc) + 4.02b (page break truoc moi nhan).
"! Bai hoc:
"! - vung lap NHIEU DONG: {{#labels}} o dong dau, {{/labels}} o dong cuoi
"! - {{*break}} chen page break truoc dong (nhan tu nhan thu 2 tro di
"!   nam dau trang moi; break truoc dong dau tien duoc engine tu bo qua)
"! - {{@index}}/{{@count}} danh so nhan
"! - PageSetup (kho giay, huong giay) giu nguyen tu template
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX02_LABELS`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex02_labels DEFINITION
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



CLASS ZCL_XLWB_EX02_LABELS IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="hd"><Font ss:Bold="1"/>` &&
      `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/></Borders>` &&
      `</Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Labels">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="220"/>` &&
      `<Row><Cell ss:StyleID="hd">` &&
      `<Data ss:Type="String">{{#labels}}{{*break}}Label {{@index}} / {{@count}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{name}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{city}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{/labels}}</Data></Cell></Row>` &&
      `</Table>` &&
      `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<PageSetup><Layout x:Orientation="Landscape"/></PageSetup>` &&
      `</WorksheetOptions>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            labels TYPE ty_labels,
          END OF ls_context.
    ls_context-labels = it_labels.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX02_LABELS`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
