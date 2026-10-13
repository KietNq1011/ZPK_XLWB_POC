"! <p class="shorttext synchronized" lang="en">XLWB ex 4.13: Gantt chart (dynamic columns + spanning header)</p>
"!
"! Bang Gantt: cot tinh (Phase / Task / Duration) + so cot NGAY dong theo
"! do dai thang; header 3 cap (Month gop het cac ngay, Week gop cac ngay
"! trong tuan, Day). Tuong duong XLWB 4.13 (Resizable pattern + row/col span).
"! Bai hoc:
"! - {{*mergeacross:days}} : gop NGANG dong theo so ngay -> header cap Month
"! - cell loop long trong header: {{#weeks>}} cho cap tuan, {{#days>}} cho ngay
"! - o than bang: cell loop {{#cells>}} trong row loop {{#tasks}} -> ma tran
"! - danh dau o co viec: ABAP quyet dinh (field mark) - engine chi bind gia tri;
"!   mau to nen theo dieu kien thi dung 2 cell + block {{?..}}/{{^..}}
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX13_GANTT`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex13_gantt DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_day,
             label TYPE string,
           END OF ty_day,
           ty_days TYPE STANDARD TABLE OF ty_day WITH EMPTY KEY,
           BEGIN OF ty_week,
             label TYPE string,
             days  TYPE ty_days,
           END OF ty_week,
           ty_weeks TYPE STANDARD TABLE OF ty_week WITH EMPTY KEY,
           BEGIN OF ty_cell,
             mark TYPE string,
           END OF ty_cell,
           ty_cells TYPE STANDARD TABLE OF ty_cell WITH EMPTY KEY,
           BEGIN OF ty_task,
             phase    TYPE string,
             task     TYPE string,
             duration TYPE i,
             cells    TYPE ty_cells,
           END OF ty_task,
           ty_tasks TYPE STANDARD TABLE OF ty_task WITH EMPTY KEY,
           BEGIN OF ty_gantt,
             monthname TYPE string,
             days      TYPE ty_days,
             weeks     TYPE ty_weeks,
             tasks     TYPE ty_tasks,
           END OF ty_gantt.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING is_gantt       TYPE ty_gantt
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX13_GANTT IMPLEMENTATION.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/>` &&
      `<Alignment ss:Horizontal="Center"/></Style>` &&
      `<Style ss:ID="wk"><Font ss:Bold="1" ss:Size="9"/><Interior ss:Color="#EDF3FA" ss:Pattern="Solid"/>` &&
      `<Alignment ss:Horizontal="Center"/></Style>` &&
      `<Style ss:ID="dy"><Font ss:Size="8"/><Alignment ss:Horizontal="Center"/></Style>` &&
      `<Style ss:ID="bar"><Interior ss:Color="#70AD47" ss:Pattern="Solid"/>` &&
      `<Alignment ss:Horizontal="Center"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Gantt">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="150"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="55"/>` &&
      " ---- header cap 1: Month gop ngang het so cot ngay -------------------
      `<Row><Cell ss:StyleID="th" ss:MergeAcross="2"><Data ss:Type="String">Project plan</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">{{*mergeacross:days}}{{monthname}}</Data></Cell></Row>` &&
      " ---- header cap 2: Week, moi tuan gop ngang so ngay cua no -----------
      `<Row><Cell ss:StyleID="th" ss:MergeAcross="2"><Data ss:Type="String"></Data></Cell>` &&
      `<Cell ss:StyleID="wk"><Data ss:Type="String">{{#weeks>}}{{*mergeacross:days}}{{label}}</Data></Cell></Row>` &&
      " ---- header cap 3: Day ------------------------------------------------
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Phase</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Task</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Days</Data></Cell>` &&
      `<Cell ss:StyleID="dy"><Data ss:Type="String">{{#days>}}{{label}}</Data></Cell></Row>` &&
      " ---- than bang: row loop + cell loop ---------------------------------
      `<Row><Cell><Data ss:Type="String">{{#tasks}}{{phase}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{task}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{duration}}</Data></Cell>` &&
      `<Cell ss:StyleID="bar"><Data ss:Type="String">{{#cells>}}{{mark}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{/tasks}}</Data></Cell></Row>` &&
      `</Table>` &&
      `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<FreezePanes/><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane>` &&
      `</WorksheetOptions>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_gantt.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX13_GANTT`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
