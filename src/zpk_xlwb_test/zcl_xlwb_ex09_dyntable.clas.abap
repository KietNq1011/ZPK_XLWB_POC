"! <p class="shorttext synchronized" lang="en">XLWB ex 4.09: dynamic table (rows x cols runtime)</p>
"!
"! Bang co SO DONG va SO COT chi biet luc runtime — dung nhu vi du 4.09
"! cua tool goc (DYNTABLE, demo bang so ngau nhien: moi lan chay ra kich
"! thuoc khac nhau). Bai hoc:
"! - context bang long bang: rows -> cells (kich thuoc tuy y)
"! - row loop x cell loop = ma tran; header cot cung la cell loop
"! - tong theo DONG bang {{sum:cells.val}} trong frame cua dong
"! - tong theo COT: tinh san o ABAP thanh bang coltotals (aggregate cua
"!   engine tinh theo frame dong, khong tinh doc theo cot duoc)
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX09_DYNTABLE`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex09_dyntable DEFINITION
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
           ty_rows TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.

    "! Render bảng động. Không truyền kích thước -> sinh NGẪU NHIÊN
    "! (dòng 3..8, cột 3..7, giá trị 1..100) như ví dụ 4.09 gốc.
    METHODS get_file
      IMPORTING iv_rows        TYPE i DEFAULT 0
                iv_cols        TYPE i DEFAULT 0
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    DATA mv_seed TYPE int8.
    METHODS rnd
      IMPORTING iv_min        TYPE i
                iv_max        TYPE i
      RETURNING VALUE(rv_val) TYPE i.
ENDCLASS.



CLASS ZCL_XLWB_EX09_DYNTABLE IMPLEMENTATION.


  METHOD rnd.
    " LCG don gian, seed tu timestamp — du "ngau nhien" cho demo,
    " khong phu thuoc API random nao
    IF mv_seed = 0.
      GET TIME STAMP FIELD DATA(lv_ts).
      mv_seed = CONV int8( lv_ts ) MOD 2147483647.
      IF mv_seed <= 0.
        mv_seed = 42.
      ENDIF.
    ENDIF.
    mv_seed = ( mv_seed * 1103515245 + 12345 ) MOD 2147483648.
    rv_val = iv_min + CONV i( mv_seed MOD CONV int8( iv_max - iv_min + 1 ) ).
  ENDMETHOD.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="ti"><Font ss:Bold="1" ss:Size="14"/></Style>` &&
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/>` &&
      `<Alignment ss:Horizontal="Center"/>` &&
      `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `<Style ss:ID="rh"><Font ss:Bold="1"/></Style>` &&
      `<Style ss:ID="num"><NumberFormat ss:Format="#,##0.00"/></Style>` &&
      `<Style ss:ID="tot"><Font ss:Bold="1"/><NumberFormat ss:Format="#,##0.00"/>` &&
      `<Interior ss:Color="#FCE4D6" ss:Pattern="Solid"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="DynTable">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="70"/>` &&
      `<Row><Cell ss:StyleID="ti" ss:MergeAcross="3">` &&
      `<Data ss:Type="String">DYNAMIC TABLE {{rows_n}} x {{cols_n}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Row \ Col</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">{{#cols>}}{{title}}</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Row total</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="rh"><Data ss:Type="String">{{#rows}}{{rowname}}</Data></Cell>` &&
      `<Cell ss:StyleID="num"><Data ss:Type="String">{{#cells>}}{{val}}</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:cells.val}}{{/rows}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="rh"><Data ss:Type="String">Col total</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{#coltotals>}}{{val}}</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{grand}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            rows_n    TYPE i,
            cols_n    TYPE i,
            cols      TYPE ty_cols,
            rows      TYPE ty_rows,
            coltotals TYPE ty_cells,
            grand     TYPE p LENGTH 11 DECIMALS 2,
          END OF ls_context.

    ls_context-rows_n = COND #( WHEN iv_rows > 0 THEN iv_rows ELSE rnd( iv_min = 3 iv_max = 8 ) ).
    ls_context-cols_n = COND #( WHEN iv_cols > 0 THEN iv_cols ELSE rnd( iv_min = 3 iv_max = 7 ) ).

    DO ls_context-cols_n TIMES.
      APPEND VALUE ty_col( title = |Col { sy-index }| ) TO ls_context-cols.
      APPEND VALUE ty_cell( ) TO ls_context-coltotals.
    ENDDO.

    DATA lv_val TYPE p LENGTH 8 DECIMALS 2.
    DO ls_context-rows_n TIMES.
      DATA(ls_row) = VALUE ty_row( rowname = |Row { sy-index }| ).
      DO ls_context-cols_n TIMES.
        lv_val = rnd( iv_min = 1 iv_max = 100 ).
        APPEND VALUE ty_cell( val = lv_val ) TO ls_row-cells.
        DATA(lv_cix) = sy-index.
        ls_context-coltotals[ lv_cix ]-val = ls_context-coltotals[ lv_cix ]-val + lv_val.
        ls_context-grand = ls_context-grand + lv_val.
      ENDDO.
      APPEND ls_row TO ls_context-rows.
    ENDDO.

    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX09_DYNTABLE`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
