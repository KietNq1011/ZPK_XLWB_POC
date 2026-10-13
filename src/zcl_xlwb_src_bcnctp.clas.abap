"! <p class="shorttext synchronized" lang="en">XLWB source: Bao cao nhu cau thanh pham (BCNCTP)</p>
"!
"! Xuất Excel app "Quản Lý Đơn Hàng" (ZCS_BCNCTP) theo dạng file
"! WeekDailyProductivity: 6 cột tĩnh (PH3/PH4/Plant + tên) và CẶP CỘT ĐỘNG
"! theo tuần — "W{w}/{y}" (Order qty) + "CX W{w}/{y}" (Unconfirmed).
"! Dữ liệu THẬT lấy qua zcl_get_bcnctp=>get_bcnctp (đúng nguồn của app).
"!
"! Gọi production:
"!   zcl_xlwb_runtime=>render_by_source(
"!     iv_form_name = 'XLWB_BCNCTP_WEEKLY'
"!     iv_keys      = '{ "week": "35", "year": "2026", "weeks_count": 18 }' ).
"! Keys hỗ trợ: week*, year* (bắt buộc), plant, ph3, weeks_count (mặc định 18,
"! tối đa 54 — số cặp cột tuần).
CLASS zcl_xlwb_src_bcnctp DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES zif_xlwb_source.

    TYPES: BEGIN OF ty_week,
             label TYPE string,
           END OF ty_week,
           ty_weeks TYPE STANDARD TABLE OF ty_week WITH EMPTY KEY,
           BEGIN OF ty_cell,
             val TYPE p LENGTH 15 DECIMALS 0,
           END OF ty_cell,
           ty_cells TYPE STANDARD TABLE OF ty_cell WITH EMPTY KEY,
           BEGIN OF ty_row,
             ph3        TYPE string,
             ph3_name   TYPE string,
             ph4        TYPE string,
             ph4_name   TYPE string,
             plant      TYPE string,
             plant_name TYPE string,
             cells      TYPE ty_cells,
           END OF ty_row,
           ty_rows TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY,
           BEGIN OF ty_ctx,
             title   TYPE string,
             week    TYPE string,
             year    TYPE string,
             lastcol TYPE i,          " cột cuối = 6 + số nhãn tuần (cho AutoFilter)
             weeks   TYPE ty_weeks,
             rows    TYPE ty_rows,
           END OF ty_ctx.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.

    "! File demo cho catalog (dữ liệu mẫu tĩnh, không đụng DB)
    METHODS get_file
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    METHODS build_ctx
      IMPORTING it_data       TYPE zcl_get_bcnctp=>gty_bcnctp
                iv_week       TYPE i
                iv_year       TYPE i
                iv_count      TYPE i
      RETURNING VALUE(rs_ctx) TYPE ty_ctx.
    METHODS sample_ctx
      RETURNING VALUE(rs_ctx) TYPE ty_ctx.
    METHODS key_of
      IMPORTING ir_keys         TYPE REF TO data
                iv_comp         TYPE string
      RETURNING VALUE(rv_value) TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_SRC_BCNCTP IMPLEMENTATION.


  METHOD zif_xlwb_source~get_context.
    IF iv_keys IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = |BCNCTP cần keys JSON, tối thiểu \{ "week": "35", "year": "2026" \}| ).
    ENDIF.

    DATA(lr_keys) = zcl_xlwb_ctx_json=>parse( iv_keys ).
    DATA(lv_week) = key_of( ir_keys = lr_keys iv_comp = 'WEEK' ).
    DATA(lv_year) = key_of( ir_keys = lr_keys iv_comp = 'YEAR' ).
    IF lv_week IS INITIAL OR lv_year IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Thiếu "week" hoặc "year" trong keys| ).
    ENDIF.

    DATA(lv_week_n) = CONV i( lv_week ).
    DATA(lv_year_n) = CONV i( lv_year ).
    DATA(lv_count)  = COND i( LET c = key_of( ir_keys = lr_keys iv_comp = 'WEEKS_COUNT' ) IN
                              WHEN c IS NOT INITIAL THEN CONV i( c ) ELSE 18 ).
    IF lv_count < 1.  lv_count = 1.  ENDIF.
    IF lv_count > 54. lv_count = 54. ENDIF.

    DATA(lv_week_c) = |{ lv_week_n WIDTH = 2 PAD = '0' ALIGN = RIGHT }|.
    DATA(lt_r_week)  = VALUE zcl_get_bcnctp=>tt_ranges( ( sign = 'I' option = 'EQ' low = lv_week_c ) ).
    DATA(lt_r_year)  = VALUE zcl_get_bcnctp=>tt_ranges( ( sign = 'I' option = 'EQ' low = lv_year ) ).
    DATA lt_r_plant TYPE zcl_get_bcnctp=>tt_ranges.
    DATA(lv_plant) = key_of( ir_keys = lr_keys iv_comp = 'PLANT' ).
    IF lv_plant IS NOT INITIAL.
      lt_r_plant = VALUE #( ( sign = 'I' option = 'EQ' low = lv_plant ) ).
    ENDIF.
    DATA lt_r_ph3 TYPE zcl_get_bcnctp=>tt_ranges.
    DATA(lv_ph3) = key_of( ir_keys = lr_keys iv_comp = 'PH3' ).
    IF lv_ph3 IS NOT INITIAL.
      lt_r_ph3 = VALUE #( ( sign = 'I' option = 'EQ' low = lv_ph3 ) ).
    ENDIF.

    zcl_get_bcnctp=>get_bcnctp(
      EXPORTING ir_hierarchy3 = lt_r_ph3
                ir_plant      = lt_r_plant
                ir_week       = lt_r_week
                ir_year       = lt_r_year
      IMPORTING et_data       = DATA(lt_data) ).

    DATA lr_ctx TYPE REF TO ty_ctx.
    CREATE DATA lr_ctx.
    lr_ctx->* = build_ctx( it_data = lt_data
                           iv_week = lv_week_n
                           iv_year = lv_year_n
                           iv_count = lv_count ).
    rr_context = lr_ctx.
  ENDMETHOD.


  METHOD zif_xlwb_source~get_sample.
    DATA lr_ctx TYPE REF TO ty_ctx.
    CREATE DATA lr_ctx.
    lr_ctx->* = sample_ctx( ).
    rr_context = lr_ctx.
  ENDMETHOD.


  METHOD key_of.
    ASSIGN ir_keys->* TO FIELD-SYMBOL(<ls_keys>).
    ASSIGN COMPONENT iv_comp OF STRUCTURE <ls_keys> TO FIELD-SYMBOL(<lv_val>).
    IF sy-subrc = 0.
      rv_value = condense( |{ <lv_val> }| ).
    ENDIF.
  ENDMETHOD.


  METHOD build_ctx.
    rs_ctx-week  = |{ iv_week }|.
    rs_ctx-year  = |{ iv_year }|.
    rs_ctx-title = |BÁO CÁO NHU CẦU THÀNH PHẨM — TỪ TUẦN W{ iv_week }/{ iv_year }|.

    " nhãn CẶP cột tuần, xen kẽ Qty / CX (chưa xác nhận) — cột W1..Wn của
    " ZCS_BCNCTP tương ứng tuần bắt đầu từ tuần được chọn
    DATA(lv_w) = iv_week.
    DATA(lv_y) = iv_year.
    DO iv_count TIMES.
      APPEND VALUE ty_week( label = |W{ lv_w }/{ lv_y }| )    TO rs_ctx-weeks.
      APPEND VALUE ty_week( label = |CX W{ lv_w }/{ lv_y }| ) TO rs_ctx-weeks.
      lv_w += 1.
      IF lv_w > 52.
        lv_w = 1.
        lv_y += 1.
      ENDIF.
    ENDDO.

    rs_ctx-lastcol = 6 + lines( rs_ctx-weeks ).

    LOOP AT it_data INTO DATA(ls_data).
      " dạng TREE như app: dòng cha (PH4 rỗng) hiện PH3; dòng con chỉ hiện
      " PH4, cột PH3 để trống
      DATA(lv_child) = xsdbool( ls_data-producthierarchy4 IS NOT INITIAL ).
      DATA(ls_row) = VALUE ty_row(
        ph3        = COND #( WHEN lv_child = abap_false THEN ls_data-producthierarchy3 )
        ph3_name   = COND #( WHEN lv_child = abap_false THEN ls_data-producthierarchy3name )
        ph4        = ls_data-producthierarchy4
        ph4_name   = COND #( WHEN lv_child = abap_true THEN ls_data-producthierarchy4name )
        plant      = ls_data-plant
        plant_name = ls_data-plantname ).

      DO iv_count TIMES.
        DATA(lv_ix) = sy-index.
        ASSIGN COMPONENT |W{ lv_ix }ORDERQUANTITY| OF STRUCTURE ls_data
          TO FIELD-SYMBOL(<lv_qty>).
        ASSIGN COMPONENT |W{ lv_ix }UNCONFIRMEDQUANTITY| OF STRUCTURE ls_data
          TO FIELD-SYMBOL(<lv_cx>).
        APPEND VALUE ty_cell( val = COND #( WHEN <lv_qty> IS ASSIGNED THEN <lv_qty> ) )
          TO ls_row-cells.
        APPEND VALUE ty_cell( val = COND #( WHEN <lv_cx> IS ASSIGNED THEN <lv_cx> ) )
          TO ls_row-cells.
        UNASSIGN: <lv_qty>, <lv_cx>.
      ENDDO.

      APPEND ls_row TO rs_ctx-rows.
    ENDLOOP.
  ENDMETHOD.


  METHOD sample_ctx.
    " dữ liệu demo tĩnh theo đúng hình app (client test không có data thật)
    DATA lt_data TYPE zcl_get_bcnctp=>gty_bcnctp.
    lt_data = VALUE #(
      ( producthierarchy3 = 'M50304' producthierarchy3name = 'Not use'
        plant = '6711' plantname = 'Nhà máy túi CASLA 1'
        w1orderquantity = 65123 w1unconfirmedquantity = 0 )
      ( producthierarchy3 = 'M50304' producthierarchy3name = 'Not use'
        producthierarchy4 = 'M50304002' producthierarchy4name = 'Not use'
        plant = '6711' plantname = 'Nhà máy túi CASLA 1'
        w1orderquantity = 65123 w1unconfirmedquantity = 0 )
      ( producthierarchy3 = 'M50602' producthierarchy3name = 'Túi dán Tshirt'
        plant = '6711' plantname = 'Nhà máy túi CASLA 1'
        w1orderquantity = 77100 w2orderquantity = 1200 w2unconfirmedquantity = 300 )
      ( producthierarchy3 = 'M50602' producthierarchy3name = 'Túi dán Tshirt'
        producthierarchy4 = 'M50602014' producthierarchy4name = 'Not use'
        plant = '6711' plantname = 'Nhà máy túi CASLA 1'
        w1orderquantity = 77100 w2orderquantity = 1200 w2unconfirmedquantity = 300 ) ).

    rs_ctx = build_ctx( it_data = lt_data iv_week = 35 iv_year = 2026 iv_count = 4 ).
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
      `<Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>` &&
      `<Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>` &&
      `<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `<Style ss:ID="ph3"><Font ss:Bold="1"/></Style>` &&
      `<Style ss:ID="num"><NumberFormat ss:Format="#,##0"/>` &&
      `<Borders><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="BCNCTP">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="80"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="110"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="80"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="110"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="45"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="120"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="75"/>` &&
      `<Row ss:Height="20"><Cell ss:StyleID="ti" ss:MergeAcross="6">` &&
      `<Data ss:Type="String">{{title}}</Data></Cell></Row>` &&
      `<Row/>` &&
      `<Row ss:Height="28">` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Product Hierarchy 3</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">PH3 Name</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Product Hierarchy 4</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">PH4 Name</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Plant</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Plant Name</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">{{#weeks>}}{{label}}</Data></Cell>` &&
      `</Row>` &&
      `<Row>` &&
      `<Cell ss:StyleID="ph3"><Data ss:Type="String">{{#rows}}{{ph3}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{ph3_name}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{ph4}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{ph4_name}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{plant}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{plant_name}}</Data></Cell>` &&
      `<Cell ss:StyleID="num"><Data ss:Type="String">{{#cells>}}{{val}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{/rows}}</Data></Cell>` &&
      `</Row>` &&
      `</Table>` &&
      `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<FreezePanes/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane>` &&
      `<SplitVertical>6</SplitVertical><LeftColumnRightPane>6</LeftColumnRightPane>` &&
      `</WorksheetOptions>` &&
      `<AutoFilter x:Range="R3C1:R3C{{lastcol}}"` &&
      ` xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_ctx) = sample_ctx( ).
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_BCNCTP_WEEKLY`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_ctx ) )-content.
  ENDMETHOD.
ENDCLASS.
