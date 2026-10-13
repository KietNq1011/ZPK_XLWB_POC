"! <p class="shorttext synchronized" lang="en">XLWB ex: source class demo (ZIF_XLWB_SOURCE)</p>
"!
"! Vi du GAN SOURCE cho template — thay the man hinh gan context cua tool
"! XLSX Workbench goc:
"! 1. Class nay implement ZIF_XLWB_SOURCE (get_context = du lieu that,
"!    o day la demo; production thi SELECT theo iv_keys).
"! 2. Trong app template: gan SOURCE_CLASS = ZCL_XLWB_SRC_ORDER cho form
"!    XLWB_EX22_SOURCE, dien SAMPLE_JSON roi bam Preview.
"! 3. Code nghiep vu chi can:
"!    zcl_xlwb_runtime=>render_by_source( iv_form_name = 'XLWB_EX22_SOURCE'
"!                                        iv_keys = '{ "orderno": "PO-1" }' ).
CLASS zcl_xlwb_src_order DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES zif_xlwb_source.

    TYPES: BEGIN OF ty_item,
             pos    TYPE i,
             matnr  TYPE string,
             qty    TYPE p LENGTH 8 DECIMALS 0,
             amount TYPE p LENGTH 11 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_order,
             orderno  TYPE string,
             customer TYPE string,
             orddate  TYPE d,
             items    TYPE ty_items,
           END OF ty_order.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    METHODS build_order
      IMPORTING iv_orderno      TYPE string
      RETURNING VALUE(rs_order) TYPE ty_order.
ENDCLASS.



CLASS ZCL_XLWB_SRC_ORDER IMPLEMENTATION.


  METHOD zif_xlwb_source~get_context.
    DATA(lv_orderno) = `PO-2026-DEMO`.

    IF iv_keys IS NOT INITIAL.
      TRY.
          DATA(lr_keys) = zcl_xlwb_ctx_json=>parse( iv_keys ).
          ASSIGN lr_keys->* TO FIELD-SYMBOL(<ls_keys>).
          ASSIGN COMPONENT 'ORDERNO' OF STRUCTURE <ls_keys> TO FIELD-SYMBOL(<lv_no>).
          IF sy-subrc = 0 AND <lv_no> IS NOT INITIAL.
            lv_orderno = |{ <lv_no> }|.
          ENDIF.
        CATCH zcx_xlwb INTO DATA(lx).
          RAISE EXCEPTION NEW zcx_xlwb(
            iv_text = |iv_keys không phải JSON hợp lệ: { lx->get_text( ) }| ).
      ENDTRY.
    ENDIF.

    " Production: thay build_order bằng SELECT theo lv_orderno
    DATA lr_order TYPE REF TO ty_order.
    CREATE DATA lr_order.
    lr_order->* = build_order( lv_orderno ).
    rr_context = lr_order.
  ENDMETHOD.


  METHOD zif_xlwb_source~get_sample.
    DATA lr_order TYPE REF TO ty_order.
    CREATE DATA lr_order.
    lr_order->* = build_order( `PO-2026-DEMO` ).
    rr_context = lr_order.
  ENDMETHOD.


  METHOD build_order.
    rs_order = VALUE ty_order(
      orderno  = iv_orderno
      customer = `ACME Corp`
      orddate  = cl_abap_context_info=>get_system_date( )
      items    = VALUE #(
        ( pos = 10 matnr = `SLAB-WHITE-3CM` qty = 20 amount = '2100.00' )
        ( pos = 20 matnr = `SLAB-GREY-2CM`  qty = 15 amount = '1327.50' ) ) ).
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
      `<Style ss:ID="th"><Font ss:Bold="1"/><Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="tot"><Font ss:Bold="1"/>` &&
      `<Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Order">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="40"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="150"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="60"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="90"/>` &&
      `<Row><Cell ss:StyleID="ti" ss:MergeAcross="3">` &&
      `<Data ss:Type="String">ORDER {{orderno}} — {{customer}} ({{orddate}})</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="th"><Data ss:Type="String">Pos</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Material</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Qty</Data></Cell>` &&
      `<Cell ss:StyleID="th"><Data ss:Type="String">Amount</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#items}}{{pos}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{matnr}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{qty}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{amount}}{{/items}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="tot" ss:MergeAcross="1"><Data ss:Type="String">TOTAL</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.qty}}</Data></Cell>` &&
      `<Cell ss:StyleID="tot"><Data ss:Type="String">{{sum:items.amount}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX22_SOURCE`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = zif_xlwb_source~get_sample( ) )-content.
  ENDMETHOD.
ENDCLASS.
