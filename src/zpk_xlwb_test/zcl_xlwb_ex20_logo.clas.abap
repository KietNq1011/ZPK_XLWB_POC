"! <p class="shorttext synchronized" lang="en">XLWB ex 4.06/4.07: Invoice có logo (engine XLSX)</p>
"!
"! Form .xlsx thật có **logo nhúng** — thứ engine SpreadsheetML không làm được.
"! Tương đương XLWB 4.06/4.07 («Drawing» với dynamic source).
"! Bài học:
"! - logo để dạng **base64 trong context**, engine tự decode và đóng vào
"!   xl/media + drawing + rels (giống pattern ZCL_QM_EXPORT_LOGO/SERVICE)
"! - marker {{*image:path}} hoặc {{*image:path@CxR}} (C cột, R dòng)
"! - cùng cú pháp placeholder với engine SpreadsheetML: loop, aggregate...
"! - ảnh có thể lấy từ bảng Z, MIME repository hay chính field trong dòng lặp
"!   (mỗi dòng một ảnh khác nhau: đặt {{*image:photo}} trong vùng lặp)
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX20_LOGO`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex20_logo DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_item,
             matnr  TYPE string,
             amount TYPE p LENGTH 11 DECIMALS 2,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_invoice,
             logo     TYPE string,   "! PNG dạng base64
             invno    TYPE string,
             customer TYPE string,
             items    TYPE ty_items,
           END OF ty_invoice.

    "! PNG 1x1 trong suốt để chạy thử — dự án thật đọc logo từ bảng Z
    "! hoặc MIME repository (xem ZCL_QM_EXPORT_LOGO)
    CONSTANTS c_demo_logo TYPE string VALUE
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='.

    "! Template .xlsx — thực tế do user Save As từ Excel rồi lưu vào bảng
    "! template; ở đây sinh bằng XCO để example tự chạy được.
    METHODS get_template RETURNING VALUE(rv_template) TYPE xstring.

    METHODS get_file
      IMPORTING is_invoice     TYPE ty_invoice
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX20_LOGO IMPLEMENTATION.


  METHOD get_template.
    DATA(lo_wa) = xco_cp_xlsx=>document->empty( )->write_access( ).
    DATA(lo_ws) = lo_wa->get_workbook( )->worksheet->at_position( 1 ).
    lo_ws->set_name( `INVOICE` ).

    DATA(lt_lines) = VALUE string_table(
      ( `{{*image:logo@2x4}}` )                    " A1: vùng logo
      ( `` )
      ( `` )
      ( `` )
      ( `INVOICE {{invno}}` )                      " A5
      ( `Customer: {{customer}}` )                 " A6
      ( `{{#items}}{{matnr}}{{/items}}` )          " A7: vùng lặp GỌN TRONG 1 DÒNG
      ( `Total: {{sum:items.amount}}` ) ).         " A8: footer, KHÔNG nằm trong vùng lặp

    LOOP AT lt_lines INTO DATA(lv_text).
      lo_ws->cursor(
        io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
        io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( sy-tabix )
        )->get_cell( )->value->write_from( lv_text ).
    ENDLOOP.

    " cột B: số tiền của từng item (cùng dòng với {{#items}})
    lo_ws->cursor(
      io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
      io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 7 )
      )->get_cell( )->value->write_from( `{{amount}}` ).

    rv_template = lo_wa->get_file_content( ).
  ENDMETHOD.


  METHOD get_file.
    DATA(ls_context) = is_invoice.
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX20_LOGO`
                iv_fallback_template = get_template( )
                iv_fallback_engine   = `XLSX`
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
