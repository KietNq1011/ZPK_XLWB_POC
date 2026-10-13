"! <p class="shorttext synchronized" lang="en">XLWB: chạy thử form mẫu trong Eclipse (F9)</p>
"!
"! Cách dùng:
"! 1. Đổi giá trị hằng {@link zcl_xlwb_run.DATA:c_example} sang form muốn xem.
"! 2. Bấm F9 (Run as ABAP Application) trong Eclipse.
"! 3. Console in ra nội dung SpreadsheetML. Copy toàn bộ, dán vào Notepad,
"!    lưu tên "test.xls" (chọn All files, encoding UTF-8) rồi mở bằng Excel.
"!
"! Với form dùng engine XLSX (EX20) file là nhị phân nên chỉ in kích thước và
"! danh sách part — muốn lấy file thật thì trả base64 qua action RAP / HTTP service.
CLASS zcl_xlwb_run DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.

    CONSTANTS:
      "! Đổi dòng này: EX00 EX01 EX02 EX03 EX05 EX08 EX11 EX13 EX16 EX20 EX21
      c_example TYPE string VALUE 'EX05',
      "! PNG 1x1 trong suốt — thay bằng logo công ty khi dùng thật
      c_logo    TYPE string VALUE 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='.
ENDCLASS.



CLASS ZCL_XLWB_RUN IMPLEMENTATION.


  METHOD if_oo_adt_classrun~main.

    TRY.
        CASE c_example.

          WHEN 'EX00'.
            out->write( |=== 4.00 Hello World ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex00_hello( )->get_file( ) ) ).

          WHEN 'EX01'.
            out->write( |=== 4.01 Shipping label ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex01_label( )->get_file( VALUE #(
                to       = VALUE #( name = 'ACME Corp' street = '1 Main St' city = 'Chicago' )
                shipdate = cl_abap_context_info=>get_system_date( )
                weight   = '125.50' ) ) ) ).

          WHEN 'EX02'.
            out->write( |=== 4.02 N labels + page break ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex02_labels( )->get_file( VALUE #(
                ( name = 'ACME Corp'    city = 'Chicago' )
                ( name = 'Globex Ltd'   city = 'London' )
                ( name = 'Initech GmbH' city = 'Berlin' ) ) ) ) ).

          WHEN 'EX03'.
            out->write( |=== 4.03 mỗi label 1 worksheet ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex03_sheets( )->get_file( VALUE #(
                ( name = 'ACME Corp'  city = 'Chicago' )
                ( name = 'Globex Ltd' city = 'London' ) ) ) ) ).

          WHEN 'EX05'.
            out->write( |=== 4.05 Order form + subtotal ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex05_order( )->get_file( VALUE #(
                orderno  = 'PO-2026-0815'
                customer = 'ACME Corp'
                orddate  = cl_abap_context_info=>get_system_date( )
                remark   = 'Rush order'
                items    = VALUE #(
                  ( pos = 10 matnr = 'SLAB-WHITE-3CM' qty = 20 price = '105.00' amount = '2100.00' )
                  ( pos = 20 matnr = 'SLAB-GREY-2CM'  qty = 15 price = '88.50'  amount = '1327.50' )
                  ( pos = 30 matnr = 'SLAB-BLACK-3CM' qty = 5  price = '120.00' amount = '600.00' ) ) ) ) ) ).

          WHEN 'EX08'.
            out->write( |=== 4.09 bảng số cột động ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex08_dyncols( )->get_file( VALUE #(
                cols = VALUE #( ( title = 'Jan' ) ( title = 'Feb' ) ( title = 'Mar' ) )
                rows = VALUE #(
                  ( rowname = 'White quartz' cells = VALUE #( ( val = 10 ) ( val = 20 ) ( val = 30 ) ) )
                  ( rowname = 'Grey quartz'  cells = VALUE #( ( val = 5 )  ( val = 6 )  ( val = 7 ) ) ) ) ) ) ) ).

          WHEN 'EX11'.
            out->write( |=== 4.11 danh sách 3 cấp + subtotal + gộp ô ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex11_multilevel( )->get_file( VALUE #(
                ( cityfrom = 'HANOI' cityto = 'SAIGON' conns = VALUE #(
                    ( connid = 'VN213' carrier = 'VN' flights = VALUE #(
                        ( fldate = '20260801' capacity = 180 occupied = 150 )
                        ( fldate = '20260802' capacity = 180 occupied = 170 ) ) ) ) ) ) ) ) ).

          WHEN 'EX13'.
            out->write( |=== 4.13 Gantt, header 3 cấp ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex13_gantt( )->get_file( VALUE #(
                monthname = 'August 2026'
                days  = VALUE #( ( label = '1' ) ( label = '2' ) ( label = '3' ) )
                weeks = VALUE #( ( label = 'W31' days = VALUE #( ( label = '1' ) ( label = '2' ) ) )
                                 ( label = 'W32' days = VALUE #( ( label = '3' ) ) ) )
                tasks = VALUE #(
                  ( phase = 'Design' task = 'Blueprint' duration = 2
                    cells = VALUE #( ( mark = 'X' ) ( mark = 'X' ) ( mark = '' ) ) ) ) ) ) ) ).

          WHEN 'EX16'.
            out->write( |=== 4.16 matrix layout + data validation ===| ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              NEW zcl_xlwb_ex16_advanced( )->get_file(
                iv_taxcode = '0101234567' iv_show_secret = abap_false ) ) ).

          WHEN 'EX20'.
            out->write( |=== 4.06 Invoice có logo (engine XLSX, nhị phân) ===| ).
            DATA(lv_xlsx) = NEW zcl_xlwb_ex20_logo( )->get_file( VALUE #(
              logo     = c_logo
              invno    = 'INV-2026-0042'
              customer = 'ACME Corp'
              items    = VALUE #( ( matnr = 'SLAB-WHITE' amount = '2100.00' )
                                  ( matnr = 'SLAB-GREY'  amount = '1327.50' ) ) ) ).
            out->write( |Kích thước file: { xstrlen( lv_xlsx ) } byte| ).
            DATA(lo_zip) = NEW cl_abap_zip( ).
            lo_zip->load( lv_xlsx ).
            LOOP AT lo_zip->files INTO DATA(ls_file).
              out->write( |part: { ls_file-name }| ).
            ENDLOOP.
            out->write( |Muốn lấy file: trả base64 qua action RAP hoặc HTTP service.| ).

          WHEN 'EX21'.
            out->write( |=== Packing List HQ (template từ file Excel thật) ===| ).
            DATA(ls_pack_ctx) = zcl_xlwb_ex21_packlist=>sample_context( ).
            out->write( cl_abap_conv_codepage=>create_in( )->convert(
              zcl_xlwb_runtime=>render_with_template(
                iv_template = cl_abap_conv_codepage=>create_out( )->convert(
                                zcl_xlwb_ex21_packlist=>get_template( ) )
                iv_engine   = `SSML`
                ir_context  = REF #( ls_pack_ctx ) )-content ) ).

          WHEN OTHERS.
            out->write( |Giá trị c_example không hợp lệ: { c_example }| ).
        ENDCASE.

      CATCH zcx_xlwb INTO DATA(lx).
        out->write( |LỖI XLWB: { lx->get_text( ) }| ).
    ENDTRY.

  ENDMETHOD.
ENDCLASS.
