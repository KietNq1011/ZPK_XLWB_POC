"! <p class="shorttext synchronized" lang="en">XLWB: nạp catalog ví dụ cho app Fiori</p>
"!
"! Bấm F9 để nạp/làm mới bảng ZXLWB_DEMO — nguồn dữ liệu của app tra cứu ví dụ.
"! Mỗi bản ghi gồm mô tả, 4 bước (context → vẽ Excel → placeholder → engine),
"! lưu ý, và file Excel mẫu render sẵn để tải về mở bằng Excel.
CLASS zcl_xlwb_demo_load DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.

    "! Nạp toàn bộ catalog (dùng được cả từ code khác, không chỉ F9)
    CLASS-METHODS load
      RETURNING VALUE(rv_count) TYPE i.

  PRIVATE SECTION.
    CLASS-DATA mt_rows TYPE STANDARD TABLE OF zxlwb_demo WITH EMPTY KEY.

    CLASS-METHODS add
      IMPORTING iv_id       TYPE string
                iv_seq      TYPE i
                iv_ref      TYPE string
                iv_title    TYPE string
                iv_form     TYPE string
                iv_purpose  TYPE string
                iv_context  TYPE string
                iv_excel    TYPE string
                iv_template TYPE string
                iv_engine   TYPE string
                iv_notes    TYPE string.
    CLASS-METHODS build.
ENDCLASS.



CLASS ZCL_XLWB_DEMO_LOAD IMPLEMENTATION.


  METHOD add.
    DATA ls TYPE zxlwb_demo.

    ls-example_id    = iv_id.
    ls-seq           = iv_seq.
    ls-xlwb_ref      = iv_ref.
    ls-title         = iv_title.
    ls-template_form = iv_form.
    ls-purpose       = iv_purpose.
    ls-step_context  = iv_context.
    ls-step_excel    = iv_excel.
    ls-step_template = iv_template.
    ls-step_engine   = iv_engine.
    ls-notes         = iv_notes.
    ls-created_by    = cl_abap_context_info=>get_user_technical_name( ).
    GET TIME STAMP FIELD ls-created_at.

    TRY.
        DATA(ls_file) = zcl_xlwb_demo_files=>render( iv_id ).
        ls-content     = ls_file-content.
        ls-file_name   = ls_file-file_name.
        ls-mime_type   = ls_file-mime_type.
        ls-engine_type = ls_file-engine.
      CATCH zcx_xlwb INTO DATA(lx).
        ls-notes = |{ ls-notes }| && |[Không render được file mẫu: | && |{ lx->get_text( ) }]|.
    ENDTRY.

    APPEND ls TO mt_rows.
  ENDMETHOD.


  METHOD build.
    CLEAR mt_rows.

    add( iv_id = `EX00` iv_seq = 10 iv_ref = `4.00`
         iv_title = `Hello World — giá trị đơn`
         iv_form  = `XLWB_EX00_HELLO`
         iv_purpose =
|Form nhỏ nhất có thể: một giá trị từ context vào một ô Excel.\n| &&
           |Dùng để hiểu ba mảnh của mọi form: TEMPLATE (hình thức),\n| &&
           |CONTEXT (dữ liệu), ENGINE (nối hai thứ đó lại).\n|
         iv_context =
|" Context là một structure ABAP thường, không cần interface gì.\n| &&
           |DATA: BEGIN OF ls_context,\n| &&
           |        message TYPE string,\n| &&
           |      END OF ls_context.\n| &&
           |ls_context-message = 'Hello world!'.\n|
         iv_excel =
|1. Mở Excel, tạo workbook mới.\n| &&
           |2. Nháy đúp vào tab sheet ở dưới, đổi tên sheet thành "Hello".\n| &&
           |3. Nháy vào ô A1 — ví dụ này chỉ cần một ô.\n| &&
           |4. File > Save As > chọn kiểu "XML Spreadsheet 2003 (*.xml)".\n| &&
           |\n| &&
           |Chưa cần định dạng gì. Các ví dụ sau sẽ thêm khung, màu, số.\n|
         iv_template =
|Gõ trực tiếp vào ô A1 trong Excel:\n| &&
           |    \{\{message\}\}\n| &&
           |\n| &&
           |Sau khi Save As XML, file sẽ chứa:\n| &&
           |    <Row><Cell><Data ss:Type="String">\{\{message\}\}</Data></Cell></Row>\n| &&
           |\n| &&
           |Tên field trong context = tên trong \{\{...\}\} (không phân biệt hoa/thường).\n|
         iv_engine =
|DATA(lv_file) = zcl_xlwb_runtime=>render_prefer_table(\n| &&
           |                  iv_form_name         = `XLWB_EX00_HELLO`\n| &&
           |                  iv_fallback_template = cl_abap_conv_codepage=>create_out(\n| &&
           |                                           )->convert( get_template( ) )\n| &&
           |                  ir_context           = REF #( ls_context ) )-content.\n| &&
           |\n| &&
           |Xem class ZCL_XLWB_EX00_HELLO.\n|
         iv_notes =
|Sai tên placeholder thì engine NÉM zcx_xlwb kèm tên path — không âm thầm ra ô rỗng.\n| &&
           |\n| &&
           |Template: ưu tiên bản trong bảng ZXLWB_TMPL (form XLWB_EX00_HELLO). Chưa upload\n| &&
           |template ở đó thì engine dùng bản dự phòng dựng bằng code.\n| ).

    add( iv_id = `EX01` iv_seq = 20 iv_ref = `4.01`
         iv_title = `Shipping label — path lồng nhau + ô có kiểu`
         iv_form  = `XLWB_EX01_LABEL`
         iv_purpose =
|Nhãn vận chuyển: phần FROM tĩnh nằm trong template, phần TO lấy từ context.\n| &&
           |Cho thấy style, độ rộng cột, chiều cao dòng, number format đều thuộc TEMPLATE.\n|
         iv_context =
|TYPES: BEGIN OF ty_party,\n| &&
           |         name   TYPE string,\n| &&
           |         street TYPE string,\n| &&
           |         city   TYPE string,\n| &&
           |       END OF ty_party.\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        to       TYPE ty_party,   " structure lồng nhau\n| &&
           |        shipdate TYPE d,\n| &&
           |        weight   TYPE p LENGTH 8 DECIMALS 2,\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel. Đặt độ rộng cột A khoảng 12, cột B khoảng 30 (kéo mép cột\n| &&
           |   hoặc chuột phải vào đầu cột > Column Width).\n| &&
           |2. Ô A1: gõ "SHIPPING LABEL", bôi đậm, cỡ 14.\n| &&
           |   Chọn A1:B1 > Home > Merge & Center.\n| &&
           |3. Cột A gõ nhãn tĩnh: FROM:, TO:, Ship date:, Weight: (bôi đậm).\n| &&
           |4. Ô ngày (B5): chuột phải > Format Cells > Date > dd/mm/yyyy.\n| &&
           |5. Ô khối lượng (B6): Format Cells > Custom > #,##0.00 "kg"\n| &&
           |6. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Gõ placeholder vào các ô bên cạnh nhãn:\n| &&
           |    B3:  \{\{to.name\}\}\n| &&
           |    B4:  \{\{to.street\}\}, \{\{to.city\}\}\n| &&
           |    B5:  \{\{shipdate\}\}       (ô đã format dd/mm/yyyy)\n| &&
           |    B6:  \{\{weight\}\}         (ô đã format #,##0.00 "kg")\n|
         iv_engine =
|NEW zcl_xlwb_ex01_label( )->get_file( VALUE #(\n| &&
           |  to       = VALUE #( name = 'ACME Corp' street = '1 Main St' city = 'Chicago' )\n| &&
           |  shipdate = cl_abap_context_info=>get_system_date( )\n| &&
           |  weight   = '125.50' ) ).\n|
         iv_notes =
|Ô mà TOÀN BỘ nội dung là một placeholder thì được ghi CÓ KIỂU (số → ss:Type\n| &&
           |Number, ngày → DateTime) nên number format của template mới ăn.\n| &&
           |Ô trộn text như "Total: \{\{amount\}\}" luôn ra String — cần định dạng số thì tách hai ô.\n| ).

    add( iv_id = `EX02` iv_seq = 30 iv_ref = `4.02 / 4.02b`
         iv_title = `N nhãn + mỗi nhãn một trang in`
         iv_form  = `XLWB_EX02_LABELS`
         iv_purpose =
|Lặp nhiều dòng: một vùng 4 dòng được nhân bản theo số dòng của internal table,\n| &&
           |mỗi nhãn chèn page break để in riêng một trang.\n|
         iv_context =
|TYPES: BEGIN OF ty_label,\n| &&
           |         name TYPE string,\n| &&
           |         city TYPE string,\n| &&
           |       END OF ty_label,\n| &&
           |       ty_labels TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY.\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        labels TYPE ty_labels,\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel, vẽ MỘT nhãn mẫu trên 4 dòng liền nhau (A1:A4).\n| &&
           |   Dòng 1: tiêu đề nhãn — bôi đậm, Home > Borders > Bottom Border.\n| &&
           |   Dòng 2: tên người nhận. Dòng 3: thành phố. Dòng 4: để trống làm khoảng cách.\n| &&
           |2. QUAN TRỌNG: chỉ vẽ MỘT nhãn. Engine tự nhân bản theo số dòng dữ liệu.\n| &&
           |3. Hướng giấy: Page Layout > Orientation > Landscape (tuỳ nhu cầu).\n| &&
           |4. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Ô đầu vùng lặp (A1) — mở vùng lặp, thêm marker ngắt trang và số thứ tự:\n| &&
           |    A1:  \{\{#labels\}\}\{\{*break\}\}Label \{\{@index\}\} / \{\{@count\}\}\n| &&
           |    A2:  \{\{name\}\}\n| &&
           |    A3:  \{\{city\}\}\n| &&
           |Ô cuối vùng lặp (A4) — đóng vùng lặp:\n| &&
           |    A4:  \{\{/labels\}\}\n| &&
           |\n| &&
           |Vùng lặp = từ dòng có \{\{#labels\}\} đến dòng có \{\{/labels\}\} (kể cả hai dòng đó).\n|
         iv_engine =
|NEW zcl_xlwb_ex02_labels( )->get_file( VALUE #(\n| &&
           |  ( name = 'ACME Corp'  city = 'Chicago' )\n| &&
           |  ( name = 'Globex Ltd' city = 'London' ) ) ).\n|
         iv_notes =
|\{\{*break\}\} ở dòng đầu tiên được bỏ qua (ngắt trang trước dòng 1 vô nghĩa).\n| &&
           |PageSetup (hướng giấy, khổ giấy, lề) giữ nguyên từ template.\n| ).

    add( iv_id = `EX02A` iv_seq = 35 iv_ref = `4.02a`
         iv_title = `N nhãn NGANG — mỗi nhãn một cột (cell loop)`
         iv_form  = `XLWB_EX02A_HLABELS`
         iv_purpose =
|Nhãn xếp NGANG sang phải thay vì dọc xuống: mỗi bản ghi một CỘT.\n| &&
           |Cell loop thuần tuý — nền tảng của mọi form cột động (EX08, EX13).\n|
         iv_context =
|" Cùng context với EX02:\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        labels TYPE ty_labels,   " name + city\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel. Cột A gõ nhãn tĩnh: TO:, Name:, City: (bôi đậm).\n| &&
           |2. Cột B là CỘT MẪU của một nhãn — 3 ô B1:B3, đặt độ rộng khoảng 18.\n| &&
           |   Ô B1 bôi đậm + Fill Color xanh nhạt làm tiêu đề nhãn.\n| &&
           |3. Chỉ vẽ MỘT cột nhãn — engine nhân bản sang phải theo số bản ghi.\n| &&
           |4. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|MỖI DÒNG có cell loop RIÊNG trên CÙNG một bảng — các cột tự thẳng hàng:\n| &&
           |    B1:  \{\{#labels>\}\}Label \{\{@index\}\}\n| &&
           |    B2:  \{\{#labels>\}\}\{\{name\}\}\n| &&
           |    B3:  \{\{#labels>\}\}\{\{city\}\}\n| &&
           |\n| &&
           |Dấu > = lặp Ô NGANG. Không có \{\{/...\}\} vì cell loop chỉ chiếm một ô.\n|
         iv_engine =
|NEW zcl_xlwb_ex02a_hlabels( )->get_file( VALUE #(\n| &&
           |  ( name = 'ACME Corp'  city = 'Chicago' )\n| &&
           |  ( name = 'Globex Ltd' city = 'London' ) ) ).\n|
         iv_notes =
|So sánh với EX02 (lặp DÒNG — dọc xuống): cùng dữ liệu, đổi hướng chỉ là đổi template.\n| &&
           |Style và độ rộng của Ô MẪU đi theo mọi bản sao.\n| ).

    add( iv_id = `EX03` iv_seq = 40 iv_ref = `4.03`
         iv_title = `Mỗi bản ghi một worksheet`
         iv_form  = `XLWB_EX03_SHEETS`
         iv_purpose =
|Sheet động: tên worksheet chứa marker lặp nên CẢ worksheet được nhân bản.\n|
         iv_context =
|DATA: BEGIN OF ls_ctx,\n| &&
           |        labels TYPE ty_labels,   " như EX02\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel, vẽ nội dung của MỘT bản ghi trên sheet đầu tiên.\n| &&
           |2. Nháy đúp vào tab sheet để đổi tên — đây là chỗ đặt marker lặp.\n| &&
           |3. Chỉ để MỘT sheet trong workbook (xoá Sheet2, Sheet3 nếu có).\n| &&
           |4. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Tên sheet (gõ khi đổi tên tab sheet):\n| &&
           |    \{\{#labels\}\}Label \{\{@index\}\}\n| &&
           |\n| &&
           |Nội dung trong sheet:\n| &&
           |    A1:  TO: \{\{name\}\}\n| &&
           |    A2:  \{\{city\}\}\n| &&
           |\n| &&
           |Phần còn lại của tên sheet là template và phải render ra tên DUY NHẤT.\n|
         iv_engine =
|NEW zcl_xlwb_ex03_sheets( )->get_file( it_labels ).\n|
         iv_notes =
|Tên sheet trùng nhau làm Excel từ chối mở file — luôn đưa \{\{@index\}\} hoặc\n| &&
           |một field khoá vào tên sheet.\n| &&
           |Excel giới hạn tên sheet 31 ký tự và không cho các ký tự : \\ / ? * [ ]\n| ).

    add( iv_id = `EX04` iv_seq = 45 iv_ref = `4.04`
         iv_title = `Workbook nhiều form — một file, nhiều worksheet`
         iv_form  = `XLWB_EX04_BUNDLE`
         iv_purpose =
|MỘT lần render ra MỘT workbook chứa NHIỀU form khác nhau (Order + Labels),\n| &&
           |mỗi form một worksheet tĩnh. Thay thế bài "save several forms into one\n| &&
           |workbook" của tool gốc mà không cần viewer.\n|
         iv_context =
|" Context GOM các nhánh, mỗi worksheet bind một nhánh:\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        orderno  TYPE string,\n| &&
           |        customer TYPE string,\n| &&
           |        items    TYPE ty_items,    " cho sheet Order\n| &&
           |        labels   TYPE ty_labels,   " cho sheet Labels\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel, tạo HAI sheet: đổi tên tab thành "Order" và "Labels".\n| &&
           |2. Sheet Order: vẽ form order nhỏ (tiêu đề, đầu bảng, MỘT dòng item, dòng tổng).\n| &&
           |3. Sheet Labels: vẽ MỘT nhãn 3 dòng như EX02.\n| &&
           |4. File > Save As > XML Spreadsheet 2003 — cả hai sheet nằm chung một file.\n|
         iv_template =
|Sheet "Order":\n| &&
           |    A1:  ORDER \{\{orderno\}\} — \{\{customer\}\}\n| &&
           |    A3:  \{\{#items\}\}\{\{matnr\}\}   ...   C3: \{\{amount\}\}\{\{/items\}\}\n| &&
           |    B4:  \{\{sum:items.qty\}\}       C4: \{\{sum:items.amount\}\}\n| &&
           |\n| &&
           |Sheet "Labels":\n| &&
           |    A1:  \{\{#labels\}\}Label \{\{@index\}\} / \{\{@count\}\}\n| &&
           |    A3:  \{\{city\}\}\{\{/labels\}\}\n| &&
           |\n| &&
           |Khác EX03: tên sheet KHÔNG chứa marker nên sheet không bị nhân bản.\n|
         iv_engine =
|NEW zcl_xlwb_ex04_bundle( )->get_file( VALUE #(\n| &&
           |  orderno = 'PO-2026-0815' customer = 'ACME Corp'\n| &&
           |  items   = ...  labels = ... ) ).\n|
         iv_notes =
|Cần gộp nhiều lần render thành một file .zip (mỗi file một workbook riêng) thì\n| &&
           |làm ở frontend bằng JSZip — pattern có sẵn ở ECUS.\n| ).

    add( iv_id = `EX05` iv_seq = 50 iv_ref = `4.05 / 4.11a`
         iv_title = `Order form: Header / bảng item / Footer`
         iv_form  = `XLWB_EX05_ORDER`
         iv_purpose =
|Mẫu chứng từ kinh điển và hay dùng nhất: header tĩnh, bảng item lặp, dòng tổng.\n| &&
           |Có cả tổng do ENGINE tính và tổng do EXCEL tính bằng công thức.\n|
         iv_context =
|TYPES: BEGIN OF ty_item,\n| &&
           |         pos    TYPE i,\n| &&
           |         matnr  TYPE string,\n| &&
           |         qty    TYPE p LENGTH 8  DECIMALS 0,\n| &&
           |         price  TYPE p LENGTH 11 DECIMALS 2,\n| &&
           |         amount TYPE p LENGTH 11 DECIMALS 2,\n| &&
           |       END OF ty_item,\n| &&
           |       ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY.\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        orderno  TYPE string,\n| &&
           |        customer TYPE string,\n| &&
           |        orddate  TYPE d,\n| &&
           |        items    TYPE ty_items,   " vùng lặp\n| &&
           |        remark   TYPE string,\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel. Đặt độ rộng 5 cột: A=6 (Pos), B=28 (Material), C=8 (Qty),\n| &&
           |   D=12 (Price), E=14 (Amount).\n| &&
           |2. Dòng 1 — tiêu đề: gõ "PURCHASE ORDER", chọn A1:E1 > Merge & Center,\n| &&
           |   bôi đậm, cỡ 14.\n| &&
           |3. Dòng 2 — thông tin: A2 "Customer:", C2 "Date:".\n| &&
           |   Ô D2: Format Cells > Date > dd/mm/yyyy.\n| &&
           |4. Dòng 3 — đầu bảng: gõ Pos / Material / Qty / Price / Amount.\n| &&
           |   Chọn A3:E3 > bôi đậm > Fill Color xanh nhạt > Borders > Bottom Border.\n| &&
           |5. Dòng 4 — MỘT dòng item mẫu (chỉ một dòng!).\n| &&
           |   Chọn D4:E4 > Format Cells > Number > 2 decimal places + Use 1000 Separator.\n| &&
           |6. Dòng 5 — dòng tổng: chọn D5:E5 > bôi đậm > Borders > Top and Double Bottom.\n| &&
           |7. Dòng 6 — dòng tổng bằng công thức Excel. Dòng 7 — ghi chú.\n| &&
           |8. Vùng in: Page Layout > Print Area > Set Print Area.\n| &&
           |9. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Header:\n| &&
           |    A1:  PURCHASE ORDER \{\{orderno\}\}\n| &&
           |    B2:  \{\{customer\}\}          D2:  \{\{orddate\}\}\n| &&
           |\n| &&
           |Dòng item (dòng 4) — mở và đóng vùng lặp TRÊN CÙNG MỘT DÒNG:\n| &&
           |    A4:  \{\{#items\}\}\{\{pos\}\}\n| &&
           |    B4:  \{\{matnr\}\}\n| &&
           |    C4:  \{\{qty\}\}\n| &&
           |    D4:  \{\{price\}\}\n| &&
           |    E4:  \{\{amount\}\}\{\{/items\}\}\n| &&
           |\n| &&
           |Dòng tổng do engine tính (dòng 5):\n| &&
           |    D5:  \{\{sum:items.qty\}\}\n| &&
           |    E5:  \{\{sum:items.amount\}\}\n| &&
           |\n| &&
           |Dòng tổng do Excel tính (dòng 6): gõ công thức thật vào ô, engine không chạm:\n| &&
           |    E6:  =SUM(E4:E5)\n|
         iv_engine =
|NEW zcl_xlwb_ex05_order( )->get_file( VALUE #(\n| &&
           |  orderno  = 'PO-2026-0815'\n| &&
           |  customer = 'ACME Corp'\n| &&
           |  orddate  = cl_abap_context_info=>get_system_date( )\n| &&
           |  items    = VALUE #( ( pos = 10 matnr = 'SLAB-WHITE' qty = 20\n| &&
           |                        price = '105.00' amount = '2100.00' ) ) ) ).\n|
         iv_notes =
|\{\{sum:...\}\} tính trong PHẠM VI vùng lặp hiện tại, nên đặt trong loop cha sẽ ra\n| &&
           |subtotal của đúng dòng cha đó.\n| &&
           |\n| &&
           |Công thức Excel tham chiếu tương đối kiểu R[-n]C không dùng được khi số dòng thay\n| &&
           |đổi — hãy quét từ dòng cố định, ví dụ SUM(R4C:R[-2]C).\n| ).

    add( iv_id = `EX08` iv_seq = 60 iv_ref = `4.09 / 4.08a`
         iv_title = `Bảng số cột động (ma trận)`
         iv_form  = `XLWB_EX08_DYNCOLS`
         iv_purpose =
|Số cột quyết định lúc runtime: mỗi dòng context chứa một bảng con các ô.\n| &&
           |Kết hợp lặp dòng và lặp ô ngang thành ma trận.\n|
         iv_context =
|TYPES: BEGIN OF ty_cell,  val TYPE p LENGTH 8 DECIMALS 2, END OF ty_cell,\n| &&
           |       ty_cells TYPE STANDARD TABLE OF ty_cell WITH EMPTY KEY,\n| &&
           |       BEGIN OF ty_row,\n| &&
           |         rowname TYPE string,\n| &&
           |         cells   TYPE ty_cells,   " bảng CON → cột động\n| &&
           |       END OF ty_row.\n|
         iv_excel =
|1. Mở Excel. Chỉ vẽ HAI cột: cột A (tên dòng) và cột B (một ô số mẫu).\n| &&
           |   Engine nhân bản cột B sang phải theo số cột thật.\n| &&
           |2. Ô A1 và B1 là hàng tiêu đề: bôi đậm, Fill Color xanh nhạt.\n| &&
           |3. Ô B2: Format Cells > Number > 2 decimal places — định dạng này được\n| &&
           |   nhân bản cho MỌI cột sinh ra.\n| &&
           |4. Đừng vẽ sẵn nhiều cột trống — đó là cách làm sai.\n| &&
           |5. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Hàng tiêu đề: ô B1 được nhân bản SANG PHẢI theo bảng cols\n| &&
           |    A1:  Product\n| &&
           |    B1:  \{\{#cols>\}\}\{\{title\}\}\n| &&
           |\n| &&
           |Thân bảng: lặp dòng + lặp ô\n| &&
           |    A2:  \{\{#rows\}\}\{\{rowname\}\}\n| &&
           |    B2:  \{\{#cells>\}\}\{\{val\}\}\n| &&
           |    C2:  \{\{/rows\}\}      (ô cuối để đóng vùng lặp dòng)\n| &&
           |\n| &&
           |Dấu > trong \{\{#cols>\}\} = lặp Ô NGANG (khác \{\{#cols\}\} là lặp DÒNG).\n|
         iv_engine =
|NEW zcl_xlwb_ex08_dyncols( )->get_file( VALUE #(\n| &&
           |  cols = VALUE #( ( title = 'Jan' ) ( title = 'Feb' ) )\n| &&
           |  rows = VALUE #( ( rowname = 'White quartz'\n| &&
           |                    cells = VALUE #( ( val = 10 ) ( val = 20 ) ) ) ) ) ).\n|
         iv_notes =
|Muốn ẩn một cột: đừng đưa cột đó vào bảng cols — không cần sửa template.\n| &&
           |Số phần tử của cells nên bằng số phần tử của cols, nếu không bảng sẽ lệch cột.\n| ).

    add( iv_id = `EX08G` iv_seq = 62 iv_ref = `4.08`
         iv_title = `Row grouping — dòng chi tiết gập được (+/-)`
         iv_form  = `XLWB_EX08G_GROUPING`
         iv_purpose =
|Danh sách nhân viên theo vị trí (đúng bài 4.08 gốc): dòng vị trí là summary,\n| &&
           |các dòng nhân viên GẬP ĐƯỢC bằng nút +/- bên lề trái Excel (outline).\n|
         iv_context =
|TYPES: BEGIN OF ty_emp, name TYPE string, role TYPE string, END OF ty_emp,\n| &&
           |       BEGIN OF ty_position,\n| &&
           |         position TYPE string,\n| &&
           |         emps     TYPE ty_emps,   " loop con = nhóm gập được\n| &&
           |       END OF ty_position.\n|
         iv_excel =
|1. Mở Excel, vẽ 2 dòng: dòng VỊ TRÍ (bôi đậm, Fill xanh nhạt, Merge 2 cột)\n| &&
           |   và MỘT dòng nhân viên (thụt đầu dòng vài dấu cách).\n| &&
           |2. KHÔNG cần Data > Group trong Excel — XML 2003 không lưu grouping;\n| &&
           |   nhóm được khai bằng marker \{\{*group=1\}\} (bước sau).\n| &&
           |3. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Dòng vị trí (summary — KHÔNG marker):\n| &&
           |    A2:  \{\{#positions\}\}\{\{position\}\} (\{\{cnt:emps\}\} employees)\n| &&
           |\n| &&
           |Dòng nhân viên — marker nhóm cấp 1:\n| &&
           |    A3:  \{\{#emps\}\}\{\{*group=1\}\}    \{\{name\}\}\n| &&
           |    B3:  \{\{role\}\}\{\{/emps\}\}      C3: \{\{/positions\}\}\n| &&
           |\n| &&
           |\{\{*group=N\}\} = outline cấp N (1..7); lồng nhóm trong nhóm bằng cấp tăng dần.\n|
         iv_engine =
|NEW zcl_xlwb_ex08g_grouping( )->get_file( VALUE #(\n| &&
           |  ( position = 'Manager' emps = VALUE #(\n| &&
           |      ( name = 'Nguyen Van A' role = 'Sales manager' ) ) ) ) ).\n|
         iv_notes =
|Summary nằm TRÊN nhóm chi tiết (outlinePr summaryBelow=0) — đúng bố cục 4.08.\n| &&
           |Marker chỉ có tác dụng khi output là .xlsx (mặc định hiện tại); nếu converter\n| &&
           |lỗi phải fallback .xls thô thì nhóm không hiển thị (XML 2003 không chứa outline).\n| &&
           |Template .xlsx cho engine XLSX: group thẳng trong Excel (Data > Group) — engine\n| &&
           |giữ nguyên outlineLevel khi nhân bản dòng, không cần marker.\n| ).

    add( iv_id = `EX09` iv_seq = 65 iv_ref = `4.09`
         iv_title = `Bảng động — số dòng VÀ số cột chỉ biết lúc runtime`
         iv_form  = `XLWB_EX09_DYNTABLE`
         iv_purpose =
|Đúng bài 4.09 của tool gốc (DYNTABLE): kích thước bảng sinh NGẪU NHIÊN mỗi lần\n| &&
           |chạy — bấm "Gen lại file mẫu" nhiều lần sẽ thấy bảng to nhỏ khác nhau.\n| &&
           |Thêm tổng theo DÒNG (engine tính) và tổng theo CỘT (ABAP tính sẵn).\n|
         iv_context =
|" Bảng lồng bảng — kích thước tuỳ ý:\n| &&
           |TYPES: BEGIN OF ty_cell, val TYPE p LENGTH 8 DECIMALS 2, END OF ty_cell,\n| &&
           |       BEGIN OF ty_row,\n| &&
           |         rowname TYPE string,\n| &&
           |         cells   TYPE ty_cells,   " số phần tử = số cột\n| &&
           |       END OF ty_row.\n| &&
           |DATA: BEGIN OF ls_ctx,\n| &&
           |        rows_n TYPE i, cols_n TYPE i,\n| &&
           |        cols   TYPE ty_cols,     " tiêu đề cột\n| &&
           |        rows   TYPE ty_rows,\n| &&
           |        coltotals TYPE ty_cells, " tổng cột — ABAP tính sẵn\n| &&
           |        grand  TYPE p LENGTH 11 DECIMALS 2,\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel, vẽ 4 dòng x 3 cột: tiêu đề, header, MỘT dòng dữ liệu, dòng tổng.\n| &&
           |2. Cột A = tên dòng; cột B là CỘT MẪU (một ô duy nhất cho header, dữ liệu,\n| &&
           |   tổng cột); cột C = tổng dòng (bôi đậm, Fill Color cam nhạt).\n| &&
           |3. Ô B của dòng dữ liệu: Format Cells > Number > 2 decimals.\n| &&
           |4. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Header cột (B2) + cột Total đứng SAU vùng lặp ngang:\n| &&
           |    B2:  \{\{#cols>\}\}\{\{title\}\}      C2: Row total\n| &&
           |\n| &&
           |Dòng dữ liệu (B3) — ma trận + tổng dòng trong frame của dòng:\n| &&
           |    A3:  \{\{#rows\}\}\{\{rowname\}\}\n| &&
           |    B3:  \{\{#cells>\}\}\{\{val\}\}\n| &&
           |    C3:  \{\{sum:cells.val\}\}\{\{/rows\}\}\n| &&
           |\n| &&
           |Dòng tổng cột (B4) — bảng coltotals tính sẵn ở ABAP:\n| &&
           |    B4:  \{\{#coltotals>\}\}\{\{val\}\}      C4: \{\{grand\}\}\n|
         iv_engine =
|" Không truyền kích thước -> random như 4.09 gốc:\n| &&
           |NEW zcl_xlwb_ex09_dyntable( )->get_file( ).\n| &&
           |" Truyền kích thước cố định (unit test / demo lặp lại được):\n| &&
           |NEW zcl_xlwb_ex09_dyntable( )->get_file( iv_rows = 4 iv_cols = 5 ).\n|
         iv_notes =
|\{\{sum:...\}\} tính trong frame HIỆN TẠI nên đặt ở dòng nào là tổng của dòng đó —\n| &&
           |vì vậy tổng theo CỘT phải tính sẵn trong ABAP (bảng coltotals).\n| &&
           |So với EX08: cùng kỹ thuật ma trận, EX09 thêm tổng 2 chiều và kích thước random.\n| ).

    add( iv_id = `EX11` iv_seq = 70 iv_ref = `4.11 / 4.12`
         iv_title = `Danh sách 3 cấp + subtotal + gộp ô dọc`
         iv_form  = `XLWB_EX11_MULTILEVEL`
         iv_purpose =
|Loop lồng nhau 3 tầng, subtotal từng tầng, và ô cấp cha tự động kéo dọc hết\n| &&
           |số dòng con mà không cần biết trước số dòng.\n|
         iv_context =
|" nested table 3 tầng: routes → conns → flights\n| &&
           |TYPES: BEGIN OF ty_conn,\n| &&
           |         connid  TYPE string,\n| &&
           |         flights TYPE ty_flights,\n| &&
           |       END OF ty_conn,\n| &&
           |       BEGIN OF ty_route,\n| &&
           |         cityfrom TYPE string,\n| &&
           |         conns    TYPE ty_conns,\n| &&
           |       END OF ty_route.\n|
         iv_excel =
|1. Mở Excel, vẽ 5 cột: Route / Connection / Date / Capacity / Occupied.\n| &&
           |2. Vẽ MỘT dòng cho mỗi CẤP (không vẽ nhiều dòng dữ liệu):\n| &&
           |   - dòng cấp 1 (route): chọn A:E > Merge & Center, bôi đậm.\n| &&
           |   - dòng cấp 2 + 3 (connection và flight): một dòng duy nhất.\n| &&
           |   - dòng subtotal cấp 2, dòng subtotal cấp 1: bôi đậm + Borders > Top Border.\n| &&
           |3. Ô cột A của dòng cấp 2: Format Cells > Alignment > Vertical = Center\n| &&
           |   để chữ nằm giữa khi ô bị gộp dọc.\n| &&
           |4. Ô ngày: Format Cells > Date > dd/mm/yyyy.\n| &&
           |5. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Cấp 1 — route:\n| &&
           |    A2:  \{\{#routes\}\}\{\{cityfrom\}\} - \{\{cityto\}\}\n| &&
           |\n| &&
           |Cấp 2 — connection, ô này tự gộp dọc theo số chuyến bay của nó:\n| &&
           |    A3:  \{\{#conns\}\}\{\{*mergedown:flights\}\}\{\{carrier\}\} \{\{connid\}\}\n| &&
           |\n| &&
           |Cấp 3 — flight (cùng dòng với cấp 2):\n| &&
           |    B3:  \{\{#flights\}\}\{\{connid\}\}\n| &&
           |    C3:  \{\{fldate\}\}      D3: \{\{capacity\}\}\n| &&
           |    E3:  \{\{occupied\}\}\{\{/flights\}\}\n| &&
           |\n| &&
           |Subtotal cấp 2 rồi cấp 1:\n| &&
           |    A4:  Subtotal \{\{connid\}\} (\{\{cnt:flights\}\} flights)\n| &&
           |    E4:  \{\{sum:flights.occupied\}\}\{\{/conns\}\}\n| &&
           |    A5:  == Route \{\{cityfrom\}\}: \{\{cnt:conns\}\} connections\{\{/routes\}\}\n|
         iv_engine =
|NEW zcl_xlwb_ex11_multilevel( )->get_file( it_routes ).\n|
         iv_notes =
|Trong loop con vẫn đọc được field của loop cha: engine tìm path ở dòng trong\n| &&
           |cùng trước, rồi ra ngoài, cuối cùng là context gốc.\n| &&
           |\{\{*mergedown:flights\}\} đặt ss:MergeDown = số dòng con − 1.\n| ).

    add( iv_id = `EX13` iv_seq = 80 iv_ref = `4.13`
         iv_title = `Gantt: header 3 cấp gộp ngang động`
         iv_form  = `XLWB_EX13_GANTT`
         iv_purpose =
|Cột ngày sinh động theo độ dài tháng, header 3 tầng (Tháng / Tuần / Ngày)\n| &&
           |với số ô gộp tính theo dữ liệu.\n|
         iv_context =
|TYPES: BEGIN OF ty_week,\n| &&
           |         label TYPE string,\n| &&
           |         days  TYPE ty_days,   " số ngày → số ô gộp\n| &&
           |       END OF ty_week.\n|
         iv_excel =
|1. Mở Excel, vẽ 3 cột tĩnh: Phase / Task / Days (A, B, C).\n| &&
           |2. Cột D là cột ngày MẪU — chỉ vẽ một cột, đặt độ rộng nhỏ (khoảng 3).\n| &&
           |3. Ba dòng header: dòng 1 (tháng), dòng 2 (tuần), dòng 3 (ngày).\n| &&
           |   Cả ba ô ở cột D đều Center, Fill Color khác nhau cho dễ phân biệt.\n| &&
           |4. Ô đánh dấu công việc (dòng 4, cột D): Fill Color xanh lá — đây là "thanh" Gantt.\n| &&
           |5. Cố định 3 cột đầu: chọn ô D1 > View > Freeze Panes > Freeze Panes.\n| &&
           |6. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Header cấp 1 — tháng, gộp ngang hết số cột ngày:\n| &&
           |    D1:  \{\{*mergeacross:days\}\}\{\{monthname\}\}\n| &&
           |\n| &&
           |Header cấp 2 — tuần: lặp ô ngang, mỗi tuần gộp số ngày của nó:\n| &&
           |    D2:  \{\{#weeks>\}\}\{\{*mergeacross:days\}\}\{\{label\}\}\n| &&
           |\n| &&
           |Header cấp 3 — ngày:\n| &&
           |    D3:  \{\{#days>\}\}\{\{label\}\}\n| &&
           |\n| &&
           |Thân bảng: lặp dòng + lặp ô\n| &&
           |    A4:  \{\{#tasks\}\}\{\{phase\}\}     B4: \{\{task\}\}     C4: \{\{duration\}\}\n| &&
           |    D4:  \{\{#cells>\}\}\{\{mark\}\}\n| &&
           |    E4:  \{\{/tasks\}\}\n|
         iv_engine =
|NEW zcl_xlwb_ex13_gantt( )->get_file( is_gantt ).\n|
         iv_notes =
|Marker gộp ô được tính cho TỪNG bản sao khi ô nằm trong vùng lặp ngang, nên mỗi\n| &&
           |tuần có số ô gộp riêng. FreezePanes của template giữ nguyên.\n| ).

    add( iv_id = `EX14` iv_seq = 85 iv_ref = `4.14 / 4.15`
         iv_title = `Cây phân cấp (ALV tree) → nested table`
         iv_form  = `XLWB_EX14_TREE`
         iv_purpose =
|Xuất dữ liệu dạng cây. Tool gốc đọc trực tiếp CL_GUI_ALV_TREE — class GUI\n| &&
           |không tồn tại trên Public Cloud. Pattern thay thế: chuyển bảng PHẲNG\n| &&
           |(node_key + parent_key) sang NESTED TABLE rồi render loop lồng nhau.\n|
         iv_context =
|" Bảng phẳng kiểu ALV tree:\n| &&
           |TYPES: BEGIN OF ty_node,\n| &&
           |         node_key   TYPE string,\n| &&
           |         parent_key TYPE string,   " rỗng = node gốc\n| &&
           |         text       TYPE string,\n| &&
           |         qty        TYPE p LENGTH 8 DECIMALS 0,\n| &&
           |       END OF ty_node.\n| &&
           |\n| &&
           |" Chuyển sang nested bằng helper dùng lại được:\n| &&
           |DATA(lt_groups) = zcl_xlwb_ex14_tree=>to_nested( lt_nodes ).\n|
         iv_excel =
|1. Mở Excel, vẽ 2 cột: Item (rộng 40) và Qty (rộng 10).\n| &&
           |2. Vẽ MỘT dòng cho mỗi cấp:\n| &&
           |   - dòng nhóm: chọn A:B > Merge, bôi đậm, Fill Color xanh nhạt.\n| &&
           |   - dòng item con: thụt đầu dòng bằng vài dấu cách trong template.\n| &&
           |   - dòng subtotal: bôi đậm nghiêng, Borders > Top Border.\n| &&
           |3. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Dòng nhóm (cấp 1):\n| &&
           |    A1:  \{\{#groups\}\}\{\{text\}\}\n| &&
           |\n| &&
           |Dòng item con (cấp 2):\n| &&
           |    A2:  \{\{#items\}\}    \{\{text\}\}      B2: \{\{qty\}\}\{\{/items\}\}\n| &&
           |\n| &&
           |Dòng subtotal nhóm:\n| &&
           |    A3:  Subtotal \{\{text\}\} (\{\{cnt:items\}\} items)\n| &&
           |    B3:  \{\{sum:items.qty\}\}\{\{/groups\}\}\n|
         iv_engine =
|NEW zcl_xlwb_ex14_tree( )->get_file( VALUE #(\n| &&
           |  ( node_key = 'G1' text = 'Raw slabs' )\n| &&
           |  ( node_key = 'I1' parent_key = 'G1' text = 'White quartz' qty = 120 ) ) ).\n|
         iv_notes =
|Cây sâu hơn 2 cấp: thêm tầng nested (xem EX11 — 3 tầng). \{\{text\}\} ở dòng\n| &&
           |subtotal đọc được field của nhóm nhờ context stack tìm từ trong ra ngoài.\n| ).

    add( iv_id = `EX16` iv_seq = 90 iv_ref = `4.16 / 4.19a / 4.08a`
         iv_title = `Matrix layout, data validation, khối điều kiện`
         iv_form  = `XLWB_EX16_ADVANCED`
         iv_purpose =
|Ba kỹ thuật nâng cao: mỗi ký tự một ô (biểu mẫu ô vuông), dropdown data\n| &&
           |validation với danh sách động, và ẩn/hiện khối theo điều kiện.\n|
         iv_context =
|DATA: BEGIN OF ls_ctx,\n| &&
           |        chars       TYPE ty_chars,   " tách từ chuỗi bằng to_chars( )\n| &&
           |        validlist   TYPE string,     " 'Draft,Released,Closed'\n| &&
           |        show_secret TYPE abap_bool,\n| &&
           |      END OF ls_ctx.\n| &&
           |ls_ctx-chars = zcl_xlwb_ex16_advanced=>to_chars(\n| &&
           |                 iv_text = '0101234567' iv_len = 10 ).\n|
         iv_excel =
|1. Mở Excel. Ô A1 gõ nhãn "Tax code:" (bôi đậm).\n| &&
           |2. Ô B1 là Ô VUÔNG MẪU cho một ký tự:\n| &&
           |   - độ rộng cột B khoảng 3, chiều cao dòng khoảng 22\n| &&
           |   - Format Cells > Alignment > Horizontal = Center\n| &&
           |   - Format Cells > Border > chọn Outline (cả 4 viền)\n| &&
           |   Chỉ vẽ MỘT ô — engine nhân bản đủ số ký tự.\n| &&
           |3. Ô dropdown (B2): Data > Data Validation > Allow = List,\n| &&
           |   Source tạm gõ "A,B,C" (giá trị thật do placeholder điền).\n| &&
           |4. Hai dòng cho khối điều kiện (dòng 3 và 4), mỗi dòng Merge A:D.\n| &&
           |5. File > Save As > XML Spreadsheet 2003.\n|
         iv_template =
|Mỗi ký tự một ô vuông (ô B1):\n| &&
           |    B1:  \{\{#chars>\}\}\{\{c\}\}\n| &&
           |\n| &&
           |Ô dropdown:\n| &&
           |    B2:  \{\{status\}\}\n| &&
           |\n| &&
           |Khối điều kiện theo dòng:\n| &&
           |    A3:  \{\{?show_secret\}\}CONFIDENTIAL: \{\{secret\}\}\{\{/show_secret\}\}\n| &&
           |    A4:  \{\{^show_secret\}\}(public copy)\{\{/show_secret\}\}\n| &&
           |\n| &&
           |Danh sách dropdown động — sửa trong file XML sau khi Save As, ở phần\n| &&
           |<DataValidation> của WorksheetOptions:\n| &&
           |    <Value>"\{\{validlist\}\}"</Value>\n|
         iv_engine =
|NEW zcl_xlwb_ex16_advanced( )->get_file(\n| &&
           |  iv_taxcode = '0101234567' iv_show_secret = abap_false ).\n|
         iv_notes =
|Placeholder NGOÀI vùng <Table> (như danh sách data validation trong\n| &&
           |WorksheetOptions) cũng được engine render.\n| &&
           |Excel không cho gõ placeholder trực tiếp vào ô Source của Data Validation, nên\n| &&
           |chỗ đó sửa trong file XML sau khi Save As.\n| ).

    add( iv_id = `EX17` iv_seq = 95 iv_ref = `4.17 / 4.18`
         iv_title = `Chart / dashboard — chart vẽ sẵn trong template`
         iv_form  = `XLWB_EX17_CHART`
         iv_purpose =
|Engine KHÔNG tạo chart lúc runtime. Cách đúng: vẽ chart TRONG template .xlsx\n| &&
           |trên vùng dữ liệu; engine XLSX bind dữ liệu vào vùng đó và GIỮ NGUYÊN chart.\n| &&
           |Mở file bằng Excel là chart tự cập nhật theo dữ liệu mới.\n|
         iv_context =
|DATA: BEGIN OF ls_ctx,\n| &&
           |        months TYPE ty_months,   " name + revenue\n| &&
           |      END OF ls_ctx.\n|
         iv_excel =
|1. Mở Excel, vẽ bảng dữ liệu 2 cột: Month / Revenue, MỘT dòng dữ liệu mẫu.\n| &&
           |2. Gõ vài dòng dữ liệu GIẢ tạm thời để vẽ được chart (sẽ xoá sau).\n| &&
           |3. Chọn vùng dữ liệu > Insert > Chart (Column/Line tuỳ nhu cầu).\n| &&
           |4. Xoá các dòng giả, CHỈ GIỮ MỘT dòng mẫu chứa placeholder.\n| &&
           |5. File > Save As > "Excel Workbook (*.xlsx)" — PHẢI là .xlsx thật.\n| &&
           |6. Upload vào app template: form XLWB_EX17_CHART, Engine = XLSX.\n|
         iv_template =
|Bảng dữ liệu (dòng 2 là dòng mẫu):\n| &&
           |    A2:  \{\{#months\}\}\{\{name\}\}      B2: \{\{revenue\}\}\{\{/months\}\}\n| &&
           |\n| &&
           |Chart không cần placeholder — nó tham chiếu vùng dữ liệu và được giữ nguyên.\n|
         iv_engine =
|NEW zcl_xlwb_ex17_chart( )->get_file( VALUE #(\n| &&
           |  ( name = 'Jan' revenue = '1200.00' )\n| &&
           |  ( name = 'Feb' revenue = '1550.50' ) ) ).\n|
         iv_notes =
|File mẫu tải ở đây là bản DỰ PHÒNG (SSML, chưa có chart) — vì chart chỉ tồn tại\n| &&
           |trong template .xlsx thật do Excel vẽ. Sau khi upload template .xlsx có chart,\n| &&
           |ví dụ này render ra đúng file có chart.\n| &&
           |Giới hạn: dữ liệu chart là vùng tĩnh trong sheet; không tạo series động.\n| ).

    add( iv_id = `EX20` iv_seq = 100 iv_ref = `4.06 / 4.07`
         iv_title = `Invoice có logo (engine XLSX)`
         iv_form  = `XLWB_EX20_LOGO`
         iv_purpose =
|File .xlsx thật có ẢNH NHÚNG — điều engine SpreadsheetML không làm được.\n| &&
           |Logo truyền vào dưới dạng base64 trong context.\n|
         iv_context =
|DATA: BEGIN OF ls_ctx,\n| &&
           |        logo  TYPE string,   " PNG dạng base64\n| &&
           |        invno TYPE string,\n| &&
           |        items TYPE ty_items,\n| &&
           |      END OF ls_ctx.\n| &&
           |\n| &&
           |" Lấy logo từ bảng Z / MIME repository, xem ZCL_QM_EXPORT_LOGO.\n| &&
           |" Đổi PNG sang base64 bằng PowerShell:\n| &&
           |"   [Convert]::ToBase64String([IO.File]::ReadAllBytes("logo.png"))\n|
         iv_excel =
|1. Mở Excel, vẽ hoá đơn: vùng logo ở góc trên trái (A1:B4), phần thông tin,\n| &&
           |   rồi bảng item.\n| &&
           |2. Chừa trống vùng A1:B4 cho logo — có thể đặt Fill Color xám nhạt để nhìn\n| &&
           |   thấy vùng khi thiết kế.\n| &&
           |3. KHÔNG cần chèn ảnh trong Excel: engine sẽ chèn ảnh vào đúng vùng đó.\n| &&
           |4. File > Save As > chọn "Excel Workbook (*.xlsx)" — KHÁC các ví dụ trên.\n| &&
           |   Engine XLSX cần template .xlsx thật, không phải XML Spreadsheet 2003.\n| &&
           |5. Upload file .xlsx vào app quản lý template với Engine = XLSX.\n|
         iv_template =
|Ô A1 — marker ảnh (rộng 2 cột, cao 4 dòng):\n| &&
           |    A1:  \{\{*image:logo@2x4\}\}\n| &&
           |\n| &&
           |Phần còn lại như bình thường:\n| &&
           |    A5:  INVOICE \{\{invno\}\}\n| &&
           |    A6:  Customer: \{\{customer\}\}\n| &&
           |    A7:  \{\{#items\}\}\{\{matnr\}\}\{\{/items\}\}      B7: \{\{amount\}\}\n| &&
           |    A8:  Total: \{\{sum:items.amount\}\}\n| &&
           |\n| &&
           |Đặt marker trong vùng lặp thì MỖI DÒNG một ảnh riêng.\n|
         iv_engine =
|" Engine XLSX được chọn tự động qua cột ENGINE của template,\n| &&
           |" hoặc chỉ định khi dùng bản dự phòng:\n| &&
           |zcl_xlwb_runtime=>render_prefer_table(\n| &&
           |  iv_form_name         = `XLWB_EX20_LOGO`\n| &&
           |  iv_fallback_template = get_template( )\n| &&
           |  iv_fallback_engine   = `XLSX`\n| &&
           |  ir_context           = REF #( ls_ctx ) ).\n|
         iv_notes =
|Engine XLSX v1 CHƯA có: lặp ô ngang, sheet động, \{\{*break\}\}. Cần những thứ đó\n| &&
           |thì dùng engine SpreadsheetML.\n| &&
           |Ngược lại, cần ảnh / chart / .xlsx thật thì dùng engine XLSX.\n| ).

    add( iv_id = `EX21` iv_seq = 110 iv_ref = `thực tế`
         iv_title = `Packing List HQ từ file Excel thật`
         iv_form  = `PACKING_LIST_HQ`
         iv_purpose =
|Bài quan trọng nhất: biến một form ECUS đang hardcode trong ABAP thành template\n| &&
           |mà người dùng tự sửa bằng Excel. File render đã đối chiếu từng dòng với file gốc\n| &&
           |và khớp hoàn toàn.\n|
         iv_context =
|" Context phẳng theo đúng các ô trên form\n| &&
           |TYPES: BEGIN OF ty_ctx,\n| &&
           |         shipper           TYPE ty_party,\n| &&
           |         buyer             TYPE ty_party,\n| &&
           |         invoice_no        TYPE string,\n| &&
           |         port_of_shipment  TYPE string,\n| &&
           |         container_summary TYPE string,\n| &&
           |         items             TYPE ty_items,\n| &&
           |       END OF ty_ctx.\n| &&
           |\n| &&
           |" Xem ZCL_XLWB_EX21_PACKLIST=>sample_context( ).\n|
         iv_excel =
|Ví dụ này KHÔNG vẽ mới — dùng lại chính file mà app hiện tại đang xuất:\n| &&
           |\n| &&
           |1. Vào app ECUS, xuất một Packing List HQ ra file (.xls).\n| &&
           |2. Mở file đó bằng Excel — đây đã là form hoàn chỉnh, đúng khung, đúng font.\n| &&
           |3. Xoá các dòng dữ liệu, CHỈ GIỮ MỘT dòng item làm mẫu.\n| &&
           |4. Gõ placeholder thay cho dữ liệu (xem bước sau).\n| &&
           |5. File > Save As > XML Spreadsheet 2003.\n| &&
           |6. Upload vào app quản lý template với form name PACKING_LIST_HQ.\n| &&
           |\n| &&
           |Cách này nhanh hơn vẽ lại: bạn thừa hưởng toàn bộ hình thức đã được nghiệp vụ\n| &&
           |chấp nhận, chỉ thay dữ liệu bằng placeholder.\n|
         iv_template =
|Thay dữ liệu bằng placeholder:\n| &&
           |    CASABLANCA JOINT STOCK COMPANY  →  \{\{shipper.name\}\}\n| &&
           |    16-Sep-2025                     →  \{\{invoice_date\}\}\n| &&
           |    FOB Haiphong, VIETNAM           →  \{\{port_of_shipment\}\}\n| &&
           |    0x CONTAINER                    →  \{\{container_summary\}\}\n| &&
           |\n| &&
           |Dòng item: ô đầu mở vùng lặp, ô cuối đóng vùng lặp\n| &&
           |    A15: \{\{#items\}\}\{\{@index\}\}\n| &&
           |    K15: \{\{container\}\}\{\{/items\}\}\n| &&
           |\n| &&
           |Dòng TOTAL: thay số bằng aggregate\n| &&
           |    F20: \{\{sum:items.qty_pcs\}\}\n| &&
           |\n| &&
           |GỘP Ô ĐỘNG — cột nào có giá trị lặp lại thì thêm \{\{*mergesame:field\}\}\n| &&
           |vào ĐẦU ô, engine sẽ tự gộp các dòng liên tiếp trùng giá trị:\n| &&
           |    D15: \{\{*mergesame:delivery_no\}\}\{\{delivery_no\}\}\n| &&
           |    E15: \{\{*mergesame:commodity\}\}\{\{commodity\}\}\n| &&
           |    K15: \{\{*mergesame:container\}\}\{\{container\}\}\{\{/items\}\}\n|
         iv_engine =
|" Template đã nạp vào bảng ZXLWB_TMPL nên code nghiệp vụ chỉ còn một dòng:\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |                  iv_form_name = 'PACKING_LIST_HQ'\n| &&
           |                  ir_context   = REF #( ls_ctx ) ).\n|
         iv_notes =
|CẠM BẪY LỚN NHẤT: file ECUS xuất ra dùng ss:Index trên <Row> để nhảy qua dòng\n| &&
           |trống. Đó là NEO DÒNG TUYỆT ĐỐI — nếu giữ lại, sau khi engine nhân bản vùng lặp\n| &&
           |thì mọi dòng phía sau nằm sai chỗ.\n| &&
           |\n| &&
           |Engine ĐÃ TỰ XỬ LÝ việc này (thay mỗi neo bằng đúng số dòng rỗng), nên bạn cứ để\n| &&
           |dòng trống thoải mái trong Excel. Lưu ý ss:Index trên <Cell> là neo CỘT và được\n| &&
           |giữ nguyên.\n| &&
           |\n| &&
           |Về \{\{*mergesame\}\}: nhóm được tính trên các dòng LIÊN TIẾP, nên phải SORT\n| &&
           |internal table theo đúng cột định gộp trước khi đưa vào context. Ô bị phủ được\n| &&
           |engine xoá nội dung — không tự viết giá trị vào đó.\n| ).

    add( iv_id = `EX22` iv_seq = 120 iv_ref = `mới`
         iv_title = `Gán Source class + Preview JSON (thay workbench gốc)`
         iv_form  = `XLWB_EX22_SOURCE`
         iv_purpose =
|Bài dành riêng cho app template mới: gán NGUỒN DỮ LIỆU cho form bằng class\n| &&
           |(thay màn hình gán context của tool gốc) và PREVIEW form bằng payload JSON\n| &&
           |ngay trong app — không cần chạy nghiệp vụ, không cần structure ABAP.\n|
         iv_context =
|" 1) Class nguồn dữ liệu — mỗi form nghiệp vụ một class:\n| &&
           |CLASS zcl_xlwb_src_order DEFINITION ... CREATE PUBLIC.\n| &&
           |  INTERFACES zif_xlwb_source.\n| &&
           |" get_context( iv_keys ): SELECT dữ liệu thật theo keys JSON\n| &&
           |" get_sample( ): dữ liệu demo cho Preview\n| &&
           |\n| &&
           |" 2) Payload JSON cho Preview (dán vào Sample JSON của template):\n| &&
           |\{ "orderno": "PO-2026-DEMO", "customer": "ACME Corp",\n| &&
           |  "orddate": "2026-08-25",\n| &&
           |  "items": [ \{ "pos": 10, "matnr": "SLAB-WHITE", "qty": 20,\n| &&
           |               "amount": 2100.00 \} ] \}\n|
         iv_excel =
|1. Vẽ form order như EX05, Save As XML Spreadsheet 2003.\n| &&
           |2. Mở app "XLWB Templates", tạo/sửa form XLWB_EX22_SOURCE:\n| &&
           |   - upload file template\n| &&
           |   - Source class = ZCL_XLWB_SRC_ORDER (validate lúc save)\n| &&
           |   - Sample JSON = payload ở bước trước (sai path là bị cảnh báo ngay)\n| &&
           |3. Bấm nút PREVIEW: bỏ trống popup = dùng Sample JSON; hoặc dán JSON khác\n| &&
           |   để thử nhánh dữ liệu rỗng, số dòng lớn...\n|
         iv_template =
|Template giống EX05 — điểm khác nằm ở NGUỒN dữ liệu:\n| &&
           |    kiểu JSON:  object → structure, array → bảng,\n| &&
           |                số → Number, "YYYY-MM-DD" → ngày (number format ăn)\n| &&
           |    tên field JSON = tên placeholder (không phân biệt hoa/thường)\n|
         iv_engine =
|" Render bằng dữ liệu THẬT từ source class đã gán:\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render_by_source(\n| &&
           |                  iv_form_name = 'XLWB_EX22_SOURCE'\n| &&
           |                  iv_keys      = '\{ "orderno": "PO-90001234" \}' ).\n| &&
           |\n| &&
           |" Xuất/nhập template giữa tenant: nút Export JSON / Import JSON trên app\n| &&
           |" (bảng Z không đi theo transport trên Public Cloud).\n|
         iv_notes =
|Preview render bằng ZCL_XLWB_CTX_JSON — JSON thành cây data động, không cần\n| &&
           |khai structure. Sai JSON hay sai placeholder đều báo lỗi kèm vị trí.\n| &&
           |Validate lúc save: cặp \{\{#\}\}/\{\{/\}\} lệch = LỖI chặn save; path không có\n| &&
           |trong Sample JSON = CẢNH BÁO.\n| ).

    add( iv_id = `EX23` iv_seq = 130 iv_ref = `thực tế`
         iv_title = `Xuất Excel app Quản Lý Đơn Hàng — cặp cột tuần động`
         iv_form  = `XLWB_BCNCTP_WEEKLY`
         iv_purpose =
|Bài thực tế thứ hai (sau Packing List): xuất báo cáo nhu cầu thành phẩm\n| &&
           |(app Quản Lý Đơn Hàng / ZCS_BCNCTP) dạng file WeekDailyProductivity:\n| &&
           |6 cột tĩnh + CẶP cột động theo tuần (Qty + CX chưa xác nhận).\n| &&
           |Dữ liệu THẬT lấy đúng nguồn của app: zcl_get_bcnctp=>get_bcnctp.\n|
         iv_context =
|" Source class ZCL_XLWB_SRC_BCNCTP đã implement zif_xlwb_source:\n| &&
           |" - get_context( iv_keys ): gọi zcl_get_bcnctp=>get_bcnctp với ranges\n| &&
           |"   từ keys JSON, rồi XEN KẼ W1..Wn Order/Unconfirmed thành bảng cells\n| &&
           |" - get_sample( ): dữ liệu demo tĩnh (chạy được cả client không có data)\n| &&
           |\n| &&
           |" Keys JSON (week/year bắt buộc):\n| &&
           |\{ "week": "35", "year": "2026", "plant": "6711", "weeks_count": 18 \}\n|
         iv_excel =
|1. Template có sẵn trong Maintain Repository (form XLWB_BCNCTP_WEEKLY,\n| &&
           |   Source class đã gán) — muốn đổi khung/màu thì tải về sửa rồi upload lại.\n| &&
           |2. Cột G là CỘT MẪU cho cả header tuần lẫn dữ liệu — engine nhân bản\n| &&
           |   2 x weeks_count cột (Qty và CX xen kẽ).\n| &&
           |3. Freeze panes: 6 cột trái + 3 dòng đầu đứng yên khi cuộn.\n|
         iv_template =
|Header tuần (G3) — MỘT cell loop, nhãn đã xen kẽ sẵn từ ABAP:\n| &&
           |    G3:  \{\{#weeks>\}\}\{\{label\}\}\n| &&
           |\n| &&
           |Dòng dữ liệu: 6 ô tĩnh + ma trận\n| &&
           |    A4:  \{\{#rows\}\}\{\{ph3\}\}   ...   F4: \{\{plant_name\}\}\n| &&
           |    G4:  \{\{#cells>\}\}\{\{val\}\}\n| &&
           |    H4:  \{\{/rows\}\}\n|
         iv_engine =
|" Gọi production — MỘT dòng, dữ liệu thật của app:\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render_by_source(\n| &&
           |  iv_form_name = 'XLWB_BCNCTP_WEEKLY'\n| &&
           |  iv_keys      = '\{ "week": "35", "year": "2026", "weeks_count": 18 \}' ).\n|
         iv_notes =
|CẶP cột động: 2 cell loop cạnh nhau sẽ ra "toàn bộ Qty rồi toàn bộ CX" —\n| &&
           |muốn XEN KẼ từng tuần thì xen kẽ ngay trong bảng context (ABAP làm 1 lần,\n| &&
           |template chỉ cần 1 cell loop). Đây là pattern chuẩn cho mọi báo cáo cặp cột.\n| &&
           |File mẫu ở đây render từ dữ liệu demo tĩnh; trên app thật dùng render_by_source.\n| ).

    " ================= Ví dụ Word (engine DOCX) =================

    add( iv_id = `DX00` iv_seq = 200 iv_ref = `DOCX`
         iv_title = `Word EX00 — Hello World (field đơn + checkbox)`
         iv_form  = `DOCX_EX00_HELLO`
         iv_purpose =
|Bản Word của EX00: giá trị đơn từ context vào content control.\n| &&
           |Khác Excel: control map vào CUSTOM XML PART (w:dataBinding) —\n| &&
           |engine chỉ thay data XML trong file, Word tự refresh khi mở,\n| &&
           |repeating section tự nhân bản theo số dòng.\n|
         iv_context =
|" Context = structure thường; component name = tên element (chữ thường)\n| &&
           |DATA: BEGIN OF ls_context,\n| &&
           |        title    TYPE string,\n| &&
           |        name     TYPE string,\n| &&
           |        docdate  TYPE string,\n| &&
           |        approved TYPE string,   " 'true'/'false' cho checkbox\n| &&
           |      END OF ls_context.\n|
         iv_excel =
|1. Tải template DOCX_EX00_HELLO từ Maintain Repository, mở bằng Word.\n| &&
           |2. Developer > XML Mapping Pane > chọn part urn:zdocx:data.\n| &&
           |3. Sửa layout tự do; thêm field = chuột phải node > Insert Content Control.\n|
         iv_template =
|Word KHÔNG dùng placeholder \{\{...\}\} — binding nằm trong control:\n| &&
           |  w:dataBinding xpath="/ns0:data/ns0:title"\n| &&
           |Checkbox: node value 'true'/'false'.\n|
         iv_engine =
|DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX00_HELLO'\n| &&
           |  ir_context   = REF #( ls_context ) ).\n| &&
           |" cột ENGINE = DOCX -> tự chọn zcl_xlwb_docx\n|
         iv_notes =
|Tạo template mới: app "XML Part từ DDIC" (nhập structure/bảng/table type)\n| &&
           |hoặc nút Prepare DOCX ngay trong Maintain Repository.\n| ).

    add( iv_id = `DX01` iv_seq = 201 iv_ref = `DOCX`
         iv_title = `Word EX01 — Công văn: path lồng nhau + multiline`
         iv_form  = `DOCX_EX01_LETTER`
         iv_purpose =
|Thư/công văn: field lồng trong structure con (customer/name, signer/title)\n| &&
           |và đoạn văn nhiều dòng (plain text control bật multiline).\n|
         iv_context =
|DATA: BEGIN OF ls_customer, name TYPE string, address TYPE string,\n| &&
           |        tax_code TYPE string, END OF ls_customer.\n| &&
           |" body_text chứa xuống dòng thật (\\n) -> Word giữ nguyên ngắt dòng\n|
         iv_excel =
|Map node con: mở rộng cây customer/signer trong XML Mapping Pane,\n| &&
           |chuột phải từng node lá > Insert Content Control > Plain Text.\n| &&
           |Với body_text: Properties control bật "Allow carriage returns".\n|
         iv_template =
|XPath lồng nhau: /ns0:data/ns0:customer/ns0:name\n| &&
           |Structure ABAP lồng bao nhiêu cấp cũng map được.\n|
         iv_engine =
|DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX01_LETTER' ir_context = REF #( ls_ctx ) ).\n|
         iv_notes =
|Giá trị mẫu trong template = tên field để designer nhìn là biết map gì.\n| ).

    add( iv_id = `DX02` iv_seq = 202 iv_ref = `DOCX`
         iv_title = `Word EX02 — Bảng lặp (Repeating Section)`
         iv_form  = `DOCX_EX02_ITEMS`
         iv_purpose =
|Bảng N dòng trong Word: Repeating Section Content Control bọc dòng mẫu,\n| &&
           |số dòng nhân bản theo số phần tử <row> trong data XML.\n|
         iv_context =
|" Component kiểu BẢNG -> element bao trùng tên + mỗi dòng <row>\n| &&
           |TYPES: BEGIN OF ty_item, stt TYPE string, product TYPE string,\n| &&
           |         unit TYPE string, qty TYPE string, price TYPE string,\n| &&
           |         amount TYPE string, END OF ty_item.\n| &&
           |DATA items TYPE STANDARD TABLE OF ty_item.   " -> <items><row>...\n|
         iv_excel =
|1. Vẽ bảng Word: 1 dòng header + 1 dòng mẫu.\n| &&
           |2. Chọn CẢ dòng mẫu, chuột phải node <row> > Insert > Repeating Section.\n| &&
           |3. Map từng cell vào field con của <row>.\n|
         iv_template =
|Repeating section bind xpath /ns0:data/ns0:items/ns0:row (không có [1]),\n| &&
           |cell bên trong bind field con. Bảng rỗng -> 0 dòng.\n|
         iv_engine =
|DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX02_ITEMS' ir_context = REF #( ls_ctx ) ).\n|
         iv_notes =
|Quy ước dòng lặp là element <row> — generator XML Part và engine dùng\n| &&
           |chung quy ước nên map một lần chạy mọi nơi.\n| ).

    add( iv_id = `DX05` iv_seq = 205 iv_ref = `DOCX`
         iv_title = `Word EX05 — Đơn hàng: header + bảng item + footer`
         iv_form  = `DOCX_EX05_ORDER`
         iv_purpose =
|Ví dụ tổng hợp tương đương EX05 bên Excel: field header, bảng item lặp,\n| &&
           |footer tổng tiền + bằng chữ, checkbox phê duyệt.\n|
         iv_context =
|" Trộn đủ loại: field đơn + structure con + bảng\n| &&
           |DATA: BEGIN OF ls_ctx, doc_no TYPE string, doc_date TYPE string,\n| &&
           |        customer TYPE ty_customer, items TYPE tt_item,\n| &&
           |        total_amount TYPE string, amount_in_words TYPE string,\n| &&
           |        approved TYPE string, approver TYPE string, END OF ls_ctx.\n|
         iv_excel =
|Template dựng sẵn trong Maintain Repository — tải về, sửa layout/logo\n| &&
           |trong Word rồi upload lại; binding giữ nguyên miễn không xoá control.\n|
         iv_template =
|Mọi kỹ thuật của DX00-DX02 gộp trong một form thật.\n|
         iv_engine =
|" Production: đọc dữ liệu -> đổ vào structure -> render 1 dòng\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX05_ORDER' ir_context = REF #( ls_ctx ) ).\n|
         iv_notes =
|File mẫu gen từ Sample JSON của template — đổi Sample JSON trong\n| &&
           |Maintain Repository rồi "Gen lại file mẫu" là thấy dữ liệu mới.\n| ).

    add( iv_id = `DX20` iv_seq = 220 iv_ref = `DOCX`
         iv_title = `Word EX20 — Nhãn slab có QR (picture control)`
         iv_form  = `DOCX_EX20_LABEL`
         iv_purpose =
|In nhãn có mã QR/barcode: Picture Content Control map vào node XML\n| &&
           |chứa ảnh PNG dạng BASE64 — Word tự hiện ảnh khi mở file.\n| &&
           |Kèm chữ xoay dọc (textDirection btLr) ở mép phải nhãn.\n|
         iv_context =
|" qr_png = chuỗi base64 của ảnh PNG (QR đã encode sẵn)\n| &&
           |DATA: BEGIN OF ls_ctx, product_title TYPE string, grade_letter TYPE string,\n| &&
           |        fn TYPE string, slab TYPE string, size TYPE string,\n| &&
           |        grade TYPE string, shade TYPE string,\n| &&
           |        batch TYPE string, stamp TYPE string,\n| &&
           |        qr_png TYPE string, END OF ls_ctx.\n|
         iv_excel =
|1. Chèn ảnh placeholder vào Word, bọc bằng Picture Content Control.\n| &&
           |2. XML Mapping Pane: chuột phải node qr_png > Insert > Picture.\n| &&
           |3. Chữ dọc: Table cell > Text Direction.\n|
         iv_template =
|Node bind của picture control chứa BASE64 PNG. Kích thước ảnh theo\n| &&
           |khung control trong template (2.5cm trong ví dụ này).\n|
         iv_engine =
|" QR sinh ở tầng cấp dữ liệu (app/CPI/service) -> đổ base64 vào context:\n| &&
           |ls_ctx-qr_png = lv_qr_base64.\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX20_LABEL' ir_context = REF #( ls_ctx ) ).\n|
         iv_notes =
|Engine không cần biết gì về QR — chỉ chép chuỗi base64 vào data XML.\n| &&
           |Barcode 1D (Code128/39) cũng làm y hệt, hoặc dùng font barcode\n| &&
           |cho field text thường.\n| ).

    add( iv_id = `DX21` iv_seq = 221 iv_ref = `DOCX`
         iv_title = `Word EX21 — Letterhead: LOGO ĐỘNG từ ZXLWB_TMPL`
         iv_form  = `DOCX_EX21_LOGO`
         iv_purpose =
|In logo công ty trong Word, logo KHÔNG nằm cứng trong template mà lưu\n| &&
           |một chỗ trong ZXLWB_TMPL (form LOGO_DEMO, ENGINE=ASST) — đổi logo\n| &&
           |một nơi, mọi form dùng chung đổi theo.\n|
         iv_context =
|" Chọn logo động lúc in:\n| &&
           |ls_ctx-logo = zcl_xlwb_runtime=>get_asset_b64( 'LOGO_DEMO' ).\n| &&
           |" logo là base64 PNG -> bind vào Picture Content Control như DX20\n|
         iv_excel =
|1. Upload logo PNG vào Maintain Repository: FormName=LOGO_DEMO, Engine=ASST.\n| &&
           |2. Template Word có Picture Content Control map node logo.\n| &&
           |3. Muốn đổi logo: upload PNG mới đè lên LOGO_DEMO — xong.\n|
         iv_template =
|Giống DX20: node <logo> chứa base64. Template chỉ giữ khung ảnh,\n| &&
           |nội dung ảnh do get_asset_b64 quyết định lúc render.\n|
         iv_engine =
|ls_ctx-logo = zcl_xlwb_runtime=>get_asset_b64( 'LOGO_DEMO' ).\n| &&
           |DATA(ls_file) = zcl_xlwb_runtime=>render(\n| &&
           |  iv_form_name = 'DOCX_EX21_LOGO' ir_context = REF #( ls_ctx ) ).\n|
         iv_notes =
|Excel dùng chung asset này: EX20 (marker \{\{*image:logo\}\}) cũng lấy logo qua\n| &&
           |get_asset_b64. SSML thì KHÔNG nhúng được ảnh — xem form XLWB_EX21_SSML.\n| ).

    add( iv_id = `DS21` iv_seq = 230 iv_ref = `SSML`
         iv_title = `SSML + logo — GIỚI HẠN: SpreadsheetML không nhúng ảnh`
         iv_form  = `XLWB_EX21_SSML`
         iv_purpose =
|Ví dụ ĐỐI CHIẾU: cùng một marker \{\{*image:logo\}\} nhưng engine SSML\n| &&
           |(XML Spreadsheet 2003) KHÔNG nhúng được ảnh — định dạng không có\n| &&
           |phần drawing/media. Muốn có logo trong Excel: dùng engine XLSX (EX20).\n| &&
           |Trong Word: dùng picture content control (DX21).\n|
         iv_context =
|" Context giống EX20 — logo base64 lấy động từ asset LOGO_DEMO\n| &&
           |ls_ctx-logo = zcl_xlwb_runtime=>get_asset_b64( 'LOGO_DEMO' ).\n|
         iv_excel =
|Template .xls (SpreadsheetML) chỉ có ô text — không có chỗ chứa ảnh.\n|
         iv_template =
|A1: \{\{*image:logo@4x6\}\}   <- SSML bỏ qua, KHÔNG ra ảnh\n| &&
           |E1: \{\{company\}\}  E2: \{\{slogan\}\}\n|
         iv_engine =
|" Cùng một dòng gọi, chỉ khác cột ENGINE của template:\n| &&
           |zcl_xlwb_runtime=>render( iv_form_name = 'XLWB_EX21_SSML' ... )  " SSML: không logo\n| &&
           |zcl_xlwb_runtime=>render( iv_form_name = 'XLWB_EX21_LOGO' ... )  " XLSX: CÓ logo\n|
         iv_notes =
|KẾT LUẬN đã kiểm chứng: logo/ảnh chỉ khả thi với XLSX (Excel) và DOCX (Word).\n| &&
           |SSML chỉ dùng khi cần file nhẹ, không ảnh. Bảng ZXLWB_TMPL giữ logo một\n| &&
           |chỗ (form LOGO_DEMO, ENGINE=ASST) nên đổi logo là mọi form đổi theo.\n| ).
  ENDMETHOD.


  METHOD load.
    build( ).
    DELETE FROM zxlwb_demo.
    INSERT zxlwb_demo FROM TABLE @mt_rows.
    rv_count = sy-dbcnt.
    COMMIT WORK AND WAIT.
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    DATA(lv_count) = load( ).
    out->write( |Đã nạp { lv_count } ví dụ vào ZXLWB_DEMO.| ).
    out->write( |—| ).

    SELECT example_id, title, engine_type, template_form, file_name,
           CAST( LENGTH( content ) AS INT4 ) AS size
      FROM zxlwb_demo
      ORDER BY example_id
      INTO TABLE @DATA(lt_check).

    LOOP AT lt_check INTO DATA(ls).
      DATA(lv_src) = COND string(
        WHEN zcl_xlwb_runtime=>exists_in_table( CONV string( ls-template_form ) ) = abap_true
        THEN |ZXLWB_TMPL| ELSE |code| ).
      out->write( |{ ls-example_id } { ls-engine_type } | &&
                  |{ ls-file_name WIDTH = 22 } { ls-size WIDTH = 6 } byte  | &&
                  |template: { lv_src WIDTH = 11 }  { ls-title }| ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
