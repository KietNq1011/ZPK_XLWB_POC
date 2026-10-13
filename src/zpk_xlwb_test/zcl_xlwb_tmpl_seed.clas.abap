"! <p class="shorttext synchronized" lang="en">XLWB: seed template repository (F9)</p>
"!
"! Bấm F9 để nạp 16 template mẫu vào bảng ZXLWB_TMPL — nguồn dữ liệu của app
"! "Maintain Repository". Template lấy từ chính get_template( ) của các example
"! class nên luôn khớp bản đã unit test. Chạy lại = ghi đè (upsert); form do
"! người dùng tự upload với tên khác không bị đụng.
CLASS zcl_xlwb_tmpl_seed DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.

    "! Nạp/làm mới toàn bộ template mẫu, trả về số dòng đã ghi
    CLASS-METHODS load
      RETURNING VALUE(rv_count) TYPE i.

  PRIVATE SECTION.
    CLASS-DATA mt_rows TYPE STANDARD TABLE OF zxlwb_tmpl WITH EMPTY KEY.

    CLASS-METHODS add
      IMPORTING iv_form         TYPE string
                iv_descr        TYPE string
                iv_engine       TYPE string DEFAULT `SSML`
                iv_template     TYPE xstring
                iv_source_class TYPE string OPTIONAL
                iv_sample_json  TYPE string OPTIONAL.

    CLASS-METHODS utf8
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_bin) TYPE xstring.

    CLASS-METHODS build.
ENDCLASS.



CLASS ZCL_XLWB_TMPL_SEED IMPLEMENTATION.


  METHOD add.
    DATA ls TYPE zxlwb_tmpl.

    ls-form_name    = iv_form.
    ls-descr        = iv_descr.
    ls-engine       = iv_engine.
    ls-mime_type    = COND #( WHEN iv_engine = `XLSX`
      THEN `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
      ELSE `application/vnd.ms-excel` ).
    ls-file_name    = |{ to_lower( iv_form ) }.| &&
                      COND #( WHEN iv_engine = `XLSX` THEN `xlsx` ELSE `xls` ).
    ls-is_active    = abap_true.
    ls-template     = iv_template.
    ls-source_class = iv_source_class.
    ls-sample_json  = iv_sample_json.
    ls-created_by   = cl_abap_context_info=>get_user_technical_name( ).
    GET TIME STAMP FIELD ls-created_at.
    ls-last_changed_by  = ls-created_by.
    ls-last_changed_at  = ls-created_at.
    ls-last_changed_glo = ls-created_at.

    APPEND ls TO mt_rows.
  ENDMETHOD.


  METHOD utf8.
    rv_bin = cl_abap_conv_codepage=>create_out( )->convert( iv_text ).
  ENDMETHOD.


  METHOD build.
    CLEAR mt_rows.

    add( iv_form = `XLWB_EX00_HELLO`     iv_descr = `EX00: Hello World — giá trị đơn`
         iv_template = utf8( NEW zcl_xlwb_ex00_hello( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX01_LABEL`     iv_descr = `EX01: Shipping label — path lồng nhau`
         iv_template = utf8( NEW zcl_xlwb_ex01_label( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX02_LABELS`    iv_descr = `EX02: N nhãn + page break`
         iv_template = utf8( NEW zcl_xlwb_ex02_labels( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX02A_HLABELS`  iv_descr = `EX02A: N nhãn ngang (cell loop)`
         iv_template = utf8( NEW zcl_xlwb_ex02a_hlabels( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX03_SHEETS`    iv_descr = `EX03: mỗi bản ghi một worksheet`
         iv_template = utf8( NEW zcl_xlwb_ex03_sheets( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX04_BUNDLE`    iv_descr = `EX04: workbook nhiều form`
         iv_template = utf8( NEW zcl_xlwb_ex04_bundle( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX05_ORDER`     iv_descr = `EX05: Order form Header/Item/Footer`
         iv_template = utf8( NEW zcl_xlwb_ex05_order( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX08_DYNCOLS`   iv_descr = `EX08: bảng số cột động (ma trận)`
         iv_template = utf8( NEW zcl_xlwb_ex08_dyncols( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX08G_GROUPING` iv_descr = `EX08G: row grouping (+/- outline)`
         iv_template = utf8( NEW zcl_xlwb_ex08g_grouping( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX09_DYNTABLE`  iv_descr = `EX09: bảng động dòng x cột runtime + tổng 2 chiều`
         iv_template = utf8( NEW zcl_xlwb_ex09_dyntable( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX11_MULTILEVEL` iv_descr = `EX11: 3 cấp + subtotal + gộp ô dọc`
         iv_template = utf8( NEW zcl_xlwb_ex11_multilevel( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX13_GANTT`     iv_descr = `EX13: Gantt header 3 cấp gộp ngang`
         iv_template = utf8( NEW zcl_xlwb_ex13_gantt( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX14_TREE`      iv_descr = `EX14: cây phân cấp -> nested table`
         iv_template = utf8( NEW zcl_xlwb_ex14_tree( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX16_ADVANCED`  iv_descr = `EX16: matrix + validation + điều kiện`
         iv_template = utf8( NEW zcl_xlwb_ex16_advanced( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX17_CHART`     iv_descr = `EX17: chart trong template (upload .xlsx có chart để thay)`
         iv_template = utf8( NEW zcl_xlwb_ex17_chart( )->get_template( ) ) ).

    " EX20: template là .xlsx THẬT dựng bằng XCO — engine XLSX
    add( iv_form = `XLWB_EX20_LOGO`      iv_descr = `EX20: invoice có logo (engine XLSX)`
         iv_engine = `XLSX`
         iv_template = NEW zcl_xlwb_ex20_logo( )->get_template( ) ).

    add( iv_form = `PACKING_LIST_HQ`     iv_descr = `EX21: Packing List HQ (form ECUS thật)`
         iv_template = utf8( zcl_xlwb_ex21_packlist=>get_template( ) ) ).

    add( iv_form = `XLWB_BCNCTP_WEEKLY`  iv_descr = `BCNCTP: xuất Excel app Quản Lý Đơn Hàng (cột tuần động)`
         iv_source_class = `ZCL_XLWB_SRC_BCNCTP`
         iv_sample_json  =
           `{ "title": "BÁO CÁO NHU CẦU THÀNH PHẨM — TỪ TUẦN W35/2026", "week": "35", "year": "2026", "lastcol": 10,` &&
           ` "weeks": [ { "label": "W35/2026" }, { "label": "CX W35/2026" },` &&
           ` { "label": "W36/2026" }, { "label": "CX W36/2026" } ],` &&
           ` "rows": [ { "ph3": "M50304", "ph3_name": "Not use", "ph4": "", "ph4_name": "",` &&
           ` "plant": "6711", "plant_name": "Nhà máy túi CASLA 1",` &&
           ` "cells": [ { "val": 65123 }, { "val": 0 }, { "val": 500 }, { "val": 100 } ] } ] }`
         iv_template = utf8( NEW zcl_xlwb_src_bcnctp( )->get_template( ) ) ).

    add( iv_form = `XLWB_EX22_SOURCE`    iv_descr = `EX22: gán source class + preview JSON`
         iv_source_class = `ZCL_XLWB_SRC_ORDER`
         iv_sample_json  = `{ "orderno": "PO-2026-DEMO", "customer": "ACME Corp",` &&
                           ` "orddate": "2026-08-26",` &&
                           ` "items": [ { "pos": 10, "matnr": "SLAB-WHITE-3CM", "qty": 20, "amount": 2100.00 },` &&
                           ` { "pos": 20, "matnr": "SLAB-GREY-2CM", "qty": 15, "amount": 1327.50 } ] }`
         iv_template = utf8( NEW zcl_xlwb_src_order( )->get_template( ) ) ).
  ENDMETHOD.


  METHOD load.
    build( ).
    MODIFY zxlwb_tmpl FROM TABLE @mt_rows.
    rv_count = sy-dbcnt.
    COMMIT WORK AND WAIT.
    zcl_xlwb_runtime=>clear_cache( ).
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    DATA(lv_count) = load( ).
    out->write( |Đã nạp { lv_count } template vào ZXLWB_TMPL (app Maintain Repository).| ).
    out->write( |—| ).

    SELECT form_name, engine, file_name, source_class,
           CAST( LENGTH( template ) AS INT4 ) AS size
      FROM zxlwb_tmpl
      ORDER BY form_name
      INTO TABLE @DATA(lt_check).

    LOOP AT lt_check INTO DATA(ls).
      out->write( |{ ls-form_name WIDTH = 22 } { ls-engine WIDTH = 5 } | &&
                  |{ ls-size WIDTH = 7 } byte  { ls-file_name WIDTH = 26 } { ls-source_class }| ).
    ENDLOOP.
    out->write( |—| ).
    out->write( |Tiếp theo: mở app XLWB Examples, bấm "Gen lại tất cả" để file mẫu render theo template trong bảng.| ).
  ENDMETHOD.
ENDCLASS.
