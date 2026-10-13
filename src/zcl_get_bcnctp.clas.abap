"! <p class="shorttext synchronized" lang="en">MOCK BCNCTP: finished-goods demand by week</p>
"!
"! Bản MÔ PHỎNG của ZCL_GET_BCNCTP — nguồn dữ liệu app "Quản Lý Đơn Hàng"
"! (ZCS_BCNCTP, Báo Cáo Nhu Cầu Thành Phẩm) — để ZPK_XLWB_POC activate và
"! chạy được trên hệ thống không có dữ liệu bán hàng (BTP trial).
"!
"! Giữ đúng interface mà ZCL_XLWB_SRC_BCNCTP đang gọi:
"! <pre>
"! zcl_get_bcnctp=>get_bcnctp( EXPORTING ir_hierarchy3 = ... ir_plant = ...
"!                                       ir_week = ...       ir_year = ...
"!                             IMPORTING et_data = DATA(lt_data) ).
"! </pre>
"! Mỗi dòng et_data:
"! <ul>
"! <li>dòng CHA: PH3 + Plant, PH4 rỗng — số lượng = tổng các dòng con</li>
"! <li>dòng CON: PH3 + PH4 + Plant</li>
"! <li>W{n}ORDERQUANTITY / W{n}UNCONFIRMEDQUANTITY (n = 1..54): số lượng đặt
"!     và phần chưa xác nhận (CX) của tuần thứ n, W1 = tuần trong ir_week/ir_year</li>
"! </ul>
"! Số liệu sinh giả, cố định theo (PH4, Plant, tuần, năm): gọi lại cho cùng kết quả.
CLASS zcl_get_bcnctp DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    CONSTANTS c_weeks TYPE i VALUE 54.

    TYPES ty_range_low TYPE c LENGTH 40.
    TYPES tt_ranges    TYPE RANGE OF ty_range_low.

    TYPES: BEGIN OF gty_line,
             producthierarchy3     TYPE c LENGTH 18,
             producthierarchy3name TYPE c LENGTH 40,
             producthierarchy4     TYPE c LENGTH 18,
             producthierarchy4name TYPE c LENGTH 40,
             plant                 TYPE c LENGTH 4,
             plantname             TYPE c LENGTH 30,
             w1orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w1unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w2orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w2unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w3orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w3unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w4orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w4unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w5orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w5unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w6orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w6unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w7orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w7unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w8orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w8unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w9orderquantity        TYPE p LENGTH 15 DECIMALS 3,
             w9unconfirmedquantity  TYPE p LENGTH 15 DECIMALS 3,
             w10orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w10unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w11orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w11unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w12orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w12unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w13orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w13unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w14orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w14unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w15orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w15unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w16orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w16unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w17orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w17unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w18orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w18unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w19orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w19unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w20orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w20unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w21orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w21unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w22orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w22unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w23orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w23unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w24orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w24unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w25orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w25unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w26orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w26unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w27orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w27unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w28orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w28unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w29orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w29unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w30orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w30unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w31orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w31unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w32orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w32unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w33orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w33unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w34orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w34unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w35orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w35unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w36orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w36unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w37orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w37unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w38orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w38unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w39orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w39unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w40orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w40unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w41orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w41unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w42orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w42unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w43orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w43unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w44orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w44unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w45orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w45unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w46orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w46unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w47orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w47unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w48orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w48unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w49orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w49unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w50orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w50unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w51orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w51unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w52orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w52unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w53orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w53unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
             w54orderquantity       TYPE p LENGTH 15 DECIMALS 3,
             w54unconfirmedquantity TYPE p LENGTH 15 DECIMALS 3,
           END OF gty_line,
           gty_bcnctp TYPE STANDARD TABLE OF gty_line WITH EMPTY KEY.

    "! Nhu cầu thành phẩm theo tuần, lọc theo PH3 / Plant; tuần bắt đầu lấy
    "! từ dòng đầu của ir_week / ir_year (rỗng = tuần hiện tại)
    CLASS-METHODS get_bcnctp
      IMPORTING ir_hierarchy3 TYPE tt_ranges OPTIONAL
                ir_plant      TYPE tt_ranges OPTIONAL
                ir_week       TYPE tt_ranges OPTIONAL
                ir_year       TYPE tt_ranges OPTIONAL
      EXPORTING et_data       TYPE gty_bcnctp.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_product,
             ph3      TYPE c LENGTH 18,
             ph3_name TYPE c LENGTH 40,
             ph4      TYPE c LENGTH 18,
             ph4_name TYPE c LENGTH 40,
           END OF ty_product,
           ty_products TYPE STANDARD TABLE OF ty_product WITH EMPTY KEY,
           BEGIN OF ty_plant,
             plant TYPE c LENGTH 4,
             name  TYPE c LENGTH 30,
           END OF ty_plant,
           ty_plants TYPE STANDARD TABLE OF ty_plant WITH EMPTY KEY.

    CLASS-METHODS start_period
      IMPORTING ir_week TYPE tt_ranges
                ir_year TYPE tt_ranges
      EXPORTING ev_week TYPE i
                ev_year TYPE i.

    "! Sinh số giả cố định 0..32767 — cùng seed cho cùng kết quả
    CLASS-METHODS pseudo
      IMPORTING iv_seed         TYPE int8
      RETURNING VALUE(rv_value) TYPE int8.

    "! Cộng dồn W1..W54 của dòng con vào dòng cha
    CLASS-METHODS add_weeks
      IMPORTING is_child  TYPE gty_line
      CHANGING  cs_parent TYPE gty_line.
ENDCLASS.



CLASS zcl_get_bcnctp IMPLEMENTATION.

  METHOD get_bcnctp.
    CLEAR et_data.

    start_period( EXPORTING ir_week = ir_week
                            ir_year = ir_year
                  IMPORTING ev_week = DATA(lv_start_week)
                            ev_year = DATA(lv_start_year) ).

    " master data mô phỏng — mã M50304 / M50602 / 6711 lấy theo sample của
    " ZCL_XLWB_SRC_BCNCTP, phần còn lại là dữ liệu giả
    DATA(lt_product) = VALUE ty_products(
      ( ph3 = 'M50304' ph3_name = 'Not use'         ph4 = 'M50304002' ph4_name = 'Not use' )
      ( ph3 = 'M50304' ph3_name = 'Not use'         ph4 = 'M50304005' ph4_name = 'Mock PH4 M50304005' )
      ( ph3 = 'M50602' ph3_name = 'Túi dán Tshirt'  ph4 = 'M50602014' ph4_name = 'Not use' )
      ( ph3 = 'M50602' ph3_name = 'Túi dán Tshirt'  ph4 = 'M50602020' ph4_name = 'Mock PH4 M50602020' )
      ( ph3 = 'M50701' ph3_name = 'Mock PH3 M50701' ph4 = 'M50701003' ph4_name = 'Mock PH4 M50701003' ) ).

    DATA(lt_plant) = VALUE ty_plants(
      ( plant = '6711' name = 'Nhà máy túi CASLA 1' )
      ( plant = '6712' name = 'Mock plant 6712' ) ).

    LOOP AT lt_plant INTO DATA(ls_plant) WHERE plant IN ir_plant.
      LOOP AT lt_product INTO DATA(ls_prod) WHERE ph3 IN ir_hierarchy3.

        DATA(ls_child) = VALUE gty_line(
          producthierarchy3     = ls_prod-ph3
          producthierarchy3name = ls_prod-ph3_name
          producthierarchy4     = ls_prod-ph4
          producthierarchy4name = ls_prod-ph4_name
          plant                 = ls_plant-plant
          plantname             = ls_plant-name ).

        DATA(lv_week) = lv_start_week.
        DATA(lv_year) = lv_start_year.
        DO c_weeks TIMES.
          DATA(lv_ix) = sy-index.

          " seed theo (PH4, Plant, năm, tuần thực) — đổi tuần bắt đầu thì
          " cột W1 dịch theo, số của 1 tuần cụ thể vẫn giữ nguyên
          DATA(lv_seed) = CONV int8( ls_prod-ph4+1(8) ) * 7919
                        + CONV int8( ls_plant-plant ) * 104729
                        + lv_year * 53 + lv_week.
          DATA(lv_r1) = pseudo( lv_seed ).
          DATA(lv_r2) = pseudo( lv_r1 ).

          " ~30% tuần không có đơn; còn lại 100..80.000, chẵn trăm
          DATA(lv_qty) = COND int8( WHEN lv_r1 MOD 10 < 3 THEN 0
                                    ELSE ( lv_r1 MOD 800 + 1 ) * 100 ).
          " tuần càng xa càng dễ có phần chưa xác nhận (10..49% số đặt)
          DATA(lv_cx) = COND int8(
            WHEN lv_qty > 0 AND lv_r2 MOD 54 < lv_ix
            THEN lv_qty * ( lv_r2 MOD 40 + 10 ) DIV 100
            ELSE 0 ).

          ASSIGN COMPONENT |W{ lv_ix }ORDERQUANTITY| OF STRUCTURE ls_child
            TO FIELD-SYMBOL(<lv_qty>).
          IF sy-subrc = 0.
            <lv_qty> = lv_qty.
          ENDIF.
          ASSIGN COMPONENT |W{ lv_ix }UNCONFIRMEDQUANTITY| OF STRUCTURE ls_child
            TO FIELD-SYMBOL(<lv_cx>).
          IF sy-subrc = 0.
            <lv_cx> = lv_cx.
          ENDIF.

          " cùng quy ước với ZCL_XLWB_SRC_BCNCTP: năm 52 tuần
          lv_week += 1.
          IF lv_week > 52.
            lv_week = 1.
            lv_year += 1.
          ENDIF.
        ENDDO.

        " dòng cha (PH4 rỗng): tạo ở PH4 đầu tiên của nhóm, các con cộng dồn vào
        READ TABLE et_data ASSIGNING FIELD-SYMBOL(<ls_parent>)
          WITH KEY producthierarchy3 = ls_prod-ph3
                   plant             = ls_plant-plant
                   producthierarchy4 = ``.
        IF sy-subrc <> 0.
          APPEND VALUE #( producthierarchy3     = ls_child-producthierarchy3
                          producthierarchy3name = ls_child-producthierarchy3name
                          plant                 = ls_child-plant
                          plantname             = ls_child-plantname )
            TO et_data ASSIGNING <ls_parent>.
        ENDIF.
        add_weeks( EXPORTING is_child = ls_child CHANGING cs_parent = <ls_parent> ).

        APPEND ls_child TO et_data.
      ENDLOOP.
    ENDLOOP.

    " cây như app: theo PH3, Plant; dòng cha (PH4 rỗng) đứng trước các con
    SORT et_data BY producthierarchy3 plant producthierarchy4.
  ENDMETHOD.


  METHOD start_period.
    " mặc định: tuần hiện tại (đếm 7 ngày từ 01/01, không theo ISO 8601)
    DATA(lv_today) = cl_abap_context_info=>get_system_date( ).
    DATA(lv_jan1)  = CONV d( |{ lv_today(4) }0101| ).
    ev_year = CONV i( lv_today(4) ).
    ev_week = ( lv_today - lv_jan1 ) DIV 7 + 1.

    TRY.
        READ TABLE ir_week INTO DATA(ls_week) INDEX 1.
        IF sy-subrc = 0 AND ls_week-low IS NOT INITIAL.
          ev_week = CONV i( condense( ls_week-low ) ).
        ENDIF.
        READ TABLE ir_year INTO DATA(ls_year) INDEX 1.
        IF sy-subrc = 0 AND ls_year-low IS NOT INITIAL.
          ev_year = CONV i( condense( ls_year-low ) ).
        ENDIF.
      CATCH cx_sy_conversion_error.
        " tuần/năm truyền vào không phải số: giữ tuần hiện tại
    ENDTRY.

    IF ev_week < 1 OR ev_week > 52.
      ev_week = 1.
    ENDIF.
  ENDMETHOD.


  METHOD pseudo.
    " 1. bình phương mod số nguyên tố 2^31-1: phá tính tuyến tính giữa các
    "    seed liền nhau (tuần kế tiếp, PH4 kế tiếp)
    " 2. LCG (hằng số glibc), lấy 15 bit cao (0..32767) vì bit thấp lặp chu kỳ ngắn
    " Mọi phép nhân đều < 4,7E18 nên không tràn int8
    DATA(lv_x) = iv_seed MOD 2147483647.
    lv_x = lv_x * lv_x MOD 2147483647.
    rv_value = ( lv_x * 1103515245 + 12345 ) MOD 2147483648 DIV 65536.
  ENDMETHOD.


  METHOD add_weeks.
    DO c_weeks TIMES.
      DATA(lv_ix) = sy-index.
      ASSIGN COMPONENT |W{ lv_ix }ORDERQUANTITY| OF STRUCTURE is_child  TO FIELD-SYMBOL(<lv_src_qty>).
      ASSIGN COMPONENT |W{ lv_ix }ORDERQUANTITY| OF STRUCTURE cs_parent TO FIELD-SYMBOL(<lv_dst_qty>).
      IF <lv_src_qty> IS ASSIGNED AND <lv_dst_qty> IS ASSIGNED.
        <lv_dst_qty> += <lv_src_qty>.
      ENDIF.
      ASSIGN COMPONENT |W{ lv_ix }UNCONFIRMEDQUANTITY| OF STRUCTURE is_child  TO FIELD-SYMBOL(<lv_src_cx>).
      ASSIGN COMPONENT |W{ lv_ix }UNCONFIRMEDQUANTITY| OF STRUCTURE cs_parent TO FIELD-SYMBOL(<lv_dst_cx>).
      IF <lv_src_cx> IS ASSIGNED AND <lv_dst_cx> IS ASSIGNED.
        <lv_dst_cx> += <lv_src_cx>.
      ENDIF.
      UNASSIGN: <lv_src_qty>, <lv_dst_qty>, <lv_src_cx>, <lv_dst_cx>.
    ENDDO.
  ENDMETHOD.

ENDCLASS.
