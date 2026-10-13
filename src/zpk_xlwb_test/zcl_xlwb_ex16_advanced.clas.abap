"! <p class="shorttext synchronized" lang="en">XLWB ex 4.16/4.19a/4.08a: Matrix layout, validation, conditions</p>
"!
"! Gom 4 ky thuat nang cao trong 1 form mau:
"! - 4.16 Matrix layout: moi KY TU cua 1 gia tri vao 1 o rieng (bieu mau
"!   hanh chinh o vuong). Engine khong co option rieng - dev tach chuoi
"!   thanh bang ky tu (helper to_chars) roi dung cell loop {{#chars>}}.
"! - 4.19a Data Validation: khai trong template (WorksheetOptions/DataValidation);
"!   engine giu nguyen. Danh sach gia tri co the DONG bang placeholder
"!   trong the <Value> (o day dung {{validlist}}).
"! - 4.08a An/hien cot & khoi: block {{?flag}}..{{/flag}} va {{^flag}}..{{/flag}}
"!   (theo DONG). An theo COT thi khong dua cot do vao bang cell loop.
"! - Attribute passthrough: moi thuoc tinh la cua template tren Row/Cell
"!   (StyleID, Height, MergeAcross tinh, x:...) deu duoc giu khi engine clone.
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX16_ADVANCED`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex16_advanced DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_char,
             c TYPE string,
           END OF ty_char,
           ty_chars TYPE STANDARD TABLE OF ty_char WITH EMPTY KEY.

    "! Tach chuoi thanh bang ky tu cho Matrix layout (4.16)
    CLASS-METHODS to_chars
      IMPORTING iv_text         TYPE string
                iv_len          TYPE i OPTIONAL
      RETURNING VALUE(rt_chars) TYPE ty_chars.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING iv_taxcode     TYPE string
                iv_show_secret TYPE abap_bool
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX16_ADVANCED IMPLEMENTATION.


  METHOD to_chars.
    DATA(lv_n) = COND i( WHEN iv_len > 0 THEN iv_len ELSE strlen( iv_text ) ).
    DO lv_n TIMES.
      DATA(lv_off) = sy-index - 1.
      APPEND VALUE ty_char(
        c = COND string( WHEN lv_off < strlen( iv_text )
                         THEN substring( val = iv_text off = lv_off len = 1 )
                         ELSE `` ) ) TO rt_chars.
    ENDDO.
  ENDMETHOD.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="box"><Alignment ss:Horizontal="Center"/>` &&
      `<Borders>` &&
      `<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>` &&
      `</Borders></Style>` &&
      `<Style ss:ID="lbl"><Font ss:Bold="1"/></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Form">` &&
      `<Table>` &&
      " --- 4.16 matrix layout: 1 ky tu / 1 o -------------------------------
      `<Row ss:AutoFitHeight="0" ss:Height="22">` &&
      `<Cell ss:StyleID="lbl"><Data ss:Type="String">Tax code:</Data></Cell>` &&
      `<Cell ss:StyleID="box"><Data ss:Type="String">{{#chars>}}{{c}}</Data></Cell></Row>` &&
      " --- 4.19a data validation: danh sach dong ---------------------------
      `<Row><Cell ss:StyleID="lbl"><Data ss:Type="String">Status:</Data></Cell>` &&
      `<Cell ss:StyleID="box"><Data ss:Type="String">{{status}}</Data></Cell></Row>` &&
      " --- 4.08a khoi co dieu kien -----------------------------------------
      `<Row><Cell ss:MergeAcross="3">` &&
      `<Data ss:Type="String">{{?show_secret}}CONFIDENTIAL: {{secret}}{{/show_secret}}</Data></Cell></Row>` &&
      `<Row><Cell ss:MergeAcross="3">` &&
      `<Data ss:Type="String">{{^show_secret}}(public copy){{/show_secret}}</Data></Cell></Row>` &&
      `</Table>` &&
      `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<DataValidation xmlns="urn:schemas-microsoft-com:office:excel">` &&
      `<Range>R2C2</Range><Type>List</Type>` &&
      `<Value>&quot;{{validlist}}&quot;</Value>` &&
      `<ErrorStyle>Stop</ErrorStyle>` &&
      `<ErrorMessage>Choose a value from the list</ErrorMessage>` &&
      `</DataValidation>` &&
      `</WorksheetOptions>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            chars       TYPE ty_chars,
            status      TYPE string,
            validlist   TYPE string,
            show_secret TYPE abap_bool,
            secret      TYPE string,
          END OF ls_context.

    " 10 o co dinh - chuoi ngan hon thi cac o cuoi de trong
    ls_context-chars       = to_chars( iv_text = iv_taxcode iv_len = 10 ).
    ls_context-status      = 'Draft'.
    ls_context-validlist   = 'Draft,Released,Closed'.
    ls_context-show_secret = iv_show_secret.
    ls_context-secret      = 'internal price list'.

    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX16_ADVANCED`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
