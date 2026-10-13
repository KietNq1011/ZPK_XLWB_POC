"! <p class="shorttext synchronized" lang="en">XLWB ex 4.14/4.15: hierarchy -> nested table</p>
"!
"! Xuat cay phan cap (kieu ALV tree / SALV tree). Tool goc doc truc tiep
"! CL_GUI_ALV_TREE — class GUI khong ton tai tren Public Cloud, nen pattern
"! thay the la: chuyen du lieu PHANG (node_key + parent_key) sang NESTED
"! TABLE roi render bang loop long nhau.
"! Bai hoc:
"! - to_nested( ): flat parent-child -> groups/items (pattern dung lai duoc)
"! - group header merge ngang tinh + subtotal cnt/sum tung nhom
"! Template: ƯU TIÊN bản trong bảng ZXLWB_TMPL (form `XLWB_EX14_TREE`);
"! chưa khai ở đó thì dùng template dự phòng trong get_template( ).
CLASS zcl_xlwb_ex14_tree DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF ty_node,
             node_key   TYPE string,
             parent_key TYPE string,
             text       TYPE string,
             qty        TYPE p LENGTH 8 DECIMALS 0,
           END OF ty_node,
           ty_nodes TYPE STANDARD TABLE OF ty_node WITH EMPTY KEY,
           BEGIN OF ty_item,
             text TYPE string,
             qty  TYPE p LENGTH 8 DECIMALS 0,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_group,
             text  TYPE string,
             items TYPE ty_items,
           END OF ty_group,
           ty_groups TYPE STANDARD TABLE OF ty_group WITH EMPTY KEY.

    "! Chuyển bảng phẳng parent-child (2 cấp) sang nested table
    CLASS-METHODS to_nested
      IMPORTING it_nodes         TYPE ty_nodes
      RETURNING VALUE(rt_groups) TYPE ty_groups.

    METHODS get_template RETURNING VALUE(rv_template) TYPE string.
    METHODS get_file
      IMPORTING it_nodes       TYPE ty_nodes
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.
ENDCLASS.



CLASS ZCL_XLWB_EX14_TREE IMPLEMENTATION.


  METHOD to_nested.
    LOOP AT it_nodes INTO DATA(ls_root) WHERE parent_key IS INITIAL.
      DATA(ls_group) = VALUE ty_group( text = ls_root-text ).
      LOOP AT it_nodes INTO DATA(ls_child) WHERE parent_key = ls_root-node_key.
        APPEND VALUE ty_item( text = ls_child-text qty = ls_child-qty ) TO ls_group-items.
      ENDLOOP.
      APPEND ls_group TO rt_groups.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_template.
    rv_template =
      `<?xml version="1.0"?>` &&
      `<?mso-application progid="Excel.Sheet"?>` &&
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"` &&
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">` &&
      `<Styles>` &&
      `<Style ss:ID="grp"><Font ss:Bold="1"/>` &&
      `<Interior ss:Color="#DDEBF7" ss:Pattern="Solid"/></Style>` &&
      `<Style ss:ID="sub"><Font ss:Bold="1" ss:Italic="1"/>` &&
      `<Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>` &&
      `</Styles>` &&
      `<Worksheet ss:Name="Tree">` &&
      `<Table>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="200"/>` &&
      `<Column ss:AutoFitWidth="0" ss:Width="70"/>` &&
      `<Row><Cell ss:StyleID="grp" ss:MergeAcross="1">` &&
      `<Data ss:Type="String">{{#groups}}{{text}}</Data></Cell></Row>` &&
      `<Row><Cell><Data ss:Type="String">{{#items}}    {{text}}</Data></Cell>` &&
      `<Cell><Data ss:Type="String">{{qty}}{{/items}}</Data></Cell></Row>` &&
      `<Row><Cell ss:StyleID="sub"><Data ss:Type="String">Subtotal {{text}} ({{cnt:items}} items)</Data></Cell>` &&
      `<Cell ss:StyleID="sub"><Data ss:Type="String">{{sum:items.qty}}{{/groups}}</Data></Cell></Row>` &&
      `</Table>` &&
      `</Worksheet>` &&
      `</Workbook>`.
  ENDMETHOD.


  METHOD get_file.
    DATA: BEGIN OF ls_context,
            groups TYPE ty_groups,
          END OF ls_context.
    ls_context-groups = to_nested( it_nodes ).
    rv_file = zcl_xlwb_runtime=>render_prefer_table(
                iv_form_name         = `XLWB_EX14_TREE`
                iv_fallback_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                ir_context           = REF #( ls_context ) )-content.
  ENDMETHOD.
ENDCLASS.
