"! <p class="shorttext synchronized" lang="en">XLWB: SpreadsheetML -> OOXML (.xlsx) converter</p>
"!
"! Chuyển KẾT QUẢ render của engine SpreadsheetML thành file .xlsx thật —
"! chạy sau zcl_xlwb_engine nên giữ nguyên 100% tính năng template SSML
"! (cột động, sheet động, page break, mergesame...) mà file cuối vẫn là OOXML.
"!
"! Phạm vi v1: giá trị có kiểu (String/Number/DateTime), style (font, fill,
"! border, number format, alignment), độ rộng cột / chiều cao dòng, merge,
"! nhiều worksheet, công thức (R1C1 -> A1), page break, page setup + margins
"! + fit-to-page, freeze panes, gridlines, data validation List.
"! KHÔNG hỗ trợ: ảnh (dùng engine XLSX), rich text trong ô (mất định dạng
"! từng chữ, giữ nguyên text), outline/grouping.
CLASS zcl_xlwb_ssml2xlsx DEFINITION
  PUBLIC FINAL CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS convert
      IMPORTING iv_ssml        TYPE string
      RETURNING VALUE(rv_xlsx) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    CONSTANTS mc_uri_ss TYPE string VALUE 'urn:schemas-microsoft-com:office:spreadsheet'.
    CONSTANTS mc_uri_x  TYPE string VALUE 'urn:schemas-microsoft-com:office:excel'.

    TYPES ty_elems TYPE STANDARD TABLE OF REF TO if_ixml_element WITH EMPTY KEY.

    " model worksheet: dòng -> ô theo cột (OOXML đòi ô tăng dần theo cột, mà
    " ô đệm của vùng merge lại rơi vào các dòng SAU -> phải dựng model rồi emit)
    TYPES: BEGIN OF ty_ocell,
             col  TYPE i,
             frag TYPE string,
           END OF ty_ocell,
           ty_ocells TYPE SORTED TABLE OF ty_ocell WITH UNIQUE KEY col,
           BEGIN OF ty_orow,
             row   TYPE i,
             attrs TYPE string,
             cells TYPE ty_ocells,
           END OF ty_orow,
           ty_orows TYPE SORTED TABLE OF ty_orow WITH UNIQUE KEY row,
           BEGIN OF ty_fill,
             row TYPE i,
             col TYPE i,
             xf  TYPE i,
           END OF ty_fill,
           ty_fills TYPE STANDARD TABLE OF ty_fill WITH EMPTY KEY.

    TYPES: BEGIN OF ty_style,
             id       TYPE string,
             font_ix  TYPE i,
             fill_ix  TYPE i,
             border_ix TYPE i,
             numfmt_id TYPE i,
             align    TYPE string,          " fragment <alignment .../> hoặc rỗng
             xf_ix    TYPE i,
             xf_date_ix TYPE i,             " biến thể numFmt ngày khi style không có numfmt
           END OF ty_style,
           ty_styles TYPE SORTED TABLE OF ty_style WITH UNIQUE KEY id.

    TYPES: BEGIN OF ty_pool,
             key TYPE string,
             ix  TYPE i,
           END OF ty_pool,
           ty_pools TYPE SORTED TABLE OF ty_pool WITH UNIQUE KEY key.

    TYPES: BEGIN OF ty_sst,
             text TYPE string,
             ix   TYPE i,
           END OF ty_sst,
           ty_ssts TYPE SORTED TABLE OF ty_sst WITH UNIQUE KEY text.

    DATA mo_doc      TYPE REF TO if_ixml_document.
    DATA mt_styles   TYPE ty_styles.
    DATA mt_fonts    TYPE string_table.      " fragment xml theo index
    DATA mt_fills    TYPE string_table.
    DATA mt_borders  TYPE string_table.
    DATA mt_font_ix  TYPE ty_pools.
    DATA mt_fill_ix  TYPE ty_pools.
    DATA mt_border_ix TYPE ty_pools.
    DATA mt_numfmts  TYPE string_table.      " custom formats (id = 163 + index)
    DATA mt_xfs      TYPE string_table.      " cellXfs fragments
    DATA mt_sst      TYPE ty_ssts.
    DATA mt_sst_list TYPE string_table.
    DATA mv_date_xf_default TYPE i VALUE -1.

    METHODS parse
      IMPORTING iv_ssml TYPE string
      RAISING   zcx_xlwb.
    METHODS read_styles.
    METHODS style_xf
      IMPORTING iv_style_id  TYPE string
                iv_date      TYPE abap_bool DEFAULT abap_false
      RETURNING VALUE(rv_ix) TYPE i.
    METHODS pool_index
      IMPORTING iv_key       TYPE string
      CHANGING  ct_ix        TYPE ty_pools
                ct_list      TYPE string_table
      RETURNING VALUE(rv_ix) TYPE i.
    METHODS numfmt_id
      IMPORTING iv_format    TYPE string
      RETURNING VALUE(rv_id) TYPE i.
    METHODS sst_index
      IMPORTING iv_text      TYPE string
      RETURNING VALUE(rv_ix) TYPE i.
    METHODS build_worksheet
      IMPORTING io_ws         TYPE REF TO if_ixml_element
      RETURNING VALUE(rv_xml) TYPE string.
    METHODS rc_to_a1
      IMPORTING iv_formula   TYPE string
                iv_row       TYPE i
                iv_col       TYPE i
      RETURNING VALUE(rv_a1) TYPE string.
    METHODS rc_range_to_a1
      IMPORTING iv_range      TYPE string
      RETURNING VALUE(rv_ref) TYPE string.
    METHODS col_letter
      IMPORTING iv_col          TYPE i
      RETURNING VALUE(rv_name) TYPE string.
    METHODS date_serial
      IMPORTING iv_iso           TYPE string
      RETURNING VALUE(rv_serial) TYPE string.
    METHODS width_chars
      IMPORTING iv_points       TYPE string
      RETURNING VALUE(rv_chars) TYPE string.
    METHODS esc
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_out) TYPE string.
    METHODS esc_attr
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_out) TYPE string.
    METHODS attr
      IMPORTING io_elem        TYPE REF TO if_ixml_element
                iv_name        TYPE string
                iv_uri         TYPE string DEFAULT 'urn:schemas-microsoft-com:office:spreadsheet'
      RETURNING VALUE(rv_val) TYPE string.
    METHODS elem_children
      IMPORTING io_parent       TYPE REF TO if_ixml_element
                iv_name         TYPE string
      RETURNING VALUE(rt_elems) TYPE ty_elems.
    METHODS child_value
      IMPORTING io_parent      TYPE REF TO if_ixml_element
                iv_name        TYPE string
      RETURNING VALUE(rv_val) TYPE string.
    METHODS find_child
      IMPORTING io_parent      TYPE REF TO if_ixml_element
                iv_name        TYPE string
      RETURNING VALUE(ro_elem) TYPE REF TO if_ixml_element.
    METHODS zip_text
      IMPORTING io_zip  TYPE REF TO cl_abap_zip
                iv_name TYPE string
                iv_text TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_SSML2XLSX IMPLEMENTATION.


  METHOD convert.
    DATA(lo) = NEW zcl_xlwb_ssml2xlsx( ).
    lo->parse( iv_ssml ).
    lo->read_styles( ).

    DATA(lo_root) = lo->mo_doc->get_root_element( ).
    IF lo_root IS INITIAL OR lo_root->get_name( ) <> 'Workbook'.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |ssml2xlsx: root không phải <Workbook>| ).
    ENDIF.

    DATA lt_names  TYPE string_table.
    DATA lt_sheets TYPE string_table.
    LOOP AT lo->elem_children( io_parent = lo_root iv_name = 'Worksheet' ) INTO DATA(lo_ws).
      APPEND lo->attr( io_elem = lo_ws iv_name = 'Name' ) TO lt_names.
      APPEND lo->build_worksheet( lo_ws ) TO lt_sheets.
    ENDLOOP.
    IF lt_sheets IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |ssml2xlsx: workbook không có worksheet| ).
    ENDIF.

    " ---- workbook.xml + rels --------------------------------------------
    DATA(lv_n) = lines( lt_sheets ).
    DATA(lv_wb) = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` &&
      ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`.
    DATA(lv_rels) = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`.
    DATA(lv_types) = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` &&
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` &&
      `<Default Extension="xml" ContentType="application/xml"/>` &&
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` &&
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` &&
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`.

    DO lv_n TIMES.
      DATA(lv_i) = sy-index.
      READ TABLE lt_names INDEX lv_i INTO DATA(lv_name).
      IF lv_name IS INITIAL.
        lv_name = |Sheet{ lv_i }|.
      ENDIF.
      lv_wb &&= |<sheet name="{ lo->esc_attr( lv_name ) }" sheetId="{ lv_i }" r:id="rId{ lv_i }"/>|.
      lv_rels &&= |<Relationship Id="rId{ lv_i }"| &&
        ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"` &&
        | Target="worksheets/sheet{ lv_i }.xml"/>|.
      lv_types &&= |<Override PartName="/xl/worksheets/sheet{ lv_i }.xml"| &&
        ` ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`.
    ENDDO.
    lv_wb &&= `</sheets></workbook>`.
    lv_rels &&= |<Relationship Id="rId{ lv_n + 1 }"| &&
      ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` &&
      |<Relationship Id="rId{ lv_n + 2 }"| &&
      ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` &&
      `</Relationships>`.
    lv_types &&= `</Types>`.

    " ---- styles.xml ------------------------------------------------------
    DATA(lv_styles) = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`.
    IF lo->mt_numfmts IS NOT INITIAL.
      lv_styles &&= |<numFmts count="{ lines( lo->mt_numfmts ) }">|.
      LOOP AT lo->mt_numfmts INTO DATA(lv_fmt).
        lv_styles &&= |<numFmt numFmtId="{ 163 + sy-tabix }" formatCode="{ lo->esc_attr( lv_fmt ) }"/>|.
      ENDLOOP.
      lv_styles &&= `</numFmts>`.
    ENDIF.
    lv_styles &&= |<fonts count="{ lines( lo->mt_fonts ) }">|.
    LOOP AT lo->mt_fonts INTO DATA(lv_frag).
      lv_styles &&= lv_frag.
    ENDLOOP.
    lv_styles &&= `</fonts>` && |<fills count="{ lines( lo->mt_fills ) }">|.
    LOOP AT lo->mt_fills INTO lv_frag.
      lv_styles &&= lv_frag.
    ENDLOOP.
    lv_styles &&= `</fills>` && |<borders count="{ lines( lo->mt_borders ) }">|.
    LOOP AT lo->mt_borders INTO lv_frag.
      lv_styles &&= lv_frag.
    ENDLOOP.
    lv_styles &&= `</borders>` &&
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` &&
      |<cellXfs count="{ lines( lo->mt_xfs ) }">|.
    LOOP AT lo->mt_xfs INTO lv_frag.
      lv_styles &&= lv_frag.
    ENDLOOP.
    lv_styles &&= `</cellXfs>` &&
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` &&
      `</styleSheet>`.

    " ---- sharedStrings.xml ----------------------------------------------
    DATA(lv_sst) = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` &&
      | count="{ lines( lo->mt_sst_list ) }" uniqueCount="{ lines( lo->mt_sst_list ) }">|.
    LOOP AT lo->mt_sst_list INTO DATA(lv_text).
      DATA(lv_space) = ``.
      IF lv_text IS NOT INITIAL
      AND (    substring( val = lv_text off = 0 len = 1 ) = ` `
            OR substring( val = lv_text off = strlen( lv_text ) - 1 len = 1 ) = ` `
            OR lv_text CA cl_abap_char_utilities=>newline ).
        lv_space = ` xml:space="preserve"`.
      ENDIF.
      lv_sst &&= |<si><t{ lv_space }>{ lo->esc( lv_text ) }</t></si>|.
    ENDLOOP.
    lv_sst &&= `</sst>`.

    " ---- zip -------------------------------------------------------------
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo->zip_text( io_zip = lo_zip iv_name = '[Content_Types].xml' iv_text = lv_types ).
    lo->zip_text( io_zip = lo_zip iv_name = '_rels/.rels' iv_text =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` &&
      `<Relationship Id="rId1"` &&
      ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"` &&
      ` Target="xl/workbook.xml"/></Relationships>` ).
    lo->zip_text( io_zip = lo_zip iv_name = 'xl/workbook.xml' iv_text = lv_wb ).
    lo->zip_text( io_zip = lo_zip iv_name = 'xl/_rels/workbook.xml.rels' iv_text = lv_rels ).
    lo->zip_text( io_zip = lo_zip iv_name = 'xl/styles.xml' iv_text = lv_styles ).
    lo->zip_text( io_zip = lo_zip iv_name = 'xl/sharedStrings.xml' iv_text = lv_sst ).
    DO lv_n TIMES.
      READ TABLE lt_sheets INDEX sy-index INTO DATA(lv_sheet).
      lo->zip_text( io_zip = lo_zip iv_name = |xl/worksheets/sheet{ sy-index }.xml| iv_text = lv_sheet ).
    ENDDO.
    rv_xlsx = lo_zip->save( ).
  ENDMETHOD.


  METHOD parse.
    DATA(lo_ixml) = cl_ixml_core=>create( ).
    mo_doc = lo_ixml->create_document( ).
    DATA(lo_sf) = lo_ixml->create_stream_factory( ).
    DATA(lv_x) = cl_abap_conv_codepage=>create_out( )->convert( iv_ssml ).
    DATA(lo_parser) = lo_ixml->create_parser(
      document       = mo_doc
      istream        = lo_sf->create_istream_xstring( lv_x )
      stream_factory = lo_sf ).
    lo_parser->set_namespace_mode( if_ixml_parser_core=>co_namespace_aware ).
    IF lo_parser->parse( ) <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |ssml2xlsx: input không phải XML hợp lệ| ).
    ENDIF.
  ENDMETHOD.


  METHOD read_styles.
    " pool mặc định: font 0, fill none + gray125, border rỗng, xf 0
    pool_index( EXPORTING iv_key = `<font><sz val="11"/><name val="Calibri"/></font>`
                CHANGING ct_ix = mt_font_ix ct_list = mt_fonts ).
    pool_index( EXPORTING iv_key = `<fill><patternFill patternType="none"/></fill>`
                CHANGING ct_ix = mt_fill_ix ct_list = mt_fills ).
    pool_index( EXPORTING iv_key = `<fill><patternFill patternType="gray125"/></fill>`
                CHANGING ct_ix = mt_fill_ix ct_list = mt_fills ).
    pool_index( EXPORTING iv_key = `<border><left/><right/><top/><bottom/><diagonal/></border>`
                CHANGING ct_ix = mt_border_ix ct_list = mt_borders ).
    APPEND `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` TO mt_xfs.

    DATA(lo_root) = mo_doc->get_root_element( ).
    DATA(lo_styles) = find_child( io_parent = lo_root iv_name = 'Styles' ).
    IF lo_styles IS INITIAL.
      RETURN.
    ENDIF.

    LOOP AT elem_children( io_parent = lo_styles iv_name = 'Style' ) INTO DATA(lo_style).
      DATA(ls_style) = VALUE ty_style( id = attr( io_elem = lo_style iv_name = 'ID' ) ).

      " ---- font ----
      DATA(lv_font) = ``.
      DATA(lo_font) = find_child( io_parent = lo_style iv_name = 'Font' ).
      IF lo_font IS BOUND.
        IF attr( io_elem = lo_font iv_name = 'Bold' ) = '1'.      lv_font &&= `<b/>`. ENDIF.
        IF attr( io_elem = lo_font iv_name = 'Italic' ) = '1'.    lv_font &&= `<i/>`. ENDIF.
        IF attr( io_elem = lo_font iv_name = 'Underline' ) IS NOT INITIAL. lv_font &&= `<u/>`. ENDIF.
        DATA(lv_size) = attr( io_elem = lo_font iv_name = 'Size' ).
        lv_font &&= |<sz val="{ COND #( WHEN lv_size IS NOT INITIAL THEN lv_size ELSE `11` ) }"/>|.
        DATA(lv_color) = attr( io_elem = lo_font iv_name = 'Color' ).
        IF lv_color IS NOT INITIAL AND lv_color(1) = '#'.
          lv_font &&= |<color rgb="FF{ to_upper( substring( val = lv_color off = 1 ) ) }"/>|.
        ENDIF.
        DATA(lv_fname) = attr( io_elem = lo_font iv_name = 'FontName' ).
        lv_font &&= |<name val="{ esc_attr( COND #( WHEN lv_fname IS NOT INITIAL THEN lv_fname ELSE `Calibri` ) ) }"/>|.
      ELSE.
        lv_font = `<sz val="11"/><name val="Calibri"/>`.
      ENDIF.
      ls_style-font_ix = pool_index( EXPORTING iv_key = |<font>{ lv_font }</font>|
                                     CHANGING ct_ix = mt_font_ix ct_list = mt_fonts ).

      " ---- fill ----
      DATA(lo_int) = find_child( io_parent = lo_style iv_name = 'Interior' ).
      DATA(lv_icolor) = COND #( WHEN lo_int IS BOUND THEN attr( io_elem = lo_int iv_name = 'Color' ) ).
      IF lv_icolor IS NOT INITIAL AND lv_icolor(1) = '#'.
        ls_style-fill_ix = pool_index(
          EXPORTING iv_key = `<fill><patternFill patternType="solid"><fgColor rgb="FF` &&
                             to_upper( substring( val = lv_icolor off = 1 ) ) &&
                             `"/></patternFill></fill>`
          CHANGING ct_ix = mt_fill_ix ct_list = mt_fills ).
      ENDIF.

      " ---- borders ----
      DATA(lo_borders) = find_child( io_parent = lo_style iv_name = 'Borders' ).
      IF lo_borders IS BOUND.
        DATA(lv_left) = ``. DATA(lv_right) = ``. DATA(lv_top) = ``. DATA(lv_bottom) = ``.
        LOOP AT elem_children( io_parent = lo_borders iv_name = 'Border' ) INTO DATA(lo_border).
          DATA(lv_ls) = attr( io_elem = lo_border iv_name = 'LineStyle' ).
          DATA(lv_wt) = attr( io_elem = lo_border iv_name = 'Weight' ).
          DATA(lv_bs) = SWITCH string( lv_ls
            WHEN 'Double' THEN `double`
            WHEN 'Dash'   THEN `dashed`
            WHEN 'Dot'    THEN `dotted`
            ELSE SWITCH string( lv_wt
              WHEN '2' THEN `medium`
              WHEN '3' THEN `thick`
              ELSE `thin` ) ).
          DATA(lv_bfrag) = |style="{ lv_bs }"|.
          CASE attr( io_elem = lo_border iv_name = 'Position' ).
            WHEN 'Left'.   lv_left   = lv_bfrag.
            WHEN 'Right'.  lv_right  = lv_bfrag.
            WHEN 'Top'.    lv_top    = lv_bfrag.
            WHEN 'Bottom'. lv_bottom = lv_bfrag.
          ENDCASE.
        ENDLOOP.
        ls_style-border_ix = pool_index(
          EXPORTING iv_key =
            `<border>` &&
            COND string( WHEN lv_left   IS INITIAL THEN `<left/>`   ELSE |<left { lv_left }><color auto="1"/></left>| ) &&
            COND string( WHEN lv_right  IS INITIAL THEN `<right/>`  ELSE |<right { lv_right }><color auto="1"/></right>| ) &&
            COND string( WHEN lv_top    IS INITIAL THEN `<top/>`    ELSE |<top { lv_top }><color auto="1"/></top>| ) &&
            COND string( WHEN lv_bottom IS INITIAL THEN `<bottom/>` ELSE |<bottom { lv_bottom }><color auto="1"/></bottom>| ) &&
            `<diagonal/></border>`
          CHANGING ct_ix = mt_border_ix ct_list = mt_borders ).
      ENDIF.

      " ---- number format ----
      DATA(lo_nf) = find_child( io_parent = lo_style iv_name = 'NumberFormat' ).
      IF lo_nf IS BOUND.
        ls_style-numfmt_id = numfmt_id( attr( io_elem = lo_nf iv_name = 'Format' ) ).
      ENDIF.

      " ---- alignment ----
      DATA(lo_al) = find_child( io_parent = lo_style iv_name = 'Alignment' ).
      IF lo_al IS BOUND.
        DATA(lv_al) = ``.
        DATA(lv_h) = attr( io_elem = lo_al iv_name = 'Horizontal' ).
        IF lv_h IS NOT INITIAL AND lv_h <> 'Automatic'.
          lv_al &&= | horizontal="{ to_lower( lv_h ) }"|.
        ENDIF.
        DATA(lv_v) = attr( io_elem = lo_al iv_name = 'Vertical' ).
        IF lv_v IS NOT INITIAL AND lv_v <> 'Automatic' AND lv_v <> 'Bottom'.
          lv_al &&= | vertical="{ to_lower( lv_v ) }"|.
        ENDIF.
        IF attr( io_elem = lo_al iv_name = 'WrapText' ) = '1'.
          lv_al &&= ` wrapText="1"`.
        ENDIF.
        IF lv_al IS NOT INITIAL.
          ls_style-align = |<alignment{ lv_al }/>|.
        ENDIF.
      ENDIF.

      " ---- xf ----
      DATA(lv_xf) = |<xf numFmtId="{ ls_style-numfmt_id }" fontId="{ ls_style-font_ix }"| &&
                    | fillId="{ ls_style-fill_ix }" borderId="{ ls_style-border_ix }" xfId="0"| &&
                    COND string( WHEN ls_style-numfmt_id > 0 THEN ` applyNumberFormat="1"` ) &&
                    ` applyFont="1"` &&
                    COND string( WHEN ls_style-fill_ix > 1 THEN ` applyFill="1"` ) &&
                    COND string( WHEN ls_style-border_ix > 0 THEN ` applyBorder="1"` ) &&
                    COND string( WHEN ls_style-align IS NOT INITIAL
                                 THEN | applyAlignment="1">{ ls_style-align }</xf>|
                                 ELSE `/>` ).
      APPEND lv_xf TO mt_xfs.
      ls_style-xf_ix = lines( mt_xfs ) - 1.
      ls_style-xf_date_ix = -1.

      INSERT ls_style INTO TABLE mt_styles.
    ENDLOOP.
  ENDMETHOD.


  METHOD style_xf.
    IF iv_style_id IS INITIAL.
      IF iv_date = abap_false.
        rv_ix = 0.
        RETURN.
      ENDIF.
      IF mv_date_xf_default < 0.
        APPEND `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` TO mt_xfs.
        mv_date_xf_default = lines( mt_xfs ) - 1.
      ENDIF.
      rv_ix = mv_date_xf_default.
      RETURN.
    ENDIF.

    READ TABLE mt_styles ASSIGNING FIELD-SYMBOL(<ls_style>) WITH TABLE KEY id = iv_style_id.
    IF sy-subrc <> 0.
      rv_ix = COND #( WHEN iv_date = abap_true THEN style_xf( iv_style_id = `` iv_date = abap_true ) ).
      RETURN.
    ENDIF.

    IF iv_date = abap_false OR <ls_style>-numfmt_id > 0.
      rv_ix = <ls_style>-xf_ix.
      RETURN.
    ENDIF.

    " ô DateTime nhưng style không khai number format -> biến thể numFmt 14
    IF <ls_style>-xf_date_ix < 0.
      APPEND |<xf numFmtId="14" fontId="{ <ls_style>-font_ix }" fillId="{ <ls_style>-fill_ix }"| &&
             | borderId="{ <ls_style>-border_ix }" xfId="0" applyNumberFormat="1" applyFont="1"| &&
             COND string( WHEN <ls_style>-fill_ix > 1 THEN ` applyFill="1"` ) &&
             COND string( WHEN <ls_style>-border_ix > 0 THEN ` applyBorder="1"` ) &&
             COND string( WHEN <ls_style>-align IS NOT INITIAL
                          THEN | applyAlignment="1">{ <ls_style>-align }</xf>|
                          ELSE `/>` )
        TO mt_xfs.
      <ls_style>-xf_date_ix = lines( mt_xfs ) - 1.
    ENDIF.
    rv_ix = <ls_style>-xf_date_ix.
  ENDMETHOD.


  METHOD pool_index.
    READ TABLE ct_ix INTO DATA(ls_hit) WITH TABLE KEY key = iv_key.
    IF sy-subrc = 0.
      rv_ix = ls_hit-ix.
      RETURN.
    ENDIF.
    APPEND iv_key TO ct_list.
    rv_ix = lines( ct_list ) - 1.
    INSERT VALUE #( key = iv_key ix = rv_ix ) INTO TABLE ct_ix.
  ENDMETHOD.


  METHOD numfmt_id.
    CASE iv_format.
      WHEN `` OR `General`.        rv_id = 0.
      WHEN `Fixed`.                rv_id = 2.
      WHEN `Standard`.             rv_id = 4.
      WHEN `Percent`.              rv_id = 10.
      WHEN `Scientific`.           rv_id = 11.
      WHEN `Short Date`.           rv_id = 14.
      WHEN `Short Time`.           rv_id = 20.
      WHEN `Long Time`.            rv_id = 21.
      WHEN `General Date`.         rv_id = 22.
      WHEN `@`.                    rv_id = 49.
      WHEN OTHERS.
        " custom format -> đăng ký numFmt riêng (id 164+)
        READ TABLE mt_numfmts TRANSPORTING NO FIELDS WITH KEY table_line = iv_format.
        IF sy-subrc = 0.
          rv_id = 163 + sy-tabix.
        ELSE.
          APPEND iv_format TO mt_numfmts.
          rv_id = 163 + lines( mt_numfmts ).
        ENDIF.
    ENDCASE.
  ENDMETHOD.


  METHOD sst_index.
    READ TABLE mt_sst INTO DATA(ls_hit) WITH TABLE KEY text = iv_text.
    IF sy-subrc = 0.
      rv_ix = ls_hit-ix.
      RETURN.
    ENDIF.
    APPEND iv_text TO mt_sst_list.
    rv_ix = lines( mt_sst_list ) - 1.
    INSERT VALUE #( text = iv_text ix = rv_ix ) INTO TABLE mt_sst.
  ENDMETHOD.


  METHOD build_worksheet.
    DATA lt_merges TYPE string_table.
    DATA lt_valids TYPE string_table.
    DATA lv_cols   TYPE string.
    DATA lv_rows   TYPE string.

    DATA(lo_table) = find_child( io_parent = io_ws iv_name = 'Table' ).

    " ---- cột: độ rộng -----------------------------------------------------
    IF lo_table IS BOUND.
      DATA(lv_cix) = 0.
      LOOP AT elem_children( io_parent = lo_table iv_name = 'Column' ) INTO DATA(lo_col).
        DATA(lv_colidx) = attr( io_elem = lo_col iv_name = 'Index' ).
        lv_cix = COND #( WHEN lv_colidx IS NOT INITIAL THEN CONV i( lv_colidx ) ELSE lv_cix + 1 ).
        DATA(lv_span) = attr( io_elem = lo_col iv_name = 'Span' ).
        DATA(lv_to) = lv_cix + COND i( WHEN lv_span IS NOT INITIAL THEN CONV i( lv_span ) ELSE 0 ).
        DATA(lv_w) = attr( io_elem = lo_col iv_name = 'Width' ).
        IF lv_w IS NOT INITIAL.
          lv_cols &&= |<col min="{ lv_cix }" max="{ lv_to }" width="{ width_chars( lv_w ) }" customWidth="1"/>|.
        ENDIF.
        lv_cix = lv_to.
      ENDLOOP.
    ENDIF.

    " ---- dòng + ô ---------------------------------------------------------
    DATA lt_orows TYPE ty_orows.
    DATA lt_fill  TYPE ty_fills.
    DATA lt_cells TYPE ty_ocells.
    DATA(lv_outline_max) = 0.
    IF lo_table IS BOUND.
      DATA(lv_r) = 0.
      LOOP AT elem_children( io_parent = lo_table iv_name = 'Row' ) INTO DATA(lo_row).
        DATA(lv_ridx) = attr( io_elem = lo_row iv_name = 'Index' ).
        lv_r = COND #( WHEN lv_ridx IS NOT INITIAL THEN CONV i( lv_ridx ) ELSE lv_r + 1 ).

        DATA(lv_rattr) = ``.
        DATA(lv_h) = attr( io_elem = lo_row iv_name = 'Height' ).
        IF lv_h IS NOT INITIAL.
          lv_rattr = | ht="{ lv_h }" customHeight="1"|.
        ENDIF.
        " outline từ marker {{*group=N}} của engine SSML
        DATA(lv_outline) = lo_row->get_attribute( 'xlwboutline' ).
        IF lv_outline CO '1234567' AND lv_outline IS NOT INITIAL.
          lv_rattr &&= | outlineLevel="{ lv_outline }"|.
          IF CONV i( lv_outline ) > lv_outline_max.
            lv_outline_max = CONV i( lv_outline ).
          ENDIF.
        ENDIF.

        CLEAR lt_cells.
        DATA(lv_c) = 0.
        LOOP AT elem_children( io_parent = lo_row iv_name = 'Cell' ) INTO DATA(lo_cell).
          DATA(lv_cellidx) = attr( io_elem = lo_cell iv_name = 'Index' ).
          lv_c = COND #( WHEN lv_cellidx IS NOT INITIAL THEN CONV i( lv_cellidx ) ELSE lv_c + 1 ).

          DATA(lv_ma) = attr( io_elem = lo_cell iv_name = 'MergeAcross' ).
          DATA(lv_md) = attr( io_elem = lo_cell iv_name = 'MergeDown' ).
          DATA(lv_man) = COND i( WHEN lv_ma IS NOT INITIAL THEN CONV i( lv_ma ) ELSE 0 ).
          DATA(lv_mdn) = COND i( WHEN lv_md IS NOT INITIAL THEN CONV i( lv_md ) ELSE 0 ).
          IF lv_man > 0 OR lv_mdn > 0.
            APPEND |{ col_letter( lv_c ) }{ lv_r }:{ col_letter( lv_c + lv_man ) }{ lv_r + lv_mdn }|
              TO lt_merges.
          ENDIF.

          DATA(lv_ref)      = |{ col_letter( lv_c ) }{ lv_r }|.
          DATA(lo_data)     = find_child( io_parent = lo_cell iv_name = 'Data' ).
          DATA(lv_formula)  = attr( io_elem = lo_cell iv_name = 'Formula' ).
          DATA(lv_style_id) = attr( io_elem = lo_cell iv_name = 'StyleID' ).
          DATA(lv_type)     = COND #( WHEN lo_data IS BOUND
                                      THEN attr( io_elem = lo_data iv_name = 'Type' ) ).
          DATA(lv_val)      = COND #( WHEN lo_data IS BOUND THEN lo_data->get_value( ) ).
          DATA(lv_s)        = style_xf( iv_style_id = lv_style_id
                                        iv_date = xsdbool( lv_type = 'DateTime' ) ).

          DATA(lv_frag) = ``.
          IF lv_formula IS NOT INITIAL.
            DATA(lv_f) = lv_formula.
            IF lv_f CP '=*'.
              lv_f = substring( val = lv_f off = 1 ).
            ENDIF.
            DATA(lv_body) = |<f>{ esc( rc_to_a1( iv_formula = lv_f
                                                 iv_row = lv_r iv_col = lv_c ) ) }</f>|.
            IF lv_type = 'Number' AND lv_val IS NOT INITIAL.
              lv_body &&= |<v>{ lv_val }</v>|.
            ENDIF.
            lv_frag = |<c r="{ lv_ref }" s="{ lv_s }">{ lv_body }</c>|.

          ELSEIF lo_data IS INITIAL.
            " ô rỗng: giữ khi có style (khung/border), bỏ khi không
            IF lv_style_id IS NOT INITIAL.
              lv_frag = |<c r="{ lv_ref }" s="{ lv_s }"/>|.
            ENDIF.

          ELSE.
            CASE lv_type.
              WHEN 'Number'.
                lv_frag = |<c r="{ lv_ref }" s="{ lv_s }"><v>{ lv_val }</v></c>|.
              WHEN 'DateTime'.
                lv_frag = |<c r="{ lv_ref }" s="{ lv_s }"><v>{ date_serial( lv_val ) }</v></c>|.
              WHEN OTHERS.
                IF lv_val IS INITIAL AND lv_style_id IS INITIAL.
                  " bỏ ô rỗng không style
                ELSEIF lv_val IS INITIAL.
                  lv_frag = |<c r="{ lv_ref }" s="{ lv_s }"/>|.
                ELSE.
                  lv_frag = |<c r="{ lv_ref }" s="{ lv_s }" t="s"><v>{ sst_index( lv_val ) }</v></c>|.
                ENDIF.
            ENDCASE.
          ENDIF.

          IF lv_frag IS NOT INITIAL.
            DELETE lt_cells WHERE col = lv_c.
            INSERT VALUE #( col = lv_c frag = lv_frag ) INTO TABLE lt_cells.
          ENDIF.

          " Ô ĐỆM của vùng merge. SpreadsheetML chỉ có ô NEO, còn Excel khi tự
          " ghi .xlsx thì ghi MỌI ô của vùng kèm style. Nếu chỉ có ô neo, Excel
          " chỉ vẽ được cạnh trên + trái -> KHUNG BỊ VỠ (kiểm chứng bằng Excel:
          " anchor-only cho Bottom=NONE Right=NONE, fill đủ cho 4 cạnh LINE).
          IF ( lv_man > 0 OR lv_mdn > 0 ) AND lv_s > 0
          AND ( lv_man + 1 ) * ( lv_mdn + 1 ) <= 2000.
            DO lv_mdn + 1 TIMES.
              DATA(lv_frr) = lv_r + sy-index - 1.
              DO lv_man + 1 TIMES.
                DATA(lv_fcc) = lv_c + sy-index - 1.
                IF lv_frr <> lv_r OR lv_fcc <> lv_c.
                  APPEND VALUE #( row = lv_frr col = lv_fcc xf = lv_s ) TO lt_fill.
                ENDIF.
              ENDDO.
            ENDDO.
          ENDIF.

          lv_c += lv_man.
        ENDLOOP.

        " gộp vào model (giữ cả dòng rỗng -> không lệch số dòng so với SSML)
        READ TABLE lt_orows ASSIGNING FIELD-SYMBOL(<ls_orow>) WITH TABLE KEY row = lv_r.
        IF sy-subrc <> 0.
          INSERT VALUE #( row = lv_r attrs = lv_rattr cells = lt_cells ) INTO TABLE lt_orows.
        ELSE.
          <ls_orow>-attrs = lv_rattr.
          LOOP AT lt_cells INTO DATA(ls_new_cell).
            DELETE <ls_orow>-cells WHERE col = ls_new_cell-col.
            INSERT ls_new_cell INTO TABLE <ls_orow>-cells.
          ENDLOOP.
        ENDIF.
      ENDLOOP.
    ENDIF.

    " điền ô đệm vào vị trí CHƯA có ô thật
    LOOP AT lt_fill INTO DATA(ls_fill).
      READ TABLE lt_orows ASSIGNING <ls_orow> WITH TABLE KEY row = ls_fill-row.
      IF sy-subrc <> 0.
        INSERT VALUE #( row = ls_fill-row ) INTO TABLE lt_orows.
        READ TABLE lt_orows ASSIGNING <ls_orow> WITH TABLE KEY row = ls_fill-row.
      ENDIF.
      IF NOT line_exists( <ls_orow>-cells[ col = ls_fill-col ] ).
        INSERT VALUE #( col  = ls_fill-col
                        frag = |<c r="{ col_letter( ls_fill-col ) }{ ls_fill-row }"| &&
                               | s="{ ls_fill-xf }"/>| )
          INTO TABLE <ls_orow>-cells.
      ENDIF.
    ENDLOOP.

    LOOP AT lt_orows INTO DATA(ls_orow_out).
      DATA(lv_rowcells) = ``.
      LOOP AT ls_orow_out-cells INTO DATA(ls_ocell).
        lv_rowcells &&= ls_ocell-frag.
      ENDLOOP.
      lv_rows &&= |<row r="{ ls_orow_out-row }"{ ls_orow_out-attrs }>{ lv_rowcells }</row>|.
    ENDLOOP.

    " ---- WorksheetOptions: view / freeze / page setup ---------------------
    DATA(lv_gridlines) = ``.
    DATA(lv_pane) = ``.
    DATA(lv_pagesetup_attr) = ``.
    DATA(lv_sheetpr) = ``.
    DATA: lv_m_top TYPE string VALUE `0.75`, lv_m_bottom TYPE string VALUE `0.75`,
          lv_m_left TYPE string VALUE `0.7`, lv_m_right TYPE string VALUE `0.7`,
          lv_m_head TYPE string VALUE `0.3`, lv_m_foot TYPE string VALUE `0.3`.

    DATA(lo_opts) = find_child( io_parent = io_ws iv_name = 'WorksheetOptions' ).
    IF lo_opts IS BOUND.
      IF find_child( io_parent = lo_opts iv_name = 'DoNotDisplayGridlines' ) IS BOUND.
        lv_gridlines = ` showGridLines="0"`.
      ENDIF.

      IF find_child( io_parent = lo_opts iv_name = 'FreezePanes' ) IS BOUND.
        DATA(lv_xs) = child_value( io_parent = lo_opts iv_name = 'SplitVertical' ).
        DATA(lv_ys) = child_value( io_parent = lo_opts iv_name = 'SplitHorizontal' ).
        DATA(lv_xn) = COND i( WHEN lv_xs IS NOT INITIAL THEN CONV i( lv_xs ) ELSE 0 ).
        DATA(lv_yn) = COND i( WHEN lv_ys IS NOT INITIAL THEN CONV i( lv_ys ) ELSE 0 ).
        IF lv_xn > 0 OR lv_yn > 0.
          DATA(lv_tl) = |{ col_letter( lv_xn + 1 ) }{ lv_yn + 1 }|.
          DATA(lv_active) = COND string( WHEN lv_xn > 0 AND lv_yn > 0 THEN `bottomRight`
                                         WHEN lv_xn > 0 THEN `topRight` ELSE `bottomLeft` ).
          lv_pane = `<pane ` &&
            COND string( WHEN lv_xn > 0 THEN |xSplit="{ lv_xn }" | ) &&
            COND string( WHEN lv_yn > 0 THEN |ySplit="{ lv_yn }" | ) &&
            |topLeftCell="{ lv_tl }" activePane="{ lv_active }" state="frozen"/>|.
        ENDIF.
      ENDIF.

      DATA(lo_ps) = find_child( io_parent = lo_opts iv_name = 'PageSetup' ).
      IF lo_ps IS BOUND.
        DATA(lo_layout) = find_child( io_parent = lo_ps iv_name = 'Layout' ).
        IF lo_layout IS BOUND.
          DATA(lv_orient) = attr( io_elem = lo_layout iv_name = 'Orientation' iv_uri = mc_uri_x ).
          IF lv_orient IS NOT INITIAL.
            lv_pagesetup_attr &&= | orientation="{ to_lower( lv_orient ) }"|.
          ENDIF.
        ENDIF.
        DATA(lo_margins) = find_child( io_parent = lo_ps iv_name = 'PageMargins' ).
        IF lo_margins IS BOUND.
          DATA(lv_mv) = attr( io_elem = lo_margins iv_name = 'Top' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_top = lv_mv. ENDIF.
          lv_mv = attr( io_elem = lo_margins iv_name = 'Bottom' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_bottom = lv_mv. ENDIF.
          lv_mv = attr( io_elem = lo_margins iv_name = 'Left' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_left = lv_mv. ENDIF.
          lv_mv = attr( io_elem = lo_margins iv_name = 'Right' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_right = lv_mv. ENDIF.
        ENDIF.
        DATA(lo_header) = find_child( io_parent = lo_ps iv_name = 'Header' ).
        IF lo_header IS BOUND.
          lv_mv = attr( io_elem = lo_header iv_name = 'Margin' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_head = lv_mv. ENDIF.
        ENDIF.
        DATA(lo_footer) = find_child( io_parent = lo_ps iv_name = 'Footer' ).
        IF lo_footer IS BOUND.
          lv_mv = attr( io_elem = lo_footer iv_name = 'Margin' iv_uri = mc_uri_x ).
          IF lv_mv IS NOT INITIAL. lv_m_foot = lv_mv. ENDIF.
        ENDIF.
      ENDIF.

      DATA(lo_print) = find_child( io_parent = lo_opts iv_name = 'Print' ).
      IF lo_print IS BOUND.
        DATA(lv_paper) = child_value( io_parent = lo_print iv_name = 'PaperSizeIndex' ).
        IF lv_paper IS NOT INITIAL.
          lv_pagesetup_attr &&= | paperSize="{ lv_paper }"|.
        ENDIF.
        IF find_child( io_parent = lo_print iv_name = 'FitToPage' ) IS BOUND.
          DATA(lv_fitpage) = abap_true.
          DATA(lv_fh) = child_value( io_parent = lo_print iv_name = 'FitHeight' ).
          DATA(lv_fw) = child_value( io_parent = lo_print iv_name = 'FitWidth' ).
          lv_pagesetup_attr &&=
            | fitToWidth="{ COND #( WHEN lv_fw IS NOT INITIAL THEN lv_fw ELSE `1` ) }"| &&
            | fitToHeight="{ COND #( WHEN lv_fh IS NOT INITIAL THEN lv_fh ELSE `0` ) }"|.
        ENDIF.
      ENDIF.
    ENDIF.

    " ---- auto filter (nút lọc trên header) ---------------------------------
    DATA(lv_autofilter) = ``.
    DATA(lo_af) = find_child( io_parent = io_ws iv_name = 'AutoFilter' ).
    IF lo_af IS BOUND.
      DATA(lv_af_range) = attr( io_elem = lo_af iv_name = 'Range' iv_uri = mc_uri_x ).
      IF lv_af_range IS NOT INITIAL.
        lv_autofilter = |<autoFilter ref="{ rc_range_to_a1( lv_af_range ) }"/>|.
      ENDIF.
    ENDIF.

    " ---- data validation List ---------------------------------------------
    " nằm được ở 2 chỗ: con trực tiếp của Worksheet HOẶC trong WorksheetOptions
    DATA(lt_dvs) = elem_children( io_parent = io_ws iv_name = 'DataValidation' ).
    IF lo_opts IS BOUND.
      APPEND LINES OF elem_children( io_parent = lo_opts iv_name = 'DataValidation' ) TO lt_dvs.
    ENDIF.
    LOOP AT lt_dvs INTO DATA(lo_dv).
      IF child_value( io_parent = lo_dv iv_name = 'Type' ) <> 'List'.
        CONTINUE.
      ENDIF.
      DATA(lv_range) = child_value( io_parent = lo_dv iv_name = 'Range' ).
      DATA(lv_list)  = child_value( io_parent = lo_dv iv_name = 'Value' ).
      IF lv_range IS INITIAL OR lv_list IS INITIAL.
        CONTINUE.
      ENDIF.
      APPEND |<dataValidation type="list" allowBlank="1" showInputMessage="1"| &&
             | showErrorMessage="1" sqref="{ rc_range_to_a1( lv_range ) }">| &&
             |<formula1>{ esc( lv_list ) }</formula1></dataValidation>| TO lt_valids.
    ENDLOOP.

    " ---- page breaks --------------------------------------------------------
    DATA(lv_breaks) = ``.
    DATA(lo_pb) = find_child( io_parent = io_ws iv_name = 'PageBreaks' ).
    IF lo_pb IS BOUND.
      DATA(lo_rbs) = find_child( io_parent = lo_pb iv_name = 'RowBreaks' ).
      IF lo_rbs IS BOUND.
        DATA lt_brk TYPE string_table.
        LOOP AT elem_children( io_parent = lo_rbs iv_name = 'RowBreak' ) INTO DATA(lo_rb).
          DATA(lv_brow) = child_value( io_parent = lo_rb iv_name = 'Row' ).
          IF lv_brow IS NOT INITIAL AND lv_brow <> '0'.
            APPEND |<brk id="{ lv_brow }" max="16383" man="1"/>| TO lt_brk.
          ENDIF.
        ENDLOOP.
        IF lt_brk IS NOT INITIAL.
          lv_breaks = |<rowBreaks count="{ lines( lt_brk ) }" manualBreakCount="{ lines( lt_brk ) }">|.
          LOOP AT lt_brk INTO DATA(lv_brk).
            lv_breaks &&= lv_brk.
          ENDLOOP.
          lv_breaks &&= `</rowBreaks>`.
        ENDIF.
      ENDIF.
    ENDIF.

    " ---- lắp ráp worksheet ---------------------------------------------------
    IF lv_outline_max > 0 OR lv_fitpage = abap_true.
      " outlinePr summaryBelow=0: dòng tổng/nhóm nằm TRÊN nhóm chi tiết (4.08)
      lv_sheetpr = `<sheetPr>` &&
        COND string( WHEN lv_outline_max > 0 THEN `<outlinePr summaryBelow="0"/>` ) &&
        COND string( WHEN lv_fitpage = abap_true THEN `<pageSetUpPr fitToPage="1"/>` ) &&
        `</sheetPr>`.
    ENDIF.

    rv_xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` &&
      lv_sheetpr &&
      `<sheetViews><sheetView workbookViewId="0"` && lv_gridlines &&
      COND string( WHEN lv_pane IS INITIAL THEN `/>` ELSE |>{ lv_pane }</sheetView>| ) &&
      `</sheetViews>` &&
      COND string( WHEN lv_outline_max > 0
        THEN |<sheetFormatPr defaultRowHeight="15" outlineLevelRow="{ lv_outline_max }"/>| ) &&
      COND string( WHEN lv_cols IS NOT INITIAL THEN |<cols>{ lv_cols }</cols>| ) &&
      |<sheetData>{ lv_rows }</sheetData>| &&
      lv_autofilter.   " schema: autoFilter đứng TRƯỚC mergeCells

    IF lt_merges IS NOT INITIAL.
      rv_xml &&= |<mergeCells count="{ lines( lt_merges ) }">|.
      LOOP AT lt_merges INTO DATA(lv_merge).
        rv_xml &&= |<mergeCell ref="{ lv_merge }"/>|.
      ENDLOOP.
      rv_xml &&= `</mergeCells>`.
    ENDIF.

    IF lt_valids IS NOT INITIAL.
      rv_xml &&= |<dataValidations count="{ lines( lt_valids ) }">|.
      LOOP AT lt_valids INTO DATA(lv_valid).
        rv_xml &&= lv_valid.
      ENDLOOP.
      rv_xml &&= `</dataValidations>`.
    ENDIF.

    rv_xml &&= |<pageMargins left="{ lv_m_left }" right="{ lv_m_right }" top="{ lv_m_top }"| &&
               | bottom="{ lv_m_bottom }" header="{ lv_m_head }" footer="{ lv_m_foot }"/>| &&
               COND string( WHEN lv_pagesetup_attr IS NOT INITIAL
                            THEN |<pageSetup{ lv_pagesetup_attr }/>| ) &&
               lv_breaks &&
               `</worksheet>`.
  ENDMETHOD.


  METHOD rc_to_a1.
    " R1C1 -> A1: R4C -> cột hiện tại dòng 4 tuyệt đối; R[-2]C[1] -> tương đối
    DATA(lv_rest) = iv_formula.
    FIND ALL OCCURRENCES OF PCRE
      '(?-x)(?<![A-Za-z0-9_])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])'
      IN iv_formula RESULTS DATA(lt_res).

    DATA(lv_pos) = 0.
    LOOP AT lt_res INTO DATA(ls_res).
      rv_a1 &&= substring( val = iv_formula off = lv_pos len = ls_res-offset - lv_pos ).

      DATA(lv_rspec) = ``.
      DATA(lv_cspec) = ``.
      READ TABLE ls_res-submatches INDEX 1 INTO DATA(ls_sub).
      IF sy-subrc = 0 AND ls_sub-length > 0.
        lv_rspec = substring( val = iv_formula off = ls_sub-offset len = ls_sub-length ).
      ENDIF.
      READ TABLE ls_res-submatches INDEX 2 INTO ls_sub.
      IF sy-subrc = 0 AND ls_sub-length > 0.
        lv_cspec = substring( val = iv_formula off = ls_sub-offset len = ls_sub-length ).
      ENDIF.

      DATA(lv_rowout) = ``.
      IF lv_rspec IS INITIAL.
        lv_rowout = |{ iv_row }|.                       " R = dòng hiện tại (tương đối)
      ELSEIF lv_rspec(1) = '['.
        DATA(lv_off) = CONV i( substring( val = lv_rspec off = 1 len = strlen( lv_rspec ) - 2 ) ).
        lv_rowout = |{ iv_row + lv_off }|.
      ELSE.
        lv_rowout = |${ lv_rspec }|.                    " tuyệt đối
      ENDIF.

      DATA(lv_colout) = ``.
      IF lv_cspec IS INITIAL.
        lv_colout = col_letter( iv_col ).
      ELSEIF lv_cspec(1) = '['.
        lv_off = CONV i( substring( val = lv_cspec off = 1 len = strlen( lv_cspec ) - 2 ) ).
        lv_colout = col_letter( iv_col + lv_off ).
      ELSE.
        lv_colout = |${ col_letter( CONV i( lv_cspec ) ) }|.
      ENDIF.

      rv_a1 &&= lv_colout && lv_rowout.
      lv_pos = ls_res-offset + ls_res-length.
    ENDLOOP.
    rv_a1 &&= substring( val = iv_formula off = lv_pos ).
  ENDMETHOD.


  METHOD rc_range_to_a1.
    " chỉ dạng tuyệt đối RnCm hoặc RnCm:RnCm (data validation range)
    SPLIT iv_range AT ':' INTO TABLE DATA(lt_parts).
    LOOP AT lt_parts INTO DATA(lv_part).
      DATA(lv_r) = ``.
      DATA(lv_c) = ``.
      FIND PCRE '(?-x)^R(\d+)C(\d+)$' IN condense( lv_part ) SUBMATCHES lv_r lv_c.
      IF sy-subrc = 0.
        rv_ref &&= COND string( WHEN rv_ref IS NOT INITIAL THEN `:` ) &&
                   col_letter( CONV i( lv_c ) ) && lv_r.
      ENDIF.
    ENDLOOP.
    IF rv_ref IS INITIAL.
      rv_ref = `A1`.
    ENDIF.
  ENDMETHOD.


  METHOD col_letter.
    DATA(lv_n) = iv_col.
    WHILE lv_n > 0.
      DATA(lv_m) = ( lv_n - 1 ) MOD 26.
      rv_name = substring( val = `ABCDEFGHIJKLMNOPQRSTUVWXYZ` off = lv_m len = 1 ) && rv_name.
      lv_n = ( lv_n - 1 ) DIV 26.
    ENDWHILE.
    IF rv_name IS INITIAL.
      rv_name = `A`.
    ENDIF.
  ENDMETHOD.


  METHOD date_serial.
    " 'YYYY-MM-DDThh:mm:ss.mmm' -> serial Excel (epoch 1899-12-30)
    IF strlen( iv_iso ) < 10.
      rv_serial = `0`.
      RETURN.
    ENDIF.
    DATA lv_date TYPE d.
    lv_date = iv_iso(4) && iv_iso+5(2) && iv_iso+8(2).
    DATA(lv_days) = lv_date - CONV d( '18991230' ).

    DATA lv_frac TYPE decfloat34.
    IF strlen( iv_iso ) >= 19 AND iv_iso+10(1) = 'T'.
      DATA(lv_secs) = CONV i( iv_iso+11(2) ) * 3600 + CONV i( iv_iso+14(2) ) * 60 + CONV i( iv_iso+17(2) ).
      lv_frac = CONV decfloat34( lv_secs ) / 86400.
    ENDIF.

    rv_serial = COND #( WHEN lv_frac = 0
                        THEN |{ lv_days }|
                        ELSE |{ CONV decfloat34( lv_days + lv_frac ) NUMBER = RAW }| ).
  ENDMETHOD.


  METHOD width_chars.
    " SSML ss:Width = points; OOXML width = số ký tự font mặc định
    TRY.
        DATA(lv_pts) = CONV decfloat34( iv_points ).
      CATCH cx_sy_conversion_error.
        rv_chars = `8.43`.
        RETURN.
    ENDTRY.
    DATA(lv_chars) = ( lv_pts / CONV decfloat34( '0.75' ) - 5 ) / 7.
    IF lv_chars < CONV decfloat34( '0.5' ).
      lv_chars = CONV decfloat34( '0.5' ).
    ENDIF.
    rv_chars = |{ CONV decfloat34( round( val = lv_chars dec = 2 ) ) NUMBER = RAW }|.
  ENDMETHOD.


  METHOD esc.
    rv_out = iv_text.
    REPLACE ALL OCCURRENCES OF `&` IN rv_out WITH `&amp;`.
    REPLACE ALL OCCURRENCES OF `<` IN rv_out WITH `&lt;`.
    REPLACE ALL OCCURRENCES OF `>` IN rv_out WITH `&gt;`.
  ENDMETHOD.


  METHOD esc_attr.
    rv_out = esc( iv_text ).
    REPLACE ALL OCCURRENCES OF `"` IN rv_out WITH `&quot;`.
  ENDMETHOD.


  METHOD attr.
    rv_val = io_elem->get_attribute_ns( name = iv_name uri = iv_uri ).
    IF rv_val IS INITIAL.
      rv_val = io_elem->get_attribute( iv_name ).      " attribute không namespace
    ENDIF.
  ENDMETHOD.


  METHOD elem_children.
    DATA(lo_node) = io_parent->get_first_child( ).
    WHILE lo_node IS BOUND.
      IF lo_node->get_type( ) = if_ixml_node=>co_node_element
      AND lo_node->get_name( ) = iv_name.
        APPEND CAST if_ixml_element( lo_node ) TO rt_elems.
      ENDIF.
      lo_node = lo_node->get_next( ).
    ENDWHILE.
  ENDMETHOD.


  METHOD find_child.
    DATA(lo_node) = io_parent->get_first_child( ).
    WHILE lo_node IS BOUND.
      IF lo_node->get_type( ) = if_ixml_node=>co_node_element
      AND lo_node->get_name( ) = iv_name.
        ro_elem = CAST if_ixml_element( lo_node ).
        RETURN.
      ENDIF.
      lo_node = lo_node->get_next( ).
    ENDWHILE.
  ENDMETHOD.


  METHOD child_value.
    DATA(lo_child) = find_child( io_parent = io_parent iv_name = iv_name ).
    IF lo_child IS BOUND.
      rv_val = lo_child->get_value( ).
    ENDIF.
  ENDMETHOD.


  METHOD zip_text.
    io_zip->add( name    = iv_name
                 content = cl_abap_conv_codepage=>create_out( )->convert( iv_text ) ).
  ENDMETHOD.
ENDCLASS.
