"! <p class="shorttext synchronized" lang="en">XLWB Cloud: SpreadsheetML template engine</p>
"!
"! Render an Excel form from a SpreadsheetML 2003 template (Excel: Save As
"! "XML Spreadsheet 2003") and an ABAP data context (REF TO DATA on a
"! structure). All layout (styles, number formats, row heights, column
"! widths, merges, print setup, data validation) comes from the template;
"! the engine only binds data via {{placeholder}} markup.
"!
"! Placeholder specification v1:
"! <ul>
"! <li> {{path}}              value; path = dot-separated components, resolved
"!                            case-insensitively against the context stack
"!                            (innermost loop line first, then outer, then root)</li>
"! <li> {{@index}} {{@count}} 1-based line number / total lines of nearest loop</li>
"! <li> {{#path}} .. {{/path}} row loop block: rows from opener row to closer
"!                            row (inclusive) are cloned once per line of the
"!                            internal table at path; blocks can be nested</li>
"! <li> {{?path}} .. {{/path}} conditional block: kept if value is not initial
"!                            (table: not empty)</li>
"! <li> {{^path}} .. {{/path}} inverse conditional block</li>
"! <li> {{#path&gt;}}             horizontal cell loop: the cell is cloned once per
"!                            line (dynamic columns)</li>
"! <li> {{sum:path.comp}} {{avg:..}} {{min:..}} {{max:..}} {{cnt:path}} aggregates</li>
"! <li> {{*break}}            page break before this row</li>
"! <li> {{*mergedown:path}}   ss:MergeDown = lines(path)-1 (dynamic row span)</li>
"! <li> {{*mergeacross:path}} ss:MergeAcross = lines(path)-1 (dynamic col span)</li>
"! <li> {{*mergedown=N}} / {{*mergeacross=N}} merge N cells; N is a literal
"!                            integer or a context path yielding a number</li>
"! <li> {{*mergesame:path}}   inside a row loop: merge consecutive lines that
"!                            share the same value of path (covered cells are
"!                            emptied). {{*mergesame&gt;:path}} does the same
"!                            horizontally inside a cell loop</li>
"! <li> Sheet loop: worksheet name contains {{#path}} - the whole worksheet is
"!                            cloned per line; rest of the name is a template
"!                            and must render to a unique sheet name</li>
"! </ul>
"! A cell whose entire content is one value/aggregate placeholder is written
"! typed (ss:Type Number / DateTime), so the template number format applies.
"! Excel formulas (ss:Formula, R1C1) are preserved untouched.
CLASS zcl_xlwb_engine DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    "! Render template + context to SpreadsheetML (string)
    METHODS render
      IMPORTING iv_template   TYPE string
                ir_context    TYPE REF TO data
      RETURNING VALUE(rv_xml) TYPE string
      RAISING   zcx_xlwb.

    "! Render to UTF-8 xstring (ready for download, extension .xls)
    METHODS render_xstring
      IMPORTING iv_template    TYPE string
                ir_context     TYPE REF TO data
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    CONSTANTS:
      mc_uri_ss TYPE string VALUE 'urn:schemas-microsoft-com:office:spreadsheet',
      mc_uri_x  TYPE string VALUE 'urn:schemas-microsoft-com:office:excel'.

    TYPES:
      ty_elems TYPE STANDARD TABLE OF REF TO if_ixml_element WITH EMPTY KEY.

    DATA:
      mo_doc    TYPE REF TO if_ixml_document,
      mo_ctx    TYPE REF TO zcl_xlwb_ctx,
      mt_breaks TYPE ty_elems.

    METHODS process_workbook
      RAISING zcx_xlwb.
    METHODS process_worksheet
      IMPORTING io_ws TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.
    METHODS normalize_row_index
      IMPORTING io_table TYPE REF TO if_ixml_element.
    METHODS process_rows
      IMPORTING it_rows TYPE ty_elems
      RAISING   zcx_xlwb.
    METHODS render_row
      IMPORTING io_row TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.
    METHODS render_cell
      IMPORTING io_cell TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.
    METHODS apply_cell_markers
      IMPORTING io_cell TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.
    METHODS render_free_nodes
      IMPORTING io_node TYPE REF TO if_ixml_node
      RAISING   zcx_xlwb.
    METHODS elem_children
      IMPORTING io_parent       TYPE REF TO if_ixml_element
                iv_name         TYPE string
      RETURNING VALUE(rt_elems) TYPE ty_elems.
    METHODS data_elems_of
      IMPORTING io_elem         TYPE REF TO if_ixml_element
      RETURNING VALUE(rt_elems) TYPE ty_elems.
    METHODS scan_strip
      IMPORTING io_row          TYPE REF TO if_ixml_element
                iv_pcre         TYPE string
      EXPORTING ev_found        TYPE abap_bool
                ev_sub1         TYPE string
                ev_sub2         TYPE string.
    METHODS apply_breaks
      IMPORTING io_ws    TYPE REF TO if_ixml_element
                io_table TYPE REF TO if_ixml_element.
ENDCLASS.



CLASS ZCL_XLWB_ENGINE IMPLEMENTATION.


  METHOD render.
    IF ir_context IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Context reference is initial' ).
    ENDIF.

    DATA(lo_ixml) = cl_ixml_core=>create( ).
    mo_doc = lo_ixml->create_document( ).
    DATA(lo_sf) = lo_ixml->create_stream_factory( ).

    DATA(lv_xtemplate) = cl_abap_conv_codepage=>create_out( )->convert( iv_template ).
    DATA(lo_parser) = lo_ixml->create_parser(
      document       = mo_doc
      istream        = lo_sf->create_istream_xstring( lv_xtemplate )
      stream_factory = lo_sf ).
    lo_parser->set_namespace_mode( if_ixml_parser_core=>co_namespace_aware ).
    IF lo_parser->parse( ) <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Template is not valid XML (SpreadsheetML expected)' ).
    ENDIF.

    CLEAR mt_breaks.
    mo_ctx = NEW zcl_xlwb_ctx( ir_context ).

    process_workbook( ).

    DATA lv_xout TYPE xstring.
    DATA(lo_renderer) = lo_ixml->create_renderer(
      document = mo_doc
      ostream  = lo_sf->create_ostream_xstring( lv_xout ) ).
    lo_renderer->render( ).
    rv_xml = cl_abap_conv_codepage=>create_in( )->convert( lv_xout ).
  ENDMETHOD.


  METHOD render_xstring.
    rv_file = cl_abap_conv_codepage=>create_out( )->convert(
      render( iv_template = iv_template ir_context = ir_context ) ).
  ENDMETHOD.


  METHOD process_workbook.
    DATA(lo_root) = mo_doc->get_root_element( ).
    IF lo_root IS INITIAL OR lo_root->get_name( ) <> 'Workbook'.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Template root element is not <Workbook>' ).
    ENDIF.

    LOOP AT elem_children( io_parent = lo_root iv_name = 'Worksheet' ) INTO DATA(lo_ws).
      DATA(lv_name) = lo_ws->get_attribute_ns( name = 'Name' uri = mc_uri_ss ).

      DATA lv_path TYPE string.
      CLEAR lv_path.
      FIND PCRE '(?-x)\{\{#([^}>]+)\}\}' IN lv_name SUBMATCHES lv_path.
      IF sy-subrc = 0.
        " ---- sheet loop: clone worksheet per table line -------------------
        DATA(lv_name_tpl) = lv_name.
        REPLACE FIRST OCCURRENCE OF PCRE '(?-x)\{\{#[^}>]+\}\}' IN lv_name_tpl WITH ''.

        FIELD-SYMBOLS <lt_sheet> TYPE INDEX TABLE.
        DATA(lr_sheet_tab) = mo_ctx->resolve_table( iv_path = lv_path iv_from = 'sheetloop' ).
        ASSIGN lr_sheet_tab->* TO <lt_sheet>.
        DATA(lv_count) = lines( <lt_sheet> ).

        DO lv_count TIMES.
          DATA(lv_idx) = sy-index.
          READ TABLE <lt_sheet> INDEX lv_idx ASSIGNING FIELD-SYMBOL(<ls_sheet>).
          DATA(lo_clone) = CAST if_ixml_element( lo_ws->clone( ) ).
          lo_root->insert_child( new_child = lo_clone ref_child = lo_ws ).
          mo_ctx->push( ir_data  = REF #( <ls_sheet> ) iv_index = lv_idx
                        iv_count = lv_count iv_kind = 'SHEET'
                        ir_table = lr_sheet_tab ).
          lo_clone->set_attribute_ns(
            name   = 'Name'
            prefix = 'ss'
            uri    = mc_uri_ss
            value  = mo_ctx->render_text( lv_name_tpl ) ).
          process_worksheet( lo_clone ).
          mo_ctx->pop( ).
        ENDDO.
        lo_root->remove_child( lo_ws ).
      ELSE.
        IF lv_name CS '{{'.
          lo_ws->set_attribute_ns(
            name   = 'Name'
            prefix = 'ss'
            uri    = mc_uri_ss
            value  = mo_ctx->render_text( lv_name ) ).
        ENDIF.
        process_worksheet( lo_ws ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD process_worksheet.
    CLEAR mt_breaks.

    DATA(lo_table) = io_ws->find_from_name_ns( name = 'Table' uri = mc_uri_ss ).
    IF lo_table IS INITIAL.
      RETURN.
    ENDIF.

    " counts become wrong after cloning - Excel recomputes when absent
    lo_table->remove_attribute_ns( name = 'ExpandedRowCount'    uri = mc_uri_ss ).
    lo_table->remove_attribute_ns( name = 'ExpandedColumnCount' uri = mc_uri_ss ).

    normalize_row_index( lo_table ).

    process_rows( elem_children( io_parent = lo_table iv_name = 'Row' ) ).

    " placeholders outside <Table> (e.g. WorksheetOptions / DataValidation lists)
    DATA(lo_child) = io_ws->get_first_child( ).
    WHILE lo_child IS BOUND.
      IF lo_child <> lo_table.
        render_free_nodes( lo_child ).
      ENDIF.
      lo_child = lo_child->get_next( ).
    ENDWHILE.

    apply_breaks( io_ws = io_ws io_table = lo_table ).
  ENDMETHOD.


  METHOD normalize_row_index.
    " Excel sinh <Row ss:Index="N"> de NHAY qua cac dong trong. Do la neo dong
    " TUYET DOI: sau khi engine nhan ban vung lap, so dong thay doi nen moi dong
    " phia sau se nam sai cho. Thay moi neo bang dung so <Row/> rong.
    " (ss:Index tren <Cell> la neo COT - giu nguyen, khong dung o day.)
    DATA(lv_current) = 0.

    LOOP AT elem_children( io_parent = io_table iv_name = 'Row' ) INTO DATA(lo_row).
      DATA(lv_index) = lo_row->get_attribute_ns( name = 'Index' uri = mc_uri_ss ).

      IF lv_index IS NOT INITIAL.
        DATA(lv_target) = CONV i( lv_index ).
        WHILE lv_current + 1 < lv_target.
          " clone( 0 ) = chi node, khong con -> giu dung namespace cua template
          " (create_element_ns sinh prefix khac lam Excel bo qua dong)
          DATA(lo_filler) = CAST if_ixml_element( lo_row->clone( 0 ) ).
          lo_filler->remove_attribute_ns( name = 'Index' uri = mc_uri_ss ).
          io_table->insert_child( new_child = lo_filler ref_child = lo_row ).
          lv_current += 1.
        ENDWHILE.
        lo_row->remove_attribute_ns( name = 'Index' uri = mc_uri_ss ).
        lv_current = lv_target.
      ELSE.
        lv_current += 1.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD process_rows.
    DATA(lv_i) = 1.
    WHILE lv_i <= lines( it_rows ).
      DATA(lo_row) = it_rows[ lv_i ].

      scan_strip(
        EXPORTING io_row  = lo_row
                  iv_pcre = '\{\{([#?^])([^}>:]+)\}\}'
        IMPORTING ev_found = DATA(lv_found)
                  ev_sub1  = DATA(lv_kind)
                  ev_sub2  = DATA(lv_path) ).

      IF lv_found = abap_false.
        render_row( lo_row ).
        lv_i += 1.
        CONTINUE.
      ENDIF.

      " ---- find the closing row {{/path}} ---------------------------------
      DATA(lv_path_pcre) = replace( val = lv_path sub = '.' with = '\.' occ = 0 ).
      DATA(lv_j) = 0.
      DATA(lv_k) = lv_i.
      WHILE lv_k <= lines( it_rows ).
        scan_strip(
          EXPORTING io_row  = it_rows[ lv_k ]
                    iv_pcre = |\\\{\\\{(/)({ lv_path_pcre })\\\}\\\}|
          IMPORTING ev_found = DATA(lv_closed) ).
        IF lv_closed = abap_true.
          lv_j = lv_k.
          EXIT.
        ENDIF.
        lv_k += 1.
      ENDWHILE.
      IF lv_j = 0.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |Missing closing marker \{\{/{ lv_path }\}\} for \{\{{ lv_kind }{ lv_path }\}\}| ).
      ENDIF.

      DATA(lt_block) = VALUE ty_elems( FOR n = lv_i THEN n + 1 UNTIL n > lv_j ( it_rows[ n ] ) ).
      DATA(lo_parent) = it_rows[ lv_i ]->get_parent( ).

      CASE lv_kind.
        WHEN '#'.
          FIELD-SYMBOLS <lt_loop> TYPE INDEX TABLE.
          DATA(lr_loop_tab) = mo_ctx->resolve_table( iv_path = lv_path iv_from = 'rowloop' ).
          ASSIGN lr_loop_tab->* TO <lt_loop>.
          DATA(lv_count) = lines( <lt_loop> ).

          DO lv_count TIMES.
            DATA(lv_idx) = sy-index.
            READ TABLE <lt_loop> INDEX lv_idx ASSIGNING FIELD-SYMBOL(<ls_line>).
            mo_ctx->push( ir_data  = REF #( <ls_line> ) iv_index = lv_idx
                          iv_count = lv_count iv_kind = 'ROW'
                          ir_table = lr_loop_tab ).

            DATA(lt_clones) = VALUE ty_elems( ).
            LOOP AT lt_block INTO DATA(lo_tpl_row).
              DATA(lo_row_clone) = CAST if_ixml_element( lo_tpl_row->clone( ) ).
              " loop bodies must be contiguous - absolute row anchors would clash
              lo_row_clone->remove_attribute_ns( name = 'Index' uri = mc_uri_ss ).
              lo_parent->insert_child( new_child = lo_row_clone ref_child = lt_block[ 1 ] ).
              APPEND lo_row_clone TO lt_clones.
            ENDLOOP.
            process_rows( lt_clones ).

            mo_ctx->pop( ).
          ENDDO.

          LOOP AT lt_block INTO DATA(lo_del_row).
            lo_parent->remove_child( lo_del_row ).
          ENDLOOP.

        WHEN '?' OR '^'.
          DATA(lv_truthy) = mo_ctx->is_truthy( lv_path ).
          IF ( lv_kind = '?' AND lv_truthy = abap_true )
          OR ( lv_kind = '^' AND lv_truthy = abap_false ).
            process_rows( lt_block ).
          ELSE.
            LOOP AT lt_block INTO DATA(lo_drop_row).
              lo_parent->remove_child( lo_drop_row ).
            ENDLOOP.
          ENDIF.
      ENDCASE.

      lv_i = lv_j + 1.
    ENDWHILE.
  ENDMETHOD.


  METHOD render_row.
    " page break marker anywhere in the row
    scan_strip(
      EXPORTING io_row  = io_row
                iv_pcre = '\{\{(\*)(break)\}\}'
      IMPORTING ev_found = DATA(lv_break) ).
    IF lv_break = abap_true.
      " KHÔNG lưu ref wrapper (iXML trả wrapper MỚI mỗi lần duyệt nên so sánh
      " ref luôn sai) — đánh dấu dòng bằng attribute tạm, apply_breaks đọc lại
      io_row->set_attribute( name = 'xlwbbreak' value = 'X' ).
    ENDIF.

    " {{*group=N}}: dòng thuộc nhóm outline cấp N (Excel +/- bên lề trái).
    " XML 2003 không chứa outline nên đánh dấu bằng attribute — converter
    " ssml2xlsx chuyển thành outlineLevel của OOXML; Excel bỏ qua attribute
    " lạ nếu file rơi về nhánh .xls thô (đã kiểm chứng)
    scan_strip(
      EXPORTING io_row  = io_row
                iv_pcre = '\{\{\*group=(\d)\}\}()'
      IMPORTING ev_found = DATA(lv_grp)
                ev_sub1  = DATA(lv_grp_level) ).
    IF lv_grp = abap_true.
      io_row->set_attribute( name = 'xlwboutline' value = lv_grp_level ).
    ENDIF.

    DATA(lv_delta) = 0.
    LOOP AT elem_children( io_parent = io_row iv_name = 'Cell' ) INTO DATA(lo_cell).

      " shift explicit column anchors after horizontal cell expansion
      IF lv_delta > 0.
        DATA(lv_cix) = lo_cell->get_attribute_ns( name = 'Index' uri = mc_uri_ss ).
        IF lv_cix IS NOT INITIAL.
          lo_cell->set_attribute_ns(
            name   = 'Index'
            prefix = 'ss'
            uri    = mc_uri_ss
            value  = |{ CONV i( lv_cix ) + lv_delta }| ).
        ENDIF.
      ENDIF.

      " ---- horizontal cell loop {{#path>}} --------------------------------
      scan_strip(
        EXPORTING io_row  = lo_cell
                  iv_pcre = '\{\{#([^}]+)(>)\}\}'
        IMPORTING ev_found = DATA(lv_cloop)
                  ev_sub1  = DATA(lv_cpath) ).
      IF lv_cloop = abap_true.
        FIELD-SYMBOLS <lt_cells> TYPE INDEX TABLE.
        DATA(lr_cell_tab) = mo_ctx->resolve_table( iv_path = lv_cpath iv_from = 'cellloop' ).
        ASSIGN lr_cell_tab->* TO <lt_cells>.
        DATA(lv_ccount) = lines( <lt_cells> ).

        DO lv_ccount TIMES.
          DATA(lv_cidx) = sy-index.
          READ TABLE <lt_cells> INDEX lv_cidx ASSIGNING FIELD-SYMBOL(<ls_cell>).
          DATA(lo_cell_clone) = CAST if_ixml_element( lo_cell->clone( ) ).
          IF lv_cidx > 1.
            " only the first clone may keep the template's column anchor
            lo_cell_clone->remove_attribute_ns( name = 'Index' uri = mc_uri_ss ).
          ENDIF.
          io_row->insert_child( new_child = lo_cell_clone ref_child = lo_cell ).
          mo_ctx->push( ir_data  = REF #( <ls_cell> ) iv_index = lv_cidx
                        iv_count = lv_ccount iv_kind = 'CELL'
                        ir_table = lr_cell_tab ).
          render_cell( lo_cell_clone ).
          mo_ctx->pop( ).
        ENDDO.

        io_row->remove_child( lo_cell ).
        lv_delta += lv_ccount - 1.
        CONTINUE.
      ENDIF.

      render_cell( lo_cell ).
    ENDLOOP.
  ENDMETHOD.


  METHOD apply_cell_markers.
    " dynamic span markers - evaluated per cell (also on cell-loop clones,
    " where the table driving the span is the one of the current frame)
    scan_strip(
      EXPORTING io_row  = io_cell
                iv_pcre = '\{\{\*mergedown:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_md)
                ev_sub1  = DATA(lv_md_path) ).
    IF lv_md = abap_true.
      FIELD-SYMBOLS <lt_md> TYPE INDEX TABLE.
      DATA(lr_md_tab) = mo_ctx->resolve_table( iv_path = lv_md_path iv_from = 'mergedown' ).
      ASSIGN lr_md_tab->* TO <lt_md>.
      IF lines( <lt_md> ) > 1.
        io_cell->set_attribute_ns(
          name   = 'MergeDown'
          prefix = 'ss'
          uri    = mc_uri_ss
          value  = |{ lines( <lt_md> ) - 1 }| ).
      ENDIF.
    ENDIF.

    scan_strip(
      EXPORTING io_row  = io_cell
                iv_pcre = '\{\{\*mergeacross:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_ma)
                ev_sub1  = DATA(lv_ma_path) ).
    IF lv_ma = abap_true.
      FIELD-SYMBOLS <lt_ma> TYPE INDEX TABLE.
      DATA(lr_ma_tab) = mo_ctx->resolve_table( iv_path = lv_ma_path iv_from = 'mergeacross' ).
      ASSIGN lr_ma_tab->* TO <lt_ma>.
      IF lines( <lt_ma> ) > 1.
        io_cell->set_attribute_ns(
          name   = 'MergeAcross'
          prefix = 'ss'
          uri    = mc_uri_ss
          value  = |{ lines( <lt_ma> ) - 1 }| ).
      ENDIF.
    ENDIF.

    " ---- span so hoc: {{*mergedown=N}} / {{*mergeacross=N}} --------------
    " N = so O duoc gop (khong phai gia tri thuoc tinh) -> attribute = N-1.
    " N viet truc tiep ('3') hoac la path tren context ('rowspan').
    scan_strip(
      EXPORTING io_row  = io_cell
                iv_pcre = '\{\{\*mergedown=([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_mdn)
                ev_sub1  = DATA(lv_mdn_spec) ).
    IF lv_mdn = abap_true.
      DATA(lv_mdn_num) = mo_ctx->span_of( lv_mdn_spec ).
      IF lv_mdn_num > 1.
        io_cell->set_attribute_ns( name   = 'MergeDown'
                                   prefix = 'ss'
                                   uri    = mc_uri_ss
                                   value  = |{ lv_mdn_num - 1 }| ).
      ENDIF.
    ENDIF.

    scan_strip(
      EXPORTING io_row  = io_cell
                iv_pcre = '\{\{\*mergeacross=([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_man)
                ev_sub1  = DATA(lv_man_spec) ).
    IF lv_man = abap_true.
      DATA(lv_man_num) = mo_ctx->span_of( lv_man_spec ).
      IF lv_man_num > 1.
        io_cell->set_attribute_ns( name   = 'MergeAcross'
                                   prefix = 'ss'
                                   uri    = mc_uri_ss
                                   value  = |{ lv_man_num - 1 }| ).
      ENDIF.
    ENDIF.

    " ---- gop o theo gia tri: {{*mergesame:path}} / {{*mergesame>:path}} ---
    " Dong dau nhom nhan MergeDown/MergeAcross = so phan tu cua nhom - 1;
    " cac o con lai bi merge PHU nen phai xoa noi dung, neu khong Excel se
    " thay hai o cung khai bao du lieu trong mot vung gop.
    scan_strip(
      EXPORTING io_row  = io_cell
                iv_pcre = '\{\{\*mergesame(>?):([^}]+)\}\}'
      IMPORTING ev_found = DATA(lv_ms)
                ev_sub1  = DATA(lv_ms_dir)
                ev_sub2  = DATA(lv_ms_path) ).
    IF lv_ms = abap_true.
      DATA(lv_ms_kind) = COND string( WHEN lv_ms_dir = '>' THEN 'CELL' ELSE 'ROW' ).
      mo_ctx->group_span( EXPORTING iv_path     = lv_ms_path
                                    iv_kind     = lv_ms_kind
                          IMPORTING ev_is_first = DATA(lv_ms_first)
                                    ev_span     = DATA(lv_ms_span) ).
      IF lv_ms_first = abap_true.
        IF lv_ms_span > 1.
          io_cell->set_attribute_ns(
            name   = COND string( WHEN lv_ms_dir = '>' THEN 'MergeAcross' ELSE 'MergeDown' )
            prefix = 'ss'
            uri    = mc_uri_ss
            value  = |{ lv_ms_span - 1 }| ).
        ENDIF.
      ELSE.
        " O bi phu: phai GO HAN <Cell> khoi dong — SpreadsheetML cam moi cell
        " (ke ca cell rong) nam trong vung MergeDown/MergeAcross, neu con thi
        " Excel bao "file is corrupt" va tu choi mo. Truoc khi go, neo ss:Index
        " cho o ke tiep de cac o phia sau giu nguyen cot.
        DATA(lo_row_parent) = CAST if_ixml_element( io_cell->get_parent( ) ).
        DATA lo_next TYPE REF TO if_ixml_element.
        CLEAR lo_next.
        DATA(lv_col) = 0.
        DATA(lv_found_self) = abap_false.
        " Nhan dien "chinh minh" bang attribute tam: iXML tra wrapper object
        " MOI cho cung mot node moi lan duyet, nen so sanh ref (=) luon sai.
        " Attribute nay bien mat cung cell khi remove_child ben duoi.
        io_cell->set_attribute( name = 'xlwbself' value = 'X' ).
        LOOP AT elem_children( io_parent = lo_row_parent iv_name = 'Cell' ) INTO DATA(lo_scan).
          IF lv_found_self = abap_true.
            lo_next = lo_scan.
            EXIT.
          ENDIF.
          DATA(lv_scan_ix) = lo_scan->get_attribute_ns( name = 'Index' uri = mc_uri_ss ).
          lv_col = COND #( WHEN lv_scan_ix IS NOT INITIAL
                           THEN CONV i( lv_scan_ix ) ELSE lv_col + 1 ).
          IF lo_scan->get_attribute( 'xlwbself' ) = 'X'.
            lv_found_self = abap_true.
          ELSE.
            DATA(lv_scan_ma) = lo_scan->get_attribute_ns( name = 'MergeAcross' uri = mc_uri_ss ).
            IF lv_scan_ma IS NOT INITIAL.
              lv_col = lv_col + CONV i( lv_scan_ma ).
            ENDIF.
          ENDIF.
        ENDLOOP.
        IF lo_next IS BOUND
        AND lo_next->get_attribute_ns( name = 'Index' uri = mc_uri_ss ) IS INITIAL.
          DATA(lv_self_ma) = io_cell->get_attribute_ns( name = 'MergeAcross' uri = mc_uri_ss ).
          DATA(lv_next_col) = lv_col + 1 +
            COND i( WHEN lv_self_ma IS NOT INITIAL THEN CONV i( lv_self_ma ) ELSE 0 ).
          lo_next->set_attribute_ns( name = 'Index' prefix = 'ss' uri = mc_uri_ss
                                     value = |{ lv_next_col }| ).
        ENDIF.
        LOOP AT data_elems_of( io_cell ) INTO DATA(lo_drop).
          io_cell->remove_child( lo_drop ).
        ENDLOOP.
        lo_row_parent->remove_child( io_cell ).
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD render_cell.
    apply_cell_markers( io_cell ).

    LOOP AT data_elems_of( io_cell ) INTO DATA(lo_data).
      DATA(lv_text) = lo_data->get_value( ).
      IF lv_text NS '{{'.
        CONTINUE.
      ENDIF.

      " whole cell = one placeholder -> typed output (template numfmt applies)
      DATA lv_token TYPE string.
      CLEAR lv_token.
      FIND PCRE '(?-x)^\s*\{\{([^}]+)\}\}\s*$' IN lv_text SUBMATCHES lv_token.
      IF sy-subrc = 0 AND substring( val = lv_token off = 0 len = 1 ) NA '#/?^*'.
        DATA(ls_value) = mo_ctx->token_value( lv_token ).
        IF ls_value-is_number = abap_true.
          lo_data->set_value( ls_value-text ).
          lo_data->set_attribute_ns( name = 'Type' prefix = 'ss' uri = mc_uri_ss
                                     value = 'Number' ).
        ELSEIF ls_value-is_date = abap_true.
          IF ls_value-date_value IS INITIAL.
            lo_data->set_value( `` ).
          ELSE.
            lo_data->set_value( |{ ls_value-date_value DATE = ISO }T00:00:00.000| ).
            lo_data->set_attribute_ns( name = 'Type' prefix = 'ss' uri = mc_uri_ss
                                       value = 'DateTime' ).
          ENDIF.
        ELSEIF ls_value-is_time = abap_true.
          lo_data->set_value( |1899-12-31T{ ls_value-time_value TIME = ISO }.000| ).
          lo_data->set_attribute_ns( name = 'Type' prefix = 'ss' uri = mc_uri_ss
                                     value = 'DateTime' ).
        ELSE.
          lo_data->set_value( ls_value-text ).
        ENDIF.
      ELSE.
        lo_data->set_value( mo_ctx->render_text( lv_text ) ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD render_free_nodes.
    IF io_node->get_type( ) = if_ixml_node=>co_node_text.
      DATA(lv_text) = io_node->get_value( ).
      IF lv_text CS '{{'.
        io_node->set_value( mo_ctx->render_text( lv_text ) ).
      ENDIF.
      RETURN.
    ENDIF.

    " placeholder trong ATTRIBUTE (vd AutoFilter x:Range="R3C1:R3C{{lastcol}}")
    IF io_node->get_type( ) = if_ixml_node=>co_node_element.
      DATA(lo_attrs) = io_node->get_attributes( ).
      IF lo_attrs IS BOUND.
        DO lo_attrs->get_length( ) TIMES.
          DATA(lo_attr) = lo_attrs->get_item( sy-index - 1 ).
          DATA(lv_aval) = lo_attr->get_value( ).
          IF lv_aval CS '{{'.
            lo_attr->set_value( mo_ctx->render_text( lv_aval ) ).
          ENDIF.
        ENDDO.
      ENDIF.
    ENDIF.

    DATA(lo_child) = io_node->get_first_child( ).
    WHILE lo_child IS BOUND.
      DATA(lo_next) = lo_child->get_next( ).
      render_free_nodes( lo_child ).
      lo_child = lo_next.
    ENDWHILE.
  ENDMETHOD.


  METHOD elem_children.
    DATA(lo_node) = io_parent->get_first_child( ).
    WHILE lo_node IS BOUND.
      IF lo_node->get_name( ) = iv_name.
        APPEND CAST if_ixml_element( lo_node ) TO rt_elems.
      ENDIF.
      lo_node = lo_node->get_next( ).
    ENDWHILE.
  ENDMETHOD.


  METHOD data_elems_of.
    DATA(lo_ncol) = io_elem->get_elements_by_tag_name_ns( name = 'Data' uri = mc_uri_ss ).
    DATA(lv_n) = lo_ncol->get_length( ).
    DO lv_n TIMES.
      APPEND CAST if_ixml_element( lo_ncol->get_item( sy-index - 1 ) ) TO rt_elems.
    ENDDO.
  ENDMETHOD.


  METHOD scan_strip.
    CLEAR: ev_found, ev_sub1, ev_sub2.
    LOOP AT data_elems_of( io_row ) INTO DATA(lo_data).
      DATA(lv_text) = lo_data->get_value( ).
      IF lv_text NS '{{'.
        CONTINUE.
      ENDIF.
      DATA(lv_pcre) = `(?-x)` && iv_pcre.
      FIND PCRE lv_pcre IN lv_text SUBMATCHES ev_sub1 ev_sub2.
      IF sy-subrc = 0.
        REPLACE FIRST OCCURRENCE OF PCRE lv_pcre IN lv_text WITH ''.
        lo_data->set_value( lv_text ).
        ev_found = abap_true.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD apply_breaks.
    " gom vị trí các dòng được đánh dấu {{*break}} (attribute tạm từ render_row)
    DATA lt_pos TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    LOOP AT elem_children( io_parent = io_table iv_name = 'Row' ) INTO DATA(lo_row).
      DATA(lv_tabix) = sy-tabix.
      IF lo_row->get_attribute( 'xlwbbreak' ) = 'X'.
        lo_row->remove_attribute( 'xlwbbreak' ).
        IF lv_tabix > 1. " break trước dòng đầu tiên vô nghĩa
          APPEND lv_tabix - 1 TO lt_pos.
        ENDIF.
      ENDIF.
    ENDLOOP.

    IF lt_pos IS INITIAL.
      RETURN.
    ENDIF.

    DATA(lo_pb)  = mo_doc->create_element_ns( name = 'PageBreaks' prefix = 'x' uri = mc_uri_x ).
    io_ws->append_child( lo_pb ).
    DATA(lo_rbs) = mo_doc->create_simple_element_ns(
      name = 'RowBreaks' parent = lo_pb prefix = 'x' uri = mc_uri_x ).

    LOOP AT lt_pos INTO DATA(lv_pos).
      DATA(lo_rb) = mo_doc->create_simple_element_ns(
        name = 'RowBreak' parent = lo_rbs prefix = 'x' uri = mc_uri_x ).
      mo_doc->create_simple_element_ns(
        name = 'Row' parent = lo_rb prefix = 'x' uri = mc_uri_x
        value = |{ lv_pos }| ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
