"! <p class="shorttext synchronized" lang="en">XLWB Cloud: XLSX (OOXML) template engine</p>
"!
"! Engine thứ hai của XLWB Cloud: template là file **.xlsx thật** (Excel:
"! Save As Workbook), nên giữ được mọi thứ SpreadsheetML 2003 không có —
"! **ảnh/logo nhúng**, chart, conditional formatting — và file mở ra không
"! có cảnh báo định dạng.
"!
"! Cú pháp placeholder giống engine SpreadsheetML (dùng chung
"! {@link zcl_xlwb_ctx}), cộng thêm marker ảnh:
"! <ul>
"! <li> {{*image:path}}        chèn ảnh base64 (giá trị tại path) neo ở ô này,
"!                            mặc định rộng 2 cột cao 4 dòng</li>
"! <li> {{*image:path@3x5}}    chỉ định số cột x số dòng</li>
"! </ul>
"! Hỗ trợ: giá trị (shared string + inline string), loop dòng lồng nhau,
"! khối điều kiện, aggregate, {{*mergedown/mergeacross}} (dạng :path và =N),
"! {{*mergesame:path}} + {{*mergesame&gt;:path}}, ảnh, cell loop ngang
"! {{#path&gt;}} (cột động), sheet loop (tên sheet chứa {{#path}}),
"! {{*break}} (page break) — ngang tính năng engine SpreadsheetML.
CLASS zcl_xlwb_xlsx DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    "! Render template .xlsx + context → .xlsx
    METHODS render
      IMPORTING iv_template    TYPE xstring
                ir_context     TYPE REF TO data
      RETURNING VALUE(rv_file) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    CONSTANTS:
      mc_main TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      mc_rel  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      mc_pkg  TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
      mc_ct   TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/content-types'.

    TYPES:
      ty_elems  TYPE STANDARD TABLE OF REF TO if_ixml_element WITH EMPTY KEY,
      ty_shared TYPE STANDARD TABLE OF string WITH EMPTY KEY,
      BEGIN OF ty_merge,
        col_from TYPE string,
        col_to   TYPE string,
        row_from TYPE i,
        row_to   TYPE i,
      END OF ty_merge,
      ty_merges TYPE STANDARD TABLE OF ty_merge WITH EMPTY KEY,
      BEGIN OF ty_rowmap,
        origin TYPE i,
        new    TYPE i,
      END OF ty_rowmap,
      ty_rowmaps TYPE STANDARD TABLE OF ty_rowmap WITH EMPTY KEY,
      BEGIN OF ty_dynmerge,
        cell  TYPE REF TO if_ixml_element,
        rows  TYPE i,
        cols  TYPE i,
      END OF ty_dynmerge,
      ty_dynmerges TYPE STANDARD TABLE OF ty_dynmerge WITH EMPTY KEY,
      BEGIN OF ty_image,
        cell   TYPE REF TO if_ixml_element,
        data   TYPE xstring,
        cols   TYPE i,
        rows   TYPE i,
      END OF ty_image,
      ty_images TYPE STANDARD TABLE OF ty_image WITH EMPTY KEY.

    DATA:
      mo_zip       TYPE REF TO cl_abap_zip,
      mo_ctx       TYPE REF TO zcl_xlwb_ctx,
      mo_ixml      TYPE REF TO if_ixml_core,
      mo_sf        TYPE REF TO if_ixml_stream_factory_core,
      mo_doc       TYPE REF TO if_ixml_document,
      mt_shared    TYPE ty_shared,
      mt_merges    TYPE ty_merges,
      mt_dynmerges TYPE ty_dynmerges,
      mt_images    TYPE ty_images,
      mv_img_count TYPE i,
      mv_ss_dirty  TYPE abap_bool.

    METHODS parse
      IMPORTING iv_xml        TYPE xstring
      RETURNING VALUE(ro_doc) TYPE REF TO if_ixml_document
      RAISING   zcx_xlwb.
    METHODS serialize
      IMPORTING io_doc        TYPE REF TO if_ixml_document
      RETURNING VALUE(rv_xml) TYPE xstring.
    METHODS zip_read
      IMPORTING iv_name       TYPE string
      RETURNING VALUE(rv_raw) TYPE xstring.
    METHODS zip_write
      IMPORTING iv_name TYPE string
                iv_raw  TYPE xstring.

    METHODS load_shared_strings.
    METHODS shared_index
      IMPORTING iv_text       TYPE string
      RETURNING VALUE(rv_idx) TYPE i.
    METHODS save_shared_strings.
    METHODS escape_xml
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_text) TYPE string.
    METHODS sheet_paths
      RETURNING VALUE(rt_paths) TYPE string_table
      RAISING   zcx_xlwb.

    METHODS process_workbook
      RAISING   zcx_xlwb.
    METHODS process_sheet
      IMPORTING iv_path TYPE string
      RAISING   zcx_xlwb.
    METHODS process_sheet_raw
      IMPORTING iv_raw  TYPE xstring
                iv_path TYPE string
      RAISING   zcx_xlwb.
    METHODS expand_cell_loop
      IMPORTING io_row  TYPE REF TO if_ixml_element
                io_cell TYPE REF TO if_ixml_element
                iv_path TYPE string
      RAISING   zcx_xlwb.
    METHODS ct_add_override
      IMPORTING iv_partname TYPE string
                iv_type     TYPE string.
    METHODS process_rows
      IMPORTING it_rows TYPE ty_elems
      RAISING   zcx_xlwb.
    METHODS render_row
      IMPORTING io_row TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.
    METHODS render_cell
      IMPORTING io_cell TYPE REF TO if_ixml_element
      RAISING   zcx_xlwb.

    METHODS cell_text
      IMPORTING io_cell        TYPE REF TO if_ixml_element
      RETURNING VALUE(rv_text) TYPE string.
    METHODS set_cell_value
      IMPORTING io_cell  TYPE REF TO if_ixml_element
                is_value TYPE zcl_xlwb_ctx=>ty_value.
    METHODS set_cell_text
      IMPORTING io_cell TYPE REF TO if_ixml_element
                iv_text TYPE string.

    METHODS scan_strip
      IMPORTING io_scope TYPE REF TO if_ixml_element
                iv_pcre  TYPE string
      EXPORTING ev_found TYPE abap_bool
                ev_sub1  TYPE string
                ev_sub2  TYPE string.

    METHODS elem_children
      IMPORTING io_parent       TYPE REF TO if_ixml_element
                iv_name         TYPE string
      RETURNING VALUE(rt_elems) TYPE ty_elems.
    METHODS cells_of
      IMPORTING io_scope        TYPE REF TO if_ixml_element
      RETURNING VALUE(rt_elems) TYPE ty_elems.

    METHODS collect_merges
      IMPORTING io_ws TYPE REF TO if_ixml_element.
    METHODS finalize_sheet
      IMPORTING io_ws   TYPE REF TO if_ixml_element
                iv_path TYPE string
      RAISING   zcx_xlwb.
    METHODS attach_images
      IMPORTING io_ws   TYPE REF TO if_ixml_element
                iv_path TYPE string
      RAISING   zcx_xlwb.
    METHODS register_png_type.

    METHODS split_ref
      IMPORTING iv_ref TYPE string
      EXPORTING ev_col TYPE string
                ev_row TYPE i.
    METHODS col_to_num
      IMPORTING iv_col        TYPE string
      RETURNING VALUE(rv_num) TYPE i.
    METHODS num_to_col
      IMPORTING iv_num        TYPE i
      RETURNING VALUE(rv_col) TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_XLSX IMPLEMENTATION.


  METHOD render.
    mo_ctx  = NEW zcl_xlwb_ctx( ir_context ).
    mo_ixml = cl_ixml_core=>create( ).
    mo_sf   = mo_ixml->create_stream_factory( ).

    mo_zip = NEW cl_abap_zip( ).
    mo_zip->load( iv_template ).
    IF lines( mo_zip->files ) = 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Template is not a valid .xlsx (zip) file' ).
    ENDIF.

    CLEAR: mt_images, mv_img_count, mv_ss_dirty.
    load_shared_strings( ).

    process_workbook( ).

    save_shared_strings( ).

    rv_file = mo_zip->save( ).
  ENDMETHOD.


  METHOD process_workbook.
    " Điều phối worksheet: sheet thường render tại chỗ; tên sheet chứa
    " {{#path}} thì CẢ sheet được nhân bản mỗi dòng của bảng (sheet loop).
    DATA(lv_rels_raw) = zip_read( 'xl/_rels/workbook.xml.rels' ).
    DATA(lv_wb_raw)   = zip_read( 'xl/workbook.xml' ).
    IF lv_rels_raw IS INITIAL OR lv_wb_raw IS INITIAL.
      LOOP AT sheet_paths( ) INTO DATA(lv_fallback).
        process_sheet( lv_fallback ).
      ENDLOOP.
      RETURN.
    ENDIF.

    DATA(lo_rels_doc) = parse( lv_rels_raw ).
    DATA(lo_wb_doc)   = parse( lv_wb_raw ).
    DATA(lo_rels_root) = lo_rels_doc->get_root_element( ).
    DATA(lo_wb_root)   = lo_wb_doc->get_root_element( ).
    DATA(lo_sheets)    = lo_wb_root->find_from_name_ns( name = 'sheets' uri = mc_main ).
    IF lo_sheets IS INITIAL.
      RETURN.
    ENDIF.

    " map rId -> target + sheetId lớn nhất đang dùng
    TYPES: BEGIN OF ty_rel, rid TYPE string, target TYPE string, END OF ty_rel.
    DATA lt_rels TYPE STANDARD TABLE OF ty_rel WITH EMPTY KEY.
    LOOP AT elem_children( io_parent = lo_rels_root iv_name = 'Relationship' ) INTO DATA(lo_rel).
      APPEND VALUE ty_rel( rid    = lo_rel->get_attribute( 'Id' )
                           target = lo_rel->get_attribute( 'Target' ) ) TO lt_rels.
    ENDLOOP.

    DATA(lv_max_id) = 0.
    LOOP AT elem_children( io_parent = lo_sheets iv_name = 'sheet' ) INTO DATA(lo_sh).
      DATA(lv_sid) = lo_sh->get_attribute( 'sheetId' ).
      IF lv_sid CO '0123456789' AND lv_sid IS NOT INITIAL AND CONV i( lv_sid ) > lv_max_id.
        lv_max_id = CONV i( lv_sid ).
      ENDIF.
    ENDLOOP.

    DATA(lv_wb_dirty) = abap_false.
    DATA(lv_clone_no) = 0.

    LOOP AT elem_children( io_parent = lo_sheets iv_name = 'sheet' ) INTO DATA(lo_sheet).
      DATA(lv_name) = lo_sheet->get_attribute( 'name' ).
      DATA(lv_rid)  = lo_sheet->get_attribute_ns( name = 'id' uri = mc_rel ).

      DATA(lv_target) = ``.
      LOOP AT lt_rels INTO DATA(ls_rel) WHERE rid = lv_rid.
        lv_target = ls_rel-target.
        EXIT.
      ENDLOOP.
      IF lv_target IS INITIAL.
        CONTINUE.
      ENDIF.
      DATA(lv_path) = COND string( WHEN lv_target CS 'xl/' THEN lv_target
                                   ELSE |xl/{ lv_target }| ).

      DATA lv_loop_path TYPE string.
      CLEAR lv_loop_path.
      FIND PCRE '(?-x)\{\{#([^}>]+)\}\}' IN lv_name SUBMATCHES lv_loop_path.

      IF sy-subrc <> 0.
        IF lv_name CS '{{'.
          lo_sheet->set_attribute( name = 'name' value = mo_ctx->render_text( lv_name ) ).
          lv_wb_dirty = abap_true.
        ENDIF.
        process_sheet( lv_path ).
        CONTINUE.
      ENDIF.

      " ---- sheet loop -----------------------------------------------------
      DATA(lv_name_tpl) = lv_name.
      REPLACE FIRST OCCURRENCE OF PCRE '(?-x)\{\{#[^}>]+\}\}' IN lv_name_tpl WITH ''.

      FIELD-SYMBOLS <lt_sheetloop> TYPE INDEX TABLE.
      DATA(lr_tab) = mo_ctx->resolve_table( iv_path = lv_loop_path iv_from = 'sheetloop' ).
      ASSIGN lr_tab->* TO <lt_sheetloop>.
      DATA(lv_count) = lines( <lt_sheetloop> ).
      DATA(lv_tpl_raw) = zip_read( lv_path ).

      DO lv_count TIMES.
        DATA(lv_idx) = sy-index.
        READ TABLE <lt_sheetloop> INDEX lv_idx ASSIGNING FIELD-SYMBOL(<ls_line>).
        mo_ctx->push( ir_data = REF #( <ls_line> ) iv_index = lv_idx
                      iv_count = lv_count iv_kind = 'SHEET' ir_table = lr_tab ).

        lv_clone_no += 1.
        DATA(lv_new_path) = lv_path.
        REPLACE FIRST OCCURRENCE OF PCRE '(?i)\.xml$' IN lv_new_path
          WITH |_xw{ lv_clone_no }.xml|.

        process_sheet_raw( iv_raw = lv_tpl_raw iv_path = lv_new_path ).

        " khai sheet mới trong workbook + rels + [Content_Types]
        lv_max_id += 1.
        DATA(lv_new_rid) = |rIdXw{ lv_clone_no }|.
        DATA(lo_new_sheet) = lo_wb_doc->create_element_ns( name = 'sheet' uri = mc_main ).
        lo_new_sheet->set_attribute( name = 'name'
          value = mo_ctx->render_text( lv_name_tpl ) ).
        lo_new_sheet->set_attribute( name = 'sheetId' value = |{ lv_max_id }| ).
        " XCO khai xmlns:r INLINE trên từng element (root có thể không có) —
        " nên phải tự declare inline kèm attribute, nếu không file ra
        " 'undeclared namespace prefix'
        lo_new_sheet->set_attribute( name = 'xmlns:r' value = mc_rel ).
        lo_new_sheet->set_attribute( name = 'r:id'    value = lv_new_rid ).
        lo_sheets->insert_child( new_child = lo_new_sheet ref_child = lo_sheet ).

        DATA(lo_new_rel) = lo_rels_doc->create_element_ns( name = 'Relationship' uri = mc_pkg ).
        lo_new_rel->set_attribute( name = 'Id' value = lv_new_rid ).
        lo_new_rel->set_attribute( name = 'Type' value = |{ mc_rel }/worksheet| ).
        lo_new_rel->set_attribute( name = 'Target'
          value = replace( val = lv_new_path sub = `xl/` with = `` occ = 1 ) ).
        lo_rels_root->append_child( lo_new_rel ).

        ct_add_override(
          iv_partname = |/{ lv_new_path }|
          iv_type = `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` ).

        mo_ctx->pop( ).
      ENDDO.

      " bỏ sheet gốc khỏi danh sách (part gốc để nguyên trong zip — không được
      " tham chiếu thì Excel bỏ qua, xoá override/relationship rủi ro hơn)
      lo_sheets->remove_child( lo_sheet ).
      lv_wb_dirty = abap_true.
    ENDLOOP.

    IF lv_wb_dirty = abap_true OR lv_clone_no > 0.
      zip_write( iv_name = 'xl/workbook.xml' iv_raw = serialize( lo_wb_doc ) ).
      zip_write( iv_name = 'xl/_rels/workbook.xml.rels' iv_raw = serialize( lo_rels_doc ) ).
    ENDIF.
  ENDMETHOD.


  METHOD ct_add_override.
    DATA(lo_conv_in)  = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8' ).
    DATA(lo_conv_out) = cl_abap_conv_codepage=>create_out( codepage = 'UTF-8' ).
    DATA(lv_types) = lo_conv_in->convert( zip_read( '[Content_Types].xml' ) ).
    IF lv_types IS INITIAL OR lv_types CS iv_partname.
      RETURN.
    ENDIF.
    REPLACE FIRST OCCURRENCE OF '</Types>' IN lv_types WITH
      |<Override PartName="{ iv_partname }" ContentType="{ iv_type }"/></Types>|.
    zip_write( iv_name = '[Content_Types].xml' iv_raw = lo_conv_out->convert( lv_types ) ).
  ENDMETHOD.


  METHOD parse.
    ro_doc = mo_ixml->create_document( ).
    DATA(lo_parser) = mo_ixml->create_parser(
      document       = ro_doc
      istream        = mo_sf->create_istream_xstring( iv_xml )
      stream_factory = mo_sf ).
    lo_parser->set_namespace_mode( if_ixml_parser_core=>co_namespace_aware ).
    IF lo_parser->parse( ) <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = 'Cannot parse XML part of the .xlsx template' ).
    ENDIF.
  ENDMETHOD.


  METHOD serialize.
    DATA lv_out TYPE xstring.
    mo_ixml->create_renderer( document = io_doc
                              ostream  = mo_sf->create_ostream_xstring( lv_out ) )->render( ).
    rv_xml = lv_out.
  ENDMETHOD.


  METHOD zip_read.
    " get( ) raise ZIP_INDEX_ERROR khi part không có trong package
    mo_zip->get( EXPORTING name            = iv_name
                 IMPORTING content         = rv_raw
                 EXCEPTIONS zip_index_error = 1
                            OTHERS          = 2 ).
    IF sy-subrc <> 0.
      CLEAR rv_raw.
    ENDIF.
  ENDMETHOD.


  METHOD zip_write.
    mo_zip->delete( EXPORTING name            = iv_name
                    EXCEPTIONS zip_index_error = 1
                               OTHERS          = 2 ).
    mo_zip->add( name = iv_name content = iv_raw ).
  ENDMETHOD.


  METHOD load_shared_strings.
    CLEAR mt_shared.
    DATA(lv_raw) = zip_read( 'xl/sharedStrings.xml' ).
    IF lv_raw IS INITIAL.
      RETURN.
    ENDIF.
    TRY.
        DATA(lo_doc) = parse( lv_raw ).
      CATCH zcx_xlwb.
        RETURN.
    ENDTRY.

    " mỗi <si> có thể gồm nhiều <t> (rich text) → nối lại
    DATA(lo_ss_root) = lo_doc->get_root_element( ).
    IF lo_ss_root IS NOT BOUND.
      RETURN.
    ENDIF.
    DATA(lo_si_col) = lo_ss_root->get_elements_by_tag_name_ns( name = 'si' uri = mc_main ).
    DO lo_si_col->get_length( ) TIMES.
      DATA(lo_si) = CAST if_ixml_element( lo_si_col->get_item( sy-index - 1 ) ).
      DATA(lo_t_col) = lo_si->get_elements_by_tag_name_ns( name = 't' uri = mc_main ).
      DATA lv_text TYPE string.
      CLEAR lv_text.
      DO lo_t_col->get_length( ) TIMES.
        lv_text = lv_text && lo_t_col->get_item( sy-index - 1 )->get_value( ).
      ENDDO.
      APPEND lv_text TO mt_shared.
    ENDDO.
  ENDMETHOD.


  METHOD sheet_paths.
    " workbook.xml.rels: rId → target; workbook.xml: sheet order
    DATA(lv_rels_raw) = zip_read( 'xl/_rels/workbook.xml.rels' ).
    DATA(lv_wb_raw)   = zip_read( 'xl/workbook.xml' ).
    IF lv_rels_raw IS INITIAL OR lv_wb_raw IS INITIAL.
      APPEND 'xl/worksheets/sheet1.xml' TO rt_paths.
      RETURN.
    ENDIF.

    DATA(lo_rels_root) = parse( lv_rels_raw )->get_root_element( ).
    DATA(lo_wb_root)   = parse( lv_wb_raw )->get_root_element( ).
    IF lo_rels_root IS NOT BOUND OR lo_wb_root IS NOT BOUND.
      APPEND 'xl/worksheets/sheet1.xml' TO rt_paths.
      RETURN.
    ENDIF.

    DATA(lo_rel_col)   = lo_rels_root->get_elements_by_tag_name_ns(
                           name = 'Relationship' uri = mc_pkg ).
    DATA(lo_sheet_col) = lo_wb_root->get_elements_by_tag_name_ns(
                           name = 'sheet' uri = mc_main ).

    DO lo_sheet_col->get_length( ) TIMES.
      DATA(lo_sheet) = CAST if_ixml_element( lo_sheet_col->get_item( sy-index - 1 ) ).
      DATA(lv_rid) = lo_sheet->get_attribute_ns( name = 'id' uri = mc_rel ).

      DO lo_rel_col->get_length( ) TIMES.
        DATA(lo_rel) = CAST if_ixml_element( lo_rel_col->get_item( sy-index - 1 ) ).
        IF lo_rel->get_attribute( 'Id' ) = lv_rid.
          DATA(lv_target) = lo_rel->get_attribute( 'Target' ).
          IF lv_target NS 'xl/'.
            lv_target = |xl/{ lv_target }|.
          ENDIF.
          APPEND lv_target TO rt_paths.
          EXIT.
        ENDIF.
      ENDDO.
    ENDDO.

    IF rt_paths IS INITIAL.
      APPEND 'xl/worksheets/sheet1.xml' TO rt_paths.
    ENDIF.
  ENDMETHOD.


  METHOD process_sheet.
    process_sheet_raw( iv_raw = zip_read( iv_path ) iv_path = iv_path ).
  ENDMETHOD.


  METHOD process_sheet_raw.
    DATA(lv_raw) = iv_raw.
    IF lv_raw IS INITIAL.
      RETURN.
    ENDIF.

    mo_doc = parse( lv_raw ).
    DATA(lo_ws) = mo_doc->get_root_element( ).
    IF lo_ws IS NOT BOUND.
      RETURN.
    ENDIF.
    CLEAR: mt_merges, mt_dynmerges.

    collect_merges( lo_ws ).

    DATA(lo_data) = lo_ws->find_from_name_ns( name = 'sheetData' uri = mc_main ).
    IF lo_data IS BOUND.
      process_rows( elem_children( io_parent = lo_data iv_name = 'row' ) ).
    ENDIF.

    finalize_sheet( io_ws = lo_ws iv_path = iv_path ).
    zip_write( iv_name = iv_path iv_raw = serialize( mo_doc ) ).
  ENDMETHOD.


  METHOD collect_merges.
    " Đọc mergeCells của template rồi bỏ element đi — finalize_sheet sinh lại
    " theo số dòng thật sau khi nhân bản.
    DATA(lo_mc) = io_ws->find_from_name_ns( name = 'mergeCells' uri = mc_main ).
    IF lo_mc IS NOT BOUND.
      RETURN.
    ENDIF.

    LOOP AT elem_children( io_parent = lo_mc iv_name = 'mergeCell' ) INTO DATA(lo_m).
      DATA(lv_ref) = lo_m->get_attribute_ns( name = 'ref' uri = '' ).
      SPLIT lv_ref AT ':' INTO DATA(lv_a) DATA(lv_b).
      IF lv_b IS INITIAL.
        CONTINUE.
      ENDIF.
      split_ref( EXPORTING iv_ref = lv_a
                 IMPORTING ev_col = DATA(lv_col_a) ev_row = DATA(lv_row_a) ).
      split_ref( EXPORTING iv_ref = lv_b
                 IMPORTING ev_col = DATA(lv_col_b) ev_row = DATA(lv_row_b) ).
      IF lv_row_a = 0 OR lv_row_b = 0.
        CONTINUE.
      ENDIF.
      APPEND VALUE ty_merge( col_from = lv_col_a col_to = lv_col_b
                             row_from = lv_row_a row_to = lv_row_b ) TO mt_merges.
    ENDLOOP.

    io_ws->remove_child( lo_mc ).
  ENDMETHOD.


  METHOD process_rows.
    DATA(lv_i) = 1.
    WHILE lv_i <= lines( it_rows ).
      DATA(lo_row) = it_rows[ lv_i ].

      scan_strip(
        EXPORTING io_scope = lo_row
                  iv_pcre  = '\{\{([#?^])([^}>:]+)\}\}'
        IMPORTING ev_found = DATA(lv_found)
                  ev_sub1  = DATA(lv_kind)
                  ev_sub2  = DATA(lv_path) ).

      IF lv_found = abap_false.
        render_row( lo_row ).
        lv_i += 1.
        CONTINUE.
      ENDIF.

      DATA(lv_path_pcre) = replace( val = lv_path sub = '.' with = '\.' occ = 0 ).
      DATA(lv_j) = 0.
      DATA(lv_k) = lv_i.
      WHILE lv_k <= lines( it_rows ).
        scan_strip(
          EXPORTING io_scope = it_rows[ lv_k ]
                    iv_pcre  = |\\\{\\\{(/)({ lv_path_pcre })\\\}\\\}|
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
          DATA(lr_tab) = mo_ctx->resolve_table( iv_path = lv_path iv_from = 'rowloop' ).
          ASSIGN lr_tab->* TO <lt_loop>.
          DATA(lv_count) = lines( <lt_loop> ).

          DO lv_count TIMES.
            DATA(lv_idx) = sy-index.
            READ TABLE <lt_loop> INDEX lv_idx ASSIGNING FIELD-SYMBOL(<ls_line>).
            mo_ctx->push( ir_data  = REF #( <ls_line> ) iv_index = lv_idx
                          iv_count = lv_count iv_kind = 'ROW'
                          ir_table = lr_tab ).

            DATA(lt_clones) = VALUE ty_elems( ).
            LOOP AT lt_block INTO DATA(lo_tpl).
              DATA(lo_clone) = CAST if_ixml_element( lo_tpl->clone( ) ).
              lo_parent->insert_child( new_child = lo_clone ref_child = lt_block[ 1 ] ).
              APPEND lo_clone TO lt_clones.
            ENDLOOP.
            process_rows( lt_clones ).

            mo_ctx->pop( ).
          ENDDO.

          LOOP AT lt_block INTO DATA(lo_del).
            lo_parent->remove_child( lo_del ).
          ENDLOOP.

        WHEN '?' OR '^'.
          DATA(lv_truthy) = mo_ctx->is_truthy( lv_path ).
          IF ( lv_kind = '?' AND lv_truthy = abap_true )
          OR ( lv_kind = '^' AND lv_truthy = abap_false ).
            process_rows( lt_block ).
          ELSE.
            LOOP AT lt_block INTO DATA(lo_drop).
              lo_parent->remove_child( lo_drop ).
            ENDLOOP.
          ENDIF.
      ENDCASE.

      lv_i = lv_j + 1.
    ENDWHILE.
  ENDMETHOD.


  METHOD render_row.
    " {{*break}}: đánh dấu dòng bằng attribute tạm — finalize_sheet chuyển
    " thành <rowBreaks> sau khi đã đánh lại số dòng
    scan_strip(
      EXPORTING io_scope = io_row
                iv_pcre  = '\{\{(\*)(break)\}\}'
      IMPORTING ev_found = DATA(lv_break) ).
    IF lv_break = abap_true.
      io_row->set_attribute( name = 'xlwbbreak' value = 'X' ).
    ENDIF.

    LOOP AT cells_of( io_row ) INTO DATA(lo_cell).
      " cell loop ngang {{#path>}}: nhân ô sang phải mỗi dòng của bảng
      scan_strip(
        EXPORTING io_scope = lo_cell
                  iv_pcre  = '\{\{#([^}>]+)>()\}\}'
        IMPORTING ev_found = DATA(lv_cl)
                  ev_sub1  = DATA(lv_cl_path) ).
      IF lv_cl = abap_true.
        expand_cell_loop( io_row = io_row io_cell = lo_cell iv_path = lv_cl_path ).
        CONTINUE.
      ENDIF.
      render_cell( lo_cell ).
    ENDLOOP.
  ENDMETHOD.


  METHOD expand_cell_loop.
    FIELD-SYMBOLS <lt_loop> TYPE INDEX TABLE.
    DATA(lr_tab) = mo_ctx->resolve_table( iv_path = iv_path iv_from = 'cellloop' ).
    ASSIGN lr_tab->* TO <lt_loop>.
    DATA(lv_count) = lines( <lt_loop> ).

    split_ref( EXPORTING iv_ref = io_cell->get_attribute_ns( name = 'r' uri = '' )
               IMPORTING ev_col = DATA(lv_col) ev_row = DATA(lv_row) ).
    IF lv_col IS INITIAL.
      lv_col = `A`.
      lv_row = 1.
    ENDIF.
    DATA(lv_colnum) = col_to_num( lv_col ).

    " dịch các ô BÊN PHẢI ô mẫu sang phải (count-1) cột để nhường chỗ bản sao
    IF lv_count > 1.
      LOOP AT cells_of( io_row ) INTO DATA(lo_sib).
        split_ref( EXPORTING iv_ref = lo_sib->get_attribute_ns( name = 'r' uri = '' )
                   IMPORTING ev_col = DATA(lv_scol) ev_row = DATA(lv_srow) ).
        IF lv_scol IS NOT INITIAL AND col_to_num( lv_scol ) > lv_colnum.
          lo_sib->set_attribute_ns( name = 'r' uri = ''
            value = |{ num_to_col( col_to_num( lv_scol ) + lv_count - 1 ) }{ lv_srow }| ).
        ENDIF.
      ENDLOOP.
    ENDIF.

    IF lv_count = 0.
      io_row->remove_child( io_cell ).
      RETURN.
    ENDIF.

    " tạo đủ bản sao TỪ Ô MẪU CHƯA RENDER trước — clone sau khi render sẽ
    " mang sẵn giá trị của dòng 1
    DATA lt_targets TYPE ty_elems.
    APPEND io_cell TO lt_targets.
    DATA(lo_ref_next) = io_cell->get_next( ).
    DATA(lv_i) = 1.
    WHILE lv_i < lv_count.
      lv_i += 1.
      DATA(lo_clone) = CAST if_ixml_element( io_cell->clone( ) ).
      IF lo_ref_next IS BOUND.
        io_row->insert_child( new_child = lo_clone ref_child = lo_ref_next ).
      ELSE.
        io_row->append_child( lo_clone ).
      ENDIF.
      lo_clone->set_attribute_ns( name = 'r' uri = ''
        value = |{ num_to_col( lv_colnum + lv_i - 1 ) }{ lv_row }| ).
      APPEND lo_clone TO lt_targets.
    ENDWHILE.

    LOOP AT lt_targets INTO DATA(lo_target).
      DATA(lv_idx) = sy-tabix.
      READ TABLE <lt_loop> INDEX lv_idx ASSIGNING FIELD-SYMBOL(<ls_line>).
      mo_ctx->push( ir_data = REF #( <ls_line> ) iv_index = lv_idx
                    iv_count = lv_count iv_kind = 'CELL' ir_table = lr_tab ).
      render_cell( lo_target ).
      mo_ctx->pop( ).
    ENDLOOP.
  ENDMETHOD.


  METHOD render_cell.
    " ---- ảnh: {{*image:path}} hoặc {{*image:path@CxR}} -------------------
    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*image:([^}@]+)(?:@(\d+)x(\d+))?\}\}'
      IMPORTING ev_found = DATA(lv_img)
                ev_sub1  = DATA(lv_img_path)
                ev_sub2  = DATA(lv_img_cols) ).
    IF lv_img = abap_true.
      DATA(ls_b64) = mo_ctx->token_value( lv_img_path ).
      IF ls_b64-text IS NOT INITIAL.
        APPEND VALUE ty_image(
          cell = io_cell
          data = xco_cp=>string( ls_b64-text
                   )->as_xstring( xco_cp_binary=>text_encoding->base64 )->value
          cols = COND i( WHEN lv_img_cols CO '0123456789' AND lv_img_cols IS NOT INITIAL
                         THEN CONV i( lv_img_cols ) ELSE 2 )
          rows = 4 ) TO mt_images.
      ENDIF.
    ENDIF.

    " ---- merge động ------------------------------------------------------
    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergedown:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_md)
                ev_sub1  = DATA(lv_md_path) ).
    IF lv_md = abap_true.
      FIELD-SYMBOLS <lt_md> TYPE INDEX TABLE.
      DATA(lr_md) = mo_ctx->resolve_table( iv_path = lv_md_path iv_from = 'mergedown' ).
      ASSIGN lr_md->* TO <lt_md>.
      IF lines( <lt_md> ) > 1.
        APPEND VALUE ty_dynmerge( cell = io_cell rows = lines( <lt_md> ) cols = 1 )
          TO mt_dynmerges.
      ENDIF.
    ENDIF.

    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergeacross:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_ma)
                ev_sub1  = DATA(lv_ma_path) ).
    IF lv_ma = abap_true.
      FIELD-SYMBOLS <lt_ma> TYPE INDEX TABLE.
      DATA(lr_ma) = mo_ctx->resolve_table( iv_path = lv_ma_path iv_from = 'mergeacross' ).
      ASSIGN lr_ma->* TO <lt_ma>.
      IF lines( <lt_ma> ) > 1.
        APPEND VALUE ty_dynmerge( cell = io_cell rows = 1 cols = lines( <lt_ma> ) )
          TO mt_dynmerges.
      ENDIF.
    ENDIF.

    " span so hoc: {{*mergedown=N}} / {{*mergeacross=N}} - N = so o duoc gop
    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergedown=([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_mdn)
                ev_sub1  = DATA(lv_mdn_spec) ).
    IF lv_mdn = abap_true.
      DATA(lv_mdn_num) = mo_ctx->span_of( lv_mdn_spec ).
      IF lv_mdn_num > 1.
        APPEND VALUE ty_dynmerge( cell = io_cell rows = lv_mdn_num cols = 1 )
          TO mt_dynmerges.
      ENDIF.
    ENDIF.

    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergeacross=([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_man)
                ev_sub1  = DATA(lv_man_spec) ).
    IF lv_man = abap_true.
      DATA(lv_man_num) = mo_ctx->span_of( lv_man_spec ).
      IF lv_man_num > 1.
        APPEND VALUE ty_dynmerge( cell = io_cell rows = 1 cols = lv_man_num )
          TO mt_dynmerges.
      ENDIF.
    ENDIF.

    " gop NGANG theo gia tri: {{*mergesame>:path}} trong vong lap o
    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergesame>:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_msh)
                ev_sub1  = DATA(lv_msh_path) ).
    IF lv_msh = abap_true.
      mo_ctx->group_span( EXPORTING iv_path     = lv_msh_path
                                    iv_kind     = 'CELL'
                          IMPORTING ev_is_first = DATA(lv_msh_first)
                                    ev_span     = DATA(lv_msh_span) ).
      IF lv_msh_first = abap_true.
        IF lv_msh_span > 1.
          APPEND VALUE ty_dynmerge( cell = io_cell rows = 1 cols = lv_msh_span )
            TO mt_dynmerges.
        ENDIF.
      ELSE.
        set_cell_text( io_cell = io_cell iv_text = `` ).
        RETURN.
      ENDIF.
    ENDIF.

    " gop o theo gia tri: {{*mergesame:path}} trong vong lap dong.
    " O bi phu duoc lam trong -> khong co hai o cung khai bao du lieu.
    scan_strip(
      EXPORTING io_scope = io_cell
                iv_pcre  = '\{\{\*mergesame:([^}]+)()\}\}'
      IMPORTING ev_found = DATA(lv_ms)
                ev_sub1  = DATA(lv_ms_path) ).
    IF lv_ms = abap_true.
      mo_ctx->group_span( EXPORTING iv_path     = lv_ms_path
                                    iv_kind     = 'ROW'
                          IMPORTING ev_is_first = DATA(lv_ms_first)
                                    ev_span     = DATA(lv_ms_span) ).
      IF lv_ms_first = abap_true.
        IF lv_ms_span > 1.
          APPEND VALUE ty_dynmerge( cell = io_cell rows = lv_ms_span cols = 1 )
            TO mt_dynmerges.
        ENDIF.
      ELSE.
        set_cell_text( io_cell = io_cell iv_text = `` ).
        RETURN.
      ENDIF.
    ENDIF.

    " ---- giá trị ---------------------------------------------------------
    DATA(lv_text) = cell_text( io_cell ).
    IF lv_text NS '{{'.
      RETURN.
    ENDIF.

    DATA lv_token TYPE string.
    FIND PCRE '(?-x)^\s*\{\{([^}]+)\}\}\s*$' IN lv_text SUBMATCHES lv_token.
    IF sy-subrc = 0 AND substring( val = lv_token off = 0 len = 1 ) NA '#/?^*'.
      set_cell_value( io_cell = io_cell is_value = mo_ctx->token_value( lv_token ) ).
    ELSE.
      set_cell_text( io_cell = io_cell iv_text = mo_ctx->render_text( lv_text ) ).
    ENDIF.
  ENDMETHOD.


  METHOD cell_text.
    DATA(lv_type) = io_cell->get_attribute_ns( name = 't' uri = '' ).

    IF lv_type = 's'.
      DATA(lo_v) = io_cell->find_from_name_ns( name = 'v' uri = mc_main ).
      IF lo_v IS BOUND.
        DATA(lv_ix) = CONV i( lo_v->get_value( ) ) + 1.
        IF lv_ix >= 1 AND lv_ix <= lines( mt_shared ).
          rv_text = mt_shared[ lv_ix ].
        ENDIF.
      ENDIF.
      RETURN.
    ENDIF.

    " inlineStr / str / không type: nối mọi <t> rồi tới <v>
    DATA(lo_t_col) = io_cell->get_elements_by_tag_name_ns( name = 't' uri = mc_main ).
    IF lo_t_col->get_length( ) > 0.
      DO lo_t_col->get_length( ) TIMES.
        rv_text = rv_text && lo_t_col->get_item( sy-index - 1 )->get_value( ).
      ENDDO.
      RETURN.
    ENDIF.

    DATA(lo_val) = io_cell->find_from_name_ns( name = 'v' uri = mc_main ).
    IF lo_val IS BOUND.
      rv_text = lo_val->get_value( ).
    ENDIF.
  ENDMETHOD.


  METHOD set_cell_text.
    " xoá con cũ rồi ghi lại dạng inlineStr (không cần sửa sharedStrings)
    DATA(lo_child) = io_cell->get_first_child( ).
    WHILE lo_child IS BOUND.
      DATA(lo_next) = lo_child->get_next( ).
      io_cell->remove_child( lo_child ).
      lo_child = lo_next.
    ENDWHILE.

    " Ghi qua sharedStrings (t="s"): tương thích rộng nhất — XCO read access
    " và một số tool khác KHÔNG đọc được inlineStr.
    io_cell->set_attribute_ns( name = 't' uri = '' value = 's' ).
    mo_doc->create_simple_element_ns(
      name = 'v' parent = io_cell uri = mc_main value = |{ shared_index( iv_text ) }| ).
  ENDMETHOD.


  METHOD shared_index.
    LOOP AT mt_shared INTO DATA(lv_s).
      IF lv_s = iv_text.
        rv_idx = sy-tabix - 1.       " sst dùng chỉ số từ 0
        RETURN.
      ENDIF.
    ENDLOOP.
    APPEND iv_text TO mt_shared.
    mv_ss_dirty = abap_true.
    rv_idx = lines( mt_shared ) - 1.
  ENDMETHOD.


  METHOD escape_xml.
    rv_text = iv_text.
    REPLACE ALL OCCURRENCES OF '&' IN rv_text WITH '&amp;'.
    REPLACE ALL OCCURRENCES OF '<' IN rv_text WITH '&lt;'.
    REPLACE ALL OCCURRENCES OF '>' IN rv_text WITH '&gt;'.
  ENDMETHOD.


  METHOD save_shared_strings.
    IF mv_ss_dirty = abap_false.
      RETURN.
    ENDIF.

    DATA(lv_xml) = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
                   |<sst xmlns="{ mc_main }" count="{ lines( mt_shared ) }"| &&
                   | uniqueCount="{ lines( mt_shared ) }">|.
    LOOP AT mt_shared INTO DATA(lv_s).
      lv_xml = lv_xml && |<si><t xml:space="preserve">{ escape_xml( lv_s ) }</t></si>|.
    ENDLOOP.
    lv_xml = lv_xml && |</sst>|.

    DATA(lo_conv) = cl_abap_conv_codepage=>create_out( codepage = 'UTF-8' ).
    zip_write( iv_name = 'xl/sharedStrings.xml' iv_raw = lo_conv->convert( lv_xml ) ).

    " template chưa từng có sharedStrings → phải khai part + quan hệ
    DATA(lv_types) = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8'
                       )->convert( zip_read( '[Content_Types].xml' ) ).
    IF lv_types IS NOT INITIAL AND lv_types NS 'sharedStrings.xml'.
      REPLACE FIRST OCCURRENCE OF '</Types>' IN lv_types WITH
        |<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxml| &&
        |formats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>|.
      zip_write( iv_name = '[Content_Types].xml' iv_raw = lo_conv->convert( lv_types ) ).
    ENDIF.

    DATA(lv_rels) = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8'
                      )->convert( zip_read( 'xl/_rels/workbook.xml.rels' ) ).
    IF lv_rels IS NOT INITIAL AND lv_rels NS 'sharedStrings.xml'.
      REPLACE FIRST OCCURRENCE OF '</Relationships>' IN lv_rels WITH
        |<Relationship Id="rIdXlwbSst" Type="{ mc_rel }/sharedStrings"| &&
        | Target="sharedStrings.xml"/></Relationships>|.
      zip_write( iv_name = 'xl/_rels/workbook.xml.rels' iv_raw = lo_conv->convert( lv_rels ) ).
    ENDIF.
  ENDMETHOD.


  METHOD set_cell_value.
    IF is_value-is_number = abap_false
    AND is_value-is_date  = abap_false
    AND is_value-is_time  = abap_false.
      set_cell_text( io_cell = io_cell iv_text = is_value-text ).
      RETURN.
    ENDIF.

    DATA lv_num TYPE string.
    IF is_value-is_number = abap_true.
      lv_num = is_value-text.
    ELSEIF is_value-is_date = abap_true.
      IF is_value-date_value IS INITIAL.
        set_cell_text( io_cell = io_cell iv_text = `` ).
        RETURN.
      ENDIF.
      " serial number Excel: 1900-date-system, gốc 1899-12-30
      lv_num = |{ is_value-date_value - CONV d( '18991230' ) }|.
    ELSE.
      DATA(lv_secs) = is_value-time_value - CONV t( '000000' ).
      lv_num = |{ CONV decfloat34( lv_secs / 86400 ) NUMBER = RAW }|.
    ENDIF.

    DATA(lo_child) = io_cell->get_first_child( ).
    WHILE lo_child IS BOUND.
      DATA(lo_next) = lo_child->get_next( ).
      io_cell->remove_child( lo_child ).
      lo_child = lo_next.
    ENDWHILE.

    " ô số: bỏ attribute t (mặc định "n"), style của template giữ nguyên
    io_cell->remove_attribute_ns( name = 't' uri = '' ).
    mo_doc->create_simple_element_ns(
      name = 'v' parent = io_cell uri = mc_main value = lv_num ).
  ENDMETHOD.


  METHOD scan_strip.
    CLEAR: ev_found, ev_sub1, ev_sub2.
    DATA(lv_pcre) = `(?-x)` && iv_pcre.

    LOOP AT cells_of( io_scope ) INTO DATA(lo_cell).
      DATA(lv_text) = cell_text( lo_cell ).
      IF lv_text NS '{{'.
        CONTINUE.
      ENDIF.
      FIND PCRE lv_pcre IN lv_text SUBMATCHES ev_sub1 ev_sub2.
      IF sy-subrc = 0.
        REPLACE FIRST OCCURRENCE OF PCRE lv_pcre IN lv_text WITH ''.
        set_cell_text( io_cell = lo_cell iv_text = lv_text ).
        ev_found = abap_true.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD finalize_sheet.
    " ---- 1. đánh lại số dòng và ref ô, lấy map dòng gốc → dòng mới -------
    DATA lt_map TYPE ty_rowmaps.
    DATA lt_breaks TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    DATA(lo_data) = io_ws->find_from_name_ns( name = 'sheetData' uri = mc_main ).
    IF lo_data IS BOUND.
      DATA(lv_new) = 0.
      LOOP AT elem_children( io_parent = lo_data iv_name = 'row' ) INTO DATA(lo_row).
        lv_new += 1.
        DATA(lv_origin) = CONV i( lo_row->get_attribute_ns( name = 'r' uri = '' ) ).
        APPEND VALUE ty_rowmap( origin = lv_origin new = lv_new ) TO lt_map.
        lo_row->set_attribute_ns( name = 'r' uri = '' value = |{ lv_new }| ).

        IF lo_row->get_attribute( 'xlwbbreak' ) = 'X'.
          lo_row->remove_attribute( 'xlwbbreak' ).
          IF lv_new > 1. " break trước dòng đầu vô nghĩa
            APPEND lv_new - 1 TO lt_breaks.
          ENDIF.
        ENDIF.

        LOOP AT cells_of( lo_row ) INTO DATA(lo_c).
          split_ref( EXPORTING iv_ref = lo_c->get_attribute_ns( name = 'r' uri = '' )
                     IMPORTING ev_col = DATA(lv_col) ).
          IF lv_col IS NOT INITIAL.
            lo_c->set_attribute_ns( name = 'r' uri = '' value = |{ lv_col }{ lv_new }| ).
          ENDIF.
        ENDLOOP.
      ENDLOOP.
    ENDIF.

    " ---- 2. dựng lại mergeCells -----------------------------------------
    DATA lt_refs TYPE string_table.

    LOOP AT mt_merges INTO DATA(ls_m).
      IF ls_m-row_from = ls_m-row_to.
        " merge trong 1 dòng: nhân theo mọi bản sao của dòng đó
        LOOP AT lt_map INTO DATA(ls_map) WHERE origin = ls_m-row_from.
          APPEND |{ ls_m-col_from }{ ls_map-new }:{ ls_m-col_to }{ ls_map-new }| TO lt_refs.
        ENDLOOP.
      ELSE.
        " merge nhiều dòng: chỉ giữ khi cả hai đầu không bị nhân bản
        DATA(lv_cnt_from) = REDUCE i( INIT x = 0 FOR m IN lt_map
                                      NEXT x = COND i( WHEN m-origin = ls_m-row_from THEN x + 1 ELSE x ) ).
        DATA(lv_cnt_to)   = REDUCE i( INIT x = 0 FOR m IN lt_map
                                      NEXT x = COND i( WHEN m-origin = ls_m-row_to THEN x + 1 ELSE x ) ).
        IF lv_cnt_from = 1 AND lv_cnt_to = 1.
          APPEND |{ ls_m-col_from }{ lt_map[ origin = ls_m-row_from ]-new }:| &&
                 |{ ls_m-col_to }{ lt_map[ origin = ls_m-row_to ]-new }| TO lt_refs.
        ENDIF.
      ENDIF.
    ENDLOOP.

    LOOP AT mt_dynmerges INTO DATA(ls_dyn).
      split_ref( EXPORTING iv_ref = ls_dyn-cell->get_attribute_ns( name = 'r' uri = '' )
                 IMPORTING ev_col = DATA(lv_dcol) ev_row = DATA(lv_drow) ).
      IF lv_drow = 0.
        CONTINUE.
      ENDIF.
      APPEND |{ lv_dcol }{ lv_drow }:| &&
             |{ num_to_col( col_to_num( lv_dcol ) + ls_dyn-cols - 1 ) }{ lv_drow + ls_dyn-rows - 1 }|
        TO lt_refs.
    ENDLOOP.

    IF lt_refs IS NOT INITIAL AND lo_data IS BOUND.
      DATA(lo_mc) = mo_doc->create_element_ns( name = 'mergeCells' uri = mc_main ).
      lo_mc->set_attribute_ns( name = 'count' uri = '' value = |{ lines( lt_refs ) }| ).
      LOOP AT lt_refs INTO DATA(lv_ref).
        mo_doc->create_simple_element_ns( name = 'mergeCell' parent = lo_mc uri = mc_main
          )->set_attribute_ns( name = 'ref' uri = '' value = lv_ref ).
      ENDLOOP.
      " mergeCells phải nằm ngay sau sheetData (schema OOXML có thứ tự cố định)
      DATA(lo_after) = lo_data->get_next( ).
      IF lo_after IS BOUND.
        io_ws->insert_child( new_child = lo_mc ref_child = lo_after ).
      ELSE.
        io_ws->append_child( lo_mc ).
      ENDIF.
    ENDIF.

    " ---- 3. dimension cũ sai sau khi nhân dòng → bỏ, Excel tự tính ------
    DATA(lo_dim) = io_ws->find_from_name_ns( name = 'dimension' uri = mc_main ).
    IF lo_dim IS BOUND.
      io_ws->remove_child( lo_dim ).
    ENDIF.

    " ---- 3b. page breaks từ {{*break}} ------------------------------------
    IF lt_breaks IS NOT INITIAL.
      DATA(lo_rbks) = mo_doc->create_element_ns( name = 'rowBreaks' uri = mc_main ).
      lo_rbks->set_attribute_ns( name = 'count' uri = '' value = |{ lines( lt_breaks ) }| ).
      lo_rbks->set_attribute_ns( name = 'manualBreakCount' uri = ''
                                 value = |{ lines( lt_breaks ) }| ).
      LOOP AT lt_breaks INTO DATA(lv_brk).
        DATA(lo_brk) = mo_doc->create_simple_element_ns( name = 'brk' parent = lo_rbks
                                                         uri = mc_main ).
        lo_brk->set_attribute_ns( name = 'id'  uri = '' value = |{ lv_brk }| ).
        lo_brk->set_attribute_ns( name = 'max' uri = '' value = '16383' ).
        lo_brk->set_attribute_ns( name = 'man' uri = '' value = '1' ).
      ENDLOOP.
      " schema OOXML: rowBreaks đứng TRƯỚC <drawing> nếu có
      DATA(lo_draw_ref) = io_ws->find_from_name_ns( name = 'drawing' uri = mc_main ).
      IF lo_draw_ref IS BOUND.
        io_ws->insert_child( new_child = lo_rbks ref_child = lo_draw_ref ).
      ELSE.
        io_ws->append_child( lo_rbks ).
      ENDIF.
    ENDIF.

    " ---- 4. ảnh ----------------------------------------------------------
    IF mt_images IS NOT INITIAL.
      attach_images( io_ws = io_ws iv_path = iv_path ).
      CLEAR mt_images.
    ENDIF.
  ENDMETHOD.


  METHOD attach_images.
    DATA(lv_base) = iv_path.
    REPLACE FIRST OCCURRENCE OF 'xl/worksheets/' IN lv_base WITH ''.
    DATA(lv_rels_path) = |xl/worksheets/_rels/{ lv_base }.rels|.

    mv_img_count += 1.
    DATA(lv_dw_idx)  = mv_img_count.
    DATA(lv_dw_path) = |xl/drawings/drawing{ lv_dw_idx }.xml|.

    " ---- drawing + media -------------------------------------------------
    DATA(lv_dw) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"| &&
      | xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"| &&
      | xmlns:r="{ mc_rel }">|.
    DATA(lv_dw_rels) =
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<Relationships xmlns="{ mc_pkg }">|.

    DATA(lv_n) = 0.
    LOOP AT mt_images INTO DATA(ls_img).
      lv_n += 1.
      split_ref( EXPORTING iv_ref = ls_img-cell->get_attribute_ns( name = 'r' uri = '' )
                 IMPORTING ev_col = DATA(lv_col) ev_row = DATA(lv_row) ).
      IF lv_row = 0.
        CONTINUE.
      ENDIF.
      DATA(lv_c0) = col_to_num( lv_col ) - 1.   " drawing dùng chỉ số từ 0
      DATA(lv_r0) = lv_row - 1.

      DATA(lv_media) = |xl/media/xlwb_image{ lv_dw_idx }_{ lv_n }.png|.
      zip_write( iv_name = lv_media iv_raw = ls_img-data ).

      lv_dw = lv_dw &&
        |<xdr:twoCellAnchor editAs="oneCell">| &&
        |<xdr:from><xdr:col>{ lv_c0 }</xdr:col><xdr:colOff>0</xdr:colOff>| &&
        |<xdr:row>{ lv_r0 }</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>| &&
        |<xdr:to><xdr:col>{ lv_c0 + ls_img-cols }</xdr:col><xdr:colOff>0</xdr:colOff>| &&
        |<xdr:row>{ lv_r0 + ls_img-rows }</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>| &&
        |<xdr:pic><xdr:nvPicPr>| &&
        |<xdr:cNvPr id="{ lv_n }" name="Image{ lv_n }"/><xdr:cNvPicPr/></xdr:nvPicPr>| &&
        |<xdr:blipFill><a:blip r:embed="rId{ lv_n }"/>| &&
        |<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>| &&
        |<xdr:spPr><a:xfrm><a:off x="0" y="0"/>| &&
        |<a:ext cx="{ ls_img-cols * 640080 }" cy="{ ls_img-rows * 190500 }"/></a:xfrm>| &&
        |<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>| &&
        |</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>|.

      DATA(lv_target) = lv_media.
      REPLACE FIRST OCCURRENCE OF 'xl/' IN lv_target WITH '../'.
      lv_dw_rels = lv_dw_rels &&
        |<Relationship Id="rId{ lv_n }"| &&
        | Type="{ mc_rel }/image" Target="{ lv_target }"/>|.
    ENDLOOP.

    lv_dw      = lv_dw && |</xdr:wsDr>|.
    lv_dw_rels = lv_dw_rels && |</Relationships>|.

    DATA(lo_conv) = cl_abap_conv_codepage=>create_out( codepage = 'UTF-8' ).
    zip_write( iv_name = lv_dw_path iv_raw = lo_conv->convert( lv_dw ) ).
    zip_write( iv_name = |xl/drawings/_rels/drawing{ lv_dw_idx }.xml.rels|
               iv_raw  = lo_conv->convert( lv_dw_rels ) ).

    " ---- rels của worksheet: thêm quan hệ tới drawing ---------------------
    DATA(lv_rid) = |rIdXlwbDrawing{ lv_dw_idx }|.
    DATA(lv_existing) = zip_read( lv_rels_path ).
    DATA lv_rels TYPE string.
    IF lv_existing IS INITIAL.
      lv_rels = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
                |<Relationships xmlns="{ mc_pkg }">| &&
                |<Relationship Id="{ lv_rid }" Type="{ mc_rel }/drawing"| &&
                | Target="../drawings/drawing{ lv_dw_idx }.xml"/>| &&
                |</Relationships>|.
    ELSE.
      lv_rels = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8' )->convert( lv_existing ).
      REPLACE FIRST OCCURRENCE OF '</Relationships>' IN lv_rels WITH
        |<Relationship Id="{ lv_rid }" Type="{ mc_rel }/drawing"| &&
        | Target="../drawings/drawing{ lv_dw_idx }.xml"/></Relationships>|.
    ENDIF.
    zip_write( iv_name = lv_rels_path iv_raw = lo_conv->convert( lv_rels ) ).

    " Template do XCO sinh có <pageSetup r:id="rIdN"/> nhưng KHÔNG có file rels
    " của worksheet -> relationship treo. Bỏ attribute đó nếu rels không khai.
    DATA(lo_ps) = io_ws->find_from_name_ns( name = 'pageSetup' uri = mc_main ).
    IF lo_ps IS BOUND.
      DATA(lv_ps_rid) = lo_ps->get_attribute_ns( name = 'id' uri = mc_rel ).
      IF lv_ps_rid IS NOT INITIAL AND lv_rels NS |Id="{ lv_ps_rid }"|.
        lo_ps->remove_attribute_ns( name = 'id' uri = mc_rel ).
      ENDIF.
    ENDIF.

    " ---- <drawing r:id="..."/> ở cuối worksheet ---------------------------
    " Template do XCO sinh KHÔNG khai xmlns:r ở root worksheet, nên phải khai
    " trước khi dùng prefix r — thiếu bước này file bị "unbound prefix" và
    " Excel từ chối mở (openpyxl cũng không đọc được).
    io_ws->set_attribute( name = 'xmlns:r' value = mc_rel ).
    DATA(lo_dw_el) = mo_doc->create_element_ns( name = 'drawing' uri = mc_main ).
    lo_dw_el->set_attribute( name = 'r:id' value = lv_rid ).
    io_ws->append_child( lo_dw_el ).

    " ---- [Content_Types].xml ---------------------------------------------
    DATA(lv_types) = cl_abap_conv_codepage=>create_in( codepage = 'UTF-8'
                       )->convert( zip_read( '[Content_Types].xml' ) ).
    IF lv_types NS 'Extension="png"'.
      REPLACE FIRST OCCURRENCE OF '<Types' IN lv_types WITH '<TypesXLWBMARK'.
      REPLACE FIRST OCCURRENCE OF PCRE '(?-x)<TypesXLWBMARK([^>]*)>' IN lv_types
        WITH '<Types$1><Default Extension="png" ContentType="image/png"/>'.
    ENDIF.
    IF lv_types NS |/{ lv_dw_path }|.
      REPLACE FIRST OCCURRENCE OF '</Types>' IN lv_types WITH
        |<Override PartName="/{ lv_dw_path }"| &&
        | ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>|.
    ENDIF.
    zip_write( iv_name = '[Content_Types].xml' iv_raw = lo_conv->convert( lv_types ) ).
  ENDMETHOD.


  METHOD register_png_type.
    RETURN. " gộp vào attach_images
  ENDMETHOD.


  METHOD cells_of.
    " <row> → các <c>; nếu io_scope chính là <c> thì trả về chính nó
    IF io_scope->get_name( ) = 'c'.
      APPEND io_scope TO rt_elems.
      RETURN.
    ENDIF.
    rt_elems = elem_children( io_parent = io_scope iv_name = 'c' ).
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


  METHOD split_ref.
    CLEAR: ev_col, ev_row.
    DATA lv_col TYPE string.
    DATA lv_row TYPE string.
    FIND PCRE '(?-x)^([A-Z]+)(\d+)$' IN iv_ref SUBMATCHES lv_col lv_row.
    IF sy-subrc = 0.
      ev_col = lv_col.
      ev_row = CONV i( lv_row ).
    ENDIF.
  ENDMETHOD.


  METHOD col_to_num.
    DATA(lv_n) = 0.
    DO strlen( iv_col ) TIMES.
      DATA(lv_ch) = substring( val = iv_col off = sy-index - 1 len = 1 ).
      lv_n = lv_n * 26 + ( find( val = `ABCDEFGHIJKLMNOPQRSTUVWXYZ` sub = lv_ch ) + 1 ).
    ENDDO.
    rv_num = lv_n.
  ENDMETHOD.


  METHOD num_to_col.
    DATA(lv_n) = iv_num.
    WHILE lv_n > 0.
      DATA(lv_r) = ( lv_n - 1 ) MOD 26.
      rv_col = substring( val = `ABCDEFGHIJKLMNOPQRSTUVWXYZ` off = lv_r len = 1 ) && rv_col.
      lv_n = ( lv_n - 1 - lv_r ) / 26.
    ENDWHILE.
    IF rv_col IS INITIAL.
      rv_col = `A`.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
