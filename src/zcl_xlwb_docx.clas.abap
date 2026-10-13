"! <p class="shorttext synchronized" lang="en">XLWB: engine DOCX (Word XML Mapping)</p>
"!
"! Render file Word .docx theo cơ chế XML Mapping (w:dataBinding): template do
"! designer tạo trong Word — content control map vào custom XML part qua
"! Developer &gt; XML Mapping Pane. Engine KHÔNG đụng word/document.xml; chỉ sinh
"! lại data XML từ ir_context rồi thay entry customXml/itemN.xml trong ZIP.
"! Word tự refresh giá trị + tự nhân bản repeating section khi mở file
"! (đã kiểm chứng POC 2026-08-27, kể cả checkbox và bảng N dòng).
"!
"! Quy ước sinh XML từ context (prepare_docx dùng CÙNG quy ước):
"! - component structure -&gt; element trùng tên component (chữ thường)
"! - component bảng      -&gt; element bao trùng tên, mỗi dòng một element &lt;row&gt;
"! - elementary          -&gt; text đã escape XML
"! Tên root + namespace lấy TỪ custom XML part sẵn có trong template, nên
"! designer đặt tuỳ ý — chỉ cần cây con khớp quy ước trên.
CLASS zcl_xlwb_docx DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    CONSTANTS c_mime_docx TYPE string
      VALUE `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.

    "! Render template .docx với dữ liệu context
    "! @parameter iv_template | file .docx (ZXLWB_TMPL-TEMPLATE)
    "! @parameter ir_context  | REF TO structure dữ liệu bind
    METHODS render
      IMPORTING iv_template    TYPE xstring
                ir_context     TYPE REF TO data
      RETURNING VALUE(rv_docx) TYPE xstring
      RAISING   zcx_xlwb.

    "! Sinh data XML từ context (public để unit test và prepare_docx dùng lại)
    METHODS build_data_xml
      IMPORTING iv_root_name  TYPE string
                iv_namespace  TYPE string
                ir_context    TYPE REF TO data
      RETURNING VALUE(rv_xml) TYPE string
      RAISING   zcx_xlwb.

    "! Chuẩn bị template Word từ DDIC structure/table: sinh XML part mẫu rồi
    "! nhúng vào docx để designer map control qua XML Mapping Pane.
    "! Spec (mỗi khai báo cách nhau bằng ';' hoặc xuống dòng):
    "!   header=ZSTRUCT_HD          - field của structure nằm ngay dưới root
    "!   table=items:ZSTRUCT_IT     - bảng lặp &lt;items&gt;&lt;row&gt;...  (khai nhiều dòng table= được)
    "!   root=data                  - tuỳ chọn, mặc định 'data'
    "!   ns=urn:zdocx:data          - tuỳ chọn, mặc định 'urn:zdocx:data'
    "! @parameter ev_docx        | docx đã nhúng part (mở Word để map)
    "! @parameter ev_sample_json | JSON mẫu cùng cấu trúc — lưu vào ZXLWB_TMPL-SAMPLE_JSON
    METHODS prepare_docx
      IMPORTING iv_docx        TYPE xstring
                iv_spec        TYPE string
      EXPORTING ev_docx        TYPE xstring
                ev_sample_json TYPE string
      RAISING   zcx_xlwb.

    "! Sinh XML part + sample JSON từ spec DDIC (không cần file docx) —
    "! dùng cho app tải XML part về add tay vào Word
    METHODS build_part_xml
      IMPORTING iv_spec        TYPE string
      EXPORTING ev_xml         TYPE string
                ev_sample_json TYPE string
                ev_namespace   TYPE string
      RAISING   zcx_xlwb.

    "! Sinh XML part + sample JSON từ MỘT tên DDIC — tự nhận diện loại:
    "! structure / bảng DB -> field nằm dưới root (bảng lồng thành &lt;comp&gt;&lt;row&gt;...);
    "! table type -> &lt;rows&gt;&lt;row&gt;... (context lúc render phải có component ROWS)
    METHODS build_part_from_ddic
      IMPORTING iv_ddic        TYPE string
                iv_root        TYPE string OPTIONAL
                iv_ns          TYPE string OPTIONAL
      EXPORTING ev_xml         TYPE string
                ev_sample_json TYPE string
                ev_kind        TYPE string
      RAISING   zcx_xlwb.

    "! Nhúng 1 custom XML part mới vào docx (part + itemProps + rels + content types)
    METHODS inject_part
      IMPORTING iv_docx        TYPE xstring
                iv_xml         TYPE string
                iv_namespace   TYPE string
      RETURNING VALUE(rv_docx) TYPE xstring
      RAISING   zcx_xlwb.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_table_spec,
             name TYPE string,
             ddic TYPE string,
           END OF ty_table_spec,
           tt_table_spec TYPE STANDARD TABLE OF ty_table_spec WITH EMPTY KEY.

    "! Tìm custom XML part dữ liệu trong docx: customXml/item*.xml (bỏ itemProps),
    "! ưu tiên part có namespace KHÔNG phải built-in của Microsoft.
    METHODS read_part_info
      IMPORTING io_zip       TYPE REF TO cl_abap_zip
      EXPORTING ev_part_name TYPE string
                ev_root_name TYPE string
                ev_namespace TYPE string
      RAISING   zcx_xlwb.

    METHODS parse_root
      IMPORTING iv_content   TYPE xstring
      EXPORTING ev_root_name TYPE string
                ev_namespace TYPE string.

    METHODS append_node
      IMPORTING iv_name TYPE string
                ir_data TYPE REF TO data
      CHANGING  cv_xml  TYPE string
      RAISING   zcx_xlwb.

    METHODS append_children
      IMPORTING ir_data TYPE REF TO data
      CHANGING  cv_xml  TYPE string
      RAISING   zcx_xlwb.

    METHODS parse_spec
      IMPORTING iv_spec   TYPE string
      EXPORTING ev_header TYPE string
                et_tables TYPE tt_table_spec
                ev_root   TYPE string
                ev_ns     TYPE string
      RAISING   zcx_xlwb.

    "! describe_by_name cho structure / DB table / table type -> structdescr dòng
    METHODS struct_descr_of
      IMPORTING iv_name          TYPE string
      RETURNING VALUE(ro_struct) TYPE REF TO cl_abap_structdescr
      RAISING   zcx_xlwb.

    "! Điền giá trị mẫu = tên field (chữ thường) cho field ký tự/chuỗi
    METHODS fill_sample
      IMPORTING ir_data TYPE REF TO data.

    "! '"f1":"v1","f2":"v2"' cho các component của structure (đệ quy)
    METHODS json_members
      IMPORTING ir_data        TYPE REF TO data
      RETURNING VALUE(rv_json) TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_DOCX IMPLEMENTATION.


  METHOD render.
    IF ir_context IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Context reference is initial` ).
    ENDIF.

    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING zip             = iv_template
                  EXCEPTIONS zip_parse_error = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = `Template không phải file .docx hợp lệ (ZIP parse lỗi)` ).
    ENDIF.

    read_part_info( EXPORTING io_zip       = lo_zip
                    IMPORTING ev_part_name = DATA(lv_part)
                              ev_root_name = DATA(lv_root)
                              ev_namespace = DATA(lv_ns) ).

    DATA(lv_xml) = build_data_xml( iv_root_name = lv_root
                                   iv_namespace = lv_ns
                                   ir_context   = ir_context ).

    lo_zip->delete( EXPORTING name            = lv_part
                    EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Không xoá được part '{ lv_part }'| ).
    ENDIF.

    lo_zip->add( name    = lv_part
                 content = cl_abap_conv_codepage=>create_out( )->convert( lv_xml ) ).

    rv_docx = lo_zip->save( ).
  ENDMETHOD.


  METHOD read_part_info.
    CLEAR: ev_part_name, ev_root_name, ev_namespace.

    DATA lv_fallback_part TYPE string.
    DATA lv_fallback_root TYPE string.
    DATA lv_fallback_ns   TYPE string.

    LOOP AT io_zip->files INTO DATA(ls_file).
      IF NOT ls_file-name CP 'customxml/item*.xml'
         OR ls_file-name CP 'customxml/itemprops*'.
        CONTINUE.
      ENDIF.

      io_zip->get( EXPORTING name            = ls_file-name
                   IMPORTING content         = DATA(lv_content)
                   EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      parse_root( EXPORTING iv_content   = lv_content
                  IMPORTING ev_root_name = DATA(lv_root)
                            ev_namespace = DATA(lv_ns) ).
      IF lv_root IS INITIAL.
        CONTINUE.
      ENDIF.

      " Part built-in của Word (coverPageProps, contentTypeSchema...) không phải data
      IF lv_ns CP 'http://schemas.openxmlformats.org*'
         OR lv_ns CP 'http://schemas.microsoft.com*'.
        IF lv_fallback_part IS INITIAL.
          lv_fallback_part = ls_file-name.
          lv_fallback_root = lv_root.
          lv_fallback_ns   = lv_ns.
        ENDIF.
        CONTINUE.
      ENDIF.

      ev_part_name = ls_file-name.
      ev_root_name = lv_root.
      ev_namespace = lv_ns.
      RETURN.
    ENDLOOP.

    IF lv_fallback_part IS NOT INITIAL.
      ev_part_name = lv_fallback_part.
      ev_root_name = lv_fallback_root.
      ev_namespace = lv_fallback_ns.
      RETURN.
    ENDIF.

    RAISE EXCEPTION NEW zcx_xlwb(
      iv_text = `Template chưa có custom XML part — chạy Prepare template để nhúng ` &&
                `data XML mẫu rồi map content control trong Word (XML Mapping Pane)` ).
  ENDMETHOD.


  METHOD parse_root.
    CLEAR: ev_root_name, ev_namespace.
    TRY.
        DATA(lo_reader) = cl_sxml_string_reader=>create( iv_content ).
        DO.
          DATA(lo_node) = lo_reader->read_next_node( ).
          IF lo_node IS INITIAL.
            EXIT.
          ENDIF.
          IF lo_node->type = if_sxml_node=>co_nt_element_open.
            DATA(lo_open) = CAST if_sxml_open_element( lo_node ).
            ev_root_name = lo_open->qname-name.
            ev_namespace = lo_open->qname-namespace.
            EXIT.
          ENDIF.
        ENDDO.
      CATCH cx_sxml_parse_error.
        CLEAR: ev_root_name, ev_namespace.
    ENDTRY.
  ENDMETHOD.


  METHOD build_data_xml.
    IF ir_context IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Context reference is initial` ).
    ENDIF.
    IF iv_root_name IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Root element của custom XML part rỗng` ).
    ENDIF.

    DATA lv_body TYPE string.
    append_children( EXPORTING ir_data = ir_context CHANGING cv_xml = lv_body ).

    DATA(lv_ns_attr) = COND string( WHEN iv_namespace IS NOT INITIAL
      THEN | xmlns="{ escape( val = iv_namespace format = cl_abap_format=>e_xml_attr ) }"|
      ELSE `` ).

    rv_xml = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
             |<{ iv_root_name }{ lv_ns_attr }>{ lv_body }</{ iv_root_name }>|.
  ENDMETHOD.


  METHOD append_children.
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data_ref( ir_data ).
    IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Context phải là REF TO structure` ).
    ENDIF.

    DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
    FIELD-SYMBOLS <ls_ctx> TYPE any.
    ASSIGN ir_data->* TO <ls_ctx>.
    LOOP AT lo_struct->components INTO DATA(ls_comp).
      IF ls_comp-name = 'MANDT' OR ls_comp-name = 'CLIENT'.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_comp-name OF STRUCTURE <ls_ctx> TO FIELD-SYMBOL(<lv_comp>).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      append_node( EXPORTING iv_name = to_lower( ls_comp-name )
                             ir_data = REF #( <lv_comp> )
                   CHANGING  cv_xml  = cv_xml ).
    ENDLOOP.
  ENDMETHOD.


  METHOD append_node.
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data_ref( ir_data ).

    CASE lo_type->kind.

      WHEN cl_abap_typedescr=>kind_struct.
        cv_xml = cv_xml && |<{ iv_name }>|.
        append_children( EXPORTING ir_data = ir_data CHANGING cv_xml = cv_xml ).
        cv_xml = cv_xml && |</{ iv_name }>|.

      WHEN cl_abap_typedescr=>kind_table.
        cv_xml = cv_xml && |<{ iv_name }>|.
        FIELD-SYMBOLS <lt> TYPE ANY TABLE.
        ASSIGN ir_data->* TO <lt>.
        LOOP AT <lt> ASSIGNING FIELD-SYMBOL(<ls_line>).
          append_node( EXPORTING iv_name = `row`
                                 ir_data = REF #( <ls_line> )
                       CHANGING  cv_xml  = cv_xml ).
        ENDLOOP.
        cv_xml = cv_xml && |</{ iv_name }>|.

      WHEN cl_abap_typedescr=>kind_elem.
        FIELD-SYMBOLS <lv> TYPE any.
        ASSIGN ir_data->* TO <lv>.
        cv_xml = cv_xml &&
          |<{ iv_name }>{ escape( val = |{ <lv> }| format = cl_abap_format=>e_xml_text ) }</{ iv_name }>|.

      WHEN cl_abap_typedescr=>kind_ref.
        FIELD-SYMBOLS <lr> TYPE any.
        ASSIGN ir_data->* TO <lr>.
        DATA lr_inner TYPE REF TO data.
        TRY.
            lr_inner = <lr>.
          CATCH cx_sy_move_cast_error.
            RETURN.
        ENDTRY.
        IF lr_inner IS BOUND.
          append_node( EXPORTING iv_name = iv_name
                                 ir_data = lr_inner
                       CHANGING  cv_xml  = cv_xml ).
        ENDIF.

    ENDCASE.
  ENDMETHOD.


  METHOD prepare_docx.
    CLEAR: ev_docx, ev_sample_json.

    build_part_xml( EXPORTING iv_spec        = iv_spec
                    IMPORTING ev_xml         = DATA(lv_xml)
                              ev_sample_json = ev_sample_json
                              ev_namespace   = DATA(lv_ns) ).

    ev_docx = inject_part( iv_docx      = iv_docx
                           iv_xml       = lv_xml
                           iv_namespace = lv_ns ).
  ENDMETHOD.


  METHOD build_part_xml.
    CLEAR: ev_xml, ev_sample_json, ev_namespace.

    parse_spec( EXPORTING iv_spec   = iv_spec
                IMPORTING ev_header = DATA(lv_header)
                          et_tables = DATA(lt_tables)
                          ev_root   = DATA(lv_root)
                          ev_ns     = DATA(lv_ns) ).

    DATA lv_body TYPE string.
    DATA lv_json TYPE string.

    " Field header nằm ngay dưới root
    IF lv_header IS NOT INITIAL.
      DATA(lo_hd) = struct_descr_of( lv_header ).
      DATA lr_hd TYPE REF TO data.
      CREATE DATA lr_hd TYPE HANDLE lo_hd.
      fill_sample( lr_hd ).
      append_children( EXPORTING ir_data = lr_hd CHANGING cv_xml = lv_body ).
      lv_json = json_members( lr_hd ).
    ENDIF.

    " Mỗi bảng: <name><row>...</row><row>...</row></name> — 2 dòng mẫu
    LOOP AT lt_tables INTO DATA(ls_table).
      DATA(lo_it) = struct_descr_of( ls_table-ddic ).
      DATA lr_row TYPE REF TO data.
      CREATE DATA lr_row TYPE HANDLE lo_it.
      fill_sample( lr_row ).

      DATA(lv_row_xml) = ``.
      append_node( EXPORTING iv_name = `row` ir_data = lr_row
                   CHANGING  cv_xml  = lv_row_xml ).
      lv_body = lv_body && |<{ ls_table-name }>{ lv_row_xml }{ lv_row_xml }</{ ls_table-name }>|.

      DATA(lv_row_json) = |\{{ json_members( lr_row ) }\}|.
      IF lv_json IS NOT INITIAL.
        lv_json = lv_json && `,`.
      ENDIF.
      lv_json = lv_json && |"{ ls_table-name }":[{ lv_row_json },{ lv_row_json }]|.
    ENDLOOP.

    ev_xml = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
             |<{ lv_root } xmlns="{ escape( val = lv_ns format = cl_abap_format=>e_xml_attr ) }">| &&
             |{ lv_body }</{ lv_root }>|.

    ev_sample_json = |\{{ lv_json }\}|.
    ev_namespace   = lv_ns.
  ENDMETHOD.


  METHOD build_part_from_ddic.
    CLEAR: ev_xml, ev_sample_json, ev_kind.

    DATA(lv_name) = to_upper( condense( iv_ddic ) ).
    IF lv_name IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Chưa nhập tên DDIC` ).
    ENDIF.

    cl_abap_typedescr=>describe_by_name(
      EXPORTING p_name         = lv_name
      RECEIVING p_descr_ref    = DATA(lo_type)
      EXCEPTIONS type_not_found = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Không tìm thấy DDIC type '{ lv_name }'| ).
    ENDIF.

    DATA lv_body TYPE string.
    DATA lv_json TYPE string.
    DATA lr_data TYPE REF TO data.

    CASE lo_type->kind.

      WHEN cl_abap_typedescr=>kind_struct.
        ev_kind = `STRUCTURE`.
        DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
        CREATE DATA lr_data TYPE HANDLE lo_struct.
        fill_sample( lr_data ).
        append_children( EXPORTING ir_data = lr_data CHANGING cv_xml = lv_body ).
        lv_json = json_members( lr_data ).

      WHEN cl_abap_typedescr=>kind_table.
        ev_kind = `TABLE TYPE`.
        DATA(lo_tab)  = CAST cl_abap_tabledescr( lo_type ).
        DATA(lo_line) = lo_tab->get_table_line_type( ).
        CREATE DATA lr_data TYPE HANDLE lo_tab.

        DATA lr_line TYPE REF TO data.
        CREATE DATA lr_line TYPE HANDLE lo_line.
        IF lo_line->kind = cl_abap_typedescr=>kind_struct.
          fill_sample( lr_line ).
        ENDIF.
        FIELD-SYMBOLS <ls_line> TYPE any.
        ASSIGN lr_line->* TO <ls_line>.
        FIELD-SYMBOLS <lt> TYPE ANY TABLE.
        ASSIGN lr_data->* TO <lt>.
        INSERT <ls_line> INTO TABLE <lt>.
        INSERT <ls_line> INTO TABLE <lt>.

        append_node( EXPORTING iv_name = `rows`
                               ir_data = lr_data
                     CHANGING  cv_xml  = lv_body ).

        DATA(lv_row_json) = COND string(
          WHEN lo_line->kind = cl_abap_typedescr=>kind_struct
          THEN |\{{ json_members( lr_line ) }\}|
          ELSE `""` ).
        lv_json = |"rows":[{ lv_row_json },{ lv_row_json }]|.

      WHEN OTHERS.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |'{ lv_name }' phải là structure, bảng DB hoặc table type| ).
    ENDCASE.

    DATA(lv_root) = COND string( WHEN iv_root IS INITIAL THEN `data` ELSE condense( iv_root ) ).
    DATA(lv_ns)   = COND string( WHEN iv_ns IS INITIAL THEN `urn:zdocx:data` ELSE condense( iv_ns ) ).

    ev_xml = |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
             |<{ lv_root } xmlns="{ escape( val = lv_ns format = cl_abap_format=>e_xml_attr ) }">| &&
             |{ lv_body }</{ lv_root }>|.
    ev_sample_json = |\{{ lv_json }\}|.
  ENDMETHOD.


  METHOD parse_spec.
    CLEAR: ev_header, et_tables.
    ev_root = `data`.
    ev_ns   = `urn:zdocx:data`.

    DATA(lv_spec) = iv_spec.
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf   IN lv_spec WITH `;`.
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN lv_spec WITH `;`.

    SPLIT lv_spec AT `;` INTO TABLE DATA(lt_lines).
    LOOP AT lt_lines INTO DATA(lv_line).
      CONDENSE lv_line.
      IF lv_line IS INITIAL.
        CONTINUE.
      ENDIF.
      SPLIT lv_line AT `=` INTO DATA(lv_key) DATA(lv_val).
      CONDENSE: lv_key, lv_val.
      CASE to_lower( lv_key ).
        WHEN `header`.
          ev_header = to_upper( lv_val ).
        WHEN `table`.
          SPLIT lv_val AT `:` INTO DATA(lv_tname) DATA(lv_tddic).
          CONDENSE: lv_tname, lv_tddic.
          IF lv_tname IS INITIAL OR lv_tddic IS INITIAL.
            RAISE EXCEPTION NEW zcx_xlwb(
              iv_text = |Spec bảng sai '{ lv_line }' — đúng dạng table=items:ZSTRUCT_IT| ).
          ENDIF.
          APPEND VALUE #( name = to_lower( lv_tname )
                          ddic = to_upper( lv_tddic ) ) TO et_tables.
        WHEN `root`.
          ev_root = lv_val.
        WHEN `ns`.
          ev_ns = lv_val.
        WHEN OTHERS.
          RAISE EXCEPTION NEW zcx_xlwb(
            iv_text = |Spec không hiểu '{ lv_line }' — chỉ hỗ trợ header=, table=, root=, ns=| ).
      ENDCASE.
    ENDLOOP.

    IF ev_header IS INITIAL AND et_tables IS INITIAL.
      RAISE EXCEPTION NEW zcx_xlwb(
        iv_text = `Spec rỗng — khai ít nhất header=ZSTRUCT hoặc table=items:ZSTRUCT` ).
    ENDIF.
  ENDMETHOD.


  METHOD struct_descr_of.
    cl_abap_typedescr=>describe_by_name(
      EXPORTING p_name         = iv_name
      RECEIVING p_descr_ref    = DATA(lo_type)
      EXCEPTIONS type_not_found = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = |Không tìm thấy DDIC type '{ iv_name }'| ).
    ENDIF.

    CASE lo_type->kind.
      WHEN cl_abap_typedescr=>kind_struct.
        ro_struct = CAST cl_abap_structdescr( lo_type ).
      WHEN cl_abap_typedescr=>kind_table.
        DATA(lo_line) = CAST cl_abap_tabledescr( lo_type )->get_table_line_type( ).
        IF lo_line->kind <> cl_abap_typedescr=>kind_struct.
          RAISE EXCEPTION NEW zcx_xlwb(
            iv_text = |Table type '{ iv_name }' không có dòng dạng structure| ).
        ENDIF.
        ro_struct = CAST cl_abap_structdescr( lo_line ).
      WHEN OTHERS.
        RAISE EXCEPTION NEW zcx_xlwb(
          iv_text = |'{ iv_name }' phải là structure, bảng DB hoặc table type| ).
    ENDCASE.
  ENDMETHOD.


  METHOD fill_sample.
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data_ref( ir_data ).
    IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
      RETURN.
    ENDIF.

    DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
    FIELD-SYMBOLS <ls> TYPE any.
    ASSIGN ir_data->* TO <ls>.

    LOOP AT lo_struct->components INTO DATA(ls_comp).
      ASSIGN COMPONENT ls_comp-name OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv>).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      CASE ls_comp-type_kind.
        WHEN cl_abap_typedescr=>typekind_char OR cl_abap_typedescr=>typekind_string.
          " giá trị mẫu = tên field để designer nhìn là biết map field nào
          <lv> = to_lower( ls_comp-name ).
        WHEN cl_abap_typedescr=>typekind_struct1 OR cl_abap_typedescr=>typekind_struct2.
          fill_sample( REF #( <lv> ) ).
        WHEN cl_abap_typedescr=>typekind_table.
          " bảng lồng: chèn 2 dòng mẫu để designer map được field trong <row>
          DATA(lo_tab)  = CAST cl_abap_tabledescr( cl_abap_typedescr=>describe_by_data( <lv> ) ).
          DATA(lo_line) = lo_tab->get_table_line_type( ).
          DATA lr_line TYPE REF TO data.
          CREATE DATA lr_line TYPE HANDLE lo_line.
          IF lo_line->kind = cl_abap_typedescr=>kind_struct.
            fill_sample( lr_line ).
          ENDIF.
          FIELD-SYMBOLS <ls_new> TYPE any.
          ASSIGN lr_line->* TO <ls_new>.
          FIELD-SYMBOLS <lt_tab> TYPE ANY TABLE.
          ASSIGN <lv> TO <lt_tab>.
          INSERT <ls_new> INTO TABLE <lt_tab>.
          INSERT <ls_new> INTO TABLE <lt_tab>.
        WHEN OTHERS.
          " numeric / date / raw... để initial
      ENDCASE.
    ENDLOOP.
  ENDMETHOD.


  METHOD json_members.
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data_ref( ir_data ).
    IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
      RETURN.
    ENDIF.

    DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
    FIELD-SYMBOLS <ls> TYPE any.
    ASSIGN ir_data->* TO <ls>.

    LOOP AT lo_struct->components INTO DATA(ls_comp).
      IF ls_comp-name = 'MANDT' OR ls_comp-name = 'CLIENT'.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_comp-name OF STRUCTURE <ls> TO FIELD-SYMBOL(<lv>).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      DATA(lv_name) = to_lower( ls_comp-name ).
      DATA(lv_member) = ``.
      CASE ls_comp-type_kind.
        WHEN cl_abap_typedescr=>typekind_struct1 OR cl_abap_typedescr=>typekind_struct2.
          lv_member = |"{ lv_name }":\{{ json_members( REF #( <lv> ) ) }\}|.
        WHEN cl_abap_typedescr=>typekind_table.
          FIELD-SYMBOLS <lt_j> TYPE ANY TABLE.
          ASSIGN <lv> TO <lt_j>.
          DATA(lv_rows) = ``.
          LOOP AT <lt_j> ASSIGNING FIELD-SYMBOL(<ls_row_j>).
            IF lv_rows IS NOT INITIAL.
              lv_rows = lv_rows && `,`.
            ENDIF.
            lv_rows = lv_rows && |\{{ json_members( REF #( <ls_row_j> ) ) }\}|.
          ENDLOOP.
          lv_member = |"{ lv_name }":[{ lv_rows }]|.
        WHEN OTHERS.
          " escape tối thiểu cho giá trị mẫu (tên field / giá trị initial)
          DATA(lv_val) = |{ <lv> }|.
          REPLACE ALL OCCURRENCES OF `\` IN lv_val WITH `\\`.
          REPLACE ALL OCCURRENCES OF `"` IN lv_val WITH `\"`.
          lv_member = |"{ lv_name }":"{ lv_val }"|.
      ENDCASE.
      IF rv_json IS NOT INITIAL.
        rv_json = rv_json && `,`.
      ENDIF.
      rv_json = rv_json && lv_member.
    ENDLOOP.
  ENDMETHOD.


  METHOD inject_part.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING zip             = iv_docx
                  EXCEPTIONS zip_parse_error = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `File không phải .docx hợp lệ (ZIP parse lỗi)` ).
    ENDIF.

    DATA(lo_in)  = cl_abap_conv_codepage=>create_in( ).
    DATA(lo_out) = cl_abap_conv_codepage=>create_out( ).

    " Số thứ tự part mới = max(itemN) + 1
    DATA(lv_max) = 0.
    LOOP AT lo_zip->files INTO DATA(ls_file).
      FIND PCRE '(?i)^customxml/item(\d+)\.xml$' IN ls_file-name SUBMATCHES DATA(lv_num_c).
      IF sy-subrc = 0.
        lv_max = nmax( val1 = lv_max val2 = CONV i( lv_num_c ) ).
      ENDIF.
    ENDLOOP.
    DATA(lv_n) = lv_max + 1.

    DATA lv_uuid TYPE string.
    TRY.
        lv_uuid = cl_system_uuid=>create_uuid_c36_static( ).
      CATCH cx_uuid_error.
        lv_uuid = |00000000-0000-4000-8000-{ lv_n WIDTH = 12 PAD = '0' ALIGN = RIGHT }|.
    ENDTRY.
    DATA(lv_guid) = |\{{ to_upper( lv_uuid ) }\}|.

    " 1. Data part + itemProps + rels của part
    lo_zip->add( name    = |customXml/item{ lv_n }.xml|
                 content = lo_out->convert( iv_xml ) ).

    DATA(lv_ns_esc) = escape( val = iv_namespace format = cl_abap_format=>e_xml_attr ).
    lo_zip->add( name    = |customXml/itemProps{ lv_n }.xml|
                 content = lo_out->convert(
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<ds:datastoreItem ds:itemID="{ lv_guid }" | &&
      |xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">| &&
      |<ds:schemaRefs><ds:schemaRef ds:uri="{ lv_ns_esc }"/></ds:schemaRefs></ds:datastoreItem>| ) ).

    lo_zip->add( name    = |customXml/_rels/item{ lv_n }.xml.rels|
                 content = lo_out->convert(
      |<?xml version="1.0" encoding="UTF-8" standalone="yes"?>| &&
      |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">| &&
      |<Relationship Id="rId1" | &&
      |Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" | &&
      |Target="itemProps{ lv_n }.xml"/></Relationships>| ) ).

    " 2. [Content_Types].xml
    lo_zip->get( EXPORTING name            = '[Content_Types].xml'
                 IMPORTING content         = DATA(lv_ct_x)
                 EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW zcx_xlwb( iv_text = `Thiếu [Content_Types].xml — không phải docx` ).
    ENDIF.
    DATA(lv_ct) = lo_in->convert( lv_ct_x ).
    IF lv_ct NS `Extension="xml"`.
      REPLACE FIRST OCCURRENCE OF `</Types>` IN lv_ct
        WITH `<Default Extension="xml" ContentType="application/xml"/></Types>`.
    ENDIF.
    REPLACE FIRST OCCURRENCE OF `</Types>` IN lv_ct
      WITH |<Override PartName="/customXml/itemProps{ lv_n }.xml" | &&
           |ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>|.
    lo_zip->delete( name = '[Content_Types].xml' ).
    lo_zip->add( name = '[Content_Types].xml' content = lo_out->convert( lv_ct ) ).

    " 3. word/_rels/document.xml.rels
    DATA lv_rels TYPE string.
    lo_zip->get( EXPORTING name            = 'word/_rels/document.xml.rels'
                 IMPORTING content         = DATA(lv_rels_x)
                 EXCEPTIONS zip_index_error = 1 OTHERS = 2 ).
    IF sy-subrc = 0.
      lv_rels = lo_in->convert( lv_rels_x ).
      lo_zip->delete( name = 'word/_rels/document.xml.rels' ).
    ELSE.
      lv_rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` &&
                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` &&
                `</Relationships>`.
    ENDIF.

    DATA(lv_max_rid) = 0.
    FIND ALL OCCURRENCES OF PCRE 'Id="rId(\d+)"' IN lv_rels RESULTS DATA(lt_matches).
    LOOP AT lt_matches INTO DATA(ls_match).
      READ TABLE ls_match-submatches INTO DATA(ls_sub) INDEX 1.
      IF sy-subrc = 0.
        lv_max_rid = nmax( val1 = lv_max_rid
                           val2 = CONV i( substring( val = lv_rels
                                                     off = ls_sub-offset
                                                     len = ls_sub-length ) ) ).
      ENDIF.
    ENDLOOP.

    REPLACE FIRST OCCURRENCE OF `</Relationships>` IN lv_rels
      WITH |<Relationship Id="rId{ lv_max_rid + 1 }" | &&
           |Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" | &&
           |Target="../customXml/item{ lv_n }.xml"/></Relationships>|.
    lo_zip->add( name = 'word/_rels/document.xml.rels' content = lo_out->convert( lv_rels ) ).

    rv_docx = lo_zip->save( ).
  ENDMETHOD.
ENDCLASS.
