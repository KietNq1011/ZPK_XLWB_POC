"! <p class="shorttext synchronized" lang="en">XLWB: app tải XML part từ DDIC</p>
"!
"! HTTP service: nhập structure DDIC / bảng DB / table type -> tải file XML part
"! để add vào Word (Developer &gt; XML Mapping Pane) và map content control.
"! URL: /sap/bc/http/sap/zxlwb_docxpart?sap-client=NNN
CLASS zcl_xlwb_part_http DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_http_service_extension.

  PRIVATE SECTION.
    METHODS render_page
      RETURNING VALUE(rv_html) TYPE string.
ENDCLASS.



CLASS ZCL_XLWB_PART_HTTP IMPLEMENTATION.


  METHOD if_http_service_extension~handle_request.
    DATA(lv_spec) = request->get_form_field( `spec` ).
    DATA(lv_mode) = request->get_form_field( `mode` ).

    " Không có spec -> trả trang nhập liệu
    IF lv_spec IS INITIAL.
      response->set_header_field( i_name  = `Content-Type`
                                  i_value = `text/html; charset=utf-8` ).
      response->set_text( render_page( ) ).
      RETURN.
    ENDIF.

    TRY.
        NEW zcl_xlwb_docx( )->build_part_xml(
          EXPORTING iv_spec        = lv_spec
          IMPORTING ev_xml         = DATA(lv_xml)
                    ev_sample_json = DATA(lv_json)
                    ev_namespace   = DATA(lv_ns) ).
      CATCH zcx_xlwb INTO DATA(lx).
        response->set_status( i_code = 400 i_reason = `Bad Request` ).
        response->set_header_field( i_name  = `Content-Type`
                                    i_value = `text/plain; charset=utf-8` ).
        response->set_text( |Lỗi: { lx->get_text( ) }| ).
        RETURN.
    ENDTRY.

    IF lv_mode = `json`.
      response->set_header_field( i_name  = `Content-Type`
                                  i_value = `application/json; charset=utf-8` ).
      response->set_header_field( i_name  = `Content-Disposition`
                                  i_value = `attachment; filename="sample.json"` ).
      response->set_text( lv_json ).
    ELSEIF lv_mode = `view`.
      response->set_header_field( i_name  = `Content-Type`
                                  i_value = `text/plain; charset=utf-8` ).
      response->set_text( lv_xml && cl_abap_char_utilities=>cr_lf &&
                          cl_abap_char_utilities=>cr_lf && lv_json ).
    ELSE.
      response->set_header_field( i_name  = `Content-Type`
                                  i_value = `application/xml; charset=utf-8` ).
      response->set_header_field( i_name  = `Content-Disposition`
                                  i_value = `attachment; filename="xlwb_part.xml"` ).
      response->set_text( lv_xml ).
    ENDIF.
  ENDMETHOD.


  METHOD render_page.
    rv_html =
      `<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8">` &&
      `<title>XLWB — XML Part từ DDIC</title>` &&
      `<style>` &&
      `body{font-family:"Segoe UI",Arial,sans-serif;max-width:860px;margin:2rem auto;padding:0 1rem;color:#222;line-height:1.5}` &&
      `h1{font-size:1.4rem}h2{font-size:1.05rem;margin-top:1.6rem}` &&
      `textarea{width:100%;height:9rem;font-family:Consolas,monospace;font-size:.95rem;padding:.5rem;box-sizing:border-box}` &&
      `button{padding:.5rem 1.2rem;margin:.6rem .4rem 0 0;font-size:1rem;cursor:pointer}` &&
      `pre{background:#f4f4f4;padding:.8rem;overflow-x:auto;font-size:.85rem}` &&
      `code{background:#f4f4f4;padding:0 .25rem}` &&
      `.note{background:#fff8e1;border-left:4px solid #f0ad4e;padding:.6rem .8rem}` &&
      `</style></head><body>` &&
      `<h1>Sinh XML Part cho Word từ structure / bảng SAP</h1>` &&
      `<p>Nhập spec — mỗi khai báo một dòng (hoặc cách nhau bằng <code>;</code>):</p>` &&
      `<form method="get">` &&
      `<textarea name="spec" placeholder="header=ZSD_HOPDONG_HD` && cl_abap_char_utilities=>cr_lf &&
      `table=items:ZSD_HOPDONG_IT` && cl_abap_char_utilities=>cr_lf &&
      `table=details:ZSD_KHAC"></textarea><br>` &&
      `<button type="submit">⬇ Tải XML Part</button>` &&
      `<button type="submit" name="mode" value="json">⬇ Tải Sample JSON</button>` &&
      `<button type="submit" name="mode" value="view">Xem trước</button>` &&
      `</form>` &&
      `<h2>Cú pháp spec</h2>` &&
      `<pre>header=ZSTRUCT_HD        &lt;- field nằm ngay dưới root (structure/bảng DB/table type)` && cl_abap_char_utilities=>cr_lf &&
      `table=items:ZSTRUCT_IT   &lt;- bảng lặp &lt;items&gt;&lt;row&gt;… — khai nhiều dòng table= được` && cl_abap_char_utilities=>cr_lf &&
      `root=data                &lt;- tuỳ chọn (mặc định data)` && cl_abap_char_utilities=>cr_lf &&
      `ns=urn:zdocx:data        &lt;- tuỳ chọn (mặc định urn:zdocx:data)</pre>` &&
      `<h2>Add file XML vào Word</h2>` &&
      `<p>Word không có nút thêm custom XML part — chạy macro sau <b>một lần</b> ` &&
      `(Alt+F11 &gt; chuột phải Normal &gt; Insert &gt; Module, dán vào, rồi Alt+F8 chạy ` &&
      `<code>AddXmlPartToDoc</code> khi đang mở file template):</p>` &&
      `<pre>Sub AddXmlPartToDoc()` && cl_abap_char_utilities=>cr_lf &&
      `    Dim f As FileDialog: Set f = Application.FileDialog(msoFileDialogFilePicker)` && cl_abap_char_utilities=>cr_lf &&
      `    f.Filters.Clear: f.Filters.Add "XML", "*.xml"` && cl_abap_char_utilities=>cr_lf &&
      `    If f.Show &lt;&gt; -1 Then Exit Sub` && cl_abap_char_utilities=>cr_lf &&
      `    Dim st As Object: Set st = CreateObject("ADODB.Stream")` && cl_abap_char_utilities=>cr_lf &&
      `    st.Type = 2: st.Charset = "utf-8": st.Open` && cl_abap_char_utilities=>cr_lf &&
      `    st.LoadFromFile f.SelectedItems(1)` && cl_abap_char_utilities=>cr_lf &&
      `    ActiveDocument.CustomXMLParts.Add st.ReadText: st.Close` && cl_abap_char_utilities=>cr_lf &&
      `    MsgBox "Đã thêm XML part. Mở Developer &gt; XML Mapping Pane để map."` && cl_abap_char_utilities=>cr_lf &&
      `End Sub</pre>` &&
      `<h2>Map trong Word</h2>` &&
      `<p>Developer &gt; XML Mapping Pane &gt; chọn part (mặc định <code>urn:zdocx:data</code>). ` &&
      `Field đơn: chuột phải node &gt; Insert Content Control &gt; Plain Text / Check Box / Picture. ` &&
      `Bảng lặp: vẽ bảng 1 dòng mẫu, chọn cả dòng, chuột phải node <b>row</b> &gt; ` &&
      `Insert Content Control &gt; Repeating, rồi map từng cell vào field con của row.</p>` &&
      `<p class="note">Giá trị mẫu trong part = tên field (chữ thường) để nhìn là biết map field nào. ` &&
      `Lưu template xong: upload vào app Maintain Repository với ENGINE = DOCX và dán Sample JSON ` &&
      `vào cột Sample JSON. Code nghiệp vụ gọi ` &&
      `<code>zcl_xlwb_runtime=&gt;render( iv_form_name = ... ir_context = REF #( ls_data ) )</code>.</p>` &&
      `</body></html>`.
  ENDMETHOD.
ENDCLASS.
