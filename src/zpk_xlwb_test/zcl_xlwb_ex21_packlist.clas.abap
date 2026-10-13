"! <p class="shorttext synchronized" lang="en">XLWB: Packing List HQ từ template thật</p>
"!
"! Bài thực hành "đưa form đang hardcode sang template": template dưới đây
"! chính là file PackingListHQ_0090000645.xls xuất từ ZCL_ECUS_FORM_PL_HQ,
"! chỉ thay dữ liệu bằng placeholder. Style, khung, độ rộng cột, chiều cao
"! dòng, print setup giữ nguyên 100%.
"!
"! Cách dùng:
"! <ul>
"! <li>Bấm F9 để NẠP template vào bảng ZXLWB_TMPL (form {@link zcl_xlwb_ex21_packlist.DATA:c_form}).</li>
"! <li>Sau đó mọi nơi chỉ cần: zcl_xlwb_runtime=>render( iv_form_name = 'PACKING_LIST_HQ' ... ).</li>
"! <li>Sửa hình thức form: mở app Fiori quản lý template, tải file về, sửa
"!     bằng Excel, Save As XML Spreadsheet 2003, upload lại — KHÔNG sửa ABAP.</li>
"! </ul>
CLASS zcl_xlwb_ex21_packlist DEFINITION
  PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_oo_adt_classrun.

    CONSTANTS c_form TYPE c LENGTH 40 VALUE 'PACKING_LIST_HQ'.

    TYPES: BEGIN OF ty_party,
             name    TYPE string,
             address TYPE string,
           END OF ty_party,
           BEGIN OF ty_item,
             so_item      TYPE string,
             item_no      TYPE string,
             delivery_no  TYPE string,
             commodity    TYPE string,
             qty_pcs      TYPE p LENGTH 13 DECIMALS 3,
             qty_ctns     TYPE p LENGTH 13 DECIMALS 3,
             pallet       TYPE p LENGTH 13 DECIMALS 3,
             gross_weight TYPE p LENGTH 13 DECIMALS 3,
             net_weight   TYPE p LENGTH 13 DECIMALS 3,
             container    TYPE string,
           END OF ty_item,
           ty_items TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_ctx,
             shipper           TYPE ty_party,
             buyer             TYPE ty_party,
             invoice_no        TYPE string,
             invoice_date      TYPE string,
             payment_term      TYPE string,
             remarks           TYPE string,
             port_of_shipment  TYPE string,
             port_of_discharge TYPE string,
             vessel            TYPE string,
             trunk_vessel      TYPE string,
             etd               TYPE string,
             container_summary TYPE string,
             package_summary   TYPE string,
             items             TYPE ty_items,
           END OF ty_ctx.

    "! Template SpreadsheetML (sinh từ file Excel thật)
    CLASS-METHODS get_template RETURNING VALUE(rv_xml) TYPE string.

    "! Dữ liệu mẫu bằng đúng dữ liệu của file gốc 0090000645
    CLASS-METHODS sample_context RETURNING VALUE(rs_ctx) TYPE ty_ctx.
ENDCLASS.



CLASS ZCL_XLWB_EX21_PACKLIST IMPLEMENTATION.


  METHOD get_template.
    rv_xml = `<?xml version="1.0"?>&#10;<?mso-application progid="Excel.Sheet"?>&#10;<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:o`.
    rv_xml = rv_xml && `ffice:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">&#10; `.
    rv_xml = rv_xml && `<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">&#10;  <WindowHeight>9492</WindowHeight>&#10;  <WindowWidth>23040</WindowWidth>&#10;  <WindowTopX>32767</Win`.
    rv_xml = rv_xml && `dowTopX>&#10;  <WindowTopY>32767</WindowTopY>&#10;  <ProtectStructure>False</ProtectStructure>&#10;  <ProtectWindows>False</ProtectWindows>&#10; </ExcelWorkbook>&#10; <St`.
    rv_xml = rv_xml && `yles>&#10;  <Style ss:ID="Default" ss:Name="Normal">&#10;   <Alignment ss:Vertical="Center"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10`.
    rv_xml = rv_xml && `;  <Style ss:ID="m2481854426404">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss`.
    rv_xml = rv_xml && `:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Con`.
    rv_xml = rv_xml && `tinuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" x:Family`.
    rv_xml = rv_xml && `="Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854426424">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Border`.
    rv_xml = rv_xml && `s>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Bo`.
    rv_xml = rv_xml && `rder ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <F`.
    rv_xml = rv_xml && `ont ss:FontName="Times New Roman" x:Family="Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854427008">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical`.
    rv_xml = rv_xml && `="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyl`.
    rv_xml = rv_xml && `e="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" s`.
    rv_xml = rv_xml && `s:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854427028">&#10;   <Alignment ss:Horizon`.
    rv_xml = rv_xml && `tal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Pos`.
    rv_xml = rv_xml && `ition="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:`.
    rv_xml = rv_xml && `LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854427048">&#10;`.
    rv_xml = rv_xml && `   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/`.
    rv_xml = rv_xml && `>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Bord`.
    rv_xml = rv_xml && `er ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:I`.
    rv_xml = rv_xml && `D="m2481854427068">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Co`.
    rv_xml = rv_xml && `ntinuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:We`.
    rv_xml = rv_xml && `ight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </`.
    rv_xml = rv_xml && `Style>&#10;  <Style ss:ID="m2481854427088">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="`.
    rv_xml = rv_xml && `Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:Line`.
    rv_xml = rv_xml && `Style="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman`.
    rv_xml = rv_xml && `" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854427108">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10; `.
    rv_xml = rv_xml && `   <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:`.
    rv_xml = rv_xml && `Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:F`.
    rv_xml = rv_xml && `ontName="Times New Roman" ss:Size="11" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425344">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:Wr`.
    rv_xml = rv_xml && `apText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" `.
    rv_xml = rv_xml && `ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&`.
    rv_xml = rv_xml && `#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425364">&#10;   <Alignment ss:Horizontal="Left" ss:`.
    rv_xml = rv_xml && `Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:`.
    rv_xml = rv_xml && `LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Conti`.
    rv_xml = rv_xml && `nuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425384">&#10;   <Alignment s`.
    rv_xml = rv_xml && `s:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border`.
    rv_xml = rv_xml && ` ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="T`.
    rv_xml = rv_xml && `op" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m248185442540`.
    rv_xml = rv_xml && `4">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weigh`.
    rv_xml = rv_xml && `t="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   `.
    rv_xml = rv_xml && ` <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Sty`.
    rv_xml = rv_xml && `le ss:ID="m2481854425444">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineSt`.
    rv_xml = rv_xml && `yle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous`.
    rv_xml = rv_xml && `" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#`.
    rv_xml = rv_xml && `10;  </Style>&#10;  <Style ss:ID="m2481854425464">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Pos`.
    rv_xml = rv_xml && `ition="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" `.
    rv_xml = rv_xml && `ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times Ne`.
    rv_xml = rv_xml && `w Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425484">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders`.
    rv_xml = rv_xml && `>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Bor`.
    rv_xml = rv_xml && `der ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Fo`.
    rv_xml = rv_xml && `nt ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425504">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapTe`.
    rv_xml = rv_xml && `xt="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:W`.
    rv_xml = rv_xml && `eight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;`.
    rv_xml = rv_xml && `   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425524">&#10;   <Alignment ss:Horizontal="Center" ss:Ver`.
    rv_xml = rv_xml && `tical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:Lin`.
    rv_xml = rv_xml && `eStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuo`.
    rv_xml = rv_xml && `us" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854425988">&#10;   <Alignment ss:Ho`.
    rv_xml = rv_xml && `rizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Po`.
    rv_xml = rv_xml && `sition="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:Fon`.
    rv_xml = rv_xml && `tName="Times New Roman" ss:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854426028">&#10;   <Alignment ss:Horizontal="Left" ss:`.
    rv_xml = rv_xml && `Vertical="Top" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:Line`.
    rv_xml = rv_xml && `Style="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854426048">&#10;   `.
    rv_xml = rv_xml && `<Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;  `.
    rv_xml = rv_xml && `  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <S`.
    rv_xml = rv_xml && `tyle ss:ID="m2481854426068">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineSt`.
    rv_xml = rv_xml && `yle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous`.
    rv_xml = rv_xml && `" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss`.
    rv_xml = rv_xml && `:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854426088">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10`.
    rv_xml = rv_xml && `;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&`.
    rv_xml = rv_xml && `#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders`.
    rv_xml = rv_xml && `>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854426108">&#10;   <Alignment ss:`.
    rv_xml = rv_xml && `Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border s`.
    rv_xml = rv_xml && `s:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top`.
    rv_xml = rv_xml && `" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style`.
    rv_xml = rv_xml && `>&#10;  <Style ss:ID="m2481854426128">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom"`.
    rv_xml = rv_xml && ` ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="`.
    rv_xml = rv_xml && `Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Si`.
    rv_xml = rv_xml && `ze="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854418936">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText`.
    rv_xml = rv_xml && `="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weig`.
    rv_xml = rv_xml && `ht="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"`.
    rv_xml = rv_xml && `    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854419036">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Bord`.
    rv_xml = rv_xml && `ers>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </B`.
    rv_xml = rv_xml && `orders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="m2481854419056">&#10;   <Alignme`.
    rv_xml = rv_xml && `nt ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border `.
    rv_xml = rv_xml && `ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:I`.
    rv_xml = rv_xml && `D="s63">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="20" ss:Bold="1"/>&#10;  </Sty`.
    rv_xml = rv_xml && `le>&#10;  <Style ss:ID="s65">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineSty`.
    rv_xml = rv_xml && `le="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" s`.
    rv_xml = rv_xml && `s:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="s73">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&`.
    rv_xml = rv_xml && `#10;   <Borders>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&`.
    rv_xml = rv_xml && `#10;  </Style>&#10;  <Style ss:ID="s76">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Left" ss`.
    rv_xml = rv_xml && `:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s79">&#10;   <Align`.
    rv_xml = rv_xml && `ment ss:Horizontal="Right" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Bor`.
    rv_xml = rv_xml && `ders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s90">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:Wr`.
    rv_xml = rv_xml && `apText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" `.
    rv_xml = rv_xml && `ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&`.
    rv_xml = rv_xml && `#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"    ss:Underline="Single"/>&#10;  </Style>&#10;  <Style ss:ID="s95">&#10;   <Alignmen`.
    rv_xml = rv_xml && `t ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Bor`.
    rv_xml = rv_xml && `der ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position`.
    rv_xml = rv_xml && `="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s103">&#10`.
    rv_xml = rv_xml && `;   <Alignment ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s104">&#10;   <Alignmen`.
    rv_xml = rv_xml && `t ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:Fo`.
    rv_xml = rv_xml && `ntName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s105">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <`.
    rv_xml = rv_xml && `Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10; `.
    rv_xml = rv_xml && `   <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10`.
    rv_xml = rv_xml && `;   <Font ss:FontName="Times New Roman" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="s115">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText=`.
    rv_xml = rv_xml && `"1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weig`.
    rv_xml = rv_xml && `ht="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   `.
    rv_xml = rv_xml && `</Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style ss:ID="s116">&#10;   <Alignment ss:Horizontal="Right" ss:Vertical="Center`.
    rv_xml = rv_xml && `" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Conti`.
    rv_xml = rv_xml && `nuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight`.
    rv_xml = rv_xml && `="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;   <NumberFormat ss:Format="#,##0"/>&#10;  </Style>&#10;  <Style ss:ID="s117">&#10;`.
    rv_xml = rv_xml && `   <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>`.
    rv_xml = rv_xml && `&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Borde`.
    rv_xml = rv_xml && `r ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;   <NumberFormat ss:Format="`.
    rv_xml = rv_xml && `Standard"/>&#10;  </Style>&#10;  <Style ss:ID="s126">&#10;   <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:P`.
    rv_xml = rv_xml && `osition="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right`.
    rv_xml = rv_xml && `" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times `.
    rv_xml = rv_xml && `New Roman" ss:Size="12" ss:Bold="1"/>&#10;   <NumberFormat ss:Format="#,##0"/>&#10;  </Style>&#10;  <Style ss:ID="s127">&#10;   <Alignment ss:Horizontal="Right" ss:Vertic`.
    rv_xml = rv_xml && `al="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineSt`.
    rv_xml = rv_xml && `yle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous"`.
    rv_xml = rv_xml && ` ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12" ss:Bold="1"/>&#10;   <NumberFormat ss:Format="Standard"/>&#10;  </Style>&#10;  `.
    rv_xml = rv_xml && `<Style ss:ID="s128">&#10;   <Alignment ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1`.
    rv_xml = rv_xml && `"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Bo`.
    rv_xml = rv_xml && `rder ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontName="Times New Roman" ss:Size="12"/>&#10;  </Style>&#10;  <Style s`.
    rv_xml = rv_xml && `s:ID="s130">&#10;   <Alignment ss:Horizontal="Right" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="11"/>&#10;  </Style>&#10;  <S`.
    rv_xml = rv_xml && `tyle ss:ID="s132">&#10;   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="11"/>&#10;  </Style>&#10`.
    rv_xml = rv_xml && `;  <Style ss:ID="s134">&#10;   <Alignment ss:Horizontal="Right" ss:Vertical="Top" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Size="11" ss:Bold="1"/>&`.
    rv_xml = rv_xml && `#10;  </Style>&#10;  <Style ss:ID="s136">&#10;   <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Font ss:FontName="Times New Roman" ss:Siz`.
    rv_xml = rv_xml && `e="14" ss:Bold="1"/>&#10;  </Style>&#10;  <Style ss:ID="s151">&#10;   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>&#10;   <Borders>&#10;    <B`.
    rv_xml = rv_xml && `order ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Posit`.
    rv_xml = rv_xml && `ion="Right" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>&#10;   </Borders>&#10;   <Font ss:FontNa`.
    rv_xml = rv_xml && `me="Times New Roman" x:Family="Roman" ss:Size="12"/>&#10;  </Style>&#10; </Styles>&#10; <Worksheet ss:Name="Packing List">&#10;  <Table x:FullColumns="1"   x:FullRows="1"`.
    rv_xml = rv_xml && ` ss:DefaultColumnWidth="52.8"   ss:DefaultRowHeight="15.600000000000001">&#10;   <Column ss:AutoFitWidth="0" ss:Width="22.799999999999997"/>&#10;   <Column ss:AutoFitWidt`.
    rv_xml = rv_xml && `h="0" ss:Width="91.199999999999989"/>&#10;   <Column ss:AutoFitWidth="0" ss:Width="124.80000000000001"/>&#10;   <Column ss:AutoFitWidth="0" ss:Width="67.8"/>&#10;   <Colu`.
    rv_xml = rv_xml && `mn ss:AutoFitWidth="0" ss:Width="124.80000000000001"/>&#10;   <Column ss:AutoFitWidth="0" ss:Width="67.8"/>&#10;   <Column ss:AutoFitWidth="0" ss:Width="79.80000000000001`.
    rv_xml = rv_xml && `1" ss:Span="1"/>&#10;   <Column ss:Index="9" ss:AutoFitWidth="0" ss:Width="67.8" ss:Span="1"/>&#10;   <Column ss:Index="11" ss:AutoFitWidth="0" ss:Width="171"/>&#10;   <R`.
    rv_xml = rv_xml && `ow ss:AutoFitHeight="0" ss:Height="28.05">&#10;    <Cell ss:MergeAcross="10" ss:StyleID="s63"><Data ss:Type="String">PACKING LIST</Data></Cell>&#10;   </Row>&#10;  <Row s`.
    rv_xml = rv_xml && `s:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:MergeAcross="5" ss:StyleID="s65"><Data ss:Type="String">Shipper / Exporter</Data></Cell>&#10;    <Cell ss:MergeAcross`.
    rv_xml = rv_xml && `="4" ss:StyleID="m2481854418936"><Data ss:Type="String">No &amp; Date of Commercial Invoice</Data></Cell>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="18">&#1`.
    rv_xml = rv_xml && `0;    <Cell ss:MergeAcross="5" ss:StyleID="s73"><Data ss:Type="String">{{shipper.name}}</Data></Cell>&#10;    <Cell ss:MergeAcross="3" ss:MergeDown="1" ss:StyleID="s76"><`.
    rv_xml = rv_xml && `Data ss:Type="String">{{invoice_no}}</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="s79"><Data ss:Type="String">{{invoice_date}}</Data></Cell>&#10;   </Row>&#10`.
    rv_xml = rv_xml && `;  <Row>&#10;    <Cell ss:MergeAcross="5" ss:MergeDown="3" ss:StyleID="s76"><Data      ss:Type="String">{{shipper.address}}</Data></Cell>&#10;   </Row>&#10;  <Row ss:Auto`.
    rv_xml = rv_xml && `FitHeight="0" ss:Height="18">&#10;    <Cell ss:Index="7" ss:MergeAcross="4" ss:StyleID="m2481854419036"><Data      ss:Type="String">Term of Payment</Data></Cell>&#10;   <`.
    rv_xml = rv_xml && `/Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:Index="7" ss:MergeAcross="4" ss:MergeDown="1"     ss:StyleID="m2481854419056"><Data ss:Type="String`.
    rv_xml = rv_xml && `">{{payment_term}}</Data></Cell>&#10;   </Row>&#10;  <Row/>&#10;  <Row ss:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:MergeAcross="5" ss:StyleID="s65"><Data ss:Typ`.
    rv_xml = rv_xml && `e="String">Buyer / Importer</Data></Cell>&#10;    <Cell ss:MergeAcross="4" ss:StyleID="m2481854425988"><Data ss:Type="String">Remarks</Data></Cell>&#10;   </Row>&#10;  <R`.
    rv_xml = rv_xml && `ow ss:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:MergeAcross="5" ss:StyleID="s73"><Data ss:Type="String">{{buyer.name}}</Data></Cell>&#10;    <Cell ss:MergeAcross`.
    rv_xml = rv_xml && `="4" ss:MergeDown="3" ss:StyleID="m2481854426028"><Data ss:Type="String">{{remarks}}</Data></Cell>&#10;   </Row>&#10;  <Row>&#10;    <Cell ss:MergeAcross="5" ss:MergeDown`.
    rv_xml = rv_xml && `="2" ss:StyleID="m2481854426048"><Data      ss:Type="String">{{buyer.address}}</Data></Cell>&#10;   </Row>&#10;  <Row/>&#10;  <Row/>&#10;  <Row ss:AutoFitHeight="0" ss:He`.
    rv_xml = rv_xml && `ight="18">&#10;    <Cell ss:MergeAcross="3" ss:StyleID="m2481854426068"><Data ss:Type="String">Port of Shipment</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID=`.
    rv_xml = rv_xml && `"m2481854426088"><Data ss:Type="String">Port of Discharge:</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854426108"><Data ss:Type="String">Vessels' Name`.
    rv_xml = rv_xml && `s</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854426128"><Data ss:Type="String">Trunk Vessel</Data></Cell>&#10;    <Cell ss:StyleID="s90"><Data ss:Typ`.
    rv_xml = rv_xml && `e="String">Estimated of Delivery</Data></Cell>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:MergeAcross="3" ss:StyleID="m2481854425344">`.
    rv_xml = rv_xml && `<Data ss:Type="String">{{port_of_shipment}}</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854425364"><Data ss:Type="String">{{port_of_discharge}}</Data>`.
    rv_xml = rv_xml && `</Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854425384"><Data ss:Type="String">{{vessel}}</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854`.
    rv_xml = rv_xml && `425404"><Data ss:Type="String">{{trunk_vessel}}</Data></Cell>&#10;    <Cell ss:StyleID="s95"><Data ss:Type="String">{{etd}}</Data></Cell>&#10;   </Row>&#10;  <Row ss:Auto`.
    rv_xml = rv_xml && `FitHeight="0" ss:Height="19.950000000000003">&#10;    <Cell ss:MergeAcross="10" ss:StyleID="s104"/>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="19.9500000000`.
    rv_xml = rv_xml && `00003">&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854425444"><Data ss:Type="String">No</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854425464"><Dat`.
    rv_xml = rv_xml && `a ss:Type="String">SO/ SO Item</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854425484"><Data ss:Type="String">Item No</Data></Cell>&#10;    <Cell ss:Merg`.
    rv_xml = rv_xml && `eDown="1" ss:StyleID="m2481854425504"><Data ss:Type="String">Outbound&#10;Delivery No.</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854425524"><Data ss:T`.
    rv_xml = rv_xml && `ype="String">Commodity</Data></Cell>&#10;    <Cell ss:MergeAcross="1" ss:StyleID="m2481854427008"><Data ss:Type="String">Quantity</Data></Cell>&#10;    <Cell ss:MergeDown`.
    rv_xml = rv_xml && `="1" ss:StyleID="m2481854427028"><Data ss:Type="String">PALLET&#10;( PCS)</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854427048"><Data ss:Type="String">`.
    rv_xml = rv_xml && `G.weight&#10;(kgs)</Data></Cell>&#10;    <Cell ss:MergeDown="1" ss:StyleID="m2481854427068"><Data ss:Type="String">N.weight&#10;(kgs)</Data></Cell>&#10;    <Cell ss:Merge`.
    rv_xml = rv_xml && `Down="1" ss:StyleID="m2481854427088"><Data ss:Type="String">CONT/SEAL.NO.</Data></Cell>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="19.950000000000003">&#10;`.
    rv_xml = rv_xml && `    <Cell ss:Index="6" ss:StyleID="s105"><Data ss:Type="String">(PCS)</Data></Cell>&#10;    <Cell ss:StyleID="s105"><Data ss:Type="String">(CTNS)</Data></Cell>&#10;   </R`.
    rv_xml = rv_xml && `ow>&#10;  <Row ss:AutoFitHeight="0" ss:Height="28.05">&#10;    <Cell ss:StyleID="s115"><Data ss:Type="String">{{#items}}{{@index}}</Data></Cell>&#10;    <Cell ss:StyleID=`.
    rv_xml = rv_xml && `"s115"><Data ss:Type="String">{{so_item}}</Data></Cell>&#10;    <Cell ss:StyleID="s95"><Data ss:Type="String">{{item_no}}</Data></Cell>&#10;    <Cell ss:StyleID="s115"><D`.
    rv_xml = rv_xml && `ata ss:Type="String">{{*mergesame:delivery_no}}{{delivery_no}}</Data></Cell>&#10;    <Cell ss:StyleID="s95"><Data ss:Type="String">{{*mergesame:commodity}}{{commodity}}</Data></Cell>&#10;    <Cell ss:StyleID="s116"><Data`.
    rv_xml = rv_xml && ` ss:Type="String">{{qty_pcs}}</Data></Cell>&#10;    <Cell ss:StyleID="s116"><Data ss:Type="String">{{qty_ctns}}</Data></Cell>&#10;    <Cell ss:StyleID="s116"><Data ss:Typ`.
    rv_xml = rv_xml && `e="String">{{pallet}}</Data></Cell>&#10;    <Cell ss:StyleID="s117"><Data ss:Type="String">{{gross_weight}}</Data></Cell>&#10;    <Cell ss:StyleID="s117"><Data ss:Type="S`.
    rv_xml = rv_xml && `tring">{{net_weight}}</Data></Cell>&#10;    <Cell ss:StyleID="m2481854426404"><Data ss:Type="String">{{*mergesame:container}}{{container}}{{/items}}</Data></Cell>&#10;   </Row>&#10;  <Row ss:Aut`.
    rv_xml = rv_xml && `oFitHeight="0" ss:Height="19.950000000000003">&#10;    <Cell ss:MergeAcross="4" ss:StyleID="m2481854427108"><Data ss:Type="String">TOTAL</Data></Cell>&#10;    <Cell ss:St`.
    rv_xml = rv_xml && `yleID="s126"><Data ss:Type="String">{{sum:items.qty_pcs}}</Data></Cell>&#10;    <Cell ss:StyleID="s126"><Data ss:Type="String">{{sum:items.qty_ctns}}</Data></Cell>&#10;  `.
    rv_xml = rv_xml && `  <Cell ss:StyleID="s126"><Data ss:Type="String">{{sum:items.pallet}}</Data></Cell>&#10;    <Cell ss:StyleID="s127"><Data ss:Type="String">{{sum:items.gross_weight}}</Dat`.
    rv_xml = rv_xml && `a></Cell>&#10;    <Cell ss:StyleID="s127"><Data ss:Type="String">{{sum:items.net_weight}}</Data></Cell>&#10;    <Cell ss:StyleID="s128"/>&#10;   </Row>&#10;  <Row ss:Auto`.
    rv_xml = rv_xml && `FitHeight="0" ss:Height="10.050000000000001">&#10;    <Cell ss:MergeAcross="10" ss:StyleID="s104"/>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="15">&#10;    `.
    rv_xml = rv_xml && `<Cell ss:MergeAcross="5" ss:StyleID="s130"><Data ss:Type="String">TYPE OF CONTAINER:</Data></Cell>&#10;    <Cell ss:MergeAcross="4" ss:StyleID="s132"><Data ss:Type="Strin`.
    rv_xml = rv_xml && `g">{{container_summary}}</Data></Cell>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height="18">&#10;    <Cell ss:MergeAcross="5" ss:StyleID="s134"><Data ss:Type="Str`.
    rv_xml = rv_xml && `ing">TOTAL NUMBER OF PACKAGES:</Data></Cell>&#10;    <Cell ss:MergeAcross="4" ss:StyleID="s132"><Data ss:Type="String">{{package_summary}}</Data></Cell>&#10;   </Row>&#10`.
    rv_xml = rv_xml && `;  <Row ss:AutoFitHeight="0" ss:Height="10.050000000000001">&#10;    <Cell ss:MergeAcross="10" ss:StyleID="s103"/>&#10;   </Row>&#10;  <Row ss:AutoFitHeight="0" ss:Height`.
    rv_xml = rv_xml && `="24">&#10;    <Cell ss:MergeAcross="10" ss:StyleID="s136"><Data ss:Type="String">{{shipper.name}}</Data></Cell>&#10;   </Row>&#10;  </Table>&#10;  <WorksheetOptions xmln`.
    rv_xml = rv_xml && `s="urn:schemas-microsoft-com:office:excel">&#10;   <PageSetup>&#10;    <Layout x:CenterHorizontal="1"/>&#10;    <Header x:Margin="0.1"/>&#10;    <Footer x:Margin="0.1"/>&`.
    rv_xml = rv_xml && `#10;    <PageMargins x:Bottom="0.5" x:Left="0.5" x:Right="0.5" x:Top="0.5"/>&#10;   </PageSetup>&#10;   <FitToPage/>&#10;   <Print>&#10;    <FitHeight>999</FitHeight>&#10`.
    rv_xml = rv_xml && `;    <ValidPrinterInfo/>&#10;    <PaperSizeIndex>9</PaperSizeIndex>&#10;    <VerticalResolution>0</VerticalResolution>&#10;   </Print>&#10;   <Selected/>&#10;   <DoNotDis`.
    rv_xml = rv_xml && `playGridlines/>&#10;   <Panes>&#10;    <Pane>&#10;     <Number>3</Number>&#10;     <ActiveRow>3</ActiveRow>&#10;     <RangeSelection>R4C1:R7C6</RangeSelection>&#10;    </`.
    rv_xml = rv_xml && `Pane>&#10;   </Panes>&#10;   <ProtectObjects>False</ProtectObjects>&#10;   <ProtectScenarios>False</ProtectScenarios>&#10;  </WorksheetOptions>&#10; </Worksheet>&#10;</Wo`.
    rv_xml = rv_xml && `rkbook>`.
  ENDMETHOD.


  METHOD sample_context.
    rs_ctx = VALUE ty_ctx(
      shipper = VALUE #( name    = `CASABLANCA JOINT STOCK COMPANY`
                         address = `NON SAO INDUSTRIAL ZONE, DINH TAN, TAN DINH, BAC NINH PROVINCE, VIET NAM` )
      buyer   = VALUE #( name    = `CASABLANCA HONGKONG LIMITED`
                         address = `UNIT S FLAT A-C 25/F, SEABRIGHT PLAZA 9-23 SHELL ST,, NORTH POINT, Hong Kong` )
      invoice_no        = `CAS-2025-0645`
      invoice_date      = `16-Sep-2025`
      payment_term      = `T/T 30 days`
      remarks           = ``
      port_of_shipment  = `FOB Haiphong, VIETNAM`
      port_of_discharge = `FOB UK`
      vessel            = ``
      trunk_vessel      = ``
      etd               = ``
      container_summary = `1x 40HC CONTAINER`
      package_summary   = `6170 CTNS`
      items = VALUE #(
        ( so_item = `10000124/10` item_no = `TTR-1-NC20081736-2503` delivery_no = `80000188`
          commodity = `RPET SHOPPING BAG` qty_pcs = 1234 qty_ctns = 0 pallet = 0
          gross_weight = 0 net_weight = 0 container = `ICT-110234N-2` )
        ( so_item = `10000124/20` item_no = `TTR-1-NC20081222-2506` delivery_no = `80000188`
          commodity = `RPET SHOPPING BAG` qty_pcs = 1234 qty_ctns = 0 pallet = 0
          gross_weight = 0 net_weight = 0 container = `ICT-110234N-2` )
        ( so_item = `10000125/10` item_no = `TTR-1-NC20081711-2513` delivery_no = `80000188`
          commodity = `RPET SHOPPING BAG` qty_pcs = 1234 qty_ctns = 0 pallet = 0
          gross_weight = 0 net_weight = 0 container = `IVF-HJ0001-7` )
        ( so_item = `10000125/20` item_no = `TTR-1-NC20081443-2503` delivery_no = `80000188`
          commodity = `RPET SHOPPING BAG` qty_pcs = 1234 qty_ctns = 0 pallet = 0
          gross_weight = 0 net_weight = 0 container = `ICT-34556LK-3` )
        ( so_item = `10000126/10` item_no = `TTR-1-NC20081645-2511` delivery_no = `80000188`
          commodity = `RPET SHOPPING BAG` qty_pcs = 1234 qty_ctns = 0 pallet = 0
          gross_weight = 0 net_weight = 0 container = `ICT-34556LK-3` ) ) ).
  ENDMETHOD.


  METHOD if_oo_adt_classrun~main.
    " Nạp template vào bảng ZXLWB_TMPL (chạy F9 trong Eclipse)
    DATA ls_row TYPE zxlwb_tmpl.

    ls_row-form_name = c_form.
    ls_row-descr     = 'Packing List HQ (tu file Excel that)'.
    ls_row-engine    = 'SSML'.
    ls_row-mime_type = 'application/vnd.ms-excel'.
    ls_row-file_name = 'PackingListHQ.xls'.
    ls_row-is_active = abap_true.
    ls_row-template  = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) ).
    ls_row-created_by       = cl_abap_context_info=>get_user_technical_name( ).
    GET TIME STAMP FIELD ls_row-created_at.
    ls_row-last_changed_by  = ls_row-created_by.
    ls_row-last_changed_at  = ls_row-created_at.
    ls_row-last_changed_glo = ls_row-created_at.

    MODIFY zxlwb_tmpl FROM @ls_row.
    IF sy-subrc = 0.
      COMMIT WORK AND WAIT.
      out->write( |Da nap template '{ c_form }' vao ZXLWB_TMPL | &&
                  |({ xstrlen( ls_row-template ) } byte).| ).
      out->write( |Gio co the goi: zcl_xlwb_runtime=>render( iv_form_name = '{ c_form }' ... )| ).
    ELSE.
      ROLLBACK WORK.
      out->write( |Nap template that bai, sy-subrc = { sy-subrc }| ).
    ENDIF.

    " In thu ket qua render de kiem tra nhanh
    TRY.
        DATA(ls_ctx) = sample_context( ).
        DATA(ls_file) = zcl_xlwb_runtime=>render_with_template(
                          iv_template = cl_abap_conv_codepage=>create_out( )->convert( get_template( ) )
                          iv_engine   = `SSML`
                          ir_context  = REF #( ls_ctx ) ).
        out->write( |--- Noi dung file render ({ xstrlen( ls_file-content ) } byte) ---| ).
        out->write( cl_abap_conv_codepage=>create_in( )->convert( ls_file-content ) ).
      CATCH zcx_xlwb INTO DATA(lx).
        out->write( |Loi render: { lx->get_text( ) }| ).
    ENDTRY.
  ENDMETHOD.
ENDCLASS.
