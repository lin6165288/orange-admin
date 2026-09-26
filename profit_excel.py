from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def _cell_ref(col: int, row: int) -> str:
    letters = ""
    n = col
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _inline_cell(ref: str, value, style: int = 0) -> str:
    if value is None:
        value = ""
    if isinstance(value, (date, datetime)):
        value = value.strftime("%Y-%m-%d")
    text = escape(str(value))
    preserve = ' xml:space="preserve"' if str(value).startswith(" ") or str(value).endswith(" ") else ""
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t{preserve}>{text}</t></is></c>'


def _number_cell(ref: str, value, style: int = 0) -> str:
    if value is None or value == "":
        return _inline_cell(ref, "", style)
    try:
        num = Decimal(str(value))
    except Exception:
        return _inline_cell(ref, value, style)
    return f'<c r="{ref}" s="{style}"><v>{num}</v></c>'


def make_profit_xlsx(rows: list[dict], start_date: str, end_date: str,
                     rmb_rate, payment_sell_rate, purchase_sell_rate) -> bytes:
    headers = [
        "訂單編號", "下單日期", "客戶姓名", "訂單類型", "平台", "物流單號",
        "人民幣金額", "人民幣匯率", "適用定價匯率", "手續費收入",
        "匯率價差利潤", "總利潤", "重量(kg)", "是否到貨", "是否已運回",
        "提前運回", "訂單狀態", "備註"
    ]

    rows_xml = []
    title = f"橘貓代購利潤報表｜{start_date} ～ {end_date}"
    rows_xml.append('<row r="1" ht="28" customHeight="1">' + _inline_cell("A1", title, 3) + '</row>')
    rate_text = f"人民幣匯率：{rmb_rate}｜代付定價匯率：{payment_sell_rate}｜代購定價匯率：{purchase_sell_rate}"
    rows_xml.append('<row r="2">' + _inline_cell("A2", rate_text, 0) + '</row>')
    header_cells = ''.join(_inline_cell(_cell_ref(i + 1, 4), h, 1) for i, h in enumerate(headers))
    rows_xml.append(f'<row r="4" ht="24" customHeight="1">{header_cells}</row>')

    numeric_cols = {1, 7, 8, 9, 10, 11, 12, 13}
    for row_idx, r in enumerate(rows, start=5):
        values = [
            r.get("order_id"), r.get("order_time"), r.get("customer_name"),
            r.get("order_type"), r.get("platform"), r.get("tracking_number"),
            r.get("amount_rmb_num"), rmb_rate, r.get("sell_rate_num"),
            r.get("service_fee_num"), r.get("rate_profit_num"), r.get("total_profit_num"),
            r.get("weight_num"), "是" if r.get("is_arrived") else "否",
            "是" if r.get("is_returned") else "否",
            "是" if r.get("is_early_returned") else "否",
            r.get("order_status") or "正常", r.get("remarks") or ""
        ]
        cells = []
        for col_idx, value in enumerate(values, start=1):
            ref = _cell_ref(col_idx, row_idx)
            cells.append(_number_cell(ref, value, 2) if col_idx in numeric_cols else _inline_cell(ref, value, 0))
        rows_xml.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    total_row = 5 + len(rows)
    total_rate_profit = sum(Decimal(str(r.get("rate_profit_num") or 0)) for r in rows)
    total_service_fee = sum(Decimal(str(r.get("service_fee_num") or 0)) for r in rows)
    total_profit = sum(Decimal(str(r.get("total_profit_num") or 0)) for r in rows)
    summary_cells = [
        _inline_cell(f"A{total_row}", "合計", 4),
        _inline_cell(f"I{total_row}", "匯率價差", 4),
        _number_cell(f"J{total_row}", total_rate_profit, 5),
        _inline_cell(f"K{total_row}", "手續費", 4),
        _number_cell(f"L{total_row}", total_service_fee, 5),
        _inline_cell(f"M{total_row}", "總利潤", 4),
        _number_cell(f"N{total_row}", total_profit, 5),
    ]
    rows_xml.append(f'<row r="{total_row}" ht="24" customHeight="1">{"".join(summary_cells)}</row>')

    widths = [12,14,18,12,12,22,14,14,16,14,16,14,12,12,12,12,12,28]
    cols_xml = ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, start=1))
    last_ref = _cell_ref(len(headers), total_row)
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{last_ref}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18"/>'
        f'<cols>{cols_xml}</cols>'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        '<mergeCells count="2"><mergeCell ref="A1:R1"/><mergeCell ref="A2:R2"/></mergeCells>'
        f'<autoFilter ref="A4:R{max(4, total_row - 1)}"/>'
        '</worksheet>'
    )

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Microsoft JhengHei"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft JhengHei"/></font>
    <font><b/><sz val="14"/><color rgb="FFF08300"/><name val="Microsoft JhengHei"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF08300"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF3E5"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFE6E6E6"/></left><right style="thin"><color rgb="FFE6E6E6"/></right><top style="thin"><color rgb="FFE6E6E6"/></top><bottom style="thin"><color rgb="FFE6E6E6"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="4" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="利潤報表" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''

    out = BytesIO()
    with ZipFile(out, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/styles.xml", styles_xml)
    return out.getvalue()
