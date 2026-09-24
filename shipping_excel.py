"""Self-contained Excel export of shipping groups and details.

Uses only the Python standard library to emit Office Open XML.  All user-supplied
strings are written as inlineStr (never treated as Excel formulas).
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", MAIN)
ET.register_namespace("r", OFFICE)


def q(local: str) -> str:
    return "{" + MAIN + "}" + local


def col_letters(number: int) -> str:
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def text(value):
    if value is None:
        return "—"
    if isinstance(value, (date, datetime)):
        return value.isoformat()[:10]
    return str(value)


def add_cell(row, index: int, number: int, value, style: int = 0):
    ref = f"{col_letters(index)}{number}"
    if isinstance(value, tuple) and len(value) == 3 and value[0] == "formula":
        cell = ET.SubElement(row, q("c"), r=ref, s=str(style))
        ET.SubElement(cell, q("f")).text = value[1][1:] if value[1].startswith("=") else value[1]
        ET.SubElement(cell, q("v")).text = str(value[2])
    elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        cell = ET.SubElement(row, q("c"), r=ref, s=str(style))
        ET.SubElement(cell, q("v")).text = str(value)
    else:
        cell = ET.SubElement(row, q("c"), r=ref, s=str(style), t="inlineStr")
        inline = ET.SubElement(cell, q("is"))
        ET.SubElement(inline, q("t")).text = text(value)
    return cell


def worksheet(headers, records, widths, *, summary=False, total=None):
    ws = ET.Element(q("worksheet"))
    views = ET.SubElement(ws, q("sheetViews"))
    view = ET.SubElement(views, q("sheetView"), workbookViewId="0")
    ET.SubElement(view, q("pane"), ySplit="1", topLeftCell="A2", activePane="bottomLeft", state="frozen")
    cols = ET.SubElement(ws, q("cols"))
    for i, width in enumerate(widths, start=1):
        ET.SubElement(cols, q("col"), min=str(i), max=str(i), width=str(width), customWidth="1")
    data = ET.SubElement(ws, q("sheetData"))
    header = ET.SubElement(data, q("row"), r="1", ht="29", customHeight="1")
    for i, value in enumerate(headers, 1):
        add_cell(header, i, 1, value, 1)
    numeric = set(range(2, 13)) if summary else {1, 7}
    for row_i, values in enumerate(records, 2):
        row = ET.SubElement(data, q("row"), r=str(row_i), ht="23", customHeight="1")
        for i, value in enumerate(values, 1):
            is_number = i in numeric
            style = 3 if is_number and (i == 1 and not summary or i in (2,4,8) and summary) else 2 if is_number else 0
            add_cell(row, i, row_i, value, style)
    if total is not None:
        i = len(records) + 2
        row = ET.SubElement(data, q("row"), r=str(i), ht="26", customHeight="1")
        for column, value in enumerate(total, 1):
            add_cell(row, column, i, value, 4 if column == 1 else 5)
    last_row = 1 + len(records)
    ET.SubElement(ws, q("autoFilter"), ref=f"A1:{col_letters(len(headers))}{last_row}")
    ET.SubElement(ws, q("pageMargins"), left="0.25", right="0.25", top="0.45", bottom="0.45", header="0.2", footer="0.2")
    return ET.tostring(ws, xml_declaration=True, encoding="utf-8")


def style_xml():
    root = ET.Element(q("styleSheet"))
    formats = ET.SubElement(root, q("numFmts"), count="1")
    ET.SubElement(formats, q("numFmt"), numFmtId="164", formatCode='#,##0.00')
    fonts = ET.SubElement(root, q("fonts"), count="2")
    for bold in (False, True):
        font = ET.SubElement(fonts, q("font"))
        ET.SubElement(font, q("sz"), val="11")
        ET.SubElement(font, q("name"), val="Microsoft JhengHei")
        if bold:
            ET.SubElement(font, q("b"))
            ET.SubElement(font, q("color"), rgb="FFFFFFFF")
    fills = ET.SubElement(root, q("fills"), count="3")
    ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="none")
    ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="gray125")
    colored = ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="solid")
    ET.SubElement(colored, q("fgColor"), rgb="FFB96B14")
    ET.SubElement(colored, q("bgColor"), indexed="64")
    borders = ET.SubElement(root, q("borders"), count="1")
    ET.SubElement(borders, q("border"))
    ET.SubElement(root, q("cellStyleXfs"), count="1").append(ET.Element(q("xf"), numFmtId="0", fontId="0", fillId="0", borderId="0"))
    xfs = ET.SubElement(root, q("cellXfs"), count="6")
    configs = [
        (0, 0, 0),   # normal
        (0, 1, 2),   # orange header
        (164, 0, 0), # decimals
        (1, 0, 0),   # integer
        (0, 0, 2),   # total label
        (164, 0, 2), # total value
    ]
    for num, font, fill in configs:
        xf = ET.SubElement(xfs, q("xf"), numFmtId=str(num), fontId=str(font), fillId=str(fill), borderId="0", xfId="0", applyNumberFormat="1" if num else "0")
        ET.SubElement(xf, q("alignment"), vertical="center", wrapText="1")
    return ET.tostring(root, xml_declaration=True, encoding="utf-8")


def make_shipping_xlsx(groups, orders, *, kind="summary") -> bytes:
    """Summary: one sheet; details: detail rows plus per-customer fee breakdown.

    groups must already be calculated using the shipping fee rules and EXACTLY
    the same order selection as orders. No fee is attributed to a single order.
    """
    summary_headers = ["客戶姓名", "未運回件數", "總實重(kg)", "純集運件數", "純集運實重(kg)",
                       "純集運計費重(kg)", "純集運運費(NT$)", "代購件數", "代購實重(kg)",
                       "代購計費重(kg)", "代購運費(NT$)", "國際運費合計(NT$)"]
    sums = [Decimal("0")] * 11
    summary_rows = []
    for ix, g in enumerate(groups, 2):
        fields = [
            g["name"], g["count"], g["weight"], g["forwarding_count"],
            g["forwarding_weight"], g["forwarding_billed_weight"],
            ("formula", f"=F{ix}*90", g["forwarding_fee"]),
            g["purchase_count"], g["purchase_weight"], g["purchase_billed_weight"],
            ("formula", f"=J{ix}*70", g["purchase_fee"]),
            ("formula", f"=G{ix}+K{ix}", g["total_fee"]),
        ]
        summary_rows.append(fields)
        values = [g["count"], g["weight"], g["forwarding_count"], g["forwarding_weight"],
                  g["forwarding_billed_weight"], g["forwarding_fee"], g["purchase_count"],
                  g["purchase_weight"], g["purchase_billed_weight"], g["purchase_fee"], g["total_fee"]]
        sums = [a + Decimal(str(b)) for a, b in zip(sums, values)]
    end = 1 + len(groups)
    total = ["合計"] + [("formula", f"=SUM({col_letters(i)}2:{col_letters(i)}{end})", sums[i-2])
                       for i in range(2, 13)] if groups else ["合計"] + [Decimal("0")]*11
    summary_xml = worksheet(summary_headers, summary_rows,
                            [22,15,17,15,19,21,21,13,18,19,19,23], summary=True, total=total)
    sheets = [("客戶運費統整", summary_xml)]
    if kind == "details":
        headers = ["訂單編號", "下單日期", "客戶姓名", "平台", "包裹類別", "物流單號",
                   "實重(kg)", "已到貨", "提前運回", "已延後", "已通知", "備註"]
        lines = []
        for o in orders:
            remarks = str(o.get("remarks") or "")
            lines.append([o["order_id"], text(o.get("order_time")), o["customer_name"], o.get("platform") or "—",
                          "純集運" if o.get("platform") == "集運" else "代購", o.get("tracking_number") or "—",
                          Decimal(str(o.get("weight_kg") or 0)), "是" if o.get("is_arrived") else "否",
                          "是" if o.get("is_early_returned") else "否",
                          "是" if o.get("delayed") or "[延後]" in remarks else "否",
                          "是" if o.get("notified") or "[已通知]" in remarks else "否", remarks or "—"])
        detail = worksheet(headers, lines, [14,17,22,15,15,27,15,12,16,13,13,38])
        sheets = [("訂單明細", detail), ("客戶運費統整", summary_xml)]
    out = BytesIO()
    with ZipFile(out, "w", compression=ZIP_DEFLATED) as z:
        content_types = ET.Element("{" + CONTENT + "}Types")
        ET.SubElement(content_types, "{" + CONTENT + "}Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Default", Extension="xml", ContentType="application/xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName="/xl/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        for i in range(1, len(sheets)+1):
            ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName=f"/xl/worksheets/sheet{i}.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        z.writestr("[Content_Types].xml", ET.tostring(content_types, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        rels = ET.Element("{" + PACKAGE + "}Relationships")
        ET.SubElement(rels, "{" + PACKAGE + "}Relationship", Id="rId1", Type=OFFICE + "/officeDocument", Target="xl/workbook.xml")
        z.writestr("_rels/.rels", ET.tostring(rels, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        wb = ET.Element(q("workbook"))
        sheets_el = ET.SubElement(wb, q("sheets"))
        for i, (name, _xml) in enumerate(sheets, 1):
            ET.SubElement(sheets_el, q("sheet"), {"name": name, "sheetId": str(i), "{" + OFFICE + "}id": f"rId{i}"})
        ET.SubElement(wb, q("calcPr"), calcId="191029", fullCalcOnLoad="1")
        z.writestr("xl/workbook.xml", ET.tostring(wb, encoding="utf-8", xml_declaration=True))
        wb_rels = ET.Element("{" + PACKAGE + "}Relationships")
        for i in range(1, len(sheets)+1):
            ET.SubElement(wb_rels, "{" + PACKAGE + "}Relationship", Id=f"rId{i}", Type=OFFICE + "/worksheet", Target=f"worksheets/sheet{i}.xml")
        ET.SubElement(wb_rels, "{" + PACKAGE + "}Relationship", Id=f"rId{len(sheets)+1}", Type=OFFICE + "/styles", Target="styles.xml")
        z.writestr("xl/_rels/workbook.xml.rels", ET.tostring(wb_rels, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        z.writestr("xl/styles.xml", style_xml())
        for i, (_name, xml) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
    return out.getvalue()
