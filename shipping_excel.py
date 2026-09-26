"""Excel exports for the shipping page.

Two user-facing formats are generated:
1) A macro-enabled 賣貨便 single-spec import workbook based on the user's template.
   The template is copied byte-for-byte except for the first worksheet's data cells,
   so the original VBA project, formatting, validation button and workbook structure
   are preserved.
2) A compact shipping detail .xlsx containing only the fields used during packing.

The module uses the Python standard library only so Railway needs no extra Excel
package at runtime.
"""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
ET.register_namespace("", MAIN)
ET.register_namespace("r", OFFICE)
ET.register_namespace("mc", MC)
ET.register_namespace("x14ac", X14AC)
ET.register_namespace("xr", XR)
ET.register_namespace("xr2", XR2)
ET.register_namespace("xr3", XR3)

SELLER_TEMPLATE_NAME = "賣貨便_批次新增商品區.xlsm"
SELLER_FIRST_ROW = 7
SELLER_MAX_SPECS = 100


def q(local: str) -> str:
    return "{" + MAIN + "}" + local


def col_letters(number: int) -> str:
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (date, datetime)):
        return value.isoformat()[:10]
    return str(value)


def _as_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def _clear_cell(cell):
    for child in list(cell):
        cell.remove(child)
    for attr in ("t",):
        cell.attrib.pop(attr, None)


def _set_inline(cell, value):
    _clear_cell(cell)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, q("is"))
    text = ET.SubElement(inline, q("t"))
    text.text = str(value)


def _set_number(cell, value):
    _clear_cell(cell)
    ET.SubElement(cell, q("v")).text = str(value)


def _find_or_create_row(sheet_data, row_number: int):
    for row in sheet_data.findall(q("row")):
        if int(row.get("r", "0")) == row_number:
            return row
    row = ET.Element(q("row"), r=str(row_number), ht="15")
    rows = list(sheet_data.findall(q("row")))
    inserted = False
    for index, existing in enumerate(rows):
        if int(existing.get("r", "0")) > row_number:
            sheet_data.insert(index, row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    return row


def _find_or_create_cell(row, ref: str, style: int | None = None):
    for cell in row.findall(q("c")):
        if cell.get("r") == ref:
            if style is not None and "s" not in cell.attrib:
                cell.set("s", str(style))
            return cell
    cell = ET.Element(q("c"), r=ref)
    if style is not None:
        cell.set("s", str(style))
    # Keep cells in column order.
    def col_num(r):
        letters = "".join(ch for ch in r if ch.isalpha())
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch.upper()) - 64
        return n
    inserted = False
    for index, existing in enumerate(list(row.findall(q("c")))):
        if col_num(existing.get("r", "")) > col_num(ref):
            row.insert(index, cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)
    return cell


def _seller_template_path(template_path=None) -> Path:
    if template_path:
        return Path(template_path)
    return Path(__file__).resolve().parent / "excel_templates" / SELLER_TEMPLATE_NAME


def make_sellnow_xlsm(groups, return_date, template_path=None) -> bytes:
    """Populate the user's 賣貨便 single-spec macro workbook.

    One selected customer becomes one specification row:
      A7 only: M/D運回下單區
      C7 only: 橘貓代購
      H7 only: 新品
      D7:D...: customer name
      E7:E...: 1
      F7:F...: calculated international shipping fee

    All sample/spec data in rows 7:106 is cleared first. VBA and all non-target
    workbook parts are kept intact.
    """
    groups = list(groups or [])
    if not groups:
        raise ValueError("沒有可匯出的客戶")
    if len(groups) > SELLER_MAX_SPECS:
        raise ValueError(f"賣貨便單次最多可匯出 {SELLER_MAX_SPECS} 位客戶")
    ship_date = _as_date(return_date)
    template = _seller_template_path(template_path)
    if not template.exists():
        raise FileNotFoundError(f"找不到賣貨便範本：{template}")

    with ZipFile(template, "r") as source:
        sheet_xml = source.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(sheet_xml)
        sheet_data = root.find(q("sheetData"))
        if sheet_data is None:
            raise ValueError("賣貨便範本格式不正確：缺少 sheetData")

        # Style ids copied from the supplied template. Keeping these styles makes
        # newly created rows look identical to the original spec rows.
        style_by_col = {
            "A": 7, "B": 7, "C": 7, "D": 47, "E": 7, "F": 48,
            "G": 7, "H": 21, "I": 7, "J": 7, "K": 21, "L": 21,
        }
        last_row = SELLER_FIRST_ROW + SELLER_MAX_SPECS - 1
        for row_no in range(SELLER_FIRST_ROW, last_row + 1):
            row = _find_or_create_row(sheet_data, row_no)
            for col in "ABCDEFGHIJKL":
                cell = _find_or_create_cell(row, f"{col}{row_no}", style_by_col.get(col))
                _clear_cell(cell)

        # Product-level cells exist only on the first specification row.
        row7 = _find_or_create_row(sheet_data, SELLER_FIRST_ROW)
        _set_inline(_find_or_create_cell(row7, "A7", 7), f"{ship_date.month}/{ship_date.day}運回下單區")
        _set_inline(_find_or_create_cell(row7, "C7", 38), "橘貓代購")
        _set_inline(_find_or_create_cell(row7, "H7", 21), "新品")

        for offset, group in enumerate(groups):
            row_no = SELLER_FIRST_ROW + offset
            row = _find_or_create_row(sheet_data, row_no)
            _set_inline(_find_or_create_cell(row, f"D{row_no}", 47), group["name"])
            _set_number(_find_or_create_cell(row, f"E{row_no}", 7), 1)
            fee = Decimal(str(group.get("total_fee") or 0))
            _set_number(_find_or_create_cell(row, f"F{row_no}", 48), int(fee) if fee == fee.to_integral() else fee)

        # Make the used range cover all supported rows, while leaving the macro,
        # drawings, validation rules and other sheets untouched.
        dimension = root.find(q("dimension"))
        if dimension is not None:
            dimension.set("ref", f"A1:L{max(25, SELLER_FIRST_ROW + len(groups) - 1)}")
        new_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        # ElementTree only writes namespace declarations that are referenced by
        # qualified element/attribute names. xr2/xr3 are referenced only inside
        # mc:Ignorable text in this Excel template, so add their declarations
        # explicitly or Excel will repair/remove the worksheet on open.
        root_tag_end = new_sheet.find(b">", new_sheet.find(b"<worksheet"))
        root_open = new_sheet[:root_tag_end]
        root_rest = new_sheet[root_tag_end:]
        extra_ns = []
        if b"xmlns:xr2=" not in root_open:
            extra_ns.append(b' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"')
        if b"xmlns:xr3=" not in root_open:
            extra_ns.append(b' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"')
        if extra_ns:
            new_sheet = root_open + b"".join(extra_ns) + root_rest

        output = BytesIO()
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as target:
            for info in source.infolist():
                data = new_sheet if info.filename == "xl/worksheets/sheet1.xml" else source.read(info.filename)
                target.writestr(info, data)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Compact detail workbook (.xlsx)
# ---------------------------------------------------------------------------

def _add_cell(row, index: int, row_number: int, value, style: int = 0):
    ref = f"{col_letters(index)}{row_number}"
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        cell = ET.SubElement(row, q("c"), r=ref, s=str(style))
        ET.SubElement(cell, q("v")).text = str(value)
    else:
        cell = ET.SubElement(row, q("c"), r=ref, s=str(style), t="inlineStr")
        inline = ET.SubElement(cell, q("is"))
        ET.SubElement(inline, q("t")).text = _text(value)
    return cell


def _detail_records(orders, return_date):
    ship_date = _as_date(return_date)
    date_text = f"{ship_date:%Y/%m/%d}"

    # One tracking number = one package. Duplicate tracking numbers for the same
    # customer are emitted only once. A rare missing tracking number is kept as a
    # unique package by order id so the package is not silently lost.
    packages = []
    seen = set()
    for order in orders or []:
        customer = str(order.get("customer_name") or "").strip()
        tracking = str(order.get("tracking_number") or "").strip()
        order_id = order.get("order_id")
        key = (customer, tracking) if tracking else (customer, f"__order_{order_id}")
        if key in seen:
            continue
        seen.add(key)
        packages.append({
            "customer": customer,
            "tracking": tracking,
            "order_id": order_id,
        })

    counts = Counter(p["customer"] for p in packages)
    rows = []
    for package in packages:
        tracking = package["tracking"]
        last4 = tracking[-4:] if tracking else ""
        rows.append([
            package["customer"],
            "",  # 下單順序由使用者在 Excel 內自行填寫
            date_text,
            counts[package["customer"]],
            last4,
        ])
    return rows


def _detail_styles_xml():
    root = ET.Element(q("styleSheet"))
    fonts = ET.SubElement(root, q("fonts"), count="3")
    for color, bold in (("FF262626", False), ("FFFFFFFF", True), ("FF9A6700", False)):
        font = ET.SubElement(fonts, q("font"))
        ET.SubElement(font, q("sz"), val="11")
        ET.SubElement(font, q("name"), val="Microsoft JhengHei")
        ET.SubElement(font, q("color"), rgb=color)
        if bold:
            ET.SubElement(font, q("b"))
    fills = ET.SubElement(root, q("fills"), count="4")
    ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="none")
    ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="gray125")
    orange = ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="solid")
    ET.SubElement(orange, q("fgColor"), rgb="FFF29A38")
    ET.SubElement(orange, q("bgColor"), indexed="64")
    yellow = ET.SubElement(ET.SubElement(fills, q("fill")), q("patternFill"), patternType="solid")
    ET.SubElement(yellow, q("fgColor"), rgb="FFFFF4CC")
    ET.SubElement(yellow, q("bgColor"), indexed="64")
    borders = ET.SubElement(root, q("borders"), count="2")
    ET.SubElement(borders, q("border"))
    border = ET.SubElement(borders, q("border"))
    for side in ("left", "right", "top", "bottom"):
        s = ET.SubElement(border, q(side), style="thin")
        ET.SubElement(s, q("color"), rgb="FFE5E7EB")
    ET.SubElement(root, q("cellStyleXfs"), count="1").append(
        ET.Element(q("xf"), numFmtId="0", fontId="0", fillId="0", borderId="0")
    )
    xfs = ET.SubElement(root, q("cellXfs"), count="4")
    configs = [
        (0, 0, 0, 0, "left"),     # normal
        (0, 1, 2, 1, "center"),   # header
        (0, 0, 3, 1, "center"),   # manual order sequence
        (0, 0, 0, 1, "center"),   # centered body
    ]
    for num, font, fill, border_id, align in configs:
        xf = ET.SubElement(xfs, q("xf"), numFmtId=str(num), fontId=str(font), fillId=str(fill),
                           borderId=str(border_id), xfId="0")
        ET.SubElement(xf, q("alignment"), horizontal=align, vertical="center")
    return ET.tostring(root, xml_declaration=True, encoding="utf-8")


def _detail_sheet_xml(records):
    ws = ET.Element(q("worksheet"))
    views = ET.SubElement(ws, q("sheetViews"))
    view = ET.SubElement(views, q("sheetView"), workbookViewId="0")
    ET.SubElement(view, q("pane"), ySplit="1", topLeftCell="A2", activePane="bottomLeft", state="frozen")
    cols = ET.SubElement(ws, q("cols"))
    for i, width in enumerate([22, 15, 16, 14, 18], start=1):
        ET.SubElement(cols, q("col"), min=str(i), max=str(i), width=str(width), customWidth="1")
    data = ET.SubElement(ws, q("sheetData"))
    headers = ["客戶姓名", "下單順序", "運回日期", "包裹總數", "單號後四碼"]
    header = ET.SubElement(data, q("row"), r="1", ht="28", customHeight="1")
    for index, title in enumerate(headers, 1):
        _add_cell(header, index, 1, title, 1)
    for row_no, record in enumerate(records, 2):
        row = ET.SubElement(data, q("row"), r=str(row_no), ht="23", customHeight="1")
        for index, value in enumerate(record, 1):
            style = 2 if index == 2 else 3
            _add_cell(row, index, row_no, value, style)
    end = max(1, len(records) + 1)
    ET.SubElement(ws, q("autoFilter"), ref=f"A1:E{end}")
    ET.SubElement(ws, q("pageMargins"), left="0.3", right="0.3", top="0.45", bottom="0.45", header="0.2", footer="0.2")
    return ET.tostring(ws, xml_declaration=True, encoding="utf-8")


def make_shipping_detail_xlsx(orders, return_date) -> bytes:
    records = _detail_records(orders, return_date)
    if not records:
        raise ValueError("沒有可匯出的包裹")
    sheet_xml = _detail_sheet_xml(records)
    out = BytesIO()
    with ZipFile(out, "w", compression=ZIP_DEFLATED) as z:
        content_types = ET.Element("{" + CONTENT + "}Types")
        ET.SubElement(content_types, "{" + CONTENT + "}Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Default", Extension="xml", ContentType="application/xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName="/xl/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        ET.SubElement(content_types, "{" + CONTENT + "}Override", PartName="/xl/worksheets/sheet1.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        z.writestr("[Content_Types].xml", ET.tostring(content_types, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        rels = ET.Element("{" + PACKAGE + "}Relationships")
        ET.SubElement(rels, "{" + PACKAGE + "}Relationship", Id="rId1", Type=OFFICE + "/officeDocument", Target="xl/workbook.xml")
        z.writestr("_rels/.rels", ET.tostring(rels, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        workbook = ET.Element(q("workbook"))
        sheets = ET.SubElement(workbook, q("sheets"))
        ET.SubElement(sheets, q("sheet"), {"name": "運回明細表", "sheetId": "1", "{" + OFFICE + "}id": "rId1"})
        z.writestr("xl/workbook.xml", ET.tostring(workbook, encoding="utf-8", xml_declaration=True))
        wb_rels = ET.Element("{" + PACKAGE + "}Relationships")
        ET.SubElement(wb_rels, "{" + PACKAGE + "}Relationship", Id="rId1", Type=OFFICE + "/worksheet", Target="worksheets/sheet1.xml")
        ET.SubElement(wb_rels, "{" + PACKAGE + "}Relationship", Id="rId2", Type=OFFICE + "/styles", Target="styles.xml")
        z.writestr("xl/_rels/workbook.xml.rels", ET.tostring(wb_rels, encoding="utf-8", xml_declaration=True).replace(b"ns0:", b"").replace(b"xmlns:ns0=", b"xmlns="))
        z.writestr("xl/styles.xml", _detail_styles_xml())
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return out.getvalue()


# Backward-compatible name retained in case an older route still imports it.
def make_shipping_xlsx(groups, orders, *, kind="summary") -> bytes:
    if kind == "details":
        return make_shipping_detail_xlsx(orders, date.today())
    raise ValueError("新版統整表請使用 make_sellnow_xlsm() 並提供運回日期")
