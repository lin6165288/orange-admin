from __future__ import annotations

from decimal import Decimal
from io import BytesIO
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def _inline_cell(ref: str, value, style: int = 0) -> str:
    text = escape(str(value if value is not None else ""))
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t>{text}</t></is></c>'


def _number_cell(ref: str, value, style: int = 0) -> str:
    num = Decimal(str(value or 0))
    return f'<c r="{ref}" s="{style}"><v>{num}</v></c>'


def make_member_balance_xlsx(rows: list[dict]) -> bytes:
    sheet_rows = []
    sheet_rows.append('<row r="1" ht="28" customHeight="1">' + _inline_cell('A1', '橘貓代購｜有儲值金餘額會員', 3) + '</row>')
    sheet_rows.append('<row r="3" ht="24" customHeight="1">' + _inline_cell('A3', '會員姓名', 1) + _inline_cell('B3', '儲值金餘額', 1) + '</row>')

    total = Decimal('0')
    for row_idx, row in enumerate(rows, start=4):
        balance = Decimal(str(row.get('balance') or 0))
        total += balance
        sheet_rows.append(
            f'<row r="{row_idx}">'
            + _inline_cell(f'A{row_idx}', row.get('customer_name') or '', 0)
            + _number_cell(f'B{row_idx}', balance, 2)
            + '</row>'
        )

    total_row = 4 + len(rows)
    sheet_rows.append(
        f'<row r="{total_row}" ht="24" customHeight="1">'
        + _inline_cell(f'A{total_row}', '合計', 4)
        + _number_cell(f'B{total_row}', total, 5)
        + '</row>'
    )

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:B{total_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="20"/>'
        '<cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/></cols>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
        f'<autoFilter ref="A3:B{max(3, total_row - 1)}"/>'
        '</worksheet>'
    )

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="NT$ #,##0.00"/></numFmts>
  <fonts count="3"><font><sz val="11"/><name val="Microsoft JhengHei"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft JhengHei"/></font><font><b/><sz val="14"/><color rgb="FFF08300"/><name val="Microsoft JhengHei"/></font></fonts>
  <fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF08300"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF3E5"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFE6E6E6"/></left><right style="thin"><color rgb="FFE6E6E6"/></right><top style="thin"><color rgb="FFE6E6E6"/></top><bottom style="thin"><color rgb="FFE6E6E6"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0"/><xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

    workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="儲值金餘額" sheetId="1" r:id="rId1"/></sheets></workbook>'
    workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

    out = BytesIO()
    with ZipFile(out, 'w', ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', workbook_xml)
        z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        z.writestr('xl/styles.xml', styles_xml)
    return out.getvalue()
