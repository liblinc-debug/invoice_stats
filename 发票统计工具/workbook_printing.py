from __future__ import annotations

import html
import tempfile
from datetime import date, datetime
from pathlib import Path

import fitz
from PyQt5.QtGui import QPageLayout, QPageSize, QTextDocument
from PyQt5.QtPrintSupport import QPrinter


def _display_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def _sheet_html(sheet) -> str:
    anchors: dict[tuple[int, int], tuple[int, int]] = {}
    merged_children: set[tuple[int, int]] = set()
    for merged_range in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        anchors[(min_row, min_col)] = (max_row - min_row + 1, max_col - min_col + 1)
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                if (row, column) != (min_row, min_col):
                    merged_children.add((row, column))

    rows = []
    for row in range(1, max(sheet.max_row, 1) + 1):
        cells = []
        for column in range(1, max(sheet.max_column, 1) + 1):
            if (row, column) in merged_children:
                continue
            cell = sheet.cell(row, column)
            rowspan, colspan = anchors.get((row, column), (1, 1))
            attributes = []
            if rowspan > 1:
                attributes.append(f'rowspan="{rowspan}"')
            if colspan > 1:
                attributes.append(f'colspan="{colspan}"')
            styles = []
            if cell.fill.fill_type and cell.fill.fgColor.type == "rgb" and cell.fill.fgColor.rgb:
                styles.append(f"background-color: #{cell.fill.fgColor.rgb[-6:]}")
            if cell.font.bold:
                styles.append("font-weight: bold")
            if cell.alignment.horizontal in {"center", "right"}:
                styles.append(f"text-align: {cell.alignment.horizontal}")
            if styles:
                attributes.append(f'style="{"; ".join(styles)}"')
            value = html.escape(_display_value(cell.value)).replace("\n", "<br>")
            cells.append(f"<td {' '.join(attributes)}>{value}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")

    return (
        f"<section><h1>{html.escape(sheet.title)}</h1>"
        f"<table>{''.join(rows)}</table></section>"
    )


def workbook_to_pdf_data(workbook) -> bytes:
    if workbook is None or not workbook.worksheets:
        raise ValueError("工作簿没有可打印的子表")
    sheets = [sheet for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
    if not sheets:
        raise ValueError("工作簿没有可打印的可见子表")
    combined = fitz.open()
    try:
        for sheet in sheets:
            document = QTextDocument()
            document.setHtml(
                "<html><head><style>"
                "body { font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif; font-size: 9pt; }"
                "h1 { font-size: 14pt; margin: 0 0 8px 0; }"
                "table { border-collapse: collapse; width: 100%; margin-bottom: 8px; }"
                "td { border: 0.5px solid #999; padding: 3px; vertical-align: middle; }"
                "</style></head><body>" + _sheet_html(sheet) + "</body></html>"
            )
            with tempfile.NamedTemporaryFile(
                prefix="invoice-workbook-", suffix=".pdf", delete=False
            ) as stream:
                temp_path = Path(stream.name)
            try:
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(str(temp_path))
                printer.setPageSize(QPageSize(QPageSize.A4))
                printer.setPageOrientation(QPageLayout.Landscape)
                printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
                document.print_(printer)
                sheet_pdf = fitz.open(str(temp_path))
                combined.insert_pdf(sheet_pdf)
                sheet_pdf.close()
            finally:
                temp_path.unlink(missing_ok=True)
        pdf_data = combined.tobytes()
        checked = fitz.open(stream=pdf_data, filetype="pdf")
        try:
            if checked.page_count == 0:
                raise ValueError("工作簿打印预览没有生成页面")
        finally:
            checked.close()
        return pdf_data
    finally:
        combined.close()
