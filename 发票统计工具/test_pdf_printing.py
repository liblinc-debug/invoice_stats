import os
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

import PyQt5
import fitz

os.environ.setdefault(
    "QT_QPA_PLATFORM_PLUGIN_PATH",
    str(Path(PyQt5.__file__).resolve().parent / "Qt5" / "plugins" / "platforms"),
)

from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtWidgets import QApplication, QToolBar

from gui import HELP_TAB, InvoiceWindow, PDF_PREVIEW_TAB
from pdf_printing import PdfPreviewWidget, PrintSettings, PrintSettingsDialog, ordered_slots, print_pdf_data
from workbook_printing import workbook_to_pdf_data


def sample_pdf(page_count: int = 3) -> bytes:
    writer = PdfWriter()
    for _ in range(page_count):
        writer.add_blank_page(width=595, height=842)
    stream = BytesIO()
    writer.write(stream)
    writer.close()
    return stream.getvalue()


class PdfPrintingTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_preview_and_settings_tabs(self):
        data = sample_pdf()
        preview = PdfPreviewWidget(data)
        self.assertEqual(preview.document.page_count, 3)
        dialog = PrintSettingsDialog(data)
        self.assertEqual([dialog.mode_tabs.tabText(i) for i in range(2)], ["页面大小", "一张多页"])
        dialog.mode_tabs.setCurrentIndex(1)
        self.assertEqual(dialog.settings().mode, "multiple")

    def test_main_window_shows_preview_and_print_button_before_exit(self):
        window = InvoiceWindow(Path(tempfile.gettempdir()))
        window.merged_pdf_data = sample_pdf()
        window._show_pdf_preview()
        actions = [action.text() for action in window.findChild(QToolBar).actions()]
        self.assertEqual(actions[-3:], ["打印", "帮助", "退出"])
        self.assertEqual(window.tabs.tabText(window.tabs.currentIndex()), PDF_PREVIEW_TAB)
        self.assertTrue(window.print_action.isEnabled())

    def test_merge_stays_in_workspace_until_export(self):
        window = InvoiceWindow(Path(tempfile.gettempdir()))
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as stream:
            temp_path = Path(stream.name)
            stream.write(sample_pdf())
        window._on_merge_finished(temp_path, 3, 3)
        self.assertFalse(temp_path.exists())
        self.assertEqual(window._current_workspace_kind(), "pdf")
        self.assertIn("请点击“导出”", window.statusBar().currentMessage())

    def test_pdf_export_uses_current_workspace(self):
        window = InvoiceWindow(Path(tempfile.gettempdir()))
        window.merged_pdf_data = sample_pdf()
        window._show_pdf_preview()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "合并发票_20261031_1200.pdf"
            with patch("gui.QFileDialog.getSaveFileName", return_value=(str(output), "")), patch(
                "gui.QMessageBox.information"
            ):
                window.export_current()
            self.assertEqual(len(PdfReader(str(output)).pages), 3)

    def test_workbook_print_help_and_close_tabs(self):
        workbook = Workbook()
        workbook.active.title = "材料费"
        workbook.active.append(["序号", "名称", "金额"])
        workbook.active.append([1, "测试材料", 12.5])
        travel = workbook.create_sheet("差旅费")
        travel.append(["日期", "发票名称", "金额"])
        travel.append(["2026-10-31", "住宿费", 88])

        pdf_data = workbook_to_pdf_data(workbook)
        document = fitz.open(stream=pdf_data, filetype="pdf")
        text = "\n".join(page.get_text() for page in document)
        self.assertGreaterEqual(document.page_count, 2)
        self.assertIn("测试材料", text)
        self.assertIn("住宿费", text)
        document.close()

        window = InvoiceWindow(Path(tempfile.gettempdir()))
        window.workbook = workbook
        window.template_path = Path("IDEA项目材料台账.xlsx")
        window._render_workbook()
        self.assertEqual(window._current_workspace_kind(), "workbook")
        self.assertTrue(window.export_action.isEnabled())
        self.assertTrue(window.print_action.isEnabled())

        window.show_help()
        self.assertEqual(window.tabs.tabText(window.tabs.currentIndex()), HELP_TAB)
        self.assertIn("商业用途时，需要付费", window.tabs.currentWidget().toPlainText())
        self.assertFalse(window.export_action.isEnabled())
        help_index = window.tabs.currentIndex()
        window._close_workspace_tab(help_index)
        self.assertFalse(window.help_open)

        window.merged_pdf_data = sample_pdf()
        window._show_pdf_preview()
        pdf_index = window.tabs.currentIndex()
        window._close_workspace_tab(pdf_index)
        self.assertIsNone(window.merged_pdf_data)

    def test_page_orders(self):
        self.assertEqual(ordered_slots(2, 2, "row"), [(0, 0), (0, 1), (1, 0), (1, 1)])
        self.assertEqual(ordered_slots(2, 2, "column"), [(0, 0), (1, 0), (0, 1), (1, 1)])

    def test_single_and_multiple_print_to_pdf(self):
        data = sample_pdf()
        with tempfile.TemporaryDirectory() as directory:
            for mode, expected_pages in (("single", 3), ("multiple", 2)):
                output = Path(directory) / f"{mode}.pdf"
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(str(output))
                printer.setResolution(150)
                settings = PrintSettings(mode=mode, pages_per_sheet=2, grayscale=False)
                self.assertEqual(print_pdf_data(data, settings, printer), expected_pages)
                self.assertEqual(len(PdfReader(str(output)).pages), expected_pages)


if __name__ == "__main__":
    unittest.main()
