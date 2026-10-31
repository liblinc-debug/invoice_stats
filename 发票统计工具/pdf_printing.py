from __future__ import annotations

from dataclasses import dataclass

import fitz
from PyQt5.QtCore import QRectF, Qt
from PyQt5.QtGui import QColor, QImage, QPageLayout, QPageSize, QPainter, QPen, QPixmap, QTransform
from PyQt5.QtPrintSupport import QPrinter, QPrinterInfo
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class PrintSettings:
    printer_name: str = ""
    mode: str = "single"
    scale_mode: str = "fit"
    scale_percent: int = 100
    pages_per_sheet: int = 2
    page_order: str = "row"
    draw_border: bool = False
    paper_size: str = "A4"
    orientation: str = "portrait"
    margin_mm: float = 10.0
    grayscale: bool = True
    center: bool = True
    auto_rotate: bool = False


def nup_grid(page_count: int) -> tuple[int, int]:
    return {
        2: (2, 1),
        4: (2, 2),
        6: (3, 2),
        9: (3, 3),
        16: (4, 4),
    }.get(page_count, (2, 1))


def ordered_slots(rows: int, columns: int, order: str) -> list[tuple[int, int]]:
    if order == "row_reverse":
        return [(row, column) for row in range(rows) for column in reversed(range(columns))]
    if order == "column":
        return [(row, column) for column in range(columns) for row in range(rows)]
    if order == "column_reverse":
        return [(row, column) for column in reversed(range(columns)) for row in range(rows)]
    return [(row, column) for row in range(rows) for column in range(columns)]


def render_pdf_page(document: fitz.Document, page_index: int, dpi: int, grayscale: bool) -> QImage:
    page = document.load_page(page_index)
    colorspace = fitz.csGRAY if grayscale else fitz.csRGB
    pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), colorspace=colorspace, alpha=False)
    image_format = QImage.Format_Grayscale8 if grayscale else QImage.Format_RGB888
    return QImage(pixmap.samples, pixmap.width, pixmap.height, pixmap.stride, image_format).copy()


def _draw_page(
    painter: QPainter,
    document: fitz.Document,
    page_index: int,
    cell: QRectF,
    settings: PrintSettings,
    resolution: int,
) -> None:
    page = document.load_page(page_index)
    render_dpi = min(max(resolution, 96), 300)
    image = render_pdf_page(document, page_index, render_dpi, settings.grayscale)
    natural_width = page.rect.width * resolution / 72
    natural_height = page.rect.height * resolution / 72

    normal_fit = min(cell.width() / image.width(), cell.height() / image.height())
    rotated_fit = min(cell.width() / image.height(), cell.height() / image.width())
    if settings.auto_rotate and rotated_fit > normal_fit:
        image = image.transformed(QTransform().rotate(90), Qt.SmoothTransformation)
        natural_width, natural_height = natural_height, natural_width

    fit_scale = min(cell.width() / image.width(), cell.height() / image.height())
    if settings.mode == "multiple" or settings.scale_mode == "fit":
        width = image.width() * fit_scale
        height = image.height() * fit_scale
    else:
        multiplier = 1.0
        if settings.scale_mode == "shrink":
            multiplier = min(1.0, cell.width() / natural_width, cell.height() / natural_height)
        elif settings.scale_mode == "custom":
            multiplier = settings.scale_percent / 100
        width = natural_width * multiplier
        height = natural_height * multiplier

    left = cell.left() + (cell.width() - width) / 2 if settings.center else cell.left()
    top = cell.top() + (cell.height() - height) / 2 if settings.center else cell.top()
    target = QRectF(left, top, width, height)
    painter.save()
    painter.setClipRect(cell)
    painter.drawImage(target, image)
    painter.restore()
    if settings.draw_border and settings.mode == "multiple":
        painter.setPen(QPen(QColor("#777777"), 1))
        painter.drawRect(cell.adjusted(1, 1, -1, -1))


def draw_pdf_sheet(
    painter: QPainter,
    document: fitz.Document,
    page_indices: list[int],
    page_rect: QRectF,
    settings: PrintSettings,
    resolution: int,
) -> None:
    rows, columns = nup_grid(settings.pages_per_sheet) if settings.mode == "multiple" else (1, 1)
    gap = max(2.0, resolution * 2 / 25.4) if settings.mode == "multiple" else 0.0
    cell_width = (page_rect.width() - gap * (columns - 1)) / columns
    cell_height = (page_rect.height() - gap * (rows - 1)) / rows
    for page_index, (row, column) in zip(page_indices, ordered_slots(rows, columns, settings.page_order)):
        cell = QRectF(
            page_rect.left() + column * (cell_width + gap),
            page_rect.top() + row * (cell_height + gap),
            cell_width,
            cell_height,
        )
        _draw_page(painter, document, page_index, cell, settings, resolution)


def _set_printer_layout(printer: QPrinter, document: fitz.Document, settings: PrintSettings) -> None:
    paper_sizes = {
        "A4": QPageSize.A4,
        "A3": QPageSize.A3,
        "Letter": QPageSize.Letter,
    }
    printer.setPageSize(QPageSize(paper_sizes.get(settings.paper_size, QPageSize.A4)))
    orientation = settings.orientation
    if orientation == "auto":
        first_page = document.load_page(0)
        orientation = "landscape" if first_page.rect.width > first_page.rect.height else "portrait"
    printer.setPageOrientation(
        QPageLayout.Landscape if orientation == "landscape" else QPageLayout.Portrait
    )
    printer.setPageMargins(
        settings.margin_mm,
        settings.margin_mm,
        settings.margin_mm,
        settings.margin_mm,
        QPrinter.Millimeter,
    )
    printer.setColorMode(QPrinter.GrayScale if settings.grayscale else QPrinter.Color)


def print_pdf_data(pdf_data: bytes, settings: PrintSettings, printer: QPrinter) -> int:
    document = fitz.open(stream=pdf_data, filetype="pdf")
    try:
        if document.page_count == 0:
            raise ValueError("PDF 文档没有可打印页面")
        _set_printer_layout(printer, document, settings)
        pages_per_sheet = settings.pages_per_sheet if settings.mode == "multiple" else 1
        painter = QPainter()
        if not painter.begin(printer):
            raise RuntimeError("无法启动所选打印机")
        try:
            sheet_count = 0
            for start in range(0, document.page_count, pages_per_sheet):
                if sheet_count and not printer.newPage():
                    raise RuntimeError("打印机无法创建下一页")
                indices = list(range(start, min(start + pages_per_sheet, document.page_count)))
                page_rect = QRectF(printer.pageRect(QPrinter.DevicePixel))
                draw_pdf_sheet(painter, document, indices, page_rect, settings, printer.resolution())
                sheet_count += 1
        finally:
            painter.end()
        if printer.printerState() == QPrinter.Error:
            raise RuntimeError("打印机报告输出错误")
        return sheet_count
    finally:
        document.close()


def render_sheet_preview(pdf_data: bytes, settings: PrintSettings) -> QPixmap:
    document = fitz.open(stream=pdf_data, filetype="pdf")
    try:
        landscape = settings.orientation == "landscape"
        if settings.orientation == "auto" and document.page_count:
            page = document.load_page(0)
            landscape = page.rect.width > page.rect.height
        width, height = (650, 460) if landscape else (460, 650)
        canvas = QImage(width + 40, height + 40, QImage.Format_RGB32)
        canvas.fill(QColor("#303030"))
        painter = QPainter(canvas)
        paper_rect = QRectF(20, 20, width, height)
        painter.fillRect(paper_rect, Qt.white)
        margin = settings.margin_mm * 3
        content_rect = paper_rect.adjusted(margin, margin, -margin, -margin)
        count = settings.pages_per_sheet if settings.mode == "multiple" else 1
        draw_pdf_sheet(
            painter,
            document,
            list(range(min(count, document.page_count))),
            content_rect,
            settings,
            96,
        )
        painter.end()
        return QPixmap.fromImage(canvas)
    finally:
        document.close()


class PdfPreviewWidget(QWidget):
    def __init__(self, pdf_data: bytes, parent: QWidget | None = None):
        super().__init__(parent)
        self.pdf_data = pdf_data
        self.document = fitz.open(stream=pdf_data, filetype="pdf")
        self.page_index = 0
        self.zoom = 1.0

        self.page_label = QLabel()
        self.page_label.setAlignment(Qt.AlignCenter)
        self.page_label.setStyleSheet("QLabel { background: #4a4a4a; padding: 18px; }")
        scroll = QScrollArea()
        scroll.setWidget(self.page_label)
        scroll.setWidgetResizable(True)

        self.first_button = QPushButton("|<")
        self.previous_button = QPushButton("<")
        self.page_text = QLabel()
        self.next_button = QPushButton(">")
        self.last_button = QPushButton(">|")
        self.zoom_out_button = QPushButton("缩小")
        self.zoom_in_button = QPushButton("放大")
        controls = QHBoxLayout()
        controls.addStretch()
        for button in (self.first_button, self.previous_button):
            controls.addWidget(button)
        controls.addWidget(self.page_text)
        for button in (self.next_button, self.last_button):
            controls.addWidget(button)
        controls.addSpacing(20)
        controls.addWidget(self.zoom_out_button)
        controls.addWidget(self.zoom_in_button)
        controls.addStretch()

        layout = QVBoxLayout(self)
        layout.addWidget(scroll, 1)
        layout.addLayout(controls)
        self.first_button.clicked.connect(lambda: self.set_page(0))
        self.previous_button.clicked.connect(lambda: self.set_page(self.page_index - 1))
        self.next_button.clicked.connect(lambda: self.set_page(self.page_index + 1))
        self.last_button.clicked.connect(lambda: self.set_page(self.document.page_count - 1))
        self.zoom_out_button.clicked.connect(lambda: self.set_zoom(self.zoom / 1.25))
        self.zoom_in_button.clicked.connect(lambda: self.set_zoom(self.zoom * 1.25))
        self._render()

    def set_page(self, page_index: int) -> None:
        self.page_index = max(0, min(page_index, self.document.page_count - 1))
        self._render()

    def set_zoom(self, zoom: float) -> None:
        self.zoom = max(0.5, min(zoom, 3.0))
        self._render()

    def _render(self) -> None:
        if not self.document.page_count:
            self.page_label.setText("PDF 文档没有页面")
            return
        image = render_pdf_page(self.document, self.page_index, round(110 * self.zoom), False)
        self.page_label.setPixmap(QPixmap.fromImage(image))
        self.page_label.resize(image.size())
        self.page_text.setText(f"{self.page_index + 1}/{self.document.page_count}  ({self.zoom:.0%})")
        self.first_button.setEnabled(self.page_index > 0)
        self.previous_button.setEnabled(self.page_index > 0)
        self.next_button.setEnabled(self.page_index + 1 < self.document.page_count)
        self.last_button.setEnabled(self.page_index + 1 < self.document.page_count)


class PrintSettingsDialog(QDialog):
    def __init__(self, pdf_data: bytes, parent: QWidget | None = None):
        super().__init__(parent)
        self.pdf_data = pdf_data
        self.setWindowTitle("打印")
        self.resize(1080, 720)

        self.printer_combo = QComboBox()
        printers = QPrinterInfo.availablePrinters()
        default_name = QPrinterInfo.defaultPrinter().printerName()
        for printer in printers:
            self.printer_combo.addItem(printer.printerName())
        if default_name:
            index = self.printer_combo.findText(default_name)
            if index >= 0:
                self.printer_combo.setCurrentIndex(index)
        self.grayscale_check = QCheckBox("灰度打印")
        self.grayscale_check.setChecked(True)
        basic_form = QFormLayout()
        basic_form.addRow("打印机：", self.printer_combo)
        basic_form.addRow("颜色：", self.grayscale_check)
        basic_group = QGroupBox("基础设置")
        basic_group.setLayout(basic_form)

        self.mode_tabs = QTabWidget()
        self.mode_tabs.addTab(self._build_single_tab(), "页面大小")
        self.mode_tabs.addTab(self._build_multiple_tab(), "一张多页")

        self.paper_combo = QComboBox()
        self.paper_combo.addItems(["A4", "A3", "Letter"])
        self.orientation_combo = QComboBox()
        self.orientation_combo.addItem("自动横向/纵向", "auto")
        self.orientation_combo.addItem("纵向", "portrait")
        self.orientation_combo.addItem("横向", "landscape")
        self.orientation_combo.setCurrentIndex(1)
        self.margin_spin = QDoubleSpinBox()
        self.margin_spin.setRange(0, 30)
        self.margin_spin.setSuffix(" mm")
        self.margin_spin.setValue(10)
        page_form = QFormLayout()
        page_form.addRow("纸张大小：", self.paper_combo)
        page_form.addRow("纸张方向：", self.orientation_combo)
        page_form.addRow("页边距：", self.margin_spin)
        page_group = QGroupBox("页面设置")
        page_group.setLayout(page_form)

        self.center_check = QCheckBox("自动居中")
        self.center_check.setChecked(True)
        self.rotate_check = QCheckBox("自动旋转")
        content_layout = QHBoxLayout()
        content_layout.addWidget(self.center_check)
        content_layout.addWidget(self.rotate_check)
        content_group = QGroupBox("内容设置")
        content_group.setLayout(content_layout)

        left = QVBoxLayout()
        left.addWidget(basic_group)
        left.addWidget(self.mode_tabs)
        left.addWidget(page_group)
        left.addWidget(content_group)
        left.addStretch()

        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(460, 540)
        self.preview_label.setStyleSheet("QLabel { background: #252525; }")
        preview_group = QGroupBox("打印预览")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.addWidget(self.preview_label)

        body = QHBoxLayout()
        body.addLayout(left, 1)
        body.addWidget(preview_group, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("确定并打印")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self._accept_if_printer_available)
        buttons.rejected.connect(self.reject)
        self.ok_button = buttons.button(QDialogButtonBox.Ok)

        layout = QVBoxLayout(self)
        layout.addLayout(body, 1)
        layout.addWidget(buttons)
        self._connect_preview_updates()
        self.ok_button.setEnabled(self.printer_combo.count() > 0)
        if not printers:
            self.printer_combo.addItem("未发现可用打印机")
        self._update_preview()

    def _build_single_tab(self) -> QWidget:
        widget = QWidget()
        self.scale_combo = QComboBox()
        self.scale_combo.addItem("适合打印边距", "fit")
        self.scale_combo.addItem("实际大小", "actual")
        self.scale_combo.addItem("缩小过大页面", "shrink")
        self.scale_combo.addItem("自定义比例", "custom")
        self.scale_spin = QSpinBox()
        self.scale_spin.setRange(10, 400)
        self.scale_spin.setSuffix("%")
        self.scale_spin.setValue(100)
        form = QFormLayout(widget)
        form.addRow("页面缩放：", self.scale_combo)
        form.addRow("自定义比例：", self.scale_spin)
        return widget

    def _build_multiple_tab(self) -> QWidget:
        widget = QWidget()
        self.pages_combo = QComboBox()
        for count, layout in ((2, "1 列 × 2 行"), (4, "2 列 × 2 行"), (6, "2 列 × 3 行"), (9, "3 列 × 3 行"), (16, "4 列 × 4 行")):
            self.pages_combo.addItem(f"{count} 页（{layout}）", count)
        self.order_combo = QComboBox()
        self.order_combo.addItem("从左到右、从上到下", "row")
        self.order_combo.addItem("从右到左、从上到下", "row_reverse")
        self.order_combo.addItem("从上到下、从左到右", "column")
        self.order_combo.addItem("从上到下、从右到左", "column_reverse")
        self.border_check = QCheckBox("打印页面边框")
        form = QFormLayout(widget)
        form.addRow("每张页数：", self.pages_combo)
        form.addRow("页面顺序：", self.order_combo)
        form.addRow("边框：", self.border_check)
        return widget

    def _connect_preview_updates(self) -> None:
        for signal in (
            self.mode_tabs.currentChanged,
            self.printer_combo.currentIndexChanged,
            self.scale_combo.currentIndexChanged,
            self.scale_spin.valueChanged,
            self.pages_combo.currentIndexChanged,
            self.order_combo.currentIndexChanged,
            self.paper_combo.currentIndexChanged,
            self.orientation_combo.currentIndexChanged,
            self.margin_spin.valueChanged,
            self.grayscale_check.toggled,
            self.center_check.toggled,
            self.rotate_check.toggled,
            self.border_check.toggled,
        ):
            signal.connect(self._update_preview)

    def settings(self) -> PrintSettings:
        return PrintSettings(
            printer_name=self.printer_combo.currentText(),
            mode="multiple" if self.mode_tabs.currentIndex() == 1 else "single",
            scale_mode=self.scale_combo.currentData(),
            scale_percent=self.scale_spin.value(),
            pages_per_sheet=self.pages_combo.currentData(),
            page_order=self.order_combo.currentData(),
            draw_border=self.border_check.isChecked(),
            paper_size=self.paper_combo.currentText(),
            orientation=self.orientation_combo.currentData(),
            margin_mm=self.margin_spin.value(),
            grayscale=self.grayscale_check.isChecked(),
            center=self.center_check.isChecked(),
            auto_rotate=self.rotate_check.isChecked(),
        )

    def _update_preview(self, *_args) -> None:
        pixmap = render_sheet_preview(self.pdf_data, self.settings())
        self.preview_label.setPixmap(
            pixmap.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )

    def _accept_if_printer_available(self) -> None:
        if not QPrinterInfo.availablePrinters():
            QMessageBox.warning(self, "无法打印", "未发现可用打印机，请先在系统中添加打印机。")
            return
        self.accept()

