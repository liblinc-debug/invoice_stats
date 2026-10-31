#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
from pathlib import Path

import PyQt5

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault(
    "QT_QPA_PLATFORM_PLUGIN_PATH",
    str(Path(PyQt5.__file__).resolve().parent / "Qt5" / "plugins" / "platforms"),
)

from PIL import Image
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QGuiApplication, QImage, QPainter
from PyQt5.QtSvg import QSvgRenderer


def render_png(renderer: QSvgRenderer, size: int, output_path: Path) -> None:
    image = QImage(size, size, QImage.Format_ARGB32)
    image.fill(Qt.transparent)
    painter = QPainter(image)
    renderer.render(painter)
    painter.end()
    if not image.save(str(output_path), "PNG"):
        raise RuntimeError(f"无法生成图标：{output_path}")


def build_windows_icon(renderer: QSvgRenderer, output_path: Path) -> None:
    png_path = output_path.with_suffix(".png")
    render_png(renderer, 1024, png_path)
    try:
        with Image.open(png_path) as image:
            image.save(output_path, format="ICO", sizes=[(size, size) for size in (16, 24, 32, 48, 64, 128, 256)])
    finally:
        png_path.unlink(missing_ok=True)


def build_macos_icon(renderer: QSvgRenderer, output_path: Path) -> None:
    png_path = output_path.with_suffix(".png")
    render_png(renderer, 1024, png_path)
    try:
        with Image.open(png_path) as image:
            image.save(output_path, format="ICNS")
    finally:
        png_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="从应用 SVG 生成平台图标")
    parser.add_argument("platform", choices=("macos", "windows"))
    parser.add_argument("svg", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    app = QGuiApplication.instance() or QGuiApplication([])
    renderer = QSvgRenderer(str(args.svg))
    if not renderer.isValid():
        raise ValueError(f"无效的 SVG 图标：{args.svg}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    if args.platform == "macos":
        build_macos_icon(renderer, args.output)
    else:
        build_windows_icon(renderer, args.output)
    if not args.output.is_file() or args.output.stat().st_size == 0:
        raise RuntimeError(f"图标生成失败：{args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
