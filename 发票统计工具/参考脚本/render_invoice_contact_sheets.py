from __future__ import annotations

import math
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
RENDER_DIR = ROOT / "tmp" / "pdfs" / "rendered"
CONTACT_DIR = ROOT / "tmp" / "pdfs" / "contacts"
PDFTOPPM = Path(
    "/Users/Lib_1/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm"
)

RENDER_DIR.mkdir(parents=True, exist_ok=True)
CONTACT_DIR.mkdir(parents=True, exist_ok=True)

try:
    font = ImageFont.truetype("/System/Library/Fonts/PingFang.ttc", 26)
except OSError:
    font = ImageFont.load_default()

entries = []
for category in ("材料费", "差旅费"):
    for index, pdf_path in enumerate(sorted((ROOT / category).glob("*.pdf")), start=1):
        prefix = RENDER_DIR / f"{category}_{index:02d}"
        image_path = prefix.with_suffix(".png")
        if not image_path.exists():
            subprocess.run(
                [str(PDFTOPPM), "-f", "1", "-singlefile", "-png", "-r", "100", str(pdf_path), str(prefix)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        entries.append((category, pdf_path.name, image_path))

tile_width, tile_height = 960, 680
cols, rows = 2, 3
for sheet_index in range(math.ceil(len(entries) / (cols * rows))):
    batch = entries[sheet_index * cols * rows : (sheet_index + 1) * cols * rows]
    canvas = Image.new("RGB", (tile_width * cols, tile_height * rows), "white")
    draw = ImageDraw.Draw(canvas)
    for offset, (category, filename, image_path) in enumerate(batch):
        row, col = divmod(offset, cols)
        label = f"{category}/{filename}"
        draw.text((col * tile_width + 12, row * tile_height + 10), label, fill="black", font=font)
        with Image.open(image_path) as source:
            source = source.convert("RGB")
            source.thumbnail((tile_width - 24, tile_height - 60), Image.Resampling.LANCZOS)
            x = col * tile_width + (tile_width - source.width) // 2
            y = row * tile_height + 52 + (tile_height - 60 - source.height) // 2
            canvas.paste(source, (x, y))
    output = CONTACT_DIR / f"contact_{sheet_index + 1:02d}.jpg"
    canvas.save(output, quality=92)
    print(output)
