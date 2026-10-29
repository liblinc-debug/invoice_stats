from __future__ import annotations

import json
from pathlib import Path

from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[1]
records = []
for category in ("材料费", "差旅费"):
    for pdf_path in sorted((ROOT / category).glob("*.pdf")):
        reader = PdfReader(str(pdf_path))
        pages = []
        for page_number, page in enumerate(reader.pages, start=1):
            pages.append({"page": page_number, "text": page.extract_text() or ""})
        records.append(
            {
                "category": category,
                "path": str(pdf_path.relative_to(ROOT)),
                "page_count": len(reader.pages),
                "pages": pages,
            }
        )

output = ROOT / "tmp" / "invoice_extraction_raw.json"
output.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"wrote {len(records)} records to {output}")
for record in records:
    chars = sum(len(page["text"]) for page in record["pages"])
    print(f"{record['category']}\t{record['page_count']}p\t{chars} chars\t{record['path']}")
