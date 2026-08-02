#!/usr/bin/env python3
"""Extract bounded text, word geometry, and table candidates with pdfplumber."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
from pathlib import Path
import re
import sys


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_evidence(path: Path) -> dict[str, str | int]:
    resolved = path.expanduser().resolve()
    return {"path": str(resolved), "bytes": resolved.stat().st_size, "sha256": sha256(resolved)}


def _word_lines(words: list[dict]) -> list[list[dict]]:
    """Group pdfplumber words into deterministic visual lines.

    This is deliberately a small locator for explicit footnote markers. It is
    not a general reading-order algorithm and is never used to infer table
    cells.
    """
    lines: list[list[dict]] = []
    for word in sorted(words, key=lambda item: (float(item.get("top", 0)), float(item.get("x0", 0)))):
        top = float(word.get("top", 0))
        line = next((candidate for candidate in reversed(lines) if abs(float(candidate[0].get("top", 0)) - top) <= 1.5), None)
        if line is None:
            lines.append([word])
        else:
            line.append(word)
    return lines


def _extract_marked_footnotes(page, table_bbox: tuple[float, float, float, float]) -> list[dict]:
    """Keep explicit `*`/`Note:` lines outside the ruled table as metadata."""
    _, _, _, table_bottom = [float(value) for value in table_bbox]
    footnotes: list[dict] = []
    for line in _word_lines(page.extract_words() or []):
        ordered = sorted(line, key=lambda item: float(item.get("x0", 0)))
        text = " ".join(str(item.get("text", "")).strip() for item in ordered).strip()
        if not text or not re.match(r"^(?:\*|Note:)", text, re.IGNORECASE):
            continue
        top = min(float(item.get("top", 0)) for item in ordered)
        if top <= table_bottom + 2:
            continue
        footnotes.append({
            "text": text,
            "bbox": [
                min(float(item.get("x0", 0)) for item in ordered),
                top,
                max(float(item.get("x1", 0)) for item in ordered),
                max(float(item.get("bottom", 0)) for item in ordered),
            ],
        })
    return footnotes


def _table_cell(value, cell, page_number: int, row_index: int, column_index: int, colspan: int) -> dict:
    if cell is None:
        raise ValueError(f"table page {page_number} row {row_index} column {column_index} has a value without geometry")
    x0, top, x1, bottom = [float(item) for item in cell]
    return {
        "page": page_number,
        "text": str(value).strip(),
        "bbox": [x0, top, x1, bottom],
        "rowspan": 1,
        "colspan": colspan,
        "confidence": 1.0,
    }


def extract_table(source: Path, table_name: str | None, max_pages: int, max_tables: int) -> dict:
    """Extract ruled table cells while keeping marked footnotes separate."""
    try:
        import pdfplumber
    except ImportError as error:
        raise ValueError("pdfplumber is not installed") from error
    if not source.is_file() or source.is_symlink():
        raise ValueError("input must be an existing regular PDF")
    if max_pages < 1 or max_tables < 1:
        raise ValueError("max-pages and max-tables must be positive")
    pages: list[dict] = []
    cells: list[dict] = []
    footnotes: list[dict] = []
    with pdfplumber.open(str(source)) as pdf:
        if len(pdf.pages) > max_pages:
            raise ValueError(f"PDF has {len(pdf.pages)} pages; max-pages is {max_pages}")
        table_count = 0
        for page_number, page in enumerate(pdf.pages, 1):
            tables = page.find_tables()
            if len(tables) != 1:
                raise ValueError(f"table primitive requires exactly one table on page {page_number}; found {len(tables)}")
            table_count += len(tables)
            if table_count > max_tables:
                raise ValueError(f"table candidates exceed max-tables {max_tables}")
            table = tables[0]
            extracted = table.extract() or []
            if not extracted or len(extracted) != len(table.rows):
                raise ValueError(f"table page {page_number} has inconsistent extracted rows")
            page_cells: list[dict] = []
            for row_index, row in enumerate(table.rows):
                values = extracted[row_index]
                row_cells = list(row.cells)
                if len(row_cells) != len(values):
                    raise ValueError(f"table page {page_number} has a cell/value topology mismatch")
                # pdfplumber represents a cell spanning the full logical row as
                # one real geometry/value followed by null placeholders for
                # the covered columns. Preserve that topology as a colspan.
                populated = [
                    index for index, (cell, value) in enumerate(zip(row_cells, values))
                    if cell is not None and value is not None and str(value).strip()
                ]
                merged = len(row_cells) > 1 and populated == [0] and all(
                    cell is None and value is None
                    for cell, value in zip(row_cells[1:], values[1:])
                )
                for column_index, (cell, value) in enumerate(zip(row_cells, values)):
                    if value is None or not str(value).strip():
                        continue
                    page_cells.append(_table_cell(
                        value,
                        cell,
                        page_number,
                        row_index,
                        column_index,
                        len(row_cells) if merged else 1,
                    ))
            if not page_cells:
                raise ValueError(f"table page {page_number} has no non-empty cells")
            if table_name and table_name not in str(page.extract_text() or ""):
                raise ValueError(f"table page {page_number} does not contain the requested table name {table_name!r}")
            page_footnotes = _extract_marked_footnotes(page, table.bbox)
            cells.extend(page_cells)
            footnotes.extend({"page": page_number, **note} for note in page_footnotes)
            pages.append({
                "page": page_number,
                "bbox": [float(value) for value in table.bbox],
                "rows": len(extracted),
                "cells": len(page_cells),
                "footnotes": page_footnotes,
            })
    return {
        "table": table_name or "table",
        "cells": cells,
        "footnotes": footnotes,
        "pages": pages,
        "source": file_evidence(source),
        "provider": "pdfplumber",
        "providerVersion": getattr(pdfplumber, "__version__", "unknown"),
        "strategy": "read-only",
    }


def table_main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Extract one ruled table per page with geometry and separate marked footnotes.")
    parser.add_argument("input", type=Path)
    parser.add_argument("json_output", type=Path)
    parser.add_argument("csv_output", type=Path)
    parser.add_argument("--table-name", default=None)
    parser.add_argument("--max-pages", type=int, default=200)
    parser.add_argument("--max-tables", type=int, default=1_000)
    args = parser.parse_args(argv)
    source = args.input.expanduser().resolve()
    json_output = args.json_output.expanduser().resolve()
    csv_output = args.csv_output.expanduser().resolve()
    if source in {json_output, csv_output} or json_output == csv_output:
        raise ValueError("table outputs must be distinct from the input and from each other")
    report = extract_table(source, args.table_name, args.max_pages, args.max_tables)
    json_output.parent.mkdir(parents=True, exist_ok=True)
    csv_output.parent.mkdir(parents=True, exist_ok=True)
    json_output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    with csv_output.open("w", encoding="utf8", newline="") as stream:
        writer = csv.writer(stream, lineterminator="\n")
        writer.writerow(["page", "text", "bbox", "rowspan", "colspan", "confidence"])
        for cell in report["cells"]:
            writer.writerow([
                cell["page"],
                cell["text"],
                json.dumps(cell["bbox"], separators=(",", ":")),
                cell["rowspan"],
                cell["colspan"],
                cell["confidence"],
            ])
    print(json.dumps({
        "ok": True,
        "table": report["table"],
        "pages": len(report["pages"]),
        "cells": len(report["cells"]),
        "footnotes": len(report["footnotes"]),
        "source": report["source"],
        "outputs": {
            "json": file_evidence(json_output),
            "csv": file_evidence(csv_output),
        },
    }, ensure_ascii=False, indent=2))
    return 0


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    if len(sys.argv) > 1 and sys.argv[1] == "table":
        try:
            return table_main(sys.argv[2:])
        except Exception as exc:
            print(json.dumps({"ok": False, "error": str(exc), "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
            return 2
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--max-pages", type=int, default=200)
    parser.add_argument("--max-words", type=int, default=50_000)
    parser.add_argument("--max-tables", type=int, default=1_000)
    parser.add_argument("--max-chars", type=int, default=2_000_000)
    parser.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    args = parser.parse_args()
    try:
        import pdfplumber
    except ImportError:
        print(json.dumps({"ok": False, "error": "pdfplumber is not installed", "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
        return 2
    try:
        source = args.input.expanduser().resolve()
        if not source.is_file():
            raise ValueError("input must be an existing PDF")
        if args.max_pages < 1 or args.max_words < 1 or args.max_tables < 1 or args.max_chars < 1 or args.max_bytes < 1:
            raise ValueError("all extraction limits must be positive")
        if source.stat().st_size > args.max_bytes:
            raise ValueError(f"PDF is {source.stat().st_size} bytes; max-bytes is {args.max_bytes}")
        pages = []
        total_words = total_tables = total_chars = 0
        with pdfplumber.open(str(source)) as pdf:
            if len(pdf.pages) > args.max_pages:
                raise ValueError(f"PDF has {len(pdf.pages)} pages; max-pages is {args.max_pages}")
            for page_number, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                total_chars += len(text)
                if total_chars > args.max_chars:
                    raise ValueError(f"extracted text exceeds max-chars {args.max_chars}")
                words = page.extract_words() or []
                total_words += len(words)
                if total_words > args.max_words:
                    raise ValueError(f"extracted words exceed max-words {args.max_words}")
                tables = page.extract_tables() or []
                total_tables += len(tables)
                if total_tables > args.max_tables:
                    raise ValueError(f"table candidates exceed max-tables {args.max_tables}")
                pages.append({
                    "page": page_number,
                    "width": page.width,
                    "height": page.height,
                    "text": text,
                    "words": [{key: word.get(key) for key in ("text", "x0", "x1", "top", "bottom", "doctop", "upright", "direction")} for word in words],
                    "tables": tables,
                    "lines": len(page.lines or []),
                    "rects": len(page.rects or []),
                    "images": len(page.images or []),
                })
        payload = {
            "provider": "pdfplumber",
            "strategy": "read-only",
            "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": sha256(source)},
            "summary": {"pages": len(pages), "chars": total_chars, "words": total_words, "tableCandidates": total_tables},
            "pages": pages,
            "warning": "table extraction is heuristic and must be checked against rendered page geometry",
        }
        rendered = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.output:
            output = args.output.expanduser().resolve()
            if output == source:
                raise ValueError("JSON report path must differ from the source PDF")
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(rendered + "\n", "utf-8")
        else:
            print(rendered)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc), "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
