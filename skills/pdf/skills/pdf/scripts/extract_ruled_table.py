#!/usr/bin/env python3
"""Extract one verified, cross-page ruled table without reconstructing the PDF.

This deliberately supports a narrow profile: an explicitly titled table with a
complete ruled grid, fixed column boundaries, repeated header geometry on
consecutive pages, and rectangular data cells.  It is not a general document
table-recognition or PDF reflow tool.  Any ambiguous candidate is rejected
before JSON or CSV is published.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import os
from pathlib import Path
import re
import secrets
import sys
from typing import Any


SCHEMA = "office-kit.pdf-ruled-table.v1"
PROFILE = "ruled-cross-page-v1"
COORDINATE_TOLERANCE = 0.75
LINE_TOLERANCE = 3.0
TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 2,
    "join_tolerance": 2,
    "edge_min_length": 8,
    "intersection_tolerance": 2,
}


class ExtractionError(RuntimeError):
    """Raised when the requested ruled-table proof does not hold."""


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_evidence(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def canonical_text(value: str) -> str:
    return " ".join(str(value or "").replace("\u00a0", " ").split())


def canonical_key(value: str) -> str:
    return canonical_text(value).casefold()


def rounded(value: float) -> float:
    return round(float(value), 3)


def close(left: float, right: float, tolerance: float = COORDINATE_TOLERANCE) -> bool:
    return math.isclose(float(left), float(right), abs_tol=tolerance)


def unique_coordinates(values: list[float]) -> list[float]:
    coordinates: list[float] = []
    for value in sorted(float(item) for item in values):
        if not coordinates or not close(coordinates[-1], value):
            coordinates.append(value)
        else:
            coordinates[-1] = (coordinates[-1] + value) / 2
    return coordinates


def coordinate_index(coordinates: list[float], value: float, label: str) -> int:
    matches = [index for index, candidate in enumerate(coordinates) if close(candidate, value)]
    if len(matches) != 1:
        raise ExtractionError(f"{label} {value:.3f} is not one unambiguous ruled-grid coordinate")
    return matches[0]


def word_lines(words: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group ordinary horizontal words into stable visual lines.

    We intentionally do not use page text-flow output here: table cells need
    their own geometry, and a two-column narrative must not influence a table
    title or footnote decision.
    """

    ordered = sorted(words, key=lambda word: (float(word["top"]), float(word["x0"])))
    groups: list[list[dict[str, Any]]] = []
    for word in ordered:
        if not bool(word.get("upright", True)):
            continue
        if not groups or float(word["top"]) - min(float(item["top"]) for item in groups[-1]) > LINE_TOLERANCE:
            groups.append([word])
        else:
            groups[-1].append(word)
    lines: list[dict[str, Any]] = []
    for group in groups:
        group.sort(key=lambda word: float(word["x0"]))
        text = canonical_text(" ".join(str(word.get("text", "")) for word in group))
        if not text:
            continue
        lines.append({
            "text": text,
            "bbox": [
                rounded(min(float(word["x0"]) for word in group)),
                rounded(min(float(word["top"]) for word in group)),
                rounded(max(float(word["x1"]) for word in group)),
                rounded(max(float(word["bottom"]) for word in group)),
            ],
        })
    return lines


def title_lines(lines: list[dict[str, Any]], title: str) -> list[dict[str, Any]]:
    expected = canonical_key(title)
    return [line for line in lines if canonical_key(line["text"]).startswith(expected)]


def candidate_for_title(page: Any, lines: list[dict[str, Any]], title: str) -> tuple[Any, dict[str, Any]]:
    matches: list[tuple[Any, dict[str, Any]]] = []
    for title_line in title_lines(lines, title):
        _, title_top, _, title_bottom = title_line["bbox"]
        for table in page.find_tables(table_settings=TABLE_SETTINGS):
            x0, top, _, _ = (float(value) for value in table.bbox)
            if top < title_bottom - COORDINATE_TOLERANCE:
                continue
            if top - title_bottom > 90:
                continue
            if abs(x0 - float(title_line["bbox"][0])) > 72:
                continue
            matches.append((table, title_line))
    unique: list[tuple[Any, dict[str, Any]]] = []
    for table, title_line in matches:
        if not any(all(close(left, right) for left, right in zip(table.bbox, candidate.bbox)) for candidate, _ in unique):
            unique.append((table, title_line))
    if len(unique) != 1:
        raise ExtractionError(f"expected exactly one ruled table below title {title!r}; found {len(unique)}")
    return unique[0]


def words_in_bbox(words: list[dict[str, Any]], bbox: tuple[float, float, float, float]) -> list[dict[str, Any]]:
    x0, top, x1, bottom = bbox
    selected = []
    for word in words:
        center_x = (float(word["x0"]) + float(word["x1"])) / 2
        center_y = (float(word["top"]) + float(word["bottom"])) / 2
        if x0 + COORDINATE_TOLERANCE < center_x < x1 - COORDINATE_TOLERANCE and top + COORDINATE_TOLERANCE < center_y < bottom - COORDINATE_TOLERANCE:
            selected.append(word)
    return selected


def cell_text(words: list[dict[str, Any]]) -> tuple[str, bool]:
    if any(not bool(word.get("upright", True)) for word in words):
        raise ExtractionError("ruled-table profile rejects rotated text inside a cell")
    lines = word_lines(words)
    return canonical_text("\n".join(line["text"] for line in lines)), bool(words)


def extract_grid_cells(table: Any, words: list[dict[str, Any]], expected_columns: int) -> tuple[list[dict[str, Any]], list[float], list[float]]:
    raw_cells = [tuple(float(value) for value in cell) for cell in table.cells]
    if not raw_cells:
        raise ExtractionError("ruled table has no detected cells")
    x_coordinates = unique_coordinates([value for cell in raw_cells for value in (cell[0], cell[2])])
    y_coordinates = unique_coordinates([value for cell in raw_cells for value in (cell[1], cell[3])])
    if len(x_coordinates) != expected_columns + 1:
        raise ExtractionError(f"ruled grid has {len(x_coordinates) - 1} columns; expected {expected_columns}")
    if len(y_coordinates) < 3:
        raise ExtractionError("ruled grid has fewer than two rows")

    seen: set[tuple[int, int, int, int]] = set()
    cells: list[dict[str, Any]] = []
    coverage: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    for cell in raw_cells:
        x0, top, x1, bottom = cell
        column = coordinate_index(x_coordinates, x0, "cell x0")
        column_end = coordinate_index(x_coordinates, x1, "cell x1")
        row = coordinate_index(y_coordinates, top, "cell top")
        row_end = coordinate_index(y_coordinates, bottom, "cell bottom")
        if column_end <= column or row_end <= row:
            raise ExtractionError("ruled grid has an empty cell")
        identity = (row, column, row_end, column_end)
        if identity in seen:
            continue
        seen.add(identity)
        for covered_row in range(row, row_end):
            for covered_column in range(column, column_end):
                key = (covered_row, covered_column)
                if key in coverage:
                    raise ExtractionError("ruled grid contains overlapping cells")
                coverage[key] = identity
        selected_words = words_in_bbox(words, cell)
        text, has_words = cell_text(selected_words)
        cells.append({
            "row": row,
            "column": column,
            "rowspan": row_end - row,
            "colspan": column_end - column,
            "text": text,
            "bbox": [rounded(x0), rounded(top), rounded(x1), rounded(bottom)],
            "confidence": 1.0 if has_words else 0.99,
            "confidenceReasons": ["ruled-grid", "center-contained-words" if has_words else "empty-ruled-cell"],
        })
    expected_coverage = {(row, column) for row in range(len(y_coordinates) - 1) for column in range(expected_columns)}
    if set(coverage) != expected_coverage:
        missing = sorted(expected_coverage - set(coverage))
        raise ExtractionError(f"ruled grid has gaps at {missing[:8]}")
    cells.sort(key=lambda cell: (cell["row"], cell["column"], cell["rowspan"], cell["colspan"]))
    return cells, x_coordinates, y_coordinates


def header_signature(cells: list[dict[str, Any]], header_rows: int) -> list[dict[str, Any]]:
    return [
        {
            "row": cell["row"],
            "column": cell["column"],
            "rowspan": cell["rowspan"],
            "colspan": cell["colspan"],
            "text": canonical_text(cell["text"]),
        }
        for cell in cells
        if cell["row"] < header_rows
    ]


def labels_from_header(cells: list[dict[str, Any]], header_rows: int, columns: int) -> list[str]:
    labels = []
    for column in range(columns):
        components = []
        for row in range(header_rows):
            matching = [
                cell for cell in cells
                if cell["row"] <= row < cell["row"] + cell["rowspan"]
                and cell["column"] <= column < cell["column"] + cell["colspan"]
                and canonical_text(cell["text"])
            ]
            if len(matching) != 1:
                raise ExtractionError(f"header cell for row {row + 1}, column {column + 1} is ambiguous")
            text = canonical_text(matching[0]["text"])
            if not components or components[-1] != text:
                components.append(text)
        labels.append(" / ".join(components))
    if len(set(labels)) != len(labels):
        raise ExtractionError("flattened header labels are not unique")
    return labels


def data_rows(cells: list[dict[str, Any]], header_rows: int, columns: int, page_number: int) -> list[dict[str, Any]]:
    by_row: dict[int, list[dict[str, Any]]] = {}
    for cell in cells:
        if cell["row"] >= header_rows:
            by_row.setdefault(cell["row"], []).append(cell)
    rows: list[dict[str, Any]] = []
    for source_row, row_cells in sorted(by_row.items()):
        row_cells.sort(key=lambda cell: cell["column"])
        if len(row_cells) != columns or any(cell["rowspan"] != 1 or cell["colspan"] != 1 for cell in row_cells):
            raise ExtractionError(f"data row {source_row + 1} on page {page_number} is not a rectangular {columns}-column row")
        if [cell["column"] for cell in row_cells] != list(range(columns)):
            raise ExtractionError(f"data row {source_row + 1} on page {page_number} has non-contiguous columns")
        if any(not canonical_text(cell["text"]) for cell in row_cells):
            raise ExtractionError(f"data row {source_row + 1} on page {page_number} contains an empty cell")
        rows.append({"page": page_number, "sourceRow": source_row + 1, "cells": row_cells})
    if not rows:
        raise ExtractionError(f"table on page {page_number} has no complete data rows")
    return rows


def nearby_footnotes(lines: list[dict[str, Any]], table_bbox: tuple[float, float, float, float], prefix: str, max_gap: float) -> list[dict[str, Any]]:
    _, _, _, table_bottom = table_bbox
    records = []
    for line in lines:
        _, top, _, _ = line["bbox"]
        text = canonical_text(line["text"])
        if top < table_bottom + COORDINATE_TOLERANCE or top > table_bottom + max_gap:
            continue
        if text.startswith(prefix):
            records.append({
                "text": text,
                "bbox": line["bbox"],
                "confidence": 1.0,
                "confidenceReasons": ["adjacent-to-ruled-grid", "explicit-prefix"],
            })
    return records


def atomic_write(path: Path, payload: str) -> Path:
    if path.exists() or path.is_symlink():
        raise ExtractionError(f"refuses to overwrite existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp-{os.getpid()}-{secrets.token_hex(8)}")
    try:
        with temporary.open("x", encoding="utf-8", newline="") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    return path


def csv_payload(labels: list[str], rows: list[dict[str, Any]]) -> str:
    from io import StringIO

    stream = StringIO(newline="")
    writer = csv.writer(stream, lineterminator="\n")
    writer.writerow(["page", "sourceRow", *labels])
    for row in rows:
        writer.writerow([row["page"], row["sourceRow"], *(cell["text"] for cell in row["cells"])])
    return stream.getvalue()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--table-title", required=True, help="Exact leading title text repeated above every table segment")
    parser.add_argument("--expected-columns", type=int, required=True)
    parser.add_argument("--header-rows", type=int, default=2)
    parser.add_argument("--min-pages", type=int, default=2)
    parser.add_argument("--footnote-prefix", help="Require and retain at least one nearby footnote beginning with this exact prefix")
    parser.add_argument("--max-footnote-gap", type=float, default=90)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--csv-output", type=Path, required=True)
    parser.add_argument("--max-pages", type=int, default=200)
    parser.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    return parser.parse_args()


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parse_args()
    try:
        import pdfplumber
    except ImportError:
        print(json.dumps({"ok": False, "error": "pdfplumber is not installed", "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
        return 2
    try:
        # The workflow already enforces this at its boundary.  Keep the
        # published primitive equally strict so a direct caller cannot bind an
        # audit to a path whose identity changes through a symlink.
        source_argument = args.input.expanduser()
        if source_argument.is_symlink() or not source_argument.is_file():
            raise ExtractionError("input must be an existing regular non-symlink PDF")
        source = source_argument.resolve()
        output = args.output.expanduser().resolve()
        csv_output = args.csv_output.expanduser().resolve()
        if len({source, output, csv_output}) != 3:
            raise ExtractionError("input, JSON output, and CSV output must be distinct paths")
        if args.expected_columns < 2 or args.header_rows < 1 or args.min_pages < 2:
            raise ExtractionError("expected-columns must be >= 2, header-rows >= 1, and min-pages >= 2")
        if args.max_pages < 1 or args.max_bytes < 1 or args.max_footnote_gap < 0:
            raise ExtractionError("all resource limits must be positive")
        if source.stat().st_size > args.max_bytes:
            raise ExtractionError(f"PDF is {source.stat().st_size} bytes; max-bytes is {args.max_bytes}")

        segments: list[dict[str, Any]] = []
        all_data_rows: list[dict[str, Any]] = []
        expected_header: list[dict[str, Any]] | None = None
        expected_columns: list[float] | None = None
        column_labels: list[str] | None = None
        footnotes: list[dict[str, Any]] = []
        with pdfplumber.open(str(source)) as document:
            if len(document.pages) > args.max_pages:
                raise ExtractionError(f"PDF has {len(document.pages)} pages; max-pages is {args.max_pages}")
            for page_number, page in enumerate(document.pages, 1):
                words = page.extract_words() or []
                lines = word_lines(words)
                if not title_lines(lines, args.table_title):
                    continue
                table, title_line = candidate_for_title(page, lines, args.table_title)
                cells, x_coordinates, y_coordinates = extract_grid_cells(table, words, args.expected_columns)
                # A segment is its own physical page, but each emitted cell
                # carries that identity as well. Consumers must never infer a
                # cell's location from its container when they join segments
                # or flatten the table for audit/review.
                for cell in cells:
                    cell["page"] = page_number
                if len(y_coordinates) - 1 <= args.header_rows:
                    raise ExtractionError(f"table on page {page_number} has no data row after its header")
                signature = header_signature(cells, args.header_rows)
                if not signature:
                    raise ExtractionError(f"table on page {page_number} has no header cells")
                labels = labels_from_header(cells, args.header_rows, args.expected_columns)
                rows = data_rows(cells, args.header_rows, args.expected_columns, page_number)
                if expected_header is None:
                    expected_header = signature
                    expected_columns = x_coordinates
                    column_labels = labels
                else:
                    if signature != expected_header:
                        raise ExtractionError(f"repeated header on page {page_number} differs from the first table segment")
                    if labels != column_labels:
                        raise ExtractionError(f"flattened header labels on page {page_number} differ from the first table segment")
                    if len(x_coordinates) != len(expected_columns or []) or any(not close(left, right) for left, right in zip(x_coordinates, expected_columns or [])):
                        raise ExtractionError(f"column boundaries on page {page_number} differ from the first table segment")
                table_bbox = tuple(float(value) for value in table.bbox)
                segment_footnotes = nearby_footnotes(lines, table_bbox, args.footnote_prefix, args.max_footnote_gap) if args.footnote_prefix else []
                for record in segment_footnotes:
                    record["page"] = page_number
                footnotes.extend(segment_footnotes)
                segments.append({
                    "page": page_number,
                    "pageSize": {"width": rounded(page.width), "height": rounded(page.height)},
                    "title": {"text": title_line["text"], "bbox": title_line["bbox"]},
                    "tableBBox": [rounded(value) for value in table_bbox],
                    "columns": [{"index": index + 1, "x0": rounded(x_coordinates[index]), "x1": rounded(x_coordinates[index + 1]), "label": labels[index]} for index in range(args.expected_columns)],
                    "headerCells": [cell for cell in cells if cell["row"] < args.header_rows],
                    "dataRows": rows,
                    "footnotes": segment_footnotes,
                })
                all_data_rows.extend(rows)
        pages = [segment["page"] for segment in segments]
        if len(segments) < args.min_pages:
            raise ExtractionError(f"found {len(segments)} titled ruled-table segments; min-pages is {args.min_pages}")
        if pages != list(range(pages[0], pages[0] + len(pages))):
            raise ExtractionError(f"titled ruled-table pages must be consecutive; found {pages}")
        if args.footnote_prefix and not footnotes:
            raise ExtractionError(f"expected a nearby footnote beginning with {args.footnote_prefix!r}")
        if expected_header is None or expected_columns is None or column_labels is None:
            raise ExtractionError("no ruled-table segment was selected")

        payload = {
            "schema": SCHEMA,
            "provider": {"actual": "pdfplumber", "version": str(pdfplumber.__version__), "silentFallback": False},
            "savePolicy": {"strategy": "read-only", "sourceOverwrite": False},
            "source": file_evidence(source),
            "operation": {
                "type": "extract-ruled-cross-page-table",
                "profile": PROFILE,
                "tableTitle": args.table_title,
                "expectedColumns": args.expected_columns,
                "headerRows": args.header_rows,
                "minimumPages": args.min_pages,
            },
            "table": {
                "title": args.table_title,
                "pageRange": pages,
                "columnLabels": column_labels,
                "header": expected_header,
                "segments": segments,
                "dataRows": all_data_rows,
                "footnotes": footnotes,
            },
            "validation": {
                "passed": True,
                "checks": [
                    {"id": "explicit-title", "passed": True, "pages": pages},
                    {"id": "consecutive-pages", "passed": True, "pages": pages},
                    {"id": "fixed-column-boundaries", "passed": True, "columns": args.expected_columns},
                    {"id": "repeated-header", "passed": True, "headerRows": args.header_rows},
                    {"id": "rectangular-data-rows", "passed": True, "rows": len(all_data_rows)},
                    {"id": "footnote-prefix", "passed": bool(footnotes) if args.footnote_prefix else None, "prefix": args.footnote_prefix},
                ],
                "warnings": [],
            },
        }
        json_payload = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
        csv_text = csv_payload(column_labels, all_data_rows)
        # Build both payloads before publishing either.  A failed validation or
        # serialization never leaves a partial delivery behind.
        atomic_write(output, json_payload)
        try:
            atomic_write(csv_output, csv_text)
        except Exception:
            output.unlink(missing_ok=True)
            raise
        print(json.dumps({"ok": True, "schema": SCHEMA, "output": file_evidence(output), "csvOutput": file_evidence(csv_output), "source": payload["source"], "silentFallback": False}, indent=2, sort_keys=True))
        return 0
    except (ExtractionError, OSError, ValueError) as error:
        print(json.dumps({"ok": False, "error": str(error), "provider": "pdfplumber", "profile": PROFILE, "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
