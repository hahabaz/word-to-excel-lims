#!/usr/bin/env python3
"""Build a baseline layout_plan.json from DOCX analysis.

The plan can be manually or model-refined before generating XLSX.
The examples directory is optional and may be empty.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def scan_examples(examples_dir: Path) -> dict:
    if not examples_dir.exists():
        return {"status": "empty", "selected_case": None, "available_cases": [], "notes": "examples directory not found"}
    cases = []
    for p in sorted(examples_dir.iterdir()):
        if not p.is_dir() or p.name.startswith("_"):
            continue
        if not p.name.startswith("case-"):
            continue
        notes_path = p / "notes.md"
        source_doc = (p / "source.doc").exists()
        source_docx = (p / "source.docx").exists()
        cases.append({
            "name": p.name,
            "has_source_doc": source_doc,
            "has_source_docx": source_docx,
            "has_source": source_doc or source_docx,
            "has_ai_output": (p / "ai-output.xlsx").exists(),
            "has_corrected": (p / "corrected.xlsx").exists(),
            "has_notes": notes_path.exists(),
            "notes_excerpt": notes_path.read_text(encoding="utf-8")[:900] if notes_path.exists() else "",
        })
    if not cases:
        return {"status": "empty", "selected_case": None, "available_cases": [], "notes": "No real case-* folders. Use default layout heuristics."}
    # Baseline selector: pick the first complete case. A model can override this.
    complete = [c for c in cases if c["has_corrected"] and c["has_notes"]]
    selected = complete[0]["name"] if complete else None
    return {
        "status": "used" if selected else "available_not_used",
        "selected_case": selected,
        "available_cases": cases,
        "notes": "Baseline selector used first complete case. Refine manually for semantic similarity if needed.",
    }


def choose_base_columns(analysis: dict) -> int:
    """Choose the smallest clean equal-width grid.

    Earlier versions escalated to 56/64 columns too quickly. That made portrait
    one-page LIMS forms visually noisy and harder to keep to one print page.
    Use high column counts only when the source is truly extreme or a selected
    corrected case proves that density is desirable.
    """
    max_cols = int(analysis.get("max_table_columns") or 0)
    total_cells = int(analysis.get("total_table_cells") or 0)
    field_like = int(analysis.get("field_like_paragraphs") or 0)
    complexity = analysis.get("complexity", "low")
    orientation = analysis.get("page_setup", {}).get("orientation", "portrait")
    estimated_pages = analysis.get("estimated_word_pages")

    # A012-style compact path: one-page portrait forms should not default to
    # 56/64 columns merely because a multi-level header has many leaf columns.
    if orientation == "portrait" and (estimated_pages in (None, 1)):
        if max_cols >= 18 or total_cells >= 260:
            return 40
        if max_cols >= 10 or total_cells >= 120 or complexity in {"medium", "high"}:
            return 32
        if field_like >= 5:
            return 28
        return 24

    if max_cols >= 20 or total_cells >= 320:
        return 56
    if max_cols >= 14 or total_cells >= 200:
        return 40
    if max_cols >= 10 or total_cells >= 120 or complexity == "high":
        return 32
    if max_cols >= 6 or field_like >= 10 or complexity == "medium":
        return 28
    if field_like >= 5:
        return 24
    return 20


def equal_width_for(base_columns: int, orientation: str) -> float:
    # openpyxl column width units are not exact printable units. These defaults
    # are intentionally modest so the content grid can fit one printed page.
    if orientation == "landscape":
        if base_columns >= 56:
            return 2.0
        if base_columns >= 40:
            return 2.4
        return 3.0
    if base_columns >= 56:
        return 1.55
    if base_columns >= 40:
        return 1.9
    if base_columns >= 32:
        return 2.35
    if base_columns >= 24:
        return 3.0
    return 3.6


def cell(start: int, span: int, text: str, border: str = "box", style: str = "normal", align: str = "left", wrap: bool = True) -> dict:
    return {
        "start_col": start,
        "colspan": max(1, span),
        "text": text or "",
        "border": border,
        "style": style,
        "align": align,
        "valign": "center",
        "wrap": wrap,
    }


def distribute_spans(n: int, base_columns: int) -> list[tuple[int, int]]:
    n = max(1, n)
    spans = []
    start = 1
    for i in range(n):
        end = round((i + 1) * base_columns / n)
        span = max(1, end - start + 1)
        spans.append((start, span))
        start = end + 1
    # Correct any rounding overshoot/undershoot.
    if spans:
        last_start, _ = spans[-1]
        spans[-1] = (last_start, base_columns - last_start + 1)
    return spans


def paragraph_section(text: str, base_columns: int, index: int) -> dict:
    is_title = index == 0 or (len(text) <= 40 and not any(mark in text for mark in ["：", ":"]))
    if is_title:
        return {
            "type": "title",
            "rows": [{"height": 26, "cells": [cell(1, base_columns, text, border="bottom", style="title", align="center", wrap=False)]}],
        }
    # Try field-like split for Chinese/English colon labels.
    m = re.match(r"^(.{1,18}?[：:])\s*(.*)$", text)
    if m:
        label, value = m.groups()
        label_span = min(max(4, math.ceil(len(label) / 3)), max(6, base_columns // 3))
        return {
            "type": "field_row",
            "rows": [{"height": 22, "cells": [
                cell(1, label_span, label, border="box", style="label", align="center", wrap=False),
                cell(label_span + 1, base_columns - label_span, value, border="box", style="normal", align="left", wrap=True),
            ]}],
        }
    return {
        "type": "text",
        "rows": [{"height": 32 if len(text) > 80 else 22, "cells": [cell(1, base_columns, text, border="box", style="normal", align="left", wrap=True)]}],
    }


def table_section(table: dict, base_columns: int) -> dict:
    rows = []
    max_cols = max(1, int(table.get("max_columns") or 1))
    spans = distribute_spans(max_cols, base_columns)
    for ridx, row in enumerate(table.get("rows", [])):
        cells = []
        for cidx in range(max_cols):
            start, span = spans[cidx]
            text = row[cidx] if cidx < len(row) else ""
            cells.append(cell(start, span, text, border="grid", style="table_header" if ridx == 0 else "normal", align="center", wrap=True))
        rows.append({"height": 24 if ridx == 0 else 22, "cells": cells})
    return {"type": "table", "rows": rows}


def build_plan(analysis: dict, examples_dir: Path) -> dict:
    case_library = scan_examples(examples_dir)
    base_columns = choose_base_columns(analysis)
    orientation = analysis.get("page_setup", {}).get("orientation", "portrait")
    page_setup = {
        "paper_size": "A4",
        "orientation": orientation,
        "horizontal_centered": True,
        "fit_to_width": 1,
        "fit_to_height": 1 if analysis.get("estimated_word_pages") == 1 else None,
        "margins": {"left": 0.35, "right": 0.35, "top": 0.45, "bottom": 0.45, "header": 0.2, "footer": 0.2},
        "print_area": None,
    }

    sections = []
    para_idx = 0
    for block in analysis.get("blocks", []):
        if block.get("type") == "paragraph":
            sections.append(paragraph_section(block.get("text", ""), base_columns, para_idx))
            para_idx += 1
        elif block.get("type") == "table":
            sections.append(table_section(block, base_columns))

    if not sections:
        sections.append({
            "type": "placeholder",
            "rows": [{"height": 24, "cells": [cell(1, base_columns, "", border="box", style="normal")]}],
        })

    summary = f"{analysis.get('paragraph_count', 0)} paragraphs, {analysis.get('table_count', 0)} tables, max table columns {analysis.get('max_table_columns', 0)}."
    return {
        "schema_version": "1.0",
        "source": {
            "docx_path": analysis.get("docx_path", ""),
            "estimated_word_pages": None,
            "summary": summary,
            "complexity": analysis.get("complexity", "unknown"),
            "limitations": analysis.get("limitations", []),
        },
        "case_library": case_library,
        "workbook": {
            "creator": "word-to-excel-lims skill",
            "sheets": [{
                "name": "Sheet1",
                "page_setup": page_setup,
                "base_grid": {
                    "base_columns": base_columns,
                    "equal_column_width": equal_width_for(base_columns, orientation),
                    "start_column": 1,
                    "default_row_height": 18,
                },
                "sections": sections,
                "qa": {
                    "equal_width_columns": True,
                    "gridlines_should_remain_visible": True,
                    "avoid_drawings_controls_textboxes": True,
                    "case_library_can_be_empty": True,
                },
            }],
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build layout_plan.json from word_analysis.json")
    parser.add_argument("analysis_json", type=Path)
    parser.add_argument("--examples", type=Path, default=Path("examples"))
    parser.add_argument("--out", type=Path, default=Path("layout_plan.json"))
    args = parser.parse_args(argv)
    try:
        analysis = load_json(args.analysis_json)
        plan = build_plan(analysis, args.examples)
        args.out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(f"Wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
