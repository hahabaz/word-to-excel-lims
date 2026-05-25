#!/usr/bin/env python3
"""Analyze visible structure in a DOCX file for Word -> Excel LIMS conversion.

This script intentionally avoids copying Word formatting. It extracts stable
signals for layout planning: visible text, table dimensions, checkboxes, and
complexity metrics. It uses only the Python standard library.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
CHECKBOX_CHARS = {
    "☐", "☑", "☒", "■", "□",
}
WINGDINGS_CHECKBOX_CODES = {"F0A3", "F0FE", "F0FC", "F0A8", "2610", "2611", "2612"}


def clean_text(value: str) -> str:
    value = ILLEGAL_XML_RE.sub("", value)
    for ch in CHECKBOX_CHARS:
        value = value.replace(ch, "□")
    value = value.replace("\u200b", "").replace("\ufeff", "")
    value = re.sub(r"[ \t]+", " ", value)
    return value.strip()


def paragraph_text(p: ET.Element) -> str:
    pieces: list[str] = []
    for node in p.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "t" and node.text:
            pieces.append(node.text)
        elif tag == "tab":
            pieces.append(" ")
        elif tag in {"br", "cr"}:
            pieces.append("\n")
        elif tag == "sym":
            code = node.attrib.get(f"{{{NS['w']}}}char", "").upper()
            if code in WINGDINGS_CHECKBOX_CODES:
                pieces.append("□")
    return clean_text("".join(pieces))


def cell_text(tc: ET.Element) -> str:
    paras = [paragraph_text(p) for p in tc.findall(".//w:p", NS)]
    return clean_text("\n".join([p for p in paras if p]))


def get_page_setup(root: ET.Element) -> dict:
    setup: dict = {"paper_size": "A4", "orientation": "portrait", "source": "default"}
    pg_sz = root.find(".//w:sectPr/w:pgSz", NS)
    if pg_sz is not None:
        orient = pg_sz.attrib.get(f"{{{NS['w']}}}orient")
        if orient in {"portrait", "landscape"}:
            setup["orientation"] = orient
            setup["source"] = "docx"
        w = pg_sz.attrib.get(f"{{{NS['w']}}}w")
        h = pg_sz.attrib.get(f"{{{NS['w']}}}h")
        if w and h:
            setup["word_twips"] = {"w": int(w), "h": int(h)}
    margins = root.find(".//w:sectPr/w:pgMar", NS)
    if margins is not None:
        setup["margins_twips"] = {k.rsplit('}', 1)[-1]: int(v) for k, v in margins.attrib.items() if str(v).isdigit()}
    return setup


def analyze_docx(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(path)
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("DOCX has no word/document.xml body")

    blocks: list[dict] = []
    paragraphs: list[dict] = []
    tables: list[dict] = []

    for child in body:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            text = paragraph_text(child)
            if text:
                item = {"text": text, "length": len(text), "checkboxes": text.count("□")}
                paragraphs.append(item)
                blocks.append({"type": "paragraph", **item})
        elif tag == "tbl":
            rows = []
            for tr in child.findall("w:tr", NS):
                cells = [cell_text(tc) for tc in tr.findall("w:tc", NS)]
                if cells:
                    rows.append(cells)
            if rows:
                max_cols = max(len(r) for r in rows)
                table = {
                    "rows": rows,
                    "row_count": len(rows),
                    "max_columns": max_cols,
                    "non_empty_cells": sum(1 for r in rows for c in r if c),
                    "checkboxes": sum(c.count("□") for r in rows for c in r),
                }
                tables.append(table)
                blocks.append({"type": "table", **table})

    max_table_cols = max([t["max_columns"] for t in tables], default=0)
    total_cells = sum(t["row_count"] * t["max_columns"] for t in tables)
    dense_table = max_table_cols >= 8 or total_cells >= 80
    paragraph_chars = sum(p["length"] for p in paragraphs)
    field_like = sum(1 for p in paragraphs if any(x in p["text"] for x in ["：", ":", "□", "签字", "日期", "编号"]))

    if max_table_cols >= 12 or total_cells >= 160:
        complexity = "high"
    elif dense_table or field_like >= 8:
        complexity = "medium"
    else:
        complexity = "low"

    return {
        "docx_path": str(path),
        "page_setup": get_page_setup(root),
        "paragraph_count": len(paragraphs),
        "table_count": len(tables),
        "max_table_columns": max_table_cols,
        "total_table_cells": total_cells,
        "paragraph_chars": paragraph_chars,
        "checkboxes": sum(p["checkboxes"] for p in paragraphs) + sum(t["checkboxes"] for t in tables),
        "field_like_paragraphs": field_like,
        "complexity": complexity,
        "blocks": blocks,
        "paragraphs": paragraphs,
        "tables": tables,
        "limitations": [
            "This script extracts document XML structure but does not render Word print preview.",
            "Exact Word page count may require Word/LibreOffice rendering outside this script."
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Analyze DOCX structure for Word -> Excel LIMS conversion")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, default=None)
    args = parser.parse_args(argv)
    try:
        data = analyze_docx(args.docx)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if args.out:
        args.out.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
