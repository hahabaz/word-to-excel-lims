#!/usr/bin/env python3
"""Lightweight DOCX structure analyzer for Word -> Excel template conversion.

Outputs JSON with visible-text/table signals that help choose an equal-width
base-column grid. This script does not attempt pixel-perfect rendering.
LibreOffice/page rendering should still be used when available.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def clean_text(value: str) -> str:
    value = ILLEGAL_XML_RE.sub("", value)
    # Normalize common checkbox glyphs/control remnants to a plain empty box.
    value = value.replace("☐", "□").replace("☒", "□").replace("☑", "□")
    return re.sub(r"[ \t]+", " ", value).strip()


def paragraph_text(p: ET.Element) -> str:
    pieces: list[str] = []
    for node in p.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "t" and node.text:
            pieces.append(node.text)
        elif tag in {"tab"}:
            pieces.append(" ")
        elif tag in {"br", "cr"}:
            pieces.append("\n")
        elif tag == "sym":
            char = node.attrib.get(f"{{{NS['w']}}}char", "")
            # Wingdings checkbox variants often appear as F0A3/F0FE/etc.
            if char.upper() in {"F0A3", "F0FE", "F0FC", "F0A8", "2610", "2611", "2612"}:
                pieces.append("□")
    return clean_text("".join(pieces))


def cell_text(tc: ET.Element) -> str:
    paras = [paragraph_text(p) for p in tc.findall(".//w:p", NS)]
    return clean_text("\n".join([p for p in paras if p]))


def analyze_docx(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(path)
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
        root = ET.fromstring(xml)
        body = root.find("w:body", NS)
        paragraphs = []
        tables = []
        max_table_cols = 0
        if body is not None:
            for child in body:
                tag = child.tag.rsplit("}", 1)[-1]
                if tag == "p":
                    text = paragraph_text(child)
                    if text:
                        paragraphs.append(text)
                elif tag == "tbl":
                    rows = []
                    for tr in child.findall("w:tr", NS):
                        row = [cell_text(tc) for tc in tr.findall("w:tc", NS)]
                        rows.append(row)
                        max_table_cols = max(max_table_cols, len(row))
                    tables.append(rows)
        rels = []
        for name in zf.namelist():
            if name.startswith("word/") and ("drawing" in name.lower() or "activex" in name.lower() or "ctrl" in name.lower()):
                rels.append(name)
        section_count = len(root.findall(".//w:sectPr", NS))
    all_text = paragraphs + [cell for tbl in tables for row in tbl for cell in row]
    long_label_count = sum(1 for t in all_text if 8 <= len(t) <= 40 and any(ch in t for ch in "：:___（）()"))
    dense_score = max_table_cols * 2 + len(tables) * 3 + long_label_count // 3
    if dense_score <= 12:
        suggested_base_columns = 24
    elif dense_score <= 28:
        suggested_base_columns = 36
    elif dense_score <= 45:
        suggested_base_columns = 48
    else:
        suggested_base_columns = 64
    return {
        "file": str(path),
        "paragraph_count": len(paragraphs),
        "table_count": len(tables),
        "max_table_columns": max_table_cols,
        "section_count": section_count,
        "possible_drawing_or_control_parts": rels,
        "suggested_base_columns": suggested_base_columns,
        "sample_paragraphs": paragraphs[:20],
        "tables_preview": [tbl[:5] for tbl in tables[:5]],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, help="Write JSON analysis to this path")
    args = parser.parse_args()
    data = analyze_docx(args.docx)
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if args.out:
        args.out.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
