#!/usr/bin/env python3
"""Structural safety checker for Excel 2019-compatible XLSX templates.

Checks for XML parse errors, risky drawing/control parts, disabled gridlines,
and duplicate/overlapping merged ranges. It is a fast structural gate, not a
replacement for opening the file in Excel and visual print-preview QA.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

RISKY_PART_PATTERNS = [
    "xl/drawings/",
    "xl/ctrlProps/",
    "xl/activeX/",
    "xl/media/",
    "vmlDrawing",
]
MERGE_REF_RE = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$")
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_to_int(col: str) -> int:
    total = 0
    for ch in col:
        total = total * 26 + (ord(ch) - 64)
    return total


@dataclass(frozen=True)
class Rect:
    ref: str
    c1: int
    r1: int
    c2: int
    r2: int

    def overlaps(self, other: "Rect") -> bool:
        return not (self.c2 < other.c1 or other.c2 < self.c1 or self.r2 < other.r1 or other.r2 < self.r1)


def parse_merge_ref(ref: str) -> Rect | None:
    m = MERGE_REF_RE.match(ref)
    if not m:
        return None
    c1, r1, c2, r2 = m.groups()
    a, b = col_to_int(c1), col_to_int(c2)
    x, y = int(r1), int(r2)
    if a > b:
        a, b = b, a
    if x > y:
        x, y = y, x
    return Rect(ref, a, x, b, y)


def validate(path: Path) -> dict:
    result = {
        "file": str(path),
        "ok": True,
        "xml_errors": [],
        "risky_parts": [],
        "merge_errors": [],
        "gridline_warnings": [],
        "worksheet_count": 0,
    }
    if not path.exists():
        return {**result, "ok": False, "xml_errors": ["file not found"]}
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            for name in names:
                lower = name.lower()
                if any(pattern.lower() in lower for pattern in RISKY_PART_PATTERNS):
                    result["risky_parts"].append(name)
                if name.endswith(".xml"):
                    try:
                        ET.fromstring(zf.read(name))
                    except ET.ParseError as exc:
                        result["xml_errors"].append(f"{name}: {exc}")
            sheet_names = [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
            result["worksheet_count"] = len(sheet_names)
            for sheet_name in sheet_names:
                root = ET.fromstring(zf.read(sheet_name))
                for sv in root.findall(".//main:sheetView", NS):
                    if sv.attrib.get("showGridLines") == "0":
                        result["gridline_warnings"].append(f"{sheet_name}: showGridLines=0")
                rects: list[Rect] = []
                seen: set[str] = set()
                for mc in root.findall(".//main:mergeCell", NS):
                    ref = mc.attrib.get("ref", "")
                    if ref in seen:
                        result["merge_errors"].append(f"{sheet_name}: duplicate merge {ref}")
                        continue
                    seen.add(ref)
                    rect = parse_merge_ref(ref)
                    if rect is None:
                        result["merge_errors"].append(f"{sheet_name}: invalid merge ref {ref}")
                        continue
                    for prev in rects:
                        if rect.overlaps(prev):
                            result["merge_errors"].append(f"{sheet_name}: overlapping merges {prev.ref} and {rect.ref}")
                    rects.append(rect)
    except zipfile.BadZipFile as exc:
        result["xml_errors"].append(f"not a valid XLSX ZIP: {exc}")

    result["ok"] = not (result["xml_errors"] or result["risky_parts"] or result["merge_errors"] or result["gridline_warnings"])
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx", type=Path)
    parser.add_argument("--json", action="store_true", help="Print full JSON result")
    args = parser.parse_args()
    data = validate(args.xlsx)
    if args.json or not data["ok"]:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        print(f"OK: {args.xlsx} passed structural XLSX checks")
    return 0 if data["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
