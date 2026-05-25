#!/usr/bin/env python3
"""Structural safety checker for Excel 2019-compatible XLSX templates.

Checks XML parse errors, risky drawing/control/media parts, disabled gridlines,
duplicate/overlapping merged ranges, and basic print settings.
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


def validate_xlsx(path: Path) -> dict:
    result = {
        "path": str(path),
        "ok": True,
        "errors": [],
        "warnings": [],
        "summary": {},
    }
    if not path.exists():
        result["ok"] = False
        result["errors"].append("File does not exist")
        return result

    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            result["summary"]["part_count"] = len(names)
            risky = [name for name in names for pat in RISKY_PART_PATTERNS if pat in name]
            if risky:
                result["ok"] = False
                result["errors"].append({"risky_parts": sorted(set(risky))})

            xml_names = [n for n in names if n.endswith(".xml")]
            for name in xml_names:
                try:
                    ET.fromstring(zf.read(name))
                except Exception as exc:
                    result["ok"] = False
                    result["errors"].append({"xml_parse_error": name, "message": str(exc)})

            sheet_names = [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
            result["summary"]["worksheet_count"] = len(sheet_names)
            for sheet_name in sheet_names:
                root = ET.fromstring(zf.read(sheet_name))
                sheet_views = root.findall("main:sheetViews/main:sheetView", NS)
                for view in sheet_views:
                    if view.attrib.get("showGridLines") == "0":
                        result["ok"] = False
                        result["errors"].append({"gridlines_disabled": sheet_name})

                merge_refs = [mc.attrib.get("ref", "") for mc in root.findall(".//main:mergeCell", NS)]
                seen = set()
                rects: list[Rect] = []
                malformed = []
                duplicates = []
                overlaps = []
                for ref in merge_refs:
                    if ref in seen:
                        duplicates.append(ref)
                    seen.add(ref)
                    rect = parse_merge_ref(ref)
                    if rect is None:
                        malformed.append(ref)
                        continue
                    for other in rects:
                        if rect.overlaps(other):
                            overlaps.append((rect.ref, other.ref))
                    rects.append(rect)
                if malformed:
                    result["ok"] = False
                    result["errors"].append({"malformed_merges": sheet_name, "refs": malformed})
                if duplicates:
                    result["ok"] = False
                    result["errors"].append({"duplicate_merges": sheet_name, "refs": duplicates})
                if overlaps:
                    result["ok"] = False
                    result["errors"].append({"overlapping_merges": sheet_name, "refs": overlaps[:20]})

                page_setup = root.find("main:pageSetup", NS)
                if page_setup is None:
                    result["warnings"].append({"missing_page_setup": sheet_name})
                print_options = root.find("main:printOptions", NS)
                if print_options is None or print_options.attrib.get("horizontalCentered") != "1":
                    result["warnings"].append({"not_horizontally_centered_or_unknown": sheet_name})
    except zipfile.BadZipFile:
        result["ok"] = False
        result["errors"].append("Not a valid XLSX ZIP package")
    except Exception as exc:
        result["ok"] = False
        result["errors"].append(str(exc))
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate XLSX structural safety")
    parser.add_argument("xlsx", type=Path)
    parser.add_argument("--json", type=Path, default=None)
    args = parser.parse_args(argv)
    report = validate_xlsx(args.xlsx)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json:
        args.json.write_text(text, encoding="utf-8")
    print(text)
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
