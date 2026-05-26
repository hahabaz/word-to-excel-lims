#!/usr/bin/env python3
"""Validate examples/case-* structure. Empty library is valid.

Also reports hidden rows/columns in corrected.xlsx because hidden helper rows in
human reference files can mislead future layout extraction. Hidden rows/columns
are warnings, not hard failures.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from zipfile import ZipFile


def inspect_xlsx_visibility(path: Path) -> dict:
    result = {"hidden_rows": 0, "hidden_cols": 0, "sheets": []}
    if not path.exists():
        return result
    try:
        with ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                result["zip_error"] = f"bad zip member: {bad}"
                return result
            for name in zf.namelist():
                if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
                    continue
                xml = zf.read(name).decode("utf-8", errors="ignore")
                hidden_rows = len(re.findall(r"<row[^>]*\bhidden=\"1\"", xml))
                hidden_cols = len(re.findall(r"<col[^>]*\bhidden=\"1\"", xml))
                result["hidden_rows"] += hidden_rows
                result["hidden_cols"] += hidden_cols
                result["sheets"].append({"xml": name, "hidden_rows": hidden_rows, "hidden_cols": hidden_cols})
    except Exception as exc:  # pragma: no cover
        result["zip_error"] = str(exc)
    return result


def validate(examples: Path) -> dict:
    report = {"ok": True, "status": "empty", "cases": [], "warnings": []}
    if not examples.exists():
        report["warnings"].append("examples directory does not exist; this is allowed but not recommended")
        return report
    for p in sorted(examples.iterdir()):
        if not p.is_dir() or p.name.startswith("_"):
            continue
        if not p.name.startswith("case-"):
            continue
        has_source_doc = (p / "source.doc").exists()
        has_source_docx = (p / "source.docx").exists()
        visibility = inspect_xlsx_visibility(p / "corrected.xlsx")
        case = {
            "name": p.name,
            "source_doc_or_docx": has_source_doc or has_source_docx,
            "source_doc": has_source_doc,
            "source_docx": has_source_docx,
            "ai_output_xlsx": (p / "ai-output.xlsx").exists(),
            "corrected_xlsx": (p / "corrected.xlsx").exists(),
            "notes_md": (p / "notes.md").exists(),
            "corrected_visibility": visibility,
        }
        required_keys = ["source_doc_or_docx", "ai_output_xlsx", "corrected_xlsx", "notes_md"]
        missing = [k for k in required_keys if not case.get(k)]
        if missing:
            report["warnings"].append({p.name: {"missing": missing}})
            report["ok"] = False
        if visibility.get("hidden_rows") or visibility.get("hidden_cols"):
            report["warnings"].append({
                p.name: {
                    "corrected_xlsx_hidden_rows": visibility.get("hidden_rows", 0),
                    "corrected_xlsx_hidden_cols": visibility.get("hidden_cols", 0),
                    "rule": "Do not treat hidden rows/columns as target output structure; use visible print layout."
                }
            })
        report["cases"].append(case)
    if report["cases"]:
        report["status"] = "has_cases"
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("examples", type=Path, nargs="?", default=Path("examples"))
    args = parser.parse_args()
    print(json.dumps(validate(args.examples), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
