#!/usr/bin/env python3
"""Validate examples/case-* structure. Empty library is valid."""
from __future__ import annotations

import argparse
import json
from pathlib import Path


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
        case = {
            "name": p.name,
            "source_doc_or_docx": has_source_doc or has_source_docx,
            "source_doc": has_source_doc,
            "source_docx": has_source_docx,
            "ai_output_xlsx": (p / "ai-output.xlsx").exists(),
            "corrected_xlsx": (p / "corrected.xlsx").exists(),
            "notes_md": (p / "notes.md").exists(),
        }
        required_keys = ["source_doc_or_docx", "ai_output_xlsx", "corrected_xlsx", "notes_md"]
        missing = [k for k in required_keys if not case.get(k)]
        if missing:
            report["warnings"].append({p.name: {"missing": missing}})
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
