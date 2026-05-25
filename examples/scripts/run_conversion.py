#!/usr/bin/env python3
"""Run Word -> layout_plan.json -> XLSX -> validation pipeline.

This script works even when examples/ contains no real cases.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str]) -> None:
    print("+", " ".join(str(x) for x in cmd))
    subprocess.run(cmd, check=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the word-to-excel-lims conversion pipeline")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--examples", type=Path, default=Path("examples"))
    parser.add_argument("--keep-going-on-validation-warning", action="store_true")
    args = parser.parse_args(argv)

    if not args.docx.exists():
        print(f"ERROR: DOCX not found: {args.docx}", file=sys.stderr)
        return 2

    script_dir = Path(__file__).resolve().parent
    stem = args.out.with_suffix("")
    analysis = stem.with_suffix(".word_analysis.json")
    plan = stem.with_suffix(".layout_plan.json")
    report = stem.with_suffix(".validation_report.json")

    try:
        run([sys.executable, str(script_dir / "analyze_docx.py"), str(args.docx), "--out", str(analysis)])
        run([sys.executable, str(script_dir / "build_layout_plan.py"), str(analysis), "--examples", str(args.examples), "--out", str(plan)])
        run([sys.executable, str(script_dir / "generate_xlsx_from_plan.py"), str(plan), "--out", str(args.out)])
        run([sys.executable, str(script_dir / "validate_xlsx_structure.py"), str(args.out), "--json", str(report)])
    except subprocess.CalledProcessError as exc:
        print(f"ERROR: pipeline step failed with code {exc.returncode}", file=sys.stderr)
        return exc.returncode

    print("Done.")
    print(f"XLSX: {args.out}")
    print(f"Analysis: {analysis}")
    print(f"Layout plan: {plan}")
    print(f"Validation report: {report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
