#!/usr/bin/env python3
"""Run Word -> layout_plan.json -> XLSX -> validation pipeline.

This script works even when examples/ contains no real cases.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str]) -> None:
    print("+", " ".join(str(x) for x in cmd))
    subprocess.run(cmd, check=True)


def ensure_docx(input_path: Path, work_dir: Path) -> Path:
    """Return a DOCX path, converting old .doc files when possible.

    The analyzer reads DOCX package XML. Old binary .doc files must first be
    converted by LibreOffice/headless soffice. The original source.doc should
    still be stored in examples/case-* for case history.
    """
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return input_path
    if suffix != ".doc":
        raise ValueError(f"Unsupported Word extension {input_path.suffix!r}; expected .doc or .docx")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Old .doc input requires LibreOffice/soffice to convert to .docx, but it was not found on PATH")
    work_dir.mkdir(parents=True, exist_ok=True)
    run([soffice, "--headless", "--convert-to", "docx", "--outdir", str(work_dir), str(input_path)])
    converted = work_dir / f"{input_path.stem}.docx"
    if not converted.exists():
        matches = list(work_dir.glob("*.docx"))
        if matches:
            converted = matches[0]
    if not converted.exists():
        raise RuntimeError(f"LibreOffice conversion did not produce a DOCX in {work_dir}")
    return converted


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the word-to-excel-lims conversion pipeline")
    parser.add_argument("word", type=Path, help="Input Word file (.doc or .docx)")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--examples", type=Path, default=Path("examples"))
    parser.add_argument("--keep-going-on-validation-warning", action="store_true")
    args = parser.parse_args(argv)

    if not args.word.exists():
        print(f"ERROR: Word file not found: {args.word}", file=sys.stderr)
        return 2

    script_dir = Path(__file__).resolve().parent
    stem = args.out.with_suffix("")
    analysis = stem.with_suffix(".word_analysis.json")
    plan = stem.with_suffix(".layout_plan.json")
    report = stem.with_suffix(".validation_report.json")

    try:
        analysis_input = ensure_docx(args.word, stem.parent / f"{stem.name}_docx_work")
        run([sys.executable, str(script_dir / "analyze_docx.py"), str(analysis_input), "--out", str(analysis)])
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
