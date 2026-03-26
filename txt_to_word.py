from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import md_to_docx
import txt_to_md


DEFAULT_INPUT = txt_to_md.DEFAULT_INPUT
DEFAULT_MD_OUTPUT = txt_to_md.DEFAULT_OUTPUT
DEFAULT_DOCX_OUTPUT = md_to_docx.DEFAULT_OUTPUT


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert TXT files to DOCX in one step."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="TXT file or directory to convert. Defaults to 原始文本.",
    )
    parser.add_argument(
        "--md-output",
        default=str(DEFAULT_MD_OUTPUT),
        help="Intermediate Markdown output path or directory. Defaults to output_md.",
    )
    parser.add_argument(
        "--docx-output",
        default=str(DEFAULT_DOCX_OUTPUT),
        help="Final DOCX output path or directory. Defaults to output_docx.",
    )
    parser.add_argument(
        "--reference-doc",
        help="Optional reference DOCX used by pandoc to copy styles.",
    )
    return parser.parse_args()


def build_docx_output_path(
    source: Path, input_path: Path, output_path: Path
) -> Path:
    if input_path.is_file():
        if output_path.suffix.lower() == ".docx":
            return output_path
        return output_path / f"{source.stem}.docx"

    if output_path.suffix.lower() == ".docx":
        raise ValueError(
            "Directory input requires a DOCX output directory, not a single .docx"
        )

    relative_path = source.relative_to(input_path).with_suffix(".docx")
    return output_path / relative_path


def main() -> int:
    args = parse_args()
    pandoc_path = shutil.which("pandoc")
    if pandoc_path is None:
        print("[ERROR] pandoc is not installed or not in PATH.", file=sys.stderr)
        return 1

    input_path = Path(args.input_path)
    md_output = Path(args.md_output)
    docx_output = Path(args.docx_output)
    reference_doc = Path(args.reference_doc) if args.reference_doc else None

    if reference_doc is not None and not reference_doc.is_file():
        print(f"[ERROR] Reference DOCX does not exist: {reference_doc}", file=sys.stderr)
        return 1

    try:
        txt_files = txt_to_md.collect_input_files(input_path)
        for source in txt_files:
            md_destination = txt_to_md.build_output_path(source, input_path, md_output)
            docx_destination = build_docx_output_path(source, input_path, docx_output)

            txt_to_md.convert_file(source, md_destination)
            md_to_docx.convert_file(
                pandoc_path, md_destination, docx_destination, reference_doc
            )
            print(f"[OK] {source} -> {md_destination} -> {docx_destination}")
    except (RuntimeError, ValueError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
