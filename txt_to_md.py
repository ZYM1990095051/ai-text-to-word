from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


DEFAULT_INPUT = Path("原始文本")
DEFAULT_OUTPUT = Path("output_md")
ENCODINGS = (
    "utf-8-sig",
    "utf-8",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "gb18030",
    "gbk",
)
INLINE_MATH_RE = re.compile(r"\\\((.+?)\\\)")
REMOVABLE_CHARS = ("\ufeff", "\u200b", "\u200c", "\u200d", "\u2060")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert TXT files into normalized Markdown files."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="TXT file or directory to convert. Defaults to 原始文本.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="Output .md file or directory. Defaults to output_md.",
    )
    return parser.parse_args()


def read_text(path: Path) -> str:
    data = path.read_bytes()

    for encoding in ENCODINGS:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue

    return data.decode("utf-8", errors="replace")


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u00a0", " ")

    for char in REMOVABLE_CHARS:
        text = text.replace(char, "")

    normalized_lines: list[str] = []
    for line in text.split("\n"):
        stripped = line.strip()
        indent = line[: len(line) - len(line.lstrip(" \t"))]

        if stripped in {r"\[", r"\]"}:
            normalized_lines.append(f"{indent}$$")
            continue

        same_line_display = re.fullmatch(r"([ \t]*)\\\[\s*(.+?)\s*\\\][ \t]*", line)
        if same_line_display:
            same_line_indent = same_line_display.group(1)
            formula_body = same_line_display.group(2).strip()
            normalized_lines.extend(
                [
                    f"{same_line_indent}$$",
                    f"{same_line_indent}{formula_body}",
                    f"{same_line_indent}$$",
                ]
            )
            continue

        normalized_lines.append(line.rstrip())

    text = "\n".join(normalized_lines)
    text = INLINE_MATH_RE.sub(lambda match: f"${match.group(1).strip()}$", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
    return text


def collect_input_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".txt":
            raise ValueError(f"Only .txt files are supported: {input_path}")
        return [input_path]

    if input_path.is_dir():
        files = sorted(path for path in input_path.rglob("*.txt") if path.is_file())
        if not files:
            raise ValueError(f"No .txt files found under: {input_path}")
        return files

    raise ValueError(f"Input path does not exist: {input_path}")


def build_output_path(source: Path, input_path: Path, output_path: Path) -> Path:
    if input_path.is_file():
        if output_path.suffix.lower() == ".md":
            return output_path
        return output_path / f"{source.stem}.md"

    if output_path.suffix.lower() == ".md":
        raise ValueError("Directory input requires an output directory, not a single .md")

    relative_path = source.relative_to(input_path).with_suffix(".md")
    return output_path / relative_path


def convert_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    markdown = normalize_text(read_text(source))
    destination.write_text(markdown, encoding="utf-8")


def main() -> int:
    args = parse_args()
    input_path = Path(args.input_path)
    output_path = Path(args.output)

    try:
        files = collect_input_files(input_path)
        for source in files:
            destination = build_output_path(source, input_path, output_path)
            convert_file(source, destination)
            print(f"[OK] {source} -> {destination}")
    except ValueError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
