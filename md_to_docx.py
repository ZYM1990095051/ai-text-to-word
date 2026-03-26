from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


DEFAULT_INPUT = Path("output_md")
DEFAULT_OUTPUT = Path("output_docx")
ENCODINGS = (
    "utf-8-sig",
    "utf-8",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "gb18030",
    "gbk",
)
ARROW_REPLACEMENTS = (
    (r"\rightleftharpoons", "⇌"),
    (r"\leftrightharpoons", "⇌"),
    (r"\rightleftarrows", "⇄"),
    (r"\leftrightarrow", "↔"),
    (r"\rightarrow", "→"),
    (r"\leftarrow", "←"),
    (r"\to", "→"),
    ("<=>", "⇌"),
    ("<->", "↔"),
    ("->", "→"),
    ("<-", "←"),
)
SUBSCRIPT_MAP = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
SUPERSCRIPT_MAP = {
    "0": "⁰",
    "1": "¹",
    "2": "²",
    "3": "³",
    "4": "⁴",
    "5": "⁵",
    "6": "⁶",
    "7": "⁷",
    "8": "⁸",
    "9": "⁹",
    "+": "⁺",
    "-": "⁻",
}
SPECIAL_TEX_TEXT = {
    "{": r"\{",
    "}": r"\}",
    "%": r"\%",
    "&": r"\&",
    "#": r"\#",
    "_": r"\_",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Markdown files into DOCX with Office-style equations."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="Markdown file or directory to convert. Defaults to output_md.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="Output .docx file or directory. Defaults to output_docx.",
    )
    parser.add_argument(
        "--reference-doc",
        help="Optional reference DOCX used by pandoc to copy styles.",
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


def collect_input_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".md":
            raise ValueError(f"Only .md files are supported: {input_path}")
        return [input_path]

    if input_path.is_dir():
        files = sorted(path for path in input_path.rglob("*.md") if path.is_file())
        if not files:
            raise ValueError(f"No .md files found under: {input_path}")
        return files

    raise ValueError(f"Input path does not exist: {input_path}")


def build_output_path(source: Path, input_path: Path, output_path: Path) -> Path:
    if input_path.is_file():
        if output_path.suffix.lower() == ".docx":
            return output_path
        return output_path / f"{source.stem}.docx"

    if output_path.suffix.lower() == ".docx":
        raise ValueError(
            "Directory input requires an output directory, not a single .docx"
        )

    relative_path = source.relative_to(input_path).with_suffix(".docx")
    return output_path / relative_path


def escape_tex_text(text: str) -> str:
    escaped_chars: list[str] = []
    for char in text:
        if char == "\\":
            continue
        escaped_chars.append(SPECIAL_TEX_TEXT.get(char, char))
    return "".join(escaped_chars)


def to_superscript(text: str) -> str:
    return "".join(SUPERSCRIPT_MAP.get(char, char) for char in text)


def format_chemical_species(token: str) -> str:
    trailing = ""
    while token and token[-1] in ",.;:":
        trailing = token[-1] + trailing
        token = token[:-1]

    charge = ""
    caret_charge = re.search(r"\^(?:\{)?([0-9]*[+-]+)(?:\})?$", token)
    if caret_charge and any(char.isalpha() for char in token[: caret_charge.start()]):
        charge = caret_charge.group(1)
        token = token[: caret_charge.start()]
    else:
        plain_charge = re.search(r"([0-9]*[+-]+)$", token)
        if plain_charge:
            body = token[: plain_charge.start()]
            if body and any(char.isalpha() for char in body):
                charge = plain_charge.group(1)
                token = body

    formatted: list[str] = []
    subscript_mode = False
    for index, char in enumerate(token):
        prev_char = token[index - 1] if index > 0 else ""
        if char.isdigit() and (
            subscript_mode or prev_char.isalpha() or prev_char in ")]}"
        ):
            formatted.append(char.translate(SUBSCRIPT_MAP))
            subscript_mode = True
            continue

        formatted.append(char)
        subscript_mode = False

    if charge:
        formatted.append(to_superscript(charge))

    return "".join(formatted) + trailing


def split_ce_piece(piece: str) -> list[str]:
    operators = {"⇌", "⇄", "↔", "→", "←", "=", "+"}
    parts: list[str] = []
    buffer: list[str] = []

    for index, char in enumerate(piece):
        is_charge_sign = char == "+" and index == len(piece) - 1
        if char in operators and not is_charge_sign:
            if buffer:
                parts.append("".join(buffer))
                buffer = []
            parts.append(char)
            continue

        buffer.append(char)

    if buffer:
        parts.append("".join(buffer))

    return parts


def convert_ce_expression(content: str) -> str:
    text = " ".join(content.strip().split())
    for source, target in ARROW_REPLACEMENTS:
        text = text.replace(source, target)

    for operator in ("⇌", "⇄", "↔", "→", "←", "="):
        text = text.replace(operator, f" {operator} ")

    text = re.sub(r"\s+", " ", text).strip()
    converted_parts: list[str] = []

    for piece in re.split(r"(\s+)", text):
        if not piece:
            continue
        if piece.isspace():
            converted_parts.append(piece)
            continue

        for part in split_ce_piece(piece):
            if part in {"⇌", "⇄", "↔", "→", "←", "=", "+"}:
                converted_parts.append(part)
            else:
                converted_parts.append(format_chemical_species(part))

    return "".join(converted_parts)


def extract_braced_content(text: str, brace_start: int) -> tuple[str | None, int]:
    if brace_start >= len(text) or text[brace_start] != "{":
        return None, brace_start

    depth = 0
    for index in range(brace_start, len(text)):
        char = text[index]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[brace_start + 1 : index], index + 1

    return None, brace_start


def replace_ce_commands(math_text: str) -> str:
    parts: list[str] = []
    index = 0

    while index < len(math_text):
        if math_text.startswith(r"\ce{", index):
            content, next_index = extract_braced_content(math_text, index + 3)
            if content is None or next_index == index + 3:
                parts.append(math_text[index])
                index += 1
                continue

            converted = convert_ce_expression(content)
            parts.append(rf"\text{{{escape_tex_text(converted)}}}")
            index = next_index
            continue

        parts.append(math_text[index])
        index += 1

    return "".join(parts)


def is_escaped(text: str, index: int) -> bool:
    backslash_count = 0
    pointer = index - 1
    while pointer >= 0 and text[pointer] == "\\":
        backslash_count += 1
        pointer -= 1
    return backslash_count % 2 == 1


def find_closing_delimiter(text: str, start: int, delimiter: str) -> int:
    index = start
    while index < len(text):
        if text.startswith(delimiter, index) and not is_escaped(text, index):
            return index
        index += 1
    return -1


def preprocess_markdown_for_docx(text: str) -> str:
    result: list[str] = []
    index = 0

    while index < len(text):
        if text.startswith("$$", index) and not is_escaped(text, index):
            end = find_closing_delimiter(text, index + 2, "$$")
            if end == -1:
                result.append(text[index:])
                break

            content = text[index + 2 : end]
            result.append("$$")
            result.append(replace_ce_commands(content))
            result.append("$$")
            index = end + 2
            continue

        if text[index] == "$" and not is_escaped(text, index):
            end = find_closing_delimiter(text, index + 1, "$")
            if end == -1:
                result.append(text[index])
                index += 1
                continue

            content = text[index + 1 : end]
            result.append("$")
            result.append(replace_ce_commands(content))
            result.append("$")
            index = end + 1
            continue

        result.append(text[index])
        index += 1

    return "".join(result)


def run_pandoc(
    pandoc_path: str, source_md: Path, destination_docx: Path, reference_doc: Path | None
) -> None:
    destination_docx.parent.mkdir(parents=True, exist_ok=True)

    command = [
        pandoc_path,
        "--from=markdown+tex_math_dollars",
        str(source_md),
        "-o",
        str(destination_docx),
    ]

    if reference_doc is not None:
        command.extend(["--reference-doc", str(reference_doc)])

    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or "pandoc conversion failed")

    if completed.stderr.strip():
        print(completed.stderr.strip(), file=sys.stderr)


def convert_file(
    pandoc_path: str, source: Path, destination: Path, reference_doc: Path | None
) -> None:
    prepared_markdown = preprocess_markdown_for_docx(read_text(source))

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_md = Path(temp_dir) / f"{source.stem}.prepared.md"
        temp_md.write_text(prepared_markdown, encoding="utf-8")
        run_pandoc(pandoc_path, temp_md, destination, reference_doc)


def main() -> int:
    args = parse_args()
    pandoc_path = shutil.which("pandoc")
    if pandoc_path is None:
        print("[ERROR] pandoc is not installed or not in PATH.", file=sys.stderr)
        return 1

    input_path = Path(args.input_path)
    output_path = Path(args.output)
    reference_doc = Path(args.reference_doc) if args.reference_doc else None

    if reference_doc is not None and not reference_doc.is_file():
        print(f"[ERROR] Reference DOCX does not exist: {reference_doc}", file=sys.stderr)
        return 1

    try:
        files = collect_input_files(input_path)
        for source in files:
            destination = build_output_path(source, input_path, output_path)
            convert_file(pandoc_path, source, destination, reference_doc)
            print(f"[OK] {source} -> {destination}")
    except (RuntimeError, ValueError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
