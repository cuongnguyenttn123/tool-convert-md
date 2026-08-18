#!/usr/bin/env python3
"""Convert DOCX or Excel files to Markdown.

Rules:
- DOCX: output a markdown file with the same base name.
- Excel: output a folder with the same base name. Each sheet becomes one markdown file.
- Images: extracted to an images folder and linked with relative paths.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook


@dataclass
class ImageRef:
    filename: str
    rel_path: str


def sanitize_name(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]", "_", name.strip())
    return name or "untitled"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_text(path: Path, content: str) -> None:
    path.write_text(content.rstrip() + "\n", encoding="utf-8")


def docx_extract_image(paragraph, image_dir: Path, image_index: int) -> tuple[list[str], int]:
    """Extract images from one paragraph and return markdown lines for inline placement."""
    lines: list[str] = []

    for run in paragraph.runs:
        drawings = run._element.xpath(".//w:drawing")
        if not drawings:
            continue

        for drawing in drawings:
            blips = drawing.xpath(".//a:blip")
            for blip in blips:
                rel_id = blip.get(qn("r:embed"))
                if not rel_id:
                    continue
                img_part = paragraph.part.related_parts.get(rel_id)
                if img_part is None:
                    continue

                ext = Path(img_part.filename).suffix or ".png"
                image_index += 1
                filename = f"img_{image_index:04d}{ext}"
                target = image_dir / filename
                target.write_bytes(img_part.blob)
                rel_path = f"{image_dir.name}/{filename}"
                lines.append(f"![image]({rel_path})")

    return lines, image_index


def paragraph_to_markdown(paragraph, image_dir: Path, image_index: int) -> tuple[list[str], int]:
    style_name = paragraph.style.name if paragraph.style else ""
    text = (paragraph.text or "").strip()
    lines: list[str] = []

    image_lines, image_index = docx_extract_image(paragraph, image_dir, image_index)

    if text:
        heading_match = re.match(r"Heading\s+(\d+)$", style_name)
        if heading_match:
            level = max(1, min(6, int(heading_match.group(1))))
            lines.append(f"{'#' * level} {text}")
        elif "List Bullet" in style_name:
            lines.append(f"- {text}")
        elif "List Number" in style_name:
            lines.append(f"1. {text}")
        else:
            lines.append(text)

    lines.extend(image_lines)
    return lines, image_index


def table_to_markdown(table) -> list[str]:
    rows: list[list[str]] = []
    for row in table.rows:
        values = [((cell.text or "").replace("\n", " ").strip()) for cell in row.cells]
        rows.append(values)

    if not rows:
        return []

    col_count = max(len(r) for r in rows)
    normalized = [r + [""] * (col_count - len(r)) for r in rows]

    header = normalized[0]
    sep = ["---"] * col_count

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(sep) + " |",
    ]

    for data_row in normalized[1:]:
        lines.append("| " + " | ".join(data_row) + " |")

    return lines


def convert_docx(input_path: Path, output_root: Path) -> Path:
    doc = Document(str(input_path))
    base = sanitize_name(input_path.stem)

    out_dir = output_root / base
    ensure_dir(out_dir)

    output_md = out_dir / f"{base}.md"
    image_dir = out_dir / "images"
    ensure_dir(image_dir)

    lines: list[str] = [f"# {input_path.stem}", ""]
    image_index = 0

    body = doc.element.body
    for element in body.iterchildren():
        if element.tag == qn("w:p"):
            paragraph = next((p for p in doc.paragraphs if p._element is element), None)
            if paragraph is None:
                continue
            paragraph_lines, image_index = paragraph_to_markdown(paragraph, image_dir, image_index)
            if paragraph_lines:
                lines.extend(paragraph_lines)
                lines.append("")
        elif element.tag == qn("w:tbl"):
            table = next((t for t in doc.tables if t._element is element), None)
            if table is None:
                continue
            table_lines = table_to_markdown(table)
            if table_lines:
                lines.extend(table_lines)
                lines.append("")

    # Remove image folder if no image was extracted.
    if image_index == 0 and image_dir.exists() and not any(image_dir.iterdir()):
        image_dir.rmdir()

    write_text(output_md, "\n".join(lines))
    return output_md


def excel_sheet_used_range(ws) -> tuple[int, int, int, int]:
    """Return min_row, max_row, min_col, max_col for non-empty cells."""
    min_row: Optional[int] = None
    max_row: Optional[int] = None
    min_col: Optional[int] = None
    max_col: Optional[int] = None

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is None:
                continue
            r, c = cell.row, cell.column
            min_row = r if min_row is None else min(min_row, r)
            max_row = r if max_row is None else max(max_row, r)
            min_col = c if min_col is None else min(min_col, c)
            max_col = c if max_col is None else max(max_col, c)

    if min_row is None:
        return 1, 1, 1, 1
    return min_row, max_row or 1, min_col or 1, max_col or 1


def format_cell(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    text = text.replace("|", "\\|")
    return text


def extract_excel_images(ws, images_dir: Path, prefix: str, image_index: int) -> tuple[list[str], int]:
    lines: list[str] = []
    images = getattr(ws, "_images", [])

    for img in images:
        ext = ".png"
        fmt = getattr(img, "format", None)
        if isinstance(fmt, str) and fmt:
            ext = f".{fmt.lower()}"

        image_index += 1
        filename = f"{prefix}_img_{image_index:04d}{ext}"
        target = images_dir / filename

        try:
            data = img._data()
            if callable(data):
                blob = data()
            else:
                blob = data
            if blob is None:
                continue
            target.write_bytes(blob)
        except Exception:
            # Skip non-exportable images, keep conversion resilient.
            continue

        anchor = getattr(img, "anchor", None)
        position = ""
        if hasattr(anchor, "_from"):
            row = anchor._from.row + 1
            col = anchor._from.col + 1
            position = f" (cell R{row}C{col})"

        lines.append(f"![{ws.title}{position}](images/{filename})")

    return lines, image_index


def convert_excel(input_path: Path, output_root: Path) -> Path:
    wb = load_workbook(filename=str(input_path), data_only=True)
    base = sanitize_name(input_path.stem)

    out_dir = output_root / base
    ensure_dir(out_dir)

    images_dir = out_dir / "images"
    ensure_dir(images_dir)

    image_index = 0

    for ws in wb.worksheets:
        sheet_name = sanitize_name(ws.title)
        md_path = out_dir / f"{sheet_name}.md"

        min_row, max_row, min_col, max_col = excel_sheet_used_range(ws)
        width = max_col - min_col + 1

        lines: list[str] = [f"# {ws.title}", ""]

        headers = [
            format_cell(ws.cell(row=min_row, column=min_col + i).value) or f"Column {i + 1}"
            for i in range(width)
        ]
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("| " + " | ".join(["---"] * width) + " |")

        for r in range(min_row + 1, max_row + 1):
            row_data = [format_cell(ws.cell(row=r, column=min_col + i).value) for i in range(width)]
            lines.append("| " + " | ".join(row_data) + " |")

        safe_prefix = sheet_name.replace(" ", "_")
        image_lines, image_index = extract_excel_images(ws, images_dir, safe_prefix, image_index)
        if image_lines:
            lines.append("")
            lines.append("## Images")
            lines.append("")
            lines.extend(image_lines)

        write_text(md_path, "\n".join(lines))

    if image_index == 0 and images_dir.exists() and not any(images_dir.iterdir()):
        images_dir.rmdir()

    return out_dir


def iter_input_files(path: Path) -> Iterable[Path]:
    if path.is_file():
        yield path
        return

    for ext in ("*.docx", "*.xlsx", "*.xlsm"):
        yield from path.rglob(ext)


def convert_file(input_path: Path, output_dir: Path) -> Path:
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return convert_docx(input_path, output_dir)
    if suffix in {".xlsx", ".xlsm"}:
        return convert_excel(input_path, output_dir)
    raise ValueError(f"Unsupported file type: {input_path}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert DOCX or Excel files to Markdown with extracted images."
    )
    parser.add_argument("input", type=Path, help="Input file or folder")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output root folder (default: same folder as input)",
    )
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input not found: {input_path}")

    output_root = (
        args.output.expanduser().resolve()
        if args.output
        else (input_path.parent if input_path.is_file() else input_path)
    )
    ensure_dir(output_root)

    converted: list[Path] = []
    for file_path in iter_input_files(input_path):
        if file_path.name.startswith("~$"):
            continue
        converted.append(convert_file(file_path, output_root))

    if not converted:
        print("No supported files found (.docx, .xlsx, .xlsm).")
        return 0

    print("Converted outputs:")
    for path in converted:
        print(f"- {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
