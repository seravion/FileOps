from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any
from zipfile import is_zipfile

from .models import OperationResult, OperationStatus
from .utils import duration_ms, ensure_workspace_path, now_iso

try:
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:  # pragma: no cover
    Document = None
    PackageNotFoundError = None
    Table = None
    Paragraph = None
    qn = None

try:
    import pytesseract
    from PIL import Image
except ImportError:  # pragma: no cover
    pytesseract = None
    Image = None


HEADING_MODES = {"h1", "h2", "h1_h2"}
INPUT_FORMATS = {"auto", "docx", "markdown", "txt"}
OUTPUT_FORMATS = {"auto", "docx", "md", "txt"}
TEXT_EXTENSIONS = {".md", ".markdown", ".txt"}


def split_documents_by_structure(
    sources: list[Path],
    destination: Path,
    workspace: Path,
    dry_run: bool,
    heading_mode: str,
    include_image_text: bool,
    input_format: str = "auto",
    output_format: str = "auto",
) -> list[OperationResult]:
    if heading_mode not in HEADING_MODES:
        raise ValueError(f"Unsupported heading mode: {heading_mode}")
    if input_format not in INPUT_FORMATS:
        raise ValueError(f"Unsupported input format: {input_format}")
    if output_format not in OUTPUT_FORMATS:
        raise ValueError(f"Unsupported output format: {output_format}")

    destination = destination.resolve(strict=False)
    ensure_workspace_path(destination, workspace)

    if not dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    results: list[OperationResult] = []

    for source in sources:
        started = datetime.now()
        started_at = now_iso()

        try:
            ensure_workspace_path(source, workspace)
            if not source.exists():
                raise FileNotFoundError(f"Source does not exist: {source}")
            if source.is_dir():
                raise IsADirectoryError(f"Document split supports files only: {source}")

            ext = source.suffix.lower()
            if not _matches_input_format(ext, input_format):
                raise ValueError(f"Source format does not match import format setting: {source.name}")

            if ext == ".docx":
                sections = _split_docx(source, heading_mode, include_image_text)
                if not sections:
                    sections = [{"title": "section", "start": 0, "end": 0}]
            elif ext in TEXT_EXTENSIONS:
                sections = _split_text_document(source, heading_mode, include_image_text)
                if not sections:
                    sections = [{"title": "section", "lines": [""]}]
            else:
                raise ValueError(f"Unsupported file type for document split: {source.name}")

            resolved_output = _resolve_output_format(ext, output_format)

            if dry_run:
                message = f"Would split into {len(sections)} section(s), output format: {resolved_output}."
                results.append(
                    _build_result("doc_split", source, destination, OperationStatus.DRY_RUN, message, started, started_at)
                )
                continue

            if ext == ".docx":
                if resolved_output == "docx":
                    created = _write_docx_sections(source, destination, sections)
                else:
                    created = _write_docx_sections_as_text(
                        source=source,
                        destination=destination,
                        sections=sections,
                        output_ext=resolved_output,
                        include_image_text=include_image_text,
                    )
            else:
                if resolved_output == "docx":
                    created = _write_text_sections_as_docx(source, destination, sections)
                else:
                    created = _write_text_sections(source, destination, sections, output_ext=resolved_output)

            message = f"Document split completed: {len(created)} file(s) generated."
            results.append(_build_result("doc_split", source, destination, OperationStatus.SUCCESS, message, started, started_at))

        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("doc_split", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def _matches_input_format(ext: str, input_format: str) -> bool:
    if input_format == "auto":
        return ext == ".docx" or ext in TEXT_EXTENSIONS
    if input_format == "docx":
        return ext == ".docx"
    if input_format == "markdown":
        return ext in {".md", ".markdown"}
    return ext == ".txt"


def _resolve_output_format(source_ext: str, output_format: str) -> str:
    if output_format != "auto":
        return output_format
    if source_ext == ".docx":
        return "docx"
    if source_ext in {".md", ".markdown"}:
        return "md"
    return "txt"


def _looks_like_legacy_doc(source: Path) -> bool:
    try:
        with source.open("rb") as stream:
            signature = stream.read(8)
    except OSError:
        return False
    return signature == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _split_docx(source: Path, heading_mode: str, include_image_text: bool) -> list[dict[str, Any]]:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    if not is_zipfile(source):
        if _looks_like_legacy_doc(source):
            raise ValueError("Detected legacy DOC or fake .docx. Please re-save as standard .docx in Word/WPS and retry.")
        raise ValueError("Invalid .docx package (file damaged or extension mismatch). Please re-save as .docx in Word/WPS and retry.")

    _ = include_image_text

    try:
        with source.open("rb") as stream:
            doc = Document(stream)

        blocks = _iter_docx_body_blocks(doc)
        if not blocks:
            return [{"title": "section", "start": 0, "end": 0}]

        heading_points: list[tuple[int, str]] = []
        for idx, block in enumerate(blocks):
            if block.tag != qn("w:p"):
                continue
            paragraph = _paragraph_from_xml(block, doc)
            heading_level = _get_docx_heading_level(_safe_docx_style_name(paragraph))
            if not _is_heading_boundary(heading_level, heading_mode):
                continue
            title = paragraph.text.strip() or f"section_{len(heading_points) + 1}"
            heading_points.append((idx, title))

        sections: list[dict[str, Any]] = []
        if not heading_points:
            sections.append({"title": "Preface", "start": 0, "end": len(blocks)})
            return sections

        first_start = heading_points[0][0]
        if first_start > 0:
            sections.append({"title": "Preface", "start": 0, "end": first_start})

        for index, (start_idx, title) in enumerate(heading_points):
            end_idx = heading_points[index + 1][0] if index + 1 < len(heading_points) else len(blocks)
            if end_idx > start_idx:
                sections.append({"title": title, "start": start_idx, "end": end_idx})

        return sections

    except Exception as exc:  # noqa: BLE001
        if PackageNotFoundError is not None and isinstance(exc, PackageNotFoundError):
            raise ValueError("Unable to read .docx package. Ensure the file exists and valid.") from exc
        raise ValueError(f".docx parse failed: {exc}") from exc


def _split_text_document(source: Path, heading_mode: str, include_image_text: bool) -> list[dict[str, Any]]:
    lines = source.read_text(encoding="utf-8", errors="ignore").splitlines()

    sections: list[dict[str, Any]] = []
    current_title = "Preface"
    current_lines: list[str] = []

    heading_regex = re.compile(r"^(#{1,6})\s+(.*)$")
    image_regex = re.compile(r"!\[(.*?)\]\((.*?)\)")

    for raw in lines:
        line = raw.rstrip()
        match = heading_regex.match(line)
        if match:
            level = len(match.group(1))
            title = match.group(2).strip() or f"section_{len(sections) + 1}"
            if _is_heading_boundary(level, heading_mode):
                if current_lines:
                    sections.append({"title": current_title, "lines": current_lines[:]})
                    current_lines.clear()
                current_title = title
            current_lines.append(line)
            continue

        current_lines.append(line)

        if include_image_text:
            image_match = image_regex.search(line)
            if image_match:
                alt_text = image_match.group(1).strip()
                if alt_text:
                    current_lines.append(f"[Image Alt] {alt_text}")

    if current_lines:
        sections.append({"title": current_title, "lines": current_lines[:]})

    return sections


def _iter_docx_body_blocks(doc: Any) -> list[Any]:
    if qn is None:
        return []
    body = doc.element.body
    return [child for child in body.iterchildren() if child.tag in {qn("w:p"), qn("w:tbl")}]


def _paragraph_from_xml(paragraph_xml: Any, doc: Any) -> Any:
    if Paragraph is None:
        raise RuntimeError("python-docx paragraph support is unavailable.")
    return Paragraph(paragraph_xml, doc)


def _table_from_xml(table_xml: Any, doc: Any) -> Any:
    if Table is None:
        raise RuntimeError("python-docx table support is unavailable.")
    return Table(table_xml, doc)


def _safe_docx_style_name(paragraph: Any) -> str:
    try:
        style = paragraph.style
    except Exception:  # noqa: BLE001
        return ""
    return str(getattr(style, "name", "") or "")


def _extract_docx_image_lines(paragraph: Any, include_image_text: bool) -> list[str]:
    if not include_image_text:
        return []

    image_lines: list[str] = []
    if qn is None:
        return image_lines

    try:
        blips = paragraph._element.xpath(".//a:blip")
    except Exception:  # noqa: BLE001
        return image_lines
    for blip in blips:
        embed_id = blip.get(qn("r:embed"))
        if not embed_id:
            continue

        try:
            image_part = paragraph.part.related_parts.get(embed_id)
        except Exception:  # noqa: BLE001
            continue
        if image_part is None:
            continue

        ocr_text = _ocr_image_blob(image_part.blob)
        if ocr_text:
            image_lines.append(f"[Image Text] {ocr_text}")
        else:
            image_lines.append("[Image Text] Not recognized")

    return image_lines


def _ocr_image_blob(blob: bytes) -> str:
    if pytesseract is None or Image is None:
        return ""

    try:
        from io import BytesIO

        image = Image.open(BytesIO(blob))
        text = pytesseract.image_to_string(image, lang="chi_sim+eng")
        return text.strip()
    except Exception:  # noqa: BLE001
        return ""


def _get_docx_heading_level(style_name: str) -> int | None:
    if not style_name:
        return None

    lower = style_name.lower()
    if "heading" not in lower:
        return None

    nums = re.findall(r"\d+", lower)
    if not nums:
        return 1
    return int(nums[0])


def _is_heading_boundary(level: int | None, mode: str) -> bool:
    if level is None:
        return False
    if mode == "h1":
        return level == 1
    if mode == "h2":
        return level == 2
    return level in {1, 2}


def _write_docx_sections(source: Path, destination: Path, sections: list[dict[str, Any]]) -> list[Path]:
    destination.mkdir(parents=True, exist_ok=True)

    created_files: list[Path] = []
    index_payload: list[dict[str, Any]] = []

    for idx, section in enumerate(sections, start=1):
        title = str(section.get("title", "section")).strip() or f"section_{idx}"
        safe_title = _safe_filename(title)
        output_path = destination / f"{source.stem}_{idx:03d}_{safe_title}.docx"

        shutil.copy2(source, output_path)
        start = int(section.get("start", 0))
        end = int(section.get("end", 0))
        _prune_docx_file(output_path, start, end)

        created_files.append(output_path)
        index_payload.append({"index": idx, "title": title, "file": output_path.name, "block_count": max(0, end - start)})

    index_file = destination / f"{source.stem}_split_index.json"
    index_file.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return created_files


def _write_docx_sections_as_text(
    source: Path,
    destination: Path,
    sections: list[dict[str, Any]],
    output_ext: str,
    include_image_text: bool,
) -> list[Path]:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    destination.mkdir(parents=True, exist_ok=True)

    doc = Document(str(source))
    blocks = _iter_docx_body_blocks(doc)

    created_files: list[Path] = []
    index_payload: list[dict[str, Any]] = []

    markdown = output_ext == "md"
    suffix = ".md" if markdown else ".txt"

    for idx, section in enumerate(sections, start=1):
        title = str(section.get("title", "section")).strip() or f"section_{idx}"
        safe_title = _safe_filename(title)
        output_path = destination / f"{source.stem}_{idx:03d}_{safe_title}{suffix}"

        start = int(section.get("start", 0))
        end = int(section.get("end", 0))
        lines = _extract_docx_section_lines(doc, blocks, start, end, include_image_text, markdown)

        text_content = "\n".join(lines).strip() + "\n"
        output_path.write_text(text_content, encoding="utf-8")

        created_files.append(output_path)
        index_payload.append({"index": idx, "title": title, "file": output_path.name, "line_count": len(lines)})

    index_file = destination / f"{source.stem}_split_index.json"
    index_file.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return created_files


def _extract_docx_section_lines(
    doc: Any,
    blocks: list[Any],
    start: int,
    end: int,
    include_image_text: bool,
    markdown: bool,
) -> list[str]:
    lines: list[str] = []

    for idx, block in enumerate(blocks):
        if idx < start or idx >= end:
            continue

        if block.tag == qn("w:p"):
            paragraph = _paragraph_from_xml(block, doc)
            text = paragraph.text.strip()
            if text:
                if markdown:
                    level = _get_docx_heading_level(_safe_docx_style_name(paragraph))
                    if level is not None:
                        text = f"{'#' * min(level, 6)} {text}"
                lines.append(text)
            lines.extend(_extract_docx_image_lines(paragraph, include_image_text))
            continue

        if block.tag == qn("w:tbl"):
            table = _table_from_xml(block, doc)
            table_lines = _table_to_lines(table, markdown)
            if table_lines:
                lines.extend(table_lines)
                lines.append("")

    while lines and not lines[-1].strip():
        lines.pop()

    return lines


def _table_to_lines(table: Any, markdown: bool) -> list[str]:
    rows: list[list[str]] = []

    for row in table.rows:
        rows.append([cell.text.replace("\n", " ").strip() for cell in row.cells])

    if not rows:
        return []

    cols = max(len(row) for row in rows)
    norm_rows = [row + [""] * (cols - len(row)) for row in rows]

    if markdown:
        header = norm_rows[0]
        lines = [f"| {' | '.join(header)} |", f"| {' | '.join(['---'] * cols)} |"]
        lines.extend(f"| {' | '.join(row)} |" for row in norm_rows[1:])
        return lines

    return ["\t".join(row).rstrip() for row in norm_rows]


def _prune_docx_file(path: Path, start: int, end: int) -> None:
    doc = Document(str(path))
    blocks = _iter_docx_body_blocks(doc)
    for idx, block in enumerate(blocks):
        if start <= idx < end:
            continue
        block.getparent().remove(block)

    if not _iter_docx_body_blocks(doc):
        doc.add_paragraph("")

    doc.save(str(path))


def _write_text_sections(source: Path, destination: Path, sections: list[dict[str, Any]], output_ext: str = "txt") -> list[Path]:
    destination.mkdir(parents=True, exist_ok=True)

    created_files: list[Path] = []
    index_payload: list[dict[str, Any]] = []
    suffix = ".md" if output_ext == "md" else ".txt"

    for idx, section in enumerate(sections, start=1):
        title = str(section.get("title", "section")).strip() or f"section_{idx}"
        lines = [str(item) for item in section.get("lines", [])]

        safe_title = _safe_filename(title)
        output_path = destination / f"{source.stem}_{idx:03d}_{safe_title}{suffix}"
        text_content = "\n".join(lines).strip() + "\n"
        output_path.write_text(text_content, encoding="utf-8")

        created_files.append(output_path)
        index_payload.append({"index": idx, "title": title, "file": output_path.name, "line_count": len(lines)})

    index_file = destination / f"{source.stem}_split_index.json"
    index_file.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return created_files


def _write_text_sections_as_docx(source: Path, destination: Path, sections: list[dict[str, Any]]) -> list[Path]:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    destination.mkdir(parents=True, exist_ok=True)

    created_files: list[Path] = []
    index_payload: list[dict[str, Any]] = []

    for idx, section in enumerate(sections, start=1):
        title = str(section.get("title", "section")).strip() or f"section_{idx}"
        lines = [str(item) for item in section.get("lines", [])]

        safe_title = _safe_filename(title)
        output_path = destination / f"{source.stem}_{idx:03d}_{safe_title}.docx"

        doc = Document()
        _append_text_lines_to_doc(doc, lines)
        if not doc.paragraphs and not doc.tables:
            doc.add_paragraph("")
        doc.save(str(output_path))

        created_files.append(output_path)
        index_payload.append({"index": idx, "title": title, "file": output_path.name, "line_count": len(lines)})

    index_file = destination / f"{source.stem}_split_index.json"
    index_file.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return created_files


def _append_text_lines_to_doc(doc: Any, lines: list[str]) -> None:
    heading_regex = re.compile(r"^(#{1,6})\s+(.*)$")

    index = 0
    while index < len(lines):
        raw = lines[index]
        line = raw.rstrip()

        if _looks_like_markdown_table_row(line):
            table_lines = [line]
            index += 1
            while index < len(lines) and _looks_like_markdown_table_row(lines[index].rstrip()):
                table_lines.append(lines[index].rstrip())
                index += 1
            _append_markdown_table_to_doc(doc, table_lines)
            continue

        heading_match = heading_regex.match(line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            if text:
                doc.add_heading(text, level=level)
            else:
                doc.add_paragraph(line)
        elif line:
            doc.add_paragraph(line)
        else:
            doc.add_paragraph("")
        index += 1


def _looks_like_markdown_table_row(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 2


def _append_markdown_table_to_doc(doc: Any, table_lines: list[str]) -> None:
    rows = [_parse_markdown_table_row(item) for item in table_lines]
    if not rows:
        return

    if len(rows) >= 2 and _is_markdown_separator_row(rows[1]):
        rows = [rows[0], *rows[2:]]

    if not rows:
        return

    cols = max(len(row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=cols)

    for row_idx, row in enumerate(rows):
        norm = row + [""] * (cols - len(row))
        for col_idx, value in enumerate(norm):
            table.cell(row_idx, col_idx).text = value


def _parse_markdown_table_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [cell.strip() for cell in stripped.split("|")]


def _is_markdown_separator_row(row: list[str]) -> bool:
    for cell in row:
        normalized = cell.replace(" ", "")
        if not normalized:
            continue
        if not re.fullmatch(r":?-{3,}:?", normalized):
            return False
    return True


def _safe_filename(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|\s]+", "_", value).strip("_")
    return cleaned[:40] or "section"


def _build_result(
    operation: str,
    source: Path,
    destination: Path | None,
    status: OperationStatus,
    message: str,
    started: datetime,
    started_at: str,
) -> OperationResult:
    finished = datetime.now()
    finished_at = now_iso()

    return OperationResult(
        operation=operation,
        source=str(source),
        destination=str(destination) if destination else None,
        status=status,
        message=message,
        started_at=started_at,
        finished_at=finished_at,
        duration_ms=duration_ms(started, finished),
    )
