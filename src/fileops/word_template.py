from __future__ import annotations

import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from .models import OperationResult, OperationStatus
from .utils import duration_ms, ensure_workspace_path, now_iso, unique_path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Emu
except ImportError:  # pragma: no cover
    Document = None
    qn = None
    Emu = None


def template_library_dir() -> Path:
    candidates: list[Path] = []

    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        candidates.append(Path(local_app_data) / "FileOps" / "templates")

    roaming_app_data = os.getenv("APPDATA")
    if roaming_app_data:
        candidates.append(Path(roaming_app_data) / "FileOps" / "templates")

    candidates.append(Path.home() / ".fileops" / "templates")
    candidates.append(Path.cwd() / ".fileops" / "templates")

    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except OSError:
            continue

    raise PermissionError("Unable to create template library directory. Check folder permissions.")


def list_word_templates() -> list[Path]:
    base = template_library_dir()
    return sorted([item for item in base.glob("*.docx") if item.is_file()], key=lambda item: item.name.lower())


def import_word_template(template_file: Path) -> Path:
    template_file = template_file.resolve(strict=False)
    if not template_file.exists():
        raise FileNotFoundError(f"Template does not exist: {template_file}")
    if template_file.suffix.lower() != ".docx":
        raise ValueError("Only .docx files can be imported as templates.")

    target_dir = template_library_dir()
    target = target_dir / template_file.name
    if target.exists():
        target = unique_path(target)
    target.write_bytes(template_file.read_bytes())
    return target


def format_word_documents(
    sources: list[Path],
    destination: Path,
    workspace: Path,
    dry_run: bool,
    template_path: Path,
) -> list[OperationResult]:
    if Document is None or qn is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    destination = destination.resolve(strict=False)
    ensure_workspace_path(destination, workspace)
    if not dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    template_path = template_path.resolve(strict=False)
    if not template_path.exists():
        raise FileNotFoundError(f"Template does not exist: {template_path}")
    if template_path.suffix.lower() != ".docx":
        raise ValueError("Template must be a .docx file.")

    results: list[OperationResult] = []
    for source in sources:
        started = datetime.now()
        started_at = now_iso()
        try:
            ensure_workspace_path(source, workspace)
            if not source.exists():
                raise FileNotFoundError(f"Source does not exist: {source}")
            if source.is_dir():
                raise IsADirectoryError(f"Word format supports files only: {source}")
            if source.suffix.lower() != ".docx":
                raise ValueError("Word format supports .docx files only.")

            output_path = destination / f"{source.stem}_formatted.docx"
            if output_path.exists():
                output_path = unique_path(output_path)

            if dry_run:
                message = f"Would format by template '{template_path.name}' -> {output_path.name}"
                results.append(_build_result("word_format", source, output_path, OperationStatus.DRY_RUN, message, started, started_at))
                continue

            _apply_template_format(source=source, template=template_path, output=output_path)
            message = f"Formatted by template '{template_path.name}' -> {output_path.name}"
            results.append(_build_result("word_format", source, output_path, OperationStatus.SUCCESS, message, started, started_at))
        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("word_format", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def _apply_template_format(source: Path, template: Path, output: Path) -> None:
    template_doc = Document(str(template))
    source_doc = Document(str(source))

    _clear_template_body_keep_sectpr(template_doc)

    paragraph_style = "Normal"
    table_style = _resolve_first_table_style(template_doc)

    source_blocks = _iter_body_blocks(source_doc)
    for block in source_blocks:
        if block.tag == qn("w:p"):
            paragraph = _paragraph_from_xml(block, source_doc)
            _copy_paragraph_to_template(
                source_paragraph=paragraph,
                target_doc=template_doc,
                default_paragraph_style=paragraph_style,
            )
            continue

        if block.tag == qn("w:tbl"):
            table = _table_from_xml(block, source_doc)
            _copy_table_to_template(source_table=table, target_doc=template_doc, table_style=table_style)

    template_doc.save(str(output))


def _clear_template_body_keep_sectpr(doc: Any) -> None:
    body = doc.element.body
    children = list(body.iterchildren())
    sectpr = next((child for child in children if child.tag == qn("w:sectPr")), None)

    for child in children:
        body.remove(child)

    if sectpr is not None:
        body.append(sectpr)
    else:
        doc.add_paragraph("")


def _resolve_first_table_style(doc: Any) -> str | None:
    for style in doc.styles:
        style_type = str(getattr(style.type, "name", "") or "")
        if style_type.lower() == "table":
            return style.name
    return None


def _iter_body_blocks(doc: Any) -> list[Any]:
    body = doc.element.body
    return [child for child in body.iterchildren() if child.tag in {qn("w:p"), qn("w:tbl")}]


def _paragraph_from_xml(paragraph_xml: Any, doc: Any) -> Any:
    from docx.text.paragraph import Paragraph

    return Paragraph(paragraph_xml, doc)


def _table_from_xml(table_xml: Any, doc: Any) -> Any:
    from docx.table import Table

    return Table(table_xml, doc)


def _copy_paragraph_to_template(source_paragraph: Any, target_doc: Any, default_paragraph_style: str) -> None:
    heading_style = _resolve_heading_style_name(source_paragraph)
    target_style = heading_style or default_paragraph_style

    target_paragraph = target_doc.add_paragraph()
    if target_style:
        try:
            target_paragraph.style = target_style
        except Exception:  # noqa: BLE001
            pass

    if not source_paragraph.runs:
        target_paragraph.add_run(source_paragraph.text)
        return

    for run in source_paragraph.runs:
        if run.text:
            new_run = target_paragraph.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
        _copy_images_from_run(source_run=run, target_paragraph=target_paragraph)


def _resolve_heading_style_name(paragraph: Any) -> str | None:
    try:
        style_name = str(getattr(paragraph.style, "name", "") or "")
    except Exception:  # noqa: BLE001
        return None

    if not style_name:
        return None
    if "heading" in style_name.lower():
        return style_name
    return None


def _copy_images_from_run(source_run: Any, target_paragraph: Any) -> None:
    blips = source_run._element.xpath(".//a:blip")
    extents = source_run._element.xpath(".//wp:extent")
    extent_pairs: list[tuple[int | None, int | None]] = []
    for item in extents:
        try:
            cx = int(item.get("cx")) if item.get("cx") else None
            cy = int(item.get("cy")) if item.get("cy") else None
        except Exception:  # noqa: BLE001
            cx = None
            cy = None
        extent_pairs.append((cx, cy))

    for index, blip in enumerate(blips):
        embed_id = blip.get(qn("r:embed"))
        if not embed_id:
            continue
        image_part = source_run.part.related_parts.get(embed_id)
        if image_part is None:
            continue

        width = None
        height = None
        if Emu is not None and index < len(extent_pairs):
            cx, cy = extent_pairs[index]
            width = Emu(cx) if cx else None
            height = Emu(cy) if cy else None

        image_stream = BytesIO(image_part.blob)
        target_run = target_paragraph.add_run()
        target_run.add_picture(image_stream, width=width, height=height)


def _copy_table_to_template(source_table: Any, target_doc: Any, table_style: str | None) -> None:
    row_count = len(source_table.rows)
    col_count = len(source_table.columns)
    if row_count == 0 or col_count == 0:
        return

    target_table = target_doc.add_table(rows=row_count, cols=col_count)
    if table_style:
        try:
            target_table.style = table_style
        except Exception:  # noqa: BLE001
            pass

    for row_idx, row in enumerate(source_table.rows):
        for col_idx, cell in enumerate(row.cells):
            target_table.cell(row_idx, col_idx).text = cell.text


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
