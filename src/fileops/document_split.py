from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import OperationResult, OperationStatus
from .utils import duration_ms, ensure_workspace_path, now_iso

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:  # pragma: no cover
    Document = None
    qn = None

try:
    import pytesseract
    from PIL import Image
except ImportError:  # pragma: no cover
    pytesseract = None
    Image = None


HEADING_MODES = {"h1", "h2", "h1_h2"}


def split_documents_by_structure(
    sources: list[Path],
    destination: Path,
    workspace: Path,
    dry_run: bool,
    heading_mode: str,
    include_image_text: bool,
) -> list[OperationResult]:
    if heading_mode not in HEADING_MODES:
        raise ValueError(f"Unsupported heading mode: {heading_mode}")

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
            if ext == ".docx":
                sections = _split_docx(source, heading_mode, include_image_text)
            elif ext in {".md", ".markdown", ".txt"}:
                sections = _split_text_document(source, heading_mode, include_image_text)
            else:
                raise ValueError(f"Unsupported file type for document split: {source.name}")

            if not sections:
                sections = [{"title": "section", "lines": [""]}]

            if dry_run:
                message = f"Would split into {len(sections)} section(s)."
                results.append(_build_result("doc_split", source, destination, OperationStatus.DRY_RUN, message, started, started_at))
                continue

            created = _write_sections(source, destination, sections)
            message = f"Document split completed: {len(created)} file(s) generated."
            results.append(_build_result("doc_split", source, destination, OperationStatus.SUCCESS, message, started, started_at))

        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("doc_split", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def _split_docx(source: Path, heading_mode: str, include_image_text: bool) -> list[dict[str, Any]]:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    doc = Document(str(source))
    sections: list[dict[str, Any]] = []
    current_title = "导言"
    current_lines: list[str] = []

    for paragraph in doc.paragraphs:
        heading_level = _get_docx_heading_level(getattr(paragraph.style, "name", ""))
        text = paragraph.text.strip()

        if _is_heading_boundary(heading_level, heading_mode):
            if current_lines:
                sections.append({"title": current_title, "lines": current_lines[:]})
                current_lines.clear()
            current_title = text or f"section_{len(sections) + 1}"
            if text:
                current_lines.append(text)
            continue

        if text:
            current_lines.append(text)

        if include_image_text:
            image_lines = _extract_docx_image_lines(paragraph, include_image_text)
            current_lines.extend(image_lines)

    if current_lines:
        sections.append({"title": current_title, "lines": current_lines[:]})

    return sections


def _split_text_document(source: Path, heading_mode: str, include_image_text: bool) -> list[dict[str, Any]]:
    lines = source.read_text(encoding="utf-8", errors="ignore").splitlines()

    sections: list[dict[str, Any]] = []
    current_title = "导言"
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
                    current_lines.append(f"[图片说明] {alt_text}")

    if current_lines:
        sections.append({"title": current_title, "lines": current_lines[:]})

    return sections


def _extract_docx_image_lines(paragraph: Any, include_image_text: bool) -> list[str]:
    if not include_image_text:
        return []

    image_lines: list[str] = []
    if qn is None:
        return image_lines

    blips = paragraph._element.xpath(".//a:blip")
    for blip in blips:
        embed_id = blip.get(qn("r:embed"))
        if not embed_id:
            continue

        image_part = paragraph.part.related_parts.get(embed_id)
        if image_part is None:
            continue

        ocr_text = _ocr_image_blob(image_part.blob)
        if ocr_text:
            image_lines.append(f"[图片文字] {ocr_text}")
        else:
            image_lines.append("[图片文字] 未识别到文字")

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


def _write_sections(source: Path, destination: Path, sections: list[dict[str, Any]]) -> list[Path]:
    destination.mkdir(parents=True, exist_ok=True)

    created_files: list[Path] = []
    index_payload: list[dict[str, Any]] = []

    for idx, section in enumerate(sections, start=1):
        title = str(section.get("title", "section")).strip() or f"section_{idx}"
        lines = [str(item) for item in section.get("lines", [])]

        safe_title = _safe_filename(title)
        output_path = destination / f"{source.stem}_{idx:03d}_{safe_title}.txt"
        text_content = "\n".join(lines).strip() + "\n"
        output_path.write_text(text_content, encoding="utf-8")

        created_files.append(output_path)
        index_payload.append({"index": idx, "title": title, "file": output_path.name, "line_count": len(lines)})

    index_file = destination / f"{source.stem}_split_index.json"
    index_file.write_text(json.dumps(index_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return created_files


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
