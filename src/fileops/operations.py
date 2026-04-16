from __future__ import annotations

import glob
import math
import shutil
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path

from .models import OperationResult, OperationStatus
from .utils import duration_ms, ensure_workspace_path, now_iso, unique_path

try:
    from send2trash import send2trash
except ImportError:  # pragma: no cover
    send2trash = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:  # pragma: no cover
    PdfReader = None
    PdfWriter = None

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:  # pragma: no cover
    Document = None
    qn = None


OVERWRITE_POLICIES = {"never", "always", "rename"}


@dataclass
class CommonOptions:
    workspace: Path
    dry_run: bool = False
    overwrite: str = "never"


def expand_sources(patterns: Sequence[str], recursive: bool = True, include_dirs: bool = True) -> list[Path]:
    results: list[Path] = []
    seen: set[Path] = set()

    for raw_pattern in patterns:
        pattern = raw_pattern.strip()
        if not pattern:
            continue

        raw_matches = _glob_or_path(pattern, recursive=recursive)
        if not raw_matches:
            continue

        for matched in raw_matches:
            candidate = matched.resolve(strict=False)
            if not include_dirs and candidate.is_dir():
                continue
            if candidate in seen:
                continue
            seen.add(candidate)
            results.append(candidate)

    return results


def copy_items(sources: Sequence[Path], destination: Path, options: CommonOptions) -> list[OperationResult]:
    _validate_overwrite(options.overwrite)
    destination = destination.resolve(strict=False)
    ensure_workspace_path(destination, options.workspace)

    multiple_sources = len(sources) > 1
    if multiple_sources and not options.dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    results: list[OperationResult] = []
    for source in sources:
        results.append(
            _run_transfer(
                operation="copy",
                source=source,
                destination=_build_target_path(source, destination, multiple_sources),
                options=options,
                transfer_impl=_copy_impl,
            )
        )
    return results


def move_items(sources: Sequence[Path], destination: Path, options: CommonOptions) -> list[OperationResult]:
    _validate_overwrite(options.overwrite)
    destination = destination.resolve(strict=False)
    ensure_workspace_path(destination, options.workspace)

    multiple_sources = len(sources) > 1
    if multiple_sources and not options.dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    results: list[OperationResult] = []
    for source in sources:
        results.append(
            _run_transfer(
                operation="move",
                source=source,
                destination=_build_target_path(source, destination, multiple_sources),
                options=options,
                transfer_impl=_move_impl,
            )
        )
    return results


def rename_items(
    sources: Sequence[Path],
    pattern: str,
    start_index: int,
    options: CommonOptions,
) -> list[OperationResult]:
    _validate_overwrite(options.overwrite)
    results: list[OperationResult] = []

    current_index = start_index
    today = datetime.now().strftime("%Y%m%d")

    for source in sources:
        started = datetime.now()
        started_at = now_iso()

        try:
            ensure_workspace_path(source, options.workspace)
            if not source.exists():
                raise FileNotFoundError(f"Source does not exist: {source}")

            values = {
                "stem": source.stem,
                "ext": source.suffix,
                "index": current_index,
                "parent": source.parent.name,
                "date": today,
                "name": source.name,
            }
            target_name = pattern.format_map(values)
            target = source.with_name(target_name)
            current_index += 1

            if source == target:
                results.append(_build_result("rename", source, target, OperationStatus.SKIPPED, "Source and destination are identical.", started, started_at))
                continue

            ensure_workspace_path(target, options.workspace)
            target, note = _resolve_conflict(target, options)
            if target is None:
                results.append(_build_result("rename", source, None, OperationStatus.SKIPPED, note or "Destination exists.", started, started_at))
                continue

            if options.dry_run:
                message = f"Would rename to {target}"
                if note:
                    message = f"{message}. {note}"
                results.append(_build_result("rename", source, target, OperationStatus.DRY_RUN, message, started, started_at))
                continue

            source.rename(target)
            message = f"Renamed to {target}"
            if note:
                message = f"{message}. {note}"
            results.append(_build_result("rename", source, target, OperationStatus.SUCCESS, message, started, started_at))

        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("rename", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def delete_items(
    sources: Sequence[Path],
    workspace: Path,
    dry_run: bool,
    use_trash: bool,
) -> list[OperationResult]:
    results: list[OperationResult] = []

    if use_trash and send2trash is None:
        raise RuntimeError("send2trash is not installed. Install dependencies or use --hard.")

    for source in sources:
        started = datetime.now()
        started_at = now_iso()

        try:
            ensure_workspace_path(source, workspace)
            if not source.exists():
                results.append(_build_result("delete", source, None, OperationStatus.SKIPPED, "Source does not exist.", started, started_at))
                continue

            if dry_run:
                action = "trash" if use_trash else "hard-delete"
                results.append(_build_result("delete", source, None, OperationStatus.DRY_RUN, f"Would {action} {source}", started, started_at))
                continue

            if use_trash:
                send2trash(str(source))
                results.append(_build_result("delete", source, None, OperationStatus.SUCCESS, "Moved to trash.", started, started_at))
                continue

            _remove_path(source)
            results.append(_build_result("delete", source, None, OperationStatus.SUCCESS, "Deleted permanently.", started, started_at))

        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("delete", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def split_items(
    sources: Sequence[Path],
    destination: Path,
    chunk_size_mb: float,
    options: CommonOptions,
) -> list[OperationResult]:
    _validate_overwrite(options.overwrite)
    if chunk_size_mb <= 0:
        raise ValueError("Split chunk size must be greater than 0 MB.")

    destination = destination.resolve(strict=False)
    ensure_workspace_path(destination, options.workspace)

    chunk_size_bytes = int(chunk_size_mb * 1024 * 1024)
    if chunk_size_bytes <= 0:
        raise ValueError("Split chunk size is too small.")

    if not options.dry_run:
        destination.mkdir(parents=True, exist_ok=True)

    results: list[OperationResult] = []

    for source in sources:
        started = datetime.now()
        started_at = now_iso()

        try:
            ensure_workspace_path(source, options.workspace)
            if not source.exists():
                raise FileNotFoundError(f"Source does not exist: {source}")
            if source.is_dir():
                raise IsADirectoryError(f"Split supports files only: {source}")

            file_size = source.stat().st_size
            desired_parts = max(1, math.ceil(file_size / chunk_size_bytes))

            pdf_groups: list[list[int]] | None = None
            docx_ranges: list[tuple[int, int]] | None = None

            if _is_pdf_source(source):
                pdf_reader = _load_pdf_reader(source)
                pdf_groups = _build_pdf_page_groups(pdf_reader, chunk_size_bytes, desired_parts)
                part_count = len(pdf_groups)
            elif _is_docx_source(source):
                docx_document = _load_docx_document(source)
                docx_ranges = _build_docx_block_ranges(docx_document, chunk_size_bytes, desired_parts)
                part_count = len(docx_ranges)
            else:
                part_count = desired_parts

            target_paths: list[Path] = []
            note_messages: list[str] = []
            conflict_message: str | None = None

            for part_index in range(1, part_count + 1):
                candidate = _build_split_part_path(source, destination, part_index)
                resolved, note = _resolve_conflict(candidate, options)
                if resolved is None:
                    conflict_message = note or f"Destination exists: {candidate}"
                    break
                target_paths.append(resolved)
                if note:
                    note_messages.append(note)

            if conflict_message:
                results.append(_build_result("split", source, None, OperationStatus.SKIPPED, conflict_message, started, started_at))
                continue

            if options.dry_run:
                message = f"Would split into {part_count} part(s) at {destination}"
                if note_messages:
                    message = f"{message}. {note_messages[0]}"
                results.append(_build_result("split", source, destination, OperationStatus.DRY_RUN, message, started, started_at))
                continue

            if pdf_groups is not None:
                pdf_reader = _load_pdf_reader(source)
                for target, page_indexes in zip(target_paths, pdf_groups):
                    target.parent.mkdir(parents=True, exist_ok=True)
                    _write_pdf_part(pdf_reader, page_indexes, target)
            elif docx_ranges is not None:
                for target, (start_block, end_block) in zip(target_paths, docx_ranges):
                    target.parent.mkdir(parents=True, exist_ok=True)
                    _write_docx_part(source, start_block, end_block, target)
            else:
                with source.open("rb") as reader:
                    for target in target_paths:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        chunk = reader.read(chunk_size_bytes)
                        with target.open("wb") as writer:
                            writer.write(chunk)

            message = f"Split completed: {len(target_paths)} part(s) in {destination}"
            if note_messages:
                message = f"{message}. {note_messages[0]}"
            results.append(_build_result("split", source, destination, OperationStatus.SUCCESS, message, started, started_at))

        except Exception as exc:  # noqa: BLE001
            results.append(_build_result("split", source, None, OperationStatus.FAILED, str(exc), started, started_at))

    return results


def _is_pdf_source(source: Path) -> bool:
    return source.suffix.lower() == ".pdf"


def _is_docx_source(source: Path) -> bool:
    return source.suffix.lower() == ".docx"


def _load_pdf_reader(source: Path):
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Please install dependencies first.")

    try:
        reader = PdfReader(str(source))
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f".pdf parse failed: {exc}") from exc

    if getattr(reader, "is_encrypted", False):
        try:
            decrypted = reader.decrypt("")
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f".pdf parse failed: {exc}") from exc
        if decrypted == 0:
            raise ValueError(".pdf parse failed: encrypted PDF requires password")

    return reader


def _load_docx_document(source: Path):
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please install dependencies first.")

    try:
        return Document(str(source))
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f".docx parse failed: {exc}") from exc


def _build_pdf_page_groups(pdf_reader, chunk_size_bytes: int, desired_parts: int) -> list[list[int]]:
    total_pages = len(pdf_reader.pages)
    if total_pages == 0:
        return [[]]

    target_parts = max(1, min(total_pages, desired_parts))

    groups: list[list[int]] = []
    current: list[int] = []

    for page_index in range(total_pages):
        candidate_pages = current + [page_index]
        candidate_size = _estimate_pdf_size_for_pages(pdf_reader, candidate_pages)

        if current and candidate_size > chunk_size_bytes:
            groups.append(current)
            current = [page_index]
        else:
            current = candidate_pages

    if current:
        groups.append(current)

    return _rebalance_groups_to_target(groups, target_parts)


def _iter_docx_body_blocks(doc) -> list:
    if qn is None:
        return []
    body = doc.element.body
    return [child for child in body.iterchildren() if child.tag in {qn("w:p"), qn("w:tbl")}]


def _build_docx_block_ranges(doc, chunk_size_bytes: int, desired_parts: int) -> list[tuple[int, int]]:
    blocks = _iter_docx_body_blocks(doc)
    total_blocks = len(blocks)
    if total_blocks == 0:
        return [(0, 0)]

    target_parts = max(1, min(total_blocks, desired_parts))

    groups: list[list[int]] = []
    current: list[int] = []
    current_size = 0

    for block_index, block in enumerate(blocks):
        block_size = max(1, len(block.xml.encode("utf-8")))
        if current and current_size + block_size > chunk_size_bytes:
            groups.append(current)
            current = [block_index]
            current_size = block_size
        else:
            current.append(block_index)
            current_size += block_size

    if current:
        groups.append(current)

    balanced = _rebalance_groups_to_target(groups, target_parts)
    return [(group[0], group[-1] + 1) for group in balanced]


def _rebalance_groups_to_target(groups: list[list[int]], target_parts: int) -> list[list[int]]:
    normalized = [group[:] for group in groups if group]
    if not normalized:
        return [[]]

    while len(normalized) < target_parts:
        split_index = -1
        split_length = 1
        for idx, group in enumerate(normalized):
            if len(group) > split_length:
                split_index = idx
                split_length = len(group)

        if split_index < 0:
            break

        group = normalized.pop(split_index)
        mid = len(group) // 2
        normalized.insert(split_index, group[:mid])
        normalized.insert(split_index + 1, group[mid:])

    return normalized


def _estimate_pdf_size_for_pages(pdf_reader, page_indexes: Sequence[int]) -> int:
    if PdfWriter is None:
        raise RuntimeError("pypdf is not installed. Please install dependencies first.")

    writer = PdfWriter()
    for page_index in page_indexes:
        writer.add_page(pdf_reader.pages[page_index])

    buffer = BytesIO()
    writer.write(buffer)
    return buffer.tell()


def _write_pdf_part(pdf_reader, page_indexes: Sequence[int], target: Path) -> None:
    if PdfWriter is None:
        raise RuntimeError("pypdf is not installed. Please install dependencies first.")

    writer = PdfWriter()
    for page_index in page_indexes:
        writer.add_page(pdf_reader.pages[page_index])

    with target.open("wb") as stream:
        writer.write(stream)


def _write_docx_part(source: Path, start_block: int, end_block: int, target: Path) -> None:
    shutil.copy2(source, target)

    doc = _load_docx_document(target)
    blocks = _iter_docx_body_blocks(doc)

    for idx, block in enumerate(blocks):
        if start_block <= idx < end_block:
            continue
        block.getparent().remove(block)

    if not _iter_docx_body_blocks(doc):
        doc.add_paragraph("")

    doc.save(str(target))




def _run_transfer(
    operation: str,
    source: Path,
    destination: Path,
    options: CommonOptions,
    transfer_impl,
) -> OperationResult:
    started = datetime.now()
    started_at = now_iso()

    try:
        ensure_workspace_path(source, options.workspace)
        if not source.exists():
            raise FileNotFoundError(f"Source does not exist: {source}")

        ensure_workspace_path(destination, options.workspace)
        if source == destination:
            return _build_result(operation, source, destination, OperationStatus.SKIPPED, "Source and destination are identical.", started, started_at)

        destination, note = _resolve_conflict(destination, options)
        if destination is None:
            return _build_result(operation, source, None, OperationStatus.SKIPPED, note or "Destination exists.", started, started_at)

        if options.dry_run:
            message = f"Would {operation} to {destination}"
            if note:
                message = f"{message}. {note}"
            return _build_result(operation, source, destination, OperationStatus.DRY_RUN, message, started, started_at)

        transfer_impl(source, destination)
        message = f"{operation.capitalize()} completed: {destination}"
        if note:
            message = f"{message}. {note}"
        return _build_result(operation, source, destination, OperationStatus.SUCCESS, message, started, started_at)

    except Exception as exc:  # noqa: BLE001
        return _build_result(operation, source, None, OperationStatus.FAILED, str(exc), started, started_at)


def _build_target_path(source: Path, destination: Path, multiple_sources: bool) -> Path:
    if multiple_sources:
        return destination / source.name
    if destination.exists() and destination.is_dir():
        return destination / source.name
    return destination


def _resolve_conflict(destination: Path, options: CommonOptions) -> tuple[Path | None, str | None]:
    if not destination.exists():
        return destination, None

    if options.overwrite == "never":
        return None, f"Destination exists: {destination}"

    if options.overwrite == "rename":
        new_path = unique_path(destination)
        return new_path, f"Destination exists. Auto-renamed to {new_path.name}"

    if options.overwrite == "always":
        if not options.dry_run:
            _remove_path(destination)
        return destination, "Existing destination overwritten."

    raise ValueError(f"Invalid overwrite policy: {options.overwrite}")


def _copy_impl(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir():
        shutil.copytree(source, destination)
        return
    shutil.copy2(source, destination)


def _move_impl(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(source), str(destination))


def _remove_path(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
        return
    path.unlink(missing_ok=True)


def _glob_or_path(pattern: str, recursive: bool) -> list[Path]:
    plain = Path(pattern)
    if plain.exists():
        return [plain]

    globbed = sorted(glob.glob(pattern, recursive=recursive))
    return [Path(item) for item in globbed]


def _build_split_part_path(source: Path, destination: Path, part_index: int) -> Path:
    return destination / f"{source.stem}.part{part_index:03d}{source.suffix}"


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


def _validate_overwrite(policy: str) -> None:
    if policy not in OVERWRITE_POLICIES:
        raise ValueError(f"Invalid overwrite policy '{policy}'. Expected one of {sorted(OVERWRITE_POLICIES)}")
