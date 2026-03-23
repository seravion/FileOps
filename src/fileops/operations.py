from __future__ import annotations

import glob
import shutil
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .models import OperationResult, OperationStatus
from .utils import duration_ms, ensure_workspace_path, now_iso, unique_path

try:
    from send2trash import send2trash
except ImportError:  # pragma: no cover
    send2trash = None


OVERWRITE_POLICIES = {"never", "always", "rename"}


@dataclass(slots=True)
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
