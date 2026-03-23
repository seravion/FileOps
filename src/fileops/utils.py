from __future__ import annotations

from datetime import datetime
from pathlib import Path


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def duration_ms(started: datetime, finished: datetime) -> int:
    delta = finished - started
    return int(delta.total_seconds() * 1000)


def ensure_workspace_path(path: Path, workspace: Path) -> None:
    resolved_path = path.resolve(strict=False)
    resolved_workspace = workspace.resolve(strict=False)
    if not _is_relative_to(resolved_path, resolved_workspace):
        raise ValueError(f"Path is outside workspace: {resolved_path}")


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent

    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1
