from __future__ import annotations

import json
from pathlib import Path

from .models import RunReport


def write_report(report: RunReport, output_path: Path | None) -> Path | None:
    if output_path is None:
        return None

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def print_summary(report: RunReport) -> None:
    summary = report.summary()
    print(f"Command: {summary['command']}")
    print(f"Workspace: {summary['workspace']}")
    print(f"Dry run mode: {summary['dry_run_mode']}")
    print(
        "Results: "
        f"total={summary['total']} "
        f"success={summary['success']} "
        f"dry_run={summary['dry_run']} "
        f"skipped={summary['skipped']} "
        f"failed={summary['failed']}"
    )
