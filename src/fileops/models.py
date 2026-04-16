from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class OperationStatus(str, Enum):
    SUCCESS = "success"
    SKIPPED = "skipped"
    FAILED = "failed"
    DRY_RUN = "dry_run"


@dataclass
class OperationResult:
    operation: str
    source: str
    destination: str | None
    status: OperationStatus
    message: str
    started_at: str
    finished_at: str
    duration_ms: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "operation": self.operation,
            "source": self.source,
            "destination": self.destination,
            "status": self.status.value,
            "message": self.message,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "duration_ms": self.duration_ms,
        }


@dataclass
class RunReport:
    command: str
    dry_run_mode: bool
    workspace: str
    results: list[OperationResult] = field(default_factory=list)

    def add(self, result: OperationResult) -> None:
        self.results.append(result)

    def summary(self) -> dict[str, Any]:
        counts = {
            OperationStatus.SUCCESS.value: 0,
            OperationStatus.SKIPPED.value: 0,
            OperationStatus.FAILED.value: 0,
            OperationStatus.DRY_RUN.value: 0,
        }
        for item in self.results:
            counts[item.status.value] += 1

        return {
            "command": self.command,
            "workspace": self.workspace,
            "dry_run_mode": self.dry_run_mode,
            "total": len(self.results),
            "success": counts[OperationStatus.SUCCESS.value],
            "skipped": counts[OperationStatus.SKIPPED.value],
            "failed": counts[OperationStatus.FAILED.value],
            "dry_run": counts[OperationStatus.DRY_RUN.value],
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": self.summary(),
            "results": [item.to_dict() for item in self.results],
        }
