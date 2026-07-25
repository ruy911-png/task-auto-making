"""여러 조회 조건을 순회 실행하고, 조건별 성공/실패 상태를 추적한다.

핵심 원칙: 조건 하나가 실패해도 전체 배치를 중단하지 않는다. 실패한 조건은
상태 파일에 남아, 나중에 `retry_failed`로 그 조건들만 다시 실행할 수 있다.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Callable


class ConditionStatus(str, Enum):
    PENDING = "pending"
    SUCCESS = "success"
    FAILED = "failed"


@dataclass
class ConditionResult:
    index: int
    condition: dict[str, Any]
    status: ConditionStatus = ConditionStatus.PENDING
    error: str | None = None
    result: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["status"] = self.status.value
        return d

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "ConditionResult":
        d = dict(d)
        d["status"] = ConditionStatus(d["status"])
        return cls(**d)


@dataclass
class BatchState:
    job_id: str
    items: list[ConditionResult] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"job_id": self.job_id, "items": [i.to_dict() for i in self.items]}

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "BatchState":
        return cls(job_id=d["job_id"], items=[ConditionResult.from_dict(i) for i in d["items"]])

    def save(self, path: str | Path) -> None:
        Path(path).write_text(
            json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )

    @classmethod
    def load(cls, path: str | Path) -> "BatchState":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(data)

    @property
    def failed_items(self) -> list[ConditionResult]:
        return [i for i in self.items if i.status == ConditionStatus.FAILED]

    @property
    def success_items(self) -> list[ConditionResult]:
        return [i for i in self.items if i.status == ConditionStatus.SUCCESS]

    def summary(self) -> dict[str, int]:
        return {
            "total": len(self.items),
            "success": len(self.success_items),
            "failed": len(self.failed_items),
        }


ExecuteFn = Callable[[dict[str, Any]], dict[str, Any]]


def run_batch(
    job_id: str,
    conditions: list[dict[str, Any]],
    execute: ExecuteFn,
    state_path: str | Path,
) -> BatchState:
    """조건 목록을 순서대로 실행한다. 개별 조건 실패는 기록만 하고 계속 진행한다."""
    state = BatchState(
        job_id=job_id,
        items=[ConditionResult(index=i, condition=c) for i, c in enumerate(conditions)],
    )
    _run_items(state.items, execute)
    state.save(state_path)
    return state


def retry_failed(execute: ExecuteFn, state_path: str | Path) -> BatchState:
    """상태 파일에 저장된 조건 중 '실패'만 골라 다시 실행한다."""
    state = BatchState.load(state_path)
    failed = state.failed_items
    _run_items(failed, execute)
    state.save(state_path)
    return state


def _run_items(items: list[ConditionResult], execute: ExecuteFn) -> None:
    for item in items:
        try:
            item.result = execute(item.condition)
            item.status = ConditionStatus.SUCCESS
            item.error = None
        except Exception as exc:  # noqa: BLE001 - 조건 단위 실패를 흡수하고 배치는 계속 진행
            item.status = ConditionStatus.FAILED
            item.error = str(exc)
            item.result = None
