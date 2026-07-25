"""건수/금액 검증, 중복 체크, 담당자 매칭 실패 판정 -> 정상/예외 분류.

정상 건은 자동으로 확정하고, 문제 있는 건(중복·담당자미확인 등)만 "예외"로 분류해서
사람이 확인하게 하는 것이 목표다 — 정상 데이터는 사람이 볼 필요가 없다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

EXCEPTION_REASON_COLUMN = "예외사유"


@dataclass
class TotalsCheck:
    actual_count: int
    actual_amount: float
    expected_count: int | None
    expected_amount: float | None

    @property
    def count_mismatch(self) -> bool:
        return self.expected_count is not None and self.actual_count != self.expected_count

    @property
    def amount_mismatch(self) -> bool:
        return self.expected_amount is not None and not _isclose(
            self.actual_amount, self.expected_amount
        )


def _isclose(a: float, b: float, tol: float = 0.01) -> bool:
    return abs(a - b) <= tol


def validate_totals(
    rows: list[dict[str, Any]],
    amount_column: str | None,
    expected_count: int | None = None,
    expected_amount: float | None = None,
) -> TotalsCheck:
    """건수·금액 합계를 계산하고, 기대값이 주어지면 대사(mismatch 여부)까지 판단한다."""
    actual_count = len(rows)
    actual_amount = 0.0
    if amount_column:
        for row in rows:
            try:
                actual_amount += float(row.get(amount_column))
            except (TypeError, ValueError):
                continue
    return TotalsCheck(
        actual_count=actual_count,
        actual_amount=actual_amount,
        expected_count=expected_count,
        expected_amount=expected_amount,
    )


def find_duplicate_indices(rows: list[dict[str, Any]], key_columns: list[str]) -> set[int]:
    """key_columns 값이 모두 같은 행들을 중복으로 판단하고, 해당 그룹 전체(첫 건 포함)의
    인덱스를 반환한다."""
    if not key_columns:
        return set()

    seen: dict[tuple[Any, ...], list[int]] = {}
    for i, row in enumerate(rows):
        key = tuple(row.get(c) for c in key_columns)
        seen.setdefault(key, []).append(i)

    duplicates: set[int] = set()
    for indices in seen.values():
        if len(indices) > 1:
            duplicates.update(indices)
    return duplicates


def classify_rows(
    rows: list[dict[str, Any]],
    key_columns: list[str] | None = None,
    name_output_column: str = "담당자명",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """중복 건이나 담당자 매칭 실패 건(name_output_column이 비어있는 행)만 예외로 분류하고,
    나머지는 정상으로 분류한다. 담당자 매칭은 lookup.hr_matcher.match_rows에서 미리 수행되어
    있어야 한다.
    """
    key_columns = key_columns or []
    duplicate_indices = find_duplicate_indices(rows, key_columns)

    normal_rows: list[dict[str, Any]] = []
    exception_rows: list[dict[str, Any]] = []

    for i, row in enumerate(rows):
        reasons = []
        if i in duplicate_indices:
            reasons.append("중복")
        if name_output_column in row and not row.get(name_output_column):
            reasons.append("담당자미확인")

        if reasons:
            exception_row = dict(row)
            exception_row[EXCEPTION_REASON_COLUMN] = ", ".join(reasons)
            exception_rows.append(exception_row)
        else:
            normal_rows.append(dict(row))

    return normal_rows, exception_rows
