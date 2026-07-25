"""인사데이터 기반 담당자/담당부서 매칭.

ERP 결과에서 담당자를 나타내는 값이 건별로 사번/이름/아이디 중 무엇으로 들어있는지가
다르기 때문에, 사번 -> 이름 -> 아이디 순서로 순차 매칭을 시도한다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from excel_io.schema import HRRecord

MATCH_PRIORITY = ("emp_id", "name", "user_id")


@dataclass(frozen=True)
class MatchResult:
    name: str | None
    department: str | None
    matched_by: str | None  # "emp_id" | "name" | "user_id" | None(매칭 실패)

    @property
    def matched(self) -> bool:
        return self.matched_by is not None


HRIndex = dict[str, dict[str, HRRecord]]


def build_index(hr_records: list[HRRecord]) -> HRIndex:
    """사번/이름/아이디 각각을 키로 하는 조회 인덱스를 만든다."""
    index: HRIndex = {key: {} for key in MATCH_PRIORITY}
    for record in hr_records:
        if record.emp_id:
            index["emp_id"][record.emp_id] = record
        if record.name:
            index["name"][record.name] = record
        if record.user_id:
            index["user_id"][record.user_id] = record
    return index


def match(raw_value: str | None, index: HRIndex) -> MatchResult:
    """raw_value(사번/이름/아이디 중 하나로 추정되는 값)를 사번 -> 이름 -> 아이디 순으로 매칭한다."""
    if not raw_value:
        return MatchResult(name=None, department=None, matched_by=None)

    value = raw_value.strip()
    for key in MATCH_PRIORITY:
        record = index[key].get(value)
        if record is not None:
            return MatchResult(name=record.name, department=record.department, matched_by=key)

    return MatchResult(name=None, department=None, matched_by=None)


def match_rows(
    rows: list[dict[str, Any]],
    hr_records: list[HRRecord],
    responsible_column: str,
    name_output_column: str = "담당자명",
    department_output_column: str = "담당부서",
) -> list[dict[str, Any]]:
    """rows의 responsible_column(원본 담당자 식별값)을 기준으로 담당자명/담당부서 컬럼을
    추가한다. 매칭 실패 시 두 컬럼 모두 None으로 남기고, 예외 분류는 validation.checks가
    담당한다 (여기서는 파이프라인을 중단하지 않는다).
    """
    index = build_index(hr_records)
    enriched = []
    for row in rows:
        raw_value = row.get(responsible_column)
        result = match(str(raw_value) if raw_value is not None else None, index)
        new_row = dict(row)
        new_row[name_output_column] = result.name
        new_row[department_output_column] = result.department
        enriched.append(new_row)
    return enriched
