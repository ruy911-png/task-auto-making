"""별도 참조 엑셀(예: 전표번호 -> 담당자 매핑표) 업로드 기반 담당자 매칭.

인사데이터 엑셀에 담당자 정보가 바로 없는 통제를 위한 방식이다. SAP를 전혀
접촉하지 않고, 사용자가 올린 참조 엑셀을 키 컬럼 기준으로 XLOOKUP처럼 대사한다.
"""
from __future__ import annotations

from typing import Any

ReferenceIndex = dict[str, dict[str, Any]]


def build_index(reference_rows: list[dict[str, Any]], key_column: str) -> ReferenceIndex:
    """참조 엑셀 행들을 key_column 값 기준 인덱스로 만든다."""
    index: ReferenceIndex = {}
    for row in reference_rows:
        key = row.get(key_column)
        if key is None:
            continue
        index[str(key).strip()] = row
    return index


def match_rows(
    rows: list[dict[str, Any]],
    reference_rows: list[dict[str, Any]],
    match_column: str,
    reference_key_column: str,
    name_column: str,
    department_column: str | None = None,
    name_output_column: str = "담당자명",
    department_output_column: str = "담당부서",
) -> list[dict[str, Any]]:
    """rows의 match_column 값을 참조 엑셀의 reference_key_column과 대사해 담당자명/담당부서
    컬럼을 추가한다. 매칭 실패 시 두 컬럼 모두 None으로 남기고 파이프라인은 중단하지 않는다
    (예외 분류는 validation.checks가 담당).
    """
    index = build_index(reference_rows, reference_key_column)
    enriched = []
    for row in rows:
        raw_value = row.get(match_column)
        key = str(raw_value).strip() if raw_value is not None else None
        reference_row = index.get(key) if key else None

        new_row = dict(row)
        new_row[name_output_column] = reference_row.get(name_column) if reference_row else None
        new_row[department_output_column] = (
            reference_row.get(department_column)
            if reference_row and department_column
            else None
        )
        enriched.append(new_row)
    return enriched
