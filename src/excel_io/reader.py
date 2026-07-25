"""조건 엑셀 및 인사데이터 엑셀 리더."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .schema import HR_COLUMN_ALIASES, HRRecord, resolve_column


def read_conditions(path: str | Path) -> list[dict[str, Any]]:
    """조회 조건 엑셀(여러 행)을 읽어 행별 딕셔너리 목록으로 반환한다.

    각 행의 컬럼은 선택된 통제(config/transactions/<control_id>.yaml)의
    조회 필드 매핑에 따라 나중에 SAP 화면 필드로 매핑된다.
    """
    df = pd.read_excel(path, dtype=str)
    df = df.dropna(how="all")
    records = df.to_dict(orient="records")
    return [{k: _clean(v) for k, v in row.items()} for row in records]


def read_hr_data(path: str | Path) -> list[HRRecord]:
    """인사데이터 엑셀(사번/이름/아이디/부서)을 읽어 HRRecord 목록으로 반환한다."""
    df = pd.read_excel(path, dtype=str)
    df = df.dropna(how="all")
    headers = list(df.columns)

    col_map: dict[str, str] = {}
    for field, aliases in HR_COLUMN_ALIASES.items():
        column = resolve_column(headers, aliases)
        if column is None:
            raise ValueError(
                f"인사데이터 엑셀에서 '{field}'에 해당하는 컬럼을 찾을 수 없습니다 "
                f"(지원 별칭: {aliases})"
            )
        col_map[field] = column

    records = []
    for _, row in df.iterrows():
        records.append(
            HRRecord(
                emp_id=_clean(row[col_map["emp_id"]]),
                name=_clean(row[col_map["name"]]),
                user_id=_clean(row[col_map["user_id"]]),
                department=_clean(row[col_map["department"]]),
            )
        )
    return records


def _clean(value: Any) -> str | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text or None
