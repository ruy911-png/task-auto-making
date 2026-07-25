"""통제 YAML의 편집 스키마(컬럼정리/필터/계산컬럼)를 다운로드 데이터에 적용하는 범용 실행기.

편집 규칙 스키마 (ControlConfig.edit_rules에 저장되는 dict):

    {
      "rename_columns": {"원본컬럼명": "새이름"},
      "drop_columns": ["컬럼명", ...],
      "column_order": ["컬럼명", ...],
      "filters": [{"column": "금액", "op": "!=", "value": 0}, ...],
      "calculated_columns": [
        {"name": "합계금액", "op": "multiply", "columns": ["단가", "수량"]},
        {"name": "비고", "op": "concat", "columns": ["회사코드", "전표번호"], "separator": "-"},
      ],
    }

새 통제를 추가할 때 코드를 수정하지 않고 이 설정만으로 컬럼정리/필터/계산컬럼을 조합할 수
있도록, 임의 코드 실행(eval)이 아닌 안전한 선언적 연산만 지원한다.
"""
from __future__ import annotations

import operator
from typing import Any

import pandas as pd

_FILTER_OPS = {
    "==": operator.eq,
    "!=": operator.ne,
    ">": operator.gt,
    ">=": operator.ge,
    "<": operator.lt,
    "<=": operator.le,
}

_CALC_OPS = {"add", "subtract", "multiply", "divide", "concat"}


def apply_edit_rules(
    rows: list[dict[str, Any]], edit_rules: dict[str, Any]
) -> list[dict[str, Any]]:
    """rows에 편집 규칙을 순서대로 적용한다: 컬럼정리 -> 필터 -> 계산컬럼 -> 순서정리."""
    if not rows:
        return rows

    df = pd.DataFrame(rows)
    df = _rename_columns(df, edit_rules.get("rename_columns", {}))
    df = _drop_columns(df, edit_rules.get("drop_columns", []))
    df = _apply_filters(df, edit_rules.get("filters", []))
    df = _add_calculated_columns(df, edit_rules.get("calculated_columns", []))
    df = _reorder_columns(df, edit_rules.get("column_order"))
    return df.to_dict(orient="records")


def _rename_columns(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    return df.rename(columns=mapping) if mapping else df


def _drop_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    if not columns:
        return df
    existing = [c for c in columns if c in df.columns]
    return df.drop(columns=existing) if existing else df


def _apply_filters(df: pd.DataFrame, filters: list[dict[str, Any]]) -> pd.DataFrame:
    for f in filters:
        column, op, value = f.get("column"), f.get("op"), f.get("value")
        if column not in df.columns or op not in _FILTER_OPS:
            continue
        df = df[_FILTER_OPS[op](df[column], value)]
    return df


def _add_calculated_columns(
    df: pd.DataFrame, calculated_columns: list[dict[str, Any]]
) -> pd.DataFrame:
    for calc in calculated_columns:
        name, op, columns = calc.get("name"), calc.get("op"), calc.get("columns", [])
        if op not in _CALC_OPS or any(c not in df.columns for c in columns) or not columns:
            continue

        if op == "concat":
            separator = calc.get("separator", "")
            df[name] = df[columns].astype(str).agg(separator.join, axis=1)
            continue

        numeric = df[columns].apply(pd.to_numeric, errors="coerce")
        result = numeric.iloc[:, 0]
        for c in numeric.columns[1:]:
            if op == "add":
                result = result + numeric[c]
            elif op == "subtract":
                result = result - numeric[c]
            elif op == "multiply":
                result = result * numeric[c]
            elif op == "divide":
                result = result / numeric[c]
        df[name] = result
    return df


def _reorder_columns(df: pd.DataFrame, column_order: list[str] | None) -> pd.DataFrame:
    if not column_order:
        return df
    ordered = [c for c in column_order if c in df.columns]
    remaining = [c for c in df.columns if c not in ordered]
    return df[ordered + remaining]
