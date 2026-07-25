"""엑셀 컬럼 스키마 및 별칭 정의.

인사데이터 엑셀은 회사/부서마다 헤더 이름이 다를 수 있어 별칭 목록으로 매칭한다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional

HR_COLUMN_ALIASES: dict[str, list[str]] = {
    "emp_id": ["사번", "사원번호", "employee_id", "emp_id"],
    "name": ["이름", "성명", "name"],
    "user_id": ["아이디", "id", "user_id", "사용자아이디"],
    "department": ["담당부서", "부서", "department", "부서명"],
}


@dataclass(frozen=True)
class HRRecord:
    emp_id: Optional[str]
    name: Optional[str]
    user_id: Optional[str]
    department: Optional[str]


def resolve_column(headers: Iterable[str], aliases: list[str]) -> Optional[str]:
    """헤더 목록에서 별칭 중 하나와 일치하는 실제 컬럼명을 찾는다 (대소문자/공백 무시)."""
    normalized = {str(h).strip().lower(): h for h in headers}
    for alias in aliases:
        key = alias.strip().lower()
        if key in normalized:
            return normalized[key]
    return None
