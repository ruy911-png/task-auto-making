"""표본 수 산정 및 무작위 표본 추출.

한국회계기준원 내부회계관리제도 모범규준 적용 FAQ 34에 근거한 표본 수 산정 기준을
구현한다. 사용자 확정에 따라 항상 위험도 "상단"(최대치)을 적용한다.

표본 수는 실행 시점의 실제 모집단 건수 하나만으로 정해진다 (통제의 명목상 주기와
무관하게, 그 조회의 결과 건수를 기준으로 판단). 1~4건인 경우는 최대치 적용으로
전수조사한다.

계산된 표본 수가 모집단 건수보다 클 수 없으므로 항상 clamp한다.
"""
from __future__ import annotations

import random
from typing import Any

# (하한, 상한, 표본수) - 상한은 포함(inclusive). 상한이 None이면 무제한(초과 구간).
SAMPLE_SIZE_BANDS: list[tuple[int, int | None, int]] = [
    (1, 4, None),  # 전수조사: population_count 그대로 사용 (아래에서 처리)
    (5, 12, 4),
    (13, 52, 15),
    (53, 250, 25),
    (251, None, 60),
]


def determine_sample_size(population_count: int) -> int:
    """모집단 건수를 기준으로 표본 수를 산정한다 (항상 min(계산값, population_count)로 clamp)."""
    if population_count <= 0:
        return 0
    if population_count <= 4:
        return population_count  # 전수조사

    for lower, upper, size in SAMPLE_SIZE_BANDS:
        if lower <= 4:
            continue  # 전수조사 구간은 위에서 처리했음
        if upper is None or population_count <= upper:
            return min(size, population_count)

    return SAMPLE_SIZE_BANDS[-1][2]


def select_sample(
    population_rows: list[dict[str, Any]],
    sample_size: int,
    random_state: int | None = None,
) -> list[dict[str, Any]]:
    """population_rows에서 sample_size개를 비복원추출한다.

    sample_size가 population_rows 길이 이상이면 전체를 반환한다.
    random_state를 주면 재현 가능하다.
    """
    if sample_size >= len(population_rows):
        return list(population_rows)

    rng = random.Random(random_state)
    return rng.sample(population_rows, sample_size)
