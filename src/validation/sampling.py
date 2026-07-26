"""표본 수 산정 및 무작위 표본 추출.

한국회계기준원 내부회계관리제도 모범규준 적용 FAQ 34에 근거한 표본 수 산정 기준을
구현한다. 사용자 확정에 따라 항상 위험도 "상단"(최대치)을 적용한다.

- 주기적 통제(periodic): 통제 실행 주기(cycle)에 따라 표본 수가 고정된다.
- 수시/비주기적 통제(event_driven): 실행 시점의 모집단 건수에 따라 표본 수가 정해진다.
  1~4건인 경우는 최대치 적용으로 전수조사한다.

계산된 표본 수가 모집단 건수보다 클 수 없으므로 항상 clamp한다.
"""
from __future__ import annotations

import random
from typing import Any

PERIODIC_SAMPLE_SIZES: dict[str, int] = {
    "annual": 1,
    "quarterly": 2,
    "monthly": 4,
    "weekly": 10,
    "daily": 40,
}

# (하한, 상한, 표본수) - 상한은 포함(inclusive). 상한이 None이면 무제한(초과 구간).
EVENT_DRIVEN_BANDS: list[tuple[int, int | None, int]] = [
    (1, 4, None),  # 전수조사: population_count 그대로 사용 (아래에서 처리)
    (5, 12, 4),
    (13, 52, 15),
    (53, 250, 25),
    (251, None, 60),
]


def determine_sample_size(control_type: str, cycle: str | None, population_count: int) -> int:
    """표본 수를 산정한다.

    control_type이 "periodic"이면 cycle에 따라, "event_driven"이면 population_count에
    따라 표본 수를 결정한 뒤, 항상 min(계산된 표본수, population_count)로 clamp한다.
    """
    if control_type == "periodic":
        if cycle not in PERIODIC_SAMPLE_SIZES:
            raise ValueError(f"알 수 없는 cycle: {cycle!r}")
        sample_size = PERIODIC_SAMPLE_SIZES[cycle]
    elif control_type == "event_driven":
        sample_size = _event_driven_sample_size(population_count)
    else:
        raise ValueError(f"알 수 없는 control_type: {control_type!r}")

    return min(sample_size, population_count)


def _event_driven_sample_size(population_count: int) -> int:
    if population_count <= 0:
        return 0
    if population_count <= 4:
        return population_count  # 전수조사
    for lower, upper, size in EVENT_DRIVEN_BANDS:
        if lower <= 4:
            continue  # 전수조사 구간은 위에서 처리했음
        if upper is None or population_count <= upper:
            return size
    return EVENT_DRIVEN_BANDS[-1][2]


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
