import pytest

from validation.sampling import determine_sample_size, select_sample


# ---- periodic ----

@pytest.mark.parametrize(
    "cycle, expected",
    [
        ("annual", 1),
        ("quarterly", 2),
        ("monthly", 4),
        ("weekly", 10),
        ("daily", 40),
    ],
)
def test_periodic_sample_size(cycle, expected):
    # population 충분히 큼 -> clamp 영향 없음
    assert determine_sample_size("periodic", cycle, population_count=1000) == expected


def test_periodic_unknown_cycle_raises():
    with pytest.raises(ValueError):
        determine_sample_size("periodic", "biweekly", population_count=100)


def test_periodic_none_cycle_raises():
    with pytest.raises(ValueError):
        determine_sample_size("periodic", None, population_count=100)


# ---- event_driven ----

@pytest.mark.parametrize(
    "population_count, expected",
    [
        (1, 1),
        (2, 2),
        (4, 4),  # 1~4건: 전수조사
        (5, 4),
        (12, 4),
        (13, 13),  # band 표본수는 15이지만 population이 13건뿐이라 clamp됨
        (20, 15),
        (52, 15),
        (53, 25),
        (250, 25),
        (251, 60),
        (1000, 60),
    ],
)
def test_event_driven_sample_size(population_count, expected):
    assert determine_sample_size("event_driven", None, population_count) == expected


def test_unknown_control_type_raises():
    with pytest.raises(ValueError):
        determine_sample_size("weird", None, population_count=10)


# ---- clamp ----

def test_periodic_clamped_to_population_count():
    # monthly 통제(최대치 4)인데 모집단이 2건뿐이면 표본은 2건
    assert determine_sample_size("periodic", "monthly", population_count=2) == 2


def test_event_driven_clamped_to_population_count():
    # 5~12건 구간 표본수는 4이지만 모집단이 3건이면 3건(전수조사 구간이므로 애초에 3)
    assert determine_sample_size("event_driven", None, population_count=3) == 3


def test_event_driven_clamped_when_band_size_exceeds_population():
    # 이례적으로 population_count가 band 하한보다 작은 경우에도 clamp 확인
    assert determine_sample_size("event_driven", None, population_count=0) == 0


# ---- select_sample ----

def test_select_sample_returns_requested_count():
    population = [{"id": i} for i in range(20)]
    sample = select_sample(population, sample_size=5, random_state=42)
    assert len(sample) == 5
    # 모두 population 안에서 뽑힌 것인지 확인
    ids = {row["id"] for row in sample}
    assert ids.issubset({row["id"] for row in population})
    # 비복원추출: 중복 없음
    assert len(ids) == 5


def test_select_sample_reproducible_with_random_state():
    population = [{"id": i} for i in range(50)]
    sample_1 = select_sample(population, sample_size=10, random_state=7)
    sample_2 = select_sample(population, sample_size=10, random_state=7)
    assert sample_1 == sample_2


def test_select_sample_different_random_state_can_differ():
    population = [{"id": i} for i in range(50)]
    sample_1 = select_sample(population, sample_size=10, random_state=1)
    sample_2 = select_sample(population, sample_size=10, random_state=2)
    assert sample_1 != sample_2


def test_select_sample_size_gte_population_returns_all():
    population = [{"id": i} for i in range(3)]
    sample = select_sample(population, sample_size=10, random_state=1)
    assert len(sample) == 3
    assert sample == population


def test_select_sample_size_equal_population_returns_all():
    population = [{"id": i} for i in range(5)]
    sample = select_sample(population, sample_size=5, random_state=1)
    assert len(sample) == 5
    assert {row["id"] for row in sample} == {row["id"] for row in population}
