import pytest

from validation.sampling import determine_sample_size, select_sample


# ---- determine_sample_size (모집단 건수 기준, 항상 최대치) ----

@pytest.mark.parametrize(
    "population_count, expected",
    [
        (0, 0),
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
def test_sample_size_by_population_count(population_count, expected):
    assert determine_sample_size(population_count) == expected


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
