from excel_io.schema import HRRecord
from lookup.hr_matcher import build_index, match, match_rows


def _sample_hr():
    return [
        HRRecord(emp_id="1001", name="홍길동", user_id="hong1", department="회계팀"),
        HRRecord(emp_id="1002", name="김철수", user_id="kim2", department="구매팀"),
    ]


def test_match_by_emp_id():
    index = build_index(_sample_hr())
    result = match("1001", index)
    assert result.matched
    assert result.name == "홍길동"
    assert result.matched_by == "emp_id"


def test_match_by_name_when_not_emp_id():
    index = build_index(_sample_hr())
    result = match("김철수", index)
    assert result.matched_by == "name"
    assert result.department == "구매팀"


def test_match_by_user_id_last_priority():
    index = build_index(_sample_hr())
    result = match("hong1", index)
    assert result.matched_by == "user_id"


def test_match_failure():
    index = build_index(_sample_hr())
    result = match("없는값", index)
    assert not result.matched
    assert result.name is None


def test_match_rows_adds_columns():
    rows = [{"담당자": "1001"}, {"담당자": "no-such"}]
    enriched = match_rows(rows, _sample_hr(), responsible_column="담당자")

    assert enriched[0]["담당자명"] == "홍길동"
    assert enriched[0]["담당부서"] == "회계팀"
    assert enriched[1]["담당자명"] is None
