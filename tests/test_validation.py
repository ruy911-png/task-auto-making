from validation.checks import classify_rows, find_duplicate_indices, validate_totals


def test_find_duplicate_indices():
    rows = [
        {"전표번호": "A1", "회사코드": "1000"},
        {"전표번호": "A1", "회사코드": "1000"},
        {"전표번호": "A2", "회사코드": "1000"},
    ]
    dupes = find_duplicate_indices(rows, ["전표번호", "회사코드"])
    assert dupes == {0, 1}


def test_validate_totals_mismatch():
    rows = [{"금액": "100"}, {"금액": "200"}]
    check = validate_totals(rows, amount_column="금액", expected_count=3, expected_amount=300)

    assert check.actual_count == 2
    assert check.actual_amount == 300.0
    assert check.count_mismatch is True
    assert check.amount_mismatch is False


def test_classify_rows_normal_vs_exception():
    rows = [
        {"전표번호": "A1", "담당자명": "홍길동"},  # 중복 (A1이 두 번)
        {"전표번호": "A1", "담당자명": "홍길동"},
        {"전표번호": "A2", "담당자명": None},  # 담당자 미확인
        {"전표번호": "A3", "담당자명": "김철수"},  # 정상
    ]
    normal, exception = classify_rows(rows, key_columns=["전표번호"])

    assert len(normal) == 1
    assert normal[0]["전표번호"] == "A3"
    assert len(exception) == 3
    reasons = {r["전표번호"]: r["예외사유"] for r in exception}
    assert reasons["A1"] == "중복"
    assert reasons["A2"] == "담당자미확인"
