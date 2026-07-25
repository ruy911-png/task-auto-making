import pandas as pd
import pytest

from excel_io import edit_rules, reader, writer
from excel_io.schema import HRRecord


def test_read_conditions(tmp_path):
    path = tmp_path / "conditions.xlsx"
    pd.DataFrame(
        [
            {"회사코드": "1000", "기간": "202401"},
            {"회사코드": "2000", "기간": "202402"},
        ]
    ).to_excel(path, index=False)

    rows = reader.read_conditions(path)

    assert rows == [
        {"회사코드": "1000", "기간": "202401"},
        {"회사코드": "2000", "기간": "202402"},
    ]


def test_read_hr_data_with_aliases(tmp_path):
    path = tmp_path / "hr.xlsx"
    pd.DataFrame(
        [{"사원번호": "1001", "성명": "홍길동", "아이디": "hong1", "부서": "회계팀"}]
    ).to_excel(path, index=False)

    records = reader.read_hr_data(path)

    assert records == [
        HRRecord(emp_id="1001", name="홍길동", user_id="hong1", department="회계팀")
    ]


def test_read_hr_data_missing_column_raises(tmp_path):
    path = tmp_path / "hr_bad.xlsx"
    pd.DataFrame([{"이름": "홍길동"}]).to_excel(path, index=False)

    with pytest.raises(ValueError):
        reader.read_hr_data(path)


def test_edit_rules_rename_drop_filter_calc():
    rows = [
        {"BUKRS": "1000", "금액": "100", "수량": "2"},
        {"BUKRS": "2000", "금액": "0", "수량": "1"},
    ]
    rules = {
        "rename_columns": {"BUKRS": "회사코드"},
        "filters": [{"column": "금액", "op": "!=", "value": "0"}],
        "calculated_columns": [
            {"name": "합계", "op": "multiply", "columns": ["금액", "수량"]}
        ],
    }

    result = edit_rules.apply_edit_rules(rows, rules)

    assert len(result) == 1
    assert result[0]["회사코드"] == "1000"
    assert result[0]["합계"] == 200.0


def test_edit_rules_concat():
    rows = [{"a": "AA", "b": "BB"}]
    rules = {
        "calculated_columns": [
            {"name": "합쳐진값", "op": "concat", "columns": ["a", "b"], "separator": "-"}
        ]
    }

    result = edit_rules.apply_edit_rules(rows, rules)

    assert result[0]["합쳐진값"] == "AA-BB"


def test_writer_creates_sheets(tmp_path):
    output = tmp_path / "result.xlsx"
    normal = [{"a": 1}]
    exception = [{"a": 2, "예외사유": "중복"}]
    captures = [{"label": "조건 1", "condition_capture": None, "result_capture": None}]

    writer.write_result_excel(output, normal, exception, captures)

    assert output.exists()
    sheets = pd.ExcelFile(output).sheet_names
    assert sheets == ["결과", "예외", "캡처"]
