"""여러 조건의 결과를 하나의 결과 엑셀로 합치고, 캡처 이미지를 삽입한다.

시트 구성 (5개 탭, 아래 순서 고정):
- "캡처": 조건별 조회조건 화면 + 조회결과 화면 이미지 (텍스트 로그 대신 감사 증적)
- "Raw data": SAP에서 다운로드한 원본 데이터를 편집규칙 적용 전 그대로 병합한 것
- "모집단 data": 편집규칙 적용 + 담당자 매칭까지 끝난 데이터 중 정상 건
  (validation.checks.classify_rows가 반환하는 normal_rows와 동일한 개념)
- "샘플 data": "모집단 data"에서 표본추출 기준(validation.sampling)에 따라 무작위로 뽑은 표본
- "예외 data": classify_rows가 반환하는 exception_rows
  (중복/금액불일치/담당자미확인/조회실패 등, 사람이 확인해야 하는 행)
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


def write_result_excel(
    output_path: str | Path,
    raw_rows: list[dict[str, Any]],
    population_rows: list[dict[str, Any]],
    sample_rows: list[dict[str, Any]],
    exception_rows: list[dict[str, Any]],
    captures: list[dict[str, Any]],
    sampling_note: str | None = None,
) -> Path:
    """captures: [{"label": str, "condition_capture": path|None, "result_capture": path|None}, ...]

    sampling_note: 표본 수 산정근거 메모 (예: "월별 통제 최대치(4건) 적용"). 주어지면
    "샘플 data" 시트 상단에 기록한다.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame([{"조건": c.get("label", "")} for c in captures]).to_excel(
            writer, sheet_name="캡처", index=False
        )
        pd.DataFrame(raw_rows).to_excel(writer, sheet_name="Raw data", index=False)
        pd.DataFrame(population_rows).to_excel(writer, sheet_name="모집단 data", index=False)

        sample_start_row = 0
        if sampling_note:
            sample_start_row = 1
        pd.DataFrame(sample_rows).to_excel(
            writer, sheet_name="샘플 data", index=False, startrow=sample_start_row
        )
        pd.DataFrame(exception_rows).to_excel(writer, sheet_name="예외 data", index=False)

    if sampling_note:
        _write_sampling_note(output_path, sampling_note)

    _insert_capture_images(output_path, captures)
    return output_path


def _write_sampling_note(output_path: Path, sampling_note: str) -> None:
    import openpyxl

    workbook = openpyxl.load_workbook(output_path)
    sheet = workbook["샘플 data"]
    sheet["A1"] = f"표본 수 산정근거: {sampling_note}"
    workbook.save(output_path)


def _insert_capture_images(output_path: Path, captures: list[dict[str, Any]]) -> None:
    import openpyxl
    from openpyxl.drawing.image import Image as XLImage

    workbook = openpyxl.load_workbook(output_path)
    sheet = workbook["캡처"]

    sheet["C1"] = "조회조건 화면"
    sheet["D1"] = "조회결과 화면"

    row = 2
    for capture in captures:
        condition_path = capture.get("condition_capture")
        result_path = capture.get("result_capture")
        if condition_path and Path(condition_path).exists():
            img = XLImage(condition_path)
            img.width, img.height = 240, 160
            sheet.add_image(img, f"C{row}")
        if result_path and Path(result_path).exists():
            img = XLImage(result_path)
            img.width, img.height = 240, 160
            sheet.add_image(img, f"D{row}")
        sheet.row_dimensions[row].height = 130
        row += 1

    sheet.column_dimensions["C"].width = 35
    sheet.column_dimensions["D"].width = 35

    workbook.save(output_path)
