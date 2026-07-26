"""여러 조건의 결과를 하나의 결과 엑셀로 합치고, 캡처 이미지를 삽입한다.

시트 구성:
- "결과": 정상 처리된 데이터 행 (사람이 확인할 필요 없음)
- "예외": 중복/금액불일치/담당자미확인/조회실패 등 사람이 확인해야 하는 행
- "캡처": 조건별 조회조건 화면 + 조회결과 화면 이미지 (텍스트 로그 대신 감사 증적)
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


def write_result_excel(
    output_path: str | Path,
    normal_rows: list[dict[str, Any]],
    exception_rows: list[dict[str, Any]],
    captures: list[dict[str, Any]],
) -> Path:
    """captures: [{"label": str, "condition_capture": path|None, "result_capture": path|None}, ...]"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(normal_rows).to_excel(writer, sheet_name="결과", index=False)
        pd.DataFrame(exception_rows).to_excel(writer, sheet_name="예외", index=False)
        pd.DataFrame([{"조건": c.get("label", "")} for c in captures]).to_excel(
            writer, sheet_name="캡처", index=False
        )

    _insert_capture_images(output_path, captures)
    return output_path


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
