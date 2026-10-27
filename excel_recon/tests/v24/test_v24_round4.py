# 4차 검토(독립 리뷰어 3차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round4.py
import os, sys, zipfile, shutil, subprocess, unittest, collections, datetime
from openpyxl import load_workbook
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet, H, TOOL_DIR
from test_v24_round3 import Round3
D = datetime.date
RAW = [["대분류", "중분류", "계정과목", "구분", "금액"],
       ["유동자산", "당좌자산", "현금", "A", 100], ["유동자산", "당좌자산", "예금", "A", 200],
       ["유동자산", "재고자산", "상품", "A", 50], ["비유동자산", "유형자산", "토지", "A", 500]]
RAWP = [["대분류", "계정과목", "기간", "금액"],
        ["유동자산", "현금", "당기", 100], ["유동자산", "예금", "당기", 200], ["유동자산", "예금", "전기", 150], ["유동자산", "상품", "당기", 50], ["유동자산", "상품", "전기", 40],
        ["비유동자산", "토지", "당기", 500], ["비유동자산", "토지", "전기", 500]]
RAW2 = [["대분류", "계정과목", "구분", "금액"], ["유동자산", "현금", "A", 100], ["유동자산", "현금", "B", 50], ["유동자산", "예금", "A", 200], ["유동자산", "예금", "B", 70]]
BIG = [(D(2026, 1 + i % 12, 1 + i % 27), f"계정{i % 5}", 1000 * (i + 1), f"적요{i}") for i in range(100)]


class Round4(Round3):
    def ok(self, raw, rep, nm):
        r, b1, u, mm = self.agg(raw, rep, nm)
        self.assertEqual(b1, "전부 일치", (nm, u, mm))

    # --- 구분 제목
    def test_h1_multilevel_headings(self):
        self.ok(RAW, [["계정과목", "금액"], ["유동자산", None], ["당좌자산", None], ["현금", 100], ["예금", 200], ["소계", 300],
                      ["재고자산", None], ["상품", 50], ["소계", 50], ["비유동자산", None], ["유형자산", None], ["토지", 500], ["소계", 500], ["합계", 850]], "h1")

    def test_h2_h3_headings_without_group_subtotal(self):
        self.ok(RAW, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50], ["비유동자산", None], ["토지", 500], ["합계", 850]], "h2")
        self.ok(RAW, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50], ["비유동자산", None], ["토지", 500]], "h3")

    def test_h5_column_direction_heading(self):
        raw = [r for r in RAW if r[0] != "비유동자산"]
        self.ok(raw, [["구분", "유동자산", "현금", "예금", "상품", "합계"], ["A", None, 100, 200, 50, 350]], "h5")

    def test_h6r_top_subtotal_form_blank_group_is_missing(self):
        r, b1, u, mm = self.agg(RAWP, [["계정과목", "당기"], ["비유동자산", 500], ["토지", 500], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50]], "h6r")
        self.assertEqual(u.get("보고서!B4"), "누락", u)

    def test_h8_heading_then_missing_detail(self):
        r, b1, u, mm = self.agg(RAW, [["계정과목", "금액"], ["유동자산", None], ["현금", None], ["예금", 200], ["상품", 50], ["소계", 250]], "h8")
        self.assertEqual(u.get("보고서!B3"), "누락", u); self.assertNotIn("보고서!B2", u)
        self.assertEqual([x["원본"] for x in r["mismatch"] if x["위치"] == "보고서!B6"], [350])

    def test_h10_heading_then_zero_blank(self):
        self.ok(RAWP, [["계정과목", "당기", "전기"], ["유동자산", None, None], ["현금", 100, None], ["예금", 200, 150], ["상품", 50, 40], ["소계", 350, 190],
                       ["비유동자산", None, None], ["토지", 500, 500], ["소계", 500, 500], ["합계", 850, 690]], "h10")

    def test_t1_heading_row_with_total_column_value(self):
        r, b1, u, mm = self.agg(RAW2, [["계정과목", "A", "B", "합계"], ["유동자산", None, None, 420], ["현금", 100, 50, 150], ["예금", 200, 70, 270], ["소계", 300, 120, 420]], "t1")
        self.assertEqual(mm, set(), (mm, u))

    def test_a11_sibling_blanks_are_plain_missing(self):
        r, b1, u, mm = self.agg(RAW, [["계정과목", "금액"], ["현금", None], ["예금", None], ["상품", None]], "a11")
        why = {x["위치"]: x["이유"] for x in r["unmatched"]}
        self.assertEqual(u.get("보고서!B2"), "누락", u)
        self.assertNotIn("구분 제목", why["보고서!B2"])

    # --- 빈 양식/메모 시트
    def test_t2_t3_memo_or_adjustment_sheet_blocks(self):
        for nm, extra in (("t2", [["검토 메모"], ["검토의견", "3월 거래 2건 누락 확인 필요"]]), ("t3", [["수기 조정분"], [D(2026, 3, 1), "계정1", 50000, "추가 반영"]])):
            rep = wx(os.path.join(self.d, f"p{nm}.xlsx"), [("본표", [H] + list(self.ROWS20)), ("메모", extra)])
            self.assertNotEqual(self.final(rep), "전부 일치", nm)

    @property
    def ROWS20(self):
        from test_v24_review import ROWS
        return ROWS

    # --- 수식 결과가 빈 글자
    def test_n3_formula_returning_empty_string_is_missing(self):
        raw = wx(os.path.join(self.d, "rn3.xlsx"), [("원본", [["계정과목", "금액"], ["현금", 100], ["예금", 200], ["적금", 300]])])
        rep = wx(os.path.join(self.d, "pn3.xlsx"), [("보고서", [["계정과목", "금액"], ["현금", 100], ["예금", '=IF(1=1,"",200)'], ["적금", 300], ["합계", 600]])])
        tmp = rep + ".tmp"
        with zipfile.ZipFile(rep) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for it in zi.infolist():
                data = zi.read(it.filename)
                if it.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b'<f>IF(1=1,"",200)</f><v></v>', b'<f>IF(1=1,"",200)</f><v></v>').replace(b'<c r="B3" t="str"><f>', b'<c r="B3" t="str"><f>')
                    s = data.decode()
                    i = s.index('r="B3"'); j = s.index("</c>", i)
                    s = s[:s.rindex("<c", 0, i)] + '<c r="B3" t="str"><f>IF(1=1,"",200)</f><v></v></c>' + s[j + 4:]
                    data = s.encode()
                zo.writestr(it, data)
        shutil.move(tmp, rep)
        u = {x["위치"]: x["구분"] for x in quiet(m.verify, raw, rep)["unmatched"]}
        self.assertEqual(u.get("보고서!B3"), "누락", u)

    # --- 행 대조 묶음
    def groups(self, n, sheets):
        self.raw = wx(os.path.join(self.d, f"rawg{n}.xlsx"), [("원본", [H] + BIG[:n])])
        return self.rows_of([(nm, [H] + [BIG[i] for i in idx]) for nm, idx in sheets])

    def test_g4_copy_missing_15(self):
        rep, res, c = self.groups(100, [("원화", range(100)), ("달러", range(85))])
        self.assertEqual(c["보고서에 없음"], 15, c)

    def test_g8_full_sheet_missing_with_monthly_split(self):
        rep, res, c = self.groups(100, [("전체", range(10, 100)), ("1월", range(30)), ("2월", range(30, 60)), ("3월", range(60, 100))])
        miss = [x for x in res if x["판정"] == "보고서에 없음"]
        self.assertEqual((len(miss), {x["시트"] for x in miss}), (10, {"전체"}), c)
        self.assertFalse([x for x in res if "중복" in x["이유"]], c)

    def test_g7_dup_flag_on_later_sheet(self):
        rep, res, c = self.groups(100, [("전체", range(100)), ("1월", range(25)), ("2월", range(25, 50)), ("3월", [30] + list(range(50, 75))), ("4월", range(75, 100))])
        d = [x for x in res if "중복" in x["이유"]]
        self.assertEqual([x["시트"] for x in d], ["3월"], d)

    def test_g9_not_in_any_sheet_label(self):
        rep, res, c = self.groups(100, [("전체", range(60)), ("요약", range(60, 99))])
        miss = [x for x in res if x["판정"] == "보고서에 없음"]
        self.assertEqual(len(miss), 1); self.assertNotEqual(miss[0]["시트"], "전체")

    def test_g1_copied_month_not_silent(self):
        rep, res, c = self.groups(100, [("1월", range(30)), ("2월", list(range(30)) + list(range(30, 60))), ("3월", range(60, 100))])
        self.assertNotEqual(self.final(rep), "전부 일치", c)

    def test_d1_small_dup_sheet_not_silent(self):
        rep, res, c = self.groups(40, [("본표", range(40)), ("추가분", [3, 7])])
        self.assertNotEqual(self.final(rep), "전부 일치", c)


if __name__ == "__main__":
    unittest.main(defaultTest="Round4")
