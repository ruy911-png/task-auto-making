# 9차 검토(독립 리뷰어 8차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round9.py
import os, sys, unittest, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet
from test_v24_round8 import Round8, RAWG
D = datetime.date
BASE = [["유동자산", "현금", 100], ["유동자산", "예금", 200], ["비유동자산", "토지", 500]]
RAWM = [["일자", "대분류", "계정과목", "금액"]] + [[D(2026, k, 5), a, b, v * k] for k in (1, 2) for a, b, v in BASE]
HDR = [["계정과목", "1월", "2월", "합계"]]
BODY = [["유동자산", 300, 600, 900], ["현금", 100, 200, 300], ["예금", 200, 400, 600], ["비유동자산", 500, 1000, 1500], ["토지", 500, 1000, 1500]]


class Round9(Round8):
    def test_r9_i_corner_cell_top_subtotal(self):
        self.ok(RAWM, HDR + BODY + [["합계", 800, 1600, 2400]], "r9i1")
        r, b1, u, mm = self.agg(RAWM, HDR + BODY + [["합계", 800, 1600, 4800]], "r9i2")
        self.assertIn("보고서!D7", mm, (mm, u))
        self.ok(RAWM, HDR + BODY[:3] + [["소계", 300, 600, 900]] + BODY[3:] + [["소계", 500, 1000, 1500], ["합계", 800, 1600, 2400]], "r9i3")

    def test_r9_g_summary_then_detail_block(self):
        self.ok(RAWG, [["계정과목", "금액"], ["유동자산", 300], ["비유동자산", 500], ["합계", 800], ["비유동자산 명세", None], ["토지", 500], ["합계", 500]], "r9g3")
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["유동자산", 300], ["비유동자산", 500], ["합계", 300], ["비유동자산 명세", None], ["토지", 500], ["합계", 500]], "r9g4")
        self.assertIn("보고서!B4", mm, (mm, u))
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["비유동자산", 500], ["유동자산", 300], ["합계", 500], ["현금", 100], ["예금", 200], ["토지", 500], ["합계", 800]], "r9g1")
        self.assertIn("보고서!B4", mm, (mm, u))

    def test_r9_b_same_account_in_two_sections(self):
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["주요 자산", None], ["현금", 100], ["토지", 500], ["소계", 600], ["전체 자산", None], ["현금", 100], ["예금", 200], ["토지", 500], ["소계", 800]], "r9b4")
        self.assertEqual(mm, set(), (mm, u))
        r, b1, u, mm = self.agg(RAWG, [["구간", "계정과목", "금액"], ["Ⅰ.주요", "현금", 100], ["Ⅰ.주요", "토지", 500], ["소계", None, 600], ["Ⅱ.전체", "현금", 100], ["Ⅱ.전체", "예금", 200], ["Ⅱ.전체", "토지", 500], ["소계", None, 800]], "r9b5")
        self.assertEqual(mm, set(), (mm, u))

    def test_r9_b8_same_account_column_twice(self):
        rawm = [["일자", "계정과목", "금액"]] + [[D(2026, k, 5), a, v * k] for k in (1, 2) for _, a, v in BASE]
        r, b1, u, mm = self.agg(rawm, [["월", "현금", "예금", "현금", "토지", "합계"], ["1월", 100, 200, 100, 500, 900], ["2월", 200, 400, 200, 1000, 1800]], "r9b8")
        self.assertNotEqual(b1, "전부 일치", (u, mm))
        self.ok([["계정과목", "당기", "전기"], ["현금", 100, 100], ["예금", 200, 200]], [["계정과목", "당기", "전기"], ["현금", 100, 100], ["예금", 200, 200], ["합계", 300, 300]], "r9b7")

    def test_r9_e_dup_location_attribute_and_range(self):
        H = ["일자", "부서", "계정명", "금액", "적요"]
        R = [(D(2026, 1, 1) + datetime.timedelta(days=i // 2), "영업부" if i % 2 == 0 else "관리부", f"계정{i % 5}", 1000 * (i + 1), f"적요{i}") for i in range(100)]
        self.raw = wx(os.path.join(self.d, "rawe.xlsx"), [("원본", [H] + R)])
        ev = list(range(0, 100, 2)); od = list(range(1, 100, 2))
        rep, res, c = self.rows_of([("영업부", [H] + [R[i] for i in ev]), ("관리부", [H] + [R[i] for i in od[:10] + [20] + od[10:]])])
        self.assertEqual([x["시트"] for x in res if "중복" in x["이유"]], ["관리부"], c)
        R2 = [(D(2026, 1 + i // 25, 1 + i % 25), f"계정{i % 5}", 1000 * (i + 1), f"적요{i}") for i in range(100)]
        H2 = ["일자", "계정명", "금액", "적요"]
        self.raw = wx(os.path.join(self.d, "rawe4.xlsx"), [("원본", [H2] + R2)])
        rep, res, c = self.rows_of([("1월", [H2] + R2[:25]), ("2_4월", [H2] + [R2[i] for i in [23, 24] + list(range(25, 100))])])
        self.assertEqual([x["시트"] for x in res if "중복" in x["이유"]], ["2_4월", "2_4월"], c)


if __name__ == "__main__":
    unittest.main(defaultTest="Round9")
