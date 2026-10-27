# 7차 검토(독립 리뷰어 6차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round7.py
import os, sys, unittest, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet
from test_v24_round6 import Round6
D = datetime.date
H2 = ["대분류", "계정과목", "당기", "전기"]
REP2 = [["계정과목", "당기", "전기"], ["유동자산", None, None], ["현금", 100, 90], ["예금", 200, 180], ["소계", 300, 270],
        ["비유동자산", None, None], ["토지", 500, 500], ["소계", 500, 500], ["합계", 800, 770]]
HR = ["일자", "계정명", "금액", "적요"]
RR = [(D(2026, 1 + i // 28, 1 + i % 28), f"계정{i % 5}", 1000 * (i + 1), f"적요{i}") for i in range(100)]


class Round7(Round6):
    def test_r7_a1_second_amount_column(self):
        raw = [H2, ["유동자산", "현금", 100, 90], ["유동자산", "예금", 200, 180], ["유동자산", "신규계정", 0, 50], ["비유동자산", "토지", 500, 500]]
        r, b1, u, mm = self.agg(raw, REP2, "r7a1")
        self.assertEqual(u.get("보고서!C2"), "누락", u); self.assertNotIn("보고서!B2", u)
        raw = [["대분류", "계정과목", "전기", "당기"], ["유동자산", "현금", 90, 100], ["유동자산", "예금", 180, 200], ["유동자산", "신규계정", 0, 50], ["비유동자산", "토지", 500, 500]]
        r, b1, u, mm = self.agg(raw, REP2, "r7a1e")
        self.assertEqual(u.get("보고서!B2"), "누락", u); self.assertNotIn("보고서!C2", u)

    def test_r7_s1_many_small_accounts_omitted(self):
        small = [["유동자산", f"잡계정{k}", 400] for k in range(10)]
        raw = [["대분류", "계정과목", "금액"], ["유동자산", "현금", 100000], ["유동자산", "예금", 200000]] + small + [["비유동자산", "토지", 500000], ["비유동자산", "건물", 300000]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["소계", 300],
                                      ["비유동자산", None], ["토지", 500], ["건물", 300], ["소계", 800], ["합계", 1100]], "r7s1")
        self.assertEqual(u.get("보고서!B2"), "누락", u)

    def test_r7_n4_uncached_formula_on_zero_raw(self):
        raw = [["대분류", "계정과목", "금액"], ["유동자산", "현금", 100], ["유동자산", "예금", 0], ["유동자산", "상품", 50], ["비유동자산", "토지", 500]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", "=999"], ["상품", 50], ["소계", 150], ["비유동자산", None], ["토지", 500], ["소계", 500], ["합계", 650]], "r7n4")
        self.assertEqual(u.get("보고서!B4"), "계산값 없음", u)

    def test_r7_m1_raw_has_subtotal_rows(self):
        raw = [["대분류", "계정과목", "금액"], ["유동자산", "현금", 100], ["유동자산", "예금", 200], ["유동자산", "소계", 300], ["비유동자산", "토지", 500], ["비유동자산", "소계", 500]]
        self.ok(raw, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["소계", 300], ["비유동자산", None], ["토지", 500], ["소계", 500], ["합계", 800]], "r7m1")

    def test_r7_v_valued_heading_with_offsetting_accounts_omitted(self):
        raw = [["대분류", "계정과목", "금액"], ["분류0", "계정00", 80], ["분류0", "계정01", 80], ["분류0", "상계A", 100], ["분류0", "상계B", -100], ["분류1", "계정10", 90]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["분류0", 160], ["계정00", 80], ["계정01", 80], ["분류1", 90], ["계정10", 90]], "r7v")
        self.assertEqual(u.get("보고서!B2"), "확인 필요", u)
        self.ok(raw[:3] + raw[5:], [["계정과목", "금액"], ["분류0", 160], ["계정00", 80], ["계정01", 80], ["분류1", 90], ["계정10", 90]], "r7v2")

    def rows_big(self, sheets):
        self.raw = wx(os.path.join(self.d, "rawrr.xlsx"), [("원본", [HR] + RR)])
        return self.rows_of([(n, [HR] + [RR[i] for i in idx]) for n, idx in sheets])

    def test_r7_e1_extract_not_absorbed(self):
        rep, res, c = self.rows_big([("전체", range(100)), ("1월", range(28)), ("2월", [i for i in range(28, 56) if i != 40]), ("3_4월", range(56, 100)), ("발췌", [5, 40, 60, 90])])
        self.assertEqual([x["원본 행"] for x in res if x["판정"] == "보고서에 없음"], [42], c)
        self.assertFalse([x for x in res if "중복" in x["이유"]], c)

    def test_r7_l1_dup_location_by_neighbours(self):
        rep, res, c = self.rows_big([("1월", range(28)), ("2_4월", list(range(28, 100)) + [3])])
        d = [x for x in res if "중복" in x["이유"]]
        self.assertEqual([(x["시트"], x["원본 행"]) for x in d], [("2_4월", 5)], c)


if __name__ == "__main__":
    unittest.main(defaultTest="Round7")
