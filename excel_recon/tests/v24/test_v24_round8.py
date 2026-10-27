# 8차 검토(독립 리뷰어 7차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round8.py
import os, sys, unittest, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet
from test_v24_round7 import Round7, RR
D = datetime.date
RAWG = [["대분류", "계정과목", "금액"], ["유동자산", "현금", 100], ["유동자산", "예금", 200], ["비유동자산", "토지", 500]]
HM = ["일자", "대분류", "계정과목", "금액"]
ROWSM = [[D(2026, k, 5), "유동자산", "현금", 100 * k] for k in (1, 2)] + [[D(2026, k, 5), "유동자산", "예금", 200 * k] for k in (1, 2)] \
        + [[D(2026, 3, 31), "유동자산", "미수수익", 50]] + [[D(2026, k, 5), "비유동자산", "토지", 500] for k in (1, 2, 3)]


class Round8(Round7):
    def test_r8_g_duplicate_account_rows(self):
        for nm, rep in (("g2", [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["현금", 100], ["예금", 200], ["소계", 400], ["비유동자산", None], ["토지", 500], ["소계", 500]]),
                        ("g4", [["계정과목", "금액"], ["현금", 100], ["현금", 100], ["예금", 200], ["합계", 400]]),
                        ("g1", [["계정과목", "금액"], ["유동자산", 300], ["현금", 100], ["현금", 100], ["예금", 200], ["비유동자산", 500], ["토지", 500]])):
            r, b1, u, mm = self.agg(RAWG, rep, nm)
            self.assertNotEqual(b1, "전부 일치", (nm, u, mm))
            self.assertEqual(u.get("보고서!B4" if nm != "g4" else "보고서!B3"), "확인 필요", (nm, u))
            if nm != "g1":
                self.assertIn("보고서!B6" if nm == "g2" else "보고서!B5", mm, (nm, mm))   # 중복분까지 더한 소계·합계는 불일치

    def test_r8_c1_child_without_rows_in_that_column(self):
        r, b1, u, mm = self.agg([HM] + ROWSM, [["계정과목", "1월", "2월", "3월"], ["유동자산", 300, 600, 50], ["현금", 100, 200, None], ["예금", 200, 400, None],
                                       ["비유동자산", 500, 500, 500], ["토지", 500, 500, 500]], "r8c1")
        self.assertEqual(u.get("보고서!D2"), "확인 필요", u)
        self.ok([HM] + ROWSM, [["계정과목", "1월", "2월", "3월"], ["유동자산", 300, 600, 50], ["현금", 100, 200, None], ["예금", 200, 400, None], ["미수수익", None, None, 50],
                        ["비유동자산", 500, 500, 500], ["토지", 500, 500, 500]], "r8c2")

    def test_r8_a14_column_summary_with_some_departments_ok(self):
        raw = [["부서", "계정과목", "금액"]] + [[dp, ac, v] for dp in ("영업부", "관리부", "생산부") for ac, v in (("급여", 100), ("복리", 30))]
        self.ok(raw, [["계정과목", "금액", "영업부", "관리부"], ["급여", 300, 100, 100], ["복리", 90, 30, 30]], "r8a14")

    def test_r8_top_subtotal_form_with_grand_total_ok(self):
        self.ok(RAWG, [["계정과목", "금액"], ["유동자산", 300], ["현금", 100], ["예금", 200], ["비유동자산", 500], ["토지", 500], ["합계", 800]], "r8t1")
        rm = [["일자", "대분류", "계정과목", "금액"]] + [[D(2026, k, 5), a, b, v * k] for k in (1, 2) for a, b, v in (("유동자산", "현금", 100), ("유동자산", "예금", 200), ("비유동자산", "토지", 500))]
        self.ok(rm, [["월", "유동자산", "현금", "예금", "비유동자산", "토지", "합계"], ["1월", 300, 100, 200, 500, 500, 800], ["2월", 600, 200, 400, 1000, 1000, 1600]], "r8t2")
        r, b1, u, mm = self.agg(rm, [["월", "유동자산", "현금", "비유동자산", "토지"], ["1월", 300, 100, 500, 500], ["2월", 600, 200, 1000, 1000]], "r8t3")
        self.assertEqual((u.get("보고서!B2"), u.get("보고서!B3")), ("확인 필요", "확인 필요"), u)

    def test_r8_c5_uncached_subtotal_on_zero_raw(self):
        raw = [["계정과목", "금액"], ["현금", 600], ["건물", 0]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["현금", 600], ["소계", 600], ["건물", 0], ["소계", "=B4"]], "r8c5")
        self.assertEqual(u.get("보고서!B5"), "계산값 없음", u)

    def test_r8_b1_dup_location_boundary(self):
        rep, res, c = self.rows_big([("1월", range(28)), ("2_4월", list(range(28, 100)) + [27])])
        self.assertEqual([x["시트"] for x in res if "중복" in x["이유"]], ["2_4월"], c)
        rep, res, c = self.rows_big([("1월", list(range(28)) + [28, 29]), ("2_4월", range(28, 100))])
        self.assertEqual(sorted(x["시트"] for x in res if "중복" in x["이유"]), ["1월", "1월"], c)

    def test_r8_h_summary_counts_review_cells(self):
        raw = [["대분류", "계정과목", "금액"], ["분류0", "계정00", 80], ["분류0", "계정01", 80], ["분류0", "상계A", 100], ["분류0", "상계B", -100], ["분류1", "계정10", 90]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["분류0", 160], ["계정00", 80], ["계정01", 80], ["분류1", 90], ["계정10", 90]], "r8h")
        line = "\n".join(m.summary(r))
        self.assertIn("확인 필요 1", str(line))


if __name__ == "__main__":
    unittest.main(defaultTest="Round8")
