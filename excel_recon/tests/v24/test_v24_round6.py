# 6차 검토(독립 리뷰어 5차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round6.py
import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet, H
from test_v24_round5 import Round5
from test_v24_round4 import RAW
REP = [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50], ["소계", 350], ["비유동자산", None], ["토지", 500], ["소계", 500], ["합계", 850]]
def R(cur=(), non=()):
    return RAW[:4] + [list(x) for x in cur] + RAW[4:] + [list(x) for x in non]


class Round6(Round5):
    def test_r6_a1_offsetting_accounts_omitted(self):
        for nm, extra in (("r6a1", [["유동자산", "당좌자산", "가지급금", "A", 3000], ["유동자산", "당좌자산", "가수금", "A", -3000]]),
                          ("r6a2", [["유동자산", "당좌자산", "매출채권", "A", 1000], ["유동자산", "당좌자산", "대손충당금", "A", -1000]])):
            r, b1, u, mm = self.agg(R(extra), REP, nm)
            self.assertEqual(u.get("보고서!B2"), "누락", (nm, u))
        self.ok(R([["유동자산", "당좌자산", "가지급금", "A", 30], ["유동자산", "당좌자산", "가지급금", "A", -30]]), REP, "r6a3")

    def test_r6_z_zero_net_heading_with_missing_account(self):
        raw = [["대분류", "계정과목", "금액"], ["분류0", "계정00", 70], ["분류0", "계정01", 50], ["분류1", "계정10", -30], ["분류1", "계정11", 30]]
        r, b1, u, mm = self.agg(raw, [["계정과목", "금액"], ["분류0", None], ["계정00", 70], ["계정01", 50], ["소계", 120], ["분류1", None], ["계정11", 30], ["소계", 30]], "r6z")
        self.assertEqual(u.get("보고서!B6"), "누락", u)
        self.ok(raw, [["계정과목", "금액"], ["분류0", None], ["계정00", 70], ["계정01", 50], ["소계", 120], ["분류1", None], ["계정10", -30], ["계정11", 30], ["소계", 0]], "r6z2")

    def test_r6_b_summary_blocks_ok(self):
        self.ok(RAW, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50], ["비유동자산", None], ["토지", 500], ["합계", 850],
                      [None, None], ["계정과목", "금액"], ["유동자산", 350], ["비유동자산", 500], ["합계", 850]], "r6b1")
        self.ok(RAW, [["구분", "유동자산", "현금", "예금", "상품", "비유동자산", "토지", None, "유동자산", "비유동자산"], ["A", None, 100, 200, 50, None, 500, None, 350, 500]], "r6b2")
        raw = [["대분류", "계정과목", "금액"], ["유동자산", "현금", 100], ["유동자산", "예금", 200], ["비유동자산", "토지", 500],
               ["매출", "제품매출", 700], ["매출", "상품매출", 300], ["비용", "급여", 400], ["비용", "임차료", 200]]
        self.ok(raw, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["비유동자산", None], ["토지", 500],
                      [None, None], ["계정과목", "금액"], ["매출", 1000], ["비용", 600]], "r6b4")

    def test_r6_c1_uncached_detail_formula(self):
        r, b1, u, mm = self.agg(RAW, [["계정과목", "금액"], ["유동자산", None], ["현금", "=50+50"], ["예금", 200], ["상품", 50], ["소계", 350], ["비유동자산", None], ["토지", 500], ["합계", 850]], "r6c1")
        self.assertEqual(u.get("보고서!B3"), "계산값 없음", u); self.assertEqual(mm, set(), mm)

    def test_r6_d1_dup_in_small_split_sheet(self):
        rep, res, c = self.groups(100, [("전체", range(100)), ("1월", list(range(10)) + [59]), ("2월", range(10, 100))])
        d = [x for x in res if "중복" in x["이유"]]
        self.assertEqual([(x["시트"], x["원본 행"]) for x in d], [("1월", 61)], c)
        self.assertEqual(c["보고서에 없음"], 0, c)
        rep, res, c = self.groups(100, [("전체", range(100)), ("1월", [i for i in range(10) if i != 3] + [59]), ("2월", range(10, 100))])
        self.assertEqual([x["원본 행"] for x in res if x["판정"] == "보고서에 없음"], [5], c)

    def test_r6_g1_two_filter_views_not_dups(self):
        rep, res, c = self.groups(100, [("전체", range(100)), ("계정0_1", [i for i in range(100) if i % 5 < 2]), ("1월", range(28))])
        self.assertFalse([x for x in res if "중복" in x["이유"]], c)
        self.assertEqual(c["보고서에 없음"], 0, c)


if __name__ == "__main__":
    unittest.main(defaultTest="Round6")
