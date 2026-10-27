# 10차 검토(독립 리뷰어 9차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round10.py
import os, sys, unittest, datetime, random
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet
from test_v24_round9 import Round9, RAWM as RAWM9
from test_v24_round8 import RAWG
D = datetime.date


class Round10(Round9):
    def test_r10_cross_section_duplicate_in_grand_total(self):
        for nm, rep in (("a1", [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["소계", 300], ["비유동자산", None], ["현금", 100], ["토지", 500], ["소계", 600], ["합계", 900]]),
                        ("b4", [["계정과목", "금액"], ["현금", 100], ["예금", 200], ["소계", 300], ["현금", 100], ["토지", 500], ["소계", 600], ["합계", 900]])):
            r, b1, u, mm = self.agg(RAWG, rep, "r10" + nm)
            self.assertNotEqual(b1, "전부 일치", (nm, u, mm))
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["주요 자산", None], ["현금", 100], ["토지", 500], ["소계", 600], ["현금", 100], ["토지", 500], ["합계", 600]], "r10a3")
        self.assertNotEqual(b1, "전부 일치")      # 같은 표 다른 구간에 같은 계정 → 확인 필요 (주요/전체 구간인지 사람이 판단)

    def test_r10_middle_heading_rows(self):
        rep = lambda tot: [["계정과목", "금액"], ["유동자산", 300], ["(1)당좌자산", None], ["현금", 100], ["예금", 200], ["비유동자산", 500], ["(1)유형자산", None], ["토지", 500], ["자산총계", tot]]
        self.ok(RAWG, rep(800), "r10g1")
        r, b1, u, mm = self.agg(RAWG, rep(1600), "r10g4")
        self.assertIn("보고서!B9", mm, (mm, u))
        rawm = [["일자", "대분류", "계정과목", "금액"]] + [[D(2026, k, 5), a, b, v * k] for k in (1, 2) for a, b, v in (("유동자산", "현금", 100), ("유동자산", "예금", 200), ("비유동자산", "토지", 500))]
        r, b1, u, mm = self.agg(rawm, [["계정과목", "1월", "2월", "합계"], ["유동자산", 300, 600, 900], ["(1)당좌자산", None, None, None], ["현금", 100, 200, 300], ["예금", 200, 400, 600],
                                       ["비유동자산", 500, 1000, 1500], ["(1)유형자산", None, None, None], ["토지", 500, 1000, 1500], ["합계", 1600, 3200, 4800]], "r10g5")
        self.assertNotEqual(b1, "전부 일치", (mm, u))

    def test_r10_column_sections_repeat_accounts(self):
        rawm = [["일자", "계정과목", "금액"]] + [[D(2026, k, 5), a, v * k] for k in (1, 2) for a, v in (("현금", 100), ("예금", 200), ("토지", 500))]
        r, b1, u, mm = self.agg(rawm, [["월", "현금", "토지", "소계", "현금", "예금", "토지", "소계"], ["1월", 100, 500, 600, 100, 200, 500, 800], ["2월", 200, 1000, 1200, 200, 400, 1000, 1600]], "r10b5")
        self.assertEqual(mm, set(), (mm, u))                       # 구간 소계는 맞음 — 다른 구간에 같은 계정은 확인 필요만 (세로 A3과 같은 처리)
        self.assertTrue(all(v == "확인 필요" for v in u.values()), u)

    def test_r10_constant_column_does_not_hide_department(self):
        H = ["일자", "부서", "계정명", "금액", "적요", "통화"]
        R = [(D(2026, 1, 1) + datetime.timedelta(days=i // 2), "영업부" if i % 2 == 0 else "관리부", f"계정{i % 5}", 1000 * (i + 1), f"적요{i}", "KRW") for i in range(100)]
        self.raw = wx(os.path.join(self.d, "rawc.xlsx"), [("원본", [H] + R)])
        ev = list(range(0, 100, 2)); od = list(range(1, 100, 2))
        rep, res, c = self.rows_of([("영업부", [H] + [R[i] for i in ev]), ("관리부", [H] + [R[i] for i in od[:10] + [20] + od[10:]])])
        self.assertEqual([x["시트"] for x in res if "중복" in x["이유"]], ["관리부"], c)

    def test_r10_page_split_copy_location(self):
        H = ["일자", "부서", "계정명", "금액", "적요"]
        rnd = random.Random(7)
        R = [(D(2026, 1, 1) + datetime.timedelta(days=i * 90 // 120), rnd.choice(["영업부", "관리부", "생산부"]), f"계정{rnd.randint(0, 4)}", 10 * rnd.randint(1, 10 ** 5), f"적요{i}") for i in range(120)]
        self.raw = wx(os.path.join(self.d, "rawp.xlsx"), [("원본", [H] + R)])
        wrong = []
        for k in (42, 51, 57, 3, 30):
            for where in ("앞", "끝"):
                p2 = ([k] + list(range(60, 120))) if where == "앞" else (list(range(60, 120)) + [k])
                rep, res, c = self.rows_of([("p1", [H] + R[:60]), ("p2", [H] + [R[i] for i in p2])])
                if [x["시트"] for x in res if "중복" in x["이유"]] != ["p2"]:
                    wrong.append((k, where))
        self.assertEqual(wrong, [])


if __name__ == "__main__":
    unittest.main(defaultTest="Round10")
