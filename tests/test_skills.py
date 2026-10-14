#!/usr/bin/env python3
"""스킬 자체 점검. 실행: python3 tests/test_skills.py

수식 계산 검증은 `pip install formulas` 가 되어 있을 때만 돈다 (없으면 건너뜀).
"""
import datetime as dt
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
BUDGET = ROOT / "skills" / "wedding-budget"
SCHEDULE = ROOT / "skills" / "wedding-schedule"
FORBIDDEN_FUNCS = ("XLOOKUP", "XMATCH", "LET(", "FILTER(", "UNIQUE(", "SORT(", "SEQUENCE(")

try:
    import formulas
except ImportError:
    formulas = None


def build(skill, data, out):
    script = skill / "scripts" / ("build_budget.py" if skill == BUDGET else "build_schedule.py")
    src = Path(out).with_suffix(".json")
    src.write_text(json.dumps(data, ensure_ascii=False), "utf-8")
    return subprocess.run([sys.executable, str(script), "--input", str(src), "--out", str(out)],
                          capture_output=True, text=True)


def sample(skill):
    return json.loads((skill / "examples" / "sample_input.json").read_text("utf-8"))


def check_formulas(wb):
    """금지 함수가 없고, 수식이 참조하는 시트가 실제로 있는지."""
    sheets = set(wb.sheetnames)
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    assert not any(f in v.upper() for f in FORBIDDEN_FUNCS), (ws.title, c.coordinate, v)
                    for name in re.findall(r"([가-힣A-Za-z]+)!", v):
                        assert name in sheets, (ws.title, c.coordinate, name)


def expect_bad(skill, cases, tmp):
    for i, case in enumerate(cases):
        r = build(skill, case, tmp / f"bad{i}.xlsx")
        assert r.returncode != 0 and "입력 오류" in r.stderr, (skill.name, i, r.stderr)


def calc(path):
    sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    vals = {}
    for k, v in sol.items():
        ref = k.split("]", 1)[1].replace("'", "").upper() if "]" in k else k
        vals[ref] = v.value[0, 0] if hasattr(v, "value") else v
    return vals


def assert_no_errors(vals, name):
    errs = {k: v for k, v in vals.items() if str(v).startswith("#")}
    assert not errs, (name, list(errs.items())[:10])


def serial(d):
    return (d - dt.date(1899, 12, 30)).days


# ---------- wedding-budget ----------

def test_budget(tmp):
    out = tmp / "b.xlsx"
    r = build(BUDGET, sample(BUDGET), out)
    assert r.returncode == 0, r.stderr
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["대시보드", "기본정보", "하객식대", "지출계획", "결제기록", "월별자금", "기준표"], wb.sheetnames
    plan = wb["지출계획"]
    for col in "FGLNOQRS":  # 대관료 행의 계산 열은 수식이어야 함
        assert str(plan[f"{col}6"].value).startswith("="), (col, plan[f"{col}6"].value)
    assert str(wb["하객식대"]["B17"].value).startswith("=")
    for ref in ("B6", "B7", "B9", "B13"):
        assert str(wb["대시보드"][ref].value).startswith("="), ref
    check_formulas(wb)

    base = sample(BUDGET)
    expect_bad(BUDGET, [dict(base, region="화성"), dict(base, wedding_date="2027/04/17"),
                        dict(base, items={"없는항목": {"budget": 1}}), dict(base, meal_price=-1),
                        dict(base, items={"대관료": {"payer": "신랑"}}),  # 결제자는 payments 로 옮겨감
                        dict(base, payments=[{"item": "없는항목", "amount": 1}]),
                        dict(base, payments=[{"item": "대관료", "amount": 1, "method": "수표"}]),
                        dict(base, extra_items=[{"name": "대관료"}]),  # 기본 항목과 이름 중복
                        dict(base, extra_items=[{"name": "꽃*"}])], tmp)
    minimal = {k: base[k] for k in ("wedding_date", "region", "meal_price", "guaranteed_guests")}
    r = build(BUDGET, minimal, tmp / "bmin.xlsx")
    assert r.returncode == 0, r.stderr

    bench = json.loads((BUDGET / "references" / "benchmarks.json").read_text("utf-8"))
    items = {it["name"]: it["values"] for it in bench["items"]}
    doc = (ROOT / "docs" / "research" / "wedding-cost-benchmarks-2026.md").read_text("utf-8")
    assert items["대관료"][1] == 630 and "| 대관료 | 350 | 630 |" in doc
    assert items["스드메 패키지"][0] == 298 and "| 스드메 패키지 | 298 | 350 |" in doc
    assert bench["meal"]["per_person"][1] == 8.0 and "| 1인당 식대 | 6.0 | 8.0 |" in doc
    assert items["신혼여행"][0] == 1030 and "1,030" in doc
    assert all(v is None or len(v) == len(bench["regions"]) for v in items.values())

    if not formulas:
        return
    for path in (out, tmp / "bmin.xlsx"):
        assert_no_errors(calc(path), path.name)
    vals = calc(out)
    # 샘플: 대인 215, 소인 7, 보증 220, 대인 68,000 / 소인 34,000
    assert vals["하객식대!B17"] == 220 * 68000 + 7 * 34000, vals["하객식대!B17"]
    assert vals["지출계획!Q6"] == 300 * 10000  # 경인 대관료 중간값
    assert vals["대시보드!B9"] == vals["대시보드!B6"] - vals["대시보드!B7"]
    assert vals["월별자금!G36"] == vals["대시보드!B9"]  # 모든 지급이 월별자금에 잡힘
    # 결제기록: 대관료 계약금 100만(공통·카드), 스드메 50만(신부·계좌이체·발급), 스냅 30만(신랑·계좌이체·미기재)
    assert vals["지출계획!N6"] == 1000000 and vals["지출계획!O6"] == 3000000 - 1000000
    assert vals["결제기록!M7"] == 500000  # 신부 × 계좌이체
    assert vals["결제기록!O6"] == 300000  # 신랑 합계
    assert vals["결제기록!L18"] == 1  # 현금영수증 확인 필요(신랑 계좌이체)
    assert vals["대시보드!B16"] == 1800000 and vals["대시보드!B18"] == 1
    assert vals["결제기록!L16"] == 500000 and vals["결제기록!O12"] == 0  # 발급완료 금액, 결제자 미입력
    plan_wb = openpyxl.load_workbook(out)["지출계획"]
    total_row = next(r for r in range(6, plan_wb.max_row + 1) if plan_wb[f"A{r}"].value == "합계")
    assert vals["대시보드!B17"] == vals[f"지출계획!O{total_row}"]
    wb = openpyxl.load_workbook(out)
    wb["기본정보"]["B7"] = "서울"
    wb.save(tmp / "seoul.xlsx")
    assert calc(tmp / "seoul.xlsx")["지출계획!Q6"] == 630 * 10000  # 지역을 바꾸면 기준값도 바뀜


# ---------- wedding-schedule ----------

def timeline_rows(wb):
    ws = wb["타임라인"]
    return {ws[f"Q{r}"].value: r for r in range(6, ws.max_row + 1) if ws[f"Q{r}"].value}


def test_schedule(tmp):
    data = json.loads((SCHEDULE / "references" / "checklists.json").read_text("utf-8"))
    ids = [t["id"] for t in data["tasks"]]
    assert len(ids) == len(set(ids)), "할 일 id 중복"
    for t in data["tasks"]:
        assert t["anchor"] in ("wedding", "shoot"), t["id"]
        assert isinstance(t["start"], int) and isinstance(t["due"], int) and t["start"] >= t["due"], t["id"]
        assert t["source"] in data["sources"] and t["phase"] and t["title"], t["id"]

    out = tmp / "s.xlsx"
    r = build(SCHEDULE, sample(SCHEDULE), out)
    assert r.returncode == 0, r.stderr
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["개요", "기본정보", "타임라인", "웨딩홀투어", "당일체크"], wb.sheetnames
    rows = timeline_rows(wb)
    assert len(rows) == len(ids)
    tl = wb["타임라인"]
    for col in "GHKL":
        assert str(tl[f"{col}6"].value).startswith("="), col
    for ref in ("B8", "B9", "B12"):
        assert str(wb["개요"][ref].value).startswith("="), ref
    assert tl[f"J{rows['hall_contract']}"].value == "완료"
    assert tl[f"J{rows['couple_hanbok']}"].value == "해당없음"  # skip 태그 '폐백'
    check_formulas(wb)

    base = sample(SCHEDULE)
    expect_bad(SCHEDULE, [dict(base, wedding_date="2027.04.17"), dict(base, skip=["없는태그"]),
                          dict(base, done=["no_such_id"]), dict(base, shoot_date="2027-05-01"),
                          {"groom": "a"}], tmp)
    r = build(SCHEDULE, {"wedding_date": base["wedding_date"]}, tmp / "smin.xlsx")
    assert r.returncode == 0, r.stderr

    if not formulas:
        return
    for path in (out, tmp / "smin.xlsx"):
        assert_no_errors(calc(path), path.name)
    vals = calc(out)
    wedding = dt.date.fromisoformat(base["wedding_date"])
    shoot = dt.date.fromisoformat(base["shoot_date"])
    flight, tour = rows["flight"], rows["dress_tour"]
    assert vals[f"타임라인!H{flight}"] == serial(wedding) - 180  # 예식 기준
    assert vals[f"타임라인!H{tour}"] == serial(shoot) - 30  # 촬영 기준
    assert vals[f"타임라인!L{rows['hall_contract']}"] == ""  # 완료면 알림 없음

    # 촬영일을 10일 미루면 촬영 기준만 10일 이동
    wb = openpyxl.load_workbook(out)
    wb["기본정보"]["B7"] = shoot + dt.timedelta(days=10)
    wb.save(tmp / "shift.xlsx")
    shifted = calc(tmp / "shift.xlsx")
    assert shifted[f"타임라인!H{tour}"] == vals[f"타임라인!H{tour}"] + 10
    assert shifted[f"타임라인!H{flight}"] == vals[f"타임라인!H{flight}"]

    # 홀 투어 가중 총점과 1위 (이름 없는 홀이 이기면 '홀 N')
    wb = openpyxl.load_workbook(out)
    ws = wb["웨딩홀투어"]
    tot = next(r for r in range(1, ws.max_row + 1) if ws[f"A{r}"].value == "가중 총점")
    first = tot - len(data["hall_scores"])
    for r in range(first, tot):
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"] = 2, 3, 4, 5
    wb.save(tmp / "tour.xlsx")
    tv = calc(tmp / "tour.xlsx")
    assert tv[f"웨딩홀투어!E{tot}"] == 2 * 5 * len(data["hall_scores"])
    assert tv[f"웨딩홀투어!C{tot + 1}"] == "홀 3", tv[f"웨딩홀투어!C{tot + 1}"]
    for r in range(first, tot):
        ws[f"E{r}"] = 1
    wb.save(tmp / "tour.xlsx")
    assert calc(tmp / "tour.xlsx")[f"웨딩홀투어!C{tot + 1}"] == base["halls"][1]

    # 예식이 20일 남았으면 마감 지난 미완료 할 일은 '지연'
    soon = {"wedding_date": (dt.date.today() + dt.timedelta(days=20)).isoformat()}
    r = build(SCHEDULE, soon, tmp / "soon.xlsx")
    assert r.returncode == 0, r.stderr
    soon_rows = timeline_rows(openpyxl.load_workbook(tmp / "soon.xlsx"))
    sv = calc(tmp / "soon.xlsx")
    assert sv[f"타임라인!L{soon_rows['flight']}"] == "지연"
    assert sv[f"타임라인!L{soon_rows['wedding_day']}"] in ("예정", "마감 임박")


# ---------- 공통 ----------

def test_package():
    for skill in (ROOT / "skills").iterdir():
        md = (skill / "SKILL.md").read_text("utf-8")
        m = re.match(r"---\nname: (.+)\ndescription: (.+?)\n---\n", md, re.S)
        assert m and m.group(1) == skill.name, skill
        assert len(m.group(2)) <= 1024
        for py in skill.rglob("*.py"):
            assert "../" not in py.read_text("utf-8"), py  # 폴더 하나로 zip 해도 동작해야 함


def test_no_personal_data():
    deny = ROOT / ".omc" / "research" / "denylist.txt"
    if not deny.exists():
        print("  (denylist 없음, 개인정보 검사 건너뜀)")
        return
    r = subprocess.run(["grep", "-rn", "-F", "-f", str(deny), "skills", "docs", "tests", "README.md"],
                       cwd=ROOT, capture_output=True, text=True)
    assert not r.stdout.strip(), f"개인정보 의심 문자열 발견:\n{r.stdout}"


def main():
    if not formulas:
        print("  (formulas 미설치, 수식 계산 검증 건너뜀)")
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        test_budget(tmp)
        test_schedule(tmp)
        test_package()
        test_no_personal_data()
    print("OK")


if __name__ == "__main__":
    main()
