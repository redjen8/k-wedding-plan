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


def check_cf_colors(wb):
    """조건부서식 배경은 Excel이 bgColor 로 읽는다. fgColor 만 있으면 색이 안 보인다."""
    for ws in wb:
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                bg = rule.dxf.fill.bgColor.rgb if rule.dxf and rule.dxf.fill else None
                assert bg and bg != "00000000", (ws.title, cf.sqref, rule.formula)


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
    check_cf_colors(wb)

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

def timeline(wb):
    """타임라인 시트의 (열 이름 → 열 문자, 할 일 id → 행)."""
    ws = wb["타임라인"]
    cols = {c.value: c.column_letter for c in ws[5] if c.value}
    idc = cols["id"]
    rows = {ws[f"{idc}{r}"].value: r for r in range(6, ws.max_row + 1) if ws[f"{idc}{r}"].value}
    return cols, rows


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
    col, rows = timeline(wb)
    assert len(rows) == len(ids)
    tl = wb["타임라인"]
    for name in ("시작 시기", "마감 시기", "시작일", "마감일", "남은 일수", "알림"):
        assert str(tl[f"{col[name]}6"].value).startswith("="), name
    for ref in ("B8", "B9", "B12"):
        assert str(wb["개요"][ref].value).startswith("="), ref
    st, due = col["상태"], col["마감일"]
    widths = {"단계": 16.83, "할 일": 42, "시작 시기": 15, "팁": 54.66, "메모": 10.66}  # Excel에서 맞춘 너비 유지
    for name, w in widths.items():
        assert tl.column_dimensions[col[name]].width == w, (name, tl.column_dimensions[col[name]].width)
    assert tl.column_dimensions[col["id"]].hidden
    assert tl[f"{st}{rows['hall_contract']}"].value == "완료"
    assert tl[f"{st}{rows['couple_hanbok']}"].value == "해당없음"  # skip 태그 '폐백'
    check_formulas(wb)
    check_cf_colors(wb)
    # 시작(D-)·마감(D-) 칸은 고정 노랑 없이 알림에 따라 색이 바뀐다
    s_col = col["시작(D-)"]
    assert tl[f"{s_col}6"].fill.fill_type is None
    assert any(str(cf.sqref).startswith(s_col) for cf in tl.conditional_formatting)
    # 마감일 → 시작일 순으로 급한 일이 위, 가장 나중 일이 맨 아래
    wd, sd = (dt.date.fromisoformat(sample(SCHEDULE)[k]) for k in ("wedding_date", "shoot_date"))
    keys = []
    for r in sorted(rows.values()):
        b = sd if tl[f"{col['기준']}{r}"].value == "촬영일" else wd
        keys.append((b - dt.timedelta(days=tl[f"{col['마감(D-)']}{r}"].value),
                     b - dt.timedelta(days=tl[f"{col['시작(D-)']}{r}"].value)))
    assert keys == sorted(keys), "타임라인이 마감일·시작일 순이 아님"

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
    assert vals[f"타임라인!{due}{flight}"] == serial(wedding) - 180  # 예식 기준
    assert vals[f"타임라인!{due}{tour}"] == serial(shoot) - 30  # 촬영 기준
    assert vals[f"타임라인!{col['알림']}{rows['hall_contract']}"] == ""  # 완료면 알림 없음
    # D- 일수 → 읽기용 시기 문구
    expect_when = {("first_greeting", "시작 시기"): "1년 6개월 전", ("first_greeting", "마감 시기"): "1년 4개월 전",
                   ("snap_book", "마감 시기"): "1년 전", ("ring", "마감 시기"): "4개월 전",
                   ("invite_mobile", "마감 시기"): "21일 전", ("shoot_day", "마감 시기"): "당일",
                   ("photo_select", "시작 시기"): "14일 후", ("photo_select", "마감 시기"): "1개월 후"}
    for (tid, name), text in expect_when.items():
        assert vals[f"타임라인!{col[name]}{rows[tid]}"] == text, (tid, name, vals[f"타임라인!{col[name]}{rows[tid]}"])

    # 촬영일을 10일 미루면 촬영 기준만 10일 이동
    wb = openpyxl.load_workbook(out)
    wb["기본정보"]["B7"] = shoot + dt.timedelta(days=10)
    wb.save(tmp / "shift.xlsx")
    shifted = calc(tmp / "shift.xlsx")
    assert shifted[f"타임라인!{due}{tour}"] == vals[f"타임라인!{due}{tour}"] + 10
    assert shifted[f"타임라인!{due}{flight}"] == vals[f"타임라인!{due}{flight}"]

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
    _, soon_rows = timeline(openpyxl.load_workbook(tmp / "soon.xlsx"))
    sv = calc(tmp / "soon.xlsx")
    assert sv[f"타임라인!{col['알림']}{soon_rows['flight']}"] == "지연"
    assert sv[f"타임라인!{col['알림']}{soon_rows['wedding_day']}"] in ("예정", "마감 임박")


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
