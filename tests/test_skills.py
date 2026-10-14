#!/usr/bin/env python3
"""스킬 자체 점검. 실행: python3 tests/test_skills.py

수식 계산 검증은 `pip install formulas` 가 되어 있을 때만 돈다 (없으면 건너뜀).
"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
BUDGET = ROOT / "skills" / "wedding-budget"
BUILD = BUDGET / "scripts" / "build_budget.py"
SAMPLE = BUDGET / "examples" / "sample_input.json"
FORBIDDEN_FUNCS = ("XLOOKUP", "XMATCH", "LET(", "FILTER(", "UNIQUE(", "SORT(", "SEQUENCE(")


def build(data, out):
    src = Path(out).with_suffix(".json")
    src.write_text(json.dumps(data, ensure_ascii=False), "utf-8")
    return subprocess.run([sys.executable, str(BUILD), "--input", str(src), "--out", str(out)],
                          capture_output=True, text=True)


def test_static(tmp):
    out = tmp / "b.xlsx"
    r = build(json.loads(SAMPLE.read_text("utf-8")), out)
    assert r.returncode == 0, r.stderr
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["대시보드", "기본정보", "하객식대", "지출계획", "월별자금", "기준표"], wb.sheetnames
    plan = wb["지출계획"]
    for col in "FGLQRS":  # 대관료 행의 계산 열은 수식이어야 함
        assert str(plan[f"{col}6"].value).startswith("="), (col, plan[f"{col}6"].value)
    assert str(wb["하객식대"]["B17"].value).startswith("=")
    for ref in ("B6", "B7", "B9", "B13"):
        assert str(wb["대시보드"][ref].value).startswith("="), ref
    sheets = set(wb.sheetnames)
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    assert not any(f in v.upper() for f in FORBIDDEN_FUNCS), (ws.title, c.coordinate, v)
                    for name in re.findall(r"([가-힣A-Za-z]+)!", v):
                        assert name in sheets, (ws.title, c.coordinate, name)
    return out


def test_bad_input(tmp):
    base = json.loads(SAMPLE.read_text("utf-8"))
    cases = [dict(base, region="화성"), dict(base, wedding_date="2027/04/17"),
             dict(base, items={"없는항목": {"budget": 1}}), dict(base, meal_price=-1)]
    for i, case in enumerate(cases):
        r = build(case, tmp / f"bad{i}.xlsx")
        assert r.returncode != 0 and "입력 오류" in r.stderr, (i, r.stderr)
    minimal = {k: base[k] for k in ("wedding_date", "region", "meal_price", "guaranteed_guests")}
    r = build(minimal, tmp / "min.xlsx")
    assert r.returncode == 0, r.stderr
    return tmp / "min.xlsx"


def test_benchmarks_match_research():
    bench = json.loads((BUDGET / "references" / "benchmarks.json").read_text("utf-8"))
    items = {it["name"]: it["values"] for it in bench["items"]}
    doc = (ROOT / "docs" / "research" / "wedding-cost-benchmarks-2026.md").read_text("utf-8")
    assert items["대관료"][1] == 630 and "| 대관료 | 350 | 630 |" in doc
    assert items["스드메 패키지"][0] == 298 and "| 스드메 패키지 | 298 | 350 |" in doc
    assert bench["meal"]["per_person"][1] == 8.0 and "| 1인당 식대 | 6.0 | 8.0 |" in doc
    assert items["신혼여행"][0] == 1030 and "1,030" in doc
    assert all(v is None or len(v) == len(bench["regions"]) for v in items.values())


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
    r = subprocess.run(["git", "grep", "-n", "-F", "-f", str(deny), "--", ".", ":!.omc"],
                       cwd=ROOT, capture_output=True, text=True)
    untracked = subprocess.run(["grep", "-rn", "-F", "-f", str(deny), "skills", "docs", "tests", "README.md"],
                               cwd=ROOT, capture_output=True, text=True)
    hits = (r.stdout + untracked.stdout).strip()
    assert not hits, f"개인정보 의심 문자열 발견:\n{hits}"


def test_recalc(sample, minimal, tmp):
    try:
        import formulas
    except ImportError:
        print("  (formulas 미설치, 수식 계산 검증 건너뜀)")
        return

    def calc(path):
        sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
        vals = {}
        for k, v in sol.items():
            ref = k.split("]", 1)[1].replace("'", "").upper() if "]" in k else k
            vals[ref] = v.value[0, 0] if hasattr(v, "value") else v
        return vals

    for path in (sample, minimal):
        vals = calc(path)
        errs = {k: v for k, v in vals.items() if str(v).startswith("#")}
        assert not errs, (path.name, list(errs.items())[:10])
    vals = calc(sample)
    # 샘플: 대인 215, 소인 7, 보증 220, 대인 68,000 / 소인 34,000
    assert vals["하객식대!B17"] == 220 * 68000 + 7 * 34000, vals["하객식대!B17"]
    assert vals["지출계획!Q6"] == 300 * 10000  # 경인 대관료 중간값
    assert vals["대시보드!B9"] == vals["대시보드!B6"] - vals["대시보드!B7"]
    assert vals["월별자금!G36"] == vals["대시보드!B9"]  # 모든 지급이 월별자금에 잡힘

    wb = openpyxl.load_workbook(sample)
    wb["기본정보"]["B7"] = "서울"
    seoul = tmp / "seoul.xlsx"
    wb.save(seoul)
    assert calc(seoul)["지출계획!Q6"] == 630 * 10000  # 지역을 바꾸면 기준값도 바뀜


def main():
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        sample = test_static(tmp)
        minimal = test_bad_input(tmp)
        test_benchmarks_match_research()
        test_package()
        test_no_personal_data()
        test_recalc(sample, minimal, tmp)
    print("OK")


if __name__ == "__main__":
    main()
