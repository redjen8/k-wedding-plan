#!/usr/bin/env python3
"""input.json → 결혼 자금 계획 엑셀(수식 포함).

사용법: python3 build_budget.py --input input.json --out budget.xlsx
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BENCH = json.loads((Path(__file__).resolve().parent.parent / "references" / "benchmarks.json").read_text("utf-8"))
REGIONS = BENCH["regions"]
MAN = 10000  # 통계는 만원, 엑셀은 원

FONT = "맑은 고딕"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=14, bold=True)
F_NOTE = Font(name=FONT, size=9, color="666666")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_RED = PatternFill("solid", fgColor="F8CBAD")
FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")
WON = "#,##0"
WRAP = Alignment(wrap_text=True, vertical="top")

GUEST_GROUPS = ["가족", "친구", "회사", "부모님 지인", "기타"]
PAYERS = "신랑,신부,공통,신랑측 부모,신부측 부모"
METHODS = "카드,계좌이체,현금"
STATUSES = "미정,상담,계약,완료"
ITEM_FIELDS = {"budget", "quote", "discount", "deposit", "deposit_date", "middle", "middle_date",
               "balance_date", "payer", "method", "status", "memo"}

# 지출계획 열
COLS = ["대분류", "항목", "내 예산", "견적", "할인", "확정액", "적용 금액", "계약금", "계약금일", "중도금", "중도금일",
        "잔금", "잔금일", "결제자", "결제수단", "진행", "지역 기준값", "기준 대비", "판정", "놓치기 쉬움", "팁·메모", "기준키", "비교"]
INPUT_COLS = "CDEHIJKMNOP"
FIRST = 6  # 지출계획 첫 데이터 행


def die(msg):
    sys.exit(f"입력 오류: {msg}")


def parse_date(v, field):
    if v in (None, ""):
        return None
    try:
        return dt.date.fromisoformat(str(v))
    except ValueError:
        die(f"{field} 날짜 형식은 YYYY-MM-DD 여야 합니다: {v!r}")


def num(v, field):
    if v in (None, ""):
        return None
    if not isinstance(v, (int, float)) or v < 0:
        die(f"{field} 는 0 이상의 숫자여야 합니다: {v!r}")
    return v


def style_input(c, fmt=None):
    c.font, c.fill = F_INPUT, FILL_INPUT
    if fmt:
        c.number_format = fmt


def header(ws, row, labels):
    for i, label in enumerate(labels, 1):
        c = ws.cell(row, i, label)
        c.font, c.fill, c.alignment = F_BOLD, FILL_HEAD, Alignment(horizontal="center", wrap_text=True)


def validation(ws, formula, rng):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def build_info(wb, d):
    ws = wb.create_sheet("기본정보")
    ws["A1"], ws["A1"].font = "기본정보", F_TITLE
    ws["A2"], ws["A2"].font = "노란 칸만 입력하세요. 나머지 시트는 자동 계산됩니다.", F_NOTE
    rows = [
        ("신랑", d.get("groom", "신랑"), None),
        ("신부", d.get("bride", "신부"), None),
        ("예식일", d["wedding_date"], "yyyy-mm-dd"),
        ("지역", d["region"], None),
        ("1인 식대(대인)", d["meal_price"], WON),
        ("1인 식대(소인)", d.get("child_meal_price", 0), WON),
        ("보증인원", d["guaranteed_guests"], "#,##0\"명\""),
        ("평균 축의금", d.get("avg_gift", BENCH["meal"]["avg_gift"] * MAN), WON),
        ("예비비 비율", d.get("reserve_rate", 0.03), "0%"),
        ("판정 기준(±)", 0.2, "0%"),
        ("자금흐름 시작월", "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "yyyy-mm"),
    ]
    for i, (label, value, fmt) in enumerate(rows, 4):
        ws.cell(i, 1, label).font = F_BOLD
        style_input(ws.cell(i, 2, value), fmt)
    validation(ws, ",".join(REGIONS), "B7")
    ws["C8"] = f"=\"지역 중간값 \"&TEXT(INDEX(기준표!$C$4:$G$4,MATCH($B$7,기준표!$C$3:$G$3,0)),\"#,##0\")&\"원\""
    ws["C11"] = "기본값은 NH농협 2025년 평균 이체액 117,000원"
    ws["C12"] = "다른 항목 합계에 곱해 '당일 추가금·예비비'를 만듭니다"
    ws["C13"] = "기준값 대비 이 비율보다 비싸거나 싸면 높음/낮음으로 표시"
    for r in (8, 11, 12, 13):
        ws.cell(r, 3).font = F_NOTE

    ws["A17"], ws["A17"].font = "예식 준비에 쓸 수 있는 돈 (신혼집 자금 제외)", F_BOLD
    header(ws, 18, ["자금원", "신랑", "신부", "합계"])
    funds = d.get("funds") or [{"name": n} for n in ["저축", "투자자산 현금화", "부모님 지원", "예식 전까지 추가 저축", "기타"]]
    r = 19
    for f in funds:
        ws.cell(r, 1, f.get("name", "")).font = F_BASE
        style_input(ws.cell(r, 2, num(f.get("groom"), "funds.groom")), WON)
        style_input(ws.cell(r, 3, num(f.get("bride"), "funds.bride")), WON)
        ws.cell(r, 4, f"=SUM(B{r}:C{r})").number_format = WON
        r += 1
    ws.cell(r, 1, "합계").font = F_BOLD
    for col in "BCD":
        c = ws[f"{col}{r}"]
        c.value, c.number_format, c.font = f"=SUM({col}19:{col}{r - 1})", WON, F_BOLD
    ws.column_dimensions["A"].width = 26
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["C"].width = 18
    return f"기본정보!$D${r}"  # 가용자금 합계


def build_guests(wb, d):
    ws = wb.create_sheet("하객식대")
    ws["A1"], ws["A1"].font = "하객 · 식대", F_TITLE
    ws["A2"], ws["A2"].font = "관계별 예상 인원을 입력하세요. 가족 단위로 오면 대인·소인으로 나눠 적습니다.", F_NOTE
    header(ws, 4, ["관계", "신랑측 대인", "신랑측 소인", "신부측 대인", "신부측 소인", "합계"])
    guests = d.get("guests", {})
    for i, g in enumerate(GUEST_GROUPS, 5):
        ws.cell(i, 1, g).font = F_BASE
        for j, side in enumerate(["신랑", "신부"]):
            adult, child = (list(guests.get(side, {}).get(g) or []) + [None, None])[:2]
            style_input(ws.cell(i, 2 + j * 2, num(adult, f"guests.{side}.{g}")), "#,##0")
            style_input(ws.cell(i, 3 + j * 2, num(child, f"guests.{side}.{g}")), "#,##0")
        ws.cell(i, 6, f"=SUM(B{i}:E{i})")
    ws.cell(10, 1, "합계").font = F_BOLD
    for col in "BCDEF":
        ws[f"{col}10"], ws[f"{col}10"].font = f"=SUM({col}5:{col}9)", F_BOLD

    calc = [
        ("예상 하객(대인)", "=B10+D10", "#,##0\"명\""),
        ("예상 하객(소인)", "=C10+E10", "#,##0\"명\""),
        ("보증인원", "=기본정보!B10", "#,##0\"명\""),
        ("식대 산정 인원(대인)", "=MAX(B13,B15)", "#,##0\"명\""),
        ("예상 식대", "=B16*기본정보!B8+B14*기본정보!B9", WON),
        ("보증인원 - 예상 하객", "=B15-B13", "#,##0\"명\";-#,##0\"명\""),
        ("빈자리 식대", "=MAX(B18,0)*기본정보!B8", WON),
        ("축의 봉투 수(추정)", "=B13", "#,##0"),
        ("예상 축의금", "=B20*기본정보!B11", WON),
        ("지역 1인 식대 중간값", "=INDEX(기준표!$C$4:$G$4,MATCH(기본정보!$B$7,기준표!$C$3:$G$3,0))", WON),
        ("지역 최소 보증인원 중간값", "=INDEX(기준표!$C$5:$G$5,MATCH(기본정보!$B$7,기준표!$C$3:$G$3,0))", "#,##0\"명\""),
    ]
    for i, (label, f, fmt) in enumerate(calc, 13):
        ws.cell(i, 1, label).font = F_BOLD
        c = ws.cell(i, 2, f)
        c.number_format = fmt
    style_input(ws["B20"], "#,##0")  # 사용자가 덮어쓸 수 있는 추정값
    ws["C18"] = '=IF(B18>0,"보증인원이 예상 하객보다 많습니다. 보증인원 조정을 홀과 협의해 보세요.","")'
    ws["C20"] = "가족은 한 봉투로 오는 경우가 많습니다. 실제에 맞게 줄여서 입력하세요."
    for r in (18, 20):
        ws.cell(r, 3).font = F_NOTE
    ws.column_dimensions["A"].width = 26
    for col in "BCDEF":
        ws.column_dimensions[col].width = 14
    return "하객식대!$B$17", "하객식대!$B$21"


def build_plan(wb, d, meal_ref):
    ws = wb.create_sheet("지출계획")
    ws["A1"], ws["A1"].font = "지출계획", F_TITLE
    ws["A2"] = ("노란 칸만 입력하세요. 견적을 넣으면 확정액이 생기고, 확정액이 없으면 내 예산으로 계산합니다. "
                "계약금·중도금을 넣으면 잔금은 자동입니다. 날짜를 넣어야 월별자금에 잡힙니다.")
    ws["A3"] = f"지역 기준값: 참가격 {BENCH['as_of']} 중간값(지역별) 또는 듀오 2026 평균(전국). ★ = 놓치기 쉬운 비용"
    for c in ("A2", "A3"):
        ws[c].font = F_NOTE
    header(ws, 5, COLS)

    items = dict(d.get("items") or {})
    known = {it["name"] for it in BENCH["items"]}
    unknown = set(items) - known
    if unknown:
        die(f"알 수 없는 항목 {sorted(unknown)}. 기본 항목은 {sorted(known)}. 새 항목은 extra_items 에 넣으세요.")
    rows = [dict(it, input=items.get(it["name"], {})) for it in BENCH["items"]]
    reserve = next(i for i, it in enumerate(rows) if it.get("reserve"))
    for e in d.get("extra_items") or []:  # 추가 항목은 예비비 행 바로 앞에
        if not e.get("name"):
            die("extra_items 의 각 항목에는 name 이 필요합니다")
        rows.insert(reserve, {"category": e.get("category", "기타"), "name": e["name"], "values": None, "input": e})
        reserve += 1

    region = "기본정보!$B$7"
    last = FIRST + len(rows) - 1
    for r, it in enumerate(rows, FIRST):
        inp = it["input"]
        bad = set(inp) - ITEM_FIELDS - {"category", "name"}
        if bad:
            die(f"{it['name']} 에 알 수 없는 필드 {sorted(bad)}")
        ws.cell(r, 1, it["category"])
        ws.cell(r, 2, it["name"])
        budget, memo = inp.get("budget"), inp.get("memo") or it.get("tip", "")
        if budget == "avg":  # 모르는 값은 지역 중간값으로 추정
            if not it.get("values"):
                die(f"{it['name']} 은 기준값이 없어 avg 를 쓸 수 없습니다")
            budget = it["values"][REGIONS.index(d["region"])] * MAN
            memo = f"[추정: 지역 기준값] {memo}"
        if it["name"] == "식대" and budget is None:
            budget = f"={meal_ref}"
        if it.get("reserve") and budget is None:
            budget = f"=ROUND(기본정보!$B$12*SUM(G{FIRST}:G{r - 1}),-4)"
        values = {"C": budget if isinstance(budget, str) else num(budget, f"{it['name']}.budget"),
                  "D": num(inp.get("quote"), f"{it['name']}.quote"),
                  "E": num(inp.get("discount"), f"{it['name']}.discount"),
                  "H": num(inp.get("deposit"), f"{it['name']}.deposit"),
                  "I": parse_date(inp.get("deposit_date"), f"{it['name']}.deposit_date"),
                  "J": num(inp.get("middle"), f"{it['name']}.middle"),
                  "K": parse_date(inp.get("middle_date"), f"{it['name']}.middle_date"),
                  "M": parse_date(inp.get("balance_date"), f"{it['name']}.balance_date"),
                  "N": inp.get("payer"), "O": inp.get("method"), "P": inp.get("status", "미정")}
        for col in INPUT_COLS:
            c = ws[f"{col}{r}"]
            c.value = values.get(col)
            style_input(c, "yyyy-mm-dd" if col in "IKM" else (WON if col in "CDEHJ" else None))
        ws[f"F{r}"] = f'=IF(D{r}="","",MAX(D{r}-E{r},0))'
        ws[f"G{r}"] = f'=IF(F{r}<>"",F{r},IF(C{r}<>"",C{r},0))'
        ws[f"L{r}"] = f'=IF(G{r}=0,"",MAX(G{r}-H{r}-J{r},0))'
        comparable = it.get("comparable", True)
        if it.get("values"):
            ws[f"Q{r}"] = (f"=INDEX(기준표!$C$7:$G$200,MATCH($V{r},기준표!$B$7:$B$200,0),"
                           f"MATCH({region},기준표!$C$3:$G$3,0))")
            ws[f"R{r}"] = f'=IF(OR(G{r}=0,Q{r}=0),"",G{r}/Q{r})'
            if comparable:
                ws[f"S{r}"] = (f'=IF(R{r}="","",IF(R{r}>1+기본정보!$B$13,"높음",'
                               f'IF(R{r}<1-기본정보!$B$13,"낮음","평균 수준")))')
            else:
                ws[f"S{r}"] = f'=IF(R{r}="","","참고")'
        else:
            ws[f"Q{r}"] = "기준 없음"
        ws[f"T{r}"] = "★" if it.get("hidden") else ""
        ws[f"U{r}"] = memo
        ws[f"V{r}"] = it["name"]
        ws[f"W{r}"] = "Y" if it.get("values") and comparable else "N"
        for col in "ABFGLQRSTUVW":
            ws[f"{col}{r}"].font = F_BASE
        for col in "FGLQ":
            ws[f"{col}{r}"].number_format = WON
        ws[f"R{r}"].number_format = '0.00"배"'

    tot = last + 1
    ws.cell(tot, 1, "합계").font = F_BOLD
    for col in "CDEFGHJL":
        c = ws[f"{col}{tot}"]
        c.value, c.number_format, c.font = f"=SUM({col}{FIRST}:{col}{last})", WON, F_BOLD

    validation(ws, PAYERS, f"N{FIRST}:N{last}")
    validation(ws, METHODS, f"O{FIRST}:O{last}")
    validation(ws, STATUSES, f"P{FIRST}:P{last}")
    s_rng = f"S{FIRST}:S{last}"
    ws.conditional_formatting.add(s_rng, CellIsRule(operator="equal", formula=['"높음"'], fill=FILL_RED))
    ws.conditional_formatting.add(s_rng, CellIsRule(operator="equal", formula=['"낮음"'], fill=FILL_GREEN))
    ws.conditional_formatting.add(f"A{FIRST}:B{last}",
                                  FormulaRule(formula=[f'AND($T{FIRST}="★",$G{FIRST}=0)'], fill=FILL_RED))
    widths = [10, 22, 13, 13, 11, 13, 13, 12, 11, 12, 11, 12, 11, 10, 10, 8, 13, 9, 10, 8, 50]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(5, i).column_letter].width = w
    ws.column_dimensions["V"].hidden = ws.column_dimensions["W"].hidden = True
    ws.freeze_panes = f"C{FIRST}"
    ws["B4"] = "빨간 항목명 = 놓치기 쉬운 비용인데 아직 0원"
    ws["B4"].font = F_NOTE
    return FIRST, last


def build_cashflow(wb, first, last, funds_ref):
    ws = wb.create_sheet("월별자금")
    ws["A1"], ws["A1"].font = "월별 자금 흐름", F_TITLE
    ws["A2"] = "지출계획의 계약금일·중도금일·잔금일 기준입니다. 날짜가 없는 금액은 '날짜 미정'에 모입니다."
    ws["A2"].font = F_NOTE
    header(ws, 4, ["월", "계약금", "중도금", "잔금", "합계", "누적", "남은 가용자금"])
    p = "지출계획!"
    pay = [(f"{p}$H${first}:$H${last}", f"{p}$I${first}:$I${last}"),
           (f"{p}$J${first}:$J${last}", f"{p}$K${first}:$K${last}"),
           (f"{p}$L${first}:$L${last}", f"{p}$M${first}:$M${last}")]
    ws["A5"] = "시작월 이전"
    ws["A6"] = "날짜 미정"
    for i, (amt, day) in enumerate(pay):
        col = "BCD"[i]
        ws[f"{col}5"] = f'=SUMIFS({amt},{day},"<"&$A$7)'
        ws[f"{col}6"] = f'=SUMIFS({amt},{day},"")'
    months = 30
    for k in range(months):
        r = 7 + k
        ws[f"A{r}"] = "=기본정보!$B$14" if k == 0 else f"=EDATE(A{r - 1},1)"
        ws[f"A{r}"].number_format = "yyyy-mm"
        for i, (amt, day) in enumerate(pay):
            ws[f"{'BCD'[i]}{r}"] = f'=SUMIFS({amt},{day},">="&$A{r},{day},"<"&EDATE($A{r},1))'
    end = 7 + months - 1
    for r in range(5, end + 1):
        ws[f"E{r}"] = f"=SUM(B{r}:D{r})"
        ws[f"F{r}"] = f"=E{r}" if r == 5 else f"=F{r - 1}+E{r}"
        ws[f"G{r}"] = f"={funds_ref}-F{r}"
        for col in "BCDEFG":
            ws[f"{col}{r}"].number_format = WON
    ws.conditional_formatting.add(f"G5:G{end}", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
    ws.column_dimensions["A"].width = 12
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 15


def build_bench(wb):
    ws = wb.create_sheet("기준표")
    ws["A1"], ws["A1"].font = f"지역 기준값 (단위: 원, {BENCH['as_of']} 기준 스냅샷)", F_TITLE
    ws["A2"] = "참가격 = 중간값(지역별), 듀오 = 설문 평균(전국 공통). 수정하지 마세요. 갱신은 스킬 배포본에서 합니다."
    ws["A2"].font = F_NOTE
    header(ws, 3, ["구분", "항목"] + REGIONS + ["비교", "출처"])
    ws["B4"], ws["B5"] = "1인 식대", "최소 보증인원"
    for i in range(5):
        ws.cell(4, 3 + i, BENCH["meal"]["per_person"][i] * MAN).number_format = WON
        ws.cell(5, 3 + i, BENCH["meal"]["min_guaranteed_guests"][i])
    ws["I4"] = ws["I5"] = "참가격"
    for r, it in enumerate([x for x in BENCH["items"] if x.get("values")], 7):
        ws.cell(r, 1, it["category"])
        ws.cell(r, 2, it["name"])
        for i, v in enumerate(it["values"]):
            ws.cell(r, 3 + i, v * MAN).number_format = WON
        ws.cell(r, 8, "비교" if it.get("comparable", True) else "참고")
        ws.cell(r, 9, it["source"])
    r += 2
    ws.cell(r, 1, "출처").font = F_BOLD
    for k, v in BENCH["sources"].items():
        r += 1
        ws.cell(r, 1, k)
        ws.cell(r, 2, v)
    r += 2
    ws.cell(r, 1, "참고(듀오 평균, 이 표 계산에는 안 씀)").font = F_BOLD
    for k, v in BENCH["reference_only"].items():
        r += 1
        ws.cell(r, 1, k)
        ws.cell(r, 3, v * MAN).number_format = WON
    ws.column_dimensions["B"].width = 24
    for col in "CDEFG":
        ws.column_dimensions[col].width = 12


def build_dashboard(ws, first, last, funds_ref, meal_ref, gift_ref):
    ws["A1"], ws["A1"].font = "결혼 자금 대시보드", F_TITLE
    ws["A2"] = f"지역 기준값은 참가격 {BENCH['as_of']} 중간값 / 듀오 2026 평균 스냅샷입니다. 최신 시세와 다를 수 있습니다."
    ws["A2"].font = F_NOTE
    p = "지출계획!"
    g, q, s, t, w, a = (f"{p}${c}${first}:${c}${last}" for c in "GQSTWA")
    kpis = [
        ("D-day", '=IF(기본정보!B6="","",기본정보!B6-TODAY())', '"D-"0;"D+"0'),
        ("지역", "=기본정보!B7", None),
        ("예식 준비 가용자금", f"={funds_ref}", WON),
        ("예상 총지출", f"=SUM({g})", WON),
        ("확정액 합계", f"=SUM({p}$F${first}:$F${last})", WON),
        ("부족(-) / 여유(+)", "=B6-B7", WON),
        ("예상 식대", f"={meal_ref}", WON),
        ("예상 축의금", f"={gift_ref}", WON),
        ("축의금 반영 순부담", "=B7-B11", WON),
        ("기준값 대비 배율 (비교 가능 항목)", f'=IFERROR(SUMIFS({g},{w},"Y",{g},">0")/SUMIFS({q},{w},"Y",{g},">0"),"")', '0.00"배"'),
        ("'높음' 항목 수", f'=COUNTIF({s},"높음")', "0\"개\""),
        ("놓치기 쉬운 비용 중 0원 항목", f'=COUNTIFS({t},"★",{g},0)', "0\"개\""),
    ]
    for i, (label, f, fmt) in enumerate(kpis, 4):
        ws.cell(i, 1, label).font = F_BOLD
        c = ws.cell(i, 2, f)
        c.font = F_BOLD
        if fmt:
            c.number_format = fmt
    ws.conditional_formatting.add("B9", CellIsRule(operator="lessThan", formula=["0"], fill=FILL_RED))
    ws["C9"] = '=IF(B9<0,"가용자금보다 예상 지출이 많습니다","")'
    ws["C15"] = '=IF(B15>0,"지출계획에서 빨간 항목명을 확인하세요","")'
    ws["C9"].font = ws["C15"].font = F_NOTE

    ws["A18"], ws["A18"].font = "대분류별", F_BOLD
    header(ws, 19, ["대분류", "예상 지출", "기준값 합계", "배율", "비중"])
    cats = list(dict.fromkeys(it["category"] for it in BENCH["items"]))
    for r, cat in enumerate(cats, 20):
        ws.cell(r, 1, cat)
        ws.cell(r, 2, f'=SUMIFS({g},{a},A{r})').number_format = WON
        ws.cell(r, 3, f'=SUMIFS({q},{a},A{r},{w},"Y",{g},">0")').number_format = WON
        ws.cell(r, 4, f'=IF(C{r}=0,"",SUMIFS({g},{a},A{r},{w},"Y",{g},">0")/C{r})').number_format = '0.00"배"'
        ws.cell(r, 5, f'=IF($B$7=0,"",B{r}/$B$7)').number_format = "0%"
    end = r + 1
    ws.cell(end, 1, "기타(추가 항목)")
    ws.cell(end, 2, f"=B7-SUM(B20:B{r})").number_format = WON
    ws.cell(end + 2, 1, "참고: 듀오 2026 평균 혼수 "
            f"{BENCH['reference_only']['혼수']:,}만원, 신혼집 {BENCH['reference_only']['신혼집']:,}만원 (이 표에는 포함 안 함)").font = F_NOTE
    ws.column_dimensions["A"].width = 32
    for col in "BCDE":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["C"].width = 34


def build(d):
    for k in ("wedding_date", "region", "meal_price", "guaranteed_guests"):
        if d.get(k) in (None, ""):
            die(f"필수 값 {k} 가 없습니다")
    if d["region"] not in REGIONS:
        die(f"region 은 {REGIONS} 중 하나여야 합니다: {d['region']!r}")
    d["wedding_date"] = parse_date(d["wedding_date"], "wedding_date")
    num(d["meal_price"], "meal_price")
    num(d["guaranteed_guests"], "guaranteed_guests")

    wb = Workbook()
    dash = wb.active
    dash.title = "대시보드"
    funds_ref = build_info(wb, d)
    meal_ref, gift_ref = build_guests(wb, d)
    first, last = build_plan(wb, d, meal_ref)
    build_cashflow(wb, first, last, funds_ref)
    build_bench(wb)
    build_dashboard(dash, first, last, funds_ref, meal_ref, gift_ref)
    return wb


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    data = json.loads(Path(a.input).read_text("utf-8"))
    build(data).save(a.out)
    print(f"생성 완료: {a.out}")


if __name__ == "__main__":
    main()
