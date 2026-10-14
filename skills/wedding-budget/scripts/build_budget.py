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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

BENCH = json.loads((Path(__file__).resolve().parent.parent / "references" / "benchmarks.json").read_text("utf-8"))
REGIONS = BENCH["regions"]
MAN = 10000  # 통계는 만원, 엑셀은 원

FONT = "맑은 고딕"
BODY = 11  # 표 본문 글자 크기. polish()는 이 크기인 셀만 표로 본다
F_BASE = Font(name=FONT, size=BODY)
F_INPUT = Font(name=FONT, size=BODY, color="0000FF")
F_BOLD = Font(name=FONT, size=BODY, bold=True)
F_TITLE = Font(name=FONT, size=16, bold=True)
F_SECTION = Font(name=FONT, size=12, bold=True, color="1F4E79")
F_NOTE = Font(name=FONT, size=10, color="666666")
THIN = Side(style="thin", color="C9D1E0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
# 조건부서식 색. Excel은 조건부서식 배경을 bgColor 로 읽어서 두 값을 모두 준다
FILL_RED = PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD")
FILL_GREEN = PatternFill("solid", fgColor="C6EFCE", bgColor="C6EFCE")
WON = "#,##0"
WRAP = Alignment(wrap_text=True, vertical="top")

GUEST_GROUPS = ["가족", "친구", "회사", "부모님 지인", "기타"]
PAYERS = ["신랑", "신부", "공통", "신랑측 부모", "신부측 부모"]
METHODS = ["카드", "계좌이체", "현금"]
RECEIPTS = ["발급완료", "미발급", "해당없음"]
PAY_TYPES = ["계약금", "중도금", "잔금", "현장 추가금", "기타"]
STATUSES = "미정,상담,계약,완료"
ITEM_FIELDS = {"budget", "quote", "discount", "deposit", "deposit_date", "middle", "middle_date",
               "balance_date", "status", "memo"}
PAYMENT_FIELDS = {"date", "item", "type", "amount", "payer", "method", "receipt", "memo"}

# 지출계획 열. 계약금~잔금일은 '예정', 실제 결제는 결제기록 시트에 쌓고 N열(납부 완료)로 합산한다
COLS = ["대분류", "항목", "내 예산", "견적", "할인", "확정액", "적용 금액", "계약금", "계약금일", "중도금", "중도금일",
        "잔금", "잔금일", "납부 완료", "남은 금액", "진행", "지역 기준값", "기준 대비", "판정", "놓치기 쉬움", "팁·메모", "기준키", "비교"]
INPUT_COLS = "CDEHIJKMP"
FIRST = 6  # 지출계획 첫 데이터 행
PAY_FIRST, PAY_ROWS = 6, 200  # 결제기록 입력 행
DRESS_SHOPS = 3  # 드레스투어 비교 샵 수. 보통 2~3곳 투어
PAY_LAST = PAY_FIRST + PAY_ROWS - 1
PAY = "결제기록!"
PAY_ITEM, PAY_AMT = f"{PAY}$B${PAY_FIRST}:$B${PAY_LAST}", f"{PAY}$D${PAY_FIRST}:$D${PAY_LAST}"


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


def header(ws, row, labels, col=1):
    for i, label in enumerate(labels, col):
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

    ws["A17"], ws["A17"].font = "예식 준비에 쓸 수 있는 돈 (신혼집 자금 제외)", F_SECTION
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
    return "하객식대!$B$17", "하객식대!$B$21"


def build_plan(wb, d, meal_ref):
    ws = wb.create_sheet("지출계획")
    ws["A1"], ws["A1"].font = "지출계획", F_TITLE
    ws["A2"] = ("노란 칸만 입력하세요. 견적을 넣으면 확정액이 생기고, 확정액이 없으면 내 예산으로 계산합니다. "
                "계약금·중도금·잔금은 '예정'이고 날짜를 넣어야 월별자금에 잡힙니다. "
                "실제로 낸 돈은 결제기록 시트에 적으면 '납부 완료'에 합산됩니다.")
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
        if e["name"] in known or any(ch in e["name"] for ch in "*?~"):  # 결제기록 SUMIFS 가 이름으로 합산
            die(f"extra_items 이름은 기본 항목과 겹치지 않고 *?~ 를 쓰지 않아야 합니다: {e['name']!r}")
        known.add(e["name"])
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
                  "P": inp.get("status", "미정")}
        for col in INPUT_COLS:
            c = ws[f"{col}{r}"]
            c.value = values.get(col)
            style_input(c, "yyyy-mm-dd" if col in "IKM" else (WON if col in "CDEHJ" else None))
        ws[f"F{r}"] = f'=IF(D{r}="","",MAX(D{r}-N(E{r}),0))'  # E 가 드레스투어 연동으로 "" 일 수 있음
        ws[f"G{r}"] = f'=IF(F{r}<>"",F{r},IF(C{r}<>"",C{r},0))'
        ws[f"L{r}"] = f'=IF(G{r}=0,"",MAX(G{r}-H{r}-J{r},0))'
        ws[f"N{r}"] = f"=SUMIFS({PAY_AMT},{PAY_ITEM},$B{r})"
        ws[f"O{r}"] = f'=IF(G{r}=0,"",MAX(G{r}-N{r},0))'
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
        for col in "ABFGLNOQRSTUVW":
            ws[f"{col}{r}"].font = F_BASE
        for col in "FGLNOQ":
            ws[f"{col}{r}"].number_format = WON
        ws[f"R{r}"].number_format = '0.00"배"'

    tot = last + 1
    ws.cell(tot, 1, "합계").font = F_BOLD
    for col in "CDEFGHJLNO":
        c = ws[f"{col}{tot}"]
        c.value, c.number_format, c.font = f"=SUM({col}{FIRST}:{col}{last})", WON, F_BOLD

    validation(ws, STATUSES, f"P{FIRST}:P{last}")
    s_rng = f"S{FIRST}:S{last}"
    ws.conditional_formatting.add(s_rng, CellIsRule(operator="equal", formula=['"높음"'], fill=FILL_RED))
    ws.conditional_formatting.add(s_rng, CellIsRule(operator="equal", formula=['"낮음"'], fill=FILL_GREEN))
    ws.conditional_formatting.add(f"A{FIRST}:B{last}",
                                  FormulaRule(formula=[f'AND($T{FIRST}="★",$G{FIRST}=0)'], fill=FILL_RED))
    ws.column_dimensions["V"].hidden = ws.column_dimensions["W"].hidden = True
    ws.freeze_panes = f"C{FIRST}"
    ws["B4"] = "빨간 항목명 = 놓치기 쉬운 비용인데 아직 0원"
    ws["B4"].font = F_NOTE
    return FIRST, last


def build_dress_tour(wb, d, first, last):
    ws = wb.create_sheet("드레스투어")
    ws["A1"], ws["A1"].font = "드레스 투어 비교", F_TITLE
    ws["A2"] = ("방식을 고르고 샵별 비용을 적은 뒤, 정한 샵을 '선택한 샵'에서 고르면 지출계획 견적에 들어갑니다. "
                "스드메 패키지로 계약했으면 '선택한 샵'을 비워 두세요. 지출계획 견적 칸에 숫자를 직접 치면 연동이 끊깁니다.")
    ws["A3"] = ("투어: 샵 2~3곳, 이름을 적은 샵마다 '샵당 피팅비'가 붙습니다(샵마다 다르면 덮어쓰세요). "
                "지정: 샵 1에만 적고 피팅비는 없습니다. 지정 혜택 중 금전 혜택은 '할인·지정 혜택'에, 그 밖의 혜택은 아래 '지정 혜택'에 적으세요.")
    for c in ("A2", "A3"):
        ws[c].font = F_NOTE
    header(ws, 4, ["분류", "항목"] + [f"샵 {i}" for i in range(1, DRESS_SHOPS + 1)])
    shops = d.get("dress_shops") or []
    last_col = chr(ord("C") + DRESS_SHOPS - 1)
    others = f"D5:{last_col}5"  # 지정이면 비워야 하는 샵 2~
    for r, label in ((5, "샵 이름"), (6, "방식"), (7, "선택한 샵"), (8, "지출계획에 들어가는 샵"), (9, "샵당 피팅비")):
        ws.cell(r, 2, label).font = F_BOLD
    for i in range(DRESS_SHOPS):
        style_input(ws.cell(5, 3 + i, shops[i] if i < len(shops) else None))  # 빈 열은 피팅비가 붙지 않음
    style_input(ws.cell(6, 3, d.get("dress_mode") or "투어"))
    style_input(ws["C7"])
    style_input(ws.cell(9, 3, BENCH["dress_tour"]["fitting_fee"]), WON)
    validation(ws, "투어,지정", "C6")
    dv = DataValidation(type="list", formula1=f"$C$5:${last_col}$5", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C7")
    dv = DataValidation(type="custom", formula1='$C$6<>"지정"', allow_blank=True, showErrorMessage=True,
                        errorTitle="지정은 샵 1곳", error="지정이면 샵 1에만 적습니다. 투어로 바꾸면 샵을 더 적을 수 있습니다.")
    ws.add_data_validation(dv)
    dv.add(others)
    # 지정이면 샵 1만 지출계획에 들어간다
    ws["C8"] = (f'=IF(C7="","",IF(ISNA(MATCH(C7,$C$5:${last_col}$5,0)),"",'
                'IF(AND(C6="지정",C7<>C5),"",C7)))')
    ws["C8"].font = F_BOLD
    plan = wb["지출계획"]
    prow = {plan[f"B{r}"].value: r for r in range(first, last + 1)}
    ws["D6"] = f'=IF(AND(C6="지정",COUNTIF({others},"?*")>0),"지정은 샵 1곳만 적습니다. 샵 2·3 이름을 지우세요","")'
    ws["D7"] = (f'=IF(C7="","",IF(C8="","선택한 샵이 목록에 없거나, 지정인데 샵 1이 아닙니다. 다시 고르세요",'
                f'IF(지출계획!G{prow["스드메 패키지"]}>0,"스드메 패키지에도 금액이 있습니다. 중복이 아닌지 확인하세요","")))')
    ws["D6"].font = ws["D7"].font = F_NOTE

    links, cost_rows, fit = {}, [], None
    costs = list(enumerate(BENCH["dress_tour"]["costs"], 10))
    for r, c in costs:
        ws.cell(r, 1, "비용")
        ws.cell(r, 2, c["label"])
        for i in range(DRESS_SHOPS):
            x = chr(ord("C") + i)
            if c.get("fitting"):  # 투어일 때 이름 있는 샵만 피팅비
                ws[f"{x}{r}"] = f'=IF(OR({x}$5="",$C$6="지정"),0,$C$9)'
            style_input(ws[f"{x}{r}"], WON)
        cost_rows.append((r, c.get("minus")))
        if c.get("fitting"):
            fit = r
        else:
            links.setdefault(c["link"], {"D": [], "E": []})["E" if c.get("minus") else "D"].append(r)
    tot, bur, ref, ratio, cheap = r + 1, r + 2, r + 3, r + 4, r + 5
    base = next(r for r, c in costs if c["link"] == "드레스(단품)")
    fits = f"$C${fit}:${last_col}${fit}"
    for i in range(DRESS_SHOPS):
        x = chr(ord("C") + i)
        plus = "+".join(f"{x}{r}" for r, m in cost_rows if not m)
        minus = "".join(f"-{x}{r}" for r, m in cost_rows if m)
        # 드레스 기본가를 적은 샵만 총액을 낸다. 피팅비·혜택만 적힌 샵이 '최저'로 뽑히지 않게. 지정이면 샵 1만
        skip = f'{x}$5="",N({x}{base})=0' + (',$C$6="지정"' if i else "")
        ws[f"{x}{tot}"] = f'=IF(OR({skip}),"",{plus}{minus})'
        # 투어는 투어한 샵 피팅비를 모두 낸다. 지정이면 피팅비가 전부 0
        ws[f"{x}{bur}"] = f'=IF({x}{tot}="","",{x}{tot}-N({x}{fit})+SUM({fits}))'
        ws[f"{x}{ratio}"] = f'=IF(OR(N({x}{base})=0,N($C${ref})=0),"",{x}{base}/$C${ref})'
        for r in (tot, bur):
            ws[f"{x}{r}"].number_format = WON
        ws[f"{x}{ratio}"].number_format = '0.00"배"'
    ws[f"C{ref}"] = f'=지출계획!Q{prow["드레스(단품)"]}'
    ws[f"C{ref}"].number_format = WON
    burs = f"C{bur}:{last_col}{bur}"
    ws[f"C{cheap}"] = (f'=IF(COUNT({burs})=0,"",INDEX($C$5:${last_col}$5,MATCH(MIN({burs}),{burs},0)))')
    for r, label in ((tot, "샵 총액"), (bur, "이 샵으로 정할 때 총부담"), (ref, "지역 기준값(드레스 단품)"),
                     (ratio, "기준 대비(드레스 기본가)"), (cheap, "총부담 최저")):
        ws.cell(r, 1, "합계")
        ws.cell(r, 2, label).font = F_BOLD
        for col in range(3, 3 + DRESS_SHOPS):
            ws.cell(r, col).font = F_BOLD

    # 선택한 샵 금액을 지출계획 견적(D)·할인(E)에 연동. input 에 견적·할인을 직접 준 항목은 그 값이 우선
    sel, names = "드레스투어!$C$8", f"드레스투어!$C$5:${last_col}$5"
    pick = lambda r: f"INDEX(드레스투어!$C${r}:${last_col}${r},MATCH({sel},{names},0))"
    items = d.get("items") or {}
    for name, cols in links.items():
        if {"quote", "discount"} & set(items.get(name, {})):
            continue
        for col, rs in cols.items():
            if rs:
                e = "+".join(pick(r) for r in rs)
                typed = ",".join(f"ISNUMBER({pick(r)})" for r in rs)  # 0원(포함)도 입력으로 본다
                plan[f"{col}{prow[name]}"] = f'=IF({sel}="","",IFERROR(IF(OR({typed}),{e},""),""))'
    name = "드레스 투어 피팅비"
    if not {"quote", "discount"} & set(items.get(name, {})):
        # 선택과 상관없이 투어한 샵 피팅비 합계(지정이면 0). 이름 있는 샵이 없으면 예산으로
        plan[f"D{prow[name]}"] = f'=IF(COUNTIF({names},"?*")=0,"",SUM(드레스투어!{fits}))'

    r = cheap + 2
    ws.cell(r, 1, "특징").font = F_SECTION
    for g in BENCH["dress_tour"]["checks"]:
        for item in g["items"]:
            r += 1
            ws.cell(r, 1, g["group"])
            ws.cell(r, 2, item)
            for i in range(DRESS_SHOPS):
                style_input(ws.cell(r, 3 + i))
    r += 2
    ws.cell(r, 1, "점수표").font = F_SECTION
    header(ws, r + 1, ["평가 기준", "가중치(1~5)"] + [f"샵 {i} 점수" for i in range(1, DRESS_SHOPS + 1)])
    sfirst = r + 2
    for k, label in enumerate(BENCH["dress_tour"]["scores"]):
        ws.cell(sfirst + k, 1, label)
        for col in range(2, 3 + DRESS_SHOPS):
            style_input(ws.cell(sfirst + k, col))
    send = sfirst + len(BENCH["dress_tour"]["scores"]) - 1
    stot = send + 1
    ws.cell(stot, 1, "가중 총점").font = F_BOLD
    for i in range(DRESS_SHOPS):
        x = chr(ord("C") + i)
        ws[f"{x}{stot}"] = f"=SUMPRODUCT($B${sfirst}:$B${send},{x}{sfirst}:{x}{send})"
        ws[f"{x}{stot}"].font = F_BOLD
    ws.cell(stot + 1, 1, "1위").font = F_BOLD
    rng = f"C{stot}:{last_col}{stot}"
    ws.cell(stot + 1, 3, f'=IF(MAX({rng})=0,"",INDEX($C$5:${last_col}$5,MATCH(MAX({rng}),{rng},0)))')
    for i in range(DRESS_SHOPS):
        ws.column_dimensions[chr(ord("C") + i)].width = 24  # 투어 메모를 적을 자리
    ws.freeze_panes = "C5"


def build_payments(wb, d, first, last):
    ws = wb.create_sheet("결제기록")
    ws["A1"], ws["A1"].font = "결제기록", F_TITLE
    ws["A2"] = ("실제로 결제할 때마다 한 줄씩 적으세요. 항목은 지출계획의 항목명에서 고릅니다. "
                "계좌이체·현금 결제는 현금영수증 발급 여부를 적어 두세요.")
    ws["A2"].font = F_NOTE
    header(ws, 5, ["날짜", "항목", "구분", "금액", "결제자", "결제수단", "현금영수증", "확인", "메모"])
    names = {wb["지출계획"][f"B{r}"].value for r in range(first, last + 1)}
    payments = d.get("payments") or []
    if len(payments) > PAY_ROWS:
        die(f"payments 는 최대 {PAY_ROWS}건입니다")
    choices = {"type": PAY_TYPES, "payer": PAYERS, "method": METHODS, "receipt": RECEIPTS}
    for r in range(PAY_FIRST, PAY_LAST + 1):
        p = payments[r - PAY_FIRST] if r - PAY_FIRST < len(payments) else {}
        bad = set(p) - PAYMENT_FIELDS
        if bad:
            die(f"payments 에 알 수 없는 필드 {sorted(bad)}")
        if p and p.get("item") not in names:
            die(f"payments.item 은 지출계획 항목명이어야 합니다: {p.get('item')!r}")
        for k, allowed in choices.items():
            if p.get(k) not in (None, "") and p[k] not in allowed:
                die(f"payments.{k} 는 {allowed} 중 하나여야 합니다: {p[k]!r}")
        row = [parse_date(p.get("date"), "payments.date"), p.get("item"), p.get("type"),
               num(p.get("amount"), "payments.amount"), p.get("payer"), p.get("method"), p.get("receipt")]
        for i, v in enumerate(row, 1):
            style_input(ws.cell(r, i, v), "yyyy-mm-dd" if i == 1 else (WON if i == 4 else None))
        style_input(ws.cell(r, 9, p.get("memo")))
        ws[f"H{r}"] = (f'=IF(D{r}="","",IF(AND(OR(F{r}="계좌이체",F{r}="현금"),'
                       f'G{r}<>"발급완료",G{r}<>"해당없음"),"영수증 확인",""))')
        ws[f"H{r}"].font = F_BASE
    dv = DataValidation(type="list", formula1=f"지출계획!$B${first}:$B${last}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B{PAY_FIRST}:B{PAY_LAST}")
    for col, opts in zip("CEFG", (PAY_TYPES, PAYERS, METHODS, RECEIPTS)):
        validation(ws, ",".join(opts), f"{col}{PAY_FIRST}:{col}{PAY_LAST}")
    ws.conditional_formatting.add(f"H{PAY_FIRST}:H{PAY_LAST}",
                                  CellIsRule(operator="equal", formula=['"영수증 확인"'], fill=FILL_RED))

    amt, payer, method, receipt, check = (f"${c}${PAY_FIRST}:${c}${PAY_LAST}" for c in "DEFGH")
    ws["K4"], ws["K4"].font = "결제자별 합계", F_SECTION
    header(ws, 5, ["결제자"] + METHODS + ["합계"], col=11)
    for r, who in enumerate(PAYERS, 6):
        ws.cell(r, 11, who)
        for i, m in enumerate(METHODS):
            ws.cell(r, 12 + i, f'=SUMIFS({amt},{payer},$K{r},{method},"{m}")').number_format = WON
        ws.cell(r, 15, f"=SUMIFS({amt},{payer},$K{r})").number_format = WON
    tot = 6 + len(PAYERS)
    ws.cell(tot, 11, "합계").font = F_BOLD
    for col in "LMNO":
        ws[f"{col}{tot}"] = f"=SUM({col}6:{col}{tot - 1})"
        ws[f"{col}{tot}"].number_format, ws[f"{col}{tot}"].font = WON, F_BOLD
    ws.cell(tot + 1, 11, "결제자 미입력")
    ws.cell(tot + 1, 15, f"=SUM({amt})-O{tot}").number_format = WON

    ws["K15"], ws["K15"].font = "현금영수증", F_SECTION
    for r, (label, f, fmt) in enumerate([
        ("발급완료 금액", f'=SUMIFS({amt},{receipt},"발급완료")', WON),
        ("미발급 금액", f'=SUMIFS({amt},{receipt},"미발급")', WON),
        ("확인 필요 건수", f'=COUNTIF({check},"영수증 확인")', '0"건"'),
    ], 16):
        ws.cell(r, 11, label)
        ws.cell(r, 12, f).number_format = fmt
    ws["K19"] = "계좌이체·현금인데 발급완료/해당없음이 아니면 '영수증 확인'으로 표시됩니다."
    ws["K19"].font = F_NOTE
    ws.column_dimensions["J"].width = 3  # 결제 목록과 합계표 사이 간격
    ws.freeze_panes = f"A{PAY_FIRST}"


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
    ws.cell(r, 1, "출처").font = F_SECTION
    for k, v in BENCH["sources"].items():
        r += 1
        ws.cell(r, 1, f"{k}: {v}").font = F_NOTE
    r += 2
    ws.cell(r, 1, "참고(듀오 평균, 이 표 계산에는 안 씀)").font = F_SECTION
    for k, v in BENCH["reference_only"].items():
        r += 1
        ws.cell(r, 1, k)
        ws.cell(r, 3, v * MAN).number_format = WON


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
        ("납부 완료 (결제기록)", f"=SUM({PAY_AMT})", WON),
        ("남은 결제액", f"=SUM({p}$O${first}:$O${last})", WON),
        ("현금영수증 확인 필요", f'=COUNTIF({PAY}$H${PAY_FIRST}:$H${PAY_LAST},"영수증 확인")', "0\"건\""),
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
    ws["C18"] = '=IF(B18>0,"결제기록 시트의 \'영수증 확인\' 줄을 확인하세요","")'
    ws["C9"].font = ws["C15"].font = ws["C18"].font = F_NOTE
    ws["A19"] = "신랑·신부별 결제 합계와 결제수단별 합계는 결제기록 시트 오른쪽에 있습니다."
    ws["A19"].font = F_NOTE

    ws["A21"], ws["A21"].font = "대분류별", F_SECTION
    header(ws, 22, ["대분류", "예상 지출", "기준값 합계", "배율", "비중"])
    cats = list(dict.fromkeys(it["category"] for it in BENCH["items"]))
    for r, cat in enumerate(cats, 23):
        ws.cell(r, 1, cat)
        ws.cell(r, 2, f'=SUMIFS({g},{a},A{r})').number_format = WON
        ws.cell(r, 3, f'=SUMIFS({q},{a},A{r},{w},"Y",{g},">0")').number_format = WON
        ws.cell(r, 4, f'=IF(C{r}=0,"",SUMIFS({g},{a},A{r},{w},"Y",{g},">0")/C{r})').number_format = '0.00"배"'
        ws.cell(r, 5, f'=IF($B$7=0,"",B{r}/$B$7)').number_format = "0%"
    end = r + 1
    ws.cell(end, 1, "기타(추가 항목)")
    ws.cell(end, 2, f"=B7-SUM(B23:B{r})").number_format = WON
    ws.cell(end + 2, 1, "참고: 듀오 2026 평균 혼수 "
            f"{BENCH['reference_only']['혼수']:,}만원, 신혼집 {BENCH['reference_only']['신혼집']:,}만원 (이 표에는 포함 안 함)").font = F_NOTE


def text_width(v):
    """엑셀 열 너비 단위로 대략 환산. 한글 등 전각 문자는 2칸."""
    return sum(2 if ord(ch) >= 0x1100 else 1 for ch in str(v))


def cell_width(c):
    v = c.value
    if isinstance(v, str) and v.startswith("="):  # 수식은 결과 형식으로 추정
        return 14 if c.number_format == "#,##0" else 11
    if isinstance(v, (dt.date, dt.datetime)):
        return 11
    if isinstance(v, (int, float)):
        return len(f"{v:,.0f}") + 2
    return text_width(v)


def polish(wb, max_width=42):
    """가독성 마무리: 글꼴 통일, 눈금선 숨김, 표 테두리, 열 너비 자동, 긴 글 줄바꿈과 행 높이.

    표 셀 = 글자 크기 BODY 인 셀. 제목·섹션·설명(크기가 다름)은 테두리와 너비 계산에서 뺀다.
    """
    wb.calculation.fullCalcOnLoad = True  # 여는 앱이 수식을 다시 계산하게
    for ws in wb:
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 30
        widths, wrapped = {}, {}
        for row in ws.iter_rows():
            for c in row:
                filled = c.fill.fill_type == "solid"
                if c.value is None and not filled:
                    continue
                if c.font.name != FONT:  # 글꼴을 따로 안 준 셀
                    c.font = Font(name=FONT, size=BODY, bold=c.font.b)
                if c.font.sz != BODY:
                    continue
                c.border = BORDER
                if c.fill == FILL_HEAD:
                    ws.row_dimensions[c.row].height = 32
                    continue
                w = cell_width(c) if c.value is not None else 0
                if w > max_width:
                    wrapped[c.coordinate] = w
                    w = max_width
                widths[c.column_letter] = max(widths.get(c.column_letter, 0), w)
                c.alignment = Alignment(horizontal=c.alignment.horizontal, vertical="center")
                if not ws.row_dimensions[c.row].height:
                    ws.row_dimensions[c.row].height = 22  # 기본 행 높이를 무시하는 앱이 있어 직접 지정
        for col, w in widths.items():
            if col not in ws.column_dimensions:  # 코드에서 너비를 직접 정한 열(이미 존재)은 그대로 둔다
                ws.column_dimensions[col].width = max(w + 3, 8)
        for ref, w in wrapped.items():
            c = ws[ref]
            c.alignment = Alignment(horizontal=c.alignment.horizontal, vertical="center", wrap_text=True)
            lines = -(-w // max(ws.column_dimensions[c.column_letter].width - 2, 1))
            ws.row_dimensions[c.row].height = max(ws.row_dimensions[c.row].height or 20, 16 * lines + 6)


def build(d):
    for k in ("wedding_date", "region", "meal_price", "guaranteed_guests"):
        if d.get(k) in (None, ""):
            die(f"필수 값 {k} 가 없습니다")
    if d["region"] not in REGIONS:
        die(f"region 은 {REGIONS} 중 하나여야 합니다: {d['region']!r}")
    d["wedding_date"] = parse_date(d["wedding_date"], "wedding_date")
    num(d["meal_price"], "meal_price")
    num(d["guaranteed_guests"], "guaranteed_guests")
    shops = d.get("dress_shops") or []
    if not isinstance(shops, list) or len(shops) > DRESS_SHOPS or not all(isinstance(s, str) and s.strip() for s in shops):
        die(f"dress_shops 는 샵 이름(문자열) 목록이고 최대 {DRESS_SHOPS}곳입니다: {shops!r}")
    if any(s.startswith("=") or any(ch in s for ch in "*?~") for s in shops):  # 수식으로 저장되거나 MATCH 와일드카드
        die(f"dress_shops 이름은 = 로 시작하거나 *?~ 를 쓸 수 없습니다: {shops!r}")
    if len(set(shops)) != len(shops):
        die(f"dress_shops 이름이 겹칩니다: {shops!r}")
    if d.get("dress_mode") not in (None, "투어", "지정"):
        die(f"dress_mode 는 투어 또는 지정입니다: {d['dress_mode']!r}")
    if d.get("dress_mode") == "지정" and len(shops) > 1:
        die(f"dress_mode 가 지정이면 dress_shops 는 1곳만 넣습니다: {shops!r}")

    wb = Workbook()
    dash = wb.active
    dash.title = "대시보드"
    funds_ref = build_info(wb, d)
    meal_ref, gift_ref = build_guests(wb, d)
    first, last = build_plan(wb, d, meal_ref)
    build_dress_tour(wb, d, first, last)
    build_payments(wb, d, first, last)
    build_cashflow(wb, first, last, funds_ref)
    build_bench(wb)
    build_dashboard(dash, first, last, funds_ref, meal_ref, gift_ref)
    polish(wb)
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
