#!/usr/bin/env python3
"""input.json → 결혼 준비 일정 엑셀(수식 포함).

사용법: python3 build_schedule.py --input input.json --out schedule.xlsx
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DATA = json.loads((Path(__file__).resolve().parent.parent / "references" / "checklists.json").read_text("utf-8"))
TASKS = DATA["tasks"]
SHOOT_DEFAULT = 150  # 촬영일을 모를 때 예식 며칠 전으로 볼지

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
FILL_RED = PatternFill("solid", fgColor="F8CBAD")
FILL_ORANGE = PatternFill("solid", fgColor="FCE4D6")
FILL_YELLOW = PatternFill("solid", fgColor="FFF9DB")
DATE = "yyyy-mm-dd"
WRAP = Alignment(wrap_text=True, vertical="top")

STATUSES = "예정,진행,완료,해당없음"
OWNERS = "신랑,신부,공통,양가"
COLS = ["단계", "할 일", "담당", "기준", "시작(D-)", "마감(D-)", "시작일", "마감일", "마감월", "상태", "남은 일수",
        "알림", "놓치기 쉬움", "선택", "팁", "메모", "id"]
FIRST = 6  # 타임라인 첫 데이터 행


def die(msg):
    sys.exit(f"입력 오류: {msg}")


def parse_date(v, field):
    if v in (None, ""):
        return None
    try:
        return dt.date.fromisoformat(str(v))
    except ValueError:
        die(f"{field} 날짜 형식은 YYYY-MM-DD 여야 합니다: {v!r}")


def style_input(c, fmt=None):
    c.font, c.fill = F_INPUT, FILL_INPUT
    if fmt:
        c.number_format = fmt


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = F_NOTE


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
    title(ws, "기본정보", "노란 칸만 입력하세요. 촬영일을 바꾸면 촬영 기준 할 일의 날짜가 모두 다시 계산됩니다.")
    shoot = d["shoot_date"] or f"=B6-{SHOOT_DEFAULT}"
    rows = [
        ("신랑", d.get("groom") or "신랑", None, None),
        ("신부", d.get("bride") or "신부", None, None),
        ("예식일", d["wedding_date"], DATE, "추정" if d.get("wedding_date_estimated") else "확정"),
        ("스튜디오 촬영일", shoot, DATE, "확정" if d["shoot_date"] else f"추정(예식 {SHOOT_DEFAULT}일 전)"),
        ("마감 임박 기준(일)", 14, "0", None),
    ]
    for i, (label, value, fmt, state) in enumerate(rows, 4):
        ws.cell(i, 1, label).font = F_BOLD
        style_input(ws.cell(i, 2, value), fmt)
        if state:
            style_input(ws.cell(i, 3, state))
    ws["D6"] = "날짜가 정해지면 '확정'으로 바꾸세요"
    ws["D7"] = "촬영일이 정해지면 날짜를 넣고 '확정'으로 바꾸세요"
    ws["D8"] = "마감까지 이 일수 이하로 남으면 '마감 임박'"
    for r in (6, 7, 8):
        ws.cell(r, 4).font = F_NOTE
    validation(ws, "확정,추정", "C6:C7")


def build_timeline(wb, d):
    ws = wb.create_sheet("타임라인")
    title(ws, "준비 타임라인",
          "기준일에서 시작(D-)·마감(D-) 일수를 빼서 날짜를 계산합니다. 음수는 기준일 이후입니다. "
          "D- 일수와 상태는 직접 고칠 수 있습니다. 순서가 바뀌면 마감일 열로 정렬하세요.")
    ws["A3"] = "알림: 지연 = 마감 지남 / 마감 임박 / 지금 할 때 = 시작일이 지남 / 예정. ★ = 놓치기 쉬운 일"
    ws["A3"].font = F_NOTE
    header(ws, 5, COLS)

    shoot = d["shoot_date"] or d["wedding_date"] - dt.timedelta(days=SHOOT_DEFAULT)
    base = {"wedding": d["wedding_date"], "shoot": shoot}
    rows = sorted(TASKS, key=lambda t: (base[t["anchor"]] - dt.timedelta(days=t["due"]),
                                        base[t["anchor"]] - dt.timedelta(days=t["start"])))
    for r, t in enumerate(rows, FIRST):
        skipped = t["id"] in d["skip_ids"] or set(t.get("tags", [])) & d["skip_tags"]
        status = "완료" if t["id"] in d["done"] else ("해당없음" if skipped else "예정")
        anchor = "촬영일" if t["anchor"] == "shoot" else "예식일"
        values = [t["phase"], t["title"], t["owner"], anchor, t["start"], t["due"]]
        for i, v in enumerate(values, 1):
            ws.cell(r, i, v).font = F_BASE
        for col in "CEFJ":
            style_input(ws[f"{col}{r}"])
        ws[f"J{r}"] = status
        base_ref = f'IF($D{r}="촬영일",기본정보!$B$7,기본정보!$B$6)'
        ws[f"G{r}"] = f"={base_ref}-E{r}"
        ws[f"H{r}"] = f"={base_ref}-F{r}"
        ws[f"I{r}"] = f"=DATE(YEAR(H{r}),MONTH(H{r}),1)"
        ws[f"K{r}"] = f'=IF(OR(J{r}="완료",J{r}="해당없음"),"",H{r}-TODAY())'
        ws[f"L{r}"] = (f'=IF(K{r}="","",IF(K{r}<0,"지연",IF(K{r}<=기본정보!$B$8,"마감 임박",'
                       f'IF(TODAY()>=G{r},"지금 할 때","예정"))))')
        ws[f"M{r}"] = "★" if t.get("hidden") else ""
        ws[f"N{r}"] = "선택" if t.get("optional") else ""
        ws[f"O{r}"] = t.get("tip", "")
        ws[f"Q{r}"] = t["id"]
        for col in "GH":
            ws[f"{col}{r}"].number_format = DATE
        ws[f"I{r}"].number_format = "yyyy-mm"
        ws[f"K{r}"].number_format = '0"일";-0"일"'
        for col in "GHIKLMNOPQ":
            ws[f"{col}{r}"].font = F_BASE
        ws[f"O{r}"].alignment = WRAP
        style_input(ws[f"P{r}"])
    last = FIRST + len(rows) - 1

    validation(ws, STATUSES, f"J{FIRST}:J{last}")
    validation(ws, OWNERS, f"C{FIRST}:C{last}")
    validation(ws, "예식일,촬영일", f"D{FIRST}:D{last}")
    rng = f"L{FIRST}:L{last}"
    for text, fill in (("지연", FILL_RED), ("마감 임박", FILL_ORANGE), ("지금 할 때", FILL_YELLOW)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill))
    ws.auto_filter.ref = f"A5:Q{last}"
    ws.freeze_panes = f"C{FIRST}"
    ws.column_dimensions["P"].width = 24  # 메모 칸은 비어 있어도 쓸 자리를 둔다
    ws.column_dimensions["Q"].hidden = True
    return last


def build_overview(ws, last):
    title(ws, "결혼 준비 일정 개요", "시점은 업계 관행 기준이며 업체·계약마다 다릅니다. 계약서 날짜가 우선입니다.")
    t = "타임라인!"
    a, j, l = (f"{t}${c}${FIRST}:${c}${last}" for c in "AJL")
    rows = [
        ("신랑 · 신부", '=기본정보!B4&" · "&기본정보!B5', None),
        ("예식일", "=기본정보!B6", DATE),
        ("스튜디오 촬영일", "=기본정보!B7", DATE),
        ("오늘", "=TODAY()", DATE),
        ("D-day", "=B5-B7", '"D-"0;"D+"0;"D-day"'),
        ("할 일 (해당없음 제외)", f'=COUNTA({t}$B${FIRST}:$B${last})-COUNTIF({j},"해당없음")', '0"개"'),
        ("완료", f'=COUNTIF({j},"완료")', '0"개"'),
        ("진행률", "=IF(B9=0,0,B10/B9)", "0%"),
        ("지연", f'=COUNTIF({l},"지연")', '0"개"'),
        ("마감 임박", f'=COUNTIF({l},"마감 임박")', '0"개"'),
        ("지금 할 때", f'=COUNTIF({l},"지금 할 때")', '0"개"'),
    ]
    for i, (label, f, fmt) in enumerate(rows, 4):
        ws.cell(i, 1, label).font = F_BOLD
        c = ws.cell(i, 2, f)
        c.font = F_BOLD
        if fmt:
            c.number_format = fmt
    ws["C5"] = "=기본정보!C6"
    ws["C6"] = "=기본정보!C7"
    ws["C12"] = '=IF(B12>0,"타임라인에서 알림 열을 \'지연\'으로 필터하세요","")'
    for ref in ("C5", "C6", "C12"):
        ws[ref].font = F_NOTE
    ws.conditional_formatting.add("B12", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_RED))

    ws["A17"] = "단계별 진행"
    ws["A17"].font = F_SECTION
    header(ws, 18, ["단계", "할 일", "완료", "해당없음", "진행률", "지연"])
    phases = list(dict.fromkeys(x["phase"] for x in TASKS))
    for r, p in enumerate(phases, 19):
        ws.cell(r, 1, p)
        ws.cell(r, 2, f"=COUNTIF({a},A{r})")
        ws.cell(r, 3, f'=COUNTIFS({a},A{r},{j},"완료")')
        ws.cell(r, 4, f'=COUNTIFS({a},A{r},{j},"해당없음")')
        ws.cell(r, 5, f"=IF(B{r}-D{r}=0,\"\",C{r}/(B{r}-D{r}))").number_format = "0%"
        ws.cell(r, 6, f'=COUNTIFS({a},A{r},{l},"지연")')


def build_hall_tour(wb, halls):
    ws = wb.create_sheet("웨딩홀투어")
    title(ws, "웨딩홀 투어 비교", "홀 이름을 넣고 투어하면서 칸을 채우세요. 아래 점수표에 가중치와 점수를 넣으면 총점이 나옵니다.")
    header(ws, 4, ["분류", "확인할 것", "홀 1", "홀 2", "홀 3", "홀 4"])
    ws["B5"] = "홀 이름"
    ws["B5"].font = F_BOLD
    for i in range(4):
        style_input(ws.cell(5, 3 + i, halls[i] if i < len(halls) else None))
    r = 6
    for g in DATA["hall_tour"]:
        for item in g["items"]:
            ws.cell(r, 1, g["group"]).font = F_BASE
            ws.cell(r, 2, item).font = F_BASE
            for i in range(4):
                style_input(ws.cell(r, 3 + i))
            r += 1
    r += 1
    ws.cell(r, 1, "점수표").font = F_SECTION
    header(ws, r + 1, ["평가 기준", "가중치(1~5)", "홀 1 점수", "홀 2 점수", "홀 3 점수", "홀 4 점수"])
    first = r + 2
    for k, name in enumerate(DATA["hall_scores"]):
        ws.cell(first + k, 1, name).font = F_BASE
        for col in range(2, 7):
            style_input(ws.cell(first + k, col))
    end = first + len(DATA["hall_scores"]) - 1
    tot = end + 1
    ws.cell(tot, 1, "가중 총점").font = F_BOLD
    for col in "CDEF":
        ws[f"{col}{tot}"] = f"=SUMPRODUCT($B${first}:$B${end},{col}{first}:{col}{end})"
        ws[f"{col}{tot}"].font = F_BOLD
    ws.cell(tot + 1, 1, "1위").font = F_BOLD
    best = f"MATCH(MAX(C{tot}:F{tot}),C{tot}:F{tot},0)"
    ws.cell(tot + 1, 3, f'=IF(MAX(C{tot}:F{tot})=0,"",IF(INDEX(C5:F5,{best})="","홀 "&{best},INDEX(C5:F5,{best})))')
    ws.cell(tot + 2, 1, "가중치는 타임라인의 '우선순위 정하기'에서 정한 순서대로 주면 됩니다.").font = F_NOTE
    for col in "CDEF":
        ws.column_dimensions[col].width = 24  # 투어 메모를 적을 자리
    ws.freeze_panes = "C6"


def build_day_of(wb):
    ws = wb.create_sheet("당일체크")
    title(ws, "예식 당일 체크", "준비물, 역할, 현금 봉투를 미리 정해 두세요.")
    day = DATA["day_of"]
    r = 4
    sections = [("준비물", ["항목", "챙길 사람", "완료"], day["items"]),
                ("역할", ["역할", "맡은 사람", "확인"], day["roles"])]
    for name, cols, items in sections:
        ws.cell(r, 1, name).font = F_SECTION
        header(ws, r + 1, cols)
        for k, item in enumerate(items, r + 2):
            ws.cell(k, 1, item).font = F_BASE
            style_input(ws.cell(k, 2))
            style_input(ws.cell(k, 3))
        validation(ws, "완료", f"C{r + 2}:C{r + 1 + len(items)}")
        r += len(items) + 3
    ws.cell(r, 1, "현금 봉투").font = F_SECTION
    header(ws, r + 1, ["대상", "받는 분", "금액", "봉투 준비"])
    first = r + 2
    for k, item in enumerate(day["cash"], first):
        ws.cell(k, 1, item).font = F_BASE
        style_input(ws.cell(k, 2))
        style_input(ws.cell(k, 3), "#,##0")
        style_input(ws.cell(k, 4))
    end = first + len(day["cash"]) - 1
    validation(ws, "완료", f"D{first}:D{end}")
    ws.cell(end + 1, 1, "합계").font = F_BOLD
    ws.cell(end + 1, 3, f"=SUM(C{first}:C{end})").number_format = "#,##0"
    ws.cell(end + 1, 3).font = F_BOLD
    for col in "BC":
        ws.column_dimensions[col].width = 18  # 이름·금액을 적을 자리


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
            preset = col in ws.column_dimensions  # 코드에서 직접 너비를 정한 열만 이미 존재한다
            cd = ws.column_dimensions[col]
            if not cd.hidden:  # 직접 정한 너비가 있으면 그보다 좁히지 않음
                cd.width = max(w + 3, 8, cd.width if preset else 0)
        for ref, w in wrapped.items():
            c = ws[ref]
            c.alignment = Alignment(horizontal=c.alignment.horizontal, vertical="center", wrap_text=True)
            lines = -(-w // (max_width - 2))
            ws.row_dimensions[c.row].height = max(ws.row_dimensions[c.row].height or 20, 16 * lines + 6)


def normalize(d):
    if not d.get("wedding_date"):
        die("필수 값 wedding_date 가 없습니다. 예식일이 미정이면 목표 월의 15일을 넣고 wedding_date_estimated 를 true 로 두세요")
    d["wedding_date"] = parse_date(d["wedding_date"], "wedding_date")
    d["shoot_date"] = parse_date(d.get("shoot_date"), "shoot_date")
    if d["shoot_date"] and d["shoot_date"] >= d["wedding_date"]:
        die("shoot_date 는 wedding_date 보다 앞이어야 합니다")
    ids = {t["id"] for t in TASKS}
    tags = {tag for t in TASKS for tag in t.get("tags", [])}
    skip = set(d.get("skip") or [])
    unknown = skip - ids - tags
    if unknown:
        die(f"skip 에 알 수 없는 값 {sorted(unknown)}. 태그는 {sorted(tags)} 또는 할 일 id 를 쓰세요")
    d["skip_tags"], d["skip_ids"] = skip & tags, skip & ids
    d["done"] = set(d.get("done") or [])
    if d["done"] - ids:
        die(f"done 에 알 수 없는 할 일 id {sorted(d['done'] - ids)}")
    halls = d.get("halls") or []
    if len(halls) > 4:
        die("halls 는 최대 4곳입니다")
    d["halls"] = halls
    return d


def build(d):
    d = normalize(d)
    wb = Workbook()
    overview = wb.active
    overview.title = "개요"
    build_info(wb, d)
    last = build_timeline(wb, d)
    build_hall_tour(wb, d["halls"])
    build_day_of(wb)
    build_overview(overview, last)
    polish(wb)
    return wb


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    build(json.loads(Path(a.input).read_text("utf-8"))).save(a.out)
    print(f"생성 완료: {a.out}")


if __name__ == "__main__":
    main()
