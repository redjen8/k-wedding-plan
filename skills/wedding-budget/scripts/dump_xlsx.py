#!/usr/bin/env python3
"""사용자가 첨부한 엑셀의 값이 있는 셀을 시트별로 출력한다. 항목 매핑은 Claude가 한다.

사용법: python3 dump_xlsx.py 파일.xlsx [최대행=300]
"""
import sys

from openpyxl import load_workbook

path, max_row = sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 300
wb = load_workbook(path, data_only=True, read_only=True)
for ws in wb.worksheets:
    print(f"\n=== {ws.title}")
    for row in ws.iter_rows(max_row=max_row):
        cells = [f"{c.coordinate}={c.value!r}" for c in row if c.value not in (None, "")]
        if cells:
            print(" | ".join(cells))
