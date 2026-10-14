# k-wedding-plan

대한민국 결혼 준비용 Claude 스킬 모음. 대화로 정보를 모아 수식이 살아 있는 엑셀을 만들어 준다.

| 스킬 | 하는 일 | 상태 |
|---|---|---|
| [`wedding-budget`](skills/wedding-budget/SKILL.md) | 항목별 예산을 지역 평균(참가격·듀오)과 비교하고, 놓치기 쉬운 추가금, 보증인원과 식대, 축의금, 월별 계약금·잔금 흐름을 계산. 결제기록 장부(결제자·결제수단·현금영수증), 드레스 투어/지정 비교표(선택한 샵 금액을 지출계획에 연동) | 사용 가능 |
| [`wedding-schedule`](skills/wedding-schedule/SKILL.md) | 예식일·촬영일에서 거꾸로 계산한 할 일 70여 개의 시작·마감일과 지연 알림, 웨딩홀 투어 비교표, 당일 체크리스트 | 사용 가능 |

두 스킬 모두 대화형이다. Claude가 이름, 예식일 같은 필요한 정보를 하나씩 묻고, 답을 모아 엑셀을 만든다.

## 설치

**claude.ai**
1. `skills/wedding-budget`, `skills/wedding-schedule` 폴더를 각각 zip으로 압축한다.
2. 설정 → 기능(Capabilities)에서 코드 실행을 켜고, 스킬 업로드에 zip을 올린다.
3. 대화에서 "결혼 예산표 만들어줘" 또는 "결혼 준비 일정표 만들어줘"라고 말한다.

**Claude Code**
```bash
cp -r skills/wedding-budget skills/wedding-schedule ~/.claude/skills/
```
엑셀 생성에 `openpyxl`이 필요하다(`pip install openpyxl`).

## 기준 데이터

- 한국소비자원 참가격 결혼서비스 통계, 2026년 8월 기준 중간값(지역별)
- 듀오 「2026 결혼비용 실태 보고서」(예물·예단·이바지·신혼여행)
- 원문 정리와 출처: [`docs/research/wedding-cost-benchmarks-2026.md`](docs/research/wedding-cost-benchmarks-2026.md)
- 준비 일정과 시점: 웨딩 컨설팅사 체크리스트와 플래너 안내(2026)의 익명 요약. [`docs/research/wedding-prep-structure.md`](docs/research/wedding-prep-structure.md)
- 갱신할 때는 `skills/wedding-budget/references/benchmarks.json`, `skills/wedding-schedule/references/checklists.json`만 고친다.

## 개발

```bash
python3 tests/test_skills.py                 # 정적 검사
pip install formulas && python3 tests/test_skills.py   # 수식 계산 검증까지
```
