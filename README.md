# k-wedding

대한민국 결혼 준비용 Claude 스킬 모음. 대화로 정보를 모아 수식이 살아 있는 엑셀을 만들어 준다.

| 스킬 | 하는 일 | 상태 |
|---|---|---|
| [`wedding-budget`](skills/wedding-budget/SKILL.md) | 항목별 예산을 지역 평균(참가격·듀오)과 비교하고, 놓치기 쉬운 추가금, 보증인원과 식대, 축의금, 월별 계약금·잔금 흐름을 계산 | 사용 가능 |
| `wedding-schedule` | 예식일에서 거꾸로 계산한 준비 일정표 | 준비 중 |

## 설치

**claude.ai**
1. `skills/wedding-budget` 폴더를 zip으로 압축한다.
2. 설정 → 기능(Capabilities)에서 코드 실행을 켜고, 스킬 업로드에 zip을 올린다.
3. 대화에서 "결혼 예산표 만들어줘"라고 말한다.

**Claude Code**
```bash
cp -r skills/wedding-budget ~/.claude/skills/
```
엑셀 생성에 `openpyxl`이 필요하다(`pip install openpyxl`).

## 기준 데이터

- 한국소비자원 참가격 결혼서비스 통계, 2026년 8월 기준 중간값(지역별)
- 듀오 「2026 결혼비용 실태 보고서」(예물·예단·이바지·신혼여행)
- 원문 정리와 출처: [`docs/research/wedding-cost-benchmarks-2026.md`](docs/research/wedding-cost-benchmarks-2026.md)
- 갱신할 때는 `skills/wedding-budget/references/benchmarks.json`만 고친다.

## 개발

```bash
python3 tests/test_skills.py                 # 정적 검사
pip install formulas && python3 tests/test_skills.py   # 수식 계산 검증까지
```
