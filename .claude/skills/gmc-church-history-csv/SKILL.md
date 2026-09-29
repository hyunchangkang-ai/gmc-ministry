---
name: gmc-church-history-csv
description: Use when searching, adding, or correcting entries in the church history database (Administration/church_history_updated.csv, 50주년 역사벽·주보 기반 교회 연혁). Triggers on "교회 연혁", "역사벽", "연혁 추가", "언제 부임", "창립 이후 행사".
---

# 교회 연혁 CSV

파일: `Administration/church_history_updated.csv` (UTF-8 BOM, 약 2,400행, 1974~2026).

## 컬럼
`날짜, 연도, 월, 구분, 내용, 출처/비고`
- 날짜는 `YYYY-MM-DD`. 연도·월은 날짜에서 파생한 정수 컬럼이므로 날짜를 고치면 함께 고친다.
- `구분` 값: 기타행사, 선교, 성례, 회의/의회, 특별행사, 세미나, 창립, 훈련/수련회, 목회자 부임, 부흥회/집회, 목회자 사임, 은퇴 등. 새 항목은 **기존 값을 재사용**하고 새 값은 꼭 필요할 때만 만든다.
- `출처/비고`는 근거 문서(예: `주보 (2026-02-01)`, `50주년 역사벽`)를 적는다. 근거 없는 항목을 추가하지 않는다.

## 작업 규칙
- 파일을 열고 저장할 때 BOM과 한글 인코딩을 유지한다(엑셀 호환).
- 인명·직함·날짜는 원문 주보와 대조한다. 확인 못 한 것은 추측으로 채우지 않고 "확인 필요"로 표시한다.
- 추가 후 날짜순 정렬, 중복(같은 날짜·같은 내용) 여부를 확인한다.
- 조회 요청은 pandas나 grep으로 필터링하고, 결과에 출처를 함께 보여 준다.
- 기존 주보 원본은 `Administration/bulletin_downloads*`(HWP)에 있다. 2008~2015년분은 용량 때문에 `.gitignore`로 저장소에서 제외되어 있어 로컬/Drive에서 확보해야 한다.
