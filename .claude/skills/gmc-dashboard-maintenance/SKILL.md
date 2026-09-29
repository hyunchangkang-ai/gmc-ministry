---
name: gmc-dashboard-maintenance
description: Use when modifying, debugging, or deploying the GMC 사역 관리 dashboard (ministry-dashboard.html, annual-dashboard.html, index.html, sw.js, Administration/gmc_backend.gs). Triggers on "대시보드 수정", "사역 관리", "PWA 버전", "Apps Script 재배포", "캘린더 동기화 오류", "중복 사역", "주차 데이터".
---

# GMC 사역 대시보드 유지보수

## 구성
| 구성요소 | 파일 | 역할 |
|---|---|---|
| 주간 대시보드(메인) | `ministry-dashboard.html` | 주간 사역·할 일 관리. PWA 시작 페이지 |
| 연간 대시보드 | `annual-dashboard.html` | 연간 사역 계획 |
| 허브 페이지 | `index.html` | 도구 모음/안내 |
| 구버전 | `index_old.html` | 참고용. 수정 대상 아님 |
| 서비스워커 | `sw.js` | 캐시 관리 (`CACHE = 'gmc-vNN'`) |
| 백엔드 | `Administration/gmc_backend.gs` | Google Apps Script 웹 앱 + Google Sheets DB |

호스팅은 GitHub Pages(`/gmc-ministry/` 경로)입니다. `manifest.json`, `sw.js`에 이 경로가 하드코딩되어 있어, 저장소 이름이나 소유자가 바뀌면 Pages 주소가 바뀌므로 함께 수정해야 합니다.

## 데이터 모델 (gmc_backend.gs)
- 시트 ID는 `SHEET_ID` 상수. 탭: `사역데이터`(주차별 JSON), `이번주할일목록`, `사역데이터베이스`(누적 DB, 컬럼은 `DB_HEADERS`).
- **주차 키는 "그 주 화요일" 날짜(YYYY-MM-DD)**. 날짜 → 주차 변환은 `getTuesdayForDate()`. 일·월요일은 직전 화요일 주에 속한다.
- 날짜 문자열은 시트에서 `MM-DD-YYYY`, 앱 내부에서는 `YYYY-MM-DD`. Date 객체가 섞여 들어올 수 있어 정규화 로직(`toMDY`, `parseMDY`)을 유지한다.
- API: `doGet` (`?action=prev_week_report`, `?action=gcal_events`, 기본=전체 데이터), `doPost` (`?action=gcal_create|gcal_update|gcal_delete`, 기본=전체 데이터 저장).
- 쓰기 가능한 캘린더는 `WRITABLE_CALENDARS` 허용 목록으로 제한한다. 캘린더를 추가·제거할 때 여기를 수정한다.
- 분류 자동 감지는 대시보드 `detectCategory`와 백엔드 `detectCategoryGAS`가 **같은 로직**이어야 한다. 한쪽만 고치면 분류가 어긋난다.

## 배포 절차
1. 프론트엔드: HTML 수정 → 커밋 → `main` 푸시 → GitHub Pages 자동 반영.
2. 캐시 갱신이 필요한 변경이면 `sw.js`의 `CACHE` 번호와 `ministry-dashboard.html` 부제목의 버전 라벨을 **함께** 올린다(과거 커밋 관례). 현재 `sw.js`는 v21, 라벨은 v20으로 어긋나 있으니 확인 후 맞춘다.
3. 백엔드(`.gs`): Apps Script 편집기에 붙여넣기 → **배포 관리 → 기존 배포 편집 → 새 버전**으로 재배포. 새 배포를 만들면 웹 앱 URL이 바뀐다.
4. 웹 앱 URL은 `ministry-dashboard.html`(`APPS_SCRIPT_URL`)과 `index_old.html`에 하드코딩되어 있다. URL이 바뀌면 모두 교체한다.

## 알려진 함정 (과거 버그 이력)
- 초기 pull이 끝나기 전에 saveData가 Apps Script로 push하면 서버 데이터를 덮어쓴다 → pull 완료 전 push 금지.
- 주차 키가 Date 객체로 들어오는 경우가 있어 정규화 필요(v19).
- 같은 사역·캘린더가 gcalId 차이로 중복 등록되므로 이름 기준 병합(`mergeDuplicateCalMinistries`) 로직이 있다. 완료된 할 일은 삭제 보호한다.
- 사역자 스케줄 표기 흔들림(맞춤법 정규화)이 있다(v20).

## 소유권 주의
Apps Script 프로젝트, 시트, 트리거는 만든 사람의 Google 계정 소유다. 담당자가 바뀌면 시트 공유 → Apps Script 프로젝트 소유권/편집 권한 이전 → 트리거 재생성 → 웹 앱 재배포 순으로 옮긴다.
