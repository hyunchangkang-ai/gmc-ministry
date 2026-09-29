---
name: gmc-gmail-task-sync
description: Use for the Gmail → Gemini → 사역 대시보드 할 일 자동 추출 (checkGmailAndExtractTasks in Administration/gmc_backend.gs). Triggers on "Gmail 동기화", "메일에서 할 일 추출", "Gemini API 키", "gmc-processed", "매일 7시 트리거".
---

# Gmail 할 일 자동 추출

`Administration/gmc_backend.gs`의 `checkGmailAndExtractTasks()`가 매일 오전 7시(7~8시 사이)에 실행된다.

## 흐름
1. Script Properties의 `GEMINI_API_KEY` 확인. 없으면 로그만 남기고 종료한다.
2. Gmail 검색 `in:inbox -label:gmc-processed newer_than:7d`, 최대 20 스레드.
3. 메일마다 `callGeminiAPI()`가 담당 사역자가 해야 할 일을 JSON 배열(`text`, `date`, `time`, `category`)로 추출한다.
4. `getTuesdayForDate()`로 주차(화요일 키)를 정하고 시트 `사역데이터`에 추가한다.
5. 처리한 스레드에 `gmc-processed` 라벨을 붙여 재처리를 막는다.

## 후임자가 반드시 바꿀 것
- **프롬프트가 특정 목사(Hyunchang Kang) 개인 비서로 하드코딩**되어 있다. 후임 담당자에 맞게 `systemPrompt`의 이름과 역할을 수정한다.
- 스크립트는 **실행하는 사람의 Gmail**을 읽는다. 담당자 교체 시 새 담당자 계정으로 프로젝트 소유/트리거를 다시 만든다.
- `category` 키 목록은 `CATEGORY_MAP`과 대시보드 분류와 일치해야 한다.
- 모델 이름이 `gemini-1.5-flash`로 고정되어 있다. 호출이 404/폐기 오류를 내면 현재 제공 중인 모델로 교체한다.

## 설정 절차
1. Google AI Studio에서 API 키 발급 → Apps Script → 프로젝트 설정 → 스크립트 속성에 `GEMINI_API_KEY` 등록. **코드나 저장소에 키를 적지 않는다.**
2. 편집기에서 `setupGmailSyncTrigger()`를 한 번 실행(기존 동일 트리거는 자동 정리).
3. 확인은 `checkGmailAndExtractTasks`를 수동 실행하고 실행 로그(`[Gmail Sync] ...`)를 본다.

## 주의
- 메일 본문이 외부 API(Gemini)로 전송된다. 목회 상담·교인 개인정보가 담긴 메일이 자동 처리되지 않도록 검색 조건이나 라벨로 제외하는 것을 고려한다.
