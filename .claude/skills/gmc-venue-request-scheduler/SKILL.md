---
name: gmc-venue-request-scheduler
description: Use for the church venue/room usage request automation (교회 장소 사용 신청) — Google Form responses → conflict check → Google Calendar registration → approval/conflict email. Triggers on "장소 사용 신청", "교실 예약", "일정 충돌", "장소 신청 스크립트", "church_venue_scheduler".
---

# 교회 장소 사용 신청 자동화

원본 코드: `Administration/church_venue_scheduler.js` (Google Apps Script).

## 동작
Google Form 응답 시트에 연결된 스크립트가 신청을 **타임스탬프 오름차순(선착순)** 으로 처리한다.
1. 처리 상태(N열)가 비어 있는 행만 대상.
2. 같은 신청자(양쪽 이메일이 있으면 이메일, 없으면 공백 제거한 이름으로 판정, `isSameApplicant`)·겹치는 일시·같은 교실이면 **중복 신청**, 다른 신청자가 같은 교실·겹치는 시간이면 **일정 충돌**로 표시한다.
3. 통과하면 장소 캘린더(`CALENDAR_ID`)에 일정을 만들고 **등록 완료** 처리한다.
4. 신청자 이메일로 **모든 결과(등록 완료·일정 충돌·중복 신청·오류 발생)** 에 대해 사유가 담긴 안내 메일을 보낸다. 발송은 `sendNoticeEmail` 한 곳으로 통일했고, 표시 이름 `GMC 행정실`, 참조(CC)·회신(replyTo) 모두 `office@gmcusa.org`. 신청자 이메일이 비었거나 형식이 틀리면 행정실로 '수동 안내 필요' 메일이 간다.

상태 값: `등록 완료`, `중복 신청`, `일정 충돌`, `오류 발생`.

## 시트 열 매핑 (0부터)
A 타임스탬프, B 신청부서, C 신청팀, D 팀장, E 신청자, F 연락처, G 사용 일자·요일·시간(자유 텍스트), H 교실 번호, I 사용 인원, J 최종 책임자, K 사용 목적, L 이메일, M 필요 장비, N 처리 상태(스크립트가 헤더 자동 생성).
설문 문항 순서를 바꾸면 스크립트 상단의 `COL_*` 상수를 같이 수정해야 한다.

## 핵심 로직
- G열은 자유 텍스트라서 `parseDatesFromText`, `parseTimeFromText`가 한글/영문 날짜·시간 표현을 해석한다. 파싱이 틀리면 충돌 판정도 틀리므로, 새 표현이 들어오면 이 함수를 먼저 점검한다.
- 충돌 판정은 `getDateTimeSlots` → `doSlotsConflict`로 시간 구간 겹침을 비교한다.

## 설치/인수
1. 설문 응답 스프레드시트 → 확장 프로그램 → Apps Script에 코드 붙여넣기.
2. `CALENDAR_ID`를 실제 장소 캘린더 ID로 교체(캘린더 설정 → 캘린더 통합).
3. 트리거 추가: `onFormSubmit`, 이벤트 소스 "스프레드시트", 이벤트 유형 "양식 제출 시".
4. 메일은 **스크립트 소유자 계정**으로 발송된다(MailApp 특성). 발신자를 `office@gmcusa.org`로 하려면 스크립트·스프레드시트·트리거를 office 계정이 소유해야 한다. 소유자 교체 시 기존 트리거는 새 소유자가 다시 만들어야 하고, 첫 실행 때 메일·캘린더 권한 승인이 필요하다. 캘린더(`CALENDAR_ID`)에는 office 계정의 편집 권한이 있어야 한다.
5. 수정 후에는 실제 신청이 아닌 테스트 행 2건(중복 1, 충돌 1)으로 상태 값을 확인한다. 테스트 후 캘린더 일정도 삭제한다.
