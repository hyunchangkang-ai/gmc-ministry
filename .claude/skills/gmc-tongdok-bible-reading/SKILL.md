---
name: gmc-tongdok-bible-reading
description: Use for the 성경 통독 (Bible reading) tools — tongdok.html 365-day tracker and the daily KakaoTalk reading message sender (통독.bat, send_bible_message.py). Triggers on "통독", "성경 읽기 계획", "통독 카톡 발송", "통독 스케줄러".
---

# 성경 통독 도구

## 1. 통독 트래커 (`tongdok.html`)
- 창세기~요한계시록 **1,189장을 365일에 고르게 배분**한 순차 읽기표.
- 기능: 오늘 본문 표시, 일별 체크(브라우저 localStorage 저장), 연간 진행률, 연속 읽기(streak), 월별 달력, 대시보드로 돌아가기 링크.
- **현재 `main`에 없다.** 브랜치 `claude/tongdok-0gckq`의 커밋 `12f7fec`에만 있다(`ministry-dashboard.html`의 "성경 통독" 버튼 포함). 반영하려면 그 브랜치를 `main`에 병합한다.
- 체크 기록은 각 사용자 기기의 localStorage에만 있다. 기기를 바꾸면 이어지지 않는다.
- 다음 해에 쓰려면 제목의 연도와 시작일 계산부를 수정한다.

## 2. 카카오톡 매일 발송 (Windows PC)
- `Administration/통독.bat`: `python send_bible_message.py --now`로 오늘의 통독 메시지를 즉시 발송.
- `Administration/setup_scheduler.bat`: 관리자 권한으로 `setup_scheduler.py`를 실행해 Windows 작업 스케줄러에 등록.
- 저장소의 `debug_kakao*.png`, `step*_*.png`는 카카오톡 PC 앱 화면 자동화(채팅 탭 → 검색 → 붙여넣기)를 디버깅한 캡처다. UI 자동화 방식으로 추정된다.
- **`send_bible_message.py`, `setup_scheduler.py`는 저장소에 없다**(`.gitignore`가 `*.py`를 제외). 원본은 사용자의 PC에 있을 가능성이 높다. 인수인계 전에 이 파일들을 확보해야 한다. 확인 전까지 발송 로직 세부는 이 스킬이 보증하지 않는다.
- 카카오톡 PC 앱 UI가 바뀌면 자동화가 깨질 수 있다. 화면 캡처와 비교해 점검한다.
