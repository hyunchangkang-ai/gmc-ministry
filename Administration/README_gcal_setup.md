# Google Calendar 동기화 설정 가이드

## 1단계 — Python 패키지 설치

터미널에서:

```bash
pip3 install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

## 2단계 — Google Cloud 자격증명 만들기

1. [Google Cloud Console](https://console.cloud.google.com/) 접속
2. 새 프로젝트 생성 (예: `GMC Dashboard`)
3. **API 및 서비스 → 라이브러리** → "Google Calendar API" 검색 후 **사용 설정**
4. **API 및 서비스 → 사용자 인증 정보** → **+ 사용자 인증 정보 만들기 → OAuth 클라이언트 ID**
5. 애플리케이션 유형: **데스크톱 앱** 선택
6. 생성 완료 후 **JSON 다운로드** 클릭
7. 내려받은 파일을 아래 경로로 이동/이름 변경:
   ```
   /Users/hyunchangkang/Claude/gcal_credentials.json
   ```

## 3단계 — 최초 인증 실행 (1회만)

터미널에서:

```bash
cd /Users/hyunchangkang/Claude
python3 gcal_sync.py
```

브라우저가 열리면 Google 계정으로 로그인 → 권한 허용.
완료 후 `gcal_token.json`이 자동 생성됩니다.

## 4단계 — 자동 실행 등록

```bash
launchctl load ~/Library/LaunchAgents/com.gmc.gcal-sync.plist
```

이후 Mac을 켜둔 상태로 자정이 되면 자동으로 실행됩니다.

## 특정 캘린더만 가져오려면

`gcal_sync.py` 상단의 `TARGET_CALENDARS` 변수를 수정:

```python
# 예: 기본 캘린더 + 팀 캘린더만
TARGET_CALENDARS = ["primary", "팀캘린더ID@group.calendar.google.com"]
```

캘린더 ID는 Google Calendar → 설정 → 캘린더 선택 → "캘린더 통합" 섹션에서 확인.

## 수동으로 지금 바로 실행하려면

```bash
python3 /Users/hyunchangkang/Claude/gcal_sync.py
```

## 로그 확인

- 정상 로그: `/Users/hyunchangkang/Claude/gcal_sync.log`
- 오류 로그: `/Users/hyunchangkang/Claude/gcal_sync_error.log`

## 자동 실행 중지하려면

```bash
launchctl unload ~/Library/LaunchAgents/com.gmc.gcal-sync.plist
```
