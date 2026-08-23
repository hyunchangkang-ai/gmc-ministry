// GMC 사역 관리 - Google Apps Script 백엔드
// 배포: 웹 앱으로 배포 (나만 실행, 누구나 액세스)
// 변경 후 반드시 새 버전으로 재배포 필요

const SHEET_ID = '140sfyyn6_NSKGCgHFn11kDZV0jqesPtVL6rZvcSR2Kk';

// ── DB 컬럼 정의 ────────────────────────────────────────────────────────────
const DB_HEADERS = [
  'ID', '주차', '사역명', '사역카테고리', '날짜', '시간',
  '사역내용', '완료', '분류', '연간사역', '참고사항', '최초등록일', '최종수정일'
];
// 인덱스 상수 (0-based)
const COL = {
  ID: 0, WEEK: 1, MINISTRY: 2, M_CAT: 3, DATE: 4, TIME: 5,
  CONTENT: 6, DONE: 7, CATEGORY: 8, ANNUAL: 9, NOTE: 10,
  CREATED: 11, UPDATED: 12
};

// ── 유틸: 날짜값 → MM-DD-YYYY 문자열 ────────────────────────────────────────
// 입력: Date 객체, YYYY-MM-DD 문자열, 이미 MM-DD-YYYY 문자열 모두 처리
function toMDY(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${m}-${d}-${val.getFullYear()}`;
  }
  const s = String(val).trim();
  // YYYY-MM-DD → MM-DD-YYYY
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[2]}-${isoMatch[3]}-${isoMatch[1]}`;
  // 이미 MM-DD-YYYY 형식이면 그대로
  return s;
}

// ── 유틸: MM-DD-YYYY 문자열 → Date 객체 ──────────────────────────────────────
function parseMDY(str) {
  if (str instanceof Date) return new Date(str.getFullYear(), str.getMonth(), str.getDate(), 12, 0, 0);
  const parts = String(str || '').split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]), 12, 0, 0);
  }
  return null;
}

const CATEGORY_MAP = {
  worship: '예배',
  sermon: '설교',
  disciple: '제자훈련',
  cell: '목장',
  mission: '선교',
  newmember: '새가족',
  meeting: '회의',
  admin: '행정',
  pastoral: '심방/목양',
  evangelism: '전도폭발',
  fellowship: '식사/교제',
  facility: '시설',
  shortmission: '단기선교',
  intercession: '중보기도',
  other: '기타'
};

function getCategoryLabel(key) {
  return CATEGORY_MAP[key] || key || '';
}

// ── 유틸: 텍스트에서 분류 자동 감지 (대시보드 detectCategory 동일 로직) ────
function detectCategoryGAS(text, ministryName) {
  if (!text) return ministryName || '';
  const t = text.toLowerCase();
  
  // 1. 설교 우선
  if (/설교|sermon|메시지|강해/.test(t)) return '설교';
  // 2. 목장/목자
  if (/목자|목장|셀/.test(t)) return '목장';
  // 3. 전도/제자훈련
  if (/전도|evangelism|전도폭발|ee/.test(t)) return '전도';
  if (/제자|훈련|discipleship/.test(t)) return '제자훈련';
  if (/상담|counseling/.test(t)) return '상담';
  if (/새벽|기도|prayer|중보/.test(t)) return '기도';
  // 4. 예배 (주일 1/2/3부 추가)
  if (/1부|2부|3부|예배|worship/.test(t)) return '예배';
  if (/선교|mission|단기/.test(t)) return '선교';
  if (/새가족|등록/.test(t)) return '새가족';
  if (/회의|meeting|미팅/.test(t)) return '회의';
  // 5. 행정 (지구촌 뉴스, 주보 추가)
  if (/지구촌 뉴스|뉴스 스크립트|주보|행정|office|admin|사무/.test(t)) return '행정';
  
  return ministryName || '목회';
}

// ── GET 핸들러 ───────────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'prev_week_report') {
    return getPrevWeekReport();
  }

  if (action === 'gcal_events') {
    return getCalendarEvents();
  }

  // 기본: 대시보드 데이터 반환
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('사역데이터');
    const values = sheet.getDataRange().getValues();
    const data = {};
    
    if (values.length > 0) {
      const firstVal = String(values[0][0]).trim();
      if (firstVal.startsWith('{') && firstVal.endsWith('}')) {
        // 이전 방식 호환: A1 셀에 전체 JSON이 들어 있는 경우
        return ContentService.createTextOutput(firstVal)
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      // 새 방식: 주차별 행 분할 파싱
      for (let i = 1; i < values.length; i++) {
        const weekKey = String(values[i][0]).trim();
        const jsonStr = String(values[i][1]).trim();
        if (weekKey && jsonStr) {
          try {
            data[weekKey] = JSON.parse(jsonStr);
          } catch(e) {
            // 파싱 실패 시 무시
          }
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST 핸들러 ──────────────────────────────────────────────────────────────
// 쓰기 가능한 캘린더 화이트리스트 (안전상 한 곳만 허용)
const WRITABLE_CALENDAR_ID = 'hyunchang.kang@gmcusa.org';

function doPost(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

    // 캘린더 쓰기 액션 분기 (사역 데이터 저장과 분리)
    if (action === 'gcal_create') return gcalCreate(e);
    if (action === 'gcal_update') return gcalUpdate(e);
    if (action === 'gcal_delete') return gcalDelete(e);

    const raw = e.postData.contents;
    const data = JSON.parse(raw); // 유효성 검사
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // 새 방식: 주차별 행 분할 저장 (50,000자 초과 방지)
    const sheet = ss.getSheetByName('사역데이터');
    sheet.clearContents();
    
    const rows = [['주차', '데이터']];
    Object.keys(data).sort().forEach(weekKey => {
      rows.push([weekKey, JSON.stringify(data[weekKey])]);
    });
    
    if (rows.length > 1) {
      sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
    
    updateTodoSheet(ss, data);
    updateDatabase(ss, data);     // ← 누적 DB 업데이트 추가
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.toString());
  }
}

// ── 캘린더 이벤트 생성 (내 스케줄만 허용) ────────────────────────────────────
function gcalCreate(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const calendarId = body.calendarId || WRITABLE_CALENDAR_ID;
    if (calendarId !== WRITABLE_CALENDAR_ID) {
      return jsonOut({ error: 'calendar not writable: ' + calendarId });
    }
    const title = (body.title || '').trim();
    const date = body.date;       // 'YYYY-MM-DD'
    const time = body.time || '';  // 'HH:MM' 또는 ''
    const recurrenceType = body.recurrence || ''; // 'daily', 'weekly', 'monthly', 'none'
    if (!title || !date) return jsonOut({ error: 'title and date required' });

    const cal = CalendarApp.getCalendarById(calendarId) || CalendarApp.getDefaultCalendar();
    if (!cal) return jsonOut({ error: 'calendar not found' });

    let event;
    const [yy, mo, dd] = date.split('-').map(Number);

    if (recurrenceType && recurrenceType !== 'none') {
      let rec = CalendarApp.newRecurrence();
      if (recurrenceType === 'daily') rec = rec.addDailyRule();
      else if (recurrenceType === 'weekly') rec = rec.addWeeklyRule();
      else if (recurrenceType === 'monthly') rec = rec.addMonthlyRule();

      if (time) {
        const [h, m] = time.split(':').map(Number);
        const start = new Date(yy, mo - 1, dd, h, m);
        const end = new Date(start.getTime() + 60 * 60 * 1000);
        event = cal.createEventSeries(title, start, end, rec);
      } else {
        event = cal.createAllDayEventSeries(title, new Date(yy, mo - 1, dd), rec);
      }
    } else {
      if (time) {
        // 시간 지정 이벤트 (1시간 기본)
        const [h, m] = time.split(':').map(Number);
        const start = new Date(yy, mo - 1, dd, h, m);
        const end = new Date(start.getTime() + 60 * 60 * 1000);
        event = cal.createEvent(title, start, end);
      } else {
        // 종일 이벤트
        event = cal.createAllDayEvent(title, new Date(yy, mo - 1, dd));
      }
    }

    return jsonOut({
      ok: true,
      eventId: 'gcal_' + event.getId().replace(/@.*$/, ''),
      calendarId: calendarId,
      calendarName: '내 스케줄'
    });
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

// ── 캘린더 이벤트 수정 (제목/시간만) ────────────────────────────────────────
function gcalUpdate(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const calendarId = body.calendarId || WRITABLE_CALENDAR_ID;
    if (calendarId !== WRITABLE_CALENDAR_ID) {
      return jsonOut({ error: 'calendar not writable: ' + calendarId });
    }
    const eventIdRaw = (body.eventId || '').replace(/^gcal_/, '');
    if (!eventIdRaw) return jsonOut({ error: 'eventId required' });
    const fullEventId = eventIdRaw + '@google.com';

    const cal = CalendarApp.getCalendarById(calendarId) || CalendarApp.getDefaultCalendar();
    if (!cal) return jsonOut({ error: 'calendar not found' });

    const event = cal.getEventById(fullEventId);
    if (!event) return jsonOut({ error: 'event not found: ' + fullEventId });

    // 제목 업데이트
    if (typeof body.title === 'string' && body.title.trim()) {
      event.setTitle(body.title.trim());
    }
    // 날짜/시간 업데이트
    if (body.date) {
      const [yy, mo, dd] = body.date.split('-').map(Number);
      if (body.time) {
        const [h, m] = body.time.split(':').map(Number);
        const start = new Date(yy, mo - 1, dd, h, m);
        const end = new Date(start.getTime() + 60 * 60 * 1000);
        event.setTime(start, end);
      } else {
        // 종일로 전환
        event.setAllDayDate(new Date(yy, mo - 1, dd));
      }
    }
    return jsonOut({ ok: true, eventId: 'gcal_' + event.getId().replace(/@.*$/, '') });
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

// ── 캘린더 이벤트 삭제 ────────────────────────────────────────────────────────
function gcalDelete(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const calendarId = body.calendarId || WRITABLE_CALENDAR_ID;
    if (calendarId !== WRITABLE_CALENDAR_ID) {
      return jsonOut({ error: 'calendar not writable: ' + calendarId });
    }
    const eventIdRaw = (body.eventId || '').replace(/^gcal_/, '');
    if (!eventIdRaw) return jsonOut({ error: 'eventId required' });
    const fullEventId = eventIdRaw + '@google.com';

    const cal = CalendarApp.getCalendarById(calendarId) || CalendarApp.getDefaultCalendar();
    if (!cal) return jsonOut({ error: 'calendar not found' });

    const event = cal.getEventById(fullEventId);
    if (event) {
      event.deleteEvent();
      return jsonOut({ ok: true });
    }
    return jsonOut({ error: 'event not found: ' + fullEventId });
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 이번 주 할 일 목록 (현재 주 스냅샷, 덮어쓰기) ──────────────────────────
function updateTodoSheet(ss, data) {
  let sheet = ss.getSheetByName('이번주할일목록');
  if (!sheet) sheet = ss.insertSheet('이번주할일목록');
  const rows = [['주차', '사역명', '날짜', '시간', '할일', '완료', '카테고리', '연간사역']];
  const timeRegex = /^\[\d{2}:\d{2}\]\s*/;
  Object.keys(data).sort().forEach(weekKey => {
    const wk = data[weekKey];
    (wk.ministries || []).forEach(m => {
      (m.todos || []).forEach(t => {
        rows.push([
          weekKey, m.name || '',
          t.date || '', t.time || '',
          (t.text || '').replace(timeRegex, ''),
          t.done ? '✅' : '',
          getCategoryLabel(t.category || m.category) || detectCategoryGAS(t.text, m.name),
          t.annualEventTitle || ''
        ]);
      });
    });
  });
  sheet.clearContents();
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ── 누적 사역 데이터베이스 (upsert by todo ID) ──────────────────────────────
function updateDatabase(ss, data) {
  let db = ss.getSheetByName('사역데이터베이스');
  if (!db) {
    db = ss.insertSheet('사역데이터베이스');
    db.getRange(1, 1, 1, DB_HEADERS.length).setValues([DB_HEADERS]);
    db.setFrozenRows(1);
    // 헤더 서식
    db.getRange(1, 1, 1, DB_HEADERS.length)
      .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 기존 ID → 행 번호 맵 구성
  const existing = db.getDataRange().getValues();
  const idMap = {};
  for (let i = 1; i < existing.length; i++) {
    const id = String(existing[i][COL.ID]);
    if (id) idMap[id] = i + 1; // 1-indexed
  }

  const now = new Date();
  const timeRegex = /^\[\d{2}:\d{2}\]\s*/;
  const toUpdate = [];
  const toAppend = [];
  const activeIds = new Set();
  const weeksInPayload = new Set();

  Object.keys(data).sort().forEach(weekKey => {
    weeksInPayload.add(weekKey);
    const wk = data[weekKey];
    (wk.ministries || []).forEach(m => {
      (m.todos || []).forEach(t => {
        if (!t.id) return;
        activeIds.add(t.id);
        const content = (t.text || '').replace(timeRegex, '').trim();
        const category = getCategoryLabel(t.category || m.category) || detectCategoryGAS(content, m.name);
        const row = [
          t.id,                          // ID
          weekKey,                       // 주차
          m.name || '',                  // 사역명
          m.category || '',              // 사역카테고리
          toMDY(t.date || ''),           // 날짜 (MM-DD-YYYY 형식으로 저장)
          t.time || '',                  // 시간
          content,                       // 사역내용
          t.done ? '✅' : '',            // 완료
          category,                      // 분류 (자동감지 포함)
          t.annualEventTitle || '',      // 연간사역
          t.needs || '',                 // 참고사항
          null,                          // 최초등록일 (아래서 설정)
          now                            // 최종수정일
        ];

        if (idMap[t.id]) {
          row[COL.CREATED] = existing[idMap[t.id] - 1][COL.CREATED]; // 최초등록일 유지
          toUpdate.push({ rowNum: idMap[t.id], values: row });
        } else {
          row[COL.CREATED] = now;
          toAppend.push(row);
        }
      });
    });
  });

  // 기존 항목 업데이트
  toUpdate.forEach(u => {
    db.getRange(u.rowNum, 1, 1, DB_HEADERS.length).setValues([u.values]);
  });

  // 새 항목 추가
  if (toAppend.length > 0) {
    db.getRange(db.getLastRow() + 1, 1, toAppend.length, DB_HEADERS.length).setValues(toAppend);
  }

  // ── 삭제 반영 (이번 페이로드의 주차 중 제외된 ID 행 물리 삭제) ────────────────
  // 행 인덱스 뒤틀림 방지를 위해 마지막 행부터 역순으로 검색하여 삭제
  for (let i = existing.length - 1; i >= 1; i--) {
    const rowNum = i + 1; // 1-indexed
    const rowWeek = String(existing[i][COL.WEEK]);
    const rowId = String(existing[i][COL.ID]);

    if (weeksInPayload.has(rowWeek) && !activeIds.has(rowId)) {
      db.deleteRow(rowNum);
    }
  }
}

// ── 구글 캘린더 이벤트 반환 (3주 범위: 지난주 화요일~다음주 월요일) ───────────────────
function getCalendarEvents() {
  try {
    const CALENDARS = [
      { id: 'hyunchang.kang@gmcusa.org',  name: '내 스케줄' },
      { id: 'office@gmcusa.org',           name: 'GMC OFFICE' },
      { id: 'c_d4379e87ef260d2f72099d423dc07792c97e0a4817599315ff8f38122ea055e2@group.calendar.google.com', name: '사역자 스케줄' }
    ];

    // 이번 주 화요일 계산 (dow: 0=일, 1=월, 2=화, ...)
    const today = new Date();
    const dow = today.getDay();
    // 화요일(2)까지 거슬러 올라가기
    const daysToTue = (dow === 0) ? 5 : (dow === 1) ? 6 : (dow - 2);
    const thisTuesday = new Date(today);
    thisTuesday.setDate(today.getDate() - daysToTue);
    thisTuesday.setHours(0, 0, 0, 0);

    // 3주 범위 계산: 지난주 화요일 ~ 다음주 월요일
    const startRange = new Date(thisTuesday);
    startRange.setDate(thisTuesday.getDate() - 7);
    startRange.setHours(0, 0, 0, 0);

    const endRange = new Date(thisTuesday);
    endRange.setDate(thisTuesday.getDate() + 13);
    endRange.setHours(23, 59, 59, 999);

    const results = [];

    CALENDARS.forEach(cal => {
      let calendar;
      try {
        calendar = CalendarApp.getCalendarById(cal.id);
      } catch (e) { return; }
      if (!calendar) return;

      const events = calendar.getEvents(startRange, endRange);
      events.forEach(ev => {
        const start = ev.getStartTime();
        const isAllDay = ev.isAllDayEvent();
        const yy = start.getFullYear();
        const mm = String(start.getMonth() + 1).padStart(2, '0');
        const dd = String(start.getDate()).padStart(2, '0');
        const hh = String(start.getHours()).padStart(2, '0');
        const mi = String(start.getMinutes()).padStart(2, '0');

        results.push({
          id:           'gcal_' + ev.getId().replace(/@.*$/, ''),
          calendarId:   cal.id,
          calendarName: cal.name,
          date:         `${yy}-${mm}-${dd}`,
          time:         isAllDay ? '' : `${hh}:${mi}`,
          title:        ev.getTitle()
        });
      });
    });

    // 날짜/시간 순 정렬
    results.sort((a, b) => {
      const da = a.date + (a.time || '00:00');
      const db = b.date + (b.time || '00:00');
      return da < db ? -1 : da > db ? 1 : 0;
    });

    return ContentService.createTextOutput(JSON.stringify(results))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 지난 주 보고서 데이터 반환 (분류|날짜|사역내용|참고사항) ────────────────
function getPrevWeekReport() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const db = ss.getSheetByName('사역데이터베이스');
    if (!db) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 지난 주 날짜 범위 계산 (지난 화요일 ~ 지난 월요일)
    const today = new Date();
    const dow = today.getDay(); // 0=일, 1=월, 2=화, ..., 6=토
    // 이번 주 화요일까지 거슬러 올라가기
    // dow=0(일)→5일전, 1(월)→6일전, 2(화)→0일전, 3(수)→1일전, ...
    const daysToThisTue = (dow === 0) ? 5 : (dow === 1) ? 6 : (dow - 2);
    const thisTuesday = new Date(today);
    thisTuesday.setDate(today.getDate() - daysToThisTue);
    thisTuesday.setHours(0, 0, 0, 0);

    const lastTuesday = new Date(thisTuesday);
    lastTuesday.setDate(thisTuesday.getDate() - 7);

    const lastMonday = new Date(lastTuesday);
    lastMonday.setDate(lastTuesday.getDate() + 6);
    lastMonday.setHours(23, 59, 59, 999);

    const allRows = db.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < allRows.length; i++) {
      const r = allRows[i];
      const dateStr = toMDY(r[COL.DATE]); // MM-DD-YYYY 형식
      if (!dateStr) continue;

      const rowDate = parseMDY(dateStr);
      if (!rowDate || isNaN(rowDate.getTime())) continue;
      if (rowDate < lastTuesday || rowDate > lastMonday) continue;

      const content = String(r[COL.CONTENT] || '').trim();
      if (!content) continue;

      results.push({
        category: String(r[COL.CATEGORY] || r[COL.M_CAT] || '').trim(),
        date:     dateStr,   // MM-DD-YYYY 형식
        content:  content,
        note:     String(r[COL.NOTE] || '').trim()
      });
    }

    // 날짜순 정렬 (MM-DD-YYYY는 문자열 정렬 불가 → Date 객체로 비교)
    results.sort((a, b) => {
      const da = parseMDY(a.date), db = parseMDY(b.date);
      return (da && db) ? da - db : 0;
    });

    return ContentService.createTextOutput(JSON.stringify(results))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Gmail 연동: 메일을 분석하여 '내 스케줄'에 할 일 자동 추가 ───────────────────────
function checkGmailAndExtractTasks() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    Logger.log("[Gmail Sync] 에러: Script Properties에 GEMINI_API_KEY가 설정되어 있지 않습니다.");
    return;
  }

  let label = GmailApp.getUserLabelByName("gmc-processed");
  if (!label) {
    label = GmailApp.createLabel("gmc-processed");
  }

  // inbox에 있고 gmc-processed 라벨이 없으며 최근 7일 내에 수신된 이메일 검색 (최대 20개 스레드)
  const threads = GmailApp.search("in:inbox -label:gmc-processed newer_than:7d", 0, 20);
  if (threads.length === 0) {
    Logger.log("[Gmail Sync] 분석할 새 이메일이 없습니다.");
    return;
  }

  Logger.log("[Gmail Sync] 분석할 이메일 스레드 발견: " + threads.length + "개");

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const dataSheet = ss.getSheetByName('사역데이터');
  const values = dataSheet.getDataRange().getValues();
  
  // 1. 기존 사역 데이터 로드
  const data = {};
  for (let i = 1; i < values.length; i++) {
    const weekKey = String(values[i][0]).trim();
    const jsonStr = String(values[i][1]).trim();
    if (weekKey && jsonStr) {
      try {
        data[weekKey] = JSON.parse(jsonStr);
      } catch(e) {}
    }
  }

  let changed = false;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    if (messages.length === 0) return;

    // 마지막 메시지(가장 최신 메일) 분석
    const lastMsg = messages[messages.length - 1];
    const subject = lastMsg.getSubject();
    const body = lastMsg.getPlainBody();
    const date = lastMsg.getDate();

    // 이메일 수신 날짜를 기준 날짜 문자열로 변환 (Gemini에게 상대 시간 계산의 기준점으로 제공)
    const yy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const emailDateStr = yy + "-" + mm + "-" + dd;

    Logger.log("[Gmail Sync] 이메일 분석 중: Subject: '" + subject + "', Date: " + emailDateStr);

    const tasks = callGeminiAPI(apiKey, subject, body, emailDateStr);
    if (Array.isArray(tasks) && tasks.length > 0) {
      Logger.log("[Gmail Sync] 추출된 할 일: " + tasks.length + "개");
      
      tasks.forEach(task => {
        if (!task.text || !task.date) return;

        // 주차 키 계산 (Tuesday)
        const weekKey = getTuesdayForDate(task.date);
        
        if (!data[weekKey]) {
          data[weekKey] = { ministries: [] };
        }
        const wk = data[weekKey];

        // '내 스케줄' 카드 찾기
        let mySchedule = wk.ministries.find(m => m.gcalId === 'hyunchang.kang@gmcusa.org' || m.name === '내 스케줄');
        if (!mySchedule) {
          const stableId = 'gcalm_' + weekKey + '_hyunchang_kang_gmcusa_org';
          mySchedule = {
            id: stableId,
            gcalId: 'hyunchang.kang@gmcusa.org',
            name: '내 스케줄',
            owner: '',
            color: '#1a3a5c',
            status: 's-planning',
            todos: [],
            progress: '',
            needs: '',
            pinned: false
          };
          wk.ministries.push(mySchedule);
        }

        // 중복 방지 검사 (동일 날짜 및 텍스트)
        const cleanText = task.text.trim();
        const exists = mySchedule.todos.some(t => 
          t.date === task.date && 
          (t.text.replace(/^\[\d{2}:\d{2}\]\s*/, '').trim() === cleanText)
        );

        if (!exists) {
          function uid() {
            return Math.random().toString(36).substring(2, 11);
          }
          const timeLabel = task.time ? `[${task.time}] ` : '';
          mySchedule.todos.push({
            id: 'gmail_' + uid(),
            text: timeLabel + cleanText,
            date: task.date,
            time: task.time || '',
            done: false,
            progress: '',
            needs: '',
            category: task.category || 'other',
            updatedAt: Date.now()
          });
          changed = true;
          Logger.log("[Gmail Sync] 할 일 추가 완료: " + cleanText + " (날짜: " + task.date + ")");
        }
      });
    }

    // 분석 완료 후 중복 처리 방지를 위해 gmc-processed 라벨 부착
    thread.addLabel(label);
  });

  // 변경 사항이 있는 경우에만 전체 시트 갱신
  if (changed) {
    dataSheet.clearContents();
    const rows = [['주차', '데이터']];
    Object.keys(data).sort().forEach(weekKey => {
      rows.push([weekKey, JSON.stringify(data[weekKey])]);
    });
    if (rows.length > 1) {
      dataSheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
    
    updateTodoSheet(ss, data);
    updateDatabase(ss, data);
    Logger.log("[Gmail Sync] 최종 변경 데이터를 사역데이터 및 누적 시트에 동기화 완료했습니다.");
  } else {
    Logger.log("[Gmail Sync] 추가된 할 일이 없어 시트 업데이트를 건너뜁니다.");
  }
}

// ── Gemini API 호출을 통해 할 일 추출 ──────────────────────────────────────────
function callGeminiAPI(apiKey, subject, body, emailDateStr) {
  const model = "gemini-1.5-flash";
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
  
  const systemPrompt = "You are a professional secretary for Pastor Hyunchang Kang. Extract actionable tasks (to-do items) that he needs to perform based on the email content. " +
                       "Analyze the email context carefully. " +
                       "Return ONLY a valid JSON array of tasks (no markdown block fences, no prose). " +
                       "If there are no tasks for him to perform, return an empty JSON array: []. " +
                       "The JSON schema must be an array of objects, where each object has:\n" +
                       "- text: a clear, concise Korean summary of the task.\n" +
                       "- date: the date when this task needs to be done, in 'YYYY-MM-DD' format. Use the email's context and the current reference date (" + emailDateStr + ") to resolve relative terms like '내일' (tomorrow), '이번주 목요일' (this Thursday), etc. If no specific date is mentioned, use today's date (" + emailDateStr + ").\n" +
                       "- time: time of the event in 'HH:MM' format if mentioned, otherwise ''.\n" +
                       "- category: choose the best category key from: worship, sermon, disciple, cell, mission, newmember, meeting, admin, pastoral, evangelism, fellowship, facility, shortmission, intercession, other.";

  const payload = {
    contents: [
      {
        parts: [
          { text: systemPrompt + "\n\nEmail to analyze:\nEmail Date: " + emailDateStr + "\nSubject: " + subject + "\nBody:\n" + body }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log("[Gmail Sync] Gemini API 호출 실패: Status " + code + ", Response: " + response.getContentText());
      return [];
    }

    let text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
    // 마크다운 펜스 제거 서포트
    if (text.startsWith("```json")) {
      text = text.substring(7);
    }
    if (text.endsWith("```")) {
      text = text.substring(0, text.length - 3);
    }
    text = text.trim();
    
    return JSON.parse(text);
  } catch (e) {
    Logger.log("[Gmail Sync] Gemini 응답 파싱 에러: " + e.toString());
    return [];
  }
}

// ── 입력된 YYYY-MM-DD 날짜에 해당하는 주차(화요일) 날짜 구하기 ───────────────────
function getTuesdayForDate(dateStr) {
  const parts = dateStr.split('-');
  const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]), 12, 0, 0);
  const day = d.getDay(); // 0=일, 1=월, 2=화, ...
  const daysToTue = (day === 0) ? 5 : (day === 1) ? 6 : (day - 2);
  const tue = new Date(d);
  tue.setDate(d.getDate() - daysToTue);
  
  const yy = tue.getFullYear();
  const mm = String(tue.getMonth() + 1).padStart(2, '0');
  const dd = String(tue.getDate()).padStart(2, '0');
  return yy + "-" + mm + "-" + dd;
}

// ── 매일 아침 7시에 실행되는 트리거 생성 헬퍼 함수 ──────────────────────────────
function setupGmailSyncTrigger() {
  const functionName = 'checkGmailAndExtractTasks';
  const triggers = ScriptApp.getProjectTriggers();
  
  // 기존 동일 함수에 대한 트리거가 있다면 중복 제거
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 매일 아침 7시 ~ 8시 사이 실행 트리거 생성 (Google 제한으로 1시간 윈도우 발생)
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  Logger.log("[Gmail Sync] 매일 아침 7시 자동 실행 트리거가 성공적으로 설정되었습니다.");
}
