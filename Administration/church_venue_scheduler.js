/**
 * 교회 장소 사용 신청 자동 연동 및 일정 충돌 관리 시스템 (Google Apps Script)
 * 
 * - 성도들의 교회 장소 사용 신청 정보를 Google Calendar에 자동 등록합니다.
 * - 타임스탬프(선착순)를 기준으로 동일 장소/시간에 중복 신청 시 자동 걸러냅니다.
 * - 다른 성도가 동일 장소/시간에 신청하면 '일정 충돌'로 처리하고 안내 이메일을 발송합니다.
 */

// =========================================================================
// 1. 설정 영역 (CONFIGURATION)
// =========================================================================
// 연동할 Google Calendar ID (교회 장소 사용 캘린더)
const CALENDAR_ID = 'c_ee046a6c02b41e714cc6c9810f69392b5ddf35a782710763e77c784fde229c27@group.calendar.google.com';

// 스프레드시트 열 매핑 (0부터 시작: A열 = 0, B열 = 1, ...)
const COL_TIMESTAMP = 0;       // A열: 타임스탬프 (신청 시각)
const COL_DEPT = 1;            // B열: 신청부서 (Ministry Department)
const COL_TEAM = 2;            // C열: 신청팀 (Ministry Team)
const COL_LEADER = 3;          // D열: 팀장 (Ministry Team Decon)
const COL_NAME = 4;            // E열: 신청자 이름(Applicant Name) - 중복/충돌 검증 기준 1
const COL_PHONE = 5;           // F열: 연락처 (Cell Phone)
const COL_DATETIME = 6;        // G열: 사용 일자, 요일, 시간 - 중복/충돌 검증 기준 2
const COL_ROOM = 7;            // H열: 교실 번호 (Room Number) - 중복/충돌 검증 기준 3
const COL_PARTICIPANTS = 8;    // I열: 사용 인원 (Number of Participants)
const COL_RESPONSIBLE = 9;     // J열: 최종 책임자 (Final Responsible Person)
const COL_PURPOSE = 10;        // K열: 사용 목적 (Purpose/ Usage Description)
const COL_EMAIL = 11;          // L열: Email (신청자 이메일 - 충돌 시 알림처)
const COL_EQUIPMENT = 12;      // M열: 필요한 장비들 (Equipment Needed)
const COL_STATUS = 13;         // N열: 처리 상태 (Status - 신설 컬럼)

// 안내 메일 발송 계정/참조 주소 (행정실)
// 주의: MailApp은 "스크립트를 소유한 계정"으로 발송됩니다.
// 발신자를 office@gmcusa.org로 하려면 이 스크립트(와 트리거)를 office@gmcusa.org 계정이 소유해야 합니다.
const OFFICE_EMAIL = 'office@gmcusa.org';
const MAIL_SENDER_NAME = 'GMC 행정실';

// 처리 상태 명칭 정의
const STATUS_APPROVED = '등록 완료';
const STATUS_DUPLICATE = '중복 신청';
const STATUS_CONFLICT = '일정 충돌';
const STATUS_ERROR = '오류 발생';

// =========================================================================
// 2. 메인 트리거 함수 (Google Sheets Trigger)
// =========================================================================

/**
 * 구글 설문지 제출 시 또는 수동 실행 시 미처리된 신청들을 순차적으로 처리합니다.
 */
function onFormSubmit(e) {
  Logger.log("신청 접수 트리거 실행...");
  initializeSheet();
  processLatestSubmissions();
}

/**
 * 시트에 '처리 상태' 열(N열) 헤더가 없는 경우 자동으로 이쁘게 생성합니다.
 */
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = sheet.getLastColumn();
  
  // A열부터 가져와서 헤더 확인
  let headers = [];
  if (lastCol > 0) {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  
  // N열(14번째 열)에 헤더가 없거나 다를 경우 자동 설정
  if (headers.length < 14 || !headers[13]) {
    const headerCell = sheet.getRange(1, 14);
    headerCell.setValue("처리 상태 (Status)");
    headerCell.setFontWeight("bold");
    headerCell.setBackground("#cfe2f3"); // 부드러운 파란색 계열 파스텔톤
    headerCell.setHorizontalAlignment("center");
    sheet.autoResizeColumn(14);
    Logger.log("N열에 '처리 상태 (Status)' 헤더 컬럼이 생성되었습니다.");
  }
}

/**
 * 처리 상태가 비어있는(새로 신청된) 모든 행을 찾아 타임스탬프 순서대로 처리합니다.
 * 이를 통해 일시적인 오류나 동시 제출 상황에서도 선착순(Timestamp) 처리를 완벽하게 보장합니다.
 */
function processLatestSubmissions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("처리할 데이터 행이 없습니다.");
    return;
  }
  
  // 1. 미처리된 행 (Status가 비어있는 행) 수집
  const unprocessedRows = [];
  for (let i = 1; i < data.length; i++) {
    const timestampVal = data[i][COL_TIMESTAMP];
    const status = String(data[i][COL_STATUS]).trim();
    
    // 타임스탬프가 존재하고 처리 상태가 비어있는 경우에만 수집
    if (timestampVal && !status) {
      unprocessedRows.push({
        rowIndex: i + 1, // Sheets API는 1-indexed이므로 행 번호는 i + 1
        data: data[i],
        timestamp: new Date(timestampVal)
      });
    }
  }
  
  if (unprocessedRows.length === 0) {
    Logger.log("새로 신청된(미처리된) 행이 없습니다.");
    return;
  }
  
  // 2. 타임스탬프 오름차순(가장 먼저 신청한 순서대로) 정렬
  unprocessedRows.sort((a, b) => a.timestamp - b.timestamp);
  Logger.log(`총 ${unprocessedRows.length}개의 미처리 신청 건을 선착순으로 처리합니다.`);
  
  // 3. 순차적으로 중복 및 충돌 검증 수행
  for (let task of unprocessedRows) {
    const idx = task.rowIndex;
    const row = task.data;
    
    const timestamp = task.timestamp;
    const name = String(row[COL_NAME]).trim();
    const phone = String(row[COL_PHONE]).trim();
    const dateTimeStr = String(row[COL_DATETIME]).trim();
    const room = String(row[COL_ROOM]).trim();
    const purpose = String(row[COL_PURPOSE]).trim();
    const email = String(row[COL_EMAIL]).trim();
    const participants = String(row[COL_PARTICIPANTS]).trim();
    const equipment = String(row[COL_EQUIPMENT]).trim();
    
    // 최신 상태 비교를 위해 시트의 현재 실시간 데이터를 매번 다시 읽어옵니다
    const currentSheetData = sheet.getDataRange().getValues();
    
    let isDuplicate = false;
    let isConflict = false;
    
    const taskSlots = getDateTimeSlots(dateTimeStr, timestamp);
    if (taskSlots.length === 0) {
      sheet.getRange(idx, COL_STATUS + 1).setValue(STATUS_ERROR);
      sheet.getRange(idx, COL_STATUS + 1).setBackground("#f8cbad"); // 연한 주황색(오류) 표시
      Logger.log(`[오류] ${idx}번 행 (${name}님, ${room}호): 날짜/시간 정보를 파싱할 수 없습니다.`);
      sendErrorEmail(email, name, room, dateTimeStr,
        "입력하신 사용 일자·요일·시간을 시스템이 인식하지 못했습니다. 예: '2026년 10월 11일 주일 오후 1시~3시' 와 같이 날짜, 요일, 시작·종료 시간을 함께 적어 다시 신청해 주시기 바랍니다.");
      continue;
    }
    
    // 기존 신청 목록과 비교
    for (let i = 1; i < currentSheetData.length; i++) {
      // 본인 행은 비교 대상에서 제외
      if (i + 1 === idx) continue;
      
      const otherTimestampVal = currentSheetData[i][COL_TIMESTAMP];
      if (!otherTimestampVal) continue;
      
      const otherTimestamp = new Date(otherTimestampVal);
      const otherStatus = String(currentSheetData[i][COL_STATUS]).trim();
      
      // 예약이 정상 '등록 완료' 되었거나, 
      // 혹은 아직 처리는 안 되었으나 우리보다 타임스탬프가 앞선(먼저 신청한) 유효한 신청건인지 확인
      const isOtherValid = (otherStatus === STATUS_APPROVED) || 
                            (!otherStatus && otherTimestamp < timestamp);
      
      // 이미 중복이나 충돌로 탈락한 예약은 비교 대상에서 제외
      if (!isOtherValid || otherStatus === STATUS_DUPLICATE || otherStatus === STATUS_CONFLICT) {
        continue;
      }
      
      const otherRoom = String(currentSheetData[i][COL_ROOM]).trim();
      
      // 방 번호가 다르면 충돌 가능성 없음
      if (cleanStr(otherRoom) !== cleanStr(room)) {
        continue;
      }
      
      const otherDateTimeStr = String(currentSheetData[i][COL_DATETIME]).trim();
      const otherSlots = getDateTimeSlots(otherDateTimeStr, otherTimestamp);
      
      // 실제 예약 시간대 겹침 확인
      if (doSlotsConflict(taskSlots, otherSlots)) {
        const otherName = String(currentSheetData[i][COL_NAME]).trim();
        const otherEmail = String(currentSheetData[i][COL_EMAIL]).trim();

        if (isSameApplicant(name, email, otherName, otherEmail)) {
          // [중복 신청] 동일한 사람, 일정 겹침, 동일한 교실인 경우
          isDuplicate = true;
          break; // 중복이 감지되면 즉시 루프 중단 (중복 우선 적용)
        } else {
          // [일정 충돌] 다른 사람, 일정 겹침, 동일한 교실인 경우
          isConflict = true;
        }
      }
    }
    
    // 4. 판단 결과에 따른 분기 처리
    if (isDuplicate) {
      // 중복 신청 처리 (캘린더 등록 안 함, 상태만 기록)
      sheet.getRange(idx, COL_STATUS + 1).setValue(STATUS_DUPLICATE);
      sheet.getRange(idx, COL_STATUS + 1).setBackground("#f3f3f3"); // 밝은 회색 표시
      Logger.log(`[중복] ${idx}번 행 (${name}님, ${room}호): 이미 본인이 신청한 내역이 존재합니다.`);

      // 중복 신청 안내 메일 발송 (이미 접수된 신청이 유효함을 알림)
      sendDuplicateEmail(email, name, room, dateTimeStr);

    } else if (isConflict) {
      // 일정 충돌 처리 (캘린더 등록 안 함, 상태 기록 + 이메일 안내 발송)
      sheet.getRange(idx, COL_STATUS + 1).setValue(STATUS_CONFLICT);
      sheet.getRange(idx, COL_STATUS + 1).setBackground("#ea9999"); // 연한 빨간색 표시
      Logger.log(`[충돌] ${idx}번 행 (${name}님, ${room}호): 선순위 예약과 충돌이 감지되어 취소되었습니다.`);
      
      // 나중에 신청한 이 성도님께 안내 메일 발송
      sendConflictEmail(email, name, room, dateTimeStr);
      
    } else {
      // 예약 등록 승인! (캘린더 등록 + 상태 완료 기록)
      const isSuccess = addToGoogleCalendar(
        room, name, dateTimeStr, purpose, phone, participants, equipment
      );
      
      if (isSuccess) {
        sheet.getRange(idx, COL_STATUS + 1).setValue(STATUS_APPROVED);
        sheet.getRange(idx, COL_STATUS + 1).setBackground("#b6d7a8"); // 연한 초록색 표시
        Logger.log(`[성공] ${idx}번 행 (${name}님, ${room}호): 구글 캘린더에 성공적으로 등록되었습니다.`);
        // 예약 승인 확인 메일 발송
        sendApprovalEmail(email, name, room, dateTimeStr);
      } else {
        sheet.getRange(idx, COL_STATUS + 1).setValue(STATUS_ERROR);
        sheet.getRange(idx, COL_STATUS + 1).setBackground("#f8cbad"); // 연한 주황색(오류) 표시
        Logger.log(`[오류] ${idx}번 행 (${name}님, ${room}호): 캘린더 등록 중 문제가 발생했습니다.`);
        sendErrorEmail(email, name, room, dateTimeStr,
          "시스템 오류로 캘린더에 예약을 등록하지 못했습니다. 행정실에서 확인 후 별도로 연락드리겠습니다.");
      }
    }
  }
}

// =========================================================================
// 3. 보조 기능 함수 (Utility Functions)
// =========================================================================

/**
 * 비교 정확도를 위해 문자열 내 모든 공백을 제거하고 소문자로 표준화합니다.
 */
function cleanStr(str) {
  if (!str) return "";
  return String(str).replace(/\s+/g, "").toLowerCase();
}

// 이메일 오타로 보고 동일인으로 간주하는 최대 편집 거리(글자 수)
const EMAIL_TYPO_MAX_DISTANCE = 2;

/** 두 문자열의 편집 거리(Levenshtein)를 반환합니다. */
function editDistance(a, b) {
  const m = a.length, n = b.length;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[n];
}

/**
 * 두 신청이 같은 신청자인지 판정합니다. (중복 신청 vs 일정 충돌 구분 기준)
 * - 양쪽 이메일이 모두 유효한 경우:
 *   · 이메일이 완전히 같으면 동일인
 *   · 이름이 같고 이메일이 EMAIL_TYPO_MAX_DISTANCE 글자 이내로만 다르면 오타로 보고 동일인
 *   · 그 외(이름이 같아도 이메일이 많이 다르면)는 동명이인으로 봅니다.
 * - 이메일이 없으면 공백을 제거한 이름으로 비교합니다.
 */
function isSameApplicant(nameA, emailA, nameB, emailB) {
  const a = cleanStr(emailA);
  const b = cleanStr(emailB);
  const sameName = cleanStr(nameA) !== "" && cleanStr(nameA) === cleanStr(nameB);
  if (a.includes("@") && b.includes("@")) {
    if (a === b) return true;
    return sameName && editDistance(a, b) <= EMAIL_TYPO_MAX_DISTANCE;
  }
  return sameName;
}

/**
 * 모든 안내 메일의 공통 발송 함수.
 * - 발신: 스크립트 소유 계정(office@gmcusa.org 계정이 소유해야 함), 표시 이름은 'GMC 행정실'
 * - 참조(CC): office@gmcusa.org, 회신(replyTo): office@gmcusa.org
 * - 신청자 이메일이 없거나 형식이 잘못되면 행정실로 '수동 안내 필요' 알림을 보냅니다.
 */
function sendNoticeEmail(email, name, subject, body) {
  if (!email || !email.includes("@")) {
    Logger.log(`[이메일 생략] '${name}'님의 이메일 주소('${email}')가 올바르지 않아 메일을 발송하지 못했습니다.`);
    try {
      MailApp.sendEmail(OFFICE_EMAIL,
        `[행정실 확인 필요] ${name}님께 안내 메일을 보내지 못했습니다`,
        `신청자 이메일 주소('${email}')가 올바르지 않아 아래 안내 메일을 발송하지 못했습니다. 직접 연락해 주세요.\n\n제목: ${subject}\n\n${body}`,
        { name: MAIL_SENDER_NAME });
    } catch (error) {
      Logger.log(`[이메일 에러] 행정실 알림 발송 실패: ` + error.toString());
    }
    return false;
  }
  try {
    MailApp.sendEmail(email, subject, body, {
      cc: OFFICE_EMAIL,
      replyTo: OFFICE_EMAIL,
      name: MAIL_SENDER_NAME
    });
    Logger.log(`[이메일 발송 완료] ${email} (${name}님): ${subject}`);
    return true;
  } catch (error) {
    Logger.log(`[이메일 에러] ${email} 발송 실패: ` + error.toString());
    return false;
  }
}

function buildMailBody(name, intro, room, dateTimeStr) {
  return `안녕하세요, ${name} 성도님.

${intro}

■ 신청 장소: ${room}
■ 사용 시간: ${dateTimeStr}

기타 문의는 office@gmcusa.org로 해주시기 바랍니다.

GMC 행정실 드림
---------------------------------------------
본 메일은 구글 스프레드시트 장소 예약 시스템에서 자동으로 발송되었습니다.`;
}

/** 예약 승인 확인 이메일 */
function sendApprovalEmail(email, name, room, dateTimeStr) {
  sendNoticeEmail(email, name,
    `[교회 장소 신청] 장소 예약이 확인되었습니다.`,
    buildMailBody(name, "장소 신청 예약이 확인되었습니다. 신청하신 요일/시간에 장소를 사용하실 수 있습니다.", room, dateTimeStr));
}

/** 일정 충돌 안내 이메일 */
function sendConflictEmail(email, name, room, dateTimeStr) {
  sendNoticeEmail(email, name,
    `[교회 장소 신청 안내] 신청하신 장소 예약 일정에 충돌이 발생했습니다.`,
    buildMailBody(name, "신청하신 장소는 같은 시간대에 이미 다른 분이 먼저 신청하여 예약되어 있습니다. 이번 신청은 등록되지 않았습니다. 다른 요일/시간으로 조정하거나 다른 장소를 정하여 다시 신청해 주시기 바랍니다.", room, dateTimeStr));
}

/** 중복 신청 안내 이메일 */
function sendDuplicateEmail(email, name, room, dateTimeStr) {
  sendNoticeEmail(email, name,
    `[교회 장소 신청 안내] 동일한 신청이 이미 접수되어 있습니다.`,
    buildMailBody(name, "본인이 같은 장소·겹치는 시간에 이미 신청하신 내역이 있어 이번 신청은 중복으로 처리되었습니다. 이전에 접수된 신청이 유효하며 다시 신청하실 필요가 없습니다. 시간을 변경하시려면 새로 신청하지 마시고 행정실로 문의해 주시기 바랍니다. 이전 신청의 확인 메일을 받지 못하셨다면 행정실로 문의해 주시기 바랍니다.", room, dateTimeStr));
}

/** 오류 안내 이메일 (reason: 오류 사유) */
function sendErrorEmail(email, name, room, dateTimeStr, reason) {
  sendNoticeEmail(email, name,
    `[교회 장소 신청 안내] 신청하신 장소 예약이 처리되지 않았습니다.`,
    buildMailBody(name, "신청하신 장소 예약이 처리되지 않았습니다.\n\n■ 사유: " + reason, room, dateTimeStr));
}

/**
 * 성도가 입력한 다양한 한글/영문 텍스트에서 날짜를 지능적으로 분석하여 반환합니다.
 */
function parseDateFromText(text, defaultDate) {
  if (!text) return defaultDate;
  
  const now = defaultDate || new Date();
  const currentYear = now.getFullYear();
  
  // 1. YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD 패턴 추출
  let match = text.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }
  
  // 2. M월 D일 또는 MM월 DD일 패턴 추출
  match = text.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  if (match) {
    return new Date(currentYear, parseInt(match[1]) - 1, parseInt(match[2]));
  }
  
  // 3. M/D, M.D, M-D (예: 5/24, 12-25) 패턴 추출
  match = text.match(/(\d{1,2})[-./](\d{1,2})/);
  if (match) {
    const month = parseInt(match[1]);
    const day = parseInt(match[2]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(currentYear, month - 1, day);
    }
  }
  
  // 4. 날짜 텍스트가 없고 '일요일', '주일', '토요일' 등 요일 텍스트만 들어있는 경우, 
  // 신청 시점(기준일) 이후로 돌아오는 가장 가까운 해당 요일을 찾아 매핑합니다.
  const weekdaysKo = ["일", "월", "화", "수", "목", "금", "토"];
  const weekdaysEn = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
  
  let targetDayIndex = -1;
  
  if (text.includes("주일") || text.includes("일요일") || text.includes("일요")) {
    targetDayIndex = 0;
  } else {
    for (let d = 0; d < weekdaysKo.length; d++) {
      if (text.includes(weekdaysKo[d] + "요일")) {
        targetDayIndex = d;
        break;
      }
    }
  }
  
  if (targetDayIndex === -1) {
    for (let d = 0; d < weekdaysEn.length; d++) {
      if (text.toLowerCase().includes(weekdaysEn[d])) {
        targetDayIndex = d;
        break;
      }
    }
  }
  
  if (targetDayIndex !== -1) {
    const resultDate = new Date(now.getTime());
    const currentDay = resultDate.getDay();
    let daysUntilTarget = targetDayIndex - currentDay;
    if (daysUntilTarget <= 0) {
      daysUntilTarget += 7; // 오늘이 지나갔거나 오늘이면 다음 주 해당 요일로 지정
    }
    resultDate.setDate(resultDate.getDate() + daysUntilTarget);
    return resultDate;
  }
  
  // 5. 파싱이 완전히 불가능한 경우 기본 오늘/제출 날짜 반환
  return now;
}

/**
 * 성도가 입력한 다양한 텍스트에서 시간대(시작/종료)를 추출하여 24시간 형식 객체로 반환합니다.
 */
function parseTimeFromText(text) {
  if (!text) return null;
  
  const normalized = text.toLowerCase().replace(/\s+/g, " "); // 모든 공백을 단일 공백으로 표준화
  
  // 오전/오후 및 AM/PM 보정 오프셋 계산기
  const getAmPmOffset = (subText) => {
    if (subText.includes("오후") || subText.includes("pm") || subText.includes("저녁") || subText.includes("밤")) {
      return 12;
    }
    if (subText.includes("오전") || subText.includes("am") || subText.includes("아침") || subText.includes("새벽")) {
      return 0;
    }
    return -1;
  };

  // 1. [특수 매칭] "7-8pm", "오후 7-8", "7시-8시", "7:00-8:30" 등 대시/물결표 범위 형태 파싱
  // 6/6 같은 날짜를 시간으로 오인하지 않도록, 뒤에 시간 단위나 기호가 명확히 결합된 것만 매칭합니다.
  const rangeMatch = normalized.match(/(\d{1,2})(?::(\d{2}))?\s*(?:시|am|pm)?\s*[-~]\s*(\d{1,2})(?::(\d{2}))?\s*(시|분|pm|am)/i);
  if (rangeMatch) {
    let startHour = parseInt(rangeMatch[1]);
    let startMinute = rangeMatch[2] ? parseInt(rangeMatch[2]) : 0;
    let endHour = parseInt(rangeMatch[3]);
    let endMinute = rangeMatch[4] ? parseInt(rangeMatch[4]) : 0;
    const endUnit = rangeMatch[5];
    
    // 전체 문맥에서 AM/PM 키워드가 존재하는지 체크
    let offset = getAmPmOffset(normalized);
    
    if (offset === 12) {
      // 오후/PM인 경우 보정
      if (startHour < 12) startHour += 12;
      if (endHour < 12) endHour += 12;
    } else if (offset === -1) {
      // 아무 표시 없으면 교회 예약 휴리스틱 적용 (1~8시는 오후로 간주)
      if (startHour >= 1 && startHour <= 8) startHour += 12;
      if (endHour >= 1 && endHour <= 8) endHour += 12;
    }
    
    return {
      startHour: startHour,
      startMinute: startMinute,
      endHour: endHour,
      endMinute: endMinute
    };
  }

  // 2. [일반 매칭] 단독 시간 패턴 정규식: "오전 10:30", "오후 2시", "14시", "9:00" 등
  const times = [];
  const regex = /(오전|오후|am|pm)?\s*(\d{1,2})\s*(?:시|:)\s*(\d{2})?\s*(분)?/gi;
  let match;
  
  while ((match = regex.exec(normalized)) !== null) {
    let ampm = match[1] || "";
    let hour = parseInt(match[2]);
    let minute = match[3] ? parseInt(match[3]) : 0;
    
    let offset = getAmPmOffset(ampm);
    if (offset === -1) {
      const startIdx = Math.max(0, match.index - 10);
      const surrounding = normalized.substring(startIdx, match.index);
      offset = getAmPmOffset(surrounding);
    }
    
    if (offset === 12 && hour < 12) {
      hour += 12;
    } else if (offset === 0 && hour === 12) {
      hour = 0;
    } else if (offset === -1) {
      if (hour >= 1 && hour <= 8) {
        hour += 12;
      }
    }
    
    times.push({ hour, minute });
  }
  
  if (times.length >= 1) {
    const startTime = times[0];
    let endTime;
    
    if (times.length >= 2) {
      endTime = times[1];
    } else {
      // 종료 시간이 명시되지 않은 경우 기본 2시간 사용 설정
      endTime = {
        hour: (startTime.hour + 2) % 24,
        minute: startTime.minute
      };
    }
    return {
      startHour: startTime.hour,
      startMinute: startTime.minute,
      endHour: endTime.hour,
      endMinute: endTime.minute
    };
  }
  
  return null;
}

/**
 * 성도가 입력한 다양한 텍스트에서 날짜 목록을 지능적으로 파싱하여 반환합니다.
 */
function parseDatesFromText(text, referenceDate) {
  const now = referenceDate || new Date();
  const currentYear = now.getFullYear();
  const dates = [];

  // A. 날짜 범위 패턴 매칭: 예, 6/21~7/5, 2026.6.21 ~ 2026.7.5, 6/21-7/5
  let rangeMatch = text.match(/(\d{1,2})[./](\d{1,2})\s*[-~]\s*(\d{1,2})[./](\d{1,2})/);
  if (!rangeMatch) {
    rangeMatch = text.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*[-~]\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  }
  if (!rangeMatch) {
    rangeMatch = text.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s*[-~]\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  }
  
  if (rangeMatch) {
    let startYear = currentYear, endYear = currentYear;
    let startMonth, startDay, endMonth, endDay;
    
    if (rangeMatch.length === 7) {
      startYear = parseInt(rangeMatch[1]);
      startMonth = parseInt(rangeMatch[2]);
      startDay = parseInt(rangeMatch[3]);
      endYear = parseInt(rangeMatch[4]);
      endMonth = parseInt(rangeMatch[5]);
      endDay = parseInt(rangeMatch[6]);
    } else {
      startMonth = parseInt(rangeMatch[1]);
      startDay = parseInt(rangeMatch[2]);
      endMonth = parseInt(rangeMatch[3]);
      endDay = parseInt(rangeMatch[4]);
      
      // 연도 경계 처리 (예: 12/30 ~ 1/5)
      if (endMonth < startMonth) {
        endYear = startYear + 1;
      }
    }
    
    const startDate = new Date(startYear, startMonth - 1, startDay);
    const endDate = new Date(endYear, endMonth - 1, endDay);
    
    const normalized = text.toLowerCase().replace(/\s+/g, "");
    
    // 요일 검출
    const weekdaysKo = ["일", "월", "화", "수", "목", "금", "토"];
    const weekdaysEn = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
    let targetDays = [];
    
    if (normalized.includes("주일") || normalized.includes("일요일") || normalized.includes("일요")) {
      targetDays.push(0);
    }
    for (let d = 1; d < 7; d++) {
      if (normalized.includes(weekdaysKo[d] + "요일") || normalized.includes(weekdaysKo[d] + "요")) {
        targetDays.push(d);
      }
    }
    if (targetDays.length === 0) {
      for (let d = 0; d < 7; d++) {
        if (normalized.includes(weekdaysEn[d])) {
          targetDays.push(d);
          break;
        }
      }
    }
    
    // 범위 내 날짜 루프 생성
    let cur = new Date(startDate.getTime());
    let safetyCounter = 0;
    while (cur <= endDate && safetyCounter < 366) {
      safetyCounter++;
      if (targetDays.length === 0 || targetDays.includes(cur.getDay())) {
        dates.push(new Date(cur.getTime()));
      }
      cur.setDate(cur.getDate() + 1);
    }
    
    if (dates.length > 0) {
      return dates;
    }
  }

  // B. 쉼표 등으로 구분된 여러 개별 날짜 매칭
  const monthDayKoRegex = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/g;
  let matchKo;
  while ((matchKo = monthDayKoRegex.exec(text)) !== null) {
    const month = parseInt(matchKo[1]);
    const day = parseInt(matchKo[2]);
    dates.push(new Date(currentYear, month - 1, day));
  }
  
  const dateRegex = /\b(\d{1,2})[\/\.](\d{1,2})\b/g;
  let matchDate;
  while ((matchDate = dateRegex.exec(text)) !== null) {
    const month = parseInt(matchDate[1]);
    const day = parseInt(matchDate[2]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const alreadyAdded = dates.some(d => d.getMonth() === month - 1 && d.getDate() === day);
      if (!alreadyAdded) {
        dates.push(new Date(currentYear, month - 1, day));
      }
    }
  }
  
  // C. 검출된 날짜가 없을 경우, 단일 날짜 파싱 fallback 실행
  if (dates.length === 0) {
    const singleDate = parseDateFromText(text, now);
    if (singleDate) {
      dates.push(singleDate);
    }
  }
  
  return dates;
}

/**
 * 날짜 문자열과 타임스탬프를 이용해 시작/종료 일시를 가진 Slot 목록을 생성합니다.
 */
function getDateTimeSlots(dateTimeStr, timestamp) {
  const refDate = timestamp ? new Date(timestamp) : new Date();
  const dates = parseDatesFromText(dateTimeStr, refDate);
  const time = parseTimeFromText(dateTimeStr);
  
  const slots = [];
  for (let date of dates) {
    const start = new Date(date.getTime());
    const end = new Date(date.getTime());
    
    if (time) {
      start.setHours(time.startHour, time.startMinute, 0, 0);
      end.setHours(time.endHour, time.endMinute, 0, 0);
      if (end <= start) {
        end.setTime(start.getTime() + (2 * 60 * 60 * 1000));
      }
    } else {
      start.setHours(0, 0, 0, 0);
      end.setHours(23, 59, 59, 999);
    }
    slots.push({ start: start, end: end, hasTime: !!time });
  }
  return slots;
}

/**
 * 두 Slot 목록 간에 겹치는 시간대가 있는지 확인합니다.
 */
function doSlotsConflict(slotsA, slotsB) {
  for (let slotA of slotsA) {
    for (let slotB of slotsB) {
      if (slotA.start < slotB.end && slotB.start < slotA.end) {
        return true;
      }
    }
  }
  return false;
}

/**
 * 파싱된 정보와 신청서 원본 데이터를 종합하여 Google Calendar에 일정을 생성합니다.
 * 여러 날짜 슬롯이 검출되는 경우 루프를 돌며 개별 일정을 생성합니다.
 */
function addToGoogleCalendar(room, name, dateTimeStr, purpose, phone, participants, equipment) {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      Logger.log(`[캘린더 에러] ID '${CALENDAR_ID}'에 해당하는 캘린더를 찾을 수 없습니다.`);
      return false;
    }
    
    // G열에서 모든 날짜와 시간 슬롯 파싱
    const slots = getDateTimeSlots(dateTimeStr, new Date());
    if (slots.length === 0) {
      Logger.log(`[캘린더 에러] 날짜/시간 형식을 분석할 수 없습니다. (${dateTimeStr})`);
      return false;
    }
    
    // 일정 제목 포맷: [101호] 모임 목적
    const title = `[${room}] ${purpose}`;
    
    // 일정 본문 설명 작성
    const description = `■ 담당자 이름: ${name}
■ 사용 인원수: ${participants ? participants : '미정'} 명
■ 필요한 장비: ${equipment ? equipment : '없음'}
■ 사용 일자 및 시간: ${dateTimeStr}
■ 담당자 연락처: ${phone}
■ 시스템 승인 시각: ${new Date().toLocaleString()}
`;
    
    let createdCount = 0;
    for (let slot of slots) {
      let event;
      if (slot.hasTime) {
        event = calendar.createEvent(title, slot.start, slot.end, { description: description });
        Logger.log(`[일정 생성 완료] 시간 예약 일정: ${title} (${slot.start.toLocaleString()} ~ ${slot.end.toLocaleString()})`);
      } else {
        // 시간 추출 실패 시 하루 종일 일정(All-day Event)으로 등록하여 가독성 유지
        event = calendar.createAllDayEvent(title, slot.start, { description: description });
        Logger.log(`[일정 생성 완료] 하루 종일 예약 일정: ${title} (날짜: ${slot.start.toLocaleDateString()})`);
      }
      if (event) createdCount++;
    }
    
    // 모든 슬롯에 대한 일정이 정상적으로 생성되었는지 확인
    return createdCount === slots.length;
  } catch (error) {
    Logger.log("[캘린더 등록 예외 오류]: " + error.toString());
    return false;
  }
}
