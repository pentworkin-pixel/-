/**
 * 광고 종료 캘린더 동기화.gs
 *
 * Google Sheets의 "MM월 DD일 종료" 시트를 읽어 Google Calendar에 광고 종료 일정을
 * 자동 등록/갱신한다. Apps Script의 시간 기반 트리거로 Google 서버에서 실행되므로
 * PC/Mac이 꺼져 있거나 잠자기 상태여도 동작한다.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * 이 파일은 기존 프로젝트 코드와 절대 충돌하지 않도록 다음 규칙을 지킨다.
 *   - 모든 전역 식별자는 adEndSync / AD_END_SYNC 접두사를 사용한다.
 *   - 기존 파일을 수정하거나 참조하지 않는다.
 *   - 전역 스코프에서 실행되는 코드가 없다(상수 선언 제외).
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * 사용자가 직접 실행할 함수는 아래 6개뿐이다.
 *
 *   adEndSyncInstall()            최초 1회. 중복 트리거 정리 + 매일 08:30 트리거 생성 + 검증.
 *   adEndSyncRun()                트리거가 매일 자동 호출하는 동기화 본체. 수동 실행도 가능.
 *   adEndSyncVerifyTriggers()     현재 등록된 트리거 확인.
 *   adEndSyncDryRun()             캘린더를 건드리지 않고 시트 ↔ 캘린더 비교 결과만 출력.
 *   adEndSyncVerifyDuplicates()   중복 키워드가 있는 날짜의 행 수 == 캘린더 건수 검증.
 *   adEndSyncCheckColumns()       B/C/H 열 매핑이 맞는지 헤더와 샘플 행으로 확인.
 *
 * 캘린더 쓰기는 Calendar REST API(UrlFetchApp)를 사용한다. 이유:
 *   - CalendarApp에는 transparency(일정 있음/없음) 설정 API가 없다(요구사항 16).
 *   - 제목·설명·비공개·투명·팝업알림을 "한 번의 API 호출"로 한꺼번에 설정할 수 있어
 *     단시간 변경 제한(요구사항 17)에 훨씬 유리하다. CalendarApp은 속성마다 별도 호출이다.
 * 고급 서비스(Calendar)를 따로 켤 필요는 없다. OAuth 스코프는 아래 CalendarApp 참조로
 * 자동 감지된다(adEndSyncAssertCalendarAccess_ 참고).
 */

var AD_END_SYNC_CONFIG = {
  // ── 대상 ────────────────────────────────────────────────────────────────
  SPREADSHEET_ID: '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8',
  CALENDAR_ID: 'ksh94love@gmail.com',
  TIMEZONE: 'Asia/Seoul',

  // ── 시트 구조 ───────────────────────────────────────────────────────────
  SHEET_NAME_PATTERN: /^\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*종료\s*$/, // "MM월 DD일 종료"
  HEADER_ROWS: 1,        // 1행은 헤더
  COL_COMPANY: 2,        // B열: 업체명 (중복 키워드 구분용 괄호 표기)
  COL_KEYWORD: 3,        // C열: 종료 광고 키워드
  COL_END_DATE: 8,       // H열: 종료일 (일정 날짜의 기준)

  // ── 일정 ────────────────────────────────────────────────────────────────
  EVENT_HOUR: 9,             // 오전 9시 시작
  EVENT_MINUTE: 0,
  EVENT_DURATION_MIN: 30,    // 30분
  TITLE_PREFIX: '[광고 종료]',
  SHEETS_LABEL: '김승환 현황', // 설명 첫 줄 "Google Sheets: [...]"
  REMINDER_MINUTES: 0,       // 팝업 알림: 0 = 시작 시각(09:00) 정각
  VISIBILITY: 'private',     // 비공개
  TRANSPARENCY: 'transparent', // "일정 있음"에 영향 없음

  // ── 동작 ────────────────────────────────────────────────────────────────
  INCLUDE_TODAY: true,             // 오늘 종료분도 포함
  STOP_AT_FIRST_PAST_SHEET: true,  // 과거 날짜 탭을 만나면 그 뒤 탭은 읽지 않음
  SYNC_FUNCTION: 'adEndSyncRun',
  TRIGGER_HOUR: 8,
  TRIGGER_NEAR_MINUTE: 30
};


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. 설치 (최초 1회만 실행)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 최초 설치. 기존 중복 트리거를 모두 정리하고 매일 08:30(±15분) 트리거를 하나만 만든다.
 * 생성 후 실제로 트리거가 존재하는지 다시 조회해 검증한다.
 */
function adEndSyncInstall() {
  var cfg = AD_END_SYNC_CONFIG;
  var lines = [];

  // 1) 접근 권한부터 확인한다. 여기서 실패하면 트리거를 만들어도 매일 실패할 뿐이다.
  var ss = adEndSyncOpenSpreadsheet_();
  lines.push('스프레드시트 접근: OK (' + ss.getName() + ')');
  adEndSyncAssertCalendarAccess_();
  lines.push('캘린더 접근: OK (' + cfg.CALENDAR_ID + ')');

  // 2) 기존 동기화 트리거 전부 삭제 (중복 방지)
  var removed = adEndSyncRemoveTriggers_();
  lines.push('기존 트리거 정리: ' + removed + '개 삭제');

  // 3) 트리거 1개 생성
  ScriptApp.newTrigger(cfg.SYNC_FUNCTION)
    .timeBased()
    .atHour(cfg.TRIGGER_HOUR)
    .nearMinute(cfg.TRIGGER_NEAR_MINUTE)
    .everyDays(1)
    .inTimezone(cfg.TIMEZONE)
    .create();

  // 4) 생성 검증 — 만들었다고 믿지 않고 다시 조회한다.
  var mine = adEndSyncListTriggers_();
  if (mine.length !== 1) {
    throw new Error('트리거 생성 검증 실패: 기대 1개, 실제 ' + mine.length + '개');
  }
  lines.push('트리거 생성 검증: OK (트리거 ID ' + mine[0].getUniqueId() + ')');
  lines.push('자동 실행 시각: 매일 오전 ' + cfg.TRIGGER_HOUR + '시 ' +
             cfg.TRIGGER_NEAR_MINUTE + '분 전후 (' + cfg.TIMEZONE + ', ±15분 변동)');
  lines.push('생성된 트리거 수: ' + mine.length + '개');

  // 5) 설치 직후 1회 동기화 — 다음 날 08:30까지 기다리지 않아도 되도록.
  var result = adEndSyncRun();
  lines.push('');
  lines.push('── 설치 직후 1회 동기화 ──');
  lines.push(adEndSyncFormatReport_(result));

  var report = '=== 광고 종료 캘린더 동기화 설치 완료 ===\n' + lines.join('\n');
  adEndSyncLog_(report);
  return report;
}

/** 등록된 동기화 트리거를 확인한다. */
function adEndSyncVerifyTriggers() {
  var mine = adEndSyncListTriggers_();
  var lines = ['등록된 동기화 트리거: ' + mine.length + '개'];
  for (var i = 0; i < mine.length; i++) {
    lines.push('  - ID ' + mine[i].getUniqueId() +
               ' / 함수 ' + mine[i].getHandlerFunction() +
               ' / 유형 ' + mine[i].getEventType());
  }
  if (mine.length === 0) {
    lines.push('  ⚠ 트리거가 없습니다. adEndSyncInstall() 을 실행하세요.');
  } else if (mine.length > 1) {
    lines.push('  ⚠ 트리거가 중복되었습니다. adEndSyncInstall() 을 다시 실행하면 정리됩니다.');
  }
  var report = lines.join('\n');
  adEndSyncLog_(report);
  return report;
}

/** 이 스크립트가 만든 시간 기반 트리거 목록. */
function adEndSyncListTriggers_() {
  var all = ScriptApp.getProjectTriggers();
  var mine = [];
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === AD_END_SYNC_CONFIG.SYNC_FUNCTION) mine.push(all[i]);
  }
  return mine;
}

/** 동기화 트리거를 모두 삭제하고 삭제 개수를 반환한다. */
function adEndSyncRemoveTriggers_() {
  var mine = adEndSyncListTriggers_();
  for (var i = 0; i < mine.length; i++) ScriptApp.deleteTrigger(mine[i]);
  return mine.length;
}

/** 스크립트 삭제 없이 자동 실행만 중단하고 싶을 때. */
function adEndSyncUninstall() {
  var removed = adEndSyncRemoveTriggers_();
  var report = '동기화 트리거 ' + removed + '개를 삭제했습니다. 기존 캘린더 일정은 유지됩니다.';
  adEndSyncLog_(report);
  return report;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. 매일 자동 실행되는 동기화 본체
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 트리거가 매일 호출한다. 시트를 읽어 계획을 만들고 캘린더와 대조해
 * 신규 생성 / 제목·설명 갱신 / 유지 를 수행한다. 삭제는 절대 하지 않는다.
 */
function adEndSyncRun() {
  return adEndSyncExecute_(false);
}

/**
 * 캘린더를 전혀 변경하지 않고 무엇이 신규/갱신/유지인지만 계산한다.
 * 설치 전후 테스트용.
 */
function adEndSyncDryRun() {
  var result = adEndSyncExecute_(true);
  adEndSyncLog_('※ 위 결과는 미리보기(dry-run)이며 캘린더는 변경되지 않았습니다.');
  return result;
}

/**
 * @param {boolean} dryRun true면 캘린더에 쓰지 않는다.
 * @return {Object} 실행 결과 요약
 */
function adEndSyncExecute_(dryRun) {
  var cfg = AD_END_SYNC_CONFIG;
  var startedAt = new Date();
  var result = {
    dryRun: !!dryRun,
    실행시각: adEndSyncFormat_(startedAt, 'yyyy-MM-dd HH:mm:ss'),
    대상날짜수: 0,
    신규: 0,
    갱신: 0,
    유지: 0,
    실패: 0,
    오류: [],
    상세: []
  };

  var today = adEndSyncTodayKey_();
  var plan;
  try {
    plan = adEndSyncBuildPlan_(adEndSyncOpenSpreadsheet_(), today);
  } catch (e) {
    result.실패++;
    result.오류.push('시트 읽기 실패: ' + (e && e.message ? e.message : e));
    adEndSyncLog_(adEndSyncFormatReport_(result));
    return result;
  }

  result.대상날짜수 = plan.length;
  if (plan.length === 0) {
    adEndSyncLog_(adEndSyncFormatReport_(result));
    return result;
  }

  // 기간 전체의 기존 일정을 API 호출 1번(+페이지네이션)으로 가져온다.
  // 날짜마다 조회하면 호출 수가 날짜 수만큼 늘어나 변경 제한에 불리하다(요구사항 17).
  var existing;
  try {
    existing = adEndSyncFetchExistingEvents_(plan[0].dateKey, plan[plan.length - 1].dateKey);
  } catch (e) {
    result.실패++;
    result.오류.push('캘린더 조회 실패: ' + (e && e.message ? e.message : e));
    adEndSyncLog_(adEndSyncFormatReport_(result));
    return result;
  }

  for (var i = 0; i < plan.length; i++) {
    var day = plan[i];
    try {
      var ev = existing[day.dateKey];

      if (!ev) {
        // 규칙 13: 신규 종료일만 새 일정으로 생성
        if (!dryRun) adEndSyncInsertEvent_(day);
        result.신규++;
        result.상세.push(day.dateKey + ' 신규 (' + day.items.length + '건)');

      } else if (adEndSyncSameContent_(ev, day)) {
        // 규칙 11: 내용이 같으면 아무 작업도 하지 않는다 (API 호출 0회)
        result.유지++;
        result.상세.push(day.dateKey + ' 유지 (' + day.items.length + '건)');

      } else {
        // 규칙 12: 제목과 설명"만" 갱신. 시간/알림/공개범위는 건드리지 않는다.
        if (!dryRun) adEndSyncPatchEvent_(ev.id, day);
        result.갱신++;
        result.상세.push(day.dateKey + ' 갱신 (' + day.items.length + '건)');
      }
    } catch (e2) {
      result.실패++;
      result.오류.push(day.dateKey + ': ' + (e2 && e2.message ? e2.message : e2));
    }
  }

  result.소요초 = Math.round((new Date().getTime() - startedAt.getTime()) / 100) / 10;
  adEndSyncLog_(adEndSyncFormatReport_(result));
  return result;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. 시트 → 계획(plan) 생성  (순수 로직: 캘린더/네트워크에 의존하지 않음)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 종료 시트들을 읽어 날짜별 일정 계획을 만든다.
 *
 * 규칙:
 *   2) 오늘(포함) 이후의 종료 시트만 확인
 *   3) 시트가 최신→과거 순이므로, 과거 날짜 탭을 만나면 그 뒤는 읽지 않는다
 *   4) H열 종료일을 기준으로 C열 키워드를 수집
 *   5) 빈 키워드만 제외하고 시트의 행 순서를 그대로 유지
 *   6) 같은 키워드가 여러 행이면 중복 제거하지 않고 각각 1건으로 두되,
 *      B열 업체명을 괄호로 붙여 구분한다
 *
 * @param {Spreadsheet} ss
 * @param {string} todayKey 'yyyy-MM-dd'
 * @return {Array<Object>} 날짜 오름차순 계획 배열
 */
function adEndSyncBuildPlan_(ss, todayKey) {
  var cfg = AD_END_SYNC_CONFIG;
  var sheets = ss.getSheets();
  var byDate = {};   // dateKey -> [{keyword, company}]
  var order = [];    // dateKey 등장 순서

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetDateKey = adEndSyncParseSheetName_(sheet.getName(), todayKey);

    // "MM월 DD일 종료" 형식이 아닌 탭(요약/설정 등)은 건너뛴다. 중단 조건이 아니다.
    if (!sheetDateKey) continue;

    // 규칙 3: 과거 탭에 도달하면 더 오래된 탭은 읽지 않는다.
    if (sheetDateKey < todayKey) {
      if (cfg.STOP_AT_FIRST_PAST_SHEET) break;
      continue;
    }

    var values = sheet.getDataRange().getValues();
    for (var r = cfg.HEADER_ROWS; r < values.length; r++) {
      var row = values[r];

      // 규칙 5: 빈 키워드만 제외
      var keyword = adEndSyncCell_(row, cfg.COL_KEYWORD);
      if (keyword === '') continue;

      // 규칙 4: H열 종료일 기준. 비어 있으면 시트 이름의 날짜로 보완한다.
      var dateKey = adEndSyncNormalizeDate_(
        row.length >= cfg.COL_END_DATE ? row[cfg.COL_END_DATE - 1] : '', todayKey);
      if (!dateKey) dateKey = sheetDateKey;

      // 규칙 2: 오늘 이전 종료일은 제외
      if (cfg.INCLUDE_TODAY ? dateKey < todayKey : dateKey <= todayKey) continue;

      if (!byDate[dateKey]) { byDate[dateKey] = []; order.push(dateKey); }
      byDate[dateKey].push({
        keyword: keyword,
        company: adEndSyncCell_(row, cfg.COL_COMPANY),
        sheet: sheet.getName(),
        row: r + 1
      });
    }
  }

  order.sort(); // 'yyyy-MM-dd' 문자열 정렬 = 날짜 오름차순

  var plan = [];
  for (var i = 0; i < order.length; i++) {
    var key = order[i];
    var items = byDate[key];
    var labels = adEndSyncBuildLabels_(items);
    plan.push({
      dateKey: key,
      items: items,
      labels: labels,
      title: adEndSyncBuildTitle_(items.length),
      description: adEndSyncBuildDescription_(labels)
    });
  }
  return plan;
}

/**
 * 규칙 6: 같은 날짜 안에서 중복되는 키워드에만 " (업체명)"을 붙인다.
 * 중복 제거는 하지 않는다 — 2행이면 2개의 라벨이 나온다.
 */
function adEndSyncBuildLabels_(items) {
  var count = {};
  var i;
  for (i = 0; i < items.length; i++) {
    count[items[i].keyword] = (count[items[i].keyword] || 0) + 1;
  }
  var labels = [];
  for (i = 0; i < items.length; i++) {
    var it = items[i];
    if (count[it.keyword] > 1 && it.company !== '') {
      labels.push(it.keyword + ' (' + it.company + ')');
    } else {
      labels.push(it.keyword);
    }
  }
  return labels;
}

/** 규칙 8: "[광고 종료] 광고 N건" */
function adEndSyncBuildTitle_(n) {
  return AD_END_SYNC_CONFIG.TITLE_PREFIX + ' 광고 ' + n + '건';
}

/**
 * 규칙 9:
 *   Google Sheets: [김승환 현황]
 *   1. 첫 번째 키워드
 *   2. 두 번째 키워드
 */
function adEndSyncBuildDescription_(labels) {
  var lines = ['Google Sheets: [' + AD_END_SYNC_CONFIG.SHEETS_LABEL + ']'];
  for (var i = 0; i < labels.length; i++) lines.push((i + 1) + '. ' + labels[i]);
  return lines.join('\n');
}

/** 셀 값을 문자열로 정규화한다(양끝 공백 제거). */
function adEndSyncCell_(row, colIndex1Based) {
  if (row.length < colIndex1Based) return '';
  var v = row[colIndex1Based - 1];
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    // 예: B열이 업체명이 아니라 날짜로 채워져 있으면 괄호 표기를 생략한다.
    return '';
  }
  return String(v).replace(/^\s+|\s+$/g, '');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. 날짜 처리 (연도가 없는 "MM월 DD일" 대응)
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 오늘 날짜 키('yyyy-MM-dd', Asia/Seoul). */
function adEndSyncTodayKey_() {
  return adEndSyncFormat_(new Date(), 'yyyy-MM-dd');
}

function adEndSyncFormat_(date, pattern) {
  return Utilities.formatDate(date, AD_END_SYNC_CONFIG.TIMEZONE, pattern);
}

/** "MM월 DD일 종료" → 'yyyy-MM-dd'. 형식이 아니면 null. */
function adEndSyncParseSheetName_(name, todayKey) {
  var m = AD_END_SYNC_CONFIG.SHEET_NAME_PATTERN.exec(name);
  if (!m) return null;
  return adEndSyncKeyFromMonthDay_(Number(m[1]), Number(m[2]), todayKey);
}

/**
 * 연도가 없는 월/일에 연도를 추론한다.
 * 오늘 기준 6개월 이상 과거면 내년, 6개월 이상 미래면 작년으로 본다.
 * (12월 → 1월 연말 전환을 자연스럽게 처리하기 위함)
 */
function adEndSyncKeyFromMonthDay_(month, day, todayKey) {
  if (!(month >= 1 && month <= 12) || !(day >= 1 && day <= 31)) return null;
  var parts = todayKey.split('-');
  var todayY = Number(parts[0]);
  var todayUtc = Date.UTC(todayY, Number(parts[1]) - 1, Number(parts[2]));

  var best = null, bestDiff = null;
  for (var offset = -1; offset <= 1; offset++) {
    var y = todayY + offset;
    var d = new Date(Date.UTC(y, month - 1, day));
    if (d.getUTCMonth() !== month - 1) continue; // 2월 30일 같은 무효 날짜
    var diff = Math.abs(d.getTime() - todayUtc);
    // 동률이면 미래를 우선한다.
    if (bestDiff === null || diff < bestDiff ||
        (diff === bestDiff && d.getTime() > todayUtc)) {
      bestDiff = diff;
      best = y;
    }
  }
  if (best === null) return null;
  return adEndSyncPad_(best, 4) + '-' + adEndSyncPad_(month, 2) + '-' + adEndSyncPad_(day, 2);
}

/** H열 값(Date / 숫자 시리얼 / 문자열)을 'yyyy-MM-dd'로 정규화한다. */
function adEndSyncNormalizeDate_(value, todayKey) {
  if (value === null || value === undefined || value === '') return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return adEndSyncFormat_(value, 'yyyy-MM-dd');
  }

  if (typeof value === 'number' && isFinite(value)) {
    // 스프레드시트 날짜 시리얼(1899-12-30 기준)
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
    return adEndSyncPad_(d.getUTCFullYear(), 4) + '-' +
           adEndSyncPad_(d.getUTCMonth() + 1, 2) + '-' +
           adEndSyncPad_(d.getUTCDate(), 2);
  }

  var s = String(value).replace(/^\s+|\s+$/g, '');
  if (s === '') return null;

  var m = /^(\d{4})[-.\/\s]+(\d{1,2})[-.\/\s]+(\d{1,2})/.exec(s);
  if (m) {
    return adEndSyncPad_(Number(m[1]), 4) + '-' +
           adEndSyncPad_(Number(m[2]), 2) + '-' + adEndSyncPad_(Number(m[3]), 2);
  }

  m = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(s);
  if (m) return adEndSyncKeyFromMonthDay_(Number(m[1]), Number(m[2]), todayKey);

  m = /^(\d{1,2})[-.\/](\d{1,2})$/.exec(s);
  if (m) return adEndSyncKeyFromMonthDay_(Number(m[1]), Number(m[2]), todayKey);

  return null;
}

function adEndSyncPad_(n, width) {
  var s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. Google Calendar 연동 (REST API)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 스코프 자동 감지를 위해 CalendarApp을 사용한다. 이 참조 덕분에
 * appsscript.json을 손대지 않아도 캘린더 권한이 요청된다.
 */
function adEndSyncAssertCalendarAccess_() {
  var cal = CalendarApp.getCalendarById(AD_END_SYNC_CONFIG.CALENDAR_ID);
  if (!cal) {
    throw new Error('캘린더에 접근할 수 없습니다: ' + AD_END_SYNC_CONFIG.CALENDAR_ID +
                    ' — 이 스크립트를 해당 계정으로 승인했는지 확인하세요.');
  }
  return cal;
}

function adEndSyncCalendarApi_(method, path, payload, params) {
  var url = 'https://www.googleapis.com/calendar/v3' + path;
  if (params) {
    var qs = [];
    for (var k in params) {
      if (params[k] === null || params[k] === undefined) continue;
      qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
    }
    if (qs.length) url += '?' + qs.join('&');
  }

  var options = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json; charset=UTF-8';
    options.payload = JSON.stringify(payload);
  }

  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code >= 200 && code < 300) return body ? JSON.parse(body) : {};
  throw new Error('Calendar API ' + method + ' ' + path + ' 실패 (HTTP ' + code + '): ' + body);
}

function adEndSyncCalendarPath_() {
  return '/calendars/' + encodeURIComponent(AD_END_SYNC_CONFIG.CALENDAR_ID) + '/events';
}

/** 'yyyy-MM-dd' + 시:분 → RFC3339 문자열 (Asia/Seoul 오프셋 포함) */
function adEndSyncRfc3339_(dateKey, hour, minute) {
  return dateKey + 'T' + adEndSyncPad_(hour, 2) + ':' + adEndSyncPad_(minute, 2) + ':00' +
         adEndSyncTzOffset_(dateKey);
}

/** 해당 날짜의 타임존 오프셋('+09:00' 형태). */
function adEndSyncTzOffset_(dateKey) {
  var p = dateKey.split('-');
  var noonUtc = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12));
  var raw = adEndSyncFormat_(noonUtc, 'ZZ'); // 예: '+0900'
  return raw.slice(0, 3) + ':' + raw.slice(3);
}

/**
 * 기간 내의 "[광고 종료]"로 시작하는 일정을 날짜별로 1건씩 반환한다.
 * 규칙 10: 같은 날짜에 이미 있으면 그 일정을 재사용(=중복 생성 금지)한다.
 *
 * @return {Object} dateKey -> {id, summary, description}
 */
function adEndSyncFetchExistingEvents_(firstDateKey, lastDateKey) {
  var cfg = AD_END_SYNC_CONFIG;
  var map = {};
  var pageToken = null;
  var guard = 0;

  do {
    var res = adEndSyncCalendarApi_('get', adEndSyncCalendarPath_(), null, {
      timeMin: adEndSyncRfc3339_(firstDateKey, 0, 0),
      timeMax: adEndSyncRfc3339_(lastDateKey, 23, 59),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 2500,
      showDeleted: false,
      pageToken: pageToken
    });

    var items = res.items || [];
    for (var i = 0; i < items.length; i++) {
      var ev = items[i];
      var summary = ev.summary || '';
      if (summary.indexOf(cfg.TITLE_PREFIX) !== 0) continue;

      var key = adEndSyncEventDateKey_(ev);
      if (!key) continue;
      // orderBy=startTime이므로 먼저 나온 것이 그 날의 첫 일정이다.
      if (!map[key]) {
        map[key] = { id: ev.id, summary: summary, description: ev.description || '' };
      }
    }
    pageToken = res.nextPageToken || null;
  } while (pageToken && ++guard < 20);

  return map;
}

function adEndSyncEventDateKey_(ev) {
  if (!ev.start) return null;
  if (ev.start.date) return ev.start.date;            // 종일 일정
  if (!ev.start.dateTime) return null;
  return adEndSyncFormat_(new Date(ev.start.dateTime), 'yyyy-MM-dd');
}

/**
 * 규칙 7·15·16: 09:00~09:30, 팝업 알림, 비공개, 투명(일정 있음에 영향 없음).
 * 모든 속성을 단 한 번의 insert 호출로 설정한다(요구사항 17).
 */
function adEndSyncInsertEvent_(day) {
  var cfg = AD_END_SYNC_CONFIG;
  var endMinuteTotal = cfg.EVENT_HOUR * 60 + cfg.EVENT_MINUTE + cfg.EVENT_DURATION_MIN;

  return adEndSyncCalendarApi_('post', adEndSyncCalendarPath_(), {
    summary: day.title,
    description: day.description,
    start: {
      dateTime: adEndSyncRfc3339_(day.dateKey, cfg.EVENT_HOUR, cfg.EVENT_MINUTE),
      timeZone: cfg.TIMEZONE
    },
    end: {
      dateTime: adEndSyncRfc3339_(day.dateKey,
                                  Math.floor(endMinuteTotal / 60), endMinuteTotal % 60),
      timeZone: cfg.TIMEZONE
    },
    visibility: cfg.VISIBILITY,
    transparency: cfg.TRANSPARENCY,
    reminders: {
      useDefault: false,
      overrides: [{ method: 'popup', minutes: cfg.REMINDER_MINUTES }]
    }
  });
}

/**
 * 규칙 12·17: 제목과 설명만 PATCH한다. 시간·알림·공개범위·투명도는 보내지 않으므로
 * 사용자가 캘린더에서 손수 조정한 값이 덮어써지지 않고, 변경 호출도 최소가 된다.
 */
function adEndSyncPatchEvent_(eventId, day) {
  return adEndSyncCalendarApi_(
    'patch',
    adEndSyncCalendarPath_() + '/' + encodeURIComponent(eventId),
    { summary: day.title, description: day.description }
  );
}

/** 규칙 11: 제목·설명이 모두 같으면 아무것도 하지 않는다. */
function adEndSyncSameContent_(ev, day) {
  return adEndSyncNormalizeText_(ev.summary) === adEndSyncNormalizeText_(day.title) &&
         adEndSyncNormalizeText_(ev.description) === adEndSyncNormalizeText_(day.description);
}

/** 캘린더가 줄바꿈을 \r\n으로 돌려주는 경우가 있어 비교 전에 정규화한다. */
function adEndSyncNormalizeText_(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/\r\n/g, '\n').replace(/\s+$/, '');
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. 검증 도구
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * B/C/H 열 매핑 확인용. 첫 번째 대상 시트의 헤더와 샘플 3행을 출력한다.
 * (요구사항에 B열이 "작업 종료일"과 "업체명" 두 가지로 적혀 있어 확인이 필요하다.)
 */
function adEndSyncCheckColumns() {
  var cfg = AD_END_SYNC_CONFIG;
  var ss = adEndSyncOpenSpreadsheet_();
  var todayKey = adEndSyncTodayKey_();
  var sheets = ss.getSheets();
  var lines = ['=== 열 매핑 확인 (스프레드시트: ' + ss.getName() + ') ==='];

  for (var s = 0; s < sheets.length; s++) {
    var key = adEndSyncParseSheetName_(sheets[s].getName(), todayKey);
    if (!key || key < todayKey) continue;

    var values = sheets[s].getDataRange().getValues();
    lines.push('시트: ' + sheets[s].getName() + ' (해석된 날짜 ' + key + ')');
    if (values.length > 0) {
      lines.push('  헤더 B/C/H: ' +
        JSON.stringify([adEndSyncCell_(values[0], cfg.COL_COMPANY),
                        adEndSyncCell_(values[0], cfg.COL_KEYWORD),
                        adEndSyncCell_(values[0], cfg.COL_END_DATE)]));
    }
    for (var r = cfg.HEADER_ROWS; r < Math.min(values.length, cfg.HEADER_ROWS + 3); r++) {
      lines.push('  ' + (r + 1) + '행 → B(업체명)=' + adEndSyncCell_(values[r], cfg.COL_COMPANY) +
                 ' / C(키워드)=' + adEndSyncCell_(values[r], cfg.COL_KEYWORD) +
                 ' / H(종료일)=' + adEndSyncNormalizeDate_(
                     values[r][cfg.COL_END_DATE - 1], todayKey));
    }
    break;
  }
  if (lines.length === 1) lines.push('오늘 이후의 "MM월 DD일 종료" 시트를 찾지 못했습니다.');

  var report = lines.join('\n');
  adEndSyncLog_(report);
  return report;
}

/**
 * 규칙 6 검증: 중복 키워드가 있는 날짜를 골라
 * "시트 행 수 == 캘린더 일정의 항목 수"인지 확인한다.
 * 캘린더를 변경하지 않는다.
 */
function adEndSyncVerifyDuplicates() {
  var plan = adEndSyncBuildPlan_(adEndSyncOpenSpreadsheet_(), adEndSyncTodayKey_());
  var lines = ['=== 중복 키워드 검증 ==='];
  var targets = [];
  var i;

  for (i = 0; i < plan.length; i++) {
    var seen = {}, dup = false;
    for (var j = 0; j < plan[i].items.length; j++) {
      var k = plan[i].items[j].keyword;
      if (seen[k]) { dup = true; break; }
      seen[k] = true;
    }
    if (dup) targets.push(plan[i]);
  }

  if (targets.length === 0) {
    lines.push('오늘 이후 대상 중 중복 키워드가 있는 날짜가 없습니다. (검증 대상 없음)');
    var noneReport = lines.join('\n');
    adEndSyncLog_(noneReport);
    return noneReport;
  }

  var existing = adEndSyncFetchExistingEvents_(plan[0].dateKey, plan[plan.length - 1].dateKey);
  var pass = 0, fail = 0;

  for (i = 0; i < targets.length; i++) {
    var day = targets[i];
    var rows = day.items.length;
    var ev = existing[day.dateKey];
    var calCount = ev ? adEndSyncCountDescriptionItems_(ev.description) : null;
    var titleCount = ev ? adEndSyncCountTitleItems_(ev.summary) : null;

    var ok = ev !== undefined && ev !== null && calCount === rows && titleCount === rows;
    if (ok) pass++; else fail++;

    lines.push((ok ? '✅' : '❌') + ' ' + day.dateKey +
               ' | 시트 행 수 ' + rows +
               ' | 캘린더 제목 ' + (titleCount === null ? '(일정 없음)' : titleCount + '건') +
               ' | 캘린더 설명 항목 ' + (calCount === null ? '-' : calCount + '개'));
    for (var d = 0; d < day.labels.length; d++) {
      lines.push('     ' + (d + 1) + '. ' + day.labels[d] +
                 '   ← ' + day.items[d].sheet + ' ' + day.items[d].row + '행');
    }
  }
  lines.push('검증 결과: 통과 ' + pass + '건 / 불일치 ' + fail + '건');
  if (fail > 0) lines.push('※ 불일치는 대부분 아직 동기화하지 않았기 때문입니다. adEndSyncRun() 실행 후 다시 확인하세요.');

  var report = lines.join('\n');
  adEndSyncLog_(report);
  return report;
}

function adEndSyncCountDescriptionItems_(description) {
  var lines = adEndSyncNormalizeText_(description).split('\n');
  var n = 0;
  for (var i = 0; i < lines.length; i++) if (/^\s*\d+\.\s/.test(lines[i])) n++;
  return n;
}

function adEndSyncCountTitleItems_(summary) {
  var m = /광고\s*(\d+)\s*건/.exec(summary || '');
  return m ? Number(m[1]) : null;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 7. 공통 유틸
 * ═══════════════════════════════════════════════════════════════════════════ */

function adEndSyncOpenSpreadsheet_() {
  try {
    var ss = SpreadsheetApp.openById(AD_END_SYNC_CONFIG.SPREADSHEET_ID);
    if (ss) return ss;
  } catch (e) {
    // @OnlyCurrentDoc 이 걸린 프로젝트라면 openById가 막힌다. 컨테이너 문서로 대체한다.
  }
  var active = SpreadsheetApp.getActive();
  if (active) return active;
  throw new Error('스프레드시트를 열 수 없습니다: ' + AD_END_SYNC_CONFIG.SPREADSHEET_ID);
}

function adEndSyncFormatReport_(r) {
  var lines = [];
  lines.push('실행 시각: ' + r.실행시각 + ' (' + AD_END_SYNC_CONFIG.TIMEZONE + ')' +
             (r.dryRun ? ' [미리보기]' : ''));
  lines.push('대상 종료일: ' + r.대상날짜수 + '일');
  lines.push('신규 ' + r.신규 + '건 / 갱신 ' + r.갱신 + '건 / 유지 ' + r.유지 + '건 / 실패 ' + r.실패 + '건');
  if (r.소요초 !== undefined) lines.push('소요 시간: ' + r.소요초 + '초');
  if (r.상세 && r.상세.length) lines.push('상세: ' + r.상세.join(', '));
  if (r.오류 && r.오류.length) lines.push('오류:\n  - ' + r.오류.join('\n  - '));
  return lines.join('\n');
}

function adEndSyncLog_(message) {
  Logger.log(message);
  try { console.log(message); } catch (e) {}
}
