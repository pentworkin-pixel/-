/**
 * 프로그램 동기화.gs
 *
 * A시트(김승환 현황)의 프로그램명·갯수를 읽어 B시트(펀트미디어 순위추적)의 E열에
 * "프로그램명 갯수" 형태로 자동 반영한다.
 *
 *   A시트  D열 = 프로그램명 / E열 = 갯수 / Q열 = 상품 식별 숫자
 *   B시트  K열 = 상품 식별 숫자 / E열 = 결과가 기록될 열   ← 이 열 하나만 건드린다
 *
 * Apps Script의 시간 기반 트리거로 Google 서버에서 매일 10시에 실행되므로
 * PC/Mac이 꺼져 있어도, 사용자가 아무것도 하지 않아도 동작한다.
 * (A시트는 매일 9시에 갱신되므로 그보다 뒤인 10시에 읽는다.)
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * 이 파일은 기존 프로젝트 코드와 절대 충돌하지 않도록 다음 규칙을 지킨다.
 *   - 모든 전역 식별자는 progSync / PROG_SYNC 접두사를 사용한다.
 *   - 기존 파일을 수정하거나 참조하지 않는다.
 *   - 전역 스코프에서 실행되는 코드가 없다(상수 선언 제외).
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * 사용자가 직접 실행할 함수는 아래 6개뿐이다.
 *
 *   progSyncInstall()         최초 1회. 중복 트리거 정리 + 매일 10시 트리거 생성 + 검증 + 즉시 1회 실행.
 *   progSyncRun()             트리거가 매일 자동 호출하는 동기화 본체. 수동 실행도 가능.
 *   progSyncCheckColumns()    A의 D/E/Q, B의 K/E 열 매핑이 맞는지 헤더·샘플 행으로 확인.
 *   progSyncDryRun()          B시트를 전혀 건드리지 않고 "무엇이 어떻게 바뀔지"만 출력.
 *   progSyncVerifyTriggers()  트리거가 정확히 1개 등록됐는지 확인.
 *   progSyncVerifyMatch()     동기화 후 B시트를 다시 읽어 E열이 기대값과 같은지 대조.
 *
 *   progSyncUninstall()       자동 실행만 중단(트리거 삭제). 시트 값은 그대로 둔다.
 */

var PROG_SYNC_CONFIG = {
  // ── 대상 ────────────────────────────────────────────────────────────────
  // A시트 (읽기 전용)
  SOURCE_SPREADSHEET_ID: '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8',
  SOURCE_SHEET_GID: 1948367989,
  SOURCE_LABEL: 'A시트',

  // B시트 (E열만 쓰기)
  TARGET_SPREADSHEET_ID: '1bLNh-zrYHHKWgH78ihbgnhpd11ItUhFgw2eT0MromFU',
  TARGET_SHEET_GID: 1003701754,
  TARGET_LABEL: 'B시트',

  TIMEZONE: 'Asia/Seoul',

  // ── 열 매핑 (A=1, B=2, ... E=5, K=11, Q=17) ─────────────────────────────
  SOURCE_HEADER_ROWS: 1,
  SOURCE_COL_PROGRAM: 4,   // D열: 프로그램명
  SOURCE_COL_COUNT: 5,     // E열: 프로그램 갯수
  SOURCE_COL_KEY: 17,      // Q열: 상품 식별 숫자

  TARGET_HEADER_ROWS: 1,
  TARGET_COL_KEY: 11,      // K열: 상품 식별 숫자
  TARGET_COL_OUTPUT: 5,    // E열: 프로그램명 + 갯수가 기록될 열

  // ── 출력 형식 ───────────────────────────────────────────────────────────
  // 한 상품에 프로그램이 여러 개면 SEPARATOR로 이어 붙인다.
  //   예: "리뷰 3개, 트래픽 2개"
  // 대괄호 표기를 원하면 PAIR_FORMAT을 '[{name},{count}]' 로 바꾸면 된다.
  PAIR_FORMAT: '{name} {count}개',
  PAIR_FORMAT_NO_COUNT: '{name}',   // 갯수가 비어 있을 때
  PAIR_SEPARATOR: ', ',
  MERGE_SAME_PROGRAM: true,         // 같은 상품에 같은 프로그램명이 여러 줄이면 갯수를 합산

  // ── 안전장치 ────────────────────────────────────────────────────────────
  // A시트에 없는 상품(=매칭 실패)의 B시트 E열은 기본적으로 "건드리지 않는다".
  // true로 바꾸면 매칭이 없을 때 E열을 빈 값으로 지운다.
  CLEAR_WHEN_NO_MATCH: false,

  // ── 자동 실행 ───────────────────────────────────────────────────────────
  SYNC_FUNCTION: 'progSyncRun',
  TRIGGER_HOUR: 10,          // 매일 오전 10시 (A시트 9시 갱신 이후)
  TRIGGER_NEAR_MINUTE: 0
};


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. 설치 (최초 1회만 실행)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 최초 설치. 기존 중복 트리거를 모두 정리하고 매일 10시(±15분) 트리거를 하나만 만든다.
 * 생성 후 실제로 트리거가 존재하는지 다시 조회해 검증하고, 설치 직후 1회 동기화한다.
 */
function progSyncInstall() {
  var cfg = PROG_SYNC_CONFIG;
  var lines = [];

  // 1) 접근 권한부터 확인한다. 여기서 실패하면 트리거를 만들어도 매일 실패할 뿐이다.
  var src = progSyncOpenSheet_(cfg.SOURCE_SPREADSHEET_ID, cfg.SOURCE_SHEET_GID, cfg.SOURCE_LABEL);
  lines.push('A시트 읽기 권한: OK (' + src.ss.getName() + ' / ' + src.sheet.getName() + ')');

  var tgt = progSyncOpenSheet_(cfg.TARGET_SPREADSHEET_ID, cfg.TARGET_SHEET_GID, cfg.TARGET_LABEL);
  progSyncAssertWritable_(tgt.sheet);
  lines.push('B시트 쓰기 권한: OK (' + tgt.ss.getName() + ' / ' + tgt.sheet.getName() + ')');

  // 2) 기존 동기화 트리거 전부 삭제 (중복 방지)
  var removed = progSyncRemoveTriggers_();
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
  var mine = progSyncListTriggers_();
  if (mine.length !== 1) {
    throw new Error('트리거 생성 검증 실패: 기대 1개, 실제 ' + mine.length + '개');
  }
  lines.push('트리거 생성 검증: OK (트리거 ID ' + mine[0].getUniqueId() + ')');
  lines.push('자동 실행 시각: 매일 오전 ' + cfg.TRIGGER_HOUR + '시 ' +
             cfg.TRIGGER_NEAR_MINUTE + '분 전후 (' + cfg.TIMEZONE + ', ±15분 변동)');
  lines.push('생성된 트리거 수: ' + mine.length + '개');

  // 5) 설치 직후 1회 동기화 — 다음 날 10시까지 기다리지 않아도 되도록.
  var result = progSyncRun();
  lines.push('');
  lines.push('── 설치 직후 1회 동기화 ──');
  lines.push(progSyncFormatReport_(result));

  var report = '=== 프로그램 동기화 설치 완료 ===\n' + lines.join('\n');
  progSyncLog_(report);
  return report;
}

/** 자동 실행을 중단한다. 트리거만 삭제하고 시트 값은 건드리지 않는다. */
function progSyncUninstall() {
  var removed = progSyncRemoveTriggers_();
  var report = '프로그램 동기화 트리거 ' + removed + '개를 삭제했습니다. ' +
               'B시트의 기존 값은 그대로 남아 있습니다.';
  progSyncLog_(report);
  return report;
}

/** 등록된 동기화 트리거를 확인한다. */
function progSyncVerifyTriggers() {
  var mine = progSyncListTriggers_();
  var lines = ['등록된 동기화 트리거: ' + mine.length + '개'];
  for (var i = 0; i < mine.length; i++) {
    lines.push('  ' + (i + 1) + '. ID ' + mine[i].getUniqueId() +
               ' → ' + mine[i].getHandlerFunction() + '()');
  }
  if (mine.length !== 1) {
    lines.push('⚠️ 트리거가 1개가 아닙니다. progSyncInstall() 을 다시 실행하세요.');
  }
  var report = lines.join('\n');
  progSyncLog_(report);
  return report;
}

/** 이 스크립트가 만든 트리거만 골라낸다. */
function progSyncListTriggers_() {
  var all = ScriptApp.getProjectTriggers();
  var mine = [];
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === PROG_SYNC_CONFIG.SYNC_FUNCTION) mine.push(all[i]);
  }
  return mine;
}

/** 이 스크립트가 만든 트리거를 모두 삭제하고 삭제 개수를 반환한다. */
function progSyncRemoveTriggers_() {
  var mine = progSyncListTriggers_();
  for (var i = 0; i < mine.length; i++) ScriptApp.deleteTrigger(mine[i]);
  return mine.length;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. 동기화 본체
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 트리거가 매일 호출하는 함수. A시트를 읽어 B시트 E열을 갱신한다.
 * B시트에서 실제로 값이 달라지는 셀만 쓴다(변경 없으면 쓰기 호출 0회).
 */
function progSyncRun() {
  var result = progSyncPlan_();

  var runs = progSyncGroupRuns_(result.updates);
  for (var i = 0; i < runs.length; i++) {
    var run = runs[i];
    result.targetSheet
      .getRange(run.startRow, PROG_SYNC_CONFIG.TARGET_COL_OUTPUT, run.values.length, 1)
      .setValues(run.values);
  }
  if (runs.length > 0) SpreadsheetApp.flush();

  result.writeCalls = runs.length;
  progSyncLog_(progSyncFormatReport_(result));
  return result;
}

/** B시트를 전혀 건드리지 않고 무엇이 어떻게 바뀔지만 출력한다. */
function progSyncDryRun() {
  var result = progSyncPlan_();
  var lines = ['=== 프로그램 동기화 미리보기 (B시트 변경 없음) ===',
               progSyncFormatReport_(result)];

  var shown = Math.min(result.updates.length, 20);
  if (shown > 0) lines.push('', '── 변경될 셀 (최대 20건) ──');
  for (var i = 0; i < shown; i++) {
    var u = result.updates[i];
    lines.push('  E' + u.row + ' (키 ' + u.key + '): "' + u.current + '" → "' + u.next + '"');
  }
  if (result.updates.length > shown) {
    lines.push('  ... 외 ' + (result.updates.length - shown) + '건');
  }

  var missing = result.unmatchedKeys.slice(0, 10);
  if (missing.length > 0) {
    lines.push('', '── A시트에서 찾지 못한 B시트 키 (최대 10건) ──');
    for (var j = 0; j < missing.length; j++) lines.push('  ' + missing[j]);
    if (result.unmatchedKeys.length > missing.length) {
      lines.push('  ... 외 ' + (result.unmatchedKeys.length - missing.length) + '건');
    }
  }

  var report = lines.join('\n');
  progSyncLog_(report);
  return report;
}

/**
 * 읽기만 해서 "무엇을 써야 하는지" 계획을 만든다. progSyncRun / progSyncDryRun 공용.
 * 쓰기는 하지 않는다.
 */
function progSyncPlan_() {
  var cfg = PROG_SYNC_CONFIG;

  var src = progSyncOpenSheet_(cfg.SOURCE_SPREADSHEET_ID, cfg.SOURCE_SHEET_GID, cfg.SOURCE_LABEL);
  var tgt = progSyncOpenSheet_(cfg.TARGET_SPREADSHEET_ID, cfg.TARGET_SHEET_GID, cfg.TARGET_LABEL);

  var sourceValues = progSyncReadRows_(src.sheet, cfg.SOURCE_HEADER_ROWS);
  var targetValues = progSyncReadRows_(tgt.sheet, cfg.TARGET_HEADER_ROWS);

  var built = progSyncBuildProgramMap_(sourceValues, cfg);
  var planned = progSyncComputeUpdates_(targetValues, built.map, cfg);

  return {
    targetSheet: tgt.sheet,
    sourceName: src.ss.getName() + ' / ' + src.sheet.getName(),
    targetName: tgt.ss.getName() + ' / ' + tgt.sheet.getName(),
    sourceRows: sourceValues.length,
    sourceUsedRows: built.usedRows,
    sourceKeys: built.keyCount,
    targetRows: targetValues.length,
    matched: planned.matched,
    unchanged: planned.matched - planned.updates.length,
    updates: planned.updates,
    blankKeyRows: planned.blankKeyRows,
    unmatchedKeys: planned.unmatchedKeys,
    writeCalls: 0
  };
}

/** 헤더를 제외한 데이터 영역을 2차원 배열로 읽는다. 데이터가 없으면 빈 배열. */
function progSyncReadRows_(sheet, headerRows) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= headerRows || lastCol < 1) return [];
  return sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol).getValues();
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. 순수 로직 (스프레드시트 API 없이 동작 — 로컬 테스트 대상)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * A시트 데이터에서 { 상품키 → "프로그램명 갯수, ..." } 맵을 만든다.
 * 시트에 나온 순서를 그대로 유지한다.
 */
function progSyncBuildProgramMap_(rows, cfg) {
  var order = [];      // 키 등장 순서
  var buckets = {};    // 키 → [{name, countText, countNum}]
  var usedRows = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var key = progSyncNormalizeKey_(progSyncCell_(row, cfg.SOURCE_COL_KEY));
    var name = progSyncTrim_(progSyncCell_(row, cfg.SOURCE_COL_PROGRAM));
    if (key === '' || name === '') continue;   // 키나 프로그램명이 비면 무시

    var count = progSyncParseCount_(progSyncCell_(row, cfg.SOURCE_COL_COUNT));
    usedRows++;

    if (!Object.prototype.hasOwnProperty.call(buckets, key)) {
      buckets[key] = [];
      order.push(key);
    }
    var list = buckets[key];

    // 같은 상품에 같은 프로그램명이 또 나오면 갯수를 합친다(둘 다 숫자일 때만).
    var merged = false;
    if (cfg.MERGE_SAME_PROGRAM) {
      for (var j = 0; j < list.length; j++) {
        if (list[j].name !== name) continue;
        if (list[j].countNum !== null && count.countNum !== null) {
          list[j].countNum += count.countNum;
          list[j].countText = String(list[j].countNum);
          merged = true;
        }
        break;
      }
    }
    if (!merged) list.push({ name: name, countText: count.countText, countNum: count.countNum });
  }

  var map = {};
  for (var k = 0; k < order.length; k++) {
    map[order[k]] = progSyncFormatEntries_(buckets[order[k]], cfg);
  }
  return { map: map, keyCount: order.length, usedRows: usedRows };
}

/** [{name, countText}] → "리뷰 3개, 트래픽 2개" */
function progSyncFormatEntries_(entries, cfg) {
  var parts = [];
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    var tpl = e.countText === '' ? cfg.PAIR_FORMAT_NO_COUNT : cfg.PAIR_FORMAT;
    parts.push(tpl.split('{name}').join(e.name).split('{count}').join(e.countText));
  }
  return parts.join(cfg.PAIR_SEPARATOR);
}

/**
 * B시트 데이터와 프로그램 맵을 대조해 실제로 값이 달라지는 셀만 골라낸다.
 * 값이 이미 같으면 updates에 넣지 않는다(쓰기 호출을 아끼기 위해).
 */
function progSyncComputeUpdates_(rows, map, cfg) {
  var updates = [];
  var matched = 0;
  var blankKeyRows = 0;
  var unmatchedKeys = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var absRow = i + cfg.TARGET_HEADER_ROWS + 1;   // 시트상의 실제 행 번호
    var key = progSyncNormalizeKey_(progSyncCell_(row, cfg.TARGET_COL_KEY));
    if (key === '') { blankKeyRows++; continue; }

    var has = Object.prototype.hasOwnProperty.call(map, key);
    if (!has) {
      unmatchedKeys.push(key);
      if (!cfg.CLEAR_WHEN_NO_MATCH) continue;      // 기본값: 건드리지 않는다
    } else {
      matched++;
    }

    var next = has ? map[key] : '';
    var current = progSyncTrim_(progSyncCell_(row, cfg.TARGET_COL_OUTPUT));
    if (current === next) continue;                // 이미 같으면 쓰지 않는다

    updates.push({ row: absRow, key: key, current: current, next: next });
  }

  return { updates: updates, matched: matched, blankKeyRows: blankKeyRows, unmatchedKeys: unmatchedKeys };
}

/**
 * 변경할 행들을 "연속된 구간"으로 묶는다. setValues 호출 횟수를 줄이기 위한 것으로,
 * 중간에 건너뛴 행은 절대 포함하지 않는다(관리 대상이 아닌 행을 덮어쓰지 않기 위해).
 */
function progSyncGroupRuns_(updates) {
  var runs = [];
  for (var i = 0; i < updates.length; i++) {
    var u = updates[i];
    var last = runs.length > 0 ? runs[runs.length - 1] : null;
    if (last && u.row === last.startRow + last.values.length) {
      last.values.push([u.next]);
    } else {
      runs.push({ startRow: u.row, values: [[u.next]] });
    }
  }
  return runs;
}

/** 상품 식별 숫자를 문자열 키로 정규화한다. 숫자/문자/앞따옴표/천단위 콤마를 모두 흡수한다. */
function progSyncNormalizeKey_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    if (!isFinite(value)) return '';
    // 90817968961 같은 큰 정수가 지수 표기로 새지 않도록 정수는 반올림해서 쓴다.
    return value === Math.floor(value) ? String(Math.round(value)) : String(value);
  }
  var s = String(value).trim();
  if (s.charAt(0) === "'") s = s.substring(1);          // 텍스트 강제용 앞따옴표
  s = s.replace(/[\s,]/g, '');
  s = s.replace(/^(\d+)\.0+$/, '$1');                   // "123.0" → "123"
  return s;
}

/** 갯수 셀을 { countText, countNum } 으로 정규화한다. "3개"/"3"/3 모두 "3"이 된다. */
function progSyncParseCount_(value) {
  if (value === null || value === undefined) return { countText: '', countNum: null };
  if (typeof value === 'number') {
    if (!isFinite(value)) return { countText: '', countNum: null };
    var n = value === Math.floor(value) ? Math.round(value) : value;
    return { countText: String(n), countNum: n };
  }
  var s = String(value).trim();
  if (s === '') return { countText: '', countNum: null };
  s = s.replace(/개\s*$/, '').replace(/,/g, '').trim();  // "3개" → "3"
  if (s === '') return { countText: '', countNum: null };
  var num = /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  return { countText: s, countNum: num };
}

/** 셀 값을 1-based 열 번호로 꺼낸다. 범위를 벗어나면 빈 문자열. */
function progSyncCell_(row, col) {
  var v = row[col - 1];
  return (v === null || v === undefined) ? '' : v;
}

/** 값을 문자열로 바꾸고 앞뒤 공백을 없앤다. */
function progSyncTrim_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' && isFinite(value) && value === Math.floor(value)) {
    return String(Math.round(value));
  }
  return String(value).trim();
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. 열 매핑 확인 · 결과 검증
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 설정한 열 번호가 실제 시트의 어떤 항목을 가리키는지 헤더와 첫 데이터 행으로 보여준다.
 * 기대와 다르면 PROG_SYNC_CONFIG 의 숫자만 고치면 된다 (A=1, B=2, ... E=5, K=11, Q=17).
 */
function progSyncCheckColumns() {
  var cfg = PROG_SYNC_CONFIG;
  var lines = ['=== 열 매핑 확인 ==='];

  var src = progSyncOpenSheet_(cfg.SOURCE_SPREADSHEET_ID, cfg.SOURCE_SHEET_GID, cfg.SOURCE_LABEL);
  lines.push('');
  lines.push('[A시트] ' + src.ss.getName() + ' / ' + src.sheet.getName());
  lines.push(progSyncDescribeColumns_(src.sheet, cfg.SOURCE_HEADER_ROWS, [
    { col: cfg.SOURCE_COL_PROGRAM, label: '프로그램명' },
    { col: cfg.SOURCE_COL_COUNT, label: '갯수' },
    { col: cfg.SOURCE_COL_KEY, label: '상품키' }
  ]));

  var tgt = progSyncOpenSheet_(cfg.TARGET_SPREADSHEET_ID, cfg.TARGET_SHEET_GID, cfg.TARGET_LABEL);
  lines.push('');
  lines.push('[B시트] ' + tgt.ss.getName() + ' / ' + tgt.sheet.getName());
  lines.push(progSyncDescribeColumns_(tgt.sheet, cfg.TARGET_HEADER_ROWS, [
    { col: cfg.TARGET_COL_KEY, label: '상품키' },
    { col: cfg.TARGET_COL_OUTPUT, label: '기록될 열' }
  ]));

  var report = lines.join('\n');
  progSyncLog_(report);
  return report;
}

/** 지정한 열들의 헤더와 첫 3개 데이터 행을 사람이 읽을 수 있게 뽑는다. */
function progSyncDescribeColumns_(sheet, headerRows, specs) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return '  (빈 시트)';

  var header = headerRows > 0
    ? sheet.getRange(headerRows, 1, 1, lastCol).getValues()[0]
    : [];
  var sampleCount = Math.min(3, lastRow - headerRows);
  var samples = sampleCount > 0
    ? sheet.getRange(headerRows + 1, 1, sampleCount, lastCol).getValues()
    : [];

  var lines = [];
  for (var i = 0; i < specs.length; i++) {
    var spec = specs[i];
    var letter = progSyncColumnLetter_(spec.col);
    lines.push('  ' + letter + '열(' + spec.label + ') 헤더: "' +
               progSyncTrim_(progSyncCell_(header, spec.col)) + '"');
    for (var r = 0; r < samples.length; r++) {
      lines.push('      ' + (headerRows + 1 + r) + '행 → "' +
                 progSyncTrim_(progSyncCell_(samples[r], spec.col)) + '"');
    }
  }
  return lines.join('\n');
}

/**
 * 동기화 결과 검증. B시트를 다시 읽어 E열이 A시트 기준 기대값과 일치하는지 대조한다.
 * 불일치가 있으면 그 행을 그대로 출력한다.
 */
function progSyncVerifyMatch() {
  var result = progSyncPlan_();
  var lines = ['=== 동기화 결과 검증 ===',
               'A시트: ' + result.sourceName,
               'B시트: ' + result.targetName,
               'B시트 데이터 행: ' + result.targetRows + '행',
               'A시트와 매칭된 행: ' + result.matched + '행',
               '기대값과 일치: ' + result.unchanged + '행',
               '기대값과 불일치: ' + result.updates.length + '행'];

  if (result.updates.length === 0) {
    lines.push('', '✅ B시트 E열이 A시트와 완전히 일치합니다.');
  } else {
    lines.push('', '⚠️ 아직 반영되지 않은 행이 있습니다. progSyncRun() 을 실행하세요.');
    var shown = Math.min(result.updates.length, 20);
    for (var i = 0; i < shown; i++) {
      var u = result.updates[i];
      lines.push('  E' + u.row + ' (키 ' + u.key + '): 현재 "' + u.current +
                 '" / 기대 "' + u.next + '"');
    }
    if (result.updates.length > shown) {
      lines.push('  ... 외 ' + (result.updates.length - shown) + '건');
    }
  }

  var report = lines.join('\n');
  progSyncLog_(report);
  return report;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. 시트 접근 · 로그
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 스프레드시트를 열고 gid로 시트 탭을 찾는다. */
function progSyncOpenSheet_(id, gid, label) {
  var ss = progSyncOpenSpreadsheet_(id, label);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return { ss: ss, sheet: sheets[i] };
  }
  var names = [];
  for (var j = 0; j < sheets.length; j++) {
    names.push(sheets[j].getName() + '(gid ' + sheets[j].getSheetId() + ')');
  }
  throw new Error(label + '에서 gid ' + gid + ' 탭을 찾을 수 없습니다. ' +
                  '존재하는 탭: ' + names.join(', ') + '. ' +
                  'PROG_SYNC_CONFIG 의 GID 값을 시트 URL 끝의 #gid=... 와 맞추세요.');
}

/**
 * 스프레드시트를 연다. 이 스크립트가 붙어 있는 문서면 openById 대신 getActive를 쓴다.
 * (기존 코드에 @OnlyCurrentDoc 주석이 있으면 openById가 막히기 때문)
 */
function progSyncOpenSpreadsheet_(id, label) {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active && active.getId() === id) return active;

  try {
    var ss = SpreadsheetApp.openById(id);
    if (!ss) throw new Error('openById가 null을 반환했습니다.');
    return ss;
  } catch (e) {
    throw new Error(label + '(' + id + ')을 열 수 없습니다: ' + e.message +
                    ' — 승인한 Google 계정에 이 스프레드시트 접근 권한이 있는지, ' +
                    '그리고 프로젝트에 @OnlyCurrentDoc 주석이 없는지 확인하세요.');
  }
}

/** 대상 시트가 실제로 쓰기 가능한지 확인한다(보호된 시트/읽기 전용 공유 방지). */
function progSyncAssertWritable_(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    if (!protections[i].canEdit()) {
      throw new Error('B시트("' + sheet.getName() + '")가 보호되어 있어 쓸 수 없습니다. ' +
                      '시트 보호 설정에서 이 계정에 편집 권한을 주세요.');
    }
  }
}

/** 사람이 읽는 열 문자(1 → A, 27 → AA). */
function progSyncColumnLetter_(col) {
  var s = '';
  var n = col;
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 실행 결과를 한국어 리포트로 만든다. */
function progSyncFormatReport_(result) {
  var now = Utilities.formatDate(new Date(), PROG_SYNC_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var lines = [
    '실행 시각: ' + now + ' (' + PROG_SYNC_CONFIG.TIMEZONE + ')',
    'A시트: ' + result.sourceName + ' — 데이터 ' + result.sourceRows + '행 중 ' +
      result.sourceUsedRows + '행 사용, 상품 ' + result.sourceKeys + '개',
    'B시트: ' + result.targetName + ' — 데이터 ' + result.targetRows + '행',
    '매칭 ' + result.matched + '행 / 갱신 ' + result.updates.length +
      '행 / 유지 ' + result.unchanged + '행',
    'K열 비어 건너뜀: ' + result.blankKeyRows + '행 / A시트에 없는 키: ' +
      result.unmatchedKeys.length + '행',
    '시트 쓰기 호출: ' + result.writeCalls + '회'
  ];
  return lines.join('\n');
}

/** Logger와 console 양쪽에 남긴다(트리거 실행 로그는 console 쪽에 쌓인다). */
function progSyncLog_(message) {
  try { Logger.log(message); } catch (e) { /* noop */ }
  try { console.log(message); } catch (e) { /* noop */ }
}
