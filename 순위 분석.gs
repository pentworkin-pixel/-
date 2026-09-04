/**
 * 순위 분석.gs
 *
 * B시트(순위추적)의 순위 데이터를 읽어, A스프레드시트에 「프로그램별 순위 분석」시트를
 * 만들고 프로그램별 반영도(상승/유지/하락)를 1일·3일·7일 전과 비교해 정리한다.
 *
 *   순위는 숫자가 작을수록 좋다.
 *   순위 변화폭 = 과거 순위 - 최신 순위   (양수=상승, 0=유지, 음수=하락)
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * 기존 프로젝트 코드와 충돌하지 않도록 다음 규칙을 지킨다.
 *   - 전역 식별자는 rankAnalysis / RANK_ANALYSIS 접두사를 쓴다.
 *     (예외 2개: 요구사항이 지정한 updateProgramRankingAnalysis, 그리고 메뉴용 onOpen)
 *   - B시트 원본과 사용자가 만든 다른 시트는 절대 수정하지 않는다.
 *     쓰기는 ANALYSIS_SHEET_NAME / DETAIL_SHEET_NAME 두 시트에만 일어난다.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * ⚠️ onOpen 주의: 이 프로젝트에 이미 onOpen 함수가 있다면 둘 중 하나만 살아남는다.
 *    그런 경우 아래 onOpen 은 지우고, 기존 onOpen 안에서 rankAnalysisBuildMenu_(ui) 를
 *    호출하도록 한 줄만 추가하면 된다.
 *
 * 사용자가 직접 실행할 함수:
 *   updateProgramRankingAnalysis()  분석 실행 · 결과 시트 갱신 (메뉴에서 호출됨)
 *   rankAnalysisCheckColumns()      어떤 열을 날짜/프로그램/순위/고유항목으로 인식했는지 확인
 */

var RANK_ANALYSIS_CONFIG = {
  // ── 대상 ────────────────────────────────────────────────────────────────
  // 분석 결과가 기록될 A 스프레드시트
  TARGET_SPREADSHEET_ID: '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8',
  // 순위 원본이 있는 B 스프레드시트. A 안의 시트라면 SOURCE_SPREADSHEET_ID 를 비우고
  // SOURCE_SHEET_NAME 만 채우면 getSheetByName 으로 찾는다.
  SOURCE_SPREADSHEET_ID: '1bLNh-zrYHHKWgH78ihbgnhpd11ItUhFgw2eT0MromFU',
  SOURCE_SHEET_GID: 1003701754,
  SOURCE_SHEET_NAME: '',        // gid 를 못 찾을 때 쓰는 예비 이름

  ANALYSIS_SHEET_NAME: '프로그램별 순위 분석',
  DETAIL_SHEET_NAME: '프로그램별 순위 상세',
  TIMEZONE: 'Asia/Seoul',

  // ── 시트 구조 ───────────────────────────────────────────────────────────
  HEADER_ROW: 1,
  PERIODS: [1, 3, 7],           // 기준일 대비 며칠 전과 비교할지

  // 열 자동 감지용 헤더 패턴 (정규식). 실제 헤더가 달라도 이 패턴으로 찾아낸다.
  PROGRAM_HEADER_PATTERN: /프로그램/,
  PROGRAM_COL_FALLBACK: 5,      // 못 찾으면 E열 (프로그램 동기화.gs 가 쓰는 열)
  DATE_HEADER_PATTERN: /날짜|일자|date|업데이트/i,
  RANK_HEADER_PATTERN: /순위|rank/i,
  // "순위 조회 키워드", "전일 비교 순위(자유 입력)" 처럼 순위 값이 아닌 열을
  // 순위 열로 오인하지 않도록 걸러낸다.
  RANK_HEADER_EXCLUDE: /키워드|검색어|조회|비교|입력/,
  DATE_HEADER_EXCLUDE: /순위|키워드/,
  // 고유키를 만들 때 쓰는 열들.
  //   [0] 강한 식별자(상품 MID 등) — 사람이 고쳐 쓰지 않는 값
  //   [1] 구분 축(키워드) — 같은 상품을 키워드별로 따로 추적하므로 함께 묶는다
  // 이름·메모 열은 사람이 수시로 고쳐 쓰기 때문에 키에 넣으면 글자 하나만 바뀌어도
  // 다른 항목으로 잡혀 비교가 끊긴다. 그래서 강한 식별자를 못 찾았을 때만 예비로 쓴다.
  KEY_HEADER_PATTERNS: [
    /상품\s*MID|^MID$|상품코드|상품번호|product\s*id/i,
    /키워드|검색어|keyword/i
  ],
  KEY_NAME_FALLBACK_PATTERNS: [
    /상품명|제품명|리워드명|메모/
  ],
  KEY_SEPARATOR: ' ｜ ',

  // ── 값 해석 ─────────────────────────────────────────────────────────────
  // 순위 칸이 이 값들이면 "순위 없음"으로 보고 숫자 순위와 섞어 계산하지 않는다.
  NO_RANK_TOKENS: ['-', '–', '—', 'x', 'X', 'n/a', 'na', '없음', '미노출',
                   '순위없음', '순위 없음', '노출없음', '검색안됨', '?'],
  MAX_VALID_RANK: 100000,       // 이보다 큰 값은 순위로 보지 않는다

  // E열이 "리뷰 3개, 트래픽 2개" 처럼 여러 프로그램을 담고 있을 때 쪼개서
  // 각 프로그램에 항목을 귀속시킨다. false 면 문자열 전체를 하나의 프로그램명으로 본다.
  SPLIT_PROGRAMS: true,
  PROGRAM_SPLIT_PATTERN: /[,;/·]/,
  PROGRAM_COUNT_SUFFIX: /\s*\d+\s*개\s*$/,   // "리뷰 3개" → "리뷰"

  // 2자리 연도 헤더("26-08-20")를 연-월-일로 읽는다. false 면 일-월-연으로 읽는다.
  TWO_DIGIT_YEAR_FIRST: true,

  // ── 출력 ────────────────────────────────────────────────────────────────
  MAX_DETAIL_ROWS: 5000,        // 상세 시트가 지나치게 커지지 않도록 상한
  DRAW_CHART: true
};

/** 프로그램별 요약표의 열 구성. 순서를 바꾸면 표와 서식이 같이 따라간다. */
var RANK_ANALYSIS_TABLE_HEADERS = [
  '순위', '프로그램명', '비교 가능 항목 수', '상승 수', '유지 수', '하락 수',
  '신규 수', '누락 또는 종료 수', '상승률', '유지율', '하락률',
  '평균 순위 변화폭', '평균 상승 폭', '평균 하락 폭', '최신 평균 순위',
  '단독 항목 수'
];

/** 상세 시트의 열 구성. */
var RANK_ANALYSIS_DETAIL_HEADERS = [
  '프로그램명', '고유 항목명', '최신 순위', '과거 순위', '순위 변화폭',
  '상태', '최신 업데이트 날짜', '실제 비교 날짜', '비교 기간'
];


/* ═══════════════════════════════════════════════════════════════════════════
 * 1. 메뉴 · 진입점
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 스프레드시트를 열 때 「순위 분석」 메뉴를 만든다. */
function onOpen() {
  rankAnalysisBuildMenu_(SpreadsheetApp.getUi());
}

/** 메뉴 구성. 기존 onOpen 이 따로 있다면 그 안에서 이 함수만 불러도 된다. */
function rankAnalysisBuildMenu_(ui) {
  ui.createMenu('순위 분석')
    .addItem('프로그램별 순위 분석 업데이트', 'updateProgramRankingAnalysis')
    .addItem('열 매핑 확인', 'rankAnalysisCheckColumns')
    .addItem('원본 헤더 그대로 보기 (디버그)', 'rankAnalysisDumpSource')
    .addToUi();
}

/**
 * 열 자동 감지를 거치지 않고, 원본 시트의 처음 몇 행을 있는 그대로 찍는다.
 * rankAnalysisCheckColumns()/updateProgramRankingAnalysis() 가 "필수 열을 찾지
 * 못했습니다" 오류로 죽을 때, 실제 헤더가 뭐라고 적혀 있는지 보려고 쓰는 함수다.
 */
function rankAnalysisDumpSource() {
  var found = rankAnalysisOpenSourceSheet_();
  var sheet = found.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var lines = ['=== 원본 헤더 그대로 보기 ===',
               '원본 시트: ' + found.label,
               '마지막 행/열: ' + lastRow + '행 / ' + lastCol + '열 (' +
                 rankAnalysisColumnLetter_(lastCol) + '열)'];

  var rowsToShow = Math.min(3, lastRow);
  if (rowsToShow === 0) {
    lines.push('', '(시트에 데이터가 없습니다)');
  } else {
    var values = sheet.getRange(1, 1, rowsToShow, lastCol).getValues();
    for (var r = 0; r < values.length; r++) {
      lines.push('', '── ' + (r + 1) + '행 ──');
      for (var c = 0; c < values[r].length; c++) {
        var raw = values[r][c];
        var type = raw instanceof Date ? 'Date' : typeof raw;
        lines.push('  ' + rankAnalysisColumnLetter_(c + 1) + '열 [' + type + ']: ' +
                   JSON.stringify(raw instanceof Date ? raw.toISOString() : raw));
      }
    }
  }

  var report = lines.join('\n');
  rankAnalysisLog_(report);
  rankAnalysisAlert_('원본 헤더 확인', report);
  return report;
}

/**
 * 메인 실행 함수. B시트를 읽어 분석하고 A스프레드시트의 분석 시트를 갱신한다.
 * 셀을 하나씩 쓰지 않고 배열을 만들어 setValues 로 일괄 기록한다.
 */
function updateProgramRankingAnalysis() {
  var cfg = RANK_ANALYSIS_CONFIG;
  var started = new Date();

  var analysis;
  try {
    analysis = rankAnalysisCompute_();
  } catch (e) {
    rankAnalysisAlert_('순위 분석 실패', e.message);
    throw e;
  }

  var ss = rankAnalysisOpenSpreadsheet_(cfg.TARGET_SPREADSHEET_ID, 'A스프레드시트');
  rankAnalysisWriteAnalysisSheet_(ss, analysis, started);
  rankAnalysisWriteDetailSheet_(ss, analysis);
  SpreadsheetApp.flush();

  var report = rankAnalysisFormatReport_(analysis, started);
  rankAnalysisLog_(report);
  rankAnalysisAlert_('순위 분석 완료', report);
  return report;
}

/**
 * 어떤 열을 무엇으로 인식했는지 보여준다. 코드를 고치기 전에 이것부터 실행하면
 * 열 매핑이 맞는지 바로 알 수 있다.
 */
function rankAnalysisCheckColumns() {
  var read = rankAnalysisReadSource_();
  var det = read.detection;
  var lines = ['=== 순위 분석 열 매핑 확인 ===',
               '원본 시트: ' + read.sheetLabel,
               '데이터 행 수: ' + read.rows.length + '행',
               '',
               '인식한 표 형태: ' + (det.layout === 'wide'
                 ? '가로형 — 날짜가 열 머리글로 있고 각 칸이 그 날짜의 순위'
                 : '세로형 — 날짜 열과 순위 열이 따로 있고 한 행이 한 관측')];

  lines.push('');
  lines.push('프로그램명 열: ' + rankAnalysisDescribeCol_(read.header, det.programCol));
  if (det.layout === 'wide') {
    lines.push('날짜 열: ' + det.dateCols.length + '개');
    for (var i = 0; i < det.dateCols.length; i++) {
      var dc = det.dateCols[i];
      lines.push('   ' + rankAnalysisColumnLetter_(dc.col) + '열 → ' + dc.dateKey +
                 ' (머리글 "' + rankAnalysisTrim_(read.header[dc.col - 1]) + '")');
    }
  } else {
    lines.push('날짜 열: ' + rankAnalysisDescribeCol_(read.header, det.dateCol));
    lines.push('순위 열: ' + rankAnalysisDescribeCol_(read.header, det.rankCol));
  }

  lines.push('');
  lines.push('고유키 구성 열: ' + (det.keyCols.length === 0
    ? '(못 찾음 — 행 번호를 고유키로 사용)'
    : det.keyCols.map(function (c) { return rankAnalysisDescribeCol_(read.header, c); }).join(' + ')));

  var sample = Math.min(3, read.rows.length);
  if (sample > 0) {
    lines.push('');
    lines.push('── 샘플 (' + sample + '행) ──');
    for (var r = 0; r < sample; r++) {
      var row = read.rows[r];
      lines.push('  고유키: ' + rankAnalysisBuildItemKey_(row, det, r, cfgOf_()));
      lines.push('    프로그램: ' + JSON.stringify(
        rankAnalysisSplitPrograms_(rankAnalysisCell_(row, det.programCol), cfgOf_())));
      if (det.layout === 'wide' && det.dateCols.length > 0) {
        var first = det.dateCols[det.dateCols.length - 1];
        lines.push('    ' + first.dateKey + ' 순위: ' +
                   JSON.stringify(rankAnalysisParseRank_(rankAnalysisCell_(row, first.col), cfgOf_())));
      } else if (det.layout === 'long') {
        lines.push('    날짜: ' + rankAnalysisParseDate_(rankAnalysisCell_(row, det.dateCol), cfgOf_()));
        lines.push('    순위: ' + JSON.stringify(
          rankAnalysisParseRank_(rankAnalysisCell_(row, det.rankCol), cfgOf_())));
      }
    }
  }

  var report = lines.join('\n');
  rankAnalysisLog_(report);
  rankAnalysisAlert_('열 매핑 확인', report);
  return report;
}

/** 설정 참조를 짧게 쓰기 위한 헬퍼. */
function cfgOf_() { return RANK_ANALYSIS_CONFIG; }


/* ═══════════════════════════════════════════════════════════════════════════
 * 2. 분석 파이프라인
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 원본을 읽고 전체 분석 결과를 만든다. 시트 쓰기는 하지 않는다. */
function rankAnalysisCompute_() {
  var cfg = RANK_ANALYSIS_CONFIG;
  var read = rankAnalysisReadSource_();
  var obs = rankAnalysisBuildObservations_(read.rows, read.detection, cfg);

  if (obs.dateKeys.length === 0) {
    throw new Error('B시트에서 유효한 날짜를 하나도 찾지 못했습니다. ' +
                    'rankAnalysisCheckColumns() 를 실행해 날짜 열이 제대로 인식되는지 확인하세요.');
  }

  var baseKey = obs.dateKeys[obs.dateKeys.length - 1];
  var picks = rankAnalysisPickComparisonDates_(obs.dateKeys, baseKey, cfg.PERIODS);

  var periods = [];
  for (var i = 0; i < cfg.PERIODS.length; i++) {
    var days = cfg.PERIODS[i];
    var pastKey = picks[days];
    periods.push({
      days: days,
      pastKey: pastKey,
      result: pastKey ? rankAnalysisAggregate_(obs, baseKey, pastKey, cfg)
                      : { programs: [], details: [] }
    });
  }

  return {
    sheetLabel: read.sheetLabel,
    detection: read.detection,
    observations: obs,
    baseKey: baseKey,
    periods: periods,
    programCount: rankAnalysisCountPrograms_(obs, baseKey),
    issues: obs.issues
  };
}

/** B시트를 읽고 열 매핑까지 감지한다. */
function rankAnalysisReadSource_() {
  var cfg = RANK_ANALYSIS_CONFIG;
  var found = rankAnalysisOpenSourceSheet_();
  var sheet = found.sheet;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < cfg.HEADER_ROW || lastCol < 1) {
    throw new Error('원본 시트("' + found.label + '")가 비어 있습니다.');
  }

  var all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var header = all[cfg.HEADER_ROW - 1] || [];
  var rows = all.slice(cfg.HEADER_ROW);

  var detection = rankAnalysisDetectColumns_(header, rows, cfg);
  return { sheetLabel: found.label, header: header, rows: rows, detection: detection };
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 3. 순수 로직 (스프레드시트 API 없이 동작 — 로컬 테스트 대상)
 * ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 헤더와 데이터를 보고 어떤 열이 무엇인지 알아낸다.
 *
 * 표 형태는 두 가지를 지원한다.
 *   가로형(wide): 날짜가 열 머리글이고 각 칸이 그 날짜의 순위 (순위추적 시트의 일반적 형태)
 *   세로형(long): "날짜" 열과 "순위" 열이 따로 있고 한 행이 한 관측
 * 날짜로 해석되는 머리글이 2개 이상이면 가로형으로 본다.
 */
function rankAnalysisDetectColumns_(header, rows, cfg) {
  var dateCols = [];
  for (var c = 1; c <= header.length; c++) {
    var key = rankAnalysisParseDate_(header[c - 1], cfg);
    if (key) dateCols.push({ col: c, dateKey: key });
  }
  dateCols.sort(function (a, b) { return a.dateKey < b.dateKey ? -1 : (a.dateKey > b.dateKey ? 1 : 0); });

  var programCol = rankAnalysisFindHeader_(header, cfg.PROGRAM_HEADER_PATTERN);
  if (!programCol) programCol = cfg.PROGRAM_COL_FALLBACK;

  // 고유키 열: 우선순위 패턴별로 하나씩 찾아 조합한다.
  // 날짜 열이나 프로그램 열은 고유키에 넣지 않는다(시간에 따라 변하므로).
  var keyCols = [];
  var strongFound = false;

  function collect(patterns) {
    for (var p = 0; p < patterns.length; p++) {
      var col = rankAnalysisFindHeader_(header, patterns[p]);
      if (!col) continue;
      if (col === programCol) continue;
      if (rankAnalysisHasCol_(dateCols, col)) continue;
      if (keyCols.indexOf(col) >= 0) continue;
      keyCols.push(col);
      if (p === 0 && patterns === cfg.KEY_HEADER_PATTERNS) strongFound = true;
    }
  }

  collect(cfg.KEY_HEADER_PATTERNS);
  // 상품 MID 같은 강한 식별자를 찾았으면 이름·메모 열은 키에 넣지 않는다.
  // (이름은 수정되는 값이라 키에 넣으면 같은 상품이 다른 항목으로 갈라진다)
  if (!strongFound) collect(cfg.KEY_NAME_FALLBACK_PATTERNS);

  var dateCol = rankAnalysisFindHeader_(header, cfg.DATE_HEADER_PATTERN, cfg.DATE_HEADER_EXCLUDE);
  var rankCol = rankAnalysisFindHeader_(header, cfg.RANK_HEADER_PATTERN, cfg.RANK_HEADER_EXCLUDE);
  if (rankAnalysisHasCol_(dateCols, dateCol)) dateCol = 0;
  if (rankAnalysisHasCol_(dateCols, rankCol)) rankCol = 0;

  // 날짜 머리글이 2개 이상이면 가로형이 확실하다.
  // 1개뿐이어도 "날짜 열 + 순위 열" 짝을 못 찾았다면 가로형으로 본다
  // (데이터가 하루치만 쌓인 초기 상태 — 비교는 못 해도 읽기는 해야 한다).
  if (dateCols.length >= 2 || (dateCols.length === 1 && !(dateCol && rankCol))) {
    return { layout: 'wide', dateCols: dateCols, programCol: programCol, keyCols: keyCols,
             dateCol: 0, rankCol: 0 };
  }

  if (!dateCol || !rankCol) {
    throw new Error('필수 열을 찾지 못했습니다. 날짜로 읽히는 머리글이 ' + dateCols.length +
                    '개뿐이고, "날짜" 열' + (dateCol ? '은 찾았지만' : '도 찾지 못했으며') +
                    ' "순위" 열' + (rankCol ? '은 찾았습니다.' : '도 찾지 못했습니다.') +
                    ' rankAnalysisCheckColumns() 로 실제 머리글을 확인한 뒤 ' +
                    'RANK_ANALYSIS_CONFIG 의 패턴을 조정하세요.');
  }
  keyCols = keyCols.filter(function (c) { return c !== dateCol && c !== rankCol; });
  return { layout: 'long', dateCols: [], programCol: programCol, keyCols: keyCols,
           dateCol: dateCol, rankCol: rankCol };
}

/**
 * 헤더에서 패턴에 맞는 첫 열의 1-based 번호를 찾는다. 없으면 0.
 * exclude 에 걸리는 머리글은 건너뛴다("순위 조회 키워드"를 순위 열로 잡지 않기 위해).
 */
function rankAnalysisFindHeader_(header, pattern, exclude) {
  for (var c = 1; c <= header.length; c++) {
    var text = rankAnalysisTrim_(header[c - 1]);
    if (text === '') continue;
    if (exclude && exclude.test(text)) continue;
    if (pattern.test(text)) return c;
  }
  return 0;
}

function rankAnalysisHasCol_(dateCols, col) {
  for (var i = 0; i < dateCols.length; i++) if (dateCols[i].col === col) return true;
  return false;
}

/**
 * 행들을 (날짜 × 고유항목) 관측으로 펼친다.
 *
 * 중복 처리: 같은 날짜 + 같은 고유항목이 여러 번 나오면 마지막 행의 값을 쓴다.
 * 프로그램 귀속: 프로그램명이 비어 있지 않은 행 중 가장 나중 것을 그 항목의 프로그램으로 본다.
 */
function rankAnalysisBuildObservations_(rows, det, cfg) {
  var byDate = {};        // dateKey → itemKey → {rank, hasRank}
  var programsByItem = {};
  var labelByItem = {};
  var dateSet = {};
  var issues = { duplicates: 0, blankProgram: 0, badDate: 0, noRank: 0, blankKey: 0 };

  function put(dateKey, itemKey, parsed) {
    if (!byDate[dateKey]) byDate[dateKey] = {};
    if (Object.prototype.hasOwnProperty.call(byDate[dateKey], itemKey)) issues.duplicates++;
    byDate[dateKey][itemKey] = parsed;   // 마지막 행이 이긴다
    dateSet[dateKey] = true;
    if (!parsed.hasRank) issues.noRank++;
  }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var itemKey = rankAnalysisBuildItemKey_(row, det, i, cfg);
    if (itemKey === '') { issues.blankKey++; continue; }

    var programCell = rankAnalysisCell_(row, det.programCol);
    var programs = rankAnalysisSplitPrograms_(programCell, cfg);
    if (programs.length === 0) issues.blankProgram++;

    if (det.layout === 'wide') {
      var wrote = false;
      for (var d = 0; d < det.dateCols.length; d++) {
        var dc = det.dateCols[d];
        put(dc.dateKey, itemKey, rankAnalysisParseRank_(rankAnalysisCell_(row, dc.col), cfg));
        wrote = true;
      }
      if (wrote) {
        labelByItem[itemKey] = itemKey;
        if (programs.length > 0) programsByItem[itemKey] = programs;
      }
    } else {
      var dateKey = rankAnalysisParseDate_(rankAnalysisCell_(row, det.dateCol), cfg);
      if (!dateKey) { issues.badDate++; continue; }
      put(dateKey, itemKey, rankAnalysisParseRank_(rankAnalysisCell_(row, det.rankCol), cfg));
      labelByItem[itemKey] = itemKey;
      // 세로형에서는 나중 날짜의 프로그램명을 우선한다.
      if (programs.length > 0) {
        var prev = programsByItem[itemKey];
        if (!prev || !prev.__dateKey || prev.__dateKey <= dateKey) {
          programs.__dateKey = dateKey;
          programsByItem[itemKey] = programs;
        }
      }
    }
  }

  var dateKeys = Object.keys(dateSet).sort();
  return { dateKeys: dateKeys, byDate: byDate, programsByItem: programsByItem,
           labelByItem: labelByItem, issues: issues };
}

/**
 * 고유키를 만든다.
 * 감지된 식별 열들(예: 상품 MID + 순위 조회 키워드)을 ' ｜ ' 로 이어 붙인다.
 * 식별 열을 하나도 못 찾으면 행 번호를 쓴다(가로형에서는 행 자체가 항목이므로 유효하다).
 */
function rankAnalysisBuildItemKey_(row, det, rowIndex, cfg) {
  var parts = [];
  for (var i = 0; i < det.keyCols.length; i++) {
    var v = rankAnalysisTrim_(rankAnalysisCell_(row, det.keyCols[i]));
    if (v !== '') parts.push(v);
  }
  if (parts.length > 0) return parts.join(cfg.KEY_SEPARATOR);
  if (det.layout === 'wide') return '행 ' + (rowIndex + 1 + cfg.HEADER_ROW);
  return '';
}

/** "리뷰 3개, 트래픽 2개" → ['리뷰', '트래픽'] */
function rankAnalysisSplitPrograms_(value, cfg) {
  var text = rankAnalysisTrim_(value);
  if (text === '') return [];
  if (!cfg.SPLIT_PROGRAMS) return [text];

  var raw = text.split(cfg.PROGRAM_SPLIT_PATTERN);
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var name = raw[i].replace(cfg.PROGRAM_COUNT_SUFFIX, '').trim();
    if (name !== '' && out.indexOf(name) < 0) out.push(name);
  }
  return out;
}

/**
 * 순위 칸을 해석한다.
 * 숫자가 아니거나 NO_RANK_TOKENS 에 해당하면 순위 없음(hasRank=false)으로 둔다.
 * 순위 없음은 숫자 순위와 섞어 평균내지 않는다.
 */
function rankAnalysisParseRank_(value, cfg) {
  if (value === null || value === undefined) return { rank: null, hasRank: false };

  if (typeof value === 'number') {
    if (!isFinite(value) || value <= 0 || value > cfg.MAX_VALID_RANK) {
      return { rank: null, hasRank: false };
    }
    return { rank: value, hasRank: true };
  }

  var s = String(value).trim();
  if (s === '') return { rank: null, hasRank: false };

  var lowered = s.toLowerCase();
  for (var i = 0; i < cfg.NO_RANK_TOKENS.length; i++) {
    if (lowered === String(cfg.NO_RANK_TOKENS[i]).toLowerCase()) {
      return { rank: null, hasRank: false };
    }
  }

  // "12위", "1,234" 같은 표기 허용. "100+" 처럼 범위를 뜻하는 건 순위 없음으로 본다.
  var cleaned = s.replace(/위\s*$/, '').replace(/,/g, '').trim();
  if (!/^\d+(\.\d+)?$/.test(cleaned)) return { rank: null, hasRank: false };

  var n = Number(cleaned);
  if (!isFinite(n) || n <= 0 || n > cfg.MAX_VALID_RANK) return { rank: null, hasRank: false };
  return { rank: n, hasRank: true };
}

/**
 * 날짜 값을 'yyyy-MM-dd' 키로 바꾼다. 해석할 수 없으면 null.
 * Date 객체 / 스프레드시트 날짜 시리얼 / "2026-08-20" / "26-08-20" / "8월 20일" 을 받는다.
 */
function rankAnalysisParseDate_(value, cfg) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return rankAnalysisDateKey_(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  if (typeof value === 'number') {
    // 스프레드시트 날짜 시리얼(1899-12-30 기준). 순위 숫자와 헷갈리지 않도록 범위를 제한한다.
    if (!isFinite(value) || value < 20000 || value > 80000) return null;
    var ms = Math.round((value - 25569) * 86400000);
    var d = new Date(ms);
    return rankAnalysisDateKey_(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }

  var s = String(value).trim();
  if (s === '') return null;

  var m = /^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/.exec(s);
  if (m) return rankAnalysisDateKey_(Number(m[1]), Number(m[2]), Number(m[3]));

  m = /^(\d{1,2})[-.\/](\d{1,2})[-.\/](\d{1,2})/.exec(s);
  if (m) {
    var a = Number(m[1]), b = Number(m[2]), c = Number(m[3]);
    return cfg.TWO_DIGIT_YEAR_FIRST
      ? rankAnalysisDateKey_(2000 + a, b, c)
      : rankAnalysisDateKey_(2000 + c, b, a);
  }

  m = /(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(s);
  if (m) return rankAnalysisDateKey_(Number(m[1]), Number(m[2]), Number(m[3]));

  m = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(s);
  if (m) return rankAnalysisDateKey_(new Date().getFullYear(), Number(m[1]), Number(m[2]));

  return null;
}

/** 연/월/일을 'yyyy-MM-dd' 로. 달력상 존재하지 않는 날짜면 null. */
function rankAnalysisDateKey_(y, mo, d) {
  if (!isFinite(y) || !isFinite(mo) || !isFinite(d)) return null;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  var dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  return y + '-' + rankAnalysisPad_(mo) + '-' + rankAnalysisPad_(d);
}

function rankAnalysisPad_(n) { return (n < 10 ? '0' : '') + n; }

/** 'yyyy-MM-dd' → 1970-01-01 기준 일수. 날짜 뺄셈을 시간대 영향 없이 하기 위한 것. */
function rankAnalysisDayNumber_(key) {
  var p = String(key).split('-');
  return Math.round(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])) / 86400000);
}

/**
 * 기준일 대비 N일 전 비교 날짜를 고른다.
 * 목표 날짜에 정확히 맞는 데이터가 없으면, 목표일보다 이전이면서 가장 가까운 날짜를 쓴다.
 * 비교할 과거 데이터가 아예 없으면 null (→ "데이터 없음").
 */
function rankAnalysisPickComparisonDates_(dateKeys, baseKey, offsets) {
  var baseDay = rankAnalysisDayNumber_(baseKey);
  var picks = {};

  for (var i = 0; i < offsets.length; i++) {
    var target = baseDay - offsets[i];
    var best = null;
    for (var j = 0; j < dateKeys.length; j++) {
      var day = rankAnalysisDayNumber_(dateKeys[j]);
      if (day >= baseDay) continue;          // 기준일 이후/당일은 비교 대상이 아니다
      if (day > target) continue;            // 목표일보다 나중이면 쓰지 않는다
      if (best === null || day > rankAnalysisDayNumber_(best)) best = dateKeys[j];
    }
    picks[offsets[i]] = best;
  }
  return picks;
}

/**
 * 한 기간(기준일 vs 과거일)에 대해 프로그램별 집계와 상세 목록을 만든다.
 *
 *   순위 변화폭 = 과거 순위 - 최신 순위   (양수=상승, 0=유지, 음수=하락)
 *   비율의 분모는 "최신·과거 모두 숫자 순위가 있는" 비교 가능 항목 수.
 *   신규(과거 순위 없음) / 누락·종료(최신 순위 없음)는 비율 계산에서 제외한다.
 */
function rankAnalysisAggregate_(obs, baseKey, pastKey, cfg) {
  var latest = obs.byDate[baseKey] || {};
  var past = obs.byDate[pastKey] || {};

  var acc = {};    // program → 누적
  var details = [];

  function bucket(program) {
    if (!acc[program]) {
      acc[program] = { program: program, comparable: 0, up: 0, same: 0, down: 0,
                       newCount: 0, lostCount: 0, deltaSum: 0, upSum: 0, downSum: 0,
                       latestRankSum: 0, latestRankCount: 0, soloComparable: 0 };
    }
    return acc[program];
  }

  // 최신과 과거에 등장하는 모든 항목을 훑는다.
  var seen = {};
  var itemKeys = [];
  var k;
  for (k in latest) if (Object.prototype.hasOwnProperty.call(latest, k) && !seen[k]) { seen[k] = true; itemKeys.push(k); }
  for (k in past) if (Object.prototype.hasOwnProperty.call(past, k) && !seen[k]) { seen[k] = true; itemKeys.push(k); }

  for (var i = 0; i < itemKeys.length; i++) {
    var itemKey = itemKeys[i];
    var l = latest[itemKey] || { rank: null, hasRank: false };
    var p = past[itemKey] || { rank: null, hasRank: false };
    var programs = obs.programsByItem[itemKey] || [];
    if (programs.length === 0) continue;   // 프로그램명이 없으면 어디에도 귀속시키지 않는다

    var status, delta = null;
    if (l.hasRank && p.hasRank) {
      delta = p.rank - l.rank;
      status = delta > 0 ? '상승' : (delta === 0 ? '유지' : '하락');
    } else if (l.hasRank && !p.hasRank) {
      status = '신규';
    } else if (!l.hasRank && p.hasRank) {
      status = '누락 또는 종료';
    } else {
      continue;                            // 양쪽 다 순위 없음 → 집계 대상 아님
    }

    for (var g = 0; g < programs.length; g++) {
      var b = bucket(programs[g]);
      if (status === '상승' || status === '유지' || status === '하락') {
        b.comparable++;
        b.deltaSum += delta;
        if (programs.length === 1) b.soloComparable++;
        if (status === '상승') { b.up++; b.upSum += delta; }
        else if (status === '유지') { b.same++; }
        else { b.down++; b.downSum += -delta; }
      } else if (status === '신규') {
        b.newCount++;
      } else {
        b.lostCount++;
      }
      if (l.hasRank) { b.latestRankSum += l.rank; b.latestRankCount++; }
    }

    details.push({
      programs: programs,
      itemLabel: obs.labelByItem[itemKey] || itemKey,
      latestRank: l.hasRank ? l.rank : null,
      pastRank: p.hasRank ? p.rank : null,
      delta: delta,
      status: status
    });
  }

  var programs = [];
  for (var name in acc) {
    if (!Object.prototype.hasOwnProperty.call(acc, name)) continue;
    var a = acc[name];
    var n = a.comparable;
    programs.push({
      program: a.program,
      comparable: n,
      up: a.up, same: a.same, down: a.down,
      newCount: a.newCount, lostCount: a.lostCount,
      upRate: n > 0 ? a.up / n : null,
      sameRate: n > 0 ? a.same / n : null,
      downRate: n > 0 ? a.down / n : null,
      avgDelta: n > 0 ? a.deltaSum / n : null,
      avgUp: a.up > 0 ? a.upSum / a.up : null,
      avgDown: a.down > 0 ? a.downSum / a.down : null,
      latestAvgRank: a.latestRankCount > 0 ? a.latestRankSum / a.latestRankCount : null,
      soloComparable: a.soloComparable
    });
  }

  rankAnalysisSortPrograms_(programs);
  return { programs: programs, details: details };
}

/**
 * 반영도 정렬.
 *   1) 상승률 높은 순  2) 평균 순위 변화폭 큰 순  3) 하락률 낮은 순  4) 비교 항목 수 많은 순
 * 비교 가능 항목이 0개인 프로그램은 항상 뒤로 보낸다(비율이 없으므로 위로 올라오면 오해를 준다).
 */
function rankAnalysisSortPrograms_(programs) {
  programs.sort(function (a, b) {
    if ((a.comparable > 0) !== (b.comparable > 0)) return a.comparable > 0 ? -1 : 1;
    var d = rankAnalysisNum_(b.upRate) - rankAnalysisNum_(a.upRate);
    if (d !== 0) return d;
    d = rankAnalysisNum_(b.avgDelta) - rankAnalysisNum_(a.avgDelta);
    if (d !== 0) return d;
    d = rankAnalysisNum_(a.downRate) - rankAnalysisNum_(b.downRate);
    if (d !== 0) return d;
    d = b.comparable - a.comparable;
    if (d !== 0) return d;
    return a.program < b.program ? -1 : (a.program > b.program ? 1 : 0);
  });
  return programs;
}

function rankAnalysisNum_(v) { return (v === null || v === undefined || isNaN(v)) ? 0 : v; }

/** 기준일에 등장한 프로그램 수. */
function rankAnalysisCountPrograms_(obs, baseKey) {
  var latest = obs.byDate[baseKey] || {};
  var set = {};
  for (var k in latest) {
    if (!Object.prototype.hasOwnProperty.call(latest, k)) continue;
    var programs = obs.programsByItem[k] || [];
    for (var i = 0; i < programs.length; i++) set[programs[i]] = true;
  }
  return Object.keys(set).length;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 4. 분석 시트 쓰기
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 분석 시트를 통째로 다시 그린다. 같은 이름의 시트가 있으면 재사용한다. */
function rankAnalysisWriteAnalysisSheet_(ss, analysis, started) {
  var cfg = RANK_ANALYSIS_CONFIG;
  var sheet = rankAnalysisGetOrCreateSheet_(ss, cfg.ANALYSIS_SHEET_NAME);
  rankAnalysisResetSheet_(sheet);

  var width = RANK_ANALYSIS_TABLE_HEADERS.length;
  var grid = [];
  var marks = { summaryRows: 0, tables: [] };

  // ── 상단 요약 ────────────────────────────────────────────────────────────
  var best = rankAnalysisBestByPeriod_(analysis);
  grid.push(rankAnalysisPad2_(['프로그램별 순위 분석'], width));
  grid.push(rankAnalysisPad2_(['분석 기준일', analysis.baseKey], width));
  for (var i = 0; i < analysis.periods.length; i++) {
    var pd = analysis.periods[i];
    grid.push(rankAnalysisPad2_([
      '실제 ' + pd.days + '일 전 비교일',
      pd.pastKey ? pd.pastKey : '데이터 없음'
    ], width));
  }
  grid.push(rankAnalysisPad2_(['전체 프로그램 수', analysis.programCount], width));
  grid.push(rankAnalysisPad2_(['비교 가능한 전체 항목 수',
                               rankAnalysisTotalComparable_(analysis)], width));
  for (var b = 0; b < best.length; b++) {
    grid.push(rankAnalysisPad2_([best[b].days + '일 전 대비 반영도 1위', best[b].text], width));
  }
  grid.push(rankAnalysisPad2_(['마지막 분석 실행 시간',
    Utilities.formatDate(started, cfg.TIMEZONE, 'yyyy-MM-dd HH:mm:ss') +
    ' (' + cfg.TIMEZONE + ')'], width));
  grid.push(rankAnalysisPad2_(['원본 시트', analysis.sheetLabel], width));
  grid.push(rankAnalysisPad2_(['고유키 기준', rankAnalysisKeyDescription_(analysis)], width));
  marks.summaryRows = grid.length;
  grid.push(rankAnalysisPad2_([], width));

  // ── 기간별 표 ────────────────────────────────────────────────────────────
  for (var t = 0; t < analysis.periods.length; t++) {
    var period = analysis.periods[t];
    grid.push(rankAnalysisPad2_([
      period.days + '일 전 대비 분석' +
      (period.pastKey ? ' (기준 ' + analysis.baseKey + ' vs 비교 ' + period.pastKey + ')'
                      : ' — 데이터 없음')
    ], width));
    var titleRow = grid.length;                     // 1-based
    grid.push(RANK_ANALYSIS_TABLE_HEADERS.slice());
    var headerRow = grid.length;

    var rows = rankAnalysisTableRows_(period);
    for (var r = 0; r < rows.length; r++) grid.push(rows[r]);

    marks.tables.push({
      days: period.days, titleRow: titleRow, headerRow: headerRow,
      firstDataRow: headerRow + 1, rowCount: rows.length
    });
    grid.push(rankAnalysisPad2_([], width));
  }

  sheet.getRange(1, 1, grid.length, width).setValues(grid);
  rankAnalysisStyleAnalysisSheet_(sheet, marks, width);
  if (cfg.DRAW_CHART) rankAnalysisDrawChart_(sheet, marks, analysis);
  return sheet;
}

/** 한 기간의 표 본문을 만든다. 비교 불가면 안내 한 줄만 넣는다. */
function rankAnalysisTableRows_(period) {
  var width = RANK_ANALYSIS_TABLE_HEADERS.length;
  if (!period.pastKey) {
    return [rankAnalysisPad2_(['-', '비교할 과거 데이터가 없습니다 (데이터 없음)'], width)];
  }
  if (period.result.programs.length === 0) {
    return [rankAnalysisPad2_(['-', '비교 가능한 항목이 0개입니다'], width)];
  }

  var out = [];
  for (var i = 0; i < period.result.programs.length; i++) {
    var p = period.result.programs[i];
    out.push([
      i + 1, p.program, p.comparable, p.up, p.same, p.down,
      p.newCount, p.lostCount,
      rankAnalysisCellNum_(p.upRate), rankAnalysisCellNum_(p.sameRate),
      rankAnalysisCellNum_(p.downRate), rankAnalysisCellNum_(p.avgDelta),
      rankAnalysisCellNum_(p.avgUp), rankAnalysisCellNum_(p.avgDown),
      rankAnalysisCellNum_(p.latestAvgRank), p.soloComparable
    ]);
  }
  return out;
}

/** 값이 없으면 빈칸(계산 불가). 0 과 구분하기 위해 null 을 그대로 쓰지 않는다. */
function rankAnalysisCellNum_(v) { return (v === null || v === undefined) ? '' : v; }

/** 분석 시트 서식. */
function rankAnalysisStyleAnalysisSheet_(sheet, marks, width) {
  var summaryRows = marks.summaryRows;

  sheet.getRange(1, 1, 1, width).setFontSize(14).setFontWeight('bold');
  sheet.getRange(1, 1, summaryRows, width)
    .setBackground('#e8f0fe')
    .setFontWeight('normal');
  sheet.getRange(1, 1, summaryRows, 1).setFontWeight('bold');

  var rules = [];
  for (var t = 0; t < marks.tables.length; t++) {
    var tb = marks.tables[t];

    sheet.getRange(tb.titleRow, 1, 1, width)
      .setFontWeight('bold').setFontSize(12).setBackground('#f1f3f4');
    sheet.getRange(tb.headerRow, 1, 1, width)
      .setFontWeight('bold').setBackground('#d9e7fd').setWrap(true);

    if (tb.rowCount > 0) {
      var body = sheet.getRange(tb.firstDataRow, 1, tb.rowCount, width);
      body.setBorder(true, true, true, true, true, true,
                     '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

      // 상승률/유지율/하락률 = 9~11열, 퍼센트 소수점 1자리
      sheet.getRange(tb.firstDataRow, 9, tb.rowCount, 3).setNumberFormat('0.0%');
      // 평균 변화폭/상승폭/하락폭/최신 평균 순위 = 12~15열, 소수점 2자리
      sheet.getRange(tb.firstDataRow, 12, tb.rowCount, 4).setNumberFormat('0.00');

      // 상승 수(초록) / 유지 수(회색) / 하락 수(빨강) / 신규 수(파랑)
      sheet.getRange(tb.firstDataRow, 4, tb.rowCount, 1).setFontColor('#188038');
      sheet.getRange(tb.firstDataRow, 5, tb.rowCount, 1).setFontColor('#7f7f7f');
      sheet.getRange(tb.firstDataRow, 6, tb.rowCount, 1).setFontColor('#c5221f');
      sheet.getRange(tb.firstDataRow, 7, tb.rowCount, 1).setFontColor('#1a73e8');
      sheet.getRange(tb.firstDataRow, 9, tb.rowCount, 1).setFontColor('#188038');
      sheet.getRange(tb.firstDataRow, 11, tb.rowCount, 1).setFontColor('#c5221f');

      // 평균 순위 변화폭(12열) 조건부 서식: 양수 초록 / 0 회색 / 음수 빨강
      var deltaRange = sheet.getRange(tb.firstDataRow, 12, tb.rowCount, 1);
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setFontColor('#188038').setBackground('#e6f4ea')
        .setRanges([deltaRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(0).setFontColor('#5f6368').setBackground('#f1f3f4')
        .setRanges([deltaRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0).setFontColor('#c5221f').setBackground('#fce8e6')
        .setRanges([deltaRange]).build());
    }
  }
  if (rules.length > 0) sheet.setConditionalFormatRules(rules);

  // 요약 영역과 프로그램명 열을 고정해 스크롤해도 기준이 보이게 한다.
  sheet.setFrozenRows(Math.min(summaryRows, 10));
  sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, width);
}

/**
 * 상승·유지·하락 비율 누적 막대 차트.
 * 비교 가능한 기간 중 가장 긴 기간의 표를 쓴다(표본이 가장 안정적이므로).
 */
function rankAnalysisDrawChart_(sheet, marks, analysis) {
  var target = null;
  for (var i = 0; i < marks.tables.length; i++) {
    var tb = marks.tables[i];
    var period = rankAnalysisFindPeriod_(analysis, tb.days);
    if (!period || !period.pastKey || period.result.programs.length === 0) continue;
    if (!target || tb.days > target.days) target = tb;
  }
  if (!target) return;

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(target.headerRow, 2, target.rowCount + 1, 1))       // 프로그램명
    .addRange(sheet.getRange(target.headerRow, 9, target.rowCount + 1, 3))       // 상승/유지/하락률
    .setNumHeaders(1)
    .setOption('title', target.days + '일 전 대비 프로그램별 상승·유지·하락 비율')
    .setOption('isStacked', 'percent')
    .setOption('colors', ['#188038', '#9aa0a6', '#c5221f'])
    .setOption('legend', { position: 'top' })
    .setPosition(target.headerRow, RANK_ANALYSIS_TABLE_HEADERS.length + 2, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function rankAnalysisFindPeriod_(analysis, days) {
  for (var i = 0; i < analysis.periods.length; i++) {
    if (analysis.periods[i].days === days) return analysis.periods[i];
  }
  return null;
}

/** 상세 시트. 각 기간 × 각 항목의 변화를 그대로 나열해 검증에 쓴다. */
function rankAnalysisWriteDetailSheet_(ss, analysis) {
  var cfg = RANK_ANALYSIS_CONFIG;
  var sheet = rankAnalysisGetOrCreateSheet_(ss, cfg.DETAIL_SHEET_NAME);
  rankAnalysisResetSheet_(sheet);

  var width = RANK_ANALYSIS_DETAIL_HEADERS.length;
  var grid = [RANK_ANALYSIS_DETAIL_HEADERS.slice()];
  var truncated = false;

  for (var i = 0; i < analysis.periods.length; i++) {
    var period = analysis.periods[i];
    if (!period.pastKey) continue;
    var details = period.result.details;
    for (var d = 0; d < details.length; d++) {
      var item = details[d];
      for (var g = 0; g < item.programs.length; g++) {
        if (grid.length > cfg.MAX_DETAIL_ROWS) { truncated = true; break; }
        grid.push([
          item.programs[g],
          item.itemLabel,
          item.latestRank === null ? '순위 없음' : item.latestRank,
          item.pastRank === null ? '순위 없음' : item.pastRank,
          item.delta === null ? '' : item.delta,
          item.status,
          analysis.baseKey,
          period.pastKey,
          period.days + '일'
        ]);
      }
      if (truncated) break;
    }
    if (truncated) break;
  }

  if (truncated) {
    grid.push(['(이하 생략)', 'MAX_DETAIL_ROWS(' + cfg.MAX_DETAIL_ROWS +
               ')을 넘어 잘렸습니다', '', '', '', '', '', '', '']);
  }

  sheet.getRange(1, 1, grid.length, width).setValues(grid);
  sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9e7fd');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  if (grid.length > 1) {
    var body = sheet.getRange(2, 1, grid.length - 1, width);
    // 변화폭(5열) 조건부 서식 + 상태(6열) 색 구분
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setFontColor('#188038')
        .setRanges([sheet.getRange(2, 5, grid.length - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0).setFontColor('#c5221f')
        .setRanges([sheet.getRange(2, 5, grid.length - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('상승').setFontColor('#188038').setBackground('#e6f4ea')
        .setRanges([sheet.getRange(2, 6, grid.length - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('유지').setFontColor('#5f6368').setBackground('#f1f3f4')
        .setRanges([sheet.getRange(2, 6, grid.length - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('하락').setFontColor('#c5221f').setBackground('#fce8e6')
        .setRanges([sheet.getRange(2, 6, grid.length - 1, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('신규').setFontColor('#1a73e8').setBackground('#e8f0fe')
        .setRanges([sheet.getRange(2, 6, grid.length - 1, 1)]).build()
    ]);
    // 필터는 시트당 1개만 가능하므로, 실제로 걸러 볼 일이 많은 상세 시트에 건다.
    body.getFilter() && body.getFilter().remove();
    sheet.getRange(1, 1, grid.length, width).createFilter();
  }
  sheet.autoResizeColumns(1, width);
  return sheet;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 5. 시트 · 스프레드시트 접근
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 원본(B) 시트를 찾는다. SOURCE_SPREADSHEET_ID 가 비어 있으면 A 안에서 이름으로 찾는다. */
function rankAnalysisOpenSourceSheet_() {
  var cfg = RANK_ANALYSIS_CONFIG;
  var ss, where;

  if (cfg.SOURCE_SPREADSHEET_ID) {
    ss = rankAnalysisOpenSpreadsheet_(cfg.SOURCE_SPREADSHEET_ID, 'B스프레드시트');
    where = ss.getName();
  } else {
    ss = rankAnalysisOpenSpreadsheet_(cfg.TARGET_SPREADSHEET_ID, 'A스프레드시트');
    where = ss.getName();
  }

  var sheets = ss.getSheets();
  var i;
  if (cfg.SOURCE_SHEET_GID) {
    for (i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === cfg.SOURCE_SHEET_GID) {
        return { sheet: sheets[i], label: where + ' / ' + sheets[i].getName() };
      }
    }
  }
  if (cfg.SOURCE_SHEET_NAME) {
    var byName = ss.getSheetByName(cfg.SOURCE_SHEET_NAME);
    if (byName) return { sheet: byName, label: where + ' / ' + byName.getName() };
  }

  var names = [];
  for (i = 0; i < sheets.length; i++) {
    names.push(sheets[i].getName() + '(gid ' + sheets[i].getSheetId() + ')');
  }
  throw new Error('원본 시트를 찾을 수 없습니다. gid ' + cfg.SOURCE_SHEET_GID +
                  (cfg.SOURCE_SHEET_NAME ? ' / 이름 "' + cfg.SOURCE_SHEET_NAME + '"' : '') +
                  ' 로 찾았습니다. 존재하는 탭: ' + names.join(', ') +
                  '. RANK_ANALYSIS_CONFIG 의 SOURCE_SHEET_GID 를 시트 URL 끝의 #gid=... 와 맞추세요.');
}

/** 스프레드시트를 연다. 스크립트가 붙어 있는 문서면 openById 대신 getActive 를 쓴다. */
function rankAnalysisOpenSpreadsheet_(id, label) {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active && active.getId() === id) return active;

  try {
    var ss = SpreadsheetApp.openById(id);
    if (!ss) throw new Error('openById 가 null 을 반환했습니다.');
    return ss;
  } catch (e) {
    throw new Error(label + '(' + id + ')을 열 수 없습니다: ' + e.message +
                    ' — 승인한 계정에 접근 권한이 있는지, 프로젝트에 @OnlyCurrentDoc 주석이 ' +
                    '없는지, appsscript.json 의 oauthScopes 에 spreadsheets 권한이 있는지 확인하세요.');
  }
}

/** 이름으로 시트를 찾고 없으면 만든다. 다른 시트는 건드리지 않는다. */
function rankAnalysisGetOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  return sheet ? sheet : ss.insertSheet(name);
}

/** 시트 내용·서식·차트·필터를 비운다. 시트 자체는 지우지 않아 위치와 링크가 유지된다. */
function rankAnalysisResetSheet_(sheet) {
  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) sheet.removeChart(charts[i]);

  var filter = sheet.getFilter();
  if (filter) filter.remove();

  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.setConditionalFormatRules([]);
  sheet.clear();
}


/* ═══════════════════════════════════════════════════════════════════════════
 * 6. 리포트 · 유틸
 * ═══════════════════════════════════════════════════════════════════════════ */

/** 기간별 1위 프로그램을 사람이 읽는 문장으로. */
function rankAnalysisBestByPeriod_(analysis) {
  var out = [];
  for (var i = 0; i < analysis.periods.length; i++) {
    var period = analysis.periods[i];
    var text;
    if (!period.pastKey) {
      text = '데이터 없음';
    } else if (period.result.programs.length === 0 || period.result.programs[0].comparable === 0) {
      text = '비교 가능한 항목이 0개';
    } else {
      var top = period.result.programs[0];
      text = top.program +
             ' (상승률 ' + rankAnalysisPercent_(top.upRate) +
             ', 평균 변화폭 ' + rankAnalysisFixed_(top.avgDelta) +
             ', 비교 ' + top.comparable + '건)';
    }
    out.push({ days: period.days, text: text });
  }
  return out;
}

function rankAnalysisTotalComparable_(analysis) {
  var total = 0;
  for (var i = 0; i < analysis.periods.length; i++) {
    var programs = analysis.periods[i].result.programs;
    for (var p = 0; p < programs.length; p++) total += programs[p].comparable;
  }
  return total;
}

function rankAnalysisKeyDescription_(analysis) {
  var det = analysis.detection;
  if (det.keyCols.length === 0) return '식별 열을 못 찾아 행 번호를 고유키로 사용';
  return det.keyCols.map(function (c) { return rankAnalysisColumnLetter_(c) + '열'; })
                    .join(' + ') + ' 조합';
}

function rankAnalysisFormatReport_(analysis, started) {
  var cfg = RANK_ANALYSIS_CONFIG;
  var lines = [
    '=== 프로그램별 순위 분석 완료 ===',
    '실행 시각: ' + Utilities.formatDate(started, cfg.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
    '원본: ' + analysis.sheetLabel,
    '분석 기준일: ' + analysis.baseKey,
    '전체 프로그램 수: ' + analysis.programCount + '개'
  ];
  for (var i = 0; i < analysis.periods.length; i++) {
    var p = analysis.periods[i];
    lines.push('  ' + p.days + '일 전 비교일: ' + (p.pastKey || '데이터 없음') +
               (p.pastKey ? ' — 프로그램 ' + p.result.programs.length + '개' : ''));
  }
  var best = rankAnalysisBestByPeriod_(analysis);
  for (var b = 0; b < best.length; b++) {
    lines.push('  ' + best[b].days + '일 1위: ' + best[b].text);
  }
  var iss = analysis.issues;
  lines.push('데이터 이슈: 중복 ' + iss.duplicates + '건 / 프로그램명 없음 ' + iss.blankProgram +
             '행 / 날짜 오류 ' + iss.badDate + '행 / 순위 없음 ' + iss.noRank +
             '칸 / 고유키 없음 ' + iss.blankKey + '행');
  return lines.join('\n');
}

function rankAnalysisPercent_(v) {
  return (v === null || v === undefined) ? '-' : (Math.round(v * 1000) / 10) + '%';
}
function rankAnalysisFixed_(v) {
  return (v === null || v === undefined) ? '-' : (Math.round(v * 100) / 100).toFixed(2);
}

/** 배열을 지정 폭까지 빈 문자열로 채운다. setValues 는 행 길이가 같아야 한다. */
function rankAnalysisPad2_(arr, width) {
  var out = arr.slice();
  while (out.length < width) out.push('');
  return out;
}

function rankAnalysisDescribeCol_(header, col) {
  if (!col) return '(없음)';
  return rankAnalysisColumnLetter_(col) + '열 "' +
         rankAnalysisTrim_(header[col - 1]) + '"';
}

function rankAnalysisCell_(row, col) {
  if (!col) return '';
  var v = row[col - 1];
  return (v === null || v === undefined) ? '' : v;
}

function rankAnalysisTrim_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return Utilities.formatDate(
    value, RANK_ANALYSIS_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (typeof value === 'number' && isFinite(value) && value === Math.floor(value)) {
    return String(Math.round(value));
  }
  return String(value).trim();
}

function rankAnalysisColumnLetter_(col) {
  var s = '', n = col;
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 메뉴에서 실행했을 때만 팝업을 띄운다. 트리거/직접 실행에서는 조용히 넘어간다. */
function rankAnalysisAlert_(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* UI 컨텍스트가 아니면 무시 */ }
}

function rankAnalysisLog_(message) {
  try { Logger.log(message); } catch (e) { /* noop */ }
  try { console.log(message); } catch (e) { /* noop */ }
}
