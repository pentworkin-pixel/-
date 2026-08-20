/**
 * 「프로그램 동기화.gs」의 동기화 · 설치 동작 테스트.
 *
 * 가짜 스프레드시트를 붙여 실제 progSyncRun / progSyncInstall 을 실행하고,
 * "B시트 E열 외에는 아무것도 건드리지 않는가", "트리거가 정확히 1개 생기는가" 를 검증한다.
 *
 * 실행: node tests/program-sync.test.js
 */
'use strict';

const {
  load, fakeSheet, fakeSpreadsheet, fakeSpreadsheetApp, fakeScriptApp,
  assertEqual, assertDeep, assertThrows, report
} = require('./program-harness');

const A_ID = '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8';
const A_GID = 1948367989;
const B_ID = '1bLNh-zrYHHKWgH78ihbgnhpd11ItUhFgw2eT0MromFU';
const B_GID = 1003701754;

/** A시트 한 행: D(4)=프로그램명, E(5)=갯수, Q(17)=상품키 */
function aRow(program, count, productKey) {
  const row = new Array(17).fill('');
  row[3] = program;
  row[4] = count;
  row[16] = productKey;
  return row;
}
function aHeader() {
  const row = new Array(17).fill('');
  row[3] = '프로그램명';
  row[4] = '수량';
  row[16] = '상품 MID';
  return row;
}

/** B시트 한 행: E(5)=출력, K(11)=상품키. 다른 열에는 지워지면 안 되는 값을 넣어 둔다. */
function bRow(current, productKey, marker) {
  const row = new Array(11).fill('');
  row[0] = marker + '-메모';      // A열
  row[1] = marker + '-리워드';    // B열
  row[4] = current;               // E열
  row[6] = marker + '-키워드';    // G열
  row[10] = productKey;           // K열
  return row;
}
function bHeader() {
  const row = new Array(11).fill('');
  row[0] = '메모1';
  row[4] = '프로그램';
  row[10] = '상품 MID';
  return row;
}

/** A/B 가짜 시트를 붙인 컨텍스트를 만든다. */
function setup(aRows, bRows) {
  const aSheet = fakeSheet('프로그램 현황', A_GID, [aHeader()].concat(aRows));
  const bSheet = fakeSheet('펀트미디어 순위추적2', B_GID, [bHeader()].concat(bRows));
  const scriptApp = fakeScriptApp();

  const gs = load({
    SpreadsheetApp: fakeSpreadsheetApp({
      [A_ID]: fakeSpreadsheet(A_ID, '김승환 현황', [aSheet]),
      [B_ID]: fakeSpreadsheet(B_ID, '펀트미디어 순위추적', [bSheet])
    }),
    ScriptApp: scriptApp
  });

  return { gs, aSheet, bSheet, scriptApp };
}

/* ── 동기화 본체 ──────────────────────────────────────────────────────────── */

{
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111), aRow('트래픽', 2, 111), aRow('찜', 5, 222)],
    [bRow('', 111, 'r2'), bRow('', 222, 'r3')]
  );
  const result = gs.progSyncRun();

  assertEqual(result.updates.length, 2, '두 행이 갱신된다');
  assertEqual(bSheet.cells[1][4], '리뷰 3개, 트래픽 2개', '2행 E열: 프로그램 2종이 합쳐진다');
  assertEqual(bSheet.cells[2][4], '찜 5개', '3행 E열');
  assertEqual(result.writeCalls, 1, '연속된 2행이므로 쓰기 호출은 1회');
}

{
  // E열 외의 다른 열은 절대 건드리면 안 된다.
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  gs.progSyncRun();

  assertEqual(bSheet.cells[1][0], 'r2-메모', 'A열 보존');
  assertEqual(bSheet.cells[1][1], 'r2-리워드', 'B열 보존');
  assertEqual(bSheet.cells[1][6], 'r2-키워드', 'G열 보존');
  assertEqual(bSheet.cells[1][10], 111, 'K열 보존');

  const cols = bSheet.writes.map((w) => w.col);
  assertDeep(cols, [5], '쓰기는 E열(5)에만 일어난다');
  const widths = bSheet.writes.map((w) => w.numCols);
  assertDeep(widths, [1], '쓰기 폭은 항상 1열');
}

{
  // 이미 값이 같으면 쓰기 호출이 아예 없어야 한다.
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('리뷰 3개', 111, 'r2')]
  );
  const result = gs.progSyncRun();

  assertEqual(result.updates.length, 0, '변경할 것이 없다');
  assertEqual(result.unchanged, 1, '유지 1행');
  assertEqual(bSheet.writes.length, 0, '쓰기 호출 0회');
}

{
  // A시트에 없는 상품(999)의 행은 손대지 않고, 그 앞뒤 행만 갱신한다.
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111), aRow('찜', 5, 222)],
    [bRow('', 111, 'r2'), bRow('손으로 적은 값', 999, 'r3'), bRow('', 222, 'r4')]
  );
  const result = gs.progSyncRun();

  assertEqual(bSheet.cells[2][4], '손으로 적은 값', 'A시트에 없는 상품의 E열은 보존된다');
  assertEqual(bSheet.cells[1][4], '리뷰 3개', '앞 행은 갱신');
  assertEqual(bSheet.cells[3][4], '찜 5개', '뒤 행도 갱신');
  assertEqual(result.writeCalls, 2, '중간이 끊겼으므로 쓰기 호출 2회');
  assertDeep(result.unmatchedKeys, ['999'], '매칭 실패 키 보고');

  const rowsWritten = bSheet.writes.map((w) => w.row);
  assertDeep(rowsWritten, [2, 4], '3행은 쓰기 대상에 포함되지 않는다');
}

{
  // K열이 빈 행은 건너뛴다.
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('그대로', '', 'r2'), bRow('', 111, 'r3')]
  );
  const result = gs.progSyncRun();

  assertEqual(result.blankKeyRows, 1, 'K열이 빈 행 1개');
  assertEqual(bSheet.cells[1][4], '그대로', 'K열이 비면 E열을 건드리지 않는다');
  assertEqual(bSheet.cells[2][4], '리뷰 3개', '나머지는 정상 갱신');
}

{
  // A시트가 바뀌면(갯수 3 → 7) 다음 실행에서 덮어써야 한다.
  const { gs, bSheet } = setup(
    [aRow('리뷰', 7, 111)],
    [bRow('리뷰 3개', 111, 'r2')]
  );
  gs.progSyncRun();
  assertEqual(bSheet.cells[1][4], '리뷰 7개', '옛 값을 새 값으로 덮어쓴다');
}

{
  // 같은 데이터로 두 번 실행해도 두 번째는 쓰기가 없어야 한다(멱등).
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  gs.progSyncRun();
  const writesAfterFirst = bSheet.writes.length;
  const second = gs.progSyncRun();

  assertEqual(writesAfterFirst, 1, '첫 실행은 1회 쓴다');
  assertEqual(bSheet.writes.length, 1, '두 번째 실행은 추가 쓰기가 없다 (멱등)');
  assertEqual(second.updates.length, 0, '두 번째 실행에는 변경 대상이 없다');
}

/* ── 미리보기는 절대 쓰지 않는다 ──────────────────────────────────────────── */

{
  const { gs, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  const out = gs.progSyncDryRun();

  assertEqual(bSheet.writes.length, 0, 'progSyncDryRun 은 시트를 쓰지 않는다');
  assertEqual(bSheet.cells[1][4], '', 'E열이 그대로다');
  assertEqual(out.indexOf('E2') >= 0, true, '무엇이 바뀔지 미리 보여준다');
}

/* ── 설치 ─────────────────────────────────────────────────────────────────── */

{
  const { gs, scriptApp, bSheet } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  const out = gs.progSyncInstall();

  assertEqual(scriptApp.triggers.length, 1, '트리거가 정확히 1개 생성된다');
  assertEqual(scriptApp.triggers[0].fn, 'progSyncRun', '트리거가 부르는 함수');
  assertEqual(scriptApp.triggers[0].hour, 10, '매일 오전 10시');
  assertEqual(scriptApp.triggers[0].days, 1, '매일');
  assertEqual(scriptApp.triggers[0].tz, 'Asia/Seoul', '서울 시간 기준');
  assertEqual(bSheet.cells[1][4], '리뷰 3개', '설치 직후 1회 동기화까지 끝난다');
  assertEqual(out.indexOf('설치 완료') >= 0, true, '완료 리포트를 남긴다');
}

{
  // 이미 트리거가 여러 개 있어도 설치 후에는 항상 1개여야 한다.
  const { gs, scriptApp } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  scriptApp.newTrigger('progSyncRun').timeBased().atHour(3).everyDays(1).create();
  scriptApp.newTrigger('progSyncRun').timeBased().atHour(4).everyDays(1).create();
  assertEqual(scriptApp.triggers.length, 2, '설치 전 중복 트리거 2개');

  gs.progSyncInstall();
  assertEqual(scriptApp.triggers.length, 1, '중복 트리거를 정리하고 1개만 남긴다');
  assertEqual(scriptApp.triggers[0].hour, 10, '남은 트리거는 10시짜리');
}

{
  // 다른 스크립트의 트리거는 건드리지 않는다.
  const { gs, scriptApp } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  scriptApp.newTrigger('adEndSyncRun').timeBased().atHour(8).everyDays(1).create();
  gs.progSyncInstall();

  const handlers = scriptApp.triggers.map((t) => t.fn).sort();
  assertDeep(handlers, ['adEndSyncRun', 'progSyncRun'], '기존 캘린더 동기화 트리거는 그대로 둔다');
}

{
  const { gs, scriptApp } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  scriptApp.newTrigger('adEndSyncRun').timeBased().atHour(8).everyDays(1).create();
  gs.progSyncInstall();
  gs.progSyncUninstall();

  const handlers = scriptApp.triggers.map((t) => t.fn);
  assertDeep(handlers, ['adEndSyncRun'], 'progSyncUninstall 은 자기 트리거만 지운다');
}

/* ── 검증 함수 ────────────────────────────────────────────────────────────── */

{
  const { gs } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  const before = gs.progSyncVerifyMatch();
  assertEqual(before.indexOf('불일치: 1행') >= 0, true, '동기화 전에는 불일치를 보고한다');

  gs.progSyncRun();
  const after = gs.progSyncVerifyMatch();
  assertEqual(after.indexOf('완전히 일치') >= 0, true, '동기화 후에는 일치를 보고한다');
}

{
  const { gs } = setup(
    [aRow('리뷰', 3, 111)],
    [bRow('', 111, 'r2')]
  );
  const out = gs.progSyncCheckColumns();
  assertEqual(out.indexOf('D열(프로그램명) 헤더: "프로그램명"') >= 0, true, 'A시트 D열 매핑 표시');
  assertEqual(out.indexOf('Q열(상품키) 헤더: "상품 MID"') >= 0, true, 'A시트 Q열 매핑 표시');
  assertEqual(out.indexOf('K열(상품키) 헤더: "상품 MID"') >= 0, true, 'B시트 K열 매핑 표시');
  assertEqual(out.indexOf('E열(기록될 열)') >= 0, true, 'B시트 E열 매핑 표시');
}

/* ── 잘못된 설정은 조용히 넘어가지 않는다 ─────────────────────────────────── */

{
  const { gs } = setup([aRow('리뷰', 3, 111)], [bRow('', 111, 'r2')]);
  gs.PROG_SYNC_CONFIG.TARGET_SHEET_GID = 999999;
  assertThrows(() => gs.progSyncRun(), 'gid가 틀리면 명확한 오류를 던진다');
}

{
  const { gs } = setup([aRow('리뷰', 3, 111)], [bRow('', 111, 'r2')]);
  gs.PROG_SYNC_CONFIG.SOURCE_SPREADSHEET_ID = '없는아이디';
  assertThrows(() => gs.progSyncRun(), '문서를 못 열면 명확한 오류를 던진다');
}

/* ── 데이터가 비어도 죽지 않는다 ──────────────────────────────────────────── */

{
  const { gs, bSheet } = setup([], []);
  const result = gs.progSyncRun();
  assertEqual(result.updates.length, 0, '빈 시트에서도 정상 동작');
  assertEqual(bSheet.writes.length, 0, '쓰기 없음');
}

report('프로그램 동기화 · 동기화/설치');
