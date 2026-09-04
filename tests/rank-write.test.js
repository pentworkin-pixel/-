/**
 * 「순위 분석.gs」의 시트 쓰기 경로 테스트.
 *
 * 가짜 스프레드시트를 붙여 updateProgramRankingAnalysis() 를 실제로 실행하고,
 * "B시트 원본과 사용자 시트를 건드리지 않는가", "분석 시트를 중복 생성하지 않는가",
 * "요약에 실제 비교일이 찍히는가" 를 검증한다.
 *
 * 실행: node tests/rank-write.test.js
 */
'use strict';

const { load, assertEqual, assertDeep, report } = require('./rank-harness');

const A_ID = '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8';
const B_ID = '1bLNh-zrYHHKWgH78ihbgnhpd11ItUhFgw2eT0MromFU';
const B_GID = 1003701754;

/* ── 가짜 시트 (체이닝 메서드는 Proxy 로 흘려보낸다) ─────────────────────── */

function fakeSheet(name, gid, grid) {
  const writes = [];
  const cells = grid.map((r) => r.slice());
  let filterCreated = 0;

  const sheet = {
    name, writes, cells,
    getName: () => name,
    getSheetId: () => gid,
    getLastRow: () => cells.length,
    getLastColumn: () => cells.reduce((m, r) => Math.max(m, r.length), 0),
    getCharts: () => [],
    removeChart: () => {},
    getFilter: () => null,
    setFrozenRows: () => {},
    setFrozenColumns: () => {},
    setConditionalFormatRules: () => {},
    autoResizeColumns: () => {},
    clear: () => { cells.length = 0; },
    insertChart: () => {},
    newChart: () => chainable({ build: () => ({}) }),
    get filterCount() { return filterCreated; },
    getRange(row, col, numRows, numCols) {
      const nR = numRows === undefined ? 1 : numRows;
      const nC = numCols === undefined ? 1 : numCols;
      return chainable({
        getValues() {
          const out = [];
          for (let r = 0; r < nR; r++) {
            const src = cells[row - 1 + r] || [];
            const line = [];
            for (let c = 0; c < nC; c++) {
              const v = src[col - 1 + c];
              line.push(v === undefined ? '' : v);
            }
            out.push(line);
          }
          return out;
        },
        setValues(values) {
          writes.push({ row, col, numRows: nR, numCols: nC });
          for (let r = 0; r < values.length; r++) {
            while (cells.length < row + r) cells.push([]);
            const target = cells[row - 1 + r];
            for (let c = 0; c < values[r].length; c++) {
              while (target.length < col - 1 + c) target.push('');
              target[col - 1 + c] = values[r][c];
            }
          }
          return this;
        },
        getFilter: () => null,
        createFilter: () => { filterCreated++; }
      });
    }
  };
  return sheet;
}

/**
 * 정의되지 않은 메서드는 자기 자신을 돌려주어 서식 체이닝(setFontWeight 등)을 흘려보낸다.
 * 테스트가 검증하는 것은 "무엇을 썼는가"이지 서식 호출 하나하나가 아니다.
 */
function chainable(base) {
  const proxy = new Proxy(base, {
    get(target, prop) {
      if (prop in target) return target[prop];
      return () => proxy;
    }
  });
  return proxy;
}

const chain = chainable;

function fakeSpreadsheet(id, name, sheets) {
  const inserted = [];
  return {
    inserted,
    getId: () => id,
    getName: () => name,
    getSheets: () => sheets.slice(),
    getSheetByName: (n) => sheets.filter((s) => s.getName() === n)[0] || null,
    insertSheet: (n) => {
      const s = fakeSheet(n, 900000 + sheets.length, []);
      sheets.push(s);
      inserted.push(n);
      return s;
    }
  };
}

function fakeSpreadsheetApp(byId) {
  const rule = () => chain({ build: () => ({}) });
  return {
    BorderStyle: { SOLID: 'SOLID' },
    getActiveSpreadsheet: () => null,
    getUi: () => { throw new Error('no ui'); },
    openById: (id) => {
      if (!byId[id]) throw new Error('없는 문서: ' + id);
      return byId[id];
    },
    newConditionalFormatRule: rule,
    flush: () => {}
  };
}

/* ── 데이터 ───────────────────────────────────────────────────────────────
 * 가로형: E(5)=프로그램, G(7)=키워드, H(8)=상품 MID, I~K(9~11)=날짜별 순위
 * 기준일 2026-08-20, 1일 전 2026-08-19, 7일 전 목표(08-13) 데이터 있음, 3일 전(08-17)은 없음
 * ───────────────────────────────────────────────────────────────────────── */

const B_HEADER = ['메모1', 'b', '판매처명', '상품링크', '프로그램',
                  '전일 비교', '순위 조회 키워드', '상품 MID', '26-08-20', '26-08-19', '26-08-13'];
const B_ROWS = [
  ['메모A', '', '', '', '리뷰 3개', '', '마스크', 111, 3, 10, 12],   // 상승
  ['메모B', '', '', '', '리뷰 2개', '', '멀티탭', 222, 5, 5, 4],     // 유지 / 7일은 하락
  ['메모C', '', '', '', '트래픽 1개', '', '광케이블', 333, 9, 4, 4], // 하락
  ['메모D', '', '', '', '', '', '프로그램없음', 444, 1, 1, 1]        // 프로그램 미기재 → 제외
];

function setup() {
  const bSheet = fakeSheet('펀트미디어 순위추적2', B_GID, [B_HEADER].concat(B_ROWS));
  const userSheet = fakeSheet('사용자가 만든 시트', 555, [['건드리면', '안 됨']]);
  const calendarSheet = fakeSheet('08월 20일 종료', 666, [['기존', '데이터']]);

  const aSheets = [userSheet, calendarSheet];
  const aSs = fakeSpreadsheet(A_ID, '김승환 현황', aSheets);
  const bSs = fakeSpreadsheet(B_ID, '펀트미디어 순위추적', [bSheet]);

  const gs = load({
    SpreadsheetApp: fakeSpreadsheetApp({ [A_ID]: aSs, [B_ID]: bSs })
  });
  // 이 파일의 픽스처는 단순화를 위해 헤더를 1행에 둔다. 실제 B시트(헤더 6행,
  // 999 상한값)에 대한 end-to-end 검증은 아래 "실제 B시트 레이아웃" 절 참고.
  gs.RANK_ANALYSIS_CONFIG.HEADER_ROW = 1;

  return { gs, aSs, bSs, bSheet, userSheet, calendarSheet };
}

/* ── 실행 ─────────────────────────────────────────────────────────────────── */

{
  const { gs, aSs, bSheet, userSheet, calendarSheet } = setup();
  const report0 = gs.updateProgramRankingAnalysis();

  // 원본 보호
  assertEqual(bSheet.writes.length, 0, 'B시트 원본에 쓰기가 전혀 없다');
  assertEqual(userSheet.writes.length, 0, '사용자가 만든 시트를 건드리지 않는다');
  assertEqual(calendarSheet.writes.length, 0, '기존 캘린더용 시트도 건드리지 않는다');
  assertDeep(userSheet.cells, [['건드리면', '안 됨']], '사용자 시트 내용 그대로');

  // 분석 시트 생성
  assertDeep(aSs.inserted, ['프로그램별 순위 분석', '프로그램별 순위 상세'],
             '분석 시트와 상세 시트를 A스프레드시트에 만든다');

  const analysis = aSs.getSheetByName('프로그램별 순위 분석');
  const flat = analysis.cells.map((r) => r.join('|')).join('\n');

  assertEqual(flat.indexOf('2026-08-20') >= 0, true, '분석 기준일이 요약에 있다');
  assertEqual(flat.indexOf('분석 기준일') >= 0, true, '요약 항목: 분석 기준일');
  assertEqual(flat.indexOf('실제 1일 전 비교일') >= 0, true, '요약 항목: 실제 1일 전 비교일');
  assertEqual(flat.indexOf('실제 3일 전 비교일') >= 0, true, '요약 항목: 실제 3일 전 비교일');
  assertEqual(flat.indexOf('실제 7일 전 비교일') >= 0, true, '요약 항목: 실제 7일 전 비교일');
  assertEqual(flat.indexOf('마지막 분석 실행 시간') >= 0, true, '요약 항목: 실행 시간');
  assertEqual(flat.indexOf('1일 전 대비 분석') >= 0, true, '1일 전 표 제목');
  assertEqual(flat.indexOf('3일 전 대비 분석') >= 0, true, '3일 전 표 제목');
  assertEqual(flat.indexOf('7일 전 대비 분석') >= 0, true, '7일 전 표 제목');

  // 3일 전 목표일(08-17) 데이터가 없으므로 그 이전인 08-13 을 쓴다
  assertEqual(report0.indexOf('3일 전 비교일: 2026-08-13') >= 0, true,
              '3일 전 데이터가 없으면 그보다 이전인 08-13 을 쓴다');
  assertEqual(report0.indexOf('1일 전 비교일: 2026-08-19') >= 0, true, '1일 전은 08-19');
  assertEqual(report0.indexOf('7일 전 비교일: 2026-08-13') >= 0, true, '7일 전은 08-13');

  // 프로그램명이 없는 행은 어느 프로그램에도 귀속되지 않는다
  assertEqual(flat.indexOf('프로그램없음') < 0, true,
              '프로그램명이 빈 행은 분석표에 나타나지 않는다');
}

{
  // 두 번 실행해도 시트를 새로 만들지 않는다
  const { gs, aSs } = setup();
  gs.updateProgramRankingAnalysis();
  const afterFirst = aSs.getSheets().length;
  gs.updateProgramRankingAnalysis();

  assertEqual(aSs.getSheets().length, afterFirst, '재실행해도 시트가 늘지 않는다');
  assertDeep(aSs.inserted, ['프로그램별 순위 분석', '프로그램별 순위 상세'],
             '시트 생성은 최초 1회뿐 — 이후에는 내용만 갱신');
}

{
  // 상세 시트에 기간별 행이 쌓이는지
  const { gs, aSs } = setup();
  gs.updateProgramRankingAnalysis();
  const detail = aSs.getSheetByName('프로그램별 순위 상세');
  const rows = detail.cells;

  assertDeep(rows[0], [
    '프로그램명', '고유 항목명', '최신 순위', '과거 순위', '순위 변화폭',
    '상태', '최신 업데이트 날짜', '실제 비교 날짜', '비교 기간'
  ], '상세 시트 머리글');

  const periods = {};
  for (let i = 1; i < rows.length; i++) periods[rows[i][8]] = true;
  assertDeep(Object.keys(periods).sort(), ['1일', '3일', '7일'],
             '상세 시트에 1·3·7일 행이 모두 있다');

  const upRow = rows.filter((r) => r[1] === '111 ｜ 마스크' && r[8] === '1일')[0];
  assertDeep(upRow.slice(0, 6), ['리뷰', '111 ｜ 마스크', 3, 10, 7, '상승'],
             '10위 → 3위는 +7 상승으로 기록된다');
}

{
  // 비교할 과거 데이터가 아예 없는 경우에도 죽지 않는다
  const bSheet = fakeSheet('순위', B_GID, [
    ['메모1', 'b', 'c', 'd', '프로그램', 'f', '순위 조회 키워드', '상품 MID', '26-08-20'],
    ['메모', '', '', '', '리뷰 1개', '', '마스크', 111, 3]
  ]);
  const aSs = fakeSpreadsheet(A_ID, '김승환 현황', []);
  const bSs = fakeSpreadsheet(B_ID, '순위추적', [bSheet]);
  const gs = load({ SpreadsheetApp: fakeSpreadsheetApp({ [A_ID]: aSs, [B_ID]: bSs }) });
  gs.RANK_ANALYSIS_CONFIG.HEADER_ROW = 1;   // 이 픽스처는 헤더가 1행

  const out = gs.updateProgramRankingAnalysis();
  assertEqual(out.indexOf('1일 전 비교일: 데이터 없음') >= 0, true,
              '과거 데이터가 없으면 "데이터 없음"으로 표시하고 억지로 계산하지 않는다');

  const analysis = aSs.getSheetByName('프로그램별 순위 분석');
  const flat = analysis.cells.map((r) => r.join('|')).join('\n');
  assertEqual(flat.indexOf('비교할 과거 데이터가 없습니다') >= 0, true,
              '표 본문에도 안내 문구가 들어간다');
}

{
  // 날짜를 하나도 못 찾으면 명확히 실패한다 (조용히 빈 시트를 만들지 않는다)
  const bSheet = fakeSheet('순위', B_GID, [
    ['메모1', 'b', 'c', 'd', '프로그램', 'f', '키워드', '상품 MID', '순위'],
    ['메모', '', '', '', '리뷰', '', '마스크', 111, 3]
  ]);
  const aSs = fakeSpreadsheet(A_ID, '김승환 현황', []);
  const bSs = fakeSpreadsheet(B_ID, '순위추적', [bSheet]);
  const gs = load({ SpreadsheetApp: fakeSpreadsheetApp({ [A_ID]: aSs, [B_ID]: bSs }) });
  gs.RANK_ANALYSIS_CONFIG.HEADER_ROW = 1;   // 이 픽스처는 헤더가 1행

  let threw = false;
  let message = '';
  try { gs.updateProgramRankingAnalysis(); } catch (e) { threw = true; message = e.message; }
  assertEqual(threw, true, '날짜 열을 못 찾으면 오류를 던진다');
  assertEqual(message.indexOf('필수 열을 찾지 못했습니다') >= 0, true,
              '실제로 "날짜 열 못 찾음" 오류인지 확인 (다른 이유로 던진 게 아님)');
  assertDeep(aSs.inserted, [], '실패했으면 분석 시트를 만들지 않는다');
}

/* ═══ 실제 B시트 레이아웃 — 기본 설정 그대로 end-to-end 검증 ═══════════════
 * 헤더 6행(1~5행은 봇 안내문), "리워드명 / 갯수" 프로그램 열, 하루 여러 체크
 * (같은 날짜에 열 여러 개), 999 상한값 — 실제 배포될 기본 설정(RANK_ANALYSIS_CONFIG)을
 * 그대로 써서 검증한다. cfg 를 손으로 바꾸지 않는다.
 * ═══════════════════════════════════════════════════════════════════════════ */
{
  var pad = (n) => new Array(n).fill('');
  const realHeader = ['상품명', 'b', 'c', '메모1', '리워드명 / 갯수', '판매처명', '상품링크',
                      '4개전', '전일 비교 순위(자유 입력)', '순위 조회 키워드',
                      '상품 MID', '가격비교 MID',
                      '26-09-04 16:12', '26-09-04 14:31', '26-09-03 23:36', '26-08-28 08:14'];
  // D=메모1(4) E=프로그램(5) ... J=키워드(10) K=MID(11) ... M~P=날짜(13~16)
  function realRow(name, mid, keyword, m, n, o, p) {
    const row = pad(16);
    row[3] = '메모'; row[4] = name; row[9] = keyword; row[10] = mid;
    row[12] = m; row[13] = n; row[14] = o; row[15] = p;
    return row;
  }

  // 1~5행은 봇 안내문(더미), 6행이 진짜 헤더(HEADER_ROW=6과 일치), 7행부터 데이터.
  const bRows = [
    ['코알라'], [], ['', '', 'mid값 찾는 방법(클릭)'], [], [],
    realHeader,
    realRow('리뷰 3개', '111', '마스크', 3, 3, 10, 12),      // 3(9/4) vs 10(9/3, 1일전) → +7 상승
    realRow('트래픽 1개', '222', '멀티탭', 999, 999, 10, 5)  // 최신(9/4)이 999=상한 → 순위 없음
  ];

  const bSheet = fakeSheet('펀트미디어 순위추적2', B_GID, bRows);
  const aSs = fakeSpreadsheet(A_ID, '김승환 현황', []);
  const bSs = fakeSpreadsheet(B_ID, '펀트미디어 순위추적', [bSheet]);
  const gs = load({ SpreadsheetApp: fakeSpreadsheetApp({ [A_ID]: aSs, [B_ID]: bSs }) });

  // cfg 는 손대지 않는다 — 배포될 기본값 그대로 검증한다.
  assertEqual(gs.RANK_ANALYSIS_CONFIG.HEADER_ROW, 6, '기본 HEADER_ROW=6 확인');
  assertDeep(gs.RANK_ANALYSIS_CONFIG.RANK_CAP_VALUES, [999], '기본 RANK_CAP_VALUES=[999] 확인');

  const out = gs.updateProgramRankingAnalysis();
  const analysis = aSs.getSheetByName('프로그램별 순위 분석');
  const detail = aSs.getSheetByName('프로그램별 순위 상세');
  const flat = analysis.cells.map((r) => r.join('|')).join('\n');

  assertEqual(out.indexOf('기준일: 2026-09-04') >= 0, true,
              '같은 날짜의 여러 시각 열 중 가장 늦은 16:12 열 기준으로 09-04가 기준일이 된다');
  assertEqual(flat.indexOf('리뷰') >= 0, true,
              '"리워드명 / 갯수" 헤더에서 프로그램명이 정상 추출된다');

  const trafficRow1d = detail.cells.filter((r) => r[1] === '222 ｜ 멀티탭' && r[8] === '1일')[0];
  assertEqual(trafficRow1d[2], '순위 없음',
              '최신값 999는 상한 취급되어 상세 시트에도 "순위 없음"으로 기록된다');
  assertEqual(trafficRow1d[5], '누락 또는 종료',
              '999(상한) 때문에 최신 순위가 없다고 보고 "누락 또는 종료"로 분류한다');
}

report('순위 분석 · 시트 쓰기');
