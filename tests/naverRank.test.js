/**
 * 「네이버 쇼핑 순위 조회.gs」의 순수 로직 테스트.
 *
 * 배포될 .gs 파일을 그대로 읽어 Apps Script 전역(UrlFetchApp/SpreadsheetApp/ScriptApp 등)만
 * 스텁으로 채운 뒤 실행한다. 즉 사본이 아니라 실제 배포 코드를 검증한다.
 *
 * 주의: 여기서 쓰는 __NEXT_DATA__ 픽스처는 실제 네이버 응답을 캡처한 것이 아니라
 * (네트워크가 봇 차단으로 막혀 있어 실물 응답을 확보할 수 없었음), 이 파일이
 * 상정하는 "이름 필드 + id 필드를 가진 객체 배열" 모양을 흉내 낸 것이다.
 * 즉 이 테스트는 파싱/광고제외/순위계산/MID매칭 로직 자체의 정확성을 검증하는
 * 것이지, 네이버의 실제 페이지 구조와 100% 일치함을 보장하지는 않는다.
 * 실제 구조가 다르면 naverRankDebugFetch()로 확인 후 조정해야 한다.
 *
 * 실행: node tests/naverRank.test.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GS_PATH = path.join(__dirname, '..', '네이버 쇼핑 순위 조회.gs');

/* ── Apps Script 전역 스텁 ────────────────────────────────────────────────── */

const logs = [];
let fetchResponses = {}; // keyword -> response 객체 또는 그 객체를 만드는 함수

function makeResponse(code, text) {
  return { getResponseCode: () => code, getContentText: () => text };
}

function makeScriptAppStub() {
  let triggers = [];
  return {
    newTrigger: (fn) => {
      const builder = { _fn: fn, _hour: null };
      builder.timeBased = () => builder;
      builder.atHour = (h) => { builder._hour = h; return builder; };
      builder.everyDays = () => builder;
      builder.inTimezone = () => builder;
      builder.create = () => {
        const trigger = {
          getHandlerFunction: () => builder._fn,
          getTriggerSourceId: () => 'hour:' + builder._hour,
        };
        triggers.push(trigger);
        return trigger;
      };
      return builder;
    },
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => { triggers = triggers.filter((x) => x !== t); },
    _reset: () => { triggers = []; },
  };
}

const sandbox = {
  Utilities: { sleep: () => {} },
  Logger: { log: (m) => logs.push(String(m)) },
  console: { log: () => {} },
  SpreadsheetApp: {},
  ScriptApp: makeScriptAppStub(),
  UrlFetchApp: {
    fetch: (url) => {
      const m = /query=([^&]+)/.exec(url);
      const keyword = decodeURIComponent(m[1]);
      const resp = fetchResponses[keyword];
      if (!resp) throw new Error('테스트에 정의되지 않은 키워드 fetch: ' + keyword);
      return typeof resp === 'function' ? resp() : resp;
    },
  },
  Date, Math, JSON, String, Number, Object, isNaN, isFinite, RegExp, Array,
  encodeURIComponent, decodeURIComponent,
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(GS_PATH, 'utf8'), sandbox, { filename: GS_PATH });

/* ── 픽스처: __NEXT_DATA__ 모양의 검색 결과 ──────────────────────────────────── */

function buildFixtureHtml(list) {
  const data = {
    props: {
      pageProps: {
        initialState: {
          collection: {
            list: list,
            unrelatedArray: ['a', 'b', 'c'],
            filters: [{ name: '필터1' }, { name: '필터2' }],
          },
        },
      },
    },
  };
  return `<html><body><script id="__NEXT_DATA__" type="application/json">${JSON.stringify(data)}</script></body></html>`;
}

// 순서대로: 광고(isAd) / 광고(cardType) / 정상 / 광고(adId) / 정상 / 정상
const SAMPLE_LIST = [
  { productId: 111, productName: '무선 이어폰 프로', mallName: '스마트몰', isAd: true },
  { productId: 222, productName: '무선 이어폰 프로 케이스', mallName: '가짓A몰', cardType: 'AD' },
  { productId: 333, title: '무선 이어폰 프로 정품', mallName: '정품몰' },
  { productId: 444, productName: '무선 이어폰 프로 (리퍼)', mallName: '리퍼몰', adId: 'xyz123' },
  { productId: 555, productName: '완전 다른 상품', mallName: '다른몰' },
  { productId: 666, productName: '무선 이어폰 프로', mallName: '베스트몰' },
];

const SAMPLE_HTML = buildFixtureHtml(SAMPLE_LIST);

/* ── 가짜 시트 (열 삽입 지원) ─────────────────────────────────────────────── */

function fakeSheet(initialGrid) {
  let grid = initialGrid.map((r) => r.slice());
  const numberFormats = [];

  return {
    getLastRow: () => grid.length,
    insertColumnBefore: (col) => {
      grid = grid.map((r) => {
        const copy = r.slice();
        copy.splice(col - 1, 0, '');
        return copy;
      });
    },
    getRange: (row, col, numRows, numCols) => {
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        getValues: () => {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const rowArr = [];
            for (let c = 0; c < numCols; c++) {
              const v = (grid[row - 1 + r] || [])[col - 1 + c];
              rowArr.push(v === undefined ? '' : v);
            }
            out.push(rowArr);
          }
          return out;
        },
        setValues: (values) => {
          for (let r = 0; r < values.length; r++) {
            if (!grid[row - 1 + r]) grid[row - 1 + r] = [];
            for (let c = 0; c < values[r].length; c++) {
              grid[row - 1 + r][col - 1 + c] = values[r][c];
            }
          }
        },
        setValue: (v) => {
          if (!grid[row - 1]) grid[row - 1] = [];
          grid[row - 1][col - 1] = v;
        },
        setNumberFormat: (fmt) => { numberFormats.push({ row, col, fmt }); },
      };
    },
    _grid: () => grid,
    _numberFormats: numberFormats,
  };
}

/* ── 미니 테스트 러너 ─────────────────────────────────────────────────────── */

let passed = 0;
const failures = [];

function test(name, fn) {
  try { fn(); passed++; console.log('  ✅ ' + name); }
  catch (e) { failures.push({ name, message: e.message }); console.log('  ❌ ' + name + '\n       ' + e.message); }
}

function eq(actual, expected, label) {
  const a = JSON.stringify(actual), b = JSON.stringify(expected);
  if (a !== b) throw new Error((label ? label + ': ' : '') + '기대 ' + b + ' / 실제 ' + a);
}

function ok(cond, label) {
  if (!cond) throw new Error(label || '조건이 참이 아님');
}

/* ── 테스트 ───────────────────────────────────────────────────────────────── */

console.log('\n[1] __NEXT_DATA__ 추출 및 상품 배열 탐색');

test('정상 HTML에서 JSON을 추출한다', () => {
  const data = sandbox.naverRankExtractNextData_(SAMPLE_HTML);
  ok(data && data.props, 'props 없음');
});

test('__NEXT_DATA__가 없으면 null', () => {
  eq(sandbox.naverRankExtractNextData_('<html><body>차단됨</body></html>'), null);
});

test('깨진 JSON이면 null', () => {
  const bad = '<script id="__NEXT_DATA__" type="application/json">{not json}</script>';
  eq(sandbox.naverRankExtractNextData_(bad), null);
});

test('상품 목록 배열을 정확히 찾는다 (무관한 배열은 무시)', () => {
  const data = sandbox.naverRankExtractNextData_(SAMPLE_HTML);
  const arr = sandbox.naverRankFindProductArray_(data);
  ok(Array.isArray(arr), '배열 아님');
  eq(arr.length, SAMPLE_LIST.length);
});

console.log('\n[2] 광고 판별');

test('isAd:true는 광고', () => ok(sandbox.naverRankIsAdItem_({ isAd: true })));
test('cardType이 AD면 광고', () => ok(sandbox.naverRankIsAdItem_({ cardType: 'AD' })));
test('adId가 있으면 광고', () => ok(sandbox.naverRankIsAdItem_({ adId: 'x' })));
test('badge에 "광고"가 있으면 광고', () => ok(sandbox.naverRankIsAdItem_({ badge: '광고' })));
test('아무 표시 없으면 광고 아님', () => ok(!sandbox.naverRankIsAdItem_({ productName: 'x' })));

console.log('\n[3] 순위 계산 (광고 제외)');

test('광고를 건너뛰고 남은 상품에만 순위를 매긴다', () => {
  fetchResponses = { 이어폰: makeResponse(200, SAMPLE_HTML) };
  const result = sandbox.naverRankFetchRanking_('이어폰');
  ok(result.ok, '실패: ' + result.note);
  eq(result.items.map((i) => i.productId), ['333', '555', '666']);
  eq(result.items.map((i) => i.rank), [1, 2, 3]);
  eq(result.items[0].name, '무선 이어폰 프로 정품');
  eq(result.items[0].mallName, '정품몰');
});

test('HTTP 418(차단)은 실패로 보고하고 이유를 남긴다', () => {
  fetchResponses = { 차단키워드: makeResponse(418, '<html>차단</html>') };
  const result = sandbox.naverRankFetchRanking_('차단키워드');
  ok(!result.ok);
  ok(result.note.indexOf('차단') !== -1, 'note: ' + result.note);
});

test('예상 못한 HTML 구조면 실패로 보고한다', () => {
  fetchResponses = { 이상한페이지: makeResponse(200, '<html><body>구조 변경됨</body></html>') };
  const result = sandbox.naverRankFetchRanking_('이상한페이지');
  ok(!result.ok);
});

console.log('\n[4] MID 매칭');

test('상품 MID가 일치하는 항목을 찾는다', () => {
  const items = [
    { rank: 1, productId: '111', name: 'A' },
    { rank: 2, productId: '222', name: 'B' },
  ];
  eq(sandbox.naverRankFindByMid_(items, 222, '').rank, 2);
});

test('상품 MID로 못 찾으면 가격비교 MID로 찾는다', () => {
  const items = [
    { rank: 1, productId: '999', name: 'A' },
    { rank: 2, productId: '888', name: 'B' },
  ];
  eq(sandbox.naverRankFindByMid_(items, '111', 888).rank, 2);
});

test('둘 다 일치하지 않으면 null', () => {
  const items = [{ rank: 1, productId: '999', name: 'A' }];
  eq(sandbox.naverRankFindByMid_(items, '111', '222'), null);
});

console.log('\n[5] naverRankUpdateAll (시트 통합 · 열 삽입)');

test('행마다 순위를 계산해 K열에 새 열을 삽입하고, 같은 키워드는 한 번만 요청한다', () => {
  fetchResponses = { 이어폰: makeResponse(200, SAMPLE_HTML) };

  // 3행: 헤더 (A~J + 기존 K열 이력 1개), 4행부터 상품 데이터 2건 (같은 키워드 공유)
  const grid = [
    ['사용법'],
    [],
    ['상품명', '리워드명', '갯수', '판매처명', '링크', '순위변동', '순위 조회 키워드', '상품 MID', '가격비교 MID', '월검색량', '기존이력'],
    ['', '', '', 'A몰', 'link1', '-', '이어폰', 333, '', 100, 5],
    ['', '', '', 'B몰', 'link2', '-', '이어폰', 666, '', 100, 3],
  ];
  const sheet = fakeSheet(grid);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: (name) => (name === sandbox.NAVER_RANK_CONFIG.SHEET_NAME ? sheet : null),
  });

  let fetchCount = 0;
  const originalFetch = sandbox.UrlFetchApp.fetch;
  sandbox.UrlFetchApp.fetch = (url) => { fetchCount++; return originalFetch(url); };

  sandbox.naverRankUpdateAll();
  sandbox.UrlFetchApp.fetch = originalFetch;

  eq(fetchCount, 1, '같은 키워드는 한 번만 요청해야 함');

  const finalGrid = sheet._grid();
  // K열(index 10)에 새 열이 삽입되어 기존 이력은 L열(index 11)로 밀려야 함
  ok(finalGrid[2][10] instanceof Date, '새 헤더 셀은 Date 객체여야 함');
  eq(finalGrid[2][11], '기존이력', '기존 이력 열이 오른쪽으로 밀려야 함');
  eq(finalGrid[3][10], 1, '상품MID 333 → 1위');
  eq(finalGrid[4][10], 3, '상품MID 666 → 3위');
});

test('키워드나 MID가 비어 있으면 건너뛴다(빈 문자열 기록)', () => {
  fetchResponses = {};
  const grid = [
    ['사용법'], [],
    ['상품명', '리워드명', '갯수', '판매처명', '링크', '순위변동', '순위 조회 키워드', '상품 MID', '가격비교 MID', '월검색량'],
    ['', '', '', 'A몰', 'link1', '-', '', '', '', 100],
  ];
  const sheet = fakeSheet(grid);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: () => sheet,
  });
  sandbox.naverRankUpdateAll();
  eq(sheet._grid()[3][10], '');
});

test('차단/오류 시 999를 기록한다', () => {
  fetchResponses = { 차단검색: makeResponse(418, '차단') };
  const grid = [
    ['사용법'], [],
    ['상품명', '리워드명', '갯수', '판매처명', '링크', '순위변동', '순위 조회 키워드', '상품 MID', '가격비교 MID', '월검색량'],
    ['', '', '', 'A몰', 'link1', '-', '차단검색', 123, '', 100],
  ];
  const sheet = fakeSheet(grid);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: () => sheet,
  });
  sandbox.naverRankUpdateAll();
  eq(sheet._grid()[3][10], sandbox.NAVER_RANK_CONFIG.NOT_FOUND_RANK);
});

test('매칭되는 MID가 없으면 999를 기록한다', () => {
  fetchResponses = { 이어폰3: makeResponse(200, SAMPLE_HTML) };
  const grid = [
    ['사용법'], [],
    ['상품명', '리워드명', '갯수', '판매처명', '링크', '순위변동', '순위 조회 키워드', '상품 MID', '가격비교 MID', '월검색량'],
    ['', '', '', 'A몰', 'link1', '-', '이어폰3', 99999, '', 100],
  ];
  const sheet = fakeSheet(grid);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: () => sheet,
  });
  sandbox.naverRankUpdateAll();
  eq(sheet._grid()[3][10], sandbox.NAVER_RANK_CONFIG.NOT_FOUND_RANK);
});

test('시트를 찾을 수 없으면 에러', () => {
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({ getSheetByName: () => null });
  let threw = false;
  try { sandbox.naverRankGetSheet_(); } catch (e) { threw = true; }
  ok(threw);
});

console.log('\n[6] 트리거 설치/삭제 (11시, 17시)');

test('naverRankInstall은 11시/17시 트리거를 정확히 2개 만든다', () => {
  sandbox.ScriptApp._reset();
  sandbox.naverRankInstall();
  const triggers = sandbox.naverRankListTriggers_();
  eq(triggers.length, 2);
  const hours = triggers.map((t) => t.getTriggerSourceId());
  ok(hours.indexOf('hour:11') !== -1, '11시 트리거 없음');
  ok(hours.indexOf('hour:17') !== -1, '17시 트리거 없음');
});

test('naverRankInstall을 다시 실행해도 트리거가 중복되지 않는다', () => {
  sandbox.naverRankInstall();
  eq(sandbox.naverRankListTriggers_().length, 2);
});

test('naverRankUninstall은 트리거를 모두 삭제한다', () => {
  sandbox.naverRankUninstall();
  eq(sandbox.naverRankListTriggers_().length, 0);
});

/* ── 결과 ─────────────────────────────────────────────────────────────────── */

console.log(`\n총 ${passed + failures.length}건 중 ${passed}건 통과, ${failures.length}건 실패`);
if (failures.length > 0) {
  console.log('\n실패 목록:');
  failures.forEach((f) => console.log(`  - ${f.name}: ${f.message}`));
  process.exit(1);
}
