/**
 * 「네이버 쇼핑 순위 조회.gs」의 순수 로직 테스트.
 *
 * 배포될 .gs 파일을 그대로 읽어 Apps Script 전역(UrlFetchApp/SpreadsheetApp 등)만
 * 스텁으로 채운 뒤 실행한다. 즉 사본이 아니라 실제 배포 코드를 검증한다.
 *
 * 주의: 여기서 쓰는 __NEXT_DATA__ 픽스처는 실제 네이버 응답을 캡처한 것이 아니라
 * (네트워크가 봇 차단으로 막혀 있어 실물 응답을 확보할 수 없었음), 이 파일이
 * 상정하는 "이름 필드 + id 필드를 가진 객체 배열" 모양을 흉내 낸 것이다.
 * 즉 이 테스트는 파싱/광고제외/순위계산/매칭 로직 자체의 정확성을 검증하는
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
const fetchCalls = [];
let fetchResponses = {}; // keyword -> response 객체 또는 그 객체를 만드는 함수

function makeResponse(code, text) {
  return { getResponseCode: () => code, getContentText: () => text };
}

const sandbox = {
  Utilities: {
    sleep: () => {},
    formatDate: () => '2026-08-03 12:00:00',
  },
  Logger: { log: (m) => logs.push(String(m)) },
  console: { log: () => {} },
  SpreadsheetApp: {},
  UrlFetchApp: {
    fetch: (url) => {
      fetchCalls.push(url);
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
  { productId: 'p1', productName: '무선 이어폰 프로', mallName: '스마트몰', isAd: true },
  { productId: 'p2', productName: '무선 이어폰 프로 케이스', mallName: '가짓A몰', cardType: 'AD' },
  { productId: 'p3', title: '무선 이어폰 프로 정품', mallName: '정품몰' },
  { productId: 'p4', productName: '무선 이어폰 프로 (리퍼)', mallName: '리퍼몰', adId: 'xyz123' },
  { productId: 'p5', productName: '완전 다른 상품', mallName: '다른몰' },
  { productId: 'p6', productName: '무선 이어폰 프로', mallName: '베스트몰' },
];

const SAMPLE_HTML = buildFixtureHtml(SAMPLE_LIST);

/* ── 가짜 스프레드시트 ────────────────────────────────────────────────────── */

function fakeSheet(rows) {
  let stored = rows.slice();
  return {
    getLastRow: () => stored.length,
    getRange: (startRow, startCol, numRows, numCols) => ({
      getValues: () => stored.slice(startRow - 1, startRow - 1 + numRows).map((r) => r.slice()),
      setValues: (values) => {
        for (let i = 0; i < values.length; i++) stored[startRow - 1 + i] = values[i];
      },
    }),
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
  eq(result.items.map((i) => i.productId), ['p3', 'p5', 'p6']);
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

console.log('\n[4] 상품 매칭');

test('상품명 부분 일치 + 대소문자 무시', () => {
  const items = [
    { rank: 1, name: 'Wireless Earbuds Pro', mallName: 'A몰' },
    { rank: 2, name: '무선 이어폰', mallName: 'B몰' },
  ];
  const found = sandbox.naverRankFindMatch_(items, 'wireless earbuds');
  eq(found.rank, 1);
});

test('판매처까지 지정하면 그 판매처인 것만 매칭', () => {
  const items = [
    { rank: 1, name: '무선 이어폰 프로', mallName: 'A몰' },
    { rank: 2, name: '무선 이어폰 프로', mallName: '베스트몰' },
  ];
  eq(sandbox.naverRankFindMatch_(items, '이어폰', '베스트').rank, 2);
});

test('매칭되는 상품이 없으면 null', () => {
  const items = [{ rank: 1, name: '전혀 다른 상품', mallName: 'A몰' }];
  eq(sandbox.naverRankFindMatch_(items, '이어폰'), null);
});

console.log('\n[5] naverRankCheckOne_ (조회 + 매칭 통합)');

test('매칭 성공 시 순위/상품명/판매처를 반환', () => {
  fetchResponses = { 이어폰2: makeResponse(200, SAMPLE_HTML) };
  const r = sandbox.naverRankCheckOne_('이어폰2', '이어폰', '베스트');
  eq(r.rank, 3);
  eq(r.mallName, '베스트몰');
  eq(r.note, '');
});

test('매칭 실패 시 미노출 사유를 남기고 순위는 빈 값', () => {
  fetchResponses = { 이어폰3: makeResponse(200, SAMPLE_HTML) };
  const r = sandbox.naverRankCheckOne_('이어폰3', '존재하지않는상품명');
  eq(r.rank, '');
  ok(r.note.indexOf('미노출') !== -1, 'note: ' + r.note);
});

test('차단되면 그 사유를 그대로 전달', () => {
  fetchResponses = { 이어폰4: makeResponse(418, '차단') };
  const r = sandbox.naverRankCheckOne_('이어폰4', '이어폰');
  eq(r.rank, '');
  ok(r.note.indexOf('차단') !== -1, 'note: ' + r.note);
});

console.log('\n[6] naverRankCheckAll (시트 통합)');

test('행마다 순위를 계산해 결과 열에 기록하고, 빈 키워드/매칭 행은 건너뛴다', () => {
  fetchResponses = {
    이어폰5: makeResponse(200, SAMPLE_HTML),
    차단검색: makeResponse(418, '차단'),
  };

  const header = ['키워드', '상품명매칭', '판매처매칭', '순위', '상품명', '판매처', '조회시각', '비고'];
  const rows = [
    header,
    ['이어폰5', '이어폰', '베스트', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''], // 빈 행 → 건너뜀
    ['차단검색', '아무거나', '', '', '', '', '', ''],
  ];

  const sheet = fakeSheet(rows);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: (name) => (name === sandbox.NAVER_RANK_CONFIG.SHEET_NAME ? sheet : null),
  });

  sandbox.naverRankCheckAll();

  const updated = sheet.getRange(2, 1, 3, 8).getValues();
  eq(updated[0][3], 3, '1행 순위'); // 이어폰5 → 3위
  eq(updated[0][4], '무선 이어폰 프로', '1행 상품명');
  eq(updated[0][5], '베스트몰', '1행 판매처');

  eq(updated[1], ['', '', '', '', '', '', '', ''], '빈 행은 그대로');

  eq(updated[2][3], '', '3행 순위(차단)');
  ok(String(updated[2][7]).indexOf('차단') !== -1, '3행 비고: ' + updated[2][7]);
});

test('시트를 찾을 수 없으면 에러', () => {
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({ getSheetByName: () => null });
  let threw = false;
  try { sandbox.naverRankGetSheet_(); } catch (e) { threw = true; }
  ok(threw);
});

/* ── 결과 ─────────────────────────────────────────────────────────────────── */

console.log(`\n총 ${passed + failures.length}건 중 ${passed}건 통과, ${failures.length}건 실패`);
if (failures.length > 0) {
  console.log('\n실패 목록:');
  failures.forEach((f) => console.log(`  - ${f.name}: ${f.message}`));
  process.exit(1);
}
