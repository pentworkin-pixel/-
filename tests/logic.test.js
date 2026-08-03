/**
 * 「광고 종료 캘린더 동기화.gs」의 순수 로직 테스트.
 *
 * 배포될 .gs 파일을 그대로 읽어 Apps Script 전역(Utilities/Logger 등)만 스텁으로
 * 채운 뒤 실행한다. 즉 사본이 아니라 실제 배포 코드를 검증한다.
 *
 * 실행: node tests/logic.test.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GS_PATH = path.join(__dirname, '..', '광고 종료 캘린더 동기화.gs');

/* ── Apps Script 전역 스텁 ────────────────────────────────────────────────── */

function formatDate(date, tz, pattern) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: false, timeZoneName: 'longOffset'
  }).formatToParts(date).reduce((a, p) => (a[p.type] = p.value, a), {});

  if (pattern === 'ZZ') {
    // 'GMT+09:00' → '+0900'
    const m = /GMT([+-]\d{2}):?(\d{2})/.exec(parts.timeZoneName);
    return m ? m[1] + m[2] : '+0000';
  }
  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('HH', parts.hour === '24' ? '00' : parts.hour)
    .replace('mm', parts.minute)
    .replace('ss', parts.second);
}

const logs = [];
const sandbox = {
  Utilities: { formatDate },
  Logger: { log: (m) => logs.push(String(m)) },
  console: { log: () => {} },
  SpreadsheetApp: {},
  CalendarApp: {},
  ScriptApp: {},
  UrlFetchApp: {},
  Date, Math, JSON, String, Number, Object, isNaN, isFinite, RegExp, Array
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(GS_PATH, 'utf8'), sandbox, { filename: GS_PATH });

/* ── 가짜 스프레드시트 ────────────────────────────────────────────────────── */

function fakeSpreadsheet(sheetDefs) {
  return {
    getName: () => '김승환 현황',
    getSheets: () => sheetDefs.map((def) => ({
      getName: () => def.name,
      getDataRange: () => ({ getValues: () => def.values })
    }))
  };
}

/** B/C/H만 채운 행을 만든다 (A~H = 8칸). */
function row(company, keyword, endDate) {
  const r = new Array(8).fill('');
  r[1] = company;   // B
  r[2] = keyword;   // C
  r[7] = endDate;   // H
  return r;
}

const HEADER = row('업체명', '종료 광고 키워드', '종료일');

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

const TODAY = '2026-08-03'; // 월요일

/* ── 테스트 ───────────────────────────────────────────────────────────────── */

console.log('\n[1] 시트 이름 파싱 및 연도 추론');

test('"08월 05일 종료" → 올해 날짜', () => {
  eq(sandbox.adEndSyncParseSheetName_('08월 05일 종료', TODAY), '2026-08-05');
});

test('"8월 5일 종료" (한 자리 월/일)도 인식', () => {
  eq(sandbox.adEndSyncParseSheetName_('8월 5일 종료', TODAY), '2026-08-05');
});

test('"01월 10일 종료" → 6개월 이상 과거이므로 내년으로 추론', () => {
  eq(sandbox.adEndSyncParseSheetName_('01월 10일 종료', TODAY), '2027-01-10');
});

test('형식이 다른 탭은 null (중단 조건 아님)', () => {
  eq(sandbox.adEndSyncParseSheetName_('전체 현황', TODAY), null);
  eq(sandbox.adEndSyncParseSheetName_('08월 05일 시작', TODAY), null);
});

console.log('\n[2] H열 종료일 정규화');

test('Date 객체', () => {
  eq(sandbox.adEndSyncNormalizeDate_(new Date(Date.UTC(2026, 7, 5, 3)), TODAY), '2026-08-05');
});

test('문자열 "2026-08-05"', () => {
  eq(sandbox.adEndSyncNormalizeDate_('2026-08-05', TODAY), '2026-08-05');
});

test('문자열 "8월 5일"', () => {
  eq(sandbox.adEndSyncNormalizeDate_('8월 5일', TODAY), '2026-08-05');
});

test('스프레드시트 날짜 시리얼', () => {
  const serial = Math.round((Date.UTC(2026, 7, 5) - Date.UTC(1899, 11, 30)) / 86400000);
  eq(sandbox.adEndSyncNormalizeDate_(serial, TODAY), '2026-08-05');
});

test('빈 값 → null', () => {
  eq(sandbox.adEndSyncNormalizeDate_('', TODAY), null);
  eq(sandbox.adEndSyncNormalizeDate_(null, TODAY), null);
});

console.log('\n[3] 규칙 2·3 — 오늘 이후만, 과거 탭 도달 시 중단');

test('과거 탭 이후의 더 오래된 탭은 읽지 않는다', () => {
  const ss = fakeSpreadsheet([
    { name: '08월 10일 종료', values: [HEADER, row('가', '미래키워드', '2026-08-10')] },
    { name: '08월 01일 종료', values: [HEADER, row('나', '과거키워드', '2026-08-01')] },
    { name: '07월 20일 종료', values: [HEADER, row('다', '더과거키워드', '2026-07-20')] }
  ]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.map((p) => p.dateKey), ['2026-08-10']);
});

test('형식이 다른 탭은 건너뛰되 중단하지 않는다', () => {
  const ss = fakeSpreadsheet([
    { name: '대시보드', values: [HEADER, row('x', 'y', 'z')] },
    { name: '08월 10일 종료', values: [HEADER, row('가', '키워드A', '2026-08-10')] }
  ]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.map((p) => p.dateKey), ['2026-08-10']);
});

test('오늘 종료분은 포함된다', () => {
  const ss = fakeSpreadsheet([
    { name: '08월 03일 종료', values: [HEADER, row('가', '오늘키워드', '2026-08-03')] }
  ]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.map((p) => p.dateKey), ['2026-08-03']);
});

console.log('\n[4] 규칙 4·5 — H열 기준 수집, 빈 키워드만 제외, 행 순서 유지');

test('H열이 시트 이름과 다르면 H열을 따른다', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 10일 종료',
    values: [HEADER,
      row('가', '키워드1', '2026-08-10'),
      row('나', '키워드2', '2026-08-12')] // H가 다름 → 별도 날짜
  }]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.map((p) => p.dateKey), ['2026-08-10', '2026-08-12']);
  eq(plan[0].labels, ['키워드1']);
  eq(plan[1].labels, ['키워드2']);
});

test('빈 키워드 행만 제외하고 나머지 순서는 그대로', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 10일 종료',
    values: [HEADER,
      row('A', '세번째로둘것아님-첫번째', '2026-08-10'),
      row('B', '', '2026-08-10'),          // 빈 키워드 → 제외
      row('C', '   ', '2026-08-10'),       // 공백뿐 → 제외
      row('D', '두번째', '2026-08-10'),
      row('E', '세번째', '2026-08-10')]
  }]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan[0].labels, ['세번째로둘것아님-첫번째', '두번째', '세번째']);
});

test('H열이 비어 있으면 시트 이름 날짜로 보완', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 10일 종료',
    values: [HEADER, row('가', '키워드X', '')]
  }]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.map((p) => p.dateKey), ['2026-08-10']);
});

console.log('\n[5] 규칙 6 — 중복 키워드는 절대 제거하지 않는다 (핵심)');

test('같은 키워드 2행 → 2건 유지 + B열 업체명 괄호 표기', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 10일 종료',
    values: [HEADER,
      row('알파상사', '여름세일', '2026-08-10'),
      row('베타상사', '여름세일', '2026-08-10'),
      row('감마상사', '가을신상', '2026-08-10')]
  }]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  const day = plan[0];

  eq(day.items.length, 3, '시트 행 수');
  eq(day.labels, ['여름세일 (알파상사)', '여름세일 (베타상사)', '가을신상']);
  eq(day.title, '[광고 종료] 광고 3건');
  eq(day.description,
     'Google Sheets: [김승환 현황]\n1. 여름세일 (알파상사)\n2. 여름세일 (베타상사)\n3. 가을신상');
});

test('같은 키워드 3행 → 3건 모두 남는다', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 11일 종료',
    values: [HEADER,
      row('A사', '동일키워드', '2026-08-11'),
      row('B사', '동일키워드', '2026-08-11'),
      row('C사', '동일키워드', '2026-08-11')]
  }]);
  const day = sandbox.adEndSyncBuildPlan_(ss, TODAY)[0];
  eq(day.items.length, 3);
  eq(day.title, '[광고 종료] 광고 3건');
  eq(day.labels, ['동일키워드 (A사)', '동일키워드 (B사)', '동일키워드 (C사)']);
});

test('중복이 아니면 업체명을 붙이지 않는다', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 12일 종료',
    values: [HEADER, row('단독상사', '유일키워드', '2026-08-12')]
  }]);
  eq(sandbox.adEndSyncBuildPlan_(ss, TODAY)[0].labels, ['유일키워드']);
});

test('중복이지만 B열이 비어 있으면 키워드만 (괄호 없음)', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 13일 종료',
    values: [HEADER,
      row('', '무명키워드', '2026-08-13'),
      row('', '무명키워드', '2026-08-13')]
  }]);
  const day = sandbox.adEndSyncBuildPlan_(ss, TODAY)[0];
  eq(day.items.length, 2, '중복 제거되면 안 됨');
  eq(day.labels, ['무명키워드', '무명키워드']);
});

test('B열이 날짜 셀이어도(요구사항 열 정의 충돌 대비) 깨지지 않는다', () => {
  const ss = fakeSpreadsheet([{
    name: '08월 14일 종료',
    values: [HEADER,
      row(new Date(Date.UTC(2026, 7, 14)), '키워드Z', '2026-08-14'),
      row(new Date(Date.UTC(2026, 7, 14)), '키워드Z', '2026-08-14')]
  }]);
  const day = sandbox.adEndSyncBuildPlan_(ss, TODAY)[0];
  eq(day.items.length, 2);
  eq(day.labels, ['키워드Z', '키워드Z']);
});

test('여러 시트에 흩어진 같은 종료일도 한 일정으로 합쳐지고 순서를 지킨다', () => {
  const ss = fakeSpreadsheet([
    { name: '08월 20일 종료', values: [HEADER, row('가', '키워드1', '2026-08-15')] },
    { name: '08월 15일 종료', values: [HEADER, row('나', '키워드2', '2026-08-15')] }
  ]);
  const plan = sandbox.adEndSyncBuildPlan_(ss, TODAY);
  eq(plan.length, 1);
  eq(plan[0].labels, ['키워드1', '키워드2']);
  eq(plan[0].title, '[광고 종료] 광고 2건');
});

console.log('\n[6] 규칙 8·9 — 제목/설명 형식');

test('제목 형식', () => {
  eq(sandbox.adEndSyncBuildTitle_(1), '[광고 종료] 광고 1건');
  eq(sandbox.adEndSyncBuildTitle_(12), '[광고 종료] 광고 12건');
});

test('설명 형식', () => {
  eq(sandbox.adEndSyncBuildDescription_(['첫 번째 키워드', '두 번째 키워드', '세 번째 키워드']),
     'Google Sheets: [김승환 현황]\n1. 첫 번째 키워드\n2. 두 번째 키워드\n3. 세 번째 키워드');
});

console.log('\n[7] 규칙 11 — 내용이 같으면 무작업');

test('제목·설명이 같으면 동일로 판정', () => {
  const day = { title: '[광고 종료] 광고 2건', description: 'Google Sheets: [김승환 현황]\n1. A\n2. B' };
  eq(sandbox.adEndSyncSameContent_({ summary: day.title, description: day.description }, day), true);
});

test('캘린더가 \\r\\n으로 돌려줘도 동일로 판정 (불필요한 갱신 방지)', () => {
  const day = { title: '[광고 종료] 광고 2건', description: 'Google Sheets: [김승환 현황]\n1. A\n2. B' };
  const ev = { summary: day.title, description: 'Google Sheets: [김승환 현황]\r\n1. A\r\n2. B\r\n' };
  eq(sandbox.adEndSyncSameContent_(ev, day), true);
});

test('키워드가 하나 늘면 다름으로 판정 → 갱신 대상', () => {
  const day = { title: '[광고 종료] 광고 3건', description: 'Google Sheets: [김승환 현황]\n1. A\n2. B\n3. C' };
  const ev = { summary: '[광고 종료] 광고 2건', description: 'Google Sheets: [김승환 현황]\n1. A\n2. B' };
  eq(sandbox.adEndSyncSameContent_(ev, day), false);
});

console.log('\n[8] 일정 시각 및 타임존');

test('RFC3339 시작/종료 시각이 09:00~09:30 (+09:00)', () => {
  eq(sandbox.adEndSyncRfc3339_('2026-08-10', 9, 0), '2026-08-10T09:00:00+09:00');
  eq(sandbox.adEndSyncRfc3339_('2026-08-10', 9, 30), '2026-08-10T09:30:00+09:00');
});

console.log('\n[9] 검증 도구 카운터');

test('설명에서 항목 수를 센다', () => {
  eq(sandbox.adEndSyncCountDescriptionItems_(
    'Google Sheets: [김승환 현황]\n1. 여름세일 (알파상사)\n2. 여름세일 (베타상사)\n3. 가을신상'), 3);
});

test('제목에서 건수를 읽는다', () => {
  eq(sandbox.adEndSyncCountTitleItems_('[광고 종료] 광고 7건'), 7);
  eq(sandbox.adEndSyncCountTitleItems_('회의'), null);
});

/* ── 결과 ─────────────────────────────────────────────────────────────────── */

console.log('\n───────────────────────────────────────────');
console.log('통과 ' + passed + '건 / 실패 ' + failures.length + '건');
if (failures.length) {
  failures.forEach((f) => console.log('  ❌ ' + f.name + ' — ' + f.message));
  process.exit(1);
}
console.log('모든 테스트 통과 ✅');
