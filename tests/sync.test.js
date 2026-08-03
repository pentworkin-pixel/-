/**
 * 동기화 엔진 시뮬레이션 테스트.
 *
 * Calendar REST API를 가짜 캘린더로 대체하고 adEndSyncRun() 을 실제로 돌려서
 * 신규/갱신/유지 판정, 중복 생성 금지(규칙 10), 삭제 금지(규칙 14),
 * 일정 속성(규칙 7·15·16), 호출 최소화(규칙 17)를 검증한다.
 *
 * 실행: node tests/sync.test.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GS_PATH = path.join(__dirname, '..', '광고 종료 캘린더 동기화.gs');
const GS_SOURCE = fs.readFileSync(GS_PATH, 'utf8');
const TODAY = '2026-08-03';

function formatDate(date, tz, pattern) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: false, timeZoneName: 'longOffset'
  }).formatToParts(date).reduce((a, p) => (a[p.type] = p.value, a), {});
  if (pattern === 'ZZ') {
    const m = /GMT([+-]\d{2}):?(\d{2})/.exec(parts.timeZoneName);
    return m ? m[1] + m[2] : '+0000';
  }
  return pattern
    .replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day)
    .replace('HH', parts.hour === '24' ? '00' : parts.hour)
    .replace('mm', parts.minute).replace('ss', parts.second);
}

/** 가짜 Google Calendar. UrlFetchApp 레벨에서 가로챈다. */
function makeFakeCalendar(seedEvents) {
  const store = new Map();          // id -> event resource
  const calls = [];                 // 호출 기록 (규칙 17 검증용)
  let nextId = 1;

  (seedEvents || []).forEach((e) => {
    const id = 'seed-' + (nextId++);
    store.set(id, Object.assign({ id }, e));
  });

  return {
    store, calls,
    deletedCount: 0,
    fetch(url, options) {
      const method = (options.method || 'get').toLowerCase();
      const [base, query] = url.split('?');
      const params = {};
      (query || '').split('&').filter(Boolean).forEach((kv) => {
        const [k, v] = kv.split('=');
        params[decodeURIComponent(k)] = decodeURIComponent(v);
      });
      calls.push({ method, path: base.replace('https://www.googleapis.com/calendar/v3', '') });

      const body = options.payload ? JSON.parse(options.payload) : null;
      const idMatch = /\/events\/([^/?]+)$/.exec(base);

      if (method === 'get') {
        const items = [...store.values()]
          .filter((e) => {
            const start = e.start.dateTime || e.start.date;
            return start >= params.timeMin.slice(0, 10) &&
                   start.slice(0, 10) <= params.timeMax.slice(0, 10);
          })
          .sort((a, b) => (a.start.dateTime || a.start.date)
                          .localeCompare(b.start.dateTime || b.start.date));
        return respond(200, { items });
      }

      if (method === 'post') {
        const id = 'new-' + (nextId++);
        store.set(id, Object.assign({ id }, body));
        return respond(200, store.get(id));
      }

      if (method === 'patch') {
        const id = decodeURIComponent(idMatch[1]);
        const existing = store.get(id);
        if (!existing) return respond(404, { error: 'not found' });
        // PATCH는 보낸 필드만 덮어쓴다 — 실제 API 동작과 동일.
        Object.assign(existing, body);
        return respond(200, existing);
      }

      if (method === 'delete') { this.deletedCount++; return respond(204, null); }
      return respond(405, { error: 'method not allowed' });
    }
  };

  function respond(code, obj) {
    return {
      getResponseCode: () => code,
      getContentText: () => (obj === null ? '' : JSON.stringify(obj))
    };
  }
}

function fakeSpreadsheet(sheetDefs) {
  return {
    getName: () => '김승환 현황',
    getSheets: () => sheetDefs.map((def) => ({
      getName: () => def.name,
      getDataRange: () => ({ getValues: () => def.values })
    }))
  };
}

function row(company, keyword, endDate) {
  const r = new Array(8).fill('');
  r[1] = company; r[2] = keyword; r[7] = endDate;
  return r;
}
const HEADER = row('업체명', '종료 광고 키워드', '종료일');

/** .gs를 새 샌드박스에 로드하고, 시트/캘린더/오늘 날짜를 주입한다. */
function loadScript(sheetDefs, calendar) {
  const logs = [];
  const sandbox = {
    Utilities: { formatDate },
    Logger: { log: (m) => logs.push(String(m)) },
    console: { log: () => {} },
    SpreadsheetApp: {
      openById: () => fakeSpreadsheet(sheetDefs),
      getActive: () => fakeSpreadsheet(sheetDefs)
    },
    CalendarApp: { getCalendarById: () => ({ getName: () => 'fake' }) },
    ScriptApp: { getOAuthToken: () => 'fake-token' },
    UrlFetchApp: { fetch: (u, o) => calendar.fetch(u, o) },
    Date, Math, JSON, String, Number, Object, isNaN, isFinite, RegExp, Array
  };
  vm.createContext(sandbox);
  vm.runInContext(GS_SOURCE, sandbox, { filename: GS_PATH });
  // "오늘"을 고정해 테스트를 재현 가능하게 만든다.
  vm.runInContext('adEndSyncTodayKey_ = function () { return ' + JSON.stringify(TODAY) + '; };', sandbox);
  sandbox.__logs = logs;
  return sandbox;
}

let passed = 0;
const failures = [];
function test(name, fn) {
  try { fn(); passed++; console.log('  ✅ ' + name); }
  catch (e) { failures.push(name); console.log('  ❌ ' + name + '\n       ' + e.message); }
}
function eq(actual, expected, label) {
  const a = JSON.stringify(actual), b = JSON.stringify(expected);
  if (a !== b) throw new Error((label ? label + ': ' : '') + '기대 ' + b + ' / 실제 ' + a);
}

/* ── 공통 시트: 8/10에 중복 키워드 2건 + 단독 1건, 8/12에 1건 ─────────────── */
const SHEETS = [
  { name: '08월 12일 종료', values: [HEADER, row('델타상사', '겨울준비', '2026-08-12')] },
  { name: '08월 10일 종료', values: [HEADER,
      row('알파상사', '여름세일', '2026-08-10'),
      row('베타상사', '여름세일', '2026-08-10'),
      row('감마상사', '가을신상', '2026-08-10')] },
  { name: '07월 30일 종료', values: [HEADER, row('과거', '읽으면안됨', '2026-07-30')] }
];

console.log('\n[1] 빈 캘린더 → 전부 신규 생성');

test('신규 2건, 갱신 0, 유지 0, 실패 0', () => {
  const cal = makeFakeCalendar([]);
  const s = loadScript(SHEETS, cal);
  const r = s.adEndSyncRun();
  eq([r.신규, r.갱신, r.유지, r.실패], [2, 0, 0, 0]);
});

test('생성된 일정 속성: 09:00~09:30 / 비공개 / 투명 / 팝업 알림', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  const ev = [...cal.store.values()].find((e) => e.start.dateTime.startsWith('2026-08-10'));
  eq(ev.start.dateTime, '2026-08-10T09:00:00+09:00', '시작');
  eq(ev.end.dateTime, '2026-08-10T09:30:00+09:00', '종료');
  eq(ev.start.timeZone, 'Asia/Seoul');
  eq(ev.visibility, 'private');
  eq(ev.transparency, 'transparent');
  eq(ev.reminders, { useDefault: false, overrides: [{ method: 'popup', minutes: 0 }] });
});

test('중복 키워드 날짜: 시트 3행 → 캘린더 3건 (제목·설명 모두)', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  const ev = [...cal.store.values()].find((e) => e.start.dateTime.startsWith('2026-08-10'));
  eq(ev.summary, '[광고 종료] 광고 3건');
  eq(ev.description,
     'Google Sheets: [김승환 현황]\n1. 여름세일 (알파상사)\n2. 여름세일 (베타상사)\n3. 가을신상');
});

test('과거 탭(07월 30일)의 행은 캘린더에 등록되지 않는다', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  const all = [...cal.store.values()].map((e) => e.description).join('\n');
  if (all.indexOf('읽으면안됨') !== -1) throw new Error('과거 탭 데이터가 포함됨');
  eq(cal.store.size, 2);
});

console.log('\n[2] 규칙 10·11 — 재실행해도 중복 생성/불필요한 변경 없음');

test('두 번째 실행: 신규 0, 갱신 0, 유지 2 (멱등성)', () => {
  const cal = makeFakeCalendar([]);
  const s = loadScript(SHEETS, cal);
  s.adEndSyncRun();
  const before = cal.store.size;
  const r2 = s.adEndSyncRun();
  eq([r2.신규, r2.갱신, r2.유지, r2.실패], [0, 0, 2, 0]);
  eq(cal.store.size, before, '일정 개수 변화 없음');
});

test('규칙 17 — 변경 없을 때 쓰기 호출 0회 (조회 1회만)', () => {
  const cal = makeFakeCalendar([]);
  const s = loadScript(SHEETS, cal);
  s.adEndSyncRun();
  cal.calls.length = 0;
  s.adEndSyncRun();
  const writes = cal.calls.filter((c) => c.method !== 'get');
  eq(writes.length, 0, '쓰기 호출');
  eq(cal.calls.filter((c) => c.method === 'get').length, 1, '조회 호출');
});

test('규칙 10 — 이미 [광고 종료] 일정이 있으면 새로 만들지 않는다', () => {
  const cal = makeFakeCalendar([{
    summary: '[광고 종료] 광고 3건',
    description: 'Google Sheets: [김승환 현황]\n1. 여름세일 (알파상사)\n2. 여름세일 (베타상사)\n3. 가을신상',
    start: { dateTime: '2026-08-10T09:00:00+09:00' },
    end: { dateTime: '2026-08-10T09:30:00+09:00' }
  }]);
  const r = loadScript(SHEETS, cal).adEndSyncRun();
  eq([r.신규, r.갱신, r.유지], [1, 0, 1]); // 8/12만 신규, 8/10은 유지
  const onAug10 = [...cal.store.values()]
    .filter((e) => (e.start.dateTime || '').startsWith('2026-08-10'));
  eq(onAug10.length, 1, '8월 10일 일정 개수');
});

test('무관한 일정([광고 종료] 아님)은 무시하고 별도로 생성한다', () => {
  const cal = makeFakeCalendar([{
    summary: '팀 회의',
    description: '',
    start: { dateTime: '2026-08-10T09:00:00+09:00' },
    end: { dateTime: '2026-08-10T10:00:00+09:00' }
  }]);
  const r = loadScript(SHEETS, cal).adEndSyncRun();
  eq([r.신규, r.갱신, r.유지], [2, 0, 0]);
  eq(cal.store.size, 3, '기존 회의 + 신규 2건');
});

console.log('\n[3] 규칙 12 — 시트가 바뀌면 제목·설명만 갱신');

test('키워드 추가 시 갱신 1건으로 잡힌다', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();

  const changed = JSON.parse(JSON.stringify(SHEETS));
  changed[1].values.push(row('입실론상사', '추가키워드', '2026-08-10'));

  const r = loadScript(changed, cal).adEndSyncRun();
  eq([r.신규, r.갱신, r.유지, r.실패], [0, 1, 1, 0]);
  const ev = [...cal.store.values()].find((e) => e.start.dateTime.startsWith('2026-08-10'));
  eq(ev.summary, '[광고 종료] 광고 4건');
  eq(ev.description,
     'Google Sheets: [김승환 현황]\n1. 여름세일 (알파상사)\n2. 여름세일 (베타상사)\n3. 가을신상\n4. 추가키워드');
});

test('갱신 시 시간·알림·공개범위·투명도는 그대로 유지된다', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();

  const changed = JSON.parse(JSON.stringify(SHEETS));
  changed[1].values.push(row('입실론상사', '추가키워드', '2026-08-10'));
  loadScript(changed, cal).adEndSyncRun();

  const ev = [...cal.store.values()].find((e) => e.start.dateTime.startsWith('2026-08-10'));
  eq(ev.start.dateTime, '2026-08-10T09:00:00+09:00');
  eq(ev.end.dateTime, '2026-08-10T09:30:00+09:00');
  eq(ev.visibility, 'private');
  eq(ev.transparency, 'transparent');
  eq(ev.reminders, { useDefault: false, overrides: [{ method: 'popup', minutes: 0 }] });
});

test('갱신 시 PATCH 1회만 호출한다 (규칙 17)', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  const changed = JSON.parse(JSON.stringify(SHEETS));
  changed[1].values.push(row('입실론상사', '추가키워드', '2026-08-10'));
  cal.calls.length = 0;
  loadScript(changed, cal).adEndSyncRun();
  eq(cal.calls.filter((c) => c.method === 'patch').length, 1, 'patch 호출');
  eq(cal.calls.filter((c) => c.method === 'post').length, 0, 'post 호출');
});

console.log('\n[4] 규칙 13·14 — 신규만 생성, 자동 삭제 없음');

test('시트에서 날짜가 사라져도 기존 일정을 삭제하지 않는다', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  eq(cal.store.size, 2);

  const shrunk = [SHEETS[1], SHEETS[2]]; // 08월 12일 시트를 제거
  const r = loadScript(shrunk, cal).adEndSyncRun();

  eq(cal.deletedCount, 0, 'DELETE 호출 수');
  eq(cal.store.size, 2, '일정은 그대로 남아 있어야 함');
  eq([r.신규, r.갱신, r.유지], [0, 0, 1]);
});

test('새 종료일이 추가되면 그 날짜만 신규 생성', () => {
  const cal = makeFakeCalendar([]);
  loadScript(SHEETS, cal).adEndSyncRun();
  const grown = JSON.parse(JSON.stringify(SHEETS));
  grown.unshift({ name: '08월 20일 종료', values: [HEADER, row('제타상사', '신규건', '2026-08-20')] });
  const r = loadScript(grown, cal).adEndSyncRun();
  eq([r.신규, r.갱신, r.유지], [1, 0, 2]);
  eq(cal.store.size, 3);
});

console.log('\n[5] 미리보기(dry-run)와 오류 처리');

test('dryRun은 캘린더를 변경하지 않는다', () => {
  const cal = makeFakeCalendar([]);
  const r = loadScript(SHEETS, cal).adEndSyncDryRun();
  eq([r.신규, r.갱신, r.유지], [2, 0, 0]);
  eq(cal.store.size, 0, '캘린더에 아무것도 생성되면 안 됨');
});

test('한 날짜가 실패해도 나머지는 계속 처리하고 실패로 집계한다', () => {
  const cal = makeFakeCalendar([]);
  const originalFetch = cal.fetch.bind(cal);
  let postCount = 0;
  cal.fetch = (url, options) => {
    if ((options.method || 'get').toLowerCase() === 'post' && ++postCount === 1) {
      return { getResponseCode: () => 403, getContentText: () => '{"error":"rateLimitExceeded"}' };
    }
    return originalFetch(url, options);
  };
  const r = loadScript(SHEETS, cal).adEndSyncRun();
  eq([r.신규, r.실패], [1, 1]);
  if (!/rateLimitExceeded/.test(r.오류.join(''))) throw new Error('오류 메시지 미기록');
});

console.log('\n[6] 중복 검증 도구');

test('adEndSyncVerifyDuplicates가 동기화 후 통과로 보고한다', () => {
  const cal = makeFakeCalendar([]);
  const s = loadScript(SHEETS, cal);
  s.adEndSyncRun();
  const report = s.adEndSyncVerifyDuplicates();
  if (!/✅ 2026-08-10 \| 시트 행 수 3 \| 캘린더 제목 3건 \| 캘린더 설명 항목 3개/.test(report)) {
    throw new Error('보고 내용 불일치:\n' + report);
  }
  if (!/통과 1건 \/ 불일치 0건/.test(report)) throw new Error('집계 불일치:\n' + report);
});

test('동기화 전에는 불일치로 보고한다', () => {
  const cal = makeFakeCalendar([]);
  const report = loadScript(SHEETS, cal).adEndSyncVerifyDuplicates();
  if (!/통과 0건 \/ 불일치 1건/.test(report)) throw new Error('보고 내용 불일치:\n' + report);
});

console.log('\n───────────────────────────────────────────');
console.log('통과 ' + passed + '건 / 실패 ' + failures.length + '건');
if (failures.length) { process.exit(1); }
console.log('모든 테스트 통과 ✅');
