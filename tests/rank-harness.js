/**
 * 「순위 분석.gs」 테스트용 공통 하네스.
 *
 * 배포될 .gs 파일을 그대로 vm 컨텍스트에 올려 실행한다.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GS_PATH = path.join(__dirname, '..', '순위 분석.gs');

function formatDate(date, tz, pattern) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  }).formatToParts(date).reduce((a, p) => (a[p.type] = p.value, a), {});

  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('HH', parts.hour === '24' ? '00' : parts.hour)
    .replace('mm', parts.minute)
    .replace('ss', parts.second);
}

function load(overrides) {
  const logs = [];
  const sandbox = Object.assign({
    Utilities: { formatDate },
    Logger: { log: (m) => logs.push(String(m)) },
    console: { log: (m) => logs.push(String(m)) },
    SpreadsheetApp: {
      getUi: () => { throw new Error('no ui'); },
      getActiveSpreadsheet: () => null
    },
    Charts: { ChartType: { BAR: 'BAR', COLUMN: 'COLUMN' } },
    Date, Math, JSON, String, Number, Object, Array, RegExp, isNaN, isFinite
  }, overrides || {});

  sandbox.__logs = logs;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(GS_PATH, 'utf8'), sandbox, { filename: GS_PATH });
  return sandbox;
}

/* ── 가짜 시트 · 스프레드시트 (체이닝 메서드는 Proxy 로 흘려보낸다) ──────── */

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
  const rule = () => chainable({ build: () => ({}) });
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

/** 트리거 생성/삭제를 기록하는 가짜 ScriptApp. (프로그램 동기화 테스트의 것과 동일한 모양) */
function fakeScriptApp() {
  const triggers = [];
  let seq = 0;

  function builder(fn) {
    const spec = { fn: fn };
    const api = {
      timeBased: () => api,
      atHour: (h) => { spec.hour = h; return api; },
      nearMinute: (m) => { spec.minute = m; return api; },
      everyDays: (d) => { spec.days = d; return api; },
      inTimezone: (tz) => { spec.tz = tz; return api; },
      create: () => {
        seq++;
        const id = 'trigger-' + seq;
        triggers.push(Object.assign({
          getUniqueId: () => id,
          getHandlerFunction: () => spec.fn
        }, spec));
        return triggers[triggers.length - 1];
      }
    };
    return api;
  }

  return {
    triggers,
    newTrigger: builder,
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => {
      const i = triggers.indexOf(t);
      if (i >= 0) triggers.splice(i, 1);
    }
  };
}

/* ── 최소 어서션 ──────────────────────────────────────────────────────────── */

let passed = 0;
const failures = [];

function assertEqual(actual, expected, message) {
  if (actual === expected) { passed++; return; }
  failures.push(message + '\n      기대: ' + JSON.stringify(expected) +
                '\n      실제: ' + JSON.stringify(actual));
}

function assertDeep(actual, expected, message) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) { passed++; return; }
  failures.push(message + '\n      기대: ' + e + '\n      실제: ' + a);
}

function assertClose(actual, expected, message, eps) {
  const tol = eps === undefined ? 1e-9 : eps;
  if (typeof actual === 'number' && Math.abs(actual - expected) <= tol) { passed++; return; }
  failures.push(message + '\n      기대: ' + expected + '\n      실제: ' + JSON.stringify(actual));
}

function assertThrows(fn, message) {
  try { fn(); } catch (e) { passed++; return; }
  failures.push(message + '\n      기대: 예외 발생\n      실제: 정상 종료');
}

function report(title) {
  if (failures.length === 0) {
    console.log('✅ ' + title + ': ' + passed + '건 통과');
    return;
  }
  console.log('❌ ' + title + ': ' + passed + '건 통과, ' + failures.length + '건 실패');
  failures.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  process.exitCode = 1;
}

module.exports = {
  load, fakeSheet, fakeSpreadsheet, fakeSpreadsheetApp, fakeScriptApp,
  assertEqual, assertDeep, assertClose, assertThrows, report, GS_PATH
};
