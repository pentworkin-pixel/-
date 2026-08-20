/**
 * 「프로그램 동기화.gs」 테스트용 공통 하네스.
 *
 * 배포될 .gs 파일을 그대로 vm 컨텍스트에 올려 실행한다. 테스트는 사본이 아니라
 * 실제로 Apps Script에 붙여넣을 코드를 검증한다.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const GS_PATH = path.join(__dirname, '..', '프로그램 동기화.gs');

/* ── Apps Script 전역 스텁 ────────────────────────────────────────────────── */

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

/**
 * .gs를 새 컨텍스트에 로드한다. overrides로 SpreadsheetApp/ScriptApp 등을 갈아끼울 수 있다.
 * 테스트마다 독립된 컨텍스트를 쓰므로 설정을 바꿔도 서로 오염되지 않는다.
 */
function load(overrides) {
  const logs = [];
  const sandbox = Object.assign({
    Utilities: { formatDate },
    Logger: { log: (m) => logs.push(String(m)) },
    console: { log: (m) => logs.push(String(m)) },
    SpreadsheetApp: {},
    ScriptApp: {},
    Date, Math, JSON, String, Number, Object, Array, RegExp, isNaN, isFinite
  }, overrides || {});

  sandbox.__logs = logs;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(GS_PATH, 'utf8'), sandbox, { filename: GS_PATH });
  return sandbox;
}

/* ── 가짜 시트 ────────────────────────────────────────────────────────────── */

/**
 * 2차원 배열을 감싸는 가짜 Sheet. getRange(...).getValues()/setValues()를 지원하고
 * 쓰기 호출을 전부 기록한다.
 */
function fakeSheet(name, gid, grid) {
  const writes = [];
  const cells = grid.map((row) => row.slice());

  function widest() {
    return cells.reduce((max, row) => Math.max(max, row.length), 0);
  }

  return {
    writes,
    cells,
    getName: () => name,
    getSheetId: () => gid,
    getLastRow: () => cells.length,
    getLastColumn: widest,
    getProtections: () => [],
    getRange(row, col, numRows, numCols) {
      const nR = numRows === undefined ? 1 : numRows;
      const nC = numCols === undefined ? 1 : numCols;
      return {
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
          writes.push({ row, col, numRows: nR, numCols: nC, values });
          for (let r = 0; r < values.length; r++) {
            while (cells.length < row - 1 + r + 1) cells.push([]);
            const target = cells[row - 1 + r];
            for (let c = 0; c < values[r].length; c++) {
              while (target.length < col - 1 + c) target.push('');
              target[col - 1 + c] = values[r][c];
            }
          }
        }
      };
    }
  };
}

/** 가짜 Spreadsheet. */
function fakeSpreadsheet(id, name, sheets) {
  return {
    getId: () => id,
    getName: () => name,
    getSheets: () => sheets
  };
}

/** openById로 여러 스프레드시트를 돌려주는 가짜 SpreadsheetApp. */
function fakeSpreadsheetApp(byId) {
  return {
    ProtectionType: { SHEET: 'SHEET' },
    getActiveSpreadsheet: () => null,
    openById: (id) => {
      if (!byId[id]) throw new Error('없는 문서: ' + id);
      return byId[id];
    },
    flush: () => {}
  };
}

/** 트리거 생성/삭제를 기록하는 가짜 ScriptApp. */
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

function assert(cond, message) {
  if (cond) { passed++; return; }
  failures.push(message);
}

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

function assertThrows(fn, message) {
  try {
    fn();
  } catch (e) {
    passed++;
    return;
  }
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
  assert, assertEqual, assertDeep, assertThrows, report, GS_PATH
};
