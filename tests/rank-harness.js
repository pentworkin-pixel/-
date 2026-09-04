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

module.exports = { load, assertEqual, assertDeep, assertClose, assertThrows, report, GS_PATH };
