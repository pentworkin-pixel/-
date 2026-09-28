/**
 * 「순위 분석.gs」의 자동 실행(트리거) 설치 테스트.
 *
 * rankAnalysisInstall()이 매일 11시 트리거를 정확히 1개 만들고, 재실행해도
 * 중복되지 않으며, 다른 스크립트(광고 종료 캘린더 동기화 등)의 트리거는
 * 건드리지 않는지 검증한다.
 *
 * 실행: node tests/rank-trigger.test.js
 */
'use strict';

const {
  load, fakeSheet, fakeSpreadsheet, fakeSpreadsheetApp, fakeScriptApp,
  assertEqual, assertDeep, report
} = require('./rank-harness');

const A_ID = '1Sfru4Lfl7cVEXjyZuaqq1qye5UNDghctsSiL6ISeuh8';
const B_ID = '1bLNh-zrYHHKWgH78ihbgnhpd11ItUhFgw2eT0MromFU';
const B_GID = 1003701754;

const B_HEADER = ['메모1', 'b', '판매처명', '상품링크', '프로그램',
                  '전일 비교', '순위 조회 키워드', '상품 MID', '26-08-20', '26-08-19'];
const B_ROWS = [
  ['메모A', '', '', '', '리뷰 3개', '', '마스크', 111, 3, 10]
];

function setup() {
  const bSheet = fakeSheet('펀트미디어 순위추적2', B_GID, [B_HEADER].concat(B_ROWS));
  const aSs = fakeSpreadsheet(A_ID, '김승환 현황', []);
  const bSs = fakeSpreadsheet(B_ID, '펀트미디어 순위추적', [bSheet]);
  const scriptApp = fakeScriptApp();

  const gs = load({
    SpreadsheetApp: fakeSpreadsheetApp({ [A_ID]: aSs, [B_ID]: bSs }),
    ScriptApp: scriptApp
  });
  gs.RANK_ANALYSIS_CONFIG.HEADER_ROW = 1;   // 이 픽스처는 헤더가 1행

  return { gs, aSs, bSs, bSheet, scriptApp };
}

/* ── 설치 ─────────────────────────────────────────────────────────────────── */

{
  const { gs, aSs, scriptApp } = setup();
  const out = gs.rankAnalysisInstall();

  assertEqual(scriptApp.triggers.length, 1, '트리거가 정확히 1개 생성된다');
  assertEqual(scriptApp.triggers[0].fn, 'updateProgramRankingAnalysis',
              '트리거가 부르는 함수는 요구사항이 지정한 이름 그대로다');
  assertEqual(scriptApp.triggers[0].hour, 11, '매일 오전 11시');
  assertEqual(scriptApp.triggers[0].days, 1, '매일');
  assertEqual(scriptApp.triggers[0].tz, 'Asia/Seoul', '서울 시간 기준');
  assertEqual(out.indexOf('설치 완료') >= 0, true, '완료 리포트를 남긴다');

  // 설치 직후 1회 실행까지 끝난다
  assertDeep(aSs.inserted, ['프로그램별 순위 분석', '프로그램별 순위 상세'],
             '설치 직후 1회 실행으로 분석 시트가 바로 만들어진다');
  const analysis = aSs.getSheetByName('프로그램별 순위 분석');
  const flat = analysis.cells.map((r) => r.join('|')).join('\n');
  assertEqual(flat.indexOf('2026-08-20') >= 0, true, '설치 직후 실행 결과가 반영된다');
}

{
  // 이미 트리거가 여러 개 있어도 설치 후에는 항상 1개여야 한다.
  const { gs, scriptApp } = setup();
  scriptApp.newTrigger('updateProgramRankingAnalysis').timeBased().atHour(3).everyDays(1).create();
  scriptApp.newTrigger('updateProgramRankingAnalysis').timeBased().atHour(4).everyDays(1).create();
  assertEqual(scriptApp.triggers.length, 2, '설치 전 중복 트리거 2개');

  gs.rankAnalysisInstall();
  assertEqual(scriptApp.triggers.length, 1, '중복 트리거를 정리하고 1개만 남긴다');
  assertEqual(scriptApp.triggers[0].hour, 11, '남은 트리거는 11시짜리');
}

{
  // 다른 스크립트(광고 종료 캘린더 동기화 · 프로그램 동기화)의 트리거는 건드리지 않는다.
  const { gs, scriptApp } = setup();
  scriptApp.newTrigger('adEndSyncRun').timeBased().atHour(8).everyDays(1).create();
  scriptApp.newTrigger('progSyncRun').timeBased().atHour(10).everyDays(1).create();

  gs.rankAnalysisInstall();

  const handlers = scriptApp.triggers.map((t) => t.fn).sort();
  assertDeep(handlers, ['adEndSyncRun', 'progSyncRun', 'updateProgramRankingAnalysis'],
             '기존 다른 스크립트의 트리거는 그대로 두고 자기 트리거만 추가한다');
}

/* ── 중단 ─────────────────────────────────────────────────────────────────── */

{
  const { gs, scriptApp } = setup();
  scriptApp.newTrigger('adEndSyncRun').timeBased().atHour(8).everyDays(1).create();
  gs.rankAnalysisInstall();
  gs.rankAnalysisUninstall();

  const handlers = scriptApp.triggers.map((t) => t.fn);
  assertDeep(handlers, ['adEndSyncRun'], 'rankAnalysisUninstall 은 자기 트리거만 지운다');
}

{
  // 중단해도 이미 만들어진 분석 시트 내용은 남는다.
  const { gs, aSs } = setup();
  gs.rankAnalysisInstall();
  const before = aSs.getSheetByName('프로그램별 순위 분석').cells.length;
  gs.rankAnalysisUninstall();
  const after = aSs.getSheetByName('프로그램별 순위 분석').cells.length;
  assertEqual(after, before, '트리거만 지우고 분석 시트 내용은 건드리지 않는다');
}

/* ── 확인 ─────────────────────────────────────────────────────────────────── */

{
  const { gs } = setup();
  const before = gs.rankAnalysisVerifyTriggers();
  assertEqual(before.indexOf('0개') >= 0, true, '설치 전에는 0개로 보고한다');

  gs.rankAnalysisInstall();
  const after = gs.rankAnalysisVerifyTriggers();
  assertEqual(after.indexOf('1개') >= 0, true, '설치 후에는 1개로 보고한다');
  assertEqual(after.indexOf('⚠️') < 0, true, '1개일 때는 경고를 붙이지 않는다');
}

/* ── 설치 실패 시 트리거를 만들지 않는다 ──────────────────────────────────── */

{
  const { gs, scriptApp } = setup();
  gs.RANK_ANALYSIS_CONFIG.TARGET_SPREADSHEET_ID = '없는아이디';

  let threw = false;
  try { gs.rankAnalysisInstall(); } catch (e) { threw = true; }
  assertEqual(threw, true, 'A스프레드시트에 접근할 수 없으면 설치가 실패한다');
  assertEqual(scriptApp.triggers.length, 0,
              '접근 확인에 실패하면 트리거를 아예 만들지 않는다 (매일 실패할 트리거를 남기지 않음)');
}

report('순위 분석 · 자동 실행 설치');
