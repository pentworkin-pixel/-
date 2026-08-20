/**
 * 「프로그램 동기화.gs」의 순수 로직 테스트.
 *
 * 배포될 .gs 파일을 그대로 읽어 Apps Script 전역만 스텁으로 채운 뒤 실행한다.
 * 즉 사본이 아니라 실제 배포 코드를 검증한다.
 *
 * 실행: node tests/program-logic.test.js
 */
'use strict';

const { load, assert, assertEqual, assertDeep, report } = require('./program-harness');

const gs = load();
const cfg = gs.PROG_SYNC_CONFIG;

/* ── 열 번호가 요구사항과 맞는지 (A=1 … E=5, K=11, Q=17) ──────────────────── */

assertEqual(cfg.SOURCE_COL_PROGRAM, 4, 'A시트 프로그램명 = D열');
assertEqual(cfg.SOURCE_COL_COUNT, 5, 'A시트 갯수 = E열');
assertEqual(cfg.SOURCE_COL_KEY, 17, 'A시트 상품키 = Q열');
assertEqual(cfg.TARGET_COL_KEY, 11, 'B시트 상품키 = K열');
assertEqual(cfg.TARGET_COL_OUTPUT, 5, 'B시트 출력 = E열');
assertEqual(cfg.TRIGGER_HOUR, 10, '트리거는 오전 10시');
assertEqual(cfg.TIMEZONE, 'Asia/Seoul', '기준 시간대는 서울');
assertEqual(cfg.CLEAR_WHEN_NO_MATCH, false, '매칭 없으면 기본적으로 건드리지 않음');

/* ── 상품키 정규화 ────────────────────────────────────────────────────────── */

const key = gs.progSyncNormalizeKey_;

assertEqual(key(90817968961), '90817968961', '큰 정수가 지수 표기로 새지 않는다');
assertEqual(key('90817968961'), '90817968961', '문자열 숫자 그대로');
assertEqual(key(' 123 '), '123', '앞뒤 공백 제거');
assertEqual(key('1,234'), '1234', '천단위 콤마 제거');
assertEqual(key("'123"), '123', '텍스트 강제용 앞따옴표 제거');
assertEqual(key('123.0'), '123', '"123.0" → "123"');
assertEqual(key(123.0), '123', '정수 float → "123"');
assertEqual(key(''), '', '빈 값');
assertEqual(key(null), '', 'null');
assertEqual(key(undefined), '', 'undefined');
assertEqual(key(123) === key('123'), true, '숫자 셀과 텍스트 셀이 같은 키가 된다');

/* ── 갯수 정규화 ──────────────────────────────────────────────────────────── */

const cnt = gs.progSyncParseCount_;

assertDeep(cnt(3), { countText: '3', countNum: 3 }, '숫자 3');
assertDeep(cnt('3'), { countText: '3', countNum: 3 }, '문자열 "3"');
assertDeep(cnt('3개'), { countText: '3', countNum: 3 }, '"3개" → 3 (개가 중복되지 않도록)');
assertDeep(cnt('1,200'), { countText: '1200', countNum: 1200 }, '콤마 제거');
assertDeep(cnt(''), { countText: '', countNum: null }, '빈 값');
assertDeep(cnt(null), { countText: '', countNum: null }, 'null');
assertEqual(cnt('미정').countText, '미정', '숫자가 아니면 원문 유지');
assertEqual(cnt('미정').countNum, null, '숫자가 아니면 합산 대상이 아니다');

/* ── A시트 → 프로그램 맵 ──────────────────────────────────────────────────── */

// A(1) … D(4)=프로그램명, E(5)=갯수 … Q(17)=상품키
function aRow(program, count, productKey) {
  const row = new Array(17).fill('');
  row[3] = program;
  row[4] = count;
  row[16] = productKey;
  return row;
}

{
  const built = gs.progSyncBuildProgramMap_([aRow('리뷰', 3, 111)], cfg);
  assertEqual(built.map['111'], '리뷰 3개', '기본 형식: "프로그램명 갯수개"');
  assertEqual(built.keyCount, 1, '상품 1개');
  assertEqual(built.usedRows, 1, '사용된 행 1개');
}

{
  const built = gs.progSyncBuildProgramMap_([
    aRow('리뷰', 3, 111),
    aRow('트래픽', 2, 111)
  ], cfg);
  assertEqual(built.map['111'], '리뷰 3개, 트래픽 2개', '한 상품의 여러 프로그램은 시트 순서대로 이어붙인다');
}

{
  const built = gs.progSyncBuildProgramMap_([
    aRow('리뷰', 3, 111),
    aRow('리뷰', 2, 111)
  ], cfg);
  assertEqual(built.map['111'], '리뷰 5개', '같은 상품 · 같은 프로그램명은 갯수를 합산한다');
}

{
  const built = gs.progSyncBuildProgramMap_([
    aRow('리뷰', 3, 111),
    aRow('리뷰', '미정', 111)
  ], cfg);
  assertEqual(built.map['111'], '리뷰 3개, 리뷰 미정개', '숫자가 아니면 합산하지 않고 따로 남긴다');
}

{
  const built = gs.progSyncBuildProgramMap_([
    aRow('리뷰', '', 111)
  ], cfg);
  assertEqual(built.map['111'], '리뷰', '갯수가 비면 프로그램명만 쓴다');
}

{
  const built = gs.progSyncBuildProgramMap_([
    aRow('', 3, 111),        // 프로그램명 없음 → 무시
    aRow('리뷰', 3, ''),     // 상품키 없음 → 무시
    aRow('트래픽', 1, 222)
  ], cfg);
  assertDeep(Object.keys(built.map), ['222'], '프로그램명이나 상품키가 비면 건너뛴다');
  assertEqual(built.usedRows, 1, '유효한 행만 센다');
}

{
  // 맵의 키 순서는 의미가 없다(JS는 숫자형 키를 오름차순으로 돌린다). 조회만 되면 된다.
  // 순서가 중요한 것은 "한 상품 안의 프로그램 나열 순서"이고, 그건 위에서 따로 검증한다.
  const built = gs.progSyncBuildProgramMap_([
    aRow('a', 1, 333),
    aRow('b', 1, 111),
    aRow('c', 1, 222)
  ], cfg);
  assertDeep(Object.keys(built.map).sort(), ['111', '222', '333'], '상품키가 모두 조회 가능하다');
  assertEqual(built.keyCount, 3, '상품 3개');
}

{
  // 한 상품 안에서는 시트 행 순서가 반드시 유지돼야 한다(가나다순으로 재정렬 금지).
  const built = gs.progSyncBuildProgramMap_([
    aRow('하편성', 1, 111),
    aRow('가편성', 1, 111),
    aRow('나편성', 1, 111)
  ], cfg);
  assertEqual(built.map['111'], '하편성 1개, 가편성 1개, 나편성 1개',
              '한 상품 안의 프로그램은 시트 행 순서 그대로 (정렬하지 않음)');
}

{
  const built = gs.progSyncBuildProgramMap_([aRow('리뷰', 3, '  111  ')], cfg);
  assertEqual(built.map['111'], '리뷰 3개', 'A시트 키의 공백도 정규화된다');
}

/* ── 대괄호 형식으로 바꿀 수 있는지 ───────────────────────────────────────── */

{
  const bracket = Object.assign({}, cfg, { PAIR_FORMAT: '[{name},{count}]' });
  const built = gs.progSyncBuildProgramMap_([
    aRow('리뷰', 3, 111),
    aRow('트래픽', 2, 111)
  ], bracket);
  assertEqual(built.map['111'], '[리뷰,3], [트래픽,2]', 'PAIR_FORMAT만 바꾸면 대괄호 표기가 된다');
}

/* ── B시트 갱신 계획 ──────────────────────────────────────────────────────── */

// A(1) … E(5)=출력 … K(11)=상품키
function bRow(current, productKey) {
  const row = new Array(11).fill('');
  row[4] = current;
  row[10] = productKey;
  return row;
}

{
  const plan = gs.progSyncComputeUpdates_([bRow('', 111)], { '111': '리뷰 3개' }, cfg);
  assertEqual(plan.updates.length, 1, '빈 셀은 채운다');
  assertEqual(plan.updates[0].row, 2, '헤더 1행이므로 첫 데이터 행은 2행');
  assertEqual(plan.updates[0].next, '리뷰 3개', '기대값');
  assertEqual(plan.matched, 1, '매칭 1건');
}

{
  const plan = gs.progSyncComputeUpdates_([bRow('리뷰 3개', 111)], { '111': '리뷰 3개' }, cfg);
  assertEqual(plan.updates.length, 0, '이미 같은 값이면 쓰지 않는다');
  assertEqual(plan.matched, 1, '매칭은 됐지만 변경은 없다');
}

{
  const plan = gs.progSyncComputeUpdates_([bRow('옛날값', 111)], { '111': '리뷰 3개' }, cfg);
  assertEqual(plan.updates[0].current, '옛날값', '현재값을 기록해 둔다');
  assertEqual(plan.updates[0].next, '리뷰 3개', '새 값으로 덮어쓴다');
}

{
  const plan = gs.progSyncComputeUpdates_([bRow('손으로 적은 값', 999)], { '111': '리뷰 3개' }, cfg);
  assertEqual(plan.updates.length, 0, 'A시트에 없는 키는 기본적으로 건드리지 않는다');
  assertDeep(plan.unmatchedKeys, ['999'], '매칭 실패 키를 보고한다');
  assertEqual(plan.matched, 0, '매칭 0건');
}

{
  const clearing = Object.assign({}, cfg, { CLEAR_WHEN_NO_MATCH: true });
  const plan = gs.progSyncComputeUpdates_([bRow('옛날값', 999)], { '111': '리뷰 3개' }, clearing);
  assertEqual(plan.updates.length, 1, 'CLEAR_WHEN_NO_MATCH=true면 지운다');
  assertEqual(plan.updates[0].next, '', '빈 값으로 지운다');
}

{
  const plan = gs.progSyncComputeUpdates_([bRow('무언가', '')], { '111': '리뷰 3개' }, cfg);
  assertEqual(plan.updates.length, 0, 'K열이 비면 그 행은 건너뛴다');
  assertEqual(plan.blankKeyRows, 1, '건너뛴 행 수를 센다');
}

{
  // B시트 키가 숫자 셀, A시트 키가 텍스트 셀이어도 매칭돼야 한다.
  const built = gs.progSyncBuildProgramMap_([aRow('리뷰', 3, '90817968961')], cfg);
  const plan = gs.progSyncComputeUpdates_([bRow('', 90817968961)], built.map, cfg);
  assertEqual(plan.updates.length, 1, '숫자 셀 ↔ 텍스트 셀도 같은 상품으로 매칭된다');
  assertEqual(plan.updates[0].next, '리뷰 3개', '매칭 결과');
}

/* ── 연속 구간 묶기 (쓰기 호출 최소화 + 중간 행 보호) ─────────────────────── */

{
  const runs = gs.progSyncGroupRuns_([
    { row: 2, next: 'a' },
    { row: 3, next: 'b' },
    { row: 4, next: 'c' }
  ]);
  assertEqual(runs.length, 1, '연속된 행은 한 번에 쓴다');
  assertEqual(runs[0].startRow, 2, '시작 행');
  assertDeep(runs[0].values, [['a'], ['b'], ['c']], '세로 배열');
}

{
  const runs = gs.progSyncGroupRuns_([
    { row: 2, next: 'a' },
    { row: 5, next: 'b' },
    { row: 6, next: 'c' }
  ]);
  assertEqual(runs.length, 2, '중간이 끊기면 구간을 나눈다');
  assertEqual(runs[0].startRow, 2, '첫 구간');
  assertDeep(runs[0].values, [['a']], '첫 구간 값');
  assertEqual(runs[1].startRow, 5, '둘째 구간은 3·4행을 건너뛴다');
  assertDeep(runs[1].values, [['b'], ['c']], '둘째 구간 값');
}

assertEqual(gs.progSyncGroupRuns_([]).length, 0, '변경이 없으면 쓰기 구간도 없다');

/* ── 열 문자 변환 ─────────────────────────────────────────────────────────── */

assertEqual(gs.progSyncColumnLetter_(1), 'A', '1 → A');
assertEqual(gs.progSyncColumnLetter_(5), 'E', '5 → E');
assertEqual(gs.progSyncColumnLetter_(11), 'K', '11 → K');
assertEqual(gs.progSyncColumnLetter_(17), 'Q', '17 → Q');
assertEqual(gs.progSyncColumnLetter_(27), 'AA', '27 → AA');

report('프로그램 동기화 · 순수 로직');
