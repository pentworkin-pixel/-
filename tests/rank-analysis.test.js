/**
 * 「순위 분석.gs」의 순수 로직 테스트.
 *
 * 실행: node tests/rank-analysis.test.js
 */
'use strict';

const { load, assertEqual, assertDeep, assertClose, assertThrows, report } = require('./rank-harness');

const gs = load();
const cfg = gs.RANK_ANALYSIS_CONFIG;

/* ═══ 1. 날짜 해석 ═══════════════════════════════════════════════════════ */

const D = (v) => gs.rankAnalysisParseDate_(v, cfg);

assertEqual(D('2026-08-20'), '2026-08-20', 'yyyy-MM-dd');
assertEqual(D('2026/08/20'), '2026-08-20', 'yyyy/MM/dd');
assertEqual(D('26-08-20'), '2026-08-20', 'yy-MM-dd (시트 머리글 형식)');
assertEqual(D('26-8-5'), '2026-08-05', '한 자리 월/일');
assertEqual(D('2026년 8월 20일'), '2026-08-20', '연·월·일 한글');
assertEqual(D(new Date(2026, 7, 20)), '2026-08-20', 'Date 객체');
assertEqual(D(46254), '2026-08-20', '스프레드시트 날짜 시리얼');
assertEqual(D(''), null, '빈 값');
assertEqual(D(null), null, 'null');
assertEqual(D('상품 MID'), null, '날짜가 아닌 머리글');
assertEqual(D('순위'), null, '"순위"는 날짜가 아니다');
assertEqual(D(3), null, '작은 숫자는 날짜로 오인하지 않는다 (순위 값 보호)');
assertEqual(D(90817968961), null, '큰 숫자(MID)도 날짜가 아니다');
assertEqual(D('2026-02-30'), null, '존재하지 않는 날짜는 거부');
assertEqual(D('2026-13-01'), null, '13월은 거부');

// 2자리 연도 해석 순서를 뒤집을 수 있다
{
  const flipped = Object.assign({}, cfg, { TWO_DIGIT_YEAR_FIRST: false });
  assertEqual(gs.rankAnalysisParseDate_('20-08-26', flipped), '2026-08-20',
              'TWO_DIGIT_YEAR_FIRST=false 면 일-월-연으로 읽는다');
}

/* ═══ 2. 날짜 계산 ═══════════════════════════════════════════════════════ */

assertEqual(gs.rankAnalysisDayNumber_('2026-08-20') - gs.rankAnalysisDayNumber_('2026-08-13'), 7,
            '7일 차이');
assertEqual(gs.rankAnalysisDayNumber_('2026-03-01') - gs.rankAnalysisDayNumber_('2026-02-28'), 1,
            '월 경계');
assertEqual(gs.rankAnalysisDayNumber_('2027-01-01') - gs.rankAnalysisDayNumber_('2026-12-31'), 1,
            '연 경계');

/* ═══ 3. 비교 날짜 선택 ═══════════════════════════════════════════════════ */

const PICK = gs.rankAnalysisPickComparisonDates_;

{
  // 매일 데이터가 있는 경우 → 정확히 1·3·7일 전
  const dates = ['2026-08-13', '2026-08-14', '2026-08-15', '2026-08-16',
                 '2026-08-17', '2026-08-18', '2026-08-19', '2026-08-20'];
  const picks = PICK(dates, '2026-08-20', [1, 3, 7]);
  assertEqual(picks[1], '2026-08-19', '1일 전 정확히 일치');
  assertEqual(picks[3], '2026-08-17', '3일 전 정확히 일치');
  assertEqual(picks[7], '2026-08-13', '7일 전 정확히 일치');
}

{
  // 요구사항 예시: 기준 9/10, 3일 전 목표 9/7 이 없고 9/6 이 있으면 9/6 사용
  const dates = ['2026-09-06', '2026-09-09', '2026-09-10'];
  const picks = PICK(dates, '2026-09-10', [1, 3, 7]);
  assertEqual(picks[1], '2026-09-09', '1일 전: 9/9 존재');
  assertEqual(picks[3], '2026-09-06', '3일 전: 9/7 없으면 그보다 이전인 9/6 사용');
  assertEqual(picks[7], null, '7일 전: 9/3 이전 데이터가 없으면 데이터 없음');
}

{
  // 목표일보다 "나중"인 날짜를 당겨 쓰지 않는다
  const dates = ['2026-09-08', '2026-09-09', '2026-09-10'];
  const picks = PICK(dates, '2026-09-10', [3]);
  assertEqual(picks[3], null, '9/7 이전 데이터가 없으면 9/8을 대신 쓰지 않는다');
}

{
  // 기준일 하나뿐이면 비교 불가
  const picks = PICK(['2026-09-10'], '2026-09-10', [1, 3, 7]);
  assertEqual(picks[1], null, '과거 데이터가 없으면 1일 전도 null');
  assertEqual(picks[7], null, '7일 전도 null');
}

{
  // 데이터가 드문 경우 여러 기간이 같은 날짜를 가리킬 수 있다 (허용 — 실제 비교일을 표시함)
  const picks = PICK(['2026-09-01', '2026-09-10'], '2026-09-10', [1, 3, 7]);
  assertEqual(picks[1], '2026-09-01', '1일 전 → 9/1');
  assertEqual(picks[3], '2026-09-01', '3일 전 → 9/1');
  assertEqual(picks[7], '2026-09-01', '7일 전 → 9/1');
}

/* ═══ 4. 순위 해석 ═══════════════════════════════════════════════════════ */

const R = (v) => gs.rankAnalysisParseRank_(v, cfg);

assertDeep(R(3), { rank: 3, hasRank: true }, '숫자 3위');
assertDeep(R('3'), { rank: 3, hasRank: true }, '문자열 "3"');
assertDeep(R('12위'), { rank: 12, hasRank: true }, '"12위"');
assertDeep(R('1,234'), { rank: 1234, hasRank: true }, '천단위 콤마');
assertEqual(R('').hasRank, false, '빈칸은 순위 없음');
assertEqual(R('-').hasRank, false, '"-"는 순위 없음');
assertEqual(R('미노출').hasRank, false, '"미노출"은 순위 없음');
assertEqual(R('순위 없음').hasRank, false, '"순위 없음"');
assertEqual(R('순위없음').hasRank, false, '"순위없음"(붙여쓰기)');
assertEqual(R('N/A').hasRank, false, '"N/A" (대소문자 무시)');
assertEqual(R('100+').hasRank, false, '"100+"는 숫자로 섞지 않는다');
assertEqual(R(0).hasRank, false, '0위는 유효하지 않다');
assertEqual(R(-5).hasRank, false, '음수는 유효하지 않다');
assertEqual(R(999999).hasRank, false, 'MAX_VALID_RANK 초과는 순위 없음');
assertEqual(R('없는값').hasRank, false, '알 수 없는 문자열은 순위 없음');

/* ═══ 5. 프로그램명 분해 ═════════════════════════════════════════════════ */

const P = (v) => gs.rankAnalysisSplitPrograms_(v, cfg);

assertDeep(P('리뷰 3개'), ['리뷰'], '갯수 접미사를 떼어 프로그램명만 남긴다');
assertDeep(P('리뷰 3개, 트래픽 2개'), ['리뷰', '트래픽'], '여러 프로그램 분해');
assertDeep(P('리뷰'), ['리뷰'], '갯수가 없어도 동작');
assertDeep(P(''), [], '빈 칸은 프로그램 없음');
assertDeep(P('   '), [], '공백만 있어도 프로그램 없음');
assertDeep(P('리뷰 3개, 리뷰 2개'), ['리뷰'], '같은 프로그램은 중복 제거');
assertDeep(P('리뷰/트래픽'), ['리뷰', '트래픽'], '슬래시 구분자');

{
  const noSplit = Object.assign({}, cfg, { SPLIT_PROGRAMS: false });
  assertDeep(gs.rankAnalysisSplitPrograms_('리뷰 3개, 트래픽 2개', noSplit),
             ['리뷰 3개, 트래픽 2개'], 'SPLIT_PROGRAMS=false 면 문자열 전체가 하나의 프로그램');
}

/* ═══ 6. 열 감지 ═════════════════════════════════════════════════════════ */

{
  // 가로형: 날짜가 열 머리글 (원 헤더에 "프로그램"이라는 단어가 그대로 있는 경우)
  const header = ['메모1', '리워드', '판매처명', '상품링크', '프로그램',
                  '전일 비교', '순위 조회 키워드', '상품 MID', '가격비교 MID',
                  '26-08-20', '26-08-19', '26-08-17', '26-08-13'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);

  assertEqual(det.layout, 'wide', '날짜 머리글이 2개 이상이면 가로형');
  assertEqual(det.dateCols.length, 4, '날짜 열 4개 감지');
  assertDeep(det.dateCols.map((d) => d.dateKey),
             ['2026-08-13', '2026-08-17', '2026-08-19', '2026-08-20'],
             '날짜 열은 오름차순으로 정렬된다');
  assertEqual(det.programCol, 5, '"프로그램" 머리글로 E열 감지');
  assertDeep(det.keyCols, [8, 7], '고유키 = 상품 MID(H) + 키워드(G) 우선순위 순');
  assertEqual(det.keyCols.indexOf(1), -1,
              '메모 열은 키에 넣지 않는다 — 메모를 고치면 항목이 갈라지기 때문');
}

{
  // 강한 식별자(MID)가 없을 때만 이름·메모 열을 예비 키로 쓴다
  const header = ['상품명', 'b', 'c', 'd', '프로그램', 'f', '순위 조회 키워드', 'h',
                  '26-08-20', '26-08-19'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  assertDeep(det.keyCols, [7, 1], 'MID가 없으면 키워드 + 상품명으로 키를 만든다');
}

{
  // 식별 열을 하나도 못 찾으면 행 번호를 키로 쓴다 (가로형에서는 행 자체가 항목)
  const header = ['a', 'b', 'c', 'd', '프로그램', 'f', 'g', 'h', '26-08-20', '26-08-19'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  assertDeep(det.keyCols, [], '식별 열 없음');
  const row = new Array(10).fill('');

  // 행 번호 계산 자체는 HEADER_ROW 가 몇이든 "헤더 다음 행부터"라는 규칙만 따른다.
  const headerAt1 = Object.assign({}, cfg, { HEADER_ROW: 1 });
  assertEqual(gs.rankAnalysisBuildItemKey_(row, det, 0, headerAt1), '행 2',
              'HEADER_ROW=1 이면 첫 데이터 행은 "행 2"');
  assertEqual(gs.rankAnalysisBuildItemKey_(row, det, 0, cfg), '행 ' + (cfg.HEADER_ROW + 1),
              '실제 기본값(HEADER_ROW=6)이면 첫 데이터 행은 "행 7"');
}

{
  // 프로그램 머리글이 없으면 설정된 대체 열(E)을 쓴다
  const header = ['a', 'b', 'c', 'd', 'e', 'f', '키워드', '상품 MID', '26-08-20', '26-08-19'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  assertEqual(det.programCol, 5, '프로그램 머리글이 없으면 PROGRAM_COL_FALLBACK(E열)');
}

{
  // 세로형: 날짜 열 + 순위 열
  const header = ['업데이트 날짜', '프로그램명', '순위', '순위 조회 키워드', '상품 MID'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);

  assertEqual(det.layout, 'long', '날짜 머리글이 없으면 세로형');
  assertEqual(det.dateCol, 1, '날짜 열 = A');
  assertEqual(det.programCol, 2, '프로그램 열 = B');
  assertEqual(det.rankCol, 3, '순위 열 = C');
  assertDeep(det.keyCols, [5, 4], '고유키 = 상품 MID + 키워드');
  assertEqual(det.keyCols.indexOf(det.dateCol), -1, '날짜 열은 고유키에 넣지 않는다');
  assertEqual(det.keyCols.indexOf(det.programCol), -1, '프로그램 열도 고유키에 넣지 않는다');
}

{
  // "순위 조회 키워드"·"전일 비교 순위(자유 입력)"를 순위 열로 오인하지 않는다
  const header = ['업데이트 날짜', '프로그램명', '순위 조회 키워드',
                  '전일 비교 순위(자유 입력)', '순위', '상품 MID'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  assertEqual(det.rankCol, 5, '진짜 "순위" 열(E)을 고른다');
  assertEqual(det.dateCol, 1, '날짜 열은 A');
}

{
  // 날짜 열이 하나뿐이어도(하루치만 쌓인 초기 상태) 가로형으로 읽는다
  const header = ['메모1', 'b', 'c', 'd', '프로그램', 'f', '순위 조회 키워드',
                  '상품 MID', '26-08-20'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  assertEqual(det.layout, 'wide', '날짜 열 1개 + 날짜/순위 짝 없음 → 가로형');
  assertEqual(det.dateCols.length, 1, '날짜 열 1개');
}

{
  // 필수 열을 못 찾으면 조용히 넘어가지 않는다
  const header = ['이것', '저것', '그것'];
  assertThrows(() => gs.rankAnalysisDetectColumns_(header, [], cfg),
               '날짜·순위 열을 못 찾으면 명확한 오류를 던진다');
}

/* ── 실제 B시트(6행 헤더) 구조 — 하루에 여러 번 체크하는 봇 생성 시트 ─────── */

// D=메모1, E=리워드명/갯수, F=판매처명, G=상품링크, H=N개전, I=전일비교, J=키워드,
// K=상품 MID, L=가격비교 MID, M~ = 타임스탬프("yy-MM-dd HH:mm")
const REAL_HEADER = ['상품명', '재고', '상품명2', '메모1', '리워드명 / 갯수', '판매처명',
                     '상품링크', '4개전', '전일 비교 순위(자유 입력)', '순위 조회 키워드',
                     '상품 MID', '가격비교 MID',
                     '26-09-04 16:12', '26-09-04 14:31', '26-09-03 23:36', '26-08-20 08:14'];

{
  const det = gs.rankAnalysisDetectColumns_(REAL_HEADER, [], cfg);
  assertEqual(det.layout, 'wide', '실제 B시트도 가로형으로 인식');
  assertEqual(det.programCol, 5, '"리워드명 / 갯수" 헤더로 E열을 프로그램 열로 잡는다');
  assertDeep(det.keyCols, [11, 10], '고유키 = 상품 MID(K) + 순위 조회 키워드(J)');
  // "4개전"/"전일 비교 순위"는 순위 열도 날짜 열도 아니어야 한다
  assertEqual(gs.rankAnalysisParseDate_(REAL_HEADER[7], cfg), null, '"4개전"은 날짜가 아니다');
}

{
  // 같은 날짜(26-09-04)에 열이 2개(16:12, 14:31) — 더 늦은 시각(16:12) 하나만 대표로 쓴다
  const det = gs.rankAnalysisDetectColumns_(REAL_HEADER, [], cfg);
  const sep04 = det.dateCols.filter((d) => d.dateKey === '2026-09-04');
  assertEqual(sep04.length, 1, '같은 날짜 열은 하나로 합쳐진다 (중복으로 세지 않음)');
  assertEqual(sep04[0].col, 13, '그날 더 늦은 시각(16:12, M열)이 대표로 선택된다');
  assertEqual(det.dateCols.length, 3, '9/4·9/3·8/20 세 날짜만 남는다 (14:31 열은 흡수됨)');
}

{
  // 시:분 스코어 비교
  assertEqual(gs.rankAnalysisHeaderTimeScore_('26-09-04 16:12'), 16 * 60 + 12, '16:12 → 972분');
  assertEqual(gs.rankAnalysisHeaderTimeScore_('26-09-04 14:31'), 14 * 60 + 31, '14:31 → 871분');
  assertEqual(gs.rankAnalysisHeaderTimeScore_('26-09-04'), -1, '시각이 없으면 -1');
  assertEqual(gs.rankAnalysisHeaderTimeScore_('26-09-04 25:99'), -1, '말이 안 되는 시각은 무시');
}

/* ── 999 같은 상한값을 순위 없음으로 취급 (실제 B시트 안내문 근거로 기본 켬) ── */

{
  // 실제 B시트 4행 안내문("1000위 밖일 경우 확인 불가")과 999가 같은 상품에
  // 수십 번 연속 반복되는 관측을 근거로, 기본 설정에서 999를 순위 없음으로 본다.
  assertDeep(cfg.RANK_CAP_VALUES, [999], '기본 설정은 999를 측정 상한으로 본다');
  assertEqual(gs.rankAnalysisParseRank_(999, cfg).hasRank, false,
              '기본 설정에서 999는 순위 없음으로 처리된다');
  assertEqual(gs.rankAnalysisParseRank_('999', cfg).hasRank, false,
              '문자열 "999" 도 동일하게 처리된다');
  assertEqual(gs.rankAnalysisParseRank_(998, cfg).hasRank, true,
              '998처럼 상한에 안 걸리는 값은 그대로 순위로 본다');

  const uncapped = Object.assign({}, cfg, { RANK_CAP_VALUES: [] });
  assertEqual(gs.rankAnalysisParseRank_(999, uncapped).hasRank, true,
              'RANK_CAP_VALUES를 빈 배열로 바꾸면 999도 정상 순위로 처리된다');
}

{
  // 날짜 열이 하나도 없고 날짜/순위 짝도 없으면 오류
  const header = ['메모1', 'b', 'c', 'd', '프로그램', 'f', '키워드', '상품 MID', '순위'];
  assertThrows(() => gs.rankAnalysisDetectColumns_(header, [], cfg),
               '날짜를 전혀 못 찾으면 오류를 던진다');
}

/* ═══ 7. 관측 만들기 ═════════════════════════════════════════════════════ */

/** 가로형 한 행: E(5)=프로그램, G(7)=키워드, H(8)=MID, I·J·K(9~11)=날짜별 순위 */
function wideHeader() {
  return ['메모', 'b', 'c', 'd', '프로그램', 'f', '순위 조회 키워드', '상품 MID',
          '26-08-20', '26-08-19', '26-08-13'];
}
function wideRow(program, keyword, mid, r20, r19, r13) {
  return ['메모', '', '', '', program, '', keyword, mid, r20, r19, r13];
}

{
  const header = wideHeader();
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  const obs = gs.rankAnalysisBuildObservations_([
    wideRow('리뷰 3개', '자외선차단마스크', 90817968961, 3, 10, 12)
  ], det, cfg);

  assertDeep(obs.dateKeys, ['2026-08-13', '2026-08-19', '2026-08-20'],
             '날짜 키는 오름차순');
  const key = '90817968961 ｜ 자외선차단마스크';
  assertDeep(obs.programsByItem[key], ['리뷰'], '항목의 프로그램');
  assertEqual(obs.byDate['2026-08-20'][key].rank, 3, '기준일 순위');
  assertEqual(obs.byDate['2026-08-13'][key].rank, 12, '7일 전 순위');
}

{
  // 중복: 같은 날짜 + 같은 고유항목이면 마지막 행이 이긴다
  const header = wideHeader();
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  const obs = gs.rankAnalysisBuildObservations_([
    wideRow('리뷰 3개', '마스크', 111, 5, 5, 5),
    wideRow('리뷰 3개', '마스크', 111, 2, 2, 2)
  ], det, cfg);

  const key = '111 ｜ 마스크';
  assertEqual(obs.byDate['2026-08-20'][key].rank, 2, '중복이면 마지막 행의 값을 쓴다');
  assertEqual(obs.issues.duplicates, 3, '중복 건수를 센다 (날짜 3개 × 1중복)');
}

{
  // 프로그램명이 비면 집계 대상에서 빠지고 카운트만 된다
  const header = wideHeader();
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  const obs = gs.rankAnalysisBuildObservations_([
    wideRow('', '마스크', 111, 5, 5, 5)
  ], det, cfg);
  assertEqual(obs.issues.blankProgram, 1, '프로그램명 없는 행을 센다');
  assertEqual(obs.programsByItem['111 ｜ 마스크'], undefined, '프로그램 귀속이 없다');
}

{
  // 세로형 관측
  const header = ['날짜', '프로그램명', '순위', '키워드', '상품 MID'];
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  const obs = gs.rankAnalysisBuildObservations_([
    ['2026-08-20', '리뷰', 3, '마스크', 111],
    ['2026-08-19', '리뷰', 10, '마스크', 111],
    ['잘못된날짜', '리뷰', 7, '마스크', 111]
  ], det, cfg);

  assertDeep(obs.dateKeys, ['2026-08-19', '2026-08-20'], '유효한 날짜만 수집');
  assertEqual(obs.issues.badDate, 1, '날짜 형식 오류 행을 센다');
  assertEqual(obs.byDate['2026-08-20']['111 ｜ 마스크'].rank, 3, '세로형 순위 수집');
}

/* ═══ 8. 집계 ════════════════════════════════════════════════════════════ */

/** 가로형 데이터로 obs 를 만드는 헬퍼. rows = [program, keyword, mid, r20, r19, r13] */
function buildObs(rows) {
  const header = wideHeader();
  const det = gs.rankAnalysisDetectColumns_(header, [], cfg);
  return gs.rankAnalysisBuildObservations_(
    rows.map((r) => wideRow(r[0], r[1], r[2], r[3], r[4], r[5])), det, cfg);
}

{
  // 요구사항의 예시 그대로: 10위→3위 상승, 3위→3위 유지, 3위→10위 하락
  const obs = buildObs([
    ['리뷰 3개', 'a', 1, 3, 10, ''],    // 10 → 3 : 7계단 상승
    ['리뷰 3개', 'b', 2, 3, 3, ''],     // 3 → 3  : 유지
    ['리뷰 3개', 'c', 3, 10, 3, '']     // 3 → 10 : 7계단 하락
  ]);
  const agg = gs.rankAnalysisAggregate_(obs, '2026-08-20', '2026-08-19', cfg);
  const p = agg.programs[0];

  assertEqual(p.program, '리뷰', '프로그램명');
  assertEqual(p.comparable, 3, '비교 가능 3건');
  assertEqual(p.up, 1, '상승 1');
  assertEqual(p.same, 1, '유지 1');
  assertEqual(p.down, 1, '하락 1');
  assertClose(p.upRate, 1 / 3, '상승률 = 1/3');
  assertClose(p.sameRate, 1 / 3, '유지율 = 1/3');
  assertClose(p.downRate, 1 / 3, '하락률 = 1/3');
  assertClose(p.avgDelta, 0, '평균 변화폭 = (7 + 0 - 7)/3 = 0');
  assertClose(p.avgUp, 7, '평균 상승 폭 = 7계단');
  assertClose(p.avgDown, 7, '평균 하락 폭 = 7계단 (양수 크기로 표시)');
  assertClose(p.latestAvgRank, (3 + 3 + 10) / 3, '최신 평균 순위');
}

{
  // 순위가 작아질수록 상승이라는 방향이 뒤집히지 않았는지 못 박아 둔다
  const obs = buildObs([['리뷰', 'a', 1, 3, 10, '']]);
  const agg = gs.rankAnalysisAggregate_(obs, '2026-08-20', '2026-08-19', cfg);
  assertEqual(agg.programs[0].up, 1, '10위 → 3위는 상승이다');
  assertEqual(agg.programs[0].down, 0, '하락이 아니다');
  assertClose(agg.programs[0].avgDelta, 7, '변화폭은 +7');
  assertEqual(agg.details[0].status, '상승', '상세 상태도 상승');
}

{
  // 신규 / 누락은 비율 분모에서 빠진다
  const obs = buildObs([
    ['리뷰', 'a', 1, 3, 10, ''],        // 비교 가능 (상승)
    ['리뷰', 'b', 2, 5, '-', ''],       // 과거 순위 없음 → 신규
    ['리뷰', 'c', 3, '미노출', 4, '']   // 최신 순위 없음 → 누락 또는 종료
  ]);
  const agg = gs.rankAnalysisAggregate_(obs, '2026-08-20', '2026-08-19', cfg);
  const p = agg.programs[0];

  assertEqual(p.comparable, 1, '비교 가능은 1건뿐');
  assertEqual(p.newCount, 1, '신규 1건');
  assertEqual(p.lostCount, 1, '누락 또는 종료 1건');
  assertClose(p.upRate, 1, '상승률 분모는 비교 가능 항목 수(1)이므로 100%');
  assertClose(p.latestAvgRank, 4, '최신 평균 순위 = (3+5)/2 = 4 (순위 없는 항목 제외)');

  const statuses = agg.details.map((d) => d.status).sort();
  assertDeep(statuses, ['누락 또는 종료', '상승', '신규'].sort(),
             '상세에 상승/신규/누락이 각각 기록된다');
}

{
  // 양쪽 다 순위 없음이면 아예 집계하지 않는다
  const obs = buildObs([['리뷰', 'a', 1, '-', '미노출', '']]);
  const agg = gs.rankAnalysisAggregate_(obs, '2026-08-20', '2026-08-19', cfg);
  assertEqual(agg.programs.length, 0, '집계할 것이 없으면 프로그램도 나오지 않는다');
  assertEqual(agg.details.length, 0, '상세도 비어 있다');
}

{
  // 한 항목에 프로그램이 둘이면 양쪽에 귀속되고, 단독 항목 수로 구분할 수 있다
  const obs = buildObs([
    ['리뷰 3개, 트래픽 2개', 'a', 1, 3, 10, ''],
    ['리뷰 1개', 'b', 2, 4, 4, '']
  ]);
  const agg = gs.rankAnalysisAggregate_(obs, '2026-08-20', '2026-08-19', cfg);
  const byName = {};
  agg.programs.forEach((p) => { byName[p.program] = p; });

  assertEqual(byName['리뷰'].comparable, 2, '리뷰는 두 항목 모두에 귀속');
  assertEqual(byName['트래픽'].comparable, 1, '트래픽은 한 항목에만 귀속');
  assertEqual(byName['리뷰'].soloComparable, 1, '리뷰 단독 항목은 1건');
  assertEqual(byName['트래픽'].soloComparable, 0, '트래픽 단독 항목은 0건 — 해석 주의 신호');
}

/* ═══ 9. 정렬 ════════════════════════════════════════════════════════════ */

{
  const programs = [
    { program: '느림', comparable: 10, upRate: 0.2, avgDelta: 1, downRate: 0.3 },
    { program: '빠름', comparable: 10, upRate: 0.8, avgDelta: 2, downRate: 0.1 },
    { program: '중간', comparable: 10, upRate: 0.5, avgDelta: 3, downRate: 0.2 }
  ];
  gs.rankAnalysisSortPrograms_(programs);
  assertDeep(programs.map((p) => p.program), ['빠름', '중간', '느림'], '상승률 높은 순');
}

{
  // 상승률이 같으면 평균 변화폭이 큰 쪽이 위
  const programs = [
    { program: 'A', comparable: 5, upRate: 0.5, avgDelta: 1, downRate: 0.2 },
    { program: 'B', comparable: 5, upRate: 0.5, avgDelta: 9, downRate: 0.2 }
  ];
  gs.rankAnalysisSortPrograms_(programs);
  assertDeep(programs.map((p) => p.program), ['B', 'A'], '2순위: 평균 순위 변화폭');
}

{
  // 상승률·변화폭이 같으면 하락률이 낮은 쪽이 위
  const programs = [
    { program: 'A', comparable: 5, upRate: 0.5, avgDelta: 1, downRate: 0.4 },
    { program: 'B', comparable: 5, upRate: 0.5, avgDelta: 1, downRate: 0.1 }
  ];
  gs.rankAnalysisSortPrograms_(programs);
  assertDeep(programs.map((p) => p.program), ['B', 'A'], '3순위: 하락률 낮은 순');
}

{
  // 앞이 모두 같으면 표본이 많은 쪽이 위
  const programs = [
    { program: 'A', comparable: 2, upRate: 1, avgDelta: 5, downRate: 0 },
    { program: 'B', comparable: 30, upRate: 1, avgDelta: 5, downRate: 0 }
  ];
  gs.rankAnalysisSortPrograms_(programs);
  assertDeep(programs.map((p) => p.program), ['B', 'A'], '4순위: 비교 항목 수 많은 순');
}

{
  // 비교 가능 항목이 0개면 비율이 없으므로 항상 뒤로 보낸다
  const programs = [
    { program: '표본없음', comparable: 0, upRate: null, avgDelta: null, downRate: null },
    { program: '보통', comparable: 4, upRate: 0.25, avgDelta: 0.5, downRate: 0.5 }
  ];
  gs.rankAnalysisSortPrograms_(programs);
  assertDeep(programs.map((p) => p.program), ['보통', '표본없음'],
             '비교 불가 프로그램이 1위로 올라오지 않는다');
}

/* ═══ 10. 유틸 ═══════════════════════════════════════════════════════════ */

assertEqual(gs.rankAnalysisColumnLetter_(5), 'E', '5 → E');
assertEqual(gs.rankAnalysisColumnLetter_(11), 'K', '11 → K');
assertEqual(gs.rankAnalysisColumnLetter_(17), 'Q', '17 → Q');
assertDeep(gs.rankAnalysisPad2_(['a'], 3), ['a', '', ''], '행 길이를 맞춘다');
assertEqual(gs.rankAnalysisPercent_(0.3333), '33.3%', '퍼센트 소수점 1자리');
assertEqual(gs.rankAnalysisPercent_(null), '-', '값이 없으면 하이픈');
assertEqual(gs.rankAnalysisFixed_(1.239), '1.24', '소수점 2자리');
assertEqual(gs.rankAnalysisFixed_(null), '-', '값이 없으면 하이픈');

/* ═══ 11. 설정이 요구사항과 일치하는지 ═══════════════════════════════════ */

assertDeep(cfg.PERIODS, [1, 3, 7], '비교 기간은 1·3·7일');
assertEqual(cfg.ANALYSIS_SHEET_NAME, '프로그램별 순위 분석', '분석 시트 이름');
assertEqual(cfg.DETAIL_SHEET_NAME, '프로그램별 순위 상세', '상세 시트 이름');
assertEqual(typeof gs.updateProgramRankingAnalysis, 'function',
            '요구사항이 지정한 메인 함수 이름이 존재한다');
assertEqual(typeof gs.onOpen, 'function', '메뉴용 onOpen 이 존재한다');
assertDeep(gs.RANK_ANALYSIS_TABLE_HEADERS.slice(0, 15), [
  '순위', '프로그램명', '비교 가능 항목 수', '상승 수', '유지 수', '하락 수',
  '신규 수', '누락 또는 종료 수', '상승률', '유지율', '하락률',
  '평균 순위 변화폭', '평균 상승 폭', '평균 하락 폭', '최신 평균 순위'
], '요약표 열 구성이 요구사항 순서와 같다');
assertDeep(gs.RANK_ANALYSIS_DETAIL_HEADERS, [
  '프로그램명', '고유 항목명', '최신 순위', '과거 순위', '순위 변화폭',
  '상태', '최신 업데이트 날짜', '실제 비교 날짜', '비교 기간'
], '상세표 열 구성이 요구사항과 같다');

report('순위 분석 · 순수 로직');
