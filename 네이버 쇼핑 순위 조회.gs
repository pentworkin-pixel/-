/**
 * 네이버 쇼핑 모바일 검색(msearch.shopping.naver.com) 결과에서
 * 광고 상품을 제외한 순위를 조회해 스프레드시트에 기록하는 Apps Script입니다.
 *
 * 실제 운영 중인 시트("시트1")의 구조에 맞춰져 있습니다:
 *   A 상품명(참고) · B 리워드명(참고) · C 갯수(참고) · D 판매처명(참고)
 *   E 상품링크(참고) · F 순위 변동(이 스크립트가 건드리지 않음)
 *   G 순위 조회 키워드 · H 상품 MID · I 가격비교 MID · J 월검색량(참고)
 *   K열부터: 조회할 때마다 왼쪽에 새 열이 하나씩 삽입되고, 헤더에 조회 시각,
 *            각 행에는 그 시점의 순위(광고 제외) 또는 999(못 찾음/차단 등)가 채워짐
 *
 * 이 파일의 모든 전역 이름은 naverRank / NAVER_RANK 접두사를 쓰므로
 * 기존 코드(광고 종료 캘린더 동기화 등)와 충돌하지 않습니다.
 *
 * ⚠️ 중요한 한계
 * msearch.shopping.naver.com은 데이터센터 IP(포함: Google Apps Script 서버)에서의
 * 접근을 비정상 트래픽으로 판단해 HTTP 418로 차단하는 경우가 있습니다. 이 스크립트는
 * 이를 우회하지 않고, 차단되거나 상품을 못 찾으면 그 행에 999를 적을 뿐입니다.
 * naverRankDebugFetch로 먼저 접속 상태를 확인하세요.
 */

var NAVER_RANK_CONFIG = {
  SHEET_NAME: '시트1',
  TIMEZONE: 'Asia/Seoul',
  HEADER_ROW: 3,           // 헤더가 있는 행
  FIRST_DATA_ROW: 4,       // 상품 데이터가 시작하는 행
  COL_KEYWORD: 7,          // G: 순위 조회 키워드
  COL_PRODUCT_MID: 8,      // H: 상품 MID
  COL_COMPARE_MID: 9,      // I: 가격비교 MID
  FIRST_HISTORY_COL: 11,   // K: 조회 결과 열이 매번 삽입되는 위치
  NOT_FOUND_RANK: 999,     // 못 찾음/차단/오류일 때 기록할 값 (시트의 기존 관례)
  MAX_SCAN: 200,           // 한 번의 응답에서 순위 계산에 사용할 최대 상품 수
  REQUEST_DELAY_MS: 1200,  // 서로 다른 키워드 요청 사이 대기 시간
};

/**
 * 모든 상품 행의 순위를 조회해 맨 왼쪽(K열)에 새 결과 열을 삽입합니다.
 * 오전 11시 / 오후 5시에 자동 실행되도록 하려면 naverRankInstall을 한 번 실행하세요.
 */
function naverRankUpdateAll() {
  const sheet = naverRankGetSheet_();
  const lastRow = sheet.getLastRow();
  const firstDataRow = NAVER_RANK_CONFIG.FIRST_DATA_ROW;
  if (lastRow < firstDataRow) {
    Logger.log('조회할 상품 행이 없습니다.');
    return;
  }

  const numRows = lastRow - firstDataRow + 1;
  const idValues = sheet.getRange(firstDataRow, NAVER_RANK_CONFIG.COL_KEYWORD, numRows, 3).getValues();

  const cache = {}; // 키워드 → naverRankFetchRanking_ 결과 (같은 키워드 중복 요청 방지)
  const ranks = new Array(numRows);
  const keywordsFetchedThisRun = [];

  for (let i = 0; i < numRows; i++) {
    const keyword = String(idValues[i][0] || '').trim();
    const mid = idValues[i][1];
    const compareMid = idValues[i][2];

    if (!keyword || !mid) {
      ranks[i] = '';
      continue;
    }

    if (!(keyword in cache)) {
      if (keywordsFetchedThisRun.length > 0) Utilities.sleep(NAVER_RANK_CONFIG.REQUEST_DELAY_MS);
      cache[keyword] = naverRankFetchRanking_(keyword);
      keywordsFetchedThisRun.push(keyword);
    }

    const fetched = cache[keyword];
    if (!fetched.ok) {
      ranks[i] = NAVER_RANK_CONFIG.NOT_FOUND_RANK;
      continue;
    }

    const items = fetched.items.slice(0, NAVER_RANK_CONFIG.MAX_SCAN);
    const found = naverRankFindByMid_(items, mid, compareMid);
    ranks[i] = found ? found.rank : NAVER_RANK_CONFIG.NOT_FOUND_RANK;
  }

  naverRankInsertResultColumn_(sheet, firstDataRow, ranks);
  Logger.log(`완료: 키워드 ${keywordsFetchedThisRun.length}건 조회, 상품 ${numRows}행 기록`);
}

/** 새 결과 열을 K열 앞에 삽입하고, 헤더(조회 시각)와 순위 값을 채운다. */
function naverRankInsertResultColumn_(sheet, firstDataRow, ranks) {
  const col = NAVER_RANK_CONFIG.FIRST_HISTORY_COL;
  sheet.insertColumnBefore(col);

  const headerCell = sheet.getRange(NAVER_RANK_CONFIG.HEADER_ROW, col);
  headerCell.setNumberFormat('yy-mm-dd h:mm');
  headerCell.setValue(new Date());

  sheet.getRange(firstDataRow, col, ranks.length, 1).setValues(ranks.map((r) => [r]));
}

/** 접속 차단 여부·페이지 구조 변화를 진단하기 위한 디버그용 함수. */
function naverRankDebugFetch(keyword) {
  const q = keyword || '테스트';
  const fetched = naverRankFetchRanking_(q);
  if (!fetched.ok) {
    Logger.log('❌ 조회 실패: ' + fetched.note);
    return fetched;
  }
  Logger.log(`✅ 광고 제외 상품 ${fetched.items.length}개 확인됨 (키워드: ${q})`);
  fetched.items.slice(0, 10).forEach((it) => {
    Logger.log(`${it.rank}위 | ${it.name} | ${it.mallName} | MID:${it.productId}`);
  });
  return fetched;
}

/** 시트 없이 키워드+MID 하나만 바로 확인하고 싶을 때 스크립트 편집기에서 직접 실행합니다. */
function naverRankCheckKeywordMid(keyword, mid, compareMid) {
  const fetched = naverRankFetchRanking_(keyword);
  if (!fetched.ok) {
    Logger.log('❌ 조회 실패: ' + fetched.note);
    return { rank: NAVER_RANK_CONFIG.NOT_FOUND_RANK, note: fetched.note };
  }
  const items = fetched.items.slice(0, NAVER_RANK_CONFIG.MAX_SCAN);
  const found = naverRankFindByMid_(items, mid, compareMid);
  const result = found
    ? { rank: found.rank, name: found.name, mallName: found.mallName }
    : { rank: NAVER_RANK_CONFIG.NOT_FOUND_RANK, note: `미노출 (광고 제외 상위 ${items.length}개 안에서 못 찾음)` };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * 네이버 쇼핑 모바일 검색 결과를 가져와 광고를 제외한 순위 목록을 만듭니다.
 * 반환값: { ok:true, items:[{rank,name,mallName,productId}] } 또는 { ok:false, note }
 */
function naverRankFetchRanking_(keyword) {
  const url = 'https://msearch.shopping.naver.com/search/all?query=' + encodeURIComponent(keyword);

  let response;
  try {
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 '
          + '(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
        'Accept-Language': 'ko-KR,ko;q=0.9',
        'Referer': 'https://msearch.shopping.naver.com/',
      },
    });
  } catch (e) {
    return { ok: false, note: '요청 실패: ' + e.message };
  }

  const code = response.getResponseCode();
  if (code === 418 || code === 403 || code === 429) {
    return { ok: false, note: `접속 차단됨 (HTTP ${code}) - 잠시 후 다시 시도하세요` };
  }
  if (code !== 200) {
    return { ok: false, note: `HTTP 오류 (${code})` };
  }

  const html = response.getContentText();
  const nextData = naverRankExtractNextData_(html);
  if (!nextData) {
    return { ok: false, note: '검색 결과 데이터를 찾지 못함 (페이지 구조가 바뀌었을 수 있음)' };
  }

  const productArray = naverRankFindProductArray_(nextData);
  if (!productArray) {
    return { ok: false, note: '상품 목록을 찾지 못함 (페이지 구조가 바뀌었을 수 있음)' };
  }

  const items = [];
  let rank = 0;
  for (const raw of productArray) {
    if (naverRankIsAdItem_(raw)) continue;
    rank++;
    items.push({
      rank: rank,
      name: naverRankPick_(raw, ['productName', 'title', 'name']),
      mallName: naverRankPick_(raw, ['mallName'])
        || naverRankPickNested_(raw, ['mall', 'name'])
        || naverRankPickNested_(raw, ['channel', 'name'])
        || '',
      productId: naverRankPick_(raw, ['productId', 'nvMid', 'id', 'itemId']),
    });
  }

  return { ok: true, items: items };
}

/** HTML에서 Next.js의 `__NEXT_DATA__` JSON 블록을 추출합니다. */
function naverRankExtractNextData_(html) {
  const m = /<script[^>]*id="__NEXT_DATA__"[^>]*>([\s\S]*?)<\/script>/.exec(html);
  if (!m) return null;
  try {
    return JSON.parse(m[1]);
  } catch (e) {
    return null;
  }
}

/**
 * __NEXT_DATA__ 트리 안에서 상품 목록으로 보이는 배열을 찾는다.
 * 정확한 경로를 하드코딩하지 않고, "이름 필드 + id 필드를 가진 객체들의 배열"이라는
 * 모양만 보고 찾는다 — 네이버가 내부 데이터 구조/경로를 바꿔도 버틸 수 있게 하기 위함.
 */
function naverRankFindProductArray_(root) {
  const NAME_KEYS = ['productName', 'title', 'name'];
  const ID_KEYS = ['productId', 'nvMid', 'id', 'itemId'];
  let best = null;

  function isProductLike(obj) {
    if (!obj || typeof obj !== 'object') return false;
    const hasName = NAME_KEYS.some((k) => typeof obj[k] === 'string' && obj[k].length > 0);
    const hasId = ID_KEYS.some((k) => obj[k] !== undefined && obj[k] !== null && obj[k] !== '');
    return hasName && hasId;
  }

  function walk(node, depth) {
    if (!node || typeof node !== 'object' || depth > 25) return;
    if (Array.isArray(node)) {
      const productLikeCount = node.filter(isProductLike).length;
      if (productLikeCount >= 2 && productLikeCount >= node.length * 0.5) {
        if (!best || node.length > best.length) best = node;
      }
      node.forEach((child) => walk(child, depth + 1));
    } else {
      Object.keys(node).forEach((k) => walk(node[k], depth + 1));
    }
  }

  walk(root, 0);
  return best;
}

/**
 * 상품 항목이 광고인지 판단한다. 필드명을 확정할 수 없어 흔히 쓰이는
 * 여러 이름 후보를 넓게 검사한다 (naverRankDebugFetch로 실제 값을 보고 조정 가능).
 */
function naverRankIsAdItem_(item) {
  if (!item || typeof item !== 'object') return false;

  const truthyAdFlags = ['isAd', 'isAdultAd', 'adFlag', 'ad'];
  for (const k of truthyAdFlags) {
    if (item[k] === true) return true;
  }

  const adPresenceFields = ['adId', 'admarker', 'adcId', 'adcExtensionParameters', 'nclickAd'];
  for (const k of adPresenceFields) {
    if (item[k]) return true;
  }

  const typeFields = ['cardType', 'type', 'itemType', 'badge', 'adBadgeType'];
  for (const k of typeFields) {
    const v = item[k];
    if (typeof v === 'string' && (v.toLowerCase().indexOf('ad') !== -1 || v.indexOf('광고') !== -1)) return true;
  }

  return false;
}

/**
 * 순위 목록에서 상품 MID 또는 가격비교 MID가 일치하는 첫 항목을 찾는다.
 * 검색 결과 카드가 개별 상품 MID 대신 가격비교 MID로 노출되는 경우가 있어 둘 다 검사한다.
 */
function naverRankFindByMid_(items, mid, compareMid) {
  const midStr = mid !== '' && mid !== null && mid !== undefined ? String(mid) : '';
  const compareMidStr = compareMid !== '' && compareMid !== null && compareMid !== undefined ? String(compareMid) : '';

  for (const item of items) {
    const itemId = String(item.productId || '');
    if (!itemId) continue;
    if (midStr && itemId === midStr) return item;
    if (compareMidStr && itemId === compareMidStr) return item;
  }
  return null;
}

function naverRankPick_(obj, keys) {
  for (const k of keys) {
    const v = obj && obj[k];
    if (typeof v === 'string' && v) return v;
    if (typeof v === 'number') return String(v);
  }
  return '';
}

function naverRankPickNested_(obj, path) {
  let cur = obj;
  for (const k of path) {
    if (!cur || typeof cur !== 'object') return '';
    cur = cur[k];
  }
  return typeof cur === 'string' ? cur : '';
}

function naverRankGetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NAVER_RANK_CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error(
      `시트를 찾을 수 없습니다: "${NAVER_RANK_CONFIG.SHEET_NAME}". `
      + 'NAVER_RANK_CONFIG.SHEET_NAME 값을 실제 시트 이름과 맞추세요.'
    );
  }
  return sheet;
}

/* ── 자동 실행(트리거) 설치 ───────────────────────────────────────────────── */

/** 매일 오전 11시 / 오후 5시(Asia/Seoul)에 naverRankUpdateAll이 자동 실행되도록 설치합니다. */
function naverRankInstall() {
  const removed = naverRankRemoveTriggers_();
  ScriptApp.newTrigger('naverRankUpdateAll').timeBased()
    .atHour(11).everyDays(1).inTimezone(NAVER_RANK_CONFIG.TIMEZONE).create();
  ScriptApp.newTrigger('naverRankUpdateAll').timeBased()
    .atHour(17).everyDays(1).inTimezone(NAVER_RANK_CONFIG.TIMEZONE).create();

  const triggers = naverRankListTriggers_();
  Logger.log('=== 네이버 쇼핑 순위 조회 설치 완료 ===');
  Logger.log(`기존 트리거 정리: ${removed}개 삭제`);
  Logger.log(`생성된 트리거 수: ${triggers.length}개 (매일 오전 11시 전후, 오후 5시 전후)`);
}

/** 등록된 트리거가 정확히 2개(11시/17시)인지 확인합니다. */
function naverRankVerifyTriggers() {
  const triggers = naverRankListTriggers_();
  Logger.log(`현재 등록된 naverRankUpdateAll 트리거: ${triggers.length}개`);
  if (triggers.length !== 2) {
    Logger.log('⚠️ 2개가 아닙니다. naverRankInstall을 다시 실행하세요.');
  }
}

/** 자동 실행을 멈춥니다. 시트에 이미 기록된 순위 이력은 그대로 남습니다. */
function naverRankUninstall() {
  const removed = naverRankRemoveTriggers_();
  Logger.log(`트리거 ${removed}개 삭제 완료`);
}

function naverRankListTriggers_() {
  return ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'naverRankUpdateAll');
}

function naverRankRemoveTriggers_() {
  const triggers = naverRankListTriggers_();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  return triggers.length;
}
