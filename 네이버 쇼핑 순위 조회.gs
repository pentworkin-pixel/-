/**
 * 네이버 쇼핑 모바일 검색(msearch.shopping.naver.com) 결과에서
 * 광고 상품을 제외한 순위를 조회해 스프레드시트에 기록하는 Apps Script입니다.
 *
 * 이 파일의 모든 전역 이름은 naverRank / NAVER_RANK 접두사를 쓰므로
 * 기존 코드(광고 종료 캘린더 동기화 등)와 충돌하지 않습니다.
 *
 * ⚠️ 중요한 한계
 * msearch.shopping.naver.com은 데이터센터 IP(포함: Google Apps Script 서버)에서의
 * 접근을 비정상 트래픽으로 판단해 HTTP 418로 차단하는 경우가 있습니다.
 * 이 스크립트는 실행할 때마다 최신 접속 가능 여부를 그대로 반영합니다 — 차단되면
 * 결과 칸에 그 사실을 적어줄 뿐, 우회하지는 않습니다. naverRankDebugFetch로
 * 먼저 접속 상태를 확인하세요.
 */

var NAVER_RANK_CONFIG = {
  SHEET_NAME: '네이버쇼핑 순위조회',
  TIMEZONE: 'Asia/Seoul',
  HEADER_ROWS: 1,
  COL_KEYWORD: 1,          // A: 검색 키워드
  COL_MATCH_NAME: 2,       // B: 상품명 매칭 문자열 (부분 일치, 필수)
  COL_MATCH_MALL: 3,       // C: 판매처 매칭 문자열 (부분 일치, 선택)
  COL_RESULT_RANK: 4,      // D: 순위 (광고 제외)
  COL_RESULT_NAME: 5,      // E: 실제로 매칭된 상품명
  COL_RESULT_MALL: 6,      // F: 실제로 매칭된 판매처
  COL_RESULT_CHECKED_AT: 7, // G: 조회 시각
  COL_RESULT_NOTE: 8,      // H: 비고 (실패/미노출 사유)
  MAX_SCAN: 200,           // 한 번의 응답에서 순위 계산에 사용할 최대 상품 수
  REQUEST_DELAY_MS: 1200,  // 키워드 간 대기 시간 (과도한 요청으로 인한 차단 완화)
};

/** 시트의 모든 행을 순회하며 순위를 조회하고 결과를 기록합니다. */
function naverRankCheckAll() {
  const sheet = naverRankGetSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= NAVER_RANK_CONFIG.HEADER_ROWS) {
    Logger.log('조회할 행이 없습니다.');
    return;
  }

  const startRow = NAVER_RANK_CONFIG.HEADER_ROWS + 1;
  const numRows = lastRow - NAVER_RANK_CONFIG.HEADER_ROWS;
  const range = sheet.getRange(startRow, 1, numRows, NAVER_RANK_CONFIG.COL_RESULT_NOTE);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const keyword = String(values[i][NAVER_RANK_CONFIG.COL_KEYWORD - 1] || '').trim();
    const matchName = String(values[i][NAVER_RANK_CONFIG.COL_MATCH_NAME - 1] || '').trim();
    const matchMall = String(values[i][NAVER_RANK_CONFIG.COL_MATCH_MALL - 1] || '').trim();
    if (!keyword || !matchName) continue;

    const result = naverRankCheckOne_(keyword, matchName, matchMall);
    values[i][NAVER_RANK_CONFIG.COL_RESULT_RANK - 1] = result.rank;
    values[i][NAVER_RANK_CONFIG.COL_RESULT_NAME - 1] = result.name;
    values[i][NAVER_RANK_CONFIG.COL_RESULT_MALL - 1] = result.mallName;
    values[i][NAVER_RANK_CONFIG.COL_RESULT_CHECKED_AT - 1] = naverRankNow_();
    values[i][NAVER_RANK_CONFIG.COL_RESULT_NOTE - 1] = result.note;

    // 키워드마다 요청 사이에 텀을 둬 짧은 시간에 몰린 요청으로 보이지 않게 한다.
    if (i < values.length - 1) Utilities.sleep(NAVER_RANK_CONFIG.REQUEST_DELAY_MS);
  }

  range.setValues(values);
}

/** 시트 없이 키워드 하나만 바로 확인하고 싶을 때 스크립트 편집기에서 직접 실행합니다. */
function naverRankCheckKeyword(keyword, matchName, matchMall) {
  const result = naverRankCheckOne_(keyword, matchName, matchMall || '');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
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
    Logger.log(`${it.rank}위 | ${it.name} | ${it.mallName}`);
  });
  return fetched;
}

/** 키워드 하나에 대해 매칭되는 상품의 순위를 계산합니다. */
function naverRankCheckOne_(keyword, matchName, matchMall) {
  const fetched = naverRankFetchRanking_(keyword);
  if (!fetched.ok) {
    return { rank: '', name: '', mallName: '', note: fetched.note };
  }

  const items = fetched.items.slice(0, NAVER_RANK_CONFIG.MAX_SCAN);
  const found = naverRankFindMatch_(items, matchName, matchMall);
  if (!found) {
    return {
      rank: '', name: '', mallName: '',
      note: `미노출 (광고 제외 상위 ${items.length}개 안에서 못 찾음)`,
    };
  }
  return { rank: found.rank, name: found.name, mallName: found.mallName, note: '' };
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

/** 순위 목록에서 상품명(+선택적으로 판매처)이 매칭되는 첫 항목을 찾는다. */
function naverRankFindMatch_(items, matchName, matchMall) {
  const nameLower = matchName.toLowerCase();
  const mallLower = matchMall ? matchMall.toLowerCase() : '';

  for (const item of items) {
    if (!item.name || item.name.toLowerCase().indexOf(nameLower) === -1) continue;
    if (mallLower && (item.mallName || '').toLowerCase().indexOf(mallLower) === -1) continue;
    return item;
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

function naverRankNow_() {
  return Utilities.formatDate(new Date(), NAVER_RANK_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
