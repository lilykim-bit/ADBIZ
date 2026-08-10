/** ===================== CONFIG ===================== */
const CONFIG = {
  ACTIVITY_SHEET: '활동로그',
  CONTRACT_SHEET: '계약현황',
  PRODUCT_SHEET: '상품설정',
  STATE_SHEET: '알람_상태',
  MAPPING_SHEET: '담당자매핑',
  SLACK_BOT_TOKEN: PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'),
  CHANNEL_ID: 'C07G05CCGSX', // #wg-광고_담당자_biz
  REMIND_HOURS: 24,
  TEST_MODE: false,              // 2026-07-09: 테스트 완료, 실운영 전환
  TEST_SOURCE_OWNER: '우은수',    // TEST_MODE:false일 때는 사용되지 않음 (참고용으로 남겨둠)
  TEST_TARGET_SLACK_ID: 'U07RBU2TKNH', // TEST_MODE:false일 때는 사용되지 않음 (참고용으로 남겨둠)
};

const STATE_HEADERS = [
  '항목ID','유형','탭','활동ID','매장시퀀스','매장명','담당자',
  '최초알림일','마지막DM발송일','DM메시지TS','확인여부','확인시각',
  '실제완료여부','완료일'
];

/** 알람별 대상 광고 타입 (광고명에 이 키워드가 포함되어 있어야 매칭됨) */
const AD_TYPES = {
  SIGN_PAYMENT: ['예약핫플', '트렌드', '웨이팅'],
  UNITPRICE: ['예약핫플', '트렌드'],
  THEME_PERIOD: ['예약핫플']
};
/** 광고 종료 후에도 서명/결제 미완료를 계속 추적하는 유예기간(일) */
const GRACE_DAYS = 90;


/** DM 표시 순서 (처리 우선순위 기준). 건수 제한 없이 전체 항목을 모두 표시함 */
const TYPE_ORDER = ['서명미완료', '결제미완료', '객단가미입력', '기획전테마미입력', '리스팅기간미설정'];

/** ===================== 0) 초기 세팅 (최초 1회만 실행) ===================== */
function initAlarmSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.STATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.STATE_SHEET);
  sheet.getRange(1, 1, 1, STATE_HEADERS.length).setValues([STATE_HEADERS]);
  sheet.setFrozenRows(1);
}

/** ===================== 0-1) 담당자매핑 탭 자동 생성 (최초 1회 실행) ===================== */
function setupMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MAPPING_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['담당자명', 'SlackID']]);
    sheet.getRange(2, 1, 1, 2).setValues([['김승현', 'U07RBU2TKNH']]);
    sheet.setFrozenRows(1);
  }
}

/** ===================== 유틸 ===================== */
function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('컬럼을 찾을 수 없음: ' + name);
  return idx;
}

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없음: ' + sheetName);
  const data = sheet.getDataRange().getValues();
  return { headers: data[0], rows: data.slice(1) };
}

/** 날짜값 → Date 객체. Date 객체 / "2026.01.30 (금)" / "2026-01-30" 형식 처리. 파싱 불가·빈값이면 null */
function parseKDate(str) {
  if (!str) return null;
  if (Object.prototype.toString.call(str) === '[object Date]') {
    return isNaN(str.getTime()) ? null : str;
  }
  const m = String(str).match(/(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

/** 광고명에 지정된 키워드 중 하나라도 포함되어 있으면 true */
function matchesAdType(adName, allowedTypes) {
  if (!adName) return false;
  const name = String(adName);
  return allowedTypes.some(type => name.includes(type));
}

/** 활동ID → {start, end} 날짜 맵 (상품설정 탭 기준) */
function getProductDatesMap() {
  const { headers, rows } = getSheetData(CONFIG.PRODUCT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iStart = colIndex(headers, '광고시작일');
  const iEnd = colIndex(headers, '광고종료일');
  const map = {};
  rows.forEach(r => {
    map[r[iAct]] = { start: parseKDate(r[iStart]), end: parseKDate(r[iEnd]) };
  });
  return map;
}

/** ===================== 1) 스캔 & 상태시트 upsert ===================== */
function scanAndUpsert() {
  const items = [
    ...scanUnitPrice(),
    ...scanTheme(),
    ...scanAdPeriod(),
    ...scanSignature(),
    ...scanPayment()
  ];
  upsertStateSheet(items);
}

/** 객단가 미입력 — 예약핫플/트렌드, 광고시작일~종료일 */
function scanUnitPrice() {
  const { headers, rows } = getSheetData(CONFIG.ACTIVITY_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '활동_매장시퀀스');
  const iName = colIndex(headers, '활동_매장명');
  const iJoin = colIndex(headers, '활동_참여여부');
  const iAdName = colIndex(headers, '활동_광고명-룩업');
  const iLunch = colIndex(headers, '런치 객단가');
  const iDinner = colIndex(headers, '디너 객단가');
  const iOwner = colIndex(headers, '활동_담당자');

  const dateMap = getProductDatesMap();
  const today = new Date();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.UNITPRICE)) return false;
      if (r[iLunch] !== '' && r[iDinner] !== '') return false;
      const d = dateMap[r[iAct]];
      if (!d || !d.start || !d.end) return false;
      return today >= d.start && today <= d.end;
    })
    .map(r => ({
      itemId: 'UNITPRICE_' + r[iAct], type: '객단가미입력', tab: CONFIG.ACTIVITY_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 기획전 테마 미선택 — 예약핫플, 광고시작일-14일~시작일 */
function scanTheme() {
  const { headers, rows } = getSheetData(CONFIG.PRODUCT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iName = colIndex(headers, '매장명');
  const iJoin = colIndex(headers, '참여여부');
  const iAdName = colIndex(headers, '상품명-룩업');
  const iTheme = colIndex(headers, '기획전 테마');
  const iStart = colIndex(headers, '광고시작일');
  const iOwner = colIndex(headers, '계약_담당자');
  const today = new Date();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여' || r[iTheme] !== '') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.THEME_PERIOD)) return false;
      const start = parseKDate(r[iStart]);
      if (!start) return false;
      const windowStart = addDays(start, -14);
      return today >= windowStart && today < start;
    })
    .map(r => ({
      itemId: 'THEME_' + r[iAct], type: '기획전테마미입력', tab: CONFIG.PRODUCT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 리스팅 기간 미설정 (구 '광고기간미설정') — 예약핫플, 광고시작일-14일~시작일 (시작일 자체가 없으면 즉시 알람) */
function scanAdPeriod() {
  const { headers, rows } = getSheetData(CONFIG.PRODUCT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iName = colIndex(headers, '매장명');
  const iJoin = colIndex(headers, '참여여부');
  const iAdName = colIndex(headers, '상품명-룩업');
  const iStart = colIndex(headers, '광고시작일');
  const iEnd = colIndex(headers, '광고종료일');
  const iOwner = colIndex(headers, '계약_담당자');
  const today = new Date();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.THEME_PERIOD)) return false;

      const rawStart = r[iStart];
      const rawEnd = r[iEnd];
      const isEmpty = v => v === '' || v === '#N/A' || v === null;
      if (!isEmpty(rawStart) && !isEmpty(rawEnd)) return false;

      const start = parseKDate(rawStart);
      if (!start) return true; // 시작일 자체가 없음 → 즉시 알람 대상

      const windowStart = addDays(start, -14);
      return today >= windowStart && today < start;
    })
    .map(r => ({
      itemId: 'ADPERIOD_' + r[iAct], type: '리스팅기간미설정', tab: CONFIG.PRODUCT_SHEET,  // itemId 접두사는 기존 상태시트 행과의 연속성을 위해 ADPERIOD_ 유지
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 계약서 서명 미완료 — 예약핫플/트렌드/웨이팅, 시작일-14일~종료일 */
function scanSignature() {
  const { headers, rows } = getSheetData(CONFIG.CONTRACT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '계약_매장시퀀스');
  const iName = colIndex(headers, '계약_매장명');
  const iJoin = colIndex(headers, '계약_참여여부');
  const iAdName = colIndex(headers, '계약_광고명');
  const iStatus = colIndex(headers, '계약상태');
  const iStart = colIndex(headers, '시작일');
  const iEnd = colIndex(headers, '종료일');
  const iOwner = colIndex(headers, '계약_담당자');
  const today = new Date();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여' || r[iStatus] === '계약서 서명 완료') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.SIGN_PAYMENT)) return false;
      const start = parseKDate(r[iStart]);
      const end = parseKDate(r[iEnd]);
      if (!start || !end) return false;
      const windowStart = addDays(start, -14);
      return today >= windowStart && today <= addDays(end, GRACE_DAYS);
    })
    .map(r => ({
      itemId: 'SIGN_' + r[iAct], type: '서명미완료', tab: CONFIG.CONTRACT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/**
 * 결제 미완료 — 예약핫플/트렌드/웨이팅, 시작일-14일~종료일
 * '결제일' 컬럼: 2026-07-09 확정 계약현황 스키마 기준 정식 헤더명(공란 = 결제 미완료).
 */
function scanPayment() {
  const { headers, rows } = getSheetData(CONFIG.CONTRACT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '계약_매장시퀀스');
  const iName = colIndex(headers, '계약_매장명');
  const iJoin = colIndex(headers, '계약_참여여부');
  const iAdName = colIndex(headers, '계약_광고명'); // ← 원본 누락 버그: 미선언 시 ReferenceError
  const iPayDate = colIndex(headers, '결제일');
  const iStart = colIndex(headers, '시작일');
  const iEnd = colIndex(headers, '종료일');
  const iOwner = colIndex(headers, '계약_담당자');
  const today = new Date();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여' || r[iPayDate] !== '') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.SIGN_PAYMENT)) return false;
      const start = parseKDate(r[iStart]);
      const end = parseKDate(r[iEnd]);
      if (!start || !end) return false;
      const windowStart = addDays(start, -14);
      return today >= windowStart && today <= addDays(end, GRACE_DAYS);
    })
    .map(r => ({
      itemId: 'PAYMENT_' + r[iAct], type: '결제미완료', tab: CONFIG.CONTRACT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 상태시트 upsert: 신규 항목 추가 + 미러링에서 사라진 항목은 '취소추정' 처리 */
function upsertStateSheet(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.STATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.STATE_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, STATE_HEADERS.length).setValues([STATE_HEADERS]);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const idIdx = colIndex(headers, '항목ID');
  const existingIds = new Set(rows.map(r => r[idIdx]));
  const now = new Date();

  // 신규 항목 추가
  const newRows = items
    .filter(item => !existingIds.has(item.itemId))
    .map(item => {
      const row = new Array(headers.length).fill('');
      row[colIndex(headers,'항목ID')] = item.itemId;
      row[colIndex(headers,'유형')] = item.type;
      row[colIndex(headers,'탭')] = item.tab;
      row[colIndex(headers,'활동ID')] = item.actId;
      row[colIndex(headers,'매장시퀀스')] = item.seq;
      row[colIndex(headers,'매장명')] = item.storeName;
      row[colIndex(headers,'담당자')] = item.owner;
      row[colIndex(headers,'최초알림일')] = now;
      row[colIndex(headers,'확인여부')] = 'X';
      row[colIndex(headers,'실제완료여부')] = 'X';
      return row;
    });
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, headers.length).setValues(newRows);
  }

  // 이번 스캔에 안 잡힌(=사라진) 미완료 항목 → 취소추정 처리
  const currentIds = new Set(items.map(i => i.itemId));
  const completeIdx = colIndex(headers,'실제완료여부');
  const completeDateIdx = colIndex(headers,'완료일');
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    if (r[completeIdx] === 'X' && !currentIds.has(r[idIdx])) {
      sheet.getRange(rowNum, completeIdx+1).setValue('취소추정');
      sheet.getRange(rowNum, completeDateIdx+1).setValue(now);
    } else if (r[completeIdx] === '취소추정' && currentIds.has(r[idIdx])) {
      // 스캔 윈도우 확장 등으로 다시 잡힌 항목 → 미완료로 복구
      sheet.getRange(rowNum, completeIdx+1).setValue('X');
      sheet.getRange(rowNum, completeDateIdx+1).setValue('');
    }
  });
}

/** ===================== 2) 담당자별 DM 발송 ===================== */
function sendDMs() {
  // 평일(월~금)에만 발송. 트리거는 매일 걸려있지만 주말엔 여기서 조용히 스킵.
  const dow = new Date().getDay(); // 0=일, 6=토
  if (dow === 0 || dow === 6) { Logger.log('주말이라 발송 스킵'); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const iConfirm = colIndex(headers,'확인여부');
  const iComplete = colIndex(headers,'실제완료여부');
  const iOwner = colIndex(headers,'담당자');
  const iLastDM = colIndex(headers,'마지막DM발송일');
  const iType = colIndex(headers,'유형');
  const iStore = colIndex(headers,'매장명');
  const iSeq = colIndex(headers,'매장시퀀스');
  const iTs = colIndex(headers,'DM메시지TS');

  const now = new Date();
  const remindMs = CONFIG.REMIND_HOURS * 60 * 60 * 1000;
  const groups = {};

  rows.forEach((r, i) => {
    if (r[iConfirm] === 'X' && r[iComplete] === 'X') {
      const owner = r[iOwner];

      // 테스트 모드면 지정한 출처 담당자(홍성혁 등) 항목만 대상으로
      if (CONFIG.TEST_MODE && owner !== CONFIG.TEST_SOURCE_OWNER) return;

      const last = r[iLastDM];
      const needSend = !last || (now - new Date(last)) > remindMs;
      if (needSend) {
        (groups[owner] = groups[owner] || []).push({ rowIndex: i + 2, type: r[iType], store: r[iStore], seq: r[iSeq] });
      }
    }
  });

  const mapping = getOwnerSlackMap();
  Object.keys(groups).forEach(owner => {
    const slackId = CONFIG.TEST_MODE ? CONFIG.TEST_TARGET_SLACK_ID : mapping[owner];
    if (!slackId) { Logger.log('Slack ID 매핑 없음: ' + owner); return; }
    const list = groups[owner];
    const ts = postSlackDM(slackId, buildDMBlocks(owner, list), buildDMFallbackText(owner, list));
    if (ts) {
      list.forEach(item => {
        sheet.getRange(item.rowIndex, iLastDM+1).setValue(now);
        sheet.getRange(item.rowIndex, iTs+1).setValue(ts);
      });
    }
  });

  sendChannelBroadcast();
}

/**
 * 유형별로 그룹핑 + Block Kit 섹션으로 DM 구성.
 * - 유형별 소제목("서명미완료 N건") 아래에 매장 전체 리스트 (TYPE_ORDER 순서로 정렬, 심각한 유형 먼저)
 * - 건수 제한/요약 없음 — 몇 건이든 전부 그대로 노출 (숨기지 않음)
 * - 매장 항목은 "매장시퀀스  매장명" 형태로 표시해서 바로 검색/식별 가능하게 함
 */
function buildDMBlocks(owner, list) {
  const byType = {};
  list.forEach(item => { (byType[item.type] = byType[item.type] || []).push(item); });

  const orderedTypes = Object.keys(byType).sort((a, b) => {
    const ai = TYPE_ORDER.indexOf(a), bi = TYPE_ORDER.indexOf(b);
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });

  const blocks = [];
  blocks.push({
    type: 'header',
    text: { type: 'plain_text', text: `${owner}님, 처리 필요한 항목 ${list.length}건`, emoji: true }
  });
  blocks.push({ type: 'divider' });

  orderedTypes.forEach(type => {
    const items = byType[type];
    const body = items.map(it => `•  ${it.seq}  ${it.store}`).join('\n');
    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: `*${type}  ${items.length}건*\n${body}` }
    });
  });

  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'context',
    elements: [{ type: 'mrkdwn', text: '✅ 확인하셨으면 이 메시지에 리액션 하나만 눌러주세요 (이모지 종류 무관) 🙏' }]
  });

  return blocks;
}

/** 알림 미리보기/스크린리더 등 fallback용 평문 텍스트 (Block Kit과 함께 text 파라미터로 전송) */
function buildDMFallbackText(owner, list) {
  return `${owner}님, 처리 필요한 항목 ${list.length}건이 있어요. DM에서 유형별로 확인해주세요.`;
}

function getOwnerSlackMap() {
  const { headers, rows } = getSheetData(CONFIG.MAPPING_SHEET);
  const iName = colIndex(headers, '담당자명');
  const iSlack = colIndex(headers, 'SlackID');
  const map = {};
  rows.forEach(r => { map[r[iName]] = r[iSlack]; });
  return map;
}

function postSlackDM(slackUserId, blocks, fallbackText) {
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: slackUserId, text: fallbackText, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) { Logger.log('Slack 발송 실패: ' + res.error); return null; }
  return res.ts;
}

/** ===================== 3) 리액션 폴링 (확인 체크) ===================== */
function pollReactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const iConfirm = colIndex(headers,'확인여부');
  const iConfirmTime = colIndex(headers,'확인시각');
  const iTs = colIndex(headers,'DM메시지TS');
  const iOwner = colIndex(headers,'담당자');
  const mapping = getOwnerSlackMap();
  const checked = new Set();

  rows.forEach((r, i) => {
    const ts = r[iTs];
    if (r[iConfirm] === 'X' && ts) {
      const slackId = CONFIG.TEST_MODE ? CONFIG.TEST_TARGET_SLACK_ID : mapping[r[iOwner]];
      if (!slackId) return;
      const key = slackId + '_' + ts;
      if (checked.has(key)) return;
      checked.add(key);

      if (checkSlackReaction(slackId, ts)) {
        rows.forEach((r2, j) => {
          if (r2[iTs] === ts) {
            const rn = j + 2;
            sheet.getRange(rn, iConfirm+1).setValue('O');
            sheet.getRange(rn, iConfirmTime+1).setValue(new Date());
          }
        });
      }
    }
  });
}

function checkSlackReaction(channel, ts) {
  const url = `https://slack.com/api/reactions.get?channel=${channel}&timestamp=${ts}`;
  const res = JSON.parse(UrlFetchApp.fetch(url, {
    method: 'get', headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    muteHttpExceptions: true
  }).getContentText());
  // 메시지가 삭제됐거나 reactions 필드가 없을 수 있으므로 방어적으로 접근
  return !!(res.ok && res.message && Array.isArray(res.message.reactions) && res.message.reactions.length > 0);
}

/** ===================== 채널 브로드캐스트 (기획전테마/리스팅기간 집계) ===================== */
function sendChannelBroadcast() {
  if (CONFIG.TEST_MODE) { Logger.log('테스트 모드 - 채널 발송 스킵'); return; }

  const dow = new Date().getDay();
  if (dow === 0 || dow === 6) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const iType = colIndex(headers, '유형');
  const iOwner = colIndex(headers, '담당자');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iStore = colIndex(headers, '매장명');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const ownerSlackMap = getOwnerSlackMap();

  const CHANNEL_TYPES = ['기획전테마미입력', '리스팅기간미설정'];
  const byType = {};

  rows.forEach(function (r) {
    const type = r[iType];
    if (CHANNEL_TYPES.indexOf(type) === -1) return;
    if (r[iConfirm] === 'X' && r[iComplete] === 'X') {
      if (!byType[type]) byType[type] = [];
      byType[type].push({ owner: r[iOwner], seq: r[iSeq], store: r[iStore] });
    }
  });

  let totalCount = 0;
  CHANNEL_TYPES.forEach(function (t) { totalCount += byType[t] ? byType[t].length : 0; });
  if (totalCount === 0) { Logger.log('채널 발송 대상 없음'); return; }

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '처리 필요한 항목 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  CHANNEL_TYPES.forEach(function (type) {
    const items = byType[type];
    if (!items || !items.length) return;
    const body = items.map(function (it) {
      const ownerTag = ownerSlackMap[it.owner] ? ('<@' + ownerSlackMap[it.owner] + '>') : it.owner;
      return '•  ' + it.seq + '  ' + it.store + '  ' + ownerTag;
    }).join('\n');
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '*' + type + '  ' + items.length + '건*\n' + body } });
  });

  blocks.push({ type: 'divider' });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '담당자분들은 개인 DM으로도 동일 항목을 받으셨어요. 확인 후 DM에 리액션 부탁드려요 🙏' }] });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: 'cc. <@U0B2CNXM7LN>' }] });

  const fallback = '처리 필요한 항목 ' + totalCount + '건 (기획전테마미입력/리스팅기간미설정)';
  postSlackChannelMessage(blocks, fallback);
}

function postSlackChannelMessage(blocks, fallbackText) {
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: CONFIG.CHANNEL_ID, text: fallbackText, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) { Logger.log('채널 발송 실패: ' + res.error); return null; }
  return res.ts;
}

function testChannelBroadcastToMe() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const iType = colIndex(headers, '유형');
  const iOwner = colIndex(headers, '담당자');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iStore = colIndex(headers, '매장명');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');

  const CHANNEL_TYPES = ['기획전테마미입력', '리스팅기간미설정'];
  const byType = {};

  rows.forEach(function (r) {
    const type = r[iType];
    if (CHANNEL_TYPES.indexOf(type) === -1) return;
    if (r[iConfirm] === 'X' && r[iComplete] === 'X') {
      if (!byType[type]) byType[type] = [];
      byType[type].push({ owner: r[iOwner], seq: r[iSeq], store: r[iStore] });
    }
  });

  let totalCount = 0;
  CHANNEL_TYPES.forEach(function (t) { totalCount += byType[t] ? byType[t].length : 0; });
  if (totalCount === 0) { Logger.log('테스트: 대상 없음'); return; }

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '[테스트] 처리 필요한 항목 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  CHANNEL_TYPES.forEach(function (type) {
    const items = byType[type];
    if (!items || !items.length) return;
    const body = items.map(function (it) { return '•  ' + it.seq + '  ' + it.store + '  (' + it.owner + ')'; }).join('\n');
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '*' + type + '  ' + items.length + '건*\n' + body } });
  });

  blocks.push({ type: 'divider' });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '(테스트 발송 - 실제 채널에는 안 나갔어요)' }] });

  const fallback = '[테스트] 처리 필요한 항목 ' + totalCount + '건';
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: CONFIG.TEST_TARGET_SLACK_ID, text: fallback, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) Logger.log('테스트 발송 실패: ' + res.error);
  else Logger.log('테스트 발송 성공');
}

function testChannelBroadcastMockToMe() {
  const byType = {
    '기획전테마미입력': [
      { owner: '홍성혁', seq: '43073', store: '앙젤리제 시당/방배' },
      { owner: '우은수', seq: '73970', store: '용과서울' }
    ],
    '리스팅기간미설정': [
      { owner: '이도은', seq: '51204', store: '스시하나레' }
    ]
  };

  const CHANNEL_TYPES = ['기획전테마미입력', '리스팅기간미설정'];
  let totalCount = 0;
  CHANNEL_TYPES.forEach(function (t) { totalCount += byType[t] ? byType[t].length : 0; });

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '[예시] 처리 필요한 항목 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  CHANNEL_TYPES.forEach(function (type) {
    const items = byType[type];
    if (!items || !items.length) return;
    const body = items.map(function (it) { return '•  ' + it.seq + '  ' + it.store + '  (' + it.owner + ')'; }).join('\n');
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '*' + type + '  ' + items.length + '건*\n' + body } });
  });

  blocks.push({ type: 'divider' });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '(예시 발송 - 가상 데이터입니다)' }] });

  const fallback = '[예시] 처리 필요한 항목 ' + totalCount + '건';
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: CONFIG.TEST_TARGET_SLACK_ID, text: fallback, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) Logger.log('예시 발송 실패: ' + res.error);
  else Logger.log('예시 발송 성공');
}

/** ===================== 트리거 설정 (최초 1회만 실행) ===================== */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scanAndUpsert').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendDMs').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('pollReactions').timeBased().everyHours(1).create();
}

function fixSendDMsTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDMs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDMs').timeBased().everyDays(1).atHour(9).create();
  Logger.log('sendDMs 트리거를 일 단위(오전 9시) 타이머로 재설정했습니다.');
}

function tempCheckSlackId() {
  const map = getOwnerSlackMap();
  Logger.log(JSON.stringify(map));
}

function tempDiagnoseState() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const iType = colIndex(headers, '유형');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const typeCounts = {};
  const targetTypes = ['기획전테마미입력', '리스팅기간미설정'];
  let targetRows = [];
  rows.forEach(function (r) {
    const t = r[iType];
    typeCounts[t] = (typeCounts[t] || 0) + 1;
    if (targetTypes.indexOf(t) !== -1) {
      targetRows.push({ confirm: JSON.stringify(r[iConfirm]), complete: JSON.stringify(r[iComplete]) });
    }
  });
  Logger.log('전체 행수: ' + rows.length);
  Logger.log('유형별 개수: ' + JSON.stringify(typeCounts));
  Logger.log('타겟유형 행 상세: ' + JSON.stringify(targetRows));
}

function sendComplianceListToMe() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const iType = colIndex(headers, '유형');
  const iOwner = colIndex(headers, '담당자');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iStore = colIndex(headers, '매장명');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const ownerSlackMap = getOwnerSlackMap();

  const TYPES = ['서명미완료', '결제미완료'];
  const byType = {};
  rows.forEach(function (r) {
    const type = r[iType];
    if (TYPES.indexOf(type) === -1) return;
    if (r[iConfirm] === 'X' && r[iComplete] === 'X') {
      if (!byType[type]) byType[type] = [];
      byType[type].push({ owner: r[iOwner], seq: r[iSeq], store: r[iStore] });
    }
  });

  let totalCount = 0;
  TYPES.forEach(function (t) { totalCount += byType[t] ? byType[t].length : 0; });
  if (totalCount === 0) { Logger.log('테스트: 대상 없음'); return; }

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '[테스트] 계약서명/결제 미완료 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  const CHUNK = 20;
  TYPES.forEach(function (type) {
    const items = byType[type];
    if (!items || !items.length) return;
    for (let i = 0; i < items.length; i += CHUNK) {
      const part = items.slice(i, i + CHUNK);
      const body = part.map(function (it) {
        const ownerTag = ownerSlackMap[it.owner] ? ('<@' + ownerSlackMap[it.owner] + '>') : it.owner;
        return '•  ' + it.seq + '  ' + it.store + '  ' + ownerTag;
      }).join('\n');
      const title = (i === 0) ? ('*' + type + '  ' + items.length + '건*\n') : '';
      blocks.push({ type: 'section', text: { type: 'mrkdwn', text: title + body } });
    }
  });

  blocks.push({ type: 'divider' });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: 'cc. <@U0B2CNXM7LN>' }] });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '(테스트 발송 - 실제 조건과 동일, 수신자만 승현님으로 전환)' }] });

  const fallback = '[테스트] 계약서명/결제 미완료 ' + totalCount + '건';
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: CONFIG.TEST_TARGET_SLACK_ID, text: fallback, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) Logger.log('테스트 발송 실패: ' + res.error);
  else Logger.log('테스트 발송 성공, 총 ' + totalCount + '건');
}

function sendComplianceListByRoundToMe() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const iType = colIndex(headers, '유형');
  const iActivity = colIndex(headers, '활동ID');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');

  const TYPES = ['서명미완료', '결제미완료'];

  const pData = getSheetData(CONFIG.PRODUCT_SHEET);
  const pHeaders = pData.headers;
  const pRows = pData.rows;
  const pAct = colIndex(pHeaders, '활동ID');
  const pName = colIndex(pHeaders, '상품명-룩업');
  const pStart = colIndex(pHeaders, '광고시작일');
  const pEnd = colIndex(pHeaders, '광고종료일');
  const productMap = {};
  pRows.forEach(function (r) {
    productMap[r[pAct]] = { name: r[pName], start: parseKDate(r[pStart]), end: parseKDate(r[pEnd]) };
  });

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const byRound = {};
  let unmatched = 0;
  rows.forEach(function (r) {
    const type = r[iType];
    if (TYPES.indexOf(type) === -1) return;
    if (!(r[iConfirm] === 'X' && r[iComplete] === 'X')) return;

    const actId = r[iActivity];
    const product = productMap[actId];
    if (!product || !product.start || !product.end) { unmatched++; return; }

    const isTargetProduct = AD_TYPES.UNITPRICE.some(function (k) { return (product.name || '').indexOf(k) !== -1; });
    if (!isTargetProduct) return;

    const isLive = product.start <= today && today <= product.end;
    if (!isLive) return;

    const rawName = product.name || '(차수명 없음)';
    const roundName = rawName.split('_')[0].trim();
    if (!byRound[roundName]) { byRound[roundName] = { 서명미완료: 0, 결제미완료: 0, start: product.start, end: product.end }; }
    else { if (product.start < byRound[roundName].start) { byRound[roundName].start = product.start; } if (product.end > byRound[roundName].end) { byRound[roundName].end = product.end; } }
    byRound[roundName][type]++;
  });

  const roundNames = Object.keys(byRound);
  if (roundNames.length === 0) {
    Logger.log('라이브 중인 차수 대상 없음 (매칭 안 된 건: ' + unmatched + ')');
    return;
  }

  roundNames.sort(function (a, b) { return byRound[a].start - byRound[b].start; });

  let totalCount = 0;
  roundNames.forEach(function (name) { totalCount += byRound[name].서명미완료 + byRound[name].결제미완료; });

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '[테스트] 라이브 차수별 서명/결제 미완료 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  roundNames.forEach(function (name) {
    const g = byRound[name];
    const dateStr = Utilities.formatDate(g.start, 'GMT+9', 'MM/dd') + '~' + Utilities.formatDate(g.end, 'GMT+9', 'MM/dd');
    const lines = [];
    lines.push('*' + name + '* (' + dateStr + ')');
    if (g.결제미완료 > 0) lines.push('결제 미완료 ' + g.결제미완료 + '건');
    if (g.서명미완료 > 0) lines.push('계약서 미서명 ' + g.서명미완료 + '건');
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: lines.join('\n') } });
  });

  blocks.push({ type: 'divider' });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: 'cc. <@U0B2CNXM7LN>' }] });
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '(테스트 발송 - 라이브 중인 광고 차수만 집계, 수신자만 승현님으로 전환)' }] });

  const fallback = '[테스트] 라이브 차수별 서명/결제 미완료 ' + totalCount + '건';
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: CONFIG.TEST_TARGET_SLACK_ID, text: fallback, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) Logger.log('테스트 발송 실패: ' + res.error);
  else Logger.log('테스트 발송 성공, 총 ' + totalCount + '건, 차수 ' + roundNames.length + '개 (매칭 안 된 건: ' + unmatched + ')');
}

function sendComplianceListByRoundToChannel(channelOverride) {
  // 트리거 실행 시에는 이벤트 객체가 들어오므로 문자열일 때만 채널 오버라이드로 취급
  const channel = (typeof channelOverride === 'string' && channelOverride) ? channelOverride : CONFIG.CHANNEL_ID;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const iType = colIndex(headers, '유형');
  const iActivity = colIndex(headers, '활동ID');
  const iComplete = colIndex(headers, '실제완료여부');

  const TYPES = ['서명미완료', '결제미완료'];

  // 차수명/기간 소스 = 계약현황 (서명·결제 판단과 동일 기준)
  const cData = getSheetData(CONFIG.CONTRACT_SHEET);
  const cAct = colIndex(cData.headers, '활동ID');
  const cName = colIndex(cData.headers, '계약_광고명');
  const cStart = colIndex(cData.headers, '시작일');
  const cEnd = colIndex(cData.headers, '종료일');
  const contractMap = {};
  cData.rows.forEach(function (r) {
    if (!r[cAct]) return;
    contractMap[r[cAct]] = { name: String(r[cName] || ''), start: parseKDate(r[cStart]), end: parseKDate(r[cEnd]) };
  });

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const byRound = {};
  let unmatched = 0;

  rows.forEach(function (r) {
    const type = r[iType];
    if (TYPES.indexOf(type) === -1) return;
    if (r[iComplete] !== 'X') return; // 확인(리액션) 여부와 무관하게 미완료면 집계

    const c = contractMap[r[iActivity]];
    if (!c || !c.name || !c.start || !c.end) { unmatched++; return; }
    if (!matchesAdType(c.name, AD_TYPES.UNITPRICE)) return; // 예약핫플·트렌드만

    const roundName = c.name.split('_')[0].trim();
    if (!byRound[roundName]) byRound[roundName] = { 서명미완료: 0, 결제미완료: 0, starts: [], ends: [] };
    byRound[roundName][type]++;
    byRound[roundName].starts.push(c.start.getTime());
    byRound[roundName].ends.push(c.end.getTime());
  });

  // 최빈값 기준 — 잘못 입력된 한 건이 차수 기간을 끌고 가지 않도록
  function modeOf(arr) {
    const cnt = {}; let best = arr[0]; let bestN = 0;
    arr.forEach(function (v) { cnt[v] = (cnt[v] || 0) + 1; if (cnt[v] > bestN) { bestN = cnt[v]; best = v; } });
    return best;
  }

  const roundList = Object.keys(byRound).map(function (n) {
    const g = byRound[n];
    g.name = n;
    g.start = new Date(modeOf(g.starts));
    g.end = new Date(modeOf(g.ends));
    // 최빈 기간에서 30일 이상 벗어난 건만 '입력 오류 의심'으로 표시 (매장별 기간 차이는 정상)
    const TOL = 30 * 24 * 60 * 60 * 1000;
    g.odd = g.starts.filter(function (v) { return Math.abs(v - g.start.getTime()) > TOL; }).length
          + g.ends.filter(function (v) { return Math.abs(v - g.end.getTime()) > TOL; }).length;
    return g;
  });

  // 라이브 차수 + 광고타입별 직전 1개 종료 차수
  // 진행 중 + 예정(시작 전) 차수 — 시작 전이 오히려 서명/결제를 챙겨야 하는 시점
  const live = roundList.filter(function (g) { return today <= g.end; });
  const ended = roundList.filter(function (g) { return g.end < today; });
  const picked = live.slice();
  AD_TYPES.UNITPRICE.forEach(function (k) {
    const cand = ended.filter(function (g) { return g.name.indexOf(k) !== -1; })
                      .sort(function (a, b) { return b.end - a.end; });
    if (cand.length > 0) picked.push(cand[0]);
  });

  if (picked.length === 0) { Logger.log('대상 차수 없음 / 미매칭 ' + unmatched + '건'); return; }
  picked.sort(function (a, b) { return a.start - b.start; });

  let totalCount = 0;
  picked.forEach(function (g) { totalCount += g.서명미완료 + g.결제미완료; });

  const fmt = function (d) { return Utilities.formatDate(d, 'GMT+9', 'yy.MM.dd'); };

  const blocks = [];
  blocks.push({ type: 'header', text: { type: 'plain_text', text: '차수별 서명/결제 미완료 ' + totalCount + '건', emoji: true } });
  blocks.push({ type: 'divider' });

  picked.forEach(function (g) {
    const closed = (g.end < today) ? ' _(종료 차수)_' : (g.start > today ? ' _(시작 예정)_' : '');
    const warn = (g.odd > 0) ? '  :warning: 계약기간 입력 확인 필요 ' + g.odd + '건' : '';
    const l = [];
    l.push('*' + g.name + '*  (' + fmt(g.start) + '~' + fmt(g.end) + ')' + closed + warn);
    if (g.결제미완료 > 0) l.push('· 결제 미완료 *' + g.결제미완료 + '건*');
    if (g.서명미완료 > 0) l.push('· 계약서 미서명 *' + g.서명미완료 + '건*');
    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: l.join('\n') } });
  });

  blocks.push({ type: 'divider' });
  let ctx = 'cc. <!subteam^S08FRBT08QN> <@U0B2CNXM7LN>';
  if (unmatched > 0) ctx += '  |  계약 날짜 미확인으로 차수 미분류 ' + unmatched + '건';
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: ctx }] });

  const fallback = '차수별 서명/결제 미완료 ' + totalCount + '건';
  const res = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
    payload: JSON.stringify({ channel: channel, text: fallback, blocks: blocks }),
    muteHttpExceptions: true
  }).getContentText());
  if (!res.ok) Logger.log('발송 실패: ' + res.error);
  else Logger.log('발송 완료 / 차수 ' + picked.length + '개 / 총 ' + totalCount + '건 / 미매칭 ' + unmatched + '건');
}

/** 테스트 채널로만 결과 미리보기 발송 */
function testComplianceListByRoundToTestChannel() {
  sendComplianceListByRoundToChannel('C09SQQ000H2');
}

function tempCheckChannelAccess() {
  const auth = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post', headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN }, muteHttpExceptions: true
  }).getContentText());
  Logger.log('auth.test: ' + JSON.stringify(auth));

  const info = JSON.parse(UrlFetchApp.fetch('https://slack.com/api/conversations.info?channel=' + CONFIG.CHANNEL_ID, {
    method: 'get', headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN }, muteHttpExceptions: true
  }).getContentText());
  Logger.log('conversations.info: ' + JSON.stringify(info));
}

function setupComplianceChannelTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendComplianceListByRoundToChannel') { ScriptApp.deleteTrigger(t); }
  });
  const days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
  days.forEach(function (d) {
    ScriptApp.newTrigger('sendComplianceListByRoundToChannel').timeBased().onWeekDay(d).atHour(9).create();
  });
  Logger.log('평일(월~금) 오전 9시 트리거 5개 생성 완료');
}

function tempDiagnoseSendDMs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const iOwner = colIndex(headers, '담당자');
  const iLastDM = colIndex(headers, '마지막DM발송일');
  const now = new Date();
  const remindMs = CONFIG.REMIND_HOURS * 60 * 60 * 1000;
  const groups = {};
  let totalMatch = 0;
  rows.forEach(function (r) {
    if (r[iConfirm] === 'X' && r[iComplete] === 'X') {
      totalMatch++;
      const owner = r[iOwner];
      const last = r[iLastDM];
      const needSend = !last || (now - new Date(last)) > remindMs;
      if (!groups[owner]) groups[owner] = { total: 0, needSend: 0, lastSample: null };
      groups[owner].total++;
      if (needSend) groups[owner].needSend++;
      if (!groups[owner].lastSample) groups[owner].lastSample = last ? new Date(last).toString() : '(never)';
    }
  });
  const mapping = getOwnerSlackMap();
  const report = {};
  Object.keys(groups).forEach(function (owner) {
    report[owner] = { total: groups[owner].total, needSend: groups[owner].needSend, lastDMSample: groups[owner].lastSample, hasSlackMapping: !!mapping[owner] };
  });
  Logger.log('TEST_MODE: ' + CONFIG.TEST_MODE);
  Logger.log('요일(0=일,6=토): ' + new Date().getDay());
  Logger.log('전체 대상 행수: ' + totalMatch);
  Logger.log('담당자별 현황: ' + JSON.stringify(report));
}

/** ===================== 외부 호출용 Slack 메시지 발송 웹앱 (catchweek-lms-send 등에서 사용) ===================== */
// 웹훅 공유 시크릿은 코드에 하드코딩하지 않고 스크립트 속성(WEBHOOK_SHARED_SECRET)에서 읽는다.
const WEBHOOK_SHARED_SECRET = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SHARED_SECRET');
const WEBHOOK_ALLOWED_CHANNELS = ['C08731XUE95', CONFIG.CHANNEL_ID, 'C09SQQ000H2', 'C095PRYDE12', 'C09C2DWQXB8']; // C09SQQ000H2 = 테스트 전용 채널

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!WEBHOOK_SHARED_SECRET || body.secret !== WEBHOOK_SHARED_SECRET) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
    }
    const channel = body.channel;
    const text = body.message;
    if (!text || !channel) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'missing message or channel' })).setMimeType(ContentService.MimeType.JSON);
    }
    if (WEBHOOK_ALLOWED_CHANNELS.indexOf(channel) === -1) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'channel not allowed' })).setMimeType(ContentService.MimeType.JSON);
    }
    const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      headers: { Authorization: "Bearer " + CONFIG.SLACK_BOT_TOKEN, "Content-Type": "application/json; charset=utf-8" },
      payload: JSON.stringify({ channel: channel, thread_ts: (body.thread_ts || undefined), text: text }),
      muteHttpExceptions: true
    });
    return ContentService.createTextOutput(res.getContentText()).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
