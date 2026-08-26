/**
 * 광고 알람 자동화 (개인 DM + 리액션 확인 추적 + 채널 요약)
 *
 * [중요] Apps Script는 프로젝트의 모든 .gs 파일이 하나의 전역 스코프로 합쳐진다.
 * 같은 프로젝트의 미수방지.gs(OverdueUnpaidDM.gs)는 OVERDUE_DM_ / overdue 접두어를
 * 쓰므로 이 파일의 식별자와 충돌하지 않는다. 전역 이름 추가 시 이 규칙을 지킬 것.
 */

/** ===================== CONFIG ===================== */
const CONFIG = {
  ACTIVITY_SHEET: '활동로그',
  CONTRACT_SHEET: '계약현황',
  PRODUCT_SHEET: '상품설정',
  STATE_SHEET: '알람_상태',
  MAPPING_SHEET: '담당자매핑',
  CHANNEL_ID: 'C07G05CCGSX',      // #wg-광고_담당자_biz
  TEST_CHANNEL_ID: 'C09SQQ000H2', // 테스트 전용 채널
  LEAD_SLACK_ID: 'U0B2CNXM7LN',   // 김지운 팀장
  CELL_GROUP_ID: 'S08FRBT08QN',   // 광고비즈니스셀 유저그룹
  REMIND_HOURS: 24,

  // 결제 미완료 중 '입금 마감일이 이미 지난 건'은 미수방지 봇(OverdueUnpaidDM.gs)이
  // 전담한다. true면 이 스크립트는 마감일 도래 전 건만 알람해서 중복 DM을 막는다.
  SKIP_OVERDUE_PAYMENT: true,

  TEST_MODE: false,                    // 2026-07-09: 테스트 완료, 실운영 전환
  TEST_SOURCE_OWNER: '우은수',          // TEST_MODE:true일 때만 사용
  TEST_TARGET_SLACK_ID: 'U07RBU2TKNH', // TEST_MODE:true일 때만 사용
};

/** 시크릿은 소스에 두지 않는다. 스크립트 속성(프로젝트 설정 > 스크립트 속성)에서 읽는다. */
function getSlackToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!t) throw new Error('스크립트 속성 SLACK_BOT_TOKEN 이 설정되지 않았습니다.');
  return t;
}

/**
 * 웹훅 시크릿 검증. 무중단 로테이션을 위해 두 개를 동시에 허용한다.
 *   WEBHOOK_SHARED_SECRET       : 현재(신규) 값
 *   WEBHOOK_SHARED_SECRET_PREV  : 직전 값 (호출 측 갱신이 끝나면 삭제)
 * 두 속성 모두 없으면 설정 누락이므로 명시적으로 실패시킨다.
 */
function isValidWebhookSecret_(provided) {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('WEBHOOK_SHARED_SECRET');
  const previous = props.getProperty('WEBHOOK_SHARED_SECRET_PREV');
  if (!current && !previous) {
    throw new Error('스크립트 속성 WEBHOOK_SHARED_SECRET 이 설정되지 않았습니다.');
  }
  if (!provided) return false;
  return provided === current || (!!previous && provided === previous);
}

const STATE_HEADERS = [
  '항목ID','유형','탭','활동ID','매장시퀀스','매장명','담당자',
  '최초알림일','마지막DM발송일','DM메시지TS','확인여부','확인시각',
  '실제완료여부','완료일','DM채널ID'
];

/** 알람별 대상 광고 타입 (광고명에 이 키워드가 포함되어 있어야 매칭됨) */
const AD_TYPES = {
  SIGN_PAYMENT: ['예약핫플', '트렌드', '웨이팅'],
  UNITPRICE: ['예약핫플', '트렌드'],
  THEME_PERIOD: ['예약핫플'],
  ROUND_REPORT: ['예약핫플', '트렌드']  // 차수별 채널 리포트 대상
};

/** 광고 종료 후에도 서명/결제 미완료를 계속 추적하는 유예기간(일) */
const GRACE_DAYS = 90;

/** DM 표시 순서 (처리 우선순위 기준). 건수 제한 없이 전체 항목을 모두 표시함 */
const TYPE_ORDER = ['서명미완료', '결제미완료', '객단가미입력', '기획전테마미입력', '리스팅기간미설정'];

/** 채널 요약에 올리는 유형 */
const CHANNEL_TYPES = ['기획전테마미입력', '리스팅기간미설정'];

/** ===================== 0) 초기 세팅 (최초 1회만 실행) ===================== */
function initAlarmSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.STATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.STATE_SHEET);
  ensureStateHeaders_(sheet);
}

/** 0-1) 담당자매핑 탭 자동 생성 (최초 1회 실행) */
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

/**
 * 알려진 담당자-SlackID 명부.
 * 미수방지 봇(OverdueUnpaidDM.gs)의 하드코딩 명부와 동일한 값이며,
 * addOwnerMappings()로 담당자매핑 탭에 채워 넣는다.
 */
const KNOWN_OWNER_SLACK_IDS = {
  '김나현': 'U0B99RC7H08', '김상하': 'U0AG600HV1U', '김연아': 'U0BN7S7C6TH', '김현수': 'U0BHGG7FP98',
  '남윤석': 'U0BQ3UL4Z5L', '남현욱': 'U07RH97HNQL', '신유빈': 'U09MXM4BV71', '우은수': 'U093FJ573FY',
  '이도은': 'U093FJ7DZ8W', '이세한': 'U09BZ6JL60G', '이승준': 'U09E3L1KFQR', '이조은': 'U09GZ0H7928',
  '이종익': 'U0AGLUM2G2V', '이하윤': 'U0AJ3LN8E3T', '이혜민': 'U0AJY0DSMPC', '전평정': 'U02QCTZT2PP',
  '최원영': 'U0BLAHC00G3', '한창완': 'U057M7S5RA9', '홍성혁': 'U02TN1U2PQR', '김이슬': 'U09ACJEAJSJ',
  '조완수': 'U027RCFP55W', '이지민': 'U0BGDNGHUBC', '김승현': 'U07RBU2TKNH'
};

/**
 * 담당자매핑 탭에 누락된 담당자를 채운다. 여러 번 실행해도 안전하다.
 *   - 이름이 없으면 새 행으로 추가
 *   - 이름은 있고 SlackID가 비어 있으면 채움
 *   - SlackID가 이미 다른 값이면 덮어쓰지 않고 충돌로 보고 (퇴사자 옛 ID 등 판단 필요)
 */
function addOwnerMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  if (!sheet) { setupMappingSheet(); sheet = ss.getSheetByName(CONFIG.MAPPING_SHEET); }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iName = colIndex(headers, '담당자명');
  const iSlack = colIndex(headers, 'SlackID');

  const existing = {};
  data.slice(1).forEach(function (r, idx) {
    const name = String(r[iName] == null ? '' : r[iName]).trim();
    if (name) existing[name] = { id: String(r[iSlack] == null ? '' : r[iSlack]).trim(), rowNum: idx + 2 };
  });

  const toAdd = [], toFill = [], conflicts = [], unchanged = [];
  Object.keys(KNOWN_OWNER_SLACK_IDS).forEach(function (name) {
    const id = KNOWN_OWNER_SLACK_IDS[name];
    const cur = existing[name];
    if (!cur) { toAdd.push([name, id]); return; }
    if (!cur.id) { toFill.push({ rowNum: cur.rowNum, name: name, id: id }); return; }
    if (cur.id !== id) { conflicts.push(name + ' (시트: ' + cur.id + ' / 명부: ' + id + ')'); return; }
    unchanged.push(name);
  });

  if (toAdd.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, iName + 1, toAdd.length, 1).setValues(toAdd.map(function (v) { return [v[0]]; }));
    sheet.getRange(startRow, iSlack + 1, toAdd.length, 1).setValues(toAdd.map(function (v) { return [v[1]]; }));
  }
  toFill.forEach(function (t) { sheet.getRange(t.rowNum, iSlack + 1).setValue(t.id); });

  Logger.log('추가 ' + toAdd.length + '명' + (toAdd.length ? ': ' + toAdd.map(function (v) { return v[0]; }).join(', ') : ''));
  Logger.log('빈 SlackID 채움 ' + toFill.length + '명' + (toFill.length ? ': ' + toFill.map(function (t) { return t.name; }).join(', ') : ''));
  Logger.log('이미 동일 ' + unchanged.length + '명');
  if (conflicts.length > 0) {
    Logger.log('⚠️ SlackID가 명부와 달라 건드리지 않은 담당자 ' + conflicts.length + '명 — 확인 필요:');
    conflicts.forEach(function (c) { Logger.log('   ' + c); });
  }

  // 명부에 없는데 시트에만 있는 담당자 (퇴사자 등) 도 함께 알린다
  const onlyInSheet = Object.keys(existing).filter(function (n) { return !(n in KNOWN_OWNER_SLACK_IDS); });
  if (onlyInSheet.length > 0) {
    Logger.log('시트에만 있는 담당자 (슬랙 계정 확인 필요): ' + onlyInSheet.join(', '));
  }
}

/**
 * 상태시트 헤더 보정. STATE_HEADERS에 새 컬럼이 추가돼도
 * 시트를 수동으로 고치지 않아도 되게 오른쪽에 자동 추가한다.
 */
function ensureStateHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, STATE_HEADERS.length).setValues([STATE_HEADERS]);
    sheet.setFrozenRows(1);
    return STATE_HEADERS.slice();
  }

  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
                     .getValues()[0].map(h => String(h == null ? '' : h).trim());
  while (headers.length > 0 && headers[headers.length - 1] === '') headers.pop();

  const missing = STATE_HEADERS.filter(h => headers.indexOf(h) === -1);
  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    Logger.log('상태시트에 컬럼 추가: ' + missing.join(', '));
  }
  return headers;
}

/** ===================== 유틸 ===================== */
function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('컬럼을 찾을 수 없음: ' + name);
  return idx;
}

/** 한 번의 실행 안에서 같은 탭을 여러 번 읽지 않도록 캐싱 */
var _sheetCache = {};
function getSheetData(sheetName) {
  if (_sheetCache[sheetName]) return _sheetCache[sheetName];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('탭을 찾을 수 없음: ' + sheetName);
  const data = sheet.getDataRange().getValues();
  _sheetCache[sheetName] = { headers: data[0], rows: data.slice(1) };
  return _sheetCache[sheetName];
}
function clearSheetCache() { _sheetCache = {}; }

/** 공란 판정 — 공백 문자열과 수식 오류값도 미입력으로 본다 */
function isBlank(v) {
  if (v === null || v === undefined) return true;
  const s = String(v).trim();
  return s === '' || s === '#N/A' || s === '#REF!' || s === '#VALUE!' || s === '#ERROR!';
}

/**
 * 날짜 값 → Date. 파싱 불가/빈값이면 null.
 * "2026.01.30 (금)", "2026.1.5", "2026-01-30", "2026/1/5", Date 객체 모두 처리.
 */
function parseKDate(str) {
  if (!str) return null;
  if (Object.prototype.toString.call(str) === '[object Date]') {
    return isNaN(str.getTime()) ? null : str;
  }
  const m = String(str).match(/(\d{4})\s*[-.\/]\s*(\d{1,2})\s*[-.\/]\s*(\d{1,2})/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function startOfToday() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
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
  clearSheetCache();

  // 한 스캐너가 실패해도 나머지는 계속 진행한다 (컬럼명 변경 등으로 전체가 멈추지 않게)
  const scanners = [
    ['객단가미입력', scanUnitPrice],
    ['기획전테마미입력', scanTheme],
    ['리스팅기간미설정', scanAdPeriod],
    ['서명미완료', scanSignature],
    ['결제미완료', scanPayment]
  ];

  let items = [];
  const failed = [];
  scanners.forEach(function (pair) {
    try {
      items = items.concat(pair[1]());
    } catch (e) {
      failed.push(pair[0] + '(' + e.message + ')');
      Logger.log('스캔 실패: ' + pair[0] + ' — ' + e.message);
    }
  });

  // 일부 스캐너가 죽은 상태로 upsert하면 살아있는 항목이 '취소추정'으로 오분류된다
  if (failed.length > 0) {
    throw new Error('스캔 실패로 upsert를 중단했습니다: ' + failed.join(', '));
  }

  upsertStateSheet(items, getPresentActIdsByTab_());
}

/**
 * 탭별로 현재 남아있는 활동ID 집합.
 * 항목이 스캔에서 사라진 이유를 구분하는 데 쓴다.
 *   - 소스 행은 그대로인데 조건을 벗어남 → 실제 완료(O)
 *   - 소스 행 자체가 사라짐(원본에서 참여→취소) → 취소추정
 */
function getPresentActIdsByTab_() {
  const result = {};
  [CONFIG.ACTIVITY_SHEET, CONFIG.PRODUCT_SHEET, CONFIG.CONTRACT_SHEET].forEach(function (tab) {
    try {
      const d = getSheetData(tab);
      const i = colIndex(d.headers, '활동ID');
      const set = new Set();
      d.rows.forEach(function (r) { if (r[i]) set.add(r[i]); });
      result[tab] = set;
    } catch (e) {
      Logger.log('활동ID 집합 수집 실패 (' + tab + '): ' + e.message);
    }
  });
  return result;
}

/**
 * 객단가 미입력 — 예약핫플/트렌드, 광고시작일~종료일
 * 런치·디너 중 하나라도 비어 있으면 알람 대상.
 * (문서상 스펙은 "둘 다 공란"이지만 실무상 둘 다 입력이 필요해 기존 동작을 유지함)
 */
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
  const today = startOfToday();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.UNITPRICE)) return false;
      if (!isBlank(r[iLunch]) && !isBlank(r[iDinner])) return false;
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
  const today = startOfToday();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여' || !isBlank(r[iTheme])) return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.THEME_PERIOD)) return false;
      const start = parseKDate(r[iStart]);
      if (!start) return false;
      return today >= addDays(start, -14) && today < start;
    })
    .map(r => ({
      itemId: 'THEME_' + r[iAct], type: '기획전테마미입력', tab: CONFIG.PRODUCT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 리스팅 기간 미설정 — 예약핫플, 시작일-14일~시작일 (시작일 자체가 없으면 즉시 알람) */
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
  const today = startOfToday();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.THEME_PERIOD)) return false;
      if (!isBlank(r[iStart]) && !isBlank(r[iEnd])) return false;

      const start = parseKDate(r[iStart]);
      if (!start) return true; // 시작일 자체가 없음 → 즉시 알람 대상
      return today >= addDays(start, -14) && today < start;
    })
    .map(r => ({
      // itemId 접두사는 기존 상태시트 행과의 연속성을 위해 ADPERIOD_ 유지
      itemId: 'ADPERIOD_' + r[iAct], type: '리스팅기간미설정', tab: CONFIG.PRODUCT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 계약서 서명 미완료 — 예약핫플/트렌드/웨이팅, 시작일-14일~종료일+GRACE_DAYS */
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
  const today = startOfToday();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여') return false;
      if (String(r[iStatus] || '').trim() === '계약서 서명 완료') return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.SIGN_PAYMENT)) return false;
      const start = parseKDate(r[iStart]);
      const end = parseKDate(r[iEnd]);
      if (!start || !end) return false;
      return today >= addDays(start, -14) && today <= addDays(end, GRACE_DAYS);
    })
    .map(r => ({
      itemId: 'SIGN_' + r[iAct], type: '서명미완료', tab: CONFIG.CONTRACT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/**
 * 결제 미완료 — 예약핫플/트렌드/웨이팅, 시작일-14일~종료일+GRACE_DAYS
 * 입금 마감일이 이미 지난 건은 미수방지 봇이 전담하므로 여기서 제외한다
 * (CONFIG.SKIP_OVERDUE_PAYMENT = false 로 두면 예전처럼 전부 알람).
 */
function scanPayment() {
  const { headers, rows } = getSheetData(CONFIG.CONTRACT_SHEET);
  const iAct = colIndex(headers, '활동ID');
  const iSeq = colIndex(headers, '계약_매장시퀀스');
  const iName = colIndex(headers, '계약_매장명');
  const iJoin = colIndex(headers, '계약_참여여부');
  const iAdName = colIndex(headers, '계약_광고명');
  const iPayDate = colIndex(headers, '결제일');
  const iDueDate = colIndex(headers, '입금 마감일');
  const iStart = colIndex(headers, '시작일');
  const iEnd = colIndex(headers, '종료일');
  const iOwner = colIndex(headers, '계약_담당자');
  const today = startOfToday();

  return rows
    .filter(r => {
      if (r[iJoin] !== '참여' || !isBlank(r[iPayDate])) return false;
      if (!matchesAdType(r[iAdName], AD_TYPES.SIGN_PAYMENT)) return false;

      if (CONFIG.SKIP_OVERDUE_PAYMENT) {
        const due = parseKDate(r[iDueDate]);
        if (due && due.getTime() < today.getTime()) return false; // 미수방지 봇 담당
      }

      const start = parseKDate(r[iStart]);
      const end = parseKDate(r[iEnd]);
      if (!start || !end) return false;
      return today >= addDays(start, -14) && today <= addDays(end, GRACE_DAYS);
    })
    .map(r => ({
      itemId: 'PAYMENT_' + r[iAct], type: '결제미완료', tab: CONFIG.CONTRACT_SHEET,
      actId: r[iAct], seq: r[iSeq], storeName: r[iName], owner: r[iOwner]
    }));
}

/** 상태시트 upsert: 신규 항목 추가 + 미러링에서 사라진 항목은 '취소추정' 처리 */
function upsertStateSheet(items, presentByTab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.STATE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const idIdx = colIndex(headers, '항목ID');
  const existingIds = new Set(rows.map(r => r[idIdx]));
  const now = new Date();

  const idx = {};
  ['항목ID','유형','탭','활동ID','매장시퀀스','매장명','담당자','최초알림일','확인여부','실제완료여부']
    .forEach(name => { idx[name] = colIndex(headers, name); });

  const tabIdx = idx['탭'];
  const actIdx = idx['활동ID'];

  /**
   * 스캔에서 빠진 항목의 사유 판정.
   * 서명 완료·결제 완료·객단가 입력처럼 담당자가 실제로 처리한 경우는
   * 소스 행이 그대로 남아 있으므로 '취소추정'이 아니라 'O'(완료)로 본다.
   */
  const resolveDisappeared = function (row) {
    const set = presentByTab && presentByTab[row[tabIdx]];
    if (!set) return '취소추정';           // 판정 근거가 없으면 기존 동작 유지
    return set.has(row[actIdx]) ? 'O' : '취소추정';
  };

  // 신규 항목 추가 (한 번에 append)
  const newRows = items
    .filter(item => !existingIds.has(item.itemId))
    .map(item => {
      const row = new Array(headers.length).fill('');
      row[idx['항목ID']] = item.itemId;
      row[idx['유형']] = item.type;
      row[idx['탭']] = item.tab;
      row[idx['활동ID']] = item.actId;
      row[idx['매장시퀀스']] = item.seq;
      row[idx['매장명']] = item.storeName;
      row[idx['담당자']] = item.owner;
      row[idx['최초알림일']] = now;
      row[idx['확인여부']] = 'X';
      row[idx['실제완료여부']] = 'X';
      return row;
    });
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  // 이번 스캔에 안 잡힌(=사라진) 미완료 항목 → 취소추정 / 다시 잡힌 항목 → 복구
  const currentIds = new Set(items.map(i => i.itemId));
  const cIdx = colIndex(headers, '실제완료여부');
  const dIdx = colIndex(headers, '완료일');

  if (rows.length > 0) {
    let changed = 0;
    if (dIdx === cIdx + 1) {
      // 인접 컬럼이면 2열 한 덩어리로 한 번에 쓴다 (셀 단위 setValue 루프 제거)
      const block = rows.map(r => [r[cIdx], r[dIdx]]);
      rows.forEach((r, i) => {
        if (r[cIdx] === 'X' && !currentIds.has(r[idIdx])) {
          block[i] = [resolveDisappeared(r), now]; changed++;
        } else if ((r[cIdx] === '취소추정' || r[cIdx] === 'O') && currentIds.has(r[idIdx])) {
          block[i] = ['X', '']; changed++;   // 다시 미완료로 잡힘 → 복구
        }
      });
      if (changed > 0) sheet.getRange(2, cIdx + 1, rows.length, 2).setValues(block);
    } else {
      rows.forEach((r, i) => {
        const rowNum = i + 2;
        if (r[cIdx] === 'X' && !currentIds.has(r[idIdx])) {
          sheet.getRange(rowNum, cIdx + 1).setValue(resolveDisappeared(r));
          sheet.getRange(rowNum, dIdx + 1).setValue(now);
          changed++;
        } else if ((r[cIdx] === '취소추정' || r[cIdx] === 'O') && currentIds.has(r[idIdx])) {
          sheet.getRange(rowNum, cIdx + 1).setValue('X');
          sheet.getRange(rowNum, dIdx + 1).setValue('');
          changed++;
        }
      });
    }
    Logger.log('신규 ' + newRows.length + '건 / 상태변경 ' + changed + '건');
  }
}

/** ===================== 2) 담당자별 DM 발송 ===================== */
function sendDMs() {
  // 평일(월~금)에만 발송. 트리거는 매일 걸려있지만 주말엔 여기서 조용히 스킵.
  const dow = new Date().getDay(); // 0=일, 6=토
  if (dow === 0 || dow === 6) { Logger.log('주말이라 발송 스킵'); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const iOwner = colIndex(headers, '담당자');
  const iLastDM = colIndex(headers, '마지막DM발송일');
  const iType = colIndex(headers, '유형');
  const iStore = colIndex(headers, '매장명');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iTs = colIndex(headers, 'DM메시지TS');
  const iChannel = colIndex(headers, 'DM채널ID');

  const now = new Date();
  const remindMs = CONFIG.REMIND_HOURS * 60 * 60 * 1000;
  const groups = {};

  rows.forEach((r, i) => {
    if (r[iConfirm] !== 'X' || r[iComplete] !== 'X') return;
    const owner = r[iOwner];

    // 테스트 모드면 지정한 출처 담당자 항목만 대상으로
    if (CONFIG.TEST_MODE && owner !== CONFIG.TEST_SOURCE_OWNER) return;

    const last = r[iLastDM];
    if (last && (now - new Date(last)) <= remindMs) return;
    (groups[owner] = groups[owner] || []).push({ rowIndex: i + 2, type: r[iType], store: r[iStore], seq: r[iSeq] });
  });

  const mapping = getOwnerSlackMap();
  let sentOwners = 0, failedOwners = 0;

  // 발송 결과를 모았다가 마지막에 범위 단위로 한 번에 쓴다.
  // (행마다 setValue를 부르면 200건 기준 수백 번의 API 왕복이 발생한다)
  const sendCols = rows.map(r => [r[iLastDM], r[iTs]]);   // 마지막DM발송일, DM메시지TS (인접)
  const chanCol = rows.map(r => [r[iChannel]]);           // DM채널ID (끝 컬럼)
  let touched = 0;

  Object.keys(groups).forEach(owner => {
    const slackId = CONFIG.TEST_MODE ? CONFIG.TEST_TARGET_SLACK_ID : mapping[owner];
    if (!slackId) { Logger.log('Slack ID 매핑 없음: ' + owner); return; }
    const list = groups[owner];

    // 한 명이 실패해도 나머지 담당자 발송이 멈추지 않게 감싼다
    let res = null;
    try {
      res = postSlackDM(slackId, buildDMBlocks(owner, list), buildDMFallbackText(owner, list));
    } catch (e) {
      Logger.log('DM 발송 예외 (' + owner + '): ' + e.message);
    }

    if (!res) { failedOwners++; return; }
    sentOwners++;
    list.forEach(item => {
      const i = item.rowIndex - 2;
      sendCols[i] = [now, res.ts];
      // 리액션 조회는 유저ID가 아니라 DM 채널ID(D…)를 요구한다. 발송 응답값을 저장해 둔다.
      chanCol[i] = [res.channel];
      touched++;
    });
  });

  if (touched > 0 && rows.length > 0) {
    if (iTs === iLastDM + 1) {
      sheet.getRange(2, iLastDM + 1, rows.length, 2).setValues(sendCols);
    } else {
      sheet.getRange(2, iLastDM + 1, rows.length, 1).setValues(sendCols.map(v => [v[0]]));
      sheet.getRange(2, iTs + 1, rows.length, 1).setValues(sendCols.map(v => [v[1]]));
    }
    sheet.getRange(2, iChannel + 1, rows.length, 1).setValues(chanCol);
  }

  Logger.log('DM 발송: 담당자 ' + sentOwners + '명 성공' + (failedOwners ? ' / ' + failedOwners + '명 실패' : ''));

  // 채널 요약 실패가 개인 DM 결과를 덮지 않게 분리해서 처리
  try {
    sendChannelBroadcast();
  } catch (e) {
    Logger.log('채널 요약 발송 실패: ' + e.message);
  }
}

/**
 * 유형별로 그룹핑 + Block Kit 섹션으로 DM 구성.
 * - 유형별 소제목("서명미완료 N건") 아래에 매장 전체 리스트 (TYPE_ORDER 순서, 심각한 유형 먼저)
 * - 건수 제한/요약 없음 — 몇 건이든 전부 그대로 노출
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
    // Block Kit section text는 3000자 제한이므로 길면 나눠 담는다
    const lines = items.map(it => `•  ${it.seq}  ${it.store}`);
    let buf = [], len = 0;
    const flush = (isFirst) => {
      if (buf.length === 0) return;
      const title = isFirst ? `*${type}  ${items.length}건*\n` : '';
      blocks.push({ type: 'section', text: { type: 'mrkdwn', text: title + buf.join('\n') } });
      buf = []; len = 0;
    };
    let first = true;
    lines.forEach(line => {
      if (len + line.length > 2800) { flush(first); first = false; }
      buf.push(line); len += line.length + 1;
    });
    flush(first);
  });

  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'context',
    elements: [{ type: 'mrkdwn', text: '✅ 확인하셨으면 이 메시지에 리액션 하나만 눌러주세요 (이모지 종류 무관) 🙏' }]
  });

  return blocks;
}

/** 알림 미리보기/스크린리더 등 fallback용 평문 텍스트 */
function buildDMFallbackText(owner, list) {
  return `${owner}님, 처리 필요한 항목 ${list.length}건이 있어요. DM에서 유형별로 확인해주세요.`;
}

function getOwnerSlackMap() {
  const { headers, rows } = getSheetData(CONFIG.MAPPING_SHEET);
  const iName = colIndex(headers, '담당자명');
  const iSlack = colIndex(headers, 'SlackID');
  const map = {};
  rows.forEach(r => {
    const name = String(r[iName] || '').trim();
    if (name) map[name] = String(r[iSlack] || '').trim();
  });
  return map;
}

/** DM 발송. 성공 시 {ts, channel} 반환 (channel은 D… 형태의 DM 채널ID) */
function postSlackDM(slackUserId, blocks, fallbackText) {
  const res = slackCall_('chat.postMessage', {
    channel: slackUserId, text: fallbackText, blocks: blocks
  });
  if (!res.ok) { Logger.log('Slack 발송 실패 (' + slackUserId + '): ' + res.error); return null; }
  return { ts: res.ts, channel: res.channel || slackUserId };
}

function postSlackChannelMessage(blocks, fallbackText, channelOverride) {
  const res = slackCall_('chat.postMessage', {
    channel: channelOverride || CONFIG.CHANNEL_ID, text: fallbackText, blocks: blocks
  });
  if (!res.ok) { Logger.log('채널 발송 실패: ' + res.error); return null; }
  return res.ts;
}

/** Slack POST 공통 처리 */
function slackCall_(method, payload) {
  const response = UrlFetchApp.fetch('https://slack.com/api/' + method, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + getSlackToken_() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  try {
    return JSON.parse(response.getContentText());
  } catch (e) {
    return { ok: false, error: 'invalid_response: ' + response.getContentText().slice(0, 200) };
  }
}

/** ===================== 3) 리액션 폴링 (확인 체크) ===================== */
function pollReactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const iConfirm = colIndex(headers, '확인여부');
  const iConfirmTime = colIndex(headers, '확인시각');
  const iTs = colIndex(headers, 'DM메시지TS');
  const iChannel = colIndex(headers, 'DM채널ID');
  const iOwner = colIndex(headers, '담당자');

  // ts → 행 번호 목록. 예전엔 행마다 전체를 재순회해서 O(n^2)였다.
  const byTs = {};
  rows.forEach((r, i) => {
    const ts = r[iTs];
    if (r[iConfirm] !== 'X' || !ts) return;
    const key = String(ts);
    if (!byTs[key]) byTs[key] = { rowIdx: [], channel: '', owner: r[iOwner] };
    byTs[key].rowIdx.push(i);
    if (!byTs[key].channel && r[iChannel]) byTs[key].channel = String(r[iChannel]).trim();
  });

  const mapping = getOwnerSlackMap();
  const confirmed = [];   // [rowIdx...]

  // 유저ID → DM채널ID 캐시. 같은 담당자의 과거 DM이 수십 건이면
  // ts 그룹마다 conversations.open을 부르게 되므로 한 번만 조회한다.
  const dmChannelCache = {};
  const resolveChannel = function (slackId) {
    if (slackId in dmChannelCache) return dmChannelCache[slackId];
    dmChannelCache[slackId] = openDmChannel_(slackId);
    return dmChannelCache[slackId];
  };

  let noChannel = 0, noMapping = 0;
  let fatal = null;

  Object.keys(byTs).forEach(ts => {
    if (fatal) return;
    const g = byTs[ts];
    let channel = g.channel;

    // 예전 행에는 DM채널ID가 없다. 유저ID로는 reactions.get이 동작하지 않으므로 변환한다.
    if (!channel) {
      const slackId = CONFIG.TEST_MODE ? CONFIG.TEST_TARGET_SLACK_ID : mapping[g.owner];
      if (!slackId) { noMapping++; return; }
      channel = resolveChannel(slackId);
      if (!channel) { noChannel++; return; }
    }

    // 스코프·토큰 문제는 트리거가 매시간 실패 메일을 보내게 만든다.
    // 예외로 터뜨리는 대신 사유를 한 번 남기고 조용히 종료한다.
    try {
      if (checkSlackReaction(channel, ts)) confirmed.push.apply(confirmed, g.rowIdx);
    } catch (e) {
      fatal = e.message;
    }
  });

  if (fatal) {
    Logger.log('리액션 조회를 중단했습니다.\n' + fatal);
    return;
  }

  if (noMapping > 0) Logger.log('담당자 매핑이 없어 리액션 확인을 건너뛴 DM: ' + noMapping + '건');
  if (noChannel > 0) Logger.log('DM 채널 조회 실패로 건너뛴 DM: ' + noChannel + '건 (봇 스코프 im:write 확인 필요)');
  Logger.log('리액션 조회 대상 DM ' + Object.keys(byTs).length + '건 / conversations.open 호출 ' +
             Object.keys(dmChannelCache).length + '회');

  if (confirmed.length === 0) { Logger.log('새로 확인된 항목 없음'); return; }

  // 확인여부/확인시각은 인접 컬럼이므로 한 덩어리로 쓴다
  const now = new Date();
  if (iConfirmTime === iConfirm + 1) {
    const block = rows.map(r => [r[iConfirm], r[iConfirmTime]]);
    confirmed.forEach(i => { block[i] = ['O', now]; });
    sheet.getRange(2, iConfirm + 1, rows.length, 2).setValues(block);
  } else {
    confirmed.forEach(i => {
      sheet.getRange(i + 2, iConfirm + 1).setValue('O');
      sheet.getRange(i + 2, iConfirmTime + 1).setValue(now);
    });
  }
  Logger.log('확인 처리: ' + confirmed.length + '건');
}

/** 유저ID → DM 채널ID(D…). 이미 열려 있으면 기존 채널을 그대로 돌려준다 */
function openDmChannel_(slackUserId) {
  const res = slackCall_('conversations.open', { users: slackUserId });
  if (!res.ok || !res.channel) {
    Logger.log('DM 채널 조회 실패 (' + slackUserId + '): ' + res.error +
               ' — 봇 스코프에 im:write 가 필요할 수 있습니다.');
    return null;
  }
  return res.channel.id;
}

// 재시도해도 결과가 같은 치명적 오류 — 만나면 즉시 중단한다
var SLACK_FATAL_ERRORS = ['missing_scope', 'invalid_auth', 'not_authed', 'token_revoked',
                          'account_inactive', 'ratelimited'];

function checkSlackReaction(channel, ts) {
  const url = 'https://slack.com/api/reactions.get?channel=' + encodeURIComponent(channel) +
              '&timestamp=' + encodeURIComponent(ts);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + getSlackToken_() },
    muteHttpExceptions: true
  });
  let res;
  try { res = JSON.parse(response.getContentText()); }
  catch (e) { Logger.log('리액션 조회 응답 파싱 실패: ' + response.getContentText().slice(0, 200)); return false; }

  // 조용히 false를 반환하면 리마인드가 영구히 반복되므로 실패 이유를 반드시 남긴다
  if (!res.ok) {
    // 스코프·토큰 문제는 1000건을 두드려도 결과가 같다. 로그를 채우고 실행시간만
    // 잡아먹으므로 첫 건에서 바로 멈추고 조치 방법을 알린다.
    if (SLACK_FATAL_ERRORS.indexOf(res.error) !== -1) {
      throw new Error('Slack API 오류로 리액션 조회를 중단합니다: ' + res.error +
        (res.needed ? ' (필요 스코프: ' + res.needed + ')' : '') +
        (res.provided ? ' / 현재 스코프: ' + res.provided : '') +
        '\n→ api.slack.com/apps > 해당 앱 > OAuth & Permissions 에서 스코프를 추가하고 ' +
        '워크스페이스에 재설치(Reinstall)한 뒤, 새 봇 토큰을 스크립트 속성 SLACK_BOT_TOKEN에 갱신하세요.');
    }
    Logger.log('리액션 조회 실패 (' + channel + '/' + ts + '): ' + res.error);
    return false;
  }
  const reactions = res.message && res.message.reactions;
  return !!(reactions && reactions.length > 0);
}

/** ===================== 4) 채널 요약 (기획전테마/리스팅기간) ===================== */
function sendChannelBroadcast() {
  if (CONFIG.TEST_MODE) { Logger.log('테스트 모드 - 채널 발송 스킵'); return; }

  const dow = new Date().getDay();
  if (dow === 0 || dow === 6) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const rows = sheet.getDataRange().getValues().slice(1);

  const iType = colIndex(headers, '유형');
  const iOwner = colIndex(headers, '담당자');
  const iSeq = colIndex(headers, '매장시퀀스');
  const iStore = colIndex(headers, '매장명');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const ownerSlackMap = getOwnerSlackMap();

  const byType = {};
  rows.forEach(function (r) {
    const type = r[iType];
    if (CHANNEL_TYPES.indexOf(type) === -1) return;
    if (r[iConfirm] !== 'X' || r[iComplete] !== 'X') return;
    if (!byType[type]) byType[type] = [];
    byType[type].push({ owner: r[iOwner], seq: r[iSeq], store: r[iStore] });
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
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: 'cc. <@' + CONFIG.LEAD_SLACK_ID + '>' }] });

  postSlackChannelMessage(blocks, '처리 필요한 항목 ' + totalCount + '건 (기획전테마미입력/리스팅기간미설정)');
}

/** ===================== 5) 차수별 서명/결제 미완료 채널 리포트 ===================== */
function sendComplianceListByRoundToChannel(channelOverride) {
  // 트리거 실행 시에는 이벤트 객체가 들어오므로 문자열일 때만 채널 오버라이드로 취급
  const channel = (typeof channelOverride === 'string' && channelOverride) ? channelOverride : CONFIG.CHANNEL_ID;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const rows = sheet.getDataRange().getValues().slice(1);
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

  const today = startOfToday();
  const byRound = {};
  let unmatched = 0;

  rows.forEach(function (r) {
    if (TYPES.indexOf(r[iType]) === -1) return;
    if (r[iComplete] !== 'X') return; // 확인(리액션) 여부와 무관하게 미완료면 집계

    const c = contractMap[r[iActivity]];
    if (!c || !c.name || !c.start || !c.end) { unmatched++; return; }
    if (!matchesAdType(c.name, AD_TYPES.ROUND_REPORT)) return; // 예약핫플·트렌드만

    const roundName = c.name.split('_')[0].trim();
    if (!byRound[roundName]) byRound[roundName] = { 서명미완료: 0, 결제미완료: 0, starts: [], ends: [] };
    byRound[roundName][r[iType]]++;
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

  // 진행 중 + 예정(시작 전) 차수 — 시작 전이 오히려 서명/결제를 챙겨야 하는 시점
  // 여기에 광고타입별 직전 종료 차수 1개씩을 덧붙인다
  const ongoing = roundList.filter(function (g) { return today <= g.end; });
  const ended = roundList.filter(function (g) { return g.end < today; });
  const picked = ongoing.slice();
  AD_TYPES.ROUND_REPORT.forEach(function (k) {
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
  let ctx = 'cc. <!subteam^' + CONFIG.CELL_GROUP_ID + '> <@' + CONFIG.LEAD_SLACK_ID + '>';
  if (unmatched > 0) ctx += '  |  계약 날짜 미확인으로 차수 미분류 ' + unmatched + '건';
  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: ctx }] });

  const ts = postSlackChannelMessage(blocks, '차수별 서명/결제 미완료 ' + totalCount + '건', channel);
  if (ts) Logger.log('발송 완료 / 차수 ' + picked.length + '개 / 총 ' + totalCount + '건 / 미매칭 ' + unmatched + '건');
}

/** 테스트 채널로만 결과 미리보기 발송 */
function previewComplianceListToTestChannel() {
  sendComplianceListByRoundToChannel(CONFIG.TEST_CHANNEL_ID);
}

/** ===================== 트리거 설정 ===================== */
function setupTriggers() {
  // 핸들러명을 지정해서 지운다. 전체 삭제하면 sendComplianceListByRoundToChannel
  // 평일 트리거까지 함께 사라진다.
  const handled = ['scanAndUpsert', 'sendDMs', 'pollReactions'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (handled.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanAndUpsert').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendDMs').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('pollReactions').timeBased().everyHours(1).create();
  Logger.log('트리거 재설정 완료 (scanAndUpsert 1시간, sendDMs 매일 9시, pollReactions 1시간)');
}

function setupComplianceChannelTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendComplianceListByRoundToChannel') ScriptApp.deleteTrigger(t);
  });
  const days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
                ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
  days.forEach(function (d) {
    ScriptApp.newTrigger('sendComplianceListByRoundToChannel').timeBased().onWeekDay(d).atHour(9).create();
  });
  Logger.log('평일(월~금) 오전 9시 트리거 5개 생성 완료');
}

function listTriggers() {
  const out = ScriptApp.getProjectTriggers().map(function (t) {
    return t.getHandlerFunction() + ' / ' + t.getEventType();
  });
  Logger.log('등록된 트리거 ' + out.length + '개:\n' + out.join('\n'));
  return out;
}

/** ===================== 진단 ===================== */
function diagnoseSlackAccess() {
  // 응답 헤더 x-oauth-scopes 에 현재 토큰에 부여된 스코프가 들어온다
  const response = UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + getSlackToken_() },
    muteHttpExceptions: true
  });
  const auth = JSON.parse(response.getContentText());
  Logger.log('auth.test: ' + JSON.stringify(auth));

  const headers = response.getHeaders();
  const granted = String(headers['x-oauth-scopes'] || headers['X-OAuth-Scopes'] || '');
  Logger.log('현재 봇 토큰 스코프: ' + (granted || '(응답에 없음)'));

  const required = ['chat:write', 'reactions:read', 'im:write', 'im:history'];
  const have = granted.split(',').map(function (v) { return v.trim(); });
  const missing = required.filter(function (r) { return have.indexOf(r) === -1; });
  Logger.log(missing.length
    ? '❌ 누락 스코프: ' + missing.join(', ') + ' → 추가 후 워크스페이스 재설치 필요'
    : '✅ 필요한 스코프 모두 보유');

  const info = JSON.parse(UrlFetchApp.fetch(
    'https://slack.com/api/conversations.info?channel=' + CONFIG.CHANNEL_ID,
    { method: 'get', headers: { Authorization: 'Bearer ' + getSlackToken_() }, muteHttpExceptions: true }
  ).getContentText());
  Logger.log('conversations.info: ' + JSON.stringify(info));
}

function diagnoseSendDMs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const rows = sheet.getDataRange().getValues().slice(1);
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');
  const iOwner = colIndex(headers, '담당자');
  const iLastDM = colIndex(headers, '마지막DM발송일');
  const iChannel = colIndex(headers, 'DM채널ID');
  const iTs = colIndex(headers, 'DM메시지TS');

  const now = new Date();
  const remindMs = CONFIG.REMIND_HOURS * 60 * 60 * 1000;
  const groups = {};
  let totalMatch = 0, missingChannel = 0;

  rows.forEach(function (r) {
    if (r[iTs] && !r[iChannel]) missingChannel++;
    if (r[iConfirm] !== 'X' || r[iComplete] !== 'X') return;
    totalMatch++;
    const owner = r[iOwner];
    const last = r[iLastDM];
    const needSend = !last || (now - new Date(last)) > remindMs;
    if (!groups[owner]) groups[owner] = { total: 0, needSend: 0, lastSample: null };
    groups[owner].total++;
    if (needSend) groups[owner].needSend++;
    if (!groups[owner].lastSample) groups[owner].lastSample = last ? new Date(last).toString() : '(never)';
  });

  const mapping = getOwnerSlackMap();
  const report = {};
  Object.keys(groups).forEach(function (owner) {
    report[owner] = {
      total: groups[owner].total, needSend: groups[owner].needSend,
      lastDMSample: groups[owner].lastSample, hasSlackMapping: !!mapping[owner]
    };
  });

  Logger.log('TEST_MODE: ' + CONFIG.TEST_MODE);
  Logger.log('요일(0=일,6=토): ' + new Date().getDay());
  Logger.log('전체 대상 행수: ' + totalMatch);
  Logger.log('DM채널ID 없는 기존 행: ' + missingChannel + ' (conversations.open으로 대체 조회됨)');
  Logger.log('담당자별 현황: ' + JSON.stringify(report));
}

function diagnoseStateSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.STATE_SHEET);
  const headers = ensureStateHeaders_(sheet);
  const rows = sheet.getDataRange().getValues().slice(1);
  const iType = colIndex(headers, '유형');
  const iConfirm = colIndex(headers, '확인여부');
  const iComplete = colIndex(headers, '실제완료여부');

  const byType = {}, byState = {};
  rows.forEach(function (r) {
    byType[r[iType]] = (byType[r[iType]] || 0) + 1;
    const key = '확인' + r[iConfirm] + '/완료' + r[iComplete];
    byState[key] = (byState[key] || 0) + 1;
  });
  Logger.log('전체 행수: ' + rows.length);
  Logger.log('유형별: ' + JSON.stringify(byType));
  Logger.log('상태별: ' + JSON.stringify(byState));
}

/** 시크릿 설정 상태만 확인 (값 자체는 출력하지 않는다) */
function diagnoseWebhookSecret() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty('WEBHOOK_SHARED_SECRET');
  const prev = props.getProperty('WEBHOOK_SHARED_SECRET_PREV');
  const mask = function (v) { return v ? '설정됨 (' + v.length + '자, 끝 4자 …' + v.slice(-4) + ')' : '없음'; };
  Logger.log('WEBHOOK_SHARED_SECRET: ' + mask(cur));
  Logger.log('WEBHOOK_SHARED_SECRET_PREV: ' + mask(prev) +
             (prev ? '  ← 호출 측 갱신이 끝나면 삭제하세요' : ''));
  if (!cur && !prev) Logger.log('⚠️ 둘 다 없습니다. doPost 호출이 전부 실패합니다.');
}

/** ===================== 외부 호출용 Slack 메시지 발송 웹앱 ===================== */
/** catchweek-lms-send 등 외부에서 호출. 시크릿은 스크립트 속성에서 읽는다. */
const WEBHOOK_ALLOWED_CHANNELS = ['C08731XUE95', 'C07G05CCGSX', 'C09SQQ000H2', 'C095PRYDE12', 'C09C2DWQXB8'];

function doPost(e) {
  const json = function (obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  };
  try {
    const body = JSON.parse(e.postData.contents);
    if (!isValidWebhookSecret_(body.secret)) return json({ ok: false, error: 'unauthorized' });

    const channel = body.channel;
    const text = body.message;
    if (!text || !channel) return json({ ok: false, error: 'missing message or channel' });
    if (WEBHOOK_ALLOWED_CHANNELS.indexOf(channel) === -1) return json({ ok: false, error: 'channel not allowed' });

    const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + getSlackToken_(), 'Content-Type': 'application/json; charset=utf-8' },
      payload: JSON.stringify({ channel: channel, thread_ts: (body.thread_ts || undefined), text: text }),
      muteHttpExceptions: true
    });
    return ContentService.createTextOutput(res.getContentText()).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}
