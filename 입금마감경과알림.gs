/**
 * 입금 마감일 경과 · 미입금 매장 담당자 개인 1:1 DM 알림봇
 *
 * "계약현황" 시트에서 아래 조건을 모두 만족하는 매장을 담당자별로 모아
 * 개인 Slack DM으로 발송한다.
 *   - 계약상태 = "계약서 서명 완료"
 *   - 결제일(AL) 공란 (= 미입금)
 *   - 광고 종료일(H)이 2026-01-01 이후
 *   - 입금 마감일(K)이 오늘 이전(경과)
 *   - 계약대금 > 0
 *
 * ※ 공통 헬퍼 sendSlackDM_() 와 담당자 매핑 SLACK_USER_IDS 는
 *   같은 프로젝트의 "웨이팅광고종료알림.gs" 에 정의된 것을 재사용한다.
 *   (같은 Apps Script 프로젝트에 두 번 정의하면 충돌하므로 여기서는 재정의하지 않음)
 *   단독 프로젝트로 쓸 경우 아래 두 블록의 주석을 해제할 것.
 */

// ── 이 스크립트를 단독 프로젝트로 쓸 때만 주석 해제 ──────────
// var SLACK_USER_IDS = {
//   "이도은": "U093FJ7DZ8W", "최원영": "U0BLAHC00G3", "권용덕": "U0BLDGBS8G5",
//   "남현욱": "U07RH97HNQL", "우은수": "U093FJ573FY", "이종익": "U0AGLUM2G2V",
//   "김나현": "U0B99RC7H08", "홍성혁": "U02TN1U2PQR", "이하윤": "U0AJ3LN8E3T",
//   "이조은": "U09GZ0H7928", "이동헌": "U0AG4JQJVKP", "김현수": "U0BHGG7FP98",
//   "전평정": "U02QCTZT2PP", "이세한": "U09BZ6JL60G", "이혜민": "U0AJY0DSMPC",
//   "김상하": "U0AG600HV1U", "이승민": "U0AG1KQQFS7", "한창완": "U057M7S5RA9",
//   "신유빈": "U09MXM4BV71", "이지민": "U0BGDNGHUBC"
// };

var UNPAID_CONFIG = {
  SPREADSHEET_ID: '1TIjXcv7E7QQNcEhoRrKKUJwJt9ftlumKUuVQSs0bD8Y',
  SHEET_NAME: '계약현황',
  MIN_END_DATE: '2026-01-01', // 광고 종료일이 이 날짜 이후인 건만 대상

  // "계약현황" 탭 0-based 컬럼 인덱스 (2026-07-09 확정 스키마 기준)
  COL: {
    seq: 2,          // 계약_매장시퀀스
    storeName: 3,    // 계약_매장명
    adName: 5,       // 계약_광고명
    manager: 6,      // 계약_담당자
    endDate: 8,      // 종료일
    dueDate: 10,     // 입금 마감일
    price: 11,       // 계약대금
    signStatus: 25,  // 계약상태
    payStatus: 37    // 결제일
  }
};

function notifyOverdueUnpaidStoresDedicatedDM() {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) {
    Logger.log('[중단] SLACK_BOT_TOKEN 스크립트 속성이 설정되지 않았습니다.');
    return;
  }

  var cfg = UNPAID_CONFIG;
  var sheet;
  try {
    sheet = SpreadsheetApp.openById(cfg.SPREADSHEET_ID).getSheetByName(cfg.SHEET_NAME);
  } catch (e) {
    Logger.log('[중단] 스프레드시트 접근 오류: ' + e.message);
    return;
  }
  if (!sheet) {
    Logger.log('[중단] "' + cfg.SHEET_NAME + '" 시트를 찾을 수 없습니다.');
    return;
  }

  var reportData = collectOverdueUnpaid_(sheet, cfg);
  var slackIds = Object.keys(reportData);
  if (slackIds.length === 0) {
    Logger.log('입금 마감 경과·미입금 대상이 없습니다. 발송 대상 없음.');
    return;
  }

  var sent = 0, failed = 0;
  slackIds.forEach(function (slackId) {
    var message = buildUnpaidMessage_(reportData[slackId]);
    if (sendSlackDM_(token, slackId, message)) {
      sent++;
    } else {
      failed++;
    }
  });

  Logger.log('입금 경과 알림 완료 — 발송 ' + sent + ' / 실패 ' + failed);
}

/**
 * 계약현황 시트를 스캔해 담당자(slackId)별 미입금 매장을 광고상품별로 취합한다.
 * @return {Object} { slackId: { name, count, ads: { 광고명: [줄문자열...] } } }
 */
function collectOverdueUnpaid_(sheet, cfg) {
  var data = sheet.getDataRange().getValues();
  var col = cfg.COL;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var minEndTime = new Date(cfg.MIN_END_DATE).getTime();

  var report = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col.adName]) continue;

    // [필터 1] 서명 완료 & 결제 미완료(결제일 공란)
    var signStatus = String(row[col.signStatus] || '').trim();
    var payStatus = String(row[col.payStatus] || '').trim();
    if (signStatus !== '계약서 서명 완료' || payStatus !== '') continue;

    // [필터 2] 광고 종료일이 기준일 이후
    var endDate = parseSheetDate_(row[col.endDate]);
    if (!endDate || endDate.getTime() < minEndTime) continue;

    // [필터 3] 입금 마감일이 오늘 이전(경과)
    var dueDate = parseSheetDate_(row[col.dueDate]);
    if (!dueDate) continue;
    dueDate.setHours(0, 0, 0, 0);
    if (dueDate.getTime() >= today.getTime()) continue;

    // [필터 4] 계약대금 0원 제외
    var contractPrice = Number(String(row[col.price] == null ? '' : row[col.price]).replace(/,/g, '')) || 0;
    if (contractPrice === 0) continue;

    // [필터 5] 명부에 등록된 담당자만 발송
    var managerName = String(row[col.manager] || '담당자미지정').trim();
    var slackId = SLACK_USER_IDS[managerName];
    if (!slackId) continue;

    if (!report[slackId]) report[slackId] = { name: managerName, count: 0, ads: {} };

    var unifiedAdName = String(row[col.adName]).trim().split('_')[0].trim();
    var formattedDueDate = (dueDate.getMonth() + 1) + '/' + dueDate.getDate();
    var storeInfo = '  • ' + row[col.storeName] + '(' + row[col.seq] + ') - '
      + contractPrice.toLocaleString() + '원 (마감일: ' + formattedDueDate + ')';

    if (!report[slackId].ads[unifiedAdName]) report[slackId].ads[unifiedAdName] = [];
    report[slackId].ads[unifiedAdName].push(storeInfo);
    report[slackId].count++;
  }
  return report;
}

/** 담당자 1명에게 보낼 DM 본문을 만든다. */
function buildUnpaidMessage_(managerInfo) {
  var detailText = '';
  for (var adTitle in managerInfo.ads) {
    var stores = managerInfo.ads[adTitle];
    detailText += '📦 *[' + adTitle + ']*\n' + stores.join('\n') + '\n\n';
  }

  var msg = '📢 *안녕하세요 ' + managerInfo.name + '님, 입금 마감일 경과 미입금 알림입니다.*\n\n';
  msg += '🔴 *담당하신 매장 중 서명 완료 / 결제 미완료 건 (총 ' + managerInfo.count + '건)*\n';
  msg += '_입금 마감일이 지났으나 아직 입금이 확인되지 않은 매장 리스트입니다._\n\n';
  msg += detailText;
  msg += '⚠️ *안내사항*\n';
  msg += '• 해당 내용은 앱시트 기반으로 가져오는 데이터입니다.\n';
  msg += '• 앱시트 상에서는 미수로 확인되니, **이미 입금이 완료되었다면 입금일자와 주문번호를 반드시 입력**해 주시기 바랍니다 🙏';
  return msg;
}

/**
 * 시트 셀 값을 Date로 변환. Date 객체 그대로 / "2026.01.15", "2026-01-15 00:00" 등
 * 문자열 모두 처리하고, #N/A 등 오류값·빈값은 null 반환.
 */
function parseSheetDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }
  var s = String(value == null ? '' : value).trim();
  if (!s || s.charAt(0) === '#') return null; // '', '#N/A', '#REF!' 등 제외
  var datePart = s.split(' ')[0].replace(/\./g, '-'); // 시간 부분 제거 + 점→하이픈
  var d = new Date(datePart);
  return isNaN(d.getTime()) ? null : d;
}
