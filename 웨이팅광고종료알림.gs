/**
 * 웨이팅 광고 종료 14일 전 알림 (개인 DM 전용)
 *
 * "상품설정" 시트에서 광고명에 "웨이팅"이 포함되고 광고종료일이
 * 오늘 ~ 14일 후 사이인 매장을 담당자별로 모아 개인 Slack DM으로 발송한다.
 *
 * 개선 포인트(원본 대비)
 *  - Slack API 응답(ok/error)을 확인해 조용한 실패(invalid_auth 등)를 로깅
 *  - SLACK_BOT_TOKEN 미설정 시 즉시 중단
 *  - 담당자별 발송을 격리(한 명 실패가 전체를 멈추지 않음)
 *  - #N/A 등 오류값/빈 날짜 방어
 *  - 설정(시트명/윈도우/컬럼)을 CONFIG로 분리
 *  - 발송/실패/미매핑 건수 요약 로깅
 */

// ── 설정 ────────────────────────────────────────────────
var WAITING_EXPIRY_CONFIG = {
  SHEET_NAME: '상품설정',
  WINDOW_DAYS: 14,
  KEYWORD: '웨이팅',
  TIMEZONE: 'GMT+9',

  // "상품설정" 탭 0-based 컬럼 인덱스 (2026-07-09 확정 스키마 기준)
  COL: {
    seq: 2,        // 매장시퀀스
    storeName: 3,  // 매장명
    adName: 5,     // 상품명-룩업 (광고명)
    endDate: 8,    // 광고종료일
    manager: 10,   // 계약_담당자
    // 참여구분(연속/단독) — "연속" 포함 여부로 연참 판단.
    // ※ 원본은 13(기타 상세)이었으나 확정 스키마상 참여구분은 14다.
    //    실제 시트가 다르면 이 값만 바꾸면 된다.
    type: 14
  }
};

// 담당자명 → Slack User ID
var SLACK_USER_IDS = {
  "이도은": "U093FJ7DZ8W", "최원영": "U0BLAHC00G3", "권용덕": "U0BLDGBS8G5",
  "남현욱": "U07RH97HNQL", "우은수": "U093FJ573FY", "이종익": "U0AGLUM2G2V",
  "김나현": "U0B99RC7H08", "홍성혁": "U02TN1U2PQR", "이하윤": "U0AJ3LN8E3T",
  "이조은": "U09GZ0H7928", "이동헌": "U0AG4JQJVKP", "김현수": "U0BHGG7FP98",
  "전평정": "U02QCTZT2PP", "이세한": "U09BZ6JL60G", "이혜민": "U0AJY0DSMPC",
  "김상하": "U0AG600HV1U", "이승민": "U0AG1KQQFS7", "한창완": "U057M7S5RA9",
  "신유빈": "U09MXM4BV71", "이지민": "U0BGDNGHUBC"
};

function notifyWaitingExpiryDMOnly() {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) {
    Logger.log('[중단] SLACK_BOT_TOKEN 스크립트 속성이 설정되지 않았습니다.');
    return;
  }

  var cfg = WAITING_EXPIRY_CONFIG;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.SHEET_NAME);
  if (!sheet) {
    Logger.log('[중단] "' + cfg.SHEET_NAME + '" 시트를 찾을 수 없습니다.');
    return;
  }

  var reportData = collectWaitingExpiries_(sheet, cfg);
  var managers = Object.keys(reportData);
  if (managers.length === 0) {
    Logger.log('종료 임박(' + cfg.WINDOW_DAYS + '일 이내) 웨이팅 매장이 없습니다. 발송 대상 없음.');
    return;
  }

  var sent = 0, failed = 0, unmapped = 0;
  managers.forEach(function (name) {
    var slackId = SLACK_USER_IDS[name];
    if (!slackId) {
      unmapped++;
      Logger.log('[건너뜀] SlackID 미매핑 담당자: ' + name + ' (' + reportData[name].length + '건)');
      return;
    }
    var message = buildWaitingExpiryMessage_(slackId, reportData[name], cfg);
    if (sendSlackDM_(token, slackId, message)) {
      sent++;
    } else {
      failed++;
    }
  });

  Logger.log('웨이팅 종료 알림 완료 — 발송 ' + sent + ' / 실패 ' + failed + ' / 미매핑 ' + unmapped);
}

/**
 * 상품설정 시트를 스캔해 담당자별 종료 임박 매장 리스트를 만든다.
 * @return {Object} { 담당자명: [ {name, seq, dateString, dDay, isContinuous}, ... ] }
 */
function collectWaitingExpiries_(sheet, cfg) {
  var data = sheet.getDataRange().getValues();
  var col = cfg.COL;

  var today = startOfDay_(new Date());
  var limit = startOfDay_(new Date());
  limit.setDate(limit.getDate() + cfg.WINDOW_DAYS);

  var report = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var adName = String(row[col.adName] || '').trim();
    if (adName.indexOf(cfg.KEYWORD) === -1) continue;

    var endDate = parseDate_(row[col.endDate]);
    if (!endDate) continue;
    endDate = startOfDay_(endDate);
    if (endDate < today || endDate > limit) continue;

    var manager = String(row[col.manager] || '담당자미지정').trim();
    if (!report[manager]) report[manager] = [];
    report[manager].push({
      name: String(row[col.storeName] || '').trim(),
      seq: row[col.seq],
      dateString: Utilities.formatDate(endDate, cfg.TIMEZONE, 'yyyy-MM-dd'),
      dDay: Math.round((endDate.getTime() - today.getTime()) / 86400000),
      isContinuous: String(row[col.type] || '').indexOf('연속') !== -1
    });
  }
  return report;
}

/** 담당자 1명에게 보낼 DM 본문을 만든다. */
function buildWaitingExpiryMessage_(slackId, stores, cfg) {
  var sorted = stores.slice().sort(function (a, b) { return a.dDay - b.dDay; });

  var lines = sorted.map(function (s) {
    var tag = s.isContinuous ? '🔥 ' : '   ';
    var ddayLabel = s.dDay <= 0 ? '오늘 종료' : 'D-' + s.dDay;
    return tag + s.name + '(' + s.seq + ') · ' + ddayLabel + ' (' + s.dateString + ' 종료)';
  }).join('\n');

  var hasContinuous = sorted.some(function (s) { return s.isContinuous; });
  var weeks = Math.round(cfg.WINDOW_DAYS / 7);

  var msg = '안녕하세요 <@' + slackId + '>님, 좋은 아침이에요 ☀️\n';
  msg += '담당하시는 웨이팅 매장 중 광고 종료가 ' + weeks + '주 내로 다가온 곳이 ' + sorted.length + '곳 있어요.\n\n';
  msg += lines + '\n\n';
  msg += hasContinuous
    ? '🔥 표시된 연참 매장은 재계약 타이밍 놓치지 않도록 특히 빠른 컨택 부탁드려요! 📞'
    : '종료 임박 매장, 빠른 컨택 부탁드려요! 📞';
  return msg;
}

/**
 * 슬랙 개인 DM 전송. 성공 여부(Boolean)를 반환한다.
 * chat.postMessage는 실패해도 HTTP 200 + {ok:false}를 주므로 응답 본문을 확인한다.
 */
function sendSlackDM_(token, userId, text) {
  var response;
  try {
    response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ channel: userId, text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('[실패] Slack 요청 오류 (' + userId + '): ' + e.message);
    return false;
  }

  var body = {};
  try { body = JSON.parse(response.getContentText()); } catch (e) { /* ignore */ }

  if (!body.ok) {
    Logger.log('[실패] Slack API 오류 (' + userId + '): ' + (body.error || response.getContentText()));
    return false;
  }
  return true;
}

// ── 유틸 ────────────────────────────────────────────────

/** 시각을 00:00:00으로 맞춘다(원본 Date를 변경 후 반환). */
function startOfDay_(d) {
  d.setHours(0, 0, 0, 0);
  return d;
}

/** 셀 값을 Date로 변환. Date 객체/문자열 모두 처리하고 #N/A 등 오류값은 null. */
function parseDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }
  var s = String(value == null ? '' : value).trim();
  if (!s || s.charAt(0) === '#') return null; // '', '#N/A', '#REF!' 등 제외
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
