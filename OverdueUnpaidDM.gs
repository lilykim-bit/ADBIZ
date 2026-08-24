/**
 * 입금 마감일 경과 미완료 매장 담당자 개인 1:1 DM 알림봇
 *
 * 대상: 계약현황 탭에서
 *   - 계약상태 = "계약서 서명 완료"
 *   - 결제일 공란 (결제 미완료)
 *   - 종료일이 MIN_END_YEAR(2026)년 이후
 *   - 입금 마감일이 오늘보다 과거
 *   - 계약대금 > 0
 * 위 조건을 모두 만족하는 건을 영업담당자별로 묶어 개인 DM 발송.
 *
 * [중요] Apps Script는 프로젝트의 모든 .gs 파일이 하나의 전역 스코프로 합쳐진다.
 * 같은 프로젝트의 알람자동화.gs가 const CONFIG / STATE_HEADERS / AD_TYPES 등을
 * 선언하므로, 이 파일의 전역 식별자는 모두 OVERDUE_DM_ / overdue 접두어를 붙여
 * 이름 충돌(SyntaxError: Identifier 'CONFIG' has already been declared)을 피한다.
 * 이 파일에 전역 변수/함수를 추가할 때도 같은 규칙을 지킬 것.
 */

// ===== 설정 =====
var OVERDUE_DM_CONFIG = {
  SPREADSHEET_ID: "1TIjXcv7E7QQNcEhoRrKKUJwJt9ftlumKUuVQSs0bD8Y",
  SHEET_NAME: "계약현황",
  MIN_END_YEAR: 2026,        // 종료일이 이 연도 이상인 건만 대상

  // 한 DM에 담을 최대 글자수. 초과하면 (1/2), (2/2) 형태로 분할 발송.
  // 리스트를 잘라내지 않고 나눠 보내므로 누락되는 매장이 없음.
  MAX_CHARS_PER_DM: 3000,

  // 테스트 모드: true면 실제 담당자에게 보내지 않고 TEST_TARGET_SLACK_ID 한 명에게만 발송
  TEST_MODE: false,
  TEST_TARGET_SLACK_ID: "U07RBU2TKNH",  // 김승현

  SEND_INTERVAL_MS: 300      // Slack rate limit 대비 발송 간 대기
};

// 슬랙 유저 ID 매핑 명부
var OVERDUE_DM_SLACK_USER_IDS = {
  "김나현": "U0B99RC7H08", "김상하": "U0AG600HV1U", "김연아": "U0BN7S7C6TH", "김현수": "U0BHGG7FP98",
  "남윤석": "U0BQ3UL4Z5L", "남현욱": "U07RH97HNQL", "신유빈": "U09MXM4BV71", "우은수": "U093FJ573FY",
  "이도은": "U093FJ7DZ8W", "이세한": "U09BZ6JL60G", "이승준": "U09E3L1KFQR", "이조은": "U09GZ0H7928",
  "이종익": "U0AGLUM2G2V", "이하윤": "U0AJ3LN8E3T", "이혜민": "U0AJY0DSMPC", "전평정": "U02QCTZT2PP",
  "최원영": "U0BLAHC00G3", "한창완": "U057M7S5RA9", "홍성혁": "U02TN1U2PQR"
};

// 계약현황 탭 컬럼 인덱스 (0-based)
var OVERDUE_DM_COL = {
  seq: 2,          // C열: 계약_매장시퀀스
  storeName: 3,    // D열: 계약_매장명
  adName: 5,       // F열: 계약_광고명
  manager: 6,      // G열: 계약_담당자 (영업담당자)
  endDate: 8,      // I열: 종료일
  dueDate: 10,     // K열: 입금 마감일
  price: 11,       // L열: 계약대금
  signStatus: 25,  // Z열: 계약상태
  payStatus: 37    // AL열: 결제일
};

// ===== 메인 =====

/** 실제 발송 (트리거 대상 함수) */
function notifyOverdueUnpaidStoresDedicatedDM() {
  return runOverdueNotification_(false);
}

/** 발송 없이 로그로만 결과 확인 (드라이런) */
function previewOverdueUnpaidDMs() {
  return runOverdueNotification_(true);
}

function runOverdueNotification_(dryRun) {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!dryRun && !token) {
    throw new Error("SLACK_BOT_TOKEN 스크립트 속성이 설정되지 않았습니다. " +
                    "프로젝트 설정 > 스크립트 속성에 xoxb- 로 시작하는 토큰을 등록해주세요.");
  }

  var report = collectOverdueItems_();
  var sent = 0, failed = 0;

  for (var managerName in report.byManager) {
    var entry = report.byManager[managerName];
    var messages = overdueBuildMessages_(managerName, entry);
    var target = OVERDUE_DM_CONFIG.TEST_MODE ? OVERDUE_DM_CONFIG.TEST_TARGET_SLACK_ID : entry.slackId;

    for (var m = 0; m < messages.length; m++) {
      if (dryRun) {
        Logger.log("[DRY RUN] → " + managerName + " (" + target + ")\n" + messages[m]);
        sent++;
        continue;
      }
      var ok = sendOverdueSlackDM_(token, target, messages[m]);
      ok ? sent++ : failed++;
      if (m < messages.length - 1 || OVERDUE_DM_CONFIG.SEND_INTERVAL_MS) {
        Utilities.sleep(OVERDUE_DM_CONFIG.SEND_INTERVAL_MS);
      }
    }
  }

  var summary = "대상 " + report.totalItems + "건 / 담당자 " +
                Object.keys(report.byManager).length + "명 / DM " + sent + "건 발송" +
                (failed ? " (실패 " + failed + "건)" : "") +
                (dryRun ? " [DRY RUN]" : "") +
                (OVERDUE_DM_CONFIG.TEST_MODE ? " [TEST_MODE]" : "");
  Logger.log(summary);

  // 슬랙 ID 미등록 담당자는 조용히 빠지면 영구 누락되므로 반드시 남긴다
  var unmapped = Object.keys(report.unmapped);
  if (unmapped.length > 0) {
    Logger.log("⚠️ OVERDUE_DM_SLACK_USER_IDS 미등록으로 발송 제외된 담당자: " +
      unmapped.map(function(n) { return n + "(" + report.unmapped[n] + "건)"; }).join(", "));
  }

  return summary;
}

/** 시트를 스캔해 담당자별 미입금 건을 취합 */
function collectOverdueItems_() {
  var sheet = SpreadsheetApp.openById(OVERDUE_DM_CONFIG.SPREADSHEET_ID).getSheetByName(OVERDUE_DM_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("'" + OVERDUE_DM_CONFIG.SHEET_NAME + "' 탭을 찾을 수 없습니다.");

  var data = sheet.getDataRange().getValues();
  var today = overdueStartOfDay_(new Date());
  var byManager = {};
  var unmapped = {};
  var totalItems = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[OVERDUE_DM_COL.adName]) continue;

    // [필터 1] 서명 완료 / 결제 미완료
    if (String(row[OVERDUE_DM_COL.signStatus] || "").trim() !== "계약서 서명 완료") continue;
    if (String(row[OVERDUE_DM_COL.payStatus] || "").trim() !== "") continue;

    // [필터 2] 종료일이 MIN_END_YEAR년 이후
    var endDate = overdueParseDate_(row[OVERDUE_DM_COL.endDate]);
    if (!endDate || endDate.getFullYear() < OVERDUE_DM_CONFIG.MIN_END_YEAR) continue;

    // [필터 3] 입금 마감일이 오늘 이전(과거)
    var dueDate = overdueParseDate_(row[OVERDUE_DM_COL.dueDate]);
    if (!dueDate || dueDate.getTime() >= today.getTime()) continue;

    // [필터 4] 계약대금 0원 제외
    var contractPrice = overdueToNumber_(row[OVERDUE_DM_COL.price]);
    if (contractPrice <= 0) continue;

    var managerName = String(row[OVERDUE_DM_COL.manager] || "").trim() || "담당자미지정";
    var slackId = OVERDUE_DM_SLACK_USER_IDS[managerName];
    if (!slackId) {
      unmapped[managerName] = (unmapped[managerName] || 0) + 1;
      continue;
    }

    if (!byManager[managerName]) {
      byManager[managerName] = { slackId: slackId, ads: {}, count: 0, amount: 0 };
    }
    var manager = byManager[managerName];

    var adTitle = String(row[OVERDUE_DM_COL.adName]).trim().split('_')[0].trim();
    if (!manager.ads[adTitle]) manager.ads[adTitle] = [];
    manager.ads[adTitle].push({
      storeName: String(row[OVERDUE_DM_COL.storeName] || "").trim(),
      seq: row[OVERDUE_DM_COL.seq],
      price: contractPrice,
      dueDate: dueDate,
      overdueDays: Math.round((today.getTime() - dueDate.getTime()) / 86400000)
    });

    manager.count++;
    manager.amount += contractPrice;
    totalItems++;
  }

  return { byManager: byManager, unmapped: unmapped, totalItems: totalItems };
}

/** 담당자 1명의 DM 본문 생성. 길면 여러 개로 분할해서 배열로 반환 */
function overdueBuildMessages_(managerName, entry) {
  // 광고 상품별 섹션: 연체가 오래된 건이 위로 오도록 정렬
  var adTitles = Object.keys(entry.ads).sort();
  var sections = adTitles.map(function(adTitle) {
    var stores = entry.ads[adTitle].sort(function(a, b) { return b.overdueDays - a.overdueDays; });
    var amount = stores.reduce(function(sum, s) { return sum + s.price; }, 0);

    var lines = ["📦 *[" + adTitle + "]* " + stores.length + "건 · " + overdueWon_(amount)];
    stores.forEach(function(s) {
      lines.push("  • " + s.storeName + "(" + s.seq + ") — " + overdueWon_(s.price) +
                 " · 마감 " + overdueMmdd_(s.dueDate) + " (D+" + s.overdueDays + ")");
    });
    return lines.join("\n");
  });

  function header(part) {
    return "📢 *안녕하세요 " + managerName + "님, 입금 마감일 경과 미입금 알림입니다.*" + part + "\n\n" +
           "🔴 *담당 매장 중 서명 완료 / 결제 미완료 " + entry.count + "건 · " + overdueWon_(entry.amount) + "*\n" +
           "_입금 마감일이 지났으나 아직 입금이 확인되지 않은 매장 리스트입니다._\n";
  }

  var footer = "⚠️ *안내사항*\n" +
               "• 해당 내용은 앱시트 기반으로 가져오는 데이터입니다.\n" +
               "• 앱시트 상에서는 미수로 확인되니, *이미 입금이 완료되었다면 입금일자와 주문번호를 반드시 입력*해 주시기 바랍니다 🙏";

  // 섹션을 글자수 기준으로 묶기 (매장을 잘라내지 않고 메시지를 나눔)
  var chunks = [];
  var current = [];
  var currentLen = 0;
  sections.forEach(function(section) {
    if (current.length > 0 && currentLen + section.length > OVERDUE_DM_CONFIG.MAX_CHARS_PER_DM) {
      chunks.push(current);
      current = [];
      currentLen = 0;
    }
    current.push(section);
    currentLen += section.length + 2;
  });
  if (current.length > 0) chunks.push(current);

  return chunks.map(function(chunk, idx) {
    var part = chunks.length > 1 ? " (" + (idx + 1) + "/" + chunks.length + ")" : "";
    var body = header(part) + "\n" + chunk.join("\n\n");
    return idx === chunks.length - 1 ? body + "\n\n" + footer : body;
  });
}

// ===== 유틸 =====

/** 시트 값(Date 객체 / "2026-03-05" / "2026. 3. 5 오전 12:00:00" 등)을 당일 0시 Date로 변환 */
function overdueParseDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : overdueStartOfDay_(value);
  }
  var text = String(value || "").trim();
  if (!text) return null;

  // 연-월-일 숫자만 추출 → 로컬 타임존 기준으로 생성 (문자열 파싱의 UTC 해석 문제 회피)
  var m = text.match(/(\d{4})\s*[-.\/]\s*(\d{1,2})\s*[-.\/]\s*(\d{1,2})/);
  if (!m) return null;
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function overdueStartOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function overdueToNumber_(value) {
  if (typeof value === "number") return value;
  var n = Number(String(value || "").replace(/[,\s원]/g, ""));
  return isNaN(n) ? 0 : n;
}

function overdueWon_(amount) {
  return String(Math.round(amount)).replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "원";
}

function overdueMmdd_(date) {
  return (date.getMonth() + 1) + "/" + date.getDate();
}

/**
 * 슬랙 API를 이용해 특정 유저 ID로 1:1 DM 발송.
 * 성공 여부를 boolean으로 반환 (한 명 실패해도 나머지 발송이 중단되지 않게 예외를 흡수)
 */
function sendOverdueSlackDM_(token, userId, text) {
  var options = {
    "method": "post",
    "contentType": "application/json; charset=utf-8",
    "headers": { "Authorization": "Bearer " + token },
    "payload": JSON.stringify({ "channel": userId, "text": text }),
    "muteHttpExceptions": true
  };

  for (var attempt = 0; attempt < 3; attempt++) {
    try {
      var response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
      var code = response.getResponseCode();

      if (code === 429) {  // rate limited
        var retryAfter = Number(response.getHeaders()['Retry-After'] || 1);
        Utilities.sleep((retryAfter + 1) * 1000);
        continue;
      }

      var resJson = JSON.parse(response.getContentText());
      if (resJson.ok) return true;
      Logger.log("슬랙 전송 실패 (" + userId + "): " + resJson.error);
      return false;   // invalid_auth 등은 재시도해도 동일하므로 즉시 중단
    } catch (e) {
      Logger.log("슬랙 전송 예외 (" + userId + ", 시도 " + (attempt + 1) + "): " + e.message);
      Utilities.sleep(1000 * (attempt + 1));
    }
  }
  return false;
}
