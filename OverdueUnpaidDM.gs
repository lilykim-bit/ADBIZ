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
  // 미러링 시트 ([광고비즈니스] 고가광고 세일즈 헬퍼) — IMPORTRANGE로 원본을 끌어온 사본
  SPREADSHEET_ID: "1TIjXcv7E7QQNcEhoRrKKUJwJt9ftlumKUuVQSs0bD8Y",
  // 원본 전사광고영업디비
  SOURCE_SPREADSHEET_ID: "1mq56pdu9hplHv1qjWT1Yf0xLYopsuRsG2CZ8ouTH8aA",
  // true면 미러링 시트를 건너뛰고 원본에서 직접 읽는다.
  // 미러링 IMPORTRANGE가 깨지거나(#VALUE!/#REF!) 범위 갱신을 놓쳐도 영향받지 않음.
  USE_SOURCE_DIRECTLY: false,

  SHEET_NAME: "계약현황",
  MIN_END_YEAR: 2026,        // 종료일이 이 연도 이상인 건만 대상

  // 한 DM에 담을 최대 글자수. 초과하면 (1/2), (2/2) 형태로 분할 발송.
  // 리스트를 잘라내지 않고 나눠 보내므로 누락되는 매장이 없음.
  MAX_CHARS_PER_DM: 3000,

  // 슬랙 ID를 못 찾은 담당자(퇴사자/외부 인력 등)의 건을 모아 보낼 수신자.
  // 담당자가 없는 미수금은 팀장이 인수해야 하므로 팀장에게 전달한다.
  // 비워두면 해당 건은 아무에게도 전달되지 않고 로그에만 남는다.
  FALLBACK_SLACK_ID: "U0B2CNXM7LN",  // 김지운 팀장 (Business Growth Leader)

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
  "최원영": "U0BLAHC00G3", "한창완": "U057M7S5RA9", "홍성혁": "U02TN1U2PQR",
  "김이슬": "U09ACJEAJSJ", "조완수": "U027RCFP55W", "이지민": "U0BGDNGHUBC"
};

// 계약현황 탭 컬럼.
// 헤더명으로 찾는 것이 원칙이고, 헤더를 못 찾으면 fallback 인덱스(0-based)를 쓴다.
// 미러링 시트의 IMPORTRANGE 범위가 바뀌어 컬럼이 밀리거나 잘려도
// 조용히 0건이 되지 않고 경고/에러로 드러나게 하는 것이 목적.
var OVERDUE_DM_COL_SPEC = {
  seq:        { header: "계약_매장시퀀스", fallback: 2  },  // C열
  storeName:  { header: "계약_매장명",     fallback: 3  },  // D열
  adName:     { header: "계약_광고명",     fallback: 5  },  // F열
  manager:    { header: "계약_담당자",     fallback: 6  },  // G열
  endDate:    { header: "종료일",          fallback: 8  },  // I열
  dueDate:    { header: "입금 마감일",     fallback: 10 },  // K열
  price:      { header: "계약대금",        fallback: 11 },  // L열
  signStatus: { header: "계약상태",        fallback: 25 },  // Z열
  payStatus:  { header: "결제일",          fallback: 37 }   // AL열
};

// 계약상태가 이 값이면 서명 완료로 본다 (공백 제거 후 비교)
var OVERDUE_DM_SIGNED_VALUE = "계약서 서명 완료";

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

  // 슬랙 ID 없는 담당자 건은 폴백 수신자에게 모아서 전달 (조용히 사라지지 않게)
  var unmappedNames = Object.keys(report.unmapped);
  if (unmappedNames.length > 0 && OVERDUE_DM_CONFIG.FALLBACK_SLACK_ID) {
    var fallbackMsgs = overdueBuildUnmappedMessages_(report.unmapped);
    for (var f = 0; f < fallbackMsgs.length; f++) {
      if (dryRun) {
        Logger.log("[DRY RUN] → 폴백(" + OVERDUE_DM_CONFIG.FALLBACK_SLACK_ID + ")\n" + fallbackMsgs[f]);
        sent++;
      } else {
        sendOverdueSlackDM_(token, OVERDUE_DM_CONFIG.FALLBACK_SLACK_ID, fallbackMsgs[f]) ? sent++ : failed++;
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

  // 0건이면 "대상 0건"만 보고 끝내지 않고 어느 필터에서 죽었는지 같이 남긴다
  if (report.totalItems === 0) {
    Logger.log("⚠️ 대상 0건 — 단계별 탈락 내역:");
    Object.keys(report.funnel).forEach(function(k) { Logger.log("   " + k + ": " + report.funnel[k]); });
    Logger.log("   계약상태 실제 값 분포: " + JSON.stringify(report.samples.계약상태값));
    if (report.fellBack.length > 0) Logger.log("   ⚠️ 고정 인덱스로 대체된 컬럼: " + report.fellBack.join(", "));
    Logger.log("   자세한 진단은 diagnoseOverdueFilters() 실행");
  }

  // 슬랙 ID 미등록 담당자는 조용히 빠지면 영구 누락되므로 반드시 남긴다
  var unmapped = Object.keys(report.unmapped);
  if (unmapped.length > 0) {
    Logger.log("⚠️ OVERDUE_DM_SLACK_USER_IDS 미등록 담당자: " +
      unmapped.map(function(n) { return n + "(" + report.unmapped[n].count + "건)"; }).join(", "));
  }

  return summary;
}

// 시트에 나타나는 수식 오류값
var OVERDUE_DM_ERROR_VALUES = ["#VALUE!", "#REF!", "#N/A", "#NAME?", "#DIV/0!", "#ERROR!", "#NULL!", "#NUM!"];

/** 설정에 따라 미러링 시트 또는 원본에서 계약현황 탭을 연다 */
function overdueOpenContractSheet_() {
  var useSource = OVERDUE_DM_CONFIG.USE_SOURCE_DIRECTLY;
  var id = useSource ? OVERDUE_DM_CONFIG.SOURCE_SPREADSHEET_ID : OVERDUE_DM_CONFIG.SPREADSHEET_ID;
  var label = useSource ? "원본 전사광고영업디비" : "미러링 시트";

  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(label + "(" + id + ")를 열 수 없습니다. 열람 권한을 확인해주세요: " + e.message);
  }

  var sheet = ss.getSheetByName(OVERDUE_DM_CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error(label + "에 '" + OVERDUE_DM_CONFIG.SHEET_NAME + "' 탭이 없습니다. 존재하는 탭: " +
      ss.getSheets().map(function(sh) { return sh.getName(); }).join(", "));
  }
  return { sheet: sheet, label: label, id: id };
}

/** 헤더 행에 수식 오류값이 있으면 그 값을 반환 (없으면 null) */
function overdueFindErrorValue_(headers) {
  for (var i = 0; i < headers.length; i++) {
    var v = String(headers[i] == null ? "" : headers[i]).trim();
    if (OVERDUE_DM_ERROR_VALUES.indexOf(v) >= 0) return v;
  }
  return null;
}

/** 헤더 행에서 컬럼 위치를 해석 */
function overdueResolveColumns_(headers) {
  var norm = function(v) { return String(v == null ? "" : v).replace(/\s+/g, ""); };
  var byName = {};
  for (var i = 0; i < headers.length; i++) {
    var k = norm(headers[i]);
    if (k && !(k in byName)) byName[k] = i;
  }

  var col = {}, missing = [], fellBack = [];
  Object.keys(OVERDUE_DM_COL_SPEC).forEach(function(key) {
    var spec = OVERDUE_DM_COL_SPEC[key];
    var k = norm(spec.header);
    if (k in byName) { col[key] = byName[k]; return; }
    if (spec.fallback < headers.length) {
      col[key] = spec.fallback;
      fellBack.push(spec.header + "→" + spec.fallback + "번째열('" + headers[spec.fallback] + "')");
      return;
    }
    col[key] = -1;
    missing.push(spec.header);
  });
  return { col: col, missing: missing, fellBack: fellBack };
}

/**
 * 시트를 스캔해 담당자별 미입금 건을 취합.
 * 단계별 탈락 건수(funnel)와 실제 값 샘플을 함께 반환해서
 * "대상 0건"이 나왔을 때 어느 필터에서 죽었는지 바로 알 수 있게 한다.
 */
function collectOverdueItems_() {
  var opened = overdueOpenContractSheet_();
  var sheet = opened.sheet;

  var data = sheet.getDataRange().getValues();
  var shape = opened.label + " '" + OVERDUE_DM_CONFIG.SHEET_NAME + "' 탭: " +
              data.length + "행 x " + (data.length ? data[0].length : 0) + "열";
  if (data.length < 2) throw new Error(shape + " — 데이터 행이 없습니다.");

  var headers = data[0];

  // 미러링 IMPORTRANGE가 깨진 경우: 필터 문제가 아니라 데이터 자체가 안 들어온 것이므로
  // 원인(수식 오류값 + A1 수식)을 그대로 드러낸다.
  var errorValue = overdueFindErrorValue_(headers);
  if (errorValue) {
    throw new Error(
      shape + " — 헤더 행이 " + errorValue + " 입니다. IMPORTRANGE 수식이 실패해서 데이터가 " +
      "들어오지 않은 상태이므로 스크립트가 잡을 대상이 없습니다.\n" +
      "A1 수식: " + (sheet.getRange(1, 1).getFormula() || "(수식 없음)") + "\n" +
      "→ 시트 수식을 고치거나, OVERDUE_DM_CONFIG.USE_SOURCE_DIRECTLY = true 로 두고 " +
      "원본 전사광고영업디비에서 직접 읽으세요.");
  }

  var resolved = overdueResolveColumns_(headers);
  var col = resolved.col;
  if (resolved.missing.length > 0) {
    throw new Error(shape + " — 다음 컬럼을 찾을 수 없습니다: " + resolved.missing.join(", ") +
      " / 실제 헤더=[" + headers.join(" | ") + "]");
  }

  var today = overdueStartOfDay_(new Date());
  var byManager = {};
  var unmapped = {};
  var totalItems = 0;

  // 진단용 집계
  var funnel = { 전체행: data.length - 1, 광고명공란: 0, 서명미완료: 0, 결제일있음: 0,
                 종료일파싱실패: 0, 종료일연도미달: 0, 마감일파싱실패: 0,
                 마감일미래: 0, 금액0: 0, 슬랙ID미등록: 0, 최종대상: 0 };
  var samples = { 계약상태값: {}, 결제일값: [], 종료일원본: [], 마감일원본: [] };
  var tally = function(map, value) {
    var k = "[" + String(value) + "]";
    if (Object.keys(map).length < 25 || (k in map)) map[k] = (map[k] || 0) + 1;
  };
  var keep = function(arr, value) {
    if (arr.length < 5) arr.push("[" + String(value) + "]");
  };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col.adName]) { funnel.광고명공란++; continue; }

    // [필터 1] 서명 완료 / 결제 미완료
    var signStatus = String(row[col.signStatus] || "").trim();
    tally(samples.계약상태값, signStatus);
    if (signStatus !== OVERDUE_DM_SIGNED_VALUE) { funnel.서명미완료++; continue; }

    var payStatus = String(row[col.payStatus] || "").trim();
    if (payStatus !== "") { funnel.결제일있음++; keep(samples.결제일값, payStatus); continue; }

    // [필터 2] 종료일이 MIN_END_YEAR년 이후
    var endDate = overdueParseDate_(row[col.endDate]);
    if (!endDate) { funnel.종료일파싱실패++; keep(samples.종료일원본, row[col.endDate]); continue; }
    if (endDate.getFullYear() < OVERDUE_DM_CONFIG.MIN_END_YEAR) { funnel.종료일연도미달++; continue; }

    // [필터 3] 입금 마감일이 오늘 이전(과거)
    var dueDate = overdueParseDate_(row[col.dueDate]);
    if (!dueDate) { funnel.마감일파싱실패++; keep(samples.마감일원본, row[col.dueDate]); continue; }
    if (dueDate.getTime() >= today.getTime()) { funnel.마감일미래++; continue; }

    // [필터 4] 계약대금 0원 제외
    var contractPrice = overdueToNumber_(row[col.price]);
    if (contractPrice <= 0) { funnel.금액0++; continue; }

    var managerName = String(row[col.manager] || "").trim() || "담당자미지정";
    var slackId = OVERDUE_DM_SLACK_USER_IDS[managerName];

    // 슬랙 ID가 없어도 건수만 세고 버리지 않는다. 상세를 모아 폴백 수신자에게 전달한다.
    var bucket = slackId ? byManager : unmapped;
    if (!bucket[managerName]) {
      bucket[managerName] = { slackId: slackId || null, ads: {}, count: 0, amount: 0 };
    }
    var manager = bucket[managerName];

    var adTitle = String(row[col.adName]).trim().split('_')[0].trim();
    if (!manager.ads[adTitle]) manager.ads[adTitle] = [];
    manager.ads[adTitle].push({
      storeName: String(row[col.storeName] || "").trim(),
      seq: row[col.seq],
      price: contractPrice,
      dueDate: dueDate,
      overdueDays: Math.round((today.getTime() - dueDate.getTime()) / 86400000)
    });

    manager.count++;
    manager.amount += contractPrice;

    if (!slackId) { funnel.슬랙ID미등록++; continue; }
    totalItems++;
  }
  funnel.최종대상 = totalItems;

  return {
    byManager: byManager, unmapped: unmapped, totalItems: totalItems,
    funnel: funnel, samples: samples, headers: headers,
    resolvedCol: col, fellBack: resolved.fellBack
  };
}

/**
 * 데이터 소스 점검: 미러링 시트와 원본을 각각 열어보고
 * 탭 목록 / 크기 / 계약현황 A1 수식 / 헤더 상태를 로그로 남긴다.
 * "대상 0건"이 필터 문제인지 시트 문제인지 가리는 데 사용.
 */
function diagnoseSourceSheet() {
  var lines = ["=== 데이터 소스 점검 ==="];

  [{ id: OVERDUE_DM_CONFIG.SPREADSHEET_ID, label: "미러링 시트" },
   { id: OVERDUE_DM_CONFIG.SOURCE_SPREADSHEET_ID, label: "원본 전사광고영업디비" }].forEach(function(target) {
    lines.push("");
    lines.push("[" + target.label + "] " + target.id);
    var ss;
    try {
      ss = SpreadsheetApp.openById(target.id);
    } catch (e) {
      lines.push("  ❌ 열 수 없음 (열람 권한 확인 필요): " + e.message);
      return;
    }
    lines.push("  파일명: " + ss.getName());
    lines.push("  탭 목록: " + ss.getSheets().map(function(sh) {
      return sh.getName() + "(" + sh.getLastRow() + "x" + sh.getLastColumn() + ")";
    }).join(", "));

    var sheet = ss.getSheetByName(OVERDUE_DM_CONFIG.SHEET_NAME);
    if (!sheet) { lines.push("  ❌ '" + OVERDUE_DM_CONFIG.SHEET_NAME + "' 탭 없음"); return; }

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var errorValue = overdueFindErrorValue_(headers);
    lines.push("  A1 수식: " + (sheet.getRange(1, 1).getFormula() || "(수식 없음 / 값 직접 입력)"));
    lines.push("  헤더(" + headers.length + "개): [" + headers.join(" | ") + "]");
    if (errorValue) {
      lines.push("  ❌ 헤더가 " + errorValue + " — 수식이 실패해 데이터가 들어오지 않았습니다.");
    } else {
      var resolved = overdueResolveColumns_(headers);
      lines.push(resolved.missing.length
        ? "  ❌ 누락 컬럼: " + resolved.missing.join(", ")
        : "  ✅ 필요한 컬럼 모두 확인 (" + JSON.stringify(resolved.col) + ")");
    }
  });

  var out = lines.join("\n");
  Logger.log(out);
  return out;
}

/** 왜 대상이 0건인지(혹은 몇 건인지) 단계별로 로그 출력 — 발송 안 함 */
function diagnoseOverdueFilters() {
  var report = collectOverdueItems_();
  var lines = [];

  lines.push("=== 계약현황 스캔 진단 ===");
  lines.push("시트 컬럼 수: " + report.headers.length);
  lines.push("해석된 컬럼 위치: " + JSON.stringify(report.resolvedCol));
  if (report.fellBack.length > 0) {
    lines.push("⚠️ 헤더명으로 못 찾아 고정 인덱스로 대체한 컬럼: " + report.fellBack.join(", "));
  }
  lines.push("");
  lines.push("--- 단계별 탈락 건수 ---");
  Object.keys(report.funnel).forEach(function(k) { lines.push("  " + k + ": " + report.funnel[k]); });
  lines.push("");
  lines.push("--- 계약상태 실제 값 분포 (기대값: '" + OVERDUE_DM_SIGNED_VALUE + "') ---");
  Object.keys(report.samples.계약상태값).forEach(function(k) {
    lines.push("  " + k + " : " + report.samples.계약상태값[k] + "건");
  });
  if (report.samples.종료일원본.length) lines.push("종료일 파싱실패 샘플: " + report.samples.종료일원본.join(", "));
  if (report.samples.마감일원본.length) lines.push("마감일 파싱실패 샘플: " + report.samples.마감일원본.join(", "));
  if (report.samples.결제일값.length) lines.push("결제일 값 샘플(=결제완료로 제외됨): " + report.samples.결제일값.join(", "));

  var unmapped = Object.keys(report.unmapped);
  if (unmapped.length) {
    lines.push("슬랙ID 미등록 담당자: " + unmapped.map(function(n) {
      return n + "(" + report.unmapped[n].count + "건)"; }).join(", "));
  }
  lines.push("");
  lines.push("최종 발송 대상: " + report.totalItems + "건 / 담당자 " + Object.keys(report.byManager).length + "명");

  var out = lines.join("\n");
  Logger.log(out);
  return out;
}

/** 슬랙 ID를 못 찾은 담당자들의 건을 폴백 수신자용으로 묶는다 */
function overdueBuildUnmappedMessages_(unmapped) {
  var names = Object.keys(unmapped).sort(function(a, b) {
    return unmapped[b].count - unmapped[a].count;
  });
  var totalCount = 0, totalAmount = 0;
  names.forEach(function(n) { totalCount += unmapped[n].count; totalAmount += unmapped[n].amount; });

  var sections = names.map(function(name) {
    var entry = unmapped[name];
    var lines = ["👤 *" + name + "* " + entry.count + "건 · " + overdueWon_(entry.amount)];
    Object.keys(entry.ads).sort().forEach(function(adTitle) {
      var stores = entry.ads[adTitle].sort(function(a, b) { return b.overdueDays - a.overdueDays; });
      lines.push("  📦 [" + adTitle + "]");
      stores.forEach(function(st) {
        lines.push("    • " + st.storeName + "(" + st.seq + ") — " + overdueWon_(st.price) +
                   " · 마감 " + overdueMmdd_(st.dueDate) + " (D+" + st.overdueDays + ")");
      });
    });
    return lines.join("\n");
  });

  function header(part) {
    return "🚨 *퇴사자 미입금 알림*" + part + "\n\n" +
           "🔴 *" + names.length + "명 / " + totalCount + "건 · " + overdueWon_(totalAmount) + "*\n" +
           "아래 건은 퇴사자가 입금 팔로우업을 하지 않고 간 건입니다. 확인 부탁드립니다.\n";
  }
  var footer = "⚠️ 담당자가 재직 중인데 목록에 올라왔다면 슬랙 ID 매핑이 누락된 경우이니 알려주세요.\n" +
               "(스크립트의 `OVERDUE_DM_SLACK_USER_IDS`에 추가하면 담당자 본인에게 직접 발송됩니다.)";

  var chunks = [], current = [], currentLen = 0;
  sections.forEach(function(section) {
    if (current.length > 0 && currentLen + section.length > OVERDUE_DM_CONFIG.MAX_CHARS_PER_DM) {
      chunks.push(current); current = []; currentLen = 0;
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
