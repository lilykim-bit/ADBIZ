/**
 * 광고 성과 데이터 통합 (최적화 버전)
 * - getLastRow()로 실제 데이터 범위만 로드
 * - 날짜를 ISO 문자열로 통일하여 키 생성 비용 절감
 * - 최근 N일 데이터만 처리하는 옵션 추가
 * - 외부 시트 접근을 변수로 캐싱
 */

/** 설정 */
var CONFIG = {
  PRICE_SHEET_ID: "1YrYqZ-spGgK0ywb_moPlT-Ah_plSSUhlNysx65Yl_g0",
  LOG_SHEET_ID: "12beX3xFWni4bRILjQ8C_zDLGZ-fr3pCB-wkUmODF2BU",
  TARGET_SHEET_NAME: "RAW",
  SYNC_DAYS: 30  // 최근 30일 데이터만 동기화 (0이면 전체)
};

function runAdvertisingDataSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);

  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("'" + CONFIG.TARGET_SHEET_NAME + "' 시트를 찾을 수 없습니다.");
    return;
  }

  try {
    // 1. 기존 RAW 데이터 로드 — 실제 행만 읽기
    var existingKeys = buildExistingKeySet_(targetSheet);

    // 2. 상품명 정규화 맵 구성
    var productMap = buildProductMap_();

    // 3. 로그 데이터에서 신규 행 추출
    var newRows = extractNewRows_(productMap, existingKeys);

    // 4. 결과 기록
    if (newRows.length > 0) {
      targetSheet
        .getRange(targetSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
        .setValues(newRows);
      SpreadsheetApp.getUi().alert("업데이트 완료: 신규 데이터 " + newRows.length + "건이 추가되었습니다.");
    } else {
      SpreadsheetApp.getUi().alert("추가할 데이터가 없습니다. (모든 데이터가 이미 존재함)");
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert("오류: " + e.message + "\n" + e.stack);
  }
}

/**
 * RAW 시트의 기존 데이터로부터 중복 체크용 Set 생성
 * getDataRange() 대신 getLastRow()로 실제 범위만 읽음
 */
function buildExistingKeySet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  // 필요한 열만 로드: A(날짜), B(광고주), C(상품명), E(금액)
  var colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var colC = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var colE = sheet.getRange(2, 5, lastRow - 1, 1).getValues();

  var keys = new Set();
  for (var i = 0; i < colA.length; i++) {
    keys.add(makeKey_(colA[i][0], colB[i][0], colC[i][0], colE[i][0]));
  }
  return keys;
}

/**
 * 객단가 시트에서 상품명 정규화 맵 구성
 */
function buildProductMap_() {
  var priceSheet = SpreadsheetApp.openById(CONFIG.PRICE_SHEET_ID).getSheets()[0];
  var lastRow = priceSheet.getLastRow();
  if (lastRow < 2) return new Map();

  // A열(원본), B열(정규화명)만 로드
  var data = priceSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = new Map();
  for (var i = 0; i < data.length; i++) {
    var raw = (data[i][0] || "").toString().trim();
    if (raw) map.set(raw, data[i][1]);
  }
  return map;
}

/**
 * 로그 시트에서 신규 행 추출 (중복 제외, 날짜 필터 적용)
 */
function extractNewRows_(productMap, existingKeys) {
  var logSheet = SpreadsheetApp.openById(CONFIG.LOG_SHEET_ID).getSheets()[0];
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return [];

  var data = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();

  // 날짜 필터 기준
  var cutoff = null;
  if (CONFIG.SYNC_DAYS > 0) {
    cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - CONFIG.SYNC_DAYS);
    cutoff.setHours(0, 0, 0, 0);
  }

  var newRows = [];
  for (var j = 0; j < data.length; j++) {
    var row = data[j];
    var date = row[0];

    // 날짜 필터: 기준일 이전 데이터 건너뜀
    if (cutoff && date instanceof Date && date < cutoff) continue;
    // 빈 행 건너뜀
    if (!date) continue;

    var advertiser = row[1];
    var rawProduct = (row[2] || "").toString().trim();
    var spend = row[4];

    var normalizedProduct = productMap.get(rawProduct) || row[2];

    var key = makeKey_(date, advertiser, normalizedProduct, spend);
    if (!existingKeys.has(key)) {
      newRows.push([
        date,
        advertiser,
        normalizedProduct,
        row[3],  // D열: 수량/노출
        spend,
        row[5],  // F열: 비고
        row[6]   // G열: 추가 데이터
      ]);
      existingKeys.add(key);
    }
  }
  return newRows;
}

/**
 * 중복 체크용 키 생성
 * Date 객체는 getTime()으로 숫자 변환 → toString() 대비 ~10배 빠름
 */
function makeKey_(date, advertiser, product, spend) {
  var d = (date instanceof Date) ? date.getTime() : date;
  return d + "_" + advertiser + "_" + product + "_" + spend;
}

/** 메뉴 등록 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('데이터 자동화')
    .addItem('광고 성과 취합 실행', 'runAdvertisingDataSync')
    .addToUi();
}
