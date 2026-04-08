/**
 * 광고 데이터 매핑 및 자동 배분 스크립트 (간소화 버전)
 * - 원본 로그 1회만 로드 (캐싱)
 * - getLastRow()로 필요한 범위만 읽기
 * - 중복 체크 시 키 열만 로드
 */

var LOG_SOURCE_ID = "12beX3xFWni4bRILjQ8C_zDLGZ-fr3pCB-wkUmODF2BU";
var OBJ_PRICE_ID  = "1YrYqZ-spGgK0ywb_moPlT-Ah_plSSUhlNysx65Yl_g0";
var DB_RAW_ID     = "15WPIAXnwzXNO6k6Z69QYSW0-MPuzszgulPKqXXSe3rA";

function onOpen() {
  SpreadsheetApp.getUi().createMenu('데이터 자동화')
    .addItem('광고 차수 선택 및 배분 실행', 'showRoundSelector')
    .addToUi();
}

/** 원본 로그 1회 로드 후 캐싱 */
var _logCache = null;
function getLogData_() {
  if (_logCache) return _logCache;
  var sheet = SpreadsheetApp.openById(LOG_SOURCE_ID).getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  _logCache = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return _logCache;
}

/** 2026년 차수 선택 팝업 */
function showRoundSelector() {
  var ui = SpreadsheetApp.getUi();
  try {
    var data = getLogData_();
    var rounds = [];
    var seen = {};
    for (var i = 0; i < data.length; i++) {
      var r = data[i][2];
      if (r && r.toString().indexOf("2026") !== -1 && !seen[r]) {
        seen[r] = true;
        rounds.push(r);
      }
    }
    rounds.sort().reverse();

    if (rounds.length === 0) {
      ui.alert("2026년 광고 차수 데이터를 찾을 수 없습니다.");
      return;
    }

    var options = rounds.map(function(r) { return '<option value="' + r + '">' + r + '</option>'; }).join('');
    var html = '<html><head><style>'
      + 'body{font-family:sans-serif;padding:15px}'
      + 'select{width:100%;padding:10px;margin-bottom:20px;font-size:14px}'
      + 'button{width:100%;padding:12px;background:#4285f4;color:white;border:none;cursor:pointer;font-weight:bold;border-radius:4px}'
      + '</style></head><body>'
      + '<div style="margin-bottom:10px;font-size:14px;font-weight:bold">가져올 2026년 차수 선택:</div>'
      + '<select id="rs">' + options + '</select>'
      + '<button onclick="google.script.run.withSuccessHandler(function(){google.script.host.close()}).processDistribution(document.getElementById(\'rs\').value)">데이터 배분 시작</button>'
      + '</body></html>';

    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(230), '광고 차수 선택');
  } catch (e) {
    ui.alert("원본 시트 접근 오류: " + e.message);
  }
}

/** 데이터 배분 실행 */
function processDistribution(selectedRound) {
  var ui = SpreadsheetApp.getUi();
  try {
    var priceSheet = SpreadsheetApp.openById(OBJ_PRICE_ID).getSheetByName("객단가정보");
    var rawSheet = SpreadsheetApp.openById(DB_RAW_ID).getSheetByName("RAW");
    if (!priceSheet || !rawSheet) {
      ui.alert("대상 시트의 탭 이름을 확인해주세요. (객단가정보 / RAW)");
      return;
    }

    // 캐싱 초기화 (팝업에서 호출되므로 새 실행 컨텍스트)
    _logCache = null;
    var logData = getLogData_();

    // 선택된 차수 필터링
    var targetData = [];
    for (var i = 0; i < logData.length; i++) {
      if (logData[i][2] === selectedRound) targetData.push(logData[i]);
    }
    if (targetData.length === 0) {
      ui.alert("데이터가 없습니다.");
      return;
    }

    // 객단가 시트용 매핑
    var priceMapped = targetData.map(function(row) {
      return [row[4], row[2], row[3], row[4], row[5], row[6], "", row[8]];
    });

    // DB RAW 시트용 매핑
    var rawMapped = targetData.map(function(row) {
      return [row[6], row[2], row[3], row[4], row[8], ""];
    });

    // 중복 제외 후 추가 (키: A열 + B열)
    var keyFn = function(r) { return r[0] + "_" + r[1]; };
    var c1 = appendUnique_(priceSheet, priceMapped, keyFn);
    var c2 = appendUnique_(rawSheet, rawMapped, keyFn);

    ui.alert("배분 완료\n\n- 선택 차수: " + selectedRound + "\n- 객단가정보 추가: " + c1 + "건\n- DB RAW 추가: " + c2 + "건");
  } catch (e) {
    ui.alert("오류: " + e.message);
  }
}

/** 중복 제외 후 하단 추가 — 키 열만 로드하여 메모리 절감 */
function appendUnique_(sheet, newData, keyFn) {
  var lastRow = sheet.getLastRow();
  var existingKeys = new Set();

  if (lastRow >= 2) {
    // A열, B열(키에 쓰이는 열)만 로드
    var colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      existingKeys.add(colA[i][0] + "_" + colB[i][0]);
    }
  }

  var toAdd = [];
  for (var j = 0; j < newData.length; j++) {
    var key = keyFn(newData[j]);
    if (!existingKeys.has(key)) {
      toAdd.push(newData[j]);
      existingKeys.add(key);
    }
  }

  if (toAdd.length > 0) {
    sheet.getRange(lastRow + 1, 1, toAdd.length, toAdd[0].length).setValues(toAdd);
  }
  return toAdd.length;
}
