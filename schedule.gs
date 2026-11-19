/**
 * 폭스러닝센터 — 오늘의 활동 (학생용 화면) 데이터 API
 *
 * '활동 내용' 탭에서 오늘 날짜가 들어 있는 주간 블록을 찾아
 * [연구소 이름, 오늘 칸] 만 그대로 잘라서 JSON으로 돌려줍니다.
 * 글자를 해석하는 일(시간·이름 나누기)은 index.html 이 합니다.
 *
 * 배포: 확장 프로그램 → Apps Script → 이 파일 붙여넣기 →
 *       배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 */

var SCHEDULE_SPREADSHEET_ID = '1Q1F7Hda6fglcwuvLKCP9S3Xf2PiwQHyp9p1iCekI2Mg';
var SCHEDULE_SHEET_NAME = '활동 내용';
var SCHEDULE_CACHE_SECONDS = 30;

// 튼튼: 학생이 화면에서 이름을 끌어 고른 시간을 저장하는 탭 (없으면 처음 저장할 때 만들어요)
var TUNTEUN_SHEET_NAME = '튼튼_배정';
var TUNTEUN_SLOTS = { '16:00~16:30': true, '16:30~17:00': true, '': true };

function doGet(e) {
  var p = (e && e.parameter) || {};
  var ss = SpreadsheetApp.openById(SCHEDULE_SPREADSHEET_ID);
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var date = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '')
    ? p.date
    : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  if (p.action === 'tunteun') {
    return json_(moveTunteun_(ss, tz, date, p.name, p.slot));
  }

  var cache = CacheService.getScriptCache();
  var key = 'schedule:' + date;
  var body = p.fresh ? null : cache.get(key);
  if (!body) {
    var result;
    try {
      var sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
      if (!sheet) throw new Error("'" + SCHEDULE_SHEET_NAME + "' 탭을 찾을 수 없어요");
      result = sliceDay_(sheet.getDataRange().getDisplayValues(), date);
      result.status = 'ok';
    } catch (err) {
      result = { status: 'error', message: String(err && err.message || err) };
    }
    result.date = date;
    result.updated = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
    body = JSON.stringify(result);
    if (result.status === 'ok' && body.length < 90000) {
      cache.put(key, body, SCHEDULE_CACHE_SECONDS);
    }
  }
  var obj = JSON.parse(body);
  if (obj.status === 'ok') {
    try { obj.tunteun = readTunteun_(ss, date); }
    catch (err) { obj.tunteunError = String(err && err.message || err); }
  }
  return json_(obj);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 튼튼 시간 고르기 ───────── */

function tunteunSheet_(ss, create) {
  var sh = ss.getSheetByName(TUNTEUN_SHEET_NAME);
  if (!sh && create) {
    sh = ss.insertSheet(TUNTEUN_SHEET_NAME);
    sh.getRange('A:D').setNumberFormat('@');   // 날짜·시간이 자동 변환되지 않게 글자로
    sh.getRange(1, 1, 1, 4).setValues([['날짜', '이름', '튼튼 시간', '바꾼 시각']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 그날의 { 이름: '16:00~16:30' | '16:30~17:00' | '' } ('' = 다시 빼 둠) */
function readTunteun_(ss, date) {
  var out = {};
  var sh = tunteunSheet_(ss, false);
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] === date && v[i][1]) out[v[i][1]] = v[i][2];
  }
  return out;
}

function moveTunteun_(ss, tz, date, name, slot) {
  name = String(name || '').trim();
  slot = String(slot == null ? '' : slot).trim();
  if (!/^[가-힣A-Za-z]{1,10}$/.test(name)) return { status: 'error', message: '이름이 올바르지 않아요' };
  if (!Object.prototype.hasOwnProperty.call(TUNTEUN_SLOTS, slot)) return { status: 'error', message: '시간 칸이 올바르지 않아요' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = tunteunSheet_(ss, true);
    var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    var last = sh.getLastRow(), row = 0;
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
      for (var i = 0; i < v.length; i++) {
        if (v[i][0] === date && v[i][1] === name) { row = i + 2; break; }
      }
    }
    if (row) {
      sh.getRange(row, 3, 1, 2).setValues([[slot, now]]);
    } else {
      sh.getRange(last + 1, 1, 1, 4).setNumberFormat('@').setValues([[date, name, slot, now]]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { status: 'ok', date: date, tunteun: readTunteun_(ss, date) };
}

/**
 * values: 탭 전체 표시값(2차원 배열), date: 'yyyy-MM-dd'
 * 반환: { found, header, week, rows: [[A열, 오늘 칸], ...] }
 */
function sliceDay_(values, date) {
  var month = Number(date.slice(5, 7));
  var day = Number(date.slice(8, 10));
  var dateRe = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/;

  function isHeaderRow(row) {
    for (var c = 1; c < row.length; c++) {
      if (dateRe.test(String(row[c]))) return true;
    }
    return false;
  }

  var headerRow = -1, col = -1;
  for (var r = 0; r < values.length; r++) {
    for (var c = 1; c < values[r].length; c++) {
      var m = dateRe.exec(String(values[r][c]));
      if (m && Number(m[1]) === month && Number(m[2]) === day) {
        headerRow = r; col = c;   // 같은 날짜가 여러 번 나오면 가장 아래(최근) 블록
      }
    }
  }
  if (headerRow < 0) return { found: false, rows: [] };

  var rows = [];
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    if (isHeaderRow(row)) break;
    if (String(row[1] || '').indexOf('연구소별 활동내용') >= 0) break;
    rows.push([String(row[0] || ''), String(row[col] || '')]);
  }
  while (rows.length && !rows[rows.length - 1][0].trim() && !rows[rows.length - 1][1].trim()) {
    rows.pop();
  }
  return {
    found: true,
    header: String(values[headerRow][col]),
    week: headerRow > 0 ? String(values[headerRow - 1][0] || '') : '',
    rows: rows
  };
}
