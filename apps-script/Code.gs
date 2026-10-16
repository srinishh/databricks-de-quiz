/**
 * Lakehouse Pro Quiz - progress API backed by this Google Sheet.
 *
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 *
 *   GET  <url>?username=alice          -> { ok, found, data }
 *   POST <url>  body: {"username":"alice","data":{...}}
 *        (send as Content-Type text/plain so the browser skips the CORS preflight)
 *
 * One row per user. Username is the row key (lowercase letters, digits, . _ -).
 */

var SHEET_NAME = 'Progress';
var HEADERS = ['Username', 'TotalAnswered', 'TotalCorrect', 'SectionScores',
               'MockBestPro', 'MockBestAssoc', 'LastUpdated', 'State'];
var USER_RE = /^[a-z0-9._-]{1,40}$/;
var MAX_CELL = 49000; // Sheets allows 50,000 characters per cell

function doGet(e) {
  try {
    var username = cleanUser_(e && e.parameter && e.parameter.username);
    var sheet = getSheet_();
    var row = findRow_(sheet, username);
    if (!row) return json_({ ok: true, found: false });
    var v = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    return json_({ ok: true, found: true, data: rowToData_(v) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var username = cleanUser_(body.username);
    var d = body.data || {};
    var updated = Number(d.lastUpdated) || Date.now();
    var sections = JSON.stringify(d.sectionScores || {});
    var state = d.state == null ? '' : JSON.stringify(d.state);
    if (sections.length > MAX_CELL) throw new Error('sectionScores too large');
    if (state.length > MAX_CELL) throw new Error('state too large for one cell (' + state.length + ' chars)');

    lock.waitLock(10000);
    var sheet = getSheet_();
    var row = findRow_(sheet, username);
    if (row) {
      // Never let an older copy (e.g. a device that was offline) overwrite newer progress.
      var current = sheet.getRange(row, 7).getValue();
      var currentMs = current instanceof Date ? current.getTime() : Number(current) || 0;
      if (currentMs > updated) {
        return json_({ ok: true, stale: true, lastUpdated: currentMs });
      }
    } else {
      row = sheet.getLastRow() + 1;
    }
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      username,
      Number(d.totalAnswered) || 0,
      Number(d.totalCorrect) || 0,
      sections,
      numOrBlank_(d.mockBestPro),
      numOrBlank_(d.mockBestAssoc),
      new Date(updated),
      state
    ]]);
    return json_({ ok: true, saved: true, lastUpdated: updated });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@'); // keep usernames as plain text
  }
  return sheet;
}

function findRow_(sheet, username) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var cell = sheet.getRange(2, 1, last - 1, 1)
    .createTextFinder(username).matchEntireCell(true).matchCase(true).findNext();
  return cell ? cell.getRow() : 0;
}

function cleanUser_(u) {
  u = String(u || '').trim().toLowerCase();
  if (!USER_RE.test(u)) throw new Error('Invalid username');
  return u;
}

function rowToData_(v) {
  var ts = v[6] instanceof Date ? v[6].getTime() : Number(v[6]) || 0;
  return {
    username: String(v[0]),
    totalAnswered: Number(v[1]) || 0,
    totalCorrect: Number(v[2]) || 0,
    sectionScores: parse_(v[3], {}),
    mockBestPro: v[4] === '' ? null : Number(v[4]),
    mockBestAssoc: v[5] === '' ? null : Number(v[5]),
    lastUpdated: ts,
    state: parse_(v[7], null)
  };
}

function parse_(s, fallback) {
  if (s === '' || s == null) return fallback;
  try { return JSON.parse(s); } catch (e) { return fallback; }
}

function numOrBlank_(n) {
  return typeof n === 'number' && n >= 0 ? n : '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
