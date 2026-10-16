/**
 * Lakehouse Pro Quiz - progress API backed by this Google Sheet.
 *
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 *
 *   GET  <url>?username=alice          -> { ok, found, data }
 *   POST <url>  body: {"username":"alice","data":{...}}
 *        (send as Content-Type text/plain so the browser skips the CORS preflight)
 *
 * Progress tab: one row per user. Username is the row key (lowercase letters, digits, . _ -).
 * Questions tab: one row per question a user added to their question bank.
 *   A POST that includes data.questions replaces that user's rows; one without it leaves them alone.
 */

var SHEET_NAME = 'Progress';
var HEADERS = ['Username', 'TotalAnswered', 'TotalCorrect', 'SectionScores',
               'MockBestPro', 'MockBestAssoc', 'LastUpdated', 'State'];
var USER_RE = /^[a-z0-9._-]{1,40}$/;
var MAX_CELL = 49000; // Sheets allows 50,000 characters per cell

var Q_SHEET_NAME = 'Questions';
var Q_HEADERS = ['Username', 'QuestionId', 'Exam', 'SectionId', 'Section', 'Question', 'Code',
                 'CorrectAnswer', 'Option2', 'Option3', 'Option4', 'Explanation', 'SeenInExam', 'Updated'];

function doGet(e) {
  try {
    var username = cleanUser_(e && e.parameter && e.parameter.username);
    var sheet = getSheet_();
    var row = findRow_(sheet, username);
    if (!row) return json_({ ok: true, found: false });
    var v = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    var data = rowToData_(v);
    data.questions = readQuestions_(username);
    return json_({ ok: true, found: true, data: data });
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
    var questions = Array.isArray(d.questions) ? d.questions.map(function (q) { return questionRow_(username, q, updated); }) : null;

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
    if (questions) writeQuestions_(username, questions);
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

function getQuestionSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(Q_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(Q_SHEET_NAME);
    sheet.appendRow(Q_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, sheet.getMaxRows(), Q_HEADERS.length - 2).setNumberFormat('@'); // text, not numbers or formulas
  }
  return sheet;
}

/* Returns the user's questions, or null when the user has never saved any (older clients kept them in State). */
function readQuestions_(username) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Q_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Q_HEADERS.length).getValues();
  var out = [], any = false;
  rows.forEach(function (r) {
    if (String(r[0]) !== username) return;
    any = true;
    if (!r[1]) return; // marker row: the user has saved an empty question bank
    out.push({
      id: String(r[1]), exam: String(r[2]), sectionId: String(r[3]), sectionName: String(r[4]),
      question: String(r[5]), code: String(r[6]), correct: String(r[7]),
      option2: String(r[8]), option3: String(r[9]), option4: String(r[10]),
      explanation: String(r[11]), seenInExam: r[12] === true || r[12] === 'TRUE'
    });
  });
  return any ? out : null;
}

/* Replace all of one user's rows with the given ones, keeping everyone else's. */
function writeQuestions_(username, newRows) {
  var sheet = getQuestionSheet_();
  var last = sheet.getLastRow();
  var keep = last > 1 ? sheet.getRange(2, 1, last - 1, Q_HEADERS.length).getValues()
    .filter(function (r) { return String(r[0]) !== username; })
    .map(function (r) { return r.map(function (v, i) { return i < Q_HEADERS.length - 2 ? text_(String(v)) : v; }); }) : [];
  // An empty bank still gets one marker row so a later load knows the bank is empty rather than never saved.
  if (!newRows.length) newRows = [[username, '', '', '', '', '', '', '', '', '', '', '', false, new Date()]];
  var all = keep.concat(newRows);
  if (last > 1) sheet.getRange(2, 1, last - 1, Q_HEADERS.length).clearContent();
  sheet.getRange(2, 1, all.length, Q_HEADERS.length - 2).setNumberFormat('@');
  sheet.getRange(2, 1, all.length, Q_HEADERS.length).setValues(all);
}

function questionRow_(username, q, updated) {
  q = q || {};
  var cells = [username, q.id, q.exam, q.sectionId, q.sectionName, q.question, q.code,
               q.correct, q.option2, q.option3, q.option4, q.explanation].map(function (v) {
    v = String(v == null ? '' : v);
    if (v.length > MAX_CELL) throw new Error('A question field is too long for one cell');
    return text_(v);
  });
  if (!cells[1]) throw new Error('Question without an id');
  return cells.concat([q.seenInExam === true, new Date(updated)]);
}

/* A leading = + - @ would make Sheets treat user text as a formula. */
function text_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
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
