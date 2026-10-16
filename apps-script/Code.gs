/**
 * Lakehouse Pro Quiz - progress API backed by this Google Sheet.
 *
 * Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 *
 *   GET  <url>?username=alice          -> { ok, found, data }        one user's progress, questions and notes
 *   GET  <url>?action=all              -> { ok, users: [...] }       summary of every user, for the group dashboard
 *   POST <url>  body: {"username":"alice","data":{...}}               save progress
 *   POST <url>  body: {"action":"upload","username":"alice","name":"x.jpg","mime":"image/jpeg","data":"<base64>"}
 *                                      -> { ok, id }                 store a notes picture in Drive
 *        (send as Content-Type text/plain so the browser skips the CORS preflight)
 *
 * Progress tab:  one row per user. Username is the row key (lowercase letters, digits, . _ -).
 * Questions tab: one row per question a user added to their question bank.
 * Notes tab:     one row per notes heading / subheading.
 *   A POST that includes data.questions (or data.notes) replaces that user's rows in that tab;
 *   one without it leaves them alone.
 * Pictures go to a Drive folder named "Lakehouse Quiz Images", shared as view-only by link.
 */

var SHEET_NAME = 'Progress';
var HEADERS = ['Username', 'TotalAnswered', 'TotalCorrect', 'SectionScores',
               'MockBestPro', 'MockBestAssoc', 'LastUpdated', 'State'];
var USER_RE = /^[a-z0-9._-]{1,40}$/;
var MAX_CELL = 49000; // Sheets allows 50,000 characters per cell

var Q_SHEET_NAME = 'Questions';
var Q_HEADERS = ['Username', 'QuestionId', 'Exam', 'SectionId', 'Section', 'Question', 'Code',
                 'CorrectAnswer', 'Option2', 'Option3', 'Option4', 'Explanation', 'SeenInExam', 'Updated'];

var N_SHEET_NAME = 'Notes';
var N_HEADERS = ['Username', 'HeadingId', 'Heading', 'SubheadingId', 'Subheading', 'Order', 'Body', 'Images', 'Updated'];

var IMAGE_FOLDER = 'Lakehouse Quiz Images';
var IMAGE_TYPES = { 'image/jpeg': 1, 'image/png': 1, 'image/webp': 1, 'image/gif': 1 };
var MAX_IMAGE_B64 = 4000000; // about 3 MB of image data

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action === 'all') return json_({ ok: true, users: allUsers_() });
    var username = cleanUser_(p.username);
    var sheet = getSheet_();
    var row = findRow_(sheet, username);
    if (!row) return json_({ ok: true, found: false });
    var v = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    var data = rowToData_(v);
    data.questions = readQuestions_(username);
    data.notes = readNotes_(username);
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
    if (body.action === 'upload') return json_(uploadImage_(username, body));

    var d = body.data || {};
    var updated = Number(d.lastUpdated) || Date.now();
    var sections = JSON.stringify(d.sectionScores || {});
    var state = d.state == null ? '' : JSON.stringify(d.state);
    if (sections.length > MAX_CELL) throw new Error('sectionScores too large');
    if (state.length > MAX_CELL) throw new Error('state too large for one cell (' + state.length + ' chars)');
    var questions = Array.isArray(d.questions) ? d.questions.map(function (q) { return questionRow_(username, q, updated); }) : null;
    var notes = Array.isArray(d.notes) ? d.notes.map(function (n) { return noteRow_(username, n, updated); }) : null;

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
    if (questions) replaceUserRows_(getTab_(Q_SHEET_NAME, Q_HEADERS), Q_HEADERS, username, questions);
    if (notes) replaceUserRows_(getTab_(N_SHEET_NAME, N_HEADERS), N_HEADERS, username, notes);
    return json_({ ok: true, saved: true, lastUpdated: updated });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- group dashboard ---------- */

function allUsers_() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var qCount = countByUser_(Q_SHEET_NAME, 1), nCount = countByUser_(N_SHEET_NAME, 3);
  return sheet.getRange(2, 1, last - 1, HEADERS.length - 1).getValues() // everything except State
    .filter(function (v) { return v[0] !== ''; })
    .map(function (v) {
      var d = rowToData_(v.concat(['']));
      delete d.state;
      d.questionsAdded = qCount[d.username] || 0;
      d.notesSections = nCount[d.username] || 0;
      return d;
    });
}

/* Number of non-marker rows per user in a tab (a row counts when column `idCol`, 0-based, is filled). */
function countByUser_(name, idCol) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name), out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, idCol + 1).getValues().forEach(function (r) {
    if (r[0] !== '' && r[idCol] !== '') out[r[0]] = (out[r[0]] || 0) + 1;
  });
  return out;
}

/* ---------- pictures ---------- */

function uploadImage_(username, body) {
  var mime = String(body.mime || '');
  var b64 = String(body.data || '');
  if (!IMAGE_TYPES[mime]) throw new Error('Only JPEG, PNG, WebP or GIF pictures can be added');
  if (!b64 || b64.length > MAX_IMAGE_B64) throw new Error('Picture is too large');
  var name = username + '-' + Date.now() + '-' + String(body.name || 'image').replace(/[^\w.-]/g, '_').slice(0, 60);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  var folders = DriveApp.getFoldersByName(IMAGE_FOLDER);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(IMAGE_FOLDER);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId() };
}

/* ---------- per-user tabs (Questions, Notes) ---------- */

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

/* The last column of every per-user tab is the Updated date; everything before it is stored as text. */
function getTab_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* Rows of one user, or null when the user has never saved any (older clients kept them in State). */
function userRows_(name, headers, username) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
    .filter(function (r) { return String(r[0]) === username; });
  return rows.length ? rows : null;
}

/* Replace all of one user's rows with the given ones, keeping everyone else's. */
function replaceUserRows_(sheet, headers, username, newRows) {
  var n = headers.length, last = sheet.getLastRow();
  var keep = last > 1 ? sheet.getRange(2, 1, last - 1, n).getValues()
    .filter(function (r) { return r[0] !== '' && String(r[0]) !== username; })
    .map(function (r) { return r.map(function (v, i) { return i < n - 1 && typeof v === 'string' ? text_(v) : v; }); }) : [];
  // An empty list still gets one marker row so a later load knows it is empty rather than never saved.
  if (!newRows.length) {
    var marker = [username];
    for (var i = 1; i < n - 1; i++) marker.push('');
    newRows = [marker.concat([new Date()])];
  }
  var all = keep.concat(newRows);
  if (last > 1) sheet.getRange(2, 1, last - 1, n).clearContent();
  sheet.getRange(2, 1, all.length, n - 1).setNumberFormat('@');
  sheet.getRange(2, 1, all.length, n).setValues(all);
}

function readQuestions_(username) {
  var rows = userRows_(Q_SHEET_NAME, Q_HEADERS, username);
  if (!rows) return null;
  return rows.filter(function (r) { return r[1] !== ''; }).map(function (r) {
    return {
      id: String(r[1]), exam: String(r[2]), sectionId: String(r[3]), sectionName: String(r[4]),
      question: String(r[5]), code: String(r[6]), correct: String(r[7]),
      option2: String(r[8]), option3: String(r[9]), option4: String(r[10]),
      explanation: String(r[11]), seenInExam: r[12] === true || r[12] === 'TRUE'
    };
  });
}

function readNotes_(username) {
  var rows = userRows_(N_SHEET_NAME, N_HEADERS, username);
  if (!rows) return null;
  return rows.filter(function (r) { return r[1] !== ''; }).map(function (r) {
    return {
      headingId: String(r[1]), heading: String(r[2]), subId: String(r[3]), subheading: String(r[4]),
      order: Number(r[5]) || 0, body: String(r[6]), images: parse_(r[7], [])
    };
  });
}

function questionRow_(username, q, updated) {
  q = q || {};
  var cells = textCells_([username, q.id, q.exam, q.sectionId, q.sectionName, q.question, q.code,
                          q.correct, q.option2, q.option3, q.option4, q.explanation]);
  if (!cells[1]) throw new Error('Question without an id');
  return cells.concat([q.seenInExam === true ? 'TRUE' : 'FALSE', new Date(updated)]);
}

function noteRow_(username, n, updated) {
  n = n || {};
  var cells = textCells_([username, n.headingId, n.heading, n.subId, n.subheading, String(Number(n.order) || 0),
                          n.body, JSON.stringify(Array.isArray(n.images) ? n.images : [])]);
  if (!cells[1]) throw new Error('Note without a heading id');
  return cells.concat([new Date(updated)]);
}

function textCells_(values) {
  return values.map(function (v) {
    v = String(v == null ? '' : v);
    if (v.length > MAX_CELL) throw new Error('A field is too long for one cell (' + v.length + ' characters)');
    return text_(v);
  });
}

/* A leading = + - @ would make Sheets treat user text as a formula. */
function text_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/* ---------- helpers ---------- */

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

/* v tells the quiz what this script supports: 2 = Questions tab, 3 = Notes tab, pictures, group dashboard. */
function json_(obj) {
  obj.v = 3;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
