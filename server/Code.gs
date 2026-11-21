/**
 * WorkAttend — shared database + translator (Google Apps Script, bound to a Google Sheet)
 *
 * Install (owner, once):
 *   1. Open the Google Sheet → Extensions → Apps Script → paste this whole file → Save.
 *   2. Run the function "setup" once (▶ Run) and allow the permissions.
 *   3. Deploy → New deployment → type "Web app"
 *        Execute as: Me      Who has access: Anyone
 *      → Deploy → copy the Web app URL (ends in /exec) into the app's Settings.
 *   4. Change the PIN in the "Settings" tab (cell B2) whenever you want.
 *
 * Handover: transfer ownership of the Sheet in Google Drive. The new owner repeats steps 2–3
 * (new URL) and shares the new setup link from the app.
 *
 * API (POST, body = JSON text):
 *   {action:"check", pin}
 *   {action:"sync", pin, since, push:{employees:[], attendance:[], messages:[]}} → {now, employees, attendance, messages}
 *   {action:"translate", pin, text} → {source, ar, en}
 */

var VERSION = 1;
var DEFAULT_PIN = 'WORK';
var TABLES = {
  Employees:  ['id', 'name', 'phone', 'title', 'dept', 'notes', 'createdAt', 'updatedAt', 'deletedAt', 'srv'],
  Attendance: ['id', 'empId', 'empName', 'date', 'status', 'statusLabel', 'note', 'updatedAt', 'deletedAt', 'srv'],
  Messages:   ['id', 'date', 'text', 'createdAt', 'updatedAt', 'deletedAt', 'srv']
};
var KEY_OF = { employees: 'Employees', attendance: 'Attendance', messages: 'Messages' };
var NUMERIC = { createdAt: 1, updatedAt: 1, deletedAt: 1, srv: 1 };
var STATUS_LABEL = {
  present: 'حضور / Present', absent: 'غياب / Absent', sick: 'مرضية / Sick leave',
  leave: 'إجازة / Leave', annual: 'سنوية / Annual leave'
};

/* ---------- setup ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (name) { sheet_(ss, name); });
  var st = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (!st.getRange('A1').getValue()) {
    st.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold');
    st.getRange('A2:B2').setValues([['PIN', DEFAULT_PIN]]);
    st.getRange('A3:B3').setValues([['Version', VERSION]]);
    st.getRange('A5').setValue('Change the PIN in B2. Everyone then enters the new PIN in the app.');
    st.setColumnWidth(1, 140); st.setColumnWidth(2, 220);
  }
  var first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return 'Ready';
}

function sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  var cols = TABLES[name];
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@'); // keep dates/phones as text
  }
  return sh;
}

function pin_() {
  var st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var p = st ? String(st.getRange('B2').getValue()) : '';
  return (p || DEFAULT_PIN).trim().toUpperCase();
}

/* ---------- web endpoints ---------- */
function doGet() {
  return out_({ ok: true, app: 'workattend', version: VERSION });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'BAD_JSON' }); }
  if (String(body.pin || '').trim().toUpperCase() !== pin_()) return out_({ error: 'BAD_PIN' });
  try {
    if (body.action === 'check') return out_({ ok: true, version: VERSION });
    if (body.action === 'translate') return out_(translate_(body.text));
    if (body.action === 'sync') {
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try { return out_(sync_(body)); } finally { lock.releaseLock(); }
    }
    return out_({ error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return out_({ error: 'SERVER', message: String(err && err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sync (last write wins on updatedAt, soft deletes) ---------- */
function sync_(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = Date.now();
  var since = Math.max(0, (Number(body.since) || 0) - 10000); // overlap; re-applying is harmless
  var push = body.push || {};
  var result = { now: now };

  Object.keys(KEY_OF).forEach(function (key) {
    var name = KEY_OF[key];
    var cols = TABLES[name];
    var sh = sheet_(ss, name);
    var last = sh.getLastRow();
    var data = last > 1 ? sh.getRange(2, 1, last - 1, cols.length).getValues() : [];
    var index = {};
    for (var i = 0; i < data.length; i++) index[String(data[i][0])] = i;

    var changed = false, appended = [];
    (push[key] || []).forEach(function (r) {
      if (!r || !r.id) return;
      if (key === 'attendance') r.statusLabel = STATUS_LABEL[r.status] || '';
      var row = cols.map(function (c) {
        if (c === 'srv') return String(now);
        var v = r[c];
        return v === undefined || v === null ? '' : String(v);
      });
      var at = index[String(r.id)];
      if (at === undefined) {
        index[String(r.id)] = data.length + appended.length;
        appended.push(row);
      } else if (Number(r.updatedAt) >= Number(data[at][cols.indexOf('updatedAt')])) {
        data[at] = row; changed = true;
      }
    });

    if (changed) sh.getRange(2, 1, data.length, cols.length).setValues(data);
    if (appended.length) {
      var start = data.length + 2;
      var need = start + appended.length - 1 - sh.getMaxRows();
      if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
      sh.getRange(start, 1, appended.length, cols.length).setNumberFormat('@').setValues(appended);
      data = data.concat(appended);
    }

    var srvCol = cols.indexOf('srv');
    result[key] = data.filter(function (row) { return Number(row[srvCol]) > since; }).map(function (row) {
      var o = {};
      cols.forEach(function (c, j) {
        if (c === 'srv' || c === 'statusLabel' || c === 'empName') return;
        var v = row[j];
        o[c] = NUMERIC[c] ? (v === '' ? null : Number(v)) : (v === '' ? null : String(v));
      });
      return o;
    });
  });
  return result;
}

/* ---------- translation (Google Translate, free inside Apps Script) ---------- */
function translate_(text) {
  text = String(text || '').trim().slice(0, 4500);
  if (!text) return { source: null, ar: '', en: '' };
  var ar = (text.match(/[؀-ۿ]/g) || []).length;
  var lat = (text.match(/[A-Za-z]/g) || []).length;
  if (ar >= lat) return { source: 'ar', ar: text, en: LanguageApp.translate(text, 'ar', 'en') };
  return { source: 'en', en: text, ar: LanguageApp.translate(text, 'en', 'ar') };
}
