/**
 * KR Lost Customer Tracker — receives sales feedback (Google Apps Script)
 *
 * Setup
 *  1. Open the Apps Script project (or Google Sheet → Extensions → Apps Script)
 *  2. Delete all existing code, paste this whole file → Save
 *  3. Choose the function "setup" → Run → allow access (the Execution log shows the Google Sheet link)
 *  4. Deploy → New deployment → Select type: Web app
 *       Execute as: Me   ·   Who has access: Anyone
 *  5. Click Deploy → copy the Web app URL (ends with /exec); opening it in a browser must show {"ok":true,...}
 *  6. Run: python kr_tracker.py build "KR May to Oct.xls" --sheet-url <URL>
 *
 * Sheet "Feedback" = latest answer per customer · Sheet "Log" = every submission
 */
const LATEST = 'Feedback';
const LOG = 'Log';
const HEAD = ['id', 'Month', 'View', 'Code', 'Customer', 'Sales', 'Type', 'result',
              'Check result', 'Note', 'Checked by', 'Saved at', 'at'];
const RESULT_EN = {
  back: 'Will book again / has cargo next month', shift: 'Shipment postponed / moved to next month vessel',
  price: 'Price not competitive', competitor: 'Moved to another carrier',
  space: 'Space / schedule / equipment issue', noorder: 'Fewer orders / no cargo',
  spot: 'Spot cargo, not regular', closed: 'Stopped this route / closed business', other: 'Other (see note)',
};
const TYPE_EN = {lost: 'Lost this month', drop: 'Dropped ≥50%', gone: 'Lost earlier', new: 'New customer'};
// labels written by the earlier Thai version, so old rows still read back correctly
const TYPE_TH = {lost: 'หายไปเดือนนี้', drop: 'ลดลง ≥50%', gone: 'หายไปก่อนหน้า', new: 'ลูกค้าใหม่'};
const ID_RE = /^(cust|org)~\d{4}-\d{2}~[a-z0-9]+$/;

/** Uses the bound spreadsheet; a standalone script creates "KR Lost Customer Feedback" in Drive */
function book_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('KR Lost Customer Feedback');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

/** Run this once in Apps Script; the Execution log shows the Google Sheet link */
function setup() {
  const ss = book_();
  sheet_(LATEST); sheet_(LOG);
  Logger.log('Feedback Google Sheet: ' + ss.getUrl());
}

function sheet_(name) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setFrozenRows(1);
  }
  const head = sh.getRange(1, 1, 1, HEAD.length);
  if (head.getValues()[0].join('|') !== HEAD.join('|')) head.setValues([HEAD]).setFontWeight('bold');
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function str_(v, max) { return String(v == null ? '' : v).slice(0, max || 200); }

function typeKey_(label) {
  return Object.keys(TYPE_EN).find(k => TYPE_EN[k] === label || TYPE_TH[k] === label) || label;
}

function toRow_(id, r) {
  const at = Number(r.at) || Date.now();
  return [id, "'" + str_(r.month, 7), r.view === 'cust' ? 'Shipper' : 'Org Shipper', "'" + str_(r.key, 60),
          str_(r.name), str_(r.sales, 80), TYPE_EN[r.type] || str_(r.type, 20), str_(r.result, 20),
          RESULT_EN[r.result] || '', str_(r.note, 2000), str_(r.by, 80), new Date(at), at];
}

/** The page loads the latest answers */
function doGet() {
  const sh = sheet_(LATEST);
  const values = sh.getDataRange().getValues().slice(1);
  const items = {};
  values.forEach(v => {
    if (!ID_RE.test(v[0])) return;
    items[v[0]] = {month: v[1] instanceof Date ? Utilities.formatDate(v[1], 'Asia/Bangkok', 'yyyy-MM') : String(v[1]),
                   view: v[2] === 'Shipper' ? 'cust' : 'org', key: String(v[3]), name: v[4],
                   sales: v[5], type: typeKey_(v[6]), result: v[7], note: v[9], by: v[10], at: Number(v[12]) || 0};
  });
  return json_({ok: true, items: items});
}

/** The page submits answers: {items: {id: {result, note, by, at, view, month, key, name, sales, type}}} */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: 'bad json'}); }
  const items = body && body.items;
  if (!items || typeof items !== 'object') return json_({ok: false, error: 'no items'});

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const latest = sheet_(LATEST), log = sheet_(LOG);
    const ids = latest.getRange(1, 1, latest.getLastRow(), 1).getValues().map(v => v[0]);
    const ats = latest.getRange(1, HEAD.length, latest.getLastRow(), 1).getValues().map(v => Number(v[0]) || 0);
    let saved = 0;
    Object.keys(items).slice(0, 500).forEach(id => {
      const r = items[id];
      if (!ID_RE.test(id) || !r || typeof r !== 'object') return;
      const row = toRow_(id, r);
      log.appendRow(row);
      const i = ids.indexOf(id);
      if (i < 0) { latest.appendRow(row); ids.push(id); ats.push(row[12]); }
      else if (row[12] >= ats[i]) { latest.getRange(i + 1, 1, 1, HEAD.length).setValues([row]); ats[i] = row[12]; }
      saved++;
    });
    return json_({ok: true, saved: saved});
  } finally {
    lock.releaseLock();
  }
}
