/**
 * KR Lost Customer Tracker — ตัวรับคำตอบจากเซล (Google Apps Script)
 *
 * ติดตั้ง
 *  1. เปิดโปรเจกต์ Apps Script (หรือ Google Sheet → Extensions → Apps Script)
 *  2. ลบโค้ดเดิมทั้งหมด วางไฟล์นี้ทั้งไฟล์ → Save
 *  3. เลือกฟังก์ชัน setup → Run → อนุญาตสิทธิ์ (Execution log จะแสดงลิงก์ Google Sheet ที่เก็บคำตอบ)
 *  4. Deploy → New deployment → Select type: Web app
 *       Execute as: Me   ·   Who has access: Anyone
 *  5. กด Deploy → คัดลอก Web app URL (ลงท้ายด้วย /exec) เปิดในเบราว์เซอร์ต้องเห็น {"ok":true,...}
 *  6. รัน: python kr_tracker.py build "KR May to Oct.xls" --sheet-url <URL>
 *
 * ชีต "Feedback" = คำตอบล่าสุดของลูกค้าแต่ละราย · ชีต "Log" = ประวัติทุกครั้งที่กดส่ง
 */
const LATEST = 'Feedback';
const LOG = 'Log';
const HEAD = ['id', 'เดือน', 'มุมมอง', 'รหัส', 'ลูกค้า', 'Sales', 'ประเภท', 'result',
              'ผลการเช็ค', 'หมายเหตุ', 'ผู้เช็ค', 'เวลาบันทึก', 'at'];
const RESULT_TH = {
  back: 'จะกลับมาจอง / มีงานเดือนหน้า', shift: 'เลื่อนชิปเมนต์ / ไปอยู่เรือเดือนถัดไป',
  price: 'ราคาสู้คู่แข่งไม่ได้', competitor: 'ย้ายไปใช้สายเรืออื่น',
  space: 'ปัญหาพื้นที่ / ตารางเรือ / อุปกรณ์', noorder: 'ออเดอร์ลด / ไม่มีสินค้า',
  spot: 'เป็นงาน spot ไม่ประจำ', closed: 'เลิกส่งเส้นทางนี้ / ปิดกิจการ', other: 'อื่นๆ (ดูหมายเหตุ)',
};
const TYPE_TH = {lost: 'หายไปเดือนนี้', drop: 'ลดลง ≥50%', gone: 'หายไปก่อนหน้า', new: 'ลูกค้าใหม่'};
const ID_RE = /^(cust|org)~\d{4}-\d{2}~[a-z0-9]+$/;

/** ใช้ชีตที่สคริปต์ผูกอยู่ ถ้าสร้างสคริปต์แยก (script.google.com) จะสร้างชีต "KR Lost Customer Feedback" ใน Drive ให้ */
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

/** กด Run ฟังก์ชันนี้ใน Apps Script เพื่อดูลิงก์ Google Sheet ที่เก็บคำตอบ (ดูใน Execution log) */
function setup() {
  const ss = book_();
  sheet_(LATEST); sheet_(LOG);
  Logger.log('Google Sheet ที่เก็บคำตอบ: ' + ss.getUrl());
}

function sheet_(name) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function str_(v, max) { return String(v == null ? '' : v).slice(0, max || 200); }

function toRow_(id, r) {
  const at = Number(r.at) || Date.now();
  return [id, str_(r.month, 7), r.view === 'cust' ? 'Shipper' : 'Org Shipper', "'" + str_(r.key, 60),
          str_(r.name), str_(r.sales, 80), TYPE_TH[r.type] || str_(r.type, 20), str_(r.result, 20),
          RESULT_TH[r.result] || '', str_(r.note, 2000), str_(r.by, 80), new Date(at), at];
}

/** หน้าเว็บดึงคำตอบล่าสุดทั้งหมด */
function doGet() {
  const sh = sheet_(LATEST);
  const values = sh.getDataRange().getValues().slice(1);
  const items = {};
  values.forEach(v => {
    if (!ID_RE.test(v[0])) return;
    items[v[0]] = {month: v[1], view: v[2] === 'Shipper' ? 'cust' : 'org', key: String(v[3]), name: v[4],
                   sales: v[5], type: Object.keys(TYPE_TH).find(k => TYPE_TH[k] === v[6]) || v[6],
                   result: v[7], note: v[9], by: v[10], at: Number(v[12]) || 0};
  });
  return json_({ok: true, items: items});
}

/** หน้าเว็บส่งคำตอบ: {items: {id: {result, note, by, at, view, month, key, name, sales, type}}} */
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
