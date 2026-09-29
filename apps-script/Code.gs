/**
 * ดันดี (DunDee) — Backend API (Google Apps Script)
 * หน้าเว็บอยู่บน GitHub Pages ส่วนไฟล์นี้ทำหน้าที่เป็น API อ่าน/เขียน Google Sheet
 *
 * ติดตั้ง: เปิด Google Sheet > ส่วนขยาย > Apps Script > วางไฟล์นี้เป็น Code.gs
 *          รันฟังก์ชัน setup() หนึ่งครั้ง แล้วดู "รหัสลับ" ใน Execution log
 *          Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone)
 */

const SHEET_NAME = 'บันทึกสุขภาพ';

const KEYS = ['id', 'datetime', 'sys', 'dia', 'pulse', 'arm', 'posture',
  'weight', 'sugar', 'sugarTiming', 'temp', 'spo2', 'sleep', 'med',
  'symptoms', 'note', 'createdAt'];

const HEADERS = ['id', 'วันเวลาที่วัด', 'ตัวบน SYS (mmHg)', 'ตัวล่าง DIA (mmHg)',
  'ชีพจร (ครั้ง/นาที)', 'แขน', 'ท่า', 'น้ำหนัก (กก.)', 'น้ำตาล (mg/dL)',
  'ช่วงวัดน้ำตาล', 'อุณหภูมิ (°C)', 'SpO2 (%)', 'นอน (ชม.)', 'กินยาแล้ว',
  'อาการ', 'โน้ต', 'บันทึกเมื่อ'];

const TEXT_FIELDS = ['arm', 'posture', 'sugarTiming', 'symptoms', 'note'];
const NUM_FIELDS = ['sys', 'dia', 'pulse', 'weight', 'sugar', 'temp', 'spo2', 'sleep'];

/* ================= ตั้งค่าครั้งแรก ================= */

/** รันครั้งแรก: สร้างชีต + สร้างรหัสลับ (ดูได้ใน Execution log) */
function setup() {
  getSheet_();
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('API_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('API_KEY', key);
  }
  Logger.log('รหัสลับ (API key) ของคุณคือ: ' + key);
  Logger.log('คัดลอกไปใส่ในหน้า "ตั้งค่า" ของเว็บแอป — ห้ามแชร์ให้คนอื่น');
}

/** ดูรหัสลับอีกครั้ง */
function showKey() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('API_KEY') || 'ยังไม่มี — รัน setup() ก่อน');
}

/** เปลี่ยนรหัสลับใหม่ (ถ้าสงสัยว่าหลุด) — อุปกรณ์เดิมต้องใส่รหัสใหม่ */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('API_KEY');
  setup();
}

/* ================= API ================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) return json_({ ok: true, app: 'health-log' });
  return handle_(p.action, p);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  return handle_(body.action, body);
}

function handle_(action, p) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!key) return json_({ ok: false, error: 'not_setup' });
    if (p.key !== key) return json_({ ok: false, error: 'bad_key' });
    switch (action) {
      case 'ping':   return json_({ ok: true });
      case 'list':   return json_({ ok: true, records: getRecords() });
      case 'add':    return json_(addRecord(p.record || {}));
      case 'delete': return json_(deleteRecord(p.id));
      default:       return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= ชีต ================= */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#FDE6E7').setWrap(true);
    sh.hideColumns(1);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('Q:Q').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sh;
}

function getRecords() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  const vals = sh.getRange(2, 1, n, KEYS.length).getValues();
  return vals
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      KEYS.forEach((k, i) => {
        let v = r[i];
        if (v instanceof Date) v = v.toISOString();
        o[k] = v === '' ? null : v;
      });
      return o;
    });
}

function clean_(rec, k) {
  const v = rec[k];
  if (v === undefined || v === null || v === '') return '';
  if (NUM_FIELDS.indexOf(k) >= 0) {
    const n = Number(v);
    return isFinite(n) ? n : '';
  }
  let s = String(v).slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // กันสูตรชีต
  return s;
}

function addRecord(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const id = Utilities.getUuid();
    const d = rec.datetime ? new Date(rec.datetime) : new Date();
    const row = KEYS.map(k => {
      if (k === 'id') return id;
      if (k === 'createdAt') return new Date();
      if (k === 'datetime') return isNaN(d.getTime()) ? new Date() : d;
      if (k === 'med') return rec.med ? 'ใช่' : '';
      return clean_(rec, k);
    });
    sh.appendRow(row);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(id) {
  if (!id) return { ok: false, error: 'no_id' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const n = sh.getLastRow() - 1;
    if (n <= 0) return { ok: false, error: 'not_found' };
    const ids = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, error: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}
