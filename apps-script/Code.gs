/**
 * ดันดี (DunDee) — Backend API (Google Apps Script)
 * หน้าเว็บอยู่บน GitHub Pages ส่วนไฟล์นี้ทำหน้าที่เป็น API อ่าน/เขียน Google Sheet
 *
 * วิธีใช้ (ไม่ต้องแตะโค้ด): เปิดชีต → เมนู "❤️ ดันดี"
 *   1) เริ่มต้นใช้งาน   → อนุญาตสิทธิ์ + สร้างรหัสลับ
 *   (Deploy เป็น Web app: Execute as = Me, Who has access = Anyone)
 *   2) เชื่อมกับแอป     → ได้ QR/ลิงก์ ไปเปิดแอปบนมือถือ ตั้งค่าอัตโนมัติ
 */

const SHEET_NAME = 'บันทึกสุขภาพ';
const APP_URL = 'https://bearpanda2012-wq.github.io/health-log/';

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

/**
 * อ่านรหัสลับของชีตนี้
 * รหัสผูกกับ ID ของชีต — ถ้าชีตถูก "ทำสำเนา" (ID เปลี่ยน) รหัสเดิมจะใช้ไม่ได้ ต้องสร้างใหม่
 * กันไม่ให้ทุกสำเนาจากชีตต้นแบบใช้รหัสเดียวกัน
 */
function getKey_() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('API_KEY');
  if (!key) return null;
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const bound = props.getProperty('API_SHEET');
  if (!bound) { props.setProperty('API_SHEET', ssId); return key; } // รุ่นเก่า: ผูกกับชีตปัจจุบัน
  return bound === ssId ? key : null;
}

/** สร้างชีต + รหัสลับ (ถ้ายังไม่มี) แล้วคืนรหัส */
function setup() {
  getSheet_();
  let key = getKey_();
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    const props = PropertiesService.getScriptProperties();
    props.setProperty('API_KEY', key);
    props.setProperty('API_SHEET', SpreadsheetApp.getActiveSpreadsheet().getId());
  }
  Logger.log('รหัสลับ (API key) ของคุณคือ: ' + key);
  Logger.log('คัดลอกไปใส่ในหน้า "ตั้งค่า" ของเว็บแอป — ห้ามแชร์ให้คนอื่น');
  return key;
}

/** ดูรหัสลับอีกครั้ง */
function showKey() {
  Logger.log(getKey_() || 'ยังไม่มี — รัน setup() ก่อน');
}

/** เปลี่ยนรหัสลับใหม่ (ถ้าสงสัยว่าหลุด) — อุปกรณ์เดิมต้องเชื่อมใหม่ */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('API_KEY');
  return setup();
}

/* ================= เมนูในชีต ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('❤️ ดันดี')
    .addItem('1) เริ่มต้นใช้งาน', 'menuStart')
    .addItem('2) เชื่อมกับแอป (QR / ลิงก์)', 'menuConnect')
    .addSeparator()
    .addItem('เปลี่ยนรหัสลับใหม่', 'menuResetKey')
    .addToUi();
}

function webAppUrl_() {
  try {
    const u = ScriptApp.getService().getUrl();
    return u && /\/exec$/.test(u) ? u : '';
  } catch (e) { return ''; }
}

const DIALOG_CSS = '<style>body{font-family:"IBM Plex Sans Thai",Arial,sans-serif;color:#1E2A3A;font-size:14px;line-height:1.6;margin:0;padding:4px 6px}'
  + 'h2{color:#E8505B;font-size:20px;margin:0 0 8px}ol{padding-left:20px;margin:8px 0}li{margin-bottom:6px}'
  + 'b.hl{background:#FDE6E7;padding:1px 6px;border-radius:6px}.ok{color:#1FA38A;font-weight:600}'
  + '.btn{display:inline-block;background:#E8505B;color:#fff;border:0;border-radius:10px;padding:10px 16px;font:inherit;font-weight:600;text-decoration:none;cursor:pointer}'
  + '.btn2{display:inline-block;background:#fff;color:#1E2A3A;border:1px solid #F0E4DC;border-radius:10px;padding:9px 14px;font:inherit;cursor:pointer}'
  + 'input{width:100%;box-sizing:border-box;font:inherit;padding:9px 10px;border:1px solid #F0E4DC;border-radius:10px}'
  + '.muted{color:#6B7280;font-size:12.5px}.center{text-align:center}#qr svg{width:220px;height:220px}</style>'
  + '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;600&display=swap" rel="stylesheet">';

function showDialog_(html, title, h) {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(DIALOG_CSS + html).setWidth(460).setHeight(h || 460), title);
}

function menuStart() {
  setup();
  const deployed = !!webAppUrl_();
  showDialog_(
    '<h2>ขั้นที่ 1 เสร็จแล้ว ✓</h2>'
    + '<p class="ok">สร้างตารางบันทึกและรหัสลับเรียบร้อย</p>'
    + (deployed
      ? '<p>ชีตนี้เผยแพร่เป็นเว็บแอปแล้ว ไปที่เมนู <b class="hl">❤️ ดันดี → 2) เชื่อมกับแอป</b> ได้เลย</p>'
      : '<p><b>ขั้นที่ 2: เผยแพร่เป็นเว็บแอป</b> (ทำครั้งเดียว)</p><ol>'
        + '<li>เมนู <b class="hl">ส่วนขยาย → Apps Script</b></li>'
        + '<li>ปุ่มสีน้ำเงินมุมขวาบน <b class="hl">การทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่</b></li>'
        + '<li>กดเฟือง ⚙ เลือก <b class="hl">เว็บแอป</b></li>'
        + '<li>ดำเนินการในฐานะ: <b class="hl">ฉัน</b> · ผู้ที่มีสิทธิ์เข้าถึง: <b class="hl">ทุกคน</b></li>'
        + '<li>กด <b class="hl">ทำให้ใช้งานได้</b> → ถ้าขอสิทธิ์ ให้ติ๊ก <b class="hl">เลือกทั้งหมด</b> → อนุญาต → <b>เสร็จสิ้น</b></li>'
        + '<li>กลับมาที่ชีตนี้ เมนู <b class="hl">❤️ ดันดี → 2) เชื่อมกับแอป</b></li></ol>'
        + '<p class="muted">⚠️ ตอนขอสิทธิ์: ถ้าขึ้น "Google ยังไม่ได้ยืนยันแอปนี้" ห้ามกดปุ่มสีน้ำเงิน ให้กดลิงก์เล็ก <b>ขั้นสูง</b> → <b>ไปที่ ดันดี</b> และต้องติ๊ก <b>เลือกทั้งหมด</b> ก่อนกดดำเนินการต่อ ไม่งั้นจะขึ้น "Access is denied" (สคริปต์นี้อยู่ในบัญชีของคุณเอง ปลอดภัย)</p>')
    + '<p class="center"><button class="btn2" onclick="google.script.host.close()">ปิด</button></p>',
    'ดันดี — เริ่มต้นใช้งาน', deployed ? 260 : 520);
}

function menuConnect() {
  const key = setup();
  const data = JSON.stringify({ url: webAppUrl_(), key: key, app: APP_URL }).replace(/</g, '\\u003c');
  showDialog_(
    '<h2>เชื่อมกับแอปดันดี</h2>'
    + '<p id="chk" class="center">⏳ กำลังตรวจสอบการเชื่อมต่อ…</p>'
    + '<div id="ask" style="display:none"><p>วาง <b>URL ของเว็บแอป</b> (ลงท้ายด้วย <code>/exec</code>)</p>'
    + '<p class="muted">หาได้ที่ ส่วนขยาย → Apps Script → การทำให้ใช้งานได้ → <b>จัดการการทำให้ใช้งานได้</b> → กด "คัดลอก" ใต้ URL</p>'
    + '<input id="u" placeholder="https://script.google.com/macros/s/…/exec"><p><button class="btn" id="go" onclick="useUrl()">ตรวจสอบและสร้าง QR</button></p>'
    + '<p id="err" style="display:none;color:#C93B46"></p>'
    + '<p class="muted">ยังไม่ได้ Deploy? ดูเมนู ❤️ ดันดี → 1) เริ่มต้นใช้งาน</p></div>'
    + '<div id="done" style="display:none" class="center">'
    + '<p>📱 <b>มือถือ:</b> เปิดกล้องสแกน QR แล้วแตะลิงก์</p><div id="qr"></div>'
    + '<p><a class="btn" id="open" target="_blank">💻 เปิดแอปบนเครื่องนี้</a> <button class="btn2" onclick="copy()">คัดลอกลิงก์</button></p>'
    + '<p class="muted">QR และลิงก์นี้มีรหัสลับอยู่ข้างใน ห้ามให้คนอื่นเห็น</p></div>'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"></script>'
    + '<script>var D=' + data + ',LINK="";'
    + 'function make(u){LINK=D.app+"#setup="+encodeURIComponent(btoa(JSON.stringify({url:u,key:D.key})));'
    + 'try{var q=qrcode(0,"M");q.addData(LINK);q.make();document.getElementById("qr").innerHTML=q.createSvgTag({cellSize:4,margin:2,scalable:true});}catch(e){document.getElementById("qr").textContent="(สร้าง QR ไม่ได้ ใช้ปุ่มคัดลอกลิงก์แทน)";}'
    + 'document.getElementById("open").href=LINK;document.getElementById("chk").style.display="none";document.getElementById("ask").style.display="none";document.getElementById("done").style.display="block";}'
    // ทดสอบ URL จริงก่อนสร้าง QR (Google บางครั้งคืน URL ที่ใช้ไม่ได้)
    + 'function $(i){return document.getElementById(i)}'
    + 'function verify(u,cb){try{fetch(u+"?action=ping&key="+encodeURIComponent(D.key)+"&t="+Date.now(),{credentials:"omit",cache:"no-store"})'
    + '.then(function(r){return r.text()}).then(function(t){var ok=false;try{ok=JSON.parse(t).ok===true}catch(e){}cb(ok?"ok":(/<html/i.test(t)?"html":"bad"))})'
    + '.catch(function(){cb("net")})}catch(e){cb("net")}}'
    + 'function ask(msg){$("chk").style.display="none";$("ask").style.display="block";if(msg){$("err").textContent=msg;$("err").style.display="block"}}'
    + 'var FORCE="";'
    + 'function useUrl(){var u=$("u").value.trim();if(!/^https:\\/\\/script\\.google\\.com\\/.+\\/exec$/.test(u)){ask("URL ต้องขึ้นต้นด้วย https://script.google.com/ และลงท้ายด้วย /exec");return;}'
    + 'if(FORCE===u){make(u);return;}$("go").disabled=true;$("go").textContent="กำลังตรวจสอบ…";'
    + 'verify(u,function(r){$("go").disabled=false;$("go").textContent="ตรวจสอบและสร้าง QR";if(r==="ok"){make(u);return;}FORCE=u;'
    + 'ask(r==="html"?"URL นี้ยังใช้ไม่ได้: ตรวจว่า Deploy แบบ ผู้ที่มีสิทธิ์เข้าถึง = ทุกคน แล้วลองใหม่ (หรือกดปุ่มอีกครั้งเพื่อใช้ URL นี้อยู่ดี)":"ตรวจสอบไม่ผ่าน ลองใหม่อีกครั้ง หรือกดปุ่มอีกครั้งเพื่อใช้ URL นี้อยู่ดี")})}'
    + 'function copy(){var t=document.createElement("textarea");t.value=LINK;document.body.appendChild(t);t.select();document.execCommand("copy");t.remove();alert("คัดลอกลิงก์แล้ว ส่งเข้ามือถือตัวเองได้เลย");}'
    + 'if(D.url){verify(D.url,function(r){if(r==="ok"){$("chk").style.display="none";make(D.url)}else ask("")})}else ask("");</script>',
    'ดันดี — เชื่อมกับแอป', 480);
}

function menuResetKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('เปลี่ยนรหัสลับใหม่?', 'ทุกเครื่องที่เชื่อมไว้จะหลุด ต้องสแกน QR ใหม่ (ข้อมูลในชีตไม่หาย)', ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;
  resetKey();
  menuConnect();
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
    const key = getKey_();
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
