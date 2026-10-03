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
  'symptoms', 'note', 'createdAt', 'sugarLevel'];

const HEADERS = ['id', 'วันเวลาที่วัด', 'ตัวบน SYS (mmHg)', 'ตัวล่าง DIA (mmHg)',
  'ชีพจร (ครั้ง/นาที)', 'แขน', 'ท่า', 'น้ำหนัก (กก.)', 'น้ำตาล (mg/dL)',
  'ช่วงวัดน้ำตาล', 'อุณหภูมิ (°C)', 'SpO2 (%)', 'นอน (ชม.)', 'กินยาแล้ว',
  'อาการ', 'โน้ต', 'บันทึกเมื่อ', 'ระดับน้ำตาล'];

const TEXT_FIELDS = ['arm', 'posture', 'sugarTiming', 'sugarLevel', 'symptoms', 'note'];
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
    .addItem('3) เปิดอ่านค่าจากรูป (AI)', 'menuOcr')
    .addItem('🔄 อัปเดตแอป (หลังวางโค้ดใหม่)', 'menuUpdate')
    .addSeparator()
    .addItem('เปลี่ยนรหัสลับใหม่', 'menuResetKey')
    .addToUi();
}

function webAppUrl_() {
  const saved = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  if (saved) return saved;
  try {
    const u = ScriptApp.getService().getUrl();
    return u && /\/exec$/.test(u) ? u : '';
  } catch (e) { return ''; }
}

/* ===== Deploy อัตโนมัติผ่าน Apps Script API (ไม่ต้องกด Deploy เอง) =====
   ต้องมีไฟล์ appsscript.json (มี webapp + scope script.projects/deployments)
   และผู้ใช้ต้องเปิดสวิตช์ "Google Apps Script API" ที่ script.google.com/home/usersettings ครั้งเดียว */
function scriptApi_(method, path, body) {
  const res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + path, {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: body ? JSON.stringify(body) : undefined
  });
  const code = res.getResponseCode(), txt = res.getContentText();
  let j = {}; try { j = JSON.parse(txt); } catch (e) {}
  if (code >= 200 && code < 300) return j;
  const msg = (j.error && j.error.message) || txt.slice(0, 200);
  const err = new Error(msg);
  // สวิตช์ระดับผู้ใช้ (script.google.com/home/usersettings) ยังไม่เปิด → ให้ผู้ใช้เปิดเอง
  // ส่วนข้อผิดพลาดอื่น (เช่น API ปิดในโปรเจกต์ Cloud) → ถอยไปใช้วิธี Deploy เอง ไม่ถามวนซ้ำ
  err.needApi = /User has not enabled the Apps Script API|usersettings/i.test(msg);
  console.error('Apps Script API ' + code + ': ' + msg);
  err.code = code;
  throw err;
}

// สร้าง/อัปเดตเว็บแอปเป็นเวอร์ชันล่าสุด → คืน URL
function autoDeploy_() {
  const props = PropertiesService.getScriptProperties();
  const ver = scriptApi_('post', '/versions', { description: 'ดันดี ' + new Date().toISOString().slice(0, 10) });
  const cfg = { versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: 'ดันดี เว็บแอป' };
  let dep, depId = props.getProperty('WEBAPP_DEPLOY_ID');
  if (!depId) {
    // เคย Deploy เองไว้แล้ว → ใช้ตัวเดิม (URL เดิม มือถือไม่หลุด)
    try {
      const list = (scriptApi_('get', '/deployments?pageSize=50').deployments || []).filter(function (d) {
        return d.deploymentConfig && d.deploymentConfig.versionNumber
          && (d.entryPoints || []).some(function (x) { return x.entryPointType === 'WEB_APP'; });
      });
      let cur = ''; try { cur = ScriptApp.getService().getUrl() || ''; } catch (e) {}
      const hit = list.filter(function (d) { return cur && cur.indexOf(d.deploymentId) >= 0; })[0]
        || list.sort(function (a, b) { return String(b.updateTime).localeCompare(String(a.updateTime)); })[0];
      if (hit) depId = hit.deploymentId;
    } catch (e) { if (e.needApi) throw e; }
  }
  if (depId) {
    try { dep = scriptApi_('put', '/deployments/' + depId, { deploymentConfig: cfg }); } catch (e) { if (e.needApi) throw e; dep = null; }
  }
  if (!dep) dep = scriptApi_('post', '/deployments', cfg);
  const ep = (dep.entryPoints || []).filter(function (x) { return x.entryPointType === 'WEB_APP'; })[0];
  const url = ep && ep.webApp && ep.webApp.url;
  if (!url) throw new Error('สร้างเว็บแอปไม่สำเร็จ (ไม่พบ URL) — ตรวจว่ามีไฟล์ appsscript.json ที่ตั้ง webapp ไว้');
  props.setProperty('WEBAPP_DEPLOY_ID', dep.deploymentId);
  props.setProperty('WEBAPP_URL', url);
  return url;
}

function needApiDialog_(retryFn, err) {
  const tries = Number(PropertiesService.getUserProperties().getProperty('API_TRIES') || 0) + 1;
  PropertiesService.getUserProperties().setProperty('API_TRIES', String(tries));
  showDialog_(
    '<h2>เหลืออีกขั้นเดียว</h2>'
    + '<p>เปิดสวิตช์ให้ดันดีตั้งค่าเว็บแอปให้อัตโนมัติ (ทำครั้งเดียว)</p><ol>'
    + '<li>กด → <a class="btn" style="padding:6px 12px" target="_blank" href="https://script.google.com/home/usersettings">เปิดหน้าตั้งค่า</a></li>'
    + '<li>📱 มือถือ: หน้าใหม่ให้ขอ "เว็บไซต์เดสก์ท็อป" ก่อน</li>'
    + '<li>แตะสวิตช์ <b class="hl">Google Apps Script API</b> ให้เป็น <b class="hl">เปิด</b></li>'
    + '<li>กลับมาที่ชีตนี้ แล้วกดปุ่มด้านล่าง</li></ol>'
    + '<p class="center"><button class="btn" id="r" onclick="this.disabled=true;this.textContent=\'กำลังตั้งค่า…\';google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message);document.getElementById(\'r\').disabled=false}).' + retryFn + '()">ลองอีกครั้ง</button></p>'
    + '<p class="muted">เพิ่งเปิดสวิตช์? บางครั้งต้องรอ 1–2 นาที</p>'
    + (tries > 1 ? '<p class="muted">ยังไม่ได้ผล? <a href="#" onclick="google.script.run.withSuccessHandler(function(){}).menuManualDeploy();return false">ข้ามไปตั้งค่าเอง (Deploy เอง)</a></p>' : '')
    + (err ? '<p class="muted" style="font-size:11px;word-break:break-all">รายละเอียด: ' + String(err.message).replace(/</g, '&lt;').slice(0, 200) + '</p>' : ''),
    'ดันดี — เปิดสวิตช์', 460);
}

function menuManualDeploy() {
  PropertiesService.getUserProperties().setProperty('SKIP_AUTO', '1');
  menuStart();
}

function menuUpdate() {
  try {
    autoDeploy_();
    SpreadsheetApp.getUi().alert('อัปเดตแอปแล้ว ✓', 'เว็บแอปใช้โค้ดล่าสุดแล้ว ลิงก์และ QR เดิมใช้ได้เหมือนเดิม', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    if (e.needApi) return needApiDialog_('menuUpdate', e);
    SpreadsheetApp.getUi().alert('อัปเดตไม่สำเร็จ', String(e.message), SpreadsheetApp.getUi().ButtonSet.OK);
  }
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
  let autoErr = null;
  const up = PropertiesService.getUserProperties();
  const skipAuto = up.getProperty('SKIP_AUTO') === '1';
  if (!skipAuto && !PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') && !webAppUrl_()) {
    try { autoDeploy_(); up.deleteProperty('API_TRIES'); } catch (e) { autoErr = e; }
  }
  if (!autoErr && PropertiesService.getScriptProperties().getProperty('WEBAPP_URL')) return menuConnect(); // พร้อมแล้ว → ไปหน้าคัดลอกลิงก์เลย
  if (autoErr && autoErr.needApi) return needApiDialog_('menuStart', autoErr);
  const deployed = !!webAppUrl_();
  showDialog_(
    '<h2>ขั้นที่ 1 เสร็จแล้ว ✓</h2>'
    + '<p class="ok">สร้างตารางบันทึกและรหัสลับเรียบร้อย</p>'
    + (autoErr && !deployed ? '<p class="muted" style="font-size:11px">ตั้งค่าอัตโนมัติไม่สำเร็จ จึงต้อง Deploy เอง (' + String(autoErr.message).replace(/</g, '&lt;').slice(0, 160) + ')</p>' : '')
    + (deployed
      ? '<p>ชีตนี้เผยแพร่เป็นเว็บแอปแล้ว ไปที่เมนู <b class="hl">❤️ ดันดี → 2) เชื่อมกับแอป</b> ได้เลย</p>'
      : '<p><b>ขั้นที่ 2: เผยแพร่เป็นเว็บแอป</b> (ทำครั้งเดียว)</p><ol>'
        + '<li>กดปุ่มนี้ → <a class="btn" style="padding:6px 12px" target="_blank" href="https://script.google.com/d/' + ScriptApp.getScriptId() + '/edit">เปิดหน้า Apps Script</a>'
        + '<br><span class="muted">📱 มือถือ: หน้าใหม่ที่เปิดขึ้น ต้องกดขอ "เว็บไซต์เดสก์ท็อป" อีกครั้ง</span></li>'
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
    + '<p style="margin:4px 0"><button class="btn" style="font-size:16px;padding:12px 18px" onclick="copy()">📋 คัดลอกลิงก์เชื่อมต่อ</button></p>'
    + '<p class="muted" style="text-align:left">📱 <b>ตั้งค่าบนมือถือเครื่องนี้:</b> กดคัดลอก → เปิดแอปดันดี (ไอคอนหน้าจอโฮม) → กดปุ่ม <b>วางลิงก์เชื่อมต่อ</b></p>'
    + '<p class="muted" style="text-align:left">💻 <b>ตั้งค่าบนคอม:</b> ใช้มือถือสแกน QR นี้ หรือ <a id="open" target="_blank">เปิดแอปบนเครื่องนี้</a></p><div id="qr"></div>'
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
    + 'function copy(){var done=function(){alert("คัดลอกแล้ว ✓ เปิดแอปดันดี แล้วกด วางลิงก์เชื่อมต่อ")};'
    + 'try{navigator.clipboard.writeText(LINK).then(done,fb)}catch(e){fb()}'
    + 'function fb(){var t=document.createElement("textarea");t.value=LINK;document.body.appendChild(t);t.select();document.execCommand("copy");t.remove();done()}}'
    + 'if(D.url){verify(D.url,function(r){if(r==="ok"){$("chk").style.display="none";make(D.url)}else ask("")})}else ask("");</script>',
    'ดันดี — เชื่อมกับแอป', 560);
}

function menuResetKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('เปลี่ยนรหัสลับใหม่?', 'ทุกเครื่องที่เชื่อมไว้จะหลุด ต้องสแกน QR ใหม่ (ข้อมูลในชีตไม่หาย)', ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;
  resetKey();
  menuConnect();
}

/* ================= อ่านค่าจากรูป (Gemini) ================= */

const GEMINI_MODELS = ['gemini-flash-lite-latest', 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-flash-latest', 'gemini-3.5-flash'];

function menuOcr() {
  const has = !!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  showDialog_(
    '<h2>อ่านค่าจากรูปด้วย AI</h2>'
    + (has ? '<p class="ok">✓ เปิดใช้งานอยู่แล้ว วางรหัสใหม่ด้านล่างถ้าต้องการเปลี่ยน</p>' : '')
    + '<ol><li>เปิด <a href="https://aistudio.google.com/apikey" target="_blank">aistudio.google.com/apikey</a> (ล็อกอินบัญชี Google เดียวกัน)</li>'
    + '<li>กด <b class="hl">Create API key</b> แล้วกด <b class="hl">Copy key</b></li>'
    + '<li>วางด้านล่างแล้วกด <b class="hl">บันทึกและทดสอบ</b></li></ol>'
    + '<input id="k" placeholder="วางรหัส Gemini ที่นี่" autocomplete="off"><p><button class="btn" id="go" onclick="save()">บันทึกและทดสอบ</button></p>'
    + '<p id="m"></p>'
    + '<p class="muted">รหัสเก็บในชีตนี้เท่านั้น · รูปจะถูกส่งให้ Google Gemini อ่านตัวเลข · รหัสฟรีมีโควตาต่อวันเพียงพอสำหรับใช้ที่บ้าน</p>'
    + '<p class="muted">⚠️ หลังบันทึกครั้งแรก ต้อง Deploy เวอร์ชันใหม่: ส่วนขยาย → Apps Script → การทำให้ใช้งานได้ → จัดการ → ✏️ → เวอร์ชันใหม่</p>'
    + '<script>function save(){var k=document.getElementById("k").value.trim(),m=document.getElementById("m"),b=document.getElementById("go");'
    + 'if(!/^[\\w.\\-]{20,}$/.test(k)){m.style.color="#C93B46";m.textContent="รหัสไม่ถูกต้อง ลองกด Copy key แล้ววางใหม่";return;}'
    + 'b.disabled=true;m.style.color="";m.textContent="⏳ กำลังทดสอบ…";'
    + 'google.script.run.withSuccessHandler(function(r){b.disabled=false;m.style.color=r.ok?"#1FA38A":"#C93B46";m.textContent=r.ok?"✓ ใช้ได้แล้ว! ("+r.model+") ในแอปกดปุ่ม 📷 อ่านค่าจากรูปได้เลย":"ใช้ไม่ได้: "+r.error;})'
    + '.withFailureHandler(function(e){b.disabled=false;m.style.color="#C93B46";m.textContent="ผิดพลาด: "+e.message;}).saveGeminiKey(k);}</script>',
    'ดันดี — อ่านค่าจากรูป', 470);
}

function saveGeminiKey(k) {
  k = String(k || '').trim();
  PropertiesService.getScriptProperties().deleteProperty('GEMINI_MODEL2');
  const r = gemini_(k, null, null, 'ตอบคำว่า ok เป็น JSON {"ok":true}');
  if (!r.ok) return r;
  PropertiesService.getScriptProperties().setProperty('GEMINI_KEY', k);
  return { ok: true, model: r.model };
}

// หารุ่น Gemini ที่บัญชีนี้ใช้ได้ (ชื่อรุ่นเปลี่ยนบ่อย) — เลือกรุ่น flash ใหม่สุด
function geminiModels_(key) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('GEMINI_MODEL2');
  let list = [];
  try {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200',
      { muteHttpExceptions: true, headers: { 'x-goog-api-key': key } });
    if (res.getResponseCode() === 200) {
      const ver = function (n) { const m = n.match(/gemini-(\d+(?:\.\d+)?)/); return m ? parseFloat(m[1]) : 0; };
      list = (JSON.parse(res.getContentText()).models || [])
        .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
        .map(function (m) { return String(m.name).replace(/^models\//, ''); })
        .filter(function (n) { return /flash/.test(n) && !/(preview|exp|tts|image|audio|live|thinking|embedding)/.test(n); })
        .sort(function (a, b) { return (/lite/.test(b) - /lite/.test(a)) || (ver(b) - ver(a)); }); // รุ่น lite ตอบเร็วกว่า พออ่านตัวเลข
    }
  } catch (e) {}
  console.log('gemini models', list.join(', ') || '(list failed)');
  const all = [saved].concat(list, GEMINI_MODELS).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
  return all;
}

function gemini_(key, b64, mime, prompt, full) {
  const parts = [{ text: prompt }];
  if (b64) parts.unshift({ inline_data: { mime_type: mime || 'image/jpeg', data: b64 } });
  const mk = function (think) {
    const gc = { temperature: 0, responseMimeType: 'application/json' };
    if (think) {
      gc.thinkingConfig = { thinkingLevel: 'minimal' }; // ให้ตอบเร็ว ไม่ต้องคิดนาน
      if (b64) gc.mediaResolution = 'MEDIA_RESOLUTION_LOW'; // รูปใช้ token น้อยลง เร็วขึ้น
    }
    return JSON.stringify({ contents: [{ parts: parts }], generationConfig: gc });
  };
  const t0 = Date.now();
  const saved = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL2');
  const models = (saved && !full) ? [saved] : geminiModels_(key).filter(function (m) { return !(full && m === saved); });
  let last = 'no_model';
  for (let i = 0; i < models.length && i < 4; i++) {
    if (Date.now() - t0 > 60000) { last = 'ช้าเกินไป (' + last + ')'; break; }
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent';
    const opt = { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key } };
    const t1 = Date.now();
    opt.payload = mk(true);
    let res = UrlFetchApp.fetch(url, opt);
    if (res.getResponseCode() === 400 && /thinking|media_?resolution/i.test(res.getContentText())) { opt.payload = mk(false); res = UrlFetchApp.fetch(url, opt); }
    if ([500, 502, 503, 504].indexOf(res.getResponseCode()) >= 0 && Date.now() - t0 < 40000) { // ไม่ว่างชั่วคราว → รอแล้วลองซ้ำ 1 ครั้ง
      Utilities.sleep(1500); res = UrlFetchApp.fetch(url, opt);
    }
    const code = res.getResponseCode(), txt = res.getContentText();
    console.log('gemini', models[i], code, (Date.now() - t1) + 'ms', code === 200 ? '' : errMsg_(txt));
    if (code === 400 && /API key not valid|API_KEY_INVALID/.test(txt)) return { ok: false, error: 'รหัส Gemini ไม่ถูกต้อง' };
    if (code === 401 || code === 403) return { ok: false, error: 'รหัสนี้ใช้กับ Gemini ไม่ได้ (' + code + ') ' + errMsg_(txt) };
    if (code === 429) return { ok: false, error: 'ใช้เกินโควตาวันนี้ ลองใหม่พรุ่งนี้' };
    if (code !== 200) { last = models[i] + ': HTTP ' + code + ' ' + errMsg_(txt); continue; }
    PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL2', models[i]);
    try {
      const j = JSON.parse(txt), t = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join('');
      return { ok: true, model: models[i], data: JSON.parse(t.replace(/^```(json)?|```$/g, '').trim()) };
    } catch (e) { return { ok: false, error: 'อ่านคำตอบไม่ได้' }; }
  }
  if (saved && !full && Date.now() - t0 < 45000) return gemini_(key, b64, mime, prompt, true); // รุ่นที่จำไว้ใช้ไม่ได้ → ลองรุ่นอื่น
  if (/HTTP 5\d\d/.test(last)) return { ok: false, error: 'busy', detail: last };
  return { ok: false, error: last };
}

function errMsg_(txt) {
  try { return String(JSON.parse(txt).error.message || '').slice(0, 160); } catch (e) { return String(txt || '').slice(0, 120); }
}

const OCR_PROMPT = 'This photo shows the display of a home medical device: either a blood pressure monitor '
  + '(SYS top/largest, DIA middle, PULSE bottom/smallest) or a blood glucose meter (one large number with mg/dL or mmol/L, e.g. Accu-Chek). '
  + 'Read the lit digits only; ignore faint unlit segments. If a date or time is shown on the display, read it too. '
  + 'Answer JSON only: {"device":"bp"|"glucose"|"other","sys":int|null,"dia":int|null,"pulse":int|null,'
  + '"glucose":number|null,"unit":"mg/dL"|"mmol/L"|null,"date":"YYYY-MM-DD"|null,"time":"HH:MM"|null,"sure":true|false}. '
  + 'Use null for anything you cannot read clearly. sure=false if any digit is uncertain or the device is not one of these.';

function ocrImage_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'ocr_off' };
  const b64 = String(p.image || '');
  if (!b64 || b64.length > 4e6) return { ok: false, error: 'bad_image' };
  const r = gemini_(key, b64, p.mime || 'image/jpeg', OCR_PROMPT);
  if (!r.ok) return r;
  const d = r.data || {}, n = function (v) { v = Number(v); return isFinite(v) && v > 0 ? Math.round(v) : null; };
  let glu = Number(d.glucose);
  if (!isFinite(glu) || glu <= 0) glu = null;
  else if (d.unit === 'mmol/L' || glu < 35) glu = Math.round(glu * 18); // แปลง mmol/L → mg/dL
  else glu = Math.round(glu);
  const dev = d.device === 'glucose' || (glu && !d.sys && !d.dia) ? 'glucose' : d.device === 'bp' ? 'bp' : (d.sys || d.dia ? 'bp' : 'other');
  return { ok: true, device: dev, glucose: dev === 'glucose' ? glu : null,
    sys: dev === 'bp' ? n(d.sys) : null, dia: dev === 'bp' ? n(d.dia) : null, pulse: dev === 'bp' ? n(d.pulse) : null,
    date: /^\d{4}-\d{2}-\d{2}$/.test(d.date || '') ? d.date : null,
    time: /^\d{1,2}:\d{2}$/.test(d.time || '') ? d.time : null, sure: d.sure !== false };
}

/* ================= ยาของฉัน (รูปยา + ชื่อ + วิธีกิน) ================= */

const MEDS_SHEET = 'ยาของฉัน';
const MED_HEADERS = ['id', 'ชื่อยา', 'วิธีกิน / หมายเหตุ', 'รูป', 'เพิ่มเมื่อ'];

function medSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MEDS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MEDS_SHEET);
    sh.appendRow(MED_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, MED_HEADERS.length).setFontWeight('bold').setBackground('#DDF3EE');
    sh.hideColumns(1); sh.setColumnWidth(4, 80);
  }
  return sh;
}

function getMeds_() {
  const sh = medSheet_(), n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, 5).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    return { id: r[0], name: String(r[1] || ''), note: String(r[2] || ''), photo: String(r[3] || ''),
      createdAt: r[4] instanceof Date ? r[4].toISOString() : r[4] };
  });
}

function addMed_(m) {
  const name = clean_({ v: m.name }, 'v'), note = clean_({ v: m.note }, 'v');
  let photo = String(m.photo || '');
  if (photo && (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(photo) || photo.length > 49000)) return { ok: false, error: 'bad_image' };
  if (!name && !photo) return { ok: false, error: 'empty' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const id = Utilities.getUuid();
    medSheet_().appendRow([id, name, note, photo, new Date()]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

function deleteMed_(id) {
  if (!id) return { ok: false, error: 'no_id' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = medSheet_(), n = sh.getLastRow() - 1;
    if (n <= 0) return { ok: false, error: 'not_found' };
    const ids = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { sh.deleteRow(i + 2); return { ok: true }; }
    return { ok: false, error: 'not_found' };
  } finally { lock.releaseLock(); }
}

const MED_PROMPT = 'This is a photo of a medicine package, blister pack, bottle or a pharmacy label (may be in Thai or English). '
  + 'Extract the drug name (generic and/or brand) with strength, and the dosing instructions if printed. '
  + 'Answer JSON only: {"name":string|null,"note":string|null}. Write "name" like "Amlodipine 5 mg" or the Thai name as printed. '
  + 'Write "note" in short Thai (e.g. "ครั้งละ 1 เม็ด วันละ 1 ครั้ง หลังอาหารเช้า"). Use null if not readable.';

function medOcr_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'ocr_off' };
  const b64 = String(p.image || '');
  if (!b64 || b64.length > 4e6) return { ok: false, error: 'bad_image' };
  const r = gemini_(key, b64, p.mime || 'image/jpeg', MED_PROMPT);
  if (!r.ok) return r;
  const d = r.data || {};
  return { ok: true, name: d.name ? String(d.name).slice(0, 120) : null, note: d.note ? String(d.note).slice(0, 300) : null };
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
      case 'ping':   return json_({ ok: true, ocr: !!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') });
      case 'ocr':    return json_(ocrImage_(p));
      case 'list':   return json_({ ok: true, records: getRecords() });
      case 'add':    return json_(addRecord(p.record || {}));
      case 'delete': return json_(deleteRecord(p.id));
      case 'update': return json_(updateRecord(p.id, p.record || {}));
      case 'meds':   return json_({ ok: true, meds: getMeds_() });
      case 'medAdd': return json_(addMed_(p.med || {}));
      case 'medDel': return json_(deleteMed_(p.id));
      case 'medOcr': return json_(medOcr_(p));
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
  // ชีตเก่าที่ยังไม่มีคอลัมน์ใหม่ (เช่น ระดับน้ำตาล) → เติมหัวคอลัมน์ให้
  const lastCol = sh.getLastColumn();
  if (lastCol < HEADERS.length) {
    sh.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol)
      .setValues([HEADERS.slice(lastCol)])
      .setFontWeight('bold').setBackground('#FDE6E7').setWrap(true);
  }
  return sh;
}

function getRecords() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  const vals = sh.getRange(2, 1, n, Math.min(KEYS.length, sh.getMaxColumns())).getValues();
  return vals
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      KEYS.forEach((k, i) => {
        let v = r[i];
        if (v instanceof Date) v = v.toISOString();
        o[k] = (v === '' || v === undefined) ? null : v;
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
      if (k === 'med') return rec.med === true ? 'ใช่' : (rec.med ? clean_(rec, 'med') : ''); // เช่น "เช้า หลังอาหาร, เย็น ก่อนอาหาร"
      return clean_(rec, k);
    });
    sh.appendRow(row);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function updateRecord(id, rec) {
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
        const old = sh.getRange(i + 2, 1, 1, KEYS.length).getValues()[0];
        const d = rec.datetime ? new Date(rec.datetime) : null;
        const row = KEYS.map((k, j) => {
          if (k === 'id') return id;
          if (k === 'createdAt') return old[j];
          if (k === 'datetime') return d && !isNaN(d.getTime()) ? d : old[j];
          if (k === 'med') return rec.med === true ? 'ใช่' : (rec.med ? clean_(rec, 'med') : '');
          return clean_(rec, k);
        });
        sh.getRange(i + 2, 1, 1, KEYS.length).setValues([row]);
        return { ok: true, id: id };
      }
    }
    return { ok: false, error: 'not_found' };
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
