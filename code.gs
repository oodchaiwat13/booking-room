const SHEET_NAME = 'Bookings';
const TZ = 'Asia/Bangkok';
const HEADER = ['ID','บันทึกเมื่อ','ห้อง','กลุ่มห้อง','สังกัด','ชื่อ-สกุล','สถานภาพ','วันที่','เริ่ม','สิ้นสุด','วัตถุประสงค์','สถานะ','ปีงบประมาณ'];

const ROOM_GROUPS = [
  { group: 'ห้องเรียน / ห้องประชุม', rooms: [
    'สทก.1109 (ห้องปฏิบัติการคอมพิวเตอร์ 40 ที่นั่ง)',
    'สทก.1110 (ห้องบรรยาย 80 ที่นั่ง)',
    'สทก.1117 (ห้องบรรยาย 20 ที่นั่ง)',
    'สทก.1421 (ห้องประชุม 20 ที่นั่ง)',
    'สทก.1423 (ห้องประชุมโต๊ะกลม 6 ที่นั่ง ด้านหน้า)',
    'ห้องประชุมสำนักงานคณบดี อาคารบริหาร ชั้น 4 (18 ที่นั่ง)'] },
  { group: 'ห้องปฏิบัติการชั้น 1', rooms: ['1104', 'wet lab'] },
  { group: 'ห้องปฏิบัติการชั้น 2', rooms: ['1202','1204','1206','1208','1210','1212','1214','1216','1216-1','1216-2','1218','1218-1','1218-2','1223','1227'] },
  { group: 'ห้องปฏิบัติการชั้น 3', rooms: ['1302','1303','1304','1305','1306','1307','1308','1310','1313','1314','1315','1316','1321','1325'] },
  { group: 'ห้องปฏิบัติการชั้น 4', rooms: ['1405','1406','1415','1419','Rumen in Vitro Lab'] }
];
const AFFILIATIONS = ['วท.บ.สัตวศาสตร์','วท.บ.วิทยาศาสตร์การประมง','วท.บ.เทคโนโลยีการผลิตพืช','วท.บ.ผู้ประกอบการธุรกิจการเกษตรและอาหาร','วท.บ.ส่งเสริมการเกษตร','วท.ม.สัตวศาสตร์','วท.ม.ชีววิทยาศาสตร์เพื่อเกษตรกรรมที่ยั่งยืน','ปร.ด.ชีววิทยาศาสตร์เพื่อเกษตรกรรมที่ยั่งยืน','สำนักงานคณบดี'];
const ROLES = ['นักศึกษา','อาจารย์','พนักงาน'];

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'dashboard') return loadPage_(['Dashboard', 'dashboard', 'Dashboard.html'], 'สถานะห้อง');
  return loadPage_(['Index', 'index', 'Index.html'], 'ระบบจองห้อง');
}

// โหลดไฟล์ HTML ถ้าไม่พบจะแสดงข้อความบอกชื่อไฟล์ที่ต้องสร้าง แทนการขึ้น Exception
function loadPage_(names, title) {
  for (let i = 0; i < names.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(names[i])
        .setTitle(title)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) { /* ลองชื่อถัดไป */ }
  }
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:24px"><h3>ไม่พบไฟล์ HTML ชื่อ "' + names[0] + '"</h3>' +
    '<p>ใน Apps Script กด + ข้าง "ไฟล์" > HTML > ตั้งชื่อว่า <b>' + names[0] + '</b> (ไม่ต้องใส่ .html) แล้ววางโค้ด จากนั้น Deploy เวอร์ชันใหม่</p></div>');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('H:J').setNumberFormat('@'); // เก็บวันที่/เวลาเป็นข้อความ กัน Sheets แปลงค่าเอง
  }
  return sh;
}

function s_(v, pattern) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, pattern) : String(v);
}

function readAll_() {
  const rows = getSheet_().getDataRange().getValues();
  rows.shift();
  return rows.filter(r => r[0]).map(r => ({
    id: r[0], room: r[2], group: r[3], affiliation: r[4], name: r[5], role: r[6],
    date: s_(r[7], 'yyyy-MM-dd'), start: s_(r[8], 'HH:mm'), end: s_(r[9], 'HH:mm'),
    purpose: r[10], status: r[11], fy: Number(r[12])
  }));
}

function fiscalYear_(dateStr) { // ปีงบประมาณไทย: 1 ต.ค. - 30 ก.ย. (พ.ศ.)
  const [y, m] = dateStr.split('-').map(Number);
  return (m >= 10 ? y + 1 : y) + 543;
}

function getConfig() {
  return { groups: ROOM_GROUPS, affiliations: AFFILIATIONS, roles: ROLES,
           today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
           nowTime: Utilities.formatDate(new Date(), TZ, 'HH:mm'),
           fy: fiscalYear_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')) };
}

function getBookings(dateStr) {
  return readAll_().filter(b => b.status === 'ใช้งาน' && b.date === dateStr)
    .sort((a, b) => a.start.localeCompare(b.start));
}

function submitBooking(b) {
  ['room','affiliation','name','role','date','start','end'].forEach(k => {
    if (!b[k] || !String(b[k]).trim()) throw new Error('กรุณากรอกข้อมูลให้ครบ');
  });
  if (b.end <= b.start) throw new Error('เวลาสิ้นสุดต้องหลังเวลาเริ่ม');
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (b.date < today) throw new Error('ไม่สามารถจองย้อนหลังได้');

  const group = ROOM_GROUPS.filter(g => g.rooms.indexOf(b.room) > -1)[0];
  if (!group) throw new Error('ไม่พบห้องที่เลือก');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const clash = getBookings(b.date).filter(x => x.room === b.room && b.start < x.end && b.end > x.start)[0];
    if (clash) throw new Error('ห้องไม่ว่าง: ' + clash.start + '-' + clash.end + ' จองโดย ' + clash.name);
    const id = Utilities.getUuid().slice(0, 8);
    getSheet_().appendRow([id, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), b.room, group.group,
      b.affiliation, b.name.trim(), b.role, b.date, b.start, b.end, (b.purpose || '').trim(), 'ใช้งาน', fiscalYear_(b.date)]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

function cancelBooking(id, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        if (String(rows[i][5]).trim() !== String(name).trim()) throw new Error('ชื่อไม่ตรงกับผู้จอง ไม่สามารถยกเลิกได้');
        sh.getRange(i + 1, 12).setValue('ยกเลิก');
        return { ok: true };
      }
    }
    throw new Error('ไม่พบรายการจอง');
  } finally { lock.releaseLock(); }
}

function getStats(fy) {
  const list = readAll_().filter(b => b.status === 'ใช้งาน' && b.fy === Number(fy));
  const hours = b => (Number(b.end.slice(0,2)) * 60 + Number(b.end.slice(3)) - Number(b.start.slice(0,2)) * 60 - Number(b.start.slice(3))) / 60;
  const tally = keyFn => {
    const m = {};
    list.forEach(b => { const k = keyFn(b); m[k] = m[k] || { key: k, count: 0, hours: 0 }; m[k].count++; m[k].hours += hours(b); });
    return Object.keys(m).map(k => ({ key: k, count: m[k].count, hours: Math.round(m[k].hours * 10) / 10 }));
  };
  const months = ['ต.ค.','พ.ย.','ธ.ค.','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.'];
  const byMonth = months.map((n, i) => ({ key: n, count: 0, hours: 0 }));
  list.forEach(b => { const mm = Number(b.date.slice(5, 7)); const i = (mm + 2) % 12; byMonth[i].count++; byMonth[i].hours += hours(b); });
  byMonth.forEach(x => x.hours = Math.round(x.hours * 10) / 10);
  const byRoom = tally(b => b.room).sort((a, b) => b.count - a.count);
  return { total: list.length, totalHours: Math.round(list.reduce((s, b) => s + hours(b), 0) * 10) / 10,
           byRoom: byRoom, byAffiliation: tally(b => b.affiliation).sort((a, b) => b.count - a.count),
           byRole: tally(b => b.role), byMonth: byMonth };
}
/* ============================================================
 * Line.gs — โมดูลแจ้งเตือน LINE Messaging API (ใช้ซ้ำข้ามโครงการได้)
 * ------------------------------------------------------------
 * วิธีใช้:
 *  1) สร้างไฟล์ Line.gs ในโครงการ Apps Script แล้ววางโค้ดทั้งหมดนี้
 *  2) ตั้งค่า Script Properties (Project Settings → Script Properties)
 *       LINE_CHANNEL_ACCESS_TOKEN = Channel access token (long-lived)
 *       LINE_TARGET_ID            = User ID (U...) / Group ID (C...) / Room ID (R...)
 *     หรือกรอกค่าใน setupLineProperties() แล้วรันครั้งเดียว (รันเสร็จให้ลบค่าออกจากโค้ด)
 *  3) รัน authorizeLineAccess() เพื่ออนุญาตสิทธิ์ + ทดสอบ Token
 *  4) รัน testLineNotification() เพื่อทดสอบส่งข้อความ
 *  5) ในโค้ดของโครงการ เรียกใช้: sendLineNotification_('ข้อความ');
 *
 * ถ้าโครงการมี appsscript.json ที่กำหนด oauthScopes เอง ต้องมี:
 *   "https://www.googleapis.com/auth/script.external_request"
 *
 * หมายเหตุ: ถ้าโครงการเดิมมีฟังก์ชัน clean_ หรือ getSetting_ อยู่แล้ว
 * ให้ลบตัวที่ซ้ำในไฟล์นี้ออก (Apps Script ไม่อนุญาตชื่อฟังก์ชันซ้ำกัน)
 * ============================================================ */

var LINE_API_BASE_ = 'https://api.line.me/v2/bot';
var LINE_MAX_TEXT_LENGTH_ = 5000;   // ข้อความตัวอักษรสูงสุดต่อ 1 บับเบิล
var LINE_MAX_MESSAGES_PER_PUSH_ = 5; // LINE จำกัด 5 บับเบิลต่อ 1 คำขอ

/* ---------- ตัวช่วยพื้นฐาน ---------- */

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

function getSetting_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/* ---------- ตั้งค่าครั้งแรก (ทางเลือก) ---------- */

/**
 * กรอกค่าแล้วรันครั้งเดียว จากนั้น "ลบค่าออกจากโค้ด" ทันที
 * เพื่อไม่ให้ Token ค้างอยู่ในซอร์สโค้ด
 */
function setupLineProperties() {
  var token = '';    // ใส่ Channel access token
  var targetId = ''; // ใส่ User ID / Group ID / Room ID
  if (!token || !targetId) throw new Error('กรอก token และ targetId ในฟังก์ชันนี้ก่อนรัน');
  PropertiesService.getScriptProperties().setProperties({
    LINE_CHANNEL_ACCESS_TOKEN: token,
    LINE_TARGET_ID: targetId
  });
  Logger.log('บันทึกค่าแล้ว — อย่าลืมลบ token/targetId ออกจากโค้ด');
}

/* ---------- ตรวจการตั้งค่า / สิทธิ์ ---------- */

function checkLineConfiguration() {
  var token = clean_(getSetting_('LINE_CHANNEL_ACCESS_TOKEN'));
  var targetId = clean_(getSetting_('LINE_TARGET_ID'));
  var result = {
    hasChannelAccessToken: Boolean(token),
    hasTargetId: Boolean(targetId),
    targetIdPrefix: targetId ? targetId.substring(0, 1) : '' // U=ผู้ใช้, C=กลุ่ม, R=ห้อง
  };
  Logger.log(JSON.stringify(result));
  if (!result.hasChannelAccessToken) throw new Error('ยังไม่ได้ตั้งค่า Script Property: LINE_CHANNEL_ACCESS_TOKEN');
  if (!result.hasTargetId) throw new Error('ยังไม่ได้ตั้งค่า Script Property: LINE_TARGET_ID');
  return result;
}

/** ทดสอบ Token + กระตุ้นให้ Apps Script ขอสิทธิ์ UrlFetchApp ครั้งแรก */
function authorizeLineAccess() {
  var token = clean_(getSetting_('LINE_CHANNEL_ACCESS_TOKEN'));
  if (!token) throw new Error('ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ใน Script Properties');
  var res = UrlFetchApp.fetch(LINE_API_BASE_ + '/info', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var status = res.getResponseCode();
  Logger.log('LINE API status: ' + status);
  Logger.log(res.getContentText());
  if (status !== 200) {
    throw new Error('ตรวจสอบ LINE Token ไม่สำเร็จ: HTTP ' + status + ' — ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

/* ---------- ส่งข้อความ ---------- */

/**
 * ส่งข้อความแบบ push
 * @param {string} message        ข้อความที่ต้องการส่ง (ยาวเกิน 5000 ตัวอักษรจะแบ่งเป็นหลายบับเบิลอัตโนมัติ)
 * @param {string=} optionalTargetId  ปลายทางเฉพาะกิจ (ถ้าไม่ระบุจะใช้ LINE_TARGET_ID)
 * @return {{success:boolean, statusCode:number, response:string}}
 */
function sendLineNotification_(message, optionalTargetId) {
  var token = clean_(getSetting_('LINE_CHANNEL_ACCESS_TOKEN'));
  var targetId = clean_(optionalTargetId) || clean_(getSetting_('LINE_TARGET_ID'));
  if (!token) throw new Error('ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ในพร็อพเพอร์ตี้ของสคริปต์');
  if (!targetId) throw new Error('ไม่พบ LINE_TARGET_ID ในพร็อพเพอร์ตี้ของสคริปต์');

  var text = String(message == null ? '' : message);
  if (!text.trim()) throw new Error('ข้อความว่างเปล่า');

  var chunks = splitLineText_(text);
  var last = null;
  // ส่งครั้งละไม่เกิน 5 บับเบิล
  for (var i = 0; i < chunks.length; i += LINE_MAX_MESSAGES_PER_PUSH_) {
    var batch = chunks.slice(i, i + LINE_MAX_MESSAGES_PER_PUSH_).map(function (t) {
      return { type: 'text', text: t };
    });
    last = linePost_('/message/push', token, { to: targetId, messages: batch });
  }
  return last;
}

/**
 * ส่งหาหลายคนพร้อมกัน (multicast) — รองรับเฉพาะ User ID (ขึ้นต้นด้วย U) สูงสุด 500 รายต่อครั้ง
 * @param {string[]} userIds
 * @param {string} message
 */
function sendLineMulticast_(userIds, message) {
  var token = clean_(getSetting_('LINE_CHANNEL_ACCESS_TOKEN'));
  if (!token) throw new Error('ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ในพร็อพเพอร์ตี้ของสคริปต์');
  var ids = (userIds || []).map(clean_).filter(String);
  if (!ids.length) throw new Error('ไม่มี User ID ปลายทาง');

  var text = String(message == null ? '' : message).substring(0, LINE_MAX_TEXT_LENGTH_);
  var last = null;
  for (var i = 0; i < ids.length; i += 500) {
    last = linePost_('/message/multicast', token, {
      to: ids.slice(i, i + 500),
      messages: [{ type: 'text', text: text }]
    });
  }
  return last;
}

/**
 * ส่งแบบไม่ให้โค้ดหลักล้ม — ถ้า LINE ผิดพลาดจะแค่บันทึก log แล้วคืน {success:false}
 * เหมาะกับการเรียกใน flow ที่ไม่อยากให้การแจ้งเตือนทำให้งานหลักหยุด
 */
function sendLineNotificationSafe_(message, optionalTargetId) {
  try {
    return sendLineNotification_(message, optionalTargetId);
  } catch (err) {
    console.error('LINE notify failed: ' + err.message);
    return { success: false, error: err.message };
  }
}

/* ---------- ตัวช่วยภายใน ---------- */

/** เรียก LINE API แบบ POST พร้อม retry เมื่อเจอ 429 / 5xx */
function linePost_(path, token, body) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };
  var maxAttempts = 3, res, status, content;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    res = UrlFetchApp.fetch(LINE_API_BASE_ + path, options);
    status = res.getResponseCode();
    content = res.getContentText();
    console.log('LINE HTTP status: ' + status + ' response: ' + content);
    var retryable = status === 429 || status >= 500;
    if (!retryable || attempt === maxAttempts) break;
    Utilities.sleep(1000 * attempt);
  }
  if (status < 200 || status >= 300) {
    throw new Error('LINE ส่งข้อความไม่สำเร็จ HTTP ' + status + ': ' + content);
  }
  return { success: true, statusCode: status, response: content };
}

/** แบ่งข้อความยาวเป็นชิ้นละไม่เกิน 5000 ตัวอักษร โดยพยายามตัดที่ขึ้นบรรทัดใหม่ */
function splitLineText_(text) {
  var chunks = [];
  var rest = text;
  while (rest.length > LINE_MAX_TEXT_LENGTH_) {
    var cut = rest.lastIndexOf('\n', LINE_MAX_TEXT_LENGTH_);
    if (cut < LINE_MAX_TEXT_LENGTH_ * 0.5) cut = LINE_MAX_TEXT_LENGTH_;
    chunks.push(rest.substring(0, cut));
    rest = rest.substring(cut).replace(/^\n/, '');
  }
  if (rest.length) chunks.push(rest);
  return chunks;
}

/* ---------- ทดสอบ ---------- */

function testLineNotification() {
  checkLineConfiguration();
  var tz = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var result = sendLineNotification_('✅ ทดสอบการแจ้งเตือนจาก Apps Script\nเวลา: ' + now);
  Logger.log(JSON.stringify(result));
  return result;
}

/* ---------- (ทางเลือก) หา Group ID / User ID ผ่าน Webhook ----------
 * 1) Deploy โครงการเป็น Web app (Execute as: Me, Access: Anyone)
 * 2) นำ URL ไปใส่ที่ LINE Developers → Messaging API → Webhook URL แล้วเปิด Use webhook
 * 3) เชิญบอทเข้ากลุ่ม แล้วพิมพ์ข้อความในกลุ่ม
 * 4) เปิด Executions/Logs จะเห็น ID ที่ต้องใช้ นำไปใส่ LINE_TARGET_ID
 *
 * ถ้าโครงการมี doPost(e) อยู่แล้ว ให้เรียก logLineSourceIds_(e) จากใน doPost เดิม
 * (หากไม่ใช้ ลบส่วนนี้ทิ้งได้)
 */
function logLineSourceIds_(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    (body.events || []).forEach(function (ev) {
      var s = ev.source || {};
      console.log('LINE source type=' + s.type +
        ' userId=' + (s.userId || '') +
        ' groupId=' + (s.groupId || '') +
        ' roomId=' + (s.roomId || ''));
    });
  } catch (err) {
    console.error('parse webhook failed: ' + err.message);
  }
}
