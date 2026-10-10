/**
 * ตารางดูแลตัวเองของแม่ — ตัวรับข้อมูลลง Google Sheet
 * วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone)
 */
const SHEET = 'บันทึก';
const KEYS = ['water','cough_m','cough_n','cough_e','meal_m','meal_n','meal_e','op1','op2','op3','exercise'];
const HEAD = ['วันที่','ดื่มน้ำ 2 ลิตร','ฝึกไอ เช้า','ฝึกไอ กลางวัน','ฝึกไอ เย็น','อาหารเช้า','อาหารกลางวัน','อาหารเย็น',
              'Once Pro 1','Once Pro 2','Once Pro 3','ออกกำลังกาย','น้ำหนัก (กก.)','ทำแล้ว (จาก 12)','อัปเดตล่าสุด'];
const NCOL = HEAD.length; // 15
const TZ = 'Asia/Bangkok';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET, 0);
    sh.getRange(1, 1, 1, NCOL).setValues([HEAD]).setFontWeight('bold').setBackground('#FDE3E6');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('M:M').setNumberFormat('0.0');
    sh.getRange('O:O').setNumberFormat('d/m/yyyy HH:mm');
    sh.setColumnWidth(1, 100);
  }
  return sh;
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function readAll_(sh) {
  const vals = sh.getDataRange().getValues();
  const days = {}, rowOf = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r], k = dateKey_(row[0]);
    if (!k) continue;
    const items = {};
    KEYS.forEach((id, i) => { if (row[i + 1] === true) items[id] = true; });
    const w = row[12] === '' || row[12] == null ? null : Number(row[12]);
    const u = row[14] instanceof Date ? row[14].getTime() : Number(row[14]) || 0;
    days[k] = { items: items, weight: isFinite(w) ? w : null, u: u };
    rowOf[k] = r + 1;
  }
  return { days: days, rowOf: rowOf };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, days: readAll_(sheet_()).days });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sh = sheet_();
    const cur = readAll_(sh);
    const incoming = body.days || {};
    Object.keys(incoming).sort().forEach(function (k) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(k)) return;
      const d = incoming[k] || {}, items = d.items || {};
      const u = Number(d.u) || Date.now();
      if (cur.days[k] && cur.days[k].u >= u) return; // ข้อมูลในชีตใหม่กว่า
      const w = d.weight == null || d.weight === '' ? '' : Math.round(Number(d.weight) * 10) / 10;
      const ticks = KEYS.map(function (id) { return items[id] === true; });
      const done = ticks.filter(Boolean).length + (w === '' ? 0 : 1);
      const row = [k].concat(ticks, [w, done, new Date(u)]);
      let r = cur.rowOf[k];
      if (!r) {
        r = sh.getLastRow() + 1;
        sh.getRange(r, 2, 1, KEYS.length).insertCheckboxes();
        cur.rowOf[k] = r;
      }
      sh.getRange(r, 1, 1, NCOL).setValues([row]);
      cur.days[k] = { items: items, weight: w === '' ? null : w, u: u };
    });
    if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, NCOL).sort({ column: 1, ascending: false });
    return out_({ ok: true, days: readAll_(sh).days });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 * แจ้งเตือนทาง LINE ถ้าแม่ยังไม่ได้ติ๊ก
 * ตั้งค่า: Project Settings > Script Properties > เพิ่ม LINE_TOKEN
 * แล้วกดรัน testLine 1 ครั้ง ตามด้วย setupReminders 1 ครั้ง
 * ========================================================= */
const APP_URL = 'https://nkowenn-cloud.github.io/Mom-care/';
const LABEL = {
  water: 'ดื่มน้ำ 2 ลิตร', cough_m: 'ฝึกไอ ตอนเช้า', cough_n: 'ฝึกไอ ตอนกลางวัน', cough_e: 'ฝึกไอ ตอนเย็น',
  meal_m: 'อาหารเช้า', meal_n: 'อาหารกลางวัน', meal_e: 'อาหารเย็น',
  op1: 'Once Pro แก้วที่ 1', op2: 'Once Pro แก้วที่ 2', op3: 'Once Pro แก้วที่ 3',
  exercise: 'ออกกำลังกาย', weight: 'ชั่งน้ำหนัก'
};
// เวลาเตือน (แก้ได้) — แต่ละรอบจะเตือนรายการที่ถึงเวลาแล้วแต่ยังไม่ติ๊ก รวมของรอบก่อนหน้าด้วย
const REMIND = [
  { fn: 'remindMorning', hour: 10, due: ['cough_m', 'meal_m', 'op1'] },
  { fn: 'remindNoon',    hour: 14, due: ['cough_n', 'meal_n', 'op2'] },
  { fn: 'remindEvening', hour: 19, due: ['cough_e', 'meal_e', 'op3', 'water', 'exercise', 'weight'], praise: true }
];

function remindMorning() { remind_(0); }
function remindNoon()    { remind_(1); }
function remindEvening() { remind_(2); }

function remind_(i) {
  const due = REMIND.slice(0, i + 1).reduce(function (a, s) { return a.concat(s.due); }, []);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const d = readAll_(sheet_()).days[today] || { items: {}, weight: null };
  const missing = due.filter(function (id) { return id === 'weight' ? d.weight == null : !d.items[id]; });
  if (!missing.length) {
    if (REMIND[i].praise) sendLine_('วันนี้แม่ทำครบทุกอย่างแล้ว เก่งมากเลยค่ะ ❤️');
    return;
  }
  sendLine_('แม่จ๋า ยังไม่ได้ติ๊กในแอพนะคะ\n\n' +
    missing.map(function (id) { return '• ' + LABEL[id]; }).join('\n') +
    '\n\nถ้าทำแล้ว อย่าลืมกดติ๊กในแอพด้วยนะคะ ❤️\n' + APP_URL);
}

function sendLine_(text) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ใส่ LINE_TOKEN ใน Script Properties');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE ส่งไม่สำเร็จ: ' + res.getResponseCode() + ' ' + res.getContentText());
}

/** รันครั้งเดียวเพื่อทดสอบว่าข้อความเข้า LINE */
function testLine() { sendLine_('ทดสอบแจ้งเตือนจากแอพดูแลแม่ ✅'); }

/** รันครั้งเดียวเพื่อตั้งเวลาเตือนทุกวัน (รันซ้ำได้ จะลบของเก่าก่อน) */
function setupReminders() {
  const names = REMIND.map(function (s) { return s.fn; });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  REMIND.forEach(function (s) {
    ScriptApp.newTrigger(s.fn).timeBased().everyDays(1).atHour(s.hour).nearMinute(0).inTimezone(TZ).create();
  });
}
