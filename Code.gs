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
