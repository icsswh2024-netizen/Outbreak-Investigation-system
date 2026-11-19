/**
 * MergeIGRA.gs — ดึงผลตรวจ IGRA จาก "ชีตอีกไฟล์หนึ่ง" มาเมิจเข้าระบบ
 *
 * วิธีใช้
 *  1) วาง "ลิงก์" หรือ "ID" ของ Google Sheet ไฟล์ IGRA ลงในค่า IGRA_SOURCE ด้านล่าง
 *     - ลิงก์เต็ม เช่น https://docs.google.com/spreadsheets/d/XXXXXXXX/edit#gid=0
 *     - หรือจะใส่เฉพาะ ID (ส่วน XXXXXXXX) ก็ได้
 *     * บัญชีที่รัน Apps Script ต้องมีสิทธิ์เปิดไฟล์ IGRA นั้น (อย่างน้อยสิทธิ์ดู)
 *  2) (ถ้าชีตมีหลายแท็บ) ระบุชื่อแท็บที่ IGRA_SHEET_NAME ไม่งั้นใช้แท็บแรก
 *  3) เมนู "🧪 ผลตรวจ → ซิงก์ IGRA จากชีตภายนอก"  (เมนูรวมอยู่ใน MatchRoster.gs)
 *
 * การจับคู่: ใช้ HN ก่อน ถ้าไม่เจอค่อยใช้ชื่อ-สกุล
 * ผลลัพธ์: เขียนลง rec.labs.igra = [{date, result}] ใน _JSON + คอลัมน์ "IGRA" ในแท็บข้อมูลแบบสอบถาม
 */

var IGRA_SOURCE = 'https://docs.google.com/spreadsheets/d/17Txw-ViGta5i-l_rRIaRcULeAJMNoUkgL6P480Ie5bQ/edit';  // ★ ลิงก์/ID ชีต IGRA
var IGRA_SHEET_NAME = 'กลุ่มเสี่ยงสัมผัสทั้งหมด';   // ชื่อแท็บ (เว้นว่าง = ใช้แท็บแรก)
var IGRA_DATA_START_ROW = 5;      // แถวแรกที่เป็นข้อมูล (หัวตารางอยู่แถว 1–4)
// ระบุคอลัมน์ตรงๆ (A,B,C,...) เพราะหัวตารางมีหลายแถวและชื่อ "ผล" ซ้ำกัน — เว้นว่างเพื่อให้เดาจากหัวตารางแทน
var IGRA_COL_HN = 'B';            // HN
var IGRA_COL_NAME = 'C';          // ชื่อ-สกุล
var IGRA_COL_RESULT = 'P';        // ผล IGRA
var IGRA_COL_DATE = 'O';          // วันที่ส่ง IGRA

function _igraNorm(s) { return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase().trim(); }
function _igraNameKey(s) { return String(s == null ? '' : s).replace(/[\s​]+/g, '').replace(/^(นาย|นาง|นางสาว|น\.ส\.|ด\.ช\.|ด\.ญ\.)/, '').trim(); }
function _igraHnKey(s) { return String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, ''); } // ตัดอักขระไม่ใช่ตัวเลข + เลขศูนย์นำหน้า
function _igraColIdx(letter) { // 'A'->0, 'B'->1, ... ; รับตัวเลขก็ได้
  var s = String(letter || '').trim().toUpperCase();
  if (/^\d+$/.test(s)) return parseInt(s, 10) - 1; // ถ้าใส่เป็นเลขคอลัมน์ 1-based
  var n = 0; for (var i = 0; i < s.length; i++) { var c = s.charCodeAt(i) - 64; if (c < 1 || c > 26) return -1; n = n * 26 + c; } return n - 1;
}
// แปลงวันที่ IGRA เป็น ISO (รองรับ Date, "M/D/YYYY", "YYYY-MM-DD", "(2026-09-08):")
function _igraDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);           // YYYY-MM-DD (อาจมี (): ครอบ)
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);            // M/D/YYYY
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}

// ดึง ID จากลิงก์หรือรับ ID ตรงๆ
function _igraSheetId(src) {
  src = String(src || '').trim();
  var m = src.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(src)) return src;
  return '';
}

// หาคอลัมน์ในหัวตารางแบบยืดหยุ่น (ตรงตัวก่อน แล้ว "มีคำว่า")
function _igraFindCol(header, names) {
  for (var i = 0; i < header.length; i++) { var h = _igraNorm(header[i]); for (var j = 0; j < names.length; j++) if (h === _igraNorm(names[j])) return i; }
  for (var i2 = 0; i2 < header.length; i2++) { var h2 = _igraNorm(header[i2]); for (var j2 = 0; j2 < names.length; j2++) { var n = _igraNorm(names[j2]); if (n && h2.indexOf(n) >= 0) return i2; } }
  return -1;
}

function syncIGRA() {
  var _ui = null; try { _ui = SpreadsheetApp.getUi(); } catch (e0) { _ui = null; }
  var alertFn = function (m) { if (_ui) { try { _alertFn(m); } catch (e1) {} } else { Logger.log(m); } };
  try {
    var id = _igraSheetId(IGRA_SOURCE);
    if (!id) { alertFn('ยังไม่ได้ตั้งค่าลิงก์ชีต IGRA\n\nเปิดไฟล์ MergeIGRA.gs แล้ววางลิงก์/ID ที่ตัวแปร IGRA_SOURCE'); return; }

    var ext;
    try { ext = SpreadsheetApp.openById(id); }
    catch (e) { alertFn('เปิดไฟล์ IGRA ไม่ได้ (ID: ' + id + ')\nตรวจสอบว่าบัญชีนี้มีสิทธิ์เข้าถึงไฟล์\n\n' + (e && e.message ? e.message : e)); return; }

    var sh = IGRA_SHEET_NAME ? ext.getSheetByName(IGRA_SHEET_NAME) : ext.getSheets()[0];
    if (!sh) { alertFn('ไม่พบแท็บ "' + IGRA_SHEET_NAME + '" ในไฟล์ IGRA'); return; }
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last < IGRA_DATA_START_ROW) { alertFn('ไฟล์ IGRA ไม่มีข้อมูล'); return; }

    var vals = sh.getRange(1, 1, last, lastCol).getValues();

    // หาตำแหน่งคอลัมน์: ใช้ที่ระบุตรงๆ ก่อน ไม่งั้นเดาจากหัวตาราง (แถวที่ 4)
    var cHn = IGRA_COL_HN ? _igraColIdx(IGRA_COL_HN) : -1;
    var cName = IGRA_COL_NAME ? _igraColIdx(IGRA_COL_NAME) : -1;
    var cResult = IGRA_COL_RESULT ? _igraColIdx(IGRA_COL_RESULT) : -1;
    var cDate = IGRA_COL_DATE ? _igraColIdx(IGRA_COL_DATE) : -1;
    if (cResult < 0 || cHn < 0 && cName < 0) {
      var header = vals[Math.max(0, IGRA_DATA_START_ROW - 2)] || vals[0];
      if (cHn < 0) cHn = _igraFindCol(header, ['HN', 'เลข HN', 'หมายเลข HN']);
      if (cName < 0) cName = _igraFindCol(header, ['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล', 'ชื่อ', 'name']);
      if (cResult < 0) cResult = _igraFindCol(header, ['ผล IGRA', 'IGRA', 'ผลตรวจ', 'ผล', 'result']);
      if (cDate < 0) cDate = _igraFindCol(header, ['วันที่ส่ง', 'วันที่ตรวจ', 'วันที่', 'date']);
    }
    if (cResult < 0) { alertFn('ไม่พบคอลัมน์ผล IGRA'); return; }
    if (cHn < 0 && cName < 0) { alertFn('ไม่พบคอลัมน์ HN หรือ ชื่อ-สกุล ในไฟล์ IGRA'); return; }

    // สร้างดัชนีจากไฟล์ IGRA: hn -> {result,date} และ name -> {result,date}
    var byHn = {}, byName = {};
    for (var r = IGRA_DATA_START_ROW - 1; r < vals.length; r++) {
      var row = vals[r];
      var result = String(cResult >= 0 ? row[cResult] : '').trim();
      if (!result) continue;
      var rec = { result: result, date: cDate >= 0 ? _igraDate(row[cDate]) : '' };
      if (cHn >= 0) { var hk = _igraHnKey(row[cHn]); if (hk) byHn[hk] = rec; }
      if (cName >= 0) { var nk = _igraNameKey(row[cName]); if (nk) byName[nk] = rec; }
    }

    // อ่านแท็บข้อมูลแบบสอบถามของระบบ แล้วเมิจ
    var dsh = ss().getSheetByName(SHEET_DATA);
    if (!dsh) { alertFn('ไม่พบแท็บ "' + SHEET_DATA + '"'); return; }
    var dlast = dsh.getLastRow(), dcol = dsh.getLastColumn();
    if (dlast < 2) { alertFn('ยังไม่มีข้อมูลผู้สัมผัสในระบบ'); return; }
    var block = dsh.getRange(1, 1, dlast, dcol).getValues();
    var dhead = block[0];
    var jsonIdx = -1, igraIdx = -1;
    for (var i = 0; i < dhead.length; i++) { var hh = _igraNorm(dhead[i]); if (hh === '_json') jsonIdx = i; if (hh === 'igra') igraIdx = i; }
    if (jsonIdx < 0) { alertFn('ไม่พบคอลัมน์ _JSON (ให้บันทึกข้อมูลสักครั้งก่อน)'); return; }

    var matched = 0, matchedHn = 0, matchedName = 0;
    for (var d = 1; d < block.length; d++) {
      var js = block[d][jsonIdx];
      if (!js) continue;
      var obj; try { obj = JSON.parse(js); } catch (e2) { continue; }
      var hit = null, via = '';
      var hk2 = _igraHnKey(obj.hn);
      if (hk2 && byHn[hk2]) { hit = byHn[hk2]; via = 'hn'; }
      if (!hit) { var nk2 = _igraNameKey(obj.name); if (nk2 && byName[nk2]) { hit = byName[nk2]; via = 'name'; } }
      if (!hit) continue;
      obj.labs = obj.labs || {};
      obj.labs.igra = [{ date: hit.date || '', result: hit.result }];
      block[d][jsonIdx] = JSON.stringify(obj);
      if (igraIdx >= 0) block[d][igraIdx] = hit.result;
      matched++; if (via === 'hn') matchedHn++; else matchedName++;
    }

    dsh.getRange(1, 1, block.length, dcol).setValues(block);
    alertFn('ซิงก์ IGRA เข้าระบบแล้ว ' + matched + ' ราย\n(จับคู่ด้วย HN ' + matchedHn + ' · ด้วยชื่อ ' + matchedName + ')\n\nรีเฟรชหน้าเว็บเพื่อดูผลในคอลัมน์ IGRA' +
      (igraIdx < 0 ? '\n\n* ยังไม่พบคอลัมน์ "IGRA" ในแท็บข้อมูล — สั่ง "ใส่ดรอปดาวน์ผลตรวจ" หรือบันทึกข้อมูลสักครั้งเพื่อสร้างคอลัมน์' : '')) ;
  } catch (err) {
    alertFn('syncIGRA error: ' + (err && err.message ? err.message : err));
  }
}

/* ---------- ซิงก์ IGRA อัตโนมัติ (time-driven trigger) ---------- */
// ฟังก์ชันที่ทริกเกอร์เรียก (ไม่มี UI) — ใช้ตัว syncIGRA เดิมซึ่งจะ Logger.log แทน alert เมื่อไม่มี UI
function syncIGRAauto() { syncIGRA(); }

// ติดตั้งให้ซิงก์ IGRA อัตโนมัติทุกวัน (เวลา ~01:00)
function installIGRAAutoSync() {
  var ui = null; try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  removeIGRAAutoSyncTriggers_();
  ScriptApp.newTrigger('syncIGRAauto').timeBased().everyDays(1).atHour(1).create();
  if (ui) ui.alert('ตั้งซิงก์ IGRA อัตโนมัติแล้ว — ระบบจะดึงผล IGRA จากชีตภายนอกให้ทุกวัน (ประมาณ 01:00 น.)\n\nยกเลิกได้ที่เมนู "🧪 ผลตรวจ → ปิดซิงก์ IGRA อัตโนมัติ"');
}

// ปิดซิงก์อัตโนมัติ
function removeIGRAAutoSync() {
  var ui = null; try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var n = removeIGRAAutoSyncTriggers_();
  if (ui) ui.alert(n > 0 ? 'ปิดซิงก์ IGRA อัตโนมัติแล้ว' : 'ไม่มีการตั้งซิงก์อัตโนมัติอยู่');
}

function removeIGRAAutoSyncTriggers_() {
  var trigs = ScriptApp.getProjectTriggers(), n = 0;
  for (var i = 0; i < trigs.length; i++) {
    if (trigs[i].getHandlerFunction() === 'syncIGRAauto') { ScriptApp.deleteTrigger(trigs[i]); n++; }
  }
  return n;
}


