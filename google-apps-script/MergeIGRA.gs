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

var IGRA_SOURCE = '';            // ★ วางลิงก์หรือ ID ชีต IGRA ที่นี่
var IGRA_SHEET_NAME = '';        // ชื่อแท็บในไฟล์ IGRA (เว้นว่าง = ใช้แท็บแรก)

function _igraNorm(s) { return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase().trim(); }
function _igraNameKey(s) { return String(s == null ? '' : s).replace(/[\s​]+/g, '').replace(/^(นาย|นาง|นางสาว|น\.ส\.|ด\.ช\.|ด\.ญ\.)/, '').trim(); }
function _igraHnKey(s) { return String(s == null ? '' : s).replace(/\D/g, '').replace(/^0+/, ''); } // ตัดอักขระไม่ใช่ตัวเลข + เลขศูนย์นำหน้า

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
  var ui = SpreadsheetApp.getUi();
  try {
    var id = _igraSheetId(IGRA_SOURCE);
    if (!id) { ui.alert('ยังไม่ได้ตั้งค่าลิงก์ชีต IGRA\n\nเปิดไฟล์ MergeIGRA.gs แล้ววางลิงก์/ID ที่ตัวแปร IGRA_SOURCE'); return; }

    var ext;
    try { ext = SpreadsheetApp.openById(id); }
    catch (e) { ui.alert('เปิดไฟล์ IGRA ไม่ได้ (ID: ' + id + ')\nตรวจสอบว่าบัญชีนี้มีสิทธิ์เข้าถึงไฟล์\n\n' + (e && e.message ? e.message : e)); return; }

    var sh = IGRA_SHEET_NAME ? ext.getSheetByName(IGRA_SHEET_NAME) : ext.getSheets()[0];
    if (!sh) { ui.alert('ไม่พบแท็บ "' + IGRA_SHEET_NAME + '" ในไฟล์ IGRA'); return; }
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last < 2) { ui.alert('ไฟล์ IGRA ไม่มีข้อมูล'); return; }

    var vals = sh.getRange(1, 1, last, lastCol).getValues();
    var header = vals[0];
    var cHn = _igraFindCol(header, ['HN', 'เลข HN', 'หมายเลข HN', 'รหัสผู้ป่วย']);
    var cName = _igraFindCol(header, ['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล', 'ชื่อ นามสกุล', 'ชื่อ', 'name']);
    var cResult = _igraFindCol(header, ['ผล IGRA', 'ผลIGRA', 'IGRA', 'ผลตรวจ IGRA', 'ผลตรวจ', 'ผล', 'result', 'interpretation']);
    var cDate = _igraFindCol(header, ['วันที่ตรวจ', 'วันที่', 'วันที่เจาะเลือด', 'date', 'collected']);
    if (cResult < 0) { ui.alert('ไม่พบคอลัมน์ผล IGRA ในไฟล์ (หัวคอลัมน์ควรมีคำว่า IGRA หรือ "ผล")'); return; }
    if (cHn < 0 && cName < 0) { ui.alert('ไม่พบคอลัมน์ HN หรือ ชื่อ-สกุล ในไฟล์ IGRA'); return; }

    // สร้างดัชนีจากไฟล์ IGRA: hn -> {result,date} และ name -> {result,date}
    var byHn = {}, byName = {};
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      var result = String(cResult >= 0 ? row[cResult] : '').trim();
      if (!result) continue;
      var dateVal = cDate >= 0 ? row[cDate] : '';
      var dateStr = '';
      if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      else dateStr = String(dateVal || '').trim();
      var rec = { result: result, date: dateStr };
      if (cHn >= 0) { var hk = _igraHnKey(row[cHn]); if (hk) byHn[hk] = rec; }
      if (cName >= 0) { var nk = _igraNameKey(row[cName]); if (nk) byName[nk] = rec; }
    }

    // อ่านแท็บข้อมูลแบบสอบถามของระบบ แล้วเมิจ
    var dsh = ss().getSheetByName(SHEET_DATA);
    if (!dsh) { ui.alert('ไม่พบแท็บ "' + SHEET_DATA + '"'); return; }
    var dlast = dsh.getLastRow(), dcol = dsh.getLastColumn();
    if (dlast < 2) { ui.alert('ยังไม่มีข้อมูลผู้สัมผัสในระบบ'); return; }
    var block = dsh.getRange(1, 1, dlast, dcol).getValues();
    var dhead = block[0];
    var jsonIdx = -1, igraIdx = -1;
    for (var i = 0; i < dhead.length; i++) { var hh = _igraNorm(dhead[i]); if (hh === '_json') jsonIdx = i; if (hh === 'igra') igraIdx = i; }
    if (jsonIdx < 0) { ui.alert('ไม่พบคอลัมน์ _JSON (ให้รัน ensureHeader/บันทึกข้อมูลสักครั้งก่อน)'); return; }

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
    ui.alert('ซิงก์ IGRA เข้าระบบแล้ว ' + matched + ' ราย\n(จับคู่ด้วย HN ' + matchedHn + ' · ด้วยชื่อ ' + matchedName + ')\n\nรีเฟรชหน้าเว็บเพื่อดูผลในคอลัมน์ IGRA' +
      (igraIdx < 0 ? '\n\n* ยังไม่พบคอลัมน์ "IGRA" ในแท็บข้อมูล — สั่ง "ใส่ดรอปดาวน์ผลตรวจ" หรือบันทึกข้อมูลสักครั้งเพื่อสร้างคอลัมน์' : '')) ;
  } catch (err) {
    ui.alert('syncIGRA error: ' + (err && err.message ? err.message : err));
  }
}
