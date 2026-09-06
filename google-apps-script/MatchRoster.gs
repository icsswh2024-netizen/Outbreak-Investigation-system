/**
 * MatchRoster.gs — แมตรายชื่อจริงจากแท็บ "รายชื่อเจ้าหน้าที่ปฏิบัติงาน"
 * เข้ากับแท็บ "ข้อมูลแบบสอบถาม" โดยจับคู่ด้วย "ชื่อ-นามสกุล" (normalize)
 * เมื่อเจอ → เขียนทับ คำนำหน้า / ชื่อ-สกุล / ตำแหน่ง / หน่วยงาน ด้วยทะเบียนจริง
 * แถวที่แมตไม่เจอ → ระบายสีเหลืองไว้ให้ตรวจเอง
 *
 * วิธีใช้:
 *  1) วางไฟล์นี้เพิ่มในโปรเจกต์ Apps Script เดียวกับ Code.gs (บันทึก 💾)
 *  2) รีเฟรชชีต จะมีเมนู "🩺 แมตรายชื่อ" โผล่ด้านบน
 *  3) กด "แมตรายชื่อจริง → ข้อมูลแบบสอบถาม"
 */

// ===== ตั้งค่าชื่อแท็บ =====
var ROSTER_SHEET = 'รายชื่อเจ้าหน้าที่ปฏิบัติงาน'; // ทะเบียนจริง (source of truth)
var DATA_SHEET   = 'ข้อมูลแบบสอบถาม';              // ปลายทางที่ถูกเขียนทับ

// ===== ชื่อหัวคอลัมน์ที่ยอมรับ (เผื่อสะกดต่างกัน) =====
// แท็บ "ข้อมูลแบบสอบถาม"
var COL_NAME_DATA = ['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล'];
var COL_POS_DATA  = ['กลุ่ม/ตำแหน่ง', 'ตำแหน่ง', 'ตำแหน่งการพยาบาล'];
var COL_UNIT_DATA = ['หน่วยงาน'];   // จับแบบ "ขึ้นต้นด้วย/มีคำว่า" หน่วยงาน
var COL_PREFIX    = 'คำนำหน้า';     // ถ้าไม่มีจะเพิ่มให้อัตโนมัติ
var COL_MATCH     = 'สถานะแมตรายชื่อ'; // คอลัมน์ audit ถ้าไม่มีจะเพิ่มให้

// แท็บ "รายชื่อเจ้าหน้าที่ปฏิบัติงาน"
var COL_NAME_ROSTER   = ['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล'];
var COL_FIRST_ROSTER  = ['ชื่อ'];
var COL_LAST_ROSTER   = ['นามสกุล', 'สกุล'];
var COL_PREFIX_ROSTER = ['คำนำหน้า'];
var COL_POS_ROSTER    = ['ตำแหน่ง', 'ตำแหน่งการพยาบาล'];
var COL_UNIT_ROSTER   = ['หน่วยงาน'];

var PREFIXES = ['นางสาว', 'น.ส.', 'นส.', 'นาย', 'นาง', 'ดร.', 'ดร',
                'นพ.', 'พญ.', 'ทพ.', 'ทพญ.', 'ภก.', 'ภญ.', 'ว่าที่ร้อยตรี', 'ว่าที่ ร.ต.'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🩺 แมตรายชื่อ')
    .addItem('แมตรายชื่อจริง → ข้อมูลแบบสอบถาม', 'matchRoster')
    .addToUi();
}

// ---------- helpers ----------
function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
         SpreadsheetApp.openById(typeof SPREADSHEET_ID !== 'undefined' ? SPREADSHEET_ID : '');
}
function _norm(s) {
  s = String(s == null ? '' : s);
  s = s.replace(/[​‌‍﻿]/g, ''); // zero-width
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}
// กุญแจเทียบ: ตัดคำนำหน้า + ตัดช่องว่างทั้งหมด
function _nameKey(s) {
  var t = _norm(s);
  for (var i = 0; i < PREFIXES.length; i++) {
    if (t.indexOf(PREFIXES[i]) === 0) { t = t.slice(PREFIXES[i].length); break; }
  }
  return t.replace(/\s+/g, '');
}
function _stripPrefix(s) {
  var t = _norm(s);
  for (var i = 0; i < PREFIXES.length; i++) {
    if (t.indexOf(PREFIXES[i]) === 0) return _norm(t.slice(PREFIXES[i].length));
  }
  return t;
}
function _findCol(headers, names) {
  for (var i = 0; i < headers.length; i++) {
    var h = _norm(headers[i]);
    for (var j = 0; j < names.length; j++) if (h === _norm(names[j])) return i;
  }
  return -1;
}
function _findColContains(headers, kw) {
  for (var i = 0; i < headers.length; i++) if (_norm(headers[i]).indexOf(kw) >= 0) return i;
  return -1;
}

// ---------- main ----------
function matchRoster() {
  var ss = _ss();
  var ui = SpreadsheetApp.getUi();
  var rs = ss.getSheetByName(ROSTER_SHEET);
  var ds = ss.getSheetByName(DATA_SHEET);
  if (!rs) { ui.alert('ไม่พบแท็บ "' + ROSTER_SHEET + '"'); return; }
  if (!ds) { ui.alert('ไม่พบแท็บ "' + DATA_SHEET + '"'); return; }

  // ---- อ่านทะเบียนจริง ----
  var rv = rs.getDataRange().getValues();
  if (rv.length < 2) { ui.alert('แท็บทะเบียนไม่มีข้อมูล'); return; }
  var rh = rv[0];
  var rNameCol  = _findCol(rh, COL_NAME_ROSTER);
  var rFirstCol = _findCol(rh, COL_FIRST_ROSTER);
  var rLastCol  = _findCol(rh, COL_LAST_ROSTER);
  var rPreCol   = _findCol(rh, COL_PREFIX_ROSTER);
  var rPosCol   = _findCol(rh, COL_POS_ROSTER);
  var rUnitCol  = _findCol(rh, COL_UNIT_ROSTER);
  if (rNameCol < 0 && (rFirstCol < 0 || rLastCol < 0)) {
    ui.alert('แท็บทะเบียนต้องมีคอลัมน์ "ชื่อ-สกุล" หรือ ("ชื่อ" และ "นามสกุล")'); return;
  }

  var reg = {}; // key -> {prefix, full, pos, unit}
  for (var i = 1; i < rv.length; i++) {
    var row = rv[i];
    var prefix = rPreCol >= 0 ? _norm(row[rPreCol]) : '';
    var full, key;
    if (rNameCol >= 0) {
      full = _stripPrefix(row[rNameCol]);
      if (!prefix) { // ดึงคำนำหน้าออกจากชื่อรวมถ้าไม่มีคอลัมน์แยก
        var raw = _norm(row[rNameCol]);
        for (var p = 0; p < PREFIXES.length; p++) if (raw.indexOf(PREFIXES[p]) === 0) { prefix = PREFIXES[p]; break; }
      }
      key = _nameKey(row[rNameCol]);
    } else {
      var f = _norm(row[rFirstCol]), l = _norm(row[rLastCol]);
      full = _norm(f + ' ' + l);
      key = (f + l).replace(/\s+/g, '');
    }
    if (!key) continue;
    reg[key] = {
      prefix: prefix,
      full: full,
      pos: rPosCol >= 0 ? _norm(row[rPosCol]) : '',
      unit: rUnitCol >= 0 ? _norm(row[rUnitCol]) : ''
    };
  }

  // ---- อ่านแท็บข้อมูลแบบสอบถาม ----
  var dv = ds.getDataRange().getValues();
  if (dv.length < 2) { ui.alert('แท็บข้อมูลแบบสอบถามไม่มีข้อมูล'); return; }
  var dh = dv[0];
  var dNameCol = _findCol(dh, COL_NAME_DATA);
  var dPosCol  = _findCol(dh, COL_POS_DATA);
  var dUnitCol = _findColContains(dh, 'หน่วยงาน');
  if (dNameCol < 0) { ui.alert('แท็บข้อมูลแบบสอบถามไม่มีคอลัมน์ "ชื่อ-สกุล"'); return; }

  // เพิ่มคอลัมน์ คำนำหน้า / สถานะแมต ถ้ายังไม่มี (ต่อท้าย)
  var dPreCol   = _findCol(dh, [COL_PREFIX]);
  var dMatchCol = _findCol(dh, [COL_MATCH]);
  var appended = [];
  if (dPreCol < 0)   { dPreCol = dh.length + appended.length; appended.push(COL_PREFIX); }
  if (dMatchCol < 0) { dMatchCol = dh.length + appended.length; appended.push(COL_MATCH); }
  if (appended.length) {
    ds.getRange(1, dh.length + 1, 1, appended.length).setValues([appended]);
  }

  var matched = 0, unmatched = 0;
  var lastRow = dv.length;               // จำนวนแถวข้อมูล (รวม header)
  var maxCol = Math.max(dNameCol, dPosCol, dUnitCol, dPreCol, dMatchCol) + 1;

  // เตรียม array สำหรับเขียนกลับทีเดียว (เร็ว + กันเขียนชนกัน)
  // อ่านช่วงตั้งแต่คอลัมน์ 1 ถึง maxCol ทุกแถวข้อมูล
  var block = ds.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var bgClear = [];
  for (var r = 0; r < block.length; r++) {
    var brow = block[r];
    var key = _nameKey(brow[dNameCol]);
    if (key && reg[key]) {
      var m = reg[key];
      // เขียนทับด้วยทะเบียนจริง
      brow[dNameCol] = m.full;                 // ชื่อ-สกุล (ไม่มีคำนำหน้า)
      if (dPosCol >= 0 && m.pos)  brow[dPosCol]  = m.pos;
      if (dUnitCol >= 0 && m.unit) brow[dUnitCol] = m.unit;
      brow[dPreCol]   = m.prefix;
      brow[dMatchCol] = 'ตรงกับทะเบียน';
      matched++;
      bgClear.push(null);
    } else {
      brow[dMatchCol] = brow[dNameCol] ? 'ไม่พบในทะเบียน' : '';
      unmatched += brow[dNameCol] ? 1 : 0;
      bgClear.push(brow[dNameCol] ? '#fff3cd' : null);
    }
    // เผื่อ block สั้นกว่า maxCol ให้เติม
    while (brow.length < maxCol) brow.push('');
  }
  ds.getRange(2, 1, block.length, maxCol).setValues(block);

  // ระบายสีชื่อแถวที่ไม่พบ (คอลัมน์ชื่อ)
  for (var r2 = 0; r2 < bgClear.length; r2++) {
    ds.getRange(r2 + 2, dNameCol + 1).setBackground(bgClear[r2]);
  }

  ui.alert('เสร็จแล้ว ✅\n\nแมตตรงกับทะเบียน: ' + matched + ' ราย\nไม่พบในทะเบียน (ระบายเหลือง): ' + unmatched + ' ราย');
}
