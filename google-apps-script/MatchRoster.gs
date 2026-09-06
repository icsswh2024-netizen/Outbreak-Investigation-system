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
var COL_FNAME     = 'ชื่อ';          // เพิ่มให้อัตโนมัติ (แยกจากทะเบียน)
var COL_LNAME     = 'นามสกุล';       // เพิ่มให้อัตโนมัติ
var COL_MATCH     = 'สถานะแมตรายชื่อ'; // คอลัมน์ audit ถ้าไม่มีจะเพิ่มให้

// แท็บ "รายชื่อเจ้าหน้าที่ปฏิบัติงาน"
var COL_NAME_ROSTER   = ['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล'];
var COL_FIRST_ROSTER  = ['ชื่อ'];
var COL_LAST_ROSTER   = ['นามสกุล', 'สกุล'];
var COL_PREFIX_ROSTER = ['คำนำหน้า'];
var COL_POS_ROSTER    = ['ตำแหน่ง', 'ตำแหน่งการพยาบาล'];
var COL_UNIT_ROSTER   = ['หน่วยงาน', 'หน่วย', 'หน่วยงาน/หอผู้ป่วย'];

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
  try { s = s.normalize('NFC'); } catch (e) {}
  s = s.replace(/[​‌‍﻿]/g, ''); // zero-width
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}
// กุญแจเทียบ: ตัดคำนำหน้า + ตัดช่องว่างทั้งหมด
function _stripAllPrefix(t) {
  var changed = true;
  while (changed) {
    changed = false;
    for (var i = 0; i < PREFIXES.length; i++) {
      var pfx = PREFIXES[i];
      if (t.indexOf(pfx) === 0) { t = t.slice(pfx.length).replace(/^[\s.]+/, ''); changed = true; break; }
    }
  }
  return t;
}
function _nameKey(s) {
  return _stripAllPrefix(_norm(s)).replace(/[\s.()\-]/g, '');
}
function _stripPrefix(s) {
  return _stripAllPrefix(_norm(s));
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
// ระยะแก้ไข (Levenshtein) สำหรับแมตชื่อพิมพ์ผิดเล็กน้อย
function _lev(a, b) {
  var m = a.length, n = b.length;
  if (!m) return n; if (!n) return m;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= n; j++) prev[j] = j;
  for (i = 1; i <= m; i++) {
    cur[0] = i;
    for (j = 1; j <= n; j++) {
      var cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
    }
    for (j = 0; j <= n; j++) prev[j] = cur[j];
  }
  return prev[n];
}
// หาคู่ใกล้เคียงที่สุดในทะเบียน: คืน key ถ้ามั่นใจพอ ไม่งั้นคืน null
function _fuzzyKey(key, regKeys) {
  if (!key) return null;
  // 1) ชื่อในแบบสอบถามเป็น "คำขึ้นต้น" ของทะเบียนแบบไม่ซ้ำ (เช่น ขาดนามสกุล)
  var pfxHits = [];
  for (var p = 0; p < regKeys.length; p++) {
    var rk = regKeys[p];
    if (key.length >= 4 && rk.indexOf(key) === 0) pfxHits.push(rk);
  }
  if (pfxHits.length === 1) return pfxHits[0];
  // 2) ระยะแก้ไขน้อยสุดและไม่กำกวม
  var best = null, bestD = 99, second = 99;
  for (var i = 0; i < regKeys.length; i++) {
    var d = _lev(key, regKeys[i]);
    if (d < bestD) { second = bestD; bestD = d; best = regKeys[i]; }
    else if (d < second) { second = d; }
  }
  var limit = key.length >= 6 ? 2 : 1;  // ชื่อสั้นเข้มงวดกว่า
  if (best && bestD <= limit && second - bestD >= 1) return best;
  return null;
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

  var reg = {}; // key -> {prefix, first, last, full, pos, unit}
  for (var i = 1; i < rv.length; i++) {
    var row = rv[i];
    var prefix = rPreCol >= 0 ? _norm(row[rPreCol]) : '';
    var full, key, first, last;
    if (rNameCol >= 0) {
      full = _stripPrefix(row[rNameCol]);          // ชื่อ+สกุล ไม่มีคำนำหน้า
      if (!prefix) { // ดึงคำนำหน้าออกจากชื่อรวมถ้าไม่มีคอลัมน์แยก
        var raw = _norm(row[rNameCol]);
        for (var p = 0; p < PREFIXES.length; p++) if (raw.indexOf(PREFIXES[p]) === 0) { prefix = PREFIXES[p]; break; }
      }
      var sp = full.indexOf(' ');                  // แยกที่เว้นวรรคแรก
      first = sp >= 0 ? full.slice(0, sp) : full;
      last  = sp >= 0 ? _norm(full.slice(sp + 1)) : '';
      key = _nameKey(row[rNameCol]);
    } else {
      first = _stripAllPrefix(_norm(row[rFirstCol]));
      last  = _norm(row[rLastCol]);
      full = _norm(first + ' ' + last);
      key = _nameKey(first + ' ' + last);
    }
    if (!key) continue;
    reg[key] = {
      prefix: prefix, first: first, last: last, full: full,
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
  var dFCol     = _findCol(dh, [COL_FNAME]);
  var dLCol     = _findCol(dh, [COL_LNAME]);
  var dJsonCol  = _findCol(dh, ['_JSON']);  // ก้อน JSON ที่เว็บแอปอ่านกลับ (ต้องแก้ด้วย)
  var dMatchCol = _findCol(dh, [COL_MATCH]);
  var appended = [];
  if (dPreCol < 0)   { dPreCol = dh.length + appended.length; appended.push(COL_PREFIX); }
  if (dFCol < 0)     { dFCol = dh.length + appended.length; appended.push(COL_FNAME); }
  if (dLCol < 0)     { dLCol = dh.length + appended.length; appended.push(COL_LNAME); }
  if (dMatchCol < 0) { dMatchCol = dh.length + appended.length; appended.push(COL_MATCH); }
  if (appended.length) {
    ds.getRange(1, dh.length + 1, 1, appended.length).setValues([appended]);
  }

  var regKeys = Object.keys(reg);
  var matched = 0, fuzzy = 0, unmatched = 0;
  var lastRow = dv.length;               // จำนวนแถวข้อมูล (รวม header)
  var maxCol = Math.max(dNameCol, dPosCol, dUnitCol, dPreCol, dFCol, dLCol, dJsonCol, dMatchCol) + 1;

  var block = ds.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var bgClear = [];
  var missList = [];
  for (var r = 0; r < block.length; r++) {
    var brow = block[r];
    while (brow.length < maxCol) brow.push('');
    var key = _nameKey(brow[dNameCol]);
    var m = null, isFuzzy = false;
    if (key && reg[key]) { m = reg[key]; }
    else if (key) { var fk = _fuzzyKey(key, regKeys); if (fk) { m = reg[fk]; isFuzzy = true; } }

    if (m) {
      // เขียนทับด้วยทะเบียนจริง
      brow[dNameCol] = _norm((m.prefix ? m.prefix + ' ' : '') + m.full);
      if (dPosCol >= 0 && m.pos)  brow[dPosCol]  = m.pos;
      if (dUnitCol >= 0 && m.unit) brow[dUnitCol] = m.unit;
      brow[dPreCol]   = m.prefix;
      brow[dFCol]     = m.first;
      brow[dLCol]     = m.last;
      // อัปเดตก้อน _JSON ที่เว็บแอป/รายงานอ่านกลับ ให้ชื่อ/ตำแหน่ง/หน่วยงานเปลี่ยนตามทะเบียน
      if (dJsonCol >= 0 && brow[dJsonCol]) {
        try {
          var rec = JSON.parse(brow[dJsonCol]);
          var fullName = _norm((m.prefix ? m.prefix + ' ' : '') + m.full);
          rec.name = fullName;
          rec.group = m.pos || rec.group;
          rec.answers = rec.answers || {};
          rec.answers.name = fullName;
          if (m.pos)  rec.answers.role = m.pos;
          if (m.unit) rec.answers.unit = m.unit;
          brow[dJsonCol] = JSON.stringify(rec);
        } catch (e) {}
      }
      if (isFuzzy) { brow[dMatchCol] = 'แมตใกล้เคียง-ตรวจสอบ'; fuzzy++; bgClear.push('#ffe8cc'); }
      else { brow[dMatchCol] = 'ตรงกับทะเบียน'; matched++; bgClear.push(null); }
    } else {
      brow[dMatchCol] = brow[dNameCol] ? 'ไม่พบในทะเบียน' : '';
      if (brow[dNameCol]) { unmatched++; if (missList.length < 20) missList.push('• ' + _norm(brow[dNameCol])); }
      bgClear.push(brow[dNameCol] ? '#fff3cd' : null);
    }
  }
  ds.getRange(2, 1, block.length, maxCol).setValues(block);

  // ระบายสีชื่อแถวที่ไม่พบ (คอลัมน์ชื่อ)
  for (var r2 = 0; r2 < bgClear.length; r2++) {
    ds.getRange(r2 + 2, dNameCol + 1).setBackground(bgClear[r2]);
  }

  var msg = 'เสร็จแล้ว ✅\n\nแมตตรงกับทะเบียน: ' + matched + ' ราย'
          + '\nแมตใกล้เคียง-ตรวจสอบ (ระบายส้ม): ' + fuzzy + ' ราย'
          + '\nไม่พบในทะเบียน (ระบายเหลือง): ' + unmatched + ' ราย';
  if (missList.length) msg += '\n\nรายชื่อที่ยังไม่พบ (ตรวจการสะกด/เว้นวรรคกับทะเบียน):\n' + missList.join('\n');
  ui.alert(msg);
}
