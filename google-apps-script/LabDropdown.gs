/**
 * LabDropdown.gs — ลงผลตรวจในชีตด้วยดรอปดาวน์ (+ / - / na)
 * คอลัมน์: Chest X-ray, Sputum AFB, TST, LAB ในแท็บ "ข้อมูลแบบสอบถาม"
 *
 *  • setupLabDropdowns()  ใส่ดรอปดาวน์ (+/-/na) ให้ 4 คอลัมน์ (สร้างคอลัมน์ถ้ายังไม่มี)
 *  • syncLabResults()     อ่านค่าในชีต -> เขียนลง _JSON (rec.labs) ให้เว็บ/รายงานแสดง
 *
 * เมนูรวมอยู่ใน onOpen ของ MatchRoster.gs (เมนู "🧪 ผลตรวจ")
 * หมายเหตุ: การลงผลในชีตเป็น "ผลเดียวต่อรายการ" (1 ครั้ง) ส่วนการลงหลายครั้งใช้ปุ่ม "ผลตรวจ" ในเว็บ
 */
var LAB_SHEET = 'ข้อมูลแบบสอบถาม';
var LAB_DEFS  = [['cxr', 'Chest X-ray'], ['afb', 'Sputum AFB'], ['tst', 'TST'], ['lab', 'LAB']];
var LAB_CHOICES = ['+', '-', 'na'];

function _labSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB_SHEET); }
function _labNorm(s) { return String(s == null ? '' : s).replace(/\s+/g, '').trim(); }
function _labFindCols(header) {
  // คืน map key -> index (0-based) ; หาแบบตรงตัวก่อน แล้วค่อย "มีคำว่า"
  var idx = {};
  LAB_DEFS.forEach(function (d) {
    var want = _labNorm(d[1]).toLowerCase(), found = -1;
    for (var i = 0; i < header.length; i++) if (_labNorm(header[i]).toLowerCase() === want) { found = i; break; }
    if (found < 0) for (var j = 0; j < header.length; j++) if (_labNorm(header[j]).toLowerCase().indexOf(want) >= 0) { found = j; break; }
    idx[d[0]] = found;
  });
  return idx;
}

function setupLabDropdowns() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sh = _labSheet();
    if (!sh) { ui.alert('ไม่พบแท็บ "' + LAB_SHEET + '"'); return; }
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last < 2) { ui.alert('ไม่มีข้อมูล'); return; }
    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var jsonIdx = -1;
    for (var i = 0; i < header.length; i++) if (_labNorm(header[i]) === '_JSON') jsonIdx = i;
    var idx = _labFindCols(header);
    // สร้างคอลัมน์ที่ยังไม่มี (แทรกก่อน _JSON ถ้ามี ไม่งั้นต่อท้าย)
    LAB_DEFS.forEach(function (d) {
      if (idx[d[0]] < 0) {
        var insertAt = (jsonIdx >= 0 ? jsonIdx + 1 : sh.getLastColumn() + 1);
        sh.insertColumnBefore(insertAt);
        sh.getRange(1, insertAt).setValue(d[1]).setFontWeight('bold');
        // อ่าน header ใหม่หลังแทรก
        lastCol = sh.getLastColumn();
        header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
        for (var k = 0; k < header.length; k++) if (_labNorm(header[k]) === '_JSON') jsonIdx = k;
        idx = _labFindCols(header);
      }
    });
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(LAB_CHOICES, true).setAllowInvalid(false).build();
    LAB_DEFS.forEach(function (d) {
      var c = idx[d[0]];
      if (c >= 0) sh.getRange(2, c + 1, last - 1, 1).setDataValidation(rule);
    });
    ui.alert('ใส่ดรอปดาวน์ผลตรวจ (+/-/na) ให้คอลัมน์ Chest X-ray, Sputum AFB, TST, LAB แล้ว\n\nลงผลในชีตได้เลย จากนั้นเลือกเมนู "🧪 ผลตรวจ → ซิงก์ผลตรวจเข้าระบบ"');
  } catch (err) { ui.alert('setupLabDropdowns error: ' + (err && err.message ? err.message : err)); }
}

function syncLabResults() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sh = _labSheet();
    if (!sh) { ui.alert('ไม่พบแท็บ "' + LAB_SHEET + '"'); return; }
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last < 2) { ui.alert('ไม่มีข้อมูล'); return; }
    var vals = sh.getRange(1, 1, last, lastCol).getValues();
    var header = vals[0];
    var jsonIdx = -1;
    for (var i = 0; i < header.length; i++) if (_labNorm(header[i]) === '_JSON') jsonIdx = i;
    if (jsonIdx < 0) { ui.alert('ไม่พบคอลัมน์ _JSON'); return; }
    var idx = _labFindCols(header);
    var block = sh.getRange(2, 1, last - 1, lastCol).getValues();
    var changed = 0;
    for (var r = 0; r < block.length; r++) {
      var row = block[r];
      var js = row[jsonIdx];
      if (!js) continue;
      var rec; try { rec = JSON.parse(js); } catch (e) { continue; }
      var labs = {}, any = false;
      LAB_DEFS.forEach(function (d) {
        var c = idx[d[0]];
        if (c >= 0) {
          var v = String(row[c] == null ? '' : row[c]).trim();
          if (v === '+' || v === '-' || v === 'na') { labs[d[0]] = [{ date: '', result: v }]; any = true; }
        }
      });
      if (any) { rec.labs = labs; row[jsonIdx] = JSON.stringify(rec); changed++; }
      else if (rec.labs) { /* ไม่มีค่าในชีต -> คงค่าเดิมใน _JSON ไว้ (ไม่ล้าง) */ }
    }
    sh.getRange(2, 1, block.length, lastCol).setValues(block);
    ui.alert('ซิงก์ผลตรวจเข้าระบบแล้ว ' + changed + ' รายการ\n(รีเฟรชหน้าเว็บเพื่อดูผลในรายชื่อ/รายงาน)');
  } catch (err) { ui.alert('syncLabResults error: ' + (err && err.message ? err.message : err)); }
}
