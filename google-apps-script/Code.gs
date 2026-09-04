/**
 * ระบบสอบสวนการระบาด (Outbreak Investigation System)
 * Google Apps Script — ฐานข้อมูลกลางบน Google Sheets
 *
 * ชีตที่ใช้ (สร้างอัตโนมัติถ้ายังไม่มี):
 *   1) "ข้อมูลแบบสอบถาม"      — เก็บคำตอบทุกฉบับ (1 แถวต่อ 1 การบันทึก)
 *   2) "การจัดการแบบสอบถาม"   — เก็บโครงสร้างแบบสอบถาม (schema) ที่แอดมินแก้ผ่านหน้าเว็บ
 *
 * วิธี deploy อยู่ในไฟล์ README.md
 */

var SHEET_DATA = 'ข้อมูลแบบสอบถาม';
var SHEET_SCHEMA = 'การจัดการแบบสอบถาม';      // เก็บ JSON สำรอง (A2)
var SHEET_STRUCT = 'โครงสร้างแบบสอบถาม';      // ตารางแก้ไขคำถามแบบมือ (source of truth)
var META_COLS = ['เวลาบันทึก', 'id', 'ประเภท', 'ชื่อ-สกุล', 'HN', 'กลุ่ม/ตำแหน่ง', 'สถานะ'];

// ★ ถ้าสคริปต์ไม่ได้ผูกกับชีต (สร้างแบบ standalone) ให้วาง ID ของ Google Sheet ที่นี่
//   ID คือส่วนระหว่าง /d/ กับ /edit ใน URL ของชีต เช่น
//   https://docs.google.com/spreadsheets/d/1AbCdEfGhIjK.../edit  ->  '1AbCdEfGhIjK...'
var SPREADSHEET_ID = '1-vvdG18uzzn9EQgSAnQ1G4aCLdUvNIQA6deWZrk92EI';

/* ---------------- HTTP entry points ---------------- */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'load';
  var out;
  if (action === 'load') out = { schema: readSchema(), records: readRecords(), settings: readSettings() };
  else out = { error: 'unknown action' };

  // รองรับ JSONP (เรียกจากเว็บสถิตข้ามโดเมนได้ ผ่าน callback)
  var callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(out);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // กันการเขียนพร้อมกันจากหลายเครื่อง
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'submit') { appendRecord(body.record); return json({ ok: true }); }
    if (body.action === 'saveSchema') { writeSchema(body.schema); return json({ ok: true }); }
    if (body.action === 'delete') { deleteRecord(body.id); return json({ ok: true }); }
    if (body.action === 'saveSettings') { writeSettings(body.settings); return json({ ok: true }); }
    return json({ error: 'unknown action' });
  } catch (err) {
    return json({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- helpers ---------------- */

function ss() {
  var s = SpreadsheetApp.getActiveSpreadsheet();
  if (!s && SPREADSHEET_ID) s = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!s) throw new Error('ไม่พบสเปรดชีต: ให้เปิด Apps Script ผ่านเมนู "ส่วนขยาย → Apps Script" ของชีต หรือกรอก SPREADSHEET_ID ในโค้ด');
  return s;
}

function getSheet(name) {
  var s = ss().getSheetByName(name);
  if (!s) s = ss().insertSheet(name);
  return s;
}

/* ---------------- SCHEMA (แท็บการจัดการแบบสอบถาม) ---------------- */

// อ่าน schema: ให้ความสำคัญกับตาราง "โครงสร้างแบบสอบถาม" ก่อน ถ้าไม่มีจึงใช้ JSON สำรอง
function readSchema() {
  var s = readSchemaFromStruct();
  if (s) return s;
  var sh = getSheet(SHEET_SCHEMA);
  var v = sh.getRange(2, 1).getValue();
  if (!v) return null;
  try { return JSON.parse(v); } catch (e) { return null; }
}

// เขียน schema: เก็บ JSON สำรอง + สร้างตารางแก้ไขให้ตรงกัน
function writeSchema(schema) {
  var sh = getSheet(SHEET_SCHEMA);
  sh.getRange(1, 1).setValue('SCHEMA_JSON (สำรอง) — แก้แบบสอบถามได้ที่แท็บ "โครงสร้างแบบสอบถาม" หรือหน้าแอดมินของเว็บแอป');
  sh.getRange(2, 1).setValue(JSON.stringify(schema));
  writeSchemaToStruct(schema);
}

var STRUCT_HEADER = ['แบบฟอร์ม', 'ส่วนที่', 'หัวข้อส่วน', 'ไอคอน', 'รหัสข้อ', 'คำถาม', 'ชนิด', 'บังคับตอบ', 'ตัวเลือก (บรรทัดละ 1)', 'ค่าตายตัว'];

// สร้าง/อัปเดตตารางโครงสร้างจาก schema
function writeSchemaToStruct(schema) {
  var sh = ss().getSheetByName(SHEET_STRUCT) || ss().insertSheet(SHEET_STRUCT);
  sh.clearContents();
  var rows = [STRUCT_HEADER];
  ['staff', 'patient'].forEach(function (form) {
    var secs = (schema[form] && schema[form].sections) ? schema[form].sections : [];
    secs.forEach(function (sec, si) {
      sec.fields.forEach(function (f) {
        var opt = '';
        if (f.type === 'select' || f.type === 'radio' || f.type === 'checkbox') opt = (f.options || []).join('\n');
        else if (f.type === 'contactlog') opt = JSON.stringify({ contactOptions: f.contactOptions || {}, contactLabels: f.contactLabels || {} });
        rows.push([form, si + 1, sec.title || '', sec.icon || '', f.id, f.label || '', f.type || 'text', !!f.required, opt, (f.fixed !== undefined ? f.fixed : '')]);
      });
    });
  });
  sh.getRange(1, 1, rows.length, STRUCT_HEADER.length).setValues(rows);
  sh.getRange(1, 1, 1, STRUCT_HEADER.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

// อ่าน schema จากตารางโครงสร้าง (คืน null ถ้ายังไม่มีข้อมูล)
function readSchemaFromStruct() {
  var sh = ss().getSheetByName(SHEET_STRUCT);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, STRUCT_HEADER.length).getValues();
  var schema = { staff: { sections: [] }, patient: { sections: [] } };
  var secMap = {};
  var hasOpts = { select: 1, radio: 1, checkbox: 1 };
  vals.forEach(function (r) {
    var form = String(r[0] || '').trim();
    if (form !== 'staff' && form !== 'patient') return;
    var secNo = String(r[1] || '').trim();
    var secTitle = String(r[2] || '').trim();
    var icon = String(r[3] || '').trim();
    var fid = String(r[4] || '').trim();
    var label = String(r[5] || '').trim();
    var type = String(r[6] || 'text').trim();
    var reqv = r[7];
    var required = (reqv === true || String(reqv).toLowerCase() === 'true' || String(reqv) === 'ใช่' || String(reqv) === '1');
    var optRaw = r[8];
    var fixed = (r[9] === null || r[9] === undefined) ? '' : String(r[9]);
    if (!fid || !label) return;
    var key = form + '#' + (secNo || secTitle);
    var sec = secMap[key];
    if (!sec) {
      sec = { id: 's' + (schema[form].sections.length + 1), title: secTitle, icon: icon || 'chevron-right', fields: [] };
      secMap[key] = sec;
      schema[form].sections.push(sec);
    }
    var field = { id: fid, label: label, type: type };
    if (required) field.required = true;
    if (hasOpts[type]) {
      field.options = String(optRaw || '').split(/\r?\n|\|/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
    } else if (type === 'contactlog') {
      try { var o = JSON.parse(String(optRaw || '{}')); if (o.contactOptions) field.contactOptions = o.contactOptions; if (o.contactLabels) field.contactLabels = o.contactLabels; } catch (e) {}
    }
    if (fixed !== '') field.fixed = fixed;
    sec.fields.push(field);
  });
  if (schema.staff.sections.length === 0 && schema.patient.sections.length === 0) return null;
  return schema;
}

/* ---------------- SETTINGS (เก็บใน SHEET_SCHEMA เซลล์ A4) ---------------- */
function readSettings() {
  var sh = getSheet(SHEET_SCHEMA);
  var v = sh.getRange(4, 1).getValue();
  if (!v) return {};
  try { return JSON.parse(v); } catch (e) { return {}; }
}
function writeSettings(obj) {
  var sh = getSheet(SHEET_SCHEMA);
  sh.getRange(3, 1).setValue('SETTINGS_JSON (การตั้งค่าเมนู ฯลฯ — จัดการผ่านหน้าแอดมิน)');
  sh.getRange(4, 1).setValue(JSON.stringify(obj || {}));
}

/* ---------------- RECORDS (แท็บข้อมูลแบบสอบถาม) ---------------- */

// รวมคอลัมน์คำถามจากทั้งแบบเจ้าหน้าที่และผู้ป่วย (ไม่ซ้ำ id)
function fieldLabels(schema) {
  var cols = [], seen = {};
  ['staff', 'patient'].forEach(function (fk) {
    var secs = (schema[fk] && schema[fk].sections) ? schema[fk].sections : [];
    secs.forEach(function (sec) {
      sec.fields.forEach(function (f) {
        if (!seen[f.id]) { seen[f.id] = true; cols.push({ id: f.id, label: f.label }); }
      });
    });
  });
  return cols;
}

// แปลงคำตอบ (array / บันทึกการสัมผัส) ให้อ่านง่ายในชีต
function flatten(v) {
  if (v == null) return '';
  if (Array.isArray(v)) {
    if (v.length && typeof v[0] === 'object') {
      return v.map(function (e, i) {
        var act = (e.activities && e.activities.length) ? e.activities.join(', ') : '-';
        return 'ครั้งที่ ' + (i + 1) + ': ' + (e.date || '-') +
          ' [กิจกรรม: ' + act + '] (N95: ' + (e.n95 || '-') +
          ', ระยะเวลา: ' + (e.duration || '-') + ', ห้อง: ' + (e.room || '-') + ')';
      }).join(' | ');
    }
    return v.join(', ');
  }
  return String(v);
}

// ให้แถวหัวตารางครบตาม schema (เพิ่ม/ปรับอัตโนมัติเมื่อ schema เปลี่ยน)
function ensureHeader(sh, cols) {
  var need = META_COLS.concat(cols.map(function (c) { return c.label; })).concat(['_JSON']);
  var lastCol = sh.getLastColumn();
  var cur = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var ok = cur.length === need.length && need.every(function (h, i) { return cur[i] === h; });
  if (!ok) {
    sh.getRange(1, 1, 1, need.length).setValues([need]);
    sh.getRange(1, 1, 1, need.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return need;
}

function appendRecord(rec) {
  var schema = readSchema() || { staff: { sections: [] }, patient: { sections: [] } };
  var cols = fieldLabels(schema);
  var sh = getSheet(SHEET_DATA);
  ensureHeader(sh, cols);
  var ans = rec.answers || {};
  var row = [
    new Date(),
    rec.id || '',
    rec.type === 'patient' ? 'ผู้ป่วย' : 'เจ้าหน้าที่',
    rec.name || '',
    rec.hn || '',
    rec.group || '',
    rec.status || ''
  ];
  cols.forEach(function (c) { row.push(flatten(ans[c.id])); });
  row.push(JSON.stringify(rec)); // เก็บ JSON เต็มไว้ให้เว็บแอปอ่านกลับได้ครบ
  sh.appendRow(row);
}

function deleteRecord(id) {
  if (!id) return;
  var sh = getSheet(SHEET_DATA);
  var last = sh.getLastRow();
  if (last < 2) return;
  var idVals = sh.getRange(2, 2, last - 1, 1).getValues(); // คอลัมน์ B = id
  for (var i = idVals.length - 1; i >= 0; i--) {
    if (String(idVals[i][0]) === String(id)) sh.deleteRow(i + 2);
  }
}

function readRecords() {
  var sh = getSheet(SHEET_DATA);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var jsonIdx = header.indexOf('_JSON');
  if (jsonIdx < 0) return [];
  var values = sh.getRange(2, jsonIdx + 1, last - 1, 1).getValues();
  var out = [];
  values.forEach(function (r) {
    if (r[0]) { try { out.push(JSON.parse(r[0])); } catch (e) {} }
  });
  return out.reverse(); // ล่าสุดอยู่บนสุด
}
