/**
 * เติม HN ลงในทะเบียน "รายชื่อเจ้าหน้าที่ปฏิบัติงาน" โดยจับคู่ตามชื่อ
 * ข้อมูล HN ฝังในไฟล์นี้ (จากไฟล์ contact_69) — รันครั้งเดียวผ่านเมนู
 *
 * วิธีใช้:
 *   1) วางไฟล์นี้เป็นสคริปต์ใหม่ในโปรเจกต์ Apps Script ของชีต
 *   2) รีเฟรชชีต แล้วเลือกเมนู "จับคู่ HN" -> "เติม HN ลงทะเบียน"
 *   3) แถวที่จับคู่ไม่ได้จะถูกไฮไลต์สีเหลืองให้ตรวจสอบ
 */

var ROSTER_SHEET = 'รายชื่อเจ้าหน้าที่ปฏิบัติงาน';

// map: ชื่อที่ normalize แล้ว -> HN
var HN_MAP = {
    "นฤมลวงศ์ประเสริฐ": "0346708",
    "สิริวรรณสีถา": "0122160",
    "สุธาสินีเต็มแบบ": "0048621",
    "บุษบาวรรณสว่างทรัพย์": "0366603",
    "เพชรรัตน์ธีปะนะ": "0118958",
    "ทิวัตถ์หอมหวน": "410011",
    "วรางคณาพุทธรักษ์": "0000737",
    "นภัทร์มณฑ์จันทร์น้อย": "0109200",
    "มาลัยทรัพย์สระนิตย์": "0285529",
    "ธีรยุทธเฉิดฉาย": "0042675",
    "วิไลวรรณชูจันทร์": "0026629",
    "ธัญชนกนันทโชติ": "417156",
    "ณัฐฐิราบุญกลิ่น": "0171649",
    "นริศราพุ่มชุ่ม": "0145955",
    "พิชญ์สินีคีรีวิเชียร": "0107050",
    "ปรัชญาตันติพลาผล": "0135818",
    "ชินดนัยวงศ์นิล": "0146954",
    "ศิริรัตน์ศรีบำรุง": "408811",
    "พุทธธิดาสวัสดิรักษา": "0026548",
    "กรกนกล้วนพฤกษ์": "0190170",
    "นาวินศักดาเดช": "0396241",
    "นาวาจันทร์โท": "0163693",
    "นนทณัฏฐ์มิ่งศิริรัตน์": "0168080",
    "ระพินอินทรมาศ": "0002779",
    "วสันติ์บุญเพชร": "0147110",
    "สำราญบุญธรรม": "0020410",
    "เฉลิมชาตินาคะเกศ": "170951",
    "รัตนาพรจันทร์พวง": "423487",
    "วรรณรายอารีย์": "0024733",
    "เยาวดีเป็งญาวงศ์": "0007488",
    "ณัฐวุฒิเปียชู": "339463",
    "พรเพ็ญวิรัตน์ชัยวรรณ": "0048596",
    "ภณิดาเมตตานี": "0257691",
    "ชาตรีชัยมั่น": "0023193",
    "ทัศน์วรรณปราณีต": "0225874",
    "ปาริฉัตรสุขสด": "287858",
    "มณฑกานต์ลาโภ": "0108166",
    "ทิพย์สุดาสุภาคุณ": "0378989",
    "สายธารขวัญดี": "414613",
    "เขษมศักดิ์อนันตะ": "192690",
    "นลินนิภาเพ็ชรนิล": "187413",
    "สุมาลีคำชัยยะ": "267649",
    "ปณิธานเปลื้องคีรีรัมย์": "0087684",
    "ศิวรัตน์ชภามวลศรี": "99084",
    "ศิริลักษณ์แสนคำ": "412001",
    "บานเย็นนงค์โภชน์": "0247401",
    "ดลนภัสบริบูรณ์": "0058916",
    "สุรางคณาประจันสุกุณี": "0352377",
    "อารีรัตน์สุธารักษ์": "0035825",
    "เกตุสุณีย์แก้วบุรี": "0041287",
    "วิรัลพัชรบุญแตง": "0078895",
    "จุฬารัตน์จรุงพันธุ์": "0398395",
    "วรวิชเตชัย": "0387427",
    "พิชญาภาไกรกิจราษฎร์": "405998",
    "วรรณวิมลมั่นคง": "0011194",
    "พงศ์ปณตโพอุทัย": "410366",
    "ปณิดานิรชน": "189813",
    "วงศ์เดือนศรีธรรม": "0260826",
    "วาสนามิดชิด": "139214",
    "วันดีสังข์วงค์": "167284",
    "วรินทรเลื่อนลอย": "196204",
    "ญาณัจฉาศรีสว่าง": "0192834",
    "ฉันทนาม่วงเกตุ": "0040216",
    "สุธาสินีทองชั่ง": "0135793",
    "วรัทยาวิเศษ": "0058913",
    "อัชราภรณ์ตั้งกิติวงศ์": "0000740",
    "อัญชันปันลุน": "221946",
    "รัชกรตุ่มน้ำ": "282122",
    "อนุสราไกรกิจราษฎร์": "299188"
};

var HN_PREFIXES = ['นางสาว', 'น.ส.', 'นส.', 'นาย', 'นาง', 'ดร.', 'นพ.', 'พญ.', 'ทพ.', 'ทพญ.', 'ภก.', 'ภญ.', 'ว่าที่ร้อยตรี', 'ว่าที่ ร.ต.'];
function hnNameKey(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFC'); } catch (e) {}
  s = s.replace(/\s+/g, ' ').trim();
  var ch = true;
  while (ch) { ch = false; for (var i = 0; i < HN_PREFIXES.length; i++) { var p = HN_PREFIXES[i]; if (s.indexOf(p) === 0) { s = s.slice(p.length).replace(/^[\s.]+/, ''); ch = true; break; } } }
  return s.replace(/[\s.()\-]/g, '');
}

// หมายเหตุ: เมนูรวมอยู่ใน onOpen ของ MatchRoster.gs แล้ว (Apps Script มี onOpen ได้ตัวเดียวต่อโปรเจกต์)
// ถ้าไม่มี MatchRoster.gs ในโปรเจกต์ ให้เอาคอมเมนต์ออกจากฟังก์ชัน onOpen ด้านล่างเพื่อสร้างเมนูเอง
// function onOpen() {
//   SpreadsheetApp.getUi().createMenu('จับคู่ HN').addItem('เติม HN ลงทะเบียน', 'fillRosterHN').addToUi();
// }

function fillRosterHN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ROSTER_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('ไม่พบแท็บ "' + ROSTER_SHEET + '"'); return; }
  var last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('ไม่มีข้อมูลในทะเบียน'); return; }
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(1, 1, last, lastCol).getValues();
  var header = vals[0];

  function findCol(names) {
    for (var i = 0; i < header.length; i++) {
      var h = String(header[i] || '').replace(/\s+/g, '').trim();
      for (var j = 0; j < names.length; j++) if (h === names[j]) return i;
    }
    return -1;
  }
  var cName = findCol(['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล']);
  var cFirst = findCol(['ชื่อ']), cLast = findCol(['นามสกุล', 'สกุล']);
  var cHN = findCol(['HN', 'hn', 'เอชเอ็น']);

  // ถ้ายังไม่มีคอลัมน์ HN ให้เพิ่มต่อท้าย
  if (cHN < 0) { cHN = lastCol; sh.getRange(1, cHN + 1).setValue('HN'); sh.getRange(1, cHN + 1).setFontWeight('bold'); }

  var matched = 0, notFound = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var name = cName >= 0 ? String(row[cName] || '')
      : (String(cFirst >= 0 ? row[cFirst] : '') + ' ' + String(cLast >= 0 ? row[cLast] : ''));
    name = name.replace(/\s+/g, ' ').trim();
    if (!name) continue;
    var k = hnNameKey(name);
    var hn = HN_MAP[k];
    var cell = sh.getRange(r + 1, cHN + 1);
    if (hn) {
      cell.setNumberFormat('@');       // เก็บ HN เป็นข้อความ กันเลข 0 นำหน้าหาย
      cell.setValue(hn);
      cell.setBackground(null);
      matched++;
    } else {
      cell.setBackground('#fff3cd');   // ไฮไลต์แถวที่ไม่พบ
      notFound.push(name);
    }
  }
  SpreadsheetApp.getUi().alert('เติม HN แล้ว ' + matched + ' คน\nไม่พบ ' + notFound.length + ' คน' +
    (notFound.length ? ':\n- ' + notFound.join('\n- ') : ''));
}

/**
 * เติม HN ลงในแท็บ "ข้อมูลแบบสอบถาม" (ที่เว็บ/รายงานอ่าน) โดยจับคู่ตามชื่อ
 * เขียนทั้งคอลัมน์ HN และก้อน _JSON (rec.hn + answers.hn) ให้ตรงกัน
 */
var HN_DATA_SHEET = 'ข้อมูลแบบสอบถาม';
function fillDataHN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HN_DATA_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('ไม่พบแท็บ "' + HN_DATA_SHEET + '"'); return; }
  var last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('ไม่มีข้อมูล'); return; }
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(1, 1, last, lastCol).getValues();
  var header = vals[0];
  function findCol(names) {
    for (var i = 0; i < header.length; i++) {
      var h = String(header[i] || '').replace(/\s+/g, '').trim();
      for (var j = 0; j < names.length; j++) if (h === names[j]) return i;
    }
    return -1;
  }
  var cName = findCol(['ชื่อ-สกุล', 'ชื่อ-นามสกุล', 'ชื่อสกุล']);
  var cHN = findCol(['HN', 'hn']);
  var cJson = findCol(['_JSON']);
  if (cName < 0) { SpreadsheetApp.getUi().alert('ไม่พบคอลัมน์ "ชื่อ-สกุล"'); return; }
  if (cHN < 0) { cHN = lastCol; sh.getRange(1, cHN + 1).setValue('HN'); sh.getRange(1, cHN + 1).setFontWeight('bold'); lastCol++; }

  var matched = 0, notFound = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var name = String(row[cName] || '').trim();
    if (!name) continue;
    var hn = HN_MAP[hnNameKey(name)];
    var cell = sh.getRange(r + 1, cHN + 1);
    if (hn) {
      cell.setNumberFormat('@'); cell.setValue(hn); cell.setBackground(null);
      // อัปเดต _JSON ให้ hn เปลี่ยนตามด้วย (เว็บ/รายงานอ่านจากตรงนี้)
      if (cJson >= 0 && row[cJson]) {
        try {
          var rec = JSON.parse(row[cJson]);
          rec.hn = hn; rec.answers = rec.answers || {}; rec.answers.hn = hn;
          sh.getRange(r + 1, cJson + 1).setValue(JSON.stringify(rec));
        } catch (e) {}
      }
      matched++;
    } else {
      cell.setBackground('#fff3cd');
      notFound.push(name);
    }
  }
  SpreadsheetApp.getUi().alert('เติม HN ในข้อมูลแบบสอบถามแล้ว ' + matched + ' คน\nไม่พบ ' + notFound.length + ' คน' +
    (notFound.length ? ':\n- ' + notFound.join('\n- ') : ''));
}
