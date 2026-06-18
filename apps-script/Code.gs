/*************************************************************************
 * Direct Area Expenses — Apps Script Web App backend
 * ----------------------------------------------------------------------
 * ติดตั้ง: เปิดชีต "ค่าใช้จ่าย" (COST) → Extensions → Apps Script → วางไฟล์นี้
 *          → Deploy → New deployment → Web app
 *          → Execute as: Me  |  Who has access: Anyone
 *          → คัดลอก URL ".../exec" ไปวางใน WEBAPP_URL ของหน้าเว็บ
 *
 * อ่าน  (doGet):  ?action=cost | targets | options | perf&m=5
 * เขียน (doPost): {action:'append'|'updateStatus'|'addOption', ...}
 *
 * ตอบกลับเป็น JSON และรองรับ JSONP (?callback=) เผื่อ CORS ของบางเบราว์เซอร์
 *************************************************************************/

// ===== Spreadsheet IDs =====
var COST_ID = '1F4djjhXTzhKrqXQhPVPDtjjccOfQHRFpQfvbbw-KCrs'; // ชีต "ค่าใช้จ่าย"
var PERF_ID = '18N5MmcIXpF0AS6DoVc1UxmvoBZkW4ViAi9nls4lsEP8'; // ชีต "HUB ADMIN PERFORMANCE"

// ===== ชื่อแผ่นงาน =====
var SH_EXP    = 'ข้อมูลค่าใช้จ่าย'; // log การเบิก (append / updateStatus)
var SH_DATA   = 'DATA';            // ตารางอ้างอิง (hub map + ตัวเลือก type/detail)
var SH_TARGET = 'เป้าหมายปีนี้';    // ⚠️ อยู่ในไฟล์ PERF — เป้าหมาย Cost/parcel รายหมวด/รายฮับ

// ===== หัวคอลัมน์ของ "ข้อมูลค่าใช้จ่าย" (ตรงกับฟอร์มกรอก) =====
var EXP_HEADERS = ['วันที่','วันที่ทำเบิก','ชื่อHUB','หมายเลขอ้างอิงการเบิก OA',
                   'จำนวนเงิน','รายละเอียดค่าใช้จ่าย','ประเภทค่าใช้จ่าย','หลักฐาน','สถานะ'];

/* ====================== ROUTER ====================== */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'cost';
  var out;
  try {
    if (action === 'cost')         out = { ok:true, rows: readExpenses() };
    else if (action === 'options') out = { ok:true, options: readOptions() };
    else if (action === 'targets') out = { ok:true, targets: readTargets() };
    else if (action === 'perf')    out = { ok:true, rows: readPerf(p.m) };
    else if (action === 'ping')    out = { ok:true, pong: true };
    else                           out = { ok:false, error: 'unknown action: ' + action };
  } catch (err) {
    out = { ok:false, error: String(err) };
  }
  return reply(out, p.callback);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (_) {}
  var action = body.action || '';
  var out;
  try {
    if (action === 'append')            out = appendExpense(body);
    else if (action === 'updateStatus') out = updateStatus(body);
    else if (action === 'addOption')    out = addOption(body);
    else                                out = { ok:false, error:'unknown action: ' + action };
  } catch (err) {
    out = { ok:false, error: String(err) };
  }
  return reply(out, (e && e.parameter && e.parameter.callback) || null);
}

/* ====================== READ ====================== */
function readExpenses() {
  var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_EXP);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function(h){ return String(h).replace(/\n/g,'').trim(); });
  var idx = colIndex(head, EXP_HEADERS);
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var v = values[r];
    if (!v[idx.hub] && !v[idx.oa] && !v[idx.amount]) continue; // ข้ามแถวว่าง
    rows.push({
      row: r + 1,
      date:   fmtDate(v[idx.date]),
      claimDate: fmtDate(v[idx.claimDate]),
      hub:    String(v[idx.hub] || '').trim(),
      oa:     String(v[idx.oa] || '').trim(),
      amount: toNum(v[idx.amount]),
      detail: String(v[idx.detail] || '').trim(),
      type:   String(v[idx.type] || '').trim(),
      evidence: String(v[idx.evidence] || '').trim(),
      status: String(v[idx.status] || '').trim()
    });
  }
  return rows;
}

function readOptions() {
  var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_DATA);
  if (!sh) return { types:[], details:[], hubMap:{} };
  var v = sh.getDataRange().getValues();
  var types = [], details = [], hubMap = {};
  var mode = ''; // '', 'hub', 'opt'
  for (var r = 0; r < v.length; r++) {
    var a = String(v[r][0] || '').trim();
    var b = String(v[r][1] || '').trim();
    if (a === 'Hub')               { mode = 'hub'; continue; }
    if (a === 'ประเภทค่าใช้จ่าย')   { mode = 'opt'; continue; }
    if (mode === 'hub' && a && b)  { hubMap[b] = a; }            // code → fullname
    if (mode === 'opt') {
      if (a) types.push(a);
      if (b) details.push(b);
    }
  }
  return { types: uniq(types), details: uniq(details), hubMap: hubMap };
}

// เป้าหมายปีนี้: อยู่ในไฟล์ PERF — layout แนวนอนหลายบล็อก ของ "Cost/parcel" (อัตราต้นทุน/ชิ้น) รายฮับ/หมวด
// แต่ละบล็อกหัวคอลัมน์บอกหมวด (ค่าใช้จ่ายทั่วไป/วัสดุสิ้นเปลือง/ค่าน้ำ-ค่าไฟ ...) คอลัมน์ HUB(+1) และ Cost/parcel(+2)
// คืน { HUBxxx:{ rc:{EXP,CON,RENT,ELEC} }, ... } (ค่าคืออัตราต้นทุน/ชิ้น → ตรงกับ rcTarget ในแอป) — ถ้าไม่มีแผ่น → {}
function readTargets() {
  var sh = SpreadsheetApp.openById(PERF_ID).getSheetByName(SH_TARGET);
  if (!sh) return {};
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return {};
  var head = v[0];
  // 1) จุดเริ่มของแต่ละบล็อก (cell หัวที่ระบุหมวด)
  var marks = [];
  for (var c = 0; c < head.length; c++) {
    var cat = catOfHeader(String(head[c]));
    if (cat) marks.push({ cat: cat, start: c });
  }
  if (!marks.length) return {};
  var out = {};
  for (var m = 0; m < marks.length; m++) {
    var start = marks[m].start;
    var end = (m + 1 < marks.length) ? marks[m + 1].start : head.length;
    var catKey = marks[m].cat;
    // 2) หาคอลัมน์ค่า/ฮับเองในช่วงบล็อก (รองรับ layout ต่างกัน เช่น RENT มีคอลัมน์ตัวย่อแทรก)
    var valCol = -1, hubCol = -1, c2, r2;
    for (c2 = start; c2 < end; c2++) {                       // valCol จากหัว: Cost/parcel | Target
      var hl = String(head[c2]).toLowerCase();
      if (/cost\/parcel|target|value|เป้าหมาย\//.test(hl)) valCol = c2;
    }
    for (c2 = start; c2 < end; c2++) {                       // hubCol = คอลัมน์ที่ data ให้ค่า HUB\d
      for (r2 = 1; r2 < v.length; r2++) { if (/^HUB\d/.test(normHub(String(v[r2][c2] || '')))) { hubCol = c2; break; } }
      if (hubCol >= 0) break;
    }
    if (valCol < 0) {                                        // fallback: คอลัมน์ตัวเลขขวาสุด (ไม่ใช่ hubCol)
      for (c2 = end - 1; c2 >= start; c2--) { if (c2 === hubCol) continue; var nz = 0; for (r2 = 1; r2 < v.length; r2++) { if (toNum(v[r2][c2])) nz++; } if (nz) { valCol = c2; break; } }
    }
    if (hubCol < 0 || valCol < 0) continue;
    // 3) อ่านค่ารายฮับ (ข้ามแถว AREA / รวม)
    for (r2 = 1; r2 < v.length; r2++) {
      var key = normHub(String(v[r2][hubCol] || ''));
      if (!/^HUB\d/.test(key)) continue;
      if (!out[key]) out[key] = { rc: {} };
      out[key].rc[catKey] = toNum(v[r2][valCol]);
    }
  }
  return out;
}
function catOfHeader(s) {
  s = String(s).replace(/น้ำหนัก/g, '');   // ตัด "น้ำหนัก" (ในหัว "ค่าน้ำหนัก 20%") กันชนกับ "น้ำ" ของหมวด ELEC
  if (/วัสดุ|consumable/i.test(s)) return 'CON';
  if (/ค่าเช่า|เช่า|rent/i.test(s)) return 'RENT';
  if (/น้ำ|ไฟ|ประปา|electric|water/i.test(s)) return 'ELEC';
  if (/ค่าใช้จ่ายทั่วไป|ทั่วไป|expense/i.test(s)) return 'EXP';
  return '';
}

// perf: คืน raw 2D array ของ Admin M{m} เพื่อให้ parsePerf() ฝั่งหน้าเว็บใช้ได้เหมือนเดิม
function readPerf(m) {
  var ss = SpreadsheetApp.openById(PERF_ID);
  var variants = ['Admin M'+m, 'AdminM'+m, 'Admin M.'+m];
  for (var i = 0; i < variants.length; i++) {
    var sh = ss.getSheetByName(variants[i]);
    if (sh) {
      var vals = sh.getDataRange().getValues();
      return vals.map(function(row){ return row.map(function(c){ return c == null ? '' : String(c); }); });
    }
  }
  return [];
}

/* ====================== WRITE ====================== */
function appendExpense(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_EXP);
    if (!sh) return { ok:false, error:'no sheet ' + SH_EXP };
    var head = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), EXP_HEADERS.length))
                 .getValues()[0].map(function(h){ return String(h).replace(/\n/g,'').trim(); });
    var idx = colIndex(head, EXP_HEADERS);
    var width = Math.max(head.length, EXP_HEADERS.length);
    var rowArr = new Array(width).fill('');
    var today = b.date || fmtDate(new Date());
    rowArr[idx.date]      = today;
    rowArr[idx.claimDate] = b.claimDate || today;
    rowArr[idx.hub]       = b.hub || '';
    rowArr[idx.oa]        = b.oa || '';
    rowArr[idx.amount]    = toNum(b.amount);
    rowArr[idx.detail]    = b.detail || '';
    rowArr[idx.type]      = b.type || '';
    rowArr[idx.evidence]  = b.evidence || '';
    rowArr[idx.status]    = b.status || 'รอ';
    sh.appendRow(rowArr);
    return { ok:true, row: sh.getLastRow() };
  } finally { lock.releaseLock(); }
}

function updateStatus(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_EXP);
    if (!sh) return { ok:false, error:'no sheet ' + SH_EXP };
    var head = sh.getDataRange().getValues()[0].map(function(h){ return String(h).replace(/\n/g,'').trim(); });
    var idx = colIndex(head, EXP_HEADERS);
    var rowNum = Number(b.row) || 0;
    if (!rowNum && b.oa) {                       // หาแถวจากเลข OA ถ้าไม่ส่ง row มา
      var col = sh.getRange(2, idx.oa+1, Math.max(sh.getLastRow()-1,1), 1).getValues();
      for (var i = 0; i < col.length; i++) {
        if (String(col[i][0]).trim() === String(b.oa).trim()) { rowNum = i + 2; break; }
      }
    }
    if (!rowNum) return { ok:false, error:'row not found' };
    sh.getRange(rowNum, idx.status+1).setValue(b.status || '');
    if (b.evidence) sh.getRange(rowNum, idx.evidence+1).setValue(b.evidence);
    return { ok:true, row: rowNum };
  } finally { lock.releaseLock(); }
}

// addOption: เพิ่มตัวเลือกใหม่เข้าแผ่น DATA (kind='detail' → คอลัมน์ B, kind='type' → คอลัมน์ A ใต้ header ตัวเลือก)
function addOption(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_DATA);
    if (!sh) return { ok:false, error:'no sheet ' + SH_DATA };
    var val = String(b.val || '').trim();
    if (!val) return { ok:false, error:'empty value' };
    var v = sh.getDataRange().getValues();
    var optHeaderRow = -1;
    for (var r = 0; r < v.length; r++) { if (String(v[r][0]).trim() === 'ประเภทค่าใช้จ่าย') { optHeaderRow = r; break; } }
    var col = (b.kind === 'type') ? 0 : 1;       // A=type, B=detail
    // กันซ้ำ
    for (var r2 = 0; r2 < v.length; r2++) { if (String(v[r2][col]).trim() === val) return { ok:true, dup:true }; }
    // หาแถวว่างถัดไปในคอลัมน์นั้น (เริ่มใต้ header ตัวเลือกถ้ามี)
    var start = optHeaderRow >= 0 ? optHeaderRow + 1 : 0;
    var target = -1;
    for (var r3 = start; r3 < v.length; r3++) { if (!String(v[r3][col]).trim()) { target = r3; break; } }
    if (target < 0) target = v.length;           // ต่อท้ายสุด
    sh.getRange(target + 1, col + 1).setValue(val);
    return { ok:true, row: target + 1, col: col + 1 };
  } finally { lock.releaseLock(); }
}

/* ====================== helpers ====================== */
function colIndex(head, headers) {
  function find(name){
    var t = name.replace(/\n/g,'').trim();
    for (var i=0;i<head.length;i++){ if (head[i] === t) return i; }
    for (var j=0;j<head.length;j++){ if (head[j].indexOf(t) === 0 || t.indexOf(head[j]) === 0) return j; }
    return -1;
  }
  return {
    date:      pick(find(headers[0]), 0),
    claimDate: pick(find(headers[1]), 1),
    hub:       pick(find(headers[2]), 2),
    oa:        pick(find(headers[3]), 3),
    amount:    pick(find(headers[4]), 4),
    detail:    pick(find(headers[5]), 5),
    type:      pick(find(headers[6]), 6),
    evidence:  pick(find(headers[7]), 7),
    status:    pick(find(headers[8]), 8)
  };
}
function pick(found, fallback){ return found >= 0 ? found : fallback; }
function normHub(s){ var c = String(s).replace(/\s/g,'').toUpperCase(); var m = c.match(/(\d{2}[A-Z]{3})/); if (m) c = m[1]; return c.indexOf('HUB')===0 ? c : 'HUB'+c; }
function toNum(v){ if (typeof v === 'number') return v; var x = parseFloat(String(v).replace(/[^0-9.\-]/g,'')); return isNaN(x) ? 0 : x; }
function uniq(a){ var seen={}, out=[]; a.forEach(function(x){ if(x && !seen[x]){ seen[x]=1; out.push(x); } }); return out; }
function fmtDate(d){
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return String(d).trim();
}
function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
