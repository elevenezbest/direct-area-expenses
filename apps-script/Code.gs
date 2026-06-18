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
var FORECAST_ID = '1Md45vwrdCXIQReh9e58OHz-uRNsyV25dZ4QnK7yIpQQ'; // ไฟล์ "Forcast KPI" (แผ่น Forcast KPI M{n}) — แหล่งเป้าหมาย
var SH_USERS  = 'ผู้ใช้งาน';        // บัญชีผู้ใช้ + รหัสผ่าน (สร้างอัตโนมัติถ้ายังไม่มี)
var DEFAULT_PW = '123456';         // รหัสเข้าครั้งแรกของทุกคน → ต้องตั้งใหม่หลังเข้าครั้งแรก

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
    else if (action === 'forecast')out = { ok:true, month: p.m, targets: readForecast(p.m) };
    else if (action === 'users')   out = { ok:true, users: readUsers() };
    else if (action === 'formulas')out = { ok:true, sheet:p.sheet, formulas: readFormulas(p.id, p.sheet, p.r, p.c) };
    else if (action === 'sheets')  out = { ok:true, names: SpreadsheetApp.openById(p.id).getSheets().map(function(s){return s.getName();}) };
    else if (action === 'ping')    out = { ok:true, pong: true };
    else if (action === 'debug')   out = { ok:true, debug: debugInfo() };
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
    else if (action === 'login')        out = loginUser(body);
    else if (action === 'setpw')        out = setUserPw(body);
    else if (action === 'resetpw')      out = adminResetPw(body);
    else if (action === 'adduser')      out = addUser(body);
    else if (action === 'writeForecast')out = writeForecast(body);
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
// หาแผ่น "เป้าหมายปีนี้" แบบยืดหยุ่น (เผื่อชื่อแท็บมีช่องว่าง/อักขระซ่อน ที่ getSheetByName เป๊ะ ๆ หาไม่เจอ)
function getTargetSheet() {
  var ss = SpreadsheetApp.openById(PERF_ID);
  var sh = ss.getSheetByName(SH_TARGET);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var n = String(all[i].getName()).replace(/\s/g, '');
    if (n === SH_TARGET.replace(/\s/g, '') || n.indexOf('เป้าหมาย') === 0) return all[i];
  }
  return null;
}
// debug: ดูชื่อแท็บทั้งหมด + หัวตารางของแผ่นเป้าหมาย เพื่อวินิจฉัยตอนต่อไม่ติด
function debugInfo() {
  var ss = SpreadsheetApp.openById(PERF_ID);
  var names = ss.getSheets().map(function(s){ return s.getName(); });
  var sh = getTargetSheet();
  var rows = sh ? sh.getDataRange().getValues().slice(0, 4).map(function(r){ return r.map(function(c){ return String(c).slice(0, 40); }); }) : null;
  return { perfSheets: names, targetFound: !!sh, targetName: sh ? sh.getName() : null, first4rows: rows };
}
function readTargets() {
  var sh = getTargetSheet();
  if (!sh) return {};
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return {};
  // 0) หาแถวหัวตารางเอง = แถวที่เจอชื่อหมวดมากสุด (หัวอาจไม่ได้อยู่แถวแรก / มีแถวว่าง/merge ด้านบน)
  var hdrRow = 0, bestHits = 0;
  for (var rr = 0; rr < Math.min(v.length, 10); rr++) {
    var hits = 0;
    for (var cc = 0; cc < v[rr].length; cc++) { if (catOfHeader(String(v[rr][cc]))) hits++; }
    if (hits > bestHits) { bestHits = hits; hdrRow = rr; }
  }
  if (!bestHits) return {};
  var head = v[hdrRow];
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

// forecast: อ่านเป้าหมายจากไฟล์ Forecast แผ่น "Forcast KPI M{m}" (label-based กันแถวเลื่อน)
// คืน { HUBxxx|ALL : { pTarget, cats:{EXP|CON|RENT|ELEC:{rc,baht}}, aTarget, rTarget, fTarget }, ... }
function readForecast(m) {
  var sh = SpreadsheetApp.openById(FORECAST_ID).getSheetByName('Forcast KPI M' + m);
  if (!sh) return {};
  var v = sh.getDataRange().getValues();
  function catOf(e) {
    if (/วัสดุ|consumable/i.test(e)) return 'CON';
    if (/เช่า|rent/i.test(e)) return 'RENT';
    if (/น้ำ|ไฟ|ประปา|electric|water/i.test(e)) return 'ELEC';
    if (/ค่าใช้จ่าย|ทั่วไป|expense/i.test(e)) return 'EXP';
    return '';
  }
  var out = {}, cur = null;
  for (var r = 0; r < v.length; r++) {
    var A = String(v[r][0] || '').trim();
    var E = String(v[r][4] || '').replace(/\n/g, ' ').trim();
    var F = v[r][5];
    if (A && (/\d{2}\s*[A-Z]{2,}/.test(A) || /DIRECT/i.test(A))) {
      cur = /DIRECT/i.test(A) ? 'ALL' : normHub(A);
      out[cur] = { cats: {} };
      continue;
    }
    if (!cur) continue;
    if (/ยอดพัสดุ/.test(E)) { out[cur].pTarget = toNum(F); }
    else if (/ชำรุด|asset/i.test(E)) { out[cur].aTarget = String(F).trim(); }
    else if (/รีไซเคิล|recycle|ขายขยะ/i.test(E)) { out[cur].rTarget = toNum(F); }
    else if (/ประเมินการทำงาน|performance/i.test(E)) { out[cur].fTarget = String(F).trim(); }
    else if (E && !/COST KPI|คาดการณ์|ต้นทุนพัสดุ/i.test(E)) {
      var c = catOf(E);
      if (c && !out[cur].cats[c]) out[cur].cats[c] = { rc: toNum(F), baht: toNum((v[r + 1] || [])[5]) };
    }
  }
  return out;
}

// หา block ของฮับในแผ่น forecast (start..end แถว) จากโค้ดฮับในคอลัมน์ A
function fcBlock(v, hubKey) {
  var starts = [];
  for (var r = 0; r < v.length; r++) {
    var A = String(v[r][0] || '').trim();
    if (A && (/\d{2}\s*[A-Z]{2,}/.test(A) || /DIRECT/i.test(A))) starts.push({ r: r, key: /DIRECT/i.test(A) ? 'ALL' : normHub(A) });
  }
  for (var i = 0; i < starts.length; i++) {
    if (starts[i].key === hubKey) return { start: starts[i].r, end: (i + 1 < starts.length) ? starts[i + 1].r : v.length };
  }
  return null;
}
// เขียนค่ากลับไฟล์ Forecast: เฉพาะช่องกรอกข้อมูล (ยอดพัสดุ / บาทรายหมวด / asset / recycle / perf) — ไม่แตะช่องสูตร
// body: { m, hub, target:'p'|'cost'|'a'|'r'|'f', cat, week(0-4), value }
function writeForecast(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var sh = SpreadsheetApp.openById(FORECAST_ID).getSheetByName('Forcast KPI M' + b.m);
    if (!sh) return { ok: false, error: 'no sheet M' + b.m };
    var w = Number(b.week); if (!(w >= 0 && w <= 4)) return { ok: false, error: 'bad week' };
    var v = sh.getDataRange().getValues();
    var blk = fcBlock(v, normHub(String(b.hub || ''))); if (!blk) return { ok: false, error: 'hub not found' };
    function catName(e, key) {
      if (key === 'CON') return /วัสดุ/.test(e);
      if (key === 'RENT') return /เช่า/.test(e);
      if (key === 'ELEC') return /น้ำ|ไฟ|ประปา/.test(e);
      if (key === 'EXP') return /ค่าใช้จ่าย|ทั่วไป/.test(e);
      return false;
    }
    var targetRow = -1;
    for (var r = blk.start; r < blk.end; r++) {
      var E = String(v[r][4] || '').replace(/\n/g, ' ');
      if (b.target === 'p' && /ยอดพัสดุ/.test(E)) { targetRow = r; break; }
      else if (b.target === 'a' && /ชำรุด|asset/i.test(E)) { targetRow = r; break; }
      else if (b.target === 'r' && (/รีไซเคิล|recycle|ขายขยะ/i.test(E))) { targetRow = r; break; }
      else if (b.target === 'f' && /ประเมินการทำงาน|performance/i.test(E)) { targetRow = r; break; }
      else if (b.target === 'cost' && !/COST KPI|คาดการณ์|ต้นทุนพัสดุ/i.test(E) && catName(E, String(b.cat || ''))) { targetRow = r + 1; break; } // baht row = ratio row + 1
    }
    if (targetRow < 0) return { ok: false, error: 'target row not found' };
    sh.getRange(targetRow + 1, 7 + w).setValue(toNum(b.value)); // col G..K = 7..11
    return { ok: true, row: targetRow + 1, col: 7 + w };
  } finally { lock.releaseLock(); }
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

// debug/tool: อ่านสูตรของชีตใด ๆ (ตามสเปรดชีต id + ชื่อชีต) เพื่อวิเคราะห์/แก้สูตร
function readFormulas(id, sheet, r, c) {
  var sh = SpreadsheetApp.openById(id).getSheetByName(sheet);
  if (!sh) return null;
  var rows = Math.min(Number(r) || 80, sh.getLastRow() || 1);
  var cols = Math.min(Number(c) || 14, sh.getLastColumn() || 1);
  return sh.getRange(1, 1, rows, cols).getFormulas();
}

/* ====================== USERS / AUTH ====================== */
// แผ่น "ผู้ใช้งาน" คอลัมน์: id | role | hub | name | password | firstLogin (สร้าง+seed อัตโนมัติถ้ายังไม่มี)
function getUsersSheet() {
  var ss = SpreadsheetApp.openById(COST_ID);
  var sh = ss.getSheetByName(SH_USERS);
  if (sh) return sh;
  sh = ss.insertSheet(SH_USERS);
  sh.appendRow(['id', 'role', 'hub', 'name', 'password', 'firstLogin']);
  var seed = [
    ['bestdiarea', 'admin',  'ALL',       'ผู้ดูแลระบบ (Direct Area)', DEFAULT_PW, true],
    ['21BPL',      'member', 'HUB21BPL',  'สมาชิก HUB21BPL',          DEFAULT_PW, true],
    ['22PDT',      'member', 'HUB22PDT',  'สมาชิก HUB22PDT',          DEFAULT_PW, true],
    ['23AYU',      'member', 'HUB23AYU',  'สมาชิก HUB23AYU',          DEFAULT_PW, true],
    ['65WNO',      'member', 'HUB65WNO',  'สมาชิก HUB65WNO',          DEFAULT_PW, true],
    ['99BAG',      'member', 'HUB99BAG',  'สมาชิก HUB99BAG',          DEFAULT_PW, true]
  ];
  seed.forEach(function(r){ sh.appendRow(r); });
  return sh;
}
function usersData() {
  var sh = getUsersSheet();
  var v = sh.getDataRange().getValues();
  var head = v[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var ci = { id:head.indexOf('id'), role:head.indexOf('role'), hub:head.indexOf('hub'), name:head.indexOf('name'), pw:head.indexOf('password'), first:head.indexOf('firstlogin') };
  return { sh:sh, v:v, ci:ci };
}
function findUserRow(d, id) {
  for (var r = 1; r < d.v.length; r++) {
    if (String(d.v[r][d.ci.id]).trim().toLowerCase() === String(id).trim().toLowerCase()) return r;
  }
  return -1;
}
function userObj(row, ci) {
  return { id:String(row[ci.id]).trim(), role:String(row[ci.role]).trim(), hub:String(row[ci.hub]).trim(),
           name:String(row[ci.name]).trim(), firstLogin: row[ci.first] === true || String(row[ci.first]).toLowerCase() === 'true' };
}
// คืนรายชื่อผู้ใช้ (ไม่มี password) — สำหรับหน้าแอดมิน
function readUsers() {
  var d = usersData(); var out = [];
  for (var r = 1; r < d.v.length; r++) { if (String(d.v[r][d.ci.id]).trim()) out.push(userObj(d.v[r], d.ci)); }
  return out;
}
// ตรวจรหัสผ่าน → คืน user (ไม่มี password) ถ้าถูก
function loginUser(b) {
  var d = usersData(); var r = findUserRow(d, b.id || '');
  if (r < 0) return { ok:false, error:'ไม่พบบัญชีผู้ใช้นี้' };
  if (String(d.v[r][d.ci.pw]) !== String(b.pw || '')) return { ok:false, error:'รหัสผ่านไม่ถูกต้อง' };
  return { ok:true, user: userObj(d.v[r], d.ci) };
}
// ตั้งรหัสใหม่ (ครั้งแรก/เปลี่ยนเอง) — ตรวจรหัสเดิมก่อนถ้าส่งมา
function setUserPw(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var d = usersData(); var r = findUserRow(d, b.id || '');
    if (r < 0) return { ok:false, error:'ไม่พบบัญชีผู้ใช้นี้' };
    if (b.oldpw != null && String(d.v[r][d.ci.pw]) !== String(b.oldpw)) return { ok:false, error:'รหัสผ่านเดิมไม่ถูกต้อง' };
    if (String(b.newpw || '').length < 6) return { ok:false, error:'รหัสผ่านใหม่ต้องอย่างน้อย 6 ตัว' };
    d.sh.getRange(r + 1, d.ci.pw + 1).setValue(String(b.newpw));
    d.sh.getRange(r + 1, d.ci.first + 1).setValue(false);
    return { ok:true, user: userObj(d.v[r], d.ci) };
  } finally { lock.releaseLock(); }
}
// แอดมินรีเซ็ตรหัสกลับเป็นค่าเริ่มต้น + บังคับตั้งใหม่
function adminResetPw(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var d = usersData(); var r = findUserRow(d, b.id || '');
    if (r < 0) return { ok:false, error:'ไม่พบบัญชีผู้ใช้นี้' };
    d.sh.getRange(r + 1, d.ci.pw + 1).setValue(b.newpw ? String(b.newpw) : DEFAULT_PW);
    d.sh.getRange(r + 1, d.ci.first + 1).setValue(b.first === false ? false : true);
    return { ok:true };
  } finally { lock.releaseLock(); }
}
// เพิ่มผู้ใช้ใหม่ (รหัสเริ่มต้น + ต้องตั้งใหม่)
function addUser(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var d = usersData();
    if (findUserRow(d, b.id || '') >= 0) return { ok:false, error:'มีรหัสผู้ใช้นี้อยู่แล้ว' };
    if (!String(b.id || '').trim()) return { ok:false, error:'กรุณากรอก ID' };
    var row = []; row[d.ci.id] = String(b.id).trim(); row[d.ci.role] = b.role || 'member';
    row[d.ci.hub] = b.hub || ''; row[d.ci.name] = b.name || ''; row[d.ci.pw] = b.pw ? String(b.pw) : DEFAULT_PW; row[d.ci.first] = true;
    for (var i = 0; i < row.length; i++) if (row[i] == null) row[i] = '';
    d.sh.appendRow(row);
    return { ok:true };
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
