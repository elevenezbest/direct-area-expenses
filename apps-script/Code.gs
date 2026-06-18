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
var SH_TARGET = 'เป้าหมายปีนี้';    // เป้าหมายรายหมวด/รายฮับ (อาจยังไม่มี → คืน {})

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

// เป้าหมายปีนี้: อ่านแบบยืดหยุ่นตามหัวคอลัมน์ — คืน { HUBxxx:{pTarget,EXP,CON,RENT,ELEC,aTarget,rTarget,fTarget}, ... }
// ถ้าแผ่นยังไม่มี → คืน {} (หน้าเว็บจะใช้ค่า seed/เดิมต่อไป)
function readTargets() {
  var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_TARGET);
  if (!sh) return {};
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return {};
  var head = v[0].map(function(h){ return String(h).replace(/\n/g,'').trim().toLowerCase(); });
  function ci(names){ for (var i=0;i<head.length;i++){ for (var j=0;j<names.length;j++){ if (head[i]===names[j].toLowerCase()) return i; } } return -1; }
  var c = {
    hub:    ci(['hub','ชื่อhub','ฮับ']),
    p:      ci(['ptarget','parcel','ยอดพัสดุ','target']),
    EXP:    ci(['exp','expenses','ค่าใช้จ่ายทั่วไป','ค่าใช้จ่ายรายวัน']),
    CON:    ci(['con','consumables','ค่าวัสดุสิ้นเปลือง','วัสดุสิ้นเปลือง']),
    RENT:   ci(['rent','rental','ค่าเช่าอุปกรณ์','ค่าเช่า']),
    ELEC:   ci(['elec','electricity','ค่าไฟฟ้า-น้ำประปา','ค่าน้ำ-ไฟฟ้า','ค่าน้ำ-ค่าไฟฟ้า']),
    a:      ci(['atarget','asset','การจัดการทรัพย์สิน']),
    r:      ci(['rtarget','recycle','ขยะรีไซเคิล']),
    f:      ci(['ftarget','perf','performance','ประสิทธิภาพแอดมิน'])
  };
  if (c.hub < 0) return {};
  var out = {};
  for (var r = 1; r < v.length; r++) {
    var hub = String(v[r][c.hub] || '').trim();
    if (!hub) continue;
    var key = normHub(hub);
    var o = {};
    if (c.p   >= 0) o.pTarget = toNum(v[r][c.p]);
    if (c.EXP >= 0) o.EXP     = toNum(v[r][c.EXP]);
    if (c.CON >= 0) o.CON     = toNum(v[r][c.CON]);
    if (c.RENT>= 0) o.RENT    = toNum(v[r][c.RENT]);
    if (c.ELEC>= 0) o.ELEC    = toNum(v[r][c.ELEC]);
    if (c.a   >= 0) o.aTarget = String(v[r][c.a]).trim();
    if (c.r   >= 0) o.rTarget = toNum(v[r][c.r]);
    if (c.f   >= 0) o.fTarget = String(v[r][c.f]).trim();
    out[key] = o;
  }
  return out;
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
