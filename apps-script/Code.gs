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
var SH_USERS  = 'ผู้ใช้งาน';        // บัญชีแอดมิน (เก็บไว้)
var SH_PERSON = 'รายบุคคลใช้งาน';   // รายชื่อสมาชิกรายคน (รหัสพนักงาน|ชื่อ-นามสกุล|ชื่อเล่น|แผนก|ตำแหน่ง|สาขา|ID|password)
var SH_AVATAR = 'รูปโปรไฟล์';        // รูปโปรไฟล์รายคน (id|dataURL) — สร้างอัตโนมัติ
var DEFAULT_PW = '123456';         // รหัสเข้าครั้งแรกของทุกคน → ต้องตั้งใหม่หลังเข้าครั้งแรก

// แปลง "สาขา" (เช่น "21 BPL_BHUB-บางพลี") → รหัสฮับของเว็บ
var BRANCH_TO_HUB = {'21':'HUB21BPL','22':'HUB22PDT','23':'HUB23AYU','65':'HUB65WNO','99':'HUB99BAG'};
var HUB_TO_BRANCH = {'HUB21BPL':'21 BPL_BHUB-บางพลี','HUB22PDT':'22 PDT_BHUB-บ่อวิน','HUB23AYU':'23 AYU_BHUB-วังน้อย','HUB65WNO':'65 WNO_BHUB-วังน้อย','HUB99BAG':'99 BAG2_HUB-บางปลา'};
function hubFromBranch(s){ var m=String(s||'').match(/(\d+)/); return (m && BRANCH_TO_HUB[m[1]]) || ''; }

// ===== หัวคอลัมน์ของ "ข้อมูลค่าใช้จ่าย" (ตรงกับฟอร์มกรอก) =====
var EXP_HEADERS = ['วันที่','วันที่ทำเบิก','ชื่อHUB','หมายเลขอ้างอิงการเบิก OA',
                   'จำนวนเงิน','รายละเอียดค่าใช้จ่าย','ประเภทค่าใช้จ่าย','หลักฐาน','สถานะ','ผู้กรอก'];

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
    else if (action === 'avatars') out = { ok:true, avatars: readAvatars() };
    else if (action === 'loadkpi') out = { ok:true, key:p.key, store: readKpiStore(p.key) };
    else if (action === 'loadtasks')out = { ok:true, tasks: loadTasksStore() };
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
    else if (action === 'setavatar')    out = setAvatar(body);
    else if (action === 'writeForecast')out = writeForecast(body);
    else if (action === 'savekpi')      out = saveKpiStore(body);
    else if (action === 'addtask')      out = addTaskSrv(body);
    else if (action === 'updatetask')   out = updateTaskSrv(body);
    else if (action === 'cleardata')    out = clearData(body);
    else if (action === 'uploadEvidence') out = uploadEvidence(body);
    else if (action === 'sheetSet')     out = sheetSet(body);
    else if (action === 'sheetSetupWeekly') out = sheetSetupWeekly(body);
    else if (action === 'sheetSetWeek') out = sheetSetWeek(body);
    else if (action === 'msRead')       out = msRead(body);
    else                                out = { ok:false, error:'unknown action: ' + action };
  } catch (err) {
    out = { ok:false, error: String(err) };
  }
  return reply(out, (e && e.parameter && e.parameter.callback) || null);
}

/* ===== ข้อมูลรายเดือน (Sheet-backed): เขียนค่า 1 ช่อง (ฮับ×ปี×เดือน) ลงชีตภายนอกตาม sid =====
   หัวตาราง: HUB, Year, M1..M12 (แถวแรก) · แถว=ฮับ · หา row จาก (HUB, Year) แล้ว set คอลัมน์ M{month}
   ไม่เจอแถว → เพิ่มใหม่ · ต้องการให้บัญชีที่รัน backend มีสิทธิ์ "แก้ไข" ชีต sid นั้น */
function sheetSet(b) {
  if (!b || !b.sid) return { ok:false, error:'no sid' };
  var m = Number(b.month); if (!(m >= 1 && m <= 12)) return { ok:false, error:'bad month' };
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var ss = SpreadsheetApp.openById(b.sid); if (!ss) return { ok:false, error:'open failed' };
    var sh = ss.getSheets()[0];   // แผ่นแรก (ตารางข้อมูลรายเดือน)
    var data = sh.getDataRange().getValues(); if (!data.length) return { ok:false, error:'empty sheet' };
    var head = data[0].map(function(h){ return String(h).replace(/\s/g,'').toUpperCase(); });
    var iHub = -1, iYear = -1, iM = -1, want = 'M' + m;
    for (var c = 0; c < head.length; c++) {
      if (head[c] === 'HUB') iHub = c;
      else if (head[c] === 'YEAR') iYear = c;
      else if (head[c] === want) iM = c;
    }
    if (iHub < 0 || iM < 0) return { ok:false, error:'missing HUB or ' + want + ' column' };
    var hub = String(b.hub || '').trim().toUpperCase();
    var year = String(b.year || '').trim();
    var val = (b.value === '' || b.value == null) ? '' : Number(b.value);
    for (var r = 1; r < data.length; r++) {
      var rh = String(data[r][iHub]).trim().toUpperCase();
      var ry = (iYear >= 0) ? String(data[r][iYear]).trim() : year;
      if (rh === hub && ry === year) {
        sh.getRange(r + 1, iM + 1).setValue(val);
        return { ok:true, row: r + 1, col: iM + 1 };
      }
    }
    // ไม่เจอแถว (ฮับ,ปี) → เพิ่มแถวใหม่
    var arr = new Array(head.length).fill('');
    arr[iHub] = b.hub; if (iYear >= 0) arr[iYear] = year; arr[iM] = val;
    sh.appendRow(arr);
    return { ok:true, appended:true, row: sh.getLastRow() };
  } catch (err) {
    return { ok:false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* ===== ตั้งค่ารายสัปดาห์ (ครั้งเดียว): สร้างแท็บ "รายสัปดาห์" (HUB,Year,Month,W1-W5) + ย้ายยอดเดิมรายเดือน→W1 =====
   รายสัปดาห์เป็นตัวจริง · ยอดรายเดือน = ผลรวม W1-5 · เรียกซ้ำได้ (ไม่ migrate ทับของที่มีแล้ว) */
function sheetSetupWeekly(b) {
  if (!b || !b.sid) return { ok:false, error:'no sid' };
  var lock = LockService.getScriptLock(); lock.tryLock(20000);
  try {
    var ss = SpreadsheetApp.openById(b.sid);
    var mon = ss.getSheets()[0];                 // แท็บรายเดือน (แผ่นแรก · แผ่นรายงานเดิม — อ่านอย่างเดียว ไม่แก้)
    var md = mon.getDataRange().getValues(); if (md.length < 2) return { ok:false, error:'monthly empty' };
    var mh = md[0].map(function(h){ return String(h).replace(/\s/g,'').toUpperCase(); });
    var iHub = mh.indexOf('HUB'), iYear = mh.indexOf('YEAR');
    if (iHub < 0) return { ok:false, error:'monthly: no HUB column' };
    // คอลัมน์เดือน: M1-M12 หรือ ชื่อเดือนอังกฤษ (JANUARY/JAN…)
    var ABBR = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    var iMon = {}, monCount = 0;
    for (var m = 1; m <= 12; m++) { var c = mh.indexOf('M'+m); if (c < 0) { for (var k = 0; k < mh.length; k++) { if (mh[k].indexOf(ABBR[m-1]) === 0) { c = k; break; } } } if (c >= 0) { iMon[m] = c; monCount++; } }
    if (monCount === 0) return { ok:false, error:'monthly: no month columns (M1-12 or JANUARY…) found' };   // กันสร้างแท็บเปล่า (เช่น backend ยังไม่อัปเดต)
    var fbY = String(b.year||'').trim();   // ปีสำรอง เมื่อชีตไม่มีคอลัมน์ YEAR (เช่น ขยะ)
    var WK = 'รายสัปดาห์';
    var wk = ss.getSheetByName(WK); if (!wk) wk = ss.insertSheet(WK);
    wk.getRange(1, 1, 1, 8).setValues([['HUB','Year','Month','W1','W2','W3','W4','W5']]);
    var wdata = wk.getDataRange().getValues(); var seen = {};
    for (var i = 1; i < wdata.length; i++) seen[String(wdata[i][0]).trim().toUpperCase()+'|'+String(wdata[i][1]).trim()+'|'+String(wdata[i][2]).trim()] = true;
    var add = [];
    for (var r = 1; r < md.length; r++) {
      var hub = String(md[r][iHub]||'').trim(); if (!hub || /^(AREA|TOTAL|รวม)$/i.test(hub)) continue;
      var year = (iYear >= 0 && String(md[r][iYear]||'').trim()) ? String(md[r][iYear]).trim() : fbY;
      for (var mm = 1; mm <= 12; mm++) { if (iMon[mm] == null) continue;
        var raw = md[r][iMon[mm]]; var v = (raw === '' || raw == null) ? 0 : (Number(String(raw).replace(/[^0-9.\-]/g,'')) || 0);
        if (v <= 0) continue;
        var key = hub.toUpperCase()+'|'+year+'|'+mm; if (seen[key]) continue;
        add.push([hub, year, mm, v, 0, 0, 0, 0]);   // ยอดเดิม → W1
        seen[key] = true;
      }
    }
    if (add.length) wk.getRange(wk.getLastRow()+1, 1, add.length, 8).setValues(add);
    return { ok:true, added: add.length };
  } catch (err) { return { ok:false, error: String(err) }; } finally { lock.releaseLock(); }
}

/* ===== เขียนค่ารายสัปดาห์ 1 ช่อง (ฮับ×ปี×เดือน×สัปดาห์) → แท็บ "รายสัปดาห์" + อัปเดตแท็บรายเดือน M{n}=ผลรวม W1-5 ===== */
function sheetSetWeek(b) {
  if (!b || !b.sid) return { ok:false, error:'no sid' };
  var w = Number(b.week), m = Number(b.month);
  if (!(w >= 1 && w <= 5)) return { ok:false, error:'bad week' };
  if (!(m >= 1 && m <= 12)) return { ok:false, error:'bad month' };
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var ss = SpreadsheetApp.openById(b.sid);
    var wk = ss.getSheetByName('รายสัปดาห์'); if (!wk) return { ok:false, error:'no weekly tab (run setup first)' };
    var data = wk.getDataRange().getValues();
    var head = data[0].map(function(h){ return String(h).replace(/\s/g,'').toUpperCase(); });
    var iHub = head.indexOf('HUB'), iYear = head.indexOf('YEAR'), iMonth = head.indexOf('MONTH');
    var iW = []; for (var k = 1; k <= 5; k++) iW.push(head.indexOf('W'+k));
    if (iHub < 0 || iMonth < 0 || iW[w-1] < 0) return { ok:false, error:'weekly: missing columns' };
    var hub = String(b.hub||'').trim().toUpperCase(), year = String(b.year||'').trim();
    var val = (b.value === '' || b.value == null) ? 0 : Number(b.value);
    var rowNum = -1, weeks = [0,0,0,0,0];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iHub]).trim().toUpperCase() === hub && (iYear < 0 || String(data[r][iYear]).trim() === year) && String(data[r][iMonth]).trim() === String(m)) {
        rowNum = r + 1; for (var k2 = 0; k2 < 5; k2++) weeks[k2] = Number(data[r][iW[k2]]) || 0; break;
      }
    }
    if (rowNum < 0) {
      var arr = new Array(head.length).fill(''); arr[iHub] = b.hub; if (iYear >= 0) arr[iYear] = year; arr[iMonth] = m;
      for (var k3 = 0; k3 < 5; k3++) arr[iW[k3]] = (k3 === w-1) ? val : 0;
      wk.appendRow(arr); weeks[w-1] = val; rowNum = wk.getLastRow();
    } else {
      wk.getRange(rowNum, iW[w-1]+1).setValue(val); weeks[w-1] = val;
    }
    var monTot = weeks.reduce(function(a,x){ return a + (Number(x)||0); }, 0);
    try {   // อัปเดตแท็บรายเดือน M{m} = ผลรวมสัปดาห์ (ให้สรุปในชีตถูกเสมอ)
      var mon = ss.getSheets()[0]; var md = mon.getDataRange().getValues();
      var mh = md[0].map(function(h){ return String(h).replace(/\s/g,'').toUpperCase(); });
      var mHub = mh.indexOf('HUB'), mYear = mh.indexOf('YEAR'), mCol = mh.indexOf('M'+m);
      if (mHub >= 0 && mCol >= 0) for (var r2 = 1; r2 < md.length; r2++) if (String(md[r2][mHub]).trim().toUpperCase() === hub && (mYear < 0 || String(md[r2][mYear]).trim() === year)) { mon.getRange(r2+1, mCol+1).setValue(monTot); break; }
    } catch (_e) {}
    return { ok:true, row: rowNum, monthTotal: monTot };
  } catch (err) { return { ok:false, error: String(err) }; } finally { lock.releaseLock(); }
}

/* ===== อ่านข้อมูลรายเดือน "หลายแท็บ": แท็บชื่อเลขปี = อ่านอย่างเดียว (รายเดือน) · แท็บ "รายสัปดาห์" = แก้ได้ (W1-W5)
   คืน rows: {hub,year,month,w:[5],ro} · ลำดับความสำคัญ (ชนกัน hub|year|month): รายสัปดาห์ > แท็บเลขปี > แท็บรายงานเริ่มต้น ===== */
function msRead(b) {
  if (!b || !b.sid) return { ok:false, error:'no sid' };
  try {
    var ss = SpreadsheetApp.openById(b.sid);
    var sheets = ss.getSheets();
    var WK = 'รายสัปดาห์';
    var ABBR = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    var fbY = String(b.fbYear||'').trim();
    var num = function(x){ return (x===''||x==null) ? 0 : (Number(String(x).replace(/[^0-9.\-]/g,''))||0); };
    var buckets = { 0:[], 1:[], 2:[] };   // 0=weekly · 1=year-tab · 2=default report
    var hasWeekly = false;
    sheets.forEach(function(sh) {
      var name = String(sh.getName()).trim();
      var data = sh.getDataRange().getValues(); if (data.length < 2) return;
      var up = data[0].map(function(h){ return String(h).replace(/\s/g,'').toUpperCase(); });
      var iHub = up.indexOf('HUB'); if (iHub < 0) return;
      var iYear = up.indexOf('YEAR'), iMonth = up.indexOf('MONTH');
      var iW = []; for (var k=1;k<=5;k++) iW.push(up.indexOf('W'+k));
      if (name === WK && iMonth >= 0 && iW[0] >= 0) {           // แท็บรายสัปดาห์ (แก้ได้)
        hasWeekly = true;
        for (var r=1;r<data.length;r++) { var hub=String(data[r][iHub]||'').trim(); if(!hub) continue;
          var yr=(iYear>=0 && String(data[r][iYear]||'').trim())?String(data[r][iYear]).trim():fbY;
          var mo=parseInt(String(data[r][iMonth]||'').replace(/[^0-9]/g,''),10); if(!(mo>=1&&mo<=12)) continue;
          buckets[0].push({hub:hub,year:yr,month:mo,w:iW.map(function(ci){return ci>=0?num(data[r][ci]):0;}),ro:false});
        }
      } else {                                                  // แท็บรายเดือน: เลขปี=ปีนั้น(อ่านอย่างเดียว) · ไม่ใช่=แท็บรายงานเริ่มต้น(ใช้ fbYear)
        var isYear = /^\d{4}$/.test(name); var tabYear = isYear ? name : fbY; var bucket = isYear ? 1 : 2;
        var iMon = {}, mc = 0;
        for (var m=1;m<=12;m++){ var c=up.indexOf('M'+m); if(c<0){ for(var kk=0;kk<up.length;kk++){ if(up[kk].indexOf(ABBR[m-1])===0){c=kk;break;} } } if(c>=0){ iMon[m]=c; mc++; } }
        if (mc === 0) return;
        for (var r2=1;r2<data.length;r2++){ var hub2=String(data[r2][iHub]||'').trim(); if(!hub2||/^(AREA|TOTAL|รวม)$/i.test(hub2)) continue;
          var yr2=(iYear>=0 && String(data[r2][iYear]||'').trim())?String(data[r2][iYear]).trim():tabYear;
          for (var mm=1;mm<=12;mm++){ if(iMon[mm]==null) continue; var v=num(data[r2][iMon[mm]]); if(v<=0) continue;
            buckets[bucket].push({hub:hub2,year:yr2,month:mm,w:[v,0,0,0,0],ro:true}); }
        }
      }
    });
    var seen = {}, out = [];
    [0,1,2].forEach(function(p){ buckets[p].forEach(function(row){ var key=row.hub.toUpperCase()+'|'+row.year+'|'+row.month; if(seen[key]) return; seen[key]=true; out.push(row); }); });
    return { ok:true, rows:out, hasWeekly:hasWeekly };
  } catch (err) { return { ok:false, error:String(err) }; }
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
      status: String(v[idx.status] || '').trim(),
      by:     String(v[idx.by] || '').trim()
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
    rowArr[idx.by]        = b.by || '';
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
    // ===== แก้ไขรายการจากปุ่ม "แก้ไข" ของแอดมิน (แก้ได้ทุกช่องที่จำเป็น) =====
    if (b.type) sh.getRange(rowNum, idx.type+1).setValue(b.type);
    if (b.hub) sh.getRange(rowNum, idx.hub+1).setValue(b.hub);
    if (b.detail) sh.getRange(rowNum, idx.detail+1).setValue(b.detail);
    if (b.amount != null && b.amount !== '') sh.getRange(rowNum, idx.amount+1).setValue(toNum(b.amount));
    if (b.oaEdit) sh.getRange(rowNum, idx.oa+1).setValue(b.oaEdit);   // เปลี่ยน OA (ใช้ oaEdit แยกจาก b.oa ที่เป็นคีย์ค้นแถว)
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
// ---- แอดมิน: แผ่น "ผู้ใช้งาน" (ใช้เฉพาะแถว role=admin; แถวฮับเก่าถูกเมิน) ----
function getUsersSheet() {
  var ss = SpreadsheetApp.openById(COST_ID);
  var sh = ss.getSheetByName(SH_USERS);
  if (sh) return sh;
  sh = ss.insertSheet(SH_USERS);
  sh.appendRow(['id', 'role', 'hub', 'name', 'password', 'firstLogin']);
  sh.appendRow(['bestdiarea', 'admin', 'ALL', 'ผู้ดูแลระบบ (Direct Area)', DEFAULT_PW, true]);
  return sh;
}
function usersData() {
  var sh = getUsersSheet(); var v = sh.getDataRange().getValues();
  var head = v[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var ci = { id:head.indexOf('id'), role:head.indexOf('role'), hub:head.indexOf('hub'), name:head.indexOf('name'), pw:head.indexOf('password'), first:head.indexOf('firstlogin') };
  return { sh:sh, v:v, ci:ci };
}
function findRowById(d, idCol, id) {
  for (var r = 1; r < d.v.length; r++) if (String(d.v[r][idCol]).trim().toLowerCase() === String(id).trim().toLowerCase()) return r;
  return -1;
}
function isAdminRow(d, r) { return r >= 0 && String(d.v[r][d.ci.role]).trim().toLowerCase() === 'admin'; }
function adminObj(d, r, firstOverride) {
  var fl = (firstOverride != null) ? firstOverride : (d.v[r][d.ci.first] === true || String(d.v[r][d.ci.first]).toLowerCase() === 'true');
  return { id:String(d.v[r][d.ci.id]).trim(), role:'admin', hub:'ALL', name:String(d.v[r][d.ci.name]).trim(), firstLogin:fl };
}

// ---- สมาชิกรายคน: แผ่น "รายบุคคลใช้งาน" ----
function personData() {
  var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_PERSON);
  if (!sh) return null;
  var v = sh.getDataRange().getValues(); if (v.length < 1) return null;
  var head = v[0].map(function(h){ return String(h).replace(/\n/g,'').trim(); });
  function find(re){ for (var i=0;i<head.length;i++) if (re.test(head[i])) return i; return -1; }
  var ci = { code:find(/รหัสพนักงาน|employee|^รหัส/i), name:find(/นามสกุล|full/i), nick:find(/เล่น|nick/i), branch:find(/สาขา|branch|hub/i), pw:find(/password|รหัสผ่าน/i) };
  return { sh:sh, v:v, ci:ci, head:head };
}
function personPw(p, r){ return (p.ci.pw>=0 && String(p.v[r][p.ci.pw]).trim()) ? String(p.v[r][p.ci.pw]).trim() : DEFAULT_PW; }
function personUser(row, ci) {
  var code = String(row[ci.code]).trim();
  var pwCell = ci.pw>=0 ? String(row[ci.pw]).trim() : '';
  var nick = ci.nick>=0 ? String(row[ci.nick]).trim() : '';
  var full = ci.name>=0 ? String(row[ci.name]).trim() : '';
  return { id:code, role:'member', hub:hubFromBranch(row[ci.branch]),
           name: full || nick || code, nick:nick, fullName:full,
           branch: ci.branch>=0 ? String(row[ci.branch]).trim() : '',
           firstLogin: (pwCell==='' || pwCell===DEFAULT_PW) };
}
// รายชื่อสมาชิก (ไม่มีรหัสผ่าน) — สำหรับหน้าแอดมิน
function readUsers() {
  var out=[]; var p=personData();
  if (p && p.ci.code>=0) for (var r=1;r<p.v.length;r++) if (String(p.v[r][p.ci.code]).trim()) out.push(personUser(p.v[r], p.ci));
  return out;
}
// ล็อกอิน: แอดมิน(แผ่นผู้ใช้งาน) ก่อน แล้วสมาชิกรายคน(แผ่นรายบุคคล) — รหัสว่าง/123456 = ครั้งแรก
function loginUser(b) {
  var id=String(b.id||'').trim(), pw=String(b.pw||'');
  var d=usersData(); var ar=findRowById(d, d.ci.id, id);
  if (isAdminRow(d, ar)) {
    if (String(d.v[ar][d.ci.pw]) !== pw) return { ok:false, error:'รหัสผ่านไม่ถูกต้อง' };
    return { ok:true, user: adminObj(d, ar) };
  }
  var p=personData();
  if (p && p.ci.code>=0) { var r=findRowById(p, p.ci.code, id);
    if (r>=0) { if (pw !== personPw(p, r)) return { ok:false, error:'รหัสผ่านไม่ถูกต้อง' }; return { ok:true, user: personUser(p.v[r], p.ci) }; }
  }
  return { ok:false, error:'ไม่พบบัญชี — ใช้รหัสพนักงานเข้าระบบ' };
}
// ตั้งรหัสใหม่ (ครั้งแรก/เปลี่ยนเอง)
function setUserPw(b) {
  var lock=LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var id=String(b.id||'').trim(), newpw=String(b.newpw||'');
    if (newpw.length<6) return { ok:false, error:'รหัสผ่านใหม่ต้องอย่างน้อย 6 ตัว' };
    if (newpw===DEFAULT_PW) return { ok:false, error:'ห้ามใช้รหัสเริ่มต้น (123456) เป็นรหัสใหม่' };
    var d=usersData(); var ar=findRowById(d, d.ci.id, id);
    if (isAdminRow(d, ar)) {
      if (b.oldpw!=null && String(d.v[ar][d.ci.pw])!==String(b.oldpw)) return { ok:false, error:'รหัสผ่านเดิมไม่ถูกต้อง' };
      d.sh.getRange(ar+1, d.ci.pw+1).setValue(newpw); d.sh.getRange(ar+1, d.ci.first+1).setValue(false);
      return { ok:true, user: adminObj(d, ar, false) };
    }
    var p=personData();
    if (p && p.ci.code>=0) { var r=findRowById(p, p.ci.code, id);
      if (r>=0) {
        if (p.ci.pw<0) return { ok:false, error:'แผ่นรายบุคคลไม่มีคอลัมน์ password' };
        if (b.oldpw!=null && String(b.oldpw)!==personPw(p, r)) return { ok:false, error:'รหัสผ่านเดิมไม่ถูกต้อง' };
        p.sh.getRange(r+1, p.ci.pw+1).setValue(newpw);
        var u=personUser(p.v[r], p.ci); u.firstLogin=false; return { ok:true, user:u };
      }
    }
    return { ok:false, error:'ไม่พบบัญชีผู้ใช้นี้' };
  } finally { lock.releaseLock(); }
}
// แอดมินรีเซ็ตรหัส → กลับเป็นค่าเริ่มต้น (ครั้งแรกใหม่)
function adminResetPw(b) {
  var lock=LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var id=String(b.id||'').trim();
    var d=usersData(); var ar=findRowById(d, d.ci.id, id);
    if (isAdminRow(d, ar)) { d.sh.getRange(ar+1, d.ci.pw+1).setValue(DEFAULT_PW); d.sh.getRange(ar+1, d.ci.first+1).setValue(true); return { ok:true }; }
    var p=personData();
    if (p && p.ci.code>=0 && p.ci.pw>=0) { var r=findRowById(p, p.ci.code, id);
      if (r>=0) { p.sh.getRange(r+1, p.ci.pw+1).setValue(''); return { ok:true }; } // ว่าง = ใช้ 123456 ครั้งแรก
    }
    return { ok:false, error:'ไม่พบบัญชีผู้ใช้นี้' };
  } finally { lock.releaseLock(); }
}
// แอดมินเพิ่มสมาชิก → เขียนลงแผ่น "รายบุคคลใช้งาน" (รหัสว่าง = ใช้ 123456 ครั้งแรก)
function addUser(b) {
  var lock=LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var p=personData();
    if (!p || p.ci.code<0) return { ok:false, error:'ไม่พบแผ่น "รายบุคคลใช้งาน"' };
    var code=String(b.id||'').trim();
    if (!code) return { ok:false, error:'กรุณากรอกรหัสพนักงาน' };
    if (findRowById(p, p.ci.code, code)>=0) return { ok:false, error:'มีรหัสพนักงานนี้แล้ว' };
    var branch = HUB_TO_BRANCH[b.hub] || String(b.branch||'');
    var row=[]; row[p.ci.code]=code;
    if (p.ci.name>=0)   row[p.ci.name]=String(b.name||'');
    if (p.ci.nick>=0)   row[p.ci.nick]=String(b.nick||'');
    if (p.ci.branch>=0) row[p.ci.branch]=branch;
    if (p.ci.pw>=0)     row[p.ci.pw]='';
    for (var i=0;i<p.head.length;i++) if (row[i]==null) row[i]='';
    p.sh.appendRow(row);
    return { ok:true };
  } finally { lock.releaseLock(); }
}
// ---- รูปโปรไฟล์ (แผ่น "รูปโปรไฟล์": id | dataURL) ----
function getAvatarSheet() {
  var ss=SpreadsheetApp.openById(COST_ID); var sh=ss.getSheetByName(SH_AVATAR);
  if (!sh) { sh=ss.insertSheet(SH_AVATAR); sh.appendRow(['id','dataURL']); }
  return sh;
}
function readAvatars() {
  var sh=getAvatarSheet(); var v=sh.getDataRange().getValues(); var out={};
  for (var r=1;r<v.length;r++) { var id=String(v[r][0]).trim(); if (id) out[id]=String(v[r][1]); }
  return out;
}
function setAvatar(b) {
  var lock=LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var id=String(b.id||'').trim(); if (!id) return { ok:false, error:'no id' };
    var sh=getAvatarSheet(); var v=sh.getDataRange().getValues();
    for (var r=1;r<v.length;r++) if (String(v[r][0]).trim()===id) { sh.getRange(r+1,2).setValue(String(b.dataURL||'')); return { ok:true }; }
    sh.appendRow([id, String(b.dataURL||'')]); return { ok:true };
  } finally { lock.releaseLock(); }
}

/* ============ KPI STORE — snapshot ตาราง KPI ส่วนกลาง (ออนไลน์ ทุกคนเห็นตรงกัน) ============ */
// เก็บตาราง KPI ที่แอดมินกรอกเป็น JSON รายเดือน ในแผ่นซ่อน "WEB_KPI" ของไฟล์ "ค่าใช้จ่าย"
// — ไม่แตะไฟล์ Forecast (สูตร) เลย; ทุกเครื่องอ่าน snapshot ชุดเดียวกัน → ค่าตรงกันทั้งหมด
var SH_KPISTORE = 'WEB_KPI';
function getKpiStoreSheet() {
  var ss = SpreadsheetApp.openById(COST_ID);
  var sh = ss.getSheetByName(SH_KPISTORE);
  if (!sh) { sh = ss.insertSheet(SH_KPISTORE); sh.appendRow(['key','json','updatedAt','updatedBy']); try { sh.hideSheet(); } catch (_e) {} }
  return sh;
}
function readKpiStore(key) {
  key = String(key || '').trim(); if (!key) return null;
  var sh = getKpiStoreSheet(); var v = sh.getDataRange().getValues();
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][0]).trim() === key) {
      var js = String(v[r][1] || ''); if (!js) return null;
      return { json: js, updatedAt: String(v[r][2] || ''), updatedBy: String(v[r][3] || '') };
    }
  }
  return null;
}
function saveKpiStore(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var key = String(b.key || '').trim(); if (!key) return { ok:false, error:'no key' };
    var json = String(b.json || ''); if (!json) return { ok:false, error:'no json' };
    if (json.length > 49000) return { ok:false, error:'snapshot too large (' + json.length + ')' };
    var who = String(b.user || '').slice(0, 60);
    var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
    var sh = getKpiStoreSheet(); var v = sh.getDataRange().getValues();
    for (var r = 1; r < v.length; r++) {
      if (String(v[r][0]).trim() === key) {
        sh.getRange(r + 1, 2, 1, 3).setValues([[json, when, who]]);
        return { ok:true, key:key, updatedAt:when, updatedBy:who, row:r + 1 };
      }
    }
    sh.appendRow([key, json, when, who]);
    return { ok:true, key:key, updatedAt:when, updatedBy:who, row:sh.getLastRow() };
  } finally { lock.releaseLock(); }
}

/* ============ TASK STORE — งาน "ติดตามงาน" ส่วนกลาง (ออนไลน์ ทุกคนเห็น/เด้งเตือนตรงกัน) ============ */
// เก็บงานทั้งหมดเป็น JSON array ในแผ่น WEB_KPI คีย์ 'TASKS' — เขียนแบบ lock+read+modify กันชนกันเวลาหลายคนทำพร้อมกัน
function loadTasksStore() {
  var s = readKpiStore('TASKS'); if (!s || !s.json) return [];
  try { var a = JSON.parse(s.json); return Object.prototype.toString.call(a) === '[object Array]' ? a : []; } catch (_e) { return []; }
}
function writeTasksStore_(arr) {   // ภายใน: ต้องถือ lock อยู่แล้ว
  var sh = getKpiStoreSheet(); var v = sh.getDataRange().getValues();
  var json = JSON.stringify(arr);
  var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][0]).trim() === 'TASKS') { sh.getRange(r + 1, 2, 1, 3).setValues([[json, when, '']]); return; }
  }
  sh.appendRow(['TASKS', json, when, '']);
}
function addTaskSrv(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var arr = loadTasksStore();
    var t = b.task || {};
    var maxId = 0; for (var i = 0; i < arr.length; i++) { var n = Number(arr[i].id); if (n > maxId) maxId = n; }
    t.id = maxId + 1;
    if (!t.createdAt) t.createdAt = (new Date()).getTime();
    arr.unshift(t);
    writeTasksStore_(arr);
    return { ok:true, id:t.id, tasks:arr };
  } finally { lock.releaseLock(); }
}
function updateTaskSrv(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var arr = loadTasksStore();
    var id = String(b.id); var patch = b.patch || {}; var found = false;
    for (var i = 0; i < arr.length; i++) {
      if (String(arr[i].id) === id) { found = true; for (var k in patch) arr[i][k] = patch[k]; break; }
    }
    if (!found) return { ok:false, error:'task not found: ' + id };
    writeTasksStore_(arr);
    return { ok:true, tasks:arr };
  } finally { lock.releaseLock(); }
}

/* ============ CLEAR DATA — ล้างข้อมูลทดสอบก่อนเริ่มใช้จริง (ไม่แตะบัญชี/เป้า KPI) ============ */
// ลบเฉพาะ: log การเบิกทั้งหมด (เก็บหัวตาราง) + งานในส่วนกลาง (TASKS) + แถวทดสอบ TESTPING
// ต้องส่ง confirm:'ERASE' มาด้วย (กันยิงพลาด). คะแนน/อันดับคิดจาก log เบิก → พอ log ว่างก็รีเซ็ตเอง
function clearData(b) {
  if (String((b && b.confirm) || '') !== 'ERASE') return { ok:false, error:"ต้องส่ง confirm:'ERASE'" };
  var lock = LockService.getScriptLock(); lock.tryLock(8000);
  try {
    var out = { txnsCleared: 0, tasksCleared: false };
    var sh = SpreadsheetApp.openById(COST_ID).getSheetByName(SH_EXP);   // 1) log เบิก (เก็บแถวหัว)
    if (sh) { var last = sh.getLastRow(); if (last > 1) { sh.deleteRows(2, last - 1); out.txnsCleared = last - 1; } }
    var ks = getKpiStoreSheet(); var v = ks.getDataRange().getValues();  // 2) งาน TASKS + TESTPING ในแผ่น WEB_KPI
    for (var r = v.length - 1; r >= 1; r--) { var key = String(v[r][0]).trim(); if (key === 'TASKS' || key === 'TESTPING') ks.deleteRow(r + 1); }
    out.tasksCleared = true;
    return { ok:true, cleared: out };
  } finally { lock.releaseLock(); }
}

/* ============ EVIDENCE UPLOAD — อัปโหลดไฟล์แนบ (รูป/PDF) ขึ้น Google Drive แล้วคืนลิงก์ ============ */
// รับ dataURL (base64) + ชื่อไฟล์ → เซฟลงโฟลเดอร์ "DirectArea_Evidence" → แชร์ "ทุกคนที่มีลิงก์ดูได้" → คืน url
// body: { dataURL:'data:<mime>;base64,....', filename:'xxx.jpg' }
function getEvidenceFolder_() {
  var name = 'DirectArea_Evidence';
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}
function uploadEvidence(b) {
  var lock = LockService.getScriptLock(); lock.tryLock(20000);
  try {
    var dataUrl = String((b && b.dataURL) || '');
    var m = dataUrl.match(/^data:([^;]+);base64,([\s\S]*)$/);
    if (!m) return { ok:false, error:'bad dataURL' };
    var mime = m[1], b64 = m[2];
    var bytes = Utilities.base64Decode(b64);
    var name = String((b && b.filename) || 'evidence').replace(/[\\\/]/g,'_').slice(0, 120) || 'evidence';
    var blob = Utilities.newBlob(bytes, mime, name);
    var folder = getEvidenceFolder_();
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_e) {}
    return { ok:true, url: file.getUrl(), id: file.getId(), name: name };
  } catch (err) {
    return { ok:false, error: String(err) };
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
    status:    pick(find(headers[8]), 8),
    by:        pick(find(headers[9]), 9)
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
