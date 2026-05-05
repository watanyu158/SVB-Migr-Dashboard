// Data: SDA_Installation_Plan_V2.xlsx updated 2026-03-31 02:04
const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const https   = require('https');
const http    = require('http');

const app = express();
app.use(cors({origin:['https://svb-migr-progress.onrender.com','http://localhost:3000'],methods:['GET','POST']}));
app.use(express.json());

// ── Config ────────────────────────────────────────────────────────────────────
const SHAREPOINT_URL = process.env.SHAREPOINT_URL || '';
const LOCAL_EXCEL    = path.join(__dirname, 'SDA_Installation_Plan_V2.xlsx');
const CACHE_PATH     = path.join(__dirname, 'sda_cache.xlsx');
const GSHEET_ID      = '1DzFekcggT71Rq_oU4cwLe9GZlVe0UUT6oz9fg7GlPbE';
const CACHE_TTL      = 5 * 60 * 1000; // 5 min

let TOTAL = 1592, TOTAL_SW = 1121, TOTAL_AP = 445, TOTAL_INF = 26;
const PROJ_START = new Date('2026-02-09');
const PROJ_END   = new Date('2026-06-23');
const FABRICS    = ['D1-041','CFZ','T1-015','D1-091','RFF','AMF','PPW'];
const FAB_COLORS = { 'D1-041':'#4361ee','CFZ':'#2bc48a','T1-015':'#ff9f43',
                     'D1-091':'#a855f7','RFF':'#22b8cf','AMF':'#f76707','PPW':'#74c0fc' };
const FAB_PLAN_T = { 'D1-041':710,'CFZ':187,'T1-015':327,'D1-091':147,'RFF':99,'AMF':114,'PPW':8 };

// Week bounds W7–W26 (20 weeks)
const WK_BOUNDS = [
  ['2026-02-09','2026-02-15','09 Feb','W7'],  ['2026-02-16','2026-02-22','16 Feb','W8'],
  ['2026-02-23','2026-03-01','23 Feb','W9'],  ['2026-03-02','2026-03-08','02 Mar','W10'],
  ['2026-03-09','2026-03-15','09 Mar','W11'], ['2026-03-16','2026-03-22','16 Mar','W12'],
  ['2026-03-23','2026-03-29','23 Mar','W13'], ['2026-03-30','2026-04-05','30 Mar','W14'],
  ['2026-04-06','2026-04-12','06 Apr','W15'], ['2026-04-13','2026-04-19','13 Apr','W16'],
  ['2026-04-20','2026-04-26','20 Apr','W17'], ['2026-04-27','2026-05-03','27 Apr','W18'],
  ['2026-05-04','2026-05-10','04 May','W19'], ['2026-05-11','2026-05-17','11 May','W20'],
  ['2026-05-18','2026-05-24','18 May','W21'], ['2026-05-25','2026-05-31','25 May','W22'],
  ['2026-06-01','2026-06-08','01 Jun','W23'], ['2026-06-09','2026-06-15','09 Jun','W24'],
  ['2026-06-16','2026-06-22','16 Jun','W25'], ['2026-06-23','2026-06-29','23 Jun','W26'],
].map(([s,e,label,name]) => ({ s:new Date(s), e:new Date(e), label, name }));
const N_WK = WK_BOUNDS.length;

let cacheTime = 0, cachedData = null;

// ── File helpers ──────────────────────────────────────────────────────────────
function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    proto.get(url, { headers:{'User-Agent':'Mozilla/5.0'} }, res => {
      if ([301,302,303,307,308].includes(res.statusCode))
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode}`));
      const f = fs.createWriteStream(dest);
      res.pipe(f);
      f.on('finish', () => f.close(resolve));
      f.on('error', reject);
    }).on('error', reject);
  });
}

async function fetchCSV(sheetName) {
  const axios = require('axios');
  const url = `https://docs.google.com/spreadsheets/d/${GSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  console.log('Fetching sheet:', sheetName);
  const res = await axios.get(url, {
    timeout: 15000,
    maxRedirects: 10,
    headers: { 'User-Agent': 'Mozilla/5.0' },
    responseType: 'text'
  });
  return res.data;
}

function parseCSVRow(line) {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"' && !inQ) { inQ = true; continue; }
    if (ch === '"' && inQ) {
      if (line[i+1] === '"') { cur += '"'; i++; }
      else inQ = false;
      continue;
    }
    if (ch === ',' && !inQ) { result.push(cur); cur = ''; continue; }
    cur += ch;
  }
  result.push(cur);
  return result.map(v => {
    v = v.trim();
    if (v === '' || v === '-') return null;
    const n = Number(v);
    return isNaN(n) ? v : n;
  });
}

function csvToRows(csv) {
  return csv.split('\n').filter(l => l.trim()).map(parseCSVRow);
}

async function getWorkbook() {
  try {
    console.log('Fetching from Google Sheets...');
    const dashCSV   = await fetchCSV('Dashboard');
    const detailCSV = await fetchCSV('All_Detail');
    console.log('Google Sheets OK');
    return { _isGSheet:true, dash:csvToRows(dashCSV), detail:csvToRows(detailCSV) };
  } catch(e) {
    console.warn('Google Sheets failed:', e.message);
  }
  if (fs.existsSync(LOCAL_EXCEL)) {
    console.log('Using local Excel fallback');
    return XLSX.readFile(LOCAL_EXCEL);
  }
  throw new Error('No Excel source available');
}


// ── Date helper — handle XLSX serial, ISO string, Date object ─────────────────
function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') return new Date((v - 25569) * 86400000);
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    // D/M/YY หรือ DD/MM/YY หรือ D/M/YYYY (Google Sheets format)
    const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashMatch) {
      let [,d,m,y] = slashMatch;
      d = parseInt(d); m = parseInt(m); y = parseInt(y);
      if (y < 100) y += 2000; // YY -> YYYY
      return new Date(y, m-1, d);
    }
    // YYYY-MM-DD (ISO)
    const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      return new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2])-1, parseInt(isoMatch[3]));
    }
    return null;
  }
  return null;
}
function toDateStr(v) {
  const d = toDate(v);
  return d ? d.toISOString().slice(0,10) : null;
}

// ── Week index helper ─────────────────────────────────────────────────────────
function wkIdx(dt) {
  if (!dt) return -1;
  const d = new Date(dt); d.setHours(0,0,0,0);
  const idx = WK_BOUNDS.findIndex(w => d >= w.s && d <= w.e);
  if (idx >= 0) return idx;
  if (d < WK_BOUNDS[0].s) return 0;   // ก่อน proj_start → นับเข้า W1
  return -1;
}

function fmtDate(dt) {
  const d = toDate(dt);
  if (!d) return '';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
}

function cumPct(arr, total) {
  let c = 0;
  return arr.map(v => Math.round((c += v) / total * 10000) / 100);
}

function cumActNull(arr, total, upto) {
  let c = 0, last = null;
  return arr.map((v, i) => {
    if (v > 0) { c += v; last = Math.round(c / total * 10000) / 100; }
    return i <= upto ? last : null;
  });
}

// ── Main calculation ──────────────────────────────────────────────────────────
function calcDashboard(wb) {
  let dRows, aRows;
  if (wb._isGSheet) {
    // Google Sheets format
    dRows = wb.dash;
    // แปลง detail rows เป็น object format เหมือน XLSX
    const hdr = wb.detail[0] || [];
    aRows = wb.detail.slice(1).map(r => {
      const obj = {};
      hdr.forEach((h,i) => { if(h) obj[h] = r[i]; });
      return obj;
    });
  } else {
    // XLSX format (fallback)
    const wsD = wb.Sheets['Dashboard'];
    const wsA = wb.Sheets['All_Detail'];
    dRows = XLSX.utils.sheet_to_json(wsD, { header:1, defval:null });
    aRows = XLSX.utils.sheet_to_json(wsA, { defval:null });
  }

  // header scan — ไม่ hardcode row index
  let _totalRow = -1, _installedRow = -1, _swRow = -1, _apRow = -1;
  for (let i = 0; i < Math.min(dRows.length, 30); i++) {
    const r = dRows[i] || [];
    const r0 = String(r[0]||'');
    if (_totalRow < 0 && typeof r[0] === 'number' && r[0] > 100) _totalRow = i;
    if (_installedRow < 0 && r0.includes('Overall')) _installedRow = i;
    if (_swRow < 0 && r0.includes('SW') && typeof r[2] === 'number') _swRow = i;
    if (_apRow < 0 && r0.includes('AP') && typeof r[2] === 'number') _apRow = i;
  }
  if (_totalRow >= 0 && dRows[_totalRow][0] > 0) TOTAL = dRows[_totalRow][0];
  const installed     = _installedRow >= 0 ? (dRows[_installedRow][3] || 0) : 0;
  const INSTALLED_SW  = _swRow >= 0 ? (dRows[_swRow][2] || 0) : 0;
  const INSTALLED_AP  = _apRow >= 0 ? (dRows[_apRow][2] || 0) : 0;
  const INSTALLED_INF = installed - INSTALLED_SW - INSTALLED_AP;

  // hold = นับจำนวน rows ที่ Status='Hold' (ไม่ใช่ qty)
  const hold = aRows.filter(r => r['Status'] === 'Hold').length;

  // overdue = นับ rows ที่ Days Until Due < 0 และยังไม่เสร็จ (ok < qty)
  const overdue = aRows.filter(r => {
    const days = r['Days Until Due'];
    const qty  = r['Qty'] || 0;
    const ok   = r['Qty. Success'] || 0;
    return typeof days === 'number' && days < 0 && ok < qty;
  }).length;

  const remaining = TOTAL - installed;

  // TODAY & insight
  const today = new Date(); today.setHours(0,0,0,0);
  // elapsed: นับรวมวันแรก (proj_start) ด้วย → +1
  const elapsed   = Math.max(1, Math.floor((today - PROJ_START) / 86400000) + 1);
  // daysLeft: นับรวมวันนี้ด้วย → +1
  const daysLeft  = Math.max(1, Math.floor((PROJ_END - today) / 86400000) + 1);
  const dailyRate = Math.round(installed / elapsed * 10) / 10;
  // reqRate ปัดขึ้นเต็มจำนวน
  const reqRate   = Math.ceil(remaining / daysLeft);
  const needMore  = Math.round((reqRate - dailyRate) * 10) / 10;
  const pctMore   = dailyRate > 0 ? Math.round((reqRate / dailyRate - 1) * 100) : 0;
  const daysNeeded= dailyRate > 0 ? Math.ceil(remaining / dailyRate) : 9999;
  const finishDt  = new Date(today); finishDt.setDate(today.getDate() + daysNeeded);
  const daysLate  = Math.max(0, Math.floor((finishDt - PROJ_END) / 86400000));
  const daysEarly = Math.max(0, Math.floor((PROJ_END - finishDt) / 86400000));
  const gaugePct  = reqRate > 0 ? Math.min(150, Math.round(dailyRate / reqRate * 100)) : 100;
  const todayWk   = Math.max(0, Math.min(N_WK - 1, Math.floor((elapsed-1) / 7)));

  // Arrays
  const planWk=new Array(N_WK).fill(0); const actWk =new Array(N_WK).fill(0);
  const swPlan=new Array(N_WK).fill(0); const swAct =new Array(N_WK).fill(0);
  const apPlan=new Array(N_WK).fill(0); const apAct =new Array(N_WK).fill(0);
  const infAct=new Array(N_WK).fill(0);

  const fabSwPlan={}, fabApPlan={}, fabSwAct={}, fabApAct={};
  FABRICS.forEach(f => {
    fabSwPlan[f]=new Array(N_WK).fill(0); fabApPlan[f]=new Array(N_WK).fill(0);
    fabSwAct[f] =new Array(N_WK).fill(0); fabApAct[f] =new Array(N_WK).fill(0);
  });

  const dailyMap = {}; // dk → {sw,ap,inf,plan}
  const fabDailyAct={}, fabDailyPlan={}, dayFabSwAct={}, dayFabApAct={};
  FABRICS.forEach(f=>{ dayFabSwAct[f]={}; dayFabApAct[f]={}; });
  FABRICS.forEach(f => { fabDailyAct[f]={}; fabDailyPlan[f]={}; });

  const typeMap = {};
  const locMap  = {}; // fab → loc → {t,d}

  // upcoming 14 วัน
  const _td = new Date(today); _td.setHours(0,0,0,0);
  const _e14 = new Date(_td); _e14.setDate(_e14.getDate()+14);
  const todayStr = _td.toISOString().slice(0,10);
  const end14Str = _e14.toISOString().slice(0,10);
  const upcoming = {};
  // นับ Qty.Success ทั้งหมด (ไม่ require Install Date) — ตรงกับ Dashboard
  let totalSwOk=0, totalApOk=0, totalInfOk=0;
  // หา min/max scheduled date per fabric
  const fabSchedMin={}, fabSchedMax={};
  aRows.forEach(r => {
    const fab  = r['Fabric'];
    const cat  = r['Category'];
    const qty  = r['Qty']           || 0;
    const ok   = r['Qty. Success']  || 0;
    const dt   = r['Device Type'];
    const loc  = r['Location'];

    // Device types (all rows)
    if (dt) {
      if (!typeMap[dt]) typeMap[dt] = {plan:0, done:0};
      typeMap[dt].plan += qty;
      typeMap[dt].done += ok;
    }

    if (!FABRICS.includes(fab)) return;

    // Location map
    if (loc) {
      if (!locMap[fab]) locMap[fab] = {};
      if (!locMap[fab][loc]) locMap[fab][loc] = {t:0, d:0};
      locMap[fab][loc].t += qty;
      locMap[fab][loc].d += ok;
    }

    // upcoming 14 วัน + track min/max scheduled per fabric
    let schedDt = toDate(r['Scheduled Date']);
    if (schedDt && qty > 0) {
      const _sd = schedDt;
      const _sds = _sd instanceof Date ? _sd.toISOString().slice(0,10) : '';
      if (_sds >= todayStr && _sds <= end14Str) {
        if (!upcoming[_sds]) upcoming[_sds] = {};
        if (!upcoming[_sds][fab]) upcoming[_sds][fab] = {qty:0,rem:0,locs:new Set(),types:new Set(),cats:new Set()};
        upcoming[_sds][fab].qty += qty;
        upcoming[_sds][fab].rem += (r['Qty. Remaining']||qty);
        if (r['Location']) upcoming[_sds][fab].locs.add(r['Location']);
        if (r['Device Type']) upcoming[_sds][fab].types.add(r['Device Type']);
        if (cat) upcoming[_sds][fab].cats.add(cat);
      }
    }
    if (schedDt && FABRICS.includes(fab)) {
      const dk=schedDt ? schedDt.getTime() : new Date(schedDt).getTime();
      if(!fabSchedMin[fab]||dk<fabSchedMin[fab]) fabSchedMin[fab]=dk;
      if(!fabSchedMax[fab]||dk>fabSchedMax[fab]) fabSchedMax[fab]=dk;
    }
    if (schedDt && qty > 0) {
      const wi = wkIdx(schedDt); const dk = fmtDate(schedDt);
      if (wi >= 0) {
        planWk[wi] += qty;
        if (cat === 'Switch') { swPlan[wi]+=qty; fabSwPlan[fab][wi]+=qty; }
        else if (cat === 'AP') { apPlan[wi]+=qty; fabApPlan[fab][wi]+=qty; }
      }
      if (!dailyMap[dk]) dailyMap[dk]={sw:0,ap:0,inf:0,plan:0};
      dailyMap[dk].plan += qty;
      fabDailyPlan[fab][dk] = (fabDailyPlan[fab][dk]||0) + qty;
    }

    // Install date → actual (daily timeline ต้องมี Install Date)
    let instDt = toDate(r['Install Date']);
    if (!instDt && ok > 0) instDt = new Date();  // ติดตั้งแล้วแต่ไม่มีวันที่
    if (instDt && ok > 0) {
      const wi = wkIdx(instDt); const dk = fmtDate(instDt);
      if (wi >= 0) {
        actWk[wi] += ok;
        if (cat === 'Switch') { swAct[wi]+=ok; fabSwAct[fab][wi]+=ok; }
        else if (cat === 'AP') { apAct[wi]+=ok; fabApAct[fab][wi]+=ok; }
        else infAct[wi] += ok;
      }
      if (!dailyMap[dk]) dailyMap[dk]={sw:0,ap:0,inf:0,plan:0};
      if (cat === 'Switch') dailyMap[dk].sw += ok;
      else if (cat === 'AP') dailyMap[dk].ap += ok;
      else dailyMap[dk].inf += ok;
      if (!fabDailyAct[fab][dk]) fabDailyAct[fab][dk]={sw:0,ap:0,inf:0};
      const _isoK = instDt ? instDt.toISOString().slice(0,10) : null;
      if (_isoK) {
        if (cat==='Switch') { fabDailyAct[fab][dk].sw+=ok; dayFabSwAct[fab][_isoK]=(dayFabSwAct[fab][_isoK]||0)+ok; }
        else if (cat==='AP') { fabDailyAct[fab][dk].ap+=ok; dayFabApAct[fab][_isoK]=(dayFabApAct[fab][_isoK]||0)+ok; }
        else fabDailyAct[fab][dk].inf+=ok;
      }
    }
    // DEBUG: ดู raw Install Date format


  // นับ ok ทุก row (เหมือน Dashboard) — ไม่ require Install Date
    if (ok > 0) {
      if (cat === 'Switch') totalSwOk += ok;
      else if (cat === 'AP') totalApOk += ok;
      else totalInfOk += ok;
    }
  });

  // Cumulative weekly %
  const PLAN_ALL = cumPct(planWk, TOTAL);
  const ACT_ALL  = cumActNull(actWk, TOTAL, todayWk);
  const PLAN_SW  = cumPct(swPlan, TOTAL_SW);
  const ACT_SW   = cumActNull(swAct, TOTAL_SW, todayWk);
  const PLAN_AP  = cumPct(apPlan, TOTAL_AP);
  const ACT_AP   = cumActNull(apAct, TOTAL_AP, todayWk);

  // Burndown — BD_PLAN นับ Scheduled Date ≤ min(w.e, เมื่อวาน GMT+7)
  // ใช้เมื่อวานเป็น cutoff เพื่อแสดง end-of-day ล่าสุดที่สมบูรณ์
  const _bdGMT7 = new Date(new Date().toLocaleString('en-US',{timeZone:'Asia/Bangkok'}));
  _bdGMT7.setHours(0,0,0,0);
  _bdGMT7.setDate(_bdGMT7.getDate() - 1);  // เมื่อวาน
  const BD_PLAN = WK_BOUNDS.map(w => {
    const cutoff = w.e < _bdGMT7 ? w.e : _bdGMT7;
    let cum = 0;
    aRows.forEach(r => {
      const sd = toDate(r['Scheduled Date']);
      if (sd && r['Qty'] > 0) {
        const d = new Date(sd); d.setHours(0,0,0,0);
        if (d <= cutoff) cum += (r['Qty'] || 0);
      }
    });
    return TOTAL - cum;
  });
  let s = 0; let last = null;
  const BD_ACT = actWk.map((v,i) => {
    if (v > 0) { s += v; last = TOTAL - s; }
    return i <= todayWk ? last : null;
  });

  // FAB_WEEKLY
  const fabWeekly = {};
  FABRICS.forEach(f => {
    const swT = fabSwPlan[f].reduce((a,v)=>a+v, 0);
    const apT = fabApPlan[f].reduce((a,v)=>a+v, 0);
    fabWeekly[f] = {
      sw_plan: cumPct(fabSwPlan[f], swT || 1),
      sw_act:  cumActNull(fabSwAct[f], swT || 1, todayWk),
      ap_plan: cumPct(fabApPlan[f], apT || 1),
      ap_act:  cumActNull(fabApAct[f], apT || 1, todayWk),
    };
  });

  // Daily sorted
  const sortedDates = Object.keys(dailyMap).sort((a,b) => {
    const [da,ma] = a.split('/').map(Number);
    const [db,mb] = b.split('/').map(Number);
    return ma !== mb ? ma - mb : da - db;
  });

  let cSW=0, cAP=0, cIN=0;
  const daily = {
    labels:    sortedDates,
    sw:        sortedDates.map(d => dailyMap[d].sw),
    ap:        sortedDates.map(d => dailyMap[d].ap),
    inf:       sortedDates.map(d => dailyMap[d].inf),
    plan:      sortedDates.map(d => dailyMap[d].plan),
    cum_sw:    sortedDates.map(d => (cSW += dailyMap[d].sw)),
    cum_ap:    sortedDates.map(d => (cAP += dailyMap[d].ap)),
    cum_inf:   sortedDates.map(d => (cIN += dailyMap[d].inf)),
  };
  daily.cum_d = daily.cum_sw.map((v,i) => v + daily.cum_ap[i] + daily.cum_inf[i]);

  // FAB_DAILY — แยก SW/AP/Inf
  const fabDaily = {}, fabDailyPlanOut = {};
  FABRICS.forEach(f => {
    fabDaily[f] = {
      sw:  sortedDates.map(d => (fabDailyAct[f][d]&&fabDailyAct[f][d].sw)  || 0),
      ap:  sortedDates.map(d => (fabDailyAct[f][d]&&fabDailyAct[f][d].ap)  || 0),
      inf: sortedDates.map(d => (fabDailyAct[f][d]&&fabDailyAct[f][d].inf) || 0),
    };
    fabDailyPlanOut[f] = sortedDates.map(d => fabDailyPlan[f][d] || 0);
  });

  // Fabrics from Dashboard rows 27-33
  // หา section rows โดย header — robust กว่า slice
  const FAB_NAMES = ['D1-041','CFZ','T1-015','D1-091','RFF','AMF','PPW'];
  // section 1: "Progress by Fabric" — cols: name,total,done,%,hold,rem,start,end
  const fabHdrIdx = dRows.findIndex(r => r && r[0]==='Fabric' && r[1]==='Total');
  const fabRows = fabHdrIdx>=0 ? dRows.slice(fabHdrIdx+1, fabHdrIdx+8) : [];
  // section 2: "Progress by Fabric - Switch vs AP" — cols: name,swT,swD,sw%,apT,apD,ap%,infT,infD,inf%
  const catHdrIdx = dRows.findIndex(r => r && r[0]==='Fabric' && r[1]==='SW Total');
  const fabCatRows = catHdrIdx>=0 ? dRows.slice(catHdrIdx+1, catHdrIdx+8) : [];
  const base = new Date(1899,11,30);
  const fmtExcelDate = v => {
    if(!v||typeof v!=='number') return '–';
    const d=new Date(base.getTime()+v*86400000);
    return `${d.getDate()}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  };
  const fmtExcelDate2 = ts => {
    if(!ts) return '–';
    const d=new Date(ts);
    return `${d.getDate()}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  };
  const fabrics = FABRICS.map(fn => {
    const fr  = fabRows.find(r => r[0] === fn) || [];
    const frc = fabCatRows.find(r => r[0] === fn) || [];
    const swT=frc[1]||0, swD=frc[2]||0, apT=frc[4]||0, apD=frc[5]||0, infT=frc[7]||0, infD=frc[8]||0;
    const tot=swT+apT+infT, done=swD+apD+infD;
    return {
      n:fn, t:tot, d:done, p:tot>0?Math.round(done/tot*10000)/100:0,
      h:fr[4]||0, r:tot-done, c:FAB_COLORS[fn],
      s:(fabSchedMin[fn]?fmtExcelDate2(fabSchedMin[fn]):'–'),
      e:(fabSchedMax[fn]?fmtExcelDate2(fabSchedMax[fn]):'–'),
      sw:{t:swT,d:swD}, ap:{t:apT,d:apD}, inf:{t:infT,d:infD},
      weekly: fabWeekly[fn],
    };
  });

  const fabTotals = {};
  fabrics.forEach(f => { fabTotals[f.n] = {sw:f.sw.d, ap:f.ap.d, inf:f.inf.d}; });

  // Types sorted by plan desc
  const types = Object.entries(typeMap)
    .map(([n,d]) => ({n, plan:d.plan, done:d.done}))
    .sort((a,b) => b.plan - a.plan);

  // Locations
  const locations = {};
  Object.entries(locMap).forEach(([fab, locs]) => {
    locations[fab] = Object.entries(locs)
      .map(([l,v]) => ({l, t:v.t, d:v.d, p:Math.round(v.d/Math.max(v.t,1)*100)}))
      .sort((a,b) => b.t - a.t);
  });

  // on-time = installed ที่ Install Date <= Scheduled Date
  let onTimeQty = 0, earlyQty = 0, lateQty = 0;
  // นับ Qty.Success ทั้งหมด (ไม่ require Install Date) — ตรงกับ Dashboard
  aRows.forEach(r => {
    const ok    = r['Qty. Success'] || 0;
    const instDt  = toDate(r['Install Date']);
    const schedDt = toDate(r['Scheduled Date']);
    if (ok > 0 && instDt && schedDt) {
      instDt.setHours(0,0,0,0); schedDt.setHours(0,0,0,0);
      if (instDt <= schedDt) {
        onTimeQty += ok;
        const diff = Math.floor((schedDt - instDt) / 86400000);
        if (diff > 0) earlyQty += ok; // ติดตั้งก่อน scheduled date
      } else {
        lateQty += ok;
      }
    }
  });
  const onTimePct = installed > 0 ? Math.round(onTimeQty / installed * 1000) / 10 : 0;

  // hold items list
  const holdItems = aRows
    .filter(r => r['Status']==='Hold')
    .map(r => ({
      fab:  r['Fabric']||'',
      loc:  r['Location']||'',
      dev:  r['Device Type']||'',
      qty:  r['Qty']||0,
      done: r['Qty. Success']||0,
      rem:  r['Qty. Remaining']||0,
    }));

  // Last install date
  let lastInstallDate = null;
  // นับ Qty.Success ทั้งหมด (ไม่ require Install Date) — ตรงกับ Dashboard
  aRows.forEach(r => {
    const d = toDate(r['Install Date']);
    if (d) {
      const ds = d.toISOString().slice(0,10);
      if (!lastInstallDate || ds > lastInstallDate) lastInstallDate = ds;
    }
  });

  // ── Daily cumulative progress ────────────────────────────────────────────────
  const PROJ_START_D = new Date('2026-02-09'); PROJ_START_D.setHours(0,0,0,0);
  const PROJ_END_D   = new Date('2026-06-23'); PROJ_END_D.setHours(0,0,0,0);

  // daily maps: all / sw / ap / per-fabric
  const dayActMap={}, dayPlanMap={};
  const daySwAct={}, daySwPlan={}, dayApAct={}, dayApPlan={};
  const dayFabAct={}, dayFabPlan={}, dayFabSwPlan={}, dayFabApPlan={};
  FABRICS.forEach(f=>{ dayFabAct[f]={}; dayFabPlan[f]={}; dayFabSwPlan[f]={}; dayFabApPlan[f]={}; });

  // นับ Qty.Success ทั้งหมด (ไม่ require Install Date) — ตรงกับ Dashboard
  aRows.forEach(r => {
    const fab   = r['Fabric'];
    const cat   = r['Category'];
    const qty   = r['Qty'] || 0;
    const ok    = r['Qty. Success'] || 0;
    const instDt  = toDate(r['Install Date']);
    const schedDt = toDate(r['Scheduled Date']);

    if (instDt && ok > 0) {
      const k = instDt.toISOString().slice(0,10);
      dayActMap[k] = (dayActMap[k]||0) + ok;
      if (cat==='Switch') daySwAct[k] = (daySwAct[k]||0) + ok;
      else if (cat==='AP') dayApAct[k] = (dayApAct[k]||0) + ok;
      if (FABRICS.includes(fab)) dayFabAct[fab][k] = (dayFabAct[fab][k]||0) + ok;
    }
    if (schedDt && qty > 0) {
      const k = schedDt.toISOString().slice(0,10);
      dayPlanMap[k] = (dayPlanMap[k]||0) + qty;
      if (cat==='Switch') daySwPlan[k] = (daySwPlan[k]||0) + qty;
      else if (cat==='AP') dayApPlan[k] = (dayApPlan[k]||0) + qty;
      if (FABRICS.includes(fab)){
        dayFabPlan[fab][k] = (dayFabPlan[fab][k]||0) + qty;
        if (cat==='Switch') dayFabSwPlan[fab][k] = (dayFabSwPlan[fab][k]||0) + qty;
        else if (cat==='AP') dayFabApPlan[fab][k] = (dayFabApPlan[fab][k]||0) + qty;
      }
    }
  });

  // sw_plan/ap_plan ใช้ sortedDates เดิม (actual+plan dates รวมกัน)
  // NOTE: sortedDates อาจไม่มี future plan dates → ใช้ projDates แทน
  // projDates ถูก build ใน dailyProgress loop ด้านล่าง — เพิ่มที่นั่นแทน


  // หา last install date per fabric
  const fabLastInstall = {};
  FABRICS.forEach(f => {
    const dates = Object.keys(dayFabAct[f]).sort();
    fabLastInstall[f] = dates.length ? dates[dates.length-1] : null;
  });

  // build daily cumulative arrays
  const lastActDt = lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : null;

  const dailyProgress = { labels:[], plan_cum:[], act_cum:[],
    sw_plan:[], sw_act:[], ap_plan:[], ap_act:[],
    bd_plan:[], bd_act:[], fab:{} };
  FABRICS.forEach(f => { dailyProgress.fab[f] = { plan:[], act:[],
    sw_plan:[], sw_act:[], ap_plan:[], ap_act:[], _spc:0, _apc:0, _swAc:0, _apAc:0 }; });

  let cAll=0, cPlan=0, cSWd=0, cSWp=0, cAPd=0, cAPp=0;
  const cFab={}, cFabP={};
  FABRICS.forEach(f=>{ cFab[f]=0; cFabP[f]=0; });

  // Burndown plan: ต้องนับ planned qty ทั้งหมดก่อนแล้วลด
  // pre-calc total plan by day
  const prePlanCum = {};
  let _pp=0;
  const _c2 = new Date(PROJ_START_D);
  while (_c2 <= PROJ_END_D) {
    const k = _c2.toISOString().slice(0,10);
    _pp += dayPlanMap[k]||0;
    prePlanCum[k] = _pp;
    _c2.setDate(_c2.getDate()+1);
  }

  // SW/AP plan total per fabric (จาก fabSwPlan weekly sum)
  const fabSwPlanTotal={}, fabApPlanTotal={};
  FABRICS.forEach(f => {
    fabSwPlanTotal[f] = Object.values(dayFabSwPlan[f]).reduce((a,v)=>a+v,0)||1;
    fabApPlanTotal[f] = Object.values(dayFabApPlan[f]).reduce((a,v)=>a+v,0)||1;
  });

  const cur = new Date(PROJ_START_D);
  while (cur <= PROJ_END_D) {
    const k  = cur.toISOString().slice(0,10);
    const dd = cur.getDate(), mm = cur.getMonth()+1;
    const lbl = `${String(dd).padStart(2,'0')}/${String(mm).padStart(2,'0')}`;

    cAll  += dayActMap[k]||0; cPlan += dayPlanMap[k]||0;
    cSWd  += daySwAct[k]||0;  cSWp  += daySwPlan[k]||0;
    cAPd  += dayApAct[k]||0;  cAPp  += dayApPlan[k]||0;
    FABRICS.forEach(f=>{ cFab[f]+=(dayFabAct[f][k]||0); cFabP[f]+=(dayFabPlan[f][k]||0); });

    const inAct  = lastActDt && cur <= lastActDt;
    const pct    = v=>Math.round(v*10000)/100;

    dailyProgress.labels.push(lbl);
    dailyProgress.plan_cum.push(pct(cPlan/TOTAL));
    dailyProgress.act_cum.push(inAct ? pct(cAll/TOTAL) : null);
    // push sw_plan/ap_plan ใน dailyProgress.fab[f] (same index กับ labels)
    FABRICS.forEach(f => {
      const fd = dailyProgress.fab[f];
      fd._spc += dayFabSwPlan[f][k]||0;
      fd._apc += dayFabApPlan[f][k]||0;
      fd._swAc += (dayFabSwAct[f]&&dayFabSwAct[f][k])||0;
      fd._apAc += (dayFabApAct[f]&&dayFabApAct[f][k])||0;
      // ใช้ lastActDt (วันสุดท้ายของทุก fabric) เหมือน overall chart
      // ทำให้เส้นทุก fabric หยุดตรงวันเดียวกัน
      const inFAct = lastActDt && cur <= lastActDt;
      fd.sw_plan.push(pct(fd._spc/(fabSwPlanTotal[f]||1)));
      fd.ap_plan.push(pct(fd._apc/(fabApPlanTotal[f]||1)));
      fd.sw_act.push(inFAct ? pct(fd._swAc/(fabSwPlanTotal[f]||1)) : null);
      fd.ap_act.push(inFAct ? pct(fd._apAc/(fabApPlanTotal[f]||1)) : null);
    });

    const swT = TOTAL_SW||1, apT = TOTAL_AP||1;
    dailyProgress.sw_plan.push(pct(cSWp/swT));
    dailyProgress.sw_act.push(inAct ? pct(cSWd/swT) : null);
    dailyProgress.ap_plan.push(pct(cAPp/apT));
    dailyProgress.ap_act.push(inAct ? pct(cAPd/apT) : null);

    // burndown
    dailyProgress.bd_plan.push(TOTAL - cPlan);
    dailyProgress.bd_act.push(inAct ? TOTAL - cAll : null);

    FABRICS.forEach(f => {
      const fLast = fabLastInstall[f];
      const fLastDt = fLast ? new Date(fLast+'T00:00:00') : null;
      const fT = (FAB_PLAN_T[f]||1);
      const inFabAct = fLastDt && cur <= fLastDt;
      dailyProgress.fab[f].plan.push(pct(cFabP[f]/fT));
      dailyProgress.fab[f].act.push(inFabAct ? pct(cFab[f]/fT) : null);
    });

    cur.setDate(cur.getDate()+1);
  }

  return {
    wk:        WK_BOUNDS.map(w => w.label),
    today_wk:  todayWk,
    last_install_date: lastInstallDate,
    upcoming:  Object.fromEntries(Object.entries(upcoming).sort().map(([d,fabs])=>[d,
               Object.fromEntries(Object.entries(fabs).map(([f,v])=>[f,
                 {qty:v.qty,rem:v.rem,locs:[...v.locs],types:[...v.types],cats:[...v.cats]}
               ]))
             ])),
    meta:      { total:TOTAL, installed, installed_sw:totalSwOk, installed_ap:totalApOk, installed_inf:totalInfOk, remaining, hold, overdue, on_time_qty:onTimeQty, on_time_pct:onTimePct, on_time_early:earlyQty, on_time_late:lateQty },
    hold_items: holdItems,
    insight:   { daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
                 pct_more:pctMore, days_late:daysLate, days_early:daysEarly, gauge_pct:gaugePct,
                 finish_date:finishDt.toISOString().slice(0,10), days_left:daysLeft },
    weekly:    { plan_all:PLAN_ALL, act_all:ACT_ALL, plan_sw:PLAN_SW, act_sw:ACT_SW,
                 plan_ap:PLAN_AP, act_ap:ACT_AP, bd_plan:BD_PLAN, bd_act:BD_ACT },
    fab_weekly: fabWeekly,
    fab_daily:  fabDaily,
    fab_daily_plan: fabDailyPlanOut,
    daily,
    daily_progress: dailyProgress,
    fabrics,
    fab_totals:      fabTotals,
    fab_plan_totals: FAB_PLAN_T,
    fab_colors:      FAB_COLORS,
    types,
    locations,
  };
}

// ── Cache wrapper ─────────────────────────────────────────────────────────────
async function getDashboard(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedData && (now - cacheTime) < CACHE_TTL) return cachedData;
  const wb = await getWorkbook();
  cachedData = calcDashboard(wb);
  cacheTime  = now;
  return cachedData;
}

// ── Routes ────────────────────────────────────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  res.set('Cache-Control','no-store');
  try   { res.json(await getDashboard()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/cache/refresh', async (req, res) => {
  try   { res.json({ success:true, data: await getDashboard(true) }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/health', (req, res) => res.json({
  status: 'ok',
  source: SHAREPOINT_URL ? 'sharepoint' : 'local_excel',
  cached_at: cacheTime ? new Date(cacheTime).toISOString() : null,
  cache_age_s: cacheTime ? Math.round((Date.now() - cacheTime) / 1000) : null,
}));

app.use(express.static(path.join(__dirname, '../frontend'), {etag:false, maxAge:0,
  setHeaders:(res)=>{ res.set('Cache-Control','no-store,no-cache,must-revalidate,proxy-revalidate');
    res.set('Pragma','no-cache'); res.set('Expires','0'); }}));
app.get('*', (req, res) => {
  res.set('Cache-Control','no-store,no-cache,must-revalidate,proxy-revalidate');
  res.set('Pragma','no-cache'); res.set('Expires','0');
  res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

process.on('uncaughtException', (err) => {
  console.error('[UNCAUGHT]', err.message, err.stack);
});
process.on('unhandledRejection', (reason) => {
  console.error('[UNHANDLED]', reason);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`SVB Dashboard running on port ${PORT}`);
  // pre-warm cache เพื่อดู error ตอน startup
  getDashboard().then(() => console.log('Cache warmed OK')).catch(e => console.error('[STARTUP ERROR]', e.message, e.stack));
});
