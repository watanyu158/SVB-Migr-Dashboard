const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const https   = require('https');
const http    = require('http');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const SHAREPOINT_URL = 'https://aitcoth-my.sharepoint.com/:x:/g/personal/suttipong_s_ait_co_th/IQB4depTDLOdRbI2UEHtAB7RAbaE9Ybz60zc_CjOHPUMkmI?e=WqeQTd&download=1';
const CACHE_PATH     = path.join(__dirname, 'sda_cache.xlsx');
const CACHE_TTL_MS   = 5 * 60 * 1000;

let cacheTime = 0;
let cachedWb  = null;

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    proto.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, res => {
      if ([301,302,303,307,308].includes(res.statusCode))
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode}`));
      const file = fs.createWriteStream(dest);
      res.pipe(file);
      file.on('finish', () => file.close(resolve));
      file.on('error', reject);
    }).on('error', reject);
  });
}

async function readWorkbook() {
  const now = Date.now();
  if (cachedWb && (now - cacheTime) < CACHE_TTL_MS) return cachedWb;
  try {
    await downloadFile(SHAREPOINT_URL, CACHE_PATH);
    cachedWb  = XLSX.readFile(CACHE_PATH);
    cacheTime = Date.now();
    return cachedWb;
  } catch (err) {
    console.error('SharePoint fetch failed:', err.message);
    const localPath = path.join(__dirname, 'SDA_Installation_Plan_V2.xlsx');
    if (fs.existsSync(localPath)) { cachedWb = XLSX.readFile(localPath); cacheTime = Date.now(); return cachedWb; }
    if (cachedWb) return cachedWb;
    throw new Error('No Excel source: ' + err.message);
  }
}

// ── Helper: Excel serial date → YYYY-MM-DD ────────────────────────────────────
const excelDate = v => v ? new Date((v - 25569) * 86400000).toISOString().slice(0,10) : null;
const num = v => (typeof v === 'number' ? v : 0);

// ── GET /api/summary ─────────────────────────────────────────────────────────
app.get('/api/summary', async (req, res) => {
  try {
    const wb   = await readWorkbook();
    const ws   = wb.Sheets['Dashboard'];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });

    // Overall
    const overall = {
      total_devices:    rows[5][1],
      completed:        rows[5][3],
      on_plan:          rows[5][5],
      hold:             rows[5][7],
      pending:          rows[5][9],
      progress_pct:     Math.round(rows[6][1] * 10000) / 100,
      actual_installed: rows[6][3],
      overdue_items:    rows[6][8],
    };

    // Fabrics from rows 10-16
    const fabrics = rows.slice(10, 17).map(r => ({
      fabric:     r[0],
      total:      r[1],
      done:       r[2],
      pct_done:   Math.round((r[3]||0) * 10000) / 100,
      hold:       r[4],
      remaining:  r[5],
      start_date: r[6] ? new Date((r[6]-25569)*86400000).toISOString().slice(0,10) : null,
      end_date:   r[7] ? new Date((r[7]-25569)*86400000).toISOString().slice(0,10) : null,
      on_plan:    r[8],
      overdue:    r[9],
    }));

    // Device category (rows 21-23)
    const categories = rows.slice(21, 24).map(r => ({
      category: r[0], total: r[1], done: r[2],
      pct_done: Math.round((r[3]||0)*10000)/100, remaining: r[4], hold: r[5],
    }));

    // Fabric SW/AP/Infra (rows 27-33)
    const fabric_detail = rows.slice(27, 34).map(r => ({
      fabric: r[0],
      sw_total: r[1], sw_done: r[2], sw_pct: Math.round((r[3]||0)*10000)/100,
      ap_total: r[4], ap_done: r[5], ap_pct: Math.round((r[6]||0)*10000)/100,
      inf_total: r[7], inf_done: r[8], inf_pct: Math.round((r[9]||0)*10000)/100,
    }));

    // Weekly velocity (rows 40-56, week rows start at index 40)
    const weeklyRows = rows.slice(40, 57).filter(r => typeof r[0] === 'number');
    const weekly = weeklyRows.map(r => ({
      week:       r[0],
      dates:      r[1],
      planned:    r[2],
      actual:     r[3],
      cum_plan:   r[4],
      cum_actual: r[5],
      sw_plan:    r[6],
      sw_actual:  r[7],
      ap_plan:    r[8],
      ap_actual:  r[9],
    }));

    // Total planned devices = cum_plan of last week
    const totalPlanned = weekly.length > 0 ? weekly[weekly.length-1].cum_plan : overall.total_devices;

    // Weekly % cumulative
    weekly.forEach(w => {
      w.cum_plan_pct   = totalPlanned > 0 ? Math.round(w.cum_plan   / totalPlanned * 10000) / 100 : 0;
      w.cum_actual_pct = totalPlanned > 0 ? Math.round(w.cum_actual / totalPlanned * 10000) / 100 : 0;
    });

    res.json({ overall, fabrics, categories, fabric_detail, weekly, cached_at: new Date(cacheTime).toISOString() });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/chart-data ───────────────────────────────────────────────────────
// Dynamic chart data computed from All_Detail (no more hardcode in frontend)
app.get('/api/chart-data', async (req, res) => {
  try {
    const wb   = await readWorkbook();
    const ws   = wb.Sheets['All_Detail'];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

    const FABRICS = ['D1-041','CFZ','T1-015','D1-091','RFF','AMF','PPW'];
    const FAB_COLORS = { 'D1-041':'#4361ee','CFZ':'#2bc48a','T1-015':'#ff9f43','D1-091':'#9b59b6','RFF':'#1abc9c','AMF':'#e74c3c','PPW':'#3498db' };

    // ── Daily installs ─────────────────────────────────────────────────────
    const dailyMap = {}; // date → { sw, ap, inf, total, fab:{} }
    for (const r of rows) {
      const inst = r['Install Date'];
      const qty  = num(r['Qty. Success']);
      const cat  = r['Category'];
      const fab  = r['Fabric'];
      if (!inst || qty <= 0) continue;
      // Convert Excel date serial or JS Date
      let d;
      if (typeof inst === 'number') {
        const dt = new Date((inst - 25569) * 86400000);
        d = `${dt.getDate()}/${String(dt.getMonth()+1).padStart(2,'0')}`;
      } else if (inst instanceof Date || (typeof inst === 'string' && inst.includes('-'))) {
        const dt = new Date(inst);
        d = `${dt.getDate()}/${String(dt.getMonth()+1).padStart(2,'0')}`;
      } else continue;

      if (!dailyMap[d]) dailyMap[d] = { sw:0, ap:0, inf:0, fab:{} };
      if (cat === 'Switch') dailyMap[d].sw += qty;
      else if (cat === 'AP') dailyMap[d].ap += qty;
      else dailyMap[d].inf += qty;
      if (!dailyMap[d].fab[fab]) dailyMap[d].fab[fab] = { sw:0, ap:0, inf:0 };
      if (cat === 'Switch') dailyMap[d].fab[fab].sw += qty;
      else if (cat === 'AP') dailyMap[d].fab[fab].ap += qty;
      else dailyMap[d].fab[fab].inf += qty;
    }

    // Sort dates chronologically
    const sortDates = dates => dates.sort((a,b) => {
      const [ad,am] = a.split('/').map(Number);
      const [bd,bm] = b.split('/').map(Number);
      return (am*100+ad) - (bm*100+bd);
    });
    const dailyDates = sortDates(Object.keys(dailyMap));

    const daily = {
      labels: dailyDates,
      sw:  dailyDates.map(d => dailyMap[d].sw),
      ap:  dailyDates.map(d => dailyMap[d].ap),
      inf: dailyDates.map(d => dailyMap[d].inf),
      total: dailyDates.map(d => dailyMap[d].sw + dailyMap[d].ap + dailyMap[d].inf),
    };

    // Per-fabric daily
    const fab_daily = {};
    for (const fab of FABRICS) {
      fab_daily[fab] = {
        sw:  dailyDates.map(d => (dailyMap[d]?.fab[fab]?.sw  || 0)),
        ap:  dailyDates.map(d => (dailyMap[d]?.fab[fab]?.ap  || 0)),
        inf: dailyDates.map(d => (dailyMap[d]?.fab[fab]?.inf || 0)),
      };
    }

    // ── Scheduled (plan) per day ───────────────────────────────────────────
    const schedMap = {};
    for (const r of rows) {
      const sched = r['Scheduled Date'];
      const qty   = num(r['Qty']);
      if (!sched || qty <= 0) continue;
      let d;
      if (typeof sched === 'number') {
        const dt = new Date((sched - 25569) * 86400000);
        d = `${dt.getDate()}/${String(dt.getMonth()+1).padStart(2,'0')}`;
      } else if (sched instanceof Date || (typeof sched === 'string' && sched.includes('-'))) {
        const dt = new Date(sched);
        d = `${dt.getDate()}/${String(dt.getMonth()+1).padStart(2,'0')}`;
      } else continue;
      schedMap[d] = (schedMap[d] || 0) + qty;
    }
    const daily_plan = dailyDates.map(d => schedMap[d] || 0);

    // ── Weekly buckets ─────────────────────────────────────────────────────
    // Use Dashboard weekly data for plan — already computed in /api/summary
    // Here just return the daily data and let frontend compute weekly from it

    // ── Burndown data ──────────────────────────────────────────────────────
    // Get weekly from Dashboard
    const wsDash  = wb.Sheets['Dashboard'];
    const dashRows = XLSX.utils.sheet_to_json(wsDash, { header:1, defval:null });
    const weeklyRows = dashRows.slice(40, 57).filter(r => typeof r[0] === 'number');
    const totalPlanned = weeklyRows.length > 0 ? weeklyRows[weeklyRows.length-1].cum_plan || weeklyRows[weeklyRows.length-1][4] : 1592;

    const wk_labels     = weeklyRows.map(r => r[1] ? String(r[1]).split(' ')[0]+' '+String(r[1]).split(' ')[1] : `W${r[0]}`);
    const wk_plan_cum   = weeklyRows.map(r => num(r[4]));
    const wk_actual_cum = weeklyRows.map(r => num(r[5]));
    const wk_plan_pct   = weeklyRows.map(r => totalPlanned>0 ? Math.round(num(r[4])/totalPlanned*10000)/100 : 0);
    const wk_actual_pct = weeklyRows.map(r => num(r[5])>0 ? Math.round(num(r[5])/totalPlanned*10000)/100 : null);
    // null out future weeks for actual
    let foundNull = false;
    const wk_actual_pct_trimmed = wk_actual_pct.map(v => {
      if (foundNull || v === null || v === 0) { foundNull=true; return null; }
      return v;
    });

    const bd_plan   = weeklyRows.map((_,i) => num(weeklyRows[weeklyRows.length-1][4]) - num(weeklyRows[i][4]) + num(weeklyRows[i][2]));
    const bd_actual = weeklyRows.map(r => num(r[5])>0 ? num(weeklyRows[weeklyRows.length-1][4]) - num(r[5]) : null);

    // SW/AP weekly %
    const sw_plan_pct   = weeklyRows.map(r => { let cum=0; weeklyRows.slice(0,weeklyRows.indexOf(r)+1).forEach(x=>cum+=num(x[6])); return Math.round(cum/num(dashRows[34][1])*10000)/100; });
    const sw_actual_pct = weeklyRows.map(r => num(r[7])>0 ? Math.round(num(r[7])/num(dashRows[34][1])*10000)/100 : null);

    res.json({
      daily: { labels: daily.labels, sw: daily.sw, ap: daily.ap, inf: daily.inf, total: daily.total, plan: daily_plan },
      fab_daily,
      weekly: {
        labels: wk_labels,
        plan_cum: wk_plan_cum,
        actual_cum: wk_actual_cum,
        plan_pct: wk_plan_pct,
        actual_pct: wk_actual_pct_trimmed,
        bd_plan, bd_actual,
      },
      fab_colors: FAB_COLORS,
      cached_at: new Date(cacheTime).toISOString(),
    });
  } catch(err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/devices ──────────────────────────────────────────────────────────
app.get('/api/devices', async (req, res) => {
  try {
    const { fabric, status, location, limit=100, offset=0 } = req.query;
    const wb   = await readWorkbook();
    const ws   = wb.Sheets['All_Detail'];
    let rows = XLSX.utils.sheet_to_json(ws, { defval:null });
    rows = rows.map(r => ({
      ...r,
      'Install Date':   r['Install Date']   ? new Date((r['Install Date']  -25569)*86400000).toISOString().slice(0,10) : null,
      'Scheduled Date': r['Scheduled Date'] ? new Date((r['Scheduled Date']-25569)*86400000).toISOString().slice(0,10) : null,
    }));
    if (fabric)   rows = rows.filter(r => r['Fabric']   === fabric);
    if (status)   rows = rows.filter(r => r['Status']   === status);
    if (location) rows = rows.filter(r => r['Location'] && r['Location'].includes(location));
    const total = rows.length;
    res.json({ total, offset:Number(offset), limit:Number(limit), data: rows.slice(Number(offset), Number(offset)+Number(limit)) });
  } catch(err) { res.status(500).json({ error:err.message }); }
});

// ── GET /api/filters ─────────────────────────────────────────────────────────
app.get('/api/filters', async (req, res) => {
  try {
    const wb   = await readWorkbook();
    const ws   = wb.Sheets['All_Detail'];
    const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
    res.json({
      fabrics:      [...new Set(rows.map(r=>r['Fabric']).filter(Boolean))].sort(),
      statuses:     [...new Set(rows.map(r=>r['Status']).filter(Boolean))].sort(),
      locations:    [...new Set(rows.map(r=>r['Location']).filter(Boolean))].sort(),
      device_types: [...new Set(rows.map(r=>r['Device Type']).filter(Boolean))].sort(),
    });
  } catch(err) { res.status(500).json({ error:err.message }); }
});

// ── POST /api/cache/refresh ───────────────────────────────────────────────────
app.post('/api/cache/refresh', async (req, res) => {
  cacheTime = 0; cachedWb = null;
  try { await readWorkbook(); res.json({ success:true, cached_at: new Date(cacheTime).toISOString() }); }
  catch(e) { res.status(500).json({ error:e.message }); }
});

app.get('/health', (req, res) => res.json({ status:'ok', cached_at: cacheTime ? new Date(cacheTime).toISOString() : null, cache_age_s: cacheTime ? Math.round((Date.now()-cacheTime)/1000) : null }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SDA API (dynamic mode) running on port ${PORT}`));
