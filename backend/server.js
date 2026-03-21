const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, 'SDA_Installation_Plan_V2.xlsx');

function readWorkbook() {
  return XLSX.readFile(EXCEL_PATH);
}

// ── GET /api/summary ─────────────────────────────────────────────────────────
// ข้อมูล Dashboard: overall + per-fabric progress
app.get('/api/summary', (req, res) => {
  const wb = readWorkbook();
  const ws = wb.Sheets['Dashboard'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Row 5 (index 4): headers, Row 6 (index 5): values
  const overall = {
    total_devices: rows[5][1],
    completed:     rows[5][3],
    on_plan:       rows[5][5],
    hold:          rows[5][7],
    pending:       rows[5][9],
    progress_pct:  Math.round(rows[6][1] * 10000) / 100,
    actual_installed: rows[6][3],
    overdue_items: rows[6][7],
  };

  // Fabric rows: index 10-16 (D1-041 … PPW)
  const fabricHeaders = rows[9]; // ['Fabric','Total','Done','% Done','Hold','Remaining','Start Date','End Date','On Plan','Overdue Qty']
  const fabrics = rows.slice(10, 17).map(r => ({
    fabric:     r[0],
    total:      r[1],
    done:       r[2],
    pct_done:   Math.round(r[3] * 10000) / 100,
    hold:       r[4],
    remaining:  r[5],
    start_date: r[6] ? new Date((r[6] - 25569) * 86400000).toISOString().slice(0,10) : null,
    end_date:   r[7] ? new Date((r[7] - 25569) * 86400000).toISOString().slice(0,10) : null,
    on_plan:    r[8],
    overdue:    r[9],
  }));

  res.json({ overall, fabrics });
});

// ── GET /api/devices ──────────────────────────────────────────────────────────
// รายการอุปกรณ์ทั้งหมดจาก All_Detail พร้อม filter: fabric, status, location
app.get('/api/devices', (req, res) => {
  const { fabric, status, location, limit = 100, offset = 0 } = req.query;
  const wb = readWorkbook();
  const ws = wb.Sheets['All_Detail'];
  let rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  // normalize date fields
  rows = rows.map(r => ({
    ...r,
    'Install Date':    r['Install Date']    ? new Date((r['Install Date']    - 25569) * 86400000).toISOString().slice(0,10) : null,
    'Scheduled Date':  r['Scheduled Date']  ? new Date((r['Scheduled Date']  - 25569) * 86400000).toISOString().slice(0,10) : null,
  }));

  if (fabric)   rows = rows.filter(r => r['Fabric']   === fabric);
  if (status)   rows = rows.filter(r => r['Status']   === status);
  if (location) rows = rows.filter(r => r['Location'] && r['Location'].includes(location));

  const total = rows.length;
  const data  = rows.slice(Number(offset), Number(offset) + Number(limit));
  res.json({ total, offset: Number(offset), limit: Number(limit), data });
});

// ── GET /api/fabric/:name ────────────────────────────────────────────────────
// รายละเอียดของ Fabric เดียว
app.get('/api/fabric/:name', (req, res) => {
  const sheetName = `${req.params.name}_Detail`;
  const wb = readWorkbook();
  if (!wb.SheetNames.includes(sheetName)) {
    return res.status(404).json({ error: `Sheet ${sheetName} not found` });
  }
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  const byStatus = rows.reduce((acc, r) => {
    const s = r['Status'] || 'Unknown';
    acc[s] = (acc[s] || 0) + 1;
    return acc;
  }, {});

  res.json({ fabric: req.params.name, total: rows.length, by_status: byStatus, rows });
});

// ── GET /api/filters ─────────────────────────────────────────────────────────
// ค่า dropdown สำหรับ filter
app.get('/api/filters', (req, res) => {
  const wb = readWorkbook();
  const ws = wb.Sheets['All_Detail'];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  const fabrics   = [...new Set(rows.map(r => r['Fabric']).filter(Boolean))].sort();
  const statuses  = [...new Set(rows.map(r => r['Status']).filter(Boolean))].sort();
  const locations = [...new Set(rows.map(r => r['Location']).filter(Boolean))].sort();
  const types     = [...new Set(rows.map(r => r['Device Type']).filter(Boolean))].sort();

  res.json({ fabrics, statuses, locations, device_types: types });
});

// ── POST /api/update ─────────────────────────────────────────────────────────
// อัพเดทสถานะการติดตั้ง (เขียนกลับ Excel)
app.post('/api/update', (req, res) => {
  const { no, qty_success, install_date, status, remark } = req.body;
  if (!no) return res.status(400).json({ error: 'no (row number) is required' });

  const wb = readWorkbook();
  const ws = wb.Sheets['All_Detail'];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  const idx = rows.findIndex(r => r['No.'] === Number(no));
  if (idx === -1) return res.status(404).json({ error: `Row No. ${no} not found` });

  if (qty_success !== undefined) rows[idx]['Qty. Success'] = qty_success;
  if (install_date)              rows[idx]['Install Date'] = install_date;
  if (status)                    rows[idx]['Status']       = status;
  if (remark !== undefined)      rows[idx]['Remark']       = remark;

  const newWs = XLSX.utils.json_to_sheet(rows);
  wb.Sheets['All_Detail'] = newWs;
  XLSX.writeFile(wb, EXCEL_PATH);

  res.json({ success: true, updated: rows[idx] });
});

// ── Health check ─────────────────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SDA API running on port ${PORT}`));
