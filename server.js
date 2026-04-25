const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');
const ExcelJS = require('exceljs');
const { dataPath, scriptsPath } = require('./src/paths');
const { extractDrawingData, buildDrawingExcel } = require('./src/server_drawing');
const { geminiAnalyzeDrawing, runCVAnalysis, RATES } = require('./src/drawing_analyzer');
const { parseDXF, extractCivilData, extractTotalAreaSqft } = require('./src/dxf_parser');
const { buildExcelFromDrawing, getDrawingPrompt } = require('./src/drawing_to_excel');
const { buildDXFExcel } = require('./src/dxf_to_excel');
const { analyzeDrawing, buildAIPrompt } = require('./src/drawing_intelligence');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const GEMINI_URL = (key) =>
  `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`;

/**
 * Calls the Gemini API with automatic retry + exponential backoff.
 * Retries on 503 (overloaded) and 429 (rate-limited) up to maxRetries times.
 */
async function fetchGeminiWithRetry(key, body, { maxRetries = 5, baseDelayMs = 2000 } = {}) {
  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const r = await fetch(GEMINI_URL(key), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data = await r.json();

    // Success — return immediately
    if (r.ok && data?.candidates?.[0]) return data;

    const code = data?.error?.code;
    const retryable = code === 503 || code === 429;

    // Non-retryable error — return the data as-is so callers handle it normally
    if (!retryable || attempt === maxRetries) return data;

    // Exponential backoff: 2s, 4s, 8s, 16s, 32s …
    const delay = baseDelayMs * Math.pow(2, attempt);
    console.warn(`Gemini ${code} on attempt ${attempt + 1}/${maxRetries + 1}. Retrying in ${delay}ms…`);
    await new Promise(resolve => setTimeout(resolve, delay));
  }
}

// ─── 1. CHAT ───────────────────────────────────────────────────────
app.post('/gemini', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { body } = req.body;
    const data = await fetchGeminiWithRetry(key, body);
    return res.json(data);
  } catch (e) { return res.status(500).json({ error: e.message }); }
});

// ─── 2. EXTRACT DATA ─────────────────────────────────────────────
// Strategy: Use AI chat response text as PRIMARY source (already has all data)
// Files only used if no aiResponse available
async function extractData(key, files, userText, aiResponse) {
  const parts = [];
  
  // If we have the AI response from chat, use it as primary input
  // This avoids re-processing files and gives much better results
  // STRATEGY: aiResponse (chat text) is PRIMARY source — it already has all data
  // Only use files as fallback if no aiResponse
  const primaryText = aiResponse || userText || '';

  if (!aiResponse) {
    // No AI response yet — send actual files to Gemini
    for (const f of (files || [])) {
      try {
        if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i))
          parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
        else if (f.type?.startsWith('image/'))
          parts.push({ inline_data: { mime_type: f.type || 'image/png', data: f.b64 } });
      } catch(e) { console.log('File skip:', e.message); }
    }
  }
  // If STILL no content, use files as fallback even when aiResponse present
  if (!aiResponse && parts.length === 0 && (files||[]).length === 0) {
    return { report_type:'general', project_title: userText||'PMC Report', company:'PMC', date:new Date().toLocaleDateString('en-IN'), summary:'', vendors:[], pricing:{old_rate:[],new_rate:[]}, commercial_terms:[], technical_specs:[], boq_items:[], recommendation:'No data provided.' };
  }

  const prompt = `You are a PMC data extraction expert. Extract ALL data from the content below into JSON.
Return ONLY raw JSON. No markdown. No backticks. Start with { end with }.

CONTENT TO EXTRACT FROM:
${primaryText}

You MUST extract real data from the content above. Do NOT use placeholder values like "v1","v2".
Extract actual vendor names, actual prices, actual specifications found in the content.

Return this exact JSON structure:
{"report_type":"comparison","project_title":"EXTRACT FROM CONTENT","company":"EXTRACT FROM CONTENT","date":"DD-MM-YYYY","summary":"2-3 lines from content",
"vendors":[{"name":"ACTUAL VENDOR NAME","vendor_name":"ACTUAL PERSON NAME","contact":"ACTUAL PHONE","quote_date":"DD-MM-YYYY","brand":"ACTUAL BRAND","product_description":"ACTUAL DESCRIPTION"}],
"pricing":{"old_rate":[{"label":"BASIC AMOUNT (OLD RATE)","values":[ACTUAL_NUMBERS]},{"label":"18% GST","values":[ACTUAL_NUMBERS]},{"label":"TOTAL AMOUNT WITH GST","values":[ACTUAL_NUMBERS]}],
"new_rate":[{"label":"BASIC AMOUNT (NEW RATE)","values":[ACTUAL_NUMBERS]},{"label":"18% GST","values":[ACTUAL_NUMBERS]},{"label":"TOTAL AMOUNT WITH GST","values":[ACTUAL_NUMBERS]}]},
"commercial_terms":[{"label":"PAYMENT TERMS","values":["ACTUAL VALUE FROM CONTENT"]},{"label":"DELIVERY TIME","values":["ACTUAL VALUE"]},{"label":"WARRANTY","values":["ACTUAL VALUE"]}],
"technical_specs":[{"label":"ACTUAL SPEC NAME","values":["ACTUAL SPEC VALUE"]}],
"boq_items":[{"sr":1,"description":"ACTUAL ITEM NAME","unit":"ACTUAL UNIT","qty":ACTUAL_NUMBER,"rate":ACTUAL_NUMBER,"amount":ACTUAL_NUMBER}],
"recommendation":"ACTUAL PMC recommendation from content"}

RULES: Use ACTUAL data from content | Numbers as numbers not strings | ONLY JSON`;

  parts.push({ text: prompt });

  const geminiData = await fetchGeminiWithRetry(key, { contents: [{ role: 'user', parts }], generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' } });
  let raw = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
  if (fb !== -1 && lb !== -1) raw = raw.slice(fb, lb + 1);
  try { return JSON.parse(raw.replace(/```json|```/g, '').trim()); }
  catch (e) {
    console.error('JSON parse fail:', raw.slice(0, 300));
    return { report_type: 'general', project_title: 'PMC Report', company: 'PMC', date: new Date().toLocaleDateString('en-IN'), summary: primaryText.slice(0, 200), vendors: [], pricing: { old_rate: [], new_rate: [] }, commercial_terms: [], technical_specs: [], boq_items: [], recommendation: primaryText.slice(0, 500) };
  }
}

// ─── 3. BUILD EXCEL — EXACT PMC FORMAT ────────────────────────────
async function buildExcel(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';
  const ws = wb.addWorksheet('Comparison');

  // Exact colors from template
  const NAVY    = 'FF1F3864';
  const MIDBLUE = 'FF2E75B6';
  const LTBLUE  = 'FFBDD7EE';
  const YELLOW  = 'FFFFD966';
  const GREEN   = 'FFE2EFDA';
  const DKGREEN = 'FF375623';
  const GREY    = 'FFF2F2F2';
  const WHITE   = 'FFFFFFFF';
  const LOWEST  = 'FF00B050';

  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const bdr  = { top: thin, left: thin, bottom: thin, right: thin };

  const vendors = d.vendors || [];
  const vc = Math.max(vendors.length, 1);
  const LC = 2 + vc; // last column index

  // Set exact col widths from template
  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 32;
  for (let i = 3; i <= LC; i++) ws.getColumn(i).width = 28;

  const sc = (cell, bgArgb, bold = false, fcArgb = 'FF000000', size = 10, align = 'left', wrap = true) => {
    cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
    cell.font   = { bold, color: { argb: fcArgb }, size, name: 'Calibri' };
    cell.alignment = { horizontal: align, vertical: 'middle', wrapText: wrap };
    cell.border = bdr;
  };

  const mergeRow = (r, text, bgArgb, fcArgb = 'FF000000', size = 10, bold = true, height = 18) => {
    ws.mergeCells(r, 1, r, LC);
    const c = ws.getCell(r, 1); c.value = text;
    sc(c, bgArgb, bold, fcArgb, size, 'center');
    ws.getRow(r).height = height;
  };

  let row = 1;

  // ROW 1 — Company title  (bg:1F3864 fc:FFFFFF size:14 bold)
  mergeRow(row++, d.company || 'VCT BHARUCH', NAVY, 'FFFFFFFF', 14, true, 22);

  // ROW 2 — Report title  (bg:2E75B6 fc:FFFFFF size:12 bold)
  mergeRow(row++, (d.project_title || 'COMPARISON REPORT').toUpperCase(), MIDBLUE, 'FFFFFFFF', 12, true, 20);

  // ROW 3 — Column headers  (bg:1F3864 fc:FFFFFF size:9 bold)
  const hRow = ws.getRow(row);
  const h1 = hRow.getCell(1); h1.value = 'SR NO';      sc(h1, NAVY, true, 'FFFFFFFF', 9, 'center');
  const h2 = hRow.getCell(2); h2.value = 'PARTICULARS'; sc(h2, NAVY, true, 'FFFFFFFF', 9, 'center');
  vendors.forEach((v, i) => {
    const c = hRow.getCell(i + 3);
    c.value = `${v.name || ''}\n(${v.brand || ''})\n${v.quote_date || ''}`;
    sc(c, NAVY, true, 'FFFFFFFF', 9, 'center');
  });
  hRow.height = 60; row++;

  // ROWS 4-8 — Vendor info
  const infoRows = [
    { lbl: 'AGENCY NAME',       bg: LTBLUE, bold: true,  vals: vendors.map(v => v.name || '') },
    { lbl: 'VENDOR NAME',       bg: GREY,   bold: false, vals: vendors.map(v => v.vendor_name || '') },
    { lbl: 'CONTACT NO',        bg: LTBLUE, bold: true,  vals: vendors.map(v => String(v.contact || '')) },
    { lbl: 'DATE OF QUOTATION', bg: GREY,   bold: false, vals: vendors.map(v => v.quote_date || '') },
    { lbl: 'BRAND',             bg: LTBLUE, bold: true,  vals: vendors.map(v => v.brand || '') },
  ];
  infoRows.forEach(({ lbl, bg, bold, vals }) => {
    const r = ws.getRow(row);
    const sr = r.getCell(1); sr.value = ''; sc(sr, bg, false, 'FF000000', 10, 'center');
    const lb = r.getCell(2); lb.value = lbl; sc(lb, bg, true, 'FF000000', 10, 'left');
    vals.forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, bold, 'FF000000', 10, 'center'); });
    ws.getRow(row).height = 16; row++;
  });

  // ROW 9 — Product desc header  A9:B9 merged = "SR NO", C9:G9 merged = "PRODUCT DESCRIPTION"
  ws.mergeCells(row, 1, row, 2);
  const pd1 = ws.getCell(row, 1); pd1.value = 'SR NO'; sc(pd1, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
  ws.mergeCells(row, 3, row, LC);
  const pd2 = ws.getCell(row, 3); pd2.value = 'PRODUCT DESCRIPTION'; sc(pd2, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
  ws.getRow(row).height = 16; row++;

  // ROW 10 — Product descriptions
  const pdRow = ws.getRow(row);
  const pdsr = pdRow.getCell(1); pdsr.value = '1'; sc(pdsr, GREY, false, 'FF000000', 10, 'center');
  const pdlb = pdRow.getCell(2); pdlb.value = 'PRODUCT DESCRIPTION'; sc(pdlb, GREY, true, 'FF000000', 10, 'left');
  vendors.forEach((v, i) => {
    const c = pdRow.getCell(i + 3); c.value = v.product_description || '';
    sc(c, WHITE, false, 'FF000000', 9, 'left');
  });
  ws.getRow(row).height = 90; row++;

  // PRICING OLD RATE
  if (d.pricing?.old_rate?.length) {
    mergeRow(row++, 'PRICING — OLD RATE', NAVY, 'FFFFFFFF', 10, true, 18);
    d.pricing.old_rate.forEach(({ label, values }, idx) => {
      const isTotal = label?.toUpperCase().includes('TOTAL');
      const bg = isTotal ? YELLOW : WHITE;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, isTotal, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => {
        const c = r.getCell(i + 3);
        const disp = (v === 0 || v === null || v === '') ? 'N/A' : v;
        c.value = disp;
        if (typeof v === 'number' && v > 0) c.numFmt = '#,##0';
        sc(c, bg, isTotal, 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 16; row++;
    });
  }

  // PRICING NEW RATE
  if (d.pricing?.new_rate?.length) {
    mergeRow(row++, 'PRICING — NEW RATE', NAVY, 'FFFFFFFF', 10, true, 18);
    let totalVals = [];
    d.pricing.new_rate.forEach(({ label, values }) => {
      const isTotal = label?.toUpperCase().includes('TOTAL');
      const isDisc  = label?.toUpperCase().includes('DISCOUNT');
      const bg = isTotal ? YELLOW : isDisc ? GREEN : WHITE;
      if (isTotal) totalVals = values || [];
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, isTotal, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => {
        const c = r.getCell(i + 3);
        c.value = (v === 0 || v === null || v === '') ? (isDisc ? '-' : 'N/A') : v;
        if (typeof v === 'number' && v > 0) c.numFmt = '#,##0';
        sc(c, bg, isTotal, 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 16; row++;
    });

    // LOWEST PRICE ROW
    if (totalVals.length) {
      const nums = totalVals.map(v => typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.]/g, '')) || 0);
      const minVal = Math.min(...nums.filter(n => n > 0));
      mergeRow(row++, 'LOWEST QUOTED PRICE (NEW RATE WITH GST)', NAVY, 'FFFFFFFF', 10, true, 18);
      const lr = ws.getRow(row);
      const lsr = lr.getCell(1); lsr.value = ''; sc(lsr, GREEN, false, 'FF000000', 10, 'center');
      const llb = lr.getCell(2); llb.value = 'TOTAL WITH GST (HIGHLIGHT = LOWEST)'; sc(llb, GREEN, true, 'FF000000', 10, 'left');
      nums.forEach((n, i) => {
        const c = lr.getCell(i + 3);
        const isLow = n === minVal && n > 0;
        if (n > 0) { c.value = n; c.numFmt = '₹#,##0'; }
        else c.value = 'N/A';
        sc(c, isLow ? LOWEST : WHITE, isLow, isLow ? 'FFFFFFFF' : 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 18; row++;
    }
  }

  // COMMERCIAL TERMS
  if (d.commercial_terms?.length) {
    mergeRow(row++, 'COMMERCIAL TERMS', NAVY, 'FFFFFFFF', 10, true, 18);
    d.commercial_terms.forEach(({ label, values }, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, true, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, false, 'FF000000', 9, 'center'); });
      ws.getRow(row).height = 40; row++;
    });
  }

  // TECHNICAL SPECS
  if (d.technical_specs?.length) {
    mergeRow(row++, 'TECHNICAL SPECIFICATIONS', NAVY, 'FFFFFFFF', 10, true, 18);
    d.technical_specs.forEach(({ label, values }, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = String(idx + 1); sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, true, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, false, 'FF000000', 10, 'center'); });
      ws.getRow(row).height = 16; row++;
    });
  }

  // BOQ
  if (d.boq_items?.length) {
    mergeRow(row++, 'BILL OF QUANTITIES', NAVY, 'FFFFFFFF', 11, true, 18);
    const bHdr = ws.getRow(row++);
    ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)'].forEach((h, i) => {
      const c = bHdr.getCell(i + 1); c.value = h; sc(c, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
    });
    let total = 0;
    d.boq_items.forEach((item, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row++);
      [item.sr, item.description, item.unit, item.qty, item.rate, item.amount].forEach((v, i) => {
        const c = r.getCell(i + 1); c.value = v;
        sc(c, bg, false, 'FF000000', 10, i === 0 || i > 1 ? 'center' : 'left');
        if (i >= 4 && typeof v === 'number') c.numFmt = '#,##0';
      });
      total += parseFloat(item.amount) || 0;
    });
    ws.mergeCells(row, 1, row, 4);
    const tc = ws.getCell(row, 1); tc.value = 'GRAND TOTAL'; sc(tc, YELLOW, true, 'FF000000', 10, 'right');
    const ta = ws.getCell(row, 6); ta.value = total; ta.numFmt = '₹#,##0'; sc(ta, YELLOW, true, 'FF000000', 10, 'center');
    ws.getRow(row).height = 18; row++;
  }

  // PMC RECOMMENDATION — dark green header + light green box
  mergeRow(row++, 'PMC RECOMMENDATION', DKGREEN, 'FFFFFFFF', 11, true, 18);
  ws.mergeCells(row, 1, row, LC);
  const recCell = ws.getCell(row, 1);
  recCell.value = d.recommendation || 'Refer to chat analysis above.';
  sc(recCell, GREEN, true, 'FF000000', 10, 'left');
  ws.getRow(row).height = 70; row++;

  // Summary
  if (d.summary) {
    ws.mergeCells(row, 1, row, LC);
    const sCell = ws.getCell(row, 1);
    sCell.value = 'SUMMARY: ' + d.summary;
    sc(sCell, LTBLUE, false, 'FF000000', 9, 'left', true);
    sCell.font = { ...sCell.font, italic: true };
    ws.getRow(row).height = 30; row++;
  }

  // Footer
  ws.mergeCells(row, 1, row, LC);
  const fCell = ws.getCell(row, 1);
  const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  fCell.value = `Prepared by: PMC Civil AI Agent  |  Date: ${today}  |  VCT Bharuch — Powered by Gemini AI`;
  fCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY } };
  fCell.font = { italic: true, size: 9, color: { argb: 'FF595959' }, name: 'Calibri' };
  fCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(row).height = 14;

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
  return wb;
}

// ─── 4. EXCEL ENDPOINT ─────────────────────────────────────────────
app.post('/export-excel', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;
    const d = await extractData(key, files, userText, aiResponse);
    const wb = await buildExcel(d);
    const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="PMC_Report_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch (err) {
    console.error('Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 5. PDF ENDPOINT (print-ready HTML) ────────────────────────────
app.post('/export-pdf', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;
    const d = await extractData(key, files, userText, aiResponse);
    const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const vendors = d.vendors || [];
    const vc = Math.max(vendors.length, 1);

    const th = (txt, bg = '#1F3864', fc = '#fff', bold = true) =>
      `<th style="background:${bg};color:${fc};padding:6px 8px;font-size:9px;border:1px solid #000;text-align:center;font-weight:${bold?'bold':'normal'};">${txt}</th>`;
    const td = (txt, bg = '#fff', align = 'center', bold = false, size = 9) =>
      `<td style="background:${bg};color:#000;padding:6px 8px;font-size:${size}px;border:1px solid #ccc;text-align:${align};font-weight:${bold?'bold':'normal'};vertical-align:top;">${txt||''}</td>`;
    const sectionHdr = (txt, bg = '#1F3864') =>
      `<tr><td colspan="${vc+2}" style="background:${bg};color:#fff;font-weight:bold;padding:7px 10px;font-size:10px;border:1px solid #000;">${txt}</td></tr>`;
    const fmtNum = (v) => typeof v === 'number' && v > 0 ? '₹' + v.toLocaleString('en-IN') : (v === 0 ? 'N/A' : (v || 'N/A'));

    let html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
@page{size:A3 landscape;margin:8mm}
*{box-sizing:border-box}
body{font-family:Calibri,Arial,sans-serif;font-size:9px;margin:0;color:#000}
h1{background:#1F3864;color:#fff;text-align:center;padding:9px;margin:0;font-size:15px}
h2{background:#2E75B6;color:#fff;text-align:center;padding:7px;margin:0 0 4px;font-size:12px}
table{width:100%;border-collapse:collapse;margin-bottom:2px}
.rec-hdr{background:#375623;color:#fff;font-weight:bold;padding:7px 10px;font-size:10px;margin-top:4px}
.rec-body{background:#E2EFDA;padding:10px;font-size:9px;border:1px solid #375623;white-space:pre-wrap}
.summ{background:#BDD7EE;padding:7px 10px;font-size:8px;font-style:italic;margin-top:3px}
.footer{text-align:center;font-size:8px;color:#595959;margin-top:6px;font-style:italic}
</style></head><body>
<h1>${d.company || 'VCT BHARUCH'}</h1>
<h2>${(d.project_title || 'COMPARISON REPORT').toUpperCase()}</h2>
<table>
<tr>${th('SR NO')}${th('PARTICULARS')}${vendors.map(v => th(`${v.name||''}<br><small>(${v.brand||''})</small><br><small>${v.quote_date||''}</small>`)).join('')}</tr>
${[['AGENCY NAME','BDD7EE',true,v=>v.name],['VENDOR NAME','F2F2F2',false,v=>v.vendor_name],['CONTACT NO','BDD7EE',false,v=>v.contact],['DATE OF QUOTATION','F2F2F2',false,v=>v.quote_date],['BRAND','BDD7EE',true,v=>v.brand]].map(([lbl,bg,bold,fn])=>`<tr>${td('',`#${bg}`)}<td style="background:#${bg};padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${lbl}</td>${vendors.map(v=>td(fn(v)||'',`#${bg}`,'center',bold)).join('')}</tr>`).join('')}
${sectionHdr('PRODUCT DESCRIPTION','#2E75B6')}
<tr>${td('1','#F2F2F2','center')}${td('<b>PRODUCT DESCRIPTION</b>','#F2F2F2','left',true)}${vendors.map(v=>td(v.product_description||'','#fff','left',false,8)).join('')}</tr>
${d.pricing?.old_rate?.length ? sectionHdr('PRICING — OLD RATE') + d.pricing.old_rate.map(({label,values})=>{const isT=label?.toUpperCase().includes('TOTAL');const bg=isT?'#FFD966':'#fff';return`<tr>${td('',bg)}${td(label,bg,'left',isT)}${(values||[]).map(v=>td(fmtNum(v),bg,'center',isT)).join('')}</tr>`;}).join('') : ''}
${d.pricing?.new_rate?.length ? sectionHdr('PRICING — NEW RATE') + d.pricing.new_rate.map(({label,values})=>{const isT=label?.toUpperCase().includes('TOTAL');const isD=label?.toUpperCase().includes('DISCOUNT');const bg=isT?'#FFD966':isD?'#E2EFDA':'#fff';return`<tr>${td('',bg)}${td(label,bg,'left',isT)}${(values||[]).map(v=>td(isT&&typeof v==='number'&&v>0?'₹'+v.toLocaleString('en-IN'):isD&&(v===0||!v)?'-':fmtNum(v),bg,'center',isT)).join('')}</tr>`;}).join('') : ''}
${(()=>{const tr=d.pricing?.new_rate?.find(r=>r.label?.toUpperCase().includes('TOTAL'));if(!tr)return'';const nums=(tr.values||[]).map(v=>typeof v==='number'?v:0);const minV=Math.min(...nums.filter(n=>n>0));return sectionHdr('LOWEST QUOTED PRICE')+`<tr>${td('')}<td style="background:#E2EFDA;padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">TOTAL WITH GST (HIGHLIGHT = LOWEST)</td>${nums.map(n=>n===minV&&n>0?`<td style="background:#00B050;color:#fff;padding:6px 8px;font-size:9px;border:1px solid #ccc;text-align:center;font-weight:bold;">₹${n.toLocaleString('en-IN')} ✓</td>`:td(n>0?'₹'+n.toLocaleString('en-IN'):'N/A','#fff','center')).join('')}</tr>`;})()}
${d.commercial_terms?.length?sectionHdr('COMMERCIAL TERMS')+d.commercial_terms.map(({label,values},i)=>{const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td('',bg)}<td style="background:${bg};padding:7px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${label}</td>${(values||[]).map(v=>td(v||'',bg,'center',false,8)).join('')}</tr>`;}).join(''):''}
${d.technical_specs?.length?sectionHdr('TECHNICAL SPECIFICATIONS')+d.technical_specs.map(({label,values},i)=>{const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td(i+1,bg,'center')}<td style="background:${bg};padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${label}</td>${(values||[]).map(v=>td(v||'',bg,'center')).join('')}</tr>`;}).join(''):''}
${d.boq_items?.length?(()=>{let tot=0;const rows=d.boq_items.map(({sr,description,unit,qty,rate,amount},i)=>{tot+=parseFloat(amount)||0;const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td(sr,bg,'center')}${td(description,bg,'left')}${td(unit,bg,'center')}${td(qty,bg,'center')}${td(rate?'₹'+rate.toLocaleString('en-IN'):'',bg,'center')}${td(amount?'₹'+amount.toLocaleString('en-IN'):'',bg,'center')}</tr>`;}).join('');return sectionHdr('BILL OF QUANTITIES')+`<tr>${['SR NO','DESCRIPTION','UNIT','QTY','RATE','AMOUNT'].map(h=>th(h,'#2E75B6')).join('')}</tr>${rows}<tr><td colspan="5" style="background:#FFD966;padding:7px;font-weight:bold;border:1px solid #000;text-align:right;">GRAND TOTAL</td><td style="background:#FFD966;padding:7px;font-weight:bold;border:1px solid #000;text-align:center;">₹${tot.toLocaleString('en-IN')}</td></tr>`;})():''}
</table>
<div class="rec-hdr">PMC RECOMMENDATION</div>
<div class="rec-body">${d.recommendation||'Refer to chat analysis.'}</div>
${d.summary?`<div class="summ">SUMMARY: ${d.summary}</div>`:''}
<div class="footer">Prepared by: PMC Civil AI Agent &nbsp;|&nbsp; Date: ${today} &nbsp;|&nbsp; VCT Bharuch — Powered by Gemini AI</div>
</body></html>`;

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="PMC_Report_${today.replace(/\//g,'-')}.html"`);
    res.send(html);
  } catch (err) {
    console.error('PDF error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 6. DRAWING ANALYSIS → MULTI-SHEET EXCEL (CV + AI) ──────────
app.post('/export-drawing', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;

    // Step 1: Run OpenCV pixel-level analysis on images
    let cvData = {};
    const imageFiles = (files||[]).filter(f => f.type?.startsWith('image/'));
    if (imageFiles.length > 0) {
      try { cvData = runCVAnalysis(imageFiles[0].b64); }
      catch(e) { console.log('CV skipped:', e.message); }
    }

    // Step 2: Gemini Vision — reads scale, dimensions, annotations + our formula engine
    let drawingData = null;
    if (files?.length > 0) {
      drawingData = await geminiAnalyzeDrawing(key, files, cvData, fetch);
    }

    // Step 3: Fallback to text-based extraction if needed
    if (!drawingData) {
      drawingData = await extractDrawingData(key, files, userText, aiResponse, fetch);
    }

    // Add CV metadata to drawing data
    drawingData.cv_analysis = cvData;
    drawingData.prepared_by = 'PMC Civil AI Agent';

    // Step 4: Build multi-sheet Excel
    const wb = await buildDrawingExcel(drawingData);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const pname = (drawingData.project_name||'Drawing').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_PMC_Analysis_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch (err) {
    console.error('Drawing Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 7. DXF UPLOAD & ANALYSIS ─────────────────────────────────────
// Uses drawing_intelligence.js — reads legend, auto-maps layers, extracts levels
app.post('/analyze-dxf', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    // ── Step 1: Drawing Intelligence — scan, detect legend, auto-map layers ──
    const analyzed = analyzeDrawing(dxfContent, filename);
    console.log(`[DXF] ${filename} | ${analyzed.total_layers} layers | ${analyzed.floor_levels.length} floor levels | ${analyzed.element_counts.wall_polylines} wall polylines | ${analyzed.unknown_layers.length} unknown layers`);

    // ── Step 2: Build rich prompt from what was actually found in drawing ─────
    const { RATES: ratesMap } = require('./src/dxf_parser');
    const ratesSummary = Object.entries(ratesMap).slice(0, 25).map(([k,v]) => `${k}:₹${v}`).join(' | ');
    const prompt = buildAIPrompt(analyzed, ratesSummary);

    // ── Step 3: Gemini AI interprets + fills BOQ ──────────────────────────────
    const geminiData = await fetchGeminiWithRetry(key, {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' }
    });
    let raw = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
    let geminiResult = {};
    if (fb !== -1) try { geminiResult = JSON.parse(raw.slice(fb, lb+1)); } catch(e) { console.error('JSON parse fail:', e.message); }

    // ── Step 4: Return everything — drawing data + AI interpretation ──────────
    res.json({
      success: true,
      dxf_data: {
        filename:         analyzed.filename,
        project_name:     analyzed.project_name,
        drawing_extents:  analyzed.drawing_extents,
        floor_levels:     analyzed.floor_levels,
        floor_heights:    analyzed.floor_heights,
        legend_items:     analyzed.legend_items,
        layer_summary:    analyzed.layer_summary,
        wall_by_thickness_m2: analyzed.wall_by_thickness_m2,
        hatch_summary:    analyzed.hatch_summary,
        element_counts:   analyzed.element_counts,
        unknown_layers:   analyzed.unknown_layers,
        unknown_blocks:   analyzed.unknown_blocks,
        all_texts:        analyzed.all_texts_sample,
        layer_names:      analyzed.layer_names,
        stats: {
          total_layers:    analyzed.total_layers,
          total_texts:     analyzed.total_texts,
          total_hatches:   analyzed.total_hatches,
          total_polylines: analyzed.total_polylines,
          total_inserts:   analyzed.total_inserts,
        }
      },
      interpretation: geminiResult
    });

  } catch (err) {
    console.error('DXF analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/export-dxf-excel', async (req, res) => {
  try {
    const { dxfContent, filename, aiResponse } = req.body;
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content.' });

    // Parse DXF
    const parsed = parseDXF(dxfContent);
    const civilData = extractCivilData(parsed, filename);

    // Get Gemini interpretation via analyze-dxf logic
    let geminiResult = {};
    try {
      const fakeReq = { body: { dxfContent, filename } };
      // reuse the prompt building
      const GEMINI_URL = k => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;
      const { RATES: rMap } = require('./src/dxf_parser');
      const rSummary = Object.entries(rMap).slice(0,20).map(([k,v])=>`${k}:${v}`).join(',');
      const specLines = (civilData.material_spec_texts || []).map(m => m.text).slice(0, 50);
      const prompt = `PMC civil engineer. Analyze DXF. Use ONLY data below, no invented values.
FILE:${filename} DECLARED_TYPE:${civilData.drawing_type} INFERRED_SHEET_KIND:${civilData.inferred_sheet_kind || '?'} SCALE:${civilData.scale || '?'}
PARSER_NOTES:${(civilData.extraction_notes || []).join(' | ')}
TEXTS:${civilData.all_texts.slice(0, 120).join(' | ')}
MATERIAL_SPECS_FROM_TEXT:${specLines.join(' | ')}
DIMS:${civilData.dimension_values.slice(0, 40).map(d => (d.count > 1 ? d.value_m + 'm×' + d.count : d.value_m + 'm') + '[' + d.layer + ']').join(', ')}
AREAS:${civilData.polyline_areas.slice(0, 20).map(p => p.area_sqm + 'sqm(' + p.layer + ')').join(', ')}
BLOCK_COUNTS:${Object.entries(civilData.block_counts || {}).slice(0, 40).map(([k, v]) => k + ':' + v).join(', ')}
ELEMENT_COUNTS:${JSON.stringify(civilData.element_counts || {})}
LAYERS:${civilData.layer_names.join(', ')}
RATES:${rSummary}
Return ONLY JSON:{"project_name":"","drawing_type":"FLOOR_PLAN|BASEMENT|PARKING|LIFT_SHAFT|STAIRCASE|STRUCTURAL_SECTION|FOUNDATION|SITE_LAYOUT|ROAD_PLAN|MEP_PLUMBING|MEP_ELECTRICAL|MEP_HVAC|ELEVATION|DETAIL_DRAWING|SECTION_ELEVATION|GENERAL","scale":"","spaces":[],"boq":[{"description":"","unit":"","qty":0,"rate":0,"amount":0}],"observations":[],"pmc_recommendation":"Explain what you used from TEXT vs block symbols vs dimensions; if section/elevation, avoid treating polyline areas as room BUA."}`;

      const geminiData4 = await fetchGeminiWithRetry(key, { contents: [{ role:'user', parts:[{text:prompt}] }], generationConfig: { maxOutputTokens:4096, temperature:0.0, responseMimeType:'application/json' } });
      let raw = geminiData4?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
      const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
      if (fb !== -1) geminiResult = JSON.parse(raw.slice(fb,lb+1));
    } catch(e) { console.log('Gemini interp fail:', e.message); }

    // Build Excel
    const wb = await buildDXFExcel(civilData, geminiResult, ExcelJS);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g,'-');
    const pname = (geminiResult.project_name||filename||'DXF').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_DXF_Analysis_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();

  } catch (err) {
    console.error('DXF Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 8. DRAWING → EXCEL (AI Analysis + Auto Excel) ───────────────
app.post('/drawing-to-excel', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });

    const { files, userText, aiResponse } = req.body;
    const GEMINI_URL = k => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;

    // Build Gemini request parts
    const parts = [];

    // Add files if provided (images/PDFs)
    for (const f of (files || [])) {
      try {
        if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i))
          parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
        else if (f.type?.startsWith('image/'))
          parts.push({ inline_data: { mime_type: f.type || 'image/png', data: f.b64 } });
      } catch(e) {}
    }

    // Use AI response text if no files (for chat-based flow)
    const promptText = getDrawingPrompt() + (aiResponse ? `

CHAT ANALYSIS:
${aiResponse}` : '') + (userText ? `
Note: ${userText}` : '');
    parts.push({ text: promptText });

    // Call Gemini
    const geminiData5 = await fetchGeminiWithRetry(key, {
        contents: [{ role: 'user', parts }],
        generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' }
      });

    let raw = geminiData5?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
    let drawingData = {};
    if (fb !== -1) {
      try { drawingData = JSON.parse(raw.slice(fb, lb+1)); }
      catch(e) { console.log('Parse fail:', e.message); }
    }

    // Build Excel
    const wb = await buildExcelFromDrawing(drawingData);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const pname = (drawingData.project_name || 'Drawing').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_PMC_Estimate_${today}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('Drawing→Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});
// ── NEW: DXF → AREA STATEMENT + OVERALL SUMMARY auto-update ──
app.post('/update-area-from-dxf', async (req, res) => {
  try {
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    const totalAreaSqft = extractTotalAreaSqft(dxfContent);
    if (!totalAreaSqft || totalAreaSqft <= 0)
      return res.status(400).json({ error: 'No closed polylines found in DXF. Area calculate nahi hui.' });

    const estimatePath = path.join(__dirname, 'data', 'templates', 'UPDATED-OVERALL-ESTIMATE-MODESTAA-10.04.2026.xlsx');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(estimatePath);

    // Update AREA STATEMENT C73
    const wsArea = wb.getWorksheet('AREA STATEMENT');
    if (wsArea) wsArea.getCell('C73').value = totalAreaSqft;

    // Update OVERALL SUMMARY
    const wsOS = wb.getWorksheet('OVERALL SUMMARY');
    if (wsOS) {
      // Row 6 display text
      wsOS.getCell('B6').value = `TOTAL AREA: ${totalAreaSqft.toLocaleString('en-IN', {maximumFractionDigits:2})} SQFT`;
      // Helper cell J6 stores area value
      wsOS.getCell('J6').value = totalAreaSqft;
      // Replace all hardcoded 273613.53 with dynamic reference to J6
      wsOS.eachRow(row => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (typeof cell.value === 'string' && cell.value.includes('273613.53')) {
            cell.value = cell.value.split('273613.53').join("'OVERALL SUMMARY'!$J$6");
          }
        });
      });
    }

    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=ESTIMATE-UPDATED-${today}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Area update error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ── NEW: Fill MODESTAA template from drawing (type-aware) ──
// Detects project type from DXF content. If high-rise residential,
// opens the MODESTAA template and fills drawing-derived cells only.
// Otherwise, builds a fresh workbook via buildExcelFromDrawing with
// the right BOQ sheet for the detected project type (cafe / institute
// / commercial / road / generic).
app.post('/fill-template-from-drawing', async (req, res) => {
  try {
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    const parsed = parseDXF(dxfContent);
    const civil  = extractCivilData(parsed, filename || 'drawing.dxf');
    const ptype  = (civil.project_type || 'generic').toLowerCase();
    const ec     = civil.element_counts || {};
    const totalAreaSqft = extractTotalAreaSqft(dxfContent) || 0;
    const totalAreaSqm  = totalAreaSqft > 0 ? Math.round((totalAreaSqft / 10.764) * 100) / 100 : 0;

    // ── HIGH-RISE: use MODESTAA template, fill only drawing-derived cells ──
    if (ptype === 'high_rise_residential') {
      const estimatePath = path.join(__dirname, 'data', 'templates', 'UPDATED-OVERALL-ESTIMATE-MODESTAA-10.04.2026.xlsx');
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(estimatePath);

      // AREA STATEMENT C73 — total area
      if (totalAreaSqft > 0) {
        const wsArea = wb.getWorksheet('AREA STATEMENT');
        if (wsArea) wsArea.getCell('C73').value = totalAreaSqft;
      }

      // OVERALL SUMMARY B6 / J6
      const wsOS = wb.getWorksheet('OVERALL SUMMARY');
      if (wsOS && totalAreaSqft > 0) {
        wsOS.getCell('B6').value = `TOTAL AREA: ${totalAreaSqft.toLocaleString('en-IN',{maximumFractionDigits:2})} SQFT`;
        wsOS.getCell('J6').value = totalAreaSqft;
        wsOS.eachRow(row => {
          row.eachCell({ includeEmpty: false }, cell => {
            if (typeof cell.value === 'string' && cell.value.includes('273613.53')) {
              cell.value = cell.value.split('273613.53').join("'OVERALL SUMMARY'!$J$6");
            }
          });
        });
      }

      // DRAWING-DERIVED COUNTS sheet (new) — record what the parser read
      let wsCounts = wb.getWorksheet('DRAWING COUNTS');
      if (!wsCounts) wsCounts = wb.addWorksheet('DRAWING COUNTS');
      wsCounts.getCell('A1').value = 'ELEMENT';
      wsCounts.getCell('B1').value = 'COUNT FROM DRAWING';
      wsCounts.getCell('C1').value = 'SOURCE';
      [['Floors', ec.floor_count || 0, (ec.floor_labels || []).join(', ')],
       ['Doors',  ec.door_count  || 0, 'block / layer match'],
       ['Windows',ec.window_count|| 0, 'block / layer match'],
       ['Lifts',  ec.lift_count  || 0, 'block / layer / text'],
       ['Staircases', ec.staircase_count || 0, 'block / layer / text'],
       ['Columns', ec.column_count || 0, 'block / layer match'],
       ['Beams',   ec.beam_count   || 0, 'block / layer match'],
       ['Footings',ec.footing_count|| 0, 'block / layer match'],
       ['Toilets', ec.toilet_count || 0, 'text annotations'],
       ['Kitchens',ec.kitchen_count|| 0, 'text annotations'],
       ['Bedrooms',ec.bedroom_count|| 0, 'text annotations'],
       ['Wall length (m)', civil.wall_length_m || 0, 'LINE entities on wall layers'],
       ['Total area (sqft)', totalAreaSqft, 'closed polylines (shoelace)'],
       ['Project type detected', civil.project_type || 'generic', 'dxf_parser.detectProjectType']
      ].forEach((row, i) => {
        row.forEach((v, j) => { wsCounts.getCell(i+2, j+1).value = v; });
      });

      const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=MODESTAA-FILLED-${today}.xlsx`);
      await wb.xlsx.write(res);
      return res.end();
    }

    // ── OTHER TYPES: build fresh type-aware workbook ──
    const data = {
      drawing_type:    civil.drawing_type === 'FLOOR_PLAN' ? 'BUILDING' : 'SITE_LAYOUT',
      project_type:    ptype,
      total_area_sqm:  totalAreaSqm,
      total_area_sqft: totalAreaSqft,
      element_counts:  ec,
      wall_length_m:   civil.wall_length_m || 0,
      buildings: totalAreaSqm > 0 ? [{ name: 'Building', area_sqm: totalAreaSqm, floors: ec.floor_count || 0 }] : [],
      roads: [],
      project_name: civil.title_block?.project_name || filename || 'Project',
      source: `DXF parser — project type: ${ptype}`
    };
    const wb = await buildExcelFromDrawing(data);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${ptype.toUpperCase()}-ESTIMATE-${today}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Fill template error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── DWG/DXF ANALYSIS — Convert to PNG + Gemini Vision ───────────
// Strategy: dwg_converter.py renders DXF/DWG to PNG using ezdxf+matplotlib
// Then Gemini SEES the actual drawing like a human engineer
app.post('/analyze-dwg', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });

    const { b64, filename, detailMode } = req.body;
    if (!b64) return res.status(400).json({ error: 'No file data provided.' });
    const useDetail = detailMode === true || detailMode === 'true' || detailMode === 1;

    const fs = require('fs');
    const { execSync } = require('child_process');
    const os = require('os');

    // Write uploaded file to temp
    const ext = filename?.match(/\.(dxf|dwg|dwf)$/i)?.[1]?.toLowerCase() || 'dxf';
    const tmpIn  = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.${ext}`);
    const tmpPng = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.png`);

    fs.writeFileSync(tmpIn, Buffer.from(b64, 'base64'));

    // Run converter. DXF/DWG → ezdxf Python script. DWF → LibreOffice fallback.
    const scriptPath = scriptsPath('dwg_converter.py');
    let converterResult = {};

    if (ext === 'dwf') {
      // DWF is a compressed vector/image bundle. Try LibreOffice → PNG.
      try {
        const soffice = process.platform === 'win32'
          ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
          : 'libreoffice';
        execSync(`${soffice} --headless --convert-to png --outdir "${os.tmpdir()}" "${tmpIn}"`,
                 { timeout: 90000 });
        const base = path.basename(tmpIn, '.dwf');
        const libreOut = path.join(os.tmpdir(), `${base}.png`);
        if (fs.existsSync(libreOut)) {
          converterResult = { success: true, png_path: libreOut, texts: [], dimensions: [], layers: [], drawing_type: 'DWF_RENDER' };
        } else {
          converterResult = { success: false, error: 'LibreOffice did not produce PNG from DWF' };
        }
      } catch (e) {
        converterResult = { success: false, error: `DWF conversion failed: ${e.message}. Gemini Vision will still attempt analysis if a PNG thumbnail is embedded.` };
      }
    } else {
      try {
        const py = process.env.PMC_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
        const dpi = useDetail ? 180 : 150;
        const tiledArg = useDetail ? 'true' : 'false';
        const out = execSync(
          `${py} "${scriptPath}" "${tmpIn}" "${tmpPng}" ${dpi} ${tiledArg}`,
          { timeout: 120000, maxBuffer: 20 * 1024 * 1024 }
        );
        converterResult = JSON.parse(out.toString());
      } catch (e) {
        converterResult = { success: false, error: e.message };
      }
    }

    // DWF or any path that has PNG but no tiles yet: split with helper script
    if (useDetail && converterResult.png_path && fs.existsSync(converterResult.png_path)
        && (!converterResult.tiles || !converterResult.tiles.length)) {
      try {
        const outDir = path.dirname(converterResult.png_path);
        const baseName = path.basename(converterResult.png_path, path.extname(converterResult.png_path));
        const tileScript = path.join(__dirname, 'scripts', 'tile_only.py');
        const py = process.env.PMC_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
        const tout = execSync(
          `${py} "${tileScript}" "${converterResult.png_path}" "${outDir}" "${baseName}"`,
          { timeout: 60000, maxBuffer: 10 * 1024 * 1024 }
        );
        const ta = JSON.parse(tout.toString().trim() || '[]');
        if (Array.isArray(ta) && ta.length) converterResult.tiles = ta;
      } catch (e) {
        console.warn('Tile split (fallback):', e.message);
      }
    }

    // Gemini: images first, then one text block (vision models handle this well)
    const parts = [];
    if (converterResult.png_path && fs.existsSync(converterResult.png_path)) {
      const pngB64 = fs.readFileSync(converterResult.png_path).toString('base64');
      parts.push({ inline_data: { mime_type: 'image/png', data: pngB64 } });
    }
    if (useDetail && Array.isArray(converterResult.tiles)) {
      for (const t of converterResult.tiles) {
        if (t.path && fs.existsSync(t.path)) {
          try {
            const tb = fs.readFileSync(t.path).toString('base64');
            parts.push({ inline_data: { mime_type: 'image/png', data: tb } });
            try { fs.unlinkSync(t.path); } catch (e) {}
          } catch (e) { /* skip bad tile */ }
        }
      }
    }
    if (converterResult.png_path) {
      try { fs.unlinkSync(converterResult.png_path); } catch (e) {}
    }

    const nDetailTiles = (converterResult.tiles || []).length;
    const visionHeader = useDetail && nDetailTiles
      ? `MULTI-IMAGE INPUT: The user message includes ${1 + nDetailTiles} images in order: (1) full sheet render, then (2–${1 + nDetailTiles}) 2×2 quadrant crops of the SAME drawing (higher effective zoom for small text, legend, dimensions). Synthesize one coherent analysis; do not treat crops as different drawings.\n\n`
      : (parts.length > 0
        ? 'SINGLE-IMAGE INPUT: The drawing render is above this text. Read like CAD: title block, legend, dimensions, hatches, symbols, notes.\n\n'
        : '');

    // Always include extracted text + dimension data
    const textSummary = (converterResult.texts || []).map(t => t.text).slice(0, 150).join(' | ');
    const dimSummary  = (converterResult.dimensions || [])
      .filter(d => d.value).map(d => `${d.value}${d.text ? ' ('+d.text+')' : ''}`).slice(0, 80).join(', ');
    const layers = (converterResult.layers || []).join(', ');

    const prompt = `${visionHeader}You are a SENIOR PMC CIVIL ENGINEER with 20 years India experience analyzing an AutoCAD drawing.
(Use every image in this user message, if any, together with the extracted text list below.)

FILE: ${filename}
LAYERS FOUND: ${layers || 'See image'}
ALL TEXT IN DRAWING: ${textSummary || 'See image'}
DIMENSIONS FOUND: ${dimSummary || 'See image'}
${converterResult.error ? 'Render note: ' + converterResult.error : ''}

══════════════════════════════════════════════════════
STEP 1 — READ LEGEND / SYMBOL TABLE FROM DRAWING
══════════════════════════════════════════════════════
Every drawing has a legend box. Find it and read:
- Each symbol/hatch pattern and its label (e.g. "230MM THK. BRICK WALL", "100MM BLOCK WALL", "RCC PARDI")
- Map each hatch/color/pattern to its material meaning
- Note which AutoCAD LAYER corresponds to each element type
- If no legend, infer from layer names (e.g. "AR-HATCH 230 MM BRICK WALL" = 230mm brick wall)

══════════════════════════════════════════════════════
STEP 2 — READ TITLE BLOCK
══════════════════════════════════════════════════════
Project name, drawing number, scale, date, architect/engineer.
If not visible: write "Not shown in drawing" — do NOT invent.

══════════════════════════════════════════════════════
STEP 3 — IDENTIFY DRAWING TYPE & READ ALL FLOOR LEVELS
══════════════════════════════════════════════════════
Drawing type: SECTION / ELEVATION / FLOOR_PLAN / STRUCTURAL / SITE_PLAN / FOUNDATION
Read every floor level annotation (e.g. "+7590 MM LEVEL", "THIRD BASEMENT LEVEL").
Calculate floor heights between consecutive levels.

══════════════════════════════════════════════════════
STEP 4 — EXTRACT QUANTITIES BASED ON WHAT YOU SEE
══════════════════════════════════════════════════════
Use the legend you read in Step 1 to identify elements.
For SECTION drawing: wall lengths × thickness × floor height = volume
For FLOOR PLAN: room areas, wall lengths, openings count
For STRUCTURAL: column sizes, beam dimensions, slab thickness
For SITE/ROAD: road lengths × widths

Calculate:
| Element | Nos | Length (m) | Width/Thk (m) | Height (m) | Qty | Unit |

══════════════════════════════════════════════════════
STEP 5 — BOQ WITH GUJARAT DSR 2025 RATES
══════════════════════════════════════════════════════
100mm block wall: ₹4200/cum | 230mm brick wall: ₹4800/cum
RCC M25: ₹5500/cum | RCC M30: ₹5800/cum | Steel Fe500: ₹56/kg
Excavation: ₹180/cum | Formwork: ₹180/sqmt | PQC road: ₹1800/sqmt
Plaster 12mm: ₹280/sqmt | Waterproofing: ₹450/sqmt

══════════════════════════════════════════════════════
STEP 6 — PMC OBSERVATIONS
══════════════════════════════════════════════════════
IS code compliance, design comments, missing information, recommendations.

CRITICAL RULES:
- Read from drawing — do NOT invent values not visible
- If something not visible: state "Not shown in drawing"  
- Date = today's actual date, not from drawing unless shown
- Work floor by floor if it's a section/multi-floor drawing`;

    parts.push({ text: prompt });

    // Call Gemini (with retry)
    const data = await fetchGeminiWithRetry(key, {
      contents: [{ role: 'user', parts }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.1 }
    });
    // Debug: log Gemini response if no candidates
    if (!data?.candidates?.[0]) {
      console.error("DWG Gemini raw response:", JSON.stringify(data).slice(0,500));
    }
    const analysis = data?.candidates?.[0]?.content?.parts?.[0]?.text ||
      (data?.error ? "Gemini API Error: " + JSON.stringify(data.error) : null) ||
      `## DWG/DXF File: ${filename}\n\n` +
      `**PNG rendered:** ${converterResult.png_path ? "Yes" : "No"}\n` +
      `**Layers:** ${layers || "none"}\n` +
      `**Texts found:** ${(converterResult.texts||[]).length}\n` +
      `**Dimensions found:** ${(converterResult.dimensions||[]).length}\n\n` +
      (textSummary ? `**Annotations:**\n${textSummary}\n` : "") +
      (dimSummary ? `**Dimensions:**\n${dimSummary}\n` : "") +
      "\n> Check Render logs: server should show DWG Gemini raw response above.";

    // Cleanup temp input
    try { fs.unlinkSync(tmpIn); } catch(e) {}

    res.json({
      success: true,
      analysis,
      converter: converterResult,
      detailMode: useDetail,
      quadrantTiles: nDetailTiles,
    });
  } catch (err) {
    console.error('DWG analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── 9. SYMBOL CLASSIFICATION — Step 1: classify known/unknown ────
// Called right after DXF upload. Returns known symbols + unknown list.
// Unknown symbols will be shown to user as questions in the chat UI.
app.post('/classify-dxf', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content.' });

    const fs = require('fs');

    // Load learned symbols from disk
    const learnedPath = dataPath('symbols-learned.json');
    let learned = { blocks: {}, layers: {} };
    try { learned = JSON.parse(fs.readFileSync(learnedPath, 'utf8')); } catch(e) {}

    // Parse DXF
    const parsed = parseDXF(dxfContent);
    const civilData = extractCivilData(parsed, filename);

    const allBlocks = Object.keys(civilData.block_counts || {});
    const allLayers = civilData.layer_names || [];

    // Split blocks into known (in learned dict) vs unknown
    const knownBlocks = {};
    const unknownBlocks = [];
    for (const b of allBlocks) {
      const bUp = b.toUpperCase();
      // Check learned dict first
      if (learned.blocks[bUp]) {
        knownBlocks[b] = learned.blocks[bUp];
        continue;
      }
      // Check common AutoCAD naming conventions
      const autoType = guessBlockType(b);
      if (autoType) {
        knownBlocks[b] = autoType;
      } else {
        unknownBlocks.push({ name: b, count: civilData.block_counts[b] || 1 });
      }
    }

    // Split layers into known vs unknown
    const knownLayers = {};
    const unknownLayers = [];
    const LAYER_PREFIXES = {
      'A-': 'architectural', 'S-': 'structural', 'E-': 'electrical',
      'P-': 'plumbing', 'M-': 'mechanical', 'C-': 'civil',
      'WALL': 'wall', 'DOOR': 'door', 'WINDOW': 'window',
      'COLUMN': 'column', 'COL': 'column', 'BEAM': 'beam',
      'SLAB': 'slab', 'STAIR': 'staircase', 'LIFT': 'lift',
      'RAMP': 'ramp', 'TOILET': 'toilet', 'KITCHEN': 'kitchen',
      'PARK': 'parking', 'ROAD': 'road', 'HATCH': 'hatch',
      'DIM': 'dimension', 'TEXT': 'text', 'TITLE': 'title-block',
      'DEFPOINTS': 'dimension-helper', '0': 'default'
    };
    for (const l of allLayers) {
      const lUp = l.toUpperCase();
      if (learned.layers[lUp]) { knownLayers[l] = learned.layers[lUp]; continue; }
      let matched = false;
      for (const [pfx, type] of Object.entries(LAYER_PREFIXES)) {
        if (lUp.startsWith(pfx) || lUp.includes(pfx)) {
          knownLayers[l] = type; matched = true; break;
        }
      }
      if (!matched) unknownLayers.push(l);
    }

    // Ask Gemini to classify ONLY the unknowns (saves tokens)
    let geminiClassified = { blocks: {}, layers: {} };
    const needsGemini = unknownBlocks.length > 0 || unknownLayers.length > 0;
    if (needsGemini) {
      const classifyPrompt = `You are a senior AutoCAD civil drawing expert.
Classify these unknown block names and layer names from a civil DXF drawing.

UNKNOWN BLOCKS (name → count in drawing):
${unknownBlocks.map(b => `${b.name} (×${b.count})`).join('\n') || 'none'}

UNKNOWN LAYERS:
${unknownLayers.join('\n') || 'none'}

DRAWING CONTEXT:
- File: ${filename}
- Drawing type: ${civilData.drawing_type}
- Texts found: ${civilData.all_texts.slice(0, 30).join(', ')}

For each item, classify as one of:
door | window | column | beam | slab | wall | staircase | lift | ramp | toilet | kitchen | bedroom | parking | road | hatch | dimension | text | furniture | equipment | unknown

Return ONLY raw JSON:
{
  "blocks": { "BLOCK_NAME": "type", ... },
  "layers": { "LAYER_NAME": "type", ... },
  "still_unknown_blocks": ["BLOCK_NAME"],
  "still_unknown_layers": ["LAYER_NAME"]
}

Rules:
- If you can reasonably guess from name → classify it
- If genuinely unclear → put in still_unknown arrays
- No markdown, no explanation, only JSON`;

      const gData = await fetchGeminiWithRetry(key, {
        contents: [{ role: 'user', parts: [{ text: classifyPrompt }] }],
        generationConfig: { maxOutputTokens: 2048, temperature: 0.0, responseMimeType: 'application/json' }
      });
      let raw = gData?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
      const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
      if (fb !== -1) {
        try { geminiClassified = JSON.parse(raw.slice(fb, lb + 1)); } catch(e) {}
      }
    }

    // Merge all known
    const finalKnownBlocks = { ...knownBlocks, ...(geminiClassified.blocks || {}) };
    const finalKnownLayers = { ...knownLayers, ...(geminiClassified.layers || {}) };

    // These still need user input
    const askUserBlocks = (geminiClassified.still_unknown_blocks || [])
      .map(name => ({ name, count: civilData.block_counts[name] || 1 }));
    const askUserLayers = geminiClassified.still_unknown_layers || [];

    res.json({
      success: true,
      filename,
      dxf_data: civilData,
      known_blocks: finalKnownBlocks,
      known_layers: finalKnownLayers,
      ask_user_blocks: askUserBlocks,
      ask_user_layers: askUserLayers,
      needs_questions: askUserBlocks.length > 0 || askUserLayers.length > 0
    });

  } catch (err) {
    console.error('classify-dxf error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Helper: guess block type from common AutoCAD naming conventions
function guessBlockType(name) {
  const n = name.toUpperCase();
  if (/^D\d+$|DR[-_]?\d|DOOR|FLUSH|SFD|DOOR[-_]/.test(n)) return 'door';
  if (/^W\d+$|WIN[-_]?\d|WINDOW|ALUM[-_]WIN|CASEMENT/.test(n)) return 'window';
  if (/COL[-_]?\d|^C\d+$|COLUMN|PILLAR/.test(n)) return 'column';
  if (/BEAM|BM[-_]?\d/.test(n)) return 'beam';
  if (/LIFT|ELEV|ELEVATOR/.test(n)) return 'lift';
  if (/STAIR|STC|STEP/.test(n)) return 'staircase';
  if (/RAMP/.test(n)) return 'ramp';
  if (/TOILET|WC|BATH/.test(n)) return 'toilet';
  if (/KITCHEN|PANTRY/.test(n)) return 'kitchen';
  if (/BED|MASTER/.test(n)) return 'bedroom';
  if (/SOFA|TABLE|CHAIR|FURN/.test(n)) return 'furniture';
  if (/TREE|SHRUB|PLANT/.test(n)) return 'landscaping';
  if (/CAR|VEHICLE|PARK/.test(n)) return 'parking';
  return null;
}

// ─── 10. ANALYZE WITH USER ANSWERS — Step 2: full BOQ after Q&A ───
// Receives: original dxf_data + all known symbols + user's answers
// Returns: full Gemini BOQ analysis → used to generate Excel
app.post('/analyze-with-answers', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });

    const { dxfContent, filename, knownBlocks, knownLayers, userAnswers, dxfData } = req.body;
    const fs = require('fs');

    // Save user answers to symbols-learned.json for future drawings
    const learnedPath = dataPath('symbols-learned.json');
    let learned = { blocks: {}, layers: {} };
    try { learned = JSON.parse(fs.readFileSync(learnedPath, 'utf8')); } catch(e) {}

    // Merge user answers into learned dict
    if (userAnswers?.blocks) {
      for (const [name, type] of Object.entries(userAnswers.blocks)) {
        if (type && type !== 'skip') learned.blocks[name.toUpperCase()] = type;
      }
    }
    if (userAnswers?.layers) {
      for (const [name, type] of Object.entries(userAnswers.layers)) {
        if (type && type !== 'skip') learned.layers[name.toUpperCase()] = type;
      }
    }
    try { fs.writeFileSync(learnedPath, JSON.stringify(learned, null, 2)); } catch(e) {}

    // Build complete symbol map (known + user answered)
    const allKnownBlocks = { ...(knownBlocks || {}), ...(userAnswers?.blocks || {}) };
    const allKnownLayers = { ...(knownLayers || {}), ...(userAnswers?.layers || {}) };

    // Use stored dxfData or re-parse if dxfContent provided
    let civilData = dxfData;
    if (!civilData && dxfContent) {
      const parsed = parseDXF(dxfContent);
      civilData = extractCivilData(parsed, filename);
    }
    if (!civilData) return res.status(400).json({ error: 'No drawing data.' });

    // Build symbol summary for Gemini
    const symbolSummary = [
      ...Object.entries(allKnownBlocks).map(([name, type]) =>
        `Block "${name}" (×${civilData.block_counts?.[name] || '?'}) = ${type}`),
      ...Object.entries(allKnownLayers).map(([name, type]) =>
        `Layer "${name}" = ${type}`)
    ].join('\n');

    const { RATES: ratesMap } = require('./src/dxf_parser');
    const ratesSummary = Object.entries(ratesMap).slice(0, 30).map(([k, v]) => `${k}:${v}`).join(', ');

    const prompt = `You are a senior PMC civil engineer generating a complete BOQ.
ALL DATA IS FROM THIS DXF FILE. DO NOT INVENT VALUES.

FILE: ${filename}
DRAWING TYPE: ${civilData.drawing_type}
SCALE: ${civilData.scale || 'not detected'}
UNITS: ${civilData.units}
DRAWING SIZE: ${civilData.drawing_extents.width_m}m × ${civilData.drawing_extents.height_m}m

SYMBOL DICTIONARY (confirmed by user + AI):
${symbolSummary || 'none'}

ELEMENT COUNTS:
Doors: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='door').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.door_count||0}
Windows: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='window').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.window_count||0}
Columns: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='column').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.column_count||0}
Lifts: ${civilData.element_counts?.lift_count||0}
Staircases: ${civilData.element_counts?.staircase_count||0}
Floors: ${civilData.element_counts?.floor_count||0}
Wall length: ${civilData.wall_length_m||0}m

FLOOR LEVELS:
${(civilData.floor_levels||[]).map(l=>`${l.label}=${l.level_m||'?'}m`).join('\n')||'none'}

TEXT ANNOTATIONS:
${civilData.all_texts.slice(0,100).join('\n')}

ROOM LABELS: ${(civilData.room_annotations||[]).map(r=>r.text).join(', ')||'none'}

DIMENSIONS (top 40): ${civilData.dimension_values.slice(0,40).map(d=>`${d.value_m}m[${d.layer}]`).join(', ')}

AREAS from polylines: ${civilData.polyline_areas.slice(0,20).map(p=>`${p.area_sqm}sqm(${p.layer})`).join(', ')}

GUJARAT DSR 2025 RATES:
${ratesSummary}

Generate complete BOQ. Return ONLY raw JSON:
{
  "project_name": "",
  "drawing_type": "",
  "scale": "",
  "building_height_m": 0,
  "floor_count": 0,
  "total_bua_sqm": 0,
  "spaces": [{"name":"","area_sqm":0}],
  "boq": [
    {"sr":1,"description":"","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed"}
  ],
  "element_counts": {"door_count":0,"window_count":0,"lift_count":0,"staircase_count":0,"column_count":0},
  "observations": [],
  "pmc_recommendation": ""
}`;

    const gData = await fetchGeminiWithRetry(key, {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' }
    });

    let raw = gData?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
    let geminiResult = {};
    if (fb !== -1) {
      try { geminiResult = JSON.parse(raw.slice(fb, lb + 1)); } catch(e) {}
    }

    res.json({ success: true, interpretation: geminiResult, dxf_data: civilData, learned_count: Object.keys(learned.blocks).length + Object.keys(learned.layers).length });

  } catch (err) {
    console.error('analyze-with-answers error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── 11. HEALTH ─────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  const key = process.env.GEMINI_API_KEY;
  res.json({ status: 'ok', gemini_key_set: !!key, key_preview: key ? key.slice(0, 8) + '...' : 'NOT SET' });
});

const APP_URL = process.env.RENDER_EXTERNAL_URL;
if (APP_URL) setInterval(() => fetch(APP_URL + '/health').catch(() => {}), 14 * 60 * 1000);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ PMC Civil AI Agent on port ${PORT}`);
  console.log(`🔑 GEMINI_API_KEY: ${process.env.GEMINI_API_KEY ? 'SET ✅' : 'NOT SET ❌'}`);
});

