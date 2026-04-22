const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');
const ExcelJS = require('exceljs');
const { extractDrawingData, buildDrawingExcel } = require('./server_drawing');
const { geminiAnalyzeDrawing, runCVAnalysis, RATES } = require('./drawing_analyzer');
const { parseDXF, extractCivilData, extractTotalAreaSqft } = require('./dxf_parser');
const { buildExcelFromDrawing, getDrawingPrompt } = require('./drawing_to_excel');
const { buildDXFExcel } = require('./dxf_to_excel');
const { buildEstimateWorkbook, RATES: ESTIMATE_RATES, DEFAULT_CATEGORIES } = require('./estimate_builder');

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
app.post('/analyze-dxf', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { dxfContent, filename } = req.body;

    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    // Step 1: Parse DXF — extract everything dynamically from drawing
    const parsed = parseDXF(dxfContent);
    const civilData = extractCivilData(parsed, filename);
    // Attach raw positioned texts for Excel (x,y coordinates per annotation)
    civilData._raw_texts = parsed.texts.map(t => ({ text: t.text, layer: t.layer, x: t.x, y: t.y }));

    // Step 2: Build rich Gemini prompt from ACTUAL extracted data (no hardcoded values)
    const { RATES: ratesMap } = require('./dxf_parser');
    const ratesSummary = Object.entries(ratesMap).slice(0, 30).map(([k,v]) => `${k}:${v}`).join(', ');

    const prompt = `You are a senior PMC civil engineer analyzing a DXF drawing.
ALL DATA BELOW IS EXTRACTED DIRECTLY FROM THE DXF FILE. DO NOT INVENT VALUES.

FILE: ${filename}
TYPE: ${civilData.drawing_type}
SCALE: ${civilData.scale || 'not detected'}
UNITS: ${civilData.units}
SIZE: ${civilData.drawing_extents.width_m}m x ${civilData.drawing_extents.height_m}m

TEXT ANNOTATIONS (${civilData.stats.total_texts}):
${civilData.all_texts.slice(0,150).join('\n')}

ROOM LABELS: ${(civilData.room_annotations||[]).map(r=>r.text).join(', ')||'none'}

DIMENSIONS (top 60): ${civilData.dimension_values.slice(0,60).map(d=>`${d.value_m}m[${d.layer}]`).join(', ')}

AREAS from polylines (${civilData.polyline_areas.length}): ${civilData.polyline_areas.slice(0,30).map(p=>`${p.area_sqm}sqm(${p.layer})`).join(', ')}

LAYERS: ${civilData.layer_names.join(', ')}
BLOCKS: ${Object.entries(civilData.block_counts||{}).slice(0,20).map(([k,v])=>`${k}x${v}`).join(', ')||'none'}
RATES AVAILABLE: ${ratesSummary}

Return ONLY raw JSON (no markdown):
{"project_name":"from texts above","drawing_type":"FLOOR_PLAN|SECTION|ELEVATION|SITE_PLAN|ROAD_PLAN|STRUCTURAL|GENERAL","scale":"detected or not","date":"from texts","spaces":[{"name":"room name from text","area_sqm":0}],"boq":[{"description":"from drawing data only","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0}],"total_bua_sqm":0,"observations":["based on actual data"],"pmc_recommendation":"based on actual extracted data"}`;

    const geminiData3 = await fetchGeminiWithRetry(key, {
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 4096, temperature: 0.0, responseMimeType: 'application/json' }
      });
    let raw = geminiData3?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
    let geminiResult = {};
    if (fb !== -1) try { geminiResult = JSON.parse(raw.slice(fb, lb+1)); } catch(e) {}

    // Return parsed data + gemini interpretation
    res.json({ success: true, dxf_data: civilData, interpretation: geminiResult });

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
      const { RATES: rMap } = require('./dxf_parser');
      const rSummary = Object.entries(rMap).slice(0,20).map(([k,v])=>`${k}:${v}`).join(',');
      const prompt = `PMC civil engineer. Analyze DXF. Use ONLY data below, no invented values.
FILE:${filename} TYPE:${civilData.drawing_type} SCALE:${civilData.scale||'?'}
TEXTS:${civilData.all_texts.slice(0,100).join(' | ')}
DIMS:${civilData.dimension_values.slice(0,30).map(d=>d.value_m+'m['+d.layer+']').join(', ')}
AREAS:${civilData.polyline_areas.slice(0,15).map(p=>p.area_sqm+'sqm('+p.layer+')').join(', ')}
LAYERS:${civilData.layer_names.join(', ')}
RATES:${rSummary}
Return ONLY JSON:{"project_name":"","drawing_type":"","scale":"","spaces":[],"boq":[{"description":"","unit":"","qty":0,"rate":0,"amount":0}],"observations":[],"pmc_recommendation":""}`;

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

    const estimatePath = path.join(__dirname, 'UPDATED-OVERALL-ESTIMATE-MODESTAA-10.04.2026.xlsx');
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

// ─── DWG/DXF ANALYSIS — Convert to PNG + Gemini Vision ───────────
// Strategy: dwg_converter.py renders DXF/DWG to PNG using ezdxf+matplotlib
// Then Gemini SEES the actual drawing like a human engineer
app.post('/analyze-dwg', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });

    const { b64, filename } = req.body;
    if (!b64) return res.status(400).json({ error: 'No file data provided.' });

    const fs = require('fs');
    const { execSync } = require('child_process');
    const os = require('os');

    // Write uploaded file to temp
    const ext = filename?.match(/\.(dxf|dwg)$/i)?.[1]?.toLowerCase() || 'dxf';
    const tmpIn  = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.${ext}`);
    const tmpPng = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.png`);

    fs.writeFileSync(tmpIn, Buffer.from(b64, 'base64'));

    // Run Python converter
    const scriptPath = path.join(__dirname, 'dwg_converter.py');
    let converterResult = {};
    try {
      // Try python3 first, fall back to python (Render/Windows compatibility)
      let out;
      try {
        out = execSync(`python3 "${scriptPath}" "${tmpIn}" "${tmpPng}"`, { timeout: 180000 });
      } catch (e1) {
        out = execSync(`python "${scriptPath}" "${tmpIn}" "${tmpPng}"`, { timeout: 180000 });
      }
      converterResult = JSON.parse(out.toString());
    } catch (e) {
      converterResult = { success: false, error: e.message, errors: [e.message] };
    }
    if ((converterResult.errors || []).length) {
      console.warn('dwg_converter errors:', converterResult.errors);
    }

    // Build Gemini parts
    const parts = [];

    // If PNG rendered successfully — send image to Gemini Vision
    // FIX-8: Blank PNG guard — prevent Gemini hallucinating on empty/failed renders
    if (converterResult.png_path && fs.existsSync(converterResult.png_path)) {
      const pngBuf = fs.readFileSync(converterResult.png_path);
      const pngSizeKB = pngBuf.length / 1024;
      if (pngSizeKB < 30) {
        console.warn('[analyze-dwg] PNG too small (' + pngSizeKB.toFixed(1) + 'KB) — likely blank render, skipping vision.');
        converterResult.png_path = null;
      } else {
        parts.push({ inline_data: { mime_type: 'image/png', data: pngBuf.toString('base64') } });
      }
      try { if (converterResult.png_path) fs.unlinkSync(converterResult.png_path); } catch(e) {}
    }

    // Always include extracted text + dimension data
    const textSummary = (converterResult.texts || []).map(t => t.text).slice(0, 150).join(' | ');
    const dimSummary  = (converterResult.dimensions || [])
      .filter(d => d.value).map(d => `${d.value}${d.text ? ' ('+d.text+')' : ''}`).slice(0, 80).join(', ');
    const layers = (converterResult.layers || []).join(', ');

    const prompt = `You are a SENIOR PMC CIVIL ENGINEER with 20 years India experience.
${converterResult.png_path ? 'The image above is the actual rendered AutoCAD drawing — analyze it directly with full vision.' : 'The drawing image could not be rendered, use the extracted data below.'}

FILE: ${filename}
DRAWING TYPE (auto-detected): ${converterResult.drawing_type || 'Unknown'}
SCALE (from drawing): ${converterResult.scale || 'Not detected — estimate from dimensions'}
DRAWING EXTENTS: ${JSON.stringify(converterResult.extents || {})}

LAYERS IN DRAWING: ${layers || 'None extracted'}

ALL TEXT ANNOTATIONS (${(converterResult.texts||[]).length} found):
${textSummary || 'None extracted'}

DIMENSION VALUES (${(converterResult.dimensions||[]).length} found):
${dimSummary || 'None extracted'}

${converterResult.error ? 'NOTE: File reading had issues: ' + converterResult.error : ''}

INSTRUCTIONS:
1. If image is shown above — read EVERY dimension, annotation, scale bar, title block directly from it
2. Use extracted text + dimensions above to supplement/verify what you see
3. Calculate ALL quantities using proper PMC formulas:
   Roads: Area=L×W | GSB=Area×1.15×0.3×1800kg/m³ | WMM=Area×1.15×0.2×2100 | PQC=Area×1.05×0.25
   Structure: Volume=L×W×H | Steel=Volume×120kg/m³(slab) or 160(beam)
4. Apply Gujarat DSR 2025 rates (from config — do NOT override unless drawing specifies a different rate):
   ${Object.values(RATES).map(v => `${v.desc}: ₹${v.rate}/${v.unit}`).join(' | ')}

OUTPUT FORMAT — Full PMC Analysis Report:
## Drawing Details (Title Block)
## Scale & Dimensions
## Element-wise Quantities Table
| Element | Length(m) | Width(m) | Area(sqmt) | Qty | Unit | Rate(₹) | Amount(₹) |
## Cost Summary
## Steel BBS (if structural drawing)
## PMC Observations & IS Code References`;

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

    res.json({ success: true, analysis, converter: converterResult });
  } catch (err) {
    console.error('DWG analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── FULL ESTIMATE — universal upload → MODESTAA-format workbook ───
// Accepts any mix of image / PDF / DXF / DWG files. Extracts structured
// quantities via Gemini Vision using a schema that mirrors the reference
// annexure layout, then writes a multi-sheet workbook with formulas
// (no hard-coded rates — all rates come from Rates.json by default, and
// Gemini may override per-item when it reads them off the drawing itself).
app.post('/full-estimate', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });

    const { files = [], userText = '', projectName = '' } = req.body;
    if (!files.length) return res.status(400).json({ error: 'No files uploaded.' });

    const fs = require('fs');
    const os = require('os');
    const { execSync } = require('child_process');

    // Build Gemini content parts. For DXF/DWG, convert to PNG via dwg_converter.py first.
    const parts = [];
    const tmpPaths = [];
    const extractedText = [];

    for (const f of files) {
      const name = (f.name || '').toLowerCase();
      const mime = f.type || '';
      try {
        if (name.endsWith('.dxf') || name.endsWith('.dwg')) {
          const ext = name.endsWith('.dwg') ? 'dwg' : 'dxf';
          const tmpIn  = path.join(os.tmpdir(), `pmc_est_${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`);
          const tmpPng = path.join(os.tmpdir(), `pmc_est_${Date.now()}_${Math.random().toString(36).slice(2)}.png`);
          fs.writeFileSync(tmpIn, Buffer.from(f.b64, 'base64'));
          tmpPaths.push(tmpIn);
          let cvt = {};
          try {
            const scriptPath = path.join(__dirname, 'dwg_converter.py');
            let out;
            try {
              out = execSync(`python3 "${scriptPath}" "${tmpIn}" "${tmpPng}"`, { timeout: 180000 });
            } catch (e1) {
              out = execSync(`python "${scriptPath}" "${tmpIn}" "${tmpPng}"`, { timeout: 180000 });
            }
            cvt = JSON.parse(out.toString());
          } catch (e) { cvt = { success: false, error: e.message, errors:[e.message] }; }
          if ((cvt.errors || []).length) console.warn('dwg_converter', f.name, cvt.errors);

          // FIX-8b: Blank PNG guard for /full-estimate too
          if (cvt.png_path && fs.existsSync(cvt.png_path)) {
            const pngBuf2 = fs.readFileSync(cvt.png_path);
            if (pngBuf2.length / 1024 >= 30) {
              parts.push({ inline_data: { mime_type: 'image/png', data: pngBuf2.toString('base64') } });
            } else {
              console.warn('[full-estimate] PNG too small (' + (pngBuf2.length/1024).toFixed(1) + 'KB) — skipping vision for this file.');
            }
            tmpPaths.push(cvt.png_path);
          }
          const txts = (cvt.texts || []).map(t => t.text).slice(0, 150).join(' | ');
          const dims = (cvt.dimensions || []).filter(d => d.value)
            .map(d => `${d.value}${d.text ? ' (' + d.text + ')' : ''}`).slice(0, 80).join(', ');
          extractedText.push(
            `FILE: ${f.name}\nTYPE: ${cvt.drawing_type || 'Unknown'}\nLAYERS: ${(cvt.layers||[]).join(', ')}\nTEXTS: ${txts}\nDIMENSIONS: ${dims}`
          );
        } else if (mime === 'application/pdf' || name.endsWith('.pdf')) {
          parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
        } else if (mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp)$/i.test(name)) {
          parts.push({ inline_data: { mime_type: mime || 'image/png', data: f.b64 } });
        }
      } catch (e) { console.log('skip file', f.name, e.message); }
    }

    // Rates context (no hard-coded rates in the prompt — rely on Rates.json)
    const rateHints = JSON.stringify(ESTIMATE_RATES, null, 0).slice(0, 6000);

    const schemaPrompt = `You are a SENIOR PMC CIVIL ENGINEER preparing a project estimate from construction drawings.
You have been given one or more drawings (images / PDF / rendered DXF-DWG). READ every dimension, title block, legend and annotation.

Your job: RETURN JSON ONLY (no prose, no markdown) matching EXACTLY this schema:
{
  "project_name": "string — read from title block, else infer",
  "drawing_type": "string — floor plan / site layout / section / foundation / structural / tiles / elevation / mixed",
  "builtup_area_sqft": number,           // total BUA in SQFT. Compute from dimensions if not printed.
  "sections": [
    {
      "title": "string — pick from: EXCAVATION, CIVIL WORK, TILES & STONE WORK, PLUMBING & WATERPROOFING, MISCELLANEOUS WORK, AMENITIES, CONSULTANT COST (you may add new sections if drawing needs them)",
      "items": [
        {
          "particular": "string — specific item e.g. 'RCC M25 slab', 'Brickwork 230mm external', 'Vitrified flooring 800x1600'",
          "qty": number,
          "unit": "string — CUM / SQMT / SQFT / RMT / NOS / KG / L/S",
          "rate": number,         // Rate without GST (₹). Use the RATES LIBRARY below as authoritative reference; only override if drawing specifies another rate.
          "gstPct": number        // as fraction: 0.18 for 18%, 0 for exempt. Default 0.18 if unknown.
        }
      ]
    }
  ],
  "observations": "string — 2-3 line PMC note about the drawing and any assumptions made."
}

RULES:
- Only include sections that are actually implied by the drawings. Do NOT invent items.
- For quantities: compute from dimensions using standard PMC formulas (L×W×H for volume, L×W for area, perimeter×height×thickness for wall volume, etc.).
- Every item MUST have a qty > 0. If you cannot compute a qty, SKIP the item.
- Use whole numbers or 2 decimals for qty. Units must be consistent within one item.
- Rates MUST come from the RATES LIBRARY below when available. If an item has no match, estimate conservatively using Gujarat DSR 2025 averages.
- DO NOT hard-code the builtup area — COMPUTE it from the drawing.
- Return ONLY the JSON object. Start with { end with }.

RATES LIBRARY (Gujarat DSR 2025, Rates.json):
${rateHints}

${extractedText.length ? `RAW EXTRACTED METADATA FROM DXF/DWG FILES:\n${extractedText.join('\n---\n')}\n` : ''}
${userText ? `USER NOTE: ${userText}\n` : ''}
${projectName ? `PROJECT NAME HINT: ${projectName}\n` : ''}`;

    // If every file was a DWG/DXF and we got no image + no text, abort early
    // with a helpful message — otherwise Gemini just sees an empty prompt.
    const hasVisual = parts.some(p => p.inline_data);
    const hasText = extractedText.some(t => /TEXTS: \S/.test(t) || /DIMENSIONS: \S/.test(t));
    if (!hasVisual && !hasText) {
      return res.status(400).json({
        error: 'Could not read the uploaded drawing(s). For DWG files the server needs LibreOffice or ODA File Converter installed, which is not available on this host. Please export your drawing to DXF, PDF, PNG or JPG and upload that instead.',
      });
    }

    parts.push({ text: schemaPrompt });

    const gem = await fetchGeminiWithRetry(key, {
      contents: [{ role: 'user', parts }],
      generationConfig: {
        maxOutputTokens: 16384,
        temperature: 0.1,
        responseMimeType: 'application/json',
      },
    });

    // Cleanup temp files
    for (const p of tmpPaths) { try { fs.unlinkSync(p); } catch {} }

    const raw = gem?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    let data = {};
    try {
      const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
      data = JSON.parse(raw.slice(fb, lb + 1));
    } catch (e) {
      console.error('Estimate JSON parse fail:', e.message, raw.slice(0, 400));
      return res.status(500).json({ error: 'AI returned invalid JSON. Try again with clearer drawings.' });
    }

    if (!data.sections || !data.sections.length) {
      return res.status(400).json({ error: 'No estimable items found in drawings.', raw: data });
    }

    const wb = await buildEstimateWorkbook(data);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const pname = (data.project_name || projectName || 'Drawing').replace(/[^a-zA-Z0-9_]/g, '_').slice(0, 24);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_MODESTAA_Estimate_${today}.xlsx"`);
    res.setHeader('X-Estimate-Meta', encodeURIComponent(JSON.stringify({
      project: data.project_name, bua: data.builtup_area_sqft,
      sections: data.sections.map(s => ({ title: s.title, items: s.items.length })),
    })));
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Full-estimate error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 9. HEALTH ─────────────────────────────────────────────────────
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
