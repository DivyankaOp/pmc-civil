const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const GEMINI_URL = (key) =>
  `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;

// ── GEMINI CHAT ENDPOINT
app.post('/gemini', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set in environment.' });
    const { body } = req.body;
    if (!body) return res.status(400).json({ error: 'No body provided.' });
    const response = await fetch(GEMINI_URL(key), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const data = await response.json();
    return res.status(response.status).json(data);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// ── EXCEL EXPORT ENDPOINT
app.post('/export-excel', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) return res.status(500).json({ error: 'GEMINI_API_KEY not set.' });
    const { files, userText, mode } = req.body;

    // Step 1: Ask Gemini to extract structured JSON from uploaded files
    const parts = [];
    if (files && files.length) {
      for (const f of files) {
        if (f.type === 'application/pdf' || f.name.match(/\.pdf$/i)) {
          parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
        } else if (f.type.startsWith('image/')) {
          parts.push({ inline_data: { mime_type: f.type, data: f.b64 } });
        }
      }
    }

    const extractPrompt = `You are a PMC (Project Management Consultant) civil construction data extraction expert.

Analyze the uploaded file(s) and extract ALL data. Return ONLY valid JSON (no markdown, no explanation).

Return this exact JSON structure:
{
  "report_type": "comparison|estimation|boq|drawing|site_report|general",
  "project_title": "string",
  "company": "VCT BHARUCH or extract from document",
  "date": "DD-MM-YYYY",
  "summary": "2-3 line summary",
  
  "vendors": [
    {
      "name": "Agency/Vendor name",
      "contact": "contact number",
      "quote_date": "date",
      "brand": "brand name",
      "product_description": "full product description"
    }
  ],
  
  "pricing": {
    "old_rate": [
      { "label": "row label", "values": [v1, v2, v3, ...] }
    ],
    "new_rate": [
      { "label": "row label", "values": [v1, v2, v3, ...] }
    ]
  },
  
  "commercial_terms": [
    { "label": "term name", "values": [v1, v2, v3, ...] }
  ],
  
  "technical_specs": [
    { "label": "spec name", "values": [v1, v2, v3, ...] }
  ],
  
  "boq_items": [
    { "sr": 1, "description": "item", "unit": "m3", "qty": 10, "rate": 5000, "amount": 50000 }
  ],
  
  "recommendation": "PMC recommendation with reasons",
  "prepared_by": "PMC Civil AI Agent"
}

If a field is not applicable, use empty array []. Extract every number and detail accurately. For comparisons extract ALL vendors/options.`;

    parts.push({ text: extractPrompt + (userText ? '\n\nUser note: ' + userText : '') });

    const geminiRes = await fetch(GEMINI_URL(key), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ role: 'user', parts }],
        generationConfig: { maxOutputTokens: 8192, temperature: 0.1 }
      })
    });

    const geminiData = await geminiRes.json();
    let rawText = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    rawText = rawText.replace(/```json|```/g, '').trim();

    let data;
    try {
      data = JSON.parse(rawText);
    } catch (e) {
      return res.status(500).json({ error: 'Could not parse AI response as JSON: ' + rawText.slice(0, 200) });
    }

    // Step 2: Build Excel from extracted JSON
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('PMC Report');

    const NAVY    = '1F3864';
    const MIDBLUE = '2E75B6';
    const LTBLUE  = 'BDD7EE';
    const YELLOW  = 'FFD966';
    const GREEN   = 'E2EFDA';
    const DKGREEN = '375623';
    const GREY    = 'F2F2F2';
    const WHITE   = 'FFFFFF';
    const RED     = 'C00000';

    const thin = { style: 'thin', color: { argb: 'FF000000' } };
    const allBorders = { top: thin, left: thin, bottom: thin, right: thin };

    function styleCell(cell, opts = {}) {
      if (opts.bg)   cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + opts.bg } };
      if (opts.bold !== undefined || opts.color || opts.size || opts.italic) {
        cell.font = {
          bold: opts.bold || false,
          italic: opts.italic || false,
          color: { argb: 'FF' + (opts.color || '000000') },
          size: opts.size || 10,
          name: 'Calibri'
        };
      }
      cell.alignment = {
        horizontal: opts.align || 'left',
        vertical: 'middle',
        wrapText: opts.wrap !== false
      };
      if (opts.border !== false) cell.border = allBorders;
    }

    function addFullRow(rowNum, text, bg, color = 'FFFFFF', size = 12, bold = true) {
      ws.mergeCells(rowNum, 1, rowNum, vendorCount + 2);
      const cell = ws.getCell(rowNum, 1);
      cell.value = text;
      styleCell(cell, { bg, color, bold, size, align: 'center', border: true });
      ws.getRow(rowNum).height = 20;
    }

    function addRow(rowNum, srVal, label, values, bg, labelBold = false) {
      const srCell = ws.getCell(rowNum, 1);
      srCell.value = srVal;
      styleCell(srCell, { bg: bg || GREY, align: 'center', border: true });

      const lblCell = ws.getCell(rowNum, 2);
      lblCell.value = label;
      styleCell(lblCell, { bg: bg || WHITE, bold: labelBold, border: true });

      values.forEach((v, i) => {
        const c = ws.getCell(rowNum, i + 3);
        c.value = v;
        styleCell(c, { bg: bg || WHITE, align: 'center', border: true });
      });
      ws.getRow(rowNum).height = 16;
    }

    const vendors = data.vendors || [];
    const vendorCount = Math.max(vendors.length, 1);
    let row = 1;

    // Set column widths
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 30;
    for (let i = 3; i <= vendorCount + 2; i++) ws.getColumn(i).width = 26;

    // ── TITLE
    addFullRow(row++, data.company || 'VCT BHARUCH', NAVY, 'FFFFFF', 14, true);
    addFullRow(row++, (data.project_title || 'COMPARISON REPORT').toUpperCase(), MIDBLUE, 'FFFFFF', 12, true);

    // ── VENDOR HEADERS
    const hdrRow = ws.getRow(row);
    hdrRow.getCell(1).value = 'SR NO';
    styleCell(hdrRow.getCell(1), { bg: NAVY, color: 'FFFFFF', bold: true, align: 'center', border: true });
    hdrRow.getCell(2).value = 'PARTICULARS';
    styleCell(hdrRow.getCell(2), { bg: NAVY, color: 'FFFFFF', bold: true, align: 'center', border: true });
    vendors.forEach((v, i) => {
      const c = hdrRow.getCell(i + 3);
      c.value = `${v.name || ''}\n(${v.brand || ''})\n${v.quote_date || ''}`;
      styleCell(c, { bg: NAVY, color: 'FFFFFF', bold: true, align: 'center', border: true });
    });
    hdrRow.height = 50;
    row++;

    // ── VENDOR INFO
    const infoRows = [
      ['', 'AGENCY NAME', vendors.map(v => v.name || '')],
      ['', 'CONTACT NO',  vendors.map(v => v.contact || '')],
      ['', 'DATE OF QUOTATION', vendors.map(v => v.quote_date || '')],
      ['', 'BRAND',       vendors.map(v => v.brand || '')],
    ];
    infoRows.forEach(([sr, lbl, vals], idx) => {
      addRow(row++, sr, lbl, vals, idx % 2 === 0 ? LTBLUE : GREY, true);
    });

    // ── PRODUCT DESCRIPTION
    if (vendors.some(v => v.product_description)) {
      addFullRow(row++, 'PRODUCT DESCRIPTION', MIDBLUE, 'FFFFFF', 10, true);
      const pdRow = ws.getRow(row);
      pdRow.getCell(1).value = '1';
      styleCell(pdRow.getCell(1), { bg: WHITE, align: 'center', border: true });
      pdRow.getCell(2).value = 'PRODUCT DESCRIPTION';
      styleCell(pdRow.getCell(2), { bg: WHITE, bold: true, border: true });
      vendors.forEach((v, i) => {
        const c = pdRow.getCell(i + 3);
        c.value = v.product_description || '';
        styleCell(c, { bg: WHITE, align: 'left', border: true, size: 9 });
      });
      pdRow.height = 80;
      row++;
    }

    // ── PRICING OLD
    if (data.pricing?.old_rate?.length) {
      addFullRow(row++, 'PRICING — OLD RATE', NAVY, 'FFFFFF', 10, true);
      data.pricing.old_rate.forEach((r, idx) => {
        const isTotal = r.label?.toUpperCase().includes('TOTAL');
        const bg = isTotal ? YELLOW : (idx % 2 === 0 ? WHITE : GREY);
        addRow(row++, '', r.label, r.values || [], bg, isTotal);
      });
    }

    // ── PRICING NEW
    if (data.pricing?.new_rate?.length) {
      addFullRow(row++, 'PRICING — NEW RATE (LATEST)', NAVY, 'FFFFFF', 10, true);
      const totals = [];
      data.pricing.new_rate.forEach((r, idx) => {
        const isTotal   = r.label?.toUpperCase().includes('TOTAL');
        const isDisc    = r.label?.toUpperCase().includes('DISCOUNT');
        const bg = isTotal ? YELLOW : isDisc ? GREEN : (idx % 2 === 0 ? WHITE : GREY);
        if (isTotal) totals.push(...(r.values || []));
        addRow(row++, '', r.label, r.values || [], bg, isTotal);
      });

      // Highlight lowest
      if (totals.length > 0) {
        const nums = totals.map(v => parseFloat(String(v).replace(/[^0-9.]/g, '')) || 0);
        const minVal = Math.min(...nums.filter(n => n > 0));
        addFullRow(row++, 'LOWEST QUOTED PRICE', MIDBLUE, 'FFFFFF', 10, true);
        const lowRow = ws.getRow(row);
        lowRow.getCell(1).value = '';
        styleCell(lowRow.getCell(1), { bg: GREEN, border: true });
        lowRow.getCell(2).value = 'TOTAL WITH GST (✓ = LOWEST)';
        styleCell(lowRow.getCell(2), { bg: GREEN, bold: true, border: true });
        nums.forEach((n, i) => {
          const c = lowRow.getCell(i + 3);
          const isLowest = n === minVal && n > 0;
          c.value = n > 0 ? `₹${n.toLocaleString('en-IN')}${isLowest ? ' ✓ LOWEST' : ''}` : '—';
          styleCell(c, {
            bg: isLowest ? '00B050' : WHITE,
            color: isLowest ? 'FFFFFF' : '000000',
            bold: isLowest,
            align: 'center',
            border: true
          });
        });
        lowRow.height = 20;
        row++;
      }
    }

    // ── BOQ ITEMS
    if (data.boq_items?.length) {
      addFullRow(row++, 'BILL OF QUANTITIES', NAVY, 'FFFFFF', 11, true);
      const boqHdr = ws.getRow(row++);
      ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)'].forEach((h, i) => {
        const c = boqHdr.getCell(i + 1);
        c.value = h;
        styleCell(c, { bg: MIDBLUE, color: 'FFFFFF', bold: true, align: 'center', border: true });
      });
      let total = 0;
      data.boq_items.forEach((item, idx) => {
        const bg = idx % 2 === 0 ? WHITE : GREY;
        const r = ws.getRow(row++);
        [item.sr, item.description, item.unit, item.qty, item.rate, item.amount].forEach((v, i) => {
          const c = r.getCell(i + 1);
          c.value = v;
          styleCell(c, { bg, align: i === 0 ? 'center' : i > 1 ? 'center' : 'left', border: true });
        });
        total += parseFloat(item.amount) || 0;
      });
      const totRow = ws.getRow(row++);
      ws.mergeCells(row - 1, 1, row - 1, 4);
      totRow.getCell(1).value = 'TOTAL';
      styleCell(totRow.getCell(1), { bg: YELLOW, bold: true, align: 'right', border: true });
      totRow.getCell(5).value = '';
      styleCell(totRow.getCell(5), { bg: YELLOW, border: true });
      totRow.getCell(6).value = total;
      styleCell(totRow.getCell(6), { bg: YELLOW, bold: true, align: 'center', border: true });
    }

    // ── COMMERCIAL TERMS
    if (data.commercial_terms?.length) {
      addFullRow(row++, 'COMMERCIAL TERMS', NAVY, 'FFFFFF', 10, true);
      data.commercial_terms.forEach((t, idx) => {
        const bg = idx % 2 === 0 ? WHITE : GREY;
        addRow(row++, '', t.label, t.values || [], bg, true);
        ws.getRow(row - 1).height = 40;
      });
    }

    // ── TECHNICAL SPECS
    if (data.technical_specs?.length) {
      addFullRow(row++, 'TECHNICAL SPECIFICATIONS', NAVY, 'FFFFFF', 10, true);
      data.technical_specs.forEach((s, idx) => {
        const bg = idx % 2 === 0 ? WHITE : GREY;
        addRow(row++, String(idx + 1), s.label, s.values || [], bg, true);
      });
    }

    // ── PMC RECOMMENDATION
    if (data.recommendation) {
      addFullRow(row++, 'PMC RECOMMENDATION', DKGREEN, 'FFFFFF', 11, true);
      ws.mergeCells(row, 1, row, vendorCount + 2);
      const recCell = ws.getCell(row, 1);
      recCell.value = data.recommendation;
      styleCell(recCell, { bg: 'E2EFDA', bold: false, align: 'left', size: 10, border: true });
      ws.getRow(row).height = 80;
      row++;
    }

    // ── SUMMARY
    if (data.summary) {
      ws.mergeCells(row, 1, row, vendorCount + 2);
      const sumCell = ws.getCell(row, 1);
      sumCell.value = 'SUMMARY: ' + data.summary;
      styleCell(sumCell, { bg: LTBLUE, bold: false, align: 'left', size: 9, italic: true, border: true });
      ws.getRow(row).height = 30;
      row++;
    }

    // ── FOOTER
    ws.mergeCells(row, 1, row, vendorCount + 2);
    const footCell = ws.getCell(row, 1);
    const today = new Date().toLocaleDateString('en-IN', { day:'2-digit', month:'2-digit', year:'numeric' });
    footCell.value = `Prepared by: PMC Civil AI Agent  |  Date: ${today}  |  Powered by Gemini AI`;
    styleCell(footCell, { bg: GREY, color: '595959', align: 'center', size: 9, italic: true, bold: false, border: false });
    ws.getRow(row).height = 14;

    // Freeze header rows
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];

    // Send as Excel file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="PMC_Report.xlsx"');
    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  const key = process.env.GEMINI_API_KEY;
  res.json({ status: 'ok', gemini_key_set: !!key, key_preview: key ? key.slice(0, 8) + '...' : 'NOT SET' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  const key = process.env.GEMINI_API_KEY;
  console.log(`\n✅ PMC Civil AI Agent running on port ${PORT}`);
  console.log(`🔑 GEMINI_API_KEY: ${key ? 'SET ✅' : 'NOT SET ❌'}`);
});
