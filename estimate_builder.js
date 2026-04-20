// estimate_builder.js
// Build a MODESTAA-style PMC civil estimate workbook from a structured items JSON.
// Output format mirrors the reference template:
//   - OVERALL SUMMARY sheet (per-sqft + with/without GST)
//   - One ANNEXURE-* sheet per category with columns:
//     SR NO | PARTICULAR | QTY | UNIT | RATE WITHOUT GST | TOTAL AMOUNT WITHOUT GST
//     | % GST | RATE WITH GST | GST AMOUNT | TOTAL AMOUNT WITH GST
//   - All amounts are real Excel FORMULAS (not hard-coded values).
// Categories, items and rates come from the caller (extracted by Gemini +
// Rates.json). Nothing about the project dimensions is hard-coded.

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const RATES = (() => {
  try {
    const p = fs.existsSync(path.join(__dirname, 'Rates.json'))
      ? path.join(__dirname, 'Rates.json')
      : path.join(__dirname, 'rates.json');
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch { return {}; }
})();

// Flat lookup: description (lowercased) -> {rate, unit}
const RATE_INDEX = (() => {
  const idx = {};
  for (const grp of Object.values(RATES)) {
    if (!grp || typeof grp !== 'object') continue;
    for (const [k, v] of Object.entries(grp)) {
      if (!v || typeof v !== 'object' || !('rate' in v)) continue;
      idx[(v.description || k).toLowerCase()] = { rate: v.rate, unit: v.unit };
      idx[k.toLowerCase()] = { rate: v.rate, unit: v.unit };
    }
  }
  return idx;
})();

function lookupRate(particular) {
  if (!particular) return null;
  const q = particular.toLowerCase();
  if (RATE_INDEX[q]) return RATE_INDEX[q];
  // Partial match — pick longest key contained in the particular
  let best = null, bestLen = 0;
  for (const key of Object.keys(RATE_INDEX)) {
    if (key.length > 4 && (q.includes(key) || key.includes(q)) && key.length > bestLen) {
      best = RATE_INDEX[key]; bestLen = key.length;
    }
  }
  return best;
}

// Default category order; only sheets with items are created.
const DEFAULT_CATEGORIES = [
  'EXCAVATION',
  'CIVIL WORK',
  'TILES & STONE WORK',
  'PLUMBING & WATERPROOFING',
  'MISCELLANEOUS WORK',
  'AMENITIES',
  'CONSULTANT COST',
];

const HEADERS = [
  'SR NO', 'PARTICULAR', 'QTY', 'UNIT', 'RATE WITHOUT GST',
  'TOTAL AMOUNT WITHOUT GST', '% GST', 'RATE WITH GST',
  'GST AMOUNT', 'TOTAL AMOUNT WITH GST',
];

function styleHeader(row) {
  row.eachCell(c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1F3A' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = allBorders();
  });
  row.height = 28;
}

function allBorders() {
  const s = { style: 'thin', color: { argb: 'FF999999' } };
  return { top: s, left: s, bottom: s, right: s };
}

function addBanner(ws, title, subtitle, cols = 10) {
  ws.mergeCells(1, 1, 1, cols);
  const r1 = ws.getCell(1, 1);
  r1.value = 'MODESTAA';
  r1.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1F3A' } };
  r1.alignment = { horizontal: 'center' };

  ws.mergeCells(2, 1, 2, cols);
  const r2 = ws.getCell(2, 1);
  r2.value = title;
  r2.font = { bold: true, size: 12 };
  r2.alignment = { horizontal: 'center' };
  r2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };

  if (subtitle) {
    ws.mergeCells(3, 1, 3, cols);
    const r3 = ws.getCell(3, 1);
    r3.value = subtitle;
    r3.font = { italic: true, size: 10, color: { argb: 'FF4B5E7A' } };
    r3.alignment = { horizontal: 'center' };
  }
}

/**
 * Build an annexure sheet.
 * @param ws worksheet
 * @param letter 'A','B',... for SR NO prefix (optional)
 * @param title section title
 * @param items [{particular, qty, unit, rate, gstPct}] — rate/gstPct optional, looked up if missing
 * @param bua built-up area in sqft (for COST PER SQFT formula)
 */
function writeAnnexure(ws, letter, title, items, bua) {
  addBanner(ws, `ANNEXURE-${letter} — ${title}`, bua ? `TOTAL AREA: ${bua} SQFT` : null);

  const headerRowIdx = 4;
  const hdr = ws.getRow(headerRowIdx);
  HEADERS.forEach((h, i) => (hdr.getCell(i + 1).value = h));
  styleHeader(hdr);

  const widths = [8, 44, 12, 10, 14, 18, 8, 14, 16, 20];
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  let rIdx = headerRowIdx + 1;
  const firstDataRow = rIdx;

  items.forEach((it, i) => {
    const row = ws.getRow(rIdx);
    const resolved = (it.rate == null || it.unit == null) ? lookupRate(it.particular) : null;
    const rate = it.rate ?? resolved?.rate ?? 0;
    const unit = it.unit ?? resolved?.unit ?? '';
    const gstPct = it.gstPct ?? 0.18;

    row.getCell(1).value = i + 1;
    row.getCell(2).value = it.particular || '';
    row.getCell(3).value = typeof it.qty === 'number' ? it.qty : (parseFloat(it.qty) || 0);
    row.getCell(4).value = unit;
    row.getCell(5).value = rate;
    row.getCell(6).value = { formula: `C${rIdx}*E${rIdx}` };         // amount without GST
    row.getCell(7).value = gstPct;
    row.getCell(8).value = { formula: `E${rIdx}*(1+G${rIdx})` };     // rate with GST
    row.getCell(9).value = { formula: `F${rIdx}*G${rIdx}` };         // GST amount
    row.getCell(10).value = { formula: `F${rIdx}+I${rIdx}` };        // total with GST

    row.eachCell({ includeEmpty: true }, c => (c.border = allBorders()));
    row.getCell(2).alignment = { wrapText: true, vertical: 'middle' };
    [3, 5, 6, 8, 9, 10].forEach(c => (row.getCell(c).numFmt = '#,##0.00'));
    row.getCell(7).numFmt = '0%';
    rIdx++;
  });

  const lastDataRow = rIdx - 1;

  // TOTAL row
  const totalRow = ws.getRow(rIdx);
  totalRow.getCell(2).value = 'TOTAL COST';
  totalRow.getCell(6).value = { formula: `SUM(F${firstDataRow}:F${lastDataRow})` };
  totalRow.getCell(9).value = { formula: `SUM(I${firstDataRow}:I${lastDataRow})` };
  totalRow.getCell(10).value = { formula: `SUM(J${firstDataRow}:J${lastDataRow})` };
  totalRow.eachCell({ includeEmpty: true }, c => {
    c.font = { bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FC' } };
    c.border = allBorders();
    c.numFmt = '#,##0.00';
  });
  const totalRowIdx = rIdx;
  rIdx++;

  // COST PER SQFT row
  if (bua && bua > 0) {
    const perRow = ws.getRow(rIdx);
    perRow.getCell(2).value = 'COST PER SQFT';
    perRow.getCell(6).value = { formula: `F${totalRowIdx}/'OVERALL SUMMARY'!$C$3` };
    perRow.getCell(9).value = { formula: `I${totalRowIdx}/'OVERALL SUMMARY'!$C$3` };
    perRow.getCell(10).value = { formula: `J${totalRowIdx}/'OVERALL SUMMARY'!$C$3` };
    perRow.eachCell({ includeEmpty: true }, c => {
      c.font = { italic: true };
      c.border = allBorders();
      c.numFmt = '#,##0.00';
    });
    rIdx++;
  }

  return { totalRowIdx };
}

function writeOverallSummary(ws, sectionRefs, bua, projectName) {
  // Banner takes rows 1-2 only so C3 stays free for the BUA value referenced
  // by all annexure "COST PER SQFT" formulas.
  ws.mergeCells(1, 1, 1, 9);
  const b1 = ws.getCell(1, 1);
  b1.value = 'MODESTAA';
  b1.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  b1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1F3A' } };
  b1.alignment = { horizontal: 'center' };

  ws.mergeCells(2, 1, 2, 9);
  const b2 = ws.getCell(2, 1);
  b2.value = `OVERALL SUMMARY — ${projectName || 'PMC Civil Estimate'}`;
  b2.font = { bold: true, size: 12 };
  b2.alignment = { horizontal: 'center' };
  b2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };

  ws.getCell('B3').value = 'BUILTUP AREA';
  ws.getCell('C3').value = bua || 0;
  ws.getCell('D3').value = 'SQFT';
  ws.getCell('B3').font = { bold: true };
  ws.getCell('C3').font = { bold: true };
  ws.getCell('C3').numFmt = '#,##0.00';

  const headers = [
    'SR NO', 'PARTICULARS', 'TOTAL AMOUNT WITHOUT GST',
    'RATE PER SQFT WITHOUT GST', 'GST AMOUNT', 'RATE PER SQFT GST',
    'TOTAL AMOUNT WITH GST', 'RATE PER SQFT WITH GST', 'REMARKS',
  ];
  const hdrRow = ws.getRow(5);
  headers.forEach((h, i) => (hdrRow.getCell(i + 1).value = h));
  styleHeader(hdrRow);
  [6, 40, 20, 18, 18, 18, 20, 18, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));

  let r = 6;
  const first = r;
  sectionRefs.forEach((s, i) => {
    const row = ws.getRow(r);
    row.getCell(1).value = i + 1;
    row.getCell(2).value = s.title;
    row.getCell(3).value = { formula: `'${s.sheet}'!F${s.totalRowIdx}` };
    row.getCell(4).value = { formula: `C${r}/$C$3` };
    row.getCell(5).value = { formula: `'${s.sheet}'!I${s.totalRowIdx}` };
    row.getCell(6).value = { formula: `E${r}/$C$3` };
    row.getCell(7).value = { formula: `'${s.sheet}'!J${s.totalRowIdx}` };
    row.getCell(8).value = { formula: `G${r}/$C$3` };
    row.eachCell({ includeEmpty: true }, c => {
      c.border = allBorders();
      c.numFmt = '#,##0.00';
    });
    r++;
  });
  const last = r - 1;

  const g = ws.getRow(r);
  g.getCell(2).value = 'GRAND TOTAL';
  g.getCell(3).value = { formula: `SUM(C${first}:C${last})` };
  g.getCell(4).value = { formula: `C${r}/$C$3` };
  g.getCell(5).value = { formula: `SUM(E${first}:E${last})` };
  g.getCell(6).value = { formula: `E${r}/$C$3` };
  g.getCell(7).value = { formula: `SUM(G${first}:G${last})` };
  g.getCell(8).value = { formula: `G${r}/$C$3` };
  g.eachCell({ includeEmpty: true }, c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1F3A' } };
    c.border = allBorders();
    c.numFmt = '#,##0.00';
  });
}

/**
 * Build the full workbook.
 * @param {Object} data
 *   data.project_name       - string
 *   data.builtup_area_sqft  - number
 *   data.sections           - [{ title, items: [{particular, qty, unit, rate?, gstPct?}] }]
 */
async function buildEstimateWorkbook(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';
  wb.created = new Date();

  const bua = Number(data.builtup_area_sqft) || 0;
  const project = data.project_name || 'PMC Civil Project';

  // Reserve overall summary sheet first so annexure formulas can reference it
  const wsOS = wb.addWorksheet('OVERALL SUMMARY', {
    properties: { tabColor: { argb: 'FF0B1F3A' } },
  });

  const sectionRefs = [];
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const sections = (data.sections || []).filter(s => (s.items || []).length);
  sections.forEach((sec, i) => {
    const letter = letters[i] || String(i + 1);
    const safe = sec.title.replace(/[\\/?*[\]]/g, '').slice(0, 25);
    const sheetName = `ANNEXURE-${letter} (${safe})`.slice(0, 31);
    const ws = wb.addWorksheet(sheetName);
    const { totalRowIdx } = writeAnnexure(ws, letter, sec.title, sec.items, bua);
    sectionRefs.push({ sheet: sheetName, title: sec.title, totalRowIdx });
  });

  writeOverallSummary(wsOS, sectionRefs, bua, project);

  return wb;
}

module.exports = { buildEstimateWorkbook, lookupRate, RATES, DEFAULT_CATEGORIES };
