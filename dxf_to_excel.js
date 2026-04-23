/**
 * PMC DXF → Excel Builder — v2.0
 * Builds a fully dynamic, multi-sheet Excel from DXF extracted data.
 * ZERO hardcoded rates — everything from rates.json via RATES object.
 * ZERO hardcoded room names — everything from DXF text annotations.
 */

'use strict';

const { RATES } = require('./dxf_parser');
const fs   = require('fs');
const path = require('path');

// Load full rates config for display (with descriptions + units)
let RATES_CONFIG = {};
try {
  RATES_CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, 'rates.json'), 'utf8'));
} catch(e) {}

// ── Excel style constants ────────────────────────────────────────
const C = {
  NAVY:    'FF1F3864', MIDBLUE: 'FF2E75B6', LTBLUE:  'FFBDD7EE',
  YELLOW:  'FFFFD966', GREEN:   'FFE2EFDA', DKGREEN: 'FF375623',
  GREY:    'FFF2F2F2', WHITE:   'FFFFFFFF', ORANGE:  'FFED7D31',
  TEAL:    'FF00B0F0', PURPLE:  'FF7030A0'
};
const thin = { style: 'thin', color: { argb: 'FF000000' } };
const bdr  = { top: thin, left: thin, bottom: thin, right: thin };

function sc(cell, bg, bold = false, fc = 'FF000000', size = 9, align = 'center', wrap = true) {
  if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  cell.font = { bold, color: { argb: fc }, size, name: 'Calibri' };
  cell.alignment = { horizontal: align, vertical: 'middle', wrapText: wrap };
  cell.border = bdr;
}

function mergeRow(ws, row, lastCol, text, bg, fc = 'FFFFFFFF', size = 10, height = 18) {
  ws.mergeCells(row, 1, row, lastCol);
  const c = ws.getCell(row, 1);
  c.value = text;
  sc(c, bg, true, fc, size, 'center');
  ws.getRow(row).height = height;
}

function hdrRow(ws, row, headers, bg = C.NAVY) {
  headers.forEach((h, i) => {
    const c = ws.getCell(row, i + 1);
    c.value = h;
    sc(c, bg, true, 'FFFFFFFF', 9, 'center');
  });
  ws.getRow(row).height = 30;
}

function dataRow(ws, row, values, bg = C.WHITE, align = 'center', bold = false) {
  values.forEach((v, i) => {
    const c = ws.getCell(row, i + 1);
    c.value = v;
    sc(c, bg, bold, 'FF000000', 9, typeof v === 'number' ? 'right' : align);
    if (typeof v === 'number' && v > 999) c.numFmt = '#,##0';
  });
  ws.getRow(row).height = 15;
}

// ─────────────────────────────────────────────────────────────────
async function buildDXFExcel(dxfData, geminiResult, ExcelJS) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';

  const gi   = geminiResult || {};
  const dd   = dxfData      || {};
  const today = new Date().toLocaleDateString('en-IN');
  const projectName = gi.project_name || dd.title_block?.project_name || dd.filename?.replace('.dxf','') || 'CIVIL PROJECT';
  const drawingType = gi.drawing_type || dd.drawing_type || 'GENERAL';

  // ══════════════════════════════════════════════════════════════
  // SHEET 1 — DRAWING SUMMARY
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('DRAWING SUMMARY');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 34;
    ws.getColumn(3).width = 38;
    ws.getColumn(4).width = 18;

    let row = 1;
    mergeRow(ws, row++, 4, projectName.toUpperCase(), C.NAVY, 'FFFFFFFF', 14, 26);
    mergeRow(ws, row++, 4, `DXF DRAWING ANALYSIS — ${drawingType}`, C.MIDBLUE, 'FFFFFFFF', 11, 20);
    mergeRow(ws, row++, 4, `PREPARED BY: PMC CIVIL AI AGENT  |  DATE: ${today}`, C.LTBLUE, 'FF1F3864', 9, 16);

    row++;
    mergeRow(ws, row++, 4, 'DRAWING INFORMATION', C.NAVY, 'FFFFFFFF', 10);
    hdrRow(ws, row++, ['SR', 'PARAMETER', 'VALUE', 'REMARKS']);

    const infoItems = [
      ['Filename',         dd.filename || '—',         ''],
      ['Drawing Type',     drawingType,                 ''],
      ['Scale',            dd.scale || 'Not detected',  ''],
      ['Units',            (dd.units || 'mm').toUpperCase(), ''],
      ['Drawing Width',    dd.drawing_extents?.width_m  ? `${dd.drawing_extents.width_m} m`  : '—', ''],
      ['Drawing Height',   dd.drawing_extents?.height_m ? `${dd.drawing_extents.height_m} m` : '—', ''],
      ['Total Layers',     dd.stats?.total_layers  || 0, ''],
      ['Total Texts',      dd.stats?.total_texts   || 0, 'Annotations & labels extracted'],
      ['Total Dimensions', dd.stats?.total_dims    || 0, 'Auto-measured from drawing'],
      ['Total Polylines',  dd.stats?.total_polylines|| 0, 'Closed regions → room areas'],
      ['Total Lines',      dd.stats?.total_lines   || 0, ''],
      ['Block Instances',  dd.stats?.total_inserts || 0, 'Doors, columns, windows etc.'],
      ['Unique Blocks',    dd.stats?.unique_blocks || 0, 'Block definitions used'],
    ];
    const tbMap = dd.title_block || {};
    if (tbMap.project_name) infoItems.push(['Project Name',  tbMap.project_name, '']);
    if (tbMap.drawn_by)     infoItems.push(['Drawn By',      tbMap.drawn_by,     '']);
    if (tbMap.date)         infoItems.push(['Drawing Date',  tbMap.date,         '']);
    if (tbMap.drawing_no)   infoItems.push(['Drawing No.',   tbMap.drawing_no,   '']);

    infoItems.forEach(([lbl, val, rem], idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      dataRow(ws, row, [idx + 1, lbl, val, rem], bg, 'left');
      row++;
    });

    // PMC Recommendation from Gemini
    row++;
    mergeRow(ws, row++, 4, 'PMC OBSERVATION', C.DKGREEN, 'FFFFFFFF', 10);
    ws.mergeCells(row, 1, row, 4);
    const rc = ws.getCell(row, 1);
    rc.value = gi.pmc_recommendation || 'Drawing successfully parsed. Review extracted data in subsequent sheets.';
    sc(rc, C.GREEN, false, 'FF000000', 9, 'left');
    ws.getRow(row).height = 50;
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 2 — ALL TEXT ANNOTATIONS (extracted directly from DXF)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('ANNOTATIONS');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 55;
    ws.getColumn(3).width = 20;
    ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 14;

    let row = 1;
    mergeRow(ws, row++, 5, 'ALL TEXT ANNOTATIONS — EXTRACTED FROM DXF', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 5, 'Every TEXT and MTEXT entity found in the drawing', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'TEXT CONTENT', 'LAYER', 'X (units)', 'Y (units)']);

    const texts = [
      ...(dd.all_texts || []).map((t, idx) => ({ text: t, layer: '—', x: 0, y: 0 }))
    ];

    // Use positioned texts if available from parsed data
    const positioned = [...(dxfData._raw_texts || [])];

    texts.forEach((t, idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      const pos = positioned[idx] || {};
      dataRow(ws, row++, [
        idx + 1,
        typeof t === 'string' ? t : t.text,
        pos.layer || '—',
        pos.x ? Math.round(pos.x) : '—',
        pos.y ? Math.round(pos.y) : '—'
      ], bg, 'left');
    });

    if (texts.length === 0) {
      mergeRow(ws, row, 5, 'No text annotations found in this DXF file.', C.GREY, 'FF595959', 9, 16);
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 3 — DIMENSIONS (from DIMENSION entities + text)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('DIMENSIONS');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 18;
    ws.getColumn(3).width = 18;
    ws.getColumn(4).width = 22;
    ws.getColumn(5).width = 18;

    let row = 1;
    mergeRow(ws, row++, 5, 'DIMENSION VALUES — AUTO-EXTRACTED FROM DXF', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 5, 'All DIMENSION entities measured from drawing geometry', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'VALUE (mm)', 'VALUE (m)', 'OVERRIDE TEXT', 'LAYER']);

    const dims = dd.dimension_values || [];
    dims.forEach((d, idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      dataRow(ws, row++, [idx + 1, d.value_mm, d.value_m, d.text || '—', d.layer || '—'], bg, 'center');
    });

    // Inline dimension annotations (e.g. "3000x4500" text)
    if ((dd.inline_dims || []).length > 0) {
      row++;
      mergeRow(ws, row++, 5, 'INLINE DIMENSION ANNOTATIONS (e.g. "3000x4500" in text)', C.TEAL, 'FFFFFFFF', 10);
      hdrRow(ws, row++, ['SR', 'LABEL TEXT', 'LENGTH (mm)', 'WIDTH (mm)', 'AREA (sqm)'], C.TEAL);
      dd.inline_dims.forEach((d, idx) => {
        dataRow(ws, row++, [idx + 1, d.label, d.length_mm, d.width_mm, d.area_sqm], idx % 2 === 0 ? C.WHITE : C.GREY, 'left');
      });
    }

    if (dims.length === 0 && (dd.inline_dims || []).length === 0) {
      mergeRow(ws, row, 5, 'No dimension entities found. Ensure DIMENSION entities exist in DXF.', C.GREY, 'FF595959', 9, 16);
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 4 — ROOM / SPACE AREAS (from closed polylines)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('AREAS');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 22;
    ws.getColumn(3).width = 16;
    ws.getColumn(4).width = 16;
    ws.getColumn(5).width = 16;
    ws.getColumn(6).width = 16;

    let row = 1;
    mergeRow(ws, row++, 6, 'SPACE / ROOM AREAS — FROM CLOSED POLYLINES', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 6, 'Area calculated using Shoelace formula on LWPOLYLINE vertices', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'LAYER', 'AREA (sqm)', 'AREA (sqft)', 'PERIMETER (m)', 'VERTICES']);

    const areas = dd.polyline_areas || [];
    let totalArea = 0;

    areas.forEach((a, idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      dataRow(ws, row++, [idx + 1, a.layer || '—', a.area_sqm, a.area_sqft, a.perimeter_m, a.vertices], bg, 'center');
      totalArea += a.area_sqm || 0;
    });

    if (areas.length > 0) {
      row++;
      const tc = ws.getCell(row, 1);
      ws.mergeCells(row, 1, row, 2);
      tc.value = 'TOTAL AREA'; sc(tc, C.YELLOW, true, 'FF000000', 9, 'right');
      const tv = ws.getCell(row, 3);
      tv.value = Math.round(totalArea * 100) / 100;
      tv.numFmt = '#,##0.00'; sc(tv, C.YELLOW, true, 'FF000000', 9, 'right');
      ws.getRow(row).height = 16;
    }

    // Room annotations from text
    if ((dd.room_annotations || []).length > 0) {
      row += 2;
      mergeRow(ws, row++, 6, 'ROOM LABELS (from text annotations in drawing)', C.TEAL, 'FFFFFFFF', 10);
      hdrRow(ws, row++, ['SR', 'LABEL TEXT', 'X (units)', 'Y (units)', 'LAYER', ''], C.TEAL);
      dd.room_annotations.forEach((r, idx) => {
        dataRow(ws, row++, [idx+1, r.text, Math.round(r.x||0), Math.round(r.y||0), r.layer||'—', ''], idx%2===0?C.WHITE:C.GREY, 'left');
      });
    }

    if (areas.length === 0) {
      mergeRow(ws, row, 6, 'No closed polylines found. Rooms must be drawn as closed LWPOLYLINE for area extraction.', C.GREY, 'FF595959', 9, 20);
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 5 — LAYERS SUMMARY
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('LAYERS');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 32;
    ws.getColumn(3).width = 12;
    ws.getColumn(4).width = 12;
    ws.getColumn(5).width = 12;
    ws.getColumn(6).width = 12;
    ws.getColumn(7).width = 12;

    let row = 1;
    mergeRow(ws, row++, 7, 'LAYER ANALYSIS — DXF LAYER STRUCTURE', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 7, 'Entity count per layer — useful for PMC layer audit', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'LAYER NAME', 'TEXTS', 'LINES', 'DIMS', 'POLYLINES', 'TOTAL']);

    const layerGroups = dd.layer_groups || {};
    const layerList = Object.entries(layerGroups)
      .map(([name, g]) => ({ name, ...g }))
      .sort((a, b) => (b.count || 0) - (a.count || 0));

    layerList.forEach((l, idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      dataRow(ws, row++, [
        idx + 1, l.name,
        l.texts?.length || 0,
        l.lines?.length || 0,
        l.dims?.length  || 0,
        l.polylines?.length || 0,
        l.count || 0
      ], bg, 'center');
    });
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 6 — BLOCK INSTANCES (doors, windows, columns etc.)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('BLOCK INSTANCES');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 38;
    ws.getColumn(3).width = 16;
    ws.getColumn(4).width = 30;

    let row = 1;
    mergeRow(ws, row++, 4, 'BLOCK INSTANCES — DOOR, WINDOW, COLUMN COUNTS', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 4, 'Auto-counted from INSERT entities in DXF', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'BLOCK NAME', 'COUNT (nos)', 'PROBABLE TYPE']);

    const blockCounts = dd.block_counts || {};
    const blockList = Object.entries(blockCounts).sort((a, b) => b[1] - a[1]);

    // Infer probable type from block name
    function inferBlockType(name) {
      const n = name.toLowerCase();
      if (/door|dr|d[0-9]/.test(n))          return 'Door';
      if (/window|wd|win|w[0-9]/.test(n))    return 'Window';
      if (/col|column|clm/.test(n))          return 'Column';
      if (/stair|st[0-9]/.test(n))           return 'Staircase';
      if (/lift|elev/.test(n))               return 'Lift / Elevator';
      if (/toilet|wc|bath/.test(n))          return 'Sanitary fixture';
      if (/sink|basin|wash/.test(n))         return 'Sanitary fixture';
      if (/tree|plant|shrub/.test(n))        return 'Landscape';
      if (/car|park|vehicle/.test(n))        return 'Parking';
      if (/bed|sofa|chair|furn/.test(n))     return 'Furniture';
      return 'General block';
    }

    blockList.forEach(([name, count], idx) => {
      const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
      dataRow(ws, row++, [idx + 1, name, count, inferBlockType(name)], bg, 'left');
    });

    if (blockList.length === 0) {
      mergeRow(ws, row, 4, 'No block instances found in this DXF.', C.GREY, 'FF595959', 9, 16);
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 7 — BOQ (from Gemini interpretation, no hardcoded items)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('BOQ ESTIMATE');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 42;
    ws.getColumn(3).width = 10;
    ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 14;
    ws.getColumn(6).width = 16;

    let row = 1;
    mergeRow(ws, row++, 6, 'BILL OF QUANTITIES — AI INTERPRETED FROM DRAWING', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 6, `Project: ${projectName}  |  Drawing: ${dd.filename || '—'}  |  Date: ${today}`, C.MIDBLUE, 'FFFFFFFF', 9, 16);
    mergeRow(ws, row++, 6, 'NOTE: Quantities extracted from DXF geometry + AI interpretation. Verify before use.', C.YELLOW, 'FF000000', 9, 16);

    hdrRow(ws, row++, ['SR', 'DESCRIPTION', 'UNIT', 'QTY', 'RATE (₹)', 'AMOUNT (₹)']);

    const boqItems = gi.boq || [];
    let grandTotal = 0;

    if (boqItems.length > 0) {
      boqItems.forEach((item, idx) => {
        const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
        const qty    = parseFloat(item.qty)    || 0;
        const rate   = parseFloat(item.rate)   || 0;
        const amount = parseFloat(item.amount) || (qty * rate);
        grandTotal += amount;
        const rc = ws.getRow(row);
        [idx+1, item.description, item.unit, qty, rate, Math.round(amount)].forEach((v, i) => {
          const c = rc.getCell(i+1); c.value = v;
          sc(c, bg, false, 'FF000000', 9, i < 2 ? 'left' : 'right');
          if (i >= 3 && typeof v === 'number') c.numFmt = '#,##0';
        });
        ws.getRow(row++).height = 15;
      });

      // Grand total row
      ws.mergeCells(row, 1, row, 4);
      const gt = ws.getCell(row, 1); gt.value = 'GRAND TOTAL';
      sc(gt, C.YELLOW, true, 'FF000000', 10, 'right');
      const gv = ws.getCell(row, 6); gv.value = Math.round(grandTotal);
      gv.numFmt = '₹#,##0'; sc(gv, C.YELLOW, true, 'FF000000', 10, 'right');
      ws.getRow(row).height = 18;
    } else {
      // Auto-generate BOQ from polyline areas if no Gemini BOQ available
      row = autoGenerateBOQ(ws, row, dd, today);
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 8 — RATES REFERENCE (from rates.json, fully editable)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('RATES REFERENCE');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 38;
    ws.getColumn(3).width = 12;
    ws.getColumn(4).width = 18;
    ws.getColumn(5).width = 22;

    let row = 1;
    mergeRow(ws, row++, 5, 'RATES REFERENCE — FROM rates.json CONFIG', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 5, 'Edit rates.json to update. No code change required.', C.MIDBLUE, 'FFFFFFFF', 9, 16);

    for (const [category, items] of Object.entries(RATES_CONFIG)) {
      if (category.startsWith('_') || typeof items !== 'object') continue;
      row++;
      const catLabel = category.toUpperCase().replace(/_/g, ' ');
      mergeRow(ws, row++, 5, catLabel, C.NAVY, 'FFFFFFFF', 10);
      hdrRow(ws, row++, ['SR', 'DESCRIPTION', 'UNIT', 'RATE (₹)', 'REMARKS']);
      let sr = 1;
      for (const [key, val] of Object.entries(items)) {
        const bg = sr % 2 === 0 ? C.WHITE : C.GREY;
        dataRow(ws, row++, [sr++, val.description || key, val.unit || '—', val.rate || 0, `rates.json → ${category}.${key}`], bg, 'left');
      }
    }

    if (Object.keys(RATES_CONFIG).filter(k => !k.startsWith('_')).length === 0) {
      mergeRow(ws, row, 5, 'rates.json not found or empty. Add rates.json to project folder.', C.GREY, 'FF595959', 9);
    }
  }

  return wb;
}

// ── Auto BOQ — driven by boq_rules.json, never hardcoded thumb-rules ──
//
// The old version assumed a generic residential project and multiplied
// BUA by hardcoded factors (RCC × 0.15, brickwork × 0.30, plaster × 3.5
// etc.). Those factors are project-specific thumb rules and must NOT be
// applied blindly to every drawing.
//
// New behaviour:
//   1. Try to load boq_rules.json (optional, per-project).
//   2. If present, use its formula list to compute BOQ items.
//   3. If absent, write a placeholder row telling the user to create it —
//      no numbers are invented.
//
// boq_rules.json format:
// {
//   "rules": [
//     { "description": "...", "unit": "cum|sqmt|kg|nos",
//       "rate_key": "rcc_m25_cum",
//       "formula": { "source": "polyline_area_sqmt", "factor": 0.15 } },
//     ...
//   ]
// }
function autoGenerateBOQ(ws, row, dd, today) {
  const areas = dd.polyline_areas || [];
  if (areas.length === 0) {
    ws.mergeCells(row, 1, row, 6);
    const c = ws.getCell(row, 1);
    c.value = 'No BOQ could be generated — no closed polylines found in drawing.';
    sc(c, C.GREY, false, 'FF595959', 9, 'left');
    return row + 1;
  }

  // Load per-project BOQ rules. No fallback thumb-rules.
  let boqRules = null;
  try {
    const rulesPath = path.join(__dirname, 'boq_rules.json');
    if (fs.existsSync(rulesPath)) {
      const raw = JSON.parse(fs.readFileSync(rulesPath, 'utf8'));
      boqRules = Array.isArray(raw.rules) ? raw.rules : null;
    }
  } catch (e) {
    console.warn('boq_rules.json load failed:', e.message);
  }

  if (!boqRules || boqRules.length === 0) {
    ws.mergeCells(row, 1, row, 6);
    const c = ws.getCell(row, 1);
    c.value = 'No BOQ generated — create boq_rules.json with per-project formulas to enable auto BOQ. '
            + 'Thumb-rule factors have been removed because they vary by project type.';
    sc(c, C.GREY, false, 'FF595959', 9, 'left');
    ws.getRow(row++).height = 30;
    return row;
  }

  ws.mergeCells(row, 1, row, 6);
  const nc = ws.getCell(row, 1);
  nc.value = 'AUTO-GENERATED FROM boq_rules.json — rates from rates.json';
  sc(nc, C.ORANGE, true, 'FFFFFFFF', 9, 'center');
  ws.getRow(row++).height = 16;

  const totalArea = areas.reduce((s, a) => s + (a.area_sqm || 0), 0);
  const wallLen   = dd.wall_length_m || 0;
  const floorCount = (dd.element_counts && dd.element_counts.floor_count) || 1;

  function resolveFormulaValue(f) {
    if (!f || typeof f !== 'object') return 0;
    const factor = typeof f.factor === 'number' ? f.factor : 1;
    switch (f.source) {
      case 'polyline_area_sqmt': return totalArea * factor;
      case 'wall_length_m':      return wallLen   * factor;
      case 'floor_count':        return floorCount * factor;
      case 'constant':           return factor;
      default:                   return 0;
    }
  }

  let grandTotal = 0;
  let sr = 1;
  boqRules.forEach((item, idx) => {
    const qtyRaw = resolveFormulaValue(item.formula);
    if (!qtyRaw || qtyRaw <= 0) return;
    const qty    = Math.round(qtyRaw * 100) / 100;
    const rate   = RATES[item.rate_key] || 0;
    if (rate <= 0) return;
    const amount = Math.round(qty * rate);
    grandTotal += amount;
    const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
    const r = ws.getRow(row);
    [sr++, item.description, item.unit, qty, rate, amount].forEach((v, i) => {
      const c = r.getCell(i+1); c.value = v;
      sc(c, bg, false, 'FF000000', 9, i < 2 ? 'left' : 'right');
      if (i >= 3 && typeof v === 'number') c.numFmt = '#,##0';
    });
    ws.getRow(row++).height = 15;
  });

  ws.mergeCells(row, 1, row, 4);
  const nt = ws.getCell(row, 1); nt.value = `GRAND TOTAL (BUA ${Math.round(totalArea)} sqm)`;
  sc(nt, C.YELLOW, true, 'FF000000', 9, 'right');
  const nv = ws.getCell(row, 6); nv.value = grandTotal;
  nv.numFmt = '₹#,##0'; sc(nv, C.YELLOW, true, 'FF000000', 10, 'right');
  ws.getRow(row++).height = 18;

  return row;
}

module.exports = { buildDXFExcel };
