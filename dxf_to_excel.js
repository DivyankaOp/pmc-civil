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
  // SHEET 2 — FLOOR LEVELS (extracted from DXF level annotations)
  // ══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('FLOOR LEVELS');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 40;
    ws.getColumn(3).width = 18;
    ws.getColumn(4).width = 18;
    ws.getColumn(5).width = 20;

    let row = 1;
    mergeRow(ws, row++, 5, 'FLOOR LEVEL SCHEDULE — AUTO-EXTRACTED FROM DXF', C.NAVY, 'FFFFFFFF', 11, 22);
    mergeRow(ws, row++, 5, 'All level annotations found in drawing (e.g. +7590 MM LEVEL)', C.MIDBLUE, 'FFFFFFFF', 9, 16);
    hdrRow(ws, row++, ['SR', 'LEVEL LABEL', 'LEVEL (mm)', 'LEVEL (m)', 'REMARKS']);

    const levels = dd.floor_levels || gi.floor_levels || [];
    // Also pull from gemini interpreted floor_levels
    const geminiLevels = gi.floor_levels || [];
    const allLevels = levels.length > 0 ? levels : geminiLevels;

    if (allLevels.length > 0) {
      allLevels.forEach((l, idx) => {
        const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
        const isBasement = (l.level_mm || l.level_m * 1000 || 0) < 0;
        const isTerrace  = /terrace|cabin|roof/i.test(l.label || '');
        const remark = isBasement ? 'Basement level' : isTerrace ? 'Top level' : idx === allLevels.length - 1 ? 'Lowest level' : '';
        dataRow(ws, row++, [
          idx + 1,
          l.label || l.name || '—',
          l.level_mm != null ? l.level_mm : (l.level_m != null ? Math.round(l.level_m * 1000) : '—'),
          l.level_m != null ? l.level_m : (l.level_mm != null ? Math.round(l.level_mm / 10) / 100 : '—'),
          remark
        ], bg, 'center');
      });

      // Building height row
      const maxLevel = allLevels.filter(l => l.level_mm != null).reduce((m, l) => Math.max(m, l.level_mm), 0);
      const minLevel = allLevels.filter(l => l.level_mm != null).reduce((m, l) => Math.min(m, l.level_mm), 0);
      row++;
      ws.mergeCells(row, 1, row, 2);
      const lc = ws.getCell(row, 1); lc.value = 'TOTAL BUILDING HEIGHT';
      sc(lc, C.YELLOW, true, 'FF000000', 9, 'right');
      const lv = ws.getCell(row, 3); lv.value = maxLevel - minLevel;
      sc(lv, C.YELLOW, true, 'FF000000', 9, 'right');
      const lvm = ws.getCell(row, 4); lvm.value = Math.round((maxLevel - minLevel) / 10) / 100;
      sc(lvm, C.YELLOW, true, 'FF000000', 9, 'right');
      ws.getRow(row).height = 16;
    } else {
      mergeRow(ws, row, 5, 'No floor level annotations found. Ensure level text exists in DXF (e.g. "+7590 MM LEVEL").', C.GREY, 'FF595959', 9, 20);
    }

    // Wall & Construction Notes
    const wallNotes = dd.wall_notes || [];
    if (wallNotes.length > 0) {
      row += 2;
      mergeRow(ws, row++, 5, 'WALL & CONSTRUCTION NOTES (from DXF annotations)', C.TEAL, 'FFFFFFFF', 10);
      hdrRow(ws, row++, ['SR', 'ANNOTATION TEXT', 'LAYER', '', ''], C.TEAL);
      wallNotes.slice(0, 40).forEach((note, idx) => {
        dataRow(ws, row++, [idx + 1, note, '—', '', ''], idx % 2 === 0 ? C.WHITE : C.GREY, 'left');
      });
    }
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 3 — ALL TEXT ANNOTATIONS (extracted directly from DXF)
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
  // SHEET 7 — BOQ ESTIMATE (dynamic per drawing type)
  // ══════════════════════════════════════════════════════════════
  {
    const { getBOQForDrawingType, detectDrawingType, DRAWING_TYPES } = require('./drawing_analyzer');

    const ws = wb.addWorksheet('BOQ ESTIMATE');
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 46;
    ws.getColumn(3).width = 10;
    ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 14;
    ws.getColumn(6).width = 16;
    ws.getColumn(7).width = 24;

    // Resolve drawing type — Gemini first, then DXF keyword detection, then general
    const detectedType = gi.drawing_type ||
      detectDrawingType(dd.all_texts || [], dd.layer_names || [], dd.filename || '') ||
      'GENERAL';

    const dtLabel = DRAWING_TYPES[detectedType] || detectedType;
    const totalBUA = Math.max(
      gi.total_bua_sqm || 0,
      dd.total_bua_sqm || 0,
      (dd.polyline_areas || []).reduce((s, a) => s + (a.area_sqm || 0), 0)
    ) || 100; // fallback 100 sqm if nothing parseable

    const elementCounts = { ...(dd.element_counts || {}), ...(gi.element_counts || {}) };
    const parsedDims    = { floor_count: gi.floor_count || dd.floor_count || 10 };

    let row = 1;
    mergeRow(ws, row++, 7, `BILL OF QUANTITIES — ${detectedType.replace(/_/g,' ')}`, C.NAVY, 'FFFFFFFF', 12, 24);
    mergeRow(ws, row++, 7, `Drawing Type: ${dtLabel}`, C.MIDBLUE, 'FFFFFFFF', 9, 16);
    mergeRow(ws, row++, 7, `Project: ${projectName}  |  File: ${dd.filename || '—'}  |  BUA: ${Math.round(totalBUA)} sqm  |  Date: ${today}`, C.LTBLUE, 'FF1F3864', 9, 14);
    mergeRow(ws, row++, 7, '⚠ ESTIMATED: Quantities from drawing geometry + engineering thumb-rules. Verify before final use.', 'FFFFC000', 'FF7F0000', 9, 16);
    row++;
    hdrRow(ws, row++, ['SR', 'DESCRIPTION OF WORK', 'UNIT', 'QTY', 'RATE (₹)', 'AMOUNT (₹)', 'BASIS / REMARK'], C.NAVY);

    // ── Layer 1: Gemini BOQ if AI gave items ──────────────────────────
    const gemBOQ = (gi.boq || []).filter(it => (parseFloat(it.qty) || 0) > 0 && (parseFloat(it.rate) || 0) > 0);
    let grandTotal = 0;
    let sr = 1;

    if (gemBOQ.length > 0) {
      mergeRow(ws, row++, 7, `── AI-EXTRACTED BOQ (from drawing annotations) ──`, C.DKGREEN, 'FFFFFFFF', 9, 15);
      gemBOQ.forEach((item, idx) => {
        const bg  = idx % 2 === 0 ? C.WHITE : C.GREY;
        const qty = parseFloat(item.qty)    || 0;
        const rate= parseFloat(item.rate)   || 0;
        const amt = Math.round(qty * rate);
        grandTotal += amt;
        const rc = ws.getRow(row);
        [sr++, item.description||item.particular||'', item.unit, qty, rate, amt, 'AI extracted'].forEach((v, i) => {
          const c = rc.getCell(i + 1);
          c.value = v;
          sc(c, bg, false, 'FF000000', 9, i < 2 ? 'left' : 'right');
          if (i >= 3 && typeof v === 'number') c.numFmt = '#,##0';
        });
        ws.getRow(row++).height = 15;
      });
    }

    // ── Layer 2: Template BOQ from drawing type (fills gaps / replaces if AI gave nothing) ──
    const templateBOQ = getBOQForDrawingType(detectedType, totalBUA, elementCounts, parsedDims);
    if (templateBOQ.length > 0) {
      const headerLabel = gemBOQ.length > 0
        ? `── SUPPLEMENTARY BOQ (${detectedType.replace(/_/g,' ')} thumb-rule quantities) ──`
        : `── BOQ FOR ${detectedType.replace(/_/g,' ')} (Gujarat DSR 2025 rates) ──`;
      mergeRow(ws, row++, 7, headerLabel, C.ORANGE, 'FFFFFFFF', 9, 15);
      templateBOQ.forEach((item, idx) => {
        const bg  = idx % 2 === 0 ? C.WHITE : C.GREY;
        grandTotal += item.amount;
        const rc = ws.getRow(row);
        [sr++, item.desc, item.unit, item.qty, item.rate, item.amount, `Template: ${detectedType}`].forEach((v, i) => {
          const c = rc.getCell(i + 1);
          c.value = v;
          sc(c, bg, false, 'FF000000', 9, i < 2 ? 'left' : 'right');
          if (i >= 3 && typeof v === 'number') c.numFmt = '#,##0';
        });
        ws.getRow(row++).height = 15;
      });
    }

    if (gemBOQ.length === 0 && templateBOQ.length === 0) {
      mergeRow(ws, row++, 7, 'No BOQ items could be generated — drawing data insufficient.', C.GREY, 'FF595959', 9, 18);
    }

    // Grand total row
    row++;
    ws.mergeCells(row, 1, row, 5);
    const gt = ws.getCell(row, 1);
    gt.value = `GRAND TOTAL — ${detectedType.replace(/_/g,' ')}  (BUA: ${Math.round(totalBUA)} sqm)`;
    sc(gt, C.NAVY, true, 'FFFFFFFF', 10, 'right');
    const gv = ws.getCell(row, 6);
    gv.value = grandTotal;
    gv.numFmt = '₹#,##0';
    sc(gv, C.NAVY, true, 'FFFFD966', 11, 'right');
    ws.getRow(row++).height = 22;

    // Lacs / Crore summary
    const lacs   = Math.round(grandTotal / 100000 * 100) / 100;
    const crores = Math.round(lacs / 100 * 100) / 100;
    ws.mergeCells(row, 1, row, 5);
    const sl = ws.getCell(row, 1);
    sl.value = `TOTAL: ₹${lacs} LACS  (₹${crores} CR)  — ${drawingType} estimate`;
    sc(sl, C.LTBLUE, true, 'FF1F3864', 9, 'right');
    ws.getRow(row++).height = 16;
  }

  // ══════════════════════════════════════════════════════════════
  // SHEET 8 — SPECIAL DATA (drawing-type specific extra info)
  // ══════════════════════════════════════════════════════════════
  {
    const specialData = gi.special_data || {};
    const hasSpecial  = Object.values(specialData).some(v => v && v !== 0);
    const drawingType = gi.drawing_type || 'GENERAL';

    if (hasSpecial) {
      const ws = wb.addWorksheet('SPECIAL DATA');
      ws.getColumn(1).width = 6;
      ws.getColumn(2).width = 36;
      ws.getColumn(3).width = 28;
      ws.getColumn(4).width = 20;

      let row = 1;
      mergeRow(ws, row++, 4, `SPECIAL DATA — ${drawingType.replace(/_/g,' ')}`, C.NAVY, 'FFFFFFFF', 12, 22);
      mergeRow(ws, row++, 4, 'Drawing-type specific parameters extracted from drawing', C.MIDBLUE, 'FFFFFFFF', 9, 15);
      hdrRow(ws, row++, ['SR', 'PARAMETER', 'VALUE', 'UNIT / NOTE'], C.NAVY);

      const paramLabels = {
        lift_car_size:          ['Lift Car Size (W×D)', ''],
        lift_shaft_size:        ['Lift Shaft Size (W×D)', ''],
        lift_pit_depth_m:       ['Lift Pit Depth', 'm'],
        lift_overhead_m:        ['Overhead Clearance', 'm'],
        stair_rise_mm:          ['Stair Rise', 'mm'],
        stair_tread_mm:         ['Stair Tread', 'mm'],
        stair_flight_width_m:   ['Stair Flight Width', 'm'],
        column_size:            ['Column Size (b×d)', 'mm'],
        beam_size:              ['Beam Size (b×d)', 'mm'],
        main_bar_dia_mm:        ['Main Bar Diameter', 'mm Fe500'],
        stirrup_spacing_mm:     ['Stirrup Spacing', 'mm c/c'],
        cover_mm:               ['Clear Cover', 'mm'],
        parking_bay_size:       ['Parking Bay Size', 'm×m'],
        ramp_grade_pct:         ['Ramp Grade', '%'],
        plot_area_sqm:          ['Plot Area', 'sqm'],
        setback_front_m:        ['Front Setback', 'm'],
        setback_side_m:         ['Side Setback', 'm'],
        road_carriageway_width_m:['Carriageway Width', 'm'],
        road_total_length_m:    ['Road Total Length', 'm'],
        pile_dia_mm:            ['Pile Diameter', 'mm'],
        pile_depth_m:           ['Pile Depth', 'm'],
        pipe_dia_mm:            ['Pipe Diameter', 'mm'],
        pipe_material:          ['Pipe Material', ''],
        invert_level_m:         ['Invert Level', 'm'],
        floor_height_m:         ['Typical Floor Height', 'm'],
        parapet_height_m:       ['Parapet Height', 'm'],
        cladding_material:      ['Cladding Material', ''],
      };

      let sr = 1;
      for (const [key, [label, unit]] of Object.entries(paramLabels)) {
        const val = specialData[key];
        if (val === undefined || val === null || val === 0 || val === '') continue;
        const bg = sr % 2 === 0 ? C.WHITE : C.GREY;
        dataRow(ws, row++, [sr++, label, val, unit], bg, 'left');
      }

      if (sr === 1) {
        mergeRow(ws, row, 4, 'No special parameters extracted for this drawing type.', C.GREY, 'FF595959', 9, 18);
      }
    }
  }

  return wb;
}

// ── Auto BOQ from polyline areas (fallback when no Gemini BOQ) ──
function autoGenerateBOQ(ws, row, dd, today) {
  const areas = dd.polyline_areas || [];
  if (areas.length === 0) {
    ws.mergeCells(row, 1, row, 6);
    const c = ws.getCell(row, 1);
    c.value = 'No BOQ could be generated — no closed polylines found in drawing.';
    sc(c, C.GREY, false, 'FF595959', 9, 'left');
    return row + 1;
  }

  ws.mergeCells(row, 1, row, 6);
  const nc = ws.getCell(row, 1);
  nc.value = 'AUTO-GENERATED FROM DXF POLYLINE AREAS — rates from rates.json';
  sc(nc, C.ORANGE, true, 'FFFFFFFF', 9, 'center');
  ws.getRow(row++).height = 16;

  // FIX-3: Disclaimer — quantities are thumb-rule estimates, not from drawing dims
  ws.mergeCells(row, 1, row, 6);
  const disc = ws.getCell(row, 1);
  disc.value = '⚠️ ESTIMATES ONLY — Quantities use engineering thumb-rule conversion factors (e.g. Area×0.15 for RCC volume). NOT extracted from drawing dimensions. Verify all quantities against actual drawing before use.';
  sc(disc, 'FFFFC000', true, 'FF7F0000', 8, 'left');
  ws.getRow(row++).height = 22;

  // BASIS column header
  const hdr = ws.getRow(row);
  ['SR', 'DESCRIPTION (ESTIMATED)', 'UNIT', 'QTY (THUMB RULE)', 'RATE (₹)', 'AMOUNT (₹)'].forEach((h, i) => {
    const c = hdr.getCell(i + 1); c.value = h;
    sc(c, C.BLUE, true, 'FFFFFFFF', 9, 'center');
  });
  ws.getRow(row++).height = 16;

  const totalArea = areas.reduce((s, a) => s + (a.area_sqm || 0), 0);
  let grandTotal = 0;
  let sr = 1;

  // Generate BOQ items from rates.json — only items that make sense for the drawing type
  const boqItems = [
    { desc: 'RCC Structure (slab, beams, columns) @ M25',          unit: 'cum',  factor: totalArea * 0.15,   rateKey: 'rcc_m25_cum' },
    { desc: 'Brickwork in walls 230mm thick',                       unit: 'cum',  factor: totalArea * 0.30,   rateKey: 'brickwork_230mm_cum' },
    { desc: 'Plaster (internal) 15mm thick both faces',             unit: 'sqmt', factor: totalArea * 3.5,    rateKey: 'plaster_15mm_sqmt' },
    { desc: 'Steel reinforcement Fe500',                            unit: 'kg',   factor: totalArea * 55,     rateKey: 'steel_fe500_kg' },
    { desc: 'Vitrified tile flooring',                              unit: 'sqmt', factor: totalArea * 0.85,   rateKey: 'flooring_vitrified_sqmt' },
    { desc: 'Internal painting (2 coats)',                          unit: 'sqmt', factor: totalArea * 3.5,    rateKey: 'painting_sqmt' },
    { desc: 'Aluminum windows',                                     unit: 'sqmt', factor: totalArea * 0.12,   rateKey: 'window_aluminum_sqmt' },
    { desc: 'Electrical works (per sqmt BUA)',                      unit: 'sqmt', factor: totalArea,          rateKey: 'electrical_sqmt' },
    { desc: 'Waterproofing (terrace + bathrooms)',                   unit: 'sqmt', factor: totalArea * 0.25,   rateKey: 'waterproofing_sqmt' },
    { desc: 'Formwork (slab + beams + columns)',                    unit: 'sqmt', factor: totalArea * 2.5,    rateKey: 'formwork_sqmt' },
  ].filter(item => RATES[item.rateKey] > 0);

  boqItems.forEach((item, idx) => {
    const qty    = Math.round(item.factor * 100) / 100;
    const rate   = RATES[item.rateKey] || 0;
    const amount = Math.round(qty * rate);
    grandTotal += amount;
    const bg = idx % 2 === 0 ? C.WHITE : C.GREY;
    const r = ws.getRow(row);
    [sr++, item.desc, item.unit, qty, rate, amount].forEach((v, i) => {
      const c = r.getCell(i+1); c.value = v;
      sc(c, bg, false, 'FF000000', 9, i < 2 ? 'left' : 'right');
      if (i >= 3 && typeof v === 'number') c.numFmt = '#,##0';
    });
    ws.getRow(row++).height = 15;
  });

  // Note row
  ws.mergeCells(row, 1, row, 4);
  const nt = ws.getCell(row, 1); nt.value = `GRAND TOTAL (${Math.round(totalArea)} sqm BUA)`;
  sc(nt, C.YELLOW, true, 'FF000000', 9, 'right');
  const nv = ws.getCell(row, 6); nv.value = grandTotal;
  nv.numFmt = '₹#,##0'; sc(nv, C.YELLOW, true, 'FF000000', 10, 'right');
  ws.getRow(row++).height = 18;

  return row;
}

module.exports = { buildDXFExcel };
