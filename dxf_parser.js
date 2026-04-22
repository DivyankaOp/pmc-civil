/**
 * PMC DXF Parser — v2.0
 * Pure JavaScript, zero dependencies.
 *
 * Extracts ALL data directly from the DXF file:
 *   - Every TEXT and MTEXT entity (room names, annotations, labels)
 *   - Every DIMENSION entity (walls, openings, lengths)
 *   - Every LWPOLYLINE / POLYLINE (room boundaries → areas)
 *   - All layer names and which entities live on each layer
 *   - Title block info (project, scale, date, drawn-by)
 *   - Drawing extents and inferred scale
 *   - INSERT references (block instances: doors, columns, windows)
 *
 * NO rates hardcoded here. All rates come from rates.json.
 */

'use strict';

// ── LOAD RATES FROM CONFIG (no hardcoding) ───────────────────────
let RATES = {};
try {
  const fs   = require('fs');
  const path = require('path');
  const raw  = JSON.parse(fs.readFileSync(path.join(__dirname, 'rates.json'), 'utf8'));
  for (const category of Object.values(raw)) {
    if (typeof category === 'object' && !Array.isArray(category)) {
      for (const [key, val] of Object.entries(category)) {
        if (val && typeof val.rate === 'number') RATES[key] = val.rate;
      }
    }
  }
} catch(e) {
  console.warn('rates.json not loaded:', e.message);
}

// ─────────────────────────────────────────────────────────────────
// SECTION 1 — RAW DXF PARSER
// ─────────────────────────────────────────────────────────────────

function parseDXF(dxfContent) {
  const lines = dxfContent.split(/\r?\n/);
  const result = {
    version: '', units: 'mm',
    layers: {}, entities: [], texts: [], dimensions: [],
    polylines: [], inserts: [], blocks: {}, title_block: {},
    extents: { xmin: Infinity, xmax: -Infinity, ymin: Infinity, ymax: -Infinity },
    header_vars: {}
  };

  let i = 0;
  const next = () => lines[i++]?.trim();
  const peek = () => lines[i]?.trim();

  function updateExtents(x, y) {
    if (!isNaN(x) && !isNaN(y)) {
      if (x < result.extents.xmin) result.extents.xmin = x;
      if (x > result.extents.xmax) result.extents.xmax = x;
      if (y < result.extents.ymin) result.extents.ymin = y;
      if (y > result.extents.ymax) result.extents.ymax = y;
    }
  }

  // Read all group-code / value pairs until next entity (code 0)
  function readPairs() {
    const props = {};
    while (i < lines.length) {
      const cRaw = peek();
      const c = parseInt(cRaw);
      if (isNaN(c)) { next(); next(); continue; }
      if (c === 0) break;
      next(); // consume code
      const v = next(); // consume value
      if (c in props) {
        if (!Array.isArray(props[c])) props[c] = [props[c]];
        props[c].push(v);
      } else {
        props[c] = v;
      }
    }
    return props;
  }

  function cleanMtext(s) {
    if (!s) return '';
    return s
      .replace(/\\P/g, '\n')
      .replace(/\\p[^;]*;/g, '')
      .replace(/\\f[^;]*;/g, '')
      .replace(/\\H[^;]*;/g, '')
      .replace(/\\W[^;]*;/g, '')
      .replace(/\\C[^;]*;/g, '')
      .replace(/\{\\[^}]*\}/g, '')
      .replace(/\\[AaLlOoSsUuQqTt]/g, '')
      .replace(/%%[cCdDpP]/g, '')
      .replace(/\\\\/g, '\\')
      .trim();
  }

  function readTextEntity() {
    const p = readPairs();
    let text = p[1] || '';
    if (p[3]) {
      const extra = Array.isArray(p[3]) ? p[3].join('') : p[3];
      text = extra + text;
    }
    text = cleanMtext(text);
    if (!text) return null;
    const x = parseFloat(p[10]) || 0, y = parseFloat(p[20]) || 0;
    updateExtents(x, y);
    return { text, x, y, height: parseFloat(p[40]) || 0, layer: p[8] || '0', rotation: parseFloat(p[50]) || 0 };
  }

  function readDimension() {
    const p = readPairs();
    const x1 = parseFloat(p[10]) || 0, y1 = parseFloat(p[20]) || 0;
    const x2 = parseFloat(p[13]) || 0, y2 = parseFloat(p[23]) || 0;
    const measured = parseFloat(p[42]);
    // FIX bug#7: DXF code 70 = dimension type. 2=angular, 3=diameter, 4=radius — skip Euclidean fallback for these
    const dimType = parseInt(p[70]) || 0;
    const isAngularOrCurved = (dimType & 7) >= 2; // bits 0-2 encode type; 2=angular,3=diam,4=rad
    const geom = (!isAngularOrCurved) ? Math.sqrt((x2-x1)**2 + (y2-y1)**2) : null;
    updateExtents(x1, y1); updateExtents(x2, y2);
    return {
      dimtext: p[1] || '',
      dim_type: dimType,
      value_mm: (!isNaN(measured) && measured > 0) ? measured : (geom || 0),
      is_estimated: isNaN(measured) || measured <= 0,
      x1, y1, x2, y2,
      xt: parseFloat(p[11]) || 0, yt: parseFloat(p[21]) || 0,
      layer: p[8] || '0', dimstyle: p[3] || ''
    };
  }

  function readLWPolyline() {
    const p = readPairs();
    const closed = (parseInt(p[70]) & 1) === 1;
    const vertices = [];
    const xs = Array.isArray(p[10]) ? p[10].map(parseFloat) : (p[10] !== undefined ? [parseFloat(p[10])] : []);
    const ys = Array.isArray(p[20]) ? p[20].map(parseFloat) : (p[20] !== undefined ? [parseFloat(p[20])] : []);
    for (let k = 0; k < Math.min(xs.length, ys.length); k++) {
      vertices.push({ x: xs[k], y: ys[k] });
      updateExtents(xs[k], ys[k]);
    }
    return vertices.length >= 2 ? { vertices, closed, layer: p[8] || '0' } : null;
  }

  function readOldPolyline() {
    const p = readPairs();
    const closed = (parseInt(p[70]) & 1) === 1;
    const layer = p[8] || '0';
    const vertices = [];
    while (i < lines.length) {
      const c = parseInt(peek());
      if (isNaN(c)) { next(); next(); continue; }
      if (c !== 0) { next(); next(); continue; }
      next(); const etype = next();
      if (!etype) continue;
      if (etype.trim().toUpperCase() === 'SEQEND') break;
      if (etype.trim().toUpperCase() === 'VERTEX') {
        const vp = readPairs();
        const x = parseFloat(vp[10]) || 0, y = parseFloat(vp[20]) || 0;
        vertices.push({ x, y }); updateExtents(x, y);
      }
    }
    return vertices.length >= 2 ? { vertices, closed, layer } : null;
  }

  function readLine() {
    const p = readPairs();
    const x1 = parseFloat(p[10]) || 0, y1 = parseFloat(p[20]) || 0;
    const x2 = parseFloat(p[11]) || 0, y2 = parseFloat(p[21]) || 0;
    updateExtents(x1, y1); updateExtents(x2, y2);
    return { x1, y1, x2, y2, layer: p[8] || '0' };
  }

  function readInsert() {
    const p = readPairs();
    const x = parseFloat(p[10]) || 0, y = parseFloat(p[20]) || 0;
    updateExtents(x, y);
    return { block: p[2] || '', x, y, sx: parseFloat(p[41]) || 1, sy: parseFloat(p[42]) || 1, rotation: parseFloat(p[50]) || 0, layer: p[8] || '0' };
  }

  function skipEntity() { readPairs(); }

  function dispatchEntity(etype, dest) {
    const t = (etype || '').trim().toUpperCase();
    if      (t === 'TEXT' || t === 'ATTDEF' || t === 'ATTRIB') {
      const e = readTextEntity(); if (e) { dest.texts.push(e); dest.entities.push({ type: t, ...e }); }
    } else if (t === 'MTEXT') {
      const e = readTextEntity(); if (e) { dest.texts.push(e); dest.entities.push({ type: t, ...e }); }
    } else if (t === 'DIMENSION') {
      const e = readDimension(); if (e) { dest.dimensions.push(e); dest.entities.push({ type: t, ...e }); }
    } else if (t === 'LWPOLYLINE') {
      const e = readLWPolyline(); if (e) { dest.polylines.push(e); dest.entities.push({ type: t, ...e }); }
    } else if (t === 'POLYLINE') {
      const e = readOldPolyline(); if (e) { dest.polylines.push(e); dest.entities.push({ type: t, ...e }); }
    } else if (t === 'LINE') {
      const e = readLine(); if (e) dest.entities.push({ type: t, ...e });
    } else if (t === 'INSERT') {
      const e = readInsert(); if (e) { dest.inserts.push(e); dest.entities.push({ type: t, ...e }); }
    } else {
      skipEntity();
    }
  }

  function parseSection(dest) {
    while (i < lines.length) {
      const c = parseInt(peek());
      if (isNaN(c)) { next(); next(); continue; }
      if (c !== 0) { next(); next(); continue; }
      next(); const etype = next();
      if (!etype) continue;
      const up = etype.trim().toUpperCase();
      if (up === 'ENDSEC' || up === 'ENDBLK') break;
      dispatchEntity(etype, dest);
    }
  }

  function parseHeader() {
    let curVar = '';
    while (i < lines.length) {
      const c = parseInt(peek());
      if (isNaN(c)) { next(); next(); continue; }
      next(); const v = next();
      if (c === 0 && v === 'ENDSEC') break;
      if (c === 9) { curVar = v; }
      else if (curVar) {
        result.header_vars[curVar] = v;
        if (curVar === '$INSUNITS') {
          const u = parseInt(v);
          result.units = u===4?'mm': u===5?'cm': u===6?'m': u===1?'in': u===2?'ft': 'mm';
        }
      }
    }
    result.version = result.header_vars['$ACADVER'] || '';
  }

  function parseTables() {
    let inLayer = false, lName = '', lColor = 0, lLtype = '';
    while (i < lines.length) {
      const c = parseInt(peek());
      if (isNaN(c)) { next(); next(); continue; }
      next(); const v = next();
      if (c === 0 && v === 'ENDSEC') break;
      if (c === 0 && v === 'TABLE') inLayer = false;
      if (c === 2 && v === 'LAYER') inLayer = true;
      if (inLayer) {
        if (c === 0 && v === 'LAYER') { lName = ''; lColor = 0; lLtype = ''; }
        if (c === 2 && !lName)        lName  = v;
        if (c === 62)                 lColor = Math.abs(parseInt(v) || 0);
        if (c === 6)                  lLtype = v;
        if (c === 0 && v !== 'LAYER' && lName) {
          result.layers[lName] = { color: lColor, linetype: lLtype, entities: [] };
          lName = '';
        }
      }
    }
  }

  function parseBlocks() {
    let cb = null;
    while (i < lines.length) {
      const c = parseInt(peek());
      if (isNaN(c)) { next(); next(); continue; }
      next(); const v = next();
      if (c === 0 && v === 'ENDSEC') break;
      if (c === 0 && v === 'BLOCK') {
        cb = { name: '', texts: [], dimensions: [], polylines: [], entities: [], inserts: [] };
      }
      if (c === 2 && cb && !cb.name) cb.name = v;
      if (c === 0 && cb && !['BLOCK','ENDSEC','ENDBLK'].includes(v)) dispatchEntity(v, cb);
      if (c === 0 && v === 'ENDBLK' && cb) {
        if (cb.name) result.blocks[cb.name] = cb;
        cb = null;
      }
    }
  }

  // ── Main parse loop ──────────────────────────────────────────
  while (i < lines.length) {
    const c = parseInt(next());
    const v = next();
    if (isNaN(c) || v === undefined) continue;
    if (c === 0 && v === 'SECTION') {
      next(); // code 2
      const sec = next();
      if      (sec === 'HEADER')   parseHeader();
      else if (sec === 'TABLES')   parseTables();
      else if (sec === 'BLOCKS')   parseBlocks();
      else if (sec === 'ENTITIES') parseSection(result);
    }
  }
  return result;
}


// ─────────────────────────────────────────────────────────────────
// SECTION 2 — INTELLIGENT EXTRACTION (100% from drawing data)
// ─────────────────────────────────────────────────────────────────

function extractCivilData(parsed, filename) {
  const allTexts = [
    ...parsed.texts,
    ...Object.values(parsed.blocks).flatMap(b => b.texts || [])
  ].filter(t => t && t.text && t.text.trim());

  const allDims = [
    ...parsed.dimensions,
    ...Object.values(parsed.blocks).flatMap(b => b.dimensions || [])
  ];

  const allPolylines = [
    ...parsed.polylines,
    ...Object.values(parsed.blocks).flatMap(b => b.polylines || [])
  ];

  const allInserts = [
    ...parsed.inserts,
    ...Object.values(parsed.blocks).flatMap(b => b.inserts || [])
  ];

  // 1. Scale — FIX-7: expanded regex handles 1:100, 1/100, 1=100, SCALE 1:100
  let scale = null, scaleFactor = 1;
  const scaleRE = /(?:scale\s*)?1\s*[:/=]\s*(\d+(?:\.\d+)?)/i;
  for (const t of allTexts) {
    const m = t.text.match(scaleRE);
    if (m) { scale = `1:${m[1]}`; scaleFactor = parseFloat(m[1]); break; }
    // Also catch "NTS" / "NOT TO SCALE"
    if (/\b(NTS|NOT\s+TO\s+SCALE)\b/i.test(t.text)) { scale = 'NTS'; break; }
  }
  // 2. Unit factor → mm
  const u2m = parsed.units === 'm' ? 1000 : parsed.units === 'cm' ? 10 : parsed.units === 'ft' ? 304.8 : parsed.units === 'in' ? 25.4 : 1;

  // 3. Extents
  const ext = parsed.extents;

  // Fallback scale inference from extents vs A1 paper (841×594mm)
  if ((!scale || scale === 'NTS') && ext) {
    const dwgW = Math.abs((ext.maxX || 0) - (ext.minX || 0)) * u2m; // drawing width in mm
    const dwgH = Math.abs((ext.maxY || 0) - (ext.minY || 0)) * u2m;
    if (dwgW > 0 && dwgH > 0) {
      const inferredFactor = Math.round(Math.max(dwgW / 841, dwgH / 594) / 50) * 50; // round to nearest 50
      if (inferredFactor >= 50 && inferredFactor <= 5000) {
        scale = scale || `1:${inferredFactor} (inferred from extents)`;
        scaleFactor = scaleFactor === 1 ? inferredFactor : scaleFactor;
      }
    }
  }
  const extW = (ext.xmax - ext.xmin) * u2m;
  const extH = (ext.ymax - ext.ymin) * u2m;

  // 4. Title block
  const titleBlock = {};
  const titleREs = {
    project_name: /project\s*(name)?|work\s*name|title/i,
    drawing_no:   /drg\s*(no\.?)?|drawing\s*(no\.?)|sheet\s*(no\.?)/i,
    date:         /\bdate\b|\bdt\b/i,
    scale:        /\bscale\b/i,
    drawn_by:     /prepared|designed|drawn\s*by|architect|engineer/i,
    client:       /client|owner/i,
    revision:     /rev(ision)?\s*(no\.?)?/i
  };
  for (const t of allTexts) {
    for (const [k, re] of Object.entries(titleREs)) {
      if (!titleBlock[k] && re.test(t.text)) titleBlock[k] = t.text.trim();
    }
  }

  // 5. Dimension values
  const dimValues = allDims
    .filter(d => d.value_mm > 0)
    .map(d => ({
      value_mm:  Math.round(d.value_mm * u2m),
      value_m:   Math.round(d.value_mm * u2m / 1000 * 100) / 100,
      text:      d.dimtext || String(Math.round(d.value_mm)),
      layer:     d.layer || ''
    }))
    .sort((a, b) => b.value_mm - a.value_mm);

  // 6. Polyline → areas
  const polylineAreas = allPolylines
    .filter(pl => pl.vertices.length >= 3)
    .map(pl => {
      const rawArea = Math.abs(shoelaceArea(pl.vertices));
      const areaMm2 = rawArea * u2m * u2m;
      return {
        area_sqm:    Math.round(areaMm2 / 1e6 * 100) / 100,
        area_sqft:   Math.round(areaMm2 / 1e6 * 10.764 * 100) / 100,
        perimeter_m: Math.round(perimeter(pl.vertices) * u2m / 1000 * 100) / 100,
        layer:       pl.layer || '',
        closed:      pl.closed || false,
        vertices:    pl.vertices.length
      };
    })
    .filter(a => a.area_sqm > 0.01)
    .sort((a, b) => b.area_sqm - a.area_sqm);

  // 7. Room annotations from text
  // FIX bug#5: Expanded space regex to catch custom names: OFFICE CABIN, GYM AREA, CLUB HOUSE, etc.
  const spaceRE = /bed\s*room|bedroom|living|drawing\s*room|dining|kitchen|bath|toilet|wc|passage|corridor|lobby|hall|store|staircase|stair|lift|balcony|terrace|utility|servant|garage|office|cabin|gym|club|reception|conference|meeting|server\s*room|pantry|lounge|flat|unit|room|area|court|garden|parking|podium|ramp|duct|shaft/i;
  const roomAnnotations = allTexts
    .filter(t => spaceRE.test(t.text))
    .map(t => ({ text: t.text.trim(), x: t.x, y: t.y, layer: t.layer }));

  // 8. Inline dimension text (e.g. "3000x4500", "3.0m × 4.5m", "3.0 x 4.5m")
  const inlineDims = [];
  // FIX-6: Handle optional unit suffix (m/mm) and spaces around separator
  const dimRE = /(\d+(?:\.\d+)?)\s*(m{1,2})?\s*[xX×]\s*(\d+(?:\.\d+)?)\s*(m{1,2})?/g;
  for (const t of allTexts) {
    let m;
    while ((m = dimRE.exec(t.text)) !== null) {
      // Normalize to mm: if unit is 'm' multiply by 1000, else use as-is (assume mm)
      const unit1 = (m[2] || '').toLowerCase(), unit2 = (m[4] || '').toLowerCase();
      const toMm = (val, unit) => unit === 'm' ? parseFloat(val) * 1000 : parseFloat(val) * u2m;
      const l = toMm(m[1], unit1), w = toMm(m[3], unit2);
      if (l > 50 && w > 50) inlineDims.push({ label: t.text, length_mm: Math.round(l), width_mm: Math.round(w), area_sqm: Math.round(l*w/1e6*100)/100, layer: t.layer });
    }
  }

  // 9. Layer groups
  const layerGroups = {};
  for (const e of parsed.entities) {
    const layer = e.layer || 'DEFAULT';
    if (!layerGroups[layer]) layerGroups[layer] = { texts: [], lines: [], dims: [], polylines: [], inserts: [], count: 0 };
    layerGroups[layer].count++;
    if (e.type === 'TEXT' || e.type === 'MTEXT') layerGroups[layer].texts.push(e.text);
    if (e.type === 'LINE')       layerGroups[layer].lines.push(e);
    if (e.type === 'DIMENSION')  layerGroups[layer].dims.push(e.value_mm);
    if (e.type === 'LWPOLYLINE' || e.type === 'POLYLINE') layerGroups[layer].polylines.push(e);
    if (e.type === 'INSERT')     layerGroups[layer].inserts.push(e.block);
  }

  // 10. Block counts
  const blockCounts = {};
  for (const ins of allInserts) blockCounts[ins.block] = (blockCounts[ins.block] || 0) + 1;

  // 11. Drawing type
  const drawingType = inferDrawingType(Object.keys(layerGroups), allTexts.map(t => t.text), filename || '');

  // 12. Unique text list
  const uniqueTexts = [...new Set(allTexts.map(t => t.text.trim()))].filter(Boolean);

  return {
    filename,
    drawing_type:    drawingType,
    scale, scale_factor: scaleFactor,
    units: parsed.units, unit_to_mm: u2m,
    drawing_extents: { width_mm: Math.round(extW), height_mm: Math.round(extH), width_m: Math.round(extW/1000*100)/100, height_m: Math.round(extH/1000*100)/100 },
    title_block:     titleBlock,
    all_texts:       uniqueTexts,
    _raw_texts:      allTexts.map(t => ({ text: t.text, x: t.x, y: t.y, layer: t.layer })),  // FIX-4: expose positioned texts for ANNOTATIONS sheet
    room_annotations: roomAnnotations,
    layer_names:     Object.keys(layerGroups).filter(Boolean),
    layer_groups:    layerGroups,
    dimension_values: dimValues,
    inline_dims:     inlineDims,
    polyline_areas:  polylineAreas.slice(0, 300),
    block_counts:    blockCounts,
    stats: {
      total_texts:     allTexts.length,
      total_dims:      allDims.length,
      total_lines:     parsed.entities.filter(e => e.type === 'LINE').length,
      total_polylines: allPolylines.length,
      total_inserts:   allInserts.length,
      total_layers:    Object.keys(layerGroups).length,
      unique_blocks:   Object.keys(blockCounts).length
    }
  };
}

function inferDrawingType(layers, texts, filename) {
  const all = [...layers, ...texts, filename].join(' ').toLowerCase();
  if (/section|sectional/.test(all))                 return 'SECTION';
  if (/floor\s*plan|flat\s*plan|layout/.test(all))   return 'FLOOR_PLAN';
  if (/elevation/.test(all))                         return 'ELEVATION';
  if (/site\s*plan|master\s*plan/.test(all))         return 'SITE_PLAN';
  if (/road|highway|carriageway/.test(all))          return 'ROAD_PLAN';
  if (/rcc|reinforcement|bbs/.test(all))             return 'STRUCTURAL';
  if (/detail/.test(all))                            return 'DETAIL';
  return 'GENERAL';
}

function shoelaceArea(pts) {
  let a = 0;
  for (let i = 0; i < pts.length; i++) { const j = (i+1)%pts.length; a += pts[i].x*pts[j].y - pts[j].x*pts[i].y; }
  return a / 2;
}

function perimeter(pts) {
  let p = 0;
  for (let i = 0; i < pts.length; i++) { const j = (i+1)%pts.length; p += Math.sqrt((pts[j].x-pts[i].x)**2+(pts[j].y-pts[i].y)**2); }
  return p;
}

// NEW: Total area from all closed polylines
function extractTotalAreaSqft(dxfContent) {
  const parsed = parseDXF(dxfContent);
  const u2m = parsed.units==='m'?1000 : parsed.units==='cm'?10 : parsed.units==='ft'?304.8 : parsed.units==='in'?25.4 : 1;
  const allPolylines = [
    ...parsed.polylines,
    ...Object.values(parsed.blocks).flatMap(b => b.polylines)
  ];
  const total = allPolylines
    .filter(pl => pl.vertices && pl.vertices.length >= 3 && pl.closed)
    .reduce((sum, pl) => {
      const rawArea = Math.abs(shoelaceArea(pl.vertices));
      const sqft = (rawArea * u2m * u2m) / 1e6 * 10.764;
      return sum + (sqft > 1 ? sqft : 0);
    }, 0);
  return Math.round(total * 100) / 100;
}

module.exports = { parseDXF, extractCivilData, extractTotalAreaSqft, RATES };
