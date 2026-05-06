'use strict';
/**
 * PMC Drawing Intelligence Engine
 * 
 * Step 1 → Scan DXF: extract ALL layers, hatches, blocks, texts, dimensions
 * Step 2 → Find Legend: read the symbol/legend table drawn INSIDE the DXF itself
 * Step 3 → Auto-Map: map layers by name patterns (works for ANY drawing)
 * Step 4 → Learn & Save: save new mappings to symbols-learned.json
 * Step 5 → Calculate: compute quantities from mapped data
 * 
 * Works for ANY DXF — no hardcoding of layer names.
 * Learned mappings grow over time across drawings.
 */

const fs   = require('fs');
const path = require('path');

const LEARNED_FILE = path.join(__dirname, 'symbols-learned.json');

// ─────────────────────────────────────────────────────────────────
// SECTION 1 — LOAD / SAVE LEARNED SYMBOLS
// ─────────────────────────────────────────────────────────────────
function loadLearned() {
  try {
    if (fs.existsSync(LEARNED_FILE)) return JSON.parse(fs.readFileSync(LEARNED_FILE, 'utf8'));
  } catch(e) {}
  return { layers: {}, blocks: {}, text_patterns: {}, floor_heights: {} };
}

function saveLearned(learned) {
  try { fs.writeFileSync(LEARNED_FILE, JSON.stringify(learned, null, 2)); }
  catch(e) { console.warn('Could not save learned symbols:', e.message); }
}

// ─────────────────────────────────────────────────────────────────
// SECTION 2 — RAW DXF SCANNER (zero-dependency)
// ─────────────────────────────────────────────────────────────────
function scanDXF(content) {
  const lines = content.split(/\r?\n/);
  const result = {
    layers:    {},   // name → { color, linetype, entity_count }
    hatches:   [],   // { layer, pattern, color }
    polylines: [],   // { layer, pts, area_m2 }
    inserts:   [],   // { block, layer, x, y }
    texts:     [],   // { text, layer, x, y }
    dims:      [],   // { value_mm, layer, x, y }
    blocks:    {},   // block name → { entities[] }
    extents:   { xmin: Infinity, xmax: -Infinity, ymin: Infinity, ymax: -Infinity }
  };

  let i = 0;
  const nxt = () => lines[i++]?.trim() ?? '';
  const pk  = () => lines[i]?.trim() ?? '';

  function readPairs() {
    const p = {};
    while (i < lines.length) {
      const cStr = pk();
      const c = parseInt(cStr);
      if (isNaN(c)) { nxt(); nxt(); continue; }
      if (c === 0) break;
      nxt();
      const v = nxt();
      if (c in p) { p[c] = Array.isArray(p[c]) ? [...p[c], v] : [p[c], v]; }
      else p[c] = v;
    }
    return p;
  }

  function cleanText(s) {
    if (!s) return '';
    return s
      .replace(/\\P/g, ' ')
      .replace(/\\p[^;]*;/g, '')
      .replace(/\\f[^;]*;/g, '')
      .replace(/\\[HWhWAaCScs][^;]*;/g, '')
      .replace(/\{\\[^}]*\}/g, m => m.replace(/\{\\[^;]*;/g, '').replace(/[{}]/g,''))
      .replace(/[{}]/g, '')
      .replace(/%%[cCdDpP]/g, '')
      .replace(/\\L|\\l|\\O|\\o|\\U|\\u/g, '')
      .trim();
  }

  function shoelace(pts) {
    const n = pts.length;
    if (n < 3) return 0;
    let a = 0;
    for (let j = 0; j < n; j++) {
      const k = (j + 1) % n;
      a += pts[j][0] * pts[k][1] - pts[k][0] * pts[j][1];
    }
    return Math.abs(a) / 2;
  }

  function updateExt(x, y) {
    if (!isNaN(x) && !isNaN(y)) {
      if (x < result.extents.xmin) result.extents.xmin = x;
      if (x > result.extents.xmax) result.extents.xmax = x;
      if (y < result.extents.ymin) result.extents.ymin = y;
      if (y > result.extents.ymax) result.extents.ymax = y;
    }
  }

  // ── TABLES section: read layer definitions ──────────────────────
  function parseTables() {
    while (i < lines.length) {
      const c = parseInt(pk());
      if (isNaN(c)) { nxt(); nxt(); continue; }
      if (c === 0) {
        nxt();
        const v = nxt();
        if (v === 'ENDSEC') break;
        if (v === 'LAYER') {
          const p = readPairs();
          const name  = p[2] || '';
          const color = Math.abs(parseInt(p[62])) || 7;
          const ltype = p[6] || 'Continuous';
          if (name) result.layers[name] = { color, linetype: ltype, entity_count: 0 };
        }
      } else { nxt(); nxt(); }
    }
  }

  // ── ENTITIES section ────────────────────────────────────────────
  function parseEntities(sectionName) {
    while (i < lines.length) {
      const cStr = pk();
      const c = parseInt(cStr);
      if (isNaN(c)) { nxt(); nxt(); continue; }
      if (c !== 0) { nxt(); nxt(); continue; }
      nxt();
      const etype = nxt();
      if (!etype) continue;
      const up = etype.toUpperCase();
      if (up === 'ENDSEC' || up === 'EOF') break;

      if (up === 'TEXT' || up === 'MTEXT' || up === 'ATTDEF') {
        const p = readPairs();
        const layer = p[8] || '0';
        let txt = cleanText(p[1] || p[3] || '');
        if (p[3] && up === 'MTEXT') txt = cleanText(Array.isArray(p[3]) ? p[3].join('') : p[3]) || txt;
        const x = parseFloat(p[10]) || 0;
        const y = parseFloat(p[20]) || 0;
        if (txt) {
          result.texts.push({ text: txt, layer, x, y });
          if (result.layers[layer]) result.layers[layer].entity_count++;
          updateExt(x, y);
        }

      } else if (up === 'DIMENSION') {
        const p = readPairs();
        const layer = p[8] || '0';
        const measured = parseFloat(p[42]);
        const x1 = parseFloat(p[13]) || 0, y1 = parseFloat(p[23]) || 0;
        const x2 = parseFloat(p[14]) || 0, y2 = parseFloat(p[24]) || 0;
        const geom = Math.sqrt((x2-x1)**2 + (y2-y1)**2);
        const val  = (!isNaN(measured) && measured > 0) ? measured : geom;
        if (val > 0) result.dims.push({ value_mm: Math.round(val), layer, x: x1, y: y1 });
        if (result.layers[layer]) result.layers[layer].entity_count++;

      } else if (up === 'HATCH') {
        const p = readPairs();
        const layer   = p[8] || '0';
        const pattern = p[2] || '';
        const color   = p[62] || '';
        result.hatches.push({ layer, pattern, color });
        if (result.layers[layer]) result.layers[layer].entity_count++;

      } else if (up === 'LWPOLYLINE' || up === 'POLYLINE') {
        const p = readPairs();
        const layer = p[8] || '0';
        const xs = Array.isArray(p[10]) ? p[10] : (p[10] ? [p[10]] : []);
        const ys = Array.isArray(p[20]) ? p[20] : (p[20] ? [p[20]] : []);
        if (xs.length >= 2) {
          const pts = xs.map((x, idx) => [parseFloat(x), parseFloat(ys[idx] || 0)]);
          const areaMm2 = shoelace(pts);
          const area_m2 = areaMm2 / 1e6;
          pts.forEach(([x, y]) => updateExt(x, y));
          if (areaMm2 > 0) result.polylines.push({ layer, pts, area_m2 });
          if (result.layers[layer]) result.layers[layer].entity_count++;
        }

      } else if (up === 'INSERT') {
        const p = readPairs();
        const block = p[2] || '';
        const layer = p[8] || '0';
        const x = parseFloat(p[10]) || 0;
        const y = parseFloat(p[20]) || 0;
        result.inserts.push({ block, layer, x, y });
        if (result.layers[layer]) result.layers[layer].entity_count++;
        updateExt(x, y);

      } else if (up === 'LINE') {
        const p = readPairs();
        const layer = p[8] || '0';
        if (result.layers[layer]) result.layers[layer].entity_count++;
        updateExt(parseFloat(p[10])||0, parseFloat(p[20])||0);
        updateExt(parseFloat(p[11])||0, parseFloat(p[21])||0);

      } else {
        readPairs();
      }
    }
  }

  // ── BLOCKS section ──────────────────────────────────────────────
  function parseBlocksSection() {
    let currentBlock = null;
    while (i < lines.length) {
      const cStr = pk();
      const c = parseInt(cStr);
      if (isNaN(c)) { nxt(); nxt(); continue; }
      if (c !== 0) { nxt(); nxt(); continue; }
      nxt();
      const v = nxt();
      if (v === 'ENDSEC') break;
      if (v === 'BLOCK') {
        const p = readPairs();
        currentBlock = p[2] || '';
        if (currentBlock && !currentBlock.startsWith('*')) {
          result.blocks[currentBlock] = { texts: [], inserts: [] };
        }
      } else if (v === 'ENDBLK') {
        currentBlock = null;
      } else if (currentBlock && result.blocks[currentBlock]) {
        if (v === 'TEXT' || v === 'MTEXT') {
          const p = readPairs();
          const txt = cleanText(p[1] || p[3] || '');
          if (txt) result.blocks[currentBlock].texts.push(txt);
        } else if (v === 'INSERT') {
          const p = readPairs();
          result.blocks[currentBlock].inserts.push(p[2] || '');
        } else {
          readPairs();
        }
      } else {
        readPairs();
      }
    }
  }

  // ── Main parser loop ────────────────────────────────────────────
  while (i < lines.length) {
    const c = parseInt(pk());
    if (isNaN(c)) { nxt(); nxt(); continue; }
    if (c !== 0) { nxt(); nxt(); continue; }
    nxt();
    const v = nxt();
    if (v === 'SECTION') {
      nxt();
      const sec = nxt();
      if (sec === 'TABLES')   parseTables();
      else if (sec === 'ENTITIES') parseEntities('MODEL');
      else if (sec === 'BLOCKS')   parseBlocksSection();
    }
  }

  return result;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 3 — LEGEND DETECTION (reads the legend drawn inside DXF)
// ─────────────────────────────────────────────────────────────────
/**
 * Many drawings contain a LEGEND / SYMBOL TABLE with rows like:
 *   [hatch symbol]   "230MM THK. BRICK WALL"
 *   [hatch symbol]   "100 MM THK. BLOCK WALL"
 *   [hatch symbol]   "R.C.C PARDI"
 *
 * Strategy: find texts that match known material/element keywords,
 * then look for which LAYER has hatches nearby → link text → layer → meaning.
 */
const LEGEND_PATTERNS = [
  { re: /(\d+)\s*MM\s+THK\.?\s*(BRICK|BLOCK|AAC|STONE|GLASS)\s+WALL/i,   cat: 'wall', extract: m => ({ thk_mm: parseInt(m[1]), material: m[2].toUpperCase() }) },
  { re: /(\d+)\s*MM\s+WALL/i,             cat: 'wall',     extract: m => ({ thk_mm: parseInt(m[1]) }) },
  { re: /WALL\s*[-–]?\s*(\d+)\s*MM/i,     cat: 'wall',     extract: m => ({ thk_mm: parseInt(m[1]) }) },
  { re: /R\.?C\.?C\.?\s*PARDI/i,          cat: 'rcc_pardi',extract: _ => ({}) },
  { re: /R\.?C\.?C\.?\s*(COLUMN|COL)/i,   cat: 'column',   extract: _ => ({}) },
  { re: /R\.?C\.?C\.?\s*(SLAB|BEAM)/i,    cat: 'slab',     extract: _ => ({}) },
  { re: /COLUMN/i,                         cat: 'column',   extract: _ => ({}) },
  { re: /SLAB/i,                           cat: 'slab',     extract: _ => ({}) },
  { re: /FOOTING|FOUNDATION/i,             cat: 'footing',  extract: _ => ({}) },
  { re: /EXCAVATION|EXCAV/i,               cat: 'excavation',extract: _=> ({}) },
  { re: /D[-\s]?WALL|DIAPHRAGM/i,         cat: 'd_wall',   extract: _ => ({}) },
  { re: /(\d+)\s*MM\s+SUNK/i,             cat: 'sunk',     extract: m => ({ depth_mm: parseInt(m[1]) }) },
  { re: /RAISED\s*PLATFORM|OTLI/i,        cat: 'raised_platform', extract: _ => ({}) },
  { re: /GRANITE|MARBLE/i,                cat: 'flooring', extract: _ => ({}) },
  { re: /TILES|TILING/i,                  cat: 'flooring', extract: _ => ({}) },
  { re: /RAMP/i,                           cat: 'ramp',     extract: _ => ({}) },
  { re: /STAIR/i,                          cat: 'stair',    extract: _ => ({}) },
  { re: /LIFT\s*SHAFT|ELEVATOR/i,         cat: 'lift',     extract: _ => ({}) },
  { re: /GLASS/i,                          cat: 'glass',    extract: _ => ({}) },
];

const FLOOR_LEVEL_RE = /^[{\\]?(?:\\C\d+;)?([+-]?\d[\d,]*)\s*MM\s*LEVEL/i;
const FLOOR_NAME_RE  = /(THIRD|SECOND|FIRST|GROUND|PLINTH|TERRACE|STAIR\s+CABIN|BASEMENT|TYPICAL|\d+TH|\d+ST|\d+ND|\d+RD)\s*(FLOOR|BASEMENT|LEVEL)?/i;
const HEIGHT_RE      = /H\s*=\s*(\d+(?:\.\d+)?)\s*(MT|M|MM)/i;

function extractLegendFromTexts(texts) {
  const legendItems = [];
  for (const t of texts) {
    const clean = t.text;
    for (const pat of LEGEND_PATTERNS) {
      const m = clean.match(pat.re);
      if (m) {
        const extra = pat.extract(m);
        legendItems.push({
          text:  clean,
          layer: t.layer,
          x: t.x, y: t.y,
          category: pat.cat,
          ...extra
        });
        break;
      }
    }
  }
  return legendItems;
}

function extractFloorLevels(texts) {
  const levels = [];
  for (const t of texts) {
    const clean = t.text;
    const mmMatch = clean.match(/([+-]?\d[\d,]*)\s*MM\s+LEVEL/i);
    if (mmMatch) {
      const mm = parseInt(mmMatch[1].replace(',',''));
      const nameMatch = clean.match(FLOOR_NAME_RE);
      const name = nameMatch ? nameMatch[0].trim().toUpperCase() : `${mm >= 0 ? '+' : ''}${mm} MM LEVEL`;
      levels.push({ label: clean.replace(/[{}\\]/g,'').trim(), name, mm, m: mm / 1000, x: t.x, y: t.y });
    }
  }
  // Deduplicate by mm value, keep unique
  const seen = new Set();
  return levels.filter(l => { if (seen.has(l.mm)) return false; seen.add(l.mm); return true; })
               .sort((a, b) => a.mm - b.mm);
}

// ─────────────────────────────────────────────────────────────────
// SECTION 4 — AUTO-MAP LAYERS (works for any drawing)
// ─────────────────────────────────────────────────────────────────
/**
 * Try to derive semantic meaning from layer NAME alone (no hardcoding).
 * E.g. "AR-HATCH 100 MM BLOCK WALL" → { cat: 'wall', thk_mm: 100 }
 *      "AR-HATCH 230 MM BRICK WALL" → { cat: 'wall', thk_mm: 230 }
 *      "AR-HATCH COLUMN - SLAB"     → { cat: 'column' }
 * Returns null if cannot determine meaning.
 */
function autoMapLayer(layerName) {
  const n = layerName.toUpperCase();

  // WALL thickness from layer name
  const wallThk = n.match(/(\d+)\s*MM\s*(BLOCK|BRICK|AAC|STONE)?\s*WALL/);
  if (wallThk) return { category: 'wall', thk_mm: parseInt(wallThk[1]), source: 'layer_name' };

  // WALL without thickness
  if (/\bWALL\b/.test(n) && !/CURTAIN|GLASS/.test(n)) return { category: 'wall', thk_mm: null, source: 'layer_name' };

  // RCC types
  if (/COLUMN\s*[-–]\s*SLAB|HATCH\s*COLUMN/.test(n)) return { category: 'column', source: 'layer_name' };
  if (/\bCOLUMN\b/.test(n)) return { category: 'column', source: 'layer_name' };
  if (/R\.?C\.?C\.?\s*PARDI|\bPARDI\b/.test(n)) return { category: 'rcc_pardi', source: 'layer_name' };
  if (/\bSLAB\b/.test(n)) return { category: 'slab', source: 'layer_name' };
  if (/\bFOOTING\b|\bFOUNDATION\b/.test(n)) return { category: 'footing', source: 'layer_name' };

  // Openings
  if (/\bDOOR\b/.test(n)) return { category: 'opening', opening_type: 'door', source: 'layer_name' };
  if (/\bWINDOW\b/.test(n)) return { category: 'opening', opening_type: 'window', source: 'layer_name' };
  if (/\bGLASS\b/.test(n)) return { category: 'glass', source: 'layer_name' };

  // Sunk / levels
  const sunk = n.match(/(\d+)\s*MM\s*SUNK/);
  if (sunk) return { category: 'sunk', depth_mm: parseInt(sunk[1]), source: 'layer_name' };

  // Stairs / lift
  if (/\bSTAIR\b/.test(n)) return { category: 'stair', source: 'layer_name' };
  if (/\bRAMP\b/.test(n))  return { category: 'ramp', source: 'layer_name' };
  if (/\bLIFT\b|\bELEV\b/.test(n)) return { category: 'lift', source: 'layer_name' };

  // Flooring
  if (/\bFLOOR\b|\bTILE\b|\bGRANITE\b|\bMARBLE\b/.test(n)) return { category: 'flooring', source: 'layer_name' };

  // Dimension / text / grid — ignore for quantities
  if (/\bDIM\b|\bDIMENSION\b/.test(n)) return { category: 'dimension', source: 'layer_name' };
  if (/\bTEXT\b|\bNOTES?\b|\bANNOT\b/.test(n)) return { category: 'annotation', source: 'layer_name' };
  if (/\bGRID\b|\bCOL\.?\s*LINE\b/.test(n)) return { category: 'grid', source: 'layer_name' };
  if (/\bFORMAT\b|\bTITLE\b|\bBORDER\b/.test(n)) return { category: 'ignore', source: 'layer_name' };

  return null;  // unknown
}

// ─────────────────────────────────────────────────────────────────
// SECTION 5 — QUANTITY COMPUTATION
// ─────────────────────────────────────────────────────────────────
/**
 * Assign each polyline (with area) to a floor based on Y coordinate.
 * Uses extracted floor levels (sorted by mm) to determine which floor a polyline belongs to.
 */
function assignPolylinesToFloors(polylines, floorLevels, scaleToMm) {
  // Build floor bands from Y coords of level annotations
  // Each floor: yBottom = Y coord of lower level, yTop = Y coord of upper level
  if (floorLevels.length < 2) return { 'ALL': polylines };

  // Sort levels by Y coord ascending
  const sorted = [...floorLevels].sort((a, b) => a.y - b.y);

  const result = {};
  for (const poly of polylines) {
    const cy = poly.pts.reduce((s, p) => s + p[1], 0) / poly.pts.length; // centroid Y
    // Find which floor band this centroid falls in
    let assigned = null;
    for (let j = 0; j < sorted.length - 1; j++) {
      if (cy >= sorted[j].y && cy <= sorted[j+1].y) {
        assigned = sorted[j+1].name; // label with upper level name
        break;
      }
    }
    if (!assigned) {
      // Assign to nearest
      const dists = sorted.map((l, idx) => ({ idx, d: Math.abs(cy - l.y) }));
      dists.sort((a,b) => a.d - b.d);
      assigned = sorted[dists[0].idx].name;
    }
    if (!result[assigned]) result[assigned] = [];
    result[assigned].push(poly);
  }
  return result;
}

function computeWallQuantities(scanned, floorLevels, layerMap) {
  const results = {}; // floor → { thk → { sqm, cum } }

  // Group polylines by floor
  const byFloor = assignPolylinesToFloors(scanned.polylines, floorLevels, 1);

  // For each floor, for each wall-category polyline
  for (const [floor, polys] of Object.entries(byFloor)) {
    for (const poly of polys) {
      const mapped = layerMap[poly.layer];
      if (!mapped || mapped.category !== 'wall') continue;
      const thk = mapped.thk_mm || 100;
      const thkM = thk / 1000;

      if (!results[floor]) results[floor] = {};
      if (!results[floor][thk]) results[floor][thk] = { area_m2: 0, vol_m3: 0, count: 0 };

      // area_m2 from polyline / thk = wall face area
      // But polyline IS the plan area of wall = length × thk (in plan)
      // So face_area = plan_area / thk × height — but we don't know height here
      // Better: plan_area = poly.area_m2, length ≈ plan_area / thkM
      const wallLength = poly.area_m2 / thkM;
      results[floor][thk].area_m2 += poly.area_m2; // plan footprint
      results[floor][thk].wallLength = (results[floor][thk].wallLength || 0) + wallLength;
      results[floor][thk].count++;
    }
  }
  return results;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 6 — MAIN EXPORT: analyzeDrawing()
// ─────────────────────────────────────────────────────────────────
function analyzeDrawing(dxfContent, filename) {
  // 1. Scan raw DXF
  const scanned = scanDXF(dxfContent);

  // 2. Extract floor levels from texts
  const floorLevels = extractFloorLevels(scanned.texts);

  // Compute floor heights from consecutive levels
  const floorHeights = [];
  for (let i = 0; i < floorLevels.length - 1; i++) {
    const h = (floorLevels[i+1].mm - floorLevels[i].mm) / 1000;
    if (h > 1.5 && h < 6.0) { // reasonable floor height
      floorHeights.push({ from_mm: floorLevels[i].mm, to_mm: floorLevels[i+1].mm, height_m: h, name: floorLevels[i+1].name });
    }
  }

  // 3. Extract legend items from drawing texts
  const legendItems = extractLegendFromTexts(scanned.texts);

  // 4. Auto-map all layers
  const learned = loadLearned();
  const layerMap = {}; // layerName → { category, thk_mm, ... }

  const unknownLayers = [];
  for (const [layerName, info] of Object.entries(scanned.layers)) {
    if (layerName.startsWith('*') || layerName === 'Defpoints') continue;

    // Priority: (a) user-confirmed learned, (b) auto-detect from name
    let mapping = learned.layers[layerName] || autoMapLayer(layerName);

    if (mapping) {
      layerMap[layerName] = mapping;
      // Save to learned if not already there
      if (!learned.layers[layerName]) {
        learned.layers[layerName] = { ...mapping, _auto: true };
      }
    } else {
      unknownLayers.push({ name: layerName, entity_count: info.entity_count, color: info.color });
    }
  }

  // Auto-map block names
  const blockMap = {};
  const unknownBlocks = [];
  for (const insertName of [...new Set(scanned.inserts.map(ins => ins.block))]) {
    const lb = learned.blocks[insertName];
    if (lb) { blockMap[insertName] = lb; continue; }

    // Auto-detect block meaning from name
    const n = insertName.toUpperCase();
    let bmap = null;
    const doorW = n.match(/^D(\d+)$/);     if (doorW)   bmap = { category: 'opening', opening_type: 'door', _auto: true };
    const winW  = n.match(/^W(\d+)$/);     if (winW)    bmap = { category: 'opening', opening_type: 'window', _auto: true };
    if (/^GD/.test(n))  bmap = { category: 'opening', opening_type: 'glass_door', _auto: true };
    if (/^SLD/.test(n)) bmap = { category: 'opening', opening_type: 'sliding_door', _auto: true };
    if (/^LD/.test(n))  bmap = { category: 'opening', opening_type: 'lift_door', _auto: true };
    if (/^V\d*$/.test(n)) bmap = { category: 'opening', opening_type: 'ventilator', _auto: true };
    if (/^MD/.test(n))  bmap = { category: 'opening', opening_type: 'main_door', _auto: true };
    if (/^FD/.test(n))  bmap = { category: 'opening', opening_type: 'fire_door', _auto: true };
    if (/^KW/.test(n))  bmap = { category: 'opening', opening_type: 'kitchen_window', _auto: true };
    if (/COL/.test(n))  bmap = { category: 'column', _auto: true };
    if (/LIFT/.test(n)) bmap = { category: 'lift', _auto: true };
    if (/STAIR/.test(n))bmap = { category: 'stair', _auto: true };

    if (bmap) {
      blockMap[insertName] = bmap;
      learned.blocks[insertName] = bmap;
    } else {
      unknownBlocks.push({ name: insertName, count: scanned.inserts.filter(ins=>ins.block===insertName).length });
    }
  }

  // Save updated learnings
  saveLearned(learned);

  // 5. Summarize all layers with meaning
  const layerSummary = Object.entries(scanned.layers).map(([name, info]) => ({
    name,
    entity_count: info.entity_count,
    color: info.color,
    category: layerMap[name]?.category || 'unknown',
    thk_mm: layerMap[name]?.thk_mm || null,
    auto_mapped: !!layerMap[name]
  })).sort((a, b) => b.entity_count - a.entity_count);

  // 6. Hatch summary (which hatches are on which meaning-layer)
  const hatchSummary = {};
  for (const h of scanned.hatches) {
    const mapped = layerMap[h.layer];
    const key = mapped ? `${mapped.category}${mapped.thk_mm ? '_' + mapped.thk_mm + 'mm' : ''}` : 'unknown';
    hatchSummary[key] = (hatchSummary[key] || 0) + 1;
  }

  // 7. Count elements
  const wallPolylines      = scanned.polylines.filter(p => layerMap[p.layer]?.category === 'wall');
  const columnPolylines    = scanned.polylines.filter(p => layerMap[p.layer]?.category === 'column');
  const slabPolylines      = scanned.polylines.filter(p => layerMap[p.layer]?.category === 'slab');
  const rccPardiPolylines  = scanned.polylines.filter(p => layerMap[p.layer]?.category === 'rcc_pardi');

  // Wall area by thickness
  const wallByThk = {};
  for (const p of wallPolylines) {
    const thk = layerMap[p.layer]?.thk_mm || 100;
    wallByThk[thk] = (wallByThk[thk] || 0) + p.area_m2;
  }

  // Opening counts
  const doorCount   = scanned.inserts.filter(ins => blockMap[ins.block]?.opening_type === 'door').length;
  const winCount    = scanned.inserts.filter(ins => blockMap[ins.block]?.opening_type === 'window').length;
  const liftDoors   = scanned.inserts.filter(ins => blockMap[ins.block]?.opening_type === 'lift_door').length;
  const allOpenings = scanned.inserts.filter(ins => blockMap[ins.block]?.category === 'opening').length;

  // 8. Extents
  const ext = scanned.extents;
  const widthMm  = isFinite(ext.xmax - ext.xmin) ? ext.xmax - ext.xmin : 0;
  const heightMm = isFinite(ext.ymax - ext.ymin) ? ext.ymax - ext.ymin : 0;

  // 9. Project info from title block texts
  const projectName = (() => {
    for (const t of scanned.texts) {
      if (/modestaa|project\s*title|project\s*name/i.test(t.text)) return t.text.replace(/\\.*/,'').trim();
    }
    return filename?.replace(/\.[^.]+$/, '') || '';
  })();

  return {
    // Raw scan
    filename,
    project_name: projectName,
    drawing_extents: { width_m: Math.round(widthMm/10)/100, height_m: Math.round(heightMm/10)/100 },
    total_texts: scanned.texts.length,
    total_hatches: scanned.hatches.length,
    total_polylines: scanned.polylines.length,
    total_inserts: scanned.inserts.length,
    total_layers: Object.keys(scanned.layers).length,

    // Derived
    floor_levels: floorLevels,
    floor_heights: floorHeights,
    legend_items: legendItems,
    layer_summary: layerSummary,
    hatch_summary: hatchSummary,
    wall_by_thickness_m2: wallByThk,

    // Counts
    element_counts: {
      wall_polylines: wallPolylines.length,
      column_polylines: columnPolylines.length,
      slab_polylines: slabPolylines.length,
      rcc_pardi_polylines: rccPardiPolylines.length,
      door_count: doorCount,
      window_count: winCount,
      lift_door_count: liftDoors,
      total_openings: allOpenings,
      floor_levels_found: floorLevels.length,
    },

    // For Gemini / AI prompt
    all_texts_sample: [...new Set(scanned.texts.map(t => t.text))].slice(0, 2000),
    layer_names: Object.keys(scanned.layers),
    block_names: Object.keys(scanned.blocks),
    unknown_layers: unknownLayers,
    unknown_blocks: unknownBlocks,
    dims_sample: scanned.dims.slice(0, 500).map(d => ({ mm: d.value_mm, m: Math.round(d.value_mm/10)/100, layer: d.layer })),

    // For BOQ Excel generation
    _scanned: scanned,
    _layerMap: layerMap,
    _blockMap: blockMap,
    _wallPolylines: wallPolylines,
  };
}

// ─────────────────────────────────────────────────────────────────
// SECTION 7 — BUILD RICH AI PROMPT from analyzed data
// ─────────────────────────────────────────────────────────────────
function buildAIPrompt(analyzed, ratesSummary) {
  const dd = analyzed;
  const floorLevelStr = dd.floor_levels.map(l => `  ${l.label} = ${l.m > 0 ? '+' : ''}${l.m}m`).join('\n') || 'none found';
  const floorHtStr    = dd.floor_heights.map(l => `  ${l.name}: H = ${l.height_m}m`).join('\n') || 'none';
  const legendStr     = dd.legend_items.map(l => `  [${l.layer}] ${l.text} → ${l.category}${l.thk_mm ? ' ' + l.thk_mm + 'mm' : ''}`).join('\n') || 'none';
  const layerStr      = dd.layer_summary.slice(0, 100).map(l => `  ${l.name} (${l.entity_count} entities) → ${l.category}${l.thk_mm ? ' ' + l.thk_mm + 'mm' : ''}`).join('\n');
  const wallStr       = Object.entries(dd.wall_by_thickness_m2).map(([thk, sqm]) => `  ${thk}mm wall: ${sqm.toFixed(2)} m² plan area`).join('\n') || 'none';
  const unknownStr    = dd.unknown_layers.slice(0, 50).map(l => `  "${l.name}" (${l.entity_count} entities) — meaning unknown`).join('\n') || 'none';

  return `You are a senior PMC civil engineer analyzing a DXF drawing: "${dd.filename}"
All data below is EXTRACTED DIRECTLY from the drawing file. Do NOT invent values.

═══════════════════════════════════════════
DRAWING OVERVIEW
═══════════════════════════════════════════
Project: ${dd.project_name || 'Not found in title block'}
Size: ${dd.drawing_extents.width_m}m × ${dd.drawing_extents.height_m}m
Total texts: ${dd.total_texts} | Hatches: ${dd.total_hatches} | Polylines: ${dd.total_polylines} | Inserts: ${dd.total_inserts}

═══════════════════════════════════════════
FLOOR LEVELS (extracted from drawing annotations)
═══════════════════════════════════════════
${floorLevelStr}

CALCULATED FLOOR HEIGHTS:
${floorHtStr}

═══════════════════════════════════════════
LEGEND TABLE (read from inside drawing)
═══════════════════════════════════════════
${legendStr}

═══════════════════════════════════════════
LAYER MEANINGS (auto-mapped from layer names)
═══════════════════════════════════════════
${layerStr}

═══════════════════════════════════════════
WALL QUANTITIES (from mapped polyline areas)
═══════════════════════════════════════════
${wallStr}

ELEMENT COUNTS:
  Wall polylines: ${dd.element_counts.wall_polylines}
  Column/slab polylines: ${dd.element_counts.column_polylines}
  Doors: ${dd.element_counts.door_count}
  Windows: ${dd.element_counts.window_count}
  Lift doors: ${dd.element_counts.lift_door_count}
  Floor levels found: ${dd.element_counts.floor_levels_found}

═══════════════════════════════════════════
UNKNOWN LAYERS (not yet mapped)
═══════════════════════════════════════════
${unknownStr}

═══════════════════════════════════════════
SAMPLE DIMENSIONS
═══════════════════════════════════════════
${dd.dims_sample.slice(0, 200).map(d => `${d.mm}mm [${d.layer}]`).join(', ')}

RATES (Gujarat DSR 2025): ${ratesSummary}

Return ONLY raw JSON:
{"project_name":"","drawing_type":"SECTION|ELEVATION|FLOOR_PLAN|STRUCTURAL|SITE_PLAN|FOUNDATION|GENERAL","scale":"","building_height_m":0,"total_floors":0,"basements":0,"floor_levels":[{"name":"","level_m":0,"height_m":0}],"wall_schedule":[{"floor":"","thk_mm":100,"sqm":0,"cum":0}],"spaces":[{"name":"","area_sqm":0}],"boq":[{"description":"","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0}],"total_bua_sqm":0,"observations":[],"pmc_recommendation":""}`;
}

module.exports = { analyzeDrawing, buildAIPrompt, scanDXF, extractFloorLevels, extractLegendFromTexts, autoMapLayer, loadLearned, saveLearned };
