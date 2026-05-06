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
const fs   = require('fs');
const path = require('path');
let RATES = {};
try {
  const ratesPath = fs.existsSync(path.join(__dirname, 'rates.json'))
    ? path.join(__dirname, 'rates.json')
    : path.join(__dirname, 'Rates.json');
  const raw = JSON.parse(fs.readFileSync(ratesPath, 'utf8'));
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

// ── LOAD LEGEND (semantic dictionary, nothing hardcoded) ─────────
let legendHelper = null;
try { legendHelper = require('./legend_helper'); }
catch(e) { console.warn('legend_helper not loaded:', e.message); }

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
    const geom = Math.sqrt((x2-x1)**2 + (y2-y1)**2);
    updateExtents(x1, y1); updateExtents(x2, y2);
    return {
      dimtext: p[1] || '',
      value_mm: (!isNaN(measured) && measured > 0) ? measured : geom,
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

  // ── HATCH entity reader ────────────────────────────────────────
  // DXF group codes for HATCH:
  //   2  = pattern name (e.g. "AR-BRSTD", "ANSI31", "SOLID", "AR-CONC")
  //   8  = layer name
  //   70 = solid fill flag (1 = solid, 0 = patterned)
  //   52 = hatch angle
  //   41 = hatch scale
  // We read raw pairs then return a lightweight object.
  function readHatch() {
    const p = readPairs();
    const layer = p[8] || '0';
    const patternName = (p[2] || '').trim().toUpperCase();
    const isSolid = parseInt(p[70] || 0) === 1;
    const angle = parseFloat(p[52]) || 0;
    const scale = parseFloat(p[41]) || 1;
    return { pattern_name: patternName, layer, is_solid: isSolid, angle, scale };
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
    } else if (t === 'HATCH') {
      const e = readHatch(); if (e) {
        if (!dest.hatches) dest.hatches = [];
        dest.hatches.push(e);
        dest.entities.push({ type: t, ...e });
      }
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

function extractCivilData(parsed, filename, legendArg) {
  const legend = null; // direct layer mapping used — no legend_helper needed
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

  const allHatches = [
    ...(parsed.hatches || []),
    ...Object.values(parsed.blocks).flatMap(b => b.hatches || [])
  ];

  // ─── 1. SMART UNIT DETECTION ──────────────────────────────────────
  // INSUNITS header can be wrong (e.g. set to Inches but drawing is in mm)
  // Strategy: look at actual dimension values + text annotations to determine real unit
  // If drawing texts say "+7590 MM LEVEL" and coords are ~7590, units are mm.
  // If coords are ~299 (7590/25.4), units are inches drawn in mm scale.
  
  let u2mm = 1; // multiply raw coords → mm
  let detectedUnit = 'mm';
  
  // Check for "MM LEVEL" or "MM THK" annotations to calibrate
  const mmTexts = allTexts.filter(t => /\d+\s*MM\s+(LEVEL|THK|THICK)/i.test(t.text));
  if (mmTexts.length > 0) {
    // Extract numeric value from text like "+7590 MM LEVEL" 
    const mmMatch = mmTexts[0].text.match(/([+-]?\d+)\s*MM/i);
    if (mmMatch) {
      const textValueMM = Math.abs(parseInt(mmMatch[1]));
      // Find coord of this text entity
      const textEntity = allTexts.find(t => t.text === mmTexts[0].text);
      if (textEntity) {
        // The Y coordinate should be close to textValueMM (if mm) or textValueMM/25.4 (if inches)
        // Actually just trust: if drawing says "MM" everywhere, it IS mm
        // INSUNITS is just a metadata field that architects often set wrong
        u2mm = 1; // raw units are already mm
        detectedUnit = 'mm';
      }
    }
  } else {
    // Fall back to INSUNITS
    const unitCode = parsed.units;
    const unitMap = { 'in': 25.4, 'ft': 304.8, 'mm': 1, 'cm': 10, 'm': 1000 };
    u2mm = unitMap[unitCode] || 1;
    detectedUnit = unitCode || 'mm';
  }

  // ─── 2. EXTENTS (correct) ────────────────────────────────────────
  const ext = parsed.extents;
  const rawW = isFinite(ext.xmax) ? ext.xmax - ext.xmin : 0;
  const rawH = isFinite(ext.ymax) ? ext.ymax - ext.ymin : 0;
  const extW_mm = rawW * u2mm;
  const extH_mm = rawH * u2mm;

  // ─── 3. SCALE ─────────────────────────────────────────────────────
  let scale = null, scaleFactor = 1;
  for (const t of allTexts) {
    const m = t.text.match(/\bscale\b[:\s]*1\s*[:/]\s*(\d+)/i) || t.text.match(/1\s*:\s*(\d+)/);
    if (m) { scale = `1:${m[1]}`; scaleFactor = parseFloat(m[1]); break; }
  }

  // ─── 4. FLOOR LEVELS (from texts like "+7590 MM LEVEL") ──────────
  const floorLevelRE = /([+-]?\d[\d,]*)\s*MM\s+LEVEL|(\bGROUND\s+LEVEL\b)|(\bPLINTH\s+LEVEL\b)|(\bTERRACE\s+LEVEL\b)/i;
  const floorNameRE  = /(THIRD\s+BASEMENT|SECOND\s+BASEMENT|FIRST\s+BASEMENT|RAISED\s+GROUND|GROUND|PLINTH|FIRST|SECOND|THIRD|FOURTH|FIFTH|SIXTH|SEVENTH|EIG[HT]{1,2}H?|NINTH|TENTH|ELEVENTH|TWELTH|TWELFTH|THIRTEENTH|FOURTEENTH|FOURTINTH|FIFTINTH|SIXTINTH|SEVENTINTH|EIGTHINTH|TERRACE|STAIR\s+CABIN)/i;
  
  const floorLevels = [];
  const seenMM = new Set();
  for (const t of allTexts) {
    const mmMatch = t.text.match(/([+-]?\d[\d,]*)\s*MM\s+LEVEL/i);
    if (mmMatch) {
      const mm = parseInt(mmMatch[1].replace(/,/g, ''));
      if (!seenMM.has(mm)) {
        seenMM.add(mm);
        const nameMatch = t.text.match(floorNameRE);
        floorLevels.push({
          label:   t.text.trim(),
          name:    nameMatch ? nameMatch[0].trim() : (mm >= 0 ? `+${mm}MM` : `${mm}MM`),
          mm,
          m:       Math.round(mm / 10) / 100,
          x:       t.x,
          y:       t.y
        });
      }
    }
  }
  floorLevels.sort((a, b) => a.mm - b.mm);

  // Floor heights between consecutive levels
  const floorHeights = [];
  for (let i = 0; i < floorLevels.length - 1; i++) {
    const h = (floorLevels[i+1].mm - floorLevels[i].mm) / 1000;
    if (h > 1.5 && h < 6.0) {
      floorHeights.push({ name: floorLevels[i+1].name, from_mm: floorLevels[i].mm, to_mm: floorLevels[i+1].mm, height_m: Math.round(h * 100) / 100 });
    }
  }

  // ─── 5. DYNAMIC LAYER MAPPER (works for ANY drawing) ────────────
  // No hardcoded layer names. Uses regex patterns on layer name itself.
  // Covers AR-, S-, ST-, A-, ARCH-, STR-, STRUC-, CIVIL-, blank prefix, etc.

  function autoMapLayer(layerName) {
    const n = layerName.toUpperCase().trim();

    // ── IGNORE layers (format, grid, defpoints) ───────────────────
    if (/\b(DEFPOINTS|FORMAT|TITLE\s*BLOCK|BORDER|VIEWPORT|VPORT|XREF|REF|GRID|CL|CENTRE\s*LINE|CENTER\s*LINE|NORTH|COMPASS)\b/.test(n))
      return null;
    if (/^(0|DEFPOINTS)$/.test(n)) return null;

    // ── WALL: extract thickness from layer name ───────────────────
    // Handles: "AR-HATCH 100 MM BLOCK WALL", "WALL-115MM", "S-WALL 230", "BRICK WALL 230MM"
    const wallThk = n.match(/(\d{2,3})\s*MM\s*(BLOCK|BRICK|AAC|FLYASH|FLY\s*ASH|STONE|LIME|LIGHT|THK|THICK)?\s*WALL/)
                 || n.match(/WALL[\s\-_]*(\d{2,3})\s*MM/)
                 || n.match(/(\d{2,3})\s*MM[\s\-_]*WALL/)
                 || n.match(/HATCH[\s\-_]+(\d{2,3})[\s\-_]*MM[\s\-_]*(BLOCK|BRICK|WALL)/);
    if (wallThk) {
      const thk = parseInt(wallThk[1]);
      const mat = wallThk[2] ? wallThk[2].trim() : (thk <= 115 ? 'BLOCK' : 'BRICK');
      return { cat: 'wall', thk_mm: thk, label: `${thk}MM ${mat} Wall`, unit: 'CUM' };
    }
    // Wall without thickness in name
    if (/\bWALL\b/.test(n) && !/CURTAIN|GLASS|PARAPET|RETAINING|PARTY/.test(n))
      return { cat: 'wall', thk_mm: null, label: 'Wall', unit: 'CUM' };

    // ── RCC / STRUCTURAL ─────────────────────────────────────────
    if (/R\.?C\.?C\.?\s*PARDI|PARDI/.test(n))
      return { cat: 'rcc_pardi', thk_mm: null, label: 'RCC Pardi', unit: 'CUM' };
    if (/COLUMN[\s\-]*SLAB|HATCH[\s\-]*COLUMN|COL[\s\-]*SLAB/.test(n))
      return { cat: 'column', thk_mm: null, label: 'RCC Column/Slab', unit: 'CUM' };
    if (/\bCOLUMN\b|\bCOL\b/.test(n) && !/COLOUR|COLOR/.test(n))
      return { cat: 'column', thk_mm: null, label: 'Column', unit: 'CUM' };
    if (/\bSLAB\b/.test(n))
      return { cat: 'slab', thk_mm: null, label: 'Slab', unit: 'CUM' };
    if (/\bBEAM\b/.test(n))
      return { cat: 'beam', thk_mm: null, label: 'Beam', unit: 'CUM' };
    if (/FOOTING|FOUNDATION|RAFT|PILE\s*CAP/.test(n))
      return { cat: 'footing', thk_mm: null, label: 'Footing/Foundation', unit: 'CUM' };
    if (/\bD[\s\-]?WALL\b|DIAPHRAGM/.test(n))
      return { cat: 'retaining_wall', thk_mm: null, label: 'D-Wall', unit: 'CUM' };
    if (/SHEAR\s*WALL|CORE\s*WALL/.test(n))
      return { cat: 'shear_wall', thk_mm: null, label: 'Shear Wall', unit: 'CUM' };

    // ── FLOORING / FINISHES ───────────────────────────────────────
    if (/(\d{2,3})\s*MM\s*SUNK|SUNK\s*(\d{2,3})\s*MM/.test(n)) {
      const d = n.match(/(\d{2,3})\s*MM\s*SUNK|SUNK\s*(\d{2,3})\s*MM/);
      return { cat: 'sunk', thk_mm: parseInt(d[1]||d[2]), label: `${d[1]||d[2]}MM Sunk`, unit: 'SQMT' };
    }
    if (/SUNK/.test(n))
      return { cat: 'sunk', thk_mm: 75, label: 'Sunk', unit: 'SQMT' };
    if (/RAISED\s*PLATFORM|OTLI|RAISED\s*DECK/.test(n))
      return { cat: 'raised_platform', thk_mm: null, label: 'Raised Platform', unit: 'SQMT' };
    if (/GRANITE/.test(n))
      return { cat: 'granite', thk_mm: null, label: 'Granite Flooring', unit: 'SQMT' };
    if (/MARBLE/.test(n))
      return { cat: 'marble', thk_mm: null, label: 'Marble Flooring', unit: 'SQMT' };
    if (/FLOORING|FLOOR\s*FINISH|FLOOR\s*TILE/.test(n))
      return { cat: 'flooring', thk_mm: null, label: 'Flooring', unit: 'SQMT' };
    if (/\bTIL(E|ING)\b|\bTILES\b/.test(n))
      return { cat: 'tiling', thk_mm: null, label: 'Tiles', unit: 'SQMT' };
    if (/PLASTER|PUNNING/.test(n))
      return { cat: 'plaster', thk_mm: null, label: 'Plaster', unit: 'SQMT' };
    if (/WATERPROOF/.test(n))
      return { cat: 'waterproofing', thk_mm: null, label: 'Waterproofing', unit: 'SQMT' };

    // ── GLASS / CLADDING ─────────────────────────────────────────
    if (/\bGLASS\b/.test(n))
      return { cat: 'glass', thk_mm: null, label: 'Glass', unit: 'SQMT' };
    if (/ACP|ALUM[IU]NIUM\s*COMPOSITE|ALUMINIUM\s*COMPOSITE/.test(n))
      return { cat: 'cladding', thk_mm: null, label: 'ACP Cladding', unit: 'SQMT' };
    if (/DRY\s*CLAD|CLADDING/.test(n))
      return { cat: 'cladding', thk_mm: null, label: 'Cladding', unit: 'SQMT' };
    if (/ALUMINIUM\s*FIN|ALUM\s*FIN|AL\s*FIN/.test(n))
      return { cat: 'cladding', thk_mm: null, label: 'Aluminium Fins', unit: 'SQMT' };
    if (/\bSTONE\b|\bSLATE\b/.test(n))
      return { cat: 'stone', thk_mm: null, label: 'Stone Cladding', unit: 'SQMT' };
    if (/ELEVATION\s*TREAT|ELE[\s\.]TREAT/.test(n))
      return { cat: 'cladding', thk_mm: null, label: 'Elevation Treatment', unit: 'SQMT' };

    // ── OPENINGS ─────────────────────────────────────────────────
    if (/\bDOOR\b/.test(n) && !/DOOR\s*FRAME|SCHEDULE/.test(n))
      return { cat: 'door', thk_mm: null, label: 'Door', unit: 'NOS' };
    if (/\bWINDOW\b|\bWIN\b/.test(n) && !/SCHEDULE/.test(n))
      return { cat: 'window', thk_mm: null, label: 'Window', unit: 'NOS' };
    if (/CURTAIN\s*WALL|CURTAIN\s*GLASS/.test(n))
      return { cat: 'curtain_wall', thk_mm: null, label: 'Curtain Wall', unit: 'SQMT' };

    // ── MEP / SERVICES ───────────────────────────────────────────
    if (/LIFT\s*SHAFT|ELEVATOR\s*SHAFT|LIFT\s*WELL/.test(n))
      return { cat: 'lift_shaft', thk_mm: null, label: 'Lift Shaft', unit: 'NOS' };
    if (/\bLIFT\b|\bELEVATOR\b|\bELE\b/.test(n) && /SHAFT|WELL|MACHINE|PIT/.test(n))
      return { cat: 'lift_shaft', thk_mm: null, label: 'Lift Shaft', unit: 'NOS' };
    if (/DUCT|SHAFT/.test(n) && !/CONDUIT/.test(n))
      return { cat: 'duct', thk_mm: null, label: 'Duct/Shaft', unit: 'NOS' };

    // ── STAIR / RAMP ─────────────────────────────────────────────
    if (/\bSTAIR|\bSTEP\b/.test(n))
      return { cat: 'stair', thk_mm: null, label: 'Staircase', unit: 'NOS' };
    if (/\bRAMP\b/.test(n))
      return { cat: 'ramp', thk_mm: null, label: 'Ramp', unit: 'SQMT' };

    // ── PARKING / ROAD ───────────────────────────────────────────
    if (/PARKING|CAR\s*PARK/.test(n))
      return { cat: 'parking', thk_mm: null, label: 'Parking', unit: 'SQMT' };
    if (/ROAD|PAVEMENT|FOOTPATH/.test(n))
      return { cat: 'road', thk_mm: null, label: 'Road/Pavement', unit: 'SQMT' };

    // ── ANNOTATION / IGNORE ───────────────────────────────────────
    if (/\bDIM\b|\bDIMENSION\b|\bTEXT\b|\bNOTE\b|\bANNOT\b|\bLEVEL\s*LINE\b|\bHATCH\s*LABEL\b/.test(n))
      return { cat: 'annotation', thk_mm: null, label: 'Annotation', unit: null };
    if (/FURNITURE|FITTINGS|SANITARY|LANDSCAPE/.test(n))
      return { cat: 'furniture', thk_mm: null, label: 'Furniture', unit: null };

    return null; // truly unknown layer
  }

  // Build dynamic layer map from ALL layers in this drawing
  const LAYER_MAP = {};
  for (const layerName of Object.keys(parsed.layers || {})) {
    const mapped = autoMapLayer(layerName);
    if (mapped) LAYER_MAP[layerName] = mapped;
  }
  // Also map layers found only in entities (not in layer table)
  for (const e of parsed.entities) {
    if (e.layer && !LAYER_MAP[e.layer]) {
      const mapped = autoMapLayer(e.layer);
      if (mapped) LAYER_MAP[e.layer] = mapped;
    }
  }

  // Polyline areas grouped by mapped layer category
  function shoelaceArea(verts) {
    let area = 0;
    const n = verts.length;
    for (let j = 0; j < n; j++) {
      const k = (j + 1) % n;
      area += verts[j].x * verts[k].y;
      area -= verts[k].x * verts[j].y;
    }
    return Math.abs(area) / 2;
  }

  const wallAreas = {}; // thk_mm → { area_mm2, count }
  const categoryAreas = {}; // cat → { area_mm2, count, label, unit }

  for (const pl of allPolylines) {
    if (pl.vertices.length < 3) continue;
    const mapped = LAYER_MAP[pl.layer];
    if (!mapped) continue;
    const areaMM2 = shoelaceArea(pl.vertices) * u2mm * u2mm;
    if (areaMM2 < 100) continue; // ignore tiny slivers

    const key = mapped.cat + (mapped.thk_mm ? '_' + mapped.thk_mm : '');
    if (!categoryAreas[key]) categoryAreas[key] = { area_mm2: 0, count: 0, label: mapped.label, unit: mapped.unit, thk_mm: mapped.thk_mm, cat: mapped.cat };
    categoryAreas[key].area_mm2 += areaMM2;
    categoryAreas[key].count++;

    if (mapped.cat === 'wall') {
      const thk = mapped.thk_mm;
      if (!wallAreas[thk]) wallAreas[thk] = { area_mm2: 0, count: 0 };
      wallAreas[thk].area_mm2 += areaMM2;
      wallAreas[thk].count++;
    }
  }

  // ─── 6. HATCH COUNTS PER LAYER (from parsed entities) ────────────
  const hatchByLayer = {};
  for (const e of parsed.entities) {
    if (e.type === 'HATCH') {
      hatchByLayer[e.layer] = (hatchByLayer[e.layer] || 0) + 1;
    }
  }

  // Count doors/windows/lifts from INSERTs and layer entities
  let doorCount = 0, windowCount = 0, liftCount = 0, stairCount = 0;
  const doorLayers   = ['AR-DOOR','DOOR'];
  const windowLayers = ['AR-WINDOW','WIN','WINDOW','I - WINDOW'];
  for (const ins of allInserts) {
    const n = ins.block.toUpperCase();
    if (/^D\d+$|DOOR/i.test(n)) doorCount++;
    else if (/^W\d+$|^GD|^SLD|WINDOW/i.test(n)) windowCount++;
    else if (/LIFT|ELEVATOR/.test(n)) liftCount++;
    else if (/STAIR/.test(n)) stairCount++;
  }
  doorLayers.forEach(l => { doorCount += hatchByLayer[l] || 0; });
  windowLayers.forEach(l => { windowCount += hatchByLayer[l] || 0; });

  // ─── 7. DRAWING TYPE ─────────────────────────────────────────────
  const allTextStr = allTexts.map(t => t.text).join(' ').toUpperCase();
  let drawingType = 'UNKNOWN';
  if (/SECTION\s+[A-Z]|\bSECTION\b.*FLOOR|FLOOR.*SECTION/i.test(allTextStr)) drawingType = 'BUILDING_SECTION';
  else if (/\bELEVATION\b/i.test(allTextStr)) drawingType = 'ELEVATION';
  else if (/FLOOR\s+PLAN|\bFLOOR\b.*PLAN/i.test(allTextStr)) drawingType = 'FLOOR_PLAN';
  else if (/FOUNDATION|FOOTING/i.test(allTextStr)) drawingType = 'FOUNDATION';
  else if (/BASEMENT/i.test(allTextStr) && /SECTION/i.test(allTextStr)) drawingType = 'BASEMENT_SECTION';
  if (filename && /section/i.test(filename)) drawingType = drawingType === 'UNKNOWN' ? 'BUILDING_SECTION' : drawingType;

  // ─── 8. PROJECT NAME ─────────────────────────────────────────────
  let projectName = '';
  for (const t of allTexts) {
    if (/MODESTAA|modestaa/i.test(t.text)) { projectName = t.text.replace(/\\.*/,'').trim(); break; }
    if (/project\s*name|project\s*title/i.test(t.text)) { projectName = t.text.trim(); break; }
  }

  // ─── 9. WALL CONSTRUCTION NOTES ─────────────────────────────────
  const wallNoteRE = /(\d+)\s*MM\s*(THK|THICK|WALL|BLOCK|BRICK)|RCC\s*PARDI|COLUMN/i;
  const wallNotes = [...new Set(allTexts.filter(t => wallNoteRE.test(t.text) && t.text.length < 200).map(t => t.text.trim()))];

  // ─── 10. LAYER SUMMARY ───────────────────────────────────────────
  const layerGroups = {};
  for (const e of parsed.entities) {
    const layer = e.layer || '0';
    if (!layerGroups[layer]) layerGroups[layer] = { count: 0, texts: [], polylines: [], inserts: [], lines: [] };
    layerGroups[layer].count++;
  }

  // ─── 11. UNIQUE TEXTS ────────────────────────────────────────────
  const uniqueTexts = [...new Set(allTexts.map(t => t.text.trim()))].filter(Boolean);

  // ─── 12. WALL BOQ ITEMS ─────────────────────────────────────────
  // For section drawing: wall volume = polyline plan area / thk × floor height
  // Better: surface area (sqmt) = plan_area / thk_m | volume (cum) = plan_area
  const wallBOQ = [];
  for (const fl of floorHeights) {
    for (const [thk, data] of Object.entries(wallAreas)) {
      const thkM  = parseInt(thk) / 1000;
      const areaSqm = data.area_mm2 / 1e6;
      const lenM = areaSqm / thkM;
      const volCum = lenM * thkM * fl.height_m;
      const faceSqm = lenM * fl.height_m;
      wallBOQ.push({
        floor:    fl.name,
        thk_mm:   parseInt(thk),
        length_m: Math.round(lenM * 100) / 100,
        height_m: fl.height_m,
        area_sqm: Math.round(faceSqm * 100) / 100,
        vol_cum:  Math.round(volCum * 100) / 100
      });
    }
  }

  // ─── 13. BLOCK COUNTS ────────────────────────────────────────────
  const blockCounts = {};
  for (const ins of allInserts) blockCounts[ins.block] = (blockCounts[ins.block] || 0) + 1;

  // ─── 14. INLINE DIMS ─────────────────────────────────────────────
  const inlineDims = [];
  const dimRE = /(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)/g;
  for (const t of allTexts) {
    let m;
    while ((m = dimRE.exec(t.text)) !== null) {
      const l = parseFloat(m[1]), w = parseFloat(m[2]);
      if (l > 200 && w > 200) inlineDims.push({ label: t.text, length_mm: Math.round(l), width_mm: Math.round(w), layer: t.layer });
    }
  }

  // ─── 15. POLYLINE AREAS (all) ─────────────────────────────────────
  const polylineAreas = allPolylines
    .filter(pl => pl.vertices.length >= 3)
    .map(pl => {
      const areaMM2 = shoelaceArea(pl.vertices) * u2mm * u2mm;
      return { area_sqm: Math.round(areaMM2 / 1e6 * 100) / 100, layer: pl.layer };
    })
    .filter(a => a.area_sqm > 0.01)
    .sort((a, b) => b.area_sqm - a.area_sqm);

  // ─── 16. DIMENSION VALUES ────────────────────────────────────────
  const dimValues = allDims
    .filter(d => d.value_mm > 0)
    .map(d => ({ value_mm: Math.round(d.value_mm * u2mm), value_m: Math.round(d.value_mm * u2mm / 1000 * 100) / 100, layer: d.layer }))
    .sort((a, b) => b.value_mm - a.value_mm);

  return {
    filename,
    drawing_type:    drawingType,
    project_name:    projectName,
    scale,
    scale_factor:    scaleFactor,
    units:           detectedUnit,
    unit_to_mm:      u2mm,
    drawing_extents: {
      width_mm:  Math.round(extW_mm),
      height_mm: Math.round(extH_mm),
      width_m:   Math.round(extW_mm / 1000 * 100) / 100,
      height_m:  Math.round(extH_mm / 1000 * 100) / 100
    },
    floor_levels:    floorLevels,
    floor_heights:   floorHeights,
    wall_notes:      wallNotes,
    wall_boq:        wallBOQ,
    wall_by_thickness: Object.fromEntries(
      Object.entries(wallAreas).map(([thk, d]) => [
        thk + 'mm',
        { area_mm2: Math.round(d.area_mm2), sqm: Math.round(d.area_mm2/1e6*100)/100, count: d.count }
      ])
    ),
    hatch_by_layer:  hatchByLayer,
    hatch_summary:   buildHatchSummary(allHatches),
    all_hatches:     allHatches,
    layer_map:       LAYER_MAP,
    layer_names:     Object.keys(layerGroups).filter(Boolean),
    all_texts:       uniqueTexts,
    room_annotations: [],
    dimension_values: dimValues,
    inline_dims:     inlineDims,
    polyline_areas:  polylineAreas.slice(0, 1000),
    block_counts:    blockCounts,
    element_counts: {
      floor_levels_found: floorLevels.length,
      floor_count:  floorLevels.filter(l => l.mm >= 0).length,
      basement_count: floorLevels.filter(l => l.mm < 0).length,
      door_count:   doorCount,
      window_count: windowCount,
      lift_count:   liftCount,
      staircase_count: stairCount,
      wall_polylines_100mm: wallAreas[100]?.count || 0,
      wall_polylines_230mm: wallAreas[230]?.count || 0,
    },
    stats: {
      total_texts:     allTexts.length,
      total_dims:      allDims.length,
      total_lines:     parsed.entities.filter(e => e.type === 'LINE').length,
      total_polylines: allPolylines.length,
      total_inserts:   allInserts.length,
      total_layers:    Object.keys(layerGroups).length,
      total_hatches:   parsed.entities.filter(e => e.type === 'HATCH').length,
    }
  };
}



// ─────────────────────────────────────────────────────────────────
// HATCH PATTERN → MATERIAL LEGEND
// Maps AutoCAD hatch pattern names (as used in Indian architectural DWGs)
// to their standard civil engineering material meanings.
// ─────────────────────────────────────────────────────────────────
const HATCH_LEGEND = {
  // Brick / Masonry
  'AR-BRSTD':  { material: '230 MM THK. BRICK WALL',   category: 'wall',       symbol_desc: 'Diagonal cross-hatch (standard brick)' },
  'AR-BRELM':  { material: '115 MM THK. BRICK WALL',   category: 'wall',       symbol_desc: 'Diagonal hatch (half-brick)' },
  'ANSI31':    { material: 'BRICK / MASONRY WALL',      category: 'wall',       symbol_desc: '45° diagonal lines' },
  'BRICK':     { material: 'BRICK MASONRY',             category: 'wall',       symbol_desc: 'Brick pattern' },
  // Block Wall
  'AR-BSTONE': { material: '100 MM THK. BLOCK WALL',   category: 'wall',       symbol_desc: 'Stone/block pattern' },
  'BLOCK':     { material: 'BLOCK WALL',                category: 'wall',       symbol_desc: 'Grid block pattern' },
  // Concrete / RCC
  'AR-CONC':   { material: 'R.C.C. / CONCRETE',        category: 'structure',  symbol_desc: 'Gravel/concrete aggregate pattern' },
  'GRAVEL':    { material: 'GRAVEL / PCC',              category: 'structure',  symbol_desc: 'Dot/gravel pattern' },
  'ANSI32':    { material: 'STEEL / RCC REINFORCEMENT', category: 'structure',  symbol_desc: 'Angled steel pattern' },
  // Earth / Fill
  'EARTH':     { material: 'SOIL FILLING / EARTH',      category: 'earthwork',  symbol_desc: 'Earth fill pattern' },
  'SAND':      { material: 'SAND FILLING',              category: 'earthwork',  symbol_desc: 'Sand/dot pattern' },
  'AR-SAND':   { material: 'SAND BED',                  category: 'earthwork',  symbol_desc: 'Sand bed pattern' },
  // Sunk / Depressed Slab
  'ANSI37':    { material: '250 MM SUNK SLAB',          category: 'slab',       symbol_desc: 'Hatch for depressed/sunk area (250mm)' },
  'ANSI36':    { material: '75 MM SUNK SLAB',           category: 'slab',       symbol_desc: 'Hatch for sunk area (75mm)' },
  // Raised Platform
  'AR-RROOF':  { material: 'RAISED PLATFORM (OTLI)',    category: 'platform',   symbol_desc: 'Raised platform / otli pattern' },
  // Insulation / Waterproofing
  'INSUL':     { material: 'INSULATION / WATERPROOFING', category: 'finishing', symbol_desc: 'Insulation pattern' },
  // Wood / Flooring
  'WOOD':      { material: 'WOODEN FLOORING / SHUTTERING', category: 'finishing', symbol_desc: 'Wood grain pattern' },
  'AR-RROOF2': { material: 'ROOF / TERRACE TREATMENT', category: 'structure',  symbol_desc: 'Roof pattern' },
  // Solid Fill
  'SOLID':     { material: 'SOLID FILL / COLUMN',       category: 'structure',  symbol_desc: 'Solid black fill (column/wall section)' },
  // Generic fallback for unmapped patterns
};

function buildHatchSummary(allHatches) {
  // Group by pattern_name × layer, count occurrences
  const summary = {};
  for (const h of allHatches) {
    const key = h.pattern_name || 'UNKNOWN';
    if (!summary[key]) {
      const leg = HATCH_LEGEND[key] || {};
      summary[key] = {
        pattern_name:  key,
        material:      leg.material      || 'REFER DRAWING LEGEND',
        category:      leg.category      || 'unknown',
        symbol_desc:   leg.symbol_desc   || '—',
        count:         0,
        layers:        new Set()
      };
    }
    summary[key].count++;
    summary[key].layers.add(h.layer);
  }
  // Convert Sets to arrays for JSON serialisation
  return Object.values(summary)
    .map(s => ({ ...s, layers: [...s.layers] }))
    .sort((a, b) => b.count - a.count);
}

// ─────────────────────────────────────────────────────────────────
// SECTION 3 — HELPERS
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
// Legend-driven element counter.
//
// Nothing is hardcoded. Every block/layer/text mapping comes from
// legend.json. If the legend is empty or missing a mapping, the
// element becomes an "unknown" surfaced via findUnknowns() — the
// user confirms the meaning once, it is saved, future drawings
// with the same block/layer are auto-recognised.
//
// Backward-compat fields (door_count / window_count / etc.) are
// still produced for downstream consumers, but they are filled
// ONLY from legend mappings — never from regex guessing.
// ─────────────────────────────────────────────────────────────────
function countDrawingElements(allTexts, allInserts, blockCounts, layerNames, legend) {
  const out = {
    door_count:       0,
    window_count:     0,
    lift_count:       0,
    staircase_count:  0,
    column_count:     0,
    beam_count:       0,
    footing_count:    0,
    toilet_count:     0,
    kitchen_count:    0,
    bedroom_count:    0,
    floor_count:      0,
    floor_labels:     [],
    category_counts:  {},       // any user-defined category from legend
    room_types:       {},
    text_extractions: {},       // regex capture results from legend.text_patterns
    source_notes:     []
  };

  const lg = legend || (legendHelper ? legendHelper.emptyLegend() : { blocks:{}, layers:{}, text_patterns:{}, abbreviations:{} });
  const noLegend = Object.keys(lg.blocks).length === 0 && Object.keys(lg.layers).length === 0;

  if (noLegend) {
    out.source_notes.push('Legend empty — run legend_helper.findUnknowns() and fill legend.json, then re-parse.');
    return out;
  }

  // 1. Count from BLOCK INSERTs using legend.blocks mapping
  for (const ins of (allInserts || [])) {
    const mapping = lg.blocks[ins.block];
    if (mapping && !mapping._comment) {
      _bumpCategory(out, mapping);
      continue;
    }
    // Fallback: try layer mapping if block name is unmapped
    const layerMap = lg.layers[ins.layer];
    if (layerMap && !layerMap._comment) {
      _bumpCategory(out, layerMap);
    }
  }

  // 2. Run legend.text_patterns against every text to extract structured data
  //    (floor heights, plinth rises, grid refs, room labels — whatever user defined)
  for (const [key, re] of Object.entries(lg.text_patterns || {})) {
    if (!(re instanceof RegExp)) continue;
    out.text_extractions[key] = [];
  }
  const seenFloorLabels = new Set();
  for (const t of (allTexts || [])) {
    const txt = (t.text || '').trim();
    if (!txt) continue;
    for (const [key, re] of Object.entries(lg.text_patterns || {})) {
      if (!(re instanceof RegExp)) continue;
      const m = txt.match(re);
      if (m) {
        const captured = m[1] !== undefined ? m[1] : m[0];
        out.text_extractions[key].push(captured);
        // Floor-scheme pattern → compute floor_count
        if (/floor.*scheme|scheme.*floor|^floor_scheme$/i.test(key)) {
          const upper = (captured.match(/\+(\d+)/) || [])[1];
          const basement = /^b/i.test(captured);
          const total = 1 + (parseInt(upper) || 0) + (basement ? 1 : 0);
          if (total > out.floor_count) {
            out.floor_count = total;
            out.source_notes.push(`floor_count: ${total} from pattern "${key}" match "${captured}"`);
          }
        }
        // Floor-label pattern → collect distinct labels
        if (/^floor_label$/i.test(key)) {
          const label = captured.toUpperCase().trim();
          if (!seenFloorLabels.has(label)) {
            seenFloorLabels.add(label);
            out.floor_labels.push(label);
          }
        }
        // Room-label pattern → per-room counts
        if (/^room_label$/i.test(key)) {
          const cat = captured.toUpperCase().trim();
          out.room_types[cat] = (out.room_types[cat] || 0) + 1;
          if (cat === 'TOILET' || cat === 'BATH' || cat === 'WC')  out.toilet_count++;
          else if (cat === 'KITCHEN' || cat === 'PANTRY')          out.kitchen_count++;
          else if (/^BED/.test(cat))                               out.bedroom_count++;
        }
      }
    }
  }

  // 3. Fallback floor_count from distinct labels
  if (out.floor_count === 0 && out.floor_labels.length > 0) {
    out.floor_count = out.floor_labels.length;
    out.source_notes.push(`floor_count: ${out.floor_count} from ${out.floor_labels.length} distinct floor_label matches`);
  }

  return out;
}

// Bump the correct backward-compat counter + category_counts based on legend mapping.
function _bumpCategory(out, mapping) {
  const cat = mapping.category;
  if (!cat) return;
  out.category_counts[cat] = (out.category_counts[cat] || 0) + 1;

  if (cat === 'opening') {
    const ot = mapping.opening_type || '';
    if (ot === 'door' || ot === 'glass_door' || ot === 'sliding_door' || ot === 'main_door' || ot === 'fire_door') out.door_count++;
    else if (ot === 'window' || ot === 'kitchen_window')          out.window_count++;
    else if (ot === 'lift_door')                                   out.lift_count++;
    else if (ot === 'ventilator')                                  out.window_count++;
  } else if (cat === 'column')    out.column_count++;
  else if (cat === 'beam')        out.beam_count++;
  else if (cat === 'footing')     out.footing_count++;
  else if (cat === 'staircase')   out.staircase_count++;
  else if (cat === 'lift')        out.lift_count++;
}

// Drawing-type and project-type are NOT inferred from hardcoded keywords
// any more. The user declares them in legend.json (project_type + optional
// drawing_type text pattern). If unset, the value is 'UNKNOWN' and the
// parser surfaces that as a user question.
function detectProjectType(allTexts, filename, floorCount, extAreaSqm, legendArg) {
  const legend = legendArg || (legendHelper ? legendHelper.loadLegend() : null);
  return {
    type: (legend && legend.project_type) ? legend.project_type : 'unknown',
    scores: {},
    notes: legend && legend.project_type
      ? [`project_type from legend.json = "${legend.project_type}"`]
      : ['project_type not set in legend.json — please set it (high_rise_residential / commercial / institute / cafe / road_site / industrial / other)']
  };
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

// ═══════════════════════════════════════════════════════════════════
// COORDINATE CLUSTERING — Table Reconstruction from DXF texts
// ─────────────────────────────────────────────────────────────────
// texts: array of {text, x, y} (from parseDXF or extractCivilData)
// yTol:  Y-band tolerance in drawing units (default 50 = 50mm)
// xTol:  X-band tolerance in drawing units (default 80 = 80mm)
//
// Returns: array of rows, each row = array of {text, x, y}
// Same Y band → same row. Within row, sorted by X → columns.
//
// Example output used to reconstruct footing/column schedule tables.
// ═══════════════════════════════════════════════════════════════════
function clusterTextsToTable(texts, yTol = 150, xTol = 80) {
  if (!texts || !texts.length) return [];

  // Filter out empty/whitespace texts
  const valid = texts.filter(t => t.text && t.text.trim().length > 0);
  if (!valid.length) return [];

  // Sort by Y descending (top of drawing first, DXF Y increases upward)
  const sorted = [...valid].sort((a, b) => b.y - a.y);

  // Group into Y-bands (rows)
  const rows = [];
  for (const t of sorted) {
    const existingRow = rows.find(r => Math.abs(r.yCenter - t.y) <= yTol);
    if (existingRow) {
      existingRow.cells.push(t);
      // Update row Y center as average
      existingRow.yCenter = existingRow.cells.reduce((s, c) => s + c.y, 0) / existingRow.cells.length;
    } else {
      rows.push({ yCenter: t.y, cells: [t] });
    }
  }

  // Within each row, sort cells by X (left to right)
  for (const row of rows) {
    row.cells.sort((a, b) => a.x - b.x);
  }

  // Return rows sorted top-to-bottom
  rows.sort((a, b) => b.yCenter - a.yCenter);
  return rows.map(r => r.cells);
}

// Detect schedule tables: find rectangular grids of text
// Returns array of named tables: { name, headers, rows }
function reconstructScheduleTables(allTexts, scaleFactor = 1, yTol = 50) {
  if (!allTexts || !allTexts.length) return [];

  // Cluster all texts into rows
  const rows = clusterTextsToTable(allTexts, yTol);
  if (rows.length < 2) return [];

  const tables = [];
  let i = 0;

  while (i < rows.length) {
    const headerRow = rows[i];
    // Detect header: row where cells contain keywords like MARK, SIZE, NOS, STEEL, TYPE, FOOTING, COLUMN
    const headerTexts = headerRow.map(c => c.text.toUpperCase());
    const isHeader = headerTexts.some(t =>
      /^(MARK|TYPE|SIZE|NOS|NO\.|NO|QTY|QUANTITY|STEEL|BAR|DIA|FOOTING|COLUMN|COL|BEAM|SLAB|DEPTH|WIDTH|LENGTH|SPACING|STIRRUP|REINF|DETAIL)$/.test(t.replace(/[^A-Z.]/g,''))
    );

    if (isHeader && headerRow.length >= 2) {
      // Collect data rows below this header until gap or new header
      const dataRows = [];
      let j = i + 1;
      while (j < rows.length) {
        const nextRow = rows[j];
        // Stop if Y gap is too large (table ended)
        const yGap = Math.abs(rows[j-1][0].y - nextRow[0].y);
        if (yGap > yTol * 4) break;
        // Stop if this row also looks like a header
        const nextTexts = nextRow.map(c => c.text.toUpperCase());
        const nextIsHeader = nextTexts.some(t =>
          /^(MARK|TYPE|SIZE|NOS|FOOTING|COLUMN|BEAM|SLAB)$/.test(t.replace(/[^A-Z]/g,''))
        );
        if (nextIsHeader && dataRows.length > 0) break;
        dataRows.push(nextRow.map(c => c.text.trim()));
        j++;
      }

      if (dataRows.length > 0) {
        // Find table name from cell above header or from header itself
        const tableName = headerTexts.find(t => /FOOTING|COLUMN|BEAM|SLAB|SCHEDULE/.test(t)) || 'SCHEDULE';
        tables.push({
          name: tableName,
          headers: headerRow.map(c => c.text.trim()),
          rows: dataRows,
          // Also output as JSON objects keyed by header
          records: dataRows.map(row =>
            Object.fromEntries(headerRow.map((h, idx) => [h.text.trim(), row[idx] || '']))
          )
        });
        i = j;
        continue;
      }
    }
    i++;
  }

  return tables;
}

// High-level: given civilData, extract all schedule tables and attach
function attachScheduleTables(civilData) {
  if (!civilData || !civilData.all_texts) return civilData;
  const sf = civilData.scale_factor || 1;
  // Use a Y tolerance based on typical text height (default 50 drawing units)
  const yTol = Math.max(30, (civilData.drawing_extents?.height_m || 10) * 1000 * 0.01);
  civilData.schedule_tables = reconstructScheduleTables(civilData.all_texts, sf, yTol);
  civilData.clustered_rows  = clusterTextsToTable(civilData.all_texts, yTol);
  return civilData;
}

module.exports = { parseDXF, extractCivilData, extractTotalAreaSqft, detectProjectType, RATES, clusterTextsToTable, reconstructScheduleTables, attachScheduleTables };
