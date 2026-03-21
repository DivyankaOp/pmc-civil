/**
 * PMC DXF Parser — Pure JavaScript, no dependencies
 * Reads AutoCAD/ZWCAD DXF files and extracts:
 * - All entities (LINE, TEXT, MTEXT, DIMENSION, LWPOLYLINE, CIRCLE, ARC)
 * - Layer names with their entities
 * - Title block info (project name, scale, date)
 * - All text annotations and dimension values
 */

function parseDXF(dxfContent) {
  const lines = dxfContent.split(/\r?\n/);
  const result = {
    version: '',
    units: 'mm',
    layers: {},
    entities: [],
    texts: [],
    dimensions: [],
    polylines: [],
    blocks: {},
    title_block: {},
    extents: { xmin: Infinity, xmax: -Infinity, ymin: Infinity, ymax: -Infinity }
  };

  let i = 0;
  const next = () => { const v = lines[i]?.trim(); i++; return v; };
  const peek = () => lines[i]?.trim();

  // ── Parse group code/value pairs ──────────────────────────────
  function readPairs() {
    const pairs = {};
    while (i < lines.length) {
      const code = next();
      const value = next();
      if (code === undefined || value === undefined) break;
      const codeNum = parseInt(code);
      if (!isNaN(codeNum)) {
        pairs[codeNum] = value;
        if (codeNum === 0) { i -= 2; break; } // Entity type marker
      }
    }
    return pairs;
  }

  // ── Parse ENTITIES section ─────────────────────────────────────
  function parseEntities(inBlock = false) {
    while (i < lines.length) {
      const code = parseInt(next());
      const value = next();

      if (isNaN(code)) continue;
      if (code !== 0) continue;

      const entityType = value?.trim().toUpperCase();

      if (entityType === 'ENDSEC' || entityType === 'ENDBLK') break;

      if (entityType === 'TEXT' || entityType === 'MTEXT') {
        const e = readTextEntity(entityType);
        if (e) {
          result.texts.push(e);
          result.entities.push({ type: entityType, ...e });
        }
      } else if (entityType === 'DIMENSION') {
        const e = readDimEntity();
        if (e) {
          result.dimensions.push(e);
          result.entities.push({ type: 'DIMENSION', ...e });
        }
      } else if (entityType === 'LWPOLYLINE' || entityType === 'POLYLINE') {
        const e = readPolyline(entityType);
        if (e) {
          result.polylines.push(e);
          updateExtents(e.vertices);
        }
      } else if (entityType === 'LINE') {
        const e = readLine();
        if (e) {
          result.entities.push({ type: 'LINE', ...e });
          updateExtents([{x: e.x1, y: e.y1}, {x: e.x2, y: e.y2}]);
        }
      } else if (entityType === 'INSERT') {
        readInsert(); // block reference
      } else if (entityType === 'CIRCLE' || entityType === 'ARC') {
        skipEntity();
      } else if (entityType === 'HATCH' || entityType === 'SOLID') {
        skipEntity();
      }
    }
  }

  function updateExtents(vertices) {
    for (const v of vertices) {
      if (v.x < result.extents.xmin) result.extents.xmin = v.x;
      if (v.x > result.extents.xmax) result.extents.xmax = v.x;
      if (v.y < result.extents.ymin) result.extents.ymin = v.y;
      if (v.y > result.extents.ymax) result.extents.ymax = v.y;
    }
  }

  function readTextEntity(type) {
    const props = {};
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); const val = next();
      if (code === 1)  props.text = val;      // actual text content
      if (code === 3)  props.text = (props.text||'') + val; // MTEXT continuation
      if (code === 10) props.x = parseFloat(val);
      if (code === 20) props.y = parseFloat(val);
      if (code === 40) props.height = parseFloat(val);
      if (code === 8)  props.layer = val;
      if (code === 50) props.rotation = parseFloat(val);
    }
    // Clean MTEXT formatting codes
    if (props.text) {
      props.text = props.text
        .replace(/\\P/g, '\n')
        .replace(/\{[^}]*\}/g, '')
        .replace(/\\[a-zA-Z][^;]*;/g, '')
        .trim();
    }
    return props.text ? props : null;
  }

  function readDimEntity() {
    const props = {};
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); const val = next();
      if (code === 1)  props.dimtext = val;         // override text
      if (code === 42) props.actual_measurement = parseFloat(val); // actual dim value
      if (code === 10) props.x1 = parseFloat(val);
      if (code === 20) props.y1 = parseFloat(val);
      if (code === 13) props.x2 = parseFloat(val);
      if (code === 23) props.y2 = parseFloat(val);
      if (code === 11) props.xtext = parseFloat(val);
      if (code === 21) props.ytext = parseFloat(val);
      if (code === 8)  props.layer = val;
      if (code === 3)  props.dimstyle = val;
    }
    return props;
  }

  function readPolyline(type) {
    const props = { vertices: [], layer: '' };
    let cx = 0, cy = 0;
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); const val = next();
      if (code === 8)  props.layer = val;
      if (code === 90) props.vertex_count = parseInt(val);
      if (type === 'LWPOLYLINE') {
        if (code === 10) { cx = parseFloat(val); }
        if (code === 20) { cy = parseFloat(val); props.vertices.push({x: cx, y: cy}); }
      }
    }
    return props.vertices.length > 0 ? props : null;
  }

  function readLine() {
    const props = {};
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); const val = next();
      if (code === 10) props.x1 = parseFloat(val);
      if (code === 20) props.y1 = parseFloat(val);
      if (code === 11) props.x2 = parseFloat(val);
      if (code === 21) props.y2 = parseFloat(val);
      if (code === 8)  props.layer = val;
    }
    return (props.x1 !== undefined) ? props : null;
  }

  function readInsert() {
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); next();
    }
  }

  function skipEntity() {
    while (i < lines.length) {
      const code = parseInt(peek());
      if (isNaN(code)) { next(); next(); continue; }
      if (code === 0) break;
      next(); next();
    }
  }

  // ── Main parse loop ────────────────────────────────────────────
  while (i < lines.length) {
    const code = parseInt(next());
    const value = next();
    if (isNaN(code) || value === undefined) continue;

    if (code === 0 && value === 'SECTION') {
      const secCode = parseInt(next());
      const secName = next();
      if (secName === 'HEADER') {
        // Parse header for units and extents
        while (i < lines.length) {
          const c = parseInt(peek());
          if (isNaN(c)) { next(); next(); continue; }
          next(); const v = next();
          if (c === 9 && v === '$INSUNITS') {
            const uc = parseInt(next()); const uv = next();
            if (uv === '4') result.units = 'mm';
            else if (uv === '5') result.units = 'cm';
            else if (uv === '6') result.units = 'm';
          }
          if (c === 0 && v === 'ENDSEC') break;
        }
      } else if (secName === 'TABLES') {
        // Parse layers
        while (i < lines.length) {
          const c = parseInt(peek());
          if (isNaN(c)) { next(); next(); continue; }
          next(); const v = next();
          if (c === 0 && v === 'LAYER') {
            let layerName = '';
            while (i < lines.length) {
              const lc = parseInt(peek());
              if (isNaN(lc)) { next(); next(); continue; }
              if (lc === 0) break;
              next(); const lv = next();
              if (lc === 2) layerName = lv;
            }
            if (layerName) result.layers[layerName] = { entities: [] };
          }
          if (c === 0 && v === 'ENDSEC') break;
        }
      } else if (secName === 'ENTITIES') {
        parseEntities();
      } else if (secName === 'BLOCKS') {
        // Parse blocks (contains reusable geometry like title blocks)
        let currentBlock = null;
        while (i < lines.length) {
          const c = parseInt(peek());
          if (isNaN(c)) { next(); next(); continue; }
          next(); const v = next();
          if (c === 0 && v === 'BLOCK') {
            currentBlock = { texts: [], name: '' };
          }
          if (c === 2 && currentBlock) currentBlock.name = v;
          if (c === 0 && v === 'TEXT') {
            const t = readTextEntity('TEXT');
            if (t && currentBlock) currentBlock.texts.push(t);
          }
          if (c === 0 && v === 'MTEXT') {
            const t = readTextEntity('MTEXT');
            if (t && currentBlock) currentBlock.texts.push(t);
          }
          if (c === 0 && v === 'ENDBLK' && currentBlock) {
            if (currentBlock.name) result.blocks[currentBlock.name] = currentBlock;
            currentBlock = null;
          }
          if (c === 0 && v === 'ENDSEC') break;
        }
      }
    }
  }

  return result;
}

// ── EXTRACT CIVIL DATA FROM PARSED DXF ──────────────────────────
function extractCivilData(parsed, filename) {
  const allTexts = [
    ...parsed.texts,
    ...Object.values(parsed.blocks).flatMap(b => b.texts || [])
  ];

  // ── Scale detection ─────────────────────────────────────────
  let scale = null;
  let scaleFactor = 1; // mm per drawing unit

  for (const t of allTexts) {
    if (!t.text) continue;
    const m = t.text.match(/1\s*:\s*(\d+)/);
    if (m) { scale = `1:${m[1]}`; scaleFactor = parseInt(m[1]); break; }
  }

  // Auto-detect from extents if no scale found
  const extW = parsed.extents.xmax - parsed.extents.xmin;
  const extH = parsed.extents.ymax - parsed.extents.ymin;

  // ── Title block extraction ──────────────────────────────────
  const titleBlock = {};
  const titleKeywords = {
    project: /project|work|name|title/i,
    drawing_no: /drg\.?\s*no|drawing\s*no|dwg/i,
    date: /date|dt/i,
    scale: /scale/i,
    prepared: /prepared|designed|drawn/i,
    location: /location|site|place/i
  };

  // Cluster texts that are near each other (title block region)
  const sortedTexts = [...allTexts].sort((a, b) => (a.y || 0) - (b.y || 0));

  for (const t of allTexts) {
    if (!t.text) continue;
    for (const [key, regex] of Object.entries(titleKeywords)) {
      if (regex.test(t.text) && !titleBlock[key]) {
        titleBlock[key] = t.text;
      }
    }
  }

  // ── Dimension extraction ────────────────────────────────────
  const dimValues = parsed.dimensions
    .filter(d => d.actual_measurement && d.actual_measurement > 0)
    .map(d => ({
      value_mm: d.actual_measurement,
      value_m: d.actual_measurement / 1000,
      text: d.dimtext || d.actual_measurement.toFixed(2),
      layer: d.layer || ''
    }));

  // ── Layer-wise geometry analysis ────────────────────────────
  const layerGroups = {};
  for (const e of parsed.entities) {
    const layer = e.layer || 'DEFAULT';
    if (!layerGroups[layer]) layerGroups[layer] = { texts: [], lines: [], polylines: [], dims: [] };
    if (e.type === 'TEXT' || e.type === 'MTEXT') layerGroups[layer].texts.push(e.text);
    if (e.type === 'LINE') layerGroups[layer].lines.push(e);
    if (e.type === 'DIMENSION') layerGroups[layer].dims.push(e.actual_measurement);
  }

  // ── Extract road/space data from text annotations ───────────
  const spacesFound = [];
  const roadPattern = /([A-Z]+-?\d+|ROAD\s*[A-Z0-9-]+|R\s*[-=]?\s*\d+)/gi;
  const dimPattern = /(\d+\.?\d*)\s*[Xx×]\s*(\d+\.?\d*)/g;
  const areaPattern = /(\d+[\.,]?\d*)\s*(sqmt|sqm|sq\.?m|m2|sqft|sft)/gi;
  const lengthPattern = /(\d+[\.,]?\d*)\s*(rmt|rmt|m|mt|meter|mtr)/gi;

  for (const t of allTexts) {
    if (!t.text) continue;
    const txt = t.text;

    // Extract L×W dimensions
    let m;
    while ((m = dimPattern.exec(txt)) !== null) {
      const l = parseFloat(m[1]), w = parseFloat(m[2]);
      if (l > 0 && w > 0 && l < 5000 && w < 5000) {
        spacesFound.push({ label: txt, length: l, width: w, area: l * w, source: 'annotation' });
      }
    }
  }

  // ── Calculate polyline areas (closed regions = rooms/plots) ──
  const polylineAreas = [];
  for (const pl of parsed.polylines) {
    if (pl.vertices.length >= 3) {
      const area = Math.abs(shoelaceArea(pl.vertices));
      if (area > 100) { // filter tiny artifacts
        const unitArea = area / (scaleFactor * scaleFactor); // convert to m²
        polylineAreas.push({
          area_sqm: Math.round(unitArea * 100) / 100,
          perimeter_m: Math.round(perimeter(pl.vertices) / scaleFactor * 100) / 100,
          layer: pl.layer,
          vertices: pl.vertices.length
        });
      }
    }
  }
  polylineAreas.sort((a, b) => b.area_sqm - a.area_sqm);

  // ── Line length analysis ─────────────────────────────────────
  const lineLengths = [];
  for (const e of parsed.entities) {
    if (e.type === 'LINE') {
      const len = Math.sqrt((e.x2-e.x1)**2 + (e.y2-e.y1)**2) / scaleFactor;
      if (len > 0.5) lineLengths.push({ length_m: Math.round(len*100)/100, layer: e.layer || '' });
    }
  }

  return {
    filename: filename || 'drawing.dxf',
    scale,
    scale_factor: scaleFactor,
    units: parsed.units,
    drawing_extents: {
      width_units: Math.round(extW),
      height_units: Math.round(extH),
      estimated_width_m: Math.round(extW / scaleFactor * 100) / 100,
      estimated_height_m: Math.round(extH / scaleFactor * 100) / 100
    },
    title_block: titleBlock,
    all_texts: allTexts.filter(t => t.text).map(t => t.text).filter((v, i, a) => a.indexOf(v) === i),
    layer_names: Object.keys(layerGroups).filter(l => l.trim()),
    dimension_values: dimValues,
    spaces_from_annotations: spacesFound,
    polyline_areas: polylineAreas.slice(0, 50),
    line_lengths: lineLengths.slice(0, 200),
    stats: {
      total_texts: allTexts.length,
      total_dims: parsed.dimensions.length,
      total_lines: parsed.entities.filter(e => e.type === 'LINE').length,
      total_polylines: parsed.polylines.length,
      total_layers: Object.keys(layerGroups).length
    }
  };
}

// ── Math helpers ─────────────────────────────────────────────────
function shoelaceArea(pts) {
  let area = 0;
  for (let i = 0; i < pts.length; i++) {
    const j = (i + 1) % pts.length;
    area += pts[i].x * pts[j].y;
    area -= pts[j].x * pts[i].y;
  }
  return area / 2;
}

function perimeter(pts) {
  let p = 0;
  for (let i = 0; i < pts.length; i++) {
    const j = (i + 1) % pts.length;
    p += Math.sqrt((pts[j].x - pts[i].x) ** 2 + (pts[j].y - pts[i].y) ** 2);
  }
  return p;
}

module.exports = { parseDXF, extractCivilData };
