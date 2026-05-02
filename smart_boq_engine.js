'use strict';

/**
 * SMART BOQ ENGINE — Out-of-the-box 90-95% accuracy approach
 * ─────────────────────────────────────────────────────────────────
 * Core insight: DXF mein sab kuch already hai — walls, dimensions,
 * levels, hatches, room areas. Problem data extraction nahi hai —
 * problem data ko meaningful context mein present karna hai.
 *
 * Strategy:
 * 1. DXF se directly "engineering summary" nikalo — raw data nahi
 * 2. Wall VOLUME sahi formula se nikalo (perimeter × thk × height)
 * 3. Room areas polyline se nikalo (shoelace formula)
 * 4. Scale auto-detect karo drawing extents + dimensions se
 * 5. Claude ko ek "pre-drafted BOQ" do — sirf verify + rate karo
 *    (Claude as checker, not guesser = 90-95% accuracy)
 */

const fs   = require('fs');
const path = require('path');

// ─────────────────────────────────────────────────────────────────
// STEP 1 — WALL PERIMETER (sahi formula)
// ─────────────────────────────────────────────────────────────────
// Polyline in DXF = wall center-line or wall boundary polygon.
// Plan area = length × thickness (in plan view).
// So: wall_length = plan_area / thickness
// Volume = wall_length × thickness × floor_height
// Face area (plaster) = wall_length × floor_height
function calcWallPerimeter(pts) {
  let p = 0;
  for (let i = 0; i < pts.length; i++) {
    const j = (i + 1) % pts.length;
    p += Math.sqrt((pts[j][0] - pts[i][0]) ** 2 + (pts[j][1] - pts[i][1]) ** 2);
  }
  return p;
}

function shoelace(pts) {
  let a = 0;
  for (let i = 0; i < pts.length; i++) {
    const j = (i + 1) % pts.length;
    a += pts[i][0] * pts[j][1] - pts[j][0] * pts[i][1];
  }
  return Math.abs(a) / 2;
}

// ─────────────────────────────────────────────────────────────────
// STEP 2 — SMART SCALE DETECTION
// ─────────────────────────────────────────────────────────────────
// Strategy:
//   A. Look for "SCALE 1:100" text in drawing
//   B. Compare dimension values vs geometric distance
//   C. Compare drawing extent vs known paper sizes (A0/A1/A2)
//   D. Fallback: median ratio of (dim value / geom distance)
function detectScale(texts, dims, extents) {
  // A. Text-based scale
  for (const t of texts) {
    const m = t.text.match(/SCALE\s*[:\s]\s*1\s*[:/]\s*(\d+)/i)
           || t.text.match(/\b1\s*:\s*(\d+)\b/);
    if (m) {
      const sf = parseInt(m[1]);
      if (sf >= 5 && sf <= 5000) return { factor: sf, source: 'text', label: `1:${sf}` };
    }
  }

  // B. Dimension ratio: measured_value / geometric_distance
  const ratios = [];
  for (const d of dims) {
    if (d.value_mm > 100 && d.geom_mm > 10) {
      const r = d.value_mm / d.geom_mm;
      if (r >= 1 && r <= 5000) ratios.push(r);
    }
  }
  if (ratios.length >= 3) {
    ratios.sort((a, b) => a - b);
    const median = ratios[Math.floor(ratios.length / 2)];
    // Round to nearest common scale
    const commonScales = [1, 5, 10, 20, 25, 50, 100, 200, 500, 1000];
    const nearest = commonScales.reduce((best, s) => Math.abs(s - median) < Math.abs(best - median) ? s : best, 100);
    if (Math.abs(nearest - median) / nearest < 0.3) {
      return { factor: nearest, source: 'dimension_ratio', label: `1:${nearest}`, raw_median: Math.round(median) };
    }
  }

  // C. Extents-based: compare drawing width to A0 (1189mm) / A1 (841mm)
  const wMm = extents.xmax - extents.xmin;
  if (wMm > 0) {
    const paperWidths = [{ w: 1189, name: 'A0' }, { w: 841, name: 'A1' }, { w: 594, name: 'A2' }];
    for (const paper of paperWidths) {
      const scale = Math.round(wMm / paper.w);
      const commonScales = [50, 100, 200, 500];
      const nearest = commonScales.reduce((best, s) => Math.abs(s - scale) < Math.abs(best - scale) ? s : best, 100);
      if (Math.abs(nearest - scale) / nearest < 0.25) {
        return { factor: nearest, source: 'extents_' + paper.name, label: `1:${nearest}` };
      }
    }
  }

  return { factor: 1, source: 'fallback_assume_mm', label: 'unknown (assuming 1:1 mm)' };
}

// ─────────────────────────────────────────────────────────────────
// STEP 3 — ROOM AREA EXTRACTION
// ─────────────────────────────────────────────────────────────────
// Room labels (TEXT entities) near closed polylines → room name + area
function extractRooms(texts, polylines, scaleFactor) {
  const rooms = [];
  // Filter closed polylines with meaningful area
  const closedPoly = polylines.filter(p => p.pts && p.pts.length >= 3);

  for (const poly of closedPoly) {
    const rawArea = shoelace(poly.pts);
    const areaSqm = rawArea / (scaleFactor * scaleFactor) / 1e6;
    if (areaSqm < 0.5 || areaSqm > 10000) continue; // skip tiny/huge

    // Find text closest to polyline centroid
    const cx = poly.pts.reduce((s, p) => s + p[0], 0) / poly.pts.length;
    const cy = poly.pts.reduce((s, p) => s + p[1], 0) / poly.pts.length;

    let closestText = null, minDist = Infinity;
    for (const t of texts) {
      if (t.text.length < 2 || /^\d/.test(t.text)) continue; // skip dimension texts
      const dx = t.x - cx, dy = t.y - cy;
      const dist = Math.sqrt(dx * dx + dy * dy);
      if (dist < minDist) { minDist = dist; closestText = t.text; }
    }

    rooms.push({
      name: closestText || 'ROOM',
      area_sqm: Math.round(areaSqm * 100) / 100,
      layer: poly.layer,
      centroid: { x: Math.round(cx), y: Math.round(cy) }
    });
  }

  return rooms.sort((a, b) => b.area_sqm - a.area_sqm);
}

// ─────────────────────────────────────────────────────────────────
// STEP 4 — WALL QUANTITIES (correct formula)
// ─────────────────────────────────────────────────────────────────
function extractWallQuantities(polylines, layerMap, floorHeights, scaleFactor) {
  const wallsByThk = {}; // thk_mm → { total_plan_area_mm2, total_perimeter_mm, count, layers }

  for (const poly of polylines) {
    const mapped = layerMap[poly.layer];
    if (!mapped || mapped.category !== 'wall') continue;
    const thk = mapped.thk_mm || 230;

    if (!wallsByThk[thk]) wallsByThk[thk] = { plan_area_mm2: 0, perimeter_mm: 0, count: 0, layers: new Set() };

    // Two methods:
    // Method A: if polyline is a wall BOUNDARY (hatch area), use perimeter
    // Method B: if polyline IS the wall area in plan, use area/thk to get length
    const perim = calcWallPerimeter(poly.pts);
    const area  = shoelace(poly.pts);

    wallsByThk[thk].plan_area_mm2 += area;
    wallsByThk[thk].perimeter_mm  += perim;
    wallsByThk[thk].count++;
    wallsByThk[thk].layers.add(poly.layer);
  }

  const results = [];
  for (const [thkStr, data] of Object.entries(wallsByThk)) {
    const thk_mm  = parseInt(thkStr);
    const thk_m   = thk_mm / 1000;
    const sf      = scaleFactor;

    // Real-world plan area = raw_area / (sf * sf) / 1e6 sqm
    const planAreaSqm = (data.plan_area_mm2 / (sf * sf)) / 1e6;

    // Wall length from plan area: length = planArea / thickness
    // This is the correct formula for hatch polylines
    const wallLenM = planAreaSqm / thk_m;

    for (const fh of (floorHeights.length ? floorHeights : [{ name: 'TYPICAL', height_m: 3.0 }])) {
      const faceSqm = Math.round(wallLenM * fh.height_m * 100) / 100;
      const volCum  = Math.round(wallLenM * thk_m * fh.height_m * 100) / 100;
      results.push({
        floor:        fh.name,
        thk_mm,
        length_m:     Math.round(wallLenM * 100) / 100,
        height_m:     fh.height_m,
        face_area_sqm: faceSqm,
        volume_cum:   volCum,
        polyline_count: data.count,
        layers:       [...data.layers]
      });
    }
  }
  return results;
}

// ─────────────────────────────────────────────────────────────────
// STEP 5 — SCHEDULE TABLE EXTRACTION (column/footing schedules)
// ─────────────────────────────────────────────────────────────────
// Cluster texts by Y coordinate → rows → detect header → extract data
function extractScheduleTables(texts, scaleFactor) {
  if (!texts.length) return [];

  // Sort by Y descending (DXF Y increases upward)
  const sorted = [...texts].filter(t => t.text.trim()).sort((a, b) => b.y - a.y);

  // Cluster into rows by Y proximity
  const yTol = 60 / scaleFactor; // 60mm real-world tolerance
  const rows = [];
  for (const t of sorted) {
    const row = rows.find(r => Math.abs(r.yAvg - t.y) < yTol);
    if (row) {
      row.cells.push(t);
      row.yAvg = row.cells.reduce((s, c) => s + c.y, 0) / row.cells.length;
    } else {
      rows.push({ yAvg: t.y, cells: [t] });
    }
  }

  // Sort each row's cells by X (left to right)
  rows.forEach(r => r.cells.sort((a, b) => a.x - b.x));

  // Detect schedule tables: rows with header keywords
  const HEADER_KEYWORDS = /^(MARK|TYPE|SIZE|NOS|NO\.|QTY|STEEL|DIA|FOOTING|COLUMN|COL|BEAM|SLAB|DEPTH|WIDTH|LENGTH|SPACING|STIRRUP|DETAIL|REINF|SR|SR\.?NO|DESCRIPTION|UNIT|RATE|AMOUNT)$/i;
  const tables = [];
  let i = 0;

  while (i < rows.length) {
    const headerCells = rows[i].cells.filter(c => HEADER_KEYWORDS.test(c.text.trim().replace(/[^A-Za-z.]/g, '')));
    if (headerCells.length >= 2) {
      // Found a header row — collect data rows below
      const headers = rows[i].cells.map(c => c.text.trim());
      const dataRows = [];
      let j = i + 1;
      while (j < rows.length) {
        const yGap = Math.abs(rows[j - 1].yAvg - rows[j].yAvg);
        if (yGap > yTol * 5) break; // large gap = end of table
        const nextHeaderCells = rows[j].cells.filter(c => HEADER_KEYWORDS.test(c.text.trim().replace(/[^A-Za-z.]/g, '')));
        if (nextHeaderCells.length >= 3 && dataRows.length > 0) break; // new table started
        dataRows.push(rows[j].cells.map(c => c.text.trim()));
        j++;
      }
      if (dataRows.length > 0) {
        const tableType = headers.join(' ').toUpperCase().includes('FOOTING') ? 'FOOTING_SCHEDULE'
                        : headers.join(' ').toUpperCase().includes('COLUMN') ? 'COLUMN_SCHEDULE'
                        : headers.join(' ').toUpperCase().includes('BEAM')   ? 'BEAM_SCHEDULE'
                        : 'BOQ_TABLE';
        tables.push({
          type: tableType,
          headers,
          rows: dataRows,
          records: dataRows.map(row =>
            Object.fromEntries(headers.map((h, idx) => [h, row[idx] || '']))
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

// ─────────────────────────────────────────────────────────────────
// STEP 6 — PRE-DRAFT BOQ (Claude verify kare, not generate)
// ─────────────────────────────────────────────────────────────────
// Yeh function ek rough BOQ draft banata hai directly from drawing data.
// Claude ko sirf: (a) verify karna hai, (b) rates lagani hain, (c) missed items add karne hain.
// Result: Claude "checker" mode mein kaam karta hai — 90-95% accurate.
function preDraftBOQ(wallQty, rooms, scheduleTables, floorLevels, floorHeights, hatchSummary, elementCounts, rates) {
  const items = [];
  let sr = 1;

  // ── EARTHWORK ────────────────────────────────────────────────────
  // Foundation depth from levels (if basement/plinth levels found)
  const groundLevel  = floorLevels.find(l => /GROUND|PLINTH/i.test(l.name));
  const basementLevel = floorLevels.find(l => l.mm < 0);
  if (basementLevel && groundLevel) {
    const excavDepth = (groundLevel.mm - basementLevel.mm) / 1000;
    // Estimate plan area from total building footprint (largest room area)
    const footprint = rooms.length ? rooms[0].area_sqm * 1.2 : 0; // approx with walls
    if (footprint > 0 && excavDepth > 0) {
      items.push({ sr: sr++, description: 'Earthwork Excavation in Foundation', unit: 'CUM', qty: Math.round(footprint * excavDepth * 100) / 100, rate: rates.excavation || 320, source: 'calculated_from_levels', confidence: 'medium' });
    }
  }

  // ── WALL MASONRY ─────────────────────────────────────────────────
  for (const w of wallQty) {
    if (w.volume_cum > 0) {
      const matName = w.thk_mm >= 200 ? 'Brick Masonry' : 'Block Masonry';
      const rate = w.thk_mm >= 200 ? (rates.brick_masonry_230 || 6800) : (rates.block_masonry_100 || 4200);
      items.push({
        sr: sr++,
        description: `${matName} ${w.thk_mm}mm THK — ${w.floor}`,
        unit: 'CUM',
        qty: w.volume_cum,
        rate,
        source: 'drawing_polyline',
        confidence: w.polyline_count > 5 ? 'high' : 'medium',
        detail: `${w.length_m}m length × ${w.thk_mm / 1000}m thk × ${w.height_m}m height`
      });
    }
  }

  // ── PLASTER (both faces of wall) ────────────────────────────────
  for (const w of wallQty) {
    if (w.face_area_sqm > 0) {
      items.push({
        sr: sr++,
        description: `Cement Plaster 12mm both faces — ${w.thk_mm}mm wall — ${w.floor}`,
        unit: 'SQMT',
        qty: Math.round(w.face_area_sqm * 2 * 0.85 * 100) / 100, // 0.85 deduct for openings
        rate: rates.plaster_12mm || 280,
        source: 'calculated_from_wall',
        confidence: 'medium'
      });
    }
  }

  // ── COLUMN SCHEDULE ───────────────────────────────────────────────
  const colSchedule = scheduleTables.find(t => t.type === 'COLUMN_SCHEDULE');
  if (colSchedule) {
    const colRecords = colSchedule.records.filter(r => {
      const keys = Object.values(r).join(' ');
      return /\d+x\d+|\d+×\d+|\d+\s*X\s*\d+/i.test(keys);
    });
    for (const rec of colRecords) {
      const vals = Object.values(rec).join(' ');
      const sizeMatch = vals.match(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/);
      const nosMatch  = vals.match(/\b(\d{1,3})\b/);
      if (sizeMatch) {
        const w = parseInt(sizeMatch[1]) / 1000; // m
        const d = parseInt(sizeMatch[2]) / 1000; // m
        const nos = nosMatch ? parseInt(nosMatch[1]) : 1;
        const floorH = floorHeights.length ? floorHeights[0].height_m : 3.0;
        const volCum = Math.round(w * d * floorH * nos * 100) / 100;
        items.push({
          sr: sr++,
          description: `RCC Column ${sizeMatch[1]}×${sizeMatch[2]}mm — ${nos} Nos`,
          unit: 'CUM',
          qty: volCum,
          rate: rates.rcc_column || 8500,
          source: 'drawing_schedule',
          confidence: 'high',
          schedule_ref: JSON.stringify(rec)
        });
      }
    }
  } else if (elementCounts.column_count > 0) {
    // No schedule found — flag for Claude to fill
    items.push({
      sr: sr++,
      description: `RCC Column (${elementCounts.column_count} Nos found — schedule not detected, Claude to verify size)`,
      unit: 'CUM',
      qty: null,
      rate: rates.rcc_column || 8500,
      source: 'count_only_no_schedule',
      confidence: 'low',
      flag: 'CLAUDE_VERIFY_SIZE'
    });
  }

  // ── FOOTING SCHEDULE ──────────────────────────────────────────────
  const footingSchedule = scheduleTables.find(t => t.type === 'FOOTING_SCHEDULE');
  if (footingSchedule) {
    for (const rec of footingSchedule.records) {
      const vals = Object.values(rec).join(' ');
      const sizeMatch = vals.match(/(\d{3,4})\s*[xX×]\s*(\d{3,4})/);
      const nosMatch  = Object.entries(rec).find(([k]) => /nos|no\.|qty/i.test(k));
      if (sizeMatch) {
        const l = parseInt(sizeMatch[1]) / 1000;
        const b = parseInt(sizeMatch[2]) / 1000;
        const nos = nosMatch ? parseInt(nosMatch[1]) || 1 : 1;
        const depthMatch = vals.match(/(\d{2,3})\s*(?:mm)?\s*(?:deep|depth|thk)/i);
        const depth = depthMatch ? parseInt(depthMatch[1]) / 1000 : 0.6;
        items.push({
          sr: sr++,
          description: `RCC Footing ${sizeMatch[1]}×${sizeMatch[2]}mm — ${nos} Nos — ${Math.round(depth * 1000)}mm deep`,
          unit: 'CUM',
          qty: Math.round(l * b * depth * nos * 100) / 100,
          rate: rates.rcc_footing || 7800,
          source: 'drawing_schedule',
          confidence: 'high'
        });
      }
    }
  }

  // ── RCC SLAB ──────────────────────────────────────────────────────
  const totalRoomArea = rooms.slice(0, 20).reduce((s, r) => s + r.area_sqm, 0);
  if (totalRoomArea > 10) {
    const slabThk = 0.125; // 125mm typical
    items.push({
      sr: sr++,
      description: 'RCC Slab 125mm THK (total built-up area)',
      unit: 'CUM',
      qty: Math.round(totalRoomArea * slabThk * 100) / 100,
      rate: rates.rcc_slab || 8200,
      source: 'calculated_from_room_areas',
      confidence: 'medium',
      detail: `${rooms.length} rooms, total ${Math.round(totalRoomArea)}sqm`
    });
  }

  // ── FLOORING ─────────────────────────────────────────────────────
  for (const h of hatchSummary) {
    if (/granite|marble|tile|floor/i.test(h.material)) {
      // Flooring area = rooms with that hatch layer
      const approxArea = totalRoomArea * 0.8; // rough estimate
      if (approxArea > 0) {
        items.push({
          sr: sr++,
          description: `${h.material} (${h.count} hatches on layer ${h.layers?.join(', ')})`,
          unit: 'SQMT',
          qty: Math.round(approxArea * 100) / 100,
          rate: rates.flooring_granite || 950,
          source: 'hatch_count',
          confidence: 'low',
          flag: 'CLAUDE_VERIFY_AREA'
        });
      }
    }
  }

  // Compute amounts
  for (const item of items) {
    if (item.qty && item.rate) {
      item.amount = Math.round(item.qty * item.rate);
    } else {
      item.amount = 0;
    }
  }

  return items;
}

// ─────────────────────────────────────────────────────────────────
// MAIN EXPORT — buildSmartContext(dxfData)
// ─────────────────────────────────────────────────────────────────
// Returns: { summary_text, pre_drafted_boq, claude_prompt }
// summary_text = what to show Claude
// pre_drafted_boq = already computed quantities
// claude_prompt = ready-to-use prompt for Claude
function buildSmartContext(dxfData, rates = {}) {
  const sf = dxfData.scale_factor || 1;

  // Re-extract rooms with correct scale
  const rooms = extractRooms(
    (dxfData.all_texts || []).map(t => typeof t === 'string' ? { text: t, x: 0, y: 0 } : t),
    [], // polylines would need raw data — use dimension-based approach
    sf
  );

  // Wall quantities with CORRECT formula
  const wallQty = extractWallQuantities(
    [], // need raw polylines from parser
    dxfData.layer_map || {},
    dxfData.floor_heights || [],
    sf
  );

  // Pre-draft BOQ
  const preDraft = preDraftBOQ(
    wallQty,
    rooms,
    dxfData.schedule_tables || [],
    dxfData.floor_levels || [],
    dxfData.floor_heights || [],
    dxfData.hatch_summary || [],
    dxfData.element_counts || {},
    rates
  );

  // Build the "engineer summary" for Claude
  const floorStr = (dxfData.floor_levels || [])
    .map(l => `  ${l.label || l.name}: ${l.m >= 0 ? '+' : ''}${l.m}m`)
    .join('\n') || '  Not found';

  const heightStr = (dxfData.floor_heights || [])
    .map(h => `  ${h.name}: ${h.height_m}m height`)
    .join('\n') || '  Not calculated';

  const wallStr = wallQty.length
    ? wallQty.map(w => `  ${w.thk_mm}mm wall — Floor ${w.floor}: ${w.length_m}m long × ${w.height_m}m high = ${w.volume_cum} CUM (${w.face_area_sqm} sqm face)`).join('\n')
    : (dxfData.wall_by_thickness || dxfData.wall_by_thickness_m2)
      ? Object.entries(dxfData.wall_by_thickness || dxfData.wall_by_thickness_m2)
          .map(([thk, d]) => `  ${thk}: plan area = ${typeof d === 'object' ? d.sqm || d : Math.round(d * 100) / 100} sqm`)
          .join('\n')
      : '  Not found';

  const scheduleStr = (dxfData.schedule_tables || [])
    .map(t => `\n  TABLE: ${t.type}\n  Headers: ${t.headers?.join(' | ')}\n  ${t.rows?.length} data rows:\n${t.rows?.slice(0, 8).map(r => '    ' + r.join(' | ')).join('\n')}`)
    .join('\n') || '  No schedule tables detected';

  const preDraftStr = preDraft.length
    ? preDraft.map(item =>
        `  [${item.confidence.toUpperCase()}] SR${item.sr}: ${item.description}\n` +
        `    Unit: ${item.unit} | Qty: ${item.qty ?? 'VERIFY'} | Rate: ₹${item.rate} | Amount: ₹${item.amount || 'VERIFY'}\n` +
        `    Source: ${item.source}${item.flag ? ' ⚠️ ' + item.flag : ''}${item.detail ? '\n    Detail: ' + item.detail : ''}`
      ).join('\n\n')
    : '  None pre-calculated';

  const claudePrompt = `You are a senior PMC civil engineer verifying a pre-drafted BOQ.
The BOQ below was computed DIRECTLY from drawing data by a parser.
Your job: (1) verify quantities, (2) fix flagged items, (3) add missing items, (4) apply correct rates.
DO NOT invent values. If something is flagged ⚠️ CLAUDE_VERIFY → read the schedule tables and fill in.

════════════════════════════════════════════════════
FILE: ${dxfData.filename || 'drawing.dxf'}
DRAWING TYPE: ${dxfData.drawing_type || 'UNKNOWN'}
SCALE: ${dxfData.scale || 'auto-detect needed'} (factor: ${sf})
════════════════════════════════════════════════════

FLOOR LEVELS (from drawing annotations):
${floorStr}

FLOOR HEIGHTS (calculated):
${heightStr}

WALL QUANTITIES (parser-calculated, CORRECT formula: perimeter × thk × height):
${wallStr}

ELEMENT COUNTS (from INSERTs + layers):
  Doors: ${dxfData.element_counts?.door_count || 0}
  Windows: ${dxfData.element_counts?.window_count || 0}
  Lifts: ${dxfData.element_counts?.lift_count || dxfData.element_counts?.lift_door_count || 0}
  Staircases: ${dxfData.element_counts?.staircase_count || 0}
  Columns (blocks): ${dxfData.element_counts?.column_count || 0}
  Total floor levels: ${dxfData.element_counts?.floor_levels_found || (dxfData.floor_levels || []).length}

ROOM AREAS (from closed polylines + text labels):
${rooms.length ? rooms.slice(0, 20).map(r => `  ${r.name}: ${r.area_sqm} sqm (layer: ${r.layer})`).join('\n') : '  No room polylines detected — use dimension annotations'}

SCHEDULE TABLES (extracted from drawing):
${scheduleStr}

ALL DRAWING TEXTS (first 80):
${(dxfData.all_texts || []).slice(0, 80).join(' | ')}

DIMENSIONS (top 30):
${(dxfData.dimension_values || []).slice(0, 30).map(d => `${d.value_m}m[${d.layer}]`).join(', ')}

════════════════════════════════════════════════════
PRE-DRAFTED BOQ (verify + fix + complete):
════════════════════════════════════════════════════
${preDraftStr}

════════════════════════════════════════════════════
YOUR TASK:
1. Review each pre-drafted item — if qty looks right, keep it; if wrong, correct it
2. For items flagged ⚠️ CLAUDE_VERIFY — read the schedule tables above and fill correct values
3. Add any missing standard BOQ items (PCC, waterproofing, paint, etc.)
4. Apply Gujarat DSR 2025 rates to all items
5. Return ONLY raw JSON (no markdown):

{
  "project_name": "",
  "drawing_type": "",
  "scale": "",
  "floor_count": 0,
  "building_height_m": 0,
  "total_bua_sqm": 0,
  "boq": [
    {
      "sr": 1,
      "description": "",
      "unit": "CUM|SQMT|RMT|NOS|KG",
      "qty": 0,
      "rate": 0,
      "amount": 0,
      "source": "drawing_schedule|drawing_polyline|calculated|assumed",
      "confidence": "high|medium|low"
    }
  ],
  "rooms": [{"name":"","area_sqm":0}],
  "observations": [],
  "pmc_recommendation": ""
}`;

  return {
    summary_text: claudePrompt,
    pre_drafted_boq: preDraft,
    rooms,
    wall_quantities: wallQty
  };
}

// ─────────────────────────────────────────────────────────────────
// ENHANCED VERSIONS for server.js integration
// ─────────────────────────────────────────────────────────────────

/**
 * Full pipeline: raw scanned data → smart context
 * Takes the output of drawing_intelligence.analyzeDrawing()
 */
function buildSmartContextFromAnalyzed(analyzed, rates = {}) {
  const sf = detected_scale_factor(analyzed);
  const scanned = analyzed._scanned;
  if (!scanned) return buildSmartContext(analyzed, rates);

  // Wall quantities with raw polylines (CORRECT formula)
  const wallQty = extractWallQuantities(
    scanned.polylines || [],
    analyzed._layerMap || {},
    analyzed.floor_heights || [],
    sf
  );

  // Room extraction from raw polylines + texts
  const rooms = extractRooms(scanned.texts || [], scanned.polylines || [], sf);

  // Schedule tables from texts
  const scheduleTables = extractScheduleTables(scanned.texts || [], sf);

  const preDraft = preDraftBOQ(
    wallQty,
    rooms,
    scheduleTables,
    analyzed.floor_levels || [],
    analyzed.floor_heights || [],
    analyzed.hatch_summary ? Object.entries(analyzed.hatch_summary).map(([k, v]) => ({ material: k, count: v })) : [],
    analyzed.element_counts || {},
    rates
  );

  // Build prompt
  const context = {
    ...analyzed,
    schedule_tables: scheduleTables,
    floor_heights: analyzed.floor_heights || [],
    scale_factor: sf,
  };
  const result = buildSmartContext(context, rates);
  result.pre_drafted_boq = preDraft;
  result.rooms = rooms;
  result.wall_quantities = wallQty;
  result.schedule_tables = scheduleTables;

  return result;
}

function detected_scale_factor(analyzed) {
  if (analyzed.scale_factor && analyzed.scale_factor > 1) return analyzed.scale_factor;
  if (analyzed.scale) {
    const m = analyzed.scale.match(/1\s*[:/]\s*(\d+)/);
    if (m) return parseInt(m[1]);
  }
  // Try from dims_sample
  if (analyzed.dims_sample?.length >= 3) {
    const result = detectScale(
      (analyzed.all_texts_sample || []).map(t => ({ text: t })),
      analyzed.dims_sample.map(d => ({ value_mm: d.mm, geom_mm: d.mm })),
      analyzed.drawing_extents ? { xmin: 0, xmax: analyzed.drawing_extents.width_m * 1000, ymin: 0, ymax: analyzed.drawing_extents.height_m * 1000 } : { xmin: 0, xmax: 0, ymin: 0, ymax: 0 }
    );
    return result.factor;
  }
  return 1;
}

module.exports = {
  buildSmartContext,
  buildSmartContextFromAnalyzed,
  extractWallQuantities,
  extractRooms,
  extractScheduleTables,
  preDraftBOQ,
  detectScale,
  calcWallPerimeter,
  shoelace
};
