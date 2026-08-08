'use strict';
/**
 * Plan measure takeoff (v1) — lengths / areas / counts from drawing text + optional DXF polys.
 * Not full vision markup like Civils.ai; uses printed dims, scale, and schedule counts.
 */

function parseScale(text) {
  const m = String(text || '').match(/scale\s*[:\-]?\s*1\s*[:\-]\s*(\d{1,4})/i)
    || String(text || '').match(/\b1\s*[:\-]\s*(\d{2,4})\b/);
  if (!m) return null;
  const den = Number(m[1]);
  if (!den || den < 1 || den > 5000) return null;
  return { ratio: `1:${den}`, denominator: den, mm_per_drawing_mm: den };
}

function extractPrintedLengthsMm(text) {
  const t = String(text || '');
  const out = [];
  // Explicit length callouts
  for (const m of t.matchAll(/(?:length|span|bay|c\/c|overall|o\/o)[^\n]{0,40}?(\d{3,5})\s*(?:mm)?/gi)) {
    const n = Number(m[1]);
    if (n >= 500 && n <= 200000) out.push({ value_mm: n, source: 'length', raw: m[0].slice(0, 60) });
  }
  // Large plan dims often printed alone (industrial bays)
  const big = [...t.matchAll(/\b(\d{4,5})\b/g)].map(m => Number(m[1])).filter(n => n >= 3000 && n <= 120000);
  const uniq = [...new Set(big)].slice(0, 12);
  for (const n of uniq) {
    if (!out.some(x => x.value_mm === n)) {
      out.push({ value_mm: n, kind: 'plan_dim', raw: String(n), confidence: 'low' });
    }
  }
  return out.slice(0, 20);
}

function extractPrintedAreas(text) {
  const t = String(text || '');
  const out = [];
  for (const m of t.matchAll(/(\d+(?:\.\d+)?)\s*(sq\.?\s*m|sqm|m\s*2|m²)/gi)) {
    out.push({ area_sqm: Number(m[1]), source: 'printed', raw: m[0] });
  }
  for (const m of t.matchAll(/(\d+(?:\.\d+)?)\s*(sq\.?\s*ft|sqft|sft)/gi)) {
    out.push({ area_sqm: Math.round(Number(m[1]) * 0.092903 * 1000) / 1000, source: 'printed-sqft', raw: m[0] });
  }
  // LxB in meters e.g. 12.5 x 8.0 m
  for (const m of t.matchAll(/(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*m\b/g)) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    if (a > 0.5 && b > 0.5 && a < 500 && b < 500) {
      out.push({ area_sqm: Math.round(a * b * 100) / 100, source: 'Lxm×Bm', raw: m[0] });
    }
  }
  return out.slice(0, 30);
}

function extractCounts(text, schedules) {
  const t = String(text || '');
  const counts = {
    footing_marks: [...new Set([...t.matchAll(/\bF\s*\d{1,3}\b/gi)].map(m => m[0].replace(/\s+/g, '').toUpperCase()))],
    column_marks: [...new Set([...t.matchAll(/\bC\s*\d{1,3}\b/gi)].map(m => m[0].replace(/\s+/g, '').toUpperCase()))],
    door_marks: [...new Set([...t.matchAll(/\bD\s*\d{1,3}\b/gi)].map(m => m[0].replace(/\s+/g, '').toUpperCase()))],
    window_marks: [...new Set([...t.matchAll(/\bW\s*\d{1,3}\b/gi)].map(m => m[0].replace(/\s+/g, '').toUpperCase()))],
  };
  const sch = schedules?.schedules || {};
  return {
    ...counts,
    schedule_footing_rows: (sch.footings || []).length,
    schedule_column_rows: (sch.columns || []).length,
    schedule_door_qty: (sch.doors || []).reduce((s, r) => s + (Number(r.qty) || 0), 0),
    schedule_window_qty: (sch.windows || []).reduce((s, r) => s + (Number(r.qty) || 0), 0),
  };
}

function measureFromDxfPolylines(polylines = [], scaleFactor = 0.001) {
  // scaleFactor: drawing unit → meters (default mm→m)
  const closed = [];
  const openLens = [];
  for (const p of polylines) {
    const pts = p.points || p.vertices || [];
    if (!pts.length) continue;
    let len = 0;
    for (let i = 1; i < pts.length; i++) {
      const dx = (pts[i].x - pts[i - 1].x) * scaleFactor;
      const dy = (pts[i].y - pts[i - 1].y) * scaleFactor;
      len += Math.hypot(dx, dy);
    }
    const isClosed = p.closed || p.is_closed;
    if (isClosed && pts.length >= 3) {
      // shoelace
      let area = 0;
      for (let i = 0; i < pts.length; i++) {
        const j = (i + 1) % pts.length;
        area += (pts[i].x * scaleFactor) * (pts[j].y * scaleFactor);
        area -= (pts[j].x * scaleFactor) * (pts[i].y * scaleFactor);
      }
      area = Math.abs(area) / 2;
      if (area > 0.5 && area < 50000) {
        closed.push({
          area_sqm: Math.round(area * 100) / 100,
          perimeter_m: Math.round(len * 100) / 100,
          layer: p.layer || '',
          source: 'dxf-polyline',
        });
      }
    } else if (len > 0.3 && len < 5000) {
      openLens.push({
        length_m: Math.round(len * 100) / 100,
        layer: p.layer || '',
        source: 'dxf-polyline',
      });
    }
  }
  closed.sort((a, b) => b.area_sqm - a.area_sqm);
  return {
    areas: closed.slice(0, 40),
    lengths: openLens.slice(0, 40),
    total_area_sqm: Math.round(closed.reduce((s, a) => s + a.area_sqm, 0) * 100) / 100,
    total_length_m: Math.round(openLens.reduce((s, a) => s + a.length_m, 0) * 100) / 100,
  };
}

/**
 * @param {object} opts
 * @param {string} opts.text
 * @param {object} [opts.schedules]
 * @param {Array} [opts.polylines] DXF polylines
 * @param {string} [opts.question]
 */
function buildPlanMeasure(opts = {}) {
  const text = opts.text || '';
  const scale = parseScale(text);
  const lengths = extractPrintedLengthsMm(text);
  const areas = extractPrintedAreas(text);
  const counts = extractCounts(text, opts.schedules);
  let dxf = null;
  if (opts.polylines?.length) {
    dxf = measureFromDxfPolylines(opts.polylines, opts.scaleFactor || 0.001);
  }

  const items = [];
  for (const a of areas) {
    items.push({
      type: 'area',
      qty: a.area_sqm,
      unit: 'sqm',
      description: `Area (${a.source})`,
      source: a.source,
      confidence: 'medium',
    });
  }
  if (dxf?.areas?.length) {
    for (const a of dxf.areas.slice(0, 15)) {
      items.push({
        type: 'area',
        qty: a.area_sqm,
        unit: 'sqm',
        description: `Closed polyline area${a.layer ? ' · ' + a.layer : ''}`,
        source: 'dxf-polyline',
        confidence: 'high',
      });
    }
  }
  for (const L of lengths.filter(x => x.kind === 'length').slice(0, 10)) {
    items.push({
      type: 'length',
      qty: Math.round((L.value_mm / 1000) * 1000) / 1000,
      unit: 'm',
      description: `Printed length ${L.value_mm} mm`,
      source: 'drawing-text',
      confidence: L.confidence || 'medium',
    });
  }
  if (dxf?.lengths?.length) {
    for (const L of dxf.lengths.slice(0, 15)) {
      items.push({
        type: 'length',
        qty: L.length_m,
        unit: 'm',
        description: `Polyline length${L.layer ? ' · ' + L.layer : ''}`,
        source: 'dxf-polyline',
        confidence: 'high',
      });
    }
  }
  // Counts
  if (counts.schedule_footing_rows) {
    items.push({
      type: 'count',
      qty: counts.schedule_footing_rows,
      unit: 'types',
      description: 'Footing types in schedule',
      source: 'schedule',
      confidence: 'high',
    });
  }
  if (counts.footing_marks.length) {
    items.push({
      type: 'count',
      qty: counts.footing_marks.length,
      unit: 'marks',
      description: `Footing marks on sheet (${counts.footing_marks.slice(0, 8).join(', ')})`,
      source: 'drawing-text',
      confidence: 'low',
    });
  }

  return {
    scale,
    counts,
    printed_lengths_mm: lengths,
    printed_areas: areas,
    dxf,
    items,
    quality: items.length >= 3 ? 'medium' : items.length ? 'weak' : 'poor',
    note: 'Plan measure v1: printed dims + DXF polylines. Full vision markup (click-measure) not yet enabled.',
  };
}

function formatPlanMeasureMarkdown(measure) {
  if (!measure) return '';
  const parts = [];
  parts.push('### Plan measure (areas / lengths / counts)');
  if (measure.scale) parts.push(`- Scale detected: **${measure.scale.ratio}**`);
  else parts.push('- Scale: **not found** — confirm if using drawing units for measure');
  parts.push(`- Quality: **${measure.quality}**`);
  parts.push('');
  if (measure.items?.length) {
    parts.push('| Type | Description | Qty | Unit | Source |');
    parts.push('|---|---|---:|---|---|');
    for (const i of measure.items) {
      parts.push(`| ${i.type} | ${i.description} | ${i.qty} | ${i.unit} | ${i.source} |`);
    }
  } else {
    parts.push('_No printable areas/lengths found. Upload DXF for polyline measure, or type dims._');
  }
  parts.push('');
  parts.push(`_${measure.note}_`);
  return parts.join('\n');
}

module.exports = {
  parseScale,
  buildPlanMeasure,
  formatPlanMeasureMarkdown,
  measureFromDxfPolylines,
};
