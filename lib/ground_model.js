'use strict';
/**
 * Simple 2D ground model from boreholes — IDW grid for GL / SPT (not Leapfrog).
 */

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function holeXY(b, i) {
  const e = num(b.easting);
  const n = num(b.northing);
  if (e != null && n != null) return { x: e, y: n, mode: 'geo' };
  return { x: (i % 5) * 50, y: Math.floor(i / 5) * 50, mode: 'schematic' };
}

function idw(points, x, y, power = 2) {
  let numW = 0;
  let den = 0;
  for (const p of points) {
    const d2 = (p.x - x) ** 2 + (p.y - y) ** 2;
    if (d2 < 1e-9) return p.v;
    const w = 1 / (d2 ** (power / 2));
    numW += w * p.v;
    den += w;
  }
  return den ? numW / den : null;
}

/**
 * @returns {{ mode, grid, cells, contours_hint, note }}
 */
function buildGroundModel(geotech = {}, opts = {}) {
  const holes = geotech.boreholes || [];
  const field = opts.field || 'avg_spt'; // or ground_level_m
  const pts = [];
  let mode = 'schematic';
  holes.forEach((b, i) => {
    const xy = holeXY(b, i);
    if (xy.mode === 'geo') mode = 'geo';
    const v = num(b[field]);
    if (v == null) return;
    pts.push({ x: xy.x, y: xy.y, v, mark: b.mark });
  });

  if (pts.length < 1) {
    return {
      mode: 'empty',
      grid: null,
      cells: [],
      holes: [],
      note: 'Need boreholes with SPT or GL for ground model.',
    };
  }

  const xs = pts.map(p => p.x);
  const ys = pts.map(p => p.y);
  const pad = mode === 'geo' ? Math.max(20, (Math.max(...xs) - Math.min(...xs)) * 0.15 || 50) : 25;
  const xmin = Math.min(...xs) - pad;
  const xmax = Math.max(...xs) + pad;
  const ymin = Math.min(...ys) - pad;
  const ymax = Math.max(...ys) + pad;
  const nx = opts.nx || 12;
  const ny = opts.ny || 12;
  const cells = [];
  for (let iy = 0; iy < ny; iy++) {
    for (let ix = 0; ix < nx; ix++) {
      const x = xmin + (ix + 0.5) * (xmax - xmin) / nx;
      const y = ymin + (iy + 0.5) * (ymax - ymin) / ny;
      const v = idw(pts, x, y);
      if (v == null) continue;
      cells.push({
        x, y,
        value: Math.round(v * 100) / 100,
        // leaflet circle radius hint
        r: Math.max(3, Math.min(18, Math.abs(v) / (field === 'avg_spt' ? 4 : 2))),
      });
    }
  }

  const vals = pts.map(p => p.v);
  return {
    mode,
    field,
    grid: { xmin, xmax, ymin, ymax, nx, ny },
    cells,
    holes: pts,
    stats: {
      min: Math.min(...vals),
      max: Math.max(...vals),
      avg: Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100) / 100,
    },
    note: 'Ground model v1 — IDW interpolation. Not a substitute for Leapfrog / OpenGround mesh.',
  };
}

function formatGroundModelMarkdown(gm) {
  if (!gm || gm.mode === 'empty') return '### Ground model\n_No data._';
  const parts = [
    '### Ground model (2D IDW)',
    `- Field: **${gm.field}** · mode: **${gm.mode}**`,
    `- Holes used: **${gm.holes.length}** · grid cells: **${gm.cells.length}**`,
  ];
  if (gm.stats) {
    parts.push(`- Range: ${gm.stats.min} … ${gm.stats.max} (avg ${gm.stats.avg})`);
  }
  parts.push('', `_${gm.note}_`);
  return parts.join('\n');
}

module.exports = { buildGroundModel, formatGroundModelMarkdown, idw };
