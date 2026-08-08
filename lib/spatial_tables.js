'use strict';
/**
 * Spatial schedule reconstruction for PDF / OCR boxes (Civils.ai-style grids).
 * PDF/image Y grows downward (unlike DXF). Cluster → pipe-separated lines for extractSchedules.
 */

const { reconstructScheduleTables, clusterTextsToTable } = require('./dxf_parser');

/** Cluster text spans where Y increases downward (PDF / raster). */
function clusterTextsPdf(texts, yTol = 12) {
  if (!texts?.length) return [];
  const valid = texts.filter(t => t.text && String(t.text).trim());
  if (!valid.length) return [];
  // Flip sort: top of page = smaller y
  const sorted = [...valid].sort((a, b) => a.y - b.y);
  const rows = [];
  for (const t of sorted) {
    const existing = rows.find(r => Math.abs(r.yCenter - t.y) <= yTol);
    if (existing) {
      existing.cells.push(t);
      existing.yCenter = existing.cells.reduce((s, c) => s + c.y, 0) / existing.cells.length;
    } else {
      rows.push({ yCenter: t.y, cells: [t] });
    }
  }
  for (const row of rows) row.cells.sort((a, b) => a.x - b.x);
  rows.sort((a, b) => a.yCenter - b.yCenter);
  return rows.map(r => r.cells);
}

/**
 * Reconstruct tables from PDF-oriented texts.
 * Converts to DXF-like coords (negate Y) so shared reconstructScheduleTables works.
 */
function reconstructScheduleTablesPdf(allTexts, yTol = 12) {
  if (!allTexts?.length) return [];
  const flipped = allTexts
    .filter(t => t?.text && String(t.text).trim())
    .map(t => ({ text: String(t.text).trim(), x: Number(t.x) || 0, y: -(Number(t.y) || 0) }));
  return reconstructScheduleTables(flipped, 1, yTol);
}

/** OCR RapidOCR box: [[x1,y1],[x2,y2],[x3,y3],[x4,y4]] + text */
function ocrBoxesToTexts(boxes) {
  const out = [];
  for (const b of boxes || []) {
    const poly = b.box || b.poly || b[0];
    const text = b.text || b[1] || '';
    if (!text || !poly) continue;
    let xs, ys;
    if (Array.isArray(poly[0])) {
      xs = poly.map(p => p[0]);
      ys = poly.map(p => p[1]);
    } else {
      continue;
    }
    const x = xs.reduce((a, v) => a + v, 0) / xs.length;
    const y = ys.reduce((a, v) => a + v, 0) / ys.length;
    out.push({ text: String(text).trim(), x, y });
  }
  return out;
}

function tablesToPipeLines(tables) {
  const lines = [];
  for (const t of tables || []) {
    const name = t.name || 'SCHEDULE';
    lines.push(String(name).toUpperCase().includes('SCHEDULE') ? name : `${name} SCHEDULE`);
    if (t.headers?.length) lines.push(t.headers.join(' | '));
    for (const row of t.rows || []) {
      if (Array.isArray(row)) lines.push(row.join(' | '));
      else if (typeof row === 'string') lines.push(row);
    }
    lines.push('');
  }
  return lines;
}

function clusteredRowsToPipeLines(rows) {
  return (rows || [])
    .filter(r => r.length >= 2)
    .map(r => r.map(c => (c.text != null ? c.text : c)).join(' | '));
}

/**
 * Build enriched text from PDF page texts (with x,y) + optional OCR boxes.
 * @returns {{ text, tables, pipeLines, plainLines }}
 */
function buildSpatialScheduleText({ pdfPages = [], ocrBoxes = [], plainLines = [] } = {}) {
  const spans = [];
  for (const page of pdfPages || []) {
    const pageOff = ((page.page || 1) - 1) * 10000;
    for (const t of page.texts || []) {
      spans.push({
        text: t.text,
        x: Number(t.x) || 0,
        y: (Number(t.y) || 0) + pageOff,
      });
    }
  }
  for (const t of ocrBoxesToTexts(ocrBoxes)) spans.push(t);

  const tables = reconstructScheduleTablesPdf(spans, 14);
  const pipeFromTables = tablesToPipeLines(tables);
  const clustered = clusterTextsPdf(spans, 14);
  const pipeFromRows = clusteredRowsToPipeLines(clustered);

  const parts = [
    plainLines.join('\n'),
    pipeFromTables.join('\n'),
    // Dense pipe rows help when header keywords missed
    pipeFromRows.slice(0, 400).join('\n'),
  ].filter(Boolean);

  return {
    text: parts.join('\n'),
    tables,
    pipeLines: pipeFromTables,
    clusteredRowCount: clustered.length,
  };
}

module.exports = {
  clusterTextsPdf,
  reconstructScheduleTablesPdf,
  ocrBoxesToTexts,
  tablesToPipeLines,
  clusteredRowsToPipeLines,
  buildSpatialScheduleText,
  clusterTextsToTable,
  reconstructScheduleTables,
};
