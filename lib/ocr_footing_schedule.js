'use strict';
/**
 * Recover footing/column schedule rows from messy RapidOCR text
 * (industrial / PEB sheets like Bhagyeshree warehouse).
 */

function numsIn(block, min, max) {
  return [...String(block || '').matchAll(/\b(\d{3,4})\b/g)]
    .map(m => Number(m[1]))
    .filter(n => n >= min && n <= max);
}

/** Score a footing-schedule window — pick the one with real size/depth bands. */
function scoreFootingWindow(win) {
  if (!win) return -1;
  const hasLong = /long\s*span/i.test(win) ? 10 : 0;
  const hasSize = /footing\s*size/i.test(win) ? 6 : 0;
  const hasDepth = /footing\s*depth/i.test(win) ? 8 : 0;
  const longs = numsIn(win, 2000, 4500).length;
  const depthHit = win.match(/footing\s*depth[\s\S]{0,220}/i)?.[0] || '';
  const depths = numsIn(depthHit, 500, 1000).length;
  return hasLong + hasSize + hasDepth + longs * 2 + depths * 4;
}

/** Cut the best footing-schedule window; drop plan/overview noise. */
function footingWindow(text) {
  const raw = String(text || '');
  const re = /schedule\s*of\s*footing|scheduleoffooting/gi;
  const candidates = [];
  let m;
  while ((m = re.exec(raw))) {
    let win = raw.slice(m.index, m.index + 2400);
    win = win.split(
      /Typical\s+Reguler|Footing\s+reinforcement|SECTION\s*[-–]?\s*["']?A-A|CHKD\.|DES\.\s*BY|BUILDING IS DE/i
    )[0];
    candidates.push({ win, score: scoreFootingWindow(win), index: m.index });
  }
  if (!candidates.length) return '';
  candidates.sort((a, b) => b.score - a.score || b.index - a.index);
  return candidates[0].score > 0 ? candidates[0].win : candidates[0].win;
}

function extractOcrFootingSchedule(text) {
  const win = footingWindow(text);
  if (!win || scoreFootingWindow(win) < 8) return [];

  // Industrial OCR often prints "Footing size Footing depth" on ONE line —
  // so short-side numbers sit AFTER that label, not between size/depth indices.
  const labelIdx = win.search(/footing\s*size/i);
  const beforeSize = labelIdx >= 0 ? win.slice(0, labelIdx) : win;
  const afterLabel = labelIdx >= 0 ? win.slice(labelIdx) : win;

  // Depths: consecutive 500–1000 after label (skip 170/200 plinth noise)
  let depths = numsIn(afterLabel, 500, 1000);
  // Prefer the last decreasing-ish depth cluster (typical schedule column)
  if (depths.length > 8) depths = depths.slice(-8);
  depths = depths.slice(0, 8);

  // Long side: 2000–4500 before "Footing size" (long-span band)
  let longSide = numsIn(
    beforeSize.match(/long\s*span[\s\S]{0,450}/i)?.[0] || beforeSize.slice(-320),
    2000,
    4500
  ).filter(n => n !== 7650 && n !== 8450);

  // Short side: 1400–3999 after label, BEFORE depth cluster
  // Cut afterLabel at first depth value occurrence when we already know depths
  let shortBlock = afterLabel;
  if (depths[0]) {
    const cut = afterLabel.search(new RegExp(`\\b${depths[0]}\\b`));
    if (cut > 0) shortBlock = afterLabel.slice(0, cut);
  }
  let shortSide = numsIn(shortBlock, 1400, 3999);

  // 1800-class shorts often OCR'd just above "Footing size" with the long band
  const preShort = numsIn(
    (beforeSize.match(/long\s*span[\s\S]{0,450}/i)?.[0] || beforeSize.slice(-220)),
    1500,
    1999
  );
  if (preShort.length) shortSide = [...preShort, ...shortSide.filter(n => !preShort.includes(n))];

  // If still short on shorts, pull mid-range from afterLabel (exclude depths)
  if (shortSide.length < longSide.length) {
    const extra = numsIn(afterLabel, 1400, 3999).filter(n => !depths.includes(n));
    for (const n of extra) {
      if (!shortSide.includes(n)) shortSide.push(n);
      if (shortSide.length >= longSide.length) break;
    }
  }

  if (!longSide.length && !shortSide.length) return [];
  if (!depths.length) return [];

  const limit = Math.min(
    longSide.length,
    shortSide.length || longSide.length,
    depths.length,
    8
  );
  if (limit < 1) return [];

  const rows = [];
  for (let i = 0; i < limit; i++) {
    const L = longSide[i];
    const B = shortSide[i] || L;
    if (!L) continue;
    const depth = depths[i];
    if (!depth) continue;
    const paired = shortSide[i] != null && shortSide[i] !== L;
    const size = `${L}x${B}`;
    rows.push({
      mark: `F${i + 1}`,
      pcc_size_mm: size,
      rcc_size_mm: size,
      depth_mm: depth,
      main_bars_x: 'not found in drawing',
      main_bars_y: 'not found in drawing',
      qty: null,
      source: 'drawing-ocr-schedule',
      confidence: paired ? 'medium' : 'low',
      raw: `OCR footing schedule ${size} depth ${depth} — CONFIRM mark/qty on drawing`,
    });
  }
  return rows;
}

function extractOcrColumnSchedule(text) {
  const raw = String(text || '');
  const start = raw.search(/schedule\s*of\s*column|scheduleofcolumn/i);
  if (start < 0) return [];

  // Sizes often OCR'd just after a glued "SCHEDULEOFFOOTING" header — keep them.
  let win = raw.slice(start, start + 1000);
  win = win.split(/footing\s+size|long\s*span\s*steel|schedule\s+of\s+footing\s*:|Typical\s+Reguler/i)[0];

  const sizes = [...win.matchAll(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/g)]
    .map(m => ({ a: Number(m[1]), b: Number(m[2]), s: `${m[1]}x${m[2]}` }))
    .filter(x => x.a >= 200 && x.a <= 800 && x.b >= 200 && x.b <= 1500)
    .map(x => x.s);
  const uniq = [...new Set(sizes)].slice(0, 10);
  if (!uniq.length) return [];

  const marks = [];
  for (const m of raw.matchAll(/["']C\s*([1-8])(?:\s*&\s*C?\s*([1-8]))?["']\s*-?\s*COLUMN/gi)) {
    marks.push(`C${m[1]}`);
    if (m[2]) marks.push(`C${m[2]}`);
  }
  const uniqMarks = [...new Set(marks)];

  return uniq.map((size, i) => ({
    mark: uniqMarks[i] || `C${i + 1}`,
    size_mm: size,
    main_bars: 'not found in drawing',
    stirrups: 'not found in drawing',
    qty: null,
    source: 'drawing-ocr-schedule',
    confidence: 'medium',
    raw: `OCR column/pedestal ${size} — CONFIRM qty/height`,
  }));
}

function findPccHints(text) {
  const t = String(text || '');
  const mix = t.match(/P\.?C\.?C\s*\(?\s*1\s*:\s*3\s*:\s*6\s*\)?/i)?.[0] || '';
  const thk = t.match(/P\.?C\.?C[^.]{0,40}?(\d{2,3})\s*mm/i)
    || t.match(/(\d{2,3})\s*mm\s*(?:thk|thick)?[^.]{0,20}P\.?C\.?C/i);
  let n = thk ? Number(thk[1]) : null;
  if (n === 50) n = null; // grouting, not PCC bed
  return {
    mix: mix || (/P\.?C\.?C/i.test(t) ? 'P.C.C mentioned' : ''),
    thickness_mm: n && n >= 75 && n <= 200 ? n : null,
  };
}

function enrichSchedulesFromOcr(extracted, text) {
  const out = extracted || {
    schedules: { columns: [], footings: [], beams: [], doors: [], windows: [], base_plates: [], slabs: [], other: [] },
    meta: {},
    total_schedule_rows: 0,
    quality: 'poor',
  };
  const schedules = out.schedules || {};
  let added = 0;

  if (!(schedules.footings || []).length) {
    const ftg = extractOcrFootingSchedule(text);
    if (ftg.length) {
      schedules.footings = ftg;
      added += ftg.length;
    }
  }
  if (!(schedules.columns || []).length) {
    const cols = extractOcrColumnSchedule(text);
    if (cols.length) {
      schedules.columns = cols;
      added += cols.length;
    }
  }

  const pcc = findPccHints(text);
  out.meta = out.meta || {};
  if (pcc.mix) out.meta.pcc_note = pcc.mix;
  if (pcc.thickness_mm) out.meta.pcc_thickness_mm = pcc.thickness_mm;
  if (/M\s*-?\s*20/i.test(text)) out.meta.concrete_grade = out.meta.concrete_grade || 'M20';

  const dwg = String(text).match(/DWG\s*NO\s*[:\-]?\s*([A-Z0-9&\-]+)/i);
  if (dwg) out.meta.drawing_no = dwg[1];

  const totalRows = Object.values(schedules).reduce((n, a) => n + (a?.length || 0), 0);
  out.schedules = schedules;
  out.total_schedule_rows = totalRows;
  out.schedule_text_chars = (out.schedule_text_chars || 0) + added * 20;
  if (totalRows >= 1) out.quality = 'weak';
  out.ocr_enriched = added > 0;
  return out;
}

module.exports = {
  extractOcrFootingSchedule,
  extractOcrColumnSchedule,
  enrichSchedulesFromOcr,
  findPccHints,
  footingWindow,
};
