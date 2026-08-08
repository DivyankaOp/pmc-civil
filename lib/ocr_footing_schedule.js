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

/** Cut a clean footing-schedule window; drop plan/overview noise. */
function footingWindow(text) {
  const raw = String(text || '');
  const start = raw.search(/schedule\s*of\s*footing|scheduleoffooting/i);
  if (start < 0) return '';
  let win = raw.slice(start, start + 1800);
  // End before typical detail / reinforcement / title noise
  win = win.split(/Typical\s+Reguler|Footing\s+reinforcement|SECTION\s*[-–]?\s*["']?A|CHKD\.|DES\.\s*BY|BUILDING IS DE/i)[0];
  return win;
}

function extractOcrFootingSchedule(text) {
  const win = footingWindow(text);
  if (!win) return [];

  // Depths: ONLY after "Footing depth" / "Footing de"
  const depthIdx = win.search(/footing\s*de(?:pth)?/i);
  const depthBlock = depthIdx >= 0 ? win.slice(depthIdx, depthIdx + 180) : '';
  let depths = numsIn(depthBlock, 450, 1200).filter(d => d !== 450 || /depth/i.test(depthBlock));
  // Prefer classic footing depths
  depths = depths.filter(d => d >= 500 && d <= 1000).slice(0, 10);

  // Size bands
  const sizeIdx = win.search(/footing\s*size/i);
  const beforeSize = sizeIdx >= 0 ? win.slice(0, sizeIdx) : win;
  const afterSize = sizeIdx >= 0
    ? win.slice(sizeIdx, depthIdx > sizeIdx ? depthIdx : sizeIdx + 400)
    : '';

  // Long-span numbers just above / in schedule (exclude tiny dims)
  let longSide = numsIn(beforeSize.match(/long\s*span[\s\S]{0,300}/i)?.[0] || beforeSize.slice(-220), 1600, 4500);
  let shortSide = numsIn(afterSize, 1400, 3500);

  // Drop grid-like 7650/8450 if they leaked (bay spacing, not footing)
  longSide = longSide.filter(n => n <= 4500 && n !== 7650 && n !== 8450);
  shortSide = shortSide.filter(n => n <= 4000);

  // If short band empty, try numbers between size label and depth
  if (shortSide.length < 2 && depthIdx > 0) {
    shortSide = numsIn(win.slice(Math.max(0, sizeIdx), depthIdx), 1400, 3500);
  }

  // Pair long[i] x short[i]; if only one band, treat as square
  const n = Math.min(
    Math.max(longSide.length, 1),
    Math.max(shortSide.length || longSide.length, 1),
    depths.length || 8,
    8
  );
  if (!longSide.length && !shortSide.length) return [];

  const rows = [];
  // Prefer paired long×short with matching depth index (avoid leftover short-only squares)
  const limit = Math.min(longSide.length, shortSide.length || longSide.length, depths.length || 6, 6);
  for (let i = 0; i < limit; i++) {
    const L = longSide[i];
    const B = shortSide[i] || L;
    if (!L) continue;
    const depth = depths[i] || null;
    if (!depth) continue;
    const paired = !!(longSide[i] && shortSide[i]);
    const size = paired ? `${L}x${B}` : `${L}x${L}`;
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
  let win = raw.slice(start, start + 1200);
  win = win.split(/schedule\s*of\s*footing|SCHEDULE OF FOOTING|Typical\s+Reguler/i)[0];

  // Pedestal / column sizes in schedule (exclude huge base plates like 600x1305 sometimes pedestal)
  const sizes = [...win.matchAll(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/g)]
    .map(m => ({ a: Number(m[1]), b: Number(m[2]), s: `${m[1]}x${m[2]}` }))
    .filter(x => x.a >= 200 && x.a <= 800 && x.b >= 200 && x.b <= 1500)
    .map(x => x.s);
  const uniq = [...new Set(sizes)].slice(0, 10);

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
  // Ignore grouting 50mm — look for PCC thickness specifically
  const thk = t.match(/P\.?C\.?C[^.]{0,40}?(\d{2,3})\s*mm/i)
    || t.match(/(\d{2,3})\s*mm\s*(?:thk|thick)?[^.]{0,20}P\.?C\.?C/i);
  let n = thk ? Number(thk[1]) : null;
  if (n === 50) n = null; // almost always grouting in these sheets
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

  // Drawing no from clean pattern
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
