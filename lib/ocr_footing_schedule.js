'use strict';
/**
 * Format-tolerant footing/column schedule recovery from messy OCR/PDF text.
 *
 * Drawings never share one layout — run several strategies, score them,
 * keep the best. Prefer asking the user over inventing values.
 */

function numsIn(block, min, max) {
  return [...String(block || '').matchAll(/\b(\d{3,4})\b/g)]
    .map(m => Number(m[1]))
    .filter(n => n >= min && n <= max);
}

function uniqNums(arr) {
  const seen = new Set();
  const out = [];
  for (const n of arr || []) {
    if (seen.has(n)) continue;
    seen.add(n);
    out.push(n);
  }
  return out;
}

function footingRow(partial) {
  const size = partial.rcc_size_mm || partial.pcc_size_mm || '';
  return {
    mark: partial.mark || 'F?',
    pcc_size_mm: partial.pcc_size_mm || size || 'not found in drawing',
    rcc_size_mm: partial.rcc_size_mm || size || 'not found in drawing',
    depth_mm: partial.depth_mm ?? null,
    main_bars_x: partial.main_bars_x || 'not found in drawing',
    main_bars_y: partial.main_bars_y || 'not found in drawing',
    qty: partial.qty ?? null,
    source: partial.source || 'drawing-ocr-schedule',
    confidence: partial.confidence || 'low',
    strategy: partial.strategy || '',
    raw: partial.raw || `OCR ${size} d${partial.depth_mm || '?'} — CONFIRM on drawing`,
  };
}

function scoreRows(rows) {
  if (!rows?.length) return -1;
  let s = rows.length * 3;
  for (const r of rows) {
    if (r.rcc_size_mm && !/not found/i.test(r.rcc_size_mm)) s += 4;
    if (r.depth_mm) s += 3;
    if (r.qty != null) s += 2;
    if (r.confidence === 'high') s += 3;
    else if (r.confidence === 'medium') s += 1;
    if (/^F\d/i.test(r.mark)) s += 1;
    const m = String(r.rcc_size_mm || '').match(/(\d+)\s*x\s*(\d+)/i);
    if (m) {
      const a = Number(m[1]);
      const b = Number(m[2]);
      if (Math.min(a, b) >= 1200) s += 2; // plan footing-like
      if (Math.min(a, b) < 900) s -= 3;   // pedestal / pile noise
    }
  }
  return s;
}

/** Find all candidate schedule windows under many header wordings. */
function findScheduleWindows(text, kind) {
  const raw = String(text || '');
  const patterns = kind === 'footing'
    ? [
      /schedule\s*of\s*footing/gi,
      /scheduleoffooting/gi,
      /footing\s*schedule/gi,
      /foundation\s*schedule/gi,
      /ftg\.?\s*sch(?:edule)?/gi,
      /schedule\s*of\s*foundation/gi,
      /raft\s*schedule/gi,
    ]
    : [
      /schedule\s*of\s*column/gi,
      /scheduleofcolumn/gi,
      /column\s*schedule/gi,
      /pedestal\s*schedule/gi,
      /col\.?\s*sch(?:edule)?/gi,
      /r\.?c\.?c\.?\s*column\s*schedule/gi,
    ];

  const cuts = kind === 'footing'
    ? /Typical\s+Reguler|Footing\s+reinforcement|SECTION\s*[-–]?\s*["']?A-A|CHKD\.|DES\.\s*BY|BUILDING IS DE|general\s*notes/i
    : /footing\s+size|long\s*span\s*steel|schedule\s+of\s+footing\s*:|Typical\s+Reguler|general\s*notes/i;

  const seen = new Set();
  const wins = [];
  for (const re of patterns) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(raw))) {
      const key = m.index;
      if (seen.has(key)) continue;
      seen.add(key);
      let win = raw.slice(m.index, m.index + 2800);
      win = win.split(cuts)[0];
      wins.push({ win, index: m.index, header: m[0] });
    }
  }
  // Also: unlabeled band if "footing size" / "footing depth" appears
  if (kind === 'footing') {
    const soft = raw.search(/footing\s*size|footing\s*depth|size\s*of\s*footing/i);
    if (soft >= 0 && ![...seen].some(i => Math.abs(i - soft) < 80)) {
      const start = Math.max(0, soft - 400);
      wins.push({ win: raw.slice(start, start + 2200).split(cuts)[0], index: start, header: 'soft-footing-labels' });
    }
  }
  return wins;
}

function scoreFootingWindow(win) {
  if (!win) return -1;
  let s = 0;
  if (/footing\s*size|size\s*\(?mm\)?/i.test(win)) s += 6;
  if (/footing\s*depth|depth\s*\(?mm\)?/i.test(win)) s += 8;
  if (/long\s*span|short\s*span/i.test(win)) s += 4;
  if (/\bF\d{1,3}\b/i.test(win)) s += 5;
  if (/\d{3,4}\s*[xX×]\s*\d{3,4}/.test(win)) s += 8;
  s += Math.min(8, numsIn(win, 1200, 6000).length);
  s += Math.min(10, numsIn(win, 300, 1200).length * 2);
  return s;
}

// ── Strategy 1: classic one-line rows ─────────────────────────────
// F1 2600x1800 900 12 | F1 | 2600 x 1800 | 900 | 12
function strategyLineRows(text) {
  const rows = [];
  const re = /\b(F\d{1,3}[A-Z]?|FTG\s*\d{0,3}|FT\d{1,3})\b[^\n]{0,80}?(\d{3,4})\s*[xX×]\s*(\d{3,4})[^\n]{0,40}?(\d{3,4})?(?:[^\n]{0,20}?\b(\d{1,3})\b)?/gi;
  let m;
  while ((m = re.exec(text))) {
    const mark = m[1].replace(/\s+/g, '').toUpperCase().replace(/^FTG/, 'F').replace(/^FT/, 'F');
    const L = Number(m[2]);
    const B = Number(m[3]);
    let depth = m[4] != null ? Number(m[4]) : null;
    let qty = m[5] != null ? Number(m[5]) : null;
    // Heuristic: if "depth" looks like qty (<200) swap
    if (depth != null && depth < 200 && qty == null) { qty = depth; depth = null; }
    if (depth != null && (depth < 200 || depth > 2500)) depth = null;
    if (L < 600 || B < 600) continue;
    rows.push(footingRow({
      mark: mark.startsWith('F') ? mark : `F${mark}`,
      pcc_size_mm: `${L}x${B}`,
      rcc_size_mm: `${L}x${B}`,
      depth_mm: depth,
      qty,
      confidence: depth ? 'high' : 'medium',
      strategy: 'line-rows',
      raw: m[0].slice(0, 120),
    }));
  }
  // Dedup by mark
  const byMark = new Map();
  for (const r of rows) {
    if (!byMark.has(r.mark) || (r.depth_mm && !byMark.get(r.mark).depth_mm)) byMark.set(r.mark, r);
  }
  return { rows: [...byMark.values()], strategy: 'line-rows' };
}

// ── Strategy 2: mark near size (multi-line OCR) ───────────────────
function strategyMarkNearSize(text) {
  const rows = [];
  const marks = [...String(text).matchAll(/\b(F\s*\d{1,3}[A-Z]?)\b/gi)].map(m => ({
    mark: m[1].replace(/\s+/g, '').toUpperCase(),
    index: m.index,
  }));
  for (const { mark, index } of marks) {
    const win = text.slice(index, index + 180);
    const size = win.match(/(\d{3,4})\s*[xX×]\s*(\d{3,4})/);
    if (!size) continue;
    const L = Number(size[1]);
    const B = Number(size[2]);
    if (L < 600 || B < 600 || L > 8000 || B > 8000) continue;
    const after = win.slice(size.index + size[0].length);
    const depthHit = after.match(/\b(\d{3,4})\b/);
    let depth = depthHit ? Number(depthHit[1]) : null;
    if (depth != null && (depth < 250 || depth > 2000)) depth = null;
    const qtyHit = after.match(/\b(\d{1,3})\b/g);
    let qty = null;
    if (qtyHit) {
      for (const q of qtyHit) {
        const n = Number(q);
        if (n >= 1 && n <= 200 && n !== depth) { qty = n; break; }
      }
    }
    rows.push(footingRow({
      mark,
      pcc_size_mm: `${L}x${B}`,
      rcc_size_mm: `${L}x${B}`,
      depth_mm: depth,
      qty,
      confidence: depth ? 'medium' : 'low',
      strategy: 'mark-near-size',
    }));
  }
  const byMark = new Map();
  for (const r of rows) if (!byMark.has(r.mark)) byMark.set(r.mark, r);
  return { rows: [...byMark.values()], strategy: 'mark-near-size' };
}

/** Pick best size-label index — skip junk early "Footing size" without plan numbers. */
function bestSizeLabelIndex(win) {
  const re = /footing\s*size|size\s*of\s*footing/gi;
  let best = -1;
  let bestScore = -1;
  let m;
  while ((m = re.exec(win))) {
    const after = win.slice(m.index, m.index + 400);
    const before = win.slice(Math.max(0, m.index - 350), m.index);
    const planAfter = numsIn(after, 900, 6000).length;
    const planBefore = numsIn(before, 900, 6000).length;
    const depths = numsIn(after, 400, 1500).length;
    const longSpan = /long\s*span/i.test(before) ? 8 : 0;
    const sc = planAfter * 2 + planBefore * 2 + depths * 3 + longSpan;
    if (sc > bestScore) {
      bestScore = sc;
      best = m.index;
    }
  }
  if (best < 0) {
    const soft = win.search(/\bsize\s*\(?mm\)?/i);
    return soft;
  }
  return best;
}

// ── Strategy 3: band / column OCR (sizes & depths on separate lines) ─
function strategyBandColumns(text) {
  const windows = findScheduleWindows(text, 'footing');
  if (!windows.length) return { rows: [], strategy: 'band-columns' };
  // Prefer windows that already look like a real schedule (long-span / plan sizes)
  windows.sort((a, b) => {
    const sa = scoreFootingWindow(a.win) + (/long\s*span/i.test(a.win) ? 12 : 0);
    const sb = scoreFootingWindow(b.win) + (/long\s*span/i.test(b.win) ? 12 : 0);
    return sb - sa;
  });

  let bestRows = [];
  for (const { win } of windows.slice(0, 3)) {
    if (scoreFootingWindow(win) < 6) continue;
    const rows = bandRowsFromWindow(win);
    if (scoreRows(rows) > scoreRows(bestRows)) bestRows = rows;
  }
  return { rows: bestRows, strategy: 'band-columns' };
}

function bandRowsFromWindow(win) {
  const labelIdx = bestSizeLabelIndex(win);
  const beforeSize = labelIdx >= 0 ? win.slice(0, labelIdx) : win;
  const afterLabel = labelIdx >= 0 ? win.slice(labelIdx) : win;

  // Depths: after depth label, or trailing 400–1500 cluster
  let depthBlock = afterLabel;
  const dLab = afterLabel.search(/footing\s*depth|depth\s*\(?mm\)?/i);
  if (dLab >= 0) depthBlock = afterLabel.slice(dLab, dLab + 280);
  let depths = numsIn(depthBlock, 400, 1500);
  if (depths.some(d => d >= 500)) depths = depths.filter(d => d >= 450);
  depths = depths.slice(0, 10);

  // Prefer size band near LONG/SHORT SPAN — avoids pile/section dims (1008, 1300…)
  const spanChunk = beforeSize.match(/long\s*span[\s\S]{0,450}/i)?.[0]
    || beforeSize.match(/(?:footing|plan)[\s\S]{0,300}$/i)?.[0]
    || beforeSize.slice(-200);
  // Drop noise near PILE / SECTION markers
  const cleanedBefore = beforeSize
    .replace(/pile[\s\S]{0,120}/gi, ' ')
    .replace(/section\s*['"]?\d[\s\S]{0,80}/gi, ' ');

  let pre = numsIn(spanChunk, 1500, 6000).filter(n => n < 7000);
  if (pre.length < 2) pre = numsIn(cleanedBefore.slice(-280), 1200, 6000).filter(n => n < 7000);

  let postBlock = afterLabel;
  if (depths[0]) {
    const cut = afterLabel.search(new RegExp(`\\b${depths[0]}\\b`));
    if (cut > 0) postBlock = afterLabel.slice(0, cut);
  }
  let post = numsIn(postBlock, 1200, 6000).filter(n => n < 7000 && !depths.includes(n));

  const avg = arr => (arr.length ? arr.reduce((s, n) => s + n, 0) / arr.length : 0);
  let longSide = [];
  let shortSide = [];
  if (pre.length >= 2 && post.length >= 1) {
    if (avg(pre) >= avg(post)) {
      longSide = pre;
      shortSide = post;
    } else {
      longSide = post;
      shortSide = pre;
    }
  } else {
    const all = uniqNums([...pre, ...post]);
    const mid = [...all].sort((a, b) => a - b)[Math.floor(all.length / 2)] || 2000;
    longSide = all.filter(n => n >= mid);
    shortSide = all.filter(n => n < mid);
  }

  // 1800-class values sitting with the long band → short side
  const moved = [];
  while (longSide.length > shortSide.length && longSide[longSide.length - 1] < 2000) {
    moved.unshift(longSide.pop());
  }
  if (moved.length) shortSide = uniqNums([...moved, ...shortSide]);

  if (!longSide.length || !depths.length) return [];

  const limit = Math.min(longSide.length, shortSide.length || longSide.length, depths.length, 10);
  const rows = [];
  for (let i = 0; i < limit; i++) {
    const L = longSide[i];
    const B = shortSide[i] || L;
    if (!L || !depths[i]) continue;
    // Reject absurd pairings (pile-ish × plan)
    if (Math.min(L, B) < 1000 && Math.max(L, B) > 2500) continue;
    const paired = shortSide[i] != null && shortSide[i] !== L;
    rows.push(footingRow({
      mark: `F${rows.length + 1}`,
      pcc_size_mm: `${L}x${B}`,
      rcc_size_mm: `${L}x${B}`,
      depth_mm: depths[i],
      confidence: paired ? 'medium' : 'low',
      strategy: 'band-columns',
    }));
  }
  return rows;
}

// ── Strategy 4: loose L×B near footing keywords + nearby depth ─────
function strategyLooseSizes(text) {
  const wins = findScheduleWindows(text, 'footing');
  const blob = wins.length
    ? wins.map(w => w.win).join('\n')
    : (text.match(/(?:footing|ftg|foundation)[\s\S]{0,2500}/i)?.[0] || '');
  if (!blob || blob.length < 40) return { rows: [], strategy: 'loose-sizes' };

  const sizes = [...blob.matchAll(/(\d{3,4})\s*[xX×]\s*(\d{3,4})/g)]
    .map(m => ({ L: Number(m[1]), B: Number(m[2]), index: m.index }))
    .filter(s => s.L >= 600 && s.B >= 600 && s.L <= 8000 && s.B <= 8000
      && !(s.L >= 200 && s.L <= 800 && s.B > 1000)); // skip pedestal-ish if huge B only

  // Prefer plan-footing sizes (both sides >= 900)
  const footingLike = sizes.filter(s => s.L >= 900 && s.B >= 900);
  const use = footingLike.length ? footingLike : sizes;
  const depths = numsIn(blob, 300, 1500).filter(d => d >= 300);
  const rows = [];
  for (let i = 0; i < Math.min(use.length, 10); i++) {
    const s = use[i];
    // depth: nearest number after size in blob
    const after = blob.slice(s.index, s.index + 80);
    let depth = numsIn(after, 300, 1500)[0] || depths[i] || null;
    rows.push(footingRow({
      mark: `F${i + 1}`,
      pcc_size_mm: `${s.L}x${s.B}`,
      rcc_size_mm: `${s.L}x${s.B}`,
      depth_mm: depth,
      confidence: depth ? 'medium' : 'low',
      strategy: 'loose-sizes',
    }));
  }
  return { rows, strategy: 'loose-sizes' };
}

function extractOcrFootingSchedule(text) {
  const strategies = [
    strategyLineRows,
    strategyMarkNearSize,
    strategyBandColumns,
    strategyLooseSizes,
  ];
  let best = { rows: [], score: -1, strategy: 'none' };
  for (const fn of strategies) {
    try {
      const r = fn(text);
      const score = scoreRows(r.rows);
      if (score > best.score) best = { rows: r.rows, score, strategy: r.strategy };
    } catch (_) { /* try next */ }
  }
  // Annotate winning strategy on rows
  return (best.rows || []).map(r => ({ ...r, strategy: best.strategy, raw: `${r.raw} [${best.strategy}]` }));
}

function extractOcrColumnSchedule(text) {
  const windows = findScheduleWindows(text, 'column');
  let win = '';
  if (windows.length) {
    // Prefer window with most LxB pedestal/column sizes
    windows.sort((a, b) => {
      const ca = [...a.win.matchAll(/\d{2,4}\s*[xX×]\s*\d{2,4}/g)].length;
      const cb = [...b.win.matchAll(/\d{2,4}\s*[xX×]\s*\d{2,4}/g)].length;
      return cb - ca;
    });
    win = windows[0].win;
  } else {
    const start = String(text).search(/column|pedestal|C\s*[1-9]/i);
    if (start < 0) return [];
    win = text.slice(start, start + 1200);
  }

  // Keep sizes even if a glued FOOTING header appears early
  const sizes = [...win.matchAll(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/g)]
    .map(m => ({ a: Number(m[1]), b: Number(m[2]), s: `${m[1]}x${m[2]}` }))
    .filter(x => x.a >= 200 && x.a <= 900 && x.b >= 200 && x.b <= 1600)
    .map(x => x.s);
  const uniq = [...new Set(sizes)].slice(0, 12);
  if (!uniq.length) return [];

  const marks = [];
  for (const m of String(text).matchAll(/\bC\s*([1-9]\d?)(?:\s*[&,]\s*C?\s*([1-9]\d?))?\b/gi)) {
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
    strategy: 'column-multi-header',
    raw: `OCR column/pedestal ${size} — CONFIRM qty/height`,
  }));
}

function findPccHints(text) {
  const t = String(text || '');
  const mix = t.match(/P\.?C\.?C\s*\(?\s*1\s*:\s*3\s*:\s*6\s*\)?/i)?.[0]
    || t.match(/P\.?C\.?C\s*\(?\s*1\s*:\s*4\s*:\s*8\s*\)?/i)?.[0]
    || '';
  const thk = t.match(/P\.?C\.?C[^.\n]{0,50}?(\d{2,3})\s*mm/i)
    || t.match(/(\d{2,3})\s*mm\s*(?:thk|thick\.?)[^.\n]{0,30}P\.?C\.?C/i)
    || t.match(/plain\s*cement\s*concrete[^.\n]{0,40}?(\d{2,3})\s*mm/i);
  let n = thk ? Number(thk[1]) : null;
  if (n === 50) n = null;
  return {
    mix: mix || (/P\.?C\.?C|plain\s*cement\s*concrete/i.test(t) ? 'P.C.C mentioned' : ''),
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
  let strategy = '';

  if (!(schedules.footings || []).length) {
    const ftg = extractOcrFootingSchedule(text);
    if (ftg.length) {
      schedules.footings = ftg;
      strategy = ftg[0]?.strategy || '';
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
  if (/M\s*-?\s*2[05]/i.test(text)) {
    const g = text.match(/M\s*-?\s*(20|25|30|35|40)/i);
    if (g) out.meta.concrete_grade = out.meta.concrete_grade || `M${g[1]}`;
  }

  const dwg = String(text).match(/DWG\s*(?:NO|NO\.|NUMBER)?\s*[:\-]?\s*([A-Z0-9&\-\/]+)/i);
  if (dwg) out.meta.drawing_no = dwg[1];
  if (strategy) out.meta.ocr_strategy = strategy;

  const totalRows = Object.values(schedules).reduce((n, a) => n + (a?.length || 0), 0);
  out.schedules = schedules;
  out.total_schedule_rows = totalRows;
  out.schedule_text_chars = (out.schedule_text_chars || 0) + added * 20;
  if (totalRows >= 1) out.quality = out.quality === 'good' ? 'good' : 'weak';
  out.ocr_enriched = added > 0;
  return out;
}

/** Back-compat helper */
function footingWindow(text) {
  const wins = findScheduleWindows(text, 'footing');
  if (!wins.length) return '';
  wins.sort((a, b) => scoreFootingWindow(b.win) - scoreFootingWindow(a.win));
  return wins[0].win;
}

module.exports = {
  extractOcrFootingSchedule,
  extractOcrColumnSchedule,
  enrichSchedulesFromOcr,
  findPccHints,
  footingWindow,
  findScheduleWindows,
  scoreRows,
};
