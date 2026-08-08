'use strict';
/**
 * Schedule-first extractor (Civils.ai-style)
 * Parses printed schedule tables from machine-extracted drawing text.
 * Never invents cell values — missing cells stay empty / not found.
 */

const SCHEDULE_HEADERS = [
  { type: 'columns',  re: /column\s*schedule|r\.?c\.?c\.?\s*column\s*schedule|pedestal\s*schedule|col\.?\s*sch/i },
  { type: 'footings', re: /footing\s*schedule|foundation\s*schedule|raft\s*schedule|ftg\.?\s*sch/i },
  { type: 'beams',    re: /beam\s*schedule|r\.?c\.?c\.?\s*beam\s*schedule/i },
  { type: 'doors',    re: /door\s*schedule|door\s*list|schedule\s*of\s*doors/i },
  { type: 'windows',  re: /window\s*schedule|window\s*list|schedule\s*of\s*windows/i },
  { type: 'base_plates', re: /base\s*plate\s*schedule|anchor\s*bolt\s*schedule/i },
  { type: 'slabs',    re: /slab\s*schedule|slab\s*panel\s*schedule/i },
];

const END_MARKERS = /^(notes?|general\s*notes?|title\s*block|revision|north|legend|typical\s*detail|section\s*[A-Z]|plan\s*at|key\s*plan|drawing\s*title)/i;

function normalizeLines(text) {
  if (!text || typeof text !== 'string') return [];
  return text
    .replace(/\r/g, '')
    .split('\n')
    .map(l => l.replace(/\u00a0/g, ' ').replace(/[|]+/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function splitCells(line) {
  // Prefer pipe-separated (GCV), else multi-space / tab, else single-space tokens
  if (line.includes('|')) {
    return line.split('|').map(c => c.trim()).filter(Boolean);
  }
  if (/\s{2,}|\t/.test(line)) {
    return line.split(/\s{2,}|\t+/).map(c => c.trim()).filter(Boolean);
  }
  return String(line).trim().split(/\s+/).filter(Boolean);
}

function parseSizeMm(raw) {
  if (!raw) return '';
  const s = String(raw).replace(/[×X]/g, 'x');
  // Do NOT strip all spaces — that glues "300x450 8" into "300x4508"
  const m = s.match(/(?:^|[^0-9])(\d{2,4})\s*[x]\s*(\d{2,4})(?:[^0-9]|$)/i);
  return m ? `${m[1]}x${m[2]}` : '';
}

function parseAllSizes(raw) {
  const s = String(raw || '').replace(/[×X]/g, 'x');
  const out = [];
  const re = /(\d{2,4})\s*[x]\s*(\d{2,4})/gi;
  let m;
  while ((m = re.exec(s))) {
    // Reject glued artifacts where digits run into the size (e.g. 1300 from C1+300)
    const start = m.index;
    const before = start > 0 ? s[start - 1] : '';
    if (before && /[0-9A-Za-z]/.test(before) && before.toLowerCase() !== 'x') {
      // allow only if previous char is whitespace or punctuation
      if (!/[\s,;:(/\-]/.test(before)) continue;
    }
    out.push(`${m[1]}x${m[2]}`);
  }
  // Fallback: token-level sizes
  if (!out.length) {
    for (const tok of splitCells(s)) {
      const one = parseSizeMm(` ${tok} `);
      if (one) out.push(one);
    }
  }
  return out;
}

function parseQty(raw) {
  if (raw == null || raw === '') return null;
  const m = String(raw).replace(/,/g, '').match(/(\d+(?:\.\d+)?)/);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : null;
}

function parseDepthMm(raw) {
  if (!raw) return null;
  const m = String(raw).replace(/,/g, '').match(/(?:^|[^0-9])(\d{3,4})(?:\s*mm)?(?:[^0-9]|$)/i);
  if (!m) return null;
  const n = Number(m[1]);
  return n >= 150 && n <= 5000 ? n : null;
}

function looksLikeHeaderRow(line) {
  return /mark|size|qty|nos|nos\.|main\s*bar|stirrup|depth|width|height|thickness|description|type|dia/i.test(line)
    && !/^\s*[A-Z]{1,3}\d{0,3}\b/i.test(line);
}

function parseBarSpec(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  // Common CAD forms: 8-12Ø, 4-16O, 10T16, 8Y12, 12#16, 8Ø @150
  const m = s.match(/(\d+)\s*[-–]\s*(\d+)\s*[ØOoφΦTYy#]?/i)
    || s.match(/(\d+)\s*[TYy#Øφ]\s*(\d+)/i)
    || s.match(/(\d+)\s*[ØOoφΦ]\s*@\s*\d+/i);
  if (m) return m[0].replace(/[Oo]/g, 'Ø').slice(0, 40);
  return '';
}

function trailingQty(tokens, line) {
  // Prefer explicit qty label, else last standalone 1–3 digit token (schedule QTY column)
  const labeled = String(line).match(/\b(?:qty|nos)\s*[:\-]?\s*(\d{1,3})\b/i);
  if (labeled) return Number(labeled[1]);
  for (let i = tokens.length - 1; i >= 0; i--) {
    const t = tokens[i];
    if (!/^\d{1,3}$/.test(t)) continue;
    // Skip stirrup spacing values that follow @ (e.g. "@ 150")
    if (i > 0 && /^@/.test(tokens[i - 1])) continue;
    if (/^@\d+$/.test(tokens[i])) continue;
    return Number(t);
  }
  return null;
}

function estimateSteelKgFromBars(barSpec, lengthM, qty) {
  // Approx: N-DØ → N bars of dia D mm; unit weight ~ d²/162 kg/m
  if (!barSpec || !lengthM || !qty) return 0;
  const m = String(barSpec).match(/(\d+)\s*[-–]\s*(\d+)/) || String(barSpec).match(/(\d+)\s*[TYy#Øφ]\s*(\d+)/i);
  if (!m) return 0;
  const nos = Number(m[1]);
  const dia = Number(m[2]);
  if (!nos || !dia) return 0;
  const kgPerM = (dia * dia) / 162;
  return Math.round(nos * kgPerM * lengthM * qty * 100) / 100;
}

function parseColumnRow(cells, line) {
  const tokens = cells.length ? cells : splitCells(line);
  const joined = tokens.join(' ');
  const mark = (tokens[0] || '').match(/^([A-Z]{1,3}\d{1,4}[A-Z]?)$/i)?.[1]
    || joined.match(/\b(C\d{1,3}[A-Z]?|P\d{1,3}[A-Z]?)\b/i)?.[1]
    || '';
  const sizes = parseAllSizes(joined);
  const size = sizes[0] || '';
  const qty = trailingQty(tokens, joined);
  const bars = [];
  for (const t of tokens) {
    const b = parseBarSpec(t);
    if (b) bars.push(b);
  }
  // Also catch "8-16Ø" spanning if split oddly
  const barFromLine = [...joined.matchAll(/(\d+)\s*[-–]\s*(\d+)\s*[ØOoφΦTYy#]?/gi)].map(m => m[0].replace(/[Oo]/g, 'Ø'));
  for (const b of barFromLine) if (!bars.includes(b)) bars.push(b);
  if (!mark && !size) return null;
  return {
    mark: mark || 'not found',
    size_mm: size || 'not found in drawing',
    main_bars: bars[0] || 'not found in drawing',
    stirrups: bars[1] || (joined.match(/\d+\s*[ØOoφΦ]\s*@\s*\d+/i)?.[0]?.replace(/[Oo]/g, 'Ø') || 'not found in drawing'),
    qty,
    source: size ? 'drawing-schedule' : 'not found',
    raw: line,
  };
}

function parseFootingRow(cells, line) {
  const tokens = cells.length ? cells : splitCells(line);
  const joined = tokens.join(' ');
  const mark = (tokens[0] || '').match(/^([A-Z]{1,3}\d{1,4}[A-Z]?)$/i)?.[1]
    || joined.match(/\b(F\d{1,3}[A-Z]?|FTG\d{0,3})\b/i)?.[1]
    || '';
  const sizes = parseAllSizes(joined);
  const qty = trailingQty(tokens, joined);
  // Depth: standalone 3–4 digit between sizes and bars (typically 300–900)
  let depth = null;
  for (const t of tokens) {
    if (/^\d{3,4}$/.test(t)) {
      const n = Number(t);
      if (n >= 200 && n <= 2000 && n !== qty) { depth = n; break; }
    }
  }
  if (depth == null) depth = parseDepthMm(joined.match(/depth\s*[:\-]?\s*(\d{3,4})/i)?.[0] || '');
  const bars = [...joined.matchAll(/(\d+)\s*[-–]\s*(\d+)\s*[ØOoφΦTYy#]?/gi)].map(m => m[0].replace(/[Oo]/g, 'Ø'));
  if (!mark && !sizes.length) return null;
  return {
    mark: mark || 'not found',
    pcc_size_mm: sizes[0] || 'not found in drawing',
    rcc_size_mm: sizes[1] || sizes[0] || 'not found in drawing',
    depth_mm: depth || null,
    main_bars_x: bars[0] || 'not found in drawing',
    main_bars_y: bars[1] || bars[0] || 'not found in drawing',
    qty,
    source: sizes.length ? 'drawing-schedule' : 'not found',
    raw: line,
  };
}

function parseOpeningRow(cells, line, kind) {
  const tokens = cells.length ? cells : splitCells(line);
  const joined = tokens.join(' ');
  const mark = (tokens[0] || '').match(/^([A-Z]{1,2}\d{1,3}[A-Z]?)$/i)?.[1]
    || joined.match(/\b([DW]\d{1,3}[A-Z]?)\b/i)?.[1]
    || '';
  const sizes = parseAllSizes(joined);
  const size = sizes[0] || '';
  const qty = trailingQty(tokens, joined);
  if (!mark && !size) return null;
  return {
    mark: mark || 'not found',
    type: kind,
    size_mm: size || 'not found in drawing',
    qty,
    source: size || mark ? 'drawing-schedule' : 'not found',
    raw: line,
  };
}

function parseGenericRow(cells, line) {
  if (cells.length < 2 && line.length < 8) return null;
  return { cells, raw: line, source: 'drawing-schedule' };
}

function extractMeta(lines) {
  const blob = lines.join('\n');
  const project = blob.match(/project\s*[:\-]\s*(.+)/i)?.[1]?.split('\n')[0]?.trim()
    || blob.match(/client\s*[:\-]\s*(.+)/i)?.[1]?.split('\n')[0]?.trim()
    || '';
  const drawingNo = blob.match(/drawing\s*no\.?\s*[:\-]?\s*([A-Z0-9\-\/_.]+)/i)?.[1] || '';
  const scale = blob.match(/scale\s*[:\-]?\s*(1\s*:\s*\d+)/i)?.[1]?.replace(/\s+/g, '') || '';
  const concrete = blob.match(/\b(M\s*2[05]|M\s*30|M\s*35)\b/i)?.[1]?.replace(/\s+/g, '').toUpperCase() || '';
  const steel = blob.match(/\b(Fe\s*500|Fe\s*415|Fe\s*550D?)\b/i)?.[1]?.replace(/\s+/g, '') || '';
  return {
    project_name: project.slice(0, 120),
    drawing_no: drawingNo,
    scale,
    concrete_grade: concrete,
    steel_grade: steel,
  };
}

/**
 * @param {string|string[]} textOrLines - PyMuPDF/GCV extracted text
 * @returns {{ schedules, meta, raw_text_chars, schedule_text_chars, quality }}
 */
function extractSchedules(textOrLines) {
  const lines = Array.isArray(textOrLines) ? textOrLines : normalizeLines(textOrLines);
  const schedules = {
    columns: [],
    footings: [],
    beams: [],
    doors: [],
    windows: [],
    base_plates: [],
    slabs: [],
    other: [],
  };

  let active = null;
  let scheduleTextChars = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    let switched = false;
    for (const h of SCHEDULE_HEADERS) {
      if (h.re.test(line)) {
        active = h.type;
        switched = true;
        scheduleTextChars += line.length;
        break;
      }
    }
    if (switched) continue;

    if (active && END_MARKERS.test(line) && !/schedule/i.test(line)) {
      active = null;
      continue;
    }

    // Also start a column schedule if we see a clear mark+size pattern without header
    if (!active && parseSizeMm(line) && /\b(C\d{1,3}|P\d{1,3}|COL)\b/i.test(line)) {
      active = 'columns';
    }
    if (!active && parseSizeMm(line) && /\b(F\d{1,3}|FTG|FOOT)/i.test(line)) {
      active = 'footings';
    }

    if (!active) continue;
    if (looksLikeHeaderRow(line)) {
      scheduleTextChars += line.length;
      continue;
    }

    const cells = splitCells(line);
    let row = null;
    if (active === 'columns' || active === 'base_plates') row = parseColumnRow(cells, line);
    else if (active === 'footings') row = parseFootingRow(cells, line);
    else if (active === 'doors') row = parseOpeningRow(cells, line, 'door');
    else if (active === 'windows') row = parseOpeningRow(cells, line, 'window');
    else row = parseGenericRow(cells, line);

    if (row) {
      schedules[active] = schedules[active] || [];
      schedules[active].push(row);
      scheduleTextChars += line.length;
    }
  }

  // Dedupe by mark+size
  for (const key of Object.keys(schedules)) {
    const seen = new Set();
    schedules[key] = (schedules[key] || []).filter(r => {
      const k = `${r.mark || ''}|${r.size_mm || r.rcc_size_mm || ''}|${r.qty ?? ''}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }

  const meta = extractMeta(lines);
  const totalRows = Object.values(schedules).reduce((n, a) => n + (a?.length || 0), 0);
  const usefulChars = scheduleTextChars;

  return {
    schedules,
    meta,
    raw_text_chars: lines.join('\n').length,
    schedule_text_chars: usefulChars,
    total_schedule_rows: totalRows,
    quality: totalRows >= 2 || usefulChars >= 200 ? 'good' : usefulChars >= 80 ? 'weak' : 'poor',
  };
}

/** Markdown tables for chat UI — schedules first */
function formatSchedulesMarkdown(extracted) {
  if (!extracted?.schedules) return '_No schedule tables detected in drawing text._';
  const parts = [];
  const { schedules, meta, quality, total_schedule_rows } = extracted;

  parts.push(`## Drawing schedules (machine-read)`);
  parts.push(`Quality: **${quality}** | Rows: **${total_schedule_rows || 0}**`);
  if (meta?.project_name) parts.push(`Project: ${meta.project_name}`);
  if (meta?.drawing_no) parts.push(`Drawing No: ${meta.drawing_no}`);
  if (meta?.scale) parts.push(`Scale: ${meta.scale}`);
  if (meta?.concrete_grade || meta?.steel_grade) {
    parts.push(`Grades: ${[meta.concrete_grade, meta.steel_grade].filter(Boolean).join(' / ')}`);
  }
  parts.push('');

  if (schedules.columns?.length) {
    parts.push('### Column / Pedestal Schedule');
    parts.push('| Mark | Size (mm) | Main bars | Stirrups | Qty | Source |');
    parts.push('|---|---|---|---|---:|---|');
    for (const r of schedules.columns) {
      parts.push(`| ${r.mark} | ${r.size_mm} | ${r.main_bars} | ${r.stirrups} | ${r.qty ?? '—'} | ${r.source} |`);
    }
    parts.push('');
  }

  if (schedules.footings?.length) {
    parts.push('### Footing Schedule');
    parts.push('| Mark | PCC size | RCC size | Depth mm | Bars X | Bars Y | Qty | Source |');
    parts.push('|---|---|---|---:|---|---|---:|---|');
    for (const r of schedules.footings) {
      parts.push(`| ${r.mark} | ${r.pcc_size_mm} | ${r.rcc_size_mm} | ${r.depth_mm ?? '—'} | ${r.main_bars_x} | ${r.main_bars_y} | ${r.qty ?? '—'} | ${r.source} |`);
    }
    parts.push('');
  }

  if (schedules.doors?.length) {
    parts.push('### Door Schedule');
    parts.push('| Mark | Size (mm) | Qty | Source |');
    parts.push('|---|---|---:|---|');
    for (const r of schedules.doors) {
      parts.push(`| ${r.mark} | ${r.size_mm} | ${r.qty ?? '—'} | ${r.source} |`);
    }
    parts.push('');
  }

  if (schedules.windows?.length) {
    parts.push('### Window Schedule');
    parts.push('| Mark | Size (mm) | Qty | Source |');
    parts.push('|---|---|---:|---|');
    for (const r of schedules.windows) {
      parts.push(`| ${r.mark} | ${r.size_mm} | ${r.qty ?? '—'} | ${r.source} |`);
    }
    parts.push('');
  }

  if (schedules.beams?.length) {
    parts.push(`### Beam Schedule (${schedules.beams.length} rows detected)`);
    parts.push('');
  }

  if (!total_schedule_rows) {
    parts.push('_No COLUMN/FOOTING/DOOR/WINDOW schedule header+rows found. Upload a clearer PDF/DXF or ensure schedule tables are text (not only scanned)._');
  }

  return parts.join('\n');
}

module.exports = {
  extractSchedules,
  formatSchedulesMarkdown,
  normalizeLines,
  parseSizeMm,
  parseQty,
  estimateSteelKgFromBars,
  SCHEDULE_HEADERS,
};
