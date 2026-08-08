'use strict';
/**
 * Civils.ai-style DRAWING READING (primary) → takeoff calcs → BOQ (secondary).
 *
 * Real Civils.ai flow:
 *  1) Upload PDF sheets
 *  2) Read / measure what is on the drawing (schedules, dims, levels, notes)
 *  3) Calculate quantities with traceable formulas
 *  4) Human review missing items
 *  5) Final Excel / BOQ
 *
 * BOQ is never the only output — reading comes first.
 */

const { detectDrawingType, formatTypeMarkdown, capabilitiesForType } = require('./drawing_types');
const { extractSchedules, formatSchedulesMarkdown } = require('./schedule_extractor');
const { buildBoqFromSchedules, formatBoqMarkdown } = require('./boq_from_schedules');
const { buildClarifications, formatQuestionsMarkdown } = require('./clarifications');

function linesOf(text) {
  return String(text || '')
    .split(/\r?\n/)
    .map(l => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function findLevels(text) {
  const levels = [];
  const blob = String(text);
  const numeric = /(?:(?:PLINTH|FFL|SFL|TFL|RFL|T\.?O\.?S|B\.?O\.?S|N\.?G\.?L|G\.?L|EGL|FGL)\s*[:=]?\s*)?([+\-]?\d+(?:\.\d+)?)\s*(?:M|MT|m)\b|(?:LEVEL|LVL)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)/gi;
  let m;
  while ((m = numeric.exec(blob))) {
    levels.push(m[0].replace(/\s+/g, ' ').trim());
  }
  const named = blob.match(
    /(?:roof\s*top|terrace|first\s*floor|second\s*floor|upper\s*ground|lower\s*ground|ground\s*level|plinth|mezzanine(?:\s*-?\s*\d+)?|road\s*level|basement|parapet)(?:\s*level)?/gi
  ) || [];
  for (const n of named) levels.push(n.replace(/\s+/g, ' ').trim());
  return [...new Set(levels.map(x => x.replace(/\s+/g, ' ')))].slice(0, 50);
}

function findSizes(text) {
  const sizes = [];
  const re = /(\d{2,4})\s*[xX×]\s*(\d{2,4})(?:\s*[xX×]\s*(\d{2,4}))?/g;
  let m;
  while ((m = re.exec(String(text)))) {
    sizes.push(m[0].replace(/\s+/g, ''));
  }
  return [...new Set(sizes)].slice(0, 80);
}

function findDimensions(text) {
  const dims = [];
  const blob = String(text);
  // Printed dimension strings: 4500, 12.5m, 150mm c/c, Ø12 @150
  const pats = [
    /\b\d{3,5}\s*mm\b/gi,
    /\b\d+(?:\.\d+)?\s*m\b/gi,
    /\b\d{2,4}\s*[ØOoφΦ]\s*@\s*\d{2,4}/gi,
    /\bc\/c\s*\d{2,4}/gi,
    /\bcover\s*\d{2,3}\s*mm/gi,
  ];
  for (const re of pats) {
    let m;
    while ((m = re.exec(blob))) dims.push(m[0].replace(/\s+/g, ' ').trim());
  }
  return [...new Set(dims)].slice(0, 40);
}

function findSectionMarks(text) {
  const marks = String(text).match(/\b(?:SEC(?:TION)?\.?\s*[A-Z]\s*[-–]\s*[A-Z]|[A-Z]\s*[-–]\s*[A-Z]\s*\(?\s*SEC)/gi) || [];
  const simple = String(text).match(/\bsection\s+[A-Z](?:-[A-Z])?\b/gi) || [];
  return [...new Set([...marks, ...simple].map(s => s.replace(/\s+/g, ' ')))].slice(0, 20);
}

function findNotes(text) {
  const lines = linesOf(text);
  const notes = [];
  let inNotes = false;
  for (const line of lines) {
    if (/^general\s*notes?|^notes?\s*:|^note\s*:/i.test(line)) {
      inNotes = true;
      notes.push(line);
      continue;
    }
    if (inNotes) {
      if (/^(revision|title|drawing\s*no|schedule|legend)/i.test(line)) break;
      if (line.length > 8) notes.push(line);
      if (notes.length > 25) break;
    }
  }
  // Also grab key spec lines even outside NOTES block
  for (const line of lines) {
    if (/clear\s*cover|concrete\s*grade|steel\s*grade|fe\s*500|m\s*2[05]|m\s*30|pcc\s*\d+\s*mm|water\s*proof/i.test(line)) {
      notes.push(line);
    }
  }
  return [...new Set(notes)].slice(0, 30);
}

function findTitleBits(text) {
  const lines = linesOf(text);
  const project = lines.find(l => /project|proposed|warehouse|residential|building|client/i.test(l)) || '';
  const drawingNo = (String(text).match(/drawing\s*no\.?\s*[:\-]?\s*([A-Z0-9\-\/_.]+)/i) || [])[1] || '';
  const scale = (String(text).match(/scale\s*[:\-]?\s*(1\s*:\s*\d+)/i) || [])[1] || '';
  const title = lines.find(l => /layout|section|elevation|footing|plan|detail|schedule/i.test(l)) || '';
  const rev = (String(text).match(/rev(?:ision)?\s*[:\-]?\s*([A-Z0-9]+)/i) || [])[1] || '';
  return { project: project.slice(0, 140), drawingNo, scale, title: title.slice(0, 140), rev };
}

function inventoryFromSchedules(schedules) {
  const s = schedules?.schedules || {};
  return {
    columns: (s.columns || []).length,
    footings: (s.footings || []).length,
    beams: (s.beams || []).length,
    doors: (s.doors || []).length,
    windows: (s.windows || []).length,
    slabs: (s.slabs || []).length,
    base_plates: (s.base_plates || []).length,
    total_rows: schedules?.total_schedule_rows || 0,
    quality: schedules?.quality || 'poor',
  };
}

function readingChecklist(typeInfo, inv, title, levels, sizes) {
  const items = [];
  items.push({ ok: !!typeInfo?.drawing_type && typeInfo.drawing_type !== 'GENERAL_DRAWING', label: 'Drawing type identified' });
  items.push({ ok: !!title?.drawingNo || !!title?.project, label: 'Title block / drawing no.' });
  items.push({ ok: !!title?.scale, label: 'Scale found' });
  items.push({ ok: inv.total_rows > 0, label: 'Schedule / table rows read' });
  items.push({ ok: levels.length > 0 || ['FLOOR_PLAN', 'FOUNDATION_FOOTING', 'COLUMN_SCHEDULE'].includes(typeInfo?.drawing_type), label: 'Levels or plan/schedule context' });
  items.push({ ok: sizes.length > 0 || inv.total_rows > 0, label: 'Sizes / marks readable' });
  return items;
}

/**
 * Full Civils.ai-style reading + takeoff report.
 */
function readDrawingFully({
  text,
  filename,
  question = '',
  hints = [],
  boqOpts = {},
  extracted = null,
} = {}) {
  const typeInfo = detectDrawingType({ text, filename, hints });
  const schedules = extracted || extractSchedules(text);
  const title = findTitleBits(text);
  const levels = findLevels(text);
  const sizes = findSizes(text);
  const dims = findDimensions(text);
  const sections = findSectionMarks(text);
  const notes = findNotes(text);
  const inv = inventoryFromSchedules(schedules);
  const checklist = readingChecklist(typeInfo, inv, title, levels, sizes);

  const q = String(question || '').toLowerCase();
  const wantsBoq = !q.trim()
    || /boq|quantity|pcc|rcc|cum|estimate|rate|bill of|calculate|volume|takeoff|qty|read|study|padho|analyze/i.test(q);

  // Always compute takeoff from what was READ (missing → not_found, not invent)
  let boqResult = null;
  if (wantsBoq && (inv.total_rows >= 1
    || /FOUNDATION|COLUMN|INDUSTRIAL|FLOOR/i.test(typeInfo.drawing_type))) {
    boqResult = buildBoqFromSchedules(schedules, boqOpts);
    boqResult.drawing_type = typeInfo.drawing_type;
  }

  const clarifications = buildClarifications({
    text,
    schedules,
    typeInfo,
    title,
    question: question || 'read drawing fully and prepare takeoff / BOQ',
    boqNotFound: boqResult?.not_found || [],
  });

  const isFinal = clarifications.questions.length === 0
    && (boqResult?.boq?.length > 0)
    && inv.quality !== 'poor';
  // Schedules read but qty/PCC pending → DRAFT (not empty READING_ONLY)
  const status = isFinal ? 'FINAL'
    : (boqResult?.boq?.length || inv.total_rows >= 1) ? 'DRAFT'
      : 'READING_ONLY';

  const parts = [];
  parts.push('# Drawing Reading Report');
  parts.push(`**Status: ${status}** · Civils.ai-style: read sheet → calculate → BOQ (no invent)`);
  parts.push('');

  // ── 1. Sheet identity ──
  parts.push('## 1. Sheet identity (read from drawing)');
  parts.push(formatTypeMarkdown(typeInfo));
  parts.push(`- File: \`${filename || 'upload'}\``);
  if (title.project) parts.push(`- Project / title: ${title.project}`);
  if (title.title) parts.push(`- Sheet title hit: ${title.title}`);
  if (title.drawingNo) parts.push(`- Drawing No: **${title.drawingNo}**`);
  else parts.push('- Drawing No: **not found**');
  if (title.scale) parts.push(`- Scale: **${title.scale}**`);
  else parts.push('- Scale: **not found** (ask only if plan measuring needed)');
  if (title.rev) parts.push(`- Revision: ${title.rev}`);
  if (schedules.meta?.concrete_grade || schedules.meta?.steel_grade) {
    parts.push(`- Grades: ${[schedules.meta.concrete_grade, schedules.meta.steel_grade].filter(Boolean).join(' / ')}`);
  }
  parts.push(`- Can answer from this sheet: ${capabilitiesForType(typeInfo.drawing_type).join('; ')}`);

  // ── 2. Inventory ──
  parts.push('', '## 2. What was found on the drawing');
  parts.push('| Item | Count / status |');
  parts.push('|---|---|');
  parts.push(`| Schedule quality | **${inv.quality}** (${inv.total_rows} rows) |`);
  parts.push(`| Column / pedestal rows | ${inv.columns} |`);
  parts.push(`| Footing rows | ${inv.footings} |`);
  parts.push(`| Beam rows | ${inv.beams} |`);
  parts.push(`| Door / window rows | ${inv.doors} / ${inv.windows} |`);
  parts.push(`| Level / height hits | ${levels.length} |`);
  parts.push(`| Size marks (mm×mm) | ${sizes.length} |`);
  parts.push(`| Dimension / spacing hits | ${dims.length} |`);
  parts.push(`| Section marks | ${sections.length} |`);
  parts.push(`| Notes / spec lines | ${notes.length} |`);

  parts.push('', '### Reading checklist');
  for (const c of checklist) {
    parts.push(`- [${c.ok ? 'x' : ' '}] ${c.label}`);
  }

  // ── 3. How to read this sheet ──
  parts.push('', '## 3. How this sheet should be read');
  if (typeInfo.drawing_type === 'SECTION') {
    parts.push(
      '1. Identify section cut marks (A-A, B-B).',
      '2. Read **printed levels** only (PLINTH / FFL / T.O.S) — do not guess from image scale.',
      '3. Member sizes from schedule/detail callouts beat sketched proportions.',
    );
  } else if (typeInfo.drawing_type.includes('FOUNDATION') || inv.footings) {
    parts.push(
      '1. Read **SCHEDULE OF FOOTING / COLUMN** first (source of truth for size & qty).',
      '2. Plan/layout = location only — **do not count** symbols as qty unless schedule empty AND user confirms.',
      '3. PCC thickness / offset / excavation surcharge only if printed or user-confirmed.',
      '4. Volume = L × B × D × Qty (mm→m) — formula shown in takeoff section.',
    );
  } else if (typeInfo.drawing_type === 'COLUMN_SCHEDULE' || inv.columns) {
    parts.push(
      '1. Column schedule: Mark → Size → Bars → Stirrups → Qty.',
      '2. Height from section/schedule or user confirm — never silent 3 m.',
      '3. RCC = B × D × H × Qty.',
    );
  } else if (typeInfo.drawing_type === 'ELEVATION') {
    parts.push('1. Outside face only.', '2. Openings/levels from printed dims.', '3. Do not invent floor heights.');
  } else if (typeInfo.drawing_type === 'FLOOR_PLAN') {
    parts.push(
      '1. Rooms / walls from plan; door-window from schedules if present.',
      '2. Areas/lengths need scale confirmation for raster sheets.',
    );
  } else {
    parts.push(`1. Capabilities: ${capabilitiesForType(typeInfo.drawing_type).join('; ')}.`, '2. Prefer printed tables over graphics.');
  }

  // ── 4. Schedules ──
  parts.push('', '## 4. Schedules / tables (machine-read)');
  if (inv.total_rows >= 1) {
    parts.push(formatSchedulesMarkdown(schedules));
  } else {
    parts.push('_No schedule table rows detected yet._');
    parts.push('Upload clearer PDF/DXF, or answer clarification questions for marks/sizes.');
  }

  // ── 5. Levels / dims / notes ──
  parts.push('', '## 5. Levels, dimensions & notes (read)');
  if (levels.length) {
    parts.push('### Levels / heights');
    for (const l of levels) parts.push(`- ${l}`);
  } else {
    parts.push('### Levels / heights');
    parts.push('- **not found** on extract');
  }
  if (sections.length) {
    parts.push('', '### Section marks');
    parts.push(sections.map(s => `\`${s}\``).join(' · '));
  }
  if (sizes.length) {
    parts.push('', '### Size marks (mm)');
    parts.push(sizes.slice(0, 40).map(s => `\`${s}\``).join(' · '));
  }
  if (dims.length) {
    parts.push('', '### Other dimensions / spacing');
    parts.push(dims.slice(0, 30).map(s => `\`${s}\``).join(' · '));
  }
  if (notes.length) {
    parts.push('', '### Notes / specifications found');
    for (const n of notes.slice(0, 15)) parts.push(`- ${n}`);
  }

  // ── 6. Takeoff calculations ──
  parts.push('', `## 6. Quantity takeoff calculations (${status === 'FINAL' ? 'FINAL' : 'DRAFT — pending confirmations'})`);
  if (boqResult?.boq?.length) {
    parts.push('| Item | Formula / basis | Qty | Unit | Confidence |');
    parts.push('|---|---|---:|---|---|');
    for (const i of boqResult.boq) {
      parts.push(`| ${i.description} | ${i.calc_note || i.source} | ${i.qty} | ${i.unit} | ${i.confidence || '—'} |`);
    }
    const tq = boqResult.total_quantities || {};
    parts.push('');
    parts.push('### Totals from reading');
    if (tq.rcc_total_cum != null) parts.push(`- RCC total: **${tq.rcc_total_cum} cum**`);
    if (tq.pcc_total_cum != null) parts.push(`- PCC total: **${tq.pcc_total_cum} cum**`);
    if (tq.steel_total_kg != null) parts.push(`- Steel total: **${tq.steel_total_kg} kg**`);
    parts.push(`- Overall confidence: **${boqResult.overall_confidence || '—'}**`);
  } else {
    parts.push('_No calculable takeoff lines yet — schedule cells incomplete or sheet type has no qty tables._');
  }
  if (boqResult?.not_found?.length) {
    parts.push('', '### Blocked / not calculated (need confirmation)');
    for (const n of boqResult.not_found) parts.push(`- ${n}`);
  }

  // ── 7. BOQ ──
  parts.push('', `## 7. BOQ (${status === 'FINAL' ? '✅ FINAL' : '⚠️ DRAFT — not final until questions answered'})`);
  if (boqResult) {
    parts.push(formatBoqMarkdown(boqResult, { status }));
  } else {
    parts.push('_BOQ not generated — finish reading / provide schedule values first._');
  }

  // ── 8. Review / ask ──
  parts.push('', '## 8. Review (human-in-the-loop, like Civils.ai)');
  parts.push(formatQuestionsMarkdown(clarifications) || '_No blocking questions — reading complete enough for FINAL takeoff._');

  if (status === 'DRAFT') {
    parts.push('', '> **DRAFT** — Drawing read hua, lekin kuch values missing. Numbered answers do → FINAL BOQ + calcs rebuild.');
  } else if (status === 'FINAL') {
    parts.push('', '> **FINAL** — Reading + calculations schedule values se. Invented sizes/qty nahi.');
  } else {
    parts.push('', '> **READING_ONLY** — Sheet padha; takeoff ke liye clearer schedule/PDF ya aapke answers chahiye.');
  }

  const answeredLocally = String(text || '').trim().length >= 80;

  return {
    answeredLocally,
    needsClaude: String(text || '').trim().length < 120 && clarifications.questions.length === 0,
    needsUserInput: clarifications.questions.length > 0,
    clarifications,
    markdown: parts.join('\n'),
    status,
    meta: {
      typeInfo,
      schedules,
      title,
      levels,
      sizes,
      dimensions: dims,
      sections,
      notes,
      inventory: inv,
      checklist,
      size_count: sizes.length,
      text_chars: String(text || '').length,
      boqResult,
      status,
    },
  };
}

module.exports = {
  readDrawingFully,
  findLevels,
  findSizes,
  findDimensions,
  findSectionMarks,
  findNotes,
  findTitleBits,
};
