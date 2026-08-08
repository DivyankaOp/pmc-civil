'use strict';
/**
 * Civils.ai-style DRAWING READING (default) → answer only what user asked.
 * BOQ only when user explicitly asks for BOQ / full takeoff.
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
  return [
    { ok: !!typeInfo?.drawing_type && typeInfo.drawing_type !== 'GENERAL_DRAWING', label: 'Drawing type identified' },
    { ok: !!title?.drawingNo || !!title?.project, label: 'Title block / drawing no.' },
    { ok: !!title?.scale, label: 'Scale found' },
    { ok: inv.total_rows > 0, label: 'Schedule / table rows read' },
    { ok: levels.length > 0 || ['FLOOR_PLAN', 'FOUNDATION_FOOTING', 'COLUMN_SCHEDULE'].includes(typeInfo?.drawing_type), label: 'Levels or plan/schedule context' },
    { ok: sizes.length > 0 || inv.total_rows > 0, label: 'Sizes / marks readable' },
  ];
}

/**
 * Intent (product default = takeoff, like professional QS tools):
 * - takeoff (default on upload): schedules + draft quantities + ask gaps
 * - footing_calc: PCC/RCC footing only
 * - schedule / levels / read / boq / qa
 */
function detectIntent(question) {
  const t = String(question || '').toLowerCase().trim();
  // Explicit read-only
  if (/read\s*(this\s*)?drawing\s*only|do not (prepare )?boq|without\s*boq|padho\s*only|study\s*only/i.test(t)) {
    return 'read';
  }
  if (/\bboq\b|bill of quant|complete\s*boq|prepare\s*boq|full\s*boq/i.test(t)) return 'boq';
  if (/pcc|rcc|footing.*(calc|volume|cum|qty)|calculate.*(pcc|rcc|footing)|footing.*(pcc|rcc)/i.test(t)) {
    return 'footing_calc';
  }
  if (/^levels?$|level|plinth|ffl|height.*floor|storey/i.test(t) && !/boq|pcc|rcc|takeoff/i.test(t)) return 'levels';
  if (/schedule|footing size|column size|table/i.test(t) && !/boq|calculate|volume|takeoff|measure/i.test(t)) {
    return 'schedule';
  }
  // Default upload / takeoff / measure → draft quantity takeoff
  if (!t
    || /takeoff|take-off|measure|quantit|material\s*list|auto\s*excel|prepare\s*takeoff|upload/i.test(t)) {
    return 'takeoff';
  }
  if (/read this|drawing read|padho|study|samjhao|analyze|analyse|sheet type|title block/i.test(t)) {
    return 'read';
  }
  return 'qa';
}

function relevantLines(text, question, limit = 20) {
  const q = String(question || '').toLowerCase();
  const keys = q.split(/[^a-z0-9]+/i).filter(w => w.length > 2
    && !['the', 'and', 'for', 'from', 'what', 'how', 'please', 'calculate', 'drawing'].includes(w));
  return linesOf(text).map(line => {
    const low = line.toLowerCase();
    let s = 0;
    for (const k of keys) if (low.includes(k)) s += 1;
    if (/schedule|section|level|footing|column|pcc|rcc/i.test(line)) s += 0.3;
    return { line, s };
  }).filter(x => x.s > 0).sort((a, b) => b.s - a.s).slice(0, limit).map(x => x.line);
}

function filterBoqForIntent(boqResult, intent) {
  if (!boqResult?.boq?.length) return boqResult;
  if (intent !== 'footing_calc') return boqResult;
  const filtered = boqResult.boq.filter(i => /footing|pcc under footing/i.test(i.description));
  const rcc = filtered.filter(i => /RCC Footing/i.test(i.description)).reduce((s, i) => s + i.qty, 0);
  const pcc = filtered.filter(i => /PCC under footing/i.test(i.description)).reduce((s, i) => s + i.qty, 0);
  return {
    ...boqResult,
    boq: filtered,
    total_quantities: {
      ...(boqResult.total_quantities || {}),
      rcc_total_cum: Math.round(rcc * 100) / 100,
      footing_rcc_cum: Math.round(rcc * 100) / 100,
      pcc_total_cum: Math.round(pcc * 100) / 100,
      calculation_note: 'Footing PCC/RCC only (as asked) — not full BOQ',
    },
  };
}

function readDrawingFully({
  text,
  filename,
  question = '',
  hints = [],
  boqOpts = {},
  extracted = null,
} = {}) {
  const intent = detectIntent(question);
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

  const needsCalc = intent === 'boq' || intent === 'takeoff' || intent === 'footing_calc';
  let boqResult = null;
  if (needsCalc && (inv.total_rows >= 1
    || /FOUNDATION|COLUMN|INDUSTRIAL|FLOOR/i.test(typeInfo.drawing_type))) {
    // takeoff uses full schedule BOQ path; footing_calc still filtered
    const calcIntent = intent === 'takeoff' ? 'boq' : intent;
    boqResult = filterBoqForIntent(buildBoqFromSchedules(schedules, boqOpts), calcIntent);
    if (boqResult) boqResult.drawing_type = typeInfo.drawing_type;
  }

  // Clarifications only for the asked scope (not full BOQ grill on a read)
  const clarifyQuestion = intent === 'read' ? 'read drawing only'
    : intent === 'footing_calc' ? (question || 'calculate pcc rcc footing')
      : intent === 'boq' || intent === 'takeoff' ? (question || 'prepare takeoff quantities')
        : question;

  const clarifications = buildClarifications({
    text,
    schedules,
    typeInfo,
    title,
    question: clarifyQuestion,
    boqNotFound: needsCalc ? (boqResult?.not_found || []) : [],
  });

  // Soften: on pure read, only ask if extract empty / type unknown
  if (intent === 'read') {
    clarifications.questions = (clarifications.questions || []).filter(q =>
      /unreadable_drawing|drawing_type/.test(q.id));
  }

  const isFinal = needsCalc
    && clarifications.questions.length === 0
    && (boqResult?.boq?.length > 0)
    && inv.quality !== 'poor';
  const status = intent === 'read' ? 'READING'
    : isFinal ? 'FINAL'
      : (boqResult?.boq?.length || (needsCalc && inv.total_rows >= 1)) ? 'DRAFT'
        : 'READING_ONLY';

  const parts = [];

  // ── FOCUSED ANSWERS (no full report dump) ─────────────────────
  if (intent === 'footing_calc') {
    parts.push('# Answer: Footing PCC / RCC (as asked)');
    parts.push('_Full BOQ nahi — sirf jo poocha gaya._');
    parts.push('', formatTypeMarkdown(typeInfo));
    if (inv.footings) {
      parts.push('', '### Footing schedule (read)');
      parts.push(formatSchedulesMarkdown({
        ...schedules,
        schedules: { ...schedules.schedules, columns: [], doors: [], windows: [], beams: [] },
        total_schedule_rows: inv.footings,
      }));
    }
    if (boqResult?.boq?.length) {
      parts.push('', '### Calculation (formulas)');
      parts.push('| Item | Formula | Qty | Unit |');
      parts.push('|---|---|---:|---|');
      for (const i of boqResult.boq) {
        parts.push(`| ${i.description} | ${i.calc_note || i.source} | ${i.qty} | ${i.unit} |`);
      }
      const tq = boqResult.total_quantities || {};
      parts.push('');
      if (tq.rcc_total_cum != null) parts.push(`- **RCC footings total: ${tq.rcc_total_cum} cum**`);
      if (tq.pcc_total_cum != null) parts.push(`- **PCC under footings: ${tq.pcc_total_cum} cum**`);
    } else if (inv.footings) {
      // Sizes/depths read — show unit RCC now; totals after qty + PCC
      parts.push('', '### RCC volume (each footing)');
      parts.push('| Mark | Size (mm) | Depth | Formula | RCC / each |');
      parts.push('|---|---|---:|---|---:|');
      let unitSum = 0;
      for (const f of (schedules.schedules?.footings || [])) {
        const size = f.rcc_size_mm || f.pcc_size_mm || '';
        const m = String(size).match(/(\d+)\s*[xX×]\s*(\d+)/);
        const d = Number(f.depth_mm);
        if (!m || !d) continue;
        const each = (Number(m[1]) / 1000) * (Number(m[2]) / 1000) * (d / 1000);
        unitSum += each;
        parts.push(`| ${f.mark} | ${size} | ${d} | L×B×D | ${each.toFixed(3)} cum |`);
      }
      if (unitSum > 0) {
        parts.push('');
        parts.push(`_Unit volumes ready. **Total RCC/PCC** = each × qty (qty + PCC thickness next)._`);
      }
    } else {
      parts.push('', '_Footing table clear nahi padhi — neeche schedule lines type karo._');
    }
    if (boqResult?.not_found?.length && !inv.footings) {
      parts.push('', '### Need from you');
      for (const n of boqResult.not_found.filter(x => /footing|pcc/i.test(x))) parts.push(`- ${n}`);
    }
    parts.push(formatQuestionsMarkdown(clarifications));
    parts.push('', '> Sirf jawab do next message me. Full BOQ alag se poochna.');
  } else if (intent === 'levels') {
    parts.push('# Answer: Levels / heights (as asked)');
    if (levels.length) parts.push(...levels.map(l => `- ${l}`));
    else parts.push('- **not found** on extract');
    parts.push(formatQuestionsMarkdown(clarifications));
  } else if (intent === 'schedule') {
    parts.push('# Answer: Schedules (as asked)');
    parts.push(formatSchedulesMarkdown(schedules));
    parts.push(formatQuestionsMarkdown(clarifications));
  } else if (intent === 'qa') {
    parts.push('# Answer (focused)');
    parts.push(formatTypeMarkdown(typeInfo));
    const hits = relevantLines(text, question, 25);
    if (hits.length) {
      parts.push('', '### From drawing');
      for (const h of hits) parts.push(`- ${h}`);
    } else {
      parts.push('', '_Direct match weak — schedules/levels below if useful._');
      if (inv.total_rows) parts.push(formatSchedulesMarkdown(schedules));
    }
    parts.push(formatQuestionsMarkdown(clarifications));
    parts.push('', '> Full **BOQ** chahiye to “BOQ” likho. Default = drawing read / focused answer.');
  } else if (intent === 'boq' || intent === 'takeoff') {
    parts.push(intent === 'takeoff' ? '# Quantity Takeoff (from drawing)' : '# BOQ (explicitly requested)');
    parts.push(`**Status: ${status}** · Read schedules → calculate → confirm gaps → Excel`);
    parts.push(formatTypeMarkdown(typeInfo));
    if (inv.total_rows) parts.push('', formatSchedulesMarkdown(schedules));
    else parts.push('', '_Schedule rows weak — confirm sizes/qty below, then volumes lock._');
    if (boqResult) {
      parts.push('', '### Calculations');
      if (boqResult.boq?.length) {
        parts.push('| Item | Formula | Qty | Unit |');
        parts.push('|---|---|---:|---|');
        for (const i of boqResult.boq) {
          parts.push(`| ${i.description} | ${i.calc_note || i.source} | ${i.qty} | ${i.unit} |`);
        }
      } else {
        parts.push('_Draft quantities pending confirmations (qty / PCC / height)._');
      }
      parts.push('', formatBoqMarkdown(boqResult, { status }));
    }
    // Footing unit volumes when takeoff has sizes but no totals yet
    if (intent === 'takeoff' && inv.footings && !(boqResult?.boq?.length)) {
      parts.push('', '### RCC volume (each footing)');
      parts.push('| Mark | Size (mm) | Depth | RCC / each |');
      parts.push('|---|---|---:|---:|');
      for (const f of (schedules.schedules?.footings || [])) {
        const size = f.rcc_size_mm || f.pcc_size_mm || '';
        const m = String(size).match(/(\d+)\s*[xX×]\s*(\d+)/);
        const d = Number(f.depth_mm);
        if (!m || !d) continue;
        const each = (Number(m[1]) / 1000) * (Number(m[2]) / 1000) * (d / 1000);
        parts.push(`| ${f.mark} | ${size} | ${d} | ${each.toFixed(3)} cum |`);
      }
    }
    parts.push(formatQuestionsMarkdown(clarifications));
  } else {
    // intent === 'read' — DEFAULT: drawing reading ONLY, no BOQ
    parts.push('# Drawing Reading Report');
    parts.push('**Status: READING** · Default = sheet padho. BOQ tabhi jab aap **BOQ** poocho.');
    parts.push('');
    parts.push('## 1. Sheet identity');
    parts.push(formatTypeMarkdown(typeInfo));
    parts.push(`- File: \`${filename || 'upload'}\``);
    if (title.project) parts.push(`- Project / title: ${title.project}`);
    if (title.title) parts.push(`- Sheet: ${title.title}`);
    if (title.drawingNo) parts.push(`- Drawing No: **${title.drawingNo}**`);
    else parts.push('- Drawing No: **not found**');
    if (title.scale) parts.push(`- Scale: **${title.scale}**`);
    else parts.push('- Scale: **not found**');
    if (schedules.meta?.concrete_grade || schedules.meta?.steel_grade) {
      parts.push(`- Grades: ${[schedules.meta.concrete_grade, schedules.meta.steel_grade].filter(Boolean).join(' / ')}`);
    }
    parts.push(`- Can answer: ${capabilitiesForType(typeInfo.drawing_type).join('; ')}`);

    parts.push('', '## 2. What was found');
    parts.push('| Item | Count / status |');
    parts.push('|---|---|');
    parts.push(`| Schedule quality | **${inv.quality}** (${inv.total_rows} rows) |`);
    parts.push(`| Column rows | ${inv.columns} |`);
    parts.push(`| Footing rows | ${inv.footings} |`);
    parts.push(`| Door / window | ${inv.doors} / ${inv.windows} |`);
    parts.push(`| Levels | ${levels.length} |`);
    parts.push(`| Size marks | ${sizes.length} |`);
    parts.push('', '### Checklist');
    for (const c of checklist) parts.push(`- [${c.ok ? 'x' : ' '}] ${c.label}`);

    parts.push('', '## 3. How to read this sheet');
    if (typeInfo.drawing_type.includes('FOUNDATION') || inv.footings) {
      parts.push(
        '1. **SCHEDULE OF FOOTING / COLUMN** pehle.',
        '2. Plan = location; qty schedule se.',
        '3. PCC/RCC volume = L×B×D×Qty — jab aap calculate/BOQ poocho.',
      );
    } else if (typeInfo.drawing_type === 'SECTION') {
      parts.push('1. Section marks.', '2. Printed levels only.', '3. Sizes from callouts/schedules.');
    } else {
      parts.push(`1. ${capabilitiesForType(typeInfo.drawing_type).join('; ')}.`);
    }

    parts.push('', '## 4. Schedules / tables');
    if (inv.total_rows >= 1) parts.push(formatSchedulesMarkdown(schedules));
    else parts.push('_No schedule rows yet — clearer PDF/screenshot try karo._');

    parts.push('', '## 5. Levels, dims & notes');
    if (levels.length) {
      parts.push('### Levels', ...levels.map(l => `- ${l}`));
    } else parts.push('### Levels', '- **not found**');
    if (sections.length) parts.push('', '### Sections', sections.map(s => `\`${s}\``).join(' · '));
    if (sizes.length) parts.push('', '### Sizes', sizes.slice(0, 40).map(s => `\`${s}\``).join(' · '));
    if (notes.length) {
      parts.push('', '### Notes');
      for (const n of notes.slice(0, 12)) parts.push(`- ${n}`);
    }

    parts.push(formatQuestionsMarkdown(clarifications));
    parts.push('', '> Next: specific poocho — e.g. **“PCC RCC footing calculate”** ya **“BOQ”**. Default me BOQ nahi banaya.');
  }

  // Always treat local report as valid — never drop ask-user into Claude polish
  const answeredLocally = true;
  const needsClaude = String(text || '').trim().length < 40
    && clarifications.questions.length === 0
    && inv.total_rows === 0;

  return {
    answeredLocally,
    needsClaude,
    needsUserInput: clarifications.questions.length > 0,
    clarifications,
    markdown: parts.join('\n'),
    status,
    intent,
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
      intent,
    },
  };
}

module.exports = {
  readDrawingFully,
  detectIntent,
  findLevels,
  findSizes,
  findDimensions,
  findSectionMarks,
  findNotes,
  findTitleBits,
};
