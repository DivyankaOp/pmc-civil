'use strict';
/**
 * Local Q&A over extracted drawing text (zero tokens when possible).
 * HARD RULE: never invent values — ask the user when unclear.
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
  const numeric = /(?:(?:PLINTH|FFL|SFL|TFL|RFL|T\.?O\.?S|B\.?O\.?S|N\.?G\.?L|G\.?L)\s*[:=]?\s*)?([+\-]?\d+(?:\.\d+)?)\s*(?:M|MT|m)\b|(?:LEVEL|LVL)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)/gi;
  let m;
  while ((m = numeric.exec(blob))) {
    levels.push(m[0].replace(/\s+/g, ' ').trim());
  }
  const named = blob.match(
    /(?:roof\s*top|terrace|first\s*floor|second\s*floor|upper\s*ground|lower\s*ground|ground\s*level|plinth|mezzanine(?:\s*-?\s*\d+)?|road\s*level|basement)(?:\s*level)?/gi
  ) || [];
  for (const n of named) levels.push(n.replace(/\s+/g, ' ').trim());
  return [...new Set(levels.map(x => x.replace(/\s+/g, ' ')))].slice(0, 40);
}

function findSizes(text) {
  const sizes = [];
  const re = /(\d{2,4})\s*[xX×]\s*(\d{2,4})(?:\s*[xX×]\s*(\d{2,4}))?/g;
  let m;
  while ((m = re.exec(String(text)))) {
    sizes.push(m[0].replace(/\s+/g, ''));
  }
  return [...new Set(sizes)].slice(0, 60);
}

function findTitleBits(text) {
  const lines = linesOf(text);
  const project = lines.find(l => /project|proposed|warehouse|residential|building/i.test(l)) || '';
  const drawingNo = (String(text).match(/drawing\s*no\.?\s*[:\-]?\s*([A-Z0-9\-\/_.]+)/i) || [])[1] || '';
  const scale = (String(text).match(/scale\s*[:\-]?\s*(1\s*:\s*\d+)/i) || [])[1] || '';
  const title = lines.find(l => /layout|section|elevation|footing|plan|detail/i.test(l)) || '';
  return { project: project.slice(0, 140), drawingNo, scale, title: title.slice(0, 140) };
}

function relevantLines(text, question, limit = 25) {
  const q = String(question || '').toLowerCase();
  const keys = q
    .split(/[^a-z0-9]+/i)
    .filter(w => w.length > 2 && !['the', 'and', 'for', 'from', 'what', 'how', 'please', 'calculate', 'drawing'].includes(w));
  const lines = linesOf(text);
  const scored = lines.map(line => {
    const low = line.toLowerCase();
    let s = 0;
    for (const k of keys) if (low.includes(k)) s += 1;
    if (/schedule|section|level|footing|column|pcc|rcc|beam|slab/i.test(line)) s += 0.5;
    return { line, s };
  }).filter(x => x.s > 0)
    .sort((a, b) => b.s - a.s);
  return scored.slice(0, limit).map(x => x.line);
}

/**
 * @returns {{ answeredLocally, markdown, needsClaude, needsUserInput, clarifications, meta }}
 */
function answerFromDrawing({ text, filename, question, hints = [], boqOpts = {} }) {
  const typeInfo = detectDrawingType({ text, filename, hints });
  const schedules = extractSchedules(text);
  const title = findTitleBits(text);
  const levels = findLevels(text);
  const sizes = findSizes(text);
  const q = String(question || '').toLowerCase();

  const parts = [];
  parts.push('## Policy: **No assumptions**');
  parts.push('Sirf drawing pe printed / aapke confirm kiye values. Jo clear nahi — neeche poochhenge.');
  parts.push('');
  parts.push(formatTypeMarkdown(typeInfo));

  if (title.project || title.drawingNo || title.scale) {
    parts.push('', '### Title block (found)');
    if (title.project) parts.push(`- Project/title: ${title.project}`);
    if (title.title) parts.push(`- Sheet: ${title.title}`);
    if (title.drawingNo) parts.push(`- Drawing No: ${title.drawingNo}`);
    if (title.scale) parts.push(`- Scale: ${title.scale}`);
    else parts.push('- Scale: **not found**');
  }

  const wantsBoq = /boq|quantity|pcc|rcc|cum|estimate|rate|bill of|calculate|volume/i.test(q);
  const wantsStudy = /study|explain|kaise|what is|indicate|padho|read|samjhao|walkthrough/i.test(q) || !q.trim();
  const wantsLevels = /level|height|plinth|ffl|elevation|storey/i.test(q);
  const wantsSchedule = /schedule|table|footing|column|beam|door|window/i.test(q);

  let answeredLocally = false;
  let boqResult = null;

  if (schedules.total_schedule_rows >= 1) {
    parts.push('', formatSchedulesMarkdown(schedules));
    answeredLocally = schedules.quality !== 'poor';
  }

  if (wantsBoq && (typeInfo.drawing_type.includes('FOUNDATION') || typeInfo.drawing_type.includes('COLUMN') || typeInfo.drawing_type === 'INDUSTRIAL_PEB' || schedules.total_schedule_rows >= 1)) {
    boqResult = buildBoqFromSchedules(schedules, boqOpts);
    parts.push('', formatBoqMarkdown(boqResult));
    answeredLocally = answeredLocally || boqResult.boq.length > 0;
  }

  if (wantsLevels || typeInfo.drawing_type === 'SECTION' || typeInfo.drawing_type === 'ELEVATION') {
    parts.push('', '### Levels / heights found on drawing');
    if (levels.length) {
      parts.push(...levels.map(l => `- ${l}`));
      answeredLocally = true;
    } else {
      parts.push('- **not found** — please type the printed levels if you can read them');
    }
  }

  if (wantsStudy || typeInfo.drawing_type === 'SECTION' || typeInfo.drawing_type === 'ELEVATION') {
    parts.push('', '### How to read this sheet (CAD-style)');
    if (typeInfo.drawing_type === 'SECTION') {
      parts.push(
        '- Section marks (A-A, B-B…) = cut through plan; use printed levels only.',
        '- Table values (if any) beat guessed dimensions from the image.',
      );
    } else if (typeInfo.drawing_type === 'ELEVATION') {
      parts.push('- Elevation = outside face; openings/levels only if printed.');
    } else if (typeInfo.drawing_type.includes('FOUNDATION')) {
      parts.push(
        '- Read SCHEDULE OF FOOTING / COLUMN first.',
        '- Plan is for location; sizes/qty from schedule cells only.',
      );
    } else {
      parts.push(`- This sheet supports: ${capabilitiesForType(typeInfo.drawing_type).join('; ')}.`);
    }
    answeredLocally = true;
  }

  if (sizes.length && (wantsSchedule || wantsStudy || wantsBoq)) {
    parts.push('', '### Size marks seen (mm)', sizes.slice(0, 30).map(s => `\`${s}\``).join(' · '));
  }

  const hits = relevantLines(text, question, 20);
  if (hits.length) {
    parts.push('', '### Lines matching your question');
    for (const h of hits) parts.push(`- ${h}`);
    answeredLocally = true;
  }

  const clarifications = buildClarifications({
    text,
    schedules,
    typeInfo,
    title,
    question,
    boqNotFound: boqResult?.not_found || [],
  });
  parts.push(formatQuestionsMarkdown(clarifications));

  const weakText = String(text || '').trim().length < 120;
  // Prefer asking user over burning Claude tokens when we already know what's missing
  const needsUserInput = clarifications.questions.length > 0;
  const needsClaude = weakText && !needsUserInput;

  if (needsUserInput) {
    parts.push('', '> **Waiting for your answers** — phir exact BOQ/final report banega. Assume nahi kiya.');
  } else if (!answeredLocally) {
    parts.push('', '_Extract weak hai. Clearer PDF/DXF bhejo ya neeche type confirm karo._');
  } else {
    parts.push('', '> Local extract se jawab (no invented numbers).');
  }

  return {
    answeredLocally: (answeredLocally || needsUserInput) && !weakText,
    needsClaude,
    needsUserInput,
    clarifications,
    markdown: parts.join('\n'),
    meta: {
      typeInfo,
      schedules,
      title,
      levels,
      size_count: sizes.length,
      text_chars: String(text || '').length,
      boqResult,
    },
  };
}

module.exports = {
  answerFromDrawing,
  findLevels,
  findSizes,
  relevantLines,
};
