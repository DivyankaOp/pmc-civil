'use strict';
/**
 * Drawing study pipeline (Civils.ai-style, multi-type).
 * 1) Local text / CAD-zoom OCR (+ spatial tables)
 * 2) Detect drawing type (section / footing / plan / …)
 * 3) Schedules → BOQ when applicable
 * 4) Local Q&A — Claude ONLY if extract is weak
 * 5) User answers → merge → re-BOQ (no invent)
 */

const { extractSchedules, formatSchedulesMarkdown, normalizeLines } = require('./schedule_extractor');
const { buildBoqFromSchedules, formatBoqMarkdown } = require('./boq_from_schedules');
const { answerFromDrawing } = require('./drawing_qa');
const { detectDrawingType } = require('./drawing_types');
const { mergeUserAnswers } = require('./clarifications');

function buildReportMarkdown(extracted, boqResult, extraNotes = []) {
  const parts = [
    formatSchedulesMarkdown(extracted),
    '',
    formatBoqMarkdown(boqResult),
  ];
  if (extraNotes?.length) {
    parts.push('', '### Notes', ...extraNotes.map(n => `- ${n}`));
  }
  parts.push('', '> Pipeline: **read → calculate → BOQ**. Local extract first; API only if needed.');
  return parts.join('\n');
}

/**
 * Run local schedule→BOQ without Claude.
 * @param {string} text - combined PyMuPDF + GCV / CAD-zoom OCR text
 * @param {object} [opts]
 */
function runScheduleFirstLocal(text, opts = {}) {
  const extracted = opts.extracted
    || extractSchedules(text, { spatialTables: opts.spatialTables || [] });
  const typeInfo = detectDrawingType({
    text,
    filename: opts.filename || '',
    hints: opts.hints || [],
  });

  const qa = answerFromDrawing({
    text,
    filename: opts.filename || '',
    question: opts.question || 'Read this drawing only — sheet type, title, schedules, levels, dims, notes. Do not prepare BOQ unless asked.',
    hints: opts.hints || typeInfo.secondary_types || [],
    boqOpts: opts.boqOpts || {},
    extracted,
  });

  const boqResult = qa.meta?.boqResult || buildBoqFromSchedules(extracted, opts.boqOpts || {});
  boqResult.drawing_type = opts.drawing_type || boqResult.drawing_type || typeInfo.drawing_type;

  const markdown = qa.markdown || buildReportMarkdown(extracted, boqResult);
  const needsVision = extracted.quality === 'poor' && String(text || '').length < 400;
  return {
    extracted,
    boqResult,
    markdown,
    needsVision,
    needsClaude: qa.needsClaude && !qa.needsUserInput,
    needsUserInput: !!qa.needsUserInput,
    clarifications: qa.clarifications || { questions: [] },
    answeredLocally: qa.answeredLocally,
    typeInfo,
    qa,
  };
}

/**
 * Apply numbered user answers, rebuild BOQ, return fresh markdown.
 */
function applyUserClarifications({ text, extracted, clarifications, userText, filename, question, hints, boqOpts }) {
  const merged = mergeUserAnswers(
    { schedules: extracted, opts: boqOpts || {} },
    userText,
    clarifications
  );
  const schedules = merged.schedules;
  const opts = merged.opts || {};
  const typeInfo = detectDrawingType({
    text: text || '',
    filename: filename || '',
    hints: hints || [],
  });
  if (opts.drawing_type) typeInfo.drawing_type = opts.drawing_type;

  const qa = answerFromDrawing({
    text: text || formatSchedulesMarkdown(schedules),
    filename,
    question: question || 'Re-read with my confirmed answers: finalize takeoff calculations and FINAL BOQ',
    hints,
    boqOpts: opts,
    extracted: schedules,
    skipClarifications: false,
  });

  // Re-check remaining questions after merge
  const still = qa.clarifications?.questions || [];
  const answeredIds = new Set(Object.keys(merged.answers || {}));
  const remaining = still.filter(q => !answeredIds.has(q.id));
  const { formatOneQuestion } = require('./clarifications');

  let markdown;
  if (remaining.length) {
    // Claude-style: acknowledge + next single question only (no full dump)
    const answeredLabel = Object.values(merged.answers || {}).join(', ');
    markdown = [
      `✅ Noted: **${answeredLabel}**`,
      '',
      formatOneQuestion({ questions: remaining }, 0),
    ].join('\n');
  } else {
    markdown = String(qa.markdown || '')
      .replace(/\n## ❓ Question[\s\S]*$/i, '')
      .replace(/\n## ❓ Please confirm[\s\S]*$/i, '')
      + '\n\n> ✅ Sab answers mil gaye — result updated (no invented values).';
  }

  return {
    extracted: schedules,
    opts,
    answers: merged.answers,
    clarifications: { ...qa.clarifications, questions: remaining },
    needsUserInput: remaining.length > 0,
    markdown,
    boqResult: qa.meta?.boqResult || buildBoqFromSchedules(schedules, opts),
    typeInfo,
    answeredLocally: true,
  };
}

async function polishWithClaude(callClaudeAPI, { markdown, extracted, system }) {
  if (!callClaudeAPI || extracted?.quality === 'poor') {
    return markdown;
  }
  const prompt = `You are a PMC civil engineer. Below is MACHINE-EXTRACTED data + user questions if any.

ABSOLUTE RULES:
1. Do NOT invent or assume any qty, size, height, PCC thickness, or level.
2. Do NOT change numbers already extracted.
3. Keep "not found" / ASK USER questions as-is — never fill them with guesses.
4. Return markdown only.

${markdown.slice(0, 60000)}`;

  try {
    const raw = await callClaudeAPI({
      system: system || 'You format civil drawing takeoff reports. Never invent quantities.',
      messages: [{ role: 'user', content: prompt }],
      maxTokens: 4096,
    });
    return (raw && raw.trim().length > 100) ? raw : markdown;
  } catch (e) {
    console.warn('[schedule_pipeline] Claude polish failed:', e.message);
    return markdown;
  }
}

function textFromDrawingParts(parts) {
  return (parts || [])
    .filter(p => p.type === 'text' && p.text)
    .map(p => p.text)
    .join('\n');
}

function combineExtractedText(blocks) {
  return normalizeLines(blocks.filter(Boolean).join('\n')).join('\n');
}

module.exports = {
  runScheduleFirstLocal,
  applyUserClarifications,
  polishWithClaude,
  buildReportMarkdown,
  textFromDrawingParts,
  combineExtractedText,
  extractSchedules,
  buildBoqFromSchedules,
  formatSchedulesMarkdown,
  formatBoqMarkdown,
  answerFromDrawing,
  detectDrawingType,
};
