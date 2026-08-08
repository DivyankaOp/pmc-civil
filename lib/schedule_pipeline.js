'use strict';
/**
 * Drawing study pipeline (Civils.ai-style, multi-type).
 * 1) Local text / CAD-zoom OCR
 * 2) Detect drawing type (section / footing / plan / …)
 * 3) Schedules → BOQ when applicable
 * 4) Local Q&A — Claude ONLY if extract is weak
 */

const { extractSchedules, formatSchedulesMarkdown, normalizeLines } = require('./schedule_extractor');
const { buildBoqFromSchedules, formatBoqMarkdown } = require('./boq_from_schedules');
const { answerFromDrawing } = require('./drawing_qa');
const { detectDrawingType } = require('./drawing_types');

function buildReportMarkdown(extracted, boqResult, extraNotes = []) {
  const parts = [
    formatSchedulesMarkdown(extracted),
    '',
    formatBoqMarkdown(boqResult),
  ];
  if (extraNotes?.length) {
    parts.push('', '### Notes', ...extraNotes.map(n => `- ${n}`));
  }
  parts.push('', '> Pipeline: **local-first** (type detect → tables → BOQ). Tokens only if text/OCR weak.');
  return parts.join('\n');
}

/**
 * Run local schedule→BOQ without Claude.
 * @param {string} text - combined PyMuPDF + GCV / CAD-zoom OCR text
 * @param {object} [opts]
 */
function runScheduleFirstLocal(text, opts = {}) {
  const extracted = extractSchedules(text);
  const typeInfo = detectDrawingType({ text, filename: opts.filename || '', hints: opts.hints || [] });

  const qa = answerFromDrawing({
    text,
    filename: opts.filename || '',
    question: opts.question || 'study drawing and extract schedules / quantities',
    hints: opts.hints || typeInfo.secondary_types || [],
    boqOpts: opts.boqOpts || {},
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
 * One small Claude call: polish wording / fill drawing_type from text only.
 * Does NOT re-invent quantities.
 */
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
