'use strict';
/**
 * Local Q&A — delegates to Civils.ai-style drawing_reader (read first, BOQ last).
 */

const { readDrawingFully, findLevels, findSizes } = require('./drawing_reader');

function relevantLines(text, question, limit = 25) {
  const q = String(question || '').toLowerCase();
  const keys = q
    .split(/[^a-z0-9]+/i)
    .filter(w => w.length > 2 && !['the', 'and', 'for', 'from', 'what', 'how', 'please', 'calculate', 'drawing'].includes(w));
  const lines = String(text || '')
    .split(/\r?\n/)
    .map(l => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
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
function answerFromDrawing(opts) {
  // Default = auto takeoff/BOQ (Civils product); pass action=read|study to skip BOQ
  const question = opts.question
    || 'Prepare quantity takeoff / BOQ from this drawing — read schedules, calculate, ask missing values one-by-one.';
  const autoBoq = opts.autoBoq !== false && opts.action !== 'read' && opts.action !== 'study';
  return readDrawingFully({
    ...opts,
    question,
    action: opts.action || (autoBoq ? 'calculate' : null),
    autoBoq,
  });
}

module.exports = {
  answerFromDrawing,
  findLevels,
  findSizes,
  relevantLines,
  readDrawingFully,
};
