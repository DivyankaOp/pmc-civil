'use strict';
/**
 * Multi-type civil drawing classifier (local, zero tokens).
 * Drawings are NOT all footing schedules — detect type so Q&A + takeoff match the sheet.
 */

const TYPE_RULES = [
  { type: 'SECTION', score: (t, f) => score(t, f, [/section\s*[a-z]/i, /sec\.?\s*[a-z]-[a-z]/i, /sectional/i, /b_sections|sections?/i], [/section/i]) },
  { type: 'ELEVATION', score: (t, f) => score(t, f, [/\belevation\b/i, /\belev\b/i], [/elev/i]) },
  { type: 'FOUNDATION_FOOTING', score: (t, f) => score(t, f, [/schedule of footing/i, /footing schedule/i, /foundation layout/i, /column footing/i], [/footing|foundation/i]) },
  { type: 'COLUMN_SCHEDULE', score: (t, f) => score(t, f, [/schedule of column/i, /column schedule/i], [/column/i]) },
  { type: 'FLOOR_PLAN', score: (t, f) => score(t, f, [/floor plan/i, /ground floor plan/i, /typical floor/i], [/floor.?plan|gf_plan/i]) },
  { type: 'STRUCTURAL_DETAIL', score: (t, f) => score(t, f, [/typical detail/i, /beam detail/i, /slab detail/i, /reinforcement detail/i], [/detail/i]) },
  { type: 'SITE_PLAN', score: (t, f) => score(t, f, [/site plan/i, /key plan/i, /master plan/i], [/site.?plan/i]) },
  { type: 'ROAD', score: (t, f) => score(t, f, [/\bgsb\b/i, /\bwmm\b/i, /\bpqc\b/i, /chainage/i], [/road|highway/i]) },
  { type: 'INDUSTRIAL_PEB', score: (t, f) => score(t, f, [/base plate/i, /anchor bolt/i, /braced bay/i, /peb|pre.?engineered/i], [/warehouse|peb|shed/i]) },
];

function score(text, filename, textPats, filePats) {
  let s = 0;
  for (const p of textPats) if (p.test(text)) s += 2;
  for (const p of filePats) if (p.test(filename)) s += 3;
  return s;
}

function detectDrawingType({ text = '', filename = '', hints = [] } = {}) {
  const blob = `${text}\n${(hints || []).join(' ')}`;
  const fname = String(filename || '');
  const scored = TYPE_RULES.map(r => ({ type: r.type, score: r.score(blob, fname) }))
    .sort((a, b) => b.score - a.score);
  const best = scored[0];
  const secondary = scored.filter(x => x.score > 0 && x.type !== best.type).slice(0, 3).map(x => x.type);

  let drawingType = best.score > 0 ? best.type : 'GENERAL_DRAWING';
  // Filename hard hints
  if (/section/i.test(fname) && best.score < 3) drawingType = 'SECTION';
  if (/elev/i.test(fname) && !/section/i.test(fname)) drawingType = 'ELEVATION';
  if (/footing|foundation/i.test(fname) && best.score < 3) drawingType = 'FOUNDATION_FOOTING';

  return {
    drawing_type: drawingType,
    confidence: best.score >= 4 ? 'high' : best.score >= 2 ? 'medium' : 'low',
    secondary_types: secondary,
    scores: Object.fromEntries(scored.map(x => [x.type, x.score])),
  };
}

/** What this drawing type can answer / compute */
function capabilitiesForType(drawingType) {
  const map = {
    SECTION: ['floor heights', 'slab levels', 'wall sections', 'beam depths', 'study/explain section marks'],
    ELEVATION: ['building height', 'floor levels', 'facade elements', 'parapet'],
    FOUNDATION_FOOTING: ['footing schedule', 'PCC/RCC volumes', 'footing sizes', 'BOQ'],
    COLUMN_SCHEDULE: ['column sizes', 'bars/stirrups', 'concrete grade'],
    FLOOR_PLAN: ['rooms/areas', 'door/window schedules', 'wall lengths'],
    STRUCTURAL_DETAIL: ['bar details', 'cover', 'typical sizes'],
    SITE_PLAN: ['plot size', 'setbacks', 'road/entry'],
    ROAD: ['GSB/WMM/PQC lengths', 'chainage quantities'],
    INDUSTRIAL_PEB: ['base plates', 'pedestals', 'footing schedule', 'braced bays'],
    GENERAL_DRAWING: ['title block', 'scale', 'printed notes'],
  };
  return map[drawingType] || map.GENERAL_DRAWING;
}

function formatTypeMarkdown(info) {
  return [
    `## Drawing type: **${info.drawing_type}** (${info.confidence})`,
    info.secondary_types?.length ? `Also looks like: ${info.secondary_types.join(', ')}` : '',
    `Can answer: ${capabilitiesForType(info.drawing_type).join('; ')}`,
  ].filter(Boolean).join('\n');
}

module.exports = {
  detectDrawingType,
  capabilitiesForType,
  formatTypeMarkdown,
  TYPE_RULES,
};
