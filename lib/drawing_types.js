'use strict';
/**
 * Multi-type civil drawing classifier (local, zero tokens).
 * Schedule-header evidence beats generic "section" keyword hits on mixed sheets.
 */

const TYPE_RULES = [
  { type: 'SECTION', score: (t, f) => score(t, f, [/section\s*[a-z]/i, /sec\.?\s*[a-z]-[a-z]/i, /sectional/i, /b_sections/i], [/section/i]) },
  { type: 'ELEVATION', score: (t, f) => score(t, f, [/\belevation\b/i, /\belev\b/i], [/elev/i]) },
  { type: 'FOUNDATION_FOOTING', score: (t, f) => score(t, f, [/schedule of footing/i, /footing schedule/i, /foundation layout/i, /column footing/i, /schedule\s*of\s*ftg/i], [/footing|foundation/i]) },
  { type: 'COLUMN_SCHEDULE', score: (t, f) => score(t, f, [/schedule of column/i, /column schedule/i, /pedestal schedule/i], [/column/i]) },
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

  // Schedule-header evidence overrides weak/generic section hits
  const hasFootingSch = /schedule\s*of\s*footing|footing\s*schedule|foundation\s*schedule/i.test(blob);
  const hasColumnSch = /schedule\s*of\s*column|column\s*schedule|pedestal\s*schedule/i.test(blob);
  const hasDoorSch = /door\s*schedule|schedule\s*of\s*doors/i.test(blob);
  if (hasFootingSch) {
    const ftg = scored.find(x => x.type === 'FOUNDATION_FOOTING');
    if (!ftg || best.type === 'SECTION' || best.score <= (ftg?.score || 0) + 2) {
      drawingType = 'FOUNDATION_FOOTING';
    }
  } else if (hasColumnSch && best.type !== 'FOUNDATION_FOOTING') {
    drawingType = 'COLUMN_SCHEDULE';
  } else if (hasDoorSch && best.score < 4) {
    drawingType = 'FLOOR_PLAN';
  }

  // Hints from CAD-zoom (prefer FOUNDATION_FOOTING enum over FOUNDATION)
  const hintBlob = (hints || []).join(' ');
  if (/FOUNDATION_FOOTING|FOUNDATION/i.test(hintBlob) && !/SECTION/i.test(fname)) {
    if (drawingType === 'GENERAL_DRAWING' || drawingType === 'SECTION') {
      drawingType = 'FOUNDATION_FOOTING';
    }
  }
  if (/COLUMN_SCHEDULE/i.test(hintBlob) && drawingType === 'GENERAL_DRAWING') {
    drawingType = 'COLUMN_SCHEDULE';
  }

  // Filename hard hints (only when scores weak)
  if (/section/i.test(fname) && best.score < 3 && !hasFootingSch) drawingType = 'SECTION';
  if (/elev/i.test(fname) && !/section/i.test(fname)) drawingType = 'ELEVATION';
  if (/footing|foundation/i.test(fname) && best.score < 3) drawingType = 'FOUNDATION_FOOTING';

  const confScore = scored.find(x => x.type === drawingType)?.score ?? best.score;
  return {
    drawing_type: drawingType,
    confidence: confScore >= 4 || hasFootingSch || hasColumnSch ? 'high' : confScore >= 2 ? 'medium' : 'low',
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
