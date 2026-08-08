'use strict';
/**
 * No-assumption policy for any drawing type.
 * If a value is missing / unclear → ASK the user. Never invent.
 */

function pushQ(list, id, question, why, field) {
  if (list.some(q => q.id === id)) return;
  list.push({
    id,
    field: field || id,
    question,
    why: why || 'Not clearly readable on the drawing',
    status: 'needs_user',
  });
}

/**
 * Build clarification questions from extract + schedules + question intent.
 * @returns {{ questions: Array, blocked: boolean, summary: string }}
 */
function buildClarifications({
  text = '',
  schedules = null,
  typeInfo = null,
  title = null,
  question = '',
  boqNotFound = [],
} = {}) {
  const questions = [];
  const q = String(question || '').toLowerCase();
  // Only when user asked calc/BOQ — not on plain "read drawing"
  const wantsBoq = /\bboq\b|quantity|pcc|rcc|cum|estimate|volume|calculate|takeoff/i.test(q)
    && !/read drawing only|do not prepare boq|without boq/i.test(q);
  const textLen = String(text || '').trim().length;

  if (textLen < 80) {
    pushQ(questions, 'unreadable_drawing',
      'Drawing text/tables clearly nahi padh paaye. Kya aap PDF/DXF (vector) upload kar sakte ho, ya batao yeh sheet kis type ki hai (section / footing / plan / elevation)?',
      'OCR/text extract almost empty');
  }

  if (typeInfo?.confidence === 'low' || typeInfo?.drawing_type === 'GENERAL_DRAWING' || typeInfo?.drawing_type === 'UNKNOWN') {
    pushQ(questions, 'drawing_type',
      'Yeh drawing kis type ki hai? (SECTION / ELEVATION / FLOOR_PLAN / FOUNDATION_FOOTING / COLUMN_SCHEDULE / SITE_PLAN / ROAD / DETAIL)',
      'Auto type detection uncertain');
  }

  // Scale only when takeoff needs geometry scaling (not for schedule-cell BOQ)
  const needsScale = /scale|measure|wall length|from plan|dimension check/i.test(q)
    || (typeInfo?.drawing_type === 'FLOOR_PLAN' && /area|length|wall/i.test(q));
  if (needsScale && !title?.scale) {
    pushQ(questions, 'scale',
      'Drawing ka scale kya hai? (jaise 1:100). Agar title block me hai to type kar do.',
      'Scale needed for plan measurements');
  }

  const cols = schedules?.schedules?.columns || [];
  const ftgs = schedules?.schedules?.footings || [];
  const doors = schedules?.schedules?.doors || [];
  const windows = schedules?.schedules?.windows || [];
  const focusFooting = /footing|pcc|ftg/i.test(q) && !/column\s*schedule|all\s*boq|full\s*boq/i.test(q);

  // Footing questions FIRST when user asked PCC/RCC footing
  for (const row of ftgs) {
    const mark = row.mark || '?';
    if ((!row.rcc_size_mm || /not found/i.test(row.rcc_size_mm)) && (!row.pcc_size_mm || /not found/i.test(row.pcc_size_mm))) {
      pushQ(questions, `ftg_size_${mark}`,
        `Footing **${mark}** ka size L×B (mm) schedule me kya hai?`,
        'Footing size missing', 'footing.size_mm');
    }
    if (row.qty == null) {
      pushQ(questions, `ftg_qty_${mark}`,
        `Footing **${mark}** (${row.rcc_size_mm || row.pcc_size_mm || '?'}) ki QTY schedule/plan me kitni hai?`,
        'Qty missing — needed for RCC/PCC volume', 'footing.qty');
    }
    if (row.depth_mm == null && wantsBoq) {
      pushQ(questions, `ftg_depth_${mark}`,
        `Footing **${mark}** ki depth D (mm) schedule/section me kya hai?`,
        'Depth needed for RCC — will not assume', 'footing.depth_mm');
    }
  }

  if (!focusFooting) {
    for (const row of cols) {
      const mark = row.mark || '?';
      if (!row.size_mm || /not found/i.test(row.size_mm)) {
        pushQ(questions, `col_size_${mark}`,
          `Column/Pedestal **${mark}** ka size (mm×mm) schedule me kya hai?`,
          'Size cell missing/unclear', 'column.size_mm');
      }
      if (row.qty == null) {
        pushQ(questions, `col_qty_${mark}`,
          `Column **${mark}** ki QTY schedule me kitni hai? (plan se count assume nahi karenge)`,
          'Qty missing — will not invent', 'column.qty');
      }
      if (wantsBoq && (!row.height_m && !row.floor_height_m)) {
        pushQ(questions, `col_height_${mark}`,
          `Column **${mark}** ki height (m) schedule/section me kya printed hai? (default 3m use nahi karenge)`,
          'Height not on schedule — needed for RCC volume', 'column.height_m');
      }
    }
  }

  if (wantsBoq && (ftgs.length || /pcc/i.test(q))) {
    const pccKnown = schedules?.meta?.pcc_thickness_mm
      || /pcc[^a-z0-9]{0,12}(\d{2,3})\s*mm|(\d{2,3})\s*mm[^a-z0-9]{0,12}pcc/i.test(text);
    if (!pccKnown) {
      pushQ(questions, 'pcc_thickness',
        'PCC thickness (mm) drawing me kya printed hai? (100 / 150 / other). Assume nahi karenge.',
        'PCC thickness not clearly printed');
    }
    if (!/pcc[^a-z0-9]{0,20}(\d{2,3})\s*mm\s*(offset|proj)|(\d{2,3})\s*mm\s*(offset|beyond|projection)/i.test(text)
      && !/115\s*mm/i.test(text)) {
      pushQ(questions, 'pcc_offset',
        'PCC offset beyond footing (mm each side) drawing me kya hai?',
        'PCC offset not confirmed');
    }
  }

  for (const row of doors) {
    if (row.qty == null) {
      pushQ(questions, `door_qty_${row.mark}`,
        `Door **${row.mark}** ki QTY schedule me kitni hai?`,
        'Door qty missing');
    }
  }
  for (const row of windows) {
    if (row.qty == null) {
      pushQ(questions, `win_qty_${row.mark}`,
        `Window **${row.mark}** ki QTY schedule me kitni hai?`,
        'Window qty missing');
    }
  }

  // Only add leftover notFound that aren't already covered by structured questions
  for (const nf of boqNotFound || []) {
    const s = String(nf);
    // Skip duplicates already asked as structured height/qty/pcc questions
    if (/height not on drawing|PCC thickness not|qty not found|size not found|depth not found/i.test(s)) continue;
    if (/excavation surcharge/i.test(s)) {
      pushQ(questions, 'excavation_method',
        'Excavation working space / surcharge drawing me diya hai kya? Agar haan to mm/m batao — warna excavation line skip rahegi (assume nahi).',
        'No printed excavation allowance');
      continue;
    }
    const id = `nf_${s.slice(0, 40).replace(/\W+/g, '_')}`;
    pushQ(questions, id, `Confirm: ${s}`, 'Flagged during takeoff');
  }

  // Cap questions so chat stays usable
  const capped = questions.slice(0, 10);
  const blocked = capped.length > 0 && wantsBoq;

  return {
    questions: capped,
    blocked,
    summary: capped.length
      ? `${capped.length} value(s) unclear — please answer before final BOQ (no assumptions).`
      : 'No blocking clarifications.',
  };
}

function formatQuestionsMarkdown(clarifications) {
  const qs = clarifications?.questions || [];
  if (!qs.length) return '';
  const lines = [
    '',
    '## ❓ Please confirm (no assumptions)',
    clarifications.summary || '',
    '',
    'Reply numbered answers, e.g.:',
    '`1) SECTION  2) 1:100  3) C1 size 300x450 qty 12 height 3.15`',
    '',
  ];
  qs.forEach((q, i) => {
    lines.push(`**${i + 1}. ${q.question}**`);
    lines.push(`   _Why:_ ${q.why}`);
    lines.push('');
  });
  lines.push('> Jab tak aap confirm nahi karte, missing values **invent/assume nahi** kiye jayenge.');
  return lines.join('\n');
}

/**
 * Apply user answers map { questionId or index: value } onto schedules / opts.
 * Lightweight parser for free-text replies.
 */
function mergeUserAnswers(base, userText, clarifications) {
  const opts = { ...(base?.opts || {}) };
  const schedules = JSON.parse(JSON.stringify(base?.schedules || { schedules: {} }));
  const answers = {};
  const text = String(userText || '');

  // Pattern: "1) value" or "1. value"
  const numbered = [...text.matchAll(/(?:^|\n)\s*(\d+)\s*[\)\.\-:]\s*(.+)/g)];
  const qs = clarifications?.questions || [];
  for (const m of numbered) {
    const idx = Number(m[1]) - 1;
    if (qs[idx]) answers[qs[idx].id] = m[2].trim();
  }

  for (const [id, val] of Object.entries(answers)) {
    if (id === 'drawing_type') opts.drawing_type = val.toUpperCase();
    if (id === 'scale') opts.scale = val;
    if (id === 'pcc_thickness') {
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      if (n) opts.pccThicknessM = n / (n > 20 ? 1000 : 1); // 100 → mm, 0.1 → m
      if (n > 1) opts.pccThicknessM = n / 1000;
    }
    if (id === 'pcc_offset') {
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      if (n) opts.pccOffsetM = n / 1000;
    }
    const colH = id.match(/^col_height_(.+)$/);
    if (colH) {
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      opts.columnHeights = opts.columnHeights || {};
      opts.columnHeights[colH[1]] = n;
    }
    const colQty = id.match(/^col_qty_(.+)$/);
    if (colQty && schedules.schedules?.columns) {
      const row = schedules.schedules.columns.find(c => c.mark === colQty[1]);
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      if (row && n) { row.qty = n; row.source = 'user-confirmed'; }
    }
    const colSize = id.match(/^col_size_(.+)$/);
    if (colSize && schedules.schedules?.columns) {
      const row = schedules.schedules.columns.find(c => c.mark === colSize[1]);
      const sz = String(val).replace(/\s+/g, '').match(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/);
      if (row && sz) { row.size_mm = `${sz[1]}x${sz[2]}`; row.source = 'user-confirmed'; }
    }
    const ftgQty = id.match(/^ftg_qty_(.+)$/);
    if (ftgQty && schedules.schedules?.footings) {
      const row = schedules.schedules.footings.find(c => c.mark === ftgQty[1]);
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      if (row && n) { row.qty = n; row.source = 'user-confirmed'; }
    }
    const ftgDepth = id.match(/^ftg_depth_(.+)$/);
    if (ftgDepth && schedules.schedules?.footings) {
      const row = schedules.schedules.footings.find(c => c.mark === ftgDepth[1]);
      const n = Number(String(val).replace(/[^\d.]/g, ''));
      if (row && n) { row.depth_mm = n; row.source = 'user-confirmed'; }
    }
    const ftgSize = id.match(/^ftg_size_(.+)$/);
    if (ftgSize && schedules.schedules?.footings) {
      const row = schedules.schedules.footings.find(c => c.mark === ftgSize[1]);
      const sz = String(val).replace(/\s+/g, '').match(/(\d{2,4})\s*[xX×]\s*(\d{2,4})/);
      if (row && sz) {
        row.rcc_size_mm = `${sz[1]}x${sz[2]}`;
        row.pcc_size_mm = row.pcc_size_mm || row.rcc_size_mm;
        row.source = 'user-confirmed';
      }
    }
  }

  return { schedules, opts, answers };
}

module.exports = {
  buildClarifications,
  formatQuestionsMarkdown,
  mergeUserAnswers,
};
