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
  const wantsBoq = (!q || /\bboq\b|quantity|pcc|rcc|cum|estimate|volume|calculate|takeoff|measure|prepare\s*takeoff/i.test(q))
    && !/read drawing only|do not prepare boq|without boq/i.test(q);
  const textLen = String(text || '').trim().length;

  const focusFootingEarly = (
    /footing|pcc|ftg|rcc/i.test(q)
    || /takeoff|measure|quantit|prepare\s*takeoff/i.test(q)
    || !q
  ) && !/column\s*schedule|all\s*boq|full\s*boq/i.test(q);

  if (textLen < 80 && !focusFootingEarly) {
    pushQ(questions, 'unreadable_drawing',
      'Drawing text/tables clearly nahi padh paaye. Kya aap PDF/DXF (vector) upload kar sakte ho, ya batao yeh sheet kis type ki hai (section / footing / plan / elevation)?',
      'OCR/text extract almost empty');
  }

  if (!focusFootingEarly && (typeInfo?.confidence === 'low' || typeInfo?.drawing_type === 'GENERAL_DRAWING' || typeInfo?.drawing_type === 'UNKNOWN')) {
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
  // On takeoff / empty scope, if footing rows exist → finish footings before column grill
  const focusFooting = (
    /footing|pcc|ftg/i.test(q)
    || ((/takeoff|measure|quantit|prepare\s*takeoff/i.test(q) || !q) && (schedules?.schedules?.footings || []).length > 0)
  ) && !/column\s*schedule|all\s*boq|full\s*boq|columns?\s*only/i.test(q);

  // Footing questions FIRST when user asked PCC/RCC footing
  for (const row of ftgs) {
    const mark = row.mark || '?';
    if ((!row.rcc_size_mm || /not found/i.test(row.rcc_size_mm)) && (!row.pcc_size_mm || /not found/i.test(row.pcc_size_mm))) {
      pushQ(questions, `ftg_size_${mark}`,
        `Footing **${mark}** ka size L×B (mm) schedule me kya hai?`,
        'Footing size missing', 'footing.size_mm');
    }
    if (row.depth_mm == null && wantsBoq) {
      pushQ(questions, `ftg_depth_${mark}`,
        `Footing **${mark}** ki depth D (mm) schedule/section me kya hai?`,
        'Depth needed for RCC — will not assume', 'footing.depth_mm');
    }
  }
  // One qty question for all footings (faster than F1, F2, F3…)
  const missingQty = ftgs.filter(r => r.qty == null);
  if (missingQty.length === 1) {
    const row = missingQty[0];
    pushQ(questions, `ftg_qty_${row.mark}`,
      `Footing **${row.mark}** (${row.rcc_size_mm || row.pcc_size_mm || '?'}) ki QTY kitni hai?`,
      'Qty missing — needed for RCC/PCC volume', 'footing.qty');
  } else if (missingQty.length > 1) {
    const hint = missingQty.map(r => `${r.mark}=?`).join(', ');
    pushQ(questions, 'ftg_qty_all',
      `Har footing ki QTY likho (jaise \`F1=12, F2=8, F3=6\`). Schedule/plan se: ${hint}`,
      'Qtys missing — needed for total RCC/PCC');
  }

  // No footing rows extracted (scanned sheet) but user asked PCC/RCC → still ask one-by-one
  if (focusFooting && wantsBoq && !ftgs.length) {
    pushQ(questions, 'footing_schedule_paste',
      'SCHEDULE OF FOOTING drawing se type karo — har line: `Mark LxB Depth Qty` (jaise `F1 2600x1800 900 12`).',
      'OCR se footing table clear nahi padhi');
    pushQ(questions, 'pcc_thickness',
      'PCC thickness (mm) drawing me kya printed hai? (100 / 150 / other). Assume nahi karenge.',
      'PCC thickness needed for volume');
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

/**
 * One-by-one like Claude chat (default): show only the first pending question.
 * Pass { all: true } only if you really want the full list.
 */
function formatQuestionsMarkdown(clarifications, opts = {}) {
  const qs = clarifications?.questions || [];
  if (!qs.length) return '';
  const oneByOne = opts.all !== true;
  if (oneByOne) {
    const q = qs[0];
    const total = qs.length;
    return [
      '',
      `## ❓ Question 1 of ${total}`,
      '',
      `**${q.question}**`,
      q.why ? `_Why:_ ${q.why}` : '',
      '',
      '> Iska jawab type karo (ek line). Phir agla question aayega — saare ek saath nahi.',
    ].filter(Boolean).join('\n');
  }
  const lines = [
    '',
    '## ❓ Please confirm (no assumptions)',
    clarifications.summary || '',
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

function formatOneQuestion(clarifications, index = 0) {
  const qs = clarifications?.questions || [];
  if (!qs.length || index < 0 || index >= qs.length) return '';
  const q = qs[index];
  const total = qs.length;
  return [
    `## ❓ Question ${index + 1} of ${total}`,
    '',
    `**${q.question}**`,
    q.why ? `_Why:_ ${q.why}` : '',
    '',
    '> Sirf iska jawab do. Next question baad me aayega.',
  ].filter(Boolean).join('\n');
}

/**
 * Apply user answers map { questionId or index: value } onto schedules / opts.
 * Lightweight parser for free-text replies.
 */
function mergeUserAnswers(base, userText, clarifications) {
  const opts = { ...(base?.opts || {}) };
  const schedules = JSON.parse(JSON.stringify(base?.schedules || { schedules: {} }));
  const answers = {};
  const text = String(userText || '').trim();
  const qs = clarifications?.questions || [];

  // Pattern: "1) value" or "1. value"
  const numbered = [...text.matchAll(/(?:^|\n)\s*(\d+)\s*[\)\.\-:]\s*(.+)/g)];
  for (const m of numbered) {
    const idx = Number(m[1]) - 1;
    if (qs[idx]) answers[qs[idx].id] = m[2].trim();
  }

  // One-by-one chat: free-text answer applies to the first (current) question
  if (!Object.keys(answers).length && qs.length && text && !/^\s*\d+\s*[\)\.\-:]/.test(text)) {
    answers[qs[0].id] = text;
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
    if (id === 'ftg_qty_all' && schedules.schedules?.footings) {
      for (const m of String(val).matchAll(/([A-Z]?\d+[A-Z]?)\s*[=:]\s*(\d{1,4})/gi)) {
        const row = schedules.schedules.footings.find(c => c.mark.toUpperCase() === m[1].toUpperCase());
        const n = Number(m[2]);
        if (row && n) { row.qty = n; row.source = 'user-confirmed'; }
      }
      // also accept plain list of numbers in mark order: 12, 8, 6
      if (schedules.schedules.footings.every(f => f.qty == null)) {
        const nums = [...String(val).matchAll(/\b(\d{1,4})\b/g)].map(x => Number(x[1]));
        schedules.schedules.footings.forEach((row, i) => {
          if (nums[i]) { row.qty = nums[i]; row.source = 'user-confirmed'; }
        });
      }
    }
    const ftgQty = id.match(/^ftg_qty_(.+)$/);
    if (ftgQty && ftgQty[1] !== 'all' && schedules.schedules?.footings) {
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
    if (id === 'footing_schedule_paste') {
      schedules.schedules = schedules.schedules || {};
      schedules.schedules.footings = schedules.schedules.footings || [];
      const lines = String(val).split(/\r?\n|;/).map(l => l.trim()).filter(Boolean);
      for (const line of lines) {
        // F1 2600x1800 900 12  OR  2600x1800 900 12
        const m = line.match(/([A-Z]?\d+[A-Z]?)?\s*(\d{3,4})\s*[xX×]\s*(\d{3,4})\s+(\d{3,4})\s+(\d{1,4})/i)
          || line.match(/([A-Z]?\d+[A-Z]?)\s+(\d{3,4})\s*[xX×]\s*(\d{3,4})\s+(\d{3,4})/i);
        if (!m) continue;
        const mark = (m[1] || `F${schedules.schedules.footings.length + 1}`).toUpperCase();
        const L = m[2]; const B = m[3];
        const depth = Number(m[4]);
        const qty = m[5] != null ? Number(m[5]) : null;
        const existing = schedules.schedules.footings.find(f => f.mark === mark);
        const row = existing || {
          mark,
          pcc_size_mm: `${L}x${B}`,
          rcc_size_mm: `${L}x${B}`,
          depth_mm: depth,
          qty,
          main_bars_x: 'not found in drawing',
          main_bars_y: 'not found in drawing',
          source: 'user-confirmed',
        };
        row.pcc_size_mm = `${L}x${B}`;
        row.rcc_size_mm = `${L}x${B}`;
        row.depth_mm = depth;
        if (qty != null) row.qty = qty;
        row.source = 'user-confirmed';
        if (!existing) schedules.schedules.footings.push(row);
      }
      schedules.total_schedule_rows = Object.values(schedules.schedules).reduce((n, a) => n + (a?.length || 0), 0);
      schedules.quality = schedules.schedules.footings.length ? 'weak' : schedules.quality;
    }
  }

  return { schedules, opts, answers };
}

module.exports = {
  buildClarifications,
  formatQuestionsMarkdown,
  formatOneQuestion,
  mergeUserAnswers,
};
