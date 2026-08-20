'use strict';
/**
 * Auto vision takeoff — measure areas/lengths/counts from drawing page images (no click).
 * Uses Claude vision when API key set; otherwise merges local plan_measure text heuristics.
 */

const { buildPlanMeasure, formatPlanMeasureMarkdown } = require('./plan_measure');

function parseVisionJson(raw) {
  const t = String(raw || '');
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const body = fence ? fence[1] : t;
  const start = body.indexOf('{');
  const end = body.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  try {
    return JSON.parse(body.slice(start, end + 1));
  } catch (_) {
    return null;
  }
}

function normalizeVisionItems(parsed) {
  const items = [];
  const list = parsed?.items || parsed?.quantities || [];
  for (const it of list) {
    const type = String(it.type || it.kind || 'area').toLowerCase();
    const qty = Number(it.qty ?? it.quantity ?? it.value);
    if (!Number.isFinite(qty) || qty <= 0) continue;
    items.push({
      type: /length|perimeter|kerb|wall/.test(type) ? 'length'
        : /count|nos|ea/.test(type) ? 'count' : 'area',
      qty: Math.round(qty * 1000) / 1000,
      unit: it.unit || (/length/.test(type) ? 'm' : /count/.test(type) ? 'nos' : 'sqm'),
      description: it.description || it.label || it.name || `Vision ${type}`,
      source: 'auto-vision',
      confidence: it.confidence || 'medium',
      editable: true,
      notes: it.notes || it.evidence || '',
    });
  }
  return items.slice(0, 40);
}

/**
 * @param {object} opts
 * @param {string[]} opts.pngTiles base64 PNG tiles (capped)
 * @param {function} opts.callClaudeAPI
 * @param {string} [opts.text] OCR/text fallback
 * @param {string} [opts.instructions]
 * @param {object} [opts.schedules]
 */
async function runAutoVisionTakeoff(opts = {}) {
  const tiles = (opts.pngTiles || []).filter(Boolean).slice(0, 4);
  const local = buildPlanMeasure({
    text: opts.text || '',
    schedules: opts.schedules,
    polylines: opts.polylines,
  });

  let visionItems = [];
  let visionRaw = '';
  let mode = 'local-text';
  let tokens = 0;

  // Claude only when caller passes callClaudeAPI + tiles (explicit vision agent)
  if (tiles.length && typeof opts.callClaudeAPI === 'function' && process.env.CLAUDE_API_KEY) {
    const content = [
      {
        type: 'text',
        text: `You are a civil quantity surveyor doing AUTO vision takeoff (no human clicks).

From the drawing image(s), extract measurable quantities ONLY if you can see them printed or clearly hatch/boundary labeled.
Scope: ${opts.instructions || 'areas, lengths, counts (roads, paving, plot, walls, manholes, trees, etc.)'}

RULES:
1. NEVER invent dimensions not visible or printed.
2. Prefer printed areas/lengths (sqm, m) over guessing pixel measure.
3. If scale is visible (e.g. 1:100), you may estimate closed areas from geometry — mark confidence "low".
4. Return ONLY JSON:
{
  "scale": "1:100 or null",
  "items": [
    {"type":"area|length|count","qty":number,"unit":"sqm|m|nos","description":"...","confidence":"high|medium|low","notes":"what you saw"}
  ],
  "not_found": ["..."]
}`,
      },
      ...tiles.map(data => ({
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data },
      })),
    ];
    try {
      visionRaw = await opts.callClaudeAPI({
        system: 'Civil QS vision takeoff. Never invent quantities. JSON only.',
        messages: [{ role: 'user', content }],
        maxTokens: 2500,
      });
      tokens = 1; // flag that vision was used
      const parsed = parseVisionJson(visionRaw);
      visionItems = normalizeVisionItems(parsed || {});
      mode = visionItems.length ? 'auto-vision' : 'vision-empty-fallback-local';
    } catch (e) {
      mode = `vision-error:${e.message}`;
    }
  }

  // Merge: vision first, then local items not overlapping descriptions
  const merged = [...visionItems];
  const seen = new Set(merged.map(i => `${i.type}|${i.description}|${i.qty}`));
  for (const i of (local.items || [])) {
    const k = `${i.type}|${i.description}|${i.qty}`;
    if (!seen.has(k)) {
      merged.push(i);
      seen.add(k);
    }
  }

  const measure = {
    ...local,
    items: merged,
    vision_items: visionItems,
    quality: merged.length >= 3 ? 'medium' : merged.length ? 'weak' : 'poor',
    note: mode.startsWith('auto-vision')
      ? 'Auto vision takeoff v1 — no click. Verify before FINAL. Not Civils 97% QA-reviewed accuracy.'
      : `Auto vision unavailable/weak (${mode}) — local text/DXF measure used.`,
    auto_vision_mode: mode,
  };

  const parts = [
    '# Auto vision takeoff (no clicking)',
    `Mode: **${mode}** · items: **${merged.length}** (vision ${visionItems.length} + local)`,
    '',
    formatPlanMeasureMarkdown(measure),
  ];
  if (opts.instructions) parts.splice(2, 0, `Instructions: ${opts.instructions}`);

  return {
    measure,
    markdown: parts.join('\n'),
    visionItems,
    mode,
    tokens,
    visionRaw: visionRaw.slice(0, 2000),
  };
}

module.exports = {
  runAutoVisionTakeoff,
  parseVisionJson,
  normalizeVisionItems,
};
