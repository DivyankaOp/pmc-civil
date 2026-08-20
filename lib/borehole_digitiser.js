'use strict';
/**
 * Handwritten / scanned borehole log digitiser → structured geotech + AGS-ready data.
 * Vision-first when tiles available; OCR text merge + noise-tolerant parse.
 */

const { extractGeotech, formatGeotechMarkdown } = require('./geotech_extract');
const { buildAgs41 } = require('./ags_export');
const { buildGroundModel, formatGroundModelMarkdown } = require('./ground_model');

function parseJson(raw) {
  const t = String(raw || '');
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const body = fence ? fence[1] : t;
  const start = body.indexOf('{');
  const end = body.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  try { return JSON.parse(body.slice(start, end + 1)); } catch (_) { return null; }
}

/** Fix common OCR / handwritten confusions in borehole text */
function denoiseBoreholeText(text) {
  return String(text || '')
    .replace(/\bB[H8]\s*[-–]?\s*(\d)/gi, 'BH-$1')
    .replace(/\bG\.?\s*[lI1]\.?\b/gi, 'GL')
    .replace(/\bW\.?\s*[T7]\.?\b/gi, 'WT')
    .replace(/\bSP[T7]\b/gi, 'SPT')
    .replace(/\bN\s*[=:]\s*[Oo](\d)/g, 'N=$1')
    .replace(/\b([Il])\s*(\d{1,2})\s*m\b/g, '1$2 m')
    .replace(/siity/gi, 'silty')
    .replace(/c1ay/gi, 'clay')
    .replace(/sand\s*y/gi, 'sandy');
}

function normalizeVisionBoreholes(parsed) {
  const holes = parsed?.boreholes || parsed?.holes || [];
  return holes.map((h, i) => {
    const spt = Array.isArray(h.spt_n) ? h.spt_n.map(Number).filter(n => n > 0 && n < 200)
      : (h.avg_spt != null ? [Number(h.avg_spt)] : []);
    const strata = h.strata || h.geology || [];
    const strataList = Array.isArray(strata)
      ? strata.map(s => (typeof s === 'string' ? s : s.desc || s.description || '')).filter(Boolean)
      : [];
    const layers = Array.isArray(strata)
      ? strata.filter(s => typeof s === 'object').map(s => ({
        top_m: Number(s.top_m ?? s.top ?? 0),
        base_m: Number(s.base_m ?? s.base ?? 1.5),
        desc: String(s.desc || s.description || s.geology || 'unknown').toLowerCase(),
      }))
      : strataList.map((d, idx) => ({ top_m: idx * 1.5, base_m: idx * 1.5 + 1.5, desc: d.toLowerCase() }));

    return {
      mark: String(h.mark || h.id || `BH-${i + 1}`).toUpperCase().replace(/\s+/g, ''),
      ground_level_m: h.ground_level_m != null ? Number(h.ground_level_m) : (h.gl != null ? Number(h.gl) : null),
      water_level_m: h.water_level_m != null ? Number(h.water_level_m) : (h.wl != null ? Number(h.wl) : null),
      final_depth_m: h.final_depth_m != null ? Number(h.final_depth_m) : (h.depth_m != null ? Number(h.depth_m) : null),
      spt_n: spt,
      spt_depths: (h.spt_depths || []).map(s => ({
        depth_m: Number(s.depth_m ?? s.depth),
        n: Number(s.n ?? s.N ?? s.value),
      })).filter(s => s.n > 0),
      avg_spt: spt.length ? Math.round(spt.reduce((a, b) => a + b, 0) / spt.length) : (h.avg_spt != null ? Number(h.avg_spt) : null),
      cpt: h.cpt || [],
      easting: h.easting ?? h.E ?? null,
      northing: h.northing ?? h.N ?? null,
      strata: strataList.map(s => s.toLowerCase()),
      strata_layers: layers,
      source: 'handwritten-vision-digitiser',
      confidence: h.confidence || 'medium',
      raw: h.raw || '',
    };
  });
}

function mergeBoreholes(primary, secondary) {
  const by = new Map();
  for (const b of [...(primary || []), ...(secondary || [])]) {
    const id = (b.mark || '').toUpperCase();
    if (!id) continue;
    if (!by.has(id)) {
      by.set(id, { ...b });
      continue;
    }
    const cur = by.get(id);
    by.set(id, {
      ...cur,
      ground_level_m: cur.ground_level_m ?? b.ground_level_m,
      water_level_m: cur.water_level_m ?? b.water_level_m,
      final_depth_m: cur.final_depth_m ?? b.final_depth_m,
      spt_n: (cur.spt_n?.length ? cur.spt_n : b.spt_n) || [],
      spt_depths: (cur.spt_depths?.length ? cur.spt_depths : b.spt_depths) || [],
      avg_spt: cur.avg_spt ?? b.avg_spt,
      easting: cur.easting ?? b.easting,
      northing: cur.northing ?? b.northing,
      strata: [...new Set([...(cur.strata || []), ...(b.strata || [])])],
      strata_layers: (cur.strata_layers?.length ? cur.strata_layers : b.strata_layers) || [],
      source: cur.source?.includes('vision') ? cur.source : (b.source || cur.source),
      confidence: cur.confidence === 'high' ? 'high' : (b.confidence || cur.confidence),
    });
  }
  return [...by.values()];
}

/** Per-hole factory validation (sanity + completeness) */
function validateBorehole(b) {
  const flags = [];
  let score = 0;
  if (b.mark) score += 10;
  if (b.ground_level_m != null && Math.abs(b.ground_level_m) < 5000) score += 20;
  else if (b.ground_level_m != null) flags.push('GL out of range');
  if (b.water_level_m != null) score += 10;
  if (b.final_depth_m != null && b.final_depth_m > 0 && b.final_depth_m < 200) score += 15;
  if (b.spt_n?.length) {
    score += Math.min(25, b.spt_n.length * 5);
    if (b.spt_n.some(n => n < 1 || n > 100)) flags.push('SPT outlier');
  }
  if (b.spt_depths?.length) {
    score += 10;
    for (let i = 1; i < b.spt_depths.length; i++) {
      if (b.spt_depths[i].depth_m < b.spt_depths[i - 1].depth_m) flags.push('SPT depths not increasing');
    }
  }
  if (b.strata_layers?.length) {
    score += Math.min(20, b.strata_layers.length * 5);
    for (const L of b.strata_layers) {
      if (L.base_m < L.top_m) flags.push('strata top/base inverted');
    }
  }
  if (b.easting != null && b.northing != null) score += 10;
  const grade = score >= 70 ? 'high' : score >= 40 ? 'medium' : 'low';
  return {
    ...b,
    factory_score: score,
    factory_grade: grade,
    factory_flags: flags,
    confidence: flags.length ? 'low' : (b.confidence || grade),
  };
}

function scoreFactoryGeotech(holes) {
  if (!holes.length) return { factory_score: 0, factory_grade: 'poor', pass_rate: 0 };
  const scored = holes.map(validateBorehole);
  const avg = Math.round(scored.reduce((s, h) => s + h.factory_score, 0) / scored.length);
  const pass = scored.filter(h => h.factory_grade !== 'low').length;
  return {
    boreholes: scored,
    factory_score: avg,
    factory_grade: avg >= 70 ? 'high' : avg >= 40 ? 'medium' : 'low',
    pass_rate: Math.round((pass / scored.length) * 100),
    flags: scored.flatMap(h => (h.factory_flags || []).map(f => `${h.mark}: ${f}`)),
  };
}

/**
 * Agreement boost when OCR + vision both saw similar GL/SPT.
 */
function applyAgreementBoost(merged, visionHoles, ocrHoles) {
  const vMap = new Map((visionHoles || []).map(h => [h.mark, h]));
  const oMap = new Map((ocrHoles || []).map(h => [h.mark, h]));
  return merged.map(h => {
    const v = vMap.get(h.mark);
    const o = oMap.get(h.mark);
    if (!v || !o) return h;
    let boost = 0;
    if (v.ground_level_m != null && o.ground_level_m != null
      && Math.abs(v.ground_level_m - o.ground_level_m) <= 0.15) boost += 10;
    if (v.avg_spt != null && o.avg_spt != null && Math.abs(v.avg_spt - o.avg_spt) <= 3) boost += 10;
    return {
      ...h,
      confidence: boost >= 10 ? 'high' : h.confidence,
      agreement_boost: boost,
      source: boost ? 'ocr+vision-agree' : h.source,
    };
  });
}

/**
 * @param {object} opts
 * @param {string[]} opts.pngTiles
 * @param {function} opts.callClaudeAPI
 * @param {string} [opts.text]
 */
async function digitiseBoreholes(opts = {}) {
  const denoised = denoiseBoreholeText(opts.text || '');
  const local = extractGeotech(denoised);
  const tiles = (opts.pngTiles || []).filter(Boolean).slice(0, 4);

  let visionHoles = [];
  let mode = 'ocr-text';
  let visionRaw = '';
  let tokens = 0;

  // Claude only when OCR weak AND caller passed tiles + callClaudeAPI
  if (tiles.length && typeof opts.callClaudeAPI === 'function' && process.env.CLAUDE_API_KEY) {
    const content = [
      {
        type: 'text',
        text: `You digitise GEOTECHNICAL BOREHOLE LOGS from scanned/handwritten/messy PDF pages into structured data (Civils.ai-style borehole digitiser).

Read every borehole. Handle handwritten SPT, crossed-out numbers, poor scans carefully — only record what you can reasonably read.

Return ONLY JSON:
{
  "boreholes": [
    {
      "mark": "BH-1",
      "ground_level_m": number|null,
      "water_level_m": number|null,
      "final_depth_m": number|null,
      "easting": number|null,
      "northing": number|null,
      "avg_spt": number|null,
      "spt_n": [numbers],
      "spt_depths": [{"depth_m":1.5,"n":12}],
      "strata": [{"top_m":0,"base_m":2,"desc":"silty clay"}],
      "confidence": "high|medium|low"
    }
  ],
  "meta": {"soil_bearing": null, "notes": "..."}
}

Never invent SPT or levels. If unreadable, use null and lower confidence.`,
      },
      ...tiles.map(data => ({
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data },
      })),
    ];
    // Also attach OCR snippet as hint
    if (denoised.slice(0, 4000)) {
      content.push({
        type: 'text',
        text: `OCR HINT (noisy, may be wrong):\n${denoised.slice(0, 4000)}`,
      });
    }
    try {
      visionRaw = await opts.callClaudeAPI({
        system: 'Geotech borehole digitiser. Never invent. JSON only. Handwritten OK.',
        messages: [{ role: 'user', content }],
        maxTokens: 3500,
      });
      tokens = 1;
      const parsed = parseJson(visionRaw);
      visionHoles = normalizeVisionBoreholes(parsed || {});
      mode = visionHoles.length ? 'handwritten-vision' : 'vision-empty-fallback-ocr';
    } catch (e) {
      mode = `vision-error:${e.message}`;
    }
  }

  let merged = mergeBoreholes(visionHoles, local.boreholes);
  merged = applyAgreementBoost(merged, visionHoles, local.boreholes);
  const factory = scoreFactoryGeotech(merged);

  const geotech = {
    boreholes: factory.boreholes,
    meta: {
      ...(local.meta || {}),
      digitiser_mode: mode,
      factory_grade: factory.factory_grade,
      factory_pass_rate: factory.pass_rate,
    },
    factory_score: factory.factory_score,
    factory_grade: factory.factory_grade,
    factory_pass_rate: factory.pass_rate,
    factory_flags: factory.flags,
    quality: !factory.boreholes.length ? 'poor'
      : (factory.factory_grade === 'high' ? 'medium' : factory.factory_grade === 'medium' ? 'medium' : 'weak'),
    note: 'Handwritten BH factory digitiser v2 — OCR+vision agree / validate / score → AGS. Human QA still required before OpenGround.',
  };

  const groundModel = buildGroundModel(geotech, { field: 'avg_spt' });
  const ags = buildAgs41(geotech, { title: opts.title || 'PMC Borehole Digitiser' });

  const markdown = [
    '# Handwritten / scanned borehole digitiser (factory)',
    `Mode: **${mode}** · boreholes: **${factory.boreholes.length}** · factory score: **${factory.factory_score}/100** (${factory.factory_grade}) · pass **${factory.pass_rate}%**`,
    '',
    formatGeotechMarkdown(geotech),
    '',
    '| Mark | Factory score | Grade | Flags |',
    '|---|---:|---|---|',
    ...factory.boreholes.map(b => `| ${b.mark} | ${b.factory_score} | ${b.factory_grade} | ${(b.factory_flags || []).join('; ') || '—'} |`),
    '',
    formatGroundModelMarkdown(groundModel),
    '',
    '### AGS 4.1 ready',
    `- Groups: ${ags.groups.join(', ')}`,
    `- Download via **AGS 4.1** after Human QA sign-off`,
    '',
    `_${geotech.note}_`,
  ].join('\n');

  return {
    geotech,
    groundModel,
    ags,
    markdown,
    mode,
    tokens,
    factory,
    visionRaw: visionRaw.slice(0, 2000),
  };
}

module.exports = {
  digitiseBoreholes,
  denoiseBoreholeText,
  normalizeVisionBoreholes,
  mergeBoreholes,
  validateBorehole,
  scoreFactoryGeotech,
  applyAgreementBoost,
};
