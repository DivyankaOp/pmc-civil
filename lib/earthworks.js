'use strict';
/**
 * Earthworks / cut-fill draft from printed levels + plan areas (text v1, not DEM).
 */

const { buildPlanMeasure } = require('./plan_measure');

function parseLevels(text) {
  const t = String(text || '');
  const levels = [];
  const re = /(?:(?:NGL|EGL|FGL|FL|FFL|PLINTH|FORMATION|RL|TOS|BOS)\s*[:=]?\s*)([+\-]?\d+(?:\.\d+)?)\s*(?:M|MT|m)?/gi;
  let m;
  while ((m = re.exec(t))) {
    levels.push({
      label: m[0].replace(/\s+/g, ' ').trim().slice(0, 40),
      value_m: Number(m[1]),
    });
  }
  // Named pairs
  const ngl = t.match(/(?:N\.?G\.?L|natural\s*ground)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)/i);
  const fgl = t.match(/(?:F\.?G\.?L|finished\s*ground|formation)\s*[:=]?\s*([+\-]?\d+(?:\.\d+)?)/i);
  return {
    levels: levels.slice(0, 40),
    ngl_m: ngl ? Number(ngl[1]) : null,
    formation_m: fgl ? Number(fgl[1]) : null,
  };
}

function pickSiteAreaSqm(measure) {
  const areas = (measure?.items || []).filter(i => i.type === 'area');
  if (!areas.length) return null;
  // Prefer largest printed / plot area
  const sorted = [...areas].sort((a, b) => b.qty - a.qty);
  return sorted[0];
}

/**
 * Build earthworks draft.
 * @param {object} opts
 * @param {string} opts.text
 * @param {object} [opts.measure]
 * @param {object} [opts.answers] { ngl_m, formation_m, area_sqm, cut_depth_m, fill_depth_m }
 */
function buildEarthworks(opts = {}) {
  const text = opts.text || '';
  const parsed = parseLevels(text);
  const measure = opts.measure || buildPlanMeasure({ text, schedules: opts.schedules, polylines: opts.polylines });
  const answers = opts.answers || {};

  const ngl = answers.ngl_m != null ? Number(answers.ngl_m) : parsed.ngl_m;
  const formation = answers.formation_m != null ? Number(answers.formation_m) : parsed.formation_m;
  const areaItem = pickSiteAreaSqm(measure);
  const area = answers.area_sqm != null ? Number(answers.area_sqm) : (areaItem?.qty ?? null);

  const items = [];
  const questions = [];
  let delta = null;
  if (ngl != null && formation != null) {
    delta = Math.round((formation - ngl) * 1000) / 1000; // +fill / -cut relative to NGL→formation
  }

  if (area == null) {
    questions.push({
      id: 'ew_area',
      question: 'Site / plot area (sqm) for earthworks? e.g. `450`',
      why: 'Need area × depth for cut/fill volume',
    });
  }
  if (ngl == null) {
    questions.push({
      id: 'ew_ngl',
      question: 'Natural ground level NGL (m)? e.g. `12.50`',
      why: 'Cut/fill needs NGL vs formation',
    });
  }
  if (formation == null) {
    questions.push({
      id: 'ew_formation',
      question: 'Formation / finished ground level FGL (m)? e.g. `12.80`',
      why: 'Cut/fill needs formation level',
    });
  }

  if (area != null && delta != null) {
    const depth = Math.abs(delta);
    const vol = Math.round(area * depth * 100) / 100;
    if (delta < 0) {
      items.push({
        type: 'cut',
        description: 'Cut to formation (NGL -> FGL)',
        qty: vol,
        unit: 'cum',
        calc_note: `${area} sqm x ${depth} m cut`,
        confidence: 'medium',
      });
    } else if (delta > 0) {
      items.push({
        type: 'fill',
        description: 'Fill to formation (NGL -> FGL)',
        qty: vol,
        unit: 'cum',
        calc_note: `${area} sqm x ${depth} m fill`,
        confidence: 'medium',
      });
    } else {
      items.push({
        type: 'note',
        description: 'NGL ≈ FGL — negligible cut/fill',
        qty: 0,
        unit: 'cum',
        calc_note: 'delta 0',
        confidence: 'medium',
      });
    }
  } else if (area != null && answers.cut_depth_m != null) {
    const d = Number(answers.cut_depth_m);
    items.push({
      type: 'cut',
      description: 'Cut (user depth)',
      qty: Math.round(area * d * 100) / 100,
      unit: 'cum',
      calc_note: `${area} × ${d}`,
      confidence: 'medium',
    });
  } else if (area != null && answers.fill_depth_m != null) {
    const d = Number(answers.fill_depth_m);
    items.push({
      type: 'fill',
      description: 'Fill (user depth)',
      qty: Math.round(area * d * 100) / 100,
      unit: 'cum',
      calc_note: `${area} × ${d}`,
      confidence: 'medium',
    });
  }

  // Excavation hint from footing schedules if present
  const footings = opts.schedules?.schedules?.footings || [];
  if (footings.length && items.every(i => i.type !== 'excavation')) {
    // Only note — full footing excav is in concrete BOQ
    items.push({
      type: 'note',
      description: `Footing excavation — use Concrete agent (${footings.length} types)`,
      qty: null,
      unit: '—',
      calc_note: 'see footing takeoff',
      confidence: 'low',
    });
  }

  return {
    levels: parsed,
    ngl_m: ngl,
    formation_m: formation,
    area_sqm: area,
    delta_m: delta,
    items,
    questions,
    measure_quality: measure.quality,
    quality: items.some(i => i.type === 'cut' || i.type === 'fill') ? 'medium' : (questions.length ? 'weak' : 'poor'),
    note: 'Earthworks v1 from printed levels + area. Not a DEM / mesh cut-fill. Confirm levels before FINAL.',
  };
}

function formatEarthworksMarkdown(ew) {
  if (!ew) return '';
  const parts = ['### Earthworks / cut-fill (draft)', `Quality: **${ew.quality}**`];
  parts.push(`- NGL: **${ew.ngl_m ?? '—'}** m · Formation: **${ew.formation_m ?? '—'}** m · Δ: **${ew.delta_m ?? '—'}** m`);
  parts.push(`- Area used: **${ew.area_sqm ?? '—'}** sqm`);
  if (ew.items?.length) {
    parts.push('');
    parts.push('| Type | Description | Qty | Unit | Formula |');
    parts.push('|---|---|---:|---|---|');
    for (const i of ew.items) {
      parts.push(`| ${i.type} | ${i.description} | ${i.qty ?? '—'} | ${i.unit} | ${i.calc_note || '—'} |`);
    }
  }
  if (ew.questions?.length) {
    parts.push('', '### Need from you');
    for (const q of ew.questions) parts.push(`- ${q.question}`);
  }
  parts.push('', `_${ew.note}_`);
  return parts.join('\n');
}

module.exports = {
  parseLevels,
  buildEarthworks,
  formatEarthworksMarkdown,
};
