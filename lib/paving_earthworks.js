'use strict';
/**
 * Product-scale earthworks + paving takeoff (Civils groundworks-style v1).
 * Areas × thicknesses → volumes; cut/fill with bulking; layer BOQ.
 */

const { buildPlanMeasure } = require('./plan_measure');
const { buildEarthworks, formatEarthworksMarkdown, parseLevels } = require('./earthworks');

/** Default layer catalogue (mm) — India / site typical; user can override */
const PAVING_LAYERS = [
  { id: 'asphalt', label: 'Asphalt / bituminous carpet', thickness_mm: 40, unit: 'cum', factor: 1 },
  { id: 'dbm', label: 'DBM', thickness_mm: 50, unit: 'cum', factor: 1 },
  { id: 'wmm', label: 'WMM / WBM', thickness_mm: 150, unit: 'cum', factor: 1 },
  { id: 'gsb', label: 'GSB', thickness_mm: 200, unit: 'cum', factor: 1 },
  { id: 'subgrade', label: 'Subgrade preparation', thickness_mm: 150, unit: 'cum', factor: 1 },
  { id: 'kerb', label: 'Kerb length', thickness_mm: null, unit: 'm', kind: 'length' },
];

const BULKING = 1.25; // cut → loose
const SHRINKAGE = 0.90; // fill compaction

function extractPavingAreas(text, measure) {
  const t = String(text || '');
  const areas = [];
  const patterns = [
    [/(?:road|carriageway|asphalt|bituminous|paving|pavement|carpet)\s*(?:area)?[^\n]{0,40}?(\d+(?:\.\d+)?)\s*(?:sq\.?\s*m|sqm|m²)/gi, 'road_paving'],
    [/(?:footpath|sidewalk|pathway)\s*(?:area)?[^\n]{0,40}?(\d+(?:\.\d+)?)\s*(?:sq\.?\s*m|sqm|m²)/gi, 'footpath'],
    [/(?:parking|hardstand)\s*(?:area)?[^\n]{0,40}?(\d+(?:\.\d+)?)\s*(?:sq\.?\s*m|sqm|m²)/gi, 'parking'],
    [/(?:landscape|green|lawn)\s*(?:area)?[^\n]{0,40}?(\d+(?:\.\d+)?)\s*(?:sq\.?\s*m|sqm|m²)/gi, 'landscape'],
  ];
  for (const [re, kind] of patterns) {
    for (const m of t.matchAll(re)) {
      areas.push({ kind, area_sqm: Number(m[1]), source: 'printed', raw: m[0].slice(0, 60) });
    }
  }
  // From measure items
  for (const it of (measure?.items || []).filter(i => i.type === 'area')) {
    const d = `${it.description} ${it.source}`.toLowerCase();
    let kind = 'general_area';
    if (/road|paving|asphalt|carpet|pavement/.test(d)) kind = 'road_paving';
    else if (/footpath|path/.test(d)) kind = 'footpath';
    else if (/park/.test(d)) kind = 'parking';
    else if (/plot|floor|built/.test(d)) kind = 'plot';
    if (!areas.some(a => a.area_sqm === it.qty && a.kind === kind)) {
      areas.push({ kind, area_sqm: it.qty, source: it.source || 'measure', raw: it.description });
    }
  }
  return areas.slice(0, 30);
}

function extractThicknesses(text) {
  const t = String(text || '');
  const out = {};
  const map = [
    [/asphalt|carpet|bitumin/i, 'asphalt'],
    [/DBM/i, 'dbm'],
    [/WMM|WBM/i, 'wmm'],
    [/GSB/i, 'gsb'],
    [/sub\s*grade|subgrade/i, 'subgrade'],
  ];
  for (const [re, id] of map) {
    const m = t.match(new RegExp(`${re.source}[^\\n]{0,30}?(\\d{2,3})\\s*mm`, 'i'));
    if (m) out[id] = Number(m[1]);
  }
  return out;
}

function buildPavingTakeoff(opts = {}) {
  const text = opts.text || '';
  const measure = opts.measure || buildPlanMeasure({ text, polylines: opts.polylines });
  const areas = extractPavingAreas(text, measure);
  const thkOverride = { ...extractThicknesses(text), ...(opts.thickness_mm || {}) };
  const answers = opts.answers || {};
  const items = [];
  const questions = [];

  const roadAreas = areas.filter(a => /road|paving|parking|footpath/.test(a.kind));
  const primaryArea = answers.paving_area_sqm != null
    ? Number(answers.paving_area_sqm)
    : (roadAreas.sort((a, b) => b.area_sqm - a.area_sqm)[0]?.area_sqm
      || areas.sort((a, b) => b.area_sqm - a.area_sqm)[0]?.area_sqm
      || null);

  if (primaryArea == null) {
    questions.push({
      id: 'paving_area',
      question: 'Paving / road area (sqm)? e.g. `1200`',
      why: 'Need area × layer thickness for paving volumes',
    });
  }

  const layers = PAVING_LAYERS.filter(l => l.kind !== 'length');
  if (primaryArea != null) {
    for (const layer of layers) {
      const thk = thkOverride[layer.id] != null ? thkOverride[layer.id] : layer.thickness_mm;
      if (thk == null) continue;
      const vol = Math.round(primaryArea * (thk / 1000) * 1000) / 1000;
      items.push({
        type: 'paving',
        layer: layer.id,
        description: `${layer.label} (${thk} mm)`,
        qty: vol,
        unit: 'cum',
        area_sqm: primaryArea,
        thickness_mm: thk,
        calc_note: `${primaryArea} sqm x ${thk}/1000 m`,
        confidence: thkOverride[layer.id] != null || extractThicknesses(text)[layer.id] ? 'medium' : 'low',
        source: 'paving-catalogue',
      });
      // also area line for measurement
      items.push({
        type: 'paving_area',
        layer: layer.id,
        description: `${layer.label} area`,
        qty: primaryArea,
        unit: 'sqm',
        calc_note: 'plan area',
        confidence: 'medium',
        source: 'paving-catalogue',
      });
    }
  }

  // Kerb / edge lengths from measure
  for (const L of (measure?.items || []).filter(i => i.type === 'length').slice(0, 8)) {
    items.push({
      type: 'length',
      description: `Kerb / edge — ${L.description}`,
      qty: L.qty,
      unit: 'm',
      calc_note: L.source,
      confidence: L.confidence || 'medium',
      source: 'measure',
    });
  }

  // Deduplicate area lines — keep one area + layer vols
  const pavingVols = items.filter(i => i.type === 'paving');
  const lengths = items.filter(i => i.type === 'length');
  const areaOnce = primaryArea != null ? [{
    type: 'paving_area',
    description: 'Paving / road plan area',
    qty: primaryArea,
    unit: 'sqm',
    calc_note: roadAreas[0]?.raw || 'measure',
    confidence: 'medium',
    source: 'paving',
  }] : [];

  return {
    areas,
    primary_area_sqm: primaryArea,
    thickness_mm: { ...Object.fromEntries(layers.map(l => [l.id, thkOverride[l.id] ?? l.thickness_mm])), ...thkOverride },
    items: [...areaOnce, ...pavingVols, ...lengths],
    questions,
    quality: pavingVols.length ? 'medium' : (questions.length ? 'weak' : 'poor'),
    note: 'Paving takeoff v1 — catalogue thicknesses × plan area. Confirm thk on drawing. Product-scale layers: asphalt/DBM/WMM/GSB/subgrade.',
  };
}

/**
 * Combined groundworks: cut/fill + paving + disposal/borrow.
 */
function buildGroundworksTakeoff(opts = {}) {
  const measure = opts.measure || buildPlanMeasure({
    text: opts.text,
    schedules: opts.schedules,
    polylines: opts.polylines,
  });
  const ew = buildEarthworks({ ...opts, measure });
  const paving = buildPavingTakeoff({ ...opts, measure });
  const items = [];

  for (const i of (ew.items || []).filter(x => x.type === 'cut' || x.type === 'fill')) {
    items.push({ ...i, trade: 'earthworks' });
    if (i.type === 'cut' && i.qty != null) {
      items.push({
        type: 'disposal',
        trade: 'earthworks',
        description: 'Excavated material disposal (bulked)',
        qty: Math.round(i.qty * BULKING * 100) / 100,
        unit: 'cum',
        calc_note: `${i.qty} x bulking ${BULKING}`,
        confidence: 'low',
      });
    }
    if (i.type === 'fill' && i.qty != null) {
      items.push({
        type: 'borrow',
        trade: 'earthworks',
        description: 'Borrow fill (compacted → bank)',
        qty: Math.round((i.qty / SHRINKAGE) * 100) / 100,
        unit: 'cum',
        calc_note: `${i.qty} / shrinkage ${SHRINKAGE}`,
        confidence: 'low',
      });
    }
  }
  for (const i of (paving.items || [])) {
    items.push({ ...i, trade: 'paving' });
  }

  const questions = [...(ew.questions || []), ...(paving.questions || [])];
  // Dedupe question ids
  const seenQ = new Set();
  const uniqQ = questions.filter(q => (seenQ.has(q.id) ? false : (seenQ.add(q.id), true)));

  return {
    earthworks: ew,
    paving,
    measure,
    levels: parseLevels(opts.text || ''),
    items,
    questions: uniqQ,
    bulking: BULKING,
    shrinkage: SHRINKAGE,
    quality: items.some(i => i.qty != null && i.type !== 'note')
      ? ((ew.quality === 'medium' || paving.quality === 'medium') ? 'medium' : 'weak')
      : 'poor',
    note: 'Groundworks product takeoff v1: cut/fill + bulking/borrow + paving layers. Not DEM mesh. Confirm before FINAL.',
  };
}

function formatGroundworksMarkdown(gw) {
  if (!gw) return '';
  const parts = [
    '### Groundworks / paving (product-scale draft)',
    `Quality: **${gw.quality}** · bulking **${gw.bulking}** · shrinkage **${gw.shrinkage}**`,
  ];
  if (gw.paving?.primary_area_sqm != null) {
    parts.push(`- Paving area: **${gw.paving.primary_area_sqm}** sqm`);
    parts.push(`- Thicknesses (mm): ${Object.entries(gw.paving.thickness_mm || {}).map(([k, v]) => `${k}=${v}`).join(', ')}`);
  }
  parts.push('', formatEarthworksMarkdown(gw.earthworks));
  if (gw.items?.length) {
    parts.push('', '### Combined BOQ lines');
    parts.push('| Trade | Type | Description | Qty | Unit | Formula |');
    parts.push('|---|---|---|---:|---|---|');
    for (const i of gw.items) {
      parts.push(`| ${i.trade || '—'} | ${i.type} | ${i.description} | ${i.qty ?? '—'} | ${i.unit} | ${i.calc_note || '—'} |`);
    }
  }
  if (gw.questions?.length) {
    parts.push('', '### Need from you');
    for (const q of gw.questions) parts.push(`- ${q.question}`);
  }
  parts.push('', `_${gw.note}_`);
  return parts.join('\n');
}

module.exports = {
  PAVING_LAYERS,
  buildPavingTakeoff,
  buildGroundworksTakeoff,
  formatGroundworksMarkdown,
  extractPavingAreas,
};
