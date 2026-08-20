'use strict';
/**
 * Form-based multi-trade takeoff agents (Civils.ai-style scopes).
 */

const { buildPlanMeasure, formatPlanMeasureMarkdown } = require('./plan_measure');
const { extractGeotech, formatGeotechMarkdown } = require('./geotech_extract');
const { buildBoqFromSchedules, formatBoqMarkdown } = require('./boq_from_schedules');
const { formatSchedulesMarkdown } = require('./schedule_extractor');
const { buildEarthworks, formatEarthworksMarkdown } = require('./earthworks');
const { buildGroundworksTakeoff, formatGroundworksMarkdown } = require('./paving_earthworks');
const { buildGroundModel, formatGroundModelMarkdown } = require('./ground_model');

const AGENTS = [
  { id: 'concrete', label: 'Concrete / structure', items: ['footings', 'columns', 'pcc', 'excavation'] },
  { id: 'openings', label: 'Doors & windows', items: ['doors', 'windows'] },
  { id: 'plan_areas', label: 'Plan areas & lengths', items: ['areas', 'lengths', 'counts'] },
  { id: 'vision_takeoff', label: 'Auto vision takeoff', items: ['areas', 'lengths', 'counts', 'vision'] },
  { id: 'earthworks', label: 'Earthworks / cut-fill', items: ['cut', 'fill', 'excavation', 'areas', 'paving'] },
  { id: 'paving', label: 'Paving / road layers', items: ['paving', 'areas', 'lengths'] },
  { id: 'groundworks', label: 'Groundworks (cut/fill + paving)', items: ['cut', 'fill', 'paving', 'areas', 'excavation'] },
  { id: 'geotech', label: 'Geotech / boreholes', items: ['boreholes', 'spt', 'water'] },
  { id: 'borehole_digitise', label: 'Handwritten borehole AGS', items: ['boreholes', 'spt', 'water', 'handwritten'] },
  { id: 'full', label: 'Full sheet takeoff', items: ['footings', 'columns', 'doors', 'windows', 'areas', 'pcc'] },
];

function buildQuestionFromScope(scope = {}) {
  const agent = scope.agent || 'full';
  const items = scope.items || (AGENTS.find(a => a.id === agent)?.items || []);
  const extra = (scope.instructions || '').trim();
  if (agent === 'vision_takeoff' || items.includes('vision')) {
    return `Auto vision takeoff — measure areas, lengths and counts from the drawing without clicking. ${extra}`.trim();
  }
  if (agent === 'borehole_digitise' || items.includes('handwritten')) {
    return `Digitise handwritten or scanned borehole logs into AGS (BH, GL, WT, SPT, strata). ${extra}`.trim();
  }
  if (agent === 'geotech' || items.includes('boreholes')) {
    return `Extract geotech borehole data (BH marks, GL, water table, SPT, strata). ${extra}`.trim();
  }
  if (agent === 'groundworks' || (items.includes('paving') && (items.includes('cut') || items.includes('fill')))) {
    return `Groundworks takeoff — cut/fill with bulking/borrow plus paving layer volumes (asphalt/DBM/WMM/GSB). ${extra}`.trim();
  }
  if (agent === 'paving' || (items.includes('paving') && !items.includes('cut') && !items.includes('fill'))) {
    return `Paving / road layer takeoff — areas × asphalt/DBM/WMM/GSB thicknesses to volumes. ${extra}`.trim();
  }
  if (agent === 'earthworks' || items.includes('cut') || items.includes('fill')) {
    return `Calculate earthworks cut and fill from levels and plan areas (include paving layers if present). ${extra}`.trim();
  }
  if (agent === 'plan_areas' || (items.includes('areas') && !items.includes('footings') && !items.includes('cut') && !items.includes('vision') && !items.includes('paving'))) {
    return `Measure plan areas, lengths and counts from this drawing. ${extra}`.trim();
  }
  if (items.includes('footings') && items.includes('pcc') && !items.includes('doors')) {
    return `Calculate PCC and RCC of footing from schedule. Show volumes. ${extra}`.trim();
  }
  if (items.includes('doors') && !items.includes('footings')) {
    return `Prepare quantity takeoff for doors and windows from schedules. ${extra}`.trim();
  }
  return `Prepare quantity takeoff from this drawing for: ${items.join(', ') || 'all trades'}. ${extra}`.trim();
}

function filterExtractedForScope(extracted, scope = {}) {
  const items = new Set(scope.items || []);
  if (!items.size || scope.agent === 'full') return extracted;
  const sch = extracted?.schedules || {};
  const out = {
    ...extracted,
    schedules: {
      ...sch,
      footings: items.has('footings') || items.has('pcc') || items.has('excavation') ? (sch.footings || []) : [],
      columns: items.has('columns') ? (sch.columns || []) : [],
      doors: items.has('doors') ? (sch.doors || []) : [],
      windows: items.has('windows') ? (sch.windows || []) : [],
      beams: items.has('beams') ? (sch.beams || []) : [],
    },
  };
  const total = Object.values(out.schedules).reduce((n, a) => n + (Array.isArray(a) ? a.length : 0), 0);
  out.total_schedule_rows = total;
  return out;
}

function runTradeTakeoff({ text, extracted, scope = {}, boqOpts = {}, polylines = [] }) {
  const agent = scope.agent || 'full';
  const items = scope.items || AGENTS.find(a => a.id === agent)?.items || [];
  const parts = [];
  parts.push(`# Trade takeoff — **${AGENTS.find(a => a.id === agent)?.label || agent}**`);
  parts.push(`Scope items: ${items.join(', ') || '—'}`);
  if (scope.instructions) parts.push(`Instructions: ${scope.instructions}`);
  parts.push('');

  let geotech = null;
  let measure = null;
  let earthworks = null;
  let paving = null;
  let groundworks = null;
  let groundModel = null;
  let boqResult = null;
  let scoped = extracted;

  if (agent === 'geotech' || items.includes('boreholes') || items.includes('spt')) {
    geotech = extractGeotech(text);
    parts.push(formatGeotechMarkdown(geotech));
    groundModel = buildGroundModel(geotech, { field: 'avg_spt' });
    parts.push('', formatGroundModelMarkdown(groundModel));
  }

  const wantsGroundworks = agent === 'groundworks' || agent === 'paving'
    || items.includes('paving')
    || (agent === 'earthworks' && items.includes('paving'));

  if (agent === 'plan_areas' || agent === 'earthworks' || agent === 'paving' || agent === 'groundworks'
    || items.includes('areas') || items.includes('lengths') || items.includes('counts')
    || items.includes('cut') || items.includes('fill') || items.includes('paving')) {
    measure = buildPlanMeasure({ text, schedules: extracted, polylines });
    if (!wantsGroundworks || items.includes('areas') || agent === 'plan_areas') {
      parts.push(formatPlanMeasureMarkdown(measure));
    }
  }

  if (wantsGroundworks) {
    groundworks = buildGroundworksTakeoff({ text, measure, schedules: extracted, polylines });
    earthworks = groundworks.earthworks;
    paving = groundworks.paving;
    parts.push(formatGroundworksMarkdown(groundworks));
  } else if (agent === 'earthworks' || items.includes('cut') || items.includes('fill')) {
    earthworks = buildEarthworks({ text, measure, schedules: extracted, polylines });
    parts.push(formatEarthworksMarkdown(earthworks));
  }

  const wantsStructure = items.some(i => ['footings', 'columns', 'pcc', 'excavation', 'doors', 'windows', 'beams'].includes(i))
    || agent === 'concrete' || agent === 'openings' || agent === 'full';

  if (wantsStructure && agent !== 'geotech' && agent !== 'earthworks' && agent !== 'paving' && agent !== 'groundworks') {
    scoped = filterExtractedForScope(extracted, { agent, items });
    parts.push(formatSchedulesMarkdown(scoped));
    boqResult = buildBoqFromSchedules(scoped, boqOpts);
    if (agent === 'openings') {
      boqResult.boq = (boqResult.boq || []).filter(i => /door|window/i.test(i.description));
    }
    if (agent === 'concrete') {
      boqResult.boq = (boqResult.boq || []).filter(i => /footing|column|pcc|excavation|rcc/i.test(i.description));
    }
    parts.push('', formatBoqMarkdown(boqResult, { status: 'DRAFT' }));
  }

  parts.push('', '_Draft trade takeoff — confirm missing qty before FINAL. Engineer check required._');

  return {
    markdown: parts.join('\n'),
    geotech,
    measure,
    earthworks,
    paving,
    groundworks,
    groundModel,
    boqResult,
    extracted: scoped || extracted,
    agent,
    items,
  };
}

module.exports = {
  AGENTS,
  buildQuestionFromScope,
  filterExtractedForScope,
  runTradeTakeoff,
};
