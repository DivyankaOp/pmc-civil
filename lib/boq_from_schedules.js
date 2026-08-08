'use strict';
/**
 * Build BOQ from schedule rows + Rates.json.
 * Quantities come from printed schedule values only — never invented sizes/qty.
 */

const fs = require('fs');
const { ratesFilePath } = require('../paths');
const { parseSizeMm, estimateSteelKgFromBars } = require('./schedule_extractor');

function loadRates() {
  try {
    return JSON.parse(fs.readFileSync(ratesFilePath(), 'utf8'));
  } catch {
    return {};
  }
}

function rateOf(rates, category, key) {
  const v = rates?.[category]?.[key];
  if (v && typeof v.rate === 'number') return v;
  // flat search
  for (const cat of Object.values(rates || {})) {
    if (!cat || typeof cat !== 'object' || Array.isArray(cat)) continue;
    if (cat[key]?.rate) return cat[key];
  }
  return { rate: 0, unit: '', description: key };
}

function sizeToM(sizeMm) {
  const s = parseSizeMm(sizeMm) || String(sizeMm || '');
  const m = s.match(/(\d+)\s*x\s*(\d+)/i);
  if (!m) return null;
  return { w: Number(m[1]) / 1000, d: Number(m[2]) / 1000 };
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function item(sr, part, description, unit, qty, rateObj, source, calc_note, confidence) {
  const rate = rateObj?.rate || 0;
  const q = Number(qty) || 0;
  return {
    sr,
    part,
    description,
    unit: unit || rateObj?.unit || '',
    qty: round2(q),
    rate,
    amount: round2(q * rate),
    source,
    confidence: confidence || (q > 0 && rate > 0 ? 'high' : 'low'),
    calc_note: calc_note || '',
  };
}

/**
 * @param {object} extracted - from extractSchedules()
 * @param {object} [opts]
 * @param {number} [opts.defaultStoreyHeightM] - ONLY if user confirmed (never silent default)
 * @param {number} [opts.pccThicknessM] - ONLY if printed or user confirmed
 * @param {number} [opts.pccOffsetM] - optional, for reporting only
 * @param {object} [opts.columnHeights] - { mark: heightM } user-confirmed
 * @param {boolean} [opts.allowExcavationSurcharge=false] - if false, skip +0.3m assumption
 */
function buildBoqFromSchedules(extracted, opts = {}) {
  const rates = loadRates();
  const schedules = extracted?.schedules || {};
  const meta = extracted?.meta || {};
  const globalHeightM = opts.defaultStoreyHeightM; // undefined unless user confirmed
  const pccThk = opts.pccThicknessM ?? opts.defaultPccThicknessM; // no silent 0.15
  const notFound = [];
  const boq = [];
  let sr = 1;

  const rccKey = /M30/i.test(meta.concrete_grade || '') ? 'rcc_m30_cum'
    : /M20/i.test(meta.concrete_grade || '') ? 'rcc_m20_cum'
    : 'rcc_m25_cum';
  const rcc = rateOf(rates, 'structure', rccKey);
  const pcc = rateOf(rates, 'structure', 'pcc_m10_cum');
  const steel = rateOf(rates, 'structure', 'steel_fe500_kg');
  const excav = rateOf(rates, 'civil', 'excavation_cum');
  const doorRate = rateOf(rates, 'finishes', 'door_flush_nos');
  const winRate = rateOf(rates, 'finishes', 'window_aluminum_sqmt');

  // ── Columns ──────────────────────────────────────────────────
  let colRcc = 0;
  let colSteel = 0;
  let colCount = 0;
  for (const row of schedules.columns || []) {
    const dims = sizeToM(row.size_mm);
    const qty = row.qty;
    if (!dims) {
      notFound.push(`Column ${row.mark}: size not found in drawing — ASK USER`);
      continue;
    }
    if (qty == null || qty <= 0) {
      notFound.push(`Column ${row.mark}: qty not found in schedule — ASK USER (not counted from plan)`);
      continue;
    }
    const heightM = opts.columnHeights?.[row.mark] ?? row.height_m ?? globalHeightM;
    if (!heightM) {
      notFound.push(`Column ${row.mark}: height not on drawing — ASK USER (no 3m assumption)`);
      continue;
    }
    const vol = dims.w * dims.d * heightM * qty;
    colRcc += vol;
    colCount += qty;
    const hNote = `size ${row.size_mm} × H ${heightM}m (confirmed) × qty ${qty}`;
    boq.push(item(sr++, 'PART A — STRUCTURE',
      `RCC Columns ${row.mark} (${row.size_mm})`, 'cum', vol, rcc,
      row.source === 'user-confirmed' ? 'user-confirmed' : 'calculated-from-schedule', hNote, 'high'));

    const skg = estimateSteelKgFromBars(row.main_bars, heightM, qty);
    if (skg > 0) {
      colSteel += skg;
      boq.push(item(sr++, 'PART A — STRUCTURE',
        `Steel in columns ${row.mark} (${row.main_bars})`, 'kg', skg, steel,
        'calculated-from-schedule', `from bar schedule ${row.main_bars}`, 'medium'));
    } else if (row.main_bars && /not found/i.test(row.main_bars) === false) {
      notFound.push(`Column ${row.mark}: bar spec present but could not convert to kg — ASK USER`);
    }
  }

  // ── Footings ─────────────────────────────────────────────────
  let ftgRcc = 0;
  let ftgPcc = 0;
  let ftgSteel = 0;
  for (const row of schedules.footings || []) {
    const rccDims = sizeToM(row.rcc_size_mm) || sizeToM(row.pcc_size_mm);
    const pccDims = sizeToM(row.pcc_size_mm) || rccDims;
    const qty = row.qty;
    if (!rccDims) {
      notFound.push(`Footing ${row.mark}: size not found in drawing`);
      continue;
    }
    if (qty == null || qty <= 0) {
      notFound.push(`Footing ${row.mark}: qty not found in schedule — skipped`);
      continue;
    }
    const depthM = row.depth_mm ? row.depth_mm / 1000 : null;
    if (!depthM) {
      notFound.push(`Footing ${row.mark}: depth not found — ASK USER (RCC not calculated)`);
    } else {
      const vol = rccDims.w * rccDims.d * depthM * qty;
      ftgRcc += vol;
      boq.push(item(sr++, 'PART A — STRUCTURE',
        `RCC Footing ${row.mark} (${row.rcc_size_mm || row.pcc_size_mm})`, 'cum', vol, rcc,
        'calculated-from-schedule',
        `${row.rcc_size_mm || row.pcc_size_mm} × depth ${row.depth_mm}mm × qty ${qty}`, 'high'));

      if (opts.allowExcavationSurcharge === true) {
        const exVol = rccDims.w * rccDims.d * (depthM + 0.3) * qty;
        boq.push(item(sr++, 'PART A — STRUCTURE',
          `Excavation for footing ${row.mark}`, 'cum', exVol, excav,
          'calculated-from-schedule', 'plan area × (depth+0.3m) — user allowed surcharge', 'medium'));
      } else {
        notFound.push(`Footing ${row.mark}: excavation surcharge not on drawing — skipped (ask if needed)`);
      }
    }

    if (pccDims && pccThk) {
      const pVol = pccDims.w * pccDims.d * pccThk * qty;
      ftgPcc += pVol;
      boq.push(item(sr++, 'PART A — STRUCTURE',
        `PCC under footing ${row.mark}`, 'cum', pVol, pcc,
        'calculated-from-schedule', `PCC thk ${Math.round(pccThk * 1000)}mm (confirmed)`, 'high'));
    } else if (pccDims && !pccThk) {
      notFound.push(`Footing ${row.mark}: PCC thickness not on drawing — ASK USER (no 100/150mm assumption)`);
    }

    const span = Math.max(rccDims.w, rccDims.d);
    const skg = estimateSteelKgFromBars(row.main_bars_x, span, qty)
      + estimateSteelKgFromBars(row.main_bars_y, span, qty);
    if (skg > 0) {
      ftgSteel += skg;
      boq.push(item(sr++, 'PART A — STRUCTURE',
        `Steel in footing ${row.mark}`, 'kg', skg, steel,
        'calculated-from-schedule', 'from footing bar schedule', 'medium'));
    }
  }

  // ── Doors / Windows ──────────────────────────────────────────
  for (const row of schedules.doors || []) {
    if (row.qty == null || row.qty <= 0) {
      notFound.push(`Door ${row.mark}: qty not found`);
      continue;
    }
    boq.push(item(sr++, 'PART B — FINISHES',
      `Door ${row.mark} (${row.size_mm})`, 'nos', row.qty, doorRate,
      'drawing-schedule', 'qty from door schedule', 'high'));
  }

  for (const row of schedules.windows || []) {
    if (row.qty == null || row.qty <= 0) {
      notFound.push(`Window ${row.mark}: qty not found`);
      continue;
    }
    const dims = sizeToM(row.size_mm);
    const area = dims ? dims.w * dims.d * row.qty : row.qty;
    const unit = dims ? 'sqmt' : 'nos';
    boq.push(item(sr++, 'PART B — FINISHES',
      `Window ${row.mark} (${row.size_mm})`, unit, area, dims ? winRate : doorRate,
      'drawing-schedule', dims ? 'area from schedule size × qty' : 'qty only (size missing)', dims ? 'high' : 'low'));
  }

  const totalInr = boq.reduce((s, i) => s + (i.amount || 0), 0);

  return {
    project_name: meta.project_name || 'CIVIL PROJECT',
    drawing_no: meta.drawing_no || '',
    drawing_type: schedules.footings?.length ? 'FOUNDATION' : schedules.columns?.length ? 'STRUCTURAL' : 'GENERAL',
    scale: meta.scale || '',
    concrete_grade: meta.concrete_grade || '',
    steel_grade: meta.steel_grade || '',
    schedule_data: {
      columns: schedules.columns || [],
      footings: schedules.footings || [],
      doors: schedules.doors || [],
      windows: schedules.windows || [],
      beams: schedules.beams || [],
      base_plates: schedules.base_plates || [],
    },
    element_counts: {
      column_count: colCount,
      footing_count: (schedules.footings || []).reduce((s, r) => s + (r.qty || 0), 0),
      door_count: (schedules.doors || []).reduce((s, r) => s + (r.qty || 0), 0),
      window_count: (schedules.windows || []).reduce((s, r) => s + (r.qty || 0), 0),
    },
    boq,
    boq_items: boq,
    cost_summary: {
      civil_total_inr: round2(totalInr),
      civil_total_lacs: round2(totalInr / 100000),
      civil_total_crores: round2(totalInr / 10000000),
      item_wise: boq,
    },
    total_quantities: {
      rcc_total_cum: round2(colRcc + ftgRcc),
      footing_rcc_cum: round2(ftgRcc),
      pcc_total_cum: round2(ftgPcc),
      steel_total_kg: round2(colSteel + ftgSteel),
      calculation_note: 'Schedule-first BOQ — quantities from printed schedule tables + Rates.json',
    },
    observations: [
      'BOQ from printed schedule values only — missing cells were NOT assumed.',
      'Qty only from schedule QTY / user confirmation — plan symbols were not counted.',
      notFound.length ? `${notFound.length} item(s) need user confirmation before final totals.` : 'All used values were found or user-confirmed.',
    ].filter(Boolean),
    not_found: notFound,
    not_legible_fields: notFound,
    overall_confidence: boq.length ? (notFound.length > boq.length ? 'MEDIUM' : 'HIGH') : 'LOW',
    prepared_by: 'PMC Civil AI — Schedule-first Pipeline',
    pipeline_info: {
      mode: 'schedule-first',
      claude_mode: 'optional polish only',
      schedule_rows: extracted?.total_schedule_rows || 0,
      quality: extracted?.quality || 'poor',
    },
  };
}

function formatBoqMarkdown(boqResult) {
  const parts = [];
  parts.push('## Bill of Quantities (from schedules)');
  if (!boqResult?.boq?.length) {
    parts.push('_No BOQ lines could be built — schedule qty/size missing._');
    if (boqResult?.not_found?.length) {
      parts.push('\n### Missing from drawing');
      for (const n of boqResult.not_found) parts.push(`- ${n}`);
    }
    return parts.join('\n');
  }
  parts.push('| Sr | Part | Description | Unit | Qty | Rate | Amount | Source |');
  parts.push('|---:|---|---|---|---:|---:|---:|---|');
  for (const i of boqResult.boq) {
    parts.push(`| ${i.sr} | ${i.part} | ${i.description} | ${i.unit} | ${i.qty} | ${i.rate} | ${i.amount} | ${i.source} |`);
  }
  const tot = boqResult.cost_summary?.civil_total_inr || 0;
  parts.push(`\n**Civil total: ₹${tot.toLocaleString('en-IN')}** (${boqResult.cost_summary?.civil_total_lacs || 0} Lacs)`);
  if (boqResult.not_found?.length) {
    parts.push('\n### Not found / skipped');
    for (const n of boqResult.not_found) parts.push(`- ${n}`);
  }
  parts.push('\n_Rates from data/Rates.json (Gujarat DSR). Edit rates file to update — no code change._');
  return parts.join('\n');
}

module.exports = {
  buildBoqFromSchedules,
  formatBoqMarkdown,
  loadRates,
};
