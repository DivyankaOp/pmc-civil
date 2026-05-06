'use strict';
/**
 * PMC Civil — Progressive Rate Store (rate_store.js)
 * ═══════════════════════════════════════════════════
 * Problem: Rates.json has static Gujarat DSR 2025 data.
 * When Claude extracts BOQ from a drawing, the rates used/found
 * should be persisted so future analyses get better rates.
 *
 * How it works:
 *  1. After every BOQ analysis, call learnRatesFromBOQ(boqItems, context)
 *  2. Rates are saved to rates-learned.json (never overwrites Rates.json)
 *  3. When building rate summary for Claude prompts, call getRatesSummary()
 *     — merges Rates.json (base) + rates-learned.json (learned) with learned taking priority
 *  4. Call getLearnedRateStats() to see what has been learned
 *
 * Rate key format: "description_unit" e.g. "RCC M25_cum", "Steel Fe500_kg"
 * Each entry stores: rate, unit, count (how many times seen), last_seen, source
 */

const fs = require('fs');
const path = require('path');

const BASE_RATES_FILE  = path.join(__dirname, 'Rates.json');
const LEARNED_FILE     = path.join(__dirname, 'rates-learned.json');

// ── Load base rates from Rates.json ──────────────────────────────
function loadBaseRates() {
  try {
    const raw = JSON.parse(fs.readFileSync(BASE_RATES_FILE, 'utf8'));
    const flat = {};
    for (const [cat, items] of Object.entries(raw)) {
      if (cat.startsWith('_') || typeof items !== 'object') continue;
      for (const [key, v] of Object.entries(items)) {
        if (v?.rate) {
          flat[key] = { rate: v.rate, unit: v.unit, description: v.description, source: 'dsr_2025' };
        }
      }
    }
    return flat;
  } catch (e) {
    console.warn('[rate_store] Could not load Rates.json:', e.message);
    return {};
  }
}

// ── Load learned rates ────────────────────────────────────────────
function loadLearnedRates() {
  try {
    if (!fs.existsSync(LEARNED_FILE)) return { _meta: { total_analyses: 0, last_updated: null }, rates: {} };
    return JSON.parse(fs.readFileSync(LEARNED_FILE, 'utf8'));
  } catch (e) {
    return { _meta: { total_analyses: 0, last_updated: null }, rates: {} };
  }
}

// ── Save learned rates ────────────────────────────────────────────
function saveLearnedRates(data) {
  try {
    fs.writeFileSync(LEARNED_FILE, JSON.stringify(data, null, 2), 'utf8');
  } catch (e) {
    console.warn('[rate_store] Could not save rates-learned.json:', e.message);
  }
}

// ── Normalize description to a lookup key ─────────────────────────
function makeKey(description, unit) {
  if (!description) return null;
  const d = description.toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, '_')
    .trim()
    .slice(0, 60);
  const u = (unit || 'nos').toLowerCase().trim();
  return `${d}__${u}`;
}

// ── Learn rates from a BOQ array ─────────────────────────────────
/**
 * @param {Array} boqItems — [{description, unit, qty, rate, amount, source}]
 * @param {Object} context — {filename, drawing_type, project_name, state}
 */
function learnRatesFromBOQ(boqItems, context = {}) {
  if (!Array.isArray(boqItems) || boqItems.length === 0) return;

  const learned = loadLearnedRates();
  learned._meta.total_analyses = (learned._meta.total_analyses || 0) + 1;
  learned._meta.last_updated = new Date().toISOString();
  if (!learned.rates) learned.rates = {};

  let newCount = 0;
  let updateCount = 0;

  for (const item of boqItems) {
    const { description, unit, rate, amount, qty, source } = item;

    // Skip if no rate or rate is 0 or item was "assumed" without real data
    if (!rate || rate <= 0 || !description) continue;

    // Sanity check — skip absurd rates (could be total amounts mistaken for rates)
    if (rate > 10000000) continue; // skip >1 crore per unit (obviously wrong)

    // OUTLIER GUARD: If existing rate known, reject if new rate is outside 0.5x–2.0x band
    const existingCheck = learned.rates[makeKey(description, unit)];
    if (existingCheck?.rate) {
      const ratio = rate / existingCheck.rate;
      if (ratio < 0.5 || ratio > 2.0) continue; // likely typing error, skip
    }

    const key = makeKey(description, unit);
    if (!key) continue;

    const existing = learned.rates[key];

    if (!existing) {
      // New rate discovered
      learned.rates[key] = {
        description: description.trim(),
        unit: unit || 'nos',
        rate,
        count: 1,
        first_seen: new Date().toISOString(),
        last_seen: new Date().toISOString(),
        seen_in: [context.filename || 'unknown'].slice(0, 10),
        source: source || 'claude_boq',
        drawing_type: context.drawing_type || null,
        // Running stats for confidence
        min_rate: rate,
        max_rate: rate,
        rates_history: [rate],
      };
      newCount++;
    } else {
      // Update existing entry — use weighted average (recent weighs more)
      const prevCount = existing.count || 1;
      existing.rate = Math.round((existing.rate * prevCount + rate) / (prevCount + 1));
      existing.count = prevCount + 1;
      existing.last_seen = new Date().toISOString();
      existing.min_rate = Math.min(existing.min_rate || rate, rate);
      existing.max_rate = Math.max(existing.max_rate || rate, rate);
      if (!Array.isArray(existing.rates_history)) existing.rates_history = [];
      existing.rates_history = [...existing.rates_history.slice(-19), rate]; // keep last 20
      if (!Array.isArray(existing.seen_in)) existing.seen_in = [];
      const fname = context.filename || 'unknown';
      if (!existing.seen_in.includes(fname)) {
        existing.seen_in = [...existing.seen_in.slice(-9), fname];
      }
      updateCount++;
    }
  }

  saveLearnedRates(learned);
  console.log(`[rate_store] Learned ${newCount} new rates, updated ${updateCount} from ${context.filename || 'drawing'}`);
  return { new: newCount, updated: updateCount };
}

// ── Build combined rates summary for Claude prompts ───────────────
/**
 * Returns a concise string of rates for Claude prompts.
 * Learned rates override base DSR rates.
 * @param {Object} options — { maxItems, drawingType, format }
 */
function getRatesSummary(options = {}) {
  const { maxItems = 40, format = 'short' } = options;

  const base = loadBaseRates();
  const learned = loadLearnedRates();
  const learnedRates = learned.rates || {};

  // Merge: base first, then overlay with learned
  const merged = { ...base };
  for (const [key, lr] of Object.entries(learnedRates)) {
    // Map learned key → find closest base key or add new
    const simpleKey = key.split('__')[0].toLowerCase();
    // Try to find a matching base key
    const baseMatch = Object.keys(base).find(bk =>
      bk.toLowerCase().includes(simpleKey.slice(0, 10)) ||
      simpleKey.includes(bk.toLowerCase().slice(0, 10))
    );
    if (baseMatch) {
      // Learned rate overrides base if seen 2+ times
      if ((lr.count || 0) >= 2) {
        merged[baseMatch] = { ...merged[baseMatch], rate: lr.rate, source: 'learned' };
      }
    } else {
      // Brand new item not in base rates
      merged[key] = { rate: lr.rate, unit: lr.unit, description: lr.description, source: 'learned' };
    }
  }

  const lines = Object.values(merged)
    .filter(v => v?.rate && v?.description)
    .slice(0, maxItems)
    .map(v => {
      const tag = v.source === 'learned' ? ' [learned]' : '';
      return format === 'short'
        ? `${v.description}: Rs.${v.rate}/${v.unit}${tag}`
        : `${v.description} → Rs.${v.rate}/${v.unit} (${v.unit})${tag}`;
    });

  return lines.join('\n');
}

// ── Get flat rate map {key: rate} for quick lookup ────────────────
function getRatesMap() {
  const base = loadBaseRates();
  const learned = loadLearnedRates();
  const out = {};
  for (const [k, v] of Object.entries(base)) out[k] = v.rate;
  for (const [, lr] of Object.entries(learned.rates || {})) {
    if (lr.count >= 2 && lr.rate > 0) {
      const simpleKey = lr.description.toLowerCase().replace(/\s+/g, '_').slice(0, 40);
      out[simpleKey] = lr.rate;
    }
  }
  return out;
}

// ── Stats for admin/debug ─────────────────────────────────────────
function getLearnedRateStats() {
  const learned = loadLearnedRates();
  const rates = learned.rates || {};
  const sorted = Object.values(rates).sort((a, b) => (b.count || 0) - (a.count || 0));
  return {
    total_analyses: learned._meta?.total_analyses || 0,
    last_updated: learned._meta?.last_updated || null,
    total_learned_items: sorted.length,
    top_items: sorted.slice(0, 20).map(r => ({
      description: r.description,
      unit: r.unit,
      rate: r.rate,
      count: r.count,
      min: r.min_rate,
      max: r.max_rate,
    })),
  };
}

// ── Auto-learn from Claude's text analysis (markdown BOQ table) ───
/**
 * Extract rates from Claude's markdown response (text analysis mode).
 * Looks for BOQ table rows like: | description | unit | qty | rate | amount |
 */
function learnRatesFromMarkdown(markdownText, context = {}) {
  if (!markdownText) return;
  const boqItems = [];

  // Match table rows with 5+ columns (standard BOQ format)
  const rowPat = /\|\s*([^|]{3,60})\s*\|\s*(sqmt|cum|rmt|nos|kg|kl|lump|sqft|rft|ltr)\s*\|\s*([\d.]+)\s*\|\s*([\d,]+)\s*\|\s*([\d,]+)/gi;
  let m;
  while ((m = rowPat.exec(markdownText)) !== null) {
    const description = m[1].trim();
    const unit = m[2].trim().toLowerCase();
    const qty = parseFloat(m[3]);
    const rate = parseFloat(m[4].replace(/,/g, ''));
    const amount = parseFloat(m[5].replace(/,/g, ''));
    if (description && unit && rate > 0) {
      boqItems.push({ description, unit, qty, rate, amount, source: 'claude_markdown' });
    }
  }

  if (boqItems.length > 0) {
    learnRatesFromBOQ(boqItems, context);
  }
  return boqItems.length;
}

module.exports = {
  learnRatesFromBOQ,
  learnRatesFromMarkdown,
  getRatesSummary,
  getRatesMap,
  getLearnedRateStats,
  loadBaseRates,
  loadLearnedRates,
};
