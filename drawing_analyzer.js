/**
 * PMC Drawing Analyzer — OPTIMIZED SINGLE-PHASE
 * ─────────────────────────────────────────────
 * BEFORE: 4 Claude API calls per drawing (Phase 2+3+4+5) = HIGH COST
 * AFTER:  1 Claude API call per drawing = 75% cost reduction
 *
 * Rule-based validation replaces Phase 5 (zero API cost).
 * server.js needs ZERO changes — same function signatures.
 */

'use strict';

const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ── RATES from Rates.json ─────────────────────────────────────────
const RATES = (() => {
  const out = {};
  try {
    const rp = fs.existsSync(path.join(__dirname, 'rates.json'))
      ? path.join(__dirname, 'rates.json') : path.join(__dirname, 'Rates.json');
    const raw = JSON.parse(fs.readFileSync(rp, 'utf8'));
    for (const cat of Object.values(raw))
      if (typeof cat === 'object' && !Array.isArray(cat))
        for (const [k, v] of Object.entries(cat))
          if (v?.rate) out[k] = v.rate;
  } catch (e) { console.warn('Rates.json not loaded:', e.message); }
  return out;
})();

const RATES_STRING = (() => {
  try {
    const rp = fs.existsSync(path.join(__dirname, 'Rates.json'))
      ? path.join(__dirname, 'Rates.json') : path.join(__dirname, 'rates.json');
    const raw = JSON.parse(fs.readFileSync(rp, 'utf8'));
    const lines = [];
    for (const [cat, items] of Object.entries(raw)) {
      if (cat.startsWith('_') || typeof items !== 'object') continue;
      for (const [, v] of Object.entries(items))
        if (v?.rate) lines.push(`${v.description} → Rs.${v.rate}/${v.unit}`);
    }
    return lines.join('\n');
  } catch { return ''; }
})();

// ── KNOWLEDGE BASE ────────────────────────────────────────────────
function knowledgeBaseHints() {
  try {
    const kb = JSON.parse(fs.readFileSync(path.join(__dirname, 'ymbols-learned.json'), 'utf8'));
    const hints = [];
    if (kb.quantity_corrections?.length) {
      hints.push('CORRECTION HISTORY:');
      for (const c of kb.quantity_corrections.slice(-5))
        hints.push(`  ${c.element}: AI=${c.ai_said}, correct=${c.correct_value}`);
    }
    const blocks = Object.entries(kb.blocks || {}).slice(0, 10);
    if (blocks.length) {
      hints.push('KNOWN BLOCKS:');
      for (const [b, m] of blocks) hints.push(`  ${b} = ${m}`);
    }
    return hints.join('\n');
  } catch { return ''; }
}

// ── SYSTEM PROMPT ─────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a senior PMC civil engineer with 20 years experience in Gujarat, India.
You read AutoCAD drawings (DWG, DXF, PDF) and generate accurate BOQ for civil works.

GOLDEN RULES:
1. NEVER invent or guess dimensions. Only values visible in THIS drawing.
2. Read the legend FIRST before counting any element.
3. Apply scale factor from title block to ALL measurements.
4. Mark every quantity: source = "drawing" | "calculated" | "assumed"
5. Return ONLY raw JSON. No markdown. No explanation.

GUJARAT DSR 2025 RATES:
${RATES_STRING}`;

// ── CLAUDE API (single call) ──────────────────────────────────────
async function callClaude({ messages, maxTokens = 8192 }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');
  const body = {
    model: 'claude-sonnet-4-6',
    max_tokens: maxTokens,
    system: SYSTEM_PROMPT,
    messages,
  };
  for (let i = 0; i <= 3; i++) {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-beta':'pdfs-2024-09-25' },
      body: JSON.stringify(body),
    });
    const data = await r.json();
    if (r.ok && data.content) return data.content.filter(b=>b.type==='text').map(b=>b.text).join('');
    if (data.error?.type !== 'overloaded_error') throw new Error(`Claude: ${data.error?.message}`);
    await new Promise(res=>setTimeout(res, 2000*(i+1)));
  }
  throw new Error('Claude API: max retries exceeded');
}

function parseJSON(raw) {
  if (!raw) return null;
  const clean = raw.replace(/```json|```/g,'').trim();
  const fb=clean.indexOf('{'), lb=clean.lastIndexOf('}');
  if (fb===-1||lb===-1) return null;
  try { return JSON.parse(clean.slice(fb,lb+1)); }
  catch(e) { console.error('[Claude] JSON parse fail:',e.message,clean.slice(0,200)); return null; }
}

function buildImageParts(files) {
  return (files||[]).flatMap(f => {
    if (f.type?.startsWith('image/'))
      return [{ type:'image', source:{ type:'base64', media_type:f.type||'image/png', data:f.b64 } }];
    if (f.type==='application/pdf' || f.name?.match(/\.pdf$/i))
      return [{ type:'document', source:{ type:'base64', media_type:'application/pdf', data:f.b64 } }];
    return [];
  });
}

// ── RULE-BASED VALIDATION (FREE — replaces Phase 5 API call) ─────
function validateBOQ(boq) {
  const warnings = [];
  const passed = [];

  for (const item of (boq.boq || [])) {
    const desc = (item.description || '').toUpperCase();
    const qty = Number(item.qty) || 0;
    const rate = Number(item.rate) || 0;
    const amount = Number(item.amount) || 0;

    // Math check: qty * rate should equal amount (within 5%)
    if (qty > 0 && rate > 0 && Math.abs(qty * rate - amount) / amount > 0.05) {
      warnings.push({ item: item.description, check: 'Math mismatch', expected: `${qty}×${rate}=${qty*rate}`, found: `${amount}`, severity: 'HIGH' });
    }

    // Steel ratio checks
    if (desc.includes('STEEL') || desc.includes('REINFORCEMENT') || desc.includes('REBAR')) {
      const steelKg = qty;
      const matchedRCC = boq.boq?.find(b => {
        const bd = (b.description||'').toUpperCase();
        return (bd.includes('SLAB') || bd.includes('BEAM') || bd.includes('COLUMN')) && b.unit === 'CUM';
      });
      if (matchedRCC && matchedRCC.qty > 0) {
        const ratio = steelKg / matchedRCC.qty;
        if (desc.includes('SLAB') && (ratio < 80 || ratio > 160))
          warnings.push({ item: item.description, check: 'Steel ratio slab', expected: '100-140 kg/CUM', found: `${Math.round(ratio)} kg/CUM`, severity: 'MEDIUM' });
        if (desc.includes('COLUMN') && (ratio < 150 || ratio > 280))
          warnings.push({ item: item.description, check: 'Steel ratio column', expected: '180-240 kg/CUM', found: `${Math.round(ratio)} kg/CUM`, severity: 'MEDIUM' });
      }
    }

    // Zero qty with amount
    if (qty === 0 && amount > 0)
      warnings.push({ item: item.description, check: 'Qty=0 but amount>0', expected: 'qty>0', found: `qty=${qty} amount=${amount}`, severity: 'HIGH' });
  }

  // Cost per sqmt sanity
  const totalInr = boq.cost_summary?.civil_total_inr || 0;
  const totalArea = boq.area_statement?.total_bua_sqmt || 0;
  if (totalInr > 0 && totalArea > 0) {
    const costPerSqmt = totalInr / totalArea;
    if (costPerSqmt < 1000 || costPerSqmt > 8000)
      warnings.push({ item: 'TOTAL COST', check: 'Cost/sqmt sanity', expected: 'Rs.1500-5000/sqmt', found: `Rs.${Math.round(costPerSqmt)}/sqmt`, severity: 'MEDIUM' });
    else
      passed.push(`Cost/sqmt = Rs.${Math.round(costPerSqmt)} (within range)`);
  }

  if (warnings.length === 0) passed.push('All rule-based checks passed');

  return {
    ...boq,
    validation_warnings: warnings,
    validation_passed: passed,
    overall_confidence: warnings.filter(w=>w.severity==='HIGH').length > 0 ? 'LOW' : warnings.length > 2 ? 'MEDIUM' : 'HIGH',
    engineer_action_required: warnings.filter(w=>w.severity==='HIGH').map(w=>w.item)
  };
}

// ════════════════════════════════════════════════════════════════
// MAIN — SINGLE API CALL (was 4 calls before)
// ════════════════════════════════════════════════════════════════
async function geminiAnalyzeDrawing(key, files, cvData, fetchFn) {
  console.log('\n[PMC] === Single-Phase Claude Analysis (1 API call) ===');

  const imgParts = buildImageParts(files);
  if (!imgParts.length) return null;

  const kb = knowledgeBaseHints();
  const cv = cvData && !cvData.error
    ? `CV data: ${cvData.image_dimensions?.width_px}x${cvData.image_dimensions?.height_px}px | spaces:${cvData.detected_spaces?.length||0}`
    : '';

  // ONE comprehensive prompt — all 4 phases merged
  const prompt = `${cv}
${kb}

You are analyzing this architectural/structural drawing. Complete ALL steps in ONE response:

STEP 1 — Read legend/symbol table + title block (project name, scale, drawing type)
STEP 2 — Extract ALL quantities floor-by-floor using the legend
STEP 3 — Calculate BOQ with Gujarat DSR 2025 rates
Return validation notes inline in observations[]

Return ONLY this JSON (no markdown):
{
  "project_name": "",
  "drawing_type": "FLOOR_PLAN|SECTION|STRUCTURAL|FOUNDATION|SITE_PLAN",
  "drawing_no": "",
  "scale": "1:100",
  "scale_factor": 100,
  "legend": [{"symbol":"","meaning":"","layer":""}],
  "boq": [
    {"sr":1,"part":"PART A — EARTHWORK","description":"","unit":"CUM|SQM|RMT|NOS|KG","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed","confidence":"high|medium|low"}
  ],
  "element_counts": {"door_count":0,"window_count":0,"column_count":0,"lift_count":0,"staircase_count":0,"bedroom_count":0,"toilet_count":0,"floor_count":0},
  "area_statement": {"total_bua_sqmt":0,"floor_wise":[]},
  "cost_summary": {"civil_total_inr":0,"civil_total_lacs":0},
  "observations": [],
  "missing_info": []
}`;

  const raw = await callClaude({
    messages: [{ role: 'user', content: [...imgParts, { type: 'text', text: prompt }] }],
    maxTokens: 8192
  });

  let boqData = parseJSON(raw);
  if (!boqData) boqData = { boq: [], observations: ['Parse failed'], overall_confidence: 'LOW' };

  // Rule-based validation — FREE, no API call
  const finalData = validateBOQ(boqData);

  console.log(`[PMC] Done — ${finalData.boq?.length||0} BOQ items | confidence:${finalData.overall_confidence} | warnings:${finalData.validation_warnings?.length||0}`);
  return buildFinalOutput(finalData, cvData);
}

function buildFinalOutput(boq, cvData) {
  if (!boq) return null;
  const totalInr = boq.cost_summary?.civil_total_inr || 0;
  return {
    project_name:     boq.project_name || 'CIVIL PROJECT',
    drawing_type:     boq.drawing_type || 'GENERAL',
    drawing_no:       boq.drawing_no   || '',
    scale:            boq.scale        || '',
    legend:           boq.legend       || [],
    boq_items:        boq.boq         || [],
    element_counts:   boq.element_counts || {},
    area_statement:   boq.area_statement || { total_bua_sqmt: 0, floor_wise: [] },
    cost_summary: {
      civil_total_inr:    totalInr,
      civil_total_lacs:   Math.round(totalInr / 100000 * 100) / 100,
      civil_total_crores: Math.round(totalInr / 10000000 * 100) / 100,
    },
    observations:             boq.observations           || [],
    missing_info:             boq.missing_info           || [],
    validation_warnings:      boq.validation_warnings    || [],
    validation_passed:        boq.validation_passed      || [],
    overall_confidence:       boq.overall_confidence     || 'MEDIUM',
    engineer_action_required: boq.engineer_action_required || [],
    cv_analysis:              cvData || {},
    prepared_by:              'PMC Civil AI Agent',
    pipeline_info: {
      api_calls: 1,
      validation: 'rule-based (free)',
      cost_vs_old: '75% reduction'
    }
  };
}

// ── CV (unchanged) ────────────────────────────────────────────────
function runCVAnalysis(b64Image) {
  try {
    const tmp = path.join(os.tmpdir(), `drawing_cv_${Date.now()}.txt`);
    fs.writeFileSync(tmp, b64Image);
    const result = execSync(`python3 ${path.join(__dirname,'drawing_cv.py')} ${tmp}`, { timeout:30000 });
    fs.unlinkSync(tmp);
    return JSON.parse(result.toString());
  } catch(e) { console.error('[CV] failed:', e.message); return { error: e.message }; }
}

module.exports = { geminiAnalyzeDrawing, runCVAnalysis, RATES };
