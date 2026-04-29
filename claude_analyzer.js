/**
 * PMC Drawing Analyzer — Claude 5-Phase Pipeline
 * ─────────────────────────────────────────────────────────────────
 * PHASE 1 → CV Pre-process   (OpenCV — existing, untouched)
 * PHASE 2 → Legend + Scale   (Claude — title block ONLY, no BOQ yet)
 * PHASE 3 → Quantity Extract (Claude — with legend context)
 * PHASE 4 → BOQ Calculate    (Claude — rates applied to clean quantities)
 * PHASE 5 → Validate         (Claude extended thinking — flags anomalies)
 * ─────────────────────────────────────────────────────────────────
 * Drop-in replacement — server.js needs ZERO changes.
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
    for (const cat of Object.values(raw)) {
      if (typeof cat === 'object' && !Array.isArray(cat))
        for (const [k, v] of Object.entries(cat))
          if (v?.rate) out[k] = v.rate;
    }
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

// ── KNOWLEDGE BASE hints ──────────────────────────────────────────
function knowledgeBaseHints() {
  try {
    const kb = JSON.parse(fs.readFileSync(path.join(__dirname, 'ymbols-learned.json'), 'utf8'));
    const hints = [];
    if (kb.quantity_corrections?.length) {
      hints.push('CORRECTION HISTORY (apply these lessons):');
      for (const c of kb.quantity_corrections.slice(-10))
        hints.push(`  WARNING ${c.element}: AI said ${c.ai_said}, correct was ${c.correct_value}. Reason: ${c.correction_reason || 'engineer corrected'}`);
    }
    if (kb.scale_corrections?.length) {
      hints.push('SCALE ISSUES HISTORY:');
      for (const s of kb.scale_corrections.slice(-5))
        hints.push(`  WARNING Scale detected ${s.detected} but actual was ${s.actual}`);
    }
    const blocks = Object.entries(kb.blocks || {}).slice(0, 15);
    if (blocks.length) {
      hints.push('CONFIRMED BLOCK MEANINGS:');
      for (const [b, m] of blocks) hints.push(`  ${b} = ${m}`);
    }
    return hints.join('\n');
  } catch { return ''; }
}

// ── SYSTEM PROMPT ─────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a senior PMC civil engineer with 20 years experience in Gujarat, India.
You read AutoCAD drawings (DWG, DXF, PDF) and generate accurate BOQ for civil works.

ABSOLUTE RULES — VIOLATION IS PROFESSIONAL MISCONDUCT:
1. SCHEDULE TABLE is the ONLY valid source for column/footing sizes, steel details, quantities.
2. Read each schedule cell EXACTLY as printed — do NOT fill in, guess, or extrapolate any value.
3. NEVER mix footing schedule into column schedule or vice versa — they are completely separate tables.
4. NEVER generate a size (400×400, 500×500 etc.) unless that exact number is printed in the schedule.
5. NEVER generate steel details (12-20Ø, 16-25Ø etc.) unless exactly printed in the schedule.
6. If a cell value is not clearly readable → output exactly: "not legible" — never guess.
7. BOQ only from values actually read — no assumed quantities.
8. Mark source: "drawing-schedule" | "calculated" | "not legible". Return ONLY raw JSON.

GUJARAT DSR 2025 RATES:
${RATES_STRING}

VALIDATION RATIOS (flag if outside):
- Steel in RCC slab: 100-140 kg/CUM
- Steel in beam: 150-200 kg/CUM
- Steel in column: 180-240 kg/CUM
- Road GSB: area x 1.15 x 0.3 x 1.8 tonnes
- Road WMM: area x 1.15 x 0.2 x 2.1 tonnes`;

// ── CLAUDE API CALL ───────────────────────────────────────────────
async function callClaude({ messages, maxTokens = 8192, thinking = false }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');
  const body = {
    model: 'claude-sonnet-4-5', max_tokens: thinking ? 16000 : maxTokens,
    system: SYSTEM_PROMPT, messages,
  };
  if (thinking) body.thinking = { type: 'enabled', budget_tokens: 8000 };

  for (let i = 0; i <= 4; i++) {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-beta':'pdfs-2024-09-25' },
      body: JSON.stringify(body),
    });
    const data = await r.json();
    if (r.ok && data.content) return data.content.filter(b=>b.type==='text').map(b=>b.text).join('');
    if (data.error?.type !== 'overloaded_error') throw new Error(`Claude: ${data.error?.message}`);
    await new Promise(r=>setTimeout(r, 2000*(i+1)));
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

// ════════════════════════════════════════════
// PHASE 2 — Legend + Title Block
// ════════════════════════════════════════════
async function phase2_legendAndScale(files, cvData) {
  console.log('[Phase 2] Reading legend + title block...');
  const kb = knowledgeBaseHints();
  const cv = cvData&&!cvData.error
    ? `CV: ${cvData.image_dimensions?.width_px}x${cvData.image_dimensions?.height_px}px | spaces:${cvData.detected_spaces?.length||0} | scale candidates:${JSON.stringify(cvData.scale_bar_candidates_px?.slice(0,4))}` : '';
  const imgParts = buildImageParts(files);
  if (!imgParts.length) return null;

  const raw = await callClaude({
    messages:[{ role:'user', content:[
      ...imgParts,
      { type:'text', text:`${cv}\n${kb}\n\nTASK PHASE 2: Read ONLY legend/symbol table and title block. No BOQ yet.\nReturn JSON:\n{"drawing_type":"","project_name":"","drawing_no":"","date":"","scale":"1:100","scale_factor":100,"north_direction":"","legend":[{"symbol":"","meaning":"","layer":""}],"floors_visible":[],"title_block_confidence":"HIGH|MEDIUM|LOW","legend_confidence":"HIGH|MEDIUM|LOW","notes":[]}` }
    ]}],
    maxTokens: 2048
  });

  const meta = parseJSON(raw);
  console.log(`[Phase 2] type:${meta?.drawing_type} scale:${meta?.scale} legend:${meta?.legend?.length||0} items`);
  return meta;
}

// ════════════════════════════════════════════
// PHASE 3 — Quantity Extraction
// ════════════════════════════════════════════
async function phase3_extractQuantities(files, meta) {
  console.log('[Phase 3] Extracting quantities with legend context...');
  const legendCtx = meta?.legend?.length
    ? `LEGEND FROM PHASE 2 (use for all element identification):\n${meta.legend.map(l=>`  ${l.symbol} = ${l.meaning} (layer:${l.layer||'?'})`).join('\n')}`
    : 'No legend — use standard CAD conventions.';
  const scaleCtx = meta?.scale
    ? `SCALE: ${meta.scale} (scale_factor=${meta.scale_factor}). Apply to ALL dimensions.`
    : 'Scale not confirmed — mark confidence LOW.';
  const imgParts = buildImageParts(files);

  const raw = await callClaude({
    messages:[{ role:'user', content:[
      ...imgParts,
      { type:'text', text:`${legendCtx}\n${scaleCtx}\nDrawing type:${meta?.drawing_type||'unknown'}\n\nTASK PHASE 3: Extract ALL quantities visible in drawing.\nReturn JSON:\n{"quantities":[{"element":"","floor":"","length_m":0,"width_m":0,"height_m":0,"thickness_m":0,"nos":1,"area_sqmt":0,"volume_cum":0,"unit":"","annotation_text":"","source":"drawing|calculated|assumed","confidence":"high|medium|low"}],"element_counts":{"door_count":0,"window_count":0,"column_count":0,"footing_count":0,"staircase_count":0,"lift_count":0,"bedroom_count":0,"toilet_count":0,"kitchen_count":0,"floor_count":0},"road_data":{"roads":[{"name":"","length_rmt":0,"total_width_m":0,"carriage_width_m":0}]},"total_built_area_sqmt":0,"observations":[]}` }
    ]}],
    maxTokens: 6000
  });

  const q = parseJSON(raw);
  console.log(`[Phase 3] ${q?.quantities?.length||0} elements extracted`);
  return q;
}

// ════════════════════════════════════════════
// PHASE 4 — BOQ Calculation
// ════════════════════════════════════════════
async function phase4_calculateBOQ(quantities, meta) {
  console.log('[Phase 4] Calculating BOQ...');
  if (!quantities) return null;

  const raw = await callClaude({
    messages:[{ role:'user', content:`TASK PHASE 4: Calculate BOQ from quantities.\n\nQUANTITIES:\n${JSON.stringify(quantities,null,2)}\n\nDRAWING: type=${meta?.drawing_type||'?'} project=${meta?.project_name||'?'} scale=${meta?.scale||'?'}\n\nUse DSR 2025 rates from system prompt. Group into PARTS.\nReturn JSON:\n{"project_name":"","drawing_type":"","drawing_no":"","date":"","scale":"","boq":[{"sr":1,"part":"PART A","description":"","unit":"","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed","confidence":"high|medium|low"}],"element_counts":{},"area_statement":{"total_bua_sqmt":0,"floor_wise":[],"road_area_sqmt":0,"road_length_rmt":0},"cost_summary":{"civil_total_inr":0,"civil_total_lacs":0,"civil_total_crores":0},"observations":[],"missing_info":[]}` }],
    maxTokens: 8192
  });

  const boq = parseJSON(raw);
  console.log(`[Phase 4] ${boq?.boq?.length||0} BOQ items, total: Rs.${boq?.cost_summary?.civil_total_lacs||0} lacs`);
  return boq;
}

// ════════════════════════════════════════════
// PHASE 5 — Validation (Extended Thinking)
// ════════════════════════════════════════════
async function phase5_validateAndFlag(boqData) {
  console.log('[Phase 5] Validating with extended thinking...');
  if (!boqData) return boqData;

  const raw = await callClaude({
    messages:[{ role:'user', content:`TASK PHASE 5: Validate this BOQ. Use extended thinking.\n\n${JSON.stringify(boqData,null,2)}\n\nCHECKS:\n1. Steel ratios (slab 100-140, beam 150-200, column 180-240 kg/CUM)\n2. Wall area vs floor area (0.6x to 1.2x)\n3. Road GSB: area x 1.15 x 0.3 x 1.8\n4. Road WMM: area x 1.15 x 0.2 x 2.1\n5. Cost per sqmt sanity (building Rs.1500-3500, road Rs.2000-4000)\n6. Any qty=0 with amount>0 (math error)\n7. Items with source=assumed\n\nAdd to existing JSON:\n- "validation_warnings":[{"item":"","check":"","expected":"","found":"","severity":"HIGH|MEDIUM|LOW"}]\n- "validation_passed":["check desc"]\n- "overall_confidence":"HIGH|MEDIUM|LOW"\n- "engineer_action_required":["what to verify manually"]` }],
    maxTokens: 10000,
    thinking: true
  });

  const validated = parseJSON(raw);
  if (!validated) {
    boqData.validation_warnings = [];
    boqData.validation_passed = ['Phase 5 parse failed — manual review recommended'];
    boqData.overall_confidence = 'MEDIUM';
    return boqData;
  }
  console.log(`[Phase 5] ${validated.validation_warnings?.length||0} warnings | confidence:${validated.overall_confidence}`);
  return validated;
}

// ════════════════════════════════════════════
// PHASE 1 — CV (existing OpenCV, unchanged)
// ════════════════════════════════════════════
function runCVAnalysis(b64Image) {
  try {
    const tmp = path.join(os.tmpdir(),`drawing_cv_${Date.now()}.txt`);
    fs.writeFileSync(tmp, b64Image);
    const result = execSync(`python3 ${path.join(__dirname,'drawing_cv.py')} ${tmp}`, {timeout:30000});
    fs.unlinkSync(tmp);
    return JSON.parse(result.toString());
  } catch(e) { console.error('[Phase 1] CV failed:',e.message); return {error:e.message}; }
}

// ════════════════════════════════════════════
// MAIN — same function name as before
// server.js needs ZERO changes
// ════════════════════════════════════════════
async function geminiAnalyzeDrawing(key, files, cvData, fetchFn) {
  console.log('\n[PMC] === 5-Phase Claude Pipeline Starting ===');

  const meta        = await phase2_legendAndScale(files, cvData);
  const quantities  = await phase3_extractQuantities(files, meta);
  const boqData     = await phase4_calculateBOQ(quantities, meta);
  const finalData   = await phase5_validateAndFlag(boqData);

  return buildFinalOutput(finalData, quantities, meta, cvData);
}

function buildFinalOutput(boq, quantities, meta, cvData) {
  if (!boq) return null;
  const totalInr   = boq.cost_summary?.civil_total_inr || 0;
  const totalLacs  = boq.cost_summary?.civil_total_lacs || Math.round(totalInr/100000*100)/100;
  const totalCr    = boq.cost_summary?.civil_total_crores || Math.round(totalLacs/100*100)/100;

  return {
    project_name: boq.project_name||meta?.project_name||'',
    drawing_no:   boq.drawing_no||meta?.drawing_no||'',
    drawing_type: boq.drawing_type||meta?.drawing_type||'',
    scale:        boq.scale||meta?.scale||'',
    date:         boq.date||meta?.date||'',
    elements: (boq.boq||[]).map((item,i)=>({
      id:`E${String(i+1).padStart(3,'0')}`, type:guessType(item.description),
      name:item.description, dimensions:{note:item.source||''},
      quantities:{
        area_sqmt:   item.unit==='sqmt'?item.qty:0,
        volume_cum:  item.unit==='cum'?item.qty:0,
        gsb_ton:     /gsb/i.test(item.description)?item.qty:0,
        wmm_ton:     /wmm/i.test(item.description)?item.qty:0,
        pqc_cum:     /pqc/i.test(item.description)?item.qty:0,
        steel_kg:    item.unit==='kg'?item.qty:0,
        brickwork_cum:/brick/i.test(item.description)?item.qty:0,
      },
      cost_inr:{total:item.amount||0,per_unit:item.rate||0},
      confidence:item.confidence||'medium', annotation_found:item.source||'',
      part:item.part||'', sr:item.sr||i+1,
    })),
    element_counts:{...quantities?.element_counts,...boq.element_counts},
    total_quantities:{
      total_area_sqmt: boq.area_statement?.total_bua_sqmt||0,
      total_road_rmt:  boq.area_statement?.road_length_rmt||0,
      gsb_total_ton:   sumKw(boq.boq,'gsb'),
      wmm_total_ton:   sumKw(boq.boq,'wmm'),
      pqc_total_cum:   sumKw(boq.boq,'pqc'),
      rcc_total_cum:   sumKw(boq.boq,'rcc'),
      steel_total_kg:  sumKw(boq.boq,'steel'),
      calculation_note:'5-phase Claude pipeline',
    },
    cost_summary:{civil_total_inr:totalInr,civil_total_lacs:totalLacs,civil_total_crores:totalCr,item_wise:boq.boq||[]},
    area_statement:    boq.area_statement||{},
    validation_warnings:      boq.validation_warnings||[],
    validation_passed:        boq.validation_passed||[],
    overall_confidence:       boq.overall_confidence||'MEDIUM',
    engineer_action_required: boq.engineer_action_required||[],
    legend: meta?.legend||[],
    observations:[...(boq.observations||[]),...(boq.missing_info||[])],
    pmc_recommendation:`5-phase Claude pipeline. Confidence:${boq.overall_confidence||'MEDIUM'}. ${boq.engineer_action_required?.length?'Verify: '+boq.engineer_action_required.join(', '):'No major issues.'}`,
    extraction_confidence: boq.overall_confidence||'MEDIUM',
    missing_info: boq.missing_info||[],
    rates_applied: RATES,
    prepared_by:'PMC Civil AI — Claude 5-Phase Pipeline',
    cv_analysis: cvData||{},
    pipeline_info:{
      phases_completed:5,
      phase2_legend_items:meta?.legend?.length||0,
      phase3_elements:quantities?.quantities?.length||0,
      phase4_boq_items:boq.boq?.length||0,
      phase5_warnings:boq.validation_warnings?.length||0,
      model:'claude-sonnet-4-5',
    },
  };
}

function guessType(desc){
  if(!desc)return'GENERAL';
  const d=desc.toLowerCase();
  if(/road|gsb|wmm|pqc|kerb/.test(d))return'ROAD';
  if(/steel|rebar|bar/.test(d))return'STEEL';
  if(/rcc|slab|beam|column|footing/.test(d))return'STRUCTURE';
  if(/brick|wall/.test(d))return'WALL';
  if(/excavat|earth/.test(d))return'EARTHWORK';
  if(/plaster|paint|floor|tile/.test(d))return'FINISH';
  return'GENERAL';
}

function sumKw(items,kw){
  if(!items)return 0;
  return Math.round(items.filter(i=>i.description?.toLowerCase().includes(kw)).reduce((s,i)=>s+(i.qty||0),0)*100)/100;
}

function calcRoadQuantities(length_m,width_m){
  const a=length_m*width_m;
  return{area_sqmt:Math.round(a*100)/100,box_cutting_sqmt:Math.round(a*1.05*100)/100,
    gsb_300mm_ton:Math.round(a*1.15*0.3*1.8*100)/100,wmm_200mm_ton:Math.round(a*1.15*0.2*2.1*100)/100,
    pqc_250mm_cum:Math.round(a*1.05*0.25*100)/100,steel_dowel_ton:Math.round(a*0.00387*100)/100,
    cost_estimate:{gsb:Math.round(a*(RATES.gsb_300mm_sqmt||655)),wmm:Math.round(a*(RATES.wmm_200mm_sqmt||515)),
      pqc:Math.round(a*(RATES.pqc_250mm_sqmt||1800)),total_sqmt:Math.round(a*((RATES.gsb_300mm_sqmt||655)+(RATES.wmm_200mm_sqmt||515)+(RATES.pqc_250mm_sqmt||1800)))}};
}

function calcStructureQuantities(dims){
  const{length=0,width=0,height=0,nos=1}=dims;
  const v=length*width*height*nos,a=length*width*nos;
  return{volume_cum:Math.round(v*1000)/1000,area_sqmt:Math.round(a*100)/100,
    steel_kg:Math.round(v*120),formwork_sqmt:Math.round((2*(length+width)*height+a)*nos*100)/100};
}

// ─── Compatibility aliases ────────────────────────────────────────────────────
// server.js imports these names — wire them to the internal implementations.

/** Alias: CIVIL_SYSTEM → SYSTEM_PROMPT */
const CIVIL_SYSTEM = SYSTEM_PROMPT;

/**
 * Alias: callClaudeAPI({ system, messages, maxTokens })
 * server.js calls this directly for /api/claude, /api/analyze-drawing, etc.
 */
async function callClaudeAPI({ system, messages, maxTokens = 8192 }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');
  const body = {
    model: 'claude-sonnet-4-5',
    max_tokens: maxTokens,
    system: system || SYSTEM_PROMPT,
    messages,
  };
  for (let i = 0; i <= 4; i++) {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': 'pdfs-2024-09-25',
      },
      body: JSON.stringify(body),
    });
    const data = await r.json();
    if (r.ok && data.content)
      return data.content.filter(b => b.type === 'text').map(b => b.text).join('');
    if (data.error?.type !== 'overloaded_error')
      throw new Error(`Claude API: ${data.error?.message}`);
    await new Promise(res => setTimeout(res, 2000 * (i + 1)));
  }
  throw new Error('Claude API: max retries exceeded');
}

/**
 * claudeAnalyzeDXF — analyse parsed DXF civil data via Claude
 * server.js calls: claudeAnalyzeDXF(civilData, filename, rSummary)
 */
async function claudeAnalyzeDXF(civilData, filename, ratesSummary) {
  console.log('[claudeAnalyzeDXF] analysing DXF data for', filename);
  const prompt = `You are a senior PMC civil engineer. Analyse this parsed DXF data and generate a BOQ.
DXF FILE: ${filename}
RATES SUMMARY: ${ratesSummary || 'Use DSR 2025 rates from system prompt.'}
DXF DATA:
${JSON.stringify(civilData, null, 2)}

Return ONLY raw JSON:
{"project_name":"","drawing_type":"","boq":[{"sr":1,"part":"PART A","description":"","unit":"","qty":0,"rate":0,"amount":0,"source":"dxf-data","confidence":"high"}],"cost_summary":{"civil_total_inr":0,"civil_total_lacs":0},"observations":[]}`;
  const raw = await callClaudeAPI({ system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }], maxTokens: 8192 });
  return parseJSON(raw);
}

/**
 * claudeClassifySymbols — classify unknown DXF blocks/layers
 * server.js calls: claudeClassifySymbols(unknownBlocks, unknownLayers, civilData, filename)
 */
async function claudeClassifySymbols(unknownBlocks, unknownLayers, civilData, filename) {
  console.log('[claudeClassifySymbols] classifying', unknownBlocks?.length, 'blocks');
  const prompt = `Classify these unknown CAD blocks and layers for a civil engineering drawing.
FILE: ${filename}
UNKNOWN BLOCKS: ${JSON.stringify(unknownBlocks)}
UNKNOWN LAYERS: ${JSON.stringify(unknownLayers)}
CONTEXT: ${JSON.stringify(civilData?.summary || {})}

Return ONLY raw JSON:
{"classified_blocks":[{"name":"","civil_meaning":"","category":"STRUCTURE|ROAD|UTILITY|ANNOTATION|UNKNOWN"}],"classified_layers":[{"name":"","civil_meaning":"","category":"STRUCTURE|ROAD|UTILITY|ANNOTATION|UNKNOWN"}]}`;
  const raw = await callClaudeAPI({ system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }], maxTokens: 4096 });
  return parseJSON(raw);
}

/**
 * claudeAnalyzeWithAnswers — full DXF analysis with symbol answers
 * server.js calls: claudeAnalyzeWithAnswers(civilData, filename, symbolSummary, ratesSummary)
 */
async function claudeAnalyzeWithAnswers(civilData, filename, symbolSummary, ratesSummary) {
  console.log('[claudeAnalyzeWithAnswers] full analysis for', filename);
  const prompt = `Generate complete BOQ for this civil drawing.
FILE: ${filename}
SYMBOL CLASSIFICATION: ${JSON.stringify(symbolSummary)}
RATES: ${ratesSummary || 'Use DSR 2025 rates from system prompt.'}
DXF DATA:
${JSON.stringify(civilData, null, 2)}

Return ONLY raw JSON BOQ (same schema as claudeAnalyzeDXF).`;
  const raw = await callClaudeAPI({ system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }], maxTokens: 8192 });
  return parseJSON(raw);
}

/**
 * claudeAnalyzeDrawingVision — vision analysis of uploaded image/PDF drawing
 * server.js calls: claudeAnalyzeDrawingVision(files, userText, aiResponse)
 *
 * CRITICAL RULES enforced in prompt:
 *  - Column sizes ONLY from printed schedule table — NEVER invent 400×400, 500×500 etc.
 *  - Steel bars ONLY as printed (8–16Ø range typical) — NEVER add 12–20Ø or 16–25Ø.
 *  - Column schedule and footing schedule are SEPARATE — never mix them.
 *  - Qty from schedule qty column only — do not guess.
 *  - Unreadable cell → output "not legible", never guess.
 */
async function claudeAnalyzeDrawingVision(files, userText, aiResponse) {
  console.log('[claudeAnalyzeDrawingVision] vision analysis, files:', files?.length);
  const kb = knowledgeBaseHints();

  const imageParts = buildImageParts(files);
  if (!imageParts.length) throw new Error('No image/PDF files provided for vision analysis');

  const strictScheduleRules = `
═══════════════════════════════════════════════════
COLUMN / FOOTING SCHEDULE READING — ABSOLUTE RULES
═══════════════════════════════════════════════════
1. ONLY read values that are PHYSICALLY PRINTED in the schedule table cells.
2. Column sizes: copy the EXACT printed value (e.g. 300×300, 230×450).
   - NEVER output 400×400, 450×450, 500×500 unless those numbers appear in the drawing.
3. Main bars / stirrups: copy EXACTLY (e.g. 8-12Ø, 4-16Ø).
   - NEVER output 12-20Ø, 16-25Ø or any bar size not printed in the schedule.
4. Column schedule and Footing schedule are COMPLETELY SEPARATE tables.
   - NEVER copy footing sizes/steel into the column table or vice versa.
5. Qty (number of columns): use ONLY the qty column in the schedule.
   - NEVER guess or count columns from the plan view.
6. If any cell is unclear / not legible → write "not legible" — do NOT guess.
7. Source field must be "drawing-schedule" for schedule values, "calculated" for derived values.
═══════════════════════════════════════════════════`;

  const content = [
    ...imageParts,
    {
      type: 'text',
      text: `${kb ? kb + '\n\n' : ''}${strictScheduleRules}

USER QUESTION: ${userText || 'Analyse this drawing and generate BOQ.'}
${aiResponse ? `\nPREVIOUS AI RESPONSE CONTEXT:\n${aiResponse}` : ''}

Return ONLY raw JSON:
{
  "drawing_type": "",
  "column_schedule": [
    {
      "col_mark": "",
      "size_mm": "",
      "main_bars": "",
      "stirrups": "",
      "qty": 0,
      "floor": "",
      "source": "drawing-schedule|not legible"
    }
  ],
  "footing_schedule": [
    {
      "footing_mark": "",
      "size_mm": "",
      "depth_mm": "",
      "main_bars": "",
      "qty": 0,
      "source": "drawing-schedule|not legible"
    }
  ],
  "boq": [
    { "sr": 1, "part": "PART A", "description": "", "unit": "", "qty": 0, "rate": 0, "amount": 0, "source": "drawing-schedule|calculated", "confidence": "high|medium|low" }
  ],
  "cost_summary": { "civil_total_inr": 0, "civil_total_lacs": 0 },
  "observations": [],
  "not_legible_fields": []
}`,
    },
  ];

  const raw = await callClaudeAPI({
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content }],
    maxTokens: 8192,
  });
  return parseJSON(raw);
}

/**
 * claudeAnalyzeDWGVision — analyse DWG converted to PNG tiles
 * server.js calls: claudeAnalyzeDWGVision(pngTiles, converterResult, filename)
 */
async function claudeAnalyzeDWGVision(pngTiles, converterResult, filename) {
  console.log('[claudeAnalyzeDWGVision] analysing DWG tiles for', filename, '| tiles:', pngTiles?.length);
  const kb = knowledgeBaseHints();

  const tileFiles = (pngTiles || []).map(b64 => ({ type: 'image/png', b64 }));
  const imageParts = buildImageParts(tileFiles);
  if (!imageParts.length) throw new Error('No PNG tiles provided for DWG vision analysis');

  const prompt = `${kb ? kb + '\n\n' : ''}FILE: ${filename}
CONVERTER INFO: ${JSON.stringify(converterResult?.summary || {})}

Analyse all tiles together as one drawing. Apply the same strict schedule-reading rules:
- Copy column/footing sizes EXACTLY as printed. Never invent sizes.
- Copy steel bar details EXACTLY. Never invent bar diameters.
- Keep column schedule and footing schedule separate.

Return ONLY raw JSON BOQ (same schema as claudeAnalyzeDrawingVision).`;

  const raw = await callClaudeAPI({
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: [...imageParts, { type: 'text', text: prompt }] }],
    maxTokens: 8192,
  });
  return parseJSON(raw);
}

// ─── Exports ──────────────────────────────────────────────────────────────────
module.exports = {
  // Original exports
  geminiAnalyzeDrawing,
  runCVAnalysis,
  calcRoadQuantities,
  calcStructureQuantities,
  RATES,
  // New exports required by server.js
  callClaudeAPI,
  CIVIL_SYSTEM,
  parseJSON,
  claudeAnalyzeDXF,
  claudeClassifySymbols,
  claudeAnalyzeWithAnswers,
  claudeAnalyzeDrawingVision,
  claudeAnalyzeDWGVision,
};
