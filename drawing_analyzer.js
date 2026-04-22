/**
 * PMC Drawing Analyzer
 * Combines OpenCV (Python) + Gemini Vision for pixel-accurate analysis
 */
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

// ── GUJARAT DSR RATES 2025 ──────────────────────────────────────────
const RATES = {
  // ROADS
  pqc_road_250mm_sqmt:     1800,
  gsb_300mm_sqmt:           655,
  wmm_200mm_sqmt:           515,
  soil_stabilization_sqmt:   82,
  soil_filling_cum:          285,
  asphalt_60mm_sqmt:         750,
  paver_block_80mm_sqmt:     750,
  service_corridor_sqmt:    1790,
  kerbing_rmt:               350,
  // STRUCTURE  
  rcc_m20_cum:              5200,
  rcc_m25_cum:              5500,
  rcc_m30_cum:              5800,
  pcc_m10_cum:              3800,
  brickwork_230mm_cum:      4500,
  brickwork_115mm_cum:      4200,
  plaster_15mm_sqmt:         120,
  steel_fe500_kg:             56,
  steel_fe415_kg:             54,
  formwork_sqmt:             180,
  // CIVIL
  excavation_cum:            180,
  backfilling_cum:           120,
  compound_wall_rmt:        8600,
  gabion_wall_rmt:         14100,
  // MEP
  streetlight_nos:         35000,
  pipeline_rmt:             4500,
  borewell_nos:            75000,
};

// ── QUANTITY FORMULAS ──────────────────────────────────────────────
function calcRoadQuantities(length_m, width_m, layers = {}) {
  const carriageWidth = layers.carriageWidth || (width_m - 3);
  const area = length_m * carriageWidth;
  const boxCut = area * 1.05; // 5% extra
  
  return {
    area_sqmt:        Math.round(area * 100) / 100,
    box_cutting_sqmt: Math.round(boxCut * 100) / 100,
    gsb_300mm_ton:    Math.round(area * 1.15 * 0.300 * 1.8 * 100) / 100, // 15% compaction
    wmm_200mm_ton:    Math.round(area * 1.15 * 0.200 * 2.1 * 100) / 100,
    pqc_250mm_cum:    Math.round(area * 1.05 * 0.250 * 100) / 100,       // 5% wastage
    steel_dowel_ton:  Math.round(area * 0.00387 * 100) / 100,            // ~3.87 kg/sqmt
    cost_estimate: {
      gsb:           Math.round(area * RATES.gsb_300mm_sqmt),
      wmm:           Math.round(area * RATES.wmm_200mm_sqmt),
      pqc:           Math.round(area * RATES.pqc_road_250mm_sqmt),
      total_sqmt:    Math.round(area * (RATES.gsb_300mm_sqmt + RATES.wmm_200mm_sqmt + RATES.pqc_road_250mm_sqmt)),
    }
  };
}

function calcStructureQuantities(dims) {
  const { length = 0, width = 0, height = 0, nos = 1, thickness = 0 } = dims;
  const volume = length * width * height * nos;
  const area = length * width * nos;
  
  return {
    volume_cum:  Math.round(volume * 1000) / 1000,
    area_sqmt:   Math.round(area * 100) / 100,
    steel_kg:    Math.round(volume * 120), // ~120 kg/CUM for beams/slabs
    formwork_sqmt: Math.round((2*(length+width)*height + area) * nos * 100) / 100,
  };
}

// ── RUN PYTHON CV ──────────────────────────────────────────────────
function runCVAnalysis(b64Image) {
  try {
    const tmpFile = path.join(os.tmpdir(), `drawing_cv_${Date.now()}.txt`);
    fs.writeFileSync(tmpFile, b64Image);
    const scriptPath = path.join(__dirname, 'drawing_cv.py');
    const result = execSync(`python3 ${scriptPath} ${tmpFile}`, { timeout: 30000 });
    fs.unlinkSync(tmpFile);
    return JSON.parse(result.toString());
  } catch (e) {
    console.error('CV analysis failed:', e.message);
    return { error: e.message };
  }
}

// ── GEMINI DRAWING ANALYSIS ────────────────────────────────────────
async function geminiAnalyzeDrawing(key, files, cvData, fetch) {
  const GEMINI_URL = k => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;
  
  const parts = [];
  
  // Add all images/PDFs
  for (const f of (files || [])) {
    if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i))
      parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
    else if (f.type?.startsWith('image/'))
      parts.push({ inline_data: { mime_type: f.type || 'image/png', data: f.b64 } });
  }

  const cvHints = cvData && !cvData.error ? `
COMPUTER VISION PRE-ANALYSIS:
- Image size: ${cvData.image_dimensions?.width_px}×${cvData.image_dimensions?.height_px} px
- Detected ${cvData.detected_spaces?.length || 0} closed spaces/rooms
- Scale bar pixel candidates: ${JSON.stringify(cvData.scale_bar_candidates_px?.slice(0,5))}
- Scale hints: ${cvData.scale_interpretation_hints?.join(' | ')}
- Dimension lines found: ${cvData.dimension_lines?.horizontal?.length || 0} horizontal, ${cvData.dimension_lines?.vertical?.length || 0} vertical
Use these pixel measurements to verify/cross-check your dimension reading.
` : '';

  const prompt = `You are a SENIOR PMC CIVIL ENGINEER with 20 years experience in India. Read THIS drawing only.

${cvHints}

CRITICAL RULES (NEVER BREAK):
1. Do NOT invent, guess, estimate, or copy numbers from prior drawings or examples.
2. If a value is not visibly annotated or directly derivable from THIS drawing, set it to 0 / "" / [].
3. The JSON schema below uses placeholder zeros — DO NOT copy them; fill with values YOU read from this drawing.
4. Every dimension you report must trace back to a specific annotation or scaled line measurement in the drawing.
5. Set confidence to LOW if you could not read the title block / scale / key dimensions.

EXTRACTION CHECKLIST:
1. SCALE — look for "1:100", "1:500", scale bar. If nothing readable, leave empty.
2. DIMENSIONS — every number with arrows/extension lines.
3. TEXT — road names, room labels, material callouts.
4. TITLE BLOCK — project name, drawing no, date, north arrow.
5. ELEMENT COUNTS — count doors, windows, lifts, staircases, columns, footings visible on drawing.

FORMULAS (only use dimensions actually read from drawing):
ROAD: area=L×Wcarriage | GSB(t)=area×1.15×0.3×1.8 | WMM(t)=area×1.15×0.2×2.1 | PQC(cum)=area×1.05×0.25 | DowelSteel(kg)=area×3.87
BUILDING: wall=L×H×T (cum brick) | slab=A×T (cum RCC) | steel=RCC×120 (slab) / ×160 (beam)
FOOTING: L×B×D (cum). Column: S×S×H (cum). Steel from BBS only if BBS shown.

Return ONLY raw JSON (no markdown):
{
  "project_name": "",
  "drawing_no": "",
  "drawing_type": "",
  "scale": "",
  "date": "",
  "north_direction": "",

  "elements": [
    {
      "id": "",
      "type": "",
      "name": "",
      "dimensions": { "length_m": 0, "width_m": 0, "height_m": 0, "thickness_m": 0, "nos": 0, "note": "" },
      "quantities":  { "area_sqmt": 0, "volume_cum": 0, "gsb_ton": 0, "wmm_ton": 0, "pqc_cum": 0, "steel_kg": 0, "brickwork_cum": 0 },
      "cost_inr":    { "material": 0, "labour": 0, "total": 0 },
      "confidence": "",
      "annotation_found": ""
    }
  ],

  "element_counts": { "door_count": 0, "window_count": 0, "lift_count": 0, "staircase_count": 0, "column_count": 0, "footing_count": 0, "bedroom_count": 0, "toilet_count": 0, "kitchen_count": 0, "floor_count": 0 },

  "total_quantities": { "total_area_sqmt": 0, "total_road_rmt": 0, "gsb_total_ton": 0, "wmm_total_ton": 0, "pqc_total_cum": 0, "rcc_total_cum": 0, "steel_total_kg": 0, "brickwork_total_cum": 0 },

  "cost_summary": { "civil_total_inr": 0, "civil_total_lacs": 0, "civil_total_crores": 0, "item_wise": [] },

  "bbs_data": [],

  "observations": [],
  "pmc_recommendation": "",
  "extraction_confidence": "LOW|MEDIUM|HIGH",
  "missing_info": []
}`;

  parts.push({ text: prompt });
  
  const r = await fetch(GEMINI_URL(key), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ role: 'user', parts }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' }
    })
  });
  
  let raw = (await r.json())?.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
  if (fb !== -1 && lb !== -1) raw = raw.slice(fb, lb + 1);
  
  try {
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    // Apply our formula calculations on top of Gemini data
    return enrichWithCalculations(parsed);
  } catch (e) {
    console.error('Parse fail:', e.message, raw.slice(0, 200));
    return null;
  }
}

// ── ENRICH: Apply exact formulas to Gemini extracted dimensions ────
function enrichWithCalculations(data) {
  if (!data?.elements) return data;
  
  let totalAreaSqmt = 0;
  let totalRoadRmt = 0;
  let totalGSBTon = 0;
  let totalWMMTon = 0;
  let totalPQCCum = 0;
  let totalSteelKg = 0;
  let totalCostInr = 0;

  data.elements = data.elements.map(el => {
    const d = el.dimensions || {};
    
    if (el.type === 'ROAD') {
      const L = d.length_m || 0;
      const W = d.width_m || 0;
      const carriageW = d.carriage_width_m || Math.max(W - 3, W * 0.65);
      const q = calcRoadQuantities(L, carriageW);
      
      el.quantities = {
        ...el.quantities,
        area_sqmt:        q.area_sqmt,
        box_cutting_sqmt: q.box_cutting_sqmt,
        gsb_300mm_ton:    q.gsb_300mm_ton,
        wmm_200mm_ton:    q.wmm_200mm_ton,
        pqc_250mm_cum:    q.pqc_250mm_cum,
        steel_dowel_kg:   Math.round(q.steel_dowel_ton * 1000),
      };
      el.cost_inr = {
        gsb:   q.cost_estimate.gsb,
        wmm:   q.cost_estimate.wmm,
        pqc:   q.cost_estimate.pqc,
        total: q.cost_estimate.total_sqmt,
        per_sqmt: RATES.gsb_300mm_sqmt + RATES.wmm_200mm_sqmt + RATES.pqc_road_250mm_sqmt
      };
      
      totalAreaSqmt  += q.area_sqmt;
      totalRoadRmt   += L;
      totalGSBTon    += q.gsb_300mm_ton;
      totalWMMTon    += q.wmm_200mm_ton;
      totalPQCCum    += q.pqc_250mm_cum;
      totalSteelKg   += Math.round(q.steel_dowel_ton * 1000);
      totalCostInr   += q.cost_estimate.total_sqmt;
    }
    
    return el;
  });

  // Update totals
  data.total_quantities = {
    ...data.total_quantities,
    total_area_sqmt:   Math.round(totalAreaSqmt * 100) / 100,
    total_road_rmt:    Math.round(totalRoadRmt * 100) / 100,
    gsb_total_ton:     Math.round(totalGSBTon * 100) / 100,
    wmm_total_ton:     Math.round(totalWMMTon * 100) / 100,
    pqc_total_cum:     Math.round(totalPQCCum * 100) / 100,
    steel_total_kg:    totalSteelKg,
    calculation_note:  'Quantities re-calculated using PMC formula engine on Gemini-extracted dimensions'
  };

  const totalLacs = Math.round(totalCostInr / 100000 * 100) / 100;
  data.cost_summary = {
    ...data.cost_summary,
    civil_total_inr:    totalCostInr,
    civil_total_lacs:   totalLacs,
    civil_total_crores: Math.round(totalLacs / 100 * 100) / 100,
  };

  data.rates_applied = RATES;
  return data;
}

module.exports = { geminiAnalyzeDrawing, runCVAnalysis, calcRoadQuantities, calcStructureQuantities, RATES };
