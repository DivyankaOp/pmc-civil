'use strict';
/**
 * PMC Civil — Claude Analyzer
 * Replaces all remaining Gemini calls with Claude API
 * Gives consistent 5-phase pipeline across ALL routes
 */

const CLAUDE_URL = 'https://api.anthropic.com/v1/messages';

/**
 * Call Claude with retry on overload
 */
async function callClaudeAPI({ system, messages, maxTokens = 8192, thinking = false, thinkingBudget = 10000 }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');

  const body = {
    model: 'claude-sonnet-4-5',
    max_tokens: thinking ? Math.max(maxTokens, thinkingBudget + 4000) : maxTokens,
    system,
    messages,
  };
  if (thinking) body.thinking = { type: 'enabled', budget_tokens: thinkingBudget };

  for (let i = 0; i <= 4; i++) {
    const r = await fetch(CLAUDE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify(body),
    });
    const data = await r.json();
    if (r.ok && data.content) {
      return data.content.filter(b => b.type === 'text').map(b => b.text).join('');
    }
    if (data.error?.type !== 'overloaded_error') throw new Error(`Claude: ${data.error?.message}`);
    await new Promise(res => setTimeout(res, 2000 * Math.pow(2, i)));
  }
  throw new Error('Claude API: max retries');
}

function parseJSON(raw) {
  if (!raw) return {};
  const clean = raw.replace(/```json|```/g, '').trim();
  const fb = clean.indexOf('{'), lb = clean.lastIndexOf('}');
  if (fb === -1 || lb === -1) return {};
  try { return JSON.parse(clean.slice(fb, lb + 1)); } catch { return {}; }
}

const CIVIL_SYSTEM = `You are a senior PMC civil engineer with 20 years experience in Gujarat, India.
You analyze civil engineering drawings and generate accurate BOQ (Bill of Quantities).
GOLDEN RULES:
1. NEVER invent or guess dimensions — only values visible in THIS drawing.
2. Read the legend FIRST before counting any element.
3. Apply scale factor from title block to ALL measurements.
4. Mark every quantity source: "drawing" | "calculated" | "assumed".
5. Return ONLY raw JSON. No markdown. No explanation.
Gujarat DSR 2025 Rates:
100mm block wall: Rs.4200/cum | 230mm brick wall: Rs.4800/cum
RCC M25: Rs.5500/cum | RCC M30: Rs.5800/cum | Steel Fe500: Rs.56/kg
Excavation: Rs.180/cum | Formwork: Rs.180/sqmt | PQC road: Rs.1800/sqmt
Plaster 12mm: Rs.280/sqmt | Waterproofing: Rs.450/sqmt`;

/**
 * Analyze DXF data with Claude (replaces Gemini in /export-dxf-excel, /classify-dxf, /analyze-with-answers)
 */
async function claudeAnalyzeDXF(civilData, filename, ratesSummary) {
  const prompt = `PMC civil engineer. Analyze DXF. Use ONLY data below, no invented values.
FILE:${filename} DECLARED_TYPE:${civilData.drawing_type} SCALE:${civilData.scale || '?'}
TEXTS:${(civilData.all_texts || []).slice(0, 120).join(' | ')}
DIMS:${(civilData.dimension_values || []).slice(0, 40).map(d => d.value_m + 'm[' + d.layer + ']').join(', ')}
AREAS:${(civilData.polyline_areas || []).slice(0, 20).map(p => p.area_sqm + 'sqm(' + p.layer + ')').join(', ')}
BLOCK_COUNTS:${Object.entries(civilData.block_counts || {}).slice(0, 40).map(([k, v]) => k + ':' + v).join(', ')}
ELEMENT_COUNTS:${JSON.stringify(civilData.element_counts || {})}
LAYERS:${(civilData.layer_names || []).join(', ')}
RATES:${ratesSummary}
Return ONLY JSON:{"project_name":"","drawing_type":"FLOOR_PLAN|BASEMENT|STRUCTURAL|FOUNDATION|SITE_LAYOUT|ROAD_PLAN|ELEVATION","scale":"","spaces":[],"boq":[{"description":"","unit":"","qty":0,"rate":0,"amount":0}],"observations":[],"pmc_recommendation":""}`;

  const raw = await callClaudeAPI({ system: CIVIL_SYSTEM, messages: [{ role: 'user', content: prompt }] });
  return parseJSON(raw);
}

/**
 * Classify unknown DXF blocks/layers with Claude (replaces Gemini in /classify-dxf)
 */
async function claudeClassifySymbols(unknownBlocks, unknownLayers, civilData, filename) {
  if (!unknownBlocks.length && !unknownLayers.length) return { blocks: {}, layers: {}, still_unknown_blocks: [], still_unknown_layers: [] };

  const prompt = `You are a senior AutoCAD civil drawing expert.
Classify these unknown block names and layer names from a civil DXF drawing.

UNKNOWN BLOCKS (name → count):
${unknownBlocks.map(b => `${b.name} (×${b.count})`).join('\n') || 'none'}

UNKNOWN LAYERS:
${unknownLayers.join('\n') || 'none'}

DRAWING CONTEXT:
- File: ${filename}
- Drawing type: ${civilData.drawing_type}
- Texts found: ${(civilData.all_texts || []).slice(0, 30).join(', ')}

Classify as: door | window | column | beam | slab | wall | staircase | lift | ramp | toilet | kitchen | bedroom | parking | road | hatch | dimension | text | furniture | equipment | unknown

Return ONLY raw JSON:
{"blocks":{"BLOCK_NAME":"type"},"layers":{"LAYER_NAME":"type"},"still_unknown_blocks":[],"still_unknown_layers":[]}`;

  const raw = await callClaudeAPI({ system: CIVIL_SYSTEM, messages: [{ role: 'user', content: prompt }] });
  return parseJSON(raw);
}

/**
 * Full BOQ from DXF with symbol answers (replaces Gemini in /analyze-with-answers)
 */
async function claudeAnalyzeWithAnswers(civilData, filename, symbolSummary, ratesSummary) {
  const prompt = `You are a senior PMC civil engineer generating a complete BOQ.
ALL DATA IS FROM THIS DXF FILE. DO NOT INVENT VALUES.

FILE: ${filename}
DRAWING TYPE: ${civilData.drawing_type}
SCALE: ${civilData.scale || 'not detected'}
DRAWING SIZE: ${civilData.drawing_extents?.width_m || 0}m × ${civilData.drawing_extents?.height_m || 0}m

SYMBOL DICTIONARY (confirmed):
${symbolSummary || 'none'}

ELEMENT COUNTS: ${JSON.stringify(civilData.element_counts || {})}
FLOOR LEVELS: ${(civilData.floor_levels || []).map(l => l.label + '=' + (l.level_m || '?') + 'm').join('\n') || 'none'}
TEXT ANNOTATIONS: ${(civilData.all_texts || []).slice(0, 100).join('\n')}
DIMENSIONS: ${(civilData.dimension_values || []).slice(0, 40).map(d => d.value_m + 'm[' + d.layer + ']').join(', ')}
AREAS: ${(civilData.polyline_areas || []).slice(0, 20).map(p => p.area_sqm + 'sqm(' + p.layer + ')').join(', ')}
GUJARAT DSR 2025 RATES: ${ratesSummary}

Return ONLY raw JSON:
{"project_name":"","drawing_type":"","scale":"","building_height_m":0,"floor_count":0,"total_bua_sqm":0,"spaces":[{"name":"","area_sqm":0}],"boq":[{"sr":1,"description":"","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed"}],"element_counts":{},"observations":[],"pmc_recommendation":""}`;

  const raw = await callClaudeAPI({ system: CIVIL_SYSTEM, messages: [{ role: 'user', content: prompt }], maxTokens: 8192 });
  return parseJSON(raw);
}

/**
 * Analyze drawing image/PDF with Claude Vision (replaces Gemini in /gemini, /export-drawing, /drawing-to-excel)
 */
async function claudeAnalyzeDrawingVision(files, userText, aiResponse) {
  const contentParts = [];

  for (const f of (files || [])) {
    try {
      if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i)) {
        contentParts.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.b64 } });
      } else if (f.type?.startsWith('image/')) {
        contentParts.push({ type: 'image', source: { type: 'base64', media_type: f.type || 'image/png', data: f.b64 } });
      }
    } catch (e) {}
  }

  const textPrompt = `You are a senior PMC civil engineer. Analyze this civil drawing and generate a complete BOQ.

STEP 1: Read legend/symbol table — identify all hatch patterns, their labels, layers.
STEP 2: Read title block — project name, drawing number, scale, date.
STEP 3: Identify drawing type (SECTION/ELEVATION/FLOOR_PLAN/STRUCTURAL/SITE_PLAN/FOUNDATION) and all floor levels.
STEP 4: Extract ALL quantities — use scale factor for all measurements.
STEP 5: Calculate BOQ with Gujarat DSR 2025 rates.
STEP 6: PMC observations — IS code compliance, missing info, recommendations.

CRITICAL: Do NOT invent values. If not visible, write "Not shown in drawing".
${userText ? '\nUser note: ' + userText : ''}
${aiResponse ? '\nPrevious analysis: ' + aiResponse : ''}

Return comprehensive analysis in markdown format with tables.`;

  contentParts.push({ type: 'text', text: textPrompt });

  const raw = await callClaudeAPI({
    system: CIVIL_SYSTEM,
    messages: [{ role: 'user', content: contentParts }],
    maxTokens: 8192
  });
  return raw;
}

/**
 * Analyze DWG file rendered as PNG tiles with Claude Vision
 * (replaces Gemini in /analyze-dwg — ZWCAD compatible)
 * v5: Multi-sheet layout support + ZWCAD text-only mode
 */
async function claudeAnalyzeDWGVision(pngB64Array, converterResult, filename) {
  const contentParts = [];

  // Add all PNG tiles as images
  for (const pngB64 of pngB64Array) {
    contentParts.push({ type: 'image', source: { type: 'base64', media_type: 'image/png', data: pngB64 } });
  }

  const textSummary = (converterResult.texts || []).map(t => t.text).slice(0, 200).join(' | ');
  const dimSummary = (converterResult.dimensions || []).filter(d => d.value || d.text)
    .map(d => `${d.value || ''}${d.text ? ' (' + d.text + ')' : ''}`).slice(0, 80).join(', ');
  const layers = (converterResult.layers || []).join(', ');
  const version = converterResult.binary_extract?.version || converterResult.version || 'Unknown';
  const sheets = (converterResult.sheets || []).join(', ');
  const xrefs = (converterResult.xrefs || []).join(', ');
  const isTextMode = converterResult.zwcad_text_mode === true;
  const layoutImages = converterResult.layout_images || [];

  // Build sheet context
  const sheetContext = sheets
    ? `SHEETS/LAYOUTS IN DRAWING: ${sheets}\nNote: Multi-sheet drawing — analyze each sheet separately if data varies.`
    : '';
  const xrefContext = xrefs
    ? `XREF FILES REFERENCED: ${xrefs}\nNote: XREF content is shown inline in the drawing.`
    : '';

  // For text-only mode (ZWCAD render failed), use a richer text-analysis prompt
  const zwcadTextModeNote = isTextMode
    ? `⚠️  PNG RENDER UNAVAILABLE (ZWCAD binary format incompatible with ezdxf).
Text + layer data extracted directly from binary — use this as your drawing data.
Analyze like an engineer reading a drawing schedule: infer elements from text annotations,
layer names, and dimension values. State "Estimated from text" for each quantity.
For accurate visual analysis, user should export to PDF/PNG from ZWCAD first.`
    : '';

  const multiSheetNote = layoutImages.length > 1
    ? `MULTI-SHEET DRAWING: ${layoutImages.length} sheets rendered — ${layoutImages.map(l => l.name).join(', ')}.
Images in this message: (1) Main sheet, then additional sheets. Analyze ALL sheets.`
    : (pngB64Array.length > 1
        ? `MULTI-IMAGE: ${pngB64Array.length} tiles — full sheet + ${pngB64Array.length - 1} zoom crops. Synthesize ONE analysis.`
        : '');

  const textPrompt = `You are a SENIOR PMC CIVIL ENGINEER analyzing a ZWCAD/AutoCAD DWG drawing.
FILE: ${filename} | DWG VERSION: ${version}
${zwcadTextModeNote}
LAYERS: ${layers || 'See image'}
ALL TEXT IN DRAWING: ${textSummary || 'See image'}
DIMENSIONS: ${dimSummary || 'See image'}
${sheetContext}
${xrefContext}
${converterResult.errors ? 'Render notes: ' + (Array.isArray(converterResult.errors) ? converterResult.errors.join(' | ') : converterResult.errors) : ''}
${multiSheetNote}

══════════════════════════════════════════════════════
STEP 1 — READ LEGEND / SYMBOL TABLE
══════════════════════════════════════════════════════
Find legend box. Read every symbol/hatch + label (e.g. "230MM THK. BRICK WALL").
Map each hatch/color/pattern → material meaning.
Note AutoCAD LAYER for each element type.

══════════════════════════════════════════════════════
STEP 2 — READ TITLE BLOCK
══════════════════════════════════════════════════════
Project name, drawing no., scale, date, engineer.
If not visible: "Not shown in drawing" — do NOT invent.

══════════════════════════════════════════════════════
STEP 3 — DRAWING TYPE + FLOOR LEVELS
══════════════════════════════════════════════════════
Type: SECTION / ELEVATION / FLOOR_PLAN / STRUCTURAL / SITE_PLAN / FOUNDATION
Read every floor level annotation (e.g. "+7590 MM LEVEL").
Calculate floor heights between levels.

══════════════════════════════════════════════════════
STEP 4 — EXTRACT QUANTITIES
══════════════════════════════════════════════════════
Use legend from Step 1 for element identification.
Apply scale factor from title block to ALL measurements.

SECTION: wall length × thickness × floor height = volume
FLOOR PLAN: room areas, wall lengths, opening counts
STRUCTURAL: column sizes, beam dimensions, slab thickness
SITE/ROAD: road lengths × widths

══════════════════════════════════════════════════════
STEP 5 — BOQ WITH GUJARAT DSR 2025 RATES
══════════════════════════════════════════════════════
100mm block: ₹4200/cum | 230mm brick: ₹4800/cum
RCC M25: ₹5500/cum | RCC M30: ₹5800/cum | Steel Fe500: ₹56/kg
Excavation: ₹180/cum | Formwork: ₹180/sqmt | PQC road: ₹1800/sqmt

══════════════════════════════════════════════════════
STEP 6 — PMC OBSERVATIONS
══════════════════════════════════════════════════════
IS code compliance, design comments, missing information.

CRITICAL: Only values VISIBLE in drawing. Never invent.`;

  contentParts.push({ type: 'text', text: textPrompt });

  const raw = await callClaudeAPI({
    system: CIVIL_SYSTEM,
    messages: [{ role: 'user', content: contentParts }],
    maxTokens: 8192
  });
  return raw;
}

module.exports = {
  callClaudeAPI,
  claudeAnalyzeDXF,
  claudeClassifySymbols,
  claudeAnalyzeWithAnswers,
  claudeAnalyzeDrawingVision,
  claudeAnalyzeDWGVision,
  parseJSON,
  CIVIL_SYSTEM,
};
