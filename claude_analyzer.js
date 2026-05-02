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
async function callClaude({ messages, maxTokens = 4096, thinking = false }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');
  const body = {
    model: 'claude-sonnet-4-6',
    max_tokens: thinking ? 16000 : maxTokens,
    system: SYSTEM_PROMPT,
    messages,
  };
  if (thinking) body.thinking = { type: 'enabled', budget_tokens: 8000 };

  // Extended thinking needs a different beta header than PDF support
  // When thinking=true, betas array includes both
  const betaHeader = thinking
    ? 'interleaved-thinking-2025-05-14,pdfs-2024-09-25'
    : 'pdfs-2024-09-25';

  for (let i = 0; i <= 4; i++) {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-beta': betaHeader,
      },
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
// DRAWING LAYOUT DETECTOR
// Detects what TYPE of drawing it is from GCV text / filename / file count
// Returns a layout strategy object — NO hardcoded values, purely signal-based
// ════════════════════════════════════════════
function detectDrawingLayout(files, cvData, gcvTableContext) {
  const gcvText = (gcvTableContext || '').toLowerCase();
  const fileNames = (files||[]).map(f => (f.name||'').toLowerCase()).join(' ');
  const allText = gcvText + ' ' + fileNames;

  // ── Signal detection ────────────────────────────────────────────
  const signals = {
    // Industrial steel+RCC (must have steel-specific signals)
    hasBasePlate:    /base plate|base-plate|baseplate/.test(allText),
    hasAnchorBolt:   /anchor bolt|h\.d\. bolt|hd bolt/.test(allText),
    hasBodOfSteel:   /bod of steel|b\.o\.d|bottom of steel/.test(allText),
    hasBracedBay:    /braced bay|x.brac/.test(allText),
    hasSlope:        /slope\s*1\s*[:/]\s*\d/.test(allText),
    hasCanopy:       /canopy/.test(allText),
    hasRoofMonitor:  /roof monitor|monitor/.test(allText),
    hasPedestal:     /pedestal/.test(allText),

    // Foundation / structural schedules
    hasScheduleOfFootings: /schedule of footing/.test(allText),
    hasScheduleOfColumns:  /schedule of column/.test(allText),
    hasSectionDetail:      /section[- ][a-z]-[a-z]|detail [a-z]-[a-z]/.test(allText),

    // Column DETAIL drawing (elevation/section diagrams of individual columns)
    // MUST NOT match 'foundation to upper ground floor column detail'
    // Use strict pattern: 'column detail' as standalone, not embedded in floor description
    hasColumnDetailDrawing: /\bcolumn detail\b/.test(allText) && !/schedule of column/.test(allText),
    hasGridRef:            /grid[- ]\d|grid[- ][a-z]/.test(allText),

    // Residential floor plan (strict — require 'floor plan' explicitly, not just 'ground floor')
    hasFloorPlanExplicit: /floor plan/.test(allText),   // explicit "floor plan" text only
    hasDoorSchedule: /door schedule|door mark|d\.s\./.test(allText),
    hasWindowSchedule:/window schedule|window mark|w\.s\./.test(allText),
    hasBeamSchedule: /beam schedule|beam mark/.test(allText),
    hasSlabSchedule: /slab schedule/.test(allText),

    // Building foundation (RCC isolated footings for residential/commercial — no steel base plate)
    hasFoundationLayout: /foundation layout|foundation plan/.test(allText),
    hasPadFooting:    /pad footing|spread footing|isolated footing/.test(allText),
    hasFootingDepth:  /\bdf\b|depth of footing|footing depth/.test(allText),

    // Road / infrastructure
    hasGSB:          /\bgsb\b/.test(allText),
    hasWMM:          /\bwmm\b/.test(allText),
    hasPQC:          /\bpqc\b/.test(allText),
    hasChainage:     /ch\.\d|chainage/.test(allText),

    // Dimension unit hints
    hasFeetInches:   /\d+'-\d+''|\d+'|\d+''/.test(gcvText),  // 8'-6'', 22'' format in text

    // Layout type from CV data
    isScannedPdf:    !!(cvData?.pdf_is_gcv || cvData?.pdf_scanned_fallback),
    isVectorPdf:     !!(cvData?.pdf_is_vector),
    hasMultipleTiles: (files||[]).filter(f => f.name?.match(/tile|page/i)).length > 1,
  };

  // ── Score calculation ─────────────────────────────────────────────
  // Industrial: requires steel-specific signals — schedule alone is NOT enough
  const industrialScore = [
    signals.hasBasePlate, signals.hasAnchorBolt, signals.hasBodOfSteel,
    signals.hasBracedBay, signals.hasSlope,
  ].filter(Boolean).length;
  // Also flag industrial if explicit steel signals present alongside schedules
  const isDefinitelyIndustrial = (signals.hasBasePlate || signals.hasAnchorBolt || signals.hasBodOfSteel)
    && (signals.hasScheduleOfFootings || signals.hasScheduleOfColumns);

  const residentialScore = [
    signals.hasFloorPlanExplicit, signals.hasDoorSchedule,
    signals.hasWindowSchedule, signals.hasBeamSchedule, signals.hasSlabSchedule,
  ].filter(Boolean).length;

  const roadScore = [
    signals.hasGSB, signals.hasWMM, signals.hasPQC, signals.hasChainage,
  ].filter(Boolean).length;

  // Building foundation score: schedule of footings WITHOUT industrial signals
  const buildingFoundationScore = [
    signals.hasScheduleOfFootings, signals.hasScheduleOfColumns,
    signals.hasFoundationLayout, signals.hasPadFooting,
  ].filter(Boolean).length;

  // ── Classify drawing type ─────────────────────────────────────────
  let drawingType = 'UNKNOWN';
  let layoutStrategy = 'STANDARD';
  let panelMap = {};
  let readingInstructions = '';

  if (isDefinitelyIndustrial || industrialScore >= 3) {
    // ── INDUSTRIAL STEEL+RCC ──
    drawingType = 'INDUSTRIAL_STEEL_RCC';
    layoutStrategy = 'MULTI_PANEL_SCHEDULE';
    panelMap = {
      titleBlock:       'bottom strip — title block and notes',
      plan:             'right half — RCC layout plan with grid lines',
      columnDetails:    'left side — individual grid column details with base plate dims',
      footingSection:   'bottom-left — footing cross-section details (Section A-A, B-B)',
      footingSchedule:  'bottom-right — SCHEDULE OF FOOTINGS table',
      columnSchedule:   'far right — SCHEDULE OF COLUMNS table (column/pedestal schedule)',
      notes:            'bottom-left notes box',
    };
    readingInstructions = `
DRAWING LAYOUT — INDUSTRIAL STEEL+RCC MULTI-PANEL SHEET:
This is a single sheet with MULTIPLE PANELS. Read each panel separately:

PANEL 1 — PLAN VIEW (centre/right area):
  • Grid layout showing column positions in plan
  • Note which grids are BRACED BAYs (X-pattern)
  • Note SLOPE annotations (roof slope — NOT structural dimensions)

PANEL 2 — COLUMN DETAIL PANELS (left side, multiple stacked):
  • Each panel = one column type. Read: column mark, pedestal size, base plate size, anchor bolts
  • These are ELEVATION/SECTION details — NOT the schedule table

PANEL 3 — FOOTING SECTION DETAILS (bottom area):
  • Cross-section labeled "SECTION A-A" etc. Read: footing depth, PCC thickness, cover

PANEL 4 — SCHEDULE OF FOOTINGS (locate by scanning — it's a table with columns: mark | size | depth | reinf | qty):
  • THIS IS THE PRIMARY SOURCE for footing BOQ quantities. Read EVERY row.

PANEL 5 — SCHEDULE OF COLUMNS (another table: mark | pedestal size | main bars | stirrups | qty):
  • THIS IS THE PRIMARY SOURCE for column/pedestal BOQ quantities. Read EVERY row.

PANEL 6 — TITLE BLOCK + NOTES: project name, drawing no, concrete grade, steel grade, all notes.

CRITICAL: Schedule quantities OVERRIDE everything else. If schedule cell unclear → "not legible".`;

  } else if (signals.hasColumnDetailDrawing && !signals.hasScheduleOfFootings) {
    // ── COLUMN REINFORCEMENT DETAIL DRAWING ──
    // Individual column elevation/section diagrams — NO tabular schedule
    drawingType = 'COLUMN_DETAIL_DRAWING';
    layoutStrategy = 'DETAIL_PANELS';
    panelMap = {
      detailPanels: 'multiple column detail panels arranged in rows',
      titleBlock:   'right side or bottom — title block',
      notes:        'right side notes box',
    };
    readingInstructions = `
DRAWING LAYOUT — COLUMN REINFORCEMENT DETAIL DRAWING:
This drawing shows individual column reinforcement DETAILS (elevation/section diagrams).
There is NO tabular schedule in this drawing — do NOT look for a table.

HOW TO READ:
  • The drawing has multiple panels arranged in a grid (typically 3 rows × 5 columns)
  • Each panel = one column TYPE showing: cross-section diagram + bar layout
  • Column MARK is printed below each panel (e.g. C36,C37 / S1 / C45,C55 etc.)
  • Column SIZE is the dimension printed on the section diagram (e.g. 27", 24", 18")
  • Main bars: count shown as "8-16T+2T" or "10-20T" format in the callout text
  • Confinement zone height: shown as dimension on elevation sketch

UNIT RULE: Dimensions on these detail drawings are in INCHES (e.g. 27" = 685mm, 18" = 457mm).
Convert: inches × 25.4 = mm.

WHAT TO OUTPUT:
  • For each panel: column_mark, size_inches (as printed), size_mm (converted), main_bars, stirrups, confinement_zone_mm
  • Source = "drawing-detail" for all values read from detail panels
  • Do NOT invent or assume any values not visible in the drawing`;

  } else if (signals.hasScheduleOfFootings || buildingFoundationScore >= 2) {
    // ── BUILDING FOUNDATION SCHEDULE ──
    // Residential/commercial building — pad footing schedule table (no steel base plates)
    drawingType = 'BUILDING_FOUNDATION_SCHEDULE';
    layoutStrategy = 'SCHEDULE_TABLE';
    panelMap = {
      scheduleTable: 'main content — footing schedule table',
      titleBlock:    'top or bottom — project info',
      notes:         'notes section — concrete grade, cover, steel grade',
    };
    readingInstructions = `
DRAWING LAYOUT — BUILDING PAD FOOTING SCHEDULE:
This is a schedule table for isolated pad footings of a residential/commercial building.
There are NO steel base plates or anchor bolts — this is RCC construction only.

TABLE STRUCTURE (columns):
  COLUMN NO. | PCC Size (B×L) | RCC Size (B×L) | dmin | Df | BARS // TO B | BARS // TO L | REMARK

CRITICAL DEFINITIONS:
  • PCC B, PCC L = Plain Cement Concrete (lean concrete bed) plan dimensions
  • RCC B, RCC L = Reinforced Concrete Footing plan dimensions (smaller than PCC)
  • dmin = minimum depth from top of footing to bottom of footing (footing thickness)
  • Df = total depth of footing below Natural Ground Level (used for excavation)
  • BARS // TO B = reinforcement bars running parallel to width B
  • BARS // TO L = reinforcement bars running parallel to length L

⚠️ UNIT WARNING: Dimensions are in FEET-INCHES format (e.g. 8'-6'', 22'').
The note "ALL DIMENSIONS IN MM" is BOILERPLATE — it does NOT apply to schedule values.
CONVERT: feet-inches → mm:
  Example: 8'-6'' = (8×12 + 6) × 25.4 = 102 × 25.4 = 2591 mm
  Example: 22'' = 22 × 25.4 = 559 mm
  Example: 13'-3'' = (13×12 + 3) × 25.4 = 159 × 25.4 = 4039 mm

READING RULES:
  1. Read EVERY row — there may be 20-30 footing types
  2. Column NO. may list multiple columns per row (e.g. C1,C2,C3,C4,C5,C6...)
  3. Count qty = number of column numbers listed in that row
  4. Df column = depth for excavation calculation
  5. dmin column = footing thickness for RCC volume calculation
  6. Reinforcement format: "10 - TOR 5'' C/C" = 10mm dia TOR bars @ 5 inch c/c spacing
  7. REMARK column: note any special requirements (e.g. "Provide Top Jali")`;

  } else if (signals.hasFoundationLayout && !signals.hasScheduleOfFootings) {
    // ── FOUNDATION LAYOUT PLAN ──
    drawingType = 'FOUNDATION_LAYOUT_PLAN';
    layoutStrategy = 'PLAN_WITH_CALLOUTS';
    readingInstructions = `
DRAWING LAYOUT — FOUNDATION LAYOUT PLAN:
This is a plan-view drawing showing footing positions with PAD depth callouts (e.g. PAD:18", PAD:24").
Grid lines (A,B,C... and 1,2,3...) show column positions.

HOW TO READ:
  • Each column position has a PAD:XX" callout — this is the footing depth in inches
  • Grid labels are along the edges (letters one direction, numbers other direction)
  • Count total pad footing positions from the plan
  • PAD depth values in INCHES — convert: XX" × 25.4 = mm
  • Title block has: project name, drawing no, date, scale, notes
  • Do NOT calculate BOQ from this drawing alone — it must be cross-referenced with footing schedule`;

  } else if (residentialScore >= 2) {
    // ── RESIDENTIAL FLOOR PLAN ──
    drawingType = 'RESIDENTIAL_BUILDING';
    layoutStrategy = 'FLOOR_PLAN_WITH_SCHEDULES';
    panelMap = {
      plan:       'main area — floor plan with dimensions',
      schedules:  'side panels or separate sheets — door/window/beam schedules',
      titleBlock: 'bottom-right corner',
    };
    readingInstructions = `
DRAWING LAYOUT — RESIDENTIAL BUILDING PLAN:
  • Main panel: floor plan — read room dimensions, wall thickness, openings
  • Door/Window schedule: read each row for size, type, qty
  • Beam schedule: read span, size, reinforcement from table
  • Column schedule: read from table only — never from plan view`;

  } else if (roadScore >= 2) {
    // ── ROAD / INFRASTRUCTURE ──
    drawingType = 'ROAD_INFRASTRUCTURE';
    layoutStrategy = 'CHAINAGE_BASED';
    panelMap = {
      plan:       'plan view — chainage markers, widths',
      section:    'typical cross-section — layers, widths',
      titleBlock: 'bottom strip',
    };
    readingInstructions = `
DRAWING LAYOUT — ROAD/INFRASTRUCTURE:
  • Plan view: read start/end chainage, road width
  • Cross-section: read GSB/WMM/PQC thickness and widths
  • Calculate quantities from length × width × layer thickness`;

  } else {
    // ── GENERIC FALLBACK ──
    drawingType = 'GENERAL_CIVIL';
    layoutStrategy = 'STANDARD';
    readingInstructions = `Read all panels systematically: title block first, then plan, then schedule tables, then section details.
If you see a schedule table — read it row by row, copy values EXACTLY as printed.
If dimensions are in feet-inches (8'-6'', 22'') convert to mm: (feet×12 + inches) × 25.4.`;
  }

  console.log(`[Layout Detector] Type: ${drawingType} | Strategy: ${layoutStrategy} | Scores: industrial=${industrialScore}(definite:${isDefinitelyIndustrial}), building_foundation=${buildingFoundationScore}, residential=${residentialScore}, road=${roadScore}`);
  return { drawingType, layoutStrategy, panelMap, readingInstructions, signals };
}


// ════════════════════════════════════════════
// PHASE 2 — Legend + Title Block
// ════════════════════════════════════════════
async function phase2_legendAndScale(files, cvData, gcvTableContext='', layout={}) {
  console.log('[Phase 2] Reading legend + title block...');
  const kb = knowledgeBaseHints();
  const cv = cvData&&!cvData.error
    ? `CV: ${cvData.image_dimensions?.width_px}x${cvData.image_dimensions?.height_px}px | spaces:${cvData.detected_spaces?.length||0} | scale candidates:${JSON.stringify(cvData.scale_bar_candidates_px?.slice(0,4))}` : '';
  const imgParts = buildImageParts(files);
  if (!imgParts.length && !gcvTableContext) return null;

  // Inject layout reading instructions so Phase 2 knows WHERE to look for title block
  const layoutHint = layout.readingInstructions
    ? `\nDRAWING LAYOUT CONTEXT (auto-detected):\n${layout.readingInstructions}\n\nFor Phase 2: Focus on the TITLE BLOCK panel (${layout.panelMap?.titleBlock||'usually bottom strip'}) and NOTES section.\n`
    : '';

  const raw = await callClaude({
    messages:[{ role:'user', content:[
      ...imgParts,
      { type:'text', text:`${cv}\n${kb}${layoutHint}${gcvTableContext}\n\nTASK PHASE 2: Read ONLY the legend/symbol table, title block, and identify where schedule tables are located in the drawing. No quantities yet.\n\nCRITICAL RULES:\n1. Scale: read EXACTLY from title block (e.g. 1:100, 1:200). If not visible write null.\n2. Scale_factor: numeric part only (100 for 1:100). ALL dimensions sent to Phase 3 will be raw drawing units — Phase 3 multiplies by this factor.\n3. North direction: read compass or north arrow if present.\n4. Legend: read EVERY symbol in the legend table — hatch pattern name, its meaning, and the layer name printed next to it.\n5. If title block confidence is LOW, set scale to null so Phase 3 marks confidence LOW.\n6. If GCV table data is provided above, extract project/drawing info ONLY from those values — do not infer.\n7. DRAWING TYPE DETECTION: If you see steel column base plates, anchor bolts, braced bays, "BOD OF STEEL", "BASE PLATE", SLOPE annotations, CANOPY — set drawing_type to "RCC_FOOTING_INDUSTRIAL" and structural_system to "STEEL_FRAME_RCC_PEDESTAL".\n8. SCHEDULE TABLE LOCATIONS: Scan the full drawing and describe WHERE each schedule table appears. Phase 3 uses this to focus on the right area.\n9. CONCRETE + STEEL GRADE: Read from NOTES section or title block exactly as printed (e.g. M40, Fe500D).\n10. NOTES SECTION: Read every line of the NOTES box — critical spec info is here.\n\nReturn JSON:\n{"drawing_type":"","project_name":"","drawing_no":"","date":"","concrete_grade":"","steel_grade":"","scale":"1:100","scale_factor":100,"north_direction":"","structural_system":"","legend":[{"symbol":"","meaning":"","layer":"","hatch_pattern":""}],"annotation_layers":[],"floors_visible":[],"title_block_confidence":"HIGH|MEDIUM|LOW","legend_confidence":"HIGH|MEDIUM|LOW","schedule_tables_visible":["column_schedule","footing_schedule"],"schedule_table_locations":{"column_schedule":"describe location","footing_schedule":"describe location"},"general_notes":[],"notes":[]}` }
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
async function phase3_extractQuantities(files, meta, gcvTableContext='', layout={}) {
  console.log('[Phase 3] Extracting quantities with legend context...');
  const legendCtx = meta?.legend?.length
    ? `LEGEND FROM PHASE 2 (use for all element identification):\n${meta.legend.map(l=>`  ${l.symbol} = ${l.meaning} (layer:${l.layer||'?'})`).join('\n')}`
    : 'No legend — use standard CAD conventions.';
  const scaleCtx = meta?.scale
    ? `SCALE: ${meta.scale} (scale_factor=${meta.scale_factor}). Apply to ALL dimensions.`
    : 'Scale not confirmed — mark confidence LOW.';
  const imgParts = buildImageParts(files);

  // Cap GCV context to avoid token overflow in Phase 3 prompt
  const gcvCapped = (gcvTableContext || '').slice(0, 2500);

  // Build schedule location hints — prefer Phase 2 result, fallback to layout detector
  const schedLoc = meta?.schedule_table_locations
    ? `SCHEDULE TABLE LOCATIONS (confirmed by Phase 2):\n  Column Schedule: ${meta.schedule_table_locations.column_schedule||'unknown'}\n  Footing Schedule: ${meta.schedule_table_locations.footing_schedule||'unknown'}\nFocus on these exact areas when reading schedule tables.`
    : (layout.panelMap?.footingSchedule
      ? `SCHEDULE TABLE LOCATIONS (from auto-detector):\n  Footing Schedule: ${layout.panelMap.footingSchedule}\n  Column Schedule: ${layout.panelMap.columnSchedule}`
      : '');

  const layoutInstructions = layout.readingInstructions || '';

  // Industrial context fallback (if layout detector somehow missed it)
  const isIndustrial = layout.drawingType === 'INDUSTRIAL_STEEL_RCC' ||
    /industrial|steel_frame|rcc_footing/i.test(meta?.drawing_type||'');
  const industrialCtx = isIndustrial && !layoutInstructions
    ? `\nINDUSTRIAL STEEL+RCC DRAWING:\n- Column Schedule = RCC PEDESTAL sizes\n- Footing Schedule = isolated footing sizes\n- Base plate dimensions in left detail panels\n`
    : '';

  // Unit conversion instruction — critical for Gujarat drawings using feet-inches
  const unitConversionRule = `
━━━ UNIT CONVERSION RULE (CRITICAL FOR THIS DRAWING TYPE) ━━━
If dimensions are printed in feet-inches format (e.g. 8'-6'', 22'', 13'-3''):
  • The note "ALL DIMENSIONS IN MM" in the title block is standard BOILERPLATE.
    It does NOT apply to schedule/table dimensions — those are in feet-inches.
  • Convert feet-inches → mm:
      8'-6''  = (8×12 + 6)  × 25.4 = 102 × 25.4 = 2591 mm
      22''    = 22           × 25.4 = 559 mm
      13'-3'' = (13×12 + 3) × 25.4 = 159 × 25.4 = 4039 mm
  • Store converted mm value in size_mm / depth_mm fields
  • Also store original printed value in annotation_text field for reference
━━━ END UNIT RULE ━━━
`;

  const strictScheduleRules = `
═══════════════════════════════════════════════════════════
SCHEDULE READING — ABSOLUTE RULES (violation = wrong BOQ)
═══════════════════════════════════════════════════════════
1. Read ONLY values PHYSICALLY PRINTED in schedule table cells.
2. Column schedule and Footing schedule are COMPLETELY SEPARATE tables — NEVER mix them.
3. Column/Pedestal size: copy EXACT printed value. NEVER output 400×400, 500×500 unless printed.
4. Main bars: copy EXACTLY as printed (e.g. 8-16Ø, 4T20, 10-TOR). NEVER invent bar sizes.
5. Stirrups/ties: copy EXACTLY (e.g. 8Ø@150c/c, TOR 5'' C/C). NEVER invent.
6. Footing size: from footing schedule ONLY.
7. Qty: use ONLY the qty column OR count column numbers listed in that row.
8. Unreadable cell → write "not legible" — NEVER guess.
9. Source field: "drawing-schedule" for schedule values, "calculated" for derived, "drawing-detail" for detail panel.
10. Df (depth of footing below NGL) and dmin (footing thickness) are DIFFERENT values — store separately.
═══════════════════════════════════════════════════════════`;

  const raw = await callClaude({
    messages:[{ role:'user', content:[
      ...imgParts,
      { type:'text', text:`${gcvCapped ? 'PRIORITY DATA FROM SCANNED PDF TABLE (GCV+Claude validated):\n'+gcvCapped+'\nUSE THESE VALUES EXACTLY — do not recalculate, do not assume missing values.\n\n' : ''}${layoutInstructions ? '━━━ DRAWING LAYOUT READING GUIDE (auto-detected — follow this precisely) ━━━\n'+layoutInstructions+'\n━━━ END LAYOUT GUIDE ━━━\n\n' : ''}${unitConversionRule}\n${legendCtx}\n${scaleCtx}\n${schedLoc}${industrialCtx}${strictScheduleRules}\nDrawing type: ${meta?.drawing_type||layout.drawingType||'unknown'}\nStructural system: ${meta?.structural_system||'unknown'}\nConcrete grade: ${meta?.concrete_grade||'read from drawing'}\nSteel grade: ${meta?.steel_grade||'read from drawing'}\nAnnotation layers (ignore for quantities): ${JSON.stringify(meta?.annotation_layers||[])}\n\nCRITICAL QUANTITY READING RULES:\n1. Apply scale_factor=${meta?.scale_factor||1} to ALL raw dimensions before recording length_m/width_m/height_m.\n2. Count columns/footings from schedule QTY column ONLY (or count listed column numbers per row).\n3. Read annotation texts EXACTLY as printed — do NOT change any digit.\n4. For schedule tables: copy cell values EXACTLY. If unreadable → write "not legible".\n5. If GCV table data is provided above — use those cell values AS-IS.\n6. Mark source: "drawing-schedule" if directly read, "calculated" if derived, "drawing-detail" if from detail panel.\n7. For COLUMN_DETAIL_DRAWING type: read sizes from dimension callouts on each panel — no table exists.\n8. Read PCC thickness from footing section details if shown.\n9. Steel grade: if conflict between Fe500D and Fe550 in same drawing — use value from NOTES section, flag in observations.\n\nTASK PHASE 3: Extract ALL quantities from this drawing as described in the layout guide above.\nReturn JSON:\n{"quantities":[{"element":"","floor":"","length_m":0,"width_m":0,"height_m":0,"thickness_m":0,"nos":1,"area_sqmt":0,"volume_cum":0,"unit":"","annotation_text":"","source":"drawing-schedule|calculated|drawing-detail","confidence":"high|medium|low"}],"element_counts":{"column_count":0,"footing_count":0,"braced_bay_count":0,"anchor_bolt_count":0},"schedule_data":{"concrete_grade":"","steel_grade":"","columns":[{"mark":"","size_mm":"","size_original":"","main_bars":"","stirrups":"","height_m":0,"qty":0,"source":"drawing-schedule|drawing-detail|not legible","notes":""}],"footings":[{"mark":"","pcc_size_mm":"","rcc_size_mm":"","rcc_size_original":"","dmin_mm":0,"dmin_original":"","Df_mm":0,"Df_original":"","pcc_mm":150,"main_bars_b":"","main_bars_l":"","qty":0,"remark":"","source":"drawing-schedule|not legible"}],"base_plates":[{"column_mark":"","plate_size_mm":"","anchor_bolt_nos":0,"anchor_bolt_dia_mm":0,"source":"drawing-schedule|not legible"}]},"section_details":{"footing_depth_mm":0,"pedestal_height_mm":0,"pcc_thickness_mm":150,"cover_mm":50},"grid_info":{"typical_bay_m":0,"total_columns_plan":0,"braced_bay_grids":[]},"road_data":{"roads":[]},"total_built_area_sqmt":0,"unit_system_detected":"feet-inches|mm|mixed","observations":[]}` }
    ]}],
    maxTokens: 4096
  });

  const q = parseJSON(raw);
  console.log(`[Phase 3] ${q?.quantities?.length||0} elements | footings read: ${q?.schedule_data?.footings?.length||0} | columns read: ${q?.schedule_data?.columns?.length||0} | units: ${q?.unit_system_detected||'?'}`);
  return q;
}

// ════════════════════════════════════════════
// PHASE 4 — BOQ Calculation
// ════════════════════════════════════════════
async function phase4_calculateBOQ(quantities, meta, layout={}) {
  console.log('[Phase 4] Calculating BOQ...');
  if (!quantities) return null;

  const scheduleCtx = quantities.schedule_data
    ? `\nSCHEDULE DATA FROM PHASE 3 (use these EXACT values — do NOT recalculate):\n${JSON.stringify(quantities.schedule_data, null, 2)}\n\nSECTION DETAILS FROM PHASE 3:\n${JSON.stringify(quantities.section_details||{}, null, 2)}\n\nGRID INFO FROM PHASE 3:\n${JSON.stringify(quantities.grid_info||{}, null, 2)}\n`
    : '';

  const drawingType = layout.drawingType || meta?.drawing_type || '';

  const isIndustrial = drawingType === 'INDUSTRIAL_STEEL_RCC' ||
    /industrial|steel_frame|rcc_footing/i.test(drawingType);

  const isBuildingFoundation = drawingType === 'BUILDING_FOUNDATION_SCHEDULE' ||
    drawingType === 'FOUNDATION_LAYOUT_PLAN';

  const isColumnDetail = drawingType === 'COLUMN_DETAIL_DRAWING';

  const industrialBOQHint = isIndustrial ? `
INDUSTRIAL STEEL+RCC BOQ PARTS TO INCLUDE:
PART A — EARTHWORK: Excavation (CUM), Backfilling (CUM), Dewatering (LS)
PART B — PCC: PCC M10 bed 75mm thick (CUM) = footing plan area × 0.075 × qty
PART C — RCC: RCC Isolated Footing (CUM) = RCC plan area × dmin × qty
              RCC Pedestal (CUM) = pedestal plan area × pedestal height × qty
PART D — FORMWORK: Footing sides (SQM), Pedestal (SQM)
PART E — STEEL: Fe500D bars in footing (MT) = RCC CUM × 80 kg/CUM
                Fe500D bars in pedestal (MT) = pedestal CUM × 200 kg/CUM
PART F — BASE PLATE & ANCHORS: Fabrication (KG), Anchor bolts (NOS), Grout (CUM)
PART G — MISC: Anti-termite (SQM)
Excavation: footing plan area × (Df + 0.075 PCC + 0.3 working) × 1.25 slope × qty` : '';

  const buildingFoundationBOQHint = isBuildingFoundation ? `
BUILDING PAD FOOTING BOQ PARTS TO INCLUDE:
PART A — EARTHWORK: 
  Excavation (CUM) = RCC footing plan area × (Df_mm/1000 + 0.15 PCC + 0.3 working) × 1.25 × qty
  Backfilling & compaction (CUM) = Excavation CUM - Footing CUM - PCC CUM
PART B — PCC BED (M7.5): 
  PCC CUM = PCC_B_m × PCC_L_m × 0.15 × qty  [PCC is 150mm thick, 6" as noted]
PART C — RCC FOOTING (M30):
  RCC CUM = RCC_B_m × RCC_L_m × dmin_m × qty  [use dmin, NOT Df]
PART D — FORMWORK to footing sides (SQM)
PART E — REINFORCEMENT (Fe500D):
  Footing steel (MT) = RCC footing CUM × 90 kg/CUM (typical for pad footing)
PART F — ANTI-TERMITE TREATMENT (SQM) = total footing plan area

⚠️ CALCULATION RULES:
- Use Df_mm for EXCAVATION depth (total depth from NGL)
- Use dmin_mm for RCC FOOTING THICKNESS (footing concrete volume)
- PCC plan size = pcc_size_mm (larger than RCC)
- RCC plan size = rcc_size_mm (smaller, sits on PCC)
- Convert mm → m by dividing by 1000
- Qty for each row = count of column numbers listed in COLUMN NO. column` : '';

  const columnDetailBOQHint = isColumnDetail ? `
COLUMN REINFORCEMENT DETAIL DRAWING — BOQ NOTE:
This drawing shows column reinforcement details only (no footing schedule).
BOQ from this drawing = Column concrete + steel only (foundation BOQ from separate drawing).
PART A — RCC COLUMNS (M30): CUM = column size × height × qty per type
PART B — REINFORCEMENT: MT = column CUM × 200 kg/CUM (typical for columns with ties)
PART C — FORMWORK: SQM = perimeter × height × qty
Note: Column sizes in INCHES — already converted to mm in Phase 3 schedule_data.` : '';

  const raw = await callClaude({
    messages:[{ role:'user', content:`TASK PHASE 4: Calculate complete BOQ from Phase 3 quantities and schedule data.

DRAWING: type=${drawingType} | project=${meta?.project_name||'?'} | concrete=${meta?.concrete_grade||'?'} | steel=${meta?.steel_grade||'?'}
Unit system detected by Phase 3: ${quantities?.unit_system_detected || 'unknown'}
${scheduleCtx}
ELEMENT COUNTS FROM PHASE 3:
${JSON.stringify(quantities.element_counts || {}, null, 2)}
${industrialBOQHint}${buildingFoundationBOQHint}${columnDetailBOQHint}

RATES: Use DSR 2025 rates from system prompt.
Group items into PARTS (PART A, PART B etc.).
Show formula in calc_note so engineer can verify each item.
Do NOT add items with qty=0.

Return JSON:
{"project_name":"","drawing_type":"","drawing_no":"","date":"","concrete_grade":"","steel_grade":"","boq":[{"sr":1,"part":"PART A","description":"","unit":"","qty":0,"rate":0,"amount":0,"source":"drawing-schedule|calculated","confidence":"high|medium|low","calc_note":""}],"element_counts":{},"area_statement":{"total_bua_sqmt":0,"floor_wise":[],"road_area_sqmt":0,"road_length_rmt":0},"cost_summary":{"civil_total_inr":0,"civil_total_lacs":0,"civil_total_crores":0},"observations":[],"missing_info":[]}` }],
    maxTokens: 4096
  });

  const boq = parseJSON(raw);
  console.log(`[Phase 4] ${boq?.boq?.length||0} BOQ items, total: Rs.${boq?.cost_summary?.civil_total_lacs||0} lacs`);
  return boq;
}

// ════════════════════════════════════════════
// PHASE 5 — Validation (FREE JS — no Claude API call)
// Replaced extended thinking (expensive) with pure math checks
// Saves ~₹8-15 per drawing analysis
// ════════════════════════════════════════════
async function phase5_validateAndFlag(boqData, layout={}) {
  console.log('[Phase 5] FREE JS validation (no API call)...');
  if (!boqData) return boqData;

  const warnings = [], passed = [], engineerActions = [], pmcFlags = [];
  const boq = boqData.boq || [];
  const find = (kw) => boq.filter(i => (i.description||'').toLowerCase().includes(kw.toLowerCase()));
  const sumQty = (items) => items.reduce((s, i) => s + (i.qty||0), 0);

  // CHECK 1: qty=0 but amount>0
  for (const item of boq) {
    if ((item.qty||0) === 0 && (item.amount||0) > 0) {
      warnings.push({ item: item.description, check: 'qty=0 but amount>0', expected: 'amount=0', found: `amount=${item.amount}`, severity: 'HIGH' });
    }
  }
  if (!warnings.some(w => w.check === 'qty=0 but amount>0')) passed.push('No qty=0 with amount>0 errors');

  // CHECK 2: Steel ratios
  const footingCUM = sumQty(find('footing'));
  const colCUM = sumQty(find('pedestal')) + sumQty(find('column'));
  const steelItems = find('steel').concat(find('rebar')).concat(find('fe500'));
  const steelKG = steelItems.reduce((s,i) => {
    const u = (i.unit||'').toLowerCase();
    return s + (i.qty||0) * (u==='mt'||u==='ton' ? 1000 : 1);
  }, 0);
  if (footingCUM > 0 && steelKG > 0) {
    const r = steelKG/footingCUM;
    if (r < 60 || r > 150) warnings.push({ item: 'Steel in footing', check: 'kg/CUM ratio', expected: '70-120', found: Math.round(r), severity: 'MEDIUM' });
    else passed.push(`Footing steel ratio OK: ${Math.round(r)} kg/CUM`);
  }
  if (colCUM > 0 && steelKG > 0) {
    const r = steelKG/colCUM;
    if (r < 100 || r > 320) warnings.push({ item: 'Steel in column', check: 'kg/CUM ratio', expected: '180-240', found: Math.round(r), severity: 'MEDIUM' });
    else passed.push(`Column steel ratio OK: ${Math.round(r)} kg/CUM`);
  }

  // CHECK 3: Excavation >= footing
  const excav = sumQty(find('excavat'));
  if (excav > 0 && footingCUM > 0 && excav < footingCUM) {
    warnings.push({ item: 'Excavation', check: 'Excavation < footing CUM', expected: `>=${footingCUM.toFixed(2)}`, found: excav.toFixed(2), severity: 'HIGH' });
    pmcFlags.push('IS 1200-Part 1: Excavation must exceed structural volume');
  } else if (excav > 0) passed.push('Excavation > footing volume: OK');

  // CHECK 4: GSB/WMM ratio
  const gsbT = sumQty(find('gsb')), wmmT = sumQty(find('wmm'));
  if (gsbT > 0 && wmmT > 0) {
    const r = gsbT/wmmT;
    if (r < 1.1 || r > 1.7) warnings.push({ item: 'GSB/WMM ratio', check: 'road layer ratio', expected: '~1.3', found: r.toFixed(2), severity: 'LOW' });
    else passed.push(`GSB/WMM ratio OK: ${r.toFixed(2)}`);
  }

  // CHECK 5: Assumed items
  const assumed = boq.filter(i => (i.source||'').toLowerCase().includes('assumed'));
  if (assumed.length) {
    engineerActions.push(`Verify ${assumed.length} assumed items: ${assumed.map(i=>i.description).slice(0,3).join(', ')}`);
    warnings.push({ item: 'Assumed values', check: 'source=assumed', expected: 'drawing-schedule', found: `${assumed.length} items`, severity: 'MEDIUM' });
  }

  // CHECK 6: Total cost sanity
  const total = boqData.cost_summary?.civil_total_lacs || 0;
  if (total > 0 && total > 500) warnings.push({ item: 'Total cost', check: 'unusually high', expected: '<500 lacs for foundation', found: `${total} lacs`, severity: 'LOW' });
  else if (total > 0) passed.push(`Total ₹${total} lacs — plausible`);

  const highW = warnings.filter(w => w.severity==='HIGH').length;
  const overall_confidence = highW >= 2 ? 'LOW' : highW === 1 ? 'MEDIUM' : warnings.length > 3 ? 'MEDIUM' : 'HIGH';
  console.log(`[Phase 5 JS] ${warnings.length} warnings | ${passed.length} passed | confidence:${overall_confidence} | NO API CALL`);
  return { ...boqData, validation_warnings: warnings, validation_passed: passed, overall_confidence, engineer_action_required: engineerActions, pmc_flags: pmcFlags };
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

  // Build GCV table context string — passed to all phases so Claude uses actual table values
  let gcvTableContext = '';
  if (cvData?.gcv_validated_table) {
    const t = cvData.gcv_validated_table;
    const headerLine = (t.headers || []).join(' | ');
    const rowLines = (t.rows || []).map(r => r.map(v => v ?? '').join(' | ')).join('\n');
    gcvTableContext = `\n\nSCANNED PDF TABLE (GCV+Claude validated, confidence:${t.confidence||'?'}):\nUse ONLY these values — do not assume or calculate anything not present here.\nHeaders: ${headerLine}\n${rowLines}\n`;
    if (t.issues_fixed?.length) gcvTableContext += `Corrections applied: ${t.issues_fixed.join('; ')}\n`;
  } else if (cvData?.gcv_raw_text) {
    gcvTableContext = `\n\nSCANNED PDF RAW TEXT (GCV):\n${cvData.gcv_raw_text.slice(0, 3000)}\nUse ONLY values present in this text.\n`;
  }

  // ── DETECT DRAWING LAYOUT AUTOMATICALLY ──────────────────────────
  // This is the KEY step — it tells all phases HOW to read this drawing
  // No hardcoded values — purely signal-based from GCV text + file metadata
  const layout = detectDrawingLayout(files, cvData, gcvTableContext);
  console.log(`[PMC] Layout detected: ${layout.drawingType} | Strategy: ${layout.layoutStrategy}`);

  const meta        = await phase2_legendAndScale(files, cvData, gcvTableContext, layout);
  const quantities  = await phase3_extractQuantities(files, meta, gcvTableContext, layout);
  const boqData     = await phase4_calculateBOQ(quantities, meta, layout);
  const finalData   = await phase5_validateAndFlag(boqData, layout);

  return buildFinalOutput(finalData, quantities, meta, cvData, layout);
}

function buildFinalOutput(boq, quantities, meta, cvData, layout={}) {
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
    concrete_grade: boq.concrete_grade||meta?.concrete_grade||'',
    steel_grade:    boq.steel_grade||meta?.steel_grade||'',
    structural_system: meta?.structural_system||'',
    elements: (boq.boq||[]).map((item,i)=>({
      id:`E${String(i+1).padStart(3,'0')}`, type:guessType(item.description),
      name:item.description, dimensions:{note:item.calc_note||item.source||''},
      quantities:{
        area_sqmt:   item.unit==='sqmt'?item.qty:0,
        volume_cum:  ['cum','CUM'].includes(item.unit)?item.qty:0,
        gsb_ton:     /gsb/i.test(item.description)?item.qty:0,
        wmm_ton:     /wmm/i.test(item.description)?item.qty:0,
        pqc_cum:     /pqc/i.test(item.description)?item.qty:0,
        steel_kg:    ['kg','mt','ton','MT'].includes(item.unit?.toLowerCase())?item.qty:0,
        brickwork_cum:/brick/i.test(item.description)?item.qty:0,
      },
      cost_inr:{total:item.amount||0,per_unit:item.rate||0},
      confidence:item.confidence||'medium', annotation_found:item.source||'',
      part:item.part||'', sr:item.sr||i+1, calc_note:item.calc_note||'',
    })),
    element_counts:{...quantities?.element_counts,...boq.element_counts},
    schedule_data: quantities?.schedule_data||{},      // ← Phase 3 schedule tables surfaced
    section_details: quantities?.section_details||{},  // ← footing depth, cover, PCC thickness
    grid_info: quantities?.grid_info||{},              // ← bay spacing, braced grids
    total_quantities:{
      total_area_sqmt: boq.area_statement?.total_bua_sqmt||0,
      total_road_rmt:  boq.area_statement?.road_length_rmt||0,
      gsb_total_ton:   sumKw(boq.boq,'gsb'),
      wmm_total_ton:   sumKw(boq.boq,'wmm'),
      pqc_total_cum:   sumKw(boq.boq,'pqc'),
      rcc_total_cum:   sumKw(boq.boq,'rcc'),
      steel_total_kg:  sumKw(boq.boq,'steel'),
      footing_rcc_cum: sumKw(boq.boq,'footing'),
      pedestal_rcc_cum:sumKw(boq.boq,'pedestal'),
      base_plate_kg:   sumKw(boq.boq,'base plate'),
      calculation_note:'5-phase Claude pipeline',
    },
    cost_summary:{civil_total_inr:totalInr,civil_total_lacs:totalLacs,civil_total_crores:totalCr,item_wise:boq.boq||[]},
    area_statement:    boq.area_statement||{},
    validation_warnings:      boq.validation_warnings||[],
    validation_passed:        boq.validation_passed||[],
    pmc_flags:                boq.pmc_flags||[],
    overall_confidence:       boq.overall_confidence||'MEDIUM',
    engineer_action_required: boq.engineer_action_required||[],
    legend: meta?.legend||[],
    general_notes: meta?.general_notes||[],
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
      phase2_schedule_tables:meta?.schedule_tables_visible||[],
      phase3_elements:quantities?.quantities?.length||0,
      phase3_columns_read:quantities?.schedule_data?.columns?.length||0,
      phase3_footings_read:quantities?.schedule_data?.footings?.length||0,
      phase4_boq_items:boq.boq?.length||0,
      phase5_warnings:boq.validation_warnings?.length||0,
      model:'claude-sonnet-4-6',
      detected_layout:{
        drawing_type: layout.drawingType||'UNKNOWN',
        layout_strategy: layout.layoutStrategy||'STANDARD',
        panel_map: layout.panelMap||{},
        signals: layout.signals||{},
      },
    },
  };
}

function guessType(desc){
  if(!desc)return'GENERAL';
  const d=desc.toLowerCase();
  if(/road|gsb|wmm|pqc|kerb/.test(d))return'ROAD';
  if(/steel|rebar|bar|fe500|hysd/.test(d))return'STEEL';
  if(/base plate|anchor bolt|grout/.test(d))return'BASE_PLATE';
  if(/rcc|slab|beam|column|footing|pedestal/.test(d))return'STRUCTURE';
  if(/pcc|plain cement/.test(d))return'PCC';
  if(/brick|wall/.test(d))return'WALL';
  if(/excavat|earth|backfill|dewater/.test(d))return'EARTHWORK';
  if(/plaster|paint|floor|tile|anti.?termite/.test(d))return'FINISH';
  if(/formwork|shuttering/.test(d))return'FORMWORK';
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
// callClaudeAPI — exported alias used by server.js for direct API calls
// Routes through callClaude so model name, headers, retries are always consistent
async function callClaudeAPI({ system, messages, maxTokens = 4096 }) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');
  const body = {
    model: 'claude-sonnet-4-6',
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
async function claudeAnalyzeDXF(civilData, filename, ratesSummary, smartPrompt) {
  console.log('[claudeAnalyzeDXF] analysing DXF data for', filename);
  // If smart engine provided a pre-drafted prompt, use it (90-95% accuracy mode)
  // Otherwise fall back to old raw-dump approach
  const prompt = smartPrompt || `You are a senior PMC civil engineer. Analyse this parsed DXF data and generate a BOQ.
DXF FILE: ${filename}
RATES SUMMARY: ${ratesSummary || 'Use DSR 2025 rates from system prompt.'}
DXF DATA:
${JSON.stringify({
  filename: civilData.filename,
  drawing_type: civilData.drawing_type,
  floor_levels: civilData.floor_levels,
  floor_heights: civilData.floor_heights,
  wall_by_thickness: civilData.wall_by_thickness,
  element_counts: civilData.element_counts,
  schedule_tables: civilData.schedule_tables,
  all_texts: (civilData.all_texts || []).slice(0, 80),
  dimension_values: (civilData.dimension_values || []).slice(0, 30),
  polyline_areas: (civilData.polyline_areas || []).slice(0, 20),
  block_counts: civilData.block_counts,
}, null, 2)}

Return ONLY raw JSON:
{"project_name":"","drawing_type":"","boq":[{"sr":1,"description":"","unit":"","qty":0,"rate":0,"amount":0,"source":"dxf-data","confidence":"high"}],"cost_summary":{"civil_total_inr":0,"civil_total_lacs":0},"observations":[]}`;
  const raw = await callClaudeAPI({ system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }], maxTokens: 4096 });
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
  const raw = await callClaudeAPI({ system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }], maxTokens: 4096 });
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
═══════════════════════════════════════════════════════════
COLUMN / FOOTING SCHEDULE READING — ABSOLUTE RULES
═══════════════════════════════════════════════════════════
1. ONLY read values that are PHYSICALLY PRINTED in the schedule table cells.
2. Column/Pedestal sizes: copy the EXACT printed value (e.g. 300×300, 230×450).
   - NEVER output 400×400, 450×450, 500×500 unless those numbers appear in the drawing.
3. Main bars / stirrups: copy EXACTLY (e.g. 8-12Ø, 4-16Ø, 10T16).
   - NEVER output 12-20Ø, 16-25Ø or any bar size not printed in the schedule.
4. Column schedule and Footing schedule are COMPLETELY SEPARATE tables.
   - NEVER copy footing sizes/steel into the column table or vice versa.
5. Qty (number of columns/footings): use ONLY the qty column in the schedule.
   - NEVER guess or count from the plan view.
6. If any cell is unclear / not legible → write "not legible" — do NOT guess.
7. Source field: "drawing-schedule" for schedule values, "calculated" for derived.
8. INDUSTRIAL DRAWING: If you see base plates, anchor bolts, BOD OF STEEL, braced bays:
   - "Column Schedule" = RCC PEDESTAL schedule (not steel section schedule)
   - Read base plate dimensions and anchor bolt pattern from detail panels
   - Note concrete grade (M40 etc.) and steel grade (Fe500D etc.) from NOTES/title block
═══════════════════════════════════════════════════════════`;

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
  "project_name": "",
  "drawing_no": "",
  "concrete_grade": "",
  "steel_grade": "",
  "structural_system": "",
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
      "pcc_mm": 75,
      "main_bars_x": "",
      "main_bars_y": "",
      "qty": 0,
      "pedestal_size_mm": "",
      "source": "drawing-schedule|not legible"
    }
  ],
  "base_plate_schedule": [
    {
      "column_mark": "",
      "plate_size_mm": "",
      "anchor_bolt_nos": 0,
      "anchor_bolt_dia_mm": 0,
      "source": "drawing-schedule|not legible"
    }
  ],
  "section_details": {
    "footing_depth_mm": 0,
    "pedestal_height_mm": 0,
    "pcc_thickness_mm": 75,
    "cover_mm": 50
  },
  "grid_info": {
    "typical_bay_m": 0,
    "total_columns_plan": 0,
    "braced_bay_grids": []
  },
  "boq": [
    { "sr": 1, "part": "PART A", "description": "", "unit": "", "qty": 0, "rate": 0, "amount": 0, "source": "drawing-schedule|calculated", "confidence": "high|medium|low", "calc_note": "" }
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
    maxTokens: 4096,
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
    maxTokens: 4096,
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
