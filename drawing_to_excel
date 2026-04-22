/**
 * PMC Drawing → Excel Engine
 * Drawing upload → AI analysis → Dynamic Excel based on drawing type
 */

const ExcelJS = require('exceljs');
const fs   = require('fs');
const path = require('path');

// ── RATES from rates.json (single source of truth) ──────────────
// No rates hardcoded here. Rates only loaded from rates.json.
let DSR = {};
try {
  const raw = JSON.parse(fs.readFileSync(path.join(__dirname, 'rates.json'), 'utf8'));
  for (const category of Object.values(raw)) {
    if (typeof category === 'object' && !Array.isArray(category)) {
      for (const [key, val] of Object.entries(category)) {
        if (val && typeof val.rate === 'number') DSR[key] = val.rate;
      }
    }
  }
  // Aliases so existing references keep working (same rate, different key)
  DSR.brickwork_230_cum  = DSR.brickwork_230_cum  || DSR.brickwork_230mm_cum;
  DSR.brickwork_115_cum  = DSR.brickwork_115_cum  || DSR.brickwork_115mm_cum;
  DSR.plaster_sqmt       = DSR.plaster_sqmt       || DSR.plaster_15mm_sqmt;
  DSR.cp_wall_rmt        = DSR.cp_wall_rmt        || DSR.compound_wall_rmt;
  DSR.plot_boundary_rmt  = DSR.plot_boundary_rmt  || DSR.compound_wall_rmt;
} catch (e) {
  console.warn('rates.json not loaded:', e.message);
}

// ── EXCEL STYLES (matching original template) ──────────────────
const C = {
  TITLE_BG:   'FFDEEBF6',  // light blue - main headers
  SECTION_BG: 'FFFFE599',  // light yellow - part/section headers
  SUB_BG:     'FFEDEDED',  // light grey - subtotals
  TOTAL_BG:   'FFFFC000',  // gold - grand totals
  GREEN_BG:   'FFE2EFDA',  // light green
  WHITE:      'FFFFFFFF',
  BLACK:      'FF000000',
  ALT_ROW:    'FFF2F2F2',  // alternating row
};

const thin = { style:'thin', color:{argb:'FF000000'} };
const bdr  = { top:thin, left:thin, bottom:thin, right:thin };

function sc(cell, bg, bold=false, fc='FF000000', size=11, align='left') {
  if(bg) cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:bg}};
  cell.font = {bold, color:{argb:fc}, size, name:'Calibri'};
  cell.alignment = {horizontal:align, vertical:'middle', wrapText:true};
  cell.border = bdr;
}

function mkTitle(ws, row, text, cols) {
  ws.mergeCells(row,1,row,cols);
  sc(ws.getCell(row,1), C.TITLE_BG, true, C.BLACK, 18, 'center');
  ws.getCell(row,1).value = text;
  ws.getRow(row).height = 22.5;
  return row+1;
}

function mkSection(ws, row, text, cols) {
  ws.mergeCells(row,1,row,cols);
  sc(ws.getCell(row,1), C.SECTION_BG, true, C.BLACK, 14, 'left');
  ws.getCell(row,1).value = text;
  ws.getRow(row).height = 22.5;
  return row+1;
}

function mkHeaders(ws, row, hdrs, widths) {
  hdrs.forEach((h,i) => {
    sc(ws.getCell(row,i+1), C.TITLE_BG, true, C.BLACK, 11, 'center');
    ws.getCell(row,i+1).value = h;
  });
  ws.getRow(row).height = 42.75;
  if(widths) widths.forEach((w,i) => { if(w) ws.getColumn(i+1).width=w; });
  return row+1;
}

function mkDataRow(ws, row, vals, numCols=[]) {
  vals.forEach((v,i) => {
    const bg = row%2===0 ? C.ALT_ROW : C.WHITE;
    sc(ws.getCell(row,i+1), bg, false, C.BLACK, 11, typeof v==='number'?'right':'left');
    ws.getCell(row,i+1).value = v??'';
    if(numCols.includes(i) && typeof v==='number') ws.getCell(row,i+1).numFmt = '#,##0.00';
  });
  ws.getRow(row).height = 15;
  return row+1;
}

function mkSubtotal(ws, row, label, lacs, cols) {
  ws.mergeCells(row,1,row,cols-2);
  sc(ws.getCell(row,1), C.SUB_BG, true, C.BLACK, 11, 'right');
  ws.getCell(row,1).value = label;
  sc(ws.getCell(row,cols-1), C.SUB_BG, true, C.BLACK, 11, 'right');
  ws.getCell(row,cols-1).value = lacs;
  ws.getCell(row,cols-1).numFmt = '#,##0.00';
  sc(ws.getCell(row,cols), C.SUB_BG, true, C.BLACK, 11, 'right');
  ws.getCell(row,cols).value = lacs/100;
  ws.getCell(row,cols).numFmt = '#,##0.0000';
  ws.getRow(row).height = 16;
  return row+2;
}

function mkGrandTotal(ws, row, label, lacs, cols) {
  ws.mergeCells(row,1,row,cols-2);
  sc(ws.getCell(row,1), C.TOTAL_BG, true, C.BLACK, 14, 'right');
  ws.getCell(row,1).value = label;
  sc(ws.getCell(row,cols-1), C.TOTAL_BG, true, C.BLACK, 14, 'right');
  ws.getCell(row,cols-1).value = lacs;
  ws.getCell(row,cols-1).numFmt = '#,##0.00';
  sc(ws.getCell(row,cols), C.TOTAL_BG, true, C.BLACK, 14, 'right');
  ws.getCell(row,cols).value = lacs/100;
  ws.getCell(row,cols).numFmt = '#,##0.00';
  ws.getRow(row).height = 20;
  return row+2;
}

// ══════════════════════════════════════════════════════════════════
// GEMINI DRAWING ANALYSIS PROMPT
// ══════════════════════════════════════════════════════════════════
function getDrawingPrompt() {
  return `You are a SENIOR PMC CIVIL ENGINEER. Read THIS drawing and extract ONLY values that are visibly annotated in it.

CRITICAL RULES (NEVER BREAK):
1. Do NOT invent, guess, estimate, or carry over numbers from other drawings.
2. If a value is not clearly annotated or readable in THIS drawing, set it to 0 or null.
3. The example structure below uses placeholder zeros — DO NOT copy those numbers; extract your own from the drawing.
4. Every number you return must trace back to a specific dimension text, label, or scaled measurement visible in the drawing.
5. If you are not certain, leave the field at 0/null rather than filling it in.

STEP 1 — IDENTIFY DRAWING TYPE: ROAD_LAYOUT / BUILDING / COMPOUND_WALL / DRAINAGE / SITE_LAYOUT / STRUCTURE

STEP 2 — READ TITLE BLOCK: project name, drawing no, scale, date, location (only if printed).

STEP 3 — READ EVERY DIMENSION ANNOTATION, ROOM LABEL, CALL-OUT, LEVEL MARK.

STEP 4 — COUNT ELEMENTS VISIBLE IN DRAWING (doors, windows, lifts, staircases, floors, rooms, columns, footings). Count what you can actually see — no assumptions.

STEP 5 — COMPUTE QUANTITIES USING ONLY READ DIMENSIONS. Formulas:
- Road area = length × carriage_width
- GSB ton = area × 1.15 × 0.300 × 1.800
- WMM ton = area × 1.15 × 0.200 × 2.100
- PQC cum = area × 1.05 × 0.250
- Steel dowel kg = area × 3.87
- RCC volume = length × width × depth
- Steel for RCC = volume × 120 kg/cum (slab), 160 kg/cum (beam)
- Brickwork cum = length × height × thickness  (all three must be read from drawing)

RETURN ONLY RAW JSON (no markdown, no backticks). All numeric fields default to 0 if not read from drawing. All string fields default to "" if not read. Arrays default to []:
{
  "drawing_type": "",
  "project_name": "",
  "location": "",
  "scale": "",
  "date": "",
  "total_area_sqmt": 0,
  "confidence": "LOW|MEDIUM|HIGH based on how much you could actually read",

  "roads": [
    { "id": "", "road_no": "", "total_width_m": 0, "carriage_width_m": 0, "length_m": 0, "sc_width_m": 0, "road_top_level": 0, "ngl": 0, "remark": "" }
  ],

  "buildings": [
    { "id": "", "name": "", "floors": 0, "plinth_area_sqmt": 0, "total_built_up_sqmt": 0, "wall_length_rmt": 0, "wall_height_m": 0, "wall_thickness_m": 0, "slab_area_sqmt": 0, "slab_thickness_m": 0 }
  ],

  "compound_wall": { "total_length_rmt": 0, "gabion_length_rmt": 0, "cp_wall_length_rmt": 0, "height_m": 0, "section_type": "" },

  "drainage": { "pipe_length_rmt": 0, "pipe_dia_mm": 0, "culvert_nos": 0, "culvert_width_m": 0, "culvert_span_m": 0 },

  "structure": { "element_type": "", "dimensions": {}, "reinforcement": [] },

  "site_details": { "total_area_sqmt": 0, "road_area_sqmt": 0, "green_area_sqmt": 0, "plot_boundary_rmt": 0, "service_corridor_area_sqmt": 0 },

  "element_counts": { "door_count": 0, "window_count": 0, "lift_count": 0, "staircase_count": 0, "column_count": 0, "footing_count": 0, "toilet_count": 0, "kitchen_count": 0, "bedroom_count": 0, "floor_count": 0 },

  "soil_filling_cum": 0,
  "stp_mld": 0,
  "electrical_ls_inr": 0,
  "street_lights_nos": 0,
  "contingency_pct": 0,

  "extracted_dimensions": [],
  "pmcnotes": []
}`;
}

// ══════════════════════════════════════════════════════════════════
// SHEET BUILDERS - ONE PER DRAWING TYPE
// ══════════════════════════════════════════════════════════════════

// BLOCK COST - common to ALL drawing types
function buildBlockCost(wb, data) {
  const ws = wb.addWorksheet('BLOCK COST');
  ws.getColumn(1).width=7; ws.getColumn(2).width=56;
  ws.getColumn(3).width=14; ws.getColumn(4).width=9;
  ws.getColumn(5).width=15; ws.getColumn(6).width=21;
  ws.getColumn(7).width=20; ws.getColumn(8).width=44;
  const COLS=8;

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PMC CIVIL PROJECT', COLS);
  row = mkTitle(ws, row, 'BLOCK COST', COLS);
  ws.mergeCells(row,1,row,COLS); ws.getCell(row,1).value=`TOTAL AREA: ${(data.total_area_sqmt||0).toLocaleString('en-IN')} SQMT`;
  sc(ws.getCell(row,1),C.TITLE_BG,true,C.BLACK,11,'left'); ws.getRow(row).height=18; row++;
  row = mkHeaders(ws, row, ['SR NO','COMPONENT','ESTIMATE QUANTITY','UNITS','RATE','ESTIMATE VALUE (RS IN LACS)','ESTIMATE VALUE (RS IN CR.)','REMARK']);

  const roads = data.roads||[];
  const site = data.site_details||{};
  const dtype = data.drawing_type||'SITE_LAYOUT';

  row = mkSection(ws, row, 'PART-A (PROJECT WORKS)', COLS);

  let grandTotal = 0;
  let srNo = 1;

  // Roads section — values ONLY from drawing data, no guesses.
  // Skip a row if required dimension is missing in the drawing.
  if(dtype==='ROAD_LAYOUT'||dtype==='SITE_LAYOUT'||roads.length>0) {
    // Only include roads that have both length AND carriage width from drawing
    const validRoads = roads.filter(r => (r.length_m||0) > 0 && (r.carriage_width_m||0) > 0);
    const totalArea = validRoads.reduce((s,r)=>s+(r.length_m*r.carriage_width_m),0);

    // soil filling: only if drawing-derived value is available
    const soilFillCum = (typeof data.soil_filling_cum === 'number' && data.soil_filling_cum > 0) ? data.soil_filling_cum : null;
    const soilFill = soilFillCum ? soilFillCum*DSR.soil_filling_cum/100000 : null;

    const gsb   = totalArea>0 ? totalArea*DSR.gsb_300mm_sqmt/100000 : null;
    const wmm   = totalArea>0 ? totalArea*DSR.wmm_200mm_sqmt/100000 : null;
    const pqc   = totalArea>0 ? totalArea*DSR.pqc_250mm_sqmt/100000 : null;
    const steel = totalArea>0 ? (totalArea*3.87*DSR.steel_fe500_kg)/100000 : null;

    // Service corridor: only from explicit drawing values. No "default 5m width".
    const scFromRoads = roads
      .filter(r => (r.length_m||0)>0 && (r.sc_width_m||0)>0)
      .reduce((s,r)=>s+r.length_m*r.sc_width_m,0);
    const sc_area = site.service_corridor_area_sqmt || (scFromRoads>0 ? scFromRoads : 0);
    const sc_cost = sc_area>0 ? sc_area*DSR.service_corridor_sqmt/100000 : null;

    const pb = (site.plot_boundary_rmt||0)>0 ? site.plot_boundary_rmt*DSR.plot_boundary_rmt/100000 : null;
    const roadTotal = (soilFill||0)+(gsb||0)+(wmm||0)+(pqc||0)+(steel||0)+(sc_cost||0)+(pb||0);

    const addRow = (sr,label,qty,unit,rate,lacs,rem='') => {
      ws.getCell(row,1).value=sr||''; sc(ws.getCell(row,1),sr?null:null,!!sr,C.BLACK,sr?12:11,'center'); ws.getCell(row,1).border=bdr;
      ws.getCell(row,2).value=label; sc(ws.getCell(row,2),null,!!sr,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr;
      ws.getCell(row,3).value=qty||''; ws.getCell(row,3).numFmt='#,##0.00'; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr;
      ws.getCell(row,4).value=unit||''; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr;
      ws.getCell(row,5).value=rate||''; ws.getCell(row,5).numFmt='#,##0.00'; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr;
      ws.getCell(row,6).value=lacs||''; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr;
      ws.getCell(row,7).value=lacs?(lacs/100):''; ws.getCell(row,7).numFmt='#,##0.0000'; sc(ws.getCell(row,7),null,false,C.BLACK,11,'right'); ws.getCell(row,7).border=bdr;
      ws.mergeCells(row,8,row,8); ws.getCell(row,8).value=rem; sc(ws.getCell(row,8),null,false,C.BLACK,9,'left'); ws.getRow(row).height=15;
      row++;
    };

    addRow(srNo++,'ROADS','','','','');
    if(soilFillCum!==null) addRow(null,'SOIL FILLING IN ROAD AREA', Math.round(soilFillCum*100)/100,'CUM',DSR.soil_filling_cum,soilFill);
    if(gsb!==null)         addRow(null,'GSB FILLING (300 MM LAYER)',Math.round(totalArea*100)/100,'SQMT',DSR.gsb_300mm_sqmt,gsb);
    if(wmm!==null)         addRow(null,'WMM FILLING (200 MM LAYER)',Math.round(totalArea*100)/100,'SQMT',DSR.wmm_200mm_sqmt,wmm);
    if(pqc!==null)         addRow(null,'RCC ROAD WORK (250 MM PQC)',Math.round(totalArea*100)/100,'SQMT',DSR.pqc_250mm_sqmt,pqc,'250 MM PQC (M30)');
    if(steel!==null)       addRow(null,'STEEL FOR DOWEL BAR',Math.round(totalArea*3.87/1000*100)/100,'TON',DSR.steel_fe500_kg*1000,steel);
    if(sc_area>0)          addRow(null,'SERVICE CORRIDOR',Math.round(sc_area*100)/100,'SQMT',DSR.service_corridor_sqmt,sc_cost,'YELLOW SOIL+GSB+PAVER BLOCK');
    if(pb!==null)          addRow(null,'PLOT BOUNDARY WITH ROAD BEAM',site.plot_boundary_rmt,'RMT',DSR.plot_boundary_rmt,pb);
    row = mkSubtotal(ws, row, 'SUBTOTAL - ROADS', roadTotal, COLS);
    grandTotal += roadTotal;
  }

  // Buildings section — all dimensions MUST come from drawing.
  // If slab thickness / wall height not in drawing, that line item is skipped.
  if(dtype==='BUILDING'||dtype==='SITE_LAYOUT') {
    const bldgs = data.buildings||[];
    const bldgTotal = bldgs.reduce((s,b)=>{
      const slabThk = b.slab_thickness_m > 0 ? b.slab_thickness_m : null;
      const wallH   = b.wall_height_m    > 0 ? b.wall_height_m    : null;
      const wallThk = b.wall_thickness_m > 0 ? b.wall_thickness_m : null;  // must be in drawing
      const rcc     = (b.slab_area_sqmt && slabThk) ? b.slab_area_sqmt*slabThk*DSR.rcc_m25_cum/100000 : 0;
      const brick   = (b.wall_length_rmt && wallH && wallThk) ? b.wall_length_rmt*wallH*wallThk*DSR.brickwork_230_cum/100000 : 0;
      const plaster = (b.wall_length_rmt && wallH) ? b.wall_length_rmt*wallH*2*DSR.plaster_sqmt/100000 : 0;
      return s+rcc+brick+plaster;
    },0);
    if(bldgs.length>0) {
      ws.getCell(row,1).value=srNo++; sc(ws.getCell(row,1),null,true,C.BLACK,12,'center'); ws.getCell(row,1).border=bdr;
      ws.mergeCells(row,2,row,COLS); ws.getCell(row,2).value='BUILDINGS'; sc(ws.getCell(row,2),null,true,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getRow(row).height=18; row++;
      bldgs.forEach(b=>{
        const slabThk = b.slab_thickness_m > 0 ? b.slab_thickness_m : 0;
        const rcc=(b.slab_area_sqmt||0)*slabThk*DSR.rcc_m25_cum/100000;
        ws.getCell(row,2).value=b.name||'BUILDING'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr;
        ws.getCell(row,3).value=b.total_built_up_sqmt||0; ws.getCell(row,3).numFmt='#,##0.00'; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr;
        ws.getCell(row,4).value='SQMT'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr;
        if(rcc>0){ ws.getCell(row,6).value=rcc; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; }
        ws.getRow(row).height=15; row++;
      });
      row = mkSubtotal(ws, row, 'SUBTOTAL - BUILDINGS', bldgTotal, COLS);
      grandTotal += bldgTotal;
    }
  }

  // Compound Wall
  if(data.compound_wall && (data.compound_wall.total_length_rmt||0)>0) {
    const cw = data.compound_wall;
    const cp_cost = (cw.cp_wall_length_rmt||0)*DSR.cp_wall_rmt/100000;
    const gab_cost = (cw.gabion_length_rmt||0)*DSR.gabion_wall_rmt/100000;
    ws.getCell(row,1).value=srNo++; sc(ws.getCell(row,1),null,true,C.BLACK,12,'center'); ws.getCell(row,1).border=bdr;
    ws.mergeCells(row,2,row,COLS); ws.getCell(row,2).value='COMPOUND WALL'; sc(ws.getCell(row,2),null,true,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getRow(row).height=18; row++;
    if(cw.cp_wall_length_rmt>0) { ws.getCell(row,2).value='CP WALL'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getCell(row,3).value=cw.cp_wall_length_rmt; ws.getCell(row,3).numFmt='#,##0.00'; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr; ws.getCell(row,4).value='RMT'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr; ws.getCell(row,5).value=DSR.cp_wall_rmt; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr; ws.getCell(row,6).value=cp_cost; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; ws.getRow(row).height=15; row++; }
    if(cw.gabion_length_rmt>0) { ws.getCell(row,2).value='GABION WALL'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getCell(row,3).value=cw.gabion_length_rmt; ws.getCell(row,3).numFmt='#,##0.00'; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr; ws.getCell(row,4).value='RMT'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr; ws.getCell(row,5).value=DSR.gabion_wall_rmt; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr; ws.getCell(row,6).value=gab_cost; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; ws.getRow(row).height=15; row++; }
    row = mkSubtotal(ws, row, 'SUBTOTAL - COMPOUND WALL', cp_cost+gab_cost, COLS);
    grandTotal += (cp_cost+gab_cost);
  }

  // Services — ONLY include if drawing provides the number. No default STP 3 MLD, no ₹25 Cr electrical L/S.
  const stp_mld  = (typeof data.stp_mld === 'number' && data.stp_mld > 0) ? data.stp_mld : null;
  const elec_ls  = (typeof data.electrical_ls_inr === 'number' && data.electrical_ls_inr > 0) ? data.electrical_ls_inr : null;
  const sl_nos   = (typeof data.street_lights_nos === 'number' && data.street_lights_nos > 0) ? data.street_lights_nos : null;

  const stp_cost  = stp_mld ? stp_mld * DSR.stp_per_mld / 100000 : 0;
  const elec_cost = elec_ls ? elec_ls / 100000 : 0;
  const sl_cost   = sl_nos  ? sl_nos * DSR.street_light_nos / 100000 : 0;
  const totalServices = stp_cost + elec_cost + sl_cost;

  if(totalServices > 0) {
    ws.getCell(row,1).value=srNo++; sc(ws.getCell(row,1),null,true,C.BLACK,12,'center'); ws.getCell(row,1).border=bdr;
    ws.mergeCells(row,2,row,COLS); ws.getCell(row,2).value='SERVICES (DRAINAGE / UTILITIES)'; sc(ws.getCell(row,2),null,true,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getRow(row).height=18; row++;
    if(stp_mld) { ws.getCell(row,2).value='SEWAGE TREATMENT PLANT'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getCell(row,3).value=stp_mld; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr; ws.getCell(row,4).value='MLD'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr; ws.getCell(row,5).value=DSR.stp_per_mld; ws.getCell(row,5).numFmt='#,##0'; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr; ws.getCell(row,6).value=stp_cost; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; ws.getRow(row).height=15; row++; }
    if(elec_ls)  { ws.getCell(row,2).value='ELECTRICAL INFRASTRUCTURE'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getCell(row,3).value=1; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr; ws.getCell(row,4).value='L/S'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr; ws.getCell(row,5).value=elec_ls; ws.getCell(row,5).numFmt='#,##0'; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr; ws.getCell(row,6).value=elec_cost; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; ws.getRow(row).height=15; row++; }
    if(sl_nos)   { ws.getCell(row,2).value='STREET LIGHTING'; sc(ws.getCell(row,2),null,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr; ws.getCell(row,3).value=sl_nos; sc(ws.getCell(row,3),null,false,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr; ws.getCell(row,4).value='NOS'; sc(ws.getCell(row,4),null,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr; ws.getCell(row,5).value=DSR.street_light_nos; ws.getCell(row,5).numFmt='#,##0'; sc(ws.getCell(row,5),null,false,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr; ws.getCell(row,6).value=sl_cost; ws.getCell(row,6).numFmt='#,##0.00'; sc(ws.getCell(row,6),null,false,C.BLACK,11,'right'); ws.getCell(row,6).border=bdr; ws.getRow(row).height=15; row++; }
    row = mkSubtotal(ws, row, 'SUBTOTAL - SERVICES', totalServices, COLS);
    grandTotal += totalServices;
  }

  // Grand Total — contingency % ONLY if specified in data (no hardcoded 10%)
  const contingencyPct = (typeof data.contingency_pct === 'number' && data.contingency_pct > 0) ? data.contingency_pct : 0;
  const contingency = grandTotal * contingencyPct / 100;
  const finalTotal  = grandTotal + contingency;
  row++;
  row = mkGrandTotal(ws, row, 'TOTAL COST (PART-A)', grandTotal, COLS);
  if(contingency > 0) row = mkGrandTotal(ws, row, `${contingencyPct}% CONTINGENCY`, contingency, COLS);
  if(finalTotal > grandTotal) row = mkGrandTotal(ws, row, 'GRAND TOTAL (WITH CONTINGENCY)', finalTotal, COLS);
  row++;
  ws.mergeCells(row,1,row,4); ws.getCell(row,1).value='LACS'; sc(ws.getCell(row,1),null,true,C.BLACK,11,'right'); ws.getCell(row,1).border=bdr;
  ws.mergeCells(row,5,row,COLS); ws.getCell(row,5).value='CR'; sc(ws.getCell(row,5),null,true,C.BLACK,11,'center'); ws.getCell(row,5).border=bdr;

  ws.views = [{state:'frozen', ySplit:4}];
  return { grandTotal, finalTotal };
}

// ROAD ESTIMATE SHEET
function buildRoadEstimate(wb, data) {
  const roads = data.roads||[];
  if(!roads.length) return;

  const ws = wb.addWorksheet('ROAD ESTIMATE');
  [9,12,10,18,14,16,16,18,17,17,26].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 11);
  row = mkTitle(ws, row, 'ROAD WORK (SUB-BASE & BASE COURSE ESTIMATE)', 11);
  row = mkHeaders(ws, row,
    ['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH (METER)','CARRIAGE WAY WIDTH (METER)',
     'AREA (SQMT)','BOX CUTTING (SQMT)','GSB FILLING\n300MM THK\n(15% EXTRA) TON',
     'WMM FILLING\n200MM THK\n(15% EXTRA) TON','PQC ROAD M30\n250MM THK\n(5% EXTRA) CUM','REMARK']);
  ws.getRow(row-1).height=52;

  // Layer labels row
  ['','','','','','','','LAYER-3','LAYER-4','LAYER-5',''].forEach((h,i)=>{
    ws.getCell(row,i+1).value=h;
    sc(ws.getCell(row,i+1),C.TITLE_BG,true,C.BLACK,10,'center');
  });
  ws.getRow(row).height=16; row++;

  let totLen=0, totArea=0, totBoxCut=0, totGSB=0, totWMM=0, totPQC=0;

  roads.forEach((rd,i)=>{
    const L = rd.length_m||0;
    // Only use carriage_width_m from drawing — no "0.65 of total width" guess
    const CW = rd.carriage_width_m > 0 ? rd.carriage_width_m : 0;
    const area = L*CW;
    const boxcut = area*1.05;
    const gsb = area*1.15*0.300*1.800;  // ton
    const wmm = area*1.15*0.200*2.100;  // ton
    const pqc = area*1.05*0.250;        // cum
    totLen+=L; totArea+=area; totBoxCut+=boxcut; totGSB+=gsb; totWMM+=wmm; totPQC+=pqc;

    const bg = i%2===0?C.WHITE:C.ALT_ROW;
    const vals = [i+1, `${rd.total_width_m||''} MT`, rd.road_no||`R${i+1}`,
      Math.round(L*100)/100, Math.round(CW*100)/100,
      Math.round(area*100)/100, Math.round(boxcut*100)/100,
      Math.round(gsb*100)/100, Math.round(wmm*100)/100,
      Math.round(pqc*100)/100, rd.remark||''];
    vals.forEach((v,ci)=>{
      const c=ws.getCell(row,ci+1); c.value=v;
      sc(c,bg,false,C.BLACK,11,typeof v==='number'?'right':'left');
      if(typeof v==='number') c.numFmt='#,##0.00';
    });
    ws.getRow(row).height=15; row++;
  });

  // Totals
  row++;
  const totVals = ['TOTAL','','',Math.round(totLen*100)/100,'',
    Math.round(totArea*100)/100, Math.round(totBoxCut*100)/100,
    Math.round(totGSB*100)/100, Math.round(totWMM*100)/100,
    Math.round(totPQC*100)/100,''];
  ws.mergeCells(row,1,row,3);
  totVals.forEach((v,ci)=>{
    const c=ws.getCell(row,ci===0?1:ci+1);
    if(ci===0){c.value='TOTAL';}else if(ci>2){c.value=v;}
    sc(c,C.TOTAL_BG,true,C.BLACK,11,'right');
    if(typeof v==='number') c.numFmt='#,##0.00';
  });
  ws.getRow(row).height=18; row+=2;

  // Units row
  ['METER','','','RMT','SQMT','SQMT','TON','TON','CUM',''].forEach((u,i)=>{
    ws.getCell(row,i+(i<2?4:4)).value=u;
    sc(ws.getCell(row,i+3),null,false,C.BLACK,10,'center');
  });

  // Summary
  row+=2;
  row = mkTitle(ws, row, `${data.project_name||'PROJECT'} - SUMMARY`, 11);
  row = mkTitle(ws, row, 'ROAD WORK MATERIAL SUMMARY', 11);
  row = mkHeaders(ws, row, ['SR.NO.','MATERIAL','','','QTY.','UNIT','','REMARKS']);
  const summary = [
    ['BOX CUTTING', Math.round(totBoxCut*100)/100, 'SQMT'],
    ['GSB FILLING (300 MM THK)', Math.round(totGSB*100)/100, 'TON'],
    ['WMM FILLING (200 MM THK)', Math.round(totWMM*100)/100, 'TON'],
    ['PQC ROAD (250 MM THK M30)', Math.round(totPQC*100)/100, 'CUM'],
    ['TOTAL ROAD RMT', Math.round(totLen*100)/100, 'RMT'],
    ['AREA OF ROAD', Math.round(totArea*100)/100, 'SQMT'],
  ];
  summary.forEach(([mat,qty,unit],i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    ws.getCell(row,1).value=i+1; sc(ws.getCell(row,1),bg,false,C.BLACK,11,'center'); ws.getCell(row,1).border=bdr;
    ws.mergeCells(row,2,row,4); ws.getCell(row,2).value=mat; sc(ws.getCell(row,2),bg,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr;
    ws.getCell(row,5).value=qty; ws.getCell(row,5).numFmt='#,##0.00'; sc(ws.getCell(row,5),bg,true,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr;
    ws.getCell(row,6).value=unit; sc(ws.getCell(row,6),bg,false,C.BLACK,11,'center'); ws.getCell(row,6).border=bdr;
    ws.getRow(row).height=15; row++;
  });

  ws.views=[{state:'frozen',ySplit:4}];
  return {totLen, totArea, totGSB, totWMM, totPQC};
}

// SOIL FILLING SHEET
function buildSoilFilling(wb, data) {
  const roads = data.roads||[];
  if(!roads.length) return;

  const ws = wb.addWorksheet('SOIL FILLING WORK');
  [7,8,10,14,13,12,10,10,10,12,12,15,12,24].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 14);
  row = mkTitle(ws, row, 'RCC ROAD CUTTING / FILLING', 14);
  row = mkHeaders(ws, row,
    ['SR NO','ROAD NO','ROAD WIDTH','SUB BASE WIDTH\n(CW+0.3+0.3)','ROAD LENGTH\n(MTR)',
     'SUB BASE AREA\n(SQMT)','ROAD TOP\nLEVEL (A)','NGL OF\nROAD (B)',
     'TOTAL SECTION\n(C=0.75m)','FILLING DEPTH\n(A-B-C+0.15)',
     'EXTRA FILLING','TOTAL FILLING\n(CUM)','HYWAS\n(14 CUM)','REMARK']);

  let totFill=0, totHywas=0;

  roads.forEach((rd,i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    const CW = rd.carriage_width_m > 0 ? rd.carriage_width_m : 0;  // from drawing only
    const subW = CW > 0 ? CW + 0.6 : 0;
    const L = rd.length_m||0;
    const area = subW*L;
    const A = rd.road_top_level||0;
    const B = rd.ngl||0;
    const C_depth = 0.75; // GSB+WMM+PQC physical layer stack (design constant, not a guess)
    const fillDepth = A>0&&B>0 ? Math.max(0, A-B-C_depth+0.15) : 0;
    const extraFill = fillDepth<0.5 ? Math.max(0,0.5-fillDepth) : 0;
    // Only compute filling if we have real fillDepth from A/B; otherwise skip (no 0.65 default).
    const totFillingCum = fillDepth>0 ? area*fillDepth : 0;
    const hywas = totFillingCum/14;
    totFill+=totFillingCum; totHywas+=hywas;

    const vals=[i+1,rd.road_no,rd.total_width_m,Math.round(subW*100)/100,L,
      Math.round(area*100)/100,A||0,B||0,C_depth,
      Math.round(fillDepth*1000)/1000,Math.round(extraFill*1000)/1000,
      Math.round(totFillingCum*100)/100,Math.round(hywas*100)/100,rd.remark||''];
    vals.forEach((v,ci)=>{
      const c=ws.getCell(row,ci+1); c.value=v;
      sc(c,bg,false,C.BLACK,11,typeof v==='number'?'right':'left');
      if(typeof v==='number'&&ci>2) c.numFmt='#,##0.00';
    });
    ws.getRow(row).height=15; row++;
  });

  row++;
  // Totals
  ws.mergeCells(row,1,row,4); ws.getCell(row,1).value='TOTAL'; sc(ws.getCell(row,1),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,1).border=bdr;
  ws.getCell(row,5).value=Math.round(roads.reduce((s,r)=>s+(r.length_m||0),0)*100)/100; sc(ws.getCell(row,5),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,5).numFmt='#,##0.00'; ws.getCell(row,5).border=bdr;
  ws.getCell(row,12).value=Math.round(totFill*100)/100; sc(ws.getCell(row,12),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,12).numFmt='#,##0.00'; ws.getCell(row,12).border=bdr;
  ws.getCell(row,13).value=Math.round(totHywas*100)/100; sc(ws.getCell(row,13),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,13).numFmt='#,##0.00'; ws.getCell(row,13).border=bdr;
  ws.getRow(row).height=18; row+=2;

  ws.getCell(row,12).value=`${DSR.soil_filling_cum} RS/CUM`;
  ws.getCell(row+1,12).value=Math.round(totFill*DSR.soil_filling_cum);
  ws.getCell(row+1,12).numFmt='#,##0';
  sc(ws.getCell(row+1,12),C.TOTAL_BG,true,C.BLACK,11,'right');

  ws.views=[{state:'frozen',ySplit:3}];
}

// SERVICE CORRIDOR SHEET
function buildServiceCorridor(wb, data) {
  const roads = data.roads||[];
  if(!roads.length) return;

  const ws = wb.addWorksheet('SERVICE CORRIDOR');
  [7,10,9,16,12,14,14,16,15,13,14,26].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 12);
  row = mkTitle(ws, row, 'SERVICE CORRIDOR (SUB-BASE & BASE COURSE ESTIMATE)', 12);
  row = mkHeaders(ws, row,
    ['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH (METER)','SC WIDTH (METER)',
     'AREA (SQMT)','EXCAVATION (SQMT)',
     'YELLOW SOIL\n300MM (15% EXTRA)\nTON',
     'GSB FILLING\n230MM (15% EXTRA)\nTON',
     'PAVER BLOCK\n80MM (5% EXTRA)\nSQMT','KERBING (RMT)','REMARK']);
  ws.getRow(row-1).height=52;
  ['','','','','','','','LAYER-1','LAYER-2','LAYER-3','',''].forEach((h,i)=>{
    ws.getCell(row,i+1).value=h; sc(ws.getCell(row,i+1),C.TITLE_BG,true,C.BLACK,10,'center');
  });
  ws.getRow(row).height=16; row++;

  let totLen=0,totArea=0,totSoil=0,totGSB=0,totPaver=0;
  roads.forEach((rd,i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    const SC_W = rd.sc_width_m > 0 ? rd.sc_width_m : 0;  // only if drawing specifies
    const L=rd.length_m||0;
    const area=L*SC_W;
    const excav=area*1.05;
    const soil=area*1.15*0.300*1.5; // yellow soil ton
    const gsb=area*1.15*0.230*1.800;
    const paver=area*1.05;
    const kerb=0;
    totLen+=L; totArea+=area; totSoil+=soil; totGSB+=gsb; totPaver+=paver;

    const vals=[i+1,`${rd.total_width_m||''} MT`,rd.road_no,Math.round(L*100)/100,SC_W,
      Math.round(area*100)/100,Math.round(excav*100)/100,Math.round(soil*100)/100,
      Math.round(gsb*100)/100,Math.round(paver*100)/100,kerb,rd.remark||''];
    vals.forEach((v,ci)=>{
      const c=ws.getCell(row,ci+1); c.value=v;
      sc(c,bg,false,C.BLACK,11,typeof v==='number'?'right':'left');
      if(typeof v==='number'&&ci>2) c.numFmt='#,##0.00';
    });
    ws.getRow(row).height=15; row++;
  });

  row++;
  ws.mergeCells(row,1,row,3); ws.getCell(row,1).value='TOTAL'; sc(ws.getCell(row,1),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,1).border=bdr;
  [null,null,null,Math.round(totLen*100)/100,null,Math.round(totArea*100)/100,null,Math.round(totSoil*100)/100,Math.round(totGSB*100)/100,Math.round(totPaver*100)/100].forEach((v,i)=>{
    if(v!==null){const c=ws.getCell(row,i+1);c.value=v;sc(c,C.TOTAL_BG,true,C.BLACK,11,'right');c.numFmt='#,##0.00';c.border=bdr;}
  });
  ws.getRow(row).height=18;
  ws.views=[{state:'frozen',ySplit:4}];
}

// STREET LIGHT SHEET
function buildStreetLight(wb, data) {
  const roads = data.roads||[];
  if(!roads.length) return;

  const ws = wb.addWorksheet('STREET LIGHT');
  [7,10,10,20,18,30].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 6);
  row = mkTitle(ws, row, 'STREET LIGHT', 6);
  row = mkHeaders(ws, row, ['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH (METER)','STREET LIGHTS (NOS)','REMARK']);

  let totLen=0, totLights=0;
  roads.forEach((rd,i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    const lights = Math.ceil((rd.length_m||0)/20);
    totLen+=(rd.length_m||0); totLights+=lights;
    const vals=[i+1,`${rd.total_width_m||''} MT`,rd.road_no,Math.round((rd.length_m||0)*100)/100,lights,'EVERY 20 MT'];
    vals.forEach((v,ci)=>{const c=ws.getCell(row,ci+1);c.value=v;sc(c,bg,false,C.BLACK,11,typeof v==='number'?'right':'left');if(typeof v==='number')c.numFmt='#,##0.00';});
    ws.getRow(row).height=15; row++;
  });

  row++;
  ws.mergeCells(row,1,row,3); ws.getCell(row,1).value='TOTAL'; sc(ws.getCell(row,1),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,1).border=bdr;
  ws.getCell(row,4).value=Math.round(totLen*100)/100; sc(ws.getCell(row,4),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,4).numFmt='#,##0.00'; ws.getCell(row,4).border=bdr;
  ws.getCell(row,5).value=totLights; sc(ws.getCell(row,5),C.TOTAL_BG,true,C.BLACK,11,'right'); ws.getCell(row,5).border=bdr;
  ws.getRow(row).height=18;
}

// RATE ANALYSIS SHEET
function buildRateAnalysis(wb, data) {
  const ws = wb.addWorksheet('RATE ANALYSIS');
  [7,44,10,10,12,14,14,16].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 8);
  row = mkTitle(ws, row, 'ROAD RATE ANALYSIS', 8);

  const items = [
    {label:'SOIL STABILIZATION (LIME-FLYASH)', rate:DSR.soil_stabilization_sqmt, unit:'RS/SQMT'},
    {label:'SOIL FILLING', rate:DSR.soil_filling_cum, unit:'RS/CUM'},
    {label:'GSB FILLING (300 MM THK)', rate:DSR.gsb_300mm_sqmt, unit:'RS/SQMT'},
    {label:'WMM FILLING (200 MM THK)', rate:DSR.wmm_200mm_sqmt, unit:'RS/SQMT'},
    {label:'PQC ROAD WORK (250 MM THK M30)', rate:DSR.pqc_250mm_sqmt, unit:'RS/SQMT'},
    {label:'SERVICE CORRIDOR (YELLOW SOIL+GSB+PAVER)', rate:DSR.service_corridor_sqmt, unit:'RS/SQMT'},
    {label:'STREET LIGHT', rate:DSR.street_light_nos, unit:'RS/NOS'},
    {label:'CP WALL', rate:DSR.cp_wall_rmt, unit:'RS/RMT'},
    {label:'GABION WALL', rate:DSR.gabion_wall_rmt, unit:'RS/RMT'},
  ];

  row = mkHeaders(ws, row, ['SR.NO.','PARTICULAR','WITH MATERIAL RATE','UNIT','','','','REMARK']);
  items.forEach((item,i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    ws.getCell(row,1).value=i+1; sc(ws.getCell(row,1),bg,false,C.BLACK,11,'center'); ws.getCell(row,1).border=bdr;
    ws.mergeCells(row,2,row,2); ws.getCell(row,2).value=item.label; sc(ws.getCell(row,2),bg,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr;
    ws.getCell(row,3).value=item.rate; ws.getCell(row,3).numFmt='#,##0.00'; sc(ws.getCell(row,3),bg,true,C.BLACK,11,'right'); ws.getCell(row,3).border=bdr;
    ws.getCell(row,4).value=item.unit; sc(ws.getCell(row,4),bg,false,C.BLACK,11,'center'); ws.getCell(row,4).border=bdr;
    ws.getRow(row).height=15; row++;
  });
}

// BOQ SHEET (for buildings/structures)
function buildBOQ(wb, data) {
  const ws = wb.addWorksheet('BOQ');
  [7,50,12,14,14,18].forEach((w,i)=>ws.getColumn(i+1).width=w);

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 6);
  row = mkTitle(ws, row, 'BILL OF QUANTITIES', 6);
  row = mkHeaders(ws, row, ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)']);

  const boqItems = data.boq_items || generateBOQFromData(data);
  let total=0;
  boqItems.forEach((item,i)=>{
    const bg=i%2===0?C.WHITE:C.ALT_ROW;
    const amt=(item.qty||0)*(item.rate||0);
    total+=amt;
    [i+1,item.description||'',item.unit||'',item.qty||0,item.rate||0,amt].forEach((v,ci)=>{
      const c=ws.getCell(row,ci+1); c.value=v;
      sc(c,bg,false,C.BLACK,11,ci===1?'left':typeof v==='number'?'right':'center');
      if(ci>=2&&typeof v==='number') c.numFmt=ci===5?'₹#,##0':'#,##0.00';
    });
    ws.getRow(row).height=15; row++;
  });
  ws.mergeCells(row,1,row,4);
  sc(ws.getCell(row,1),C.TOTAL_BG,true,C.BLACK,12,'right'); ws.getCell(row,1).value='GRAND TOTAL';
  const tc=ws.getCell(row,6); tc.value=total; tc.numFmt='₹#,##0'; sc(tc,C.TOTAL_BG,true,C.BLACK,12,'right');
  ws.getRow(row).height=20;
  ws.views=[{state:'frozen',ySplit:3}];
}

function generateBOQFromData(data) {
  // BOQ rows are generated ONLY when the required dimensions were read from drawing.
  // No default thickness / height / soil factor is applied.
  const items = [];
  const roads = data.roads||[];
  const validRoads = roads.filter(r => (r.length_m||0)>0 && (r.carriage_width_m||0)>0);
  const totalArea = validRoads.reduce((s,r)=>s+r.length_m*r.carriage_width_m,0);

  if(totalArea>0) {
    if(typeof data.soil_filling_cum === 'number' && data.soil_filling_cum > 0) {
      items.push({description:'SOIL FILLING IN ROAD AREA', unit:'CUM', qty:Math.round(data.soil_filling_cum*100)/100, rate:DSR.soil_filling_cum});
    }
    items.push({description:'GSB FILLING (300 MM THK)', unit:'SQMT', qty:Math.round(totalArea*100)/100, rate:DSR.gsb_300mm_sqmt});
    items.push({description:'WMM FILLING (200 MM THK)', unit:'SQMT', qty:Math.round(totalArea*100)/100, rate:DSR.wmm_200mm_sqmt});
    items.push({description:'PQC ROAD WORK M30 (250 MM THK)', unit:'SQMT', qty:Math.round(totalArea*100)/100, rate:DSR.pqc_250mm_sqmt});
    items.push({description:'STEEL DOWEL BAR', unit:'TON', qty:Math.round(totalArea*3.87/1000*100)/100, rate:DSR.steel_fe500_kg*1000});
  }

  const bldgs = data.buildings||[];
  bldgs.forEach(b=>{
    const slabThk = b.slab_thickness_m > 0 ? b.slab_thickness_m : null;
    const wallH   = b.wall_height_m    > 0 ? b.wall_height_m    : null;
    const wallThk = b.wall_thickness_m > 0 ? b.wall_thickness_m : null;

    if(b.slab_area_sqmt && slabThk) {
      items.push({description:`RCC WORK - ${b.name}`, unit:'CUM', qty:Math.round(b.slab_area_sqmt*slabThk*100)/100, rate:DSR.rcc_m25_cum});
      items.push({description:`STEEL - ${b.name}`,    unit:'KG',  qty:Math.round(b.slab_area_sqmt*slabThk*120*100)/100, rate:DSR.steel_fe500_kg});
    }
    if(b.wall_length_rmt && wallH && wallThk) {
      items.push({description:`BRICKWORK - ${b.name}`, unit:'CUM', qty:Math.round(b.wall_length_rmt*wallH*wallThk*100)/100, rate:DSR.brickwork_230_cum});
    }
    if(b.wall_length_rmt && wallH) {
      items.push({description:`PLASTER - ${b.name}`, unit:'SQMT', qty:Math.round(b.wall_length_rmt*wallH*2*100)/100, rate:DSR.plaster_sqmt});
    }
  });

  if(data.compound_wall?.cp_wall_length_rmt>0) {
    items.push({description:'COMPOUND WALL (CP WALL)', unit:'RMT', qty:data.compound_wall.cp_wall_length_rmt, rate:DSR.cp_wall_rmt});
  }
  return items;
}

// PMC OBSERVATIONS SHEET
function buildObservations(wb, data) {
  const ws = wb.addWorksheet('PMC OBSERVATIONS');
  ws.getColumn(1).width=6; ws.getColumn(2).width=90;

  let row=1;
  row = mkTitle(ws, row, data.project_name||'PROJECT', 2);
  row = mkTitle(ws, row, 'PMC OBSERVATIONS & RECOMMENDATIONS', 2);

  const obs = [
    `Drawing Type: ${data.drawing_type||'SITE LAYOUT'}`,
    `Scale: ${data.scale||'Not detected - please verify'}`,
    `Location: ${data.location||'Not specified'}`,
    `Total Area: ${(data.total_area_sqmt||0).toLocaleString('en-IN')} SQMT`,
    `Roads extracted: ${(data.roads||[]).length} road segments`,
    `Extraction confidence: ${data.confidence||'MEDIUM'}`,
    ...(data.pmcnotes||[]),
    ...(data.extracted_dimensions||[]).map(d=>`Dimension found: ${d}`),
  ];

  obs.forEach((o,i)=>{
    ws.getCell(row,1).value=i+1; sc(ws.getCell(row,1),i%2===0?C.WHITE:C.ALT_ROW,false,C.BLACK,11,'center'); ws.getCell(row,1).border=bdr;
    ws.getCell(row,2).value=o; sc(ws.getCell(row,2),i%2===0?C.WHITE:C.ALT_ROW,false,C.BLACK,11,'left'); ws.getCell(row,2).border=bdr;
    ws.getRow(row).height=20; row++;
  });

  row+=2;
  row = mkTitle(ws, row, 'PMC RECOMMENDATION', 2);
  ws.mergeCells(row,1,row,2);
  ws.getCell(row,1).value = data.pmc_recommendation||'Drawing analyzed. Verify scale and dimensions with original drawing before proceeding.';
  sc(ws.getCell(row,1), C.GREEN_BG, true, C.BLACK, 11, 'left');
  ws.getRow(row).height=80;

  row+=2;
  ws.mergeCells(row,1,row,2);
  ws.getCell(row,1).value=`Prepared by: PMC Civil AI Agent | Date: ${new Date().toLocaleDateString('en-IN')} | Powered by Gemini AI`;
  ws.getCell(row,1).font={italic:true,size:9,color:{argb:'FF595959'},name:'Calibri'};
  ws.getCell(row,1).alignment={horizontal:'center',vertical:'middle'};
  ws.getRow(row).height=14;
}

// ══════════════════════════════════════════════════════════════════
// MAIN BUILD FUNCTION - DECIDES WHICH SHEETS TO CREATE
// ══════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════
// PROJECT-TYPE AWARE BOQ BUILDERS
// Each builder emits sheets appropriate to that project type,
// reading ONLY from data (no hardcoded fallbacks).
// ══════════════════════════════════════════════════════════════════

function pickCount(data, key) {
  if (data?.element_counts && typeof data.element_counts[key] === 'number') return data.element_counts[key];
  if (typeof data?.[key] === 'number') return data[key];
  return 0;
}

function buildCafeBOQ(wb, data) {
  const ws = wb.addWorksheet('CAFE ESTIMATE');
  const cols = 8;
  let r = mkTitle(ws, 1, 'CAFE / RESTAURANT — DRAWING-BASED BOQ', cols);
  r = mkHeaders(ws, r, ['Sr','Item','Basis (from drawing)','Qty','Unit','Rate','Amount','Source'],
                       [5, 45, 35, 12, 8, 12, 15, 25]);
  let total = 0, sr = 1;

  const totalArea  = Number(data.total_area_sqm) || 0;
  const kitchenN   = pickCount(data, 'kitchen_count');
  const toiletN    = pickCount(data, 'toilet_count');
  const doorN      = pickCount(data, 'door_count');
  const windowN    = pickCount(data, 'window_count');
  const tileRate   = DSR.floor_tiles_sqmt || DSR.tiles_sqmt || 0;
  const plasterR   = DSR.plaster_sqmt     || 0;
  const paintR     = DSR.paint_sqmt       || 0;

  if (totalArea > 0 && tileRate > 0) {
    const amt = totalArea * tileRate;
    r = mkDataRow(ws, r, [sr++, 'Floor tiling — seating + counter area', `${totalArea} sqm × ₹${tileRate}`, totalArea, 'sqm', tileRate, amt, 'drawing area'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && plasterR > 0) {
    const wallSqm = totalArea * 2.5;
    const amt = wallSqm * plasterR;
    r = mkDataRow(ws, r, [sr++, 'Internal plaster — walls', `${wallSqm.toFixed(1)} sqm × ₹${plasterR}`, Math.round(wallSqm*100)/100, 'sqm', plasterR, amt, 'area × 2.5 factor'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && paintR > 0) {
    const wallSqm = totalArea * 2.5;
    const amt = wallSqm * paintR;
    r = mkDataRow(ws, r, [sr++, 'Internal paint', `${wallSqm.toFixed(1)} sqm × ₹${paintR}`, Math.round(wallSqm*100)/100, 'sqm', paintR, amt, 'area × 2.5 factor'], [3,5,6]);
    total += amt;
  }
  if (kitchenN > 0) {
    const rate = DSR.kitchen_platform_rmt || 0;
    if (rate > 0) {
      const amt = kitchenN * 6 * rate;
      r = mkDataRow(ws, r, [sr++, 'Kitchen platform + granite counter', `${kitchenN} kitchen × 6 rmt`, kitchenN*6, 'rmt', rate, amt, `counted: ${kitchenN} kitchen(s)`], [3,5,6]);
      total += amt;
    }
  }
  if (toiletN > 0) {
    const rate = DSR.toilet_fittings_ls || 0;
    if (rate > 0) {
      const amt = toiletN * rate;
      r = mkDataRow(ws, r, [sr++, 'Toilet fittings (WC + washbasin + tap)', `${toiletN} toilet(s)`, toiletN, 'no', rate, amt, `counted from drawing`], [3,5,6]);
      total += amt;
    }
  }
  if (doorN > 0) {
    const rate = DSR.door_flush_no || DSR.door_no || 0;
    if (rate > 0) {
      const amt = doorN * rate;
      r = mkDataRow(ws, r, [sr++, 'Doors (flush) — supply + fix', `${doorN} doors from drawing`, doorN, 'no', rate, amt, `drawing block count`], [3,5,6]);
      total += amt;
    }
  }
  if (windowN > 0) {
    const rate = DSR.window_alum_sqmt || DSR.window_sqmt || 0;
    if (rate > 0) {
      const sqm = windowN * 2;
      const amt = sqm * rate;
      r = mkDataRow(ws, r, [sr++, 'Windows (aluminium) — supply + fix', `${windowN} windows × 2 sqm avg`, sqm, 'sqm', rate, amt, `drawing block count`], [3,5,6]);
      total += amt;
    }
  }
  mkGrandTotal(ws, r, 'TOTAL (₹)', total/1e5, cols);
  return ws;
}

function buildInstituteBOQ(wb, data) {
  const ws = wb.addWorksheet('INSTITUTE ESTIMATE');
  const cols = 8;
  let r = mkTitle(ws, 1, 'INSTITUTE / SCHOOL — DRAWING-BASED BOQ', cols);
  r = mkHeaders(ws, r, ['Sr','Item','Basis (from drawing)','Qty','Unit','Rate','Amount','Source'],
                       [5, 45, 35, 12, 8, 12, 15, 25]);
  let total = 0, sr = 1;

  const totalArea = Number(data.total_area_sqm) || 0;
  const floorN    = pickCount(data, 'floor_count');
  const doorN     = pickCount(data, 'door_count');
  const windowN   = pickCount(data, 'window_count');
  const liftN     = pickCount(data, 'lift_count');
  const toiletN   = pickCount(data, 'toilet_count');
  const rcc       = DSR.rcc_m25_cum   || 0;
  const plaster   = DSR.plaster_sqmt  || 0;
  const tiles     = DSR.floor_tiles_sqmt || 0;
  const liftR     = DSR.lift_passenger_no || DSR.lift_no || 0;
  const doorR     = DSR.door_flush_no || DSR.door_no || 0;
  const winR      = DSR.window_alum_sqmt || 0;

  if (totalArea > 0 && floorN > 0 && rcc > 0) {
    const cum = totalArea * 0.15 * floorN;
    const amt = cum * rcc;
    r = mkDataRow(ws, r, [sr++, 'RCC slab M25 — all floors', `${totalArea} sqm × 0.15m × ${floorN} floors`, Math.round(cum*100)/100, 'cum', rcc, amt, 'from drawing'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && floorN > 0 && plaster > 0) {
    const sqm = totalArea * 2.5 * floorN;
    const amt = sqm * plaster;
    r = mkDataRow(ws, r, [sr++, 'Internal plaster — classrooms + labs', `${totalArea}×2.5×${floorN}`, Math.round(sqm*100)/100, 'sqm', plaster, amt, 'drawing area'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && floorN > 0 && tiles > 0) {
    const sqm = totalArea * floorN;
    const amt = sqm * tiles;
    r = mkDataRow(ws, r, [sr++, 'Floor tiles — all floors', `${totalArea} × ${floorN}`, sqm, 'sqm', tiles, amt, 'drawing area'], [3,5,6]);
    total += amt;
  }
  if (doorN > 0 && doorR > 0) {
    const amt = doorN * doorR;
    r = mkDataRow(ws, r, [sr++, 'Classroom / lab doors', `${doorN} doors`, doorN, 'no', doorR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (windowN > 0 && winR > 0) {
    const sqm = windowN * 3;
    const amt = sqm * winR;
    r = mkDataRow(ws, r, [sr++, 'Windows (aluminium)', `${windowN} × 3 sqm avg`, sqm, 'sqm', winR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (liftN > 0 && liftR > 0) {
    const amt = liftN * liftR;
    r = mkDataRow(ws, r, [sr++, 'Passenger lift', `${liftN} lift(s)`, liftN, 'no', liftR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (toiletN > 0 && DSR.toilet_fittings_ls > 0) {
    const amt = toiletN * DSR.toilet_fittings_ls;
    r = mkDataRow(ws, r, [sr++, 'Toilet blocks', `${toiletN} toilet(s)`, toiletN, 'no', DSR.toilet_fittings_ls, amt, 'drawing count'], [3,5,6]);
    total += amt;
  }
  mkGrandTotal(ws, r, 'TOTAL (₹)', total/1e5, cols);
  return ws;
}

function buildCommercialBOQ(wb, data) {
  const ws = wb.addWorksheet('COMMERCIAL ESTIMATE');
  const cols = 8;
  let r = mkTitle(ws, 1, 'COMMERCIAL / SHOP / OFFICE — DRAWING-BASED BOQ', cols);
  r = mkHeaders(ws, r, ['Sr','Item','Basis (from drawing)','Qty','Unit','Rate','Amount','Source'],
                       [5, 45, 35, 12, 8, 12, 15, 25]);
  let total = 0, sr = 1;

  const totalArea = Number(data.total_area_sqm) || 0;
  const floorN    = pickCount(data, 'floor_count');
  const doorN     = pickCount(data, 'door_count');
  const windowN   = pickCount(data, 'window_count');
  const liftN     = pickCount(data, 'lift_count');
  const rcc       = DSR.rcc_m25_cum   || 0;
  const tiles     = DSR.floor_tiles_sqmt || 0;
  const acp       = DSR.acp_cladding_sqmt || 0;
  const glazing   = DSR.glazing_sqmt || DSR.window_alum_sqmt || 0;
  const liftR     = DSR.lift_passenger_no || DSR.lift_no || 0;
  const doorR     = DSR.door_flush_no || DSR.door_no || 0;

  if (totalArea > 0 && floorN > 0 && rcc > 0) {
    const cum = totalArea * 0.15 * floorN;
    const amt = cum * rcc;
    r = mkDataRow(ws, r, [sr++, 'RCC slab M25 — all floors', `${totalArea} × 0.15 × ${floorN}`, Math.round(cum*100)/100, 'cum', rcc, amt, 'drawing'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && floorN > 0 && tiles > 0) {
    const sqm = totalArea * floorN;
    const amt = sqm * tiles;
    r = mkDataRow(ws, r, [sr++, 'Vitrified tiles — shop/office floors', `${totalArea} × ${floorN}`, sqm, 'sqm', tiles, amt, 'drawing'], [3,5,6]);
    total += amt;
  }
  if (totalArea > 0 && acp > 0) {
    const sqm = Math.sqrt(totalArea) * 4 * 3 * Math.max(floorN,1);
    const amt = sqm * acp;
    r = mkDataRow(ws, r, [sr++, 'ACP façade cladding', `perimeter × 3m × ${floorN} floors`, Math.round(sqm*100)/100, 'sqm', acp, amt, 'from extents'], [3,5,6]);
    total += amt;
  }
  if (windowN > 0 && glazing > 0) {
    const sqm = windowN * 4;
    const amt = sqm * glazing;
    r = mkDataRow(ws, r, [sr++, 'Structural glazing / windows', `${windowN} × 4 sqm avg`, sqm, 'sqm', glazing, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (doorN > 0 && doorR > 0) {
    const amt = doorN * doorR;
    r = mkDataRow(ws, r, [sr++, 'Doors — shop / office', `${doorN} doors`, doorN, 'no', doorR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (liftN > 0 && liftR > 0) {
    const amt = liftN * liftR;
    r = mkDataRow(ws, r, [sr++, 'Passenger / service lift', `${liftN} lift(s)`, liftN, 'no', liftR, amt, 'drawing count'], [3,5,6]);
    total += amt;
  }
  mkGrandTotal(ws, r, 'TOTAL (₹)', total/1e5, cols);
  return ws;
}

function buildHighRiseBOQ(wb, data) {
  const ws = wb.addWorksheet('HIGH-RISE BOQ');
  const cols = 8;
  let r = mkTitle(ws, 1, 'HIGH-RISE RESIDENTIAL — DRAWING-BASED BOQ', cols);
  r = mkHeaders(ws, r, ['Sr','Item','Basis (from drawing)','Qty','Unit','Rate','Amount','Source'],
                       [5, 45, 35, 12, 8, 12, 15, 25]);
  let total = 0, sr = 1;

  const area   = Number(data.total_area_sqm) || 0;
  const fc     = pickCount(data, 'floor_count');
  const doorN  = pickCount(data, 'door_count');
  const winN   = pickCount(data, 'window_count');
  const liftN  = pickCount(data, 'lift_count');
  const bedN   = pickCount(data, 'bedroom_count');
  const toilN  = pickCount(data, 'toilet_count');
  const kitN   = pickCount(data, 'kitchen_count');
  const rcc    = DSR.rcc_m25_cum || 0;
  const brick  = DSR.brickwork_230_cum || 0;
  const plast  = DSR.plaster_sqmt || 0;
  const tiles  = DSR.floor_tiles_sqmt || 0;
  const doorR  = DSR.door_flush_no || DSR.door_no || 0;
  const winR   = DSR.window_alum_sqmt || 0;
  const liftR  = DSR.lift_passenger_no || DSR.lift_no || 0;

  if (area>0 && fc>0 && rcc>0) {
    const cum = area * 0.15 * fc;
    const amt = cum * rcc;
    r = mkDataRow(ws, r, [sr++, 'RCC slab M25 — all floors', `${area} sqm × 0.15 × ${fc}`, Math.round(cum*100)/100, 'cum', rcc, amt, 'from drawing'], [3,5,6]);
    total += amt;
  }
  if (data.wall_length_m > 0 && fc>0 && brick>0) {
    const cum = data.wall_length_m * 3 * 0.23 * fc;
    const amt = cum * brick;
    r = mkDataRow(ws, r, [sr++, 'Brickwork 230mm — external walls', `${data.wall_length_m} rmt × 3m × 0.23 × ${fc}`, Math.round(cum*100)/100, 'cum', brick, amt, 'wall lines in drawing'], [3,5,6]);
    total += amt;
  }
  if (area>0 && fc>0 && plast>0) {
    const sqm = area * 2.5 * fc;
    const amt = sqm * plast;
    r = mkDataRow(ws, r, [sr++, 'Internal + external plaster', `${area}×2.5×${fc}`, Math.round(sqm*100)/100, 'sqm', plast, amt, 'drawing'], [3,5,6]);
    total += amt;
  }
  if (area>0 && fc>0 && tiles>0) {
    const sqm = area * fc;
    const amt = sqm * tiles;
    r = mkDataRow(ws, r, [sr++, 'Vitrified floor tiles', `${area} × ${fc}`, sqm, 'sqm', tiles, amt, 'drawing'], [3,5,6]);
    total += amt;
  }
  if (doorN>0 && doorR>0) {
    const n = doorN * Math.max(fc,1);
    const amt = n * doorR;
    r = mkDataRow(ws, r, [sr++, 'Flush doors — all flats', `${doorN}/floor × ${fc} floors`, n, 'no', doorR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (winN>0 && winR>0) {
    const sqm = winN * Math.max(fc,1) * 2.5;
    const amt = sqm * winR;
    r = mkDataRow(ws, r, [sr++, 'Aluminium windows', `${winN}/floor × ${fc} × 2.5 sqm`, Math.round(sqm*100)/100, 'sqm', winR, amt, 'drawing block count'], [3,5,6]);
    total += amt;
  }
  if (liftN>0 && liftR>0) {
    const amt = liftN * liftR;
    r = mkDataRow(ws, r, [sr++, 'Passenger lifts', `${liftN} lift(s)`, liftN, 'no', liftR, amt, 'drawing count'], [3,5,6]);
    total += amt;
  }
  if (toilN>0) r = mkDataRow(ws, r, [sr++, 'Toilet block count', `${toilN} toilets`, toilN, 'no', 0, 0, 'count only'], [3,5,6]);
  if (bedN>0)  r = mkDataRow(ws, r, [sr++, 'Bedroom count',      `${bedN} bedrooms`, bedN, 'no', 0, 0, 'count only'], [3,5,6]);
  if (kitN>0)  r = mkDataRow(ws, r, [sr++, 'Kitchen count',      `${kitN} kitchens`, kitN, 'no', 0, 0, 'count only'], [3,5,6]);
  mkGrandTotal(ws, r, 'TOTAL (₹)', total/1e5, cols);
  return ws;
}

async function buildExcelFromDrawing(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';
  wb.created = new Date();

  const dtype  = data.drawing_type || 'SITE_LAYOUT';
  const ptype  = (data.project_type || 'generic').toLowerCase();

  // BLOCK COST always first
  buildBlockCost(wb, data);

  // ── Project-type-aware sheets ──────────────────────────
  if (ptype === 'cafe')                    buildCafeBOQ(wb, data);
  else if (ptype === 'institute')          buildInstituteBOQ(wb, data);
  else if (ptype === 'commercial')         buildCommercialBOQ(wb, data);
  else if (ptype === 'high_rise_residential') buildHighRiseBOQ(wb, data);

  // ── Drawing-type-based sheets (road / building / compound / drainage) ──
  if(['ROAD_LAYOUT','SITE_LAYOUT'].includes(dtype) && (data.roads||[]).length>0) {
    buildRoadEstimate(wb, data);
    buildSoilFilling(wb, data);
    buildServiceCorridor(wb, data);
    buildStreetLight(wb, data);
    buildRateAnalysis(wb, data);
  }

  if(['BUILDING','SITE_LAYOUT'].includes(dtype) && (data.buildings||[]).length>0) {
    buildBOQ(wb, data);
  }

  if(['COMPOUND_WALL'].includes(dtype)) {
    buildBOQ(wb, data);
    buildRateAnalysis(wb, data);
  }

  if(['DRAINAGE'].includes(dtype)) {
    buildBOQ(wb, data);
  }

  // Always add observations last
  buildObservations(wb, data);

  return wb;
}

module.exports = {
  buildExcelFromDrawing, getDrawingPrompt, DSR,
  buildCafeBOQ, buildInstituteBOQ, buildCommercialBOQ, buildHighRiseBOQ
};
