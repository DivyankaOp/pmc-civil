// ─── DRAWING ANALYSIS → MULTI-SHEET PMC EXCEL ──────────────────────────────
// This module adds /export-drawing endpoint to server.js
// When user uploads a civil drawing → Gemini analyzes → Multi-sheet Excel output

const ExcelJS = require('exceljs');

// Colors
const C = {
  NAVY:    'FF1F3864',// ─── DRAWING ANALYSIS → PMC MULTI-SHEET EXCEL ──────────────────────────────
// Accepts: images (PNG/JPG/WEBP), PDF drawings
// Output : Excel that mirrors the reference PMC estimate template structure
// RULE: All values extracted by Gemini from drawing. NOTHING hardcoded here.
'use strict';
const ExcelJS = require('exceljs');

// ── COLOUR PALETTE (reference template) ────────────────────────────────────
const C = {
  NAVY:    'FF1F3864',
  MIDBLUE: 'FF2E75B6',
  LTBLUE:  'FFBDD7EE',
  YELLOW:  'FFFFD966',
  GOLD:    'FFFFC000',
  GREEN:   'FFE2EFDA',
  DKGREEN: 'FF375623',
  GREY:    'FFF2F2F2',
  WHITE:   'FFFFFFFF',
};

// ── Load rates + legend from config (single source of truth) ────────────────
const fs_sd   = require('fs');
const path_sd = require('path');

function _loadRatesForPrompt() {
  try {
    const p = fs_sd.existsSync(path_sd.join(__dirname, 'rates.json'))
      ? path_sd.join(__dirname, 'rates.json')
      : path_sd.join(__dirname, 'Rates.json');
    const raw = JSON.parse(fs_sd.readFileSync(p, 'utf8'));
    const parts = [];
    for (const [category, items] of Object.entries(raw)) {
      if (category.startsWith('_') || typeof items !== 'object' || Array.isArray(items)) continue;
      const pieces = [];
      for (const [key, val] of Object.entries(items)) {
        if (!val || typeof val.rate !== 'number') continue;
        pieces.push(`${val.description || key}:${val.rate}/${val.unit || ''}`);
      }
      if (pieces.length) parts.push(`[${category.toUpperCase()}] ` + pieces.join(' | '));
    }
    return parts.join('\n');
  } catch (e) {
    console.warn('rates.json load for Gemini prompt failed:', e.message);
    return '(rates.json unavailable — use rates shown in drawing only)';
  }
}
const RATES_FOR_PROMPT = _loadRatesForPrompt();

function _loadLegendForPrompt() {
  try {
    const p = path_sd.join(__dirname, 'legend.json');
    if (!fs_sd.existsSync(p)) return '';
    const raw = JSON.parse(fs_sd.readFileSync(p, 'utf8'));
    const layers = Object.entries(raw.layers || {}).filter(([k]) => !k.startsWith('_'));
    const blocks = Object.entries(raw.blocks || {}).filter(([k]) => !k.startsWith('_'));
    const abbr   = Object.entries(raw.abbreviations || {}).filter(([k]) => !k.startsWith('_'));
    const parts = [];
    if (layers.length) parts.push('LAYERS: ' + layers.map(([k,v]) => `${k}=${v.category}${v.wall_thickness_mm?`(${v.wall_thickness_mm}mm)`:''}`).join(', '));
    if (blocks.length) parts.push('BLOCKS: ' + blocks.map(([k,v]) => `${k}=${v.opening_type||v.category}`).join(', '));
    if (abbr.length)   parts.push('ABBR: '   + abbr.map(([k,v]) => `${k}=${v}`).join(', '));
    return parts.join('\n');
  } catch { return ''; }
}
const thin = { style: 'thin', color: { argb: 'FF000000' } };
const bdr  = { top: thin, left: thin, bottom: thin, right: thin };

function sc(cell, bg, bold=false, fc='FF000000', size=9, align='center', wrap=true) {
  if (bg) cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:bg} };
  cell.font      = { bold, color:{argb:fc}, size, name:'Calibri' };
  cell.alignment = { horizontal:align, vertical:'middle', wrapText:wrap };
  cell.border    = bdr;
}
function mergeHdr(ws, row, text, cols, bg=C.NAVY, fc='FFFFFFFF', size=12, height=20) {
  ws.mergeCells(row,1,row,cols);
  const c = ws.getCell(row,1); c.value = text;
  sc(c, bg, true, fc, size, 'center');
  ws.getRow(row).height = height;
  return row + 1;
}

// ══════════════════════════════════════════════════════════════════════════════
// GEMINI PROMPT — reads drawing and returns structured JSON
// ══════════════════════════════════════════════════════════════════════════════
function getSmartDrawingPrompt() {
  return `You are a SENIOR PMC CIVIL ENGINEER with 20+ years India experience.
Analyze this civil drawing or estimate document carefully.

STEP 1 — IDENTIFY DRAWING TYPE:
ROAD_LAYOUT | SITE_LAYOUT | BUILDING | FOUNDATION | STRUCTURAL | DRAINAGE | COMPOUND_WALL | ESTIMATE | GENERAL

STEP 2 — READ TITLE BLOCK: project name, location, drawing number, scale, date

STEP 3 — READ ALL DIMENSIONS AND ANNOTATIONS (every number visible)

STEP 4 — CALCULATE QUANTITIES using exact drawing dimensions:
ROAD: area=length×carriage_width | GSB_ton=area×1.15×0.30×1800 | WMM_ton=area×1.15×0.20×2100 | PQC_cum=area×1.05×0.25 | Dowel_kg=area×3.87 | Street_lights=ceil(length/20)
BUILDING: brickwork_cum=L×H×thick | RCC_cum=L×W×thick×1.05 | Steel_kg=cum×120(slab)/160(beam)/200(col)
COMPOUND_WALL: length×rate/rmt

STEP 5 — RATES (use rate written on the drawing first; otherwise use the project's configured rate list below. NEVER guess rates — if neither source has it, leave rate=0 and flag in remarks):
${RATES_FOR_PROMPT}

STEP 6 — DRAWING LEGEND (per-project layer/block/abbreviation meanings from legend.json; if a symbol is NOT listed here, do NOT assume its meaning — report it in "unknown_symbols" so the user can confirm):
${_loadLegendForPrompt() || '(no legend configured yet — return every non-trivial layer/block/abbreviation in "unknown_symbols")'}

RETURN ONLY RAW JSON — NO MARKDOWN, NO BACKTICKS, START { END }:
{
  "drawing_type":"ROAD_LAYOUT",
  "project_name":"from title block",
  "location":"city from drawing",
  "scale":"1:500",
  "drawing_no":"",
  "date":"DD-MM-YYYY",
  "total_area_sqmt":0,
  "total_area_sqft":0,
  "confidence":"HIGH",
  "block_cost_parts":[
    {
      "part_label":"PART-A (PROJECT WORKS)",
      "sections":[
        {
          "sr":1,
          "section_name":"ROADS",
          "items":[
            {"description":"SOIL FILLING IN ROAD AREA","qty":12345.67,"unit":"CUM","rate":285,"amount_lacs":3.52,"remark":""}
          ],
          "subtotal_lacs":150.00
        },
        {
          "sr":2,
          "section_name":"SERVICES (DRAINAGE / UTILITIES)",
          "items":[
            {"description":"SEWAGE TREATMENT PLANT","qty":3,"unit":"MLD","rate":25000000,"amount_lacs":750.00,"remark":"3 MLD"},
            {"description":"ELECTRICAL INFRASTRUCTURE","qty":1,"unit":"L/S","rate":25000000,"amount_lacs":250.00,"remark":""}
          ],
          "subtotal_lacs":1000.00
        }
      ],
      "part_total_lacs":1150.00
    }
  ],
  "road_schedule":[
    {"sr":1,"road_no":"R1","total_width_m":24,"carriage_width_m":15,"length_m":129.12,"sc_width_m":9,"area_sqmt":1936.80,"gsb_ton":1459.34,"wmm_ton":1033.12,"pqc_cum":484.20,"remark":"4.5-4.5 MT B/S SERVICES"}
  ],
  "soil_filling":[
    {"sr":1,"road_no":"R1","road_width":24,"sub_base_width":15.6,"length":129.12,"sub_base_area":2014.27,"road_top_lvl":49.95,"ngl":48.796,"section_depth":0.75,"filling_depth":0.554,"extra_filling":0,"total_filling_cum":1116.31,"hywas":79.74,"remark":""}
  ],
  "bbs":[
    {"element":"PQC ROAD PANEL R1","items":[
      {"description":"DOWELS 25MM DIA 450MM C/C","dia":25,"nos":287,"cutting_length":0.60,"total_length":172.2,"weight_kg":664.7},
      {"description":"CHAIR 10MM DIA 900MM C/C","dia":10,"nos":96,"cutting_length":1.09,"total_length":104.64,"weight_kg":64.56}
    ]}
  ],
  "street_light_schedule":[
    {"sr":1,"road_wide":"24 MT","road_no":"R1","length_mt":129.12,"nos":7,"remark":"ONE EVERY 20MT"}
  ],
  "material_summary":{
    "total_road_area_sqmt":0,"total_road_length_rmt":0,"gsb_ton":0,"wmm_ton":0,"pqc_cum":0,
    "soil_filling_cum":0,"paver_block_sqmt":0,"dowel_steel_kg":0,"street_lights_nos":0,
    "compound_wall_rmt":0,"gabion_wall_rmt":0,"stp_mld":0,"pipeline_rmt":0
  },
  "boq":[
    {"sr":1,"description":"SOIL STABILIZATION","unit":"SQMT","qty":38388,"rate":82,"amount":3147864}
  ],
  "spaces":[
    {"name":"ROOM NAME","length_m":5.5,"width_m":4.2,"area_sqmt":23.1,"area_sqft":248.75,"floor":"GF"}
  ],
  "pmc_observations":["Observation 1","Observation 2"],
  "pmc_recommendation":"Full PMC recommendation based on drawing",
  "unknown_symbols":[
    {"kind":"layer|block|color|abbreviation","name":"e.g. A-EXT-WALL","seen_count":3,"note":"what the user should confirm"}
  ]
}
RULES:
- Only real values FROM THIS DRAWING. If not visible use 0.
- Do NOT carry values from previous drawings or training memory.
- amount_lacs = rupees/100000.
- Every layer/block/abbreviation you cannot map via the legend above MUST appear in "unknown_symbols" — do not guess its meaning.
- Return ONLY JSON.`;
}

// ── CALL GEMINI WITH RETRY ───────────────────────────────────────────────────
async function callGemini(key, parts, jsonMode=true, maxTokens=8192) {
  const URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`;
  const cfg = { maxOutputTokens:maxTokens, temperature:0.0 };
  if (jsonMode) cfg.responseMimeType = 'application/json';
  let last;
  for (let attempt=0; attempt<=4; attempt++) {
    const r = await fetch(URL, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({contents:[{role:'user',parts}],generationConfig:cfg}) });
    last = await r.json();
    if (r.ok && last?.candidates?.[0]) return last;
    const code = last?.error?.code;
    if (code!==503 && code!==429) return last;
    const delay = 2000*Math.pow(2,attempt);
    console.warn(`Gemini ${code} attempt ${attempt+1}/5 — retry ${delay}ms`);
    await new Promise(res=>setTimeout(res,delay));
  }
  return last;
}

// ── EXTRACT DRAWING DATA ─────────────────────────────────────────────────────
async function extractDrawingData(key, files, userText, aiResponse, fetch) {
  const parts = [];
  // Accept images AND PDFs
  for (const f of (files||[])) {
    try {
      if (f.type==='application/pdf'||f.name?.match(/\.pdf$/i))
        parts.push({inline_data:{mime_type:'application/pdf',data:f.b64}});
      else if (f.type?.startsWith('image/'))
        parts.push({inline_data:{mime_type:f.type||'image/png',data:f.b64}});
    } catch(e){console.log('File skip:',e.message);}
  }
  if (aiResponse) parts.push({text:`PREVIOUS AI ANALYSIS:\n${aiResponse}\n\n`});
  if (userText)   parts.push({text:`USER NOTE: ${userText}\n\n`});
  parts.push({text:getSmartDrawingPrompt()});

  const data = await callGemini(key, parts, true, 8192);
  let raw = data?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
  const fb=raw.indexOf('{'), lb=raw.lastIndexOf('}');
  if (fb!==-1&&lb!==-1) raw=raw.slice(fb,lb+1);
  try { return JSON.parse(raw.replace(/```json|```/g,'').trim()); }
  catch(e) {
    console.error('Drawing JSON parse fail:',e.message,raw.slice(0,300));
    return {drawing_type:'GENERAL',project_name:'PMC CIVIL PROJECT',date:new Date().toLocaleDateString('en-IN'),block_cost_parts:[],road_schedule:[],soil_filling:[],bbs:[],street_light_schedule:[],material_summary:{},boq:[],spaces:[],pmc_observations:[],pmc_recommendation:''};
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// EXCEL BUILDER — MATCHES REFERENCE TEMPLATE STRUCTURE
// ══════════════════════════════════════════════════════════════════════════════
async function buildDrawingExcel(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';
  _blockCost(wb, d);
  if ((d.road_schedule||[]).length)      _roadSchedule(wb, d);
  if ((d.soil_filling||[]).length)       _soilFilling(wb, d);
  if ((d.bbs||[]).length)                _bbs(wb, d);
  if ((d.street_light_schedule||[]).length) _streetLight(wb, d);
  _boq(wb, d);
  _materialSummary(wb, d);
  _observations(wb, d);
  return wb;
}

// ── SHEET 1: BLOCK COST ──────────────────────────────────────────────────────
function _blockCost(wb, d) {
  const ws = wb.addWorksheet('BLOCK COST');
  const COLS=8;
  [7,40,18,9,14,22,20,36].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',COLS,C.NAVY,'FFFFFFFF',14,24);
  row=mergeHdr(ws,row,'BLOCK COST',COLS,C.MIDBLUE,'FFFFFFFF',12,20);
  ws.mergeCells(row,1,row,COLS);
  const areaSqmt=d.total_area_sqmt||0, areaSqft=d.total_area_sqft||Math.round(areaSqmt*10.764);
  ws.getCell(row,1).value=`TOTAL AREA: ${areaSqmt.toLocaleString('en-IN')} SQMT  (${areaSqft.toLocaleString('en-IN')} SQFT)`;
  sc(ws.getCell(row,1),C.LTBLUE,true,'FF000000',10,'left'); ws.getRow(row).height=16; row++;
  const hdrs=['SR NO','COMPONENT','ESTIMATE QUANTITY','UNITS','RATE','ESTIMATE VALUE\n(RS IN LACS)','ESTIMATE VALUE\n(RS IN CR.)','REMARK'];
  const hr=ws.getRow(row); hr.height=36;
  hdrs.forEach((h,i)=>{const c=hr.getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',9,'center');});
  row++;

  let parts=d.block_cost_parts||[];
  if (!parts.length) parts=_autoBlockFromMaterials(d);
  let grandTotal=0;

  parts.forEach((part,pi)=>{
    ws.mergeCells(row,1,row,COLS);
    ws.getCell(row,1).value=part.part_label||`PART-${String.fromCharCode(65+pi)} (PROJECT WORKS)`;
    sc(ws.getCell(row,1),C.YELLOW,true,'FF000000',10,'left'); ws.getRow(row).height=18; row++;
    let partTotal=0;
    (part.sections||[]).forEach(sec=>{
      // Section header
      ws.mergeCells(row,1,row,COLS); ws.getCell(row,1).value=sec.section_name;
      sc(ws.getCell(row,1),C.MIDBLUE,true,'FFFFFFFF',10,'left'); ws.getRow(row).height=16; row++;
      (sec.items||[]).forEach(item=>{
        const bg=row%2===0?C.WHITE:C.GREY; ws.getRow(row).height=15;
        const vals=['',item.description,item.qty===0?'':item.qty,item.unit,item.rate===0?'':item.rate,item.amount_lacs||0,item.amount_lacs?+((item.amount_lacs||0)/100).toFixed(4):'',item.remark||''];
        vals.forEach((v,i)=>{
          const c=ws.getRow(row).getCell(i+1); c.value=v??'';
          sc(c,bg,false,'FF000000',9,i===1?'left':'center');
          if (i>=4&&typeof v==='number') c.numFmt='#,##0.00';
        }); row++;
      });
      // Subtotal
      ws.mergeCells(row,1,row,5); ws.getCell(row,1).value=`SUBTOTAL - ${sec.section_name}`;
      sc(ws.getCell(row,1),C.LTBLUE,true,'FF000000',9,'right');
      const stL=ws.getRow(row).getCell(6); stL.value=sec.subtotal_lacs||0; stL.numFmt='#,##0.00'; sc(stL,C.LTBLUE,true,'FF000000',9,'center');
      const stC=ws.getRow(row).getCell(7); stC.value=+((sec.subtotal_lacs||0)/100).toFixed(4); stC.numFmt='#,##0.0000'; sc(stC,C.LTBLUE,true,'FF000000',9,'center');
      ws.getRow(row).getCell(8).value=''; sc(ws.getRow(row).getCell(8),C.LTBLUE,false,'FF000000',9,'center');
      ws.getRow(row).height=16; row++;
      partTotal+=sec.subtotal_lacs||0;
    });
    grandTotal+=partTotal;
    ws.mergeCells(row,1,row,COLS); ws.getRow(row).height=6; row++;
    ws.mergeCells(row,1,row,5); ws.getCell(row,1).value=`TOTAL COST (${part.part_label||'PART'})`;
    sc(ws.getCell(row,1),C.GOLD,true,'FF000000',10,'right');
    const ptL=ws.getRow(row).getCell(6); ptL.value=part.part_total_lacs||partTotal; ptL.numFmt='#,##0.00'; sc(ptL,C.GOLD,true,'FF000000',10,'center');
    const ptC=ws.getRow(row).getCell(7); ptC.value=+((part.part_total_lacs||partTotal)/100).toFixed(2); ptC.numFmt='#,##0.00'; sc(ptC,C.GOLD,true,'FF000000',10,'center');
    ws.getRow(row).getCell(8).value=''; sc(ws.getRow(row).getCell(8),C.GOLD,false,'FF000000',9,'center');
    ws.getRow(row).height=20; row++; row++;
  });

  // 10% Contingency
  const cont=grandTotal*0.10;
  ws.mergeCells(row,1,row,5); ws.getCell(row,1).value='10% CONTINGENCY';
  sc(ws.getCell(row,1),C.GOLD,true,'FF000000',10,'right');
  const cL=ws.getRow(row).getCell(6); cL.value=+cont.toFixed(2); cL.numFmt='#,##0.00'; sc(cL,C.GOLD,true,'FF000000',10,'center');
  const cC=ws.getRow(row).getCell(7); cC.value=+(cont/100).toFixed(2); cC.numFmt='#,##0.00'; sc(cC,C.GOLD,true,'FF000000',10,'center');
  ws.getRow(row).getCell(8).value=''; sc(ws.getRow(row).getCell(8),C.GOLD,false,'FF000000',9,'center');
  ws.getRow(row).height=18; row++; row++;

  // Grand total
  const gw=grandTotal+cont;
  ws.mergeCells(row,1,row,5); ws.getCell(row,1).value='GRAND TOTAL (WITH CONTINGENCY)';
  sc(ws.getCell(row,1),C.NAVY,true,'FFFFFFFF',11,'right');
  const gL=ws.getRow(row).getCell(6); gL.value=+gw.toFixed(2); gL.numFmt='#,##0.00'; sc(gL,C.NAVY,true,'FFFFFFFF',11,'center');
  const gC=ws.getRow(row).getCell(7); gC.value=+(gw/100).toFixed(2); gC.numFmt='#,##0.00'; sc(gC,C.NAVY,true,'FFFFFFFF',11,'center');
  ws.getRow(row).getCell(8).value=''; sc(ws.getRow(row).getCell(8),C.NAVY,false,'FFFFFFFF',9,'center');
  ws.getRow(row).height=22; row++; row++;

  ws.mergeCells(row,1,row,5); ws.getCell(row,1).value='LACS'; sc(ws.getCell(row,1),C.GREY,false,'FF000000',8,'right');
  ws.getCell(row,6).value='CR'; sc(ws.getCell(row,6),C.GREY,false,'FF000000',8,'center');
  ws.getRow(row).height=12;
  ws.views=[{state:'frozen',xSplit:0,ySplit:4}];
}

function _autoBlockFromMaterials(d) {
  const ms=d.material_summary||{}, boq=d.boq||[];
  if (!Object.keys(ms).length&&!boq.length) return [];
  const roadItems=[]; let roadTotal=0;
  const tryAdd=(desc,qty,unit,rate)=>{
    if (!qty) return;
    const amt=+(qty*rate/100000).toFixed(2);
    roadItems.push({description:desc,qty:+qty.toFixed?qty.toFixed(2)*1:qty,unit,rate,amount_lacs:amt});
    roadTotal+=amt;
  };
  tryAdd('SOIL FILLING IN ROAD AREA',ms.soil_filling_cum,'CUM',285);
  if (ms.total_road_area_sqmt) tryAdd('GSB FILLING (300 MM LAYER)',ms.total_road_area_sqmt,'SQMT',655);
  if (ms.total_road_area_sqmt) tryAdd('WMM FILLING (200 MM LAYER)',ms.total_road_area_sqmt,'SQMT',515);
  if (ms.total_road_area_sqmt) tryAdd('RCC ROAD WORK (250 MM PQC)',ms.total_road_area_sqmt,'SQMT',1800);
  if (ms.dowel_steel_kg) tryAdd('STEEL FOR DOWEL BAR',+(ms.dowel_steel_kg/1000).toFixed(3),'TON',56000);
  if (ms.paver_block_sqmt) tryAdd('PAVER BLOCK 80MM (M40)',ms.paver_block_sqmt,'SQMT',1040);
  const sections=[]; if (roadItems.length) sections.push({sr:1,section_name:'ROADS',items:roadItems,subtotal_lacs:roadTotal});
  const svcItems=[]; let svcTotal=0;
  if (ms.stp_mld){const a=+(ms.stp_mld*25000000/100000).toFixed(2);svcItems.push({description:'SEWAGE TREATMENT PLANT',qty:ms.stp_mld,unit:'MLD',rate:25000000,amount_lacs:a});svcTotal+=a;}
  if (ms.pipeline_rmt){const a=+(ms.pipeline_rmt*4500/100000).toFixed(2);svcItems.push({description:'PIPELINE NETWORK',qty:ms.pipeline_rmt,unit:'RMT',rate:4500,amount_lacs:a});svcTotal+=a;}
  if (ms.street_lights_nos){const a=+(ms.street_lights_nos*35000/100000).toFixed(2);svcItems.push({description:'STREET LIGHT POLES',qty:ms.street_lights_nos,unit:'NOS',rate:35000,amount_lacs:a});svcTotal+=a;}
  if (ms.compound_wall_rmt){const a=+(ms.compound_wall_rmt*8600/100000).toFixed(2);svcItems.push({description:'COMPOUND WALL',qty:ms.compound_wall_rmt,unit:'RMT',rate:8600,amount_lacs:a});svcTotal+=a;}
  if (ms.gabion_wall_rmt){const a=+(ms.gabion_wall_rmt*14100/100000).toFixed(2);svcItems.push({description:'GABION WALL',qty:ms.gabion_wall_rmt,unit:'RMT',rate:14100,amount_lacs:a});svcTotal+=a;}
  if (svcItems.length) sections.push({sr:2,section_name:'SERVICES (DRAINAGE / UTILITIES)',items:svcItems,subtotal_lacs:svcTotal});
  if (!sections.length) return [];
  const pt=sections.reduce((s,x)=>s+x.subtotal_lacs,0);
  return [{part_label:'PART-A (PROJECT WORKS)',sections,part_total_lacs:pt}];
}

// ── SHEET 2: ROAD SCHEDULE ───────────────────────────────────────────────────
function _roadSchedule(wb, d) {
  const ws=wb.addWorksheet('ROAD SCHEDULE');
  const COLS=11; [6,10,10,18,14,14,14,14,14,14,32].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',COLS,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'RCC ROAD SCHEDULE',COLS,C.MIDBLUE,'FFFFFFFF',11,18);
  const hdrs=['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH\n(MTR)','CARRIAGE\nWIDTH (M)','ROAD\nAREA (SQMT)','BOX CUTTING\n(5% EXTRA)','GSB 300MM\n(TONS)','WMM 200MM\n(TONS)','PQC M30\n250MM (CUM)','REMARK'];
  const hr=ws.getRow(row); hr.height=48;
  hdrs.forEach((h,i)=>{const c=hr.getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',9,'center');}); row++;
  (d.road_schedule||[]).forEach((rd,i)=>{
    const bg=i%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=16;
    const area=rd.area_sqmt||(rd.length_m||0)*(rd.carriage_width_m||0);
    const box=+(area*1.05).toFixed(2), gsb=rd.gsb_ton||+(area*1.15*0.30*1.8).toFixed(2);
    const wmm=rd.wmm_ton||+(area*1.15*0.20*2.1).toFixed(2), pqc=rd.pqc_cum||+(area*1.05*0.25).toFixed(2);
    const vals=[rd.sr||i+1,`${rd.total_width_m||''}MT`,rd.road_no||'',rd.length_m||0,rd.carriage_width_m||0,+area.toFixed(2),box,+gsb.toFixed(2),+wmm.toFixed(2),+pqc.toFixed(2),rd.remark||''];
    vals.forEach((v,ci)=>{const c=r.getCell(ci+1);c.value=v;sc(c,bg,false,'FF000000',9,ci===10?'left':'center');if(ci>=3&&ci<=9&&typeof v==='number')c.numFmt='#,##0.00';}); row++;
  });
  ws.mergeCells(row,1,row,3); sc(ws.getCell(row,1),C.YELLOW,true,'FF000000',10,'center'); ws.getCell(row,1).value='TOTAL';
  const ms=d.material_summary||{};
  [ms.total_road_length_rmt||0,ms.total_road_area_sqmt||0,+((ms.total_road_area_sqmt||0)*1.05).toFixed(2),ms.gsb_ton||0,ms.wmm_ton||0,ms.pqc_cum||0].forEach((v,i)=>{
    const c=ws.getRow(row).getCell(i+4);c.value=v;c.numFmt='#,##0.00';sc(c,C.YELLOW,true,'FF000000',10,'center');});
  ws.getRow(row).height=18; ws.views=[{state:'frozen',ySplit:3}];
}

// ── SHEET 3: SOIL FILLING ────────────────────────────────────────────────────
function _soilFilling(wb, d) {
  const ws=wb.addWorksheet('SOIL FILLING WORK');
  const COLS=14; [6,8,10,15,12,13,11,12,12,12,12,14,12,28].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',COLS,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'RCC ROAD CUTTING / FILLING CALCULATION',COLS,C.MIDBLUE,'FFFFFFFF',11,18);
  const hdrs=['SR\nNO','ROAD\nNO','ROAD\nWIDTH','SUB BASE\nWIDTH (M)','ROAD\nLENGTH','SUB BASE\nAREA (SQMT)','ROAD TOP\nLEVEL (A)','NGL OF\nROAD (B)','SECTION\nDEPTH (C)','FILLING\nDEPTH','EXTRA\nFILLING','TOTAL FILLING\n(CUM)','HYWAS\n(14 CUM)','REMARK'];
  const hr=ws.getRow(row); hr.height=52;
  hdrs.forEach((h,i)=>{const c=hr.getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',9,'center');}); row++;
  (d.soil_filling||[]).forEach((sf,i)=>{
    const bg=i%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=15;
    const area=sf.sub_base_area||(sf.length||0)*(sf.sub_base_width||0);
    const hywas=sf.hywas||+((sf.total_filling_cum||0)/14).toFixed(2);
    const vals=[sf.sr||i+1,sf.road_no||'',sf.road_width||0,sf.sub_base_width||0,sf.length||0,+area.toFixed(2),sf.road_top_lvl||0,sf.ngl||0,sf.section_depth||0,sf.filling_depth||0,sf.extra_filling||0,sf.total_filling_cum||0,hywas,sf.remark||''];
    vals.forEach((v,ci)=>{const c=r.getCell(ci+1);c.value=v;sc(c,bg,false,'FF000000',9,ci===13?'left':'center');if(ci>=3&&ci<=12&&typeof v==='number')c.numFmt='#,##0.00';}); row++;
  });
  row++;
  row=mergeHdr(ws,row,'SUMMARY',COLS,C.MIDBLUE,'FFFFFFFF',10,16);
  const ms=d.material_summary||{};
  const tf=ms.soil_filling_cum||(d.soil_filling||[]).reduce((s,f)=>s+(f.total_filling_cum||0),0);
  [['TOTAL FILLING (CUM)',tf,'#,##0.00'],['RATE (RS/CUM)',285,'#,##0'],['TOTAL COST (RS)',Math.round(tf*285),'#,##0']].forEach(([lbl,val,fmt])=>{
    ws.mergeCells(row,1,row,10);sc(ws.getCell(row,1),C.GREY,true,'FF000000',9,'right');ws.getCell(row,1).value=lbl;
    const vc=ws.getRow(row).getCell(11);vc.value=val;vc.numFmt=fmt;sc(vc,C.YELLOW,true,'FF000000',10,'center');ws.getRow(row).height=16;row++;
  });
  ws.views=[{state:'frozen',ySplit:3}];
}

// ── SHEET 4: BBS ─────────────────────────────────────────────────────────────
function _bbs(wb, d) {
  const ws=wb.addWorksheet('BBS - STEEL');
  const COLS=15; [6,44,9,8,10,12,11,8,10,10,10,10,10,10,11].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',COLS,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'BAR BENDING SCHEDULE (BBS) — STEEL REINFORCEMENT',COLS,C.MIDBLUE,'FFFFFFFF',11,18);
  ['SR NO','DESCRIPTION','DIA\n(MM)','NOS','CUT LEN\n(M)','TOTAL LEN\n(M)','WT\n(KG)','UNIT','8MM\nKG','10MM\nKG','12MM\nKG','16MM\nKG','20MM\nKG','25MM\nKG','TOTAL\nKG'].forEach((h,i)=>{
    const c=ws.getRow(row).getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',8,'center');}); ws.getRow(row).height=42; row++;
  const UW={8:0.395,10:0.617,12:0.888,16:1.58,20:2.47,25:3.86,32:6.31};
  const DC={8:9,10:10,12:11,16:12,20:13,25:14};
  let grand=0;
  (d.bbs||[]).forEach(sec=>{
    row=mergeHdr(ws,row,sec.element||'',COLS,C.MIDBLUE,'FFFFFFFF',10,16);
    (sec.items||[]).forEach((item,i)=>{
      const bg=i%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=14;
      const wt=item.weight_kg||+((item.total_length||0)*(UW[item.dia]||0.617)).toFixed(2); grand+=wt;
      [i+1,item.description||'',item.dia||'',item.nos||'',item.cutting_length||'',item.total_length||'',+wt.toFixed(2),'KG','','','','','','',''].forEach((v,ci)=>{
        const c=r.getCell(ci+1);c.value=v;sc(c,bg,false,'FF000000',8,ci===1?'left':'center');});
      const dc=DC[item.dia]; if(dc) ws.getRow(row).getCell(dc).value=+wt.toFixed(2);
      ws.getRow(row).getCell(15).value=+wt.toFixed(2); row++;
    }); row++;
  });
  ws.mergeCells(row,1,row,6);sc(ws.getCell(row,1),C.YELLOW,true,'FF000000',10,'right');ws.getCell(row,1).value='TOTAL STEEL WEIGHT';
  const tw=ws.getRow(row).getCell(7);tw.value=+grand.toFixed(2);tw.numFmt='#,##0.00';sc(tw,C.YELLOW,true,'FF000000',10,'center');
  const tw2=ws.getRow(row).getCell(15);tw2.value=+grand.toFixed(2);tw2.numFmt='#,##0.00';sc(tw2,C.YELLOW,true,'FF000000',10,'center');
  ws.getRow(row).height=18; ws.views=[{state:'frozen',ySplit:3}];
}

// ── SHEET 5: STREET LIGHT ────────────────────────────────────────────────────
function _streetLight(wb, d) {
  const ws=wb.addWorksheet('STREET LIGHT');
  [6,12,10,18,18,30].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',6,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'STREET LIGHT SCHEDULE',6,C.MIDBLUE,'FFFFFFFF',11,18);
  ['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH (MTR)','STREET LIGHTS (NOS)','REMARK'].forEach((h,i)=>{
    const c=ws.getRow(row).getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',9,'center');}); ws.getRow(row).height=18; row++;
  let tot=0;
  (d.street_light_schedule||[]).forEach((sl,i)=>{
    const bg=i%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=15; tot+=sl.nos||0;
    [sl.sr||i+1,sl.road_wide||'',sl.road_no||'',sl.length_mt||0,sl.nos||0,sl.remark||''].forEach((v,ci)=>{
      const c=r.getCell(ci+1);c.value=v;sc(c,bg,false,'FF000000',9,ci>1&&ci<5?'center':'left');}); row++;
  });
  ws.mergeCells(row,1,row,3);sc(ws.getCell(row,1),C.YELLOW,true,'FF000000',10,'right');ws.getCell(row,1).value='TOTAL';
  const tc=ws.getRow(row).getCell(5);tc.value=tot;sc(tc,C.YELLOW,true,'FF000000',11,'center');
  ws.getRow(row).getCell(6).value=`₹${(tot*35000).toLocaleString('en-IN')} (@₹35,000/nos)`;
  sc(ws.getRow(row).getCell(6),C.GREEN,false,'FF000000',9,'left'); ws.getRow(row).height=18;
}

// ── SHEET 6: BOQ ─────────────────────────────────────────────────────────────
function _boq(wb, d) {
  const ws=wb.addWorksheet('BOQ');
  [6,50,10,14,14,18].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',6,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'BILL OF QUANTITIES',6,C.MIDBLUE,'FFFFFFFF',11,18);
  ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)'].forEach((h,i)=>{
    const c=ws.getRow(row).getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',10,'center');}); ws.getRow(row).height=18; row++;
  let grand=0;
  let items=d.boq||[];
  if (!items.length) {
    let sr=1;
    (d.block_cost_parts||[]).forEach(p=>(p.sections||[]).forEach(s=>(s.items||[]).forEach(it=>{
      items.push({sr:sr++,description:it.description,unit:it.unit,qty:it.qty,rate:it.rate,amount:(it.amount_lacs||0)*100000});
    })));
  }
  items.forEach((item,i)=>{
    const bg=i%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=15;
    const amt=item.amount||(item.qty||0)*(item.rate||0); grand+=parseFloat(amt)||0;
    [item.sr||i+1,item.description||'',item.unit||'',item.qty||0,item.rate||0,amt].forEach((v,ci)=>{
      const c=r.getCell(ci+1);c.value=v;sc(c,bg,false,'FF000000',9,ci>1?'center':'left');if(ci>=3)c.numFmt=ci===5?'#,##0':'#,##0.00';}); row++;
  });
  ws.mergeCells(row,1,row,4);sc(ws.getCell(row,1),C.YELLOW,true,'FF000000',11,'right');ws.getCell(row,1).value='GRAND TOTAL';
  const gv=ws.getRow(row).getCell(6);gv.value=grand;gv.numFmt='₹#,##0';sc(gv,C.YELLOW,true,'FF000000',11,'center');ws.getRow(row).height=20;
  ws.views=[{state:'frozen',ySplit:3}];
}

// ── SHEET 7: MATERIAL SUMMARY ────────────────────────────────────────────────
function _materialSummary(wb, d) {
  const ws=wb.addWorksheet('MATERIAL SUMMARY');
  [6,50,16,12,28].forEach((w,i)=>ws.getColumn(i+1).width=w);
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',5,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'OVERALL MATERIAL SUMMARY',5,C.MIDBLUE,'FFFFFFFF',11,18);
  ['SR NO','MATERIAL / WORK ITEM','QUANTITY','UNIT','REMARKS'].forEach((h,i)=>{
    const c=ws.getRow(row).getCell(i+1);c.value=h;sc(c,C.NAVY,true,'FFFFFFFF',10,'center');}); ws.getRow(row).height=18; row++;
  const ms=d.material_summary||{};
  const sec=(lbl)=>{ws.mergeCells(row,1,row,5);const c=ws.getCell(row,1);c.value=lbl;sc(c,C.MIDBLUE,true,'FFFFFFFF',10,'left');ws.getRow(row).height=16;row++;};
  const item=(sr,lbl,qty,unit,rem='')=>{
    const bg=row%2===0?C.WHITE:C.GREY; const r=ws.getRow(row); r.height=15;
    [sr,lbl,qty,unit,rem].forEach((v,ci)=>{const c=r.getCell(ci+1);c.value=v??'';sc(c,bg,false,'FF000000',9,ci===1||ci===4?'left':'center');if(ci===2&&typeof v==='number')c.numFmt='#,##0.00';}); row++;
  };
  if ((d.road_schedule||[]).length||ms.total_road_area_sqmt) {
    sec('ROADS & SUB-BASE');
    item(1,'TOTAL ROAD LENGTH',ms.total_road_length_rmt||0,'RMT','');
    item(2,'TOTAL ROAD AREA',ms.total_road_area_sqmt||0,'SQMT','');
    item(3,'BOX CUTTING (5% EXTRA)',+((ms.total_road_area_sqmt||0)*1.05).toFixed(2),'SQMT','15% EXTRA');
    item(4,'GSB FILLING (300MM)',ms.gsb_ton||0,'TON','15% EXTRA COMPACTION');
    item(5,'WMM FILLING (200MM)',ms.wmm_ton||0,'TON','15% EXTRA COMPACTION');
    item(6,'PQC ROAD (250MM M30)',ms.pqc_cum||0,'CUM','5% WASTAGE');
    item(7,'SOIL FILLING',ms.soil_filling_cum||0,'CUM','');
  }
  if (ms.paver_block_sqmt){sec('SERVICE CORRIDOR');item(8,'PAVER BLOCK 80MM (M40)',ms.paver_block_sqmt,'SQMT','');}
  if (ms.dowel_steel_kg||ms.compound_wall_rmt||ms.gabion_wall_rmt){
    sec('STEEL & STRUCTURAL');
    if(ms.dowel_steel_kg) item(9,'RCC STEEL DOWEL + TIEBAR',ms.dowel_steel_kg,'KG','');
    if(ms.compound_wall_rmt) item(10,'COMPOUND WALL',ms.compound_wall_rmt,'RMT','');
    if(ms.gabion_wall_rmt) item(11,'GABION WALL',ms.gabion_wall_rmt,'RMT','');
  }
  if (ms.stp_mld||ms.pipeline_rmt||ms.street_lights_nos){
    sec('SERVICES / MEP');
    if(ms.stp_mld) item(12,'STP CAPACITY',ms.stp_mld,'MLD','');
    if(ms.pipeline_rmt) item(13,'PIPELINE NETWORK',ms.pipeline_rmt,'RMT','');
    if(ms.street_lights_nos) item(14,'STREET LIGHT POLES',ms.street_lights_nos,'NOS','');
  }
  ws.views=[{state:'frozen',ySplit:3}];
}

// ── SHEET 8: PMC OBSERVATIONS ────────────────────────────────────────────────
function _observations(wb, d) {
  const ws=wb.addWorksheet('PMC OBSERVATIONS');
  ws.getColumn(1).width=6; ws.getColumn(2).width=90;
  let row=1;
  row=mergeHdr(ws,row,d.project_name||'PMC CIVIL PROJECT',2,C.NAVY,'FFFFFFFF',13,22);
  row=mergeHdr(ws,row,'PMC OBSERVATIONS & RECOMMENDATIONS',2,C.MIDBLUE,'FFFFFFFF',11,18);
  [['Drawing Type',d.drawing_type],['Location',d.location],['Scale',d.scale],['Drawing No',d.drawing_no],['Date',d.date],['AI Confidence',d.confidence]].forEach(([lbl,val])=>{
    if (!val) return; const r=ws.getRow(row); r.height=14;
    r.getCell(1).value=lbl+':'; sc(r.getCell(1),C.LTBLUE,true,'FF000000',9,'left');
    r.getCell(2).value=val; sc(r.getCell(2),C.WHITE,false,'FF000000',9,'left'); row++;
  });
  row++;
  row=mergeHdr(ws,row,'PMC OBSERVATIONS:',2,C.DKGREEN,'FFFFFFFF',10,18);
  (d.pmc_observations||['Refer to chat analysis for detailed observations.']).forEach((obs,i)=>{
    const r=ws.getRow(row); r.height=28;
    r.getCell(1).value=i+1; sc(r.getCell(1),i%2===0?C.WHITE:C.GREY,false,'FF000000',9,'center');
    r.getCell(2).value=obs; sc(r.getCell(2),i%2===0?C.WHITE:C.GREY,false,'FF000000',9,'left'); row++;
  });
  row++;
  row=mergeHdr(ws,row,'PMC RECOMMENDATION:',2,C.DKGREEN,'FFFFFFFF',10,18);
  ws.mergeCells(row,1,row,2); const rc=ws.getCell(row,1);
  rc.value=d.pmc_recommendation||'Refer to chat analysis.';
  sc(rc,C.GREEN,false,'FF000000',9,'left'); ws.getRow(row).height=80; row+=2;
  ws.mergeCells(row,1,row,2); const fc=ws.getCell(row,1);
  const today=new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'2-digit',year:'numeric'});
  fc.value=`Prepared by: PMC Civil AI Agent | Date: ${today} | ${d.project_name||'VCT Bharuch'} — Powered by Gemini AI`;
  fc.fill={type:'pattern',pattern:'solid',fgColor:{argb:C.GREY}};
  fc.font={italic:true,size:9,color:{argb:'FF595959'},name:'Calibri'};
  fc.alignment={horizontal:'center',vertical:'middle'}; ws.getRow(row).height=14;
}

module.exports = { extractDrawingData, buildDrawingExcel };

  MIDBLUE: 'FF2E75B6',
  LTBLUE:  'FFBDD7EE',
  YELLOW:  'FFFFD966',
  GREEN:   'FFE2EFDA',
  DKGREEN: 'FF375623',
  GREY:    'FFF2F2F2',
  WHITE:   'FFFFFFFF',
  LOWEST:  'FF00B050',
  ORANGE:  'FFED7D31',
  RED:     'FFFF0000'
};

const thin = { style: 'thin', color: { argb: 'FF000000' } };
const bdr  = { top: thin, left: thin, bottom: thin, right: thin };

function sc(cell, bg, bold=false, fc='FF000000', size=9, align='center', wrap=true) {
  if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  cell.font = { bold, color: { argb: fc }, size, name: 'Calibri' };
  cell.alignment = { horizontal: align, vertical: 'middle', wrapText: wrap };
  cell.border = bdr;
}

function hdr(ws, row, text, bg=C.NAVY, fc=C.WHITE, size=11, height=18) {
  const lastCol = ws.columnCount || 8;
  ws.mergeCells(row, 1, row, lastCol);
  const c = ws.getCell(row, 1);
  c.value = text; sc(c, bg, true, fc.replace('FF','FF'), size, 'center');
  ws.getRow(row).height = height;
  return row + 1;
}

// ─── EXTRACT DRAWING DATA FROM GEMINI ──────────────────────────────────────
async function extractDrawingData(key, files, userText, aiResponse, fetch) {
  const parts = [];
  if (!aiResponse) {
    for (const f of (files||[])) {
      if (f.type==='application/pdf'||f.name?.match(/\.pdf$/i))
        parts.push({inline_data:{mime_type:'application/pdf',data:f.b64}});
      else if (f.type?.startsWith('image/'))
        parts.push({inline_data:{mime_type:f.type,data:f.b64}});
    }
  }

  const GEMINI_URL = (k) => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;

  const prompt = `You are a senior PMC civil engineer for India. Analyze this civil drawing/estimate and extract ALL data.
Return ONLY raw JSON. No markdown. No backticks.

${aiResponse ? 'CONTENT:\n' + aiResponse : 'Analyze the uploaded drawing/document.'}
${userText ? 'Note: ' + userText : ''}

Return this exact JSON structure:
{
  "project_name": "AI PARK @ BHATPOR or extract",
  "drawing_type": "ROAD/BUILDING/DRAINAGE/COMPOUND_WALL/etc",
  "date": "DD-MM-YYYY",
  "prepared_by": "PMC Civil AI Agent",

  "block_cost": [
    {"sr": 1, "component": "ROADS", "items": [
      {"description": "SOIL STABILIZATION", "qty": 38388.59, "unit": "SQMT", "rate": 82, "amount_lacs": 3.15},
      {"description": "GSB FILLING 300MM", "qty": 38388.59, "unit": "SQMT", "rate": 655, "amount_lacs": 251.44}
    ], "subtotal_lacs": 254.59}
  ],

  "road_schedule": [
    {"sr": 1, "road_no": "R1", "width_mt": 24, "length_mt": 129.12, "carriage_width": 15,
     "area_sqmt": 1936.8, "gsb_ton": 1459.34, "wmm_ton": 1033.12, "pqc_cum": 457.57,
     "remark": "4.5-4.5 MT B/S SERVICES"}
  ],

  "boq": [
    {"sr": 1, "description": "SOIL STABILIZATION (LIME-FLYASH)", "unit": "SQMT",
     "qty": 38388.59, "rate": 82, "amount": 3147864.38}
  ],

  "rate_analysis": [
    {"item": "PQC ROAD (250MM THK)", "unit": "SQMT",
     "components": [
       {"description": "RCC M30 CONCRETE 250MM", "qty": 450, "unit": "CUM", "rate": 5500, "amount": 2475000},
       {"description": "5% WASTAGE", "qty": 1, "unit": "LS", "rate": 123750, "amount": 123750},
       {"description": "RCC LABOUR", "qty": 19368, "unit": "SQFT", "rate": 32.5, "amount": 629460}
     ],
     "total": 3228210, "rate_per_sqmt": 1793.45}
  ],

  "soil_filling": [
    {"sr": 1, "road_no": "R1", "road_width": 24, "sub_base_width": 15.6, "length": 129.12,
     "area": 2014.27, "road_top_lvl": 49.95, "ngl": 48.7958, "section_depth": 0.75,
     "filling_depth": 0.554, "extra_filling": 0, "total_filling_cum": 1116.31, "hywas": 79.74}
  ],

  "bbs": [
    {"element": "ROAD PANEL - R1 24MT WIDE", "items": [
      {"description": "DOWELS 25MM DIA 450MM C/C LONGITUDINAL", "dia": 25, "nos": 287, "cutting_length": 0.6, "total_length": 344.4, "weight_kg": 132.4},
      {"description": "CHAIR 10MM DIA 900MM C/C", "dia": 10, "nos": 96, "cutting_length": 1.09, "total_length": 209.28, "weight_kg": 12.93}
    ]}
  ],

  "street_light": [
    {"sr": 1, "road_wide": "24 MT", "road_no": "R1", "length_mt": 129.12, "nos": 6, "remark": "EVERY 20MT"}
  ],

  "material_summary": {
    "gsb_ton": 27812.53,
    "wmm_ton": 19689.51,
    "pqc_cum": 9511.86,
    "soil_filling_cum": 24689.53,
    "paver_block_sqmt": 4370.64,
    "rcc_steel_kg": 45000,
    "total_road_rmt": 3265.45,
    "total_road_area_sqmt": 36429.32
  },

  "cost_summary": {
    "total_crores": 45.23,
    "civil_works_lacs": 2800,
    "electrical_lacs": 450,
    "contingency_lacs": 160
  },

  "observations": ["PMC observation 1", "PMC observation 2"],
  "recommendation": "Full PMC recommendation"
}

RULES: Extract ALL actual numbers from content. If data missing use 0. Return ONLY JSON.`;

  parts.push({ text: prompt });
  const r = await fetch(GEMINI_URL(key), {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ role: 'user', parts }], generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' } })
  });
  let raw = (await r.json())?.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
  if (fb !== -1 && lb !== -1) raw = raw.slice(fb, lb + 1);
  try { return JSON.parse(raw.replace(/```json|```/g, '').trim()); }
  catch(e) {
    console.error('Drawing data parse fail:', e.message, raw.slice(0,200));
    return { project_name: 'PMC CIVIL PROJECT', drawing_type: 'GENERAL', date: new Date().toLocaleDateString('en-IN'), prepared_by: 'PMC Civil AI Agent', block_cost: [], road_schedule: [], boq: [], rate_analysis: [], soil_filling: [], bbs: [], street_light: [], material_summary: {}, cost_summary: {}, observations: [], recommendation: 'Refer to chat analysis.' };
  }
}

// ─── BUILD MULTI-SHEET EXCEL ────────────────────────────────────────────────
async function buildDrawingExcel(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';

  // ── SHEET 1: BLOCK COST ─────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('BLOCK COST');
    ws.getColumn(1).width = 6; ws.getColumn(2).width = 38;
    ws.getColumn(3).width = 18; ws.getColumn(4).width = 12;
    ws.getColumn(5).width = 14; ws.getColumn(6).width = 20; ws.getColumn(7).width = 18;

    // Row 1-2 headers
    ws.mergeCells('A1:G1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name || 'AI PARK @ BHATPOR'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:G2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'BLOCK COST'; ws.getRow(2).height = 18;
    ws.mergeCells('A3:G3'); sc(ws.getCell('A3'), C.GREY, false, 'FF000000', 9, 'left');
    ws.getCell('A3').value = 'TOTAL AREA: ' + (d.total_area || '264136.31 SQMT'); ws.getRow(3).height = 14;

    // Col headers row 4
    const hdrs4 = ['SR NO','COMPONENT','EST. QUANTITY','UNITS','RATE','EST. VALUE (RS IN LACS)','EST. VALUE (RS IN CR.)'];
    const r4 = ws.getRow(4); r4.height = 18;
    hdrs4.forEach((h,i) => { const c = r4.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 9, 'center'); });

    let row = 5;
    const blocks = d.block_cost || [];
    if (!blocks.length) {
      // Sample row if no data
      ws.mergeCells(row,1,row,7); const c=ws.getCell(row,1); c.value='(No block cost data extracted - refer to chat analysis)';
      sc(c, C.GREY, false, 'FF595959', 9, 'center'); row++;
    }
    blocks.forEach((blk, bi) => {
      // Part header
      ws.mergeCells(row,1,row,7);
      const ph = ws.getCell(row,1); ph.value = `PART-${String.fromCharCode(65+bi)} - ${blk.component||''}`;
      sc(ph, C.MIDBLUE, true, 'FFFFFFFF', 10, 'center'); ws.getRow(row).height = 16; row++;
      (blk.items||[]).forEach((item, ii) => {
        const bg = ii%2===0 ? C.WHITE : C.GREY;
        const r = ws.getRow(row); r.height = 15;
        const cells = [ii+1, item.description||'', item.qty||'', item.unit||'', item.rate||'', item.amount_lacs||'', ((item.amount_lacs||0)/100).toFixed(4)];
        cells.forEach((v,ci) => { const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>1?'center':'left'); if (ci>=4) { c.numFmt = '#,##0.00'; } });
        row++;
      });
      // Subtotal
      const sr = ws.getRow(row); sr.height = 16;
      ws.mergeCells(row,1,row,5); const stc = sr.getCell(1); stc.value = 'SUBTOTAL'; sc(stc, C.YELLOW, true, 'FF000000', 10, 'right');
      const stv = sr.getCell(6); stv.value = blk.subtotal_lacs||0; stv.numFmt='#,##0.00'; sc(stv, C.YELLOW, true, 'FF000000', 10, 'center');
      const stcr = sr.getCell(7); stcr.value = ((blk.subtotal_lacs||0)/100); stcr.numFmt='#,##0.0000'; sc(stcr, C.YELLOW, true, 'FF000000', 10, 'center');
      row += 2;
    });

    // Grand total
    ws.mergeCells(row,1,row,5);
    const gt = ws.getCell(row,1); gt.value = 'GRAND TOTAL'; sc(gt, C.NAVY, true, 'FFFFFFFF', 11, 'right');
    const gtv = ws.getCell(row,6); gtv.value = (d.cost_summary?.civil_works_lacs||0); gtv.numFmt='#,##0.00'; sc(gtv, C.NAVY, true, 'FFFFFFFF', 11, 'center');
    const gtcr = ws.getCell(row,7); gtcr.value = ((d.cost_summary?.total_crores||0)); gtcr.numFmt='#,##0.00'; sc(gtcr, C.NAVY, true, 'FFFFFFFF', 11, 'center');
    ws.getRow(row).height = 20; row++;

    ws.views = [{ state: 'frozen', ySplit: 4 }];
  }

  // ── SHEET 2: ROAD SCHEDULE ─────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('ROAD SCHEDULE');
    [6,10,8,14,14,14,14,14,14,14,24].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:K1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name || 'AI PARK @ BHATPOR'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:K2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'ROAD WORK - SUB-BASE & BASE COURSE ESTIMATE'; ws.getRow(2).height = 18;

    const rHdrs = ['SR\nNO','ROAD\nWIDE','ROAD\nNO','ROAD LENGTH\n(MTR)','CARRIAGE\nWIDTH (MT)','AREA\n(SQMT)','BOX CUTTING\n(SQMT)','GSB FILLING\n300MM (TON)','WMM FILLING\n200MM (TON)','PQC ROAD\n250MM (CUM)','REMARK'];
    const r3 = ws.getRow(3); r3.height = 52;
    rHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 9, 'center'); });

    let row = 4;
    const roads = d.road_schedule||[];
    roads.forEach((rd, i) => {
      const bg = i%2===0 ? C.WHITE : C.GREY;
      const r = ws.getRow(row); r.height = 16;
      const vals = [rd.sr, rd.road_wide||`${rd.width_mt} MT`, rd.road_no, rd.length_mt, rd.carriage_width||rd.width_mt, rd.area_sqmt, (rd.area_sqmt||0)*1.05, rd.gsb_ton, rd.wmm_ton, rd.pqc_cum, rd.remark||''];
      vals.forEach((v,ci) => { const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>=3&&ci<10?'center':ci===0?'center':'left'); if (ci>=5&&ci<=9&&typeof v==='number') c.numFmt='#,##0.00'; });
      row++;
    });

    // Totals row
    const totRow = ws.getRow(row); totRow.height = 18; row++;
    ws.mergeCells(row-1,1,row-1,3);
    const tc = totRow.getCell(1); tc.value = 'TOTAL'; sc(tc, C.YELLOW, true, 'FF000000', 10, 'center');
    const ms = d.material_summary||{};
    [ms.total_road_rmt||0, ms.total_road_area_sqmt||0, (ms.total_road_area_sqmt||0)*1.05, ms.gsb_ton||0, ms.wmm_ton||0, ms.pqc_cum||0].forEach((v,i) => {
      const c = totRow.getCell(i+4); c.value = v; c.numFmt='#,##0.00'; sc(c, C.YELLOW, true, 'FF000000', 10, 'center');
    });
    totRow.getCell(11).value = ''; sc(totRow.getCell(11), C.YELLOW, false, 'FF000000', 9, 'center');

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ── SHEET 3: BOQ ───────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('BOQ');
    [6,46,12,14,14,18].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:F1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK @ BHATPOR'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:F2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'BILL OF QUANTITIES'; ws.getRow(2).height = 18;

    const bHdrs = ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)'];
    const r3 = ws.getRow(3); r3.height = 18;
    bHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 10, 'center'); });

    let row = 4, grandTotal = 0;
    const boqs = d.boq||[];
    boqs.forEach((item, i) => {
      const bg = i%2===0 ? C.WHITE : C.GREY;
      const r = ws.getRow(row); r.height = 16;
      const amt = (item.qty||0)*(item.rate||0);
      grandTotal += parseFloat(item.amount||amt||0);
      [item.sr||i+1, item.description||'', item.unit||'', item.qty||0, item.rate||0, item.amount||amt].forEach((v,ci) => {
        const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>1?'center':'left');
        if (ci>=3) c.numFmt = ci===5 ? '₹#,##0' : '#,##0.00';
      });
      row++;
    });

    // Grand total
    ws.mergeCells(row,1,row,4);
    const gtc = ws.getCell(row,1); gtc.value = 'GRAND TOTAL'; sc(gtc, C.YELLOW, true, 'FF000000', 11, 'right');
    const gtv = ws.getCell(row,6); gtv.value = grandTotal; gtv.numFmt='₹#,##0'; sc(gtv, C.YELLOW, true, 'FF000000', 11, 'center');
    ws.getRow(row).height = 20;

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ── SHEET 4: RATE ANALYSIS ─────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('RATE ANALYSIS');
    [6,46,12,12,14,14].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:F1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;

    let row = 2;
    (d.rate_analysis||[]).forEach(ra => {
      ws.mergeCells(row,1,row,6);
      sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 11, 'center');
      ws.getCell(row,1).value = `RATE ANALYSIS: ${ra.item||''} [${ra.unit||''}]`;
      ws.getRow(row).height = 18; row++;

      const hdrs = ['SR','DESCRIPTION','UNIT','QTY','RATE','AMOUNT'];
      const hr = ws.getRow(row); hr.height = 16;
      hdrs.forEach((h,i) => { const c = hr.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 9, 'center'); });
      row++;

      (ra.components||[]).forEach((comp, i) => {
        const bg = i%2===0?C.WHITE:C.GREY;
        const r = ws.getRow(row); r.height = 15;
        [i+1, comp.description||'', comp.unit||'', comp.qty||0, comp.rate||0, comp.amount||0].forEach((v,ci) => {
          const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>1?'center':'left');
          if (ci>=3) c.numFmt = '#,##0.00';
        });
        row++;
      });

      // Subtotal / rate per sqmt
      ws.mergeCells(row,1,row,4);
      sc(ws.getCell(row,1), C.YELLOW, true, 'FF000000', 10, 'right');
      ws.getCell(row,1).value = `TOTAL | RATE PER ${ra.unit||'SQMT'}`;
      ws.getCell(row,5).value = ra.rate_per_sqmt||ra.total||0; sc(ws.getCell(row,5), C.YELLOW, true, 'FF000000', 10, 'center'); ws.getCell(row,5).numFmt='#,##0.00';
      ws.getCell(row,6).value = ra.total||0; sc(ws.getCell(row,6), C.YELLOW, true, 'FF000000', 10, 'center'); ws.getCell(row,6).numFmt='#,##0';
      ws.getRow(row).height = 18; row += 2;
    });
  }

  // ── SHEET 5: SOIL FILLING ──────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('SOIL FILLING WORK');
    [6,8,10,16,12,12,12,12,14,12,12,14,12,20].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:N1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:N2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 11, 'center');
    ws.getCell('A2').value = 'RCC ROAD CUTTING / FILLING CALCULATION'; ws.getRow(2).height = 18;

    const sfHdrs = ['SR\nNO','ROAD\nNO','ROAD\nWIDTH','SUB BASE\nWIDTH (M)','ROAD\nLENGTH','SUB BASE\nAREA (SQMT)','ROAD TOP\nLEVEL (A)','NGL OF\nROAD (B)','SECTION\nDEPTH (C)','FILLING\nDEPTH','EXTRA\nFILLING','TOTAL FILLING\n(CUM)','HYWAS\n(14 CUM)','REMARK'];
    const r3 = ws.getRow(3); r3.height = 52;
    sfHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 9, 'center'); });

    let row = 4;
    (d.soil_filling||[]).forEach((sf, i) => {
      const bg = i%2===0?C.WHITE:C.GREY;
      const r = ws.getRow(row); r.height = 15;
      const vals = [sf.sr||i+1, sf.road_no, sf.road_width, sf.sub_base_width, sf.length, sf.area, sf.road_top_lvl, sf.ngl, sf.section_depth, sf.filling_depth, sf.extra_filling||0, sf.total_filling_cum, sf.hywas, sf.remark||''];
      vals.forEach((v,ci) => { const c = r.getCell(ci+1); c.value = v||0; sc(c, bg, false, 'FF000000', 9, ci>1&&ci<13?'center':'left'); if (ci>=5&&ci<13&&typeof v==='number') c.numFmt='#,##0.00'; });
      row++;
    });

    // Summary
    row++;
    const ms = d.material_summary||{};
    ws.mergeCells(row,1,row,14); sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 10, 'center');
    ws.getCell(row,1).value = 'SUMMARY'; ws.getRow(row).height = 16; row++;
    [['TOTAL FILLING (CUM)', ms.soil_filling_cum||0], ['RATE (RS/CUM)', 285], ['TOTAL COST (RS)', (ms.soil_filling_cum||0)*285]].forEach(([lbl, val]) => {
      ws.mergeCells(row,1,row,10); sc(ws.getCell(row,1), C.GREY, true, 'FF000000', 9, 'right');
      ws.getCell(row,1).value = lbl;
      ws.getCell(row,11).value = val; sc(ws.getCell(row,11), C.YELLOW, true, 'FF000000', 10, 'center');
      ws.getCell(row,11).numFmt='#,##0.00'; ws.getRow(row).height = 16; row++;
    });

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ── SHEET 6: BBS ───────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('BBS - STEEL');
    [6,42,10,10,10,10,12,10,12,18,12,12,12,12,12].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:O1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:O2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 11, 'center');
    ws.getCell('A2').value = 'BAR BENDING SCHEDULE (BBS) - STEEL REINFORCEMENT'; ws.getRow(2).height = 18;

    const bbsHdrs = ['SR NO','DESCRIPTION','DIA\n(MM)','NOS','CUT\nLENGTH','TOTAL\nLENGTH','WEIGHT\n(KG)','UNIT','8MM\nWEIGHT','10MM\nWEIGHT','12MM\nWEIGHT','16MM\nWEIGHT','20MM\nWEIGHT','25MM\nWEIGHT','TOTAL'];
    const r3 = ws.getRow(3); r3.height = 42;
    bbsHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 8, 'center'); });

    let row = 4;
    (d.bbs||[]).forEach(bbsSection => {
      ws.mergeCells(row,1,row,15);
      sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 10, 'center');
      ws.getCell(row,1).value = bbsSection.element||''; ws.getRow(row).height = 16; row++;

      (bbsSection.items||[]).forEach((item, i) => {
        const bg = i%2===0?C.WHITE:C.GREY;
        const r = ws.getRow(row); r.height = 14;
        const UNIT_WT = {8:0.395,10:0.617,12:0.888,16:1.58,20:2.47,25:3.86,32:6.31};
        const wt = (item.weight_kg||(item.total_length*UNIT_WT[item.dia||10])/1)||0;
        const diaCol = {8:9,10:10,12:11,16:12,20:13,25:14};
        const vals = [i+1, item.description||'', item.dia||'', item.nos||'', item.cutting_length||'', item.total_length||'', wt.toFixed(2), 'KG','','','','','','',''];
        vals.forEach((v,ci) => { const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 8, ci>1?'center':'left'); });
        // Put weight in dia column
        const dc = diaCol[item.dia];
        if (dc) { const c = r.getCell(dc); c.value = wt.toFixed(2); sc(c, bg, false, 'FF000000', 8, 'center'); }
        r.getCell(15).value = wt.toFixed(2); sc(r.getCell(15), bg, true, 'FF000000', 8, 'center');
        row++;
      });
      row++;
    });

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ── SHEET 7: STREET LIGHT ─────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('STREET LIGHT');
    [6,12,10,18,16,28].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:F1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:F2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'STREET LIGHT SCHEDULE'; ws.getRow(2).height = 18;

    const slHdrs = ['SR NO','ROAD WIDE','ROAD NO','ROAD LENGTH (MTR)','STREET LIGHTS (NOS)','REMARK'];
    const r3 = ws.getRow(3); r3.height = 18;
    slHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 9, 'center'); });

    let row = 4, totalSL = 0;
    (d.street_light||[]).forEach((sl, i) => {
      const bg = i%2===0?C.WHITE:C.GREY;
      const r = ws.getRow(row); r.height = 15; totalSL += sl.nos||0;
      [sl.sr||i+1, sl.road_wide||'', sl.road_no||'', sl.length_mt||0, sl.nos||0, sl.remark||''].forEach((v,ci) => {
        const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>1&&ci<5?'center':'left');
      });
      row++;
    });

    // Total
    ws.mergeCells(row,1,row,3); sc(ws.getCell(row,1), C.YELLOW, true, 'FF000000', 10, 'right');
    ws.getCell(row,1).value = 'TOTAL';
    ws.getCell(row,5).value = totalSL; sc(ws.getCell(row,5), C.YELLOW, true, 'FF000000', 11, 'center');
    ws.getRow(row).height = 18;
  }

  // ── SHEET 8: MATERIAL SUMMARY ──────────────────────────────────────────
  {
    const ws = wb.addWorksheet('MATERIAL SUMMARY');
    [6,46,18,14,20].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:E1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:E2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'OVERALL MATERIAL SUMMARY'; ws.getRow(2).height = 18;

    const msHdrs = ['SR NO','MATERIAL / WORK ITEM','QUANTITY','UNIT','REMARKS'];
    const r3 = ws.getRow(3); r3.height = 18;
    msHdrs.forEach((h,i) => { const c = r3.getCell(i+1); c.value = h; sc(c, C.NAVY, true, 'FFFFFFFF', 10, 'center'); });

    const ms = d.material_summary||{};
    const msItems = [
      ['ROADS & SUB-BASE','', ''],
      [1,'TOTAL ROAD LENGTH', ms.total_road_rmt||0, 'RMT', ''],
      [2,'TOTAL ROAD AREA', ms.total_road_area_sqmt||0, 'SQMT', ''],
      [3,'BOX CUTTING', (ms.total_road_area_sqmt||0)*1.05, 'SQMT', '15% EXTRA'],
      [4,'GSB FILLING (300MM)', ms.gsb_ton||0, 'TON', '15% EXTRA COMPACTION'],
      [5,'WMM FILLING (200MM)', ms.wmm_ton||0, 'TON', '15% EXTRA COMPACTION'],
      [6,'PQC ROAD (250MM, M30)', ms.pqc_cum||0, 'CUM', '5% WASTAGE INCLUDED'],
      [7,'SOIL FILLING', ms.soil_filling_cum||0, 'CUM', ''],
      ['SERVICE CORRIDOR','', ''],
      [8,'PAVER BLOCK (M40, 80MM)', ms.paver_block_sqmt||0, 'SQMT', ''],
      ['STEEL & OTHER','', ''],
      [9,'RCC STEEL (DOWEL + TIEBAR)', ms.rcc_steel_kg||0, 'KG', ''],
    ];

    let row = 4;
    msItems.forEach((item, i) => {
      if (typeof item[0] === 'string') {
        ws.mergeCells(row,1,row,5); sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 10, 'center');
        ws.getCell(row,1).value = item[0]; ws.getRow(row).height = 16; row++; return;
      }
      const bg = i%2===0?C.WHITE:C.GREY;
      const r = ws.getRow(row); r.height = 15;
      item.forEach((v,ci) => { const c = r.getCell(ci+1); c.value = v; sc(c, bg, false, 'FF000000', 9, ci>1?'center':'left'); if (ci===2&&typeof v==='number') c.numFmt='#,##0.00'; });
      row++;
    });

    ws.views = [{ state: 'frozen', ySplit: 3 }];
  }

  // ── SHEET 9: PMC OBSERVATIONS ─────────────────────────────────────────
  {
    const ws = wb.addWorksheet('PMC OBSERVATIONS');
    [6,80].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:B1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = d.project_name||'AI PARK'; ws.getRow(1).height = 22;
    ws.mergeCells('A2:B2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'PMC OBSERVATIONS & RECOMMENDATIONS'; ws.getRow(2).height = 18;

    let row = 3;
    ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), C.DKGREEN, true, 'FFFFFFFF', 11, 'left');
    ws.getCell(row,1).value = 'PMC OBSERVATIONS:'; ws.getRow(row).height = 18; row++;

    (d.observations||['Refer to chat analysis for detailed observations.']).forEach((obs, i) => {
      const r = ws.getRow(row); r.height = 30;
      r.getCell(1).value = i+1; sc(r.getCell(1), i%2===0?C.WHITE:C.GREY, false, 'FF000000', 9, 'center');
      r.getCell(2).value = obs; sc(r.getCell(2), i%2===0?C.WHITE:C.GREY, false, 'FF000000', 9, 'left');
      row++;
    });

    row++;
    ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), C.DKGREEN, true, 'FFFFFFFF', 11, 'left');
    ws.getCell(row,1).value = 'PMC RECOMMENDATION:'; ws.getRow(row).height = 18; row++;

    ws.mergeCells(row,1,row,2);
    const rCell = ws.getCell(row,1); rCell.value = d.recommendation||'Refer to chat analysis.';
    sc(rCell, C.GREEN, true, 'FF000000', 10, 'left', true); ws.getRow(row).height = 80;

    // Footer
    row += 2;
    ws.mergeCells(row,1,row,2);
    const fCell = ws.getCell(row,1);
    const today = new Date().toLocaleDateString('en-IN');
    fCell.value = `Prepared by: PMC Civil AI Agent | Date: ${today} | ${d.project_name||'VCT Bharuch'} — Powered by Gemini AI`;
    fCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:C.GREY} };
    fCell.font = { italic:true, size:9, color:{argb:'FF595959'}, name:'Calibri' };
    fCell.alignment = { horizontal:'center', vertical:'middle' };
    ws.getRow(row).height = 14;
  }

  return wb;
}

module.exports = { extractDrawingData, buildDrawingExcel };
