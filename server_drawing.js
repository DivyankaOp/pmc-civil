// ─── DRAWING ANALYSIS → MULTI-SHEET PMC EXCEL ──────────────────────────────
// This module adds /export-drawing endpoint to server.js
// When user uploads a civil drawing → Gemini analyzes → Multi-sheet Excel output

const ExcelJS = require('exceljs');

// Colors
const C = {
  NAVY:    'FF1F3864',
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
