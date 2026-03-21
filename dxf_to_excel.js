/**
 * DXF → PMC Excel with full quantity calculations
 * Uses parsed DXF data + Gemini for interpretation + formula engine
 */

const RATES = {
  // ROADS (Gujarat DSR 2025)
  soil_stabilization:  82,    // ₹/sqmt
  soil_filling:       285,    // ₹/cum
  gsb_300mm:          655,    // ₹/sqmt
  wmm_200mm:          515,    // ₹/sqmt
  pqc_road_250mm:    1800,    // ₹/sqmt
  asphalt_60mm:       750,    // ₹/sqmt
  service_corridor:  1790,    // ₹/sqmt
  paver_block_80mm:   750,    // ₹/sqmt
  kerbing:            350,    // ₹/rmt
  steel_dowel:      56000,    // ₹/ton
  street_light:     35000,    // ₹/nos
  // STRUCTURE
  rcc_m20:           5200,    // ₹/cum
  rcc_m25:           5500,    // ₹/cum
  rcc_m30:           5800,    // ₹/cum
  pcc_m10:           3800,    // ₹/cum
  brickwork_230:     4500,    // ₹/cum
  plaster:            120,    // ₹/sqmt
  steel_fe500:         56,    // ₹/kg
  // CIVIL
  excavation:         180,    // ₹/cum
  compound_wall:     8600,    // ₹/rmt (CP wall)
  gabion_wall:      14100,    // ₹/rmt
};

function calcRoad(len, width) {
  const cw = Math.max(width * 0.65, width - 3); // carriage way
  const area = len * cw;
  return {
    road_no: '', length_m: len, total_width_m: width, carriage_width_m: Math.round(cw*100)/100,
    area_sqmt:      Math.round(area * 100) / 100,
    box_cut_sqmt:   Math.round(area * 1.05 * 100) / 100,
    gsb_ton:        Math.round(area * 1.15 * 0.300 * 1.800 * 100) / 100,
    wmm_ton:        Math.round(area * 1.15 * 0.200 * 2.100 * 100) / 100,
    pqc_cum:        Math.round(area * 1.05 * 0.250 * 100) / 100,
    steel_dowel_ton: Math.round(area * 0.00387 * 100) / 100,
    gsb_cost:   Math.round(area * RATES.gsb_300mm),
    wmm_cost:   Math.round(area * RATES.wmm_200mm),
    pqc_cost:   Math.round(area * RATES.pqc_road_250mm),
    total_cost: Math.round(area * (RATES.gsb_300mm + RATES.wmm_200mm + RATES.pqc_road_250mm))
  };
}

async function buildDXFExcel(dxfData, geminiInterpretation, ExcelJS) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent — DXF Analysis';

  const C = {
    NAVY: 'FF1F3864', MIDBLUE: 'FF2E75B6', LTBLUE: 'FFBDD7EE',
    YELLOW: 'FFFFD966', GREEN: 'FFE2EFDA', DKGREEN: 'FF375623',
    GREY: 'FFF2F2F2', WHITE: 'FFFFFFFF', LOWEST: 'FF00B050'
  };
  const thin = { style:'thin', color:{argb:'FF000000'} };
  const bdr = { top:thin, left:thin, bottom:thin, right:thin };

  function sc(cell, bg, bold=false, fc='FF000000', size=9, align='center', wrap=true) {
    if (bg) cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:bg} };
    cell.font = { bold, color:{argb:fc}, size, name:'Calibri' };
    cell.alignment = { horizontal:align, vertical:'middle', wrapText:wrap };
    cell.border = bdr;
  }
  function topRow(ws, cols, text, bg, fc='FFFFFFFF', size=12) {
    ws.mergeCells(1, 1, 1, cols);
    const c = ws.getCell(1,1); c.value = text;
    sc(c, bg, true, fc, size, 'center');
    ws.getRow(1).height = 22;
  }
  function hdrRow(ws, row, headers, bg=C.NAVY) {
    headers.forEach((h,i) => {
      const c = ws.getCell(row, i+1); c.value = h;
      sc(c, bg, true, 'FFFFFFFF', 9, 'center');
    });
    ws.getRow(row).height = 36;
  }

  const gi = geminiInterpretation || {};
  const projectName = gi.project_name || dxfData?.title_block?.project || 'CIVIL PROJECT';
  const today = new Date().toLocaleDateString('en-IN');

  // ═══════════════════════════════════════════════════════════════
  // SHEET 1: DRAWING SUMMARY
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('DRAWING SUMMARY');
    ws.getColumn(1).width = 6; ws.getColumn(2).width = 35;
    ws.getColumn(3).width = 22; ws.getColumn(4).width = 16; ws.getColumn(5).width = 14;

    ws.mergeCells('A1:E1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = projectName.toUpperCase(); ws.getRow(1).height = 24;

    ws.mergeCells('A2:E2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'DXF DRAWING ANALYSIS — PMC CIVIL REPORT'; ws.getRow(2).height = 18;

    // Drawing info table
    const info = [
      ['PROJECT NAME', projectName],
      ['DRAWING FILE', dxfData?.filename || ''],
      ['SCALE', dxfData?.scale || gi.scale || 'Not detected'],
      ['DATE', gi.date || today],
      ['DRAWING TYPE', gi.drawing_type || 'SITE LAYOUT'],
      ['TOTAL LAYERS', dxfData?.stats?.total_layers || 0],
      ['TOTAL ENTITIES', (dxfData?.stats?.total_lines||0) + (dxfData?.stats?.total_polylines||0)],
      ['TOTAL TEXTS', dxfData?.stats?.total_texts || 0],
      ['TOTAL DIMENSIONS', dxfData?.stats?.total_dims || 0],
      ['DRAWING EXTENTS', `${dxfData?.drawing_extents?.estimated_width_m||0}m × ${dxfData?.drawing_extents?.estimated_height_m||0}m`],
      ['PREPARED BY', 'PMC Civil AI Agent'],
      ['ANALYSIS DATE', today],
    ];

    let row = 3;
    info.forEach(([lbl, val], i) => {
      const bg = i%2===0 ? C.LTBLUE : C.GREY;
      ws.mergeCells(row,1,row,2); const lc = ws.getCell(row,1); lc.value = lbl;
      sc(lc, bg, true, 'FF000000', 9, 'left');
      ws.mergeCells(row,3,row,5); const vc = ws.getCell(row,3); vc.value = String(val);
      sc(vc, bg, false, 'FF000000', 9, 'left');
      ws.getRow(row).height = 15; row++;
    });

    // Layers found
    row++;
    ws.mergeCells(row,1,row,5); sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 10, 'center');
    ws.getCell(row,1).value = 'LAYERS FOUND IN DRAWING'; ws.getRow(row).height = 16; row++;

    const layers = dxfData?.layer_names || [];
    layers.forEach((lyr, i) => {
      const bg = i%2===0 ? C.WHITE : C.GREY;
      ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), bg, false, 'FF000000', 9, 'center');
      ws.getCell(row,1).value = i+1;
      ws.mergeCells(row,3,row,5); sc(ws.getCell(row,3), bg, false, 'FF000000', 9, 'left');
      ws.getCell(row,3).value = lyr;
      ws.getRow(row).height = 14; row++;
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 2: ALL TEXTS EXTRACTED
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('EXTRACTED TEXT');
    ws.getColumn(1).width = 6; ws.getColumn(2).width = 80; ws.getColumn(3).width = 20;

    ws.mergeCells('A1:C1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 13, 'center');
    ws.getCell('A1').value = projectName + ' — ALL TEXT EXTRACTED FROM DXF'; ws.getRow(1).height = 20;
    hdrRow(ws, 2, ['SR NO', 'TEXT CONTENT', 'CATEGORY'], C.NAVY);

    const allTexts = dxfData?.all_texts || [];
    const catKeywords = {
      'DIMENSION': /^\d+\.?\d*$|×|x\d|\d+\s*m\b/i,
      'ROAD/AREA LABEL': /road|r-\d|r\d|block|plot|zone|sector/i,
      'MATERIAL/SPEC': /gsb|wmm|pqc|rcc|pcc|m-?\d+|thk|mm|cm/i,
      'TITLE BLOCK': /scale|date|drg|drawing|project|prepared/i,
      'ANNOTATION': /.*/
    };

    let row = 3;
    allTexts.forEach((txt, i) => {
      if (!txt || txt.length < 2) return;
      const bg = i%2===0 ? C.WHITE : C.GREY;
      let cat = 'ANNOTATION';
      for (const [c, re] of Object.entries(catKeywords)) {
        if (re.test(txt)) { cat = c; break; }
      }
      sc(ws.getCell(row,1), bg, false, 'FF000000', 9, 'center'); ws.getCell(row,1).value = i+1;
      sc(ws.getCell(row,2), bg, false, 'FF000000', 9, 'left', true); ws.getCell(row,2).value = txt;
      sc(ws.getCell(row,3), bg, false, 'FF000000', 9, 'center'); ws.getCell(row,3).value = cat;
      ws.getRow(row).height = 14; row++;
    });

    ws.views = [{ state:'frozen', ySplit:2 }];
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 3: DIMENSIONS EXTRACTED
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('DIMENSIONS');
    [6,20,18,18,18,20].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:F1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 13, 'center');
    ws.getCell('A1').value = projectName + ' — DIMENSION VALUES FROM DXF'; ws.getRow(1).height = 20;
    hdrRow(ws, 2, ['SR','LAYER','VALUE (MM)','VALUE (M)','VALUE (FT)','NOTES'], C.NAVY);

    const dims = dxfData?.dimension_values || [];
    let row = 3;
    dims.forEach((d, i) => {
      const bg = i%2===0 ? C.WHITE : C.GREY;
      const vals = [i+1, d.layer||'', d.value_mm||0, d.value_m||0, Math.round((d.value_m||0)*3.281*100)/100, d.text||''];
      vals.forEach((v,ci) => {
        sc(ws.getCell(row,ci+1), bg, false, 'FF000000', 9, ci>0?'center':'center');
        ws.getCell(row,ci+1).value = v;
        if (ci>=2&&ci<=4) ws.getCell(row,ci+1).numFmt = '#,##0.00';
      });
      ws.getRow(row).height = 14; row++;
    });

    // Annotations from text
    if (dxfData?.spaces_from_annotations?.length) {
      row += 2;
      ws.mergeCells(row,1,row,6); sc(ws.getCell(row,1), C.MIDBLUE, true, 'FFFFFFFF', 10, 'center');
      ws.getCell(row,1).value = 'DIMENSIONS FROM TEXT ANNOTATIONS (L × W)'; ws.getRow(row).height = 16; row++;
      hdrRow(ws, row, ['SR','ANNOTATION TEXT','LENGTH (M)','WIDTH (M)','AREA (SQMT)','SOURCE'], C.MIDBLUE); row++;
      dxfData.spaces_from_annotations.forEach((s, i) => {
        const bg = i%2===0?C.WHITE:C.GREY;
        [i+1, s.label||'', s.length||0, s.width||0, s.area||0, 'TEXT ANNOTATION'].forEach((v,ci) => {
          sc(ws.getCell(row,ci+1), bg, false, 'FF000000', 9, ci>0?'center':'center');
          ws.getCell(row,ci+1).value = v;
          if (ci>=2) ws.getCell(row,ci+1).numFmt = '#,##0.00';
        });
        ws.getRow(row).height = 14; row++;
      });
    }

    ws.views = [{ state:'frozen', ySplit:2 }];
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 4: ROAD SCHEDULE & QUANTITIES (AI interpreted)
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('ROAD QUANTITIES');
    [6,8,10,14,14,14,15,15,14,15,16,24].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:L1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = projectName.toUpperCase(); ws.getRow(1).height = 22;
    ws.mergeCells('A2:L2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 11, 'center');
    ws.getCell('A2').value = 'ROAD WORK — QUANTITIES & COST ESTIMATE'; ws.getRow(2).height = 18;

    const hdrs = ['SR','ROAD\nNO','TOTAL\nWIDTH(M)','CARRIAGE\nWIDTH(M)','LENGTH\n(M)','AREA\n(SQMT)','GSB\n300MM(TON)','WMM\n200MM(TON)','PQC\n250MM(CUM)','STEEL\nDOWEL(TON)','TOTAL COST\n(₹)','REMARK'];
    hdrRow(ws, 3, hdrs); ws.getRow(3).height = 50;

    const roads = gi.roads || gi.elements?.filter(e => e.type==='ROAD') || [];
    let row = 4;
    let totArea=0, totGSB=0, totWMM=0, totPQC=0, totSteel=0, totCost=0;

    if (roads.length === 0) {
      // Use polyline areas as proxy for spaces
      const bigAreas = dxfData?.polyline_areas?.slice(0,20) || [];
      bigAreas.forEach((pa, i) => {
        if (pa.area_sqm < 10) return;
        const bg = i%2===0?C.WHITE:C.GREY;
        const r = ws.getRow(row); r.height = 15;
        const side = Math.sqrt(pa.area_sqm);
        const q = calcRoad(side, Math.min(side, 24));
        totArea+=q.area_sqmt; totGSB+=q.gsb_ton; totWMM+=q.wmm_ton; totPQC+=q.pqc_cum; totSteel+=q.steel_dowel_ton; totCost+=q.total_cost;
        [i+1,'AREA '+(i+1),'-','-',Math.round(side), pa.area_sqm, q.gsb_ton, q.wmm_ton, q.pqc_cum, q.steel_dowel_ton, q.total_cost, `Layer: ${pa.layer}`].forEach((v,ci) => {
          sc(r.getCell(ci+1), bg, false, 'FF000000', 9, 'center');
          r.getCell(ci+1).value = v;
          if (ci>=5) r.getCell(ci+1).numFmt = '#,##0.00';
        });
        row++;
      });
    } else {
      roads.forEach((rd, i) => {
        const bg = i%2===0?C.WHITE:C.GREY;
        const r = ws.getRow(row); r.height = 15;
        const dim = rd.dimensions || {};
        const q = calcRoad(dim.length_m||0, dim.width_m||0);
        q.road_no = rd.id || rd.name || `R${i+1}`;
        totArea+=q.area_sqmt; totGSB+=q.gsb_ton; totWMM+=q.wmm_ton; totPQC+=q.pqc_cum; totSteel+=q.steel_dowel_ton; totCost+=q.total_cost;
        [i+1, q.road_no, dim.width_m||0, q.carriage_width_m, dim.length_m||0, q.area_sqmt, q.gsb_ton, q.wmm_ton, q.pqc_cum, q.steel_dowel_ton, q.total_cost, rd.remark||''].forEach((v,ci) => {
          sc(r.getCell(ci+1), bg, false, 'FF000000', 9, 'center');
          r.getCell(ci+1).value = v;
          if (ci>=5&&ci<11) r.getCell(ci+1).numFmt = '#,##0.00';
          if (ci===10) r.getCell(ci+1).numFmt = '₹#,##0';
        });
        row++;
      });
    }

    // Totals
    ws.mergeCells(row,1,row,5); sc(ws.getCell(row,1), C.NAVY, true, 'FFFFFFFF', 10, 'right');
    ws.getCell(row,1).value = 'GRAND TOTAL';
    [totArea, totGSB, totWMM, totPQC, totSteel, totCost].forEach((v,i) => {
      const c = ws.getCell(row, i+6); c.value = v;
      c.numFmt = i===5 ? '₹#,##0' : '#,##0.00';
      sc(c, C.NAVY, true, 'FFFFFFFF', 10, 'center');
    });
    ws.getCell(row,12).value = ''; sc(ws.getCell(row,12), C.NAVY, false, 'FFFFFFFF', 9, 'center');
    ws.getRow(row).height = 18;

    ws.views = [{ state:'frozen', ySplit:3 }];
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 5: BOQ
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('BOQ');
    [6,50,12,14,14,18].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:F1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = projectName.toUpperCase(); ws.getRow(1).height = 22;
    ws.mergeCells('A2:F2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'BILL OF QUANTITIES'; ws.getRow(2).height = 18;
    hdrRow(ws, 3, ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (₹)','AMOUNT (₹)']);

    const boqItems = gi.boq || gi.cost_summary?.item_wise || [];
    const roads = gi.roads || gi.elements?.filter(e => e.type==='ROAD') || [];

    // Generate BOQ from extracted data
    const defaultBOQ = [
      { desc:'BOX CUTTING / EXCAVATION', unit:'SQMT', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};return s+(d.length_m||0)*(d.width_m||0)*1.05;},0) || (dxfData?.polyline_areas?.[0]?.area_sqm||0)*1.05, rate: RATES.excavation },
      { desc:'SOIL STABILIZATION (LIME-FLYASH)', unit:'SQMT', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};return s+(d.length_m||0)*(d.carriage_width_m||(d.width_m||0)*0.65);},0), rate: RATES.soil_stabilization },
      { desc:'SOIL FILLING IN ROAD AREA', unit:'CUM', qty: 0, rate: RATES.soil_filling },
      { desc:'GSB FILLING (300 MM LAYER)', unit:'SQMT', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};const a=(d.length_m||0)*(d.carriage_width_m||(d.width_m||0)*0.65);return s+a;},0), rate: RATES.gsb_300mm },
      { desc:'WMM FILLING (200 MM LAYER)', unit:'SQMT', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};const a=(d.length_m||0)*(d.carriage_width_m||(d.width_m||0)*0.65);return s+a;},0), rate: RATES.wmm_200mm },
      { desc:'PQC ROAD M30 (250 MM THICK)', unit:'SQMT', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};const a=(d.length_m||0)*(d.carriage_width_m||(d.width_m||0)*0.65);return s+a;},0), rate: RATES.pqc_road_250mm },
      { desc:'STEEL FOR DOWEL BARS', unit:'TON', qty: roads.reduce((s,r)=>{const d=r.dimensions||{};const a=(d.length_m||0)*(d.carriage_width_m||(d.width_m||0)*0.65);return s+a*0.00387;},0), rate: RATES.steel_dowel },
      { desc:'SERVICE CORRIDOR (PAVER BLOCK)', unit:'SQMT', qty: 0, rate: RATES.service_corridor },
      { desc:'COMPOUND WALL (CP WALL)', unit:'RMT', qty: 0, rate: RATES.compound_wall },
      { desc:'STREET LIGHTS', unit:'NOS', qty: 0, rate: RATES.street_light },
    ];

    const finalBOQ = boqItems.length > 0 ? boqItems.map(b => ({
      desc: b.description || b.item,
      unit: b.unit, qty: b.qty || b.quantity || 0, rate: b.rate || 0
    })) : defaultBOQ;

    let row = 4, grandTotal = 0;
    finalBOQ.forEach((item, i) => {
      const bg = i%2===0?C.WHITE:C.GREY;
      const amt = Math.round((item.qty||0) * (item.rate||0));
      grandTotal += amt;
      const r = ws.getRow(row); r.height = 15;
      [i+1, item.desc||'', item.unit||'', Math.round((item.qty||0)*100)/100, item.rate||0, amt].forEach((v,ci) => {
        sc(r.getCell(ci+1), bg, false, 'FF000000', 9, ci===1?'left':'center');
        r.getCell(ci+1).value = v;
        if (ci>=3) r.getCell(ci+1).numFmt = ci===5?'₹#,##0':'#,##0.00';
      });
      row++;
    });

    ws.mergeCells(row,1,row,4); sc(ws.getCell(row,1), C.YELLOW, true, 'FF000000', 11, 'right');
    ws.getCell(row,1).value = 'GRAND TOTAL';
    sc(ws.getCell(row,6), C.YELLOW, true, 'FF000000', 11, 'center');
    ws.getCell(row,6).value = grandTotal; ws.getCell(row,6).numFmt = '₹#,##0';
    ws.getRow(row).height = 20;

    ws.views = [{ state:'frozen', ySplit:3 }];
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 6: POLYLINE AREAS (detected closed regions)
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('DETECTED AREAS');
    [6,18,18,18,18,18,14].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:G1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 13, 'center');
    ws.getCell('A1').value = projectName + ' — CLOSED REGIONS DETECTED (POLYLINES)'; ws.getRow(1).height = 20;
    hdrRow(ws, 2, ['SR','LAYER','AREA (SQMT)','AREA (SQFT)','PERIMETER (M)','VERTICES','REMARKS']);

    const areas = dxfData?.polyline_areas || [];
    let row = 3, totalArea = 0;
    areas.forEach((pa, i) => {
      const bg = i%2===0?C.WHITE:C.GREY;
      const r = ws.getRow(row); r.height = 14; totalArea += pa.area_sqm||0;
      [i+1, pa.layer||'DEFAULT', pa.area_sqm||0, Math.round((pa.area_sqm||0)*10.764), pa.perimeter_m||0, pa.vertices||0, ''].forEach((v,ci) => {
        sc(r.getCell(ci+1), bg, false, 'FF000000', 9, ci===1?'left':'center');
        r.getCell(ci+1).value = v;
        if (ci>=2&&ci<=4) r.getCell(ci+1).numFmt = '#,##0.00';
      });
      row++;
    });

    ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), C.YELLOW, true, 'FF000000', 10, 'right');
    ws.getCell(row,1).value = 'TOTAL AREA';
    sc(ws.getCell(row,3), C.YELLOW, true, 'FF000000', 11, 'center');
    ws.getCell(row,3).value = Math.round(totalArea*100)/100; ws.getCell(row,3).numFmt='#,##0.00';
    ws.getRow(row).height = 18;

    ws.views = [{ state:'frozen', ySplit:2 }];
  }

  // ═══════════════════════════════════════════════════════════════
  // SHEET 7: PMC RECOMMENDATIONS
  // ═══════════════════════════════════════════════════════════════
  {
    const ws = wb.addWorksheet('PMC OBSERVATIONS');
    [6,85].forEach((w,i) => ws.getColumn(i+1).width = w);

    ws.mergeCells('A1:B1'); sc(ws.getCell('A1'), C.NAVY, true, 'FFFFFFFF', 14, 'center');
    ws.getCell('A1').value = projectName.toUpperCase(); ws.getRow(1).height = 22;
    ws.mergeCells('A2:B2'); sc(ws.getCell('A2'), C.MIDBLUE, true, 'FFFFFFFF', 12, 'center');
    ws.getCell('A2').value = 'PMC OBSERVATIONS & RECOMMENDATIONS'; ws.getRow(2).height = 18;

    let row = 3;
    ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), C.DKGREEN, true, 'FFFFFFFF', 11, 'left');
    ws.getCell(row,1).value = '  DXF ANALYSIS SUMMARY'; ws.getRow(row).height = 18; row++;

    const obsLines = [
      `Drawing: ${dxfData?.filename || 'DXF File'}`,
      `Scale detected: ${dxfData?.scale || 'Not found — please verify'}`,
      `Total layers in drawing: ${dxfData?.stats?.total_layers || 0}`,
      `Text annotations extracted: ${dxfData?.stats?.total_texts || 0}`,
      `Dimension entities: ${dxfData?.stats?.total_dims || 0}`,
      `Closed polylines (rooms/plots): ${dxfData?.polyline_areas?.length || 0}`,
      ...(gi.observations || []),
    ];

    obsLines.forEach((obs, i) => {
      const bg = i%2===0?C.WHITE:C.GREY;
      const r = ws.getRow(row); r.height = 22;
      r.getCell(1).value = i+1; sc(r.getCell(1), bg, false, 'FF000000', 9, 'center');
      r.getCell(2).value = obs; sc(r.getCell(2), bg, false, 'FF000000', 9, 'left', true);
      row++;
    });

    row++;
    ws.mergeCells(row,1,row,2); sc(ws.getCell(row,1), C.DKGREEN, true, 'FFFFFFFF', 11, 'left');
    ws.getCell(row,1).value = '  PMC RECOMMENDATION'; ws.getRow(row).height = 18; row++;

    ws.mergeCells(row,1,row,2);
    const rc = ws.getCell(row,1);
    rc.value = gi.pmc_recommendation || gi.recommendation ||
      `DXF file has been parsed. ${dxfData?.stats?.total_texts||0} text annotations and ${dxfData?.stats?.total_dims||0} dimension entities extracted. ` +
      `Verify scale factor (${dxfData?.scale||'not detected'}) and cross-check quantities with original drawing. ` +
      `All calculations based on Gujarat DSR 2025 rates.`;
    sc(rc, C.GREEN, true, 'FF000000', 10, 'left', true);
    ws.getRow(row).height = 80;

    row += 2;
    ws.mergeCells(row,1,row,2);
    ws.getCell(row,1).value = `Prepared by: PMC Civil AI Agent  |  Date: ${today}  |  Powered by Gemini AI`;
    ws.getCell(row,1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:C.GREY} };
    ws.getCell(row,1).font = { italic:true, size:9, color:{argb:'FF595959'}, name:'Calibri' };
    ws.getCell(row,1).alignment = { horizontal:'center', vertical:'middle' };
    ws.getRow(row).height = 14;
  }

  return wb;
}

module.exports = { buildDXFExcel, calcRoad, RATES };
