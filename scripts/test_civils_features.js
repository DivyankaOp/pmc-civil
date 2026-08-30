'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { buildPlanMeasure } = require('../lib/plan_measure');
const { buildAnnotatedTakeoffPdf } = require('../lib/annotated_takeoff_pdf');
const { createProject, addSheet, mergeProjectMarkdown } = require('../lib/project_store');
const { detectIntent, readDrawingFully } = require('../lib/drawing_reader');
const { extractGeotech } = require('../lib/geotech_extract');
const { buildQuestionFromScope, runTradeTakeoff, AGENTS } = require('../lib/takeoff_agents');
const { buildAgs41 } = require('../lib/ags_export');
const { buildGeoJson } = require('../lib/geojson_export');
const { buildEarthworks } = require('../lib/earthworks');
const { buildShapefileZip } = require('../lib/shapefile_export');
const { buildGroundModel } = require('../lib/ground_model');
const { buildQaChecklist, approveQaSession, assertExportAllowed } = require('../lib/qa_review');
const { buildGroundworksTakeoff } = require('../lib/paving_earthworks');
const { runAutoVisionTakeoff, parseVisionJson } = require('../lib/vision_takeoff');
const { digitiseBoreholes, denoiseBoreholeText, scoreFactoryGeotech } = require('../lib/borehole_digitiser');

async function main() {
  assert.strictEqual(detectIntent('measure plan areas and lengths'), 'measure');
  assert.strictEqual(detectIntent('extract geotech borehole SPT'), 'geotech');
  assert.strictEqual(detectIntent('calculate earthworks cut and fill'), 'earthworks');
  assert.strictEqual(detectIntent('Paving / road layer takeoff asphalt DBM'), 'paving');
  assert.strictEqual(detectIntent('Groundworks takeoff cut/fill plus paving'), 'groundworks');
  assert.strictEqual(detectIntent('Analyze uploaded file(s) and pull all proper details.'), 'takeoff');
  assert.strictEqual(detectIntent('', 'calculate'), 'takeoff');
  const autoBoq = readDrawingFully({
    text: 'SCHEDULE OF FOOTING\nF1 1500x1500 450 4\nF2 1800x1800 500 2',
    filename: 'F.pdf',
    question: 'Analyze uploaded file',
    autoBoq: true,
  });
  assert.strictEqual(autoBoq.intent, 'takeoff');
  assert(/Quantity Takeoff|BOQ|RCC|cum/i.test(autoBoq.markdown), 'auto boq md');
  assert((autoBoq.boqResult?.boq?.length || 0) >= 1 || (autoBoq.extracted?.schedules?.footings?.length || 0) >= 1, 'auto boq rows');
  assert(AGENTS.some(a => a.id === 'earthworks'), 'earthworks agent');
  assert(AGENTS.some(a => a.id === 'paving'), 'paving agent');
  assert(AGENTS.some(a => a.id === 'groundworks'), 'groundworks agent');

  const study = readDrawingFully({
    text: 'SCHEDULE OF FOOTING\nF1 1500x1500 450\nSCALE 1:100\nNGL = 12.5',
    filename: 'F.pdf',
    question: 'study',
    action: 'study',
  });
  assert.strictEqual(study.intent, 'study');
  assert(/Drawing Study|How to read|quiz/i.test(study.markdown), 'study md');
  assert(!/Quantity Takeoff/i.test(study.markdown), 'study no takeoff');

  const readOnly = readDrawingFully({
    text: 'SCHEDULE OF FOOTING\nF1 1500x1500 450',
    filename: 'F.pdf',
    action: 'read',
    scope: { agent: 'full', items: ['footings'] }, // must be ignored
  });
  assert.strictEqual(readOnly.intent, 'read');
  assert(/Drawing Reading Report/i.test(readOnly.markdown), 'read md');

  const text = `
SCALE : 1:100
Overall length 12500 mm
Bay 6850
Floor area 450 sqm
Plot area 1200 sqm
Road area 320 sqm
NGL = 12.50
FGL = 12.80
SCHEDULE OF FOOTING
F1 2600x1800 900
BH-1 GL = 12.5 m WT = 10.2 m SPT = 18 clay sand Easting = 234567.12 Northing = 1234567.89
BH-2 GL = 12.4 m SPT = 22 murum
`;
  const m = buildPlanMeasure({ text });
  assert(m.scale?.denominator === 100, 'scale');
  assert(m.items.some(i => i.type === 'area'), 'area item');
  assert(m.items.some(i => /plot|floor|paving|road|printed/i.test(i.source + i.description)), 'labeled area');
  assert(m.geometry?.length >= 0, 'geometry');

  const geo = extractGeotech(text);
  assert(geo.boreholes.length >= 2, 'boreholes');

  const ags = buildAgs41(geo, { title: 'Test' });
  assert(/^\*\*PROJ/m.test(ags.text), 'ags proj');
  assert(/\*\*HOLE/.test(ags.text), 'ags hole');
  assert(/\*\*ISPT/.test(ags.text), 'ags ispt');
  assert(/\*\*SAMP/.test(ags.text), 'ags samp');
  assert(ags.holes >= 2, 'ags holes');

  const gj = buildGeoJson({ geotech: geo, measure: m, title: 'Test' });
  assert(gj.type === 'FeatureCollection', 'geojson');
  assert(gj.features.length >= 2, 'geo features');

  const shp = buildShapefileZip({ geotech: geo, measure: m, title: 'Test' });
  assert(shp.bytes && shp.bytes.length > 100, 'shapefile zip');
  assert(shp.points >= 1, 'shp points');

  const gm = buildGroundModel(geo, { field: 'avg_spt' });
  assert(gm.cells.length >= 1, 'ground model cells');

  const qa = buildQaChecklist({ extracted: { schedules: { footings: [{ mark: 'F1', depth_mm: 900, qty: 2 }], columns: [] } }, measure: m, geotech: geo, markdown: '# ok' });
  assert(qa.items.length >= 5, 'qa items');
  assert(qa.requires_human === true, 'human qa required');
  const qaBad = approveQaSession(qa, { reviewer: '', confirmedIds: [] });
  assert(!qaBad.ok, 'qa reject empty reviewer');
  const qaOk = approveQaSession(qa, { reviewer: 'QS Test', confirmedIds: ['second_pass', 'engineer_signoff'] });
  assert(qaOk.ok && qaOk.approval.approved, 'qa approve');
  assert(assertExportAllowed(qaOk.approval).ok, 'export allowed after qa');
  assert(!assertExportAllowed(null).ok, 'export blocked without qa');

  assert.strictEqual(detectIntent('Auto vision takeoff without clicking'), 'vision_takeoff');
  assert.strictEqual(detectIntent('Digitise handwritten borehole logs into AGS'), 'borehole_digitise');
  assert(/BH-1/.test(denoiseBoreholeText('B8-1 GL 12 SPT 18')), 'denoise bh');

  const parsed = parseVisionJson('```json\n{"items":[{"type":"area","qty":120,"unit":"sqm","description":"Road"}]}\n```');
  assert(parsed?.items?.[0]?.qty === 120, 'vision json');

  const vt = await runAutoVisionTakeoff({ text, schedules: null }); // no API — local fallback
  assert(vt.measure?.items?.length >= 1, 'vision local fallback');
  assert(/local|vision/i.test(vt.mode), 'vision mode');

  const dig = await digitiseBoreholes({ text }); // OCR path without tiles/API
  assert(dig.geotech.boreholes.length >= 1, 'digitise ocr holes');
  assert(dig.ags.holes >= 1, 'digitise ags');
  assert(dig.factory?.factory_score >= 0, 'factory score');
  assert(dig.geotech.boreholes[0].factory_grade, 'per-hole factory grade');
  const scored = scoreFactoryGeotech(dig.geotech.boreholes);
  assert(scored.factory_grade, 'factory grade aggregate');

  const ew = buildEarthworks({ text, measure: m });
  assert(ew.ngl_m === 12.5, 'ngl');
  assert(ew.formation_m === 12.8, 'fgl');
  assert(ew.items.some(i => i.type === 'fill'), 'fill item');

  const gw = buildGroundworksTakeoff({ text, measure: m });
  assert(gw.paving, 'paving takeoff');
  assert(gw.items.some(i => i.trade === 'paving' || i.type === 'paving' || /asphalt|WMM|GSB|paving/i.test(i.description || '')), 'paving lines');
  assert(gw.items.some(i => i.type === 'disposal' || i.type === 'borrow' || i.type === 'fill' || i.type === 'cut'), 'cut/fill/disposal');

  const q = buildQuestionFromScope({ agent: 'earthworks' });
  assert(/earthwork|cut/i.test(q), 'ew question');
  assert(/paving|asphalt|DBM/i.test(buildQuestionFromScope({ agent: 'paving' })), 'paving question');

  const trade = runTradeTakeoff({
    text,
    extracted: { schedules: { footings: [{ mark: 'F1', rcc_size_mm: '2600x1800', depth_mm: 900, qty: 4 }], columns: [], doors: [], windows: [] }, quality: 'medium', total_schedule_rows: 1 },
    scope: { agent: 'earthworks', items: ['cut', 'fill', 'areas', 'paving'] },
  });
  assert(trade.earthworks, 'trade earthworks');
  assert(trade.paving || trade.groundworks, 'trade paving/groundworks');
  assert(/Earthworks|cut|fill|Groundworks|paving/i.test(trade.markdown), 'trade md');

  const p = createProject('Test');
  addSheet(p.id, { filename: 'A.pdf', markdown: '# Sheet A', drawing_type: 'FOUNDATION_FOOTING', schedule_rows: 2 });
  addSheet(p.id, { filename: 'B.pdf', markdown: '# Sheet B', drawing_type: 'FLOOR_PLAN', schedule_rows: 1 });
  const merged = mergeProjectMarkdown(p.id);
  assert(/Sheet 1/.test(merged.markdown) && /Sheet 2/.test(merged.markdown), 'merge');

  const pdf = await buildAnnotatedTakeoffPdf({
    markdown: '# Takeoff\nF1 1500x1500',
    extracted: { schedules: { footings: [{ mark: 'F1', rcc_size_mm: '1500x1500', depth_mm: 450, qty: 4 }], columns: [] } },
    measure: m,
    geotech: geo,
    earthworks: ew,
    title: 'Test',
  });
  assert(pdf.bytes.length > 500, 'pdf bytes');
  assert(pdf.pages >= 1, 'pages');
  assert(pdf.callouts >= 1, 'callouts');

  const fixture = path.join(__dirname, '../data/fixtures/bhagyeshree_extract.txt');
  if (fs.existsSync(fixture)) {
    const r = readDrawingFully({ text: fs.readFileSync(fixture, 'utf8'), filename: 'F.pdf', question: 'measure plan areas' });
    assert(r.intent === 'measure', 'intent measure');
    assert(/Plan measure/i.test(r.markdown), 'measure md');
  }

  console.log('PASS: civils gaps v4 (human QA gate + paving/groundworks + BH factory score)');
}

main().catch(e => { console.error('FAIL', e); process.exit(1); });
