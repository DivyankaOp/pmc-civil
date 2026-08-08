'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { buildPlanMeasure } = require('../lib/plan_measure');
const { buildAnnotatedTakeoffPdf } = require('../lib/annotated_takeoff_pdf');
const { createProject, addSheet, mergeProjectMarkdown } = require('../lib/project_store');
const { detectIntent, readDrawingFully } = require('../lib/drawing_reader');

async function main() {
  assert.strictEqual(detectIntent('measure plan areas and lengths'), 'measure');

  const text = `
SCALE : 1:100
Overall length 12500 mm
Bay 6850
Floor area 450 sqm
SCHEDULE OF FOOTING
F1 2600x1800 900
`;
  const m = buildPlanMeasure({ text });
  assert(m.scale?.denominator === 100, 'scale');
  assert(m.items.some(i => i.type === 'area'), 'area item');
  assert(m.items.some(i => i.type === 'length' || i.type === 'count'), 'length/count');

  const p = createProject('Test');
  addSheet(p.id, { filename: 'A.pdf', markdown: '# Sheet A', drawing_type: 'FOUNDATION_FOOTING', schedule_rows: 2 });
  addSheet(p.id, { filename: 'B.pdf', markdown: '# Sheet B', drawing_type: 'FLOOR_PLAN', schedule_rows: 1 });
  const merged = mergeProjectMarkdown(p.id);
  assert(/Sheet 1/.test(merged.markdown) && /Sheet 2/.test(merged.markdown), 'merge');

  const pdf = await buildAnnotatedTakeoffPdf({
    markdown: '# Takeoff\nF1 1500x1500',
    extracted: { schedules: { footings: [{ mark: 'F1', rcc_size_mm: '1500x1500', depth_mm: 450, qty: 4 }], columns: [] } },
    measure: m,
    title: 'Test',
  });
  assert(pdf.bytes.length > 500, 'pdf bytes');
  assert(pdf.pages >= 1, 'pages');

  const fixture = path.join(__dirname, '../data/fixtures/bhagyeshree_extract.txt');
  if (fs.existsSync(fixture)) {
    const r = readDrawingFully({ text: fs.readFileSync(fixture, 'utf8'), filename: 'F.pdf', question: 'measure plan areas' });
    assert(r.intent === 'measure', 'intent measure');
    assert(/Plan measure/i.test(r.markdown), 'measure md');
  }

  console.log('PASS: civils features (measure + annotated pdf + multi-sheet)');
}

main().catch(e => { console.error('FAIL', e); process.exit(1); });
