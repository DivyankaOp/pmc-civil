'use strict';
/** Quick offline checks for spatial tables + type detect + clarifications */
const assert = (c, m) => { if (!c) { console.error('FAIL:', m); process.exit(1); } };

const { buildSpatialScheduleText } = require('../lib/spatial_tables');
const { extractSchedules, buildHeaderMap } = require('../lib/schedule_extractor');
const { detectDrawingType } = require('../lib/drawing_types');
const { applyUserClarifications } = require('../lib/schedule_pipeline');

// Header map
const hm = buildHeaderMap(['Mark', 'PCC Size', 'RCC Size', 'Depth', 'Qty']);
assert(hm.mark === 0 && hm.pcc === 1 && hm.rcc === 2 && hm.depth === 3 && hm.qty === 4, 'header map');

// Spatial PDF-like spans → footing table
const pages = [{
  page: 1,
  texts: [
    { text: 'SCHEDULE OF FOOTING', x: 10, y: 20 },
    { text: 'Mark', x: 10, y: 40 }, { text: 'Size', x: 80, y: 40 }, { text: 'Depth', x: 160, y: 40 }, { text: 'Qty', x: 220, y: 40 },
    { text: 'F1', x: 10, y: 55 }, { text: '1500x1500', x: 80, y: 55 }, { text: '450', x: 160, y: 55 }, { text: '12', x: 220, y: 55 },
    { text: 'F2', x: 10, y: 70 }, { text: '1800x1800', x: 80, y: 70 }, { text: '500', x: 160, y: 70 }, { text: '8', x: 220, y: 70 },
  ],
}];
const spatial = buildSpatialScheduleText({ pdfPages: pages, plainLines: [] });
assert(spatial.tables.length >= 1, 'spatial tables detected');
const ex = extractSchedules(spatial.text, { spatialTables: spatial.tables });
assert(ex.schedules.footings.length >= 2, 'footing rows from spatial: ' + ex.schedules.footings.length);

// Type: footing schedule beats section keyword
const t = detectDrawingType({
  text: 'SECTION A-A\nSCHEDULE OF FOOTING\nF1 1500x1500',
  filename: 'mixed.pdf',
});
assert(t.drawing_type === 'FOUNDATION_FOOTING', 'type prefer footing schedule, got ' + t.drawing_type);

// Clarifications merge
const extracted = extractSchedules(`COLUMN SCHEDULE
Mark Size Qty
C1 300x450 
`);
const clarifications = {
  questions: [
    { id: 'col_qty_C1', question: 'qty?' },
    { id: 'col_height_C1', question: 'height?' },
  ],
};
const applied = applyUserClarifications({
  text: 'COLUMN SCHEDULE',
  extracted,
  clarifications,
  userText: '1) 12\n2) 3.15',
  filename: 't.pdf',
  question: 'boq calculate volumes',
});
const c1 = applied.extracted.schedules.columns.find(c => c.mark === 'C1');
assert(c1 && c1.qty === 12, 'merged qty');
assert(applied.opts.columnHeights?.C1 === 3.15, 'merged height');

console.log('PASS: improvements checks OK', {
  spatial_tables: spatial.tables.length,
  footings: ex.schedules.footings.length,
  type: t.drawing_type,
  answers: applied.answers,
});
