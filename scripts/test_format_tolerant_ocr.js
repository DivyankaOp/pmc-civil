'use strict';
/**
 * Multi-format footing OCR — drawings never share one layout.
 */
const assert = require('assert');
const { extractOcrFootingSchedule, extractOcrColumnSchedule } = require('../lib/ocr_footing_schedule');
const { extractSchedules } = require('../lib/schedule_extractor');
const fs = require('fs');
const path = require('path');

function marks(rows) {
  return rows.map(r => `${r.mark}:${r.rcc_size_mm}:d${r.depth_mm || '?'}:q${r.qty ?? '?'}`);
}

// Format A — clean residential table lines
const A = `
FOOTING SCHEDULE
Mark Size Depth Qty
F1 1500x1500 450 12
F2 1800x1800 500 8
`;

// Format B — schedule of footing + pipe cells
const B = `
SCHEDULE OF FOOTING
F1 | 2000 x 2000 | 600 | 6
F2 | 2500 x 2000 | 700 | 4
`;

// Format C — industrial band (Bhagyeshree-like)
const C = fs.readFileSync(
  path.join(__dirname, '../data/fixtures/bhagyeshree_extract.txt'),
  'utf8'
);

// Format D — marks and sizes on separate OCR lines
const D = `
FTG SCH
F1
2100x2100
550
F2
2400x1800
600
`;

let failed = 0;
function check(name, fn) {
  try {
    fn();
    console.log('PASS', name);
  } catch (e) {
    failed++;
    console.error('FAIL', name, e.message);
  }
}

check('format-A line table', () => {
  const r = extractOcrFootingSchedule(A);
  assert(r.length >= 2, `expected >=2 got ${r.length}`);
  assert(r.some(x => x.mark === 'F1' && /1500x1500/.test(x.rcc_size_mm)));
  assert(r.find(x => x.mark === 'F1').qty === 12);
});

check('format-B schedule of footing pipes', () => {
  const r = extractOcrFootingSchedule(B);
  assert(r.length >= 2, marks(r).join(','));
  assert(r.some(x => /2000x2000/.test(x.rcc_size_mm)));
});

check('format-C industrial band (Bhagyeshree)', () => {
  const r = extractOcrFootingSchedule(C);
  assert(r.length >= 3, marks(r).join(','));
  assert(r.every(x => x.depth_mm), 'depths required');
  const cols = extractOcrColumnSchedule(C);
  assert(cols.length >= 2, 'columns');
});

check('format-D mark-near-size multiline', () => {
  const r = extractOcrFootingSchedule(D);
  assert(r.length >= 2, marks(r).join(','));
  assert(r.some(x => /2100x2100/.test(x.rcc_size_mm)));
});

check('extractSchedules merges OCR when empty', () => {
  const s = extractSchedules(C, {});
  assert(s.schedules.footings.length >= 3, 'footings from OCR enrich');
});

if (failed) {
  console.error(`\n${failed} failed`);
  process.exit(1);
}
console.log('\nAll format-tolerant OCR checks passed');
