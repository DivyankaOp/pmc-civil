'use strict';
/**
 * Offline tune/test for schedule-first pipeline (no Claude API).
 * Usage: node scripts/test_schedule_pipeline.js [path-to-text-fixture]
 */
const fs = require('fs');
const path = require('path');
const { extractSchedules, formatSchedulesMarkdown } = require('../lib/schedule_extractor');
const { buildBoqFromSchedules, formatBoqMarkdown } = require('../lib/boq_from_schedules');

const fixture = process.argv[2] || path.join(__dirname, '..', 'data', 'fixtures', 'sample_schedule_text.txt');
const text = fs.readFileSync(fixture, 'utf8');
const extracted = extractSchedules(text);
// Test uses USER-CONFIRMED height/PCC — production never assumes these
const boq = buildBoqFromSchedules(extracted, {
  defaultStoreyHeightM: 3.0,
  pccThicknessM: 0.15,
  allowExcavationSurcharge: true,
});

console.log('=== SCHEDULES ===');
console.log(formatSchedulesMarkdown(extracted));
console.log('\n=== BOQ ===');
console.log(formatBoqMarkdown(boq));
console.log('\n=== SUMMARY ===');
console.log({
  quality: extracted.quality,
  rows: extracted.total_schedule_rows,
  columns: extracted.schedules.columns.length,
  footings: extracted.schedules.footings.length,
  doors: extracted.schedules.doors.length,
  windows: extracted.schedules.windows.length,
  boq_items: boq.boq.length,
  total_inr: boq.cost_summary.civil_total_inr,
});

const ok = extracted.schedules.columns.length >= 3
  && extracted.schedules.footings.length >= 2
  && boq.boq.length >= 5
  && extracted.schedules.columns[0].size_mm === '300x450'
  && extracted.schedules.columns[0].qty === 12;

if (!ok) {
  console.error('\nFAIL: fixture expectations not met — tune parsers');
  process.exit(1);
}
console.log('\nPASS: schedule-first fixture OK');
