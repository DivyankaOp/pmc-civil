'use strict';
const assert = require('assert');
const path = require('path');

// Isolate env
process.env.CLAUDE_API_KEY = 'sk-test';
process.env.PMC_CLAUDE_MODE = 'strict';
delete process.env.PMC_ALLOW_CLAUDE_DXF;
delete process.env.PMC_ALLOW_CLAUDE_EXPORT;
delete process.env.PMC_ALLOW_CLAUDE_VISION;

const policy = require('../lib/claude_policy');

assert.strictEqual(policy.mode(), 'strict');
assert.strictEqual(policy.allowClaude('vision_takeoff_local_weak').ok, true);
assert.strictEqual(policy.allowClaude('borehole_ocr_weak').ok, true);
assert.strictEqual(policy.allowClaude('freeform_chat').ok, true);
assert.strictEqual(policy.allowClaude('schedule_weak_fallback').ok, false);
assert.strictEqual(policy.allowClaude('analyze_dwg_vision').ok, false);
assert.strictEqual(policy.allowClaude('drawing_to_excel').ok, false);
assert.strictEqual(policy.allowClaude('export_excel').ok, false);
assert.strictEqual(policy.allowClaude('dxf_unknown_symbols').ok, false);

process.env.PMC_ALLOW_CLAUDE_DXF = '1';
assert.strictEqual(policy.allowClaude('dxf_unknown_symbols').ok, true);

process.env.PMC_ALLOW_CLAUDE_EXPORT = '1';
assert.strictEqual(policy.allowClaude('export_excel').ok, true);

process.env.PMC_ALLOW_CLAUDE_VISION = '0';
assert.strictEqual(policy.allowClaude('vision_takeoff_local_weak').ok, false);

process.env.PMC_CLAUDE_MODE = 'off';
assert.strictEqual(policy.allowClaude('freeform_chat').ok, false);

console.log('claude_policy tests OK');
