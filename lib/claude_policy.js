'use strict';
/**
 * Strict Claude token gate — tokens are costly.
 *
 * Modes (env PMC_CLAUDE_MODE):
 *   off    — never call Claude (local + ask-user only)
 *   strict — DEFAULT: Claude only for allowlisted last-resort reasons
 *   on     — legacy permissive (still logs every call)
 *
 * Also: PMC_ALLOW_CLAUDE_VISION=0 disables vision/BH Claude even in strict.
 *       PMC_ALLOW_CLAUDE_EXPORT=1 required for Claude-powered Excel/PDF extract.
 */

const ALLOWED_STRICT = new Set([
  // Explicit agent: local measure failed
  'vision_takeoff_local_weak',
  // Explicit agent: OCR borehole weak
  'borehole_ocr_weak',
  // Chat with NO drawing file (compare quotes, site report text, etc.)
  'freeform_chat',
  // Optional: many unknown DXF symbols (only if PMC_ALLOW_CLAUDE_DXF=1)
  'dxf_unknown_symbols',
]);

function mode() {
  const m = String(process.env.PMC_CLAUDE_MODE || 'strict').toLowerCase().trim();
  if (m === 'off' || m === '0' || m === 'never') return 'off';
  if (m === 'on' || m === 'all' || m === 'permissive') return 'on';
  return 'strict';
}

function hasKey() {
  return !!process.env.CLAUDE_API_KEY;
}

/**
 * @param {string} reason
 * @param {object} [meta]
 * @returns {{ ok: boolean, reason: string, mode: string, detail?: string }}
 */
function allowClaude(reason, meta = {}) {
  const m = mode();
  if (!hasKey()) {
    return { ok: false, reason, mode: m, detail: 'CLAUDE_API_KEY not set' };
  }
  if (m === 'off') {
    return { ok: false, reason, mode: m, detail: 'PMC_CLAUDE_MODE=off' };
  }

  // Vision / BH master kill-switch
  if ((reason === 'vision_takeoff_local_weak' || reason === 'borehole_ocr_weak')
    && process.env.PMC_ALLOW_CLAUDE_VISION === '0') {
    return { ok: false, reason, mode: m, detail: 'PMC_ALLOW_CLAUDE_VISION=0' };
  }

  if (reason === 'dxf_unknown_symbols' && process.env.PMC_ALLOW_CLAUDE_DXF !== '1') {
    return { ok: false, reason, mode: m, detail: 'PMC_ALLOW_CLAUDE_DXF not enabled (default off)' };
  }

  if (reason === 'export_excel' || reason === 'export_pdf' || reason === 'drawing_to_excel') {
    if (process.env.PMC_ALLOW_CLAUDE_EXPORT !== '1') {
      return { ok: false, reason, mode: m, detail: 'PMC_ALLOW_CLAUDE_EXPORT not enabled — use local takeoff Excel' };
    }
  }

  // Drawing analysis / auto BOQ / polish — NEVER in strict
  if (reason === 'drawing_polish'
    || reason === 'analyze_drawing_vision'
    || reason === 'analyze_dwg_vision'
    || reason === 'schedule_weak_fallback') {
    if (m === 'strict') {
      return { ok: false, reason, mode: m, detail: 'strict: drawing BOQ is local-only (ask user if weak)' };
    }
  }

  if (m === 'on') {
    return { ok: true, reason, mode: m, detail: meta.detail || 'permissive' };
  }

  // strict
  if (!ALLOWED_STRICT.has(reason) && reason !== 'export_excel' && reason !== 'export_pdf' && reason !== 'drawing_to_excel') {
    return { ok: false, reason, mode: m, detail: `reason "${reason}" not allowlisted in strict mode` };
  }
  if ((reason === 'export_excel' || reason === 'export_pdf' || reason === 'drawing_to_excel')
    && process.env.PMC_ALLOW_CLAUDE_EXPORT !== '1') {
    return { ok: false, reason, mode: m, detail: 'export Claude disabled' };
  }

  return { ok: true, reason, mode: m, detail: meta.detail || 'allowlisted last-resort' };
}

let _spendLog = [];

function logClaudeCall({ reason, ok, tokensEstimate = 0, detail = '' }) {
  const entry = {
    at: new Date().toISOString(),
    reason,
    ok: !!ok,
    tokensEstimate,
    detail,
    mode: mode(),
  };
  _spendLog.push(entry);
  if (_spendLog.length > 200) _spendLog = _spendLog.slice(-100);
  const tag = ok ? 'SPEND' : 'SKIP';
  console.log(`[claude-policy] ${tag} reason=${reason} mode=${entry.mode} est_tokens=${tokensEstimate} ${detail || ''}`);
  return entry;
}

function getSpendSummary() {
  const spent = _spendLog.filter(e => e.ok);
  const skipped = _spendLog.filter(e => !e.ok);
  return {
    mode: mode(),
    key_set: hasKey(),
    allow_vision: process.env.PMC_ALLOW_CLAUDE_VISION !== '0',
    allow_export: process.env.PMC_ALLOW_CLAUDE_EXPORT === '1',
    allow_dxf: process.env.PMC_ALLOW_CLAUDE_DXF === '1',
    calls_allowed: spent.length,
    calls_skipped: skipped.length,
    recent: _spendLog.slice(-20),
  };
}

/**
 * Wrap callClaudeAPI — refuse if policy blocks.
 * Usage: const api = gatedClaude(callClaudeAPI, 'vision_takeoff_local_weak');
 */
function gatedClaude(callClaudeAPI, reason, meta = {}) {
  const gate = allowClaude(reason, meta);
  if (!gate.ok) {
    logClaudeCall({ reason, ok: false, detail: gate.detail });
    return null;
  }
  return async (opts) => {
    logClaudeCall({
      reason,
      ok: true,
      tokensEstimate: opts?.maxTokens || 0,
      detail: meta.detail || '',
    });
    return callClaudeAPI(opts);
  };
}

module.exports = {
  mode,
  hasKey,
  allowClaude,
  logClaudeCall,
  getSpendSummary,
  gatedClaude,
  ALLOWED_STRICT,
};
