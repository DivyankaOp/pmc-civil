'use strict';
/**
 * Human QA gate on EVERY takeoff (Civils human-in-the-loop lite).
 * Auto checks + mandatory engineer sign-off before export.
 */

function buildQaChecklist({ extracted, measure, geotech, earthworks, paving, markdown } = {}) {
  const items = [];
  const ft = extracted?.schedules?.footings || [];
  const cols = extracted?.schedules?.columns || [];
  const missingQty = ft.filter(f => f.qty == null || f.qty === '').length;
  const missingDepth = ft.filter(f => !f.depth_mm).length;
  const marks = ft.map(f => f.mark).filter(Boolean);
  const dupMarks = marks.filter((m, i) => marks.indexOf(m) !== i);

  // Consistency: huge unit volumes
  const wildVol = ft.some(f => {
    const size = String(f.rcc_size_mm || f.pcc_size_mm || '');
    const m = size.match(/(\d+)\s*[xX×]\s*(\d+)/);
    const d = Number(f.depth_mm);
    if (!m || !d) return false;
    const each = (Number(m[1]) / 1000) * (Number(m[2]) / 1000) * (d / 1000);
    return each > 50 || each < 0.01;
  });

  items.push({
    id: 'extract_present',
    label: 'Something was extracted from the drawing',
    ok: (ft.length + cols.length) > 0 || (measure?.items?.length || 0) > 0
      || (geotech?.boreholes?.length || 0) > 0 || (earthworks?.items?.length || 0) > 0
      || (paving?.items?.length || 0) > 0,
    detail: `FTG ${ft.length} · COL ${cols.length} · measure ${(measure?.items || []).length} · BH ${(geotech?.boreholes || []).length}`,
    severity: 'block',
  });
  items.push({
    id: 'footing_qty',
    label: 'Footing quantities confirmed (no blanks)',
    ok: ft.length === 0 || missingQty === 0,
    detail: missingQty ? `${missingQty} footing(s) missing qty — ASK USER` : 'OK / N/A',
    severity: ft.length ? 'block' : 'info',
  });
  items.push({
    id: 'footing_depth',
    label: 'Footing depths present',
    ok: ft.length === 0 || missingDepth === 0,
    detail: missingDepth ? `${missingDepth} missing depth` : 'OK / N/A',
    severity: 'warn',
  });
  items.push({
    id: 'no_dup_marks',
    label: 'No duplicate footing marks',
    ok: dupMarks.length === 0,
    detail: dupMarks.length ? `Duplicates: ${[...new Set(dupMarks)].join(', ')}` : 'OK',
    severity: 'warn',
  });
  items.push({
    id: 'volume_sanity',
    label: 'Unit volumes look sane (not wild)',
    ok: !wildVol,
    detail: wildVol ? 'Check size/depth — volume outlier detected' : 'OK / N/A',
    severity: 'warn',
  });
  items.push({
    id: 'measure_scale',
    label: 'Plan / paving measure scale checked',
    ok: !measure?.items?.length || !!measure.scale
      || (measure.items || []).some(i => /click-measure|auto-vision/i.test(i.source || '')),
    detail: measure?.scale ? measure.scale.ratio : (measure?.items?.length ? 'Scale not found — confirm' : 'N/A'),
    severity: 'warn',
  });
  items.push({
    id: 'geotech_bh',
    label: 'Geotech boreholes reviewed',
    ok: !geotech?.boreholes?.length || geotech.quality !== 'poor',
    detail: geotech?.boreholes?.length
      ? `${geotech.boreholes.length} BH · quality ${geotech.quality} · factory ${geotech.factory_score ?? '—'}`
      : 'N/A',
    severity: 'warn',
  });
  items.push({
    id: 'earthworks_levels',
    label: 'Earthworks / paving levels & thicknesses checked',
    ok: !earthworks || ((earthworks.ngl_m != null && earthworks.formation_m != null) || !(earthworks.questions?.length))
      || (paving?.items?.length > 0),
    detail: earthworks
      ? `NGL ${earthworks.ngl_m ?? '—'} / FGL ${earthworks.formation_m ?? '—'} · paving ${(paving?.items || []).length}`
      : (paving?.items?.length ? `${paving.items.length} paving lines` : 'N/A'),
    severity: 'warn',
  });
  items.push({
    id: 'no_invent',
    label: 'Gaps flagged — nothing silently invented',
    ok: true,
    detail: /ASK USER|Need from you|not found/i.test(String(markdown || ''))
      ? 'Open gaps still listed — close before FINAL'
      : 'No open invent flags in report (still verify numbers)',
    severity: 'info',
    soft: true,
  });
  items.push({
    id: 'second_pass',
    label: 'Second-pass review (re-read schedule vs quantities)',
    ok: false,
    detail: 'Reviewer must tick after comparing drawing to draft',
    requiresConfirm: true,
    severity: 'block',
  });
  items.push({
    id: 'engineer_signoff',
    label: 'Engineer / QS human sign-off',
    ok: false,
    detail: 'Required before Excel / AGS / annotated / SHP export',
    requiresConfirm: true,
    severity: 'block',
  });

  const autoItems = items.filter(i => !i.requiresConfirm);
  const autoOk = autoItems.filter(i => i.ok).length;
  const blockers = autoItems.filter(i => !i.ok && i.severity === 'block');
  return {
    items,
    score: `${autoOk}/${autoItems.length}`,
    blockers: blockers.map(b => b.id),
    ready_for_export: blockers.length === 0,
    requires_human: true,
    note: 'HUMAN QA GATE — every takeoff. You are the reviewer (Civils-style HITL lite). Exports locked until sign-off.',
  };
}

/**
 * Validate a QA approval payload from the UI.
 */
function approveQaSession(checklist, opts = {}) {
  const reviewer = String(opts.reviewer || '').trim();
  const confirmed = new Set(opts.confirmedIds || []);
  const required = (checklist?.items || []).filter(i => i.requiresConfirm).map(i => i.id);
  const missing = required.filter(id => !confirmed.has(id));
  const blockers = (checklist?.items || []).filter(i => !i.requiresConfirm && !i.ok && i.severity === 'block');

  if (!reviewer || reviewer.length < 2) {
    return { ok: false, error: 'Reviewer name required (human QA)' };
  }
  if (missing.length) {
    return { ok: false, error: `Confirm required QA ticks: ${missing.join(', ')}` };
  }
  if (blockers.length && !opts.force) {
    return { ok: false, error: `Fix blockers first: ${blockers.map(b => b.id).join(', ')}` };
  }

  return {
    ok: true,
    approval: {
      approved: true,
      reviewer,
      confirmedIds: [...confirmed],
      at: new Date().toISOString(),
      score: checklist?.score,
      note: opts.note || 'Human QA approved for export',
    },
  };
}

function assertExportAllowed(approval) {
  if (!approval?.approved || !approval?.reviewer) {
    return { ok: false, error: 'Human QA sign-off required before export' };
  }
  // approvals older than 24h still OK for session; no expiry for v1
  return { ok: true };
}

module.exports = {
  buildQaChecklist,
  approveQaSession,
  assertExportAllowed,
};
