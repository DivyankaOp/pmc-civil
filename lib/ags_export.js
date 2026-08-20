'use strict';
/**
 * AGS 4.1 export v2 — PROJ, HOLE, GEOL, ISPT, WSTG, SAMP, ABBR
 */

function esc(v) {
  if (v == null || v === '') return '""';
  const s = String(v).replace(/"/g, '""');
  return `"${s}"`;
}

function buildAgs41(geotech = {}, opts = {}) {
  const project = opts.project || opts.title || 'PMC Geotech Export';
  const holes = geotech.boreholes || [];
  const lines = [];

  lines.push('**PROJ');
  lines.push('*PROJ_ID,PROJ_NAME,PROJ_LOC,PROJ_CLNT,PROJ_CONT,PROJ_ENG');
  lines.push('"UNIT","","","","",""');
  lines.push('"TYPE","","","","",""');
  lines.push([esc('PMC1'), esc(project), esc(geotech.meta?.report_ref || ''), esc(''), esc('PMC Civil AI'), esc('')].join(','));

  lines.push('**ABBR');
  lines.push('*ABBR_HDNG,ABBR_DESC');
  lines.push('"UNIT","",""');
  lines.push('"TYPE","X","X"');
  lines.push('"CP","Cable Percussion / unknown method"');
  lines.push('"PMC","PMC Civil AI OCR extract"');

  lines.push('**HOLE');
  lines.push('*HOLE_ID,HOLE_TYPE,HOLE_NATE,HOLE_NATN,HOLE_GL,HOLE_FDEP,HOLE_STAR,HOLE_LOG');
  lines.push('"UNIT","","m","m","m","m","yyyy-mm-dd",""');
  lines.push('"TYPE","ID","2DP","2DP","2DP","2DP","DT","X"');
  for (const b of holes) {
    const id = (b.mark || 'BH').replace(/\s+/g, '');
    lines.push([
      esc(id),
      esc('CP'),
      b.easting != null ? esc(b.easting) : '""',
      b.northing != null ? esc(b.northing) : '""',
      b.ground_level_m != null ? esc(b.ground_level_m) : '""',
      b.final_depth_m != null ? esc(b.final_depth_m) : '""',
      '""',
      esc(b.source || 'PMC-OCR'),
    ].join(','));
  }

  lines.push('**GEOL');
  lines.push('*HOLE_ID,GEOL_TOP,GEOL_BASE,GEOL_DESC,GEOL_LEG,GEOL_GEOL');
  lines.push('"UNIT","m","m","","",""');
  lines.push('"TYPE","ID","2DP","2DP","X","PA","PA"');
  for (const b of holes) {
    const id = (b.mark || 'BH').replace(/\s+/g, '');
    const layers = b.strata_layers?.length
      ? b.strata_layers
      : (b.strata?.length ? b.strata.map((s, i) => ({ top_m: i * 1.5, base_m: i * 1.5 + 1.5, desc: s })) : [{ top_m: 0, base_m: 1.5, desc: 'unknown' }]);
    for (const s of layers) {
      const code = String(s.desc || 'UNK').slice(0, 4).toUpperCase();
      lines.push([
        esc(id),
        esc(Number(s.top_m).toFixed(2)),
        esc(Number(s.base_m).toFixed(2)),
        esc(s.desc),
        esc(code),
        esc(code),
      ].join(','));
    }
  }

  lines.push('**ISPT');
  lines.push('*HOLE_ID,ISPT_TOP,ISPT_NVAL,ISPT_REP');
  lines.push('"UNIT","m","",""');
  lines.push('"TYPE","ID","2DP","2DP","X"');
  for (const b of holes) {
    const id = (b.mark || 'BH').replace(/\s+/g, '');
    if (b.spt_depths?.length) {
      for (const s of b.spt_depths) {
        lines.push([esc(id), esc(s.depth_m.toFixed(2)), esc(s.n), esc('PMC')].join(','));
      }
    } else {
      const spts = b.spt_n?.length ? b.spt_n : (b.avg_spt != null ? [b.avg_spt] : []);
      spts.forEach((n, i) => {
        lines.push([esc(id), esc((i * 1.5).toFixed(2)), esc(n), esc('PMC')].join(','));
      });
    }
  }

  lines.push('**WSTG');
  lines.push('*HOLE_ID,WSTG_DATE,WSTG_DEPTH,WSTG_CAS');
  lines.push('"UNIT","","m","m"');
  lines.push('"TYPE","ID","DT","2DP","2DP"');
  for (const b of holes) {
    if (b.water_level_m == null) continue;
    const id = (b.mark || 'BH').replace(/\s+/g, '');
    let depth = b.water_level_m;
    if (b.ground_level_m != null && b.water_level_m < b.ground_level_m) {
      depth = Math.round((b.ground_level_m - b.water_level_m) * 100) / 100;
    }
    lines.push([esc(id), '""', esc(depth), '""'].join(','));
  }

  lines.push('**SAMP');
  lines.push('*HOLE_ID,SAMP_TOP,SAMP_BASE,SAMP_TYPE,SAMP_REF');
  lines.push('"UNIT","m","m","",""');
  lines.push('"TYPE","ID","2DP","2DP","PA","X"');
  for (const b of holes) {
    const id = (b.mark || 'BH').replace(/\s+/g, '');
    const layers = (b.strata_layers || []).slice(0, 4);
    layers.forEach((s, i) => {
      lines.push([
        esc(id),
        esc(Number(s.top_m).toFixed(2)),
        esc(Number(s.base_m).toFixed(2)),
        esc('D'),
        esc(`${id}-S${i + 1}`),
      ].join(','));
    });
  }

  lines.push('**FILE');
  lines.push('*FILE_DESC,FILE_TYPE');
  lines.push('"UNIT","",""');
  lines.push('"TYPE","X","X"');
  lines.push([esc('PMC Civil AI geotech v2'), esc('AGS4')].join(','));

  const body = lines.join('\r\n') + '\r\n';
  return {
    text: body,
    bytes: Buffer.from(body, 'utf8'),
    holes: holes.length,
    groups: ['PROJ', 'ABBR', 'HOLE', 'GEOL', 'ISPT', 'WSTG', 'SAMP', 'FILE'],
    note: 'AGS 4.1 v2 — OCR-derived. Verify before OpenGround/Leapfrog import.',
  };
}

module.exports = { buildAgs41 };
