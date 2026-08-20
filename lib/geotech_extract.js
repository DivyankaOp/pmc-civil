'use strict';
/**
 * Geotech / borehole extract from PDF OCR text (deeper v2).
 */

function extractGeotech(text) {
  const t = String(text || '');
  const boreholes = [];

  const starts = [...t.matchAll(/\b(?:BH|Bore\s*hole|Borehole)\s*[-–]?\s*(\d{1,3}[A-Z]?)\b/gi)];
  for (let i = 0; i < starts.length; i++) {
    const mark = `BH-${starts[i][1]}`.toUpperCase().replace('BH-BH-', 'BH-');
    const from = starts[i].index;
    const to = i + 1 < starts.length ? starts[i + 1].index : Math.min(t.length, from + 1600);
    const block = t.slice(from, to);

    const gl = block.match(/(?:G\.?L\.?|ground\s*level|EGL)\s*[=:]?\s*([+-]?\d+(?:\.\d+)?)\s*m?/i);
    const wl = block.match(/(?:W\.?T\.?|water\s*(?:table|level)|GWT)\s*[=:]?\s*([+-]?\d+(?:\.\d+)?)\s*m?/i);
    const spt = [...block.matchAll(/SPT\s*[=:]?\s*(\d{1,3})|N\s*=\s*(\d{1,3})/gi)]
      .map(m => Number(m[1] || m[2]))
      .filter(n => n > 0 && n < 200)
      .slice(0, 30);
    // Depth-tagged SPT: 1.5m N=12
    const sptDepths = [...block.matchAll(/(\d+(?:\.\d+)?)\s*m[^\n]{0,20}?(?:SPT|N)\s*[=:]?\s*(\d{1,3})/gi)]
      .map(m => ({ depth_m: Number(m[1]), n: Number(m[2]) }))
      .filter(x => x.n > 0 && x.n < 200)
      .slice(0, 30);
    const coords = block.match(/(?:Easting|E)\s*[=:]?\s*([\d.]+).{0,60}?(?:Northing|N)\s*[=:]?\s*([\d.]+)/i)
      || block.match(/([\d.]{5,})\s*[,]\s*([\d.]{5,})/);
    const strataHits = [...block.matchAll(
      /(\d+(?:\.\d+)?)\s*(?:-|to|–)\s*(\d+(?:\.\d+)?)\s*m[^\n]{0,40}?\b(clay|silty\s*clay|sand|silty\s*sand|gravel|rock|murum|fill|black\s*cotton|silt|laterite)\b/gi
    )];
    let strataLayers = strataHits.map(m => ({
      top_m: Number(m[1]),
      base_m: Number(m[2]),
      desc: m[3].toLowerCase(),
    }));
    const strata = [...block.matchAll(/\b(clay|silty\s*clay|sand|silty\s*sand|gravel|rock|murum|fill|black\s*cotton|silt|laterite)\b/gi)]
      .map(m => m[1].toLowerCase());
    const uniqStrata = [...new Set(strata)].slice(0, 10);
    if (!strataLayers.length && uniqStrata.length) {
      strataLayers = uniqStrata.map((s, idx) => ({
        top_m: idx * 1.5,
        base_m: idx * 1.5 + 1.5,
        desc: s,
      }));
    }
    const depth = block.match(/(?:depth|terminated|borehole\s*depth)\s*[=:]?\s*(\d+(?:\.\d+)?)\s*m/i);
    const cpt = [...block.matchAll(/CPT\s*[=:]?\s*(\d+(?:\.\d+)?)/gi)].map(m => Number(m[1])).slice(0, 10);

    boreholes.push({
      mark,
      ground_level_m: gl ? Number(gl[1]) : null,
      water_level_m: wl ? Number(wl[1]) : null,
      final_depth_m: depth ? Number(depth[1]) : null,
      spt_n: spt,
      spt_depths: sptDepths,
      avg_spt: spt.length ? Math.round(spt.reduce((a, b) => a + b, 0) / spt.length) : null,
      cpt,
      easting: coords ? coords[1] : null,
      northing: coords ? coords[2] : null,
      strata: uniqStrata,
      strata_layers: strataLayers,
      source: 'drawing-ocr-geotech',
      confidence: (gl || wl || spt.length || strataLayers.length) ? 'medium' : 'low',
      raw: block.slice(0, 220).replace(/\s+/g, ' '),
    });
  }

  const meta = {
    soil_bearing: t.match(/SBC\s*[=:]?\s*(\d+(?:\.\d+)?)\s*(?:t\/m|T\/m|kN)/i)?.[0] || null,
    foundation_note: /foundation|footing/i.test(t) && /soil|bearing/i.test(t)
      ? (t.match(/[^\n.]{0,40}soil bearing[^\n.]{0,60}/i)?.[0] || null)
      : null,
    report_ref: t.match(/(?:geo.?tech|soil)\s*report[^\n]{0,40}/i)?.[0] || null,
  };

  return {
    boreholes,
    meta,
    quality: boreholes.length >= 1
      ? (boreholes.some(b => b.spt_n.length || b.ground_level_m != null || b.strata_layers?.length) ? 'medium' : 'weak')
      : 'poor',
    note: 'Geotech v2 from text/OCR (+ strata layers, SPT depths). Not a substitute for Geotech engineer / full AGS digitiser.',
  };
}

function formatGeotechMarkdown(geo) {
  if (!geo) return '';
  const parts = ['### Geotech / borehole extract', `Quality: **${geo.quality}**`];
  if (geo.meta?.soil_bearing) parts.push(`- ${geo.meta.soil_bearing}`);
  if (geo.meta?.foundation_note) parts.push(`- Note: ${geo.meta.foundation_note}`);
  if (geo.meta?.report_ref) parts.push(`- ${geo.meta.report_ref}`);
  if (!geo.boreholes?.length) {
    parts.push('_No borehole marks (BH-1…) found in extract._');
  } else {
    parts.push('');
    parts.push('| Mark | GL (m) | WL (m) | Depth | Avg SPT | Strata |');
    parts.push('|---|---:|---:|---:|---:|---|');
    for (const b of geo.boreholes) {
      parts.push(`| ${b.mark} | ${b.ground_level_m ?? '—'} | ${b.water_level_m ?? '—'} | ${b.final_depth_m ?? '—'} | ${b.avg_spt ?? '—'} | ${(b.strata || []).join(', ') || '—'} |`);
    }
  }
  parts.push('', `_${geo.note}_`);
  return parts.join('\n');
}

module.exports = { extractGeotech, formatGeotechMarkdown };
