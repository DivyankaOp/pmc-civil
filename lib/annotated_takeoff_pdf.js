'use strict';
/**
 * Annotated takeoff PDF — cover summary + markup stamps on every drawing page.
 */
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

function sanitizePdfText(text) {
  return String(text || '')
    .replace(/[→←↔⇒⇐]/g, '->')
    .replace(/[×✕✖]/g, 'x')
    .replace(/[–—]/g, '-')
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[^\x00-\x7E]/g, '?');
}

function wrapLines(text, maxChars) {
  const words = sanitizePdfText(text).split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    if ((cur + ' ' + w).trim().length > maxChars) {
      if (cur) lines.push(cur);
      cur = w;
    } else cur = (cur + ' ' + w).trim();
  }
  if (cur) lines.push(cur);
  return lines;
}

function collectCallouts(opts) {
  const out = [];
  const footings = opts.extracted?.schedules?.footings || [];
  const columns = opts.extracted?.schedules?.columns || [];
  const measureItems = opts.measure?.items || [];
  const boreholes = opts.geotech?.boreholes || [];
  const ewItems = opts.earthworks?.items || [];

  for (const f of footings.slice(0, 8)) {
    out.push(sanitizePdfText(`FTG ${f.mark}: ${f.rcc_size_mm || f.pcc_size_mm || '?'} D${f.depth_mm || '?'} Q${f.qty ?? '?'}`));
  }
  for (const c of columns.slice(0, 6)) {
    out.push(sanitizePdfText(`COL ${c.mark}: ${c.size_mm || '?'} Q${c.qty ?? '?'}`));
  }
  for (const m of measureItems.filter(i => i.type === 'area' || i.type === 'length').slice(0, 6)) {
    out.push(sanitizePdfText(`${m.type.toUpperCase()}: ${m.qty} ${m.unit}`));
  }
  for (const e of ewItems.filter(i => i.type === 'cut' || i.type === 'fill').slice(0, 4)) {
    out.push(sanitizePdfText(`${e.type.toUpperCase()}: ${e.qty} ${e.unit}`));
  }
  for (const b of boreholes.slice(0, 4)) {
    out.push(sanitizePdfText(`${b.mark} GL=${b.ground_level_m ?? '?'} SPT=${b.avg_spt ?? '?'}`));
  }
  return out;
}

async function buildAnnotatedTakeoffPdf(opts = {}) {
  const title = opts.title || 'PMC Quantity Takeoff';
  const md = String(opts.markdown || '').slice(0, 12000);
  const footings = opts.extracted?.schedules?.footings || [];
  const columns = opts.extracted?.schedules?.columns || [];
  const measureItems = opts.measure?.items || [];
  const boreholes = opts.geotech?.boreholes || [];
  const ewItems = opts.earthworks?.items || [];
  const callouts = collectCallouts(opts);

  let doc;
  let stampedOnOriginal = false;
  if (opts.pdfBytes && opts.pdfBytes.length > 100) {
    try {
      doc = await PDFDocument.load(opts.pdfBytes, { ignoreEncryption: true });
      stampedOnOriginal = true;
    } catch (_) {
      doc = await PDFDocument.create();
    }
  } else {
    doc = await PDFDocument.create();
  }

  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);

  const page = doc.insertPage(0, [595.28, 841.89]);
  const margin = 40;
  let y = 800;
  const draw = (text, size = 10, bold = false, color = rgb(0.05, 0.12, 0.23)) => {
    const f = bold ? fontBold : font;
    for (const line of wrapLines(text, 88)) {
      if (y < 50) return;
      page.drawText(line, { x: margin, y, size, font: f, color });
      y -= size + 4;
    }
  };

  page.drawRectangle({ x: 0, y: 780, width: 595.28, height: 62, color: rgb(0.04, 0.12, 0.23) });
  page.drawText('PMC CIVIL AI — ANNOTATED TAKEOFF', {
    x: margin, y: 805, size: 14, font: fontBold, color: rgb(1, 1, 1),
  });
  page.drawText(stampedOnOriginal ? 'Summary + marked-up original sheet(s)' : 'Takeoff summary PDF', {
    x: margin, y: 788, size: 9, font, color: rgb(0.75, 0.82, 0.92),
  });
  y = 760;
  draw(title, 12, true);
  draw(`Generated: ${new Date().toLocaleString('en-IN')}`, 9, false, rgb(0.4, 0.45, 0.5));
  y -= 4;
  draw('LEGEND: FTG=footing  COL=column  AREA/LEN=plan measure  CUT/FILL=earthworks  BH=borehole', 7, false, rgb(0.35, 0.4, 0.45));
  // (WinAnsi-safe ASCII only in draw())
  y -= 6;

  if (footings.length) {
    draw('FOOTINGS', 11, true, rgb(0.1, 0.45, 0.9));
    for (const f of footings.slice(0, 12)) {
      draw(`${f.mark}: ${f.rcc_size_mm || f.pcc_size_mm || '?'}  D=${f.depth_mm || '?'}  Qty=${f.qty ?? '—'}`, 9);
    }
    y -= 4;
  }
  if (columns.length) {
    draw('COLUMNS / PEDESTALS', 11, true, rgb(0.1, 0.45, 0.9));
    for (const c of columns.slice(0, 10)) draw(`${c.mark}: ${c.size_mm || '?'}  Qty=${c.qty ?? '—'}`, 9);
    y -= 4;
  }
  if (measureItems.length) {
    draw('PLAN MEASURE', 11, true, rgb(0.1, 0.45, 0.9));
    for (const m of measureItems.slice(0, 12)) draw(`${m.type}: ${m.description} = ${m.qty} ${m.unit}`, 9);
    y -= 4;
  }
  if (ewItems.length) {
    draw('EARTHWORKS', 11, true, rgb(0.1, 0.45, 0.9));
    if (opts.earthworks?.ngl_m != null || opts.earthworks?.formation_m != null) {
      draw(`NGL=${opts.earthworks.ngl_m ?? '—'}  FGL=${opts.earthworks.formation_m ?? '—'}  Area=${opts.earthworks.area_sqm ?? '—'} sqm`, 9);
    }
    for (const e of ewItems.slice(0, 8)) {
      draw(`${e.type}: ${e.description} = ${e.qty ?? '—'} ${e.unit}`, 9);
    }
    y -= 4;
  }
  if (boreholes.length) {
    draw('GEOTECH / BOREHOLES', 11, true, rgb(0.1, 0.45, 0.9));
    for (const b of boreholes.slice(0, 8)) {
      draw(`${b.mark}: GL=${b.ground_level_m ?? '—'} WL=${b.water_level_m ?? '—'} SPT=${b.avg_spt ?? '—'}`, 9);
    }
    y -= 4;
  }

  draw('REPORT EXTRACT', 11, true, rgb(0.1, 0.45, 0.9));
  const plain = md.replace(/[#|*`>_]/g, ' ').replace(/\|/g, ' ').split('\n').map(l => l.trim()).filter(Boolean).slice(0, 36);
  for (const line of plain) draw(line, 8, false, rgb(0.2, 0.25, 0.3));
  page.drawText('DRAFT — confirm qty/PCC. Not a substitute for engineer / geotech check.', {
    x: margin, y: 28, size: 7, font, color: rgb(0.55, 0.35, 0.1),
  });

  if (stampedOnOriginal) {
    const n = doc.getPageCount();
    for (let i = 1; i < n; i++) {
      try {
        const p = doc.getPage(i);
        const { width, height } = p.getSize();
        p.drawRectangle({
          x: 0, y: height - 22, width, height: 22,
          color: rgb(0.04, 0.12, 0.23), opacity: 0.9,
        });
        p.drawText(`PMC TAKEOFF · p.${i}/${n - 1} · ${title.slice(0, 36)} · cover`, {
          x: 12, y: height - 15, size: 8, font: fontBold, color: rgb(1, 1, 1),
        });
        let cy = height - 48;
        const boxW = Math.min(210, width * 0.32);
        const boxX = width - boxW - 10;
        const boxH = Math.min(callouts.length * 14 + 22, height - 60);
        p.drawRectangle({
          x: boxX, y: Math.max(40, cy - callouts.length * 14 - 16),
          width: boxW, height: boxH,
          color: rgb(1, 1, 0.92), opacity: 0.92,
          borderColor: rgb(0.85, 0.65, 0.1),
          borderWidth: 1,
        });
        p.drawText('MARKED QUANTITIES', {
          x: boxX + 6, y: cy, size: 7, font: fontBold, color: rgb(0.45, 0.3, 0.05),
        });
        cy -= 12;
        for (const c of callouts.slice(0, 14)) {
          p.drawText(c.slice(0, 42), {
            x: boxX + 6, y: cy, size: 6.5, font, color: rgb(0.1, 0.15, 0.2),
          });
          cy -= 12;
          if (cy < 50) break;
        }
        p.drawText(`Sheet ${i} of ${n - 1}`, {
          x: 12, y: 14, size: 7, font, color: rgb(0.3, 0.35, 0.4),
        });
      } catch (_) { /* skip page */ }
    }
  }

  const bytes = await doc.save();
  return {
    bytes: Buffer.from(bytes),
    stampedOnOriginal,
    pages: doc.getPageCount(),
    callouts: callouts.length,
  };
}

module.exports = {
  buildAnnotatedTakeoffPdf,
};
