'use strict';
/**
 * Annotated takeoff PDF — stamp schedule/qty summary onto original PDF (or create summary PDF).
 */
const fs = require('fs');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

function wrapLines(text, maxChars) {
  const words = String(text || '').split(/\s+/);
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

/**
 * @param {object} opts
 * @param {Buffer|Uint8Array} [opts.pdfBytes] original drawing PDF
 * @param {string} opts.markdown takeoff report
 * @param {object} [opts.extracted] schedules
 * @param {string} [opts.title]
 */
async function buildAnnotatedTakeoffPdf(opts = {}) {
  const title = opts.title || 'PMC Quantity Takeoff';
  const md = String(opts.markdown || '').slice(0, 12000);
  const footings = opts.extracted?.schedules?.footings || [];
  const columns = opts.extracted?.schedules?.columns || [];
  const measureItems = opts.measure?.items || [];

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

  // Cover / annotation page at front
  const page = doc.insertPage(0, [595.28, 841.89]); // A4
  const margin = 40;
  let y = 800;
  const draw = (text, size = 10, bold = false, color = rgb(0.05, 0.12, 0.23)) => {
    const f = bold ? fontBold : font;
    const lines = wrapLines(text, 88);
    for (const line of lines) {
      if (y < 50) return;
      page.drawText(line, { x: margin, y, size, font: f, color });
      y -= size + 4;
    }
  };

  page.drawRectangle({
    x: 0, y: 780, width: 595.28, height: 62,
    color: rgb(0.04, 0.12, 0.23),
  });
  page.drawText('PMC CIVIL AI — QUANTITY TAKEOFF', {
    x: margin, y: 805, size: 14, font: fontBold, color: rgb(1, 1, 1),
  });
  page.drawText(stampedOnOriginal ? 'Annotated pack (summary + original sheet)' : 'Takeoff summary PDF', {
    x: margin, y: 788, size: 9, font, color: rgb(0.75, 0.82, 0.92),
  });
  y = 760;

  draw(title, 12, true);
  draw(`Generated: ${new Date().toLocaleString('en-IN')}`, 9, false, rgb(0.4, 0.45, 0.5));
  y -= 8;

  if (footings.length) {
    draw('FOOTING SCHEDULE (machine-read)', 11, true, rgb(0.1, 0.45, 0.9));
    for (const f of footings.slice(0, 12)) {
      draw(
        `${f.mark}: ${f.rcc_size_mm || f.pcc_size_mm || '?'}  D=${f.depth_mm || '?'}mm  Qty=${f.qty ?? '—'}`,
        9
      );
    }
    y -= 6;
  }
  if (columns.length) {
    draw('COLUMN / PEDESTAL SCHEDULE', 11, true, rgb(0.1, 0.45, 0.9));
    for (const c of columns.slice(0, 10)) {
      draw(`${c.mark}: ${c.size_mm || '?'}  Qty=${c.qty ?? '—'}`, 9);
    }
    y -= 6;
  }
  if (measureItems.length) {
    draw('PLAN MEASURE', 11, true, rgb(0.1, 0.45, 0.9));
    for (const m of measureItems.slice(0, 12)) {
      draw(`${m.type}: ${m.description} = ${m.qty} ${m.unit}`, 9);
    }
    y -= 6;
  }

  draw('REPORT EXTRACT', 11, true, rgb(0.1, 0.45, 0.9));
  const plain = md
    .replace(/[#|*`>_]/g, ' ')
    .replace(/\|/g, ' ')
    .replace(/\n{2,}/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean)
    .slice(0, 45);
  for (const line of plain) draw(line, 8, false, rgb(0.2, 0.25, 0.3));

  page.drawText('DRAFT — confirm missing qty/PCC before FINAL. Not a substitute for engineer check.', {
    x: margin, y: 28, size: 7, font, color: rgb(0.55, 0.35, 0.1),
  });

  // Small stamp on original page 1 (now page index 1)
  if (stampedOnOriginal && doc.getPageCount() > 1) {
    try {
      const orig = doc.getPage(1);
      const { width, height } = orig.getSize();
      orig.drawRectangle({
        x: width - 210, y: height - 36, width: 200, height: 28,
        color: rgb(0.04, 0.12, 0.23), opacity: 0.88,
      });
      orig.drawText('PMC TAKEOFF · SEE COVER PAGE', {
        x: width - 200, y: height - 24, size: 8, font: fontBold, color: rgb(1, 1, 1),
      });
    } catch (_) { /* ignore stamp failure */ }
  }

  const bytes = await doc.save();
  return {
    bytes: Buffer.from(bytes),
    stampedOnOriginal,
    pages: doc.getPageCount(),
  };
}

async function buildAnnotatedTakeoffPdfFromPath(pdfPath, opts = {}) {
  let pdfBytes = null;
  if (pdfPath && fs.existsSync(pdfPath)) pdfBytes = fs.readFileSync(pdfPath);
  return buildAnnotatedTakeoffPdf({ ...opts, pdfBytes });
}

module.exports = {
  buildAnnotatedTakeoffPdf,
  buildAnnotatedTakeoffPdfFromPath,
};
