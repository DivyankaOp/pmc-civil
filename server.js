require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');
const ExcelJS = require('exceljs');
const { dataPath, scriptsPath, publicDir } = require('./paths');
const { extractDrawingData, buildDrawingExcel } = require('./lib/server_drawing');
const { geminiAnalyzeDrawing, runCVAnalysis, RATES } = require('./lib/drawing_analyzer');
const { parseDXF, extractCivilData, extractTotalAreaSqft, attachScheduleTables } = require('./lib/dxf_parser');
const { buildExcelFromDrawing, getDrawingPrompt } = require('./lib/drawing_to_excel');
const { buildDXFExcel } = require('./lib/dxf_to_excel');
const { analyzeDrawing, buildAIPrompt } = require('./lib/drawing_intelligence');
const { claudeAnalyzeDXF, claudeClassifySymbols, claudeAnalyzeWithAnswers, claudeAnalyzeDrawingVision, claudeAnalyzeDWGVision, callClaudeAPI, CIVIL_SYSTEM, parseJSON } = require('./lib/claude_analyzer');
const { learnRatesFromBOQ, learnRatesFromMarkdown, getRatesSummary, getRatesMap, getLearnedRateStats } = require('./lib/rate_store');
const { buildSmartContextFromAnalyzed, buildSmartContext } = require('./lib/smart_boq_engine');
const {
  runScheduleFirstLocal,
  polishWithClaude,
  combineExtractedText,
  applyUserClarifications,
  extractSchedules,
} = require('./lib/schedule_pipeline');
const { runCadZoomOcrFromBase64, runCadZoomOcrOnFile } = require('./lib/cad_local_reader');
const { buildSpatialScheduleText } = require('./lib/spatial_tables');
const multer = require('multer');
const fs = require('fs');
const os = require('os');

const app = express();
app.use(cors());
// Large drawings: JSON fallback up to 200MB; prefer multipart /analyze-drawing (any size up to 500MB)
app.use(express.json({ limit: '200mb' }));
app.use(express.static(publicDir()));

const uploadDrawing = multer({
  dest: path.join(os.tmpdir(), 'pmc_uploads'),
  limits: { fileSize: 500 * 1024 * 1024 }, // 500MB — Civils.ai-style: size not a hard product limit
});
try { fs.mkdirSync(path.join(os.tmpdir(), 'pmc_uploads'), { recursive: true }); } catch (_) {}

function pyExec() {
  return process.env.PMC_PYTHON || (process.platform === 'win32' ? 'py -3' : 'python3');
}

async function extractPdfTextFromPath(pdfPath) {
  const { execFileSync } = require('child_process');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pmc_pdf_'));
  try {
    const scriptPath = path.join(tmpDir, 'extract.py');
    fs.writeFileSync(scriptPath, `
import fitz, json, sys
doc = fitz.open(sys.argv[1])
pages = []
for page_num in range(len(doc)):
    page = doc[page_num]
    blocks = page.get_text("dict")["blocks"]
    texts = []
    for b in blocks:
        if b.get("type") == 0:
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    t = span.get("text","").strip()
                    if t:
                        x, y = span["origin"]
                        texts.append({"text": t, "x": round(x,2), "y": round(y,2), "size": round(span.get("size",10),1)})
    pages.append({"page": page_num+1, "texts": texts, "width": page.rect.width, "height": page.rect.height})
doc.close()
total = sum(len(p["texts"]) for p in pages)
print(json.dumps({"pages": pages, "is_vector": any(len(p["texts"])>10 for p in pages), "total_texts": total}))
`.trim());
    const { execSync } = require('child_process');
    const out = execSync(`${pyExec()} "${scriptPath}" "${pdfPath}"`, {
      timeout: 120000,
      maxBuffer: 40 * 1024 * 1024,
      windowsHide: true,
    });
    return JSON.parse(out.toString());
  } catch (e) {
    console.error('PDF text extract error:', e.message);
    return null;
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch (_) {}
  }
}

async function extractPdfText(pdfBase64) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pmc_pdf_'));
  const pdfPath = path.join(tmpDir, 'input.pdf');
  try {
    fs.writeFileSync(pdfPath, Buffer.from(pdfBase64, 'base64'));
    return await extractPdfTextFromPath(pdfPath);
  } catch (e) {
    console.error('PDF text extract error:', e.message);
    return null;
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch (_) {}
  }
}

async function extractLargePdfViaImageOCR(pdfBase64, gcvKey) {
  const { execSync } = require('child_process');
  const fs = require('fs');
  const os = require('os');
  const tmpDir = fs.mkdtempSync(os.tmpdir() + '/pmc_large_');
  const pdfPath = tmpDir + '/input.pdf';
  try {
    fs.writeFileSync(pdfPath, Buffer.from(pdfBase64, 'base64'));


    const script = `
import fitz, base64, json
doc = fitz.open('${pdfPath}')
tiles = []
for i in range(len(doc)):
    page = doc[i]
    pix = page.get_pixmap(matrix=fitz.Matrix(400/72, 400/72), alpha=False)
    tiles.append(base64.b64encode(pix.tobytes('png')).decode())
doc.close()
print(json.dumps(tiles))
`.trim();
    const sp = tmpDir + '/r.py';
    fs.writeFileSync(sp, script);
    const out = execSync(`${pyExec()} "${sp}"`, { timeout: 60000, maxBuffer: 100 * 1024 * 1024, env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' } });
    const tiles = JSON.parse(out.toString());
    console.log(`[GCV-Large] Rendered ${tiles.length} page tiles from large PDF`);

    // OCR each tile using images:annotate (no size restriction)
    const pages = [];
    for (let i = 0; i < tiles.length; i++) {
      try {
        const gcvRes = await fetch(`https://vision.googleapis.com/v1/images:annotate?key=${gcvKey}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ requests: [{ image: { content: tiles[i] }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }] }] }),
          signal: AbortSignal.timeout(30000)
        });
        if (!gcvRes.ok) continue;
        const data = await gcvRes.json();
        const text = data.responses?.[0]?.fullTextAnnotation?.text || '';
        if (text.trim()) {
          const lines = text.split('\n').filter(l => l.trim());
          pages.push({ table_rows: lines.map(l => [l]), raw_text: lines.join('\n'), is_rotated: false });
          console.log(`[GCV-Large] Page ${i+1}: ${text.length} chars`);
        }
      } catch(e) { console.error(`[GCV-Large] Page ${i+1} failed:`, e.message); }
    }
    return pages.length ? { pages, is_gcv: true } : null;
  } catch(e) {
    console.error('[GCV-Large] Error:', e.message);
    return null;
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch(e) {}
  }
}


async function extractScannedPdfWithGCV(pdfBase64) {
  const gcvKey = process.env.GOOGLE_CLOUD_VISION_API_KEY;

  // ── Path 1: GCV (if API key present) ──────────────────────────────
  if (gcvKey) {
    try {
      console.log('[OCR] GCV key found — trying Google Cloud Vision first');
      const gcvResult = await extractLargePdfViaImageOCR(pdfBase64, gcvKey);
      if (gcvResult?.pages?.length) {
        console.log(`[OCR] GCV success: ${gcvResult.pages.length} pages`);
        return gcvResult; // already has { pages, is_gcv: true }
      }
      console.warn('[OCR] GCV returned no data — falling back to Tesseract');
    } catch(gcvErr) {
      console.error('[OCR] GCV failed:', gcvErr.message, '— falling back to Tesseract');
    }
  } else {
    console.log('[OCR] No GOOGLE_CLOUD_VISION_API_KEY — using Tesseract pipeline');
  }

  const { execSync } = require('child_process');
  const fs = require('fs');
  const os = require('os');
  const path = require('path');
  const tmpDir = fs.mkdtempSync(os.tmpdir() + '/pmc_ocr_');
  try {
    const pdfPath = path.join(tmpDir, 'input.pdf');
    const outPath = path.join(tmpDir, 'ocr_out.json');
    fs.writeFileSync(pdfPath, Buffer.from(pdfBase64, 'base64'));

    const scriptPath = scriptsPath('ocr_pipeline.py');
    const py = pyExec();

    execSync(`${py} "${scriptPath}" "${pdfPath}" "${outPath}"`, {
      timeout: 180000,
      maxBuffer: 50 * 1024 * 1024,
      env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' }
    });

    if (!fs.existsSync(outPath)) return null;
    const result = JSON.parse(fs.readFileSync(outPath, 'utf8'));

    if (!result.pages?.length) return null;

    console.log(`[OCR Pipeline] ${result.pages.length} pages, ${result.total_chars} total chars`);

    return {
      pages: result.pages.map(p => ({
        table_rows: p.table_rows || [],
        raw_text:   p.raw_text   || '',
        is_rotated: p.is_rotated || false,
        crops_processed: p.crops_processed || 0
      })),
      is_gcv: false,
      engine: result.engine || 'tesseract-opencv-pipeline'
    };

  } catch(e) {
    console.error('[OCR Pipeline] Error:', e.message);
    // Last resort: simple single-pass Tesseract on full page
    try {
      const pdfPath2 = path.join(tmpDir, 'input2.pdf');
      if (!fs.existsSync(pdfPath2)) {
        fs.writeFileSync(pdfPath2, Buffer.from(pdfBase64, 'base64'));
      }
      const fallbackScript = `
import fitz,base64,subprocess,json,tempfile,os
doc=fitz.open('${tmpDir.replace(/\\/g,'/')}/input.pdf')
pages=[]
for i in range(len(doc)):
    pix=doc[i].get_pixmap(matrix=fitz.Matrix(300/72,300/72),alpha=False)
    tmp=tempfile.NamedTemporaryFile(suffix='.png',delete=False)
    pix.save(tmp.name); tmp.close()
    r=subprocess.run(['tesseract',tmp.name,'stdout','--oem','1','--psm','6','-l','eng'],capture_output=True,text=True,timeout=30)
    t=r.stdout.strip()
    if t: pages.append({'raw_text':t,'table_rows':[[l] for l in t.split('\\n') if l.strip()],'is_rotated':False})
    os.unlink(tmp.name)
doc.close()
print(json.dumps({'pages':pages}))
`.trim();
      const fbScript = path.join(tmpDir, 'fallback.py');
      fs.writeFileSync(fbScript, fallbackScript);
      const py2 = pyExec();
      const fbOut = execSync(`${py2} "${fbScript}"`, { timeout: 60000, maxBuffer: 10*1024*1024, env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' } });
      const fbData = JSON.parse(fbOut.toString());
      if (fbData.pages?.length) {
        console.log('[OCR Fallback] Used simple Tesseract PSM6 fallback');
        return { pages: fbData.pages, is_gcv: false, engine: 'tesseract-fallback' };
      }
    } catch(e2) { console.error('[OCR Fallback] also failed:', e2.message); }
    return null;
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch(e) {}
  }
}

async function pdfToImageTiles(pdfBase64) {
  const { execSync } = require('child_process');
  const fs = require('fs');
  const os = require('os');
  const tmpDir = fs.mkdtempSync(os.tmpdir() + '/pmc_pdf_');
  const pdfPath = tmpDir + '/input.pdf';
  try {
    fs.writeFileSync(pdfPath, Buffer.from(pdfBase64, 'base64'));

    // FIXED: this template string was previously missing its own declaration
    // (no "const script = `" before the python code), so it threw
    // "script is not defined" on EVERY call, was silently caught below, and
    // returned null. Net effect: Claude never received any image at all for
    // scanned/rasterized drawings — only broken OCR text.
    //
    // ALSO FIXED: previously only cropped a hardcoded bottom-right quadrant,
    // assuming that's where schedule tables live. Real drawings vary — this
    // project's column/footing schedule sits top-right. We now render an
    // overlapping 3x2 grid across the WHOLE sheet at high DPI, so whichever
    // corner the table is actually in, it gets a sharp legible close-up.
    // Blank tiles are skipped; every tile is capped to ~1568px (Claude's
    // optimal edge) so text stays crisp instead of being crushed by
    // whole-page downscaling.
    const script = `
import fitz, json, base64, io
from PIL import Image

MAX_EDGE = 1568
RENDER_DPI = 400

def encode_capped(pix):
    img = Image.open(io.BytesIO(pix.tobytes('png')))
    if max(img.size) > MAX_EDGE:
        scale = MAX_EDGE / max(img.size)
        img = img.resize((max(1, int(img.size[0]*scale)), max(1, int(img.size[1]*scale))), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='PNG', optimize=True)
    return base64.b64encode(buf.getvalue()).decode()

doc = fitz.open('${pdfPath}')
tiles = []
mat = fitz.Matrix(RENDER_DPI/72, RENDER_DPI/72)

for page_num in range(len(doc)):
    page = doc[page_num]
    w, h = page.rect.width, page.rect.height

    overview = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72), alpha=False)
    tiles.append({'label': f'page_{page_num+1}_overview', 'data': encode_capped(overview)})

    cols, rows = 3, 2
    overlap = 0.06
    for row in range(rows):
        for col in range(cols):
            x1 = max(0, (col / cols) - overlap) * w
            x2 = min(1, ((col + 1) / cols) + overlap) * w
            y1 = max(0, (row / rows) - overlap) * h
            y2 = min(1, ((row + 1) / rows) + overlap) * h
            pix = page.get_pixmap(matrix=mat, alpha=False, clip=fitz.Rect(x1, y1, x2, y2))

            samp = pix.samples
            step = max(1, len(samp) // 20000)
            dark = sum(1 for i in range(0, len(samp), step) if samp[i] < 200)
            if dark < 5:
                continue

            tiles.append({'label': f'page_{page_num+1}_r{row}c{col}', 'data': encode_capped(pix)})

doc.close()
print(json.dumps(tiles))
`.trim();
    const scriptPath = tmpDir + '/convert.py';
    fs.writeFileSync(scriptPath, script);
    const out = execSync(`${pyExec()} "${scriptPath}"`, { timeout: 120000, maxBuffer: 300 * 1024 * 1024, env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' } });
    const result = JSON.parse(out.toString());
    return result.map(t => typeof t === 'object' ? t.data : t);
  } catch(e) {
    console.error('PDF tile error:', e.message);
    return null;
  } finally {
    try { require('fs').rmSync(tmpDir, { recursive: true }); } catch(e) {}
  }
}

/**
 * Text-first drawing context (Civils.ai-style).
 * Default: NO image tiles (saves tokens).
 * Vision fallback: 1 overview + up to 2 crops only when schedule text is weak.
 */
function linesFromExtractedPdf(extracted) {
  const plainLines = [];
  let extractedTextBlock = '';
  if (extracted?.pages?.length) {
    for (const page of extracted.pages) {
      const byY = {};
      for (const t of (page.texts || [])) {
        const row = Math.round(t.y / 15) * 15;
        if (!byY[row]) byY[row] = [];
        byY[row].push(t);
      }
      for (const row of Object.keys(byY).sort((a, b) => Number(a) - Number(b))) {
        const line = byY[row].sort((a, b) => a.x - b.x).map(t => t.text).join('  ');
        if (line.trim()) plainLines.push(line);
      }
    }
    const totalTexts = extracted.total_texts || 0;
    extractedTextBlock = `=== MACHINE-EXTRACTED TEXT FROM DRAWING (${totalTexts} items) ===\n${plainLines.join('\n')}\n=== END EXTRACTED TEXT ===`;
    console.log(`[drawing-context] Vector PDF: ${totalTexts} texts extracted`);
  }
  return { plainLines, extractedTextBlock };
}

/**
 * Civils.ai-style: work from a file on disk (any size).
 * Spatial tables + schedule-quality OCR gate (not char-count).
 */
async function buildDrawingContextFromFile(filePath, opts = {}) {
  const forceVision = opts.forceVision === true;
  const parts = [];
  const ext = path.extname(filePath || '').toLowerCase();
  const isPdf = ext === '.pdf' || opts.mime === 'application/pdf';
  const isImage = ['.png', '.jpg', '.jpeg', '.webp', '.bmp', '.tif', '.tiff', '.gif'].includes(ext)
    || (opts.mime || '').startsWith('image/');

  let extracted = null;
  if (isPdf) {
    extracted = await extractPdfTextFromPath(filePath);
  }
  const { plainLines, extractedTextBlock } = linesFromExtractedPdf(extracted);

  // Spatial rebuild from PDF coords (column-aligned schedules)
  let spatial = buildSpatialScheduleText({
    pdfPages: extracted?.pages || [],
    plainLines,
  });
  let probe = extractSchedules(spatial.text || plainLines.join('\n'), {
    spatialTables: spatial.tables || [],
  });
  console.log(`[drawing-context] Vector spatial: quality=${probe.quality} rows=${probe.total_schedule_rows} tables=${spatial.tables?.length || 0}`);

  let gcvBlock = '';
  let cadZoomText = '';
  let cadHints = [];
  let ocrBoxes = [];
  const scheduleWeak = probe.quality === 'poor' || (probe.total_schedule_rows || 0) < 1;
  const vectorThin = plainLines.join('\n').length < 200;

  if (forceVision || isImage || scheduleWeak || vectorThin) {
    try {
      console.log('[drawing-context] Running CAD-zoom OCR (multi-page + adaptive schedule crops)...');
      const outDir = path.join(path.dirname(filePath), 'crops_' + Date.now());
      fs.mkdirSync(outDir, { recursive: true });
      const zoom = runCadZoomOcrOnFile(filePath, outDir);
      if (zoom?.success && zoom.full_text) {
        cadZoomText = zoom.full_text;
        cadHints = zoom.drawing_hints || [];
        ocrBoxes = zoom.boxes || [];
        plainLines.push(...String(zoom.full_text).split('\n'));
        console.log(`[drawing-context] CAD-zoom OCR: ${zoom.char_count || cadZoomText.length} chars | boxes=${ocrBoxes.length} | pages=${zoom.pages_processed || '?'} | hints=${cadHints.join(',')}`);
      } else {
        console.warn('[drawing-context] CAD-zoom OCR failed:', zoom?.error || 'empty');
      }
    } catch (e) {
      console.warn('[drawing-context] CAD-zoom OCR error:', e.message);
    }

    if ((cadZoomText.length < 200) && process.env.GOOGLE_CLOUD_VISION_API_KEY && isPdf) {
      try {
        const stat = fs.statSync(filePath);
        if (stat.size <= 40 * 1024 * 1024) {
          const pdfB64 = fs.readFileSync(filePath).toString('base64');
          const gcvResult = await extractScannedPdfWithGCV(pdfB64);
          if (gcvResult?.pages?.length) {
            const gcvLines = [];
            gcvBlock = gcvResult.pages.map((p, i) => {
              const rotNote = p.is_rotated ? ' [rotated]' : '';
              const pageText = p.raw_text || p.text || '';
              gcvLines.push(...String(pageText).split('\n'));
              return `=== PAGE ${i + 1}${rotNote} ===\n${pageText}`;
            }).join('\n\n');
            gcvBlock = `=== SCANNED PDF OCR (GCV) ===\n${gcvBlock}\n=== END OCR ===`;
            plainLines.push(...gcvLines);
            console.log(`[drawing-context] GCV fallback: ${gcvResult.pages.length} pages`);
          }
        } else {
          console.log('[drawing-context] Skipping GCV — file >40MB; local OCR only');
        }
      } catch (e) {
        console.warn('[drawing-context] GCV skip:', e.message);
      }
    }

    // Rebuild spatial with OCR boxes
    spatial = buildSpatialScheduleText({
      pdfPages: extracted?.pages || [],
      ocrBoxes,
      plainLines,
    });
  } else {
    console.log('[drawing-context] Skipping OCR — schedule quality already good from vector text');
  }

  const combinedText = combineExtractedText([
    spatial.text,
    plainLines.join('\n'),
    extractedTextBlock,
    gcvBlock,
    cadZoomText,
  ]);
  const userQuestion = opts.question || '';
  const local = runScheduleFirstLocal(combinedText, {
    filename: opts.filename || path.basename(filePath) || 'drawing.pdf',
    question: userQuestion,
    hints: cadHints,
    spatialTables: spatial.tables || [],
  });
  console.log(`[drawing-context] type=${local.typeInfo?.drawing_type} schedule=${local.extracted.quality} rows=${local.extracted.total_schedule_rows} localQA=${!!local.answeredLocally}`);

  const needVision = forceVision || (local.needsClaude && local.needsVision);
  let fileBytes = 0;
  try { fileBytes = fs.statSync(filePath).size; } catch (_) {}

  if (needVision && isPdf && fileBytes > 0 && fileBytes <= 25 * 1024 * 1024) {
    const pdfB64 = fs.readFileSync(filePath).toString('base64');
    const pngTiles = await pdfToImageTiles(pdfB64);
    const limited = (pngTiles || []).slice(0, 3);
    if (limited.length) {
      for (const tile of limited) {
        parts.push({ type: 'image', source: { type: 'base64', media_type: 'image/png', data: tile } });
      }
      console.log(`[drawing-context] Vision fallback: ${limited.length} image(s) (capped)`);
    }
  } else if (needVision && fileBytes > 25 * 1024 * 1024) {
    console.log('[drawing-context] Large file — vision skipped; ask-user / local text only');
  } else {
    console.log('[drawing-context] Text-only mode — zero image tokens');
  }

  parts.push({
    type: 'text',
    text: `${local.markdown}\n\n=== RAW EXTRACTED TEXT (reference) ===\n${combinedText.slice(0, 50000)}\n=== END RAW TEXT ===\n\nRULES: Use ONLY schedule values above for quantities. Never invent sizes/qty. Prefer drawing-schedule source.`,
  });

  return {
    parts,
    scheduleFirst: local,
    combinedText,
    needsVision: needVision,
    fileBytes,
    spatialTables: spatial.tables || [],
  };
}

async function buildDrawingContext(pdfB64, opts = {}) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pmc_ctx_'));
  const pdfPath = path.join(tmpDir, 'input.pdf');
  try {
    fs.writeFileSync(pdfPath, Buffer.from(pdfB64, 'base64'));
    return await buildDrawingContextFromFile(pdfPath, { ...opts, mime: 'application/pdf' });
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch (_) {}
  }
}

// ─── DIRECT CLAUDE CHAT ROUTE (no Gemini wrapper) ────────────────
app.post('/claude', async (req, res) => {
  try {
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { system, messages, max_tokens } = req.body;
    if (!messages?.length) return res.status(400).json({ error: 'No messages.' });
    const systemToUse = (system && system.trim().length > 50) ? system : CIVIL_SYSTEM;
    const maxTokens = Math.min(Number(max_tokens) || 4096, 4096);

    let scheduleBundle = null;
    const processedMessages = [];
    // Pull last user text question for local Q&A routing
    let userQuestion = '';
    for (const msg of messages) {
      if (!Array.isArray(msg.content)) {
        if (msg.role === 'user' && typeof msg.content === 'string') userQuestion = msg.content;
        continue;
      }
      for (const part of msg.content) {
        if (part.type === 'text' && part.text) userQuestion = part.text;
      }
    }

    for (const msg of messages) {
      if (!Array.isArray(msg.content)) { processedMessages.push(msg); continue; }

      const newParts = [];
      for (const part of msg.content) {
        if (part.type === 'document' && part.source?.media_type === 'application/pdf') {
          const pdfB64 = part.source.data;
          console.log('[/claude] PDF — local CAD-zoom + multi-type study (tokens only if needed)');
          try {
            const ctx = await buildDrawingContext(pdfB64, { question: userQuestion, filename: 'upload.pdf' });
            scheduleBundle = ctx.scheduleFirst;
            newParts.push(...ctx.parts);
          } catch (e) {
            console.error('[/claude] Drawing context build failed:', e.message);
            newParts.push(part);
          }
        } else if (part.type === 'image' && part.source?.type === 'base64') {
          const imgB64 = part.source.data;
          console.log('[/claude] Image — local CAD-zoom OCR first');
          try {
            let ocrText = '';
            try {
              const zoom = runCadZoomOcrFromBase64(imgB64, part.source.media_type || 'image/png');
              if (zoom?.success) ocrText = zoom.full_text || '';
            } catch (_) {}
            if (!ocrText || ocrText.length < 40) {
              const gcvKey = process.env.GOOGLE_CLOUD_VISION_API_KEY;
              if (gcvKey) {
                const gcvRes = await fetch(`https://vision.googleapis.com/v1/images:annotate?key=${gcvKey}`, {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ requests: [{ image: { content: imgB64 }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }] }] }),
                  signal: AbortSignal.timeout(30000)
                });
                if (gcvRes.ok) {
                  const gcvData = await gcvRes.json();
                  ocrText = gcvData.responses?.[0]?.fullTextAnnotation?.text || '';
                }
              }
            }
            if (ocrText.trim().length > 40) {
              scheduleBundle = runScheduleFirstLocal(ocrText, { filename: 'upload.png', question: userQuestion });
              newParts.push({ type: 'text', text: scheduleBundle.markdown });
              if (scheduleBundle.needsClaude && scheduleBundle.needsVision) newParts.push(part);
            } else {
              newParts.push(part);
              newParts.push({ type: 'text', text: 'WARNING: Local OCR got little text. Read only what is clearly visible. Never invent schedule qty.' });
            }
          } catch (e) {
            console.error('[/claude] Image OCR failed:', e.message);
            newParts.push(part);
          }
        } else {
          newParts.push(part);
        }
      }
      processedMessages.push({ ...msg, content: newParts });
    }

    // Fast path: local multi-type Q&A — NEVER invent; ask user if unclear
    if (scheduleBundle?.answeredLocally && scheduleBundle.markdown) {
      // If we need user answers, do NOT call Claude (saves tokens + prevents guessing)
      if (scheduleBundle.needsUserInput) {
        console.log('[/claude] Asking user — ZERO Claude tokens (no assumptions)');
        return res.json({
          content: [{ type: 'text', text: scheduleBundle.markdown }],
          schedule_first: {
            quality: scheduleBundle.extracted?.quality,
            rows: scheduleBundle.extracted?.total_schedule_rows,
            drawing_type: scheduleBundle.typeInfo?.drawing_type,
            mode: 'ask-user',
            tokens: 0,
            needs_user_input: true,
            questions: scheduleBundle.clarifications?.questions || [],
          },
        });
      }

      console.log('[/claude] Local-only short-circuit — ZERO Claude tokens');
      try {
        learnRatesFromMarkdown(scheduleBundle.markdown, {
          filename: 'chat',
          drawing_type: scheduleBundle.typeInfo?.drawing_type || scheduleBundle.boqResult?.drawing_type,
        });
      } catch (e) {}
      return res.json({
        content: [{ type: 'text', text: scheduleBundle.markdown }],
        schedule_first: {
          quality: scheduleBundle.extracted?.quality,
          rows: scheduleBundle.extracted?.total_schedule_rows,
          drawing_type: scheduleBundle.typeInfo?.drawing_type,
          boq_items: scheduleBundle.boqResult?.boq?.length || 0,
          mode: 'local-only',
          tokens: 0,
          needs_user_input: false,
        },
      });
    }

    // Weak extract: ONE Claude call with already-capped parts (few/no images)
    const raw = await callClaudeAPI({ system: systemToUse, messages: processedMessages, maxTokens });
    try { learnRatesFromMarkdown(raw, { filename: 'chat', drawing_type: scheduleBundle?.typeInfo?.drawing_type || 'GENERAL' }); } catch (e) {}
    return res.json({
      content: [{ type: 'text', text: raw }],
      schedule_first: scheduleBundle ? {
        quality: scheduleBundle.extracted?.quality,
        rows: scheduleBundle.extracted?.total_schedule_rows,
        drawing_type: scheduleBundle.typeInfo?.drawing_type,
        mode: 'claude-assisted',
      } : { mode: 'claude-only' },
    });
  } catch (e) {
    console.error('[/claude]', e.message);
    return res.status(500).json({ error: e.message });
  }
});

app.post('/gemini', async (req, res) => {
  try {
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { body } = req.body;

    //
    const frontendSystem = body?.system_instruction?.parts?.[0]?.text;
    const systemToUse = (frontendSystem && frontendSystem.trim().length > 50) ? frontendSystem : CIVIL_SYSTEM;

    // Extract all message parts (text + images/PDFs) from Gemini-format body
    const claudeMessages = [];
    for (const content of (body?.contents || [])) {
      const claudeParts = [];
      for (const part of (content.parts || [])) {
        if (part.text) {
          claudeParts.push({ type: 'text', text: part.text });
        } else if (part.inline_data) {
          const mt = part.inline_data.mime_type;
          if (mt === 'application/pdf') {
            console.log('[/gemini] PDF — schedule-first buildDrawingContext');
            try {
              const ctx = await buildDrawingContext(part.inline_data.data);
              claudeParts.push(...ctx.parts);
            } catch (e) {
              console.error('[/gemini] buildDrawingContext failed:', e.message);
              claudeParts.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: part.inline_data.data } });
            }
          } else if (mt?.startsWith('image/')) {
            // FIX A: Was duplicated — second handler had empty body, so direct images were DROPPED
            claudeParts.push({ type: 'image', source: { type: 'base64', media_type: mt, data: part.inline_data.data } });
          }
        }
      }
      if (claudeParts.length) claudeMessages.push({ role: content.role === 'user' ? 'user' : 'assistant', content: claudeParts });
    }
    if (!claudeMessages.length) return res.status(400).json({ error: 'No messages.' });

    const raw = await callClaudeAPI({ system: systemToUse, messages: claudeMessages, maxTokens: 4096 });
    // Auto-learn rates from chat responses (BOQ markdown tables)
    try { learnRatesFromMarkdown(raw, { filename: 'chat', drawing_type: 'GENERAL' }); } catch(e) {}
    // Return in Gemini-compatible format so the frontend doesn't need changes
    return res.json({ candidates: [{ content: { parts: [{ text: raw }] }, finishReason: 'STOP' }] });
  } catch (e) {
    console.error('[/gemini → Claude]', e.message);
    return res.status(500).json({ error: e.message });
  }
});

// ─── 2. EXTRACT DATA ─────────────────────────────────────────────
// Strategy: Use AI chat response text as PRIMARY source (already has all data)
// Files only used if no aiResponse available
async function extractData(_key, files, userText, aiResponse) {
  const parts = [];
  

  const primaryText = aiResponse || userText || '';

  if (!aiResponse) {
    // No AI response yet — send actual files to Gemini
    for (const f of (files || [])) {
      try {
        if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i))
          parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
        else if (f.type?.startsWith('image/'))
          parts.push({ inline_data: { mime_type: f.type || 'image/png', data: f.b64 } });
      } catch(e) { console.log('File skip:', e.message); }
    }
  }
  // If STILL no content, use files as fallback even when aiResponse present
  if (!aiResponse && parts.length === 0 && (files||[]).length === 0) {
    return { report_type:'general', project_title: userText||'PMC Report', company:'PMC', date:new Date().toLocaleDateString('en-IN'), summary:'', vendors:[], pricing:{old_rate:[],new_rate:[]}, commercial_terms:[], technical_specs:[], boq_items:[], recommendation:'No data provided.' };
  }

  const prompt = `You are a PMC data extraction expert. Extract ALL data from the content below into JSON.
Return ONLY raw JSON. No markdown. No backticks. Start with { end with }.

CONTENT TO EXTRACT FROM:
${primaryText}

You MUST extract real data from the content above. Do NOT use placeholder values like "v1","v2".
Extract actual vendor names, actual prices, actual specifications found in the content.

Return this exact JSON structure:
{"report_type":"comparison","project_title":"EXTRACT FROM CONTENT","company":"EXTRACT FROM CONTENT","date":"DD-MM-YYYY","summary":"2-3 lines from content",
"vendors":[{"name":"ACTUAL VENDOR NAME","vendor_name":"ACTUAL PERSON NAME","contact":"ACTUAL PHONE","quote_date":"DD-MM-YYYY","brand":"ACTUAL BRAND","product_description":"ACTUAL DESCRIPTION"}],
"pricing":{"old_rate":[{"label":"BASIC AMOUNT (OLD RATE)","values":[ACTUAL_NUMBERS]},{"label":"18% GST","values":[ACTUAL_NUMBERS]},{"label":"TOTAL AMOUNT WITH GST","values":[ACTUAL_NUMBERS]}],
"new_rate":[{"label":"BASIC AMOUNT (NEW RATE)","values":[ACTUAL_NUMBERS]},{"label":"18% GST","values":[ACTUAL_NUMBERS]},{"label":"TOTAL AMOUNT WITH GST","values":[ACTUAL_NUMBERS]}]},
"commercial_terms":[{"label":"PAYMENT TERMS","values":["ACTUAL VALUE FROM CONTENT"]},{"label":"DELIVERY TIME","values":["ACTUAL VALUE"]},{"label":"WARRANTY","values":["ACTUAL VALUE"]}],
"technical_specs":[{"label":"ACTUAL SPEC NAME","values":["ACTUAL SPEC VALUE"]}],
"boq_items":[{"sr":1,"description":"ACTUAL ITEM NAME","unit":"ACTUAL UNIT","qty":ACTUAL_NUMBER,"rate":ACTUAL_NUMBER,"amount":ACTUAL_NUMBER}],
"recommendation":"ACTUAL PMC recommendation from content"}

RULES: Use ACTUAL data from content | Numbers as numbers not strings | ONLY JSON`;

  parts.push({ text: prompt });

  // ✅ CONVERTED: Claude replaces Gemini for data extraction
  const claudeRaw = await callClaudeAPI({ system: CIVIL_SYSTEM, messages: [{ role: 'user', content: parts.map(p => p.text ? { type: 'text', text: p.text } : (p.inline_data?.mime_type === 'application/pdf' ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: p.inline_data.data } } : { type: 'image', source: { type: 'base64', media_type: p.inline_data?.mime_type || 'image/png', data: p.inline_data?.data } })) }], maxTokens: 4096 });
  let raw = claudeRaw || '';
  const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
  if (fb !== -1 && lb !== -1) raw = raw.slice(fb, lb + 1);
  try { return JSON.parse(raw.replace(/```json|```/g, '').trim()); }
  catch (e) {
    console.error('JSON parse fail:', raw.slice(0, 300));
    return { report_type: 'general', project_title: 'PMC Report', company: 'PMC', date: new Date().toLocaleDateString('en-IN'), summary: primaryText.slice(0, 200), vendors: [], pricing: { old_rate: [], new_rate: [] }, commercial_terms: [], technical_specs: [], boq_items: [], recommendation: primaryText.slice(0, 500) };
  }
}

// ─── 3. BUILD EXCEL — EXACT PMC FORMAT ────────────────────────────
async function buildExcel(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'PMC Civil AI Agent';
  const ws = wb.addWorksheet('Comparison');

  // Exact colors from template
  const NAVY    = 'FF1F3864';
  const MIDBLUE = 'FF2E75B6';
  const LTBLUE  = 'FFBDD7EE';
  const YELLOW  = 'FFFFD966';
  const GREEN   = 'FFE2EFDA';
  const DKGREEN = 'FF375623';
  const GREY    = 'FFF2F2F2';
  const WHITE   = 'FFFFFFFF';
  const LOWEST  = 'FF00B050';

  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const bdr  = { top: thin, left: thin, bottom: thin, right: thin };

  const vendors = d.vendors || [];
  const vc = Math.max(vendors.length, 1);
  const LC = 2 + vc; // last column index

  // Set exact col widths from template
  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 32;
  for (let i = 3; i <= LC; i++) ws.getColumn(i).width = 28;

  const sc = (cell, bgArgb, bold = false, fcArgb = 'FF000000', size = 10, align = 'left', wrap = true) => {
    cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
    cell.font   = { bold, color: { argb: fcArgb }, size, name: 'Calibri' };
    cell.alignment = { horizontal: align, vertical: 'middle', wrapText: wrap };
    cell.border = bdr;
  };

  const mergeRow = (r, text, bgArgb, fcArgb = 'FF000000', size = 10, bold = true, height = 18) => {
    ws.mergeCells(r, 1, r, LC);
    const c = ws.getCell(r, 1); c.value = text;
    sc(c, bgArgb, bold, fcArgb, size, 'center');
    ws.getRow(r).height = height;
  };

  let row = 1;

  // ROW 1 — Company title  (bg:1F3864 fc:FFFFFF size:14 bold)
  mergeRow(row++, d.company || 'VCT BHARUCH', NAVY, 'FFFFFFFF', 14, true, 22);

  // ROW 2 — Report title  (bg:2E75B6 fc:FFFFFF size:12 bold)
  mergeRow(row++, (d.project_title || 'COMPARISON REPORT').toUpperCase(), MIDBLUE, 'FFFFFFFF', 12, true, 20);

  // ROW 3 — Column headers  (bg:1F3864 fc:FFFFFF size:9 bold)
  const hRow = ws.getRow(row);
  const h1 = hRow.getCell(1); h1.value = 'SR NO';      sc(h1, NAVY, true, 'FFFFFFFF', 9, 'center');
  const h2 = hRow.getCell(2); h2.value = 'PARTICULARS'; sc(h2, NAVY, true, 'FFFFFFFF', 9, 'center');
  vendors.forEach((v, i) => {
    const c = hRow.getCell(i + 3);
    c.value = `${v.name || ''}\n(${v.brand || ''})\n${v.quote_date || ''}`;
    sc(c, NAVY, true, 'FFFFFFFF', 9, 'center');
  });
  hRow.height = 60; row++;

  // ROWS 4-8 — Vendor info
  const infoRows = [
    { lbl: 'AGENCY NAME',       bg: LTBLUE, bold: true,  vals: vendors.map(v => v.name || '') },
    { lbl: 'VENDOR NAME',       bg: GREY,   bold: false, vals: vendors.map(v => v.vendor_name || '') },
    { lbl: 'CONTACT NO',        bg: LTBLUE, bold: true,  vals: vendors.map(v => String(v.contact || '')) },
    { lbl: 'DATE OF QUOTATION', bg: GREY,   bold: false, vals: vendors.map(v => v.quote_date || '') },
    { lbl: 'BRAND',             bg: LTBLUE, bold: true,  vals: vendors.map(v => v.brand || '') },
  ];
  infoRows.forEach(({ lbl, bg, bold, vals }) => {
    const r = ws.getRow(row);
    const sr = r.getCell(1); sr.value = ''; sc(sr, bg, false, 'FF000000', 10, 'center');
    const lb = r.getCell(2); lb.value = lbl; sc(lb, bg, true, 'FF000000', 10, 'left');
    vals.forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, bold, 'FF000000', 10, 'center'); });
    ws.getRow(row).height = 16; row++;
  });

  // ROW 9 — Product desc header  A9:B9 merged = "SR NO", C9:G9 merged = "PRODUCT DESCRIPTION"
  ws.mergeCells(row, 1, row, 2);
  const pd1 = ws.getCell(row, 1); pd1.value = 'SR NO'; sc(pd1, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
  ws.mergeCells(row, 3, row, LC);
  const pd2 = ws.getCell(row, 3); pd2.value = 'PRODUCT DESCRIPTION'; sc(pd2, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
  ws.getRow(row).height = 16; row++;

  // ROW 10 — Product descriptions
  const pdRow = ws.getRow(row);
  const pdsr = pdRow.getCell(1); pdsr.value = '1'; sc(pdsr, GREY, false, 'FF000000', 10, 'center');
  const pdlb = pdRow.getCell(2); pdlb.value = 'PRODUCT DESCRIPTION'; sc(pdlb, GREY, true, 'FF000000', 10, 'left');
  vendors.forEach((v, i) => {
    const c = pdRow.getCell(i + 3); c.value = v.product_description || '';
    sc(c, WHITE, false, 'FF000000', 9, 'left');
  });
  ws.getRow(row).height = 90; row++;

  // PRICING OLD RATE
  if (d.pricing?.old_rate?.length) {
    mergeRow(row++, 'PRICING — OLD RATE', NAVY, 'FFFFFFFF', 10, true, 18);
    d.pricing.old_rate.forEach(({ label, values }, idx) => {
      const isTotal = label?.toUpperCase().includes('TOTAL');
      const bg = isTotal ? YELLOW : WHITE;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, isTotal, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => {
        const c = r.getCell(i + 3);
        const disp = (v === 0 || v === null || v === '') ? 'N/A' : v;
        c.value = disp;
        if (typeof v === 'number' && v > 0) c.numFmt = '#,##0';
        sc(c, bg, isTotal, 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 16; row++;
    });
  }

  // PRICING NEW RATE
  if (d.pricing?.new_rate?.length) {
    mergeRow(row++, 'PRICING — NEW RATE', NAVY, 'FFFFFFFF', 10, true, 18);
    let totalVals = [];
    d.pricing.new_rate.forEach(({ label, values }) => {
      const isTotal = label?.toUpperCase().includes('TOTAL');
      const isDisc  = label?.toUpperCase().includes('DISCOUNT');
      const bg = isTotal ? YELLOW : isDisc ? GREEN : WHITE;
      if (isTotal) totalVals = values || [];
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, isTotal, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => {
        const c = r.getCell(i + 3);
        c.value = (v === 0 || v === null || v === '') ? (isDisc ? '-' : 'N/A') : v;
        if (typeof v === 'number' && v > 0) c.numFmt = '#,##0';
        sc(c, bg, isTotal, 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 16; row++;
    });

    // LOWEST PRICE ROW
    if (totalVals.length) {
      const nums = totalVals.map(v => typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.]/g, '')) || 0);
      const minVal = Math.min(...nums.filter(n => n > 0));
      mergeRow(row++, 'LOWEST QUOTED PRICE (NEW RATE WITH GST)', NAVY, 'FFFFFFFF', 10, true, 18);
      const lr = ws.getRow(row);
      const lsr = lr.getCell(1); lsr.value = ''; sc(lsr, GREEN, false, 'FF000000', 10, 'center');
      const llb = lr.getCell(2); llb.value = 'TOTAL WITH GST (HIGHLIGHT = LOWEST)'; sc(llb, GREEN, true, 'FF000000', 10, 'left');
      nums.forEach((n, i) => {
        const c = lr.getCell(i + 3);
        const isLow = n === minVal && n > 0;
        if (n > 0) { c.value = n; c.numFmt = '₹#,##0'; }
        else c.value = 'N/A';
        sc(c, isLow ? LOWEST : WHITE, isLow, isLow ? 'FFFFFFFF' : 'FF000000', 10, 'center');
      });
      ws.getRow(row).height = 18; row++;
    }
  }

  // COMMERCIAL TERMS
  if (d.commercial_terms?.length) {
    mergeRow(row++, 'COMMERCIAL TERMS', NAVY, 'FFFFFFFF', 10, true, 18);
    d.commercial_terms.forEach(({ label, values }, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = ''; sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, true, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, false, 'FF000000', 9, 'center'); });
      ws.getRow(row).height = 40; row++;
    });
  }

  // TECHNICAL SPECS
  if (d.technical_specs?.length) {
    mergeRow(row++, 'TECHNICAL SPECIFICATIONS', NAVY, 'FFFFFFFF', 10, true, 18);
    d.technical_specs.forEach(({ label, values }, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row);
      const src = r.getCell(1); src.value = String(idx + 1); sc(src, bg, false, 'FF000000', 10, 'center');
      const lc = r.getCell(2); lc.value = label; sc(lc, bg, true, 'FF000000', 10, 'left');
      (values || []).forEach((v, i) => { const c = r.getCell(i + 3); c.value = v; sc(c, bg, false, 'FF000000', 10, 'center'); });
      ws.getRow(row).height = 16; row++;
    });
  }

  // BOQ
  if (d.boq_items?.length) {
    mergeRow(row++, 'BILL OF QUANTITIES', NAVY, 'FFFFFFFF', 11, true, 18);
    const bHdr = ws.getRow(row++);
    ['SR NO','DESCRIPTION OF WORK','UNIT','QUANTITY','RATE (INR)','AMOUNT (INR)'].forEach((h, i) => {
      const c = bHdr.getCell(i + 1); c.value = h; sc(c, MIDBLUE, true, 'FFFFFFFF', 10, 'center');
    });
    let total = 0;
    d.boq_items.forEach((item, idx) => {
      const bg = idx % 2 === 0 ? WHITE : GREY;
      const r = ws.getRow(row++);
      [item.sr, item.description, item.unit, item.qty, item.rate, item.amount].forEach((v, i) => {
        const c = r.getCell(i + 1); c.value = v;
        sc(c, bg, false, 'FF000000', 10, i === 0 || i > 1 ? 'center' : 'left');
        if (i >= 4 && typeof v === 'number') c.numFmt = '#,##0';
      });
      total += parseFloat(item.amount) || 0;
    });
    ws.mergeCells(row, 1, row, 4);
    const tc = ws.getCell(row, 1); tc.value = 'GRAND TOTAL'; sc(tc, YELLOW, true, 'FF000000', 10, 'right');
    const ta = ws.getCell(row, 6); ta.value = total; ta.numFmt = '₹#,##0'; sc(ta, YELLOW, true, 'FF000000', 10, 'center');
    ws.getRow(row).height = 18; row++;
  }

  // PMC RECOMMENDATION — dark green header + light green box
  mergeRow(row++, 'PMC RECOMMENDATION', DKGREEN, 'FFFFFFFF', 11, true, 18);
  ws.mergeCells(row, 1, row, LC);
  const recCell = ws.getCell(row, 1);
  recCell.value = d.recommendation || 'Refer to chat analysis above.';
  sc(recCell, GREEN, true, 'FF000000', 10, 'left');
  ws.getRow(row).height = 70; row++;

  // Summary
  if (d.summary) {
    ws.mergeCells(row, 1, row, LC);
    const sCell = ws.getCell(row, 1);
    sCell.value = 'SUMMARY: ' + d.summary;
    sc(sCell, LTBLUE, false, 'FF000000', 9, 'left', true);
    sCell.font = { ...sCell.font, italic: true };
    ws.getRow(row).height = 30; row++;
  }

  // Footer
  ws.mergeCells(row, 1, row, LC);
  const fCell = ws.getCell(row, 1);
  const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
  fCell.value = `Prepared by: PMC Civil AI Agent  |  Date: ${today}  |  VCT Bharuch — Powered by Claude AI`;
  fCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY } };
  fCell.font = { italic: true, size: 9, color: { argb: 'FF595959' }, name: 'Calibri' };
  fCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(row).height = 14;

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
  return wb;
}

// ─── 4. EXCEL ENDPOINT ─────────────────────────────────────────────
app.post('/export-excel', async (req, res) => {
  try {
    const key = process.env.CLAUDE_API_KEY;
    if (!key) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;
    const d = await extractData(key, files, userText, aiResponse);
    const wb = await buildExcel(d);
    const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="PMC_Report_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch (err) {
    console.error('Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 5. PDF ENDPOINT (print-ready HTML) ────────────────────────────
app.post('/export-pdf', async (req, res) => {
  try {
    const key = process.env.CLAUDE_API_KEY;
    if (!key) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;
    const d = await extractData(key, files, userText, aiResponse);
    const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const vendors = d.vendors || [];
    const vc = Math.max(vendors.length, 1);

    const th = (txt, bg = '#1F3864', fc = '#fff', bold = true) =>
      `<th style="background:${bg};color:${fc};padding:6px 8px;font-size:9px;border:1px solid #000;text-align:center;font-weight:${bold?'bold':'normal'};">${txt}</th>`;
    const td = (txt, bg = '#fff', align = 'center', bold = false, size = 9) =>
      `<td style="background:${bg};color:#000;padding:6px 8px;font-size:${size}px;border:1px solid #ccc;text-align:${align};font-weight:${bold?'bold':'normal'};vertical-align:top;">${txt||''}</td>`;
    const sectionHdr = (txt, bg = '#1F3864') =>
      `<tr><td colspan="${vc+2}" style="background:${bg};color:#fff;font-weight:bold;padding:7px 10px;font-size:10px;border:1px solid #000;">${txt}</td></tr>`;
    const fmtNum = (v) => typeof v === 'number' && v > 0 ? '₹' + v.toLocaleString('en-IN') : (v === 0 ? 'N/A' : (v || 'N/A'));

    let html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
@page{size:A3 landscape;margin:8mm}
*{box-sizing:border-box}
body{font-family:Calibri,Arial,sans-serif;font-size:9px;margin:0;color:#000}
h1{background:#1F3864;color:#fff;text-align:center;padding:9px;margin:0;font-size:15px}
h2{background:#2E75B6;color:#fff;text-align:center;padding:7px;margin:0 0 4px;font-size:12px}
table{width:100%;border-collapse:collapse;margin-bottom:2px}
.rec-hdr{background:#375623;color:#fff;font-weight:bold;padding:7px 10px;font-size:10px;margin-top:4px}
.rec-body{background:#E2EFDA;padding:10px;font-size:9px;border:1px solid #375623;white-space:pre-wrap}
.summ{background:#BDD7EE;padding:7px 10px;font-size:8px;font-style:italic;margin-top:3px}
.footer{text-align:center;font-size:8px;color:#595959;margin-top:6px;font-style:italic}
</style></head><body>
<h1>${d.company || 'VCT BHARUCH'}</h1>
<h2>${(d.project_title || 'COMPARISON REPORT').toUpperCase()}</h2>
<table>
<tr>${th('SR NO')}${th('PARTICULARS')}${vendors.map(v => th(`${v.name||''}<br><small>(${v.brand||''})</small><br><small>${v.quote_date||''}</small>`)).join('')}</tr>
${[['AGENCY NAME','BDD7EE',true,v=>v.name],['VENDOR NAME','F2F2F2',false,v=>v.vendor_name],['CONTACT NO','BDD7EE',false,v=>v.contact],['DATE OF QUOTATION','F2F2F2',false,v=>v.quote_date],['BRAND','BDD7EE',true,v=>v.brand]].map(([lbl,bg,bold,fn])=>`<tr>${td('',`#${bg}`)}<td style="background:#${bg};padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${lbl}</td>${vendors.map(v=>td(fn(v)||'',`#${bg}`,'center',bold)).join('')}</tr>`).join('')}
${sectionHdr('PRODUCT DESCRIPTION','#2E75B6')}
<tr>${td('1','#F2F2F2','center')}${td('<b>PRODUCT DESCRIPTION</b>','#F2F2F2','left',true)}${vendors.map(v=>td(v.product_description||'','#fff','left',false,8)).join('')}</tr>
${d.pricing?.old_rate?.length ? sectionHdr('PRICING — OLD RATE') + d.pricing.old_rate.map(({label,values})=>{const isT=label?.toUpperCase().includes('TOTAL');const bg=isT?'#FFD966':'#fff';return`<tr>${td('',bg)}${td(label,bg,'left',isT)}${(values||[]).map(v=>td(fmtNum(v),bg,'center',isT)).join('')}</tr>`;}).join('') : ''}
${d.pricing?.new_rate?.length ? sectionHdr('PRICING — NEW RATE') + d.pricing.new_rate.map(({label,values})=>{const isT=label?.toUpperCase().includes('TOTAL');const isD=label?.toUpperCase().includes('DISCOUNT');const bg=isT?'#FFD966':isD?'#E2EFDA':'#fff';return`<tr>${td('',bg)}${td(label,bg,'left',isT)}${(values||[]).map(v=>td(isT&&typeof v==='number'&&v>0?'₹'+v.toLocaleString('en-IN'):isD&&(v===0||!v)?'-':fmtNum(v),bg,'center',isT)).join('')}</tr>`;}).join('') : ''}
${(()=>{const tr=d.pricing?.new_rate?.find(r=>r.label?.toUpperCase().includes('TOTAL'));if(!tr)return'';const nums=(tr.values||[]).map(v=>typeof v==='number'?v:0);const minV=Math.min(...nums.filter(n=>n>0));return sectionHdr('LOWEST QUOTED PRICE')+`<tr>${td('')}<td style="background:#E2EFDA;padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">TOTAL WITH GST (HIGHLIGHT = LOWEST)</td>${nums.map(n=>n===minV&&n>0?`<td style="background:#00B050;color:#fff;padding:6px 8px;font-size:9px;border:1px solid #ccc;text-align:center;font-weight:bold;">₹${n.toLocaleString('en-IN')} ✓</td>`:td(n>0?'₹'+n.toLocaleString('en-IN'):'N/A','#fff','center')).join('')}</tr>`;})()}
${d.commercial_terms?.length?sectionHdr('COMMERCIAL TERMS')+d.commercial_terms.map(({label,values},i)=>{const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td('',bg)}<td style="background:${bg};padding:7px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${label}</td>${(values||[]).map(v=>td(v||'',bg,'center',false,8)).join('')}</tr>`;}).join(''):''}
${d.technical_specs?.length?sectionHdr('TECHNICAL SPECIFICATIONS')+d.technical_specs.map(({label,values},i)=>{const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td(i+1,bg,'center')}<td style="background:${bg};padding:6px 8px;font-size:9px;border:1px solid #ccc;font-weight:bold;">${label}</td>${(values||[]).map(v=>td(v||'',bg,'center')).join('')}</tr>`;}).join(''):''}
${d.boq_items?.length?(()=>{let tot=0;const rows=d.boq_items.map(({sr,description,unit,qty,rate,amount},i)=>{tot+=parseFloat(amount)||0;const bg=i%2===0?'#fff':'#F2F2F2';return`<tr>${td(sr,bg,'center')}${td(description,bg,'left')}${td(unit,bg,'center')}${td(qty,bg,'center')}${td(rate?'₹'+rate.toLocaleString('en-IN'):'',bg,'center')}${td(amount?'₹'+amount.toLocaleString('en-IN'):'',bg,'center')}</tr>`;}).join('');return sectionHdr('BILL OF QUANTITIES')+`<tr>${['SR NO','DESCRIPTION','UNIT','QTY','RATE','AMOUNT'].map(h=>th(h,'#2E75B6')).join('')}</tr>${rows}<tr><td colspan="5" style="background:#FFD966;padding:7px;font-weight:bold;border:1px solid #000;text-align:right;">GRAND TOTAL</td><td style="background:#FFD966;padding:7px;font-weight:bold;border:1px solid #000;text-align:center;">₹${tot.toLocaleString('en-IN')}</td></tr>`;})():''}
</table>
<div class="rec-hdr">PMC RECOMMENDATION</div>
<div class="rec-body">${d.recommendation||'Refer to chat analysis.'}</div>
${d.summary?`<div class="summ">SUMMARY: ${d.summary}</div>`:''}
<div class="footer">Prepared by: PMC Civil AI Agent &nbsp;|&nbsp; Date: ${today} &nbsp;|&nbsp; VCT Bharuch — Powered by Claude AI</div>
</body></html>`;

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="PMC_Report_${today.replace(/\//g,'-')}.html"`);
    res.send(html);
  } catch (err) {
    console.error('PDF error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 6. DRAWING ANALYSIS → MULTI-SHEET EXCEL (CV + AI) ──────────
app.post('/export-drawing', async (req, res) => {
  try {
    const key = process.env.CLAUDE_API_KEY;
    if (!key) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;

    // Step 1: Run OpenCV pixel-level analysis on images
    let cvData = {};
    const imageFiles = (files||[]).filter(f => f.type?.startsWith('image/'));
    if (imageFiles.length > 0) {
      try { cvData = runCVAnalysis(imageFiles[0].b64); }
      catch(e) { console.log('CV skipped:', e.message); }
    }

    // Step 2: Schedule-first from PDF text (Civils.ai-style)
    let drawingData = null;
    const pdfFiles = (files||[]).filter(f => f.type === 'application/pdf' || f.name?.match(/\.pdf$/i));
    if (pdfFiles.length > 0) {
      try {
        const ctx = await buildDrawingContext(pdfFiles[0].b64);
        cvData.drawing_context_text = ctx.combinedText || '';
        cvData.schedule_first = ctx.scheduleFirst;
        if (ctx.scheduleFirst?.boqResult?.boq?.length) {
          drawingData = {
            ...ctx.scheduleFirst.boqResult,
            elements: (ctx.scheduleFirst.boqResult.boq || []).map((item, i) => ({
              id: `E${String(i + 1).padStart(3, '0')}`,
              type: 'STRUCTURE',
              name: item.description,
              quantities: { volume_cum: item.unit === 'cum' ? item.qty : 0, steel_kg: item.unit === 'kg' ? item.qty : 0 },
              cost_inr: { total: item.amount || 0, per_unit: item.rate || 0 },
              confidence: item.confidence,
              annotation_found: item.source,
            })),
            prepared_by: 'PMC Civil AI — Schedule-first Pipeline',
          };
          console.log(`[export-drawing] Schedule-first BOQ: ${drawingData.boq.length} items`);
        }
        // Only keep overview image(s) if schedules were weak
        if (ctx.needsVision) {
          let tileCount = 0;
          for (const part of ctx.parts) {
            if (part.type === 'image' && part.source?.type === 'base64' && tileCount < 2) {
              files.push({ type: 'image/png', b64: part.source.data, name: `pdf_tile_${++tileCount}.png` });
            }
          }
        }
      } catch (e) {
        console.warn('[export-drawing] schedule-first failed:', e.message);
      }
    }

    // Step 2b: Fallback — text-only Claude analyzer (no 5-phase vision)
    if (!drawingData && files?.length > 0) {
      drawingData = await geminiAnalyzeDrawing(key, files, cvData, fetch);
    }

    if (!drawingData) {
      drawingData = await extractDrawingData(key, files, userText, aiResponse, fetch);
    }

    drawingData.cv_analysis = cvData;
    drawingData.prepared_by = drawingData.prepared_by || 'PMC Civil AI Agent';

    const wb = await buildDrawingExcel(drawingData);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const pname = (drawingData.project_name||'Drawing').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_PMC_Analysis_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch (err) {
    console.error('Drawing Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 7. DXF UPLOAD & ANALYSIS ─────────────────────────────────────
// Uses drawing_intelligence.js — reads legend, auto-maps layers, extracts levels
app.post('/analyze-dxf', async (req, res) => {
  try {
    const claudeKey = process.env.CLAUDE_API_KEY;
    if (!claudeKey) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    // ── Step 1: Drawing Intelligence — scan, detect legend, auto-map layers ──
    const analyzed = analyzeDrawing(dxfContent, filename);
    console.log(`[DXF] ${filename} | ${analyzed.total_layers} layers | ${analyzed.floor_levels.length} floor levels | ${analyzed.element_counts.wall_polylines} wall polylines | ${analyzed.unknown_layers.length} unknown layers`);

    // ── Step 2: Smart BOQ Engine — pre-digest drawing into structured engineering data ──
    // Out-of-the-box approach: Claude gets a PRE-DRAFTED BOQ to verify, not raw data to guess from
    // This shifts Claude from "guesser" to "checker" — 90-95% accuracy
    const ratesMap = getRatesMap();
    const smartCtx = buildSmartContextFromAnalyzed(analyzed, ratesMap);
    const prompt = smartCtx.summary_text;
    console.log(`[DXF Smart] Pre-drafted ${smartCtx.pre_drafted_boq?.length || 0} BOQ items, ${smartCtx.rooms?.length || 0} rooms, ${smartCtx.wall_quantities?.length || 0} wall entries`);

    // ── Step 3: Claude verifies + fixes + completes the pre-drafted BOQ ──────────────
    const claudeResp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type':'application/json','x-api-key':claudeKey,'anthropic-version':'2023-06-01','anthropic-beta':'pdfs-2024-09-25' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6', max_tokens: 4096,
        system: CIVIL_SYSTEM,
        messages: [{ role:'user', content: prompt }]
      })
    });
    const claudeData = await claudeResp.json();
    let raw = claudeData?.content?.find(b=>b.type==='text')?.text || '{}';
    const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
    let geminiResult = {};
    if (fb !== -1) try { geminiResult = JSON.parse(raw.slice(fb, lb+1)); } catch(e) { console.error('JSON parse fail:', e.message); }
    // Attach pre-drafted data for fallback
    if (!geminiResult.boq?.length && smartCtx.pre_drafted_boq?.length) {
      geminiResult.boq = smartCtx.pre_drafted_boq;
      geminiResult._source = 'pre_draft_fallback';
    }

    // ── Step 4: Return everything — drawing data + AI interpretation ──────────
    res.json({
      success: true,
      dxf_data: {
        filename:         analyzed.filename,
        project_name:     analyzed.project_name,
        drawing_extents:  analyzed.drawing_extents,
        floor_levels:     analyzed.floor_levels,
        floor_heights:    analyzed.floor_heights,
        legend_items:     analyzed.legend_items,
        layer_summary:    analyzed.layer_summary,
        wall_by_thickness_m2: analyzed.wall_by_thickness_m2,
        hatch_summary:    analyzed.hatch_summary,
        element_counts:   analyzed.element_counts,
        unknown_layers:   analyzed.unknown_layers,
        unknown_blocks:   analyzed.unknown_blocks,
        all_texts:        analyzed.all_texts_sample,
        layer_names:      analyzed.layer_names,
        stats: {
          total_layers:    analyzed.total_layers,
          total_texts:     analyzed.total_texts,
          total_hatches:   analyzed.total_hatches,
          total_polylines: analyzed.total_polylines,
          total_inserts:   analyzed.total_inserts,
        }
      },
      interpretation: geminiResult
    });

  } catch (err) {
    console.error('DXF analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/export-dxf-excel', async (req, res) => {
  try {
    const { dxfContent, filename, aiResponse } = req.body;
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content.' });

    // Parse DXF + attach coordinate-clustered schedule tables
    const parsed = parseDXF(dxfContent);
    let civilData = extractCivilData(parsed, filename);
    civilData = attachScheduleTables(civilData); // adds schedule_tables[] for accurate BOQ

    // ✅ SMART ENGINE: Pre-draft BOQ from drawing data, Claude only verifies
    let geminiResult = {};
    try {
      const ratesMap = getRatesMap();
      const smartCtx = buildSmartContext(civilData, ratesMap);
      console.log(`[DXF-Excel Smart] Pre-drafted ${smartCtx.pre_drafted_boq?.length || 0} BOQ items`);
      
      geminiResult = await claudeAnalyzeDXF(civilData, filename, getRatesSummary({ maxItems: 40 }), smartCtx.summary_text);
      console.log('[DXF-Excel] Claude analysis done:', geminiResult.drawing_type);
      
      // Fallback: use pre-draft if Claude fails
      if (!geminiResult.boq?.length && smartCtx.pre_drafted_boq?.length) {
        geminiResult.boq = smartCtx.pre_drafted_boq;
        geminiResult._source = 'smart_pre_draft_fallback';
      }
    } catch(e) { console.log('Claude DXF interp fail:', e.message); }

    // Build Excel
    const wb = await buildDXFExcel(civilData, geminiResult, ExcelJS);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g,'-');
    const pname = (geminiResult.project_name||filename||'DXF').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_DXF_Analysis_${today}.xlsx"`);
    await wb.xlsx.write(res); res.end();

  } catch (err) {
    console.error('DXF Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── 8. DRAWING → EXCEL (AI Analysis + Auto Excel) ───────────────
app.post('/drawing-to-excel', async (req, res) => {
  try {
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { files, userText, aiResponse } = req.body;

    // FIX BUG-1: claudeAnalyzeDrawingVision() already returns a parsed JS object
    // (parseJSON is called internally). Never call .replace() on the result.
    let drawingData = {};
    try {
      const analysisResult = await claudeAnalyzeDrawingVision(files, userText, aiResponse);
      if (analysisResult && typeof analysisResult === 'object') {
        drawingData = analysisResult;
      } else if (typeof analysisResult === 'string') {
        // Defensive: if somehow a string comes back, parse it
        const clean = analysisResult.replace(/```json|```/g, '').trim();
        const fb2 = clean.indexOf('{'), lb2 = clean.lastIndexOf('}');
        if (fb2 !== -1) { try { drawingData = JSON.parse(clean.slice(fb2, lb2+1)); } catch(e2) {} }
      }
      console.log('[drawing-to-excel] Claude done | type:', drawingData.drawing_type || '?', '| boq items:', drawingData.boq?.length || 0);
    } catch(e) { console.log('Claude drawing-to-excel fail:', e.message); }

    // Build Excel
    const wb = await buildExcelFromDrawing(drawingData);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    const pname = (drawingData.project_name || 'Drawing').replace(/[^a-zA-Z0-9_]/g,'_').slice(0,20);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${pname}_PMC_Estimate_${today}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('Drawing→Excel error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});
// ── NEW: DXF → AREA STATEMENT + OVERALL SUMMARY auto-update ──
app.post('/update-area-from-dxf', async (req, res) => {
  try {
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    const totalAreaSqft = extractTotalAreaSqft(dxfContent);
    if (!totalAreaSqft || totalAreaSqft <= 0)
      return res.status(400).json({ error: 'No closed polylines found in DXF. Area calculate nahi hui.' });

    const estimatePath = dataPath('templates', 'UPDATED-OVERALL-ESTIMATE-MODESTAA-10.04.2026.xlsx');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(estimatePath);

    // Update AREA STATEMENT C73
    const wsArea = wb.getWorksheet('AREA STATEMENT');
    if (wsArea) wsArea.getCell('C73').value = totalAreaSqft;

    // Update OVERALL SUMMARY
    const wsOS = wb.getWorksheet('OVERALL SUMMARY');
    if (wsOS) {
      // Row 6 display text
      wsOS.getCell('B6').value = `TOTAL AREA: ${totalAreaSqft.toLocaleString('en-IN', {maximumFractionDigits:2})} SQFT`;
      // Helper cell J6 stores area value
      wsOS.getCell('J6').value = totalAreaSqft;
      // Replace all hardcoded 273613.53 with dynamic reference to J6
      wsOS.eachRow(row => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (typeof cell.value === 'string' && cell.value.includes('273613.53')) {
            cell.value = cell.value.split('273613.53').join("'OVERALL SUMMARY'!$J$6");
          }
        });
      });
    }

    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=ESTIMATE-UPDATED-${today}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Area update error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ── NEW: Fill MODESTAA template from drawing (type-aware) ──
// Detects project type from DXF content. If high-rise residential,
// opens the MODESTAA template and fills drawing-derived cells only.
// Otherwise, builds a fresh workbook via buildExcelFromDrawing with
// the right BOQ sheet for the detected project type (cafe / institute
// / commercial / road / generic).
app.post('/fill-template-from-drawing', async (req, res) => {
  try {
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content provided.' });

    const parsed = parseDXF(dxfContent);
    const civil  = extractCivilData(parsed, filename || 'drawing.dxf');
    const ptype  = (civil.project_type || 'generic').toLowerCase();
    const ec     = civil.element_counts || {};
    const totalAreaSqft = extractTotalAreaSqft(dxfContent) || 0;
    const totalAreaSqm  = totalAreaSqft > 0 ? Math.round((totalAreaSqft / 10.764) * 100) / 100 : 0;

    // ── HIGH-RISE: use MODESTAA template, fill only drawing-derived cells ──
    if (ptype === 'high_rise_residential') {
      const estimatePath = dataPath('templates', 'UPDATED-OVERALL-ESTIMATE-MODESTAA-10.04.2026.xlsx');
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(estimatePath);

      // AREA STATEMENT C73 — total area
      if (totalAreaSqft > 0) {
        const wsArea = wb.getWorksheet('AREA STATEMENT');
        if (wsArea) wsArea.getCell('C73').value = totalAreaSqft;
      }

      // OVERALL SUMMARY B6 / J6
      const wsOS = wb.getWorksheet('OVERALL SUMMARY');
      if (wsOS && totalAreaSqft > 0) {
        wsOS.getCell('B6').value = `TOTAL AREA: ${totalAreaSqft.toLocaleString('en-IN',{maximumFractionDigits:2})} SQFT`;
        wsOS.getCell('J6').value = totalAreaSqft;
        wsOS.eachRow(row => {
          row.eachCell({ includeEmpty: false }, cell => {
            if (typeof cell.value === 'string' && cell.value.includes('273613.53')) {
              cell.value = cell.value.split('273613.53').join("'OVERALL SUMMARY'!$J$6");
            }
          });
        });
      }

      // DRAWING-DERIVED COUNTS sheet (new) — record what the parser read
      let wsCounts = wb.getWorksheet('DRAWING COUNTS');
      if (!wsCounts) wsCounts = wb.addWorksheet('DRAWING COUNTS');
      wsCounts.getCell('A1').value = 'ELEMENT';
      wsCounts.getCell('B1').value = 'COUNT FROM DRAWING';
      wsCounts.getCell('C1').value = 'SOURCE';
      [['Floors', ec.floor_count || 0, (ec.floor_labels || []).join(', ')],
       ['Doors',  ec.door_count  || 0, 'block / layer match'],
       ['Windows',ec.window_count|| 0, 'block / layer match'],
       ['Lifts',  ec.lift_count  || 0, 'block / layer / text'],
       ['Staircases', ec.staircase_count || 0, 'block / layer / text'],
       ['Columns', ec.column_count || 0, 'block / layer match'],
       ['Beams',   ec.beam_count   || 0, 'block / layer match'],
       ['Footings',ec.footing_count|| 0, 'block / layer match'],
       ['Toilets', ec.toilet_count || 0, 'text annotations'],
       ['Kitchens',ec.kitchen_count|| 0, 'text annotations'],
       ['Bedrooms',ec.bedroom_count|| 0, 'text annotations'],
       ['Wall length (m)', civil.wall_length_m || 0, 'LINE entities on wall layers'],
       ['Total area (sqft)', totalAreaSqft, 'closed polylines (shoelace)'],
       ['Project type detected', civil.project_type || 'generic', 'dxf_parser.detectProjectType']
      ].forEach((row, i) => {
        row.forEach((v, j) => { wsCounts.getCell(i+2, j+1).value = v; });
      });

      const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=MODESTAA-FILLED-${today}.xlsx`);
      await wb.xlsx.write(res);
      return res.end();
    }

    // ── OTHER TYPES: build fresh type-aware workbook ──
    const data = {
      drawing_type:    civil.drawing_type === 'FLOOR_PLAN' ? 'BUILDING' : 'SITE_LAYOUT',
      project_type:    ptype,
      total_area_sqm:  totalAreaSqm,
      total_area_sqft: totalAreaSqft,
      element_counts:  ec,
      wall_length_m:   civil.wall_length_m || 0,
      buildings: totalAreaSqm > 0 ? [{ name: 'Building', area_sqm: totalAreaSqm, floors: ec.floor_count || 0 }] : [],
      roads: [],
      project_name: civil.title_block?.project_name || filename || 'Project',
      source: `DXF parser — project type: ${ptype}`
    };
    const wb = await buildExcelFromDrawing(data);
    const today = new Date().toLocaleDateString('en-IN').replace(/\//g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${ptype.toUpperCase()}-ESTIMATE-${today}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Fill template error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ─── DWG/DXF ANALYSIS — Convert to PNG + Claude Vision ────────────
// Strategy: dwg_converter.py renders DXF/DWG to PNG using ezdxf+matplotlib
// Then Claude SEES the actual drawing like a human engineer (ZWCAD compatible)
app.post('/analyze-dwg', async (req, res) => {
  try {
    const key = process.env.CLAUDE_API_KEY;
    if (!key) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });

    const { b64, filename, detailMode } = req.body;
    if (!b64) return res.status(400).json({ error: 'No file data provided.' });
    const useDetail = detailMode === true || detailMode === 'true' || detailMode === 1;

    const fs = require('fs');
    const { execSync } = require('child_process');
    const os = require('os');

    // Write uploaded file to temp
    const ext = filename?.match(/\.(dxf|dwg|dwf)$/i)?.[1]?.toLowerCase() || 'dxf';
    const tmpIn  = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.${ext}`);
    const tmpPng = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.png`);

    fs.writeFileSync(tmpIn, Buffer.from(b64, 'base64'));

    // Run converter. DXF/DWG → ezdxf Python script. DWF → LibreOffice fallback.
    const scriptPath = scriptsPath('dwg_converter.py');
    let converterResult = {};

    if (ext === 'dwf') {
      // DWF support is weak industry-wide. Try LibreOffice first; if it fails, tell user to re-export.
      try {
        const soffice = process.platform === 'win32'
          ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
          : 'libreoffice';
        execSync(`${soffice} --headless --convert-to png --outdir "${os.tmpdir()}" "${tmpIn}"`,
                 { timeout: 90000 });
        const base = path.basename(tmpIn, '.dwf');
        const libreOut = path.join(os.tmpdir(), `${base}.png`);
        if (fs.existsSync(libreOut)) {
          converterResult = { success: true, png_path: libreOut, texts: [], dimensions: [], layers: [], drawing_type: 'DWF_RENDER' };
        } else {
          converterResult = {
            success: false,
            needsPdfOrDxf: true,
            error: 'DWF format is not supported by this system. ' +
              'Please re-export your drawing from ZWCAD or AutoCAD as PDF or DXF:\n' +
              '  ZWCAD: File → Export → PDF  (or File → Save As → DXF 2018)\n' +
              '  AutoCAD: File → Export → PDF (or SaveAs → DXF 2018)\n' +
              'Then re-upload the PDF or DXF file.'
          };
        }
      } catch (e) {
        converterResult = {
          success: false,
          needsPdfOrDxf: true,
          error: 'DWF format could not be converted (LibreOffice is not available or failed). ' +
            'Please re-export your drawing as PDF or DXF:\n' +
            '  ZWCAD: File → Export → PDF  (or File → Save As → DXF 2018)\n' +
            '  AutoCAD: File → Export → PDF (or SaveAs → DXF 2018)\n' +
            'Then re-upload the PDF or DXF file.'
        };
      }
    } else {
      try {
        const py = pyExec();
        // FIXED: tiling used to only run when the user explicitly enabled
        // "detail mode" — so by default every DWG/DXF was flattened into one
        // low-res image and small schedule-table numbers were unreadable.
        // dwg_converter.py now always tiles the main sheet and scales the
        // grid to the drawing's density; detailMode just requests a higher
        // base render DPI on top of that for extra-dense sheets.
        const dpi = useDetail ? 350 : 300;
        const tiledArg = 'true';
        const out = execSync(
          `${py} "${scriptPath}" "${tmpIn}" "${tmpPng}" ${dpi} ${tiledArg}`,
          { timeout: 180000, maxBuffer: 40 * 1024 * 1024, env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' } }
        );
        converterResult = JSON.parse(out.toString());
      } catch (e) {
        const isDwg = ext === 'dwg';
        const userMsg = isDwg
          ? `DWG file could not be converted using ezdxf. ` +
            `Please open the file in ZWCAD or AutoCAD and re-save as DXF:\n` +
            `  ZWCAD: File → Save As → File type: "AutoCAD 2018 DXF (*.dxf)"\n` +
            `  AutoCAD: File → Save As → DXF 2018\n` +
            `Then re-upload the saved .dxf file.`
          : `DXF conversion failed: ${e.message}`;
        converterResult = { success: false, error: userMsg, needsDxfExport: isDwg };
      }
    }

    // ── Early exit: if conversion failed AND no PNG was produced, return clear error to user ──
    if (!converterResult.success && !converterResult.png_path) {
      try { fs.unlinkSync(tmpIn); } catch(e) {}
      return res.status(422).json({
        success: false,
        error: converterResult.error || 'File could not be converted.',
        needsDxfExport: !!converterResult.needsDxfExport,
        needsPdfOrDxf:  !!converterResult.needsPdfOrDxf,
        converter: converterResult
      });
    }

    // DWF or any path that has PNG but no tiles yet: split with helper script
    if (useDetail && converterResult.png_path && fs.existsSync(converterResult.png_path)
        && (!converterResult.tiles || !converterResult.tiles.length)) {
      try {
        const outDir = path.dirname(converterResult.png_path);
        const baseName = path.basename(converterResult.png_path, path.extname(converterResult.png_path));
        const tileScript = scriptsPath('tile_only.py');
        const py = pyExec();
        const tout = execSync(
          `${py} "${tileScript}" "${converterResult.png_path}" "${outDir}" "${baseName}"`,
          { timeout: 60000, maxBuffer: 10 * 1024 * 1024, env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' } }
        );
        const ta = JSON.parse(tout.toString().trim() || '[]');
        if (Array.isArray(ta) && ta.length) converterResult.tiles = ta;
      } catch (e) {
        console.warn('Tile split (fallback):', e.message);
      }
    }

    // Schedule-first from converter text; images only if schedules weak (token save)
    const textSummary = (converterResult.texts || []).map(t => t.text).slice(0, 2000).join('\n');
    const dimSummary  = (converterResult.dimensions || [])
      .filter(d => d.value).map(d => `${d.value}${d.text ? ' ('+d.text+')' : ''}`).slice(0, 500).join(', ');
    const layers = (converterResult.layers || []).join(', ');
    const dwgText = [textSummary, dimSummary, layers].filter(Boolean).join('\n');
    const scheduleBundle = runScheduleFirstLocal(dwgText, {
      filename: filename || 'drawing.dwg',
      question: 'study drawing: identify type, levels, schedules, tables',
      hints: [],
    });
    console.log(`[DWG] type=${scheduleBundle.typeInfo?.drawing_type} schedule=${scheduleBundle.extracted.quality} rows=${scheduleBundle.extracted.total_schedule_rows} localQA=${!!scheduleBundle.answeredLocally}`);

    const parts = [];
    // Always keep overview PNG
    if (converterResult.png_path && fs.existsSync(converterResult.png_path)) {
      const pngB64 = fs.readFileSync(converterResult.png_path).toString('base64');
      if (scheduleBundle.needsVision || useDetail) {
        parts.push({ inline_data: { mime_type: 'image/png', data: pngB64 } });
      }
    }
    // Cap detail tiles to 2 when vision needed
    let tileKept = 0;
    if ((scheduleBundle.needsVision || useDetail) && Array.isArray(converterResult.tiles)) {
      for (const t of converterResult.tiles) {
        if (tileKept >= 2) {
          try { if (t.path) fs.unlinkSync(t.path); } catch (e) {}
          continue;
        }
        if (t.path && fs.existsSync(t.path)) {
          try {
            const tb = fs.readFileSync(t.path).toString('base64');
            parts.push({ inline_data: { mime_type: 'image/png', data: tb } });
            tileKept++;
            try { fs.unlinkSync(t.path); } catch (e) {}
          } catch (e) { /* skip */ }
        }
      }
    } else if (Array.isArray(converterResult.tiles)) {
      for (const t of converterResult.tiles) {
        try { if (t.path) fs.unlinkSync(t.path); } catch (e) {}
      }
    }
    for (const li of (converterResult.layout_images || [])) {
      try { if (li.path && li.path !== converterResult.png_path) fs.unlinkSync(li.path); } catch (e) {}
    }
    if (converterResult.png_path) {
      try { fs.unlinkSync(converterResult.png_path); } catch (e) {}
    }

    let analysisRaw = null;

    // Fast path: local multi-type Q&A — skip Claude vision when possible
    if (scheduleBundle.answeredLocally && !scheduleBundle.needsClaude) {
      console.log('[DWG] Local Q&A short-circuit — ZERO Claude vision tokens');
      analysisRaw = {
        ...(scheduleBundle.boqResult || {}),
        drawing_type: scheduleBundle.typeInfo?.drawing_type,
        column_schedule: (scheduleBundle.boqResult?.schedule_data?.columns || []).map(c => ({
          col_mark: c.mark, size_mm: c.size_mm, main_bars: c.main_bars, stirrups: c.stirrups, qty: c.qty, source: c.source,
        })),
        footing_schedule: (scheduleBundle.boqResult?.schedule_data?.footings || []).map(f => ({
          footing_mark: f.mark, pcc_size_mm: f.pcc_size_mm, rcc_size_mm: f.rcc_size_mm,
          depth_mm: f.depth_mm, main_bars_x: f.main_bars_x, main_bars_y: f.main_bars_y, qty: f.qty, source: f.source,
        })),
        _markdown: scheduleBundle.markdown,
      };
    } else if (!parts.length && scheduleBundle.markdown) {
      // DWG binary-only (no PNG): return local study + tip to export PDF/DXF for zoom OCR
      analysisRaw = {
        ...(scheduleBundle.boqResult || {}),
        drawing_type: scheduleBundle.typeInfo?.drawing_type || 'GENERAL_DRAWING',
        _markdown: `${scheduleBundle.markdown}\n\n> **Note:** Native DWG did not render to PNG on this server. For AutoCAD-like zoom reading of tables/sections, export from ZWCAD/AutoCAD as **PDF or DXF** and re-upload.`,
      };
    } else {
      const pngTiles = parts.filter(p => p.inline_data?.mime_type === 'image/png').map(p => p.inline_data.data);
      try {
        if (pngTiles.length) {
          analysisRaw = await claudeAnalyzeDWGVision(pngTiles.slice(0, 3), {
            ...converterResult,
            texts: converterResult.texts,
            _schedule_markdown: scheduleBundle.markdown,
          }, filename);
          console.log('[DWG] Claude vision (capped tiles) done');
        } else {
          analysisRaw = scheduleBundle.boqResult;
          analysisRaw._markdown = scheduleBundle.markdown;
        }
      } catch (e) {
        console.error('Claude DWG analysis fail:', e.message);
        analysisRaw = scheduleBundle.boqResult;
        if (analysisRaw) analysisRaw._markdown = scheduleBundle.markdown;
      }
    }

    // FIX: claudeAnalyzeDWGVision returns parsed JSON object (via parseJSON).
    // Convert it to a human-readable markdown string for the frontend to display.
    // Also keep the raw structured data in response for Excel export.
    function boqToMarkdown(d) {
      if (!d) return null;
      const lines = [];
      lines.push(`## DWG Analysis: ${filename}`);
      if (d.project_name) lines.push(`**Project:** ${d.project_name}`);
      if (d.drawing_no) lines.push(`**Drawing No:** ${d.drawing_no}`);
      if (d.drawing_type) lines.push(`**Drawing Type:** ${d.drawing_type}`);
      if (d.scale) lines.push(`**Scale:** ${d.scale}`);
      if (d.concrete_grade) lines.push(`**Concrete Grade:** ${d.concrete_grade}`);
      if (d.steel_grade) lines.push(`**Steel Grade:** ${d.steel_grade}`);
      if (d.structural_system) lines.push(`**Structural System:** ${d.structural_system}`);
      lines.push('');

      // Column Schedule
      if (d.column_schedule?.length) {
        lines.push('### Column Schedule');
        lines.push('| Mark | Size (mm) | Main Bars | Stirrups | Qty | Floor | Source |');
        lines.push('|------|-----------|-----------|----------|-----|-------|--------|');
        for (const c of d.column_schedule) {
          lines.push(`| ${c.col_mark||''} | ${c.size_mm||''} | ${c.main_bars||''} | ${c.stirrups||''} | ${c.qty||0} | ${c.floor||''} | ${c.source||''} |`);
        }
        lines.push('');
      }

      // Footing Schedule
      if (d.footing_schedule?.length) {
        lines.push('### Footing Schedule');
        lines.push('| Mark | PCC Size | RCC Size | Depth | PCC Thk | Bars X | Bars Y | Qty | Pedestal | Source |');
        lines.push('|------|----------|----------|-------|---------|--------|--------|-----|----------|--------|');
        for (const f of d.footing_schedule) {
          lines.push(`| ${f.footing_mark||''} | ${f.pcc_size_mm||''} | ${f.rcc_size_mm||''} | ${f.depth_mm||0} | ${f.pcc_thickness_mm||150} | ${f.main_bars_x||''} | ${f.main_bars_y||''} | ${f.qty||0} | ${f.pedestal_size_mm||''} | ${f.source||''} |`);
        }
        lines.push('');
      }

      // Base Plate Schedule
      if (d.base_plate_schedule?.length) {
        lines.push('### Base Plate Schedule');
        lines.push('| Column Mark | Plate Size (mm) | Anchor Bolts | Bolt Dia (mm) | Source |');
        lines.push('|-------------|----------------|--------------|----------------|--------|');
        for (const b of d.base_plate_schedule) {
          lines.push(`| ${b.column_mark||''} | ${b.plate_size_mm||''} | ${b.anchor_bolt_nos||0} | ${b.anchor_bolt_dia_mm||0} | ${b.source||''} |`);
        }
        lines.push('');
      }

      // Section Details
      if (d.section_details) {
        const s = d.section_details;
        lines.push('### Section Details');
        if (s.footing_depth_mm) lines.push(`- Footing Depth: **${s.footing_depth_mm} mm**`);
        if (s.pedestal_height_mm) lines.push(`- Pedestal Height: **${s.pedestal_height_mm} mm**`);
        if (s.pcc_thickness_mm) lines.push(`- PCC Thickness: **${s.pcc_thickness_mm} mm**`);
        if (s.cover_mm) lines.push(`- Clear Cover: **${s.cover_mm} mm**`);
        lines.push('');
      }

      // Grid Info
      if (d.grid_info?.total_columns_plan) {
        lines.push('### Grid Information');
        lines.push(`- Total Columns (Plan): **${d.grid_info.total_columns_plan}**`);
        if (d.grid_info.typical_bay_m) lines.push(`- Typical Bay: **${d.grid_info.typical_bay_m} m**`);
        if (d.grid_info.braced_bay_grids?.length) lines.push(`- Braced Bays: ${d.grid_info.braced_bay_grids.join(', ')}`);
        lines.push('');
      }

      // BOQ Table
      if (d.boq?.length) {
        lines.push('### Bill of Quantities');
        lines.push('| Sr | Description | Unit | Qty | Rate (₹) | Amount (₹) | Confidence |');
        lines.push('|----|-------------|------|-----|----------|------------|------------|');
        for (const b of d.boq) {
          if (b.part && !b.description) { lines.push(`| **${b.part}** | | | | | | |`); continue; }
          const amt = b.amount ? b.amount.toLocaleString('en-IN') : '0';
          const rate = b.rate ? b.rate.toLocaleString('en-IN') : '0';
          lines.push(`| ${b.sr||''} | ${b.description||''} | ${b.unit||''} | ${b.qty||0} | ${rate} | ${amt} | ${b.confidence||''} |`);
        }
        lines.push('');
        if (d.cost_summary?.civil_total_lacs) {
          lines.push(`**Total Civil Cost: ₹${d.cost_summary.civil_total_inr?.toLocaleString('en-IN')||0} (₹${d.cost_summary.civil_total_lacs} Lacs)**`);
        }
        lines.push('');
      }

      // Observations
      if (d.observations?.length) {
        lines.push('### PMC Observations');
        for (const o of d.observations) lines.push(`- ${o}`);
        lines.push('');
      }
      if (d.not_legible_fields?.length) {
        lines.push('### Not Legible / Not Found');
        for (const nf of d.not_legible_fields) lines.push(`- ${nf}`);
        lines.push('');
      }

      lines.push('> Analyzed by PMC Civil AI (Claude Vision — ZWCAD/AutoCAD DWG compatible)');
      return lines.join('\n');
    }

    const analysisMarkdown = analysisRaw?._markdown
      || (analysisRaw ? boqToMarkdown(analysisRaw) : null)
      || scheduleBundle.markdown;

    const fallbackAnalysis =
      `## DWG/DXF File: ${filename}\n\n` +
      `**Schedule quality:** ${scheduleBundle.extracted.quality}\n` +
      `**Layers:** ${layers || "none"}\n` +
      `**Texts found:** ${(converterResult.texts||[]).length}\n\n` +
      scheduleBundle.markdown +
      "\n> Schedule-first DWG pipeline (Civils.ai-style).";

    // Cleanup temp input
    try { fs.unlinkSync(tmpIn); } catch(e) {}

    if (analysisMarkdown) {
      try {
        const learnedCount = learnRatesFromMarkdown(analysisMarkdown, {
          filename,
          drawing_type: converterResult.drawing_type || scheduleBundle.boqResult?.drawing_type || 'UNKNOWN',
        });
        if (learnedCount > 0) console.log(`[rate_store] Learned ${learnedCount} rates from DWG`);
      } catch (e) { console.warn('[rate_store] learn failed:', e.message); }
    }

    res.json({
      success: true,
      analysis: analysisMarkdown || fallbackAnalysis,
      structured: analysisRaw || scheduleBundle.boqResult || null,
      converter: converterResult,
      detailMode: useDetail,
      quadrantTiles: tileKept,
      schedule_first: {
        quality: scheduleBundle.extracted.quality,
        rows: scheduleBundle.extracted.total_schedule_rows,
        mode: scheduleBundle.needsVision ? 'vision-assisted' : 'local',
      },
      ai_engine: 'Claude Vision (ZWCAD compatible)',
    });
  } catch (err) {
    console.error('DWG analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── 9. SYMBOL CLASSIFICATION — Step 1: classify known/unknown ────
// Called right after DXF upload. Returns known symbols + unknown list.
// Unknown symbols will be shown to user as questions in the chat UI.
app.post('/classify-dxf', async (req, res) => {
  try {
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });
    const { dxfContent, filename } = req.body;
    if (!dxfContent) return res.status(400).json({ error: 'No DXF content.' });

    const fs = require('fs');

    // Load learned symbols from disk
    const learnedPath = dataPath('symbols-learned.json');
    let learned = { blocks: {}, layers: {} };
    try { learned = JSON.parse(fs.readFileSync(learnedPath, 'utf8')); } catch(e) {}

    // Parse DXF + attach coordinate-clustered schedule tables
    const parsed = parseDXF(dxfContent);
    let civilData = extractCivilData(parsed, filename);
    civilData = attachScheduleTables(civilData); // adds schedule_tables[] for accurate BOQ

    const allBlocks = Object.keys(civilData.block_counts || {});
    const allLayers = civilData.layer_names || [];

    // Split blocks into known (in learned dict) vs unknown
    const knownBlocks = {};
    const unknownBlocks = [];
    for (const b of allBlocks) {
      const bUp = b.toUpperCase();
      // Check learned dict first
      if (learned.blocks[bUp]) {
        knownBlocks[b] = learned.blocks[bUp];
        continue;
      }
      // Check common AutoCAD naming conventions
      const autoType = guessBlockType(b);
      if (autoType) {
        knownBlocks[b] = autoType;
      } else {
        unknownBlocks.push({ name: b, count: civilData.block_counts[b] || 1 });
      }
    }

    // Split layers into known vs unknown
    const knownLayers = {};
    const unknownLayers = [];
    const LAYER_PREFIXES = {
      'A-': 'architectural', 'S-': 'structural', 'E-': 'electrical',
      'P-': 'plumbing', 'M-': 'mechanical', 'C-': 'civil',
      'WALL': 'wall', 'DOOR': 'door', 'WINDOW': 'window',
      'COLUMN': 'column', 'COL': 'column', 'BEAM': 'beam',
      'SLAB': 'slab', 'STAIR': 'staircase', 'LIFT': 'lift',
      'RAMP': 'ramp', 'TOILET': 'toilet', 'KITCHEN': 'kitchen',
      'PARK': 'parking', 'ROAD': 'road', 'HATCH': 'hatch',
      'DIM': 'dimension', 'TEXT': 'text', 'TITLE': 'title-block',
      'DEFPOINTS': 'dimension-helper', '0': 'default'
    };
    for (const l of allLayers) {
      const lUp = l.toUpperCase();
      if (learned.layers[lUp]) { knownLayers[l] = learned.layers[lUp]; continue; }
      let matched = false;
      for (const [pfx, type] of Object.entries(LAYER_PREFIXES)) {
        if (lUp.startsWith(pfx) || lUp.includes(pfx)) {
          knownLayers[l] = type; matched = true; break;
        }
      }
      if (!matched) unknownLayers.push(l);
    }

    // Call Claude ONLY if > threshold truly unknown blocks — saves 70% classify calls
    let geminiClassified = { blocks: {}, layers: {} };
    const CLAUDE_CLASSIFY_THRESHOLD = 3;
    const needsClaude = unknownBlocks.length > CLAUDE_CLASSIFY_THRESHOLD;
    if (needsClaude) {
      try {
        geminiClassified = await claudeClassifySymbols(unknownBlocks, unknownLayers, civilData, filename);
        console.log('[classify-dxf] Claude classified', Object.keys(geminiClassified.blocks||{}).length, 'blocks');
      } catch(e) { console.log('Claude classify fail:', e.message); }
    } else {
      console.log(`[classify-dxf] Skipped Claude — only ${unknownBlocks.length} unknown blocks (threshold:${CLAUDE_CLASSIFY_THRESHOLD})`);
    }

    // Merge all known
    const finalKnownBlocks = { ...knownBlocks, ...(geminiClassified.blocks || {}) };
    const finalKnownLayers = { ...knownLayers, ...(geminiClassified.layers || {}) };

    // These still need user input
    const askUserBlocks = (geminiClassified.still_unknown_blocks || [])
      .map(name => ({ name, count: civilData.block_counts[name] || 1 }));
    const askUserLayers = geminiClassified.still_unknown_layers || [];

    res.json({
      success: true,
      filename,
      dxf_data: civilData,
      known_blocks: finalKnownBlocks,
      known_layers: finalKnownLayers,
      ask_user_blocks: askUserBlocks,
      ask_user_layers: askUserLayers,
      needs_questions: askUserBlocks.length > 0 || askUserLayers.length > 0
    });

  } catch (err) {
    console.error('classify-dxf error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Helper: guess block type from common AutoCAD naming conventions
function guessBlockType(name) {
  const n = name.toUpperCase();
  if (/^D\d+$|DR[-_]?\d|DOOR|FLUSH|SFD|DOOR[-_]/.test(n)) return 'door';
  if (/^W\d+$|WIN[-_]?\d|WINDOW|ALUM[-_]WIN|CASEMENT/.test(n)) return 'window';
  if (/COL[-_]?\d|^C\d+$|COLUMN|PILLAR/.test(n)) return 'column';
  if (/BEAM|BM[-_]?\d/.test(n)) return 'beam';
  if (/LIFT|ELEV|ELEVATOR/.test(n)) return 'lift';
  if (/STAIR|STC|STEP/.test(n)) return 'staircase';
  if (/RAMP/.test(n)) return 'ramp';
  if (/TOILET|WC|BATH/.test(n)) return 'toilet';
  if (/KITCHEN|PANTRY/.test(n)) return 'kitchen';
  if (/BED|MASTER/.test(n)) return 'bedroom';
  if (/SOFA|TABLE|CHAIR|FURN/.test(n)) return 'furniture';
  if (/TREE|SHRUB|PLANT/.test(n)) return 'landscaping';
  if (/CAR|VEHICLE|PARK/.test(n)) return 'parking';
  return null;
}

// ─── 10. ANALYZE WITH USER ANSWERS — Step 2: full BOQ after Q&A ───
// Receives: original dxf_data + all known symbols + user's answers
// Returns: full Gemini BOQ analysis → used to generate Excel
app.post('/analyze-with-answers', async (req, res) => {
  try {
    if (!process.env.CLAUDE_API_KEY) return res.status(500).json({ error: 'CLAUDE_API_KEY not set.' });

    const { dxfContent, filename, knownBlocks, knownLayers, userAnswers, dxfData } = req.body;
    const fs = require('fs');

    // Save user answers to symbols-learned.json for future drawings
    const learnedPath = dataPath('symbols-learned.json');
    let learned = { blocks: {}, layers: {} };
    try { learned = JSON.parse(fs.readFileSync(learnedPath, 'utf8')); } catch(e) {}

    // Merge user answers into learned dict
    if (userAnswers?.blocks) {
      for (const [name, type] of Object.entries(userAnswers.blocks)) {
        if (type && type !== 'skip') learned.blocks[name.toUpperCase()] = type;
      }
    }
    if (userAnswers?.layers) {
      for (const [name, type] of Object.entries(userAnswers.layers)) {
        if (type && type !== 'skip') learned.layers[name.toUpperCase()] = type;
      }
    }
    try { fs.writeFileSync(learnedPath, JSON.stringify(learned, null, 2)); } catch(e) {}

    // Build complete symbol map (known + user answered)
    const allKnownBlocks = { ...(knownBlocks || {}), ...(userAnswers?.blocks || {}) };
    const allKnownLayers = { ...(knownLayers || {}), ...(userAnswers?.layers || {}) };

    // Use stored dxfData or re-parse if dxfContent provided
    let civilData = dxfData;
    if (!civilData && dxfContent) {
      const parsed = parseDXF(dxfContent);
      civilData = extractCivilData(parsed, filename);
    }
    if (!civilData) return res.status(400).json({ error: 'No drawing data.' });

    // Build symbol summary for Gemini
    const symbolSummary = [
      ...Object.entries(allKnownBlocks).map(([name, type]) =>
        `Block "${name}" (×${civilData.block_counts?.[name] || '?'}) = ${type}`),
      ...Object.entries(allKnownLayers).map(([name, type]) =>
        `Layer "${name}" = ${type}`)
    ].join('\n');

    const ratesSummary = getRatesSummary({ maxItems: 40 });

    const prompt = `You are a senior PMC civil engineer generating a complete BOQ.
ALL DATA IS FROM THIS DXF FILE. DO NOT INVENT VALUES.

FILE: ${filename}
DRAWING TYPE: ${civilData.drawing_type}
SCALE: ${civilData.scale || 'not detected'}
UNITS: ${civilData.units}
DRAWING SIZE: ${civilData.drawing_extents.width_m}m × ${civilData.drawing_extents.height_m}m

SYMBOL DICTIONARY (confirmed by user + AI):
${symbolSummary || 'none'}

ELEMENT COUNTS:
Doors: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='door').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.door_count||0}
Windows: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='window').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.window_count||0}
Columns: ${Object.entries(allKnownBlocks).filter(([,t])=>t==='column').map(([n])=>`${n}(×${civilData.block_counts?.[n]||0})`).join(', ')||civilData.element_counts?.column_count||0}
Lifts: ${civilData.element_counts?.lift_count||0}
Staircases: ${civilData.element_counts?.staircase_count||0}
Floors: ${civilData.element_counts?.floor_count||0}
Wall length: ${civilData.wall_length_m||0}m

FLOOR LEVELS:
${(civilData.floor_levels||[]).map(l=>`${l.label}=${l.level_m||'?'}m`).join('\n')||'none'}

TEXT ANNOTATIONS:
${civilData.all_texts.slice(0,500).join('\n')}

ROOM LABELS: ${(civilData.room_annotations||[]).map(r=>r.text).join(', ')||'none'}

DIMENSIONS (top 200): ${civilData.dimension_values.slice(0,200).map(d=>`${d.value_m}m[${d.layer}]`).join(', ')}

AREAS from polylines: ${civilData.polyline_areas.slice(0,100).map(p=>`${p.area_sqm}sqm(${p.layer})`).join(', ')}

GUJARAT DSR 2025 RATES:
${ratesSummary}

Generate complete BOQ. Return ONLY raw JSON:
{
  "project_name": "",
  "drawing_type": "",
  "scale": "",
  "building_height_m": 0,
  "floor_count": 0,
  "total_bua_sqm": 0,
  "spaces": [{"name":"","area_sqm":0}],
  "boq": [
    {"sr":1,"description":"","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed"}
  ],
  "element_counts": {"door_count":0,"window_count":0,"lift_count":0,"staircase_count":0,"column_count":0},
  "observations": [],
  "pmc_recommendation": ""
}`;

    // ✅ FIX: Claude replaces Gemini for final BOQ analysis
    let geminiResult = {};
    try {
      geminiResult = await claudeAnalyzeWithAnswers(civilData, filename, symbolSummary, ratesSummary);
      console.log('[analyze-with-answers] Claude done, BOQ items:', geminiResult.boq?.length || 0);
      // ── NEW: Auto-learn rates from BOQ result ──
      if (geminiResult.boq?.length) {
        try {
          learnRatesFromBOQ(geminiResult.boq, { filename, drawing_type: geminiResult.drawing_type });
        } catch(e) { console.warn('[rate_store]', e.message); }
      }
    } catch(e) { console.log('Claude analyze-with-answers fail:', e.message); }

    res.json({ success: true, interpretation: geminiResult, dxf_data: civilData, learned_count: Object.keys(learned.blocks).length + Object.keys(learned.layers).length });

  } catch (err) {
    console.error('analyze-with-answers error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─── 11. RATES STATS — Admin endpoint to see learned rates ─────────
app.get('/rates-stats', (req, res) => {
  try {
    const stats = getLearnedRateStats();
    const baseCount = Object.keys(require('./lib/rate_store').loadBaseRates()).length;
    res.json({ ...stats, base_dsr_items: baseCount, message: 'PMC Rate Store stats' });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── ANALYZE DRAWING (multipart) — any size, Civils.ai-style local-first ───
// Prefer this over stuffing PDF base64 into /claude JSON (breaks on large sheets).
app.post('/analyze-drawing', (req, res) => {
  uploadDrawing.single('file')(req, res, async (err) => {
    if (err) {
      console.error('[/analyze-drawing] upload error:', err.message);
      return res.status(400).json({
        error: err.code === 'LIMIT_FILE_SIZE'
          ? 'File too large for this server (max 500MB). Compress PDF or split sheets.'
          : err.message,
      });
    }
    let tmpPath = req.file?.path;
    try {
      const question = (req.body?.question || req.body?.userText || '').trim();
      let filename = req.file?.originalname || req.body?.filename || 'drawing.pdf';
      let mime = req.file?.mimetype || req.body?.mime || '';

      // JSON fallback: { b64, filename, question } for older clients
      if (!tmpPath && req.body?.b64) {
        const ext = path.extname(filename).toLowerCase() || '.pdf';
        tmpPath = path.join(os.tmpdir(), `pmc_b64_${Date.now()}${ext}`);
        fs.writeFileSync(tmpPath, Buffer.from(req.body.b64, 'base64'));
        mime = mime || (ext === '.pdf' ? 'application/pdf' : 'application/octet-stream');
      }
      if (!tmpPath || !fs.existsSync(tmpPath)) {
        return res.status(400).json({ error: 'No file uploaded. Use FormData field "file".' });
      }

      const ext = path.extname(filename).toLowerCase();
      const sizeMb = (fs.statSync(tmpPath).size / (1024 * 1024)).toFixed(1);
      console.log(`[/analyze-drawing] ${filename} (${sizeMb} MB) q="${question.slice(0, 80)}"`);

      // DWG/DXF/DWF → reuse converter path then local OCR on PNG if produced
      if (['.dwg', '.dxf', '.dwf'].includes(ext)) {
        const { execSync } = require('child_process');
        const scriptPath = scriptsPath('dwg_converter.py');
        const tmpPng = path.join(os.tmpdir(), `pmc_dwg_${Date.now()}.png`);
        let converterResult = {};
        try {
          const py = pyExec();
          const out = execSync(
            `${py} "${scriptPath}" "${tmpPath}" "${tmpPng}" 300 true`,
            { timeout: 180000, maxBuffer: 40 * 1024 * 1024 }
          );
          converterResult = JSON.parse(out.toString());
        } catch (e) {
          converterResult = { success: false, error: e.message };
        }
        if (!converterResult.success && !converterResult.png_path) {
          return res.status(422).json({
            success: false,
            needsDxfExport: ext === '.dwg',
            needsPdfOrDxf: true,
            error: converterResult.error ||
              'CAD file could not be rendered. Export PDF/DXF from ZWCAD/AutoCAD and re-upload (any size OK).',
          });
        }
        const pngPath = converterResult.png_path || tmpPng;
        const texts = (converterResult.texts || []).map(t => (typeof t === 'string' ? t : t.text || '')).filter(Boolean);
        let combined = texts.join('\n');
        let ocrBoxes = [];
        let hints = [];
        if (fs.existsSync(pngPath)) {
          const zoom = runCadZoomOcrOnFile(pngPath);
          if (zoom?.success && zoom.full_text) {
            combined = [combined, zoom.full_text].filter(Boolean).join('\n');
            ocrBoxes = zoom.boxes || [];
            hints = zoom.drawing_hints || [];
          }
        }
        const spatial = buildSpatialScheduleText({ ocrBoxes, plainLines: combined.split('\n') });
        const scheduleBundle = runScheduleFirstLocal(spatial.text || combined, {
          filename,
          question,
          hints,
          spatialTables: spatial.tables || [],
        });
        const md = scheduleBundle.markdown || 'No schedule data found. Upload PDF export of the same sheet.';
        return res.json({
          success: true,
          content: [{ type: 'text', text: md }],
          analysis: md,
          extracted: scheduleBundle.extracted,
          clarifications: scheduleBundle.clarifications,
          combined_text: (spatial.text || combined).slice(0, 80000),
          schedule_first: {
            quality: scheduleBundle.extracted?.quality,
            rows: scheduleBundle.extracted?.total_schedule_rows,
            drawing_type: scheduleBundle.typeInfo?.drawing_type,
            mode: scheduleBundle.needsUserInput ? 'ask-user' : 'local-only',
            tokens: 0,
            file_mb: Number(sizeMb),
            needs_user_input: !!scheduleBundle.needsUserInput,
            questions: scheduleBundle.clarifications?.questions || [],
          },
        });
      }

      // PDF / image — disk-based local pipeline
      const ctx = await buildDrawingContextFromFile(tmpPath, {
        question,
        filename,
        mime: mime || (ext === '.pdf' ? 'application/pdf' : undefined),
      });
      const scheduleBundle = ctx.scheduleFirst;

      // ALWAYS return local read/ask-user first — never drop clarifications via Claude polish
      if (scheduleBundle?.markdown) {
        const payload = {
          success: true,
          content: [{ type: 'text', text: scheduleBundle.markdown }],
          analysis: scheduleBundle.markdown,
          extracted: scheduleBundle.extracted,
          clarifications: scheduleBundle.clarifications,
          combined_text: (ctx.combinedText || '').slice(0, 80000),
          filename,
          schedule_first: {
            quality: scheduleBundle.extracted?.quality,
            rows: scheduleBundle.extracted?.total_schedule_rows,
            drawing_type: scheduleBundle.typeInfo?.drawing_type,
            boq_items: scheduleBundle.boqResult?.boq?.length || 0,
            status: scheduleBundle.qa?.meta?.status || scheduleBundle.qa?.status || (scheduleBundle.needsUserInput ? 'DRAFT' : 'FINAL'),
            mode: scheduleBundle.needsUserInput ? 'ask-user' : (scheduleBundle.qa?.meta?.intent || 'read-calc'),
            tokens: 0,
            file_mb: Number(sizeMb),
            needs_user_input: !!scheduleBundle.needsUserInput,
            questions: scheduleBundle.clarifications?.questions || [],
            spatial_tables: (ctx.spatialTables || []).length,
            text_chars: (ctx.combinedText || '').length,
          },
        };
        if (!scheduleBundle.needsUserInput) {
          try {
            learnRatesFromMarkdown(scheduleBundle.markdown, {
              filename,
              drawing_type: scheduleBundle.typeInfo?.drawing_type,
            });
          } catch (_) {}
        }
        return res.json(payload);
      }

      const fallback = `Drawing received (${sizeMb} MB) but local extract was weak.\n\nPlease:\n1. Export a clearer PDF from CAD\n2. Or type footing lines: Mark LxB Depth Qty\n3. Or ask again after screenshot of schedule table`;
      return res.json({
        success: true,
        content: [{ type: 'text', text: fallback }],
        analysis: fallback,
        clarifications: {
          questions: [{
            id: 'footing_schedule_paste',
            question: 'SCHEDULE OF FOOTING type karo — har line: `F1 2600x1800 900 12`',
            why: 'Extract empty',
          }],
        },
        extracted: { schedules: { footings: [], columns: [] }, quality: 'poor', total_schedule_rows: 0 },
        schedule_first: {
          mode: 'weak-extract',
          file_mb: Number(sizeMb),
          tokens: 0,
          needs_user_input: true,
          questions: [{ id: 'footing_schedule_paste', question: 'SCHEDULE OF FOOTING type karo' }],
        },
      });
    } catch (e) {
      console.error('[/analyze-drawing]', e.message);
      return res.status(500).json({ error: e.message });
    } finally {
      if (tmpPath) {
        try { fs.unlinkSync(tmpPath); } catch (_) {}
      }
    }
  });
});

// ─── RESOLVE CLARIFICATIONS — merge user answers → rebuild BOQ ───
app.post('/resolve-clarifications', (req, res) => {
  try {
    const {
      userText,
      extracted,
      clarifications,
      combined_text,
      filename,
      question,
      hints,
      boqOpts,
    } = req.body || {};
    if (!userText || !extracted || !clarifications) {
      return res.status(400).json({ error: 'Need userText + extracted + clarifications' });
    }
    const result = applyUserClarifications({
      text: combined_text || '',
      extracted,
      clarifications,
      userText,
      filename: filename || 'drawing.pdf',
      question: question || 'finalize with my answers',
      hints: hints || [],
      boqOpts: boqOpts || {},
    });
    return res.json({
      success: true,
      analysis: result.markdown,
      content: [{ type: 'text', text: result.markdown }],
      extracted: result.extracted,
      clarifications: result.clarifications,
      answers: result.answers,
      schedule_first: {
        mode: result.needsUserInput ? 'ask-user' : 'user-confirmed',
        needs_user_input: !!result.needsUserInput,
        questions: result.clarifications?.questions || [],
        drawing_type: result.typeInfo?.drawing_type,
        boq_items: result.boqResult?.boq?.length || 0,
        tokens: 0,
      },
    });
  } catch (e) {
    console.error('[/resolve-clarifications]', e.message);
    return res.status(500).json({ error: e.message });
  }
});

// ─── 12. HEALTH ─────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  const claudeKey = process.env.CLAUDE_API_KEY;
  res.json({
    status: 'ok',
    claude_key_set: !!claudeKey,
    claude_preview: claudeKey ? claudeKey.slice(0, 12) + '...' : 'NOT SET ❌',
    pipeline: 'Civils.ai-style: READ drawing → formula takeoff → DRAFT/FINAL BOQ (spatial+OCR+ask-user)',
    max_upload_mb: 500,
    routes: ['/analyze-drawing','/resolve-clarifications','/claude','/gemini','/export-excel','/export-pdf','/export-drawing','/analyze-dxf','/export-dxf-excel','/drawing-to-excel','/update-area-from-dxf','/fill-template-from-drawing','/analyze-dwg','/classify-dxf','/analyze-with-answers','/rates-stats'],
    dwg_support: 'ZWCAD + AutoCAD — prefer PDF/DXF; size up to 500MB via /analyze-drawing'
  });
});

const APP_URL = process.env.RENDER_EXTERNAL_URL;
if (APP_URL) setInterval(() => fetch(APP_URL + '/health').catch(() => {}), 14 * 60 * 1000);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ PMC Civil AI Agent on port ${PORT}`);
  console.log(`🔑 CLAUDE_API_KEY: ${process.env.CLAUDE_API_KEY ? 'SET ✅' : 'NOT SET ❌'}`);
  console.log('✅ Local-first: CAD-zoom OCR → drawing type → Q&A/BOQ (Claude only if needed)');
  console.log('🏗️  PDF/DXF/DWG/ZWCAD: multi-type sheets (section, footing, plan, elevation…)');
});
