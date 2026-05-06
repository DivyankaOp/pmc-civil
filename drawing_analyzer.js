/**
 * PMC Drawing Analyzer — SMART PIPELINE v2
 * ─────────────────────────────────────────────────────────────────
 * Drawing READING  → PyMuPDF + GCV (Google Cloud Vision) — free/cheap
 * Drawing REASONING → Claude TEXT ONLY — ZERO image tokens
 *
 * Pipeline:
 *   PDF (vector)  → PyMuPDF text extraction → structured text
 *   PDF (scanned) → GCV Document AI        → table rows
 *   Image (PNG)   → GCV Vision API         → OCR text
 *   DXF data      → passed via cvData      → entity text
 *   All above     → Claude (TEXT ONLY)     → BOQ JSON
 *
 * server.js: ZERO changes needed — same exported function signatures.
 */

'use strict';

const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ── RATES ─────────────────────────────────────────────────────────
const RATES = (() => {
  const out = {};
  try {
    const rp = fs.existsSync(path.join(__dirname, 'rates.json'))
      ? path.join(__dirname, 'rates.json') : path.join(__dirname, 'Rates.json');
    const raw = JSON.parse(fs.readFileSync(rp, 'utf8'));
    for (const cat of Object.values(raw))
      if (typeof cat === 'object' && !Array.isArray(cat))
        for (const [k, v] of Object.entries(cat))
          if (v?.rate) out[k] = v.rate;
  } catch (e) { console.warn('Rates.json not loaded:', e.message); }
  return out;
})();

const RATES_STRING = (() => {
  try {
    const rp = fs.existsSync(path.join(__dirname, 'Rates.json'))
      ? path.join(__dirname, 'Rates.json') : path.join(__dirname, 'rates.json');
    const raw = JSON.parse(fs.readFileSync(rp, 'utf8'));
    const lines = [];
    for (const [cat, items] of Object.entries(raw)) {
      if (cat.startsWith('_') || typeof items !== 'object') continue;
      for (const [, v] of Object.entries(items))
        if (v?.rate) lines.push(`${v.description} -> Rs.${v.rate}/${v.unit}`);
    }
    return lines.join('\n');
  } catch { return ''; }
})();

// ── KNOWLEDGE BASE ────────────────────────────────────────────────
function knowledgeBaseHints() {
  try {
    const kb = JSON.parse(fs.readFileSync(path.join(__dirname, 'ymbols-learned.json'), 'utf8'));
    const hints = [];
    if (kb.quantity_corrections?.length) {
      hints.push('CORRECTION HISTORY:');
      for (const c of kb.quantity_corrections.slice(-5))
        hints.push(`  ${c.element}: AI=${c.ai_said}, correct=${c.correct_value}`);
    }
    const blocks = Object.entries(kb.blocks || {}).slice(0, 10);
    if (blocks.length) {
      hints.push('KNOWN BLOCKS:');
      for (const [b, m] of blocks) hints.push(`  ${b} = ${m}`);
    }
    return hints.join('\n');
  } catch { return ''; }
}

// ── SYSTEM PROMPT ─────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a senior PMC civil engineer with 20 years experience in Gujarat, India.
You receive PRE-EXTRACTED TEXT from engineering drawings and generate accurate BOQ.

ABSOLUTE RULES:
1. Use ONLY values present in EXTRACTED TEXT below — never invent or assume.
2. If a value is not in extracted text -> write "not found in drawing".
3. Apply scale factor to ALL measurements before calculating quantities.
4. Schedule table cells: copy EXACTLY as extracted — never fill in missing values.
5. Return ONLY raw JSON. No markdown. No explanation.

GUJARAT DSR 2025 RATES:
${RATES_STRING}`;

// ════════════════════════════════════════════════════════════════
// READER 1: PyMuPDF — vector PDF text extraction (free)
// ════════════════════════════════════════════════════════════════
async function extractVectorPDF(pdfB64) {
  const tmpDir = fs.mkdtempSync(os.tmpdir() + '/pmc_read_');
  const pdfPath = path.join(tmpDir, 'input.pdf');
  try {
    fs.writeFileSync(pdfPath, Buffer.from(pdfB64, 'base64'));
    const script = `
import fitz, json, sys
doc = fitz.open('${pdfPath.replace(/\\/g, '/')}')
pages = []
for pnum in range(len(doc)):
    page = doc[pnum]
    blocks = page.get_text("dict")["blocks"]
    texts = []
    for b in blocks:
        if b.get("type") == 0:
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    t = span.get("text","").strip()
                    if t and len(t) >= 1:
                        x, y = span["origin"]
                        texts.append({"text": t, "x": round(x,1), "y": round(y,1), "size": round(span.get("size",10),1)})
    pages.append({"page": pnum+1, "texts": texts})
doc.close()
total = sum(len(p["texts"]) for p in pages)
is_vec = total > 5
print(json.dumps({"pages": pages, "is_vector": is_vec, "total_texts": total}))
`.trim();
    const sp = path.join(tmpDir, 'r.py');
    fs.writeFileSync(sp, script);
    const out = execSync(`python3 "${sp}"`, { timeout: 30000, maxBuffer: 20*1024*1024 });
    return JSON.parse(out.toString());
  } catch(e) {
    console.error('[PyMuPDF]', e.message);
    return null;
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch(e) {}
  }
}

// ════════════════════════════════════════════════════════════════
// READER 2: GCV — scanned PDF table extraction
// ════════════════════════════════════════════════════════════════
async function extractScannedPDF_GCV(pdfB64) {
  const gcvKey = process.env.GOOGLE_CLOUD_VISION_API_KEY;
  if (!gcvKey) { console.log('[GCV-PDF] No key'); return null; }
  try {
    const res = await fetch(`https://vision.googleapis.com/v1/files:annotate?key=${gcvKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: [{ inputConfig: { content: pdfB64, mimeType: 'application/pdf' }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }] }] }),
      signal: AbortSignal.timeout(60000)
    });
    if (!res.ok) return null;
    const gcvData = await res.json();
    const pages = [];
    for (const resp of (gcvData.responses || [])) {
      for (const pageAnnot of (resp.fullTextAnnotation?.pages || [])) {
        const cells = [];
        for (const block of (pageAnnot.blocks || [])) {
          for (const para of (block.paragraphs || [])) {
            const txt = (para.words || []).map(w => (w.symbols||[]).map(s=>s.text).join('')).join(' ').trim();
            if (txt) {
              const v = para.boundingBox?.vertices || [];
              cells.push({ text: txt, x: v[0]?.x||0, y: v[0]?.y||0, w: Math.abs((v[1]?.x||0)-(v[0]?.x||0)), h: Math.abs((v[2]?.y||0)-(v[0]?.y||0)) });
            }
          }
        }
        cells.sort((a,b) => a.y-b.y || a.x-b.x);
        const avgH = cells.reduce((s,c) => s+c.h, 0) / (cells.length||1);
        const thresh = Math.max(20, avgH*0.6);
        const rows = []; let curY=-1, curRow=[];
        for (const cell of cells) {
          if (Math.abs(cell.y-curY) > thresh) {
            if (curRow.length) { curRow.sort((a,b)=>a.x-b.x); rows.push(curRow.map(c=>c.text)); }
            curRow=[cell]; curY=cell.y;
          } else curRow.push(cell);
        }
        if (curRow.length) { curRow.sort((a,b)=>a.x-b.x); rows.push(curRow.map(c=>c.text)); }
        pages.push({ rows, text: rows.map(r=>r.join(' | ')).join('\n') });
      }
    }
    console.log(`[GCV-PDF] ${pages.length} pages extracted`);
    return pages.length ? { pages } : null;
  } catch(e) { console.error('[GCV-PDF]', e.message); return null; }
}

// ════════════════════════════════════════════════════════════════
// READER 3: GCV — image OCR (PNG/JPG drawings)
// ════════════════════════════════════════════════════════════════
async function extractImage_GCV(imageB64, mimeType='image/png') {
  const gcvKey = process.env.GOOGLE_CLOUD_VISION_API_KEY;
  if (!gcvKey) { console.log('[GCV-IMG] No key'); return null; }
  try {
    const res = await fetch(`https://vision.googleapis.com/v1/images:annotate?key=${gcvKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: [{ image: { content: imageB64 }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }] }] }),
      signal: AbortSignal.timeout(30000)
    });
    if (!res.ok) return null;
    const data = await res.json();
    const fullText = data.responses?.[0]?.fullTextAnnotation?.text || '';
    const words = data.responses?.[0]?.textAnnotations?.slice(1) || [];
    console.log(`[GCV-IMG] ${words.length} words extracted`);
    return fullText ? { full_text: fullText, word_count: words.length } : null;
  } catch(e) { console.error('[GCV-IMG]', e.message); return null; }
}

// ════════════════════════════════════════════════════════════════
// BUILD CONTEXT STRING — what Claude reads
// ════════════════════════════════════════════════════════════════
function buildContext(extracted) {
  const parts = [];

  // ── FIX 1: drawing_context_text from server.js PDF pipeline ──
  // This carries the full PyMuPDF spatial text extracted by buildDrawingContext()
  // It was being saved in cvData but never read here — now it is.
  if (extracted.drawing_context_text) {
    parts.push('=== PDF EXTRACTED TEXT (server pipeline — primary source) ===');
    parts.push(extracted.drawing_context_text);
  }

  if (extracted.pymupdf) {
    const d = extracted.pymupdf;
    parts.push(`=== PDF TEXT (PyMuPDF — ${d.total_texts} items) ===`);
    for (const page of (d.pages||[])) {
      parts.push(`-- Page ${page.page} --`);
      const byY = {};
      for (const t of page.texts) {
        const row = Math.round(t.y/12)*12;
        if (!byY[row]) byY[row] = [];
        byY[row].push(t);
      }
      for (const row of Object.keys(byY).sort((a,b)=>a-b)) {
        parts.push(byY[row].sort((a,b)=>a.x-b.x).map(t=>t.text).join('  '));
      }
    }
  }

  if (extracted.gcv_pdf) {
    parts.push(`=== SCANNED PDF (GCV — ${extracted.gcv_pdf.pages?.length} pages) ===`);
    for (const [i, page] of (extracted.gcv_pdf.pages||[]).entries()) {
      parts.push(`-- Page ${i+1} (columns pipe-separated) --`);
      parts.push(page.text);
    }
  }

  if (extracted.gcv_image) {
    parts.push('=== IMAGE OCR (GCV) ===');
    parts.push(extracted.gcv_image.full_text);
  }

  if (extracted.dxf_texts?.length) {
    parts.push(`=== DXF TEXTS (${extracted.dxf_texts.length}) ===`);
    parts.push(extracted.dxf_texts.slice(0,200).map(t=>t.text||t).join(' | '));
  }

  if (extracted.dxf_dimensions?.length) {
    parts.push('=== DXF DIMENSIONS ===');
    parts.push(extracted.dxf_dimensions.slice(0,80).map(d=>d.value_m+'m').join(', '));
  }

  if (extracted.dxf_layers?.length) {
    parts.push('=== DXF LAYERS ===');
    parts.push(extracted.dxf_layers.join(', '));
  }

  if (extracted.dxf_walls) {
    parts.push('=== WALL AREAS BY THICKNESS ===');
    parts.push(JSON.stringify(extracted.dxf_walls));
  }

  if (extracted.dxf_blocks) {
    parts.push('=== BLOCK COUNTS ===');
    parts.push(JSON.stringify(extracted.dxf_blocks));
  }

  return parts.join('\n');
}

// ════════════════════════════════════════════════════════════════
// CLAUDE — TEXT ONLY (zero image tokens)
// ════════════════════════════════════════════════════════════════
async function callClaudeTextOnly(context, filename, kb) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set');

  const prompt = `${kb ? kb+'\n\n' : ''}FILE: ${filename||'drawing'}

EXTRACTED DRAWING DATA:
${'='.repeat(60)}
${context || 'NO TEXT EXTRACTED — drawing could not be read'}
${'='.repeat(60)}

TASK:
1. From extracted text above: find project name, scale, drawing type, floor levels
2. Find legend/symbol table
3. Find schedule tables (column, footing, door, window) — copy cell values EXACTLY
4. Count elements: doors, windows, columns, lifts, staircases
5. Calculate areas and BOQ quantities
6. Apply Gujarat DSR 2025 rates from system prompt

RULE: Use ONLY values found in extracted text. Missing value = "not found in drawing".

Return ONLY raw JSON:
{
  "project_name": "",
  "drawing_type": "FLOOR_PLAN|SECTION|STRUCTURAL|FOUNDATION|SITE_PLAN|ROAD",
  "scale": "1:100",
  "scale_factor": 100,
  "legend": [{"symbol":"","meaning":"","layer":""}],
  "schedule_data": {
    "columns": [{"mark":"","size_mm":"","main_bars":"","stirrups":"","qty":0,"source":"drawing-schedule|not found"}],
    "footings": [{"mark":"","size_mm":"","depth_mm":"","main_bars":"","qty":0,"source":"drawing-schedule|not found"}]
  },
  "boq": [{"sr":1,"part":"PART A","description":"","unit":"CUM|SQM|RMT|NOS|KG","qty":0,"rate":0,"amount":0,"source":"drawing|calculated|assumed","confidence":"high|medium|low"}],
  "element_counts": {"door_count":0,"window_count":0,"column_count":0,"lift_count":0,"staircase_count":0,"floor_count":0},
  "area_statement": {"total_bua_sqmt":0,"floor_wise":[]},
  "cost_summary": {"civil_total_inr":0,"civil_total_lacs":0},
  "observations": [],
  "not_found": []
}`;

  const body = {
    model: 'claude-sonnet-4-6',
    max_tokens: 8192,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: prompt }]
  };

  for (let i = 0; i <= 3; i++) {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01' },
      body: JSON.stringify(body)
    });
    const data = await r.json();
    if (r.ok && data.content) return data.content.filter(b=>b.type==='text').map(b=>b.text).join('');
    if (data.error?.type !== 'overloaded_error') throw new Error(`Claude: ${data.error?.message}`);
    await new Promise(res => setTimeout(res, 2000*(i+1)));
  }
  throw new Error('Claude: max retries exceeded');
}

function parseJSON(raw) {
  if (!raw) return null;
  const clean = raw.replace(/```json|```/g,'').trim();
  const fb=clean.indexOf('{'), lb=clean.lastIndexOf('}');
  if (fb===-1||lb===-1) return null;
  try { return JSON.parse(clean.slice(fb,lb+1)); }
  catch(e) { console.error('[JSON parse]', e.message, clean.slice(0,200)); return null; }
}

// ════════════════════════════════════════════════════════════════
// MAIN EXPORT — same signature as before, server.js unchanged
// ════════════════════════════════════════════════════════════════
async function geminiAnalyzeDrawing(key, files, cvData, fetchFn) {
  console.log('\n[PMC] === Smart Pipeline v2: GCV+PyMuPDF reading, Claude text-only reasoning ===');

  const extracted = {};
  const filename = files?.[0]?.name || 'drawing';

  // ── PDF ───────────────────────────────────────────────────────
  const pdfFiles = (files||[]).filter(f => f.type==='application/pdf' || f.name?.match(/\.pdf$/i));
  if (pdfFiles.length > 0) {
    const b64 = pdfFiles[0].b64;
    console.log('[PMC] PDF detected — trying PyMuPDF...');
    const pymupdf = await extractVectorPDF(b64);
    if (pymupdf?.is_vector && pymupdf.total_texts > 5) {
      console.log(`[PMC] Vector PDF OK — ${pymupdf.total_texts} texts`);
      extracted.pymupdf = pymupdf;
    } else {
      console.log('[PMC] Not vector — trying GCV scanned PDF...');
      const gcv = await extractScannedPDF_GCV(b64);
      if (gcv) {
        extracted.gcv_pdf = gcv;
      } else {
        // ── Tesseract fallback: GCV key missing or GCV returned null ──
        console.log('[PMC] GCV unavailable — Tesseract fallback for scanned PDF...');
        try {
          const { execSync: _exec } = require('child_process');
          const _fs = require('fs');
          const _os = require('os');
          const _tmpDir = _fs.mkdtempSync(_os.tmpdir() + '/pmc_scan_');
          const _pdfPath = _tmpDir + '/input.pdf';
          _fs.writeFileSync(_pdfPath, Buffer.from(b64, 'base64'));
          const _rScript = `
import fitz, json
doc = fitz.open('${_pdfPath.replace(/\\/g, '/')}')
paths = []
for i in range(min(len(doc), 5)):
    pix = doc[i].get_pixmap(matrix=fitz.Matrix(200/72, 200/72), alpha=False)
    p = '${_tmpDir.replace(/\\/g, '/')}/page_{}.png'.format(i)
    pix.save(p)
    paths.append(p)
doc.close()
print(json.dumps(paths))
`.trim();
          _fs.writeFileSync(_tmpDir + '/r.py', _rScript);
          const _pngPaths = JSON.parse(_exec(`python3 "${_tmpDir}/r.py"`, { timeout: 30000 }).toString());
          const _tessPages = [];
          for (const _png of _pngPaths) {
            try {
              const _out = _exec(`tesseract "${_png}" stdout --oem 1 --psm 6 -l eng`, { timeout: 30000, maxBuffer: 5*1024*1024 });
              const _text = _out.toString().trim();
              if (_text.length > 20) {
                _tessPages.push({ rows: _text.split('\n').filter(l=>l.trim()).map(l=>[l]), text: _text });
                console.log(`[PMC-Tesseract] Page ${_tessPages.length}: ${_text.length} chars`);
              }
            } catch(e) { console.error('[PMC-Tesseract] page failed:', e.message); }
          }
          if (_tessPages.length) extracted.gcv_pdf = { pages: _tessPages, engine: 'tesseract' };
          try { _fs.rmSync(_tmpDir, { recursive: true }); } catch(e) {}
        } catch(e) { console.error('[PMC-Tesseract] Error:', e.message); }
      }
    }
  }

  // ── Images ────────────────────────────────────────────────────
  // PNG tiles (from buildDrawingContext or direct image upload)
  // Priority: GCV OCR → Tesseract fallback (if GCV key missing or fails)
  const imgFiles = (files||[]).filter(f => f.type?.startsWith('image/'));
  if (imgFiles.length > 0) {
    console.log(`[PMC] ${imgFiles.length} image tile(s) — running GCV OCR on all...`);
    const allTexts = [];
    const gcvAvailable = !!process.env.GOOGLE_CLOUD_VISION_API_KEY;

    for (const imgFile of imgFiles) {
      let tileText = null;

      // ── Try GCV first ──
      if (gcvAvailable) {
        const gcv = await extractImage_GCV(imgFile.b64, imgFile.type||'image/png');
        if (gcv?.full_text) tileText = gcv.full_text;
      }

      // ── Tesseract fallback: GCV unavailable OR GCV returned nothing ──
      if (!tileText) {
        try {
          const { execSync } = require('child_process');
          const fs = require('fs');
          const os = require('os');
          const tmpDir = fs.mkdtempSync(os.tmpdir() + '/pmc_tess_');
          const pngPath = `${tmpDir}/tile.png`;
          fs.writeFileSync(pngPath, Buffer.from(imgFile.b64, 'base64'));
          const out = execSync(
            `tesseract "${pngPath}" stdout --oem 1 --psm 6 -l eng`,
            { timeout: 30000, maxBuffer: 5 * 1024 * 1024 }
          );
          const tessText = out.toString().trim();
          if (tessText.length > 20) {
            tileText = tessText;
            console.log(`[PMC] Tesseract fallback: ${tessText.length} chars from tile`);
          }
          try { fs.rmSync(tmpDir, { recursive: true }); } catch(e) {}
        } catch(e) {
          console.error('[PMC] Tesseract fallback failed:', e.message);
        }
      }

      if (tileText) allTexts.push(tileText);
    }

    if (allTexts.length) {
      extracted.gcv_image = {
        full_text: allTexts.join('\n\n--- NEXT TILE ---\n\n'),
        word_count: allTexts.length
      };
      console.log(`[PMC] Image OCR: ${allTexts.length}/${imgFiles.length} tiles extracted (GCV=${gcvAvailable})`);
    }
  }

  // ── DXF data from cvData ──────────────────────────────────────
  if (cvData?.pdf_extracted_texts?.length) extracted.dxf_texts = cvData.pdf_extracted_texts;
  if (cvData?.texts?.length) extracted.dxf_texts = [...(extracted.dxf_texts||[]), ...cvData.texts];
  if (cvData?.dimensions?.length) extracted.dxf_dimensions = cvData.dimensions;
  if (cvData?.layers?.length) extracted.dxf_layers = cvData.layers;
  if (cvData?.wall_by_thickness_m2) extracted.dxf_walls = cvData.wall_by_thickness_m2;
  if (cvData?.block_counts) extracted.dxf_blocks = cvData.block_counts;

  // ── FIX 2: drawing_context_text from server.js buildDrawingContext() ──
  // server.js saves this in cvData but drawing_analyzer never read it — fixed now.
  if (cvData?.drawing_context_text) {
    extracted.drawing_context_text = cvData.drawing_context_text;
    console.log(`[PMC] drawing_context_text received: ${cvData.drawing_context_text.length} chars`);
  }

  // ── Build context ─────────────────────────────────────────────
  const context = buildContext(extracted);
  const sources = [extracted.drawing_context_text&&'ServerPipeline', extracted.pymupdf&&'PyMuPDF', extracted.gcv_pdf&&'GCV-PDF', extracted.gcv_image&&'GCV-Image', extracted.dxf_texts&&'ezdxf'].filter(Boolean);
  console.log(`[PMC] Context: ${context.length} chars from: ${sources.join(', ')||'none'}`);

  // ── Claude text-only ──────────────────────────────────────────
  const kb = knowledgeBaseHints();
  const raw = await callClaudeTextOnly(context, filename, kb);
  let boq = parseJSON(raw) || { boq:[], observations:['Parse failed: '+raw?.slice(0,100)], overall_confidence:'LOW' };

  boq = validateBOQ(boq);
  console.log(`[PMC] BOQ: ${boq.boq?.length||0} items | confidence: ${boq.overall_confidence} | sources: ${sources.join('+')}`);

  const totalInr = boq.cost_summary?.civil_total_inr || 0;
  return {
    project_name:     boq.project_name || 'CIVIL PROJECT',
    drawing_type:     boq.drawing_type || 'GENERAL',
    drawing_no:       boq.drawing_no   || '',
    scale:            boq.scale        || '',
    legend:           boq.legend       || [],
    schedule_data:    boq.schedule_data || {},
    boq_items:        boq.boq          || [],
    element_counts:   boq.element_counts || {},
    area_statement:   boq.area_statement || { total_bua_sqmt:0, floor_wise:[] },
    cost_summary:     { civil_total_inr:totalInr, civil_total_lacs:Math.round(totalInr/100000*100)/100, civil_total_crores:Math.round(totalInr/10000000*100)/100 },
    observations:     boq.observations || [],
    not_found:        boq.not_found || [],
    validation_warnings:      boq.validation_warnings || [],
    validation_passed:        boq.validation_passed || [],
    overall_confidence:       boq.overall_confidence || 'MEDIUM',
    engineer_action_required: boq.engineer_action_required || [],
    cv_analysis:      cvData || {},
    prepared_by:      'PMC Civil AI Agent',
    pipeline_info:    { reading_sources: sources, claude_mode: 'TEXT ONLY — zero image tokens' }
  };
}

// ── Validation ────────────────────────────────────────────────────
function validateBOQ(boq) {
  const warnings=[], passed=[];
  for (const item of (boq.boq||[])) {
    const qty=Number(item.qty)||0, rate=Number(item.rate)||0, amount=Number(item.amount)||0;
    if (qty>0&&rate>0&&amount>0&&Math.abs(qty*rate-amount)/amount>0.05)
      warnings.push({ item:item.description, check:'Math mismatch', severity:'HIGH' });
    if (qty===0&&amount>0)
      warnings.push({ item:item.description, check:'Qty=0 but amount>0', severity:'HIGH' });
  }
  const totalInr=boq.cost_summary?.civil_total_inr||0, totalArea=boq.area_statement?.total_bua_sqmt||0;
  if (totalInr>0&&totalArea>0) {
    const cpm=totalInr/totalArea;
    if (cpm<1000||cpm>8000) warnings.push({ item:'TOTAL COST', check:`Rs.${Math.round(cpm)}/sqmt out of range`, severity:'MEDIUM' });
    else passed.push(`Cost/sqmt Rs.${Math.round(cpm)} OK`);
  }
  if (!warnings.length) passed.push('All checks passed');
  return { ...boq, validation_warnings:warnings, validation_passed:passed, overall_confidence:warnings.filter(w=>w.severity==='HIGH').length>0?'LOW':warnings.length>2?'MEDIUM':'HIGH', engineer_action_required:warnings.filter(w=>w.severity==='HIGH').map(w=>w.item) };
}

// ── CV (unchanged) ────────────────────────────────────────────────
function runCVAnalysis(b64Image) {
  try {
    const tmp = path.join(os.tmpdir(), `drawing_cv_${Date.now()}.txt`);
    fs.writeFileSync(tmp, b64Image);
    const result = execSync(`python3 ${path.join(__dirname,'drawing_cv.py')} ${tmp}`, { timeout:30000 });
    fs.unlinkSync(tmp);
    return JSON.parse(result.toString());
  } catch(e) { console.error('[CV] failed:', e.message); return { error: e.message }; }
}

module.exports = { geminiAnalyzeDrawing, runCVAnalysis, RATES };
