"""
PMC OCR Pipeline — Rasterized PDF / Scanned Drawing
=====================================================
Best-effort text extraction from scanned/rasterized engineering drawings.

Strategy:
  1. Render PDF pages at 400 DPI using PyMuPDF
  2. For each page: extract 6 crops (full + schedule zones)
  3. Preprocess each crop: grayscale → upscale → adaptive binarize → denoise
  4. Run Tesseract PSM 6 (table mode) + PSM 11 (sparse) on each crop
  5. Merge & deduplicate all lines
  6. Output structured JSON for server.js to consume

Usage:
  python3 ocr_pipeline.py <pdf_path> <output_json_path>
  python3 ocr_pipeline.py <pdf_path> -   (prints JSON to stdout)

Output JSON:
  {
    "pages": [
      {
        "page_num": 1,
        "raw_text": "...",
        "table_rows": [["line1"], ["line2"], ...],
        "is_rotated": false,
        "crops_processed": 6,
        "chars_extracted": 1234
      }
    ],
    "engine": "tesseract-opencv-pipeline",
    "total_chars": 5678
  }
"""

import sys
import os
import json
import tempfile
import subprocess

try:
    import fitz  # PyMuPDF
except ImportError:
    print(json.dumps({"error": "PyMuPDF not installed. Run: pip install pymupdf --break-system-packages"}))
    sys.exit(1)

try:
    import cv2
    import numpy as np
except ImportError:
    print(json.dumps({"error": "OpenCV not installed. Run: pip install opencv-python --break-system-packages"}))
    sys.exit(1)


# ── CONFIG ────────────────────────────────────────────────────────
RENDER_DPI      = 400      # PDF → PNG render resolution
MIN_WIDTH_PX    = 3000     # upscale if narrower (ensures small text readable)
ADAPTIVE_BLOCK  = 31       # adaptive threshold block size
ADAPTIVE_C      = 10       # adaptive threshold constant
DENOISE_H       = 10       # NlMeans denoising strength
MIN_LINE_LEN    = 2        # discard lines shorter than this
MAX_PAGES       = 999      # UNLOCKED: process all pages (was 3)


# ── PREPROCESS ───────────────────────────────────────────────────
def preprocess(img_bgr):
    """
    Grayscale → upscale → adaptive binarize → denoise.
    Handles: uneven scan lighting, skewed text, thin lines.
    """
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

    # Upscale to ensure min width for small text readability
    h, w = gray.shape
    if w < MIN_WIDTH_PX:
        scale = MIN_WIDTH_PX / w
        gray = cv2.resize(gray, None, fx=scale, fy=scale,
                          interpolation=cv2.INTER_CUBIC)

    # Adaptive binarize — handles shadows/stains on scanned drawings
    binary = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        blockSize=ADAPTIVE_BLOCK,
        C=ADAPTIVE_C
    )

    # Denoise — removes scan artifacts that confuse Tesseract
    binary = cv2.fastNlMeansDenoising(binary, h=DENOISE_H)

    return binary


# ── CROP ZONES ───────────────────────────────────────────────────
def get_crops(img_bgr):
    """
    Engineering drawings: schedule tables appear in predictable zones.
    Returns list of (label, crop_array).
    """
    h, w = img_bgr.shape[:2]
    crops = []

    # 1. Full page — always process (catches anything missed by crops)
    crops.append(('full',         img_bgr))

    # 2. Bottom half — schedules almost always below centerline
    crops.append(('bottom_half',  img_bgr[h // 2:, :]))

    # 3. Bottom-right quadrant — most common schedule location (Surat industrial)
    crops.append(('btm_right',    img_bgr[h // 2:, w // 2:]))

    # 4. Bottom-left quadrant — section details, notes
    crops.append(('btm_left',     img_bgr[h // 2:, :w // 2]))

    # 5. Right third — column/footing schedule sometimes on far right strip
    crops.append(('right_third',  img_bgr[:, 2 * w // 3:]))

    # 6. Bottom strip (bottom 20%) — title block + notes box
    crops.append(('btm_strip',    img_bgr[4 * h // 5:, :]))

    return [(lbl, c) for lbl, c in crops if c.size > 0]


# ── TESSERACT RUNNER ─────────────────────────────────────────────
def run_tess(img_array, psm):
    """Run Tesseract on a numpy array, return text string."""
    tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    try:
        cv2.imwrite(tmp.name, img_array)
        tmp.close()
        r = subprocess.run(
            ['tesseract', tmp.name, 'stdout',
             '--oem', '1', '--psm', str(psm), '-l', 'eng'],
            capture_output=True, text=True, timeout=45
        )
        return r.stdout.strip()
    except Exception:
        return ''
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass


# ── PER-PAGE OCR ─────────────────────────────────────────────────
def ocr_page(png_path):
    """
    Full OCR pipeline for one rendered page PNG.
    Returns dict with raw_text, table_rows, crops_processed.
    """
    img = cv2.imread(png_path)
    if img is None:
        return {'raw_text': '', 'table_rows': [], 'crops_processed': 0, 'chars_extracted': 0}

    seen   = set()
    lines  = []
    crops_done = 0

    for label, crop in get_crops(img):
        processed = preprocess(crop)
        crops_done += 1

        # PSM 6: assume uniform block of text — best for schedule tables
        for psm in (6, 11):
            text = run_tess(processed, psm)
            for line in text.split('\n'):
                line = line.strip()
                if len(line) < MIN_LINE_LEN:
                    continue
                if line not in seen:
                    seen.add(line)
                    lines.append(line)

    raw = '\n'.join(lines)
    return {
        'raw_text':       raw,
        'table_rows':     [[l] for l in lines],
        'is_rotated':     False,
        'crops_processed': crops_done,
        'chars_extracted': len(raw)
    }


# ── MAIN ─────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 3:
        print(json.dumps({'error': 'Usage: ocr_pipeline.py <pdf_path> <output_json|->'  }))
        sys.exit(1)

    pdf_path    = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(pdf_path):
        print(json.dumps({'error': f'File not found: {pdf_path}'}))
        sys.exit(1)

    # Render PDF pages to PNG
    tmp_dir  = tempfile.mkdtemp(prefix='pmc_ocr_')
    png_paths = []

    try:
        doc = fitz.open(pdf_path)
        mat = fitz.Matrix(RENDER_DPI / 72, RENDER_DPI / 72)

        for i in range(min(len(doc), MAX_PAGES)):
            page = doc[i]
            pix  = page.get_pixmap(matrix=mat, alpha=False)
            p    = os.path.join(tmp_dir, f'page_{i}.png')
            pix.save(p)
            png_paths.append(p)

        doc.close()
    except Exception as e:
        print(json.dumps({'error': f'PDF render failed: {e}'}))
        sys.exit(1)

    # OCR each page
    pages      = []
    total_chars = 0

    for i, png_path in enumerate(png_paths):
        result = ocr_page(png_path)
        result['page_num'] = i + 1
        pages.append(result)
        total_chars += result['chars_extracted']
        print(f'[OCR] Page {i+1}: {result["chars_extracted"]} chars, '
              f'{result["crops_processed"]} crops', file=sys.stderr)

    # Cleanup temp PNGs
    for p in png_paths:
        try:
            os.unlink(p)
        except Exception:
            pass
    try:
        os.rmdir(tmp_dir)
    except Exception:
        pass

    output = {
        'pages':        pages,
        'engine':       'tesseract-opencv-pipeline',
        'total_chars':  total_chars
    }

    if output_path == '-':
        print(json.dumps(output))
    else:
        with open(output_path, 'w') as f:
            json.dump(output, f)
        print(json.dumps({'success': True, 'output': output_path,
                          'total_chars': total_chars}))


if __name__ == '__main__':
    main()
