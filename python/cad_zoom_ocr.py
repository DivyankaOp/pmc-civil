#!/usr/bin/env python3
"""
CAD-style zoom OCR (Civils.ai / AutoCAD-like reading).
- Render overview + zoomed crops (tables / title / sections)
- Local RapidOCR — no cloud tokens
- stdout JSON: { success, drawing_hints, regions:[{label,text,lines}], full_text }

Usage:
  python cad_zoom_ocr.py <input.pdf|png> [out_dir]
"""
from __future__ import annotations
import json, os, sys, re, tempfile

def render_pdf_regions(pdf_path, out_dir):
    import fitz
    from PIL import Image
    doc = fitz.open(pdf_path)
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    regions = {
        "overview": (0, 0, w, h, 110),
        # AutoCAD-like zooms: top band (often schedules/title), right tables, bottom title, mid sections
        "zoom_top": (0, 0, w, h * 0.38, 220),
        "zoom_right": (w * 0.55, 0, w, h * 0.70, 220),
        "zoom_left_detail": (0, h * 0.35, w * 0.55, h * 0.88, 220),
        "zoom_title": (w * 0.45, h * 0.72, w, h, 200),
        "zoom_center": (w * 0.20, h * 0.25, w * 0.80, h * 0.75, 200),
    }
    paths = []
    for name, (x1, y1, x2, y2, dpi) in regions.items():
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        clip = fitz.Rect(x1, y1, x2, y2)
        pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip)
        path = os.path.join(out_dir, f"{name}.png")
        pix.save(path)
        # Cap edge for OCR speed
        im = Image.open(path)
        if max(im.size) > 2200:
            im.thumbnail((2200, 2200), Image.LANCZOS)
            im.save(path, optimize=True)
        paths.append((name, path))
    doc.close()
    return paths

def render_image_regions(img_path, out_dir):
    from PIL import Image
    im = Image.open(img_path).convert("RGB")
    W, H = im.size
    regions = {
        "overview": (0, 0, W, H),
        "zoom_top": (0, 0, W, int(H * 0.38)),
        "zoom_right": (int(W * 0.55), 0, W, int(H * 0.70)),
        "zoom_left_detail": (0, int(H * 0.35), int(W * 0.55), int(H * 0.88)),
        "zoom_title": (int(W * 0.45), int(H * 0.72), W, H),
        "zoom_center": (int(W * 0.20), int(H * 0.25), int(W * 0.80), int(H * 0.75)),
    }
    paths = []
    for name, box in regions.items():
        crop = im.crop(box)
        if max(crop.size) > 2200:
            crop.thumbnail((2200, 2200), Image.LANCZOS)
        path = os.path.join(out_dir, f"{name}.png")
        crop.save(path, optimize=True)
        paths.append((name, path))
    return paths

def ocr_image(path, ocr):
    result, _ = ocr(path)
    lines = [r[1] for r in (result or []) if r and r[1]]
    # Drop obvious garbage short tokens
    clean = []
    for ln in lines:
        s = str(ln).strip()
        if len(s) < 2:
            continue
        clean.append(s)
    return clean

def detect_hints(text: str):
    t = text.lower()
    hints = []
    rules = [
        ("SECTION", r"\bsection\b|sec\.?\s*[a-z]-[a-z]|sectional"),
        ("ELEVATION", r"\belevation\b|\belev\b"),
        ("FLOOR_PLAN", r"floor plan|ground floor plan|typical floor"),
        ("FOUNDATION", r"footing|foundation layout|schedule of footing"),
        ("COLUMN_SCHEDULE", r"schedule of column|column schedule"),
        ("STRUCTURAL", r"r\.?c\.?c|beam schedule|slab schedule|reinforcement"),
        ("SITE_PLAN", r"site plan|key plan|layout plan"),
        ("ROAD", r"\bgsb\b|\bwmm\b|\bpqc\b|chainage"),
        ("DETAIL", r"\bdetail\b|typical detail"),
        ("TITLE_BLOCK", r"drawing no|drawn by|checked by|scale\s*1\s*:"),
    ]
    for name, pat in rules:
        if re.search(pat, t, re.I):
            hints.append(name)
    return hints or ["UNKNOWN"]

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"success": False, "error": "Usage: cad_zoom_ocr.py <pdf|png> [out_dir]"}))
        return 1
    src = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else tempfile.mkdtemp(prefix="pmc_cad_")
    os.makedirs(out_dir, exist_ok=True)

    ext = os.path.splitext(src)[1].lower()
    try:
        if ext == ".pdf":
            paths = render_pdf_regions(src, out_dir)
        elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
            paths = render_image_regions(src, out_dir)
        else:
            print(json.dumps({"success": False, "error": f"Unsupported for zoom OCR: {ext}. Convert DWG→DXF/PDF/PNG first."}))
            return 1

        from rapidocr_onnxruntime import RapidOCR
        ocr = RapidOCR()
        regions = []
        all_lines = []
        for label, path in paths:
            lines = ocr_image(path, ocr)
            regions.append({"label": label, "path": path, "lines": lines, "text": "\n".join(lines)})
            # Prefer zoomed regions over overview for table fidelity (skip duplicating overview into primary)
            if label != "overview":
                all_lines.extend(lines)
            else:
                all_lines.extend(lines[:80])

        full_text = "\n".join(dict.fromkeys(all_lines))  # de-dupe preserve order
        hints = detect_hints(full_text)
        print(json.dumps({
            "success": True,
            "engine": "rapidocr+cad_zoom",
            "drawing_hints": hints,
            "regions": [{"label": r["label"], "text": r["text"], "line_count": len(r["lines"])} for r in regions],
            "full_text": full_text,
            "char_count": len(full_text),
            "out_dir": out_dir,
        }, ensure_ascii=False))
        return 0
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}))
        return 1

if __name__ == "__main__":
    sys.exit(main() or 0)
