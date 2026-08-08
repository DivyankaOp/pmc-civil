#!/usr/bin/env python3
"""
CAD-style zoom OCR (Civils.ai / AutoCAD-like reading).
- Multi-page PDF support
- Adaptive crops around SCHEDULE / table-dense regions
- Keeps OCR bounding boxes for spatial table rebuild
- Local RapidOCR — no cloud tokens

Usage:
  python cad_zoom_ocr.py <input.pdf|png> [out_dir]
"""
from __future__ import annotations
import json, os, sys, re, tempfile

MAX_PAGES = 5
SCHEDULE_RE = re.compile(
    r"schedule\s*of|footing\s*schedule|column\s*schedule|beam\s*schedule|"
    r"door\s*schedule|window\s*schedule|base\s*plate|mark\s+size|qty|nos\.?",
    re.I,
)


def _save_clip(page, clip, dpi, path):
    import fitz
    from PIL import Image
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip)
    pix.save(path)
    im = Image.open(path)
    if max(im.size) > 2400:
        im.thumbnail((2400, 2400), Image.LANCZOS)
        im.save(path, optimize=True)
    return path


def _find_schedule_clips(page, w, h):
    """Use vector text hits to zoom schedule regions; fallback fixed bands."""
    import fitz
    clips = []
    try:
        blocks = page.get_text("blocks") or []
    except Exception:
        blocks = []
    hits = []
    for b in blocks:
        if len(b) < 5:
            continue
        txt = str(b[4] or "")
        if SCHEDULE_RE.search(txt) or re.search(r"\b(MARK|SIZE|QTY|NOS|DEPTH)\b", txt, re.I):
            hits.append(fitz.Rect(b[0], b[1], b[2], b[3]))
    # Merge nearby hits into larger table crops
    for r in hits[:8]:
        pad_x, pad_y = w * 0.08, h * 0.12
        clips.append(fitz.Rect(
            max(0, r.x0 - pad_x),
            max(0, r.y0 - pad_y * 0.3),
            min(w, r.x1 + w * 0.55),
            min(h, r.y1 + h * 0.45),
        ))
    return clips


def render_pdf_regions(pdf_path, out_dir):
    import fitz
    doc = fitz.open(pdf_path)
    paths = []
    n_pages = min(len(doc), MAX_PAGES)
    for pi in range(n_pages):
        page = doc[pi]
        w, h = page.rect.width, page.rect.height
        prefix = f"p{pi + 1}"
        # Overview per page (lower DPI)
        ov = os.path.join(out_dir, f"{prefix}_overview.png")
        _save_clip(page, page.rect, 100 if n_pages > 1 else 110, ov)
        paths.append((f"{prefix}_overview", ov, pi))

        adaptive = _find_schedule_clips(page, w, h)
        if adaptive:
            for ai, clip in enumerate(adaptive[:3]):
                path = os.path.join(out_dir, f"{prefix}_sched{ai}.png")
                _save_clip(page, clip, 200, path)
                paths.append((f"{prefix}_sched{ai}", path, pi))
        else:
            # Scanned CAD: fewer high-value crops (speed) — schedules usually top/right
            import fitz as _fitz
            bands = {
                "zoom_top_right": (w * 0.42, 0, w, h * 0.55, 200),
                "zoom_mid_right": (w * 0.48, h * 0.20, w, h * 0.78, 200),
                "zoom_left_detail": (0, h * 0.35, w * 0.50, h * 0.92, 190),
            }
            for name, (x1, y1, x2, y2, dpi) in bands.items():
                path = os.path.join(out_dir, f"{prefix}_{name}.png")
                _save_clip(page, _fitz.Rect(x1, y1, x2, y2), dpi, path)
                paths.append((f"{prefix}_{name}", path, pi))
    doc.close()
    return paths


def render_image_regions(img_path, out_dir):
    from PIL import Image
    im = Image.open(img_path).convert("RGB")
    W, H = im.size
    regions = {
        "overview": (0, 0, W, H),
        "zoom_top": (0, 0, W, int(H * 0.40)),
        "zoom_right": (int(W * 0.50), 0, W, int(H * 0.72)),
        "zoom_left": (0, int(H * 0.30), int(W * 0.55), int(H * 0.90)),
        "zoom_center": (int(W * 0.15), int(H * 0.20), int(W * 0.85), int(H * 0.80)),
        "zoom_title": (int(W * 0.40), int(H * 0.70), W, H),
    }
    paths = []
    for name, box in regions.items():
        crop = im.crop(box)
        if max(crop.size) > 2400:
            crop.thumbnail((2400, 2400), Image.LANCZOS)
        path = os.path.join(out_dir, f"{name}.png")
        crop.save(path, optimize=True)
        paths.append((name, path, 0))
    return paths


def ocr_image(path, ocr):
    """Return lines + boxes [{text, box, conf}]."""
    result, _ = ocr(path)
    lines = []
    boxes = []
    for r in (result or []):
        if not r:
            continue
        # RapidOCR: [box, text, score]
        box, text, score = None, None, None
        if isinstance(r, (list, tuple)) and len(r) >= 2:
            box, text = r[0], r[1]
            score = r[2] if len(r) > 2 else None
        if not text:
            continue
        s = str(text).strip()
        if len(s) < 2:
            continue
        lines.append(s)
        boxes.append({
            "text": s,
            "box": box,
            "conf": float(score) if score is not None else None,
        })
    return lines, boxes


def detect_hints(text: str):
    t = text.lower()
    hints = []
    rules = [
        ("FOUNDATION_FOOTING", r"schedule of footing|footing schedule|foundation layout|column footing"),
        ("COLUMN_SCHEDULE", r"schedule of column|column schedule"),
        ("SECTION", r"\bsection\b|sec\.?\s*[a-z]-[a-z]|sectional"),
        ("ELEVATION", r"\belevation\b|\belev\b"),
        ("FLOOR_PLAN", r"floor plan|ground floor plan|typical floor"),
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
        elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"):
            paths = render_image_regions(src, out_dir)
        else:
            print(json.dumps({
                "success": False,
                "error": f"Unsupported for zoom OCR: {ext}. Convert DWG→DXF/PDF/PNG first.",
            }))
            return 1

        from rapidocr_onnxruntime import RapidOCR
        ocr = RapidOCR()
        regions = []
        all_lines = []
        all_boxes = []
        for label, path, page_idx in paths:
            lines, boxes = ocr_image(path, ocr)
            # Offset boxes by page index so spatial merge can separate pages
            for b in boxes:
                b["page"] = page_idx + 1
                b["region"] = label
            regions.append({
                "label": label,
                "path": path,
                "page": page_idx + 1,
                "lines": lines,
                "text": "\n".join(lines),
                "box_count": len(boxes),
            })
            if "overview" in label:
                all_lines.extend(lines[:100])
            else:
                all_lines.extend(lines)
            all_boxes.extend(boxes)

        # Prefer schedule-region text order; de-dupe
        full_text = "\n".join(dict.fromkeys(all_lines))
        hints = detect_hints(full_text)
        # ensure_ascii=True → safe for Windows Node stdout parsers
        payload = {
            "success": True,
            "engine": "rapidocr+cad_zoom",
            "drawing_hints": hints,
            "pages_processed": len({p[2] for p in paths}),
            "regions": [{
                "label": r["label"],
                "page": r["page"],
                "text": r["text"],
                "line_count": len(r["lines"]),
                "box_count": r["box_count"],
            } for r in regions],
            "boxes": all_boxes[:4000],
            "full_text": full_text,
            "char_count": len(full_text),
            "out_dir": out_dir,
        }
        out_json = os.path.join(out_dir, "ocr_result.json")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=True)
        payload["result_path"] = out_json
        print(json.dumps(payload, ensure_ascii=True))
        return 0
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=True))
        return 1


if __name__ == "__main__":
    sys.exit(main() or 0)
