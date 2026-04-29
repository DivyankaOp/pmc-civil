#!/usr/bin/env python3
"""
PMC Civil — DWG/DXF Converter (v5 — ZWCAD + Multi-Sheet)
Strategy (first-that-works wins for PNG):
  1. DXF -> PNG via ezdxf + matplotlib (ALL layouts/sheets)
  2. DXF/DWG -> PNG via LibreOffice --headless
  3. DWG binary text extraction (ZWCAD AC1032 fallback — text-only mode)

New in v5:
  - Multi-sheet: renders all paperspace layouts, not just modelspace
  - XREF expansion: text inside INSERT blocks is extracted too
  - ZWCAD text-only mode: rich text context when PNG render fails
  - Sheet names + XREF references returned for Claude context
"""
import sys, json, os, traceback, subprocess, tempfile, shutil, re, struct

def extract_text_from_dwg_binary(dwg_path):
    """Extract ALL readable strings from DWG binary — ZWCAD + AutoCAD compatible."""
    try:
        with open(dwg_path, "rb") as f:
            data = f.read()
    except Exception as e:
        return {"texts": [], "layers": [], "dims": [], "sheets": [], "xrefs": [], "error": str(e)}

    version = data[:6].decode("ascii", errors="replace")
    results = {"version": version, "texts": [], "layers": [], "dims": [], "sheets": [], "xrefs": []}

    # UTF-16LE strings (ZWCAD/AutoCAD R2013+ encoding)
    utf16_strings = []
    i = 0
    while i < len(data) - 3:
        if 32 <= data[i] <= 126 and data[i + 1] == 0:
            j = i
            chars = []
            while j < len(data) - 1 and data[j + 1] == 0 and 32 <= data[j] <= 126:
                chars.append(chr(data[j]))
                j += 2
            if len(chars) >= 3:
                s = "".join(chars).strip()
                if s:
                    utf16_strings.append(s)
            i = j + 2
        else:
            i += 1

    # ASCII strings
    ascii_strings = []
    i = 0
    current = []
    while i < len(data):
        b = data[i]
        if 32 <= b <= 126:
            current.append(chr(b))
        else:
            if len(current) >= 4:
                s = "".join(current).strip()
                if s:
                    ascii_strings.append(s)
            current = []
        i += 1

    all_strings = list(dict.fromkeys(utf16_strings + ascii_strings))

    layer_pat = re.compile(r'^[A-Z0-9_\-\.]{2,30}$')
    eng_pat = re.compile(
        r'(FOOTING|COLUMN|COL|BEAM|SLAB|RCC|GRID|LEVEL|FLOOR|WALL|'
        r'SECTION|DETAIL|PLAN|ELEVATION|SCHEDULE|REINFORCEMENT|'
        r'STIRRUP|MAIN.?BAR|TIE|LINK|DIA|THK|WIDTH|DEPTH|HEIGHT|'
        r'FOUNDATION|PILE|RAFT|GRADE|M\d0|Fe\d{3}|'
        r'NOTES?|SPEC|DESCRIPTION|DRAWING|TITLE|PROJECT|'
        r'ROAD|GSB|WMM|PQC|KERB|DRAIN|CULVERT|BRIDGE|'
        r'PLINTH|LINTEL|PARAPET|STAIRCASE|LIFT|RAMP|'
        r'\d+[xX]\d+|\d+mm|\d+\.\d+m|\d+ MM)',
        re.IGNORECASE
    )
    dim_pat = re.compile(r'^\d+(\.\d+)?$|^\d+[xX]\d+$')
    sheet_pat = re.compile(
        r'^(Layout\s*\d+|Sheet[\-\s]*\d+|Model|Plan[\-\s]*\d+|Drawing[\-\s]*\d+|'
        r'GF|FF|SF|TF|RF|Basement|Ground\s*Floor|First\s*Floor|Site\s*Plan|'
        r'[A-Z]{1,3}-\d{1,4})$',
        re.IGNORECASE
    )

    for s in all_strings:
        s = s.strip()
        if not s or len(s) < 2:
            continue
        if re.search(r'[^\x20-\x7E]', s):
            continue
        if re.search(r'[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}', s):
            continue
        if re.search(r'\\.*\.(dwg|dxf)', s, re.IGNORECASE):
            results["xrefs"].append(s.split("\\")[-1])
            continue
        if sheet_pat.match(s):
            results["sheets"].append(s)
        elif eng_pat.search(s):
            results["texts"].append({"text": s, "source": "binary_extract"})
        elif layer_pat.match(s) and len(s) >= 3:
            results["layers"].append(s)
        elif dim_pat.match(s):
            results["dims"].append(s)

    for s in utf16_strings:
        s = s.strip()
        if len(s) >= 2 and not any(t["text"] == s for t in results["texts"]):
            if re.match(r'^[A-Z][0-9A-Z\-_\.]+$', s) and len(s) <= 20:
                results["texts"].append({"text": s, "source": "utf16_label"})

    seen = set()
    unique_texts = []
    for t in results["texts"]:
        if t["text"] not in seen:
            seen.add(t["text"])
            unique_texts.append(t)

    results["texts"] = unique_texts[:400]
    results["layers"] = list(dict.fromkeys(results["layers"]))[:120]
    results["dims"] = list(dict.fromkeys(results["dims"]))[:120]
    results["sheets"] = list(dict.fromkeys(results["sheets"]))[:50]
    results["xrefs"] = list(dict.fromkeys(results["xrefs"]))[:20]
    return results


def render_dxf_to_png(dxf_path, png_path, dpi=120, tiled=False):
    """Render DXF → PNG (modelspace + ALL paperspace layouts = multi-sheet support)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import ezdxf
    from ezdxf import recover
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

    try:
        doc, _ = recover.readfile(dxf_path)
    except Exception:
        try:
            doc = ezdxf.readfile(dxf_path)
        except Exception as e:
            return False, [], []

    layout_images = []

    def render_one(layout, out_path):
        try:
            fig = plt.figure(figsize=(16, 12), dpi=dpi)
            ax = fig.add_axes([0, 0, 1, 1])
            ctx = RenderContext(doc)
            out = MatplotlibBackend(ax)
            Frontend(ctx, out).draw_layout(layout, finalize=True)
            fig.savefig(out_path, dpi=dpi, bbox_inches="tight", facecolor="white")
            plt.close(fig)
            return os.path.exists(out_path) and os.path.getsize(out_path) > 5000
        except Exception:
            plt.close('all')
            return False

    # Modelspace
    render_one(doc.modelspace(), png_path)
    if os.path.exists(png_path) and os.path.getsize(png_path) > 5000:
        layout_images.append({"name": "ModelSpace", "path": png_path})

    # All paperspace layouts (multi-sheet)
    base_dir = os.path.dirname(png_path) or tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(png_path))[0]
    for layout in doc.layouts:
        if layout.name.upper() == "MODEL":
            continue
        sheet_png = os.path.join(base_dir, f"{base_name}_sheet_{re.sub(r'[^A-Za-z0-9]','_',layout.name)}.png")
        if render_one(layout, sheet_png):
            layout_images.append({"name": layout.name, "path": sheet_png})

    # If model failed but sheets worked, copy first sheet as main
    if not (os.path.exists(png_path) and os.path.getsize(png_path) > 5000) and layout_images:
        shutil.copy(layout_images[0]["path"], png_path)

    main_ok = os.path.exists(png_path) and os.path.getsize(png_path) > 5000

    # Tiles from main PNG
    tiles = []
    if tiled and main_ok:
        try:
            from PIL import Image
            img = Image.open(png_path)
            W, H = img.size
            for row in range(2):
                for col in range(2):
                    x1, y1 = col * W // 2, row * H // 2
                    x2, y2 = (col+1) * W // 2, (row+1) * H // 2
                    tile = img.crop((x1, y1, x2, y2))
                    tp = png_path.replace(".png", f"_tile_{row}{col}.png")
                    tile.save(tp)
                    tiles.append({"path": tp, "row": row, "col": col,
                                  "position": f"{'Top' if row==0 else 'Bottom'}-{'Left' if col==0 else 'Right'}"})
        except Exception:
            pass

    return main_ok, tiles, layout_images


def libreoffice_to_png(input_path, output_png):
    out_dir = os.path.dirname(output_png) or tempfile.gettempdir()
    base = os.path.splitext(os.path.basename(input_path))[0]
    lo_png = os.path.join(out_dir, base + ".png")
    subprocess.run(["libreoffice", "--headless", "--convert-to", "png",
                    "--outdir", out_dir, input_path],
                   capture_output=True, text=True, timeout=160)
    if os.path.exists(lo_png):
        if lo_png != output_png:
            shutil.move(lo_png, output_png)
        return True
    return False


def dwg_to_dxf_via_oda(dwg_path):
    try:
        from ezdxf.addons import odafc
        out_dxf = os.path.join(tempfile.gettempdir(),
                               os.path.splitext(os.path.basename(dwg_path))[0] + "_conv.dxf")
        odafc.convert(dwg_path, out_dxf, version="R2018")
        return out_dxf if os.path.exists(out_dxf) else None
    except Exception:
        return None


def extract_ezdxf_meta(dxf_path):
    import ezdxf
    from ezdxf import recover
    try:
        doc, _ = recover.readfile(dxf_path)
    except Exception:
        doc = ezdxf.readfile(dxf_path)

    msp = doc.modelspace()
    layers = [l.dxf.name for l in doc.layers if l.dxf.name != "0"]
    sheet_names = [layout.name for layout in doc.layouts if layout.name.upper() != "MODEL"]

    texts = []
    visited_blocks = set()

    def collect_texts(entities):
        for e in entities:
            try:
                if e.dxftype() in ("TEXT", "ATTDEF", "ATTRIB"):
                    t = e.dxf.text.strip()
                    if t:
                        pos = e.dxf.insert
                        texts.append({"text": t, "layer": e.dxf.layer,
                                      "x": round(pos.x, 3), "y": round(pos.y, 3)})
                elif e.dxftype() == "MTEXT":
                    t = e.plain_mtext().strip()
                    if t:
                        pos = e.dxf.insert
                        texts.append({"text": t, "layer": e.dxf.layer,
                                      "x": round(pos.x, 3), "y": round(pos.y, 3)})
                elif e.dxftype() == "INSERT":
                    bname = e.dxf.name
                    if bname not in visited_blocks:
                        visited_blocks.add(bname)
                        try:
                            collect_texts(doc.blocks[bname])
                        except Exception:
                            pass
            except Exception:
                pass

    collect_texts(msp)
    # Collect from all paper-space layouts (multi-sheet)
    for layout in doc.layouts:
        if layout.name.upper() != "MODEL":
            collect_texts(layout)

    dims = []
    for e in msp:
        try:
            if e.dxftype() == "DIMENSION":
                val = None
                try:
                    val = round(e.dxf.actual_measurement, 4)
                except Exception:
                    pass
                txt = ""
                try:
                    txt = e.dxf.text.strip()
                except Exception:
                    pass
                dims.append({"value": val, "text": txt, "layer": e.dxf.layer})
        except Exception:
            pass

    all_text = " ".join(t["text"] for t in texts)
    scale = "Not detected"
    m = re.search(r'(?:scale|sc)[:\s]*1\s*[:/]\s*(\d+)', all_text, re.IGNORECASE)
    if not m:
        m = re.search(r'\b1\s*:\s*(\d+)\b', all_text)
    if m:
        scale = f"1:{m.group(1)}"

    combined = (" ".join(layers) + " " + all_text).upper()
    dtype = "UNKNOWN"
    if any(k in combined for k in ["ROAD", "GSB", "WMM", "PQC", "CARRIAGEWAY", "KERB"]):
        dtype = "ROAD_PLAN"
    elif any(k in combined for k in ["FOUNDATION", "FOOTING", "PILE", "RAFT"]):
        dtype = "FOUNDATION"
    elif any(k in combined for k in ["SLAB", "BEAM", "COLUMN", "RCC", "STRUCTURAL"]):
        dtype = "STRUCTURAL"
    elif any(k in combined for k in ["FLOOR", "ROOM", "TOILET", "KITCHEN", "LIVING"]):
        dtype = "FLOOR_PLAN"
    elif any(k in combined for k in ["SITE", "LAYOUT", "PLOT", "BOUNDARY", "MASTER"]):
        dtype = "SITE_LAYOUT"
    elif any(k in combined for k in ["SECTION", "ELEVATION", "ELEV", "CROSS"]):
        dtype = "SECTION"

    return {"layers": layers, "texts": texts[:300], "dimensions": dims[:200],
            "scale": scale, "drawing_type": dtype, "sheets": sheet_names}


def run(input_path, output_png, dpi=120, tiled=False):
    result = {
        "success": False, "png_path": None, "tiles": [],
        "layout_images": [], "texts": [], "dimensions": [],
        "layers": [], "sheets": [], "xrefs": [],
        "drawing_type": "UNKNOWN", "scale": "Not detected",
        "extents": {}, "errors": [], "binary_extract": None,
        "zwcad_text_mode": False,
    }

    ext = os.path.splitext(input_path)[1].lower()
    dxf_for_meta = input_path if ext == ".dxf" else None

    # STEP 0: Binary text extraction (always — ZWCAD fallback context)
    if ext in (".dwg", ".dxf"):
        try:
            bin_result = extract_text_from_dwg_binary(input_path)
            result["binary_extract"] = bin_result
            if bin_result.get("texts"):
                result["texts"] = bin_result["texts"]
            if bin_result.get("layers"):
                result["layers"] = bin_result["layers"]
            if bin_result.get("dims"):
                result["dimensions"] = [{"value": None, "text": d, "layer": "binary"} for d in bin_result["dims"]]
            if bin_result.get("sheets"):
                result["sheets"] = bin_result["sheets"]
            if bin_result.get("xrefs"):
                result["xrefs"] = bin_result["xrefs"]
            print(f"[binary_extract] texts={len(result['texts'])} layers={len(result['layers'])} "
                  f"sheets={len(result['sheets'])} xrefs={len(result['xrefs'])}", file=sys.stderr)
        except Exception as e:
            result["errors"].append(f"Binary extract failed: {e}")

    # STEP 1: ODA DWG→DXF
    if ext == ".dwg":
        conv = dwg_to_dxf_via_oda(input_path)
        if conv:
            dxf_for_meta = conv
            result["errors"].append("ODA conversion successful")

    # STEP 2: DXF→PNG (multi-sheet)
    if dxf_for_meta:
        try:
            ok, tiles, layout_images = render_dxf_to_png(dxf_for_meta, output_png, dpi=dpi, tiled=tiled)
            if ok:
                result["png_path"] = output_png
                result["tiles"] = tiles
                result["layout_images"] = [{"name": li["name"], "path": li["path"]} for li in layout_images]
                result["success"] = True
                if len(layout_images) > 1:
                    result["errors"].append(
                        f"Multi-sheet drawing: {len(layout_images)} layouts rendered — "
                        + ", ".join(li["name"] for li in layout_images))
        except Exception as e:
            result["errors"].append(f"ezdxf render failed: {e}")

    # STEP 3: LibreOffice fallback
    if not result["png_path"]:
        try:
            if libreoffice_to_png(input_path, output_png):
                result["png_path"] = output_png
                result["success"] = True
        except FileNotFoundError:
            result["errors"].append("LibreOffice not installed")
        except subprocess.TimeoutExpired:
            result["errors"].append("LibreOffice timed out")
        except Exception as e:
            result["errors"].append(f"LibreOffice failed: {e}")

    # STEP 4: ezdxf metadata
    if dxf_for_meta:
        try:
            meta = extract_ezdxf_meta(dxf_for_meta)
            if meta.get("texts"):
                existing = set(t["text"] for t in result["texts"])
                for t in meta["texts"]:
                    if t["text"] not in existing:
                        result["texts"].append(t)
                        existing.add(t["text"])
            if meta.get("layers"):
                result["layers"] = list(set(result["layers"]) | set(meta["layers"]))
            if meta.get("dimensions") and not result["dimensions"]:
                result["dimensions"] = meta["dimensions"]
            result["scale"] = meta.get("scale", result["scale"])
            result["drawing_type"] = meta.get("drawing_type", result["drawing_type"])
            if meta.get("sheets"):
                result["sheets"] = list(set(result["sheets"]) | set(meta["sheets"]))
            result["success"] = True
        except Exception as e:
            result["errors"].append(f"ezdxf metadata failed: {e}")

    # STEP 5: ZWCAD text-only mode
    if not result["success"] and (result["texts"] or result["layers"]):
        result["success"] = True
        result["zwcad_text_mode"] = True
        sheet_info = (f" Sheets: {', '.join(result['sheets'][:10])}." if result["sheets"] else "")
        xref_info = (f" XREFs: {', '.join(result['xrefs'][:5])}." if result["xrefs"] else "")
        result["errors"].append(
            f"ZWCAD DWG: PNG render failed (format incompatible with ezdxf without ODA). "
            f"Extracted {len(result['texts'])} texts, {len(result['layers'])} layers from binary.{sheet_info}{xref_info} "
            "Claude will analyze via text-context mode. For visual accuracy, export to PDF/PNG from ZWCAD."
        )

    print(json.dumps(result))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: dwg_converter.py <input> <output.png> [dpi] [tiled]", "success": False}))
        sys.exit(1)
    try:
        dpi_arg = int(sys.argv[3]) if len(sys.argv) > 3 else 120
        tiled_arg = sys.argv[4].lower() in ("true", "1", "yes") if len(sys.argv) > 4 else False
        run(sys.argv[1], sys.argv[2], dpi=dpi_arg, tiled=tiled_arg)
    except Exception as e:
        print(json.dumps({"success": False, "errors": [f"Fatal: {e}", traceback.format_exc()[:1000]]}))
