#!/usr/bin/env python3
"""
PMC Civil — DWG/DXF Converter (v4)
Strategy (in order, first that works wins for PNG):
  1. DXF -> PNG via ezdxf + matplotlib  (no external binary required)
  2. DXF/DWG -> PNG via LibreOffice --headless  (if installed)
  3. DWG -> DXF via ezdxf.addons.odafc   (if ODA File Converter installed)
  4. [NEW] DWG AC1032 -> text extraction from binary (no converter needed)
     Extracts UTF-16LE strings + ASCII strings → sends to Gemini as text context

Usage: python3 dwg_converter.py <input_file> <output_png> [dpi] [tiled]
"""
import sys, json, os, traceback, subprocess, tempfile, shutil, re, struct

# ── NEW: Binary DWG text extractor (works on AC1027, AC1032, any version) ──
def extract_text_from_dwg_binary(dwg_path):
    """
    Extract readable text directly from DWG binary.
    Works on ALL DWG versions including AC1032 (R2018) without ODA.
    Extracts ASCII + UTF-16LE strings — catches annotations, layer names,
    dimension text, notes, schedule tables, column IDs, everything.
    """
    try:
        with open(dwg_path, "rb") as f:
            data = f.read()
    except Exception as e:
        return {"texts": [], "layers": [], "error": str(e)}

    version = data[:6].decode("ascii", errors="replace")
    results = {"version": version, "texts": [], "layers": [], "dims": []}

    # ── Pass 1: UTF-16LE strings (AutoCAD R2013+ stores text as UTF-16LE) ──
    utf16_strings = []
    i = 0
    while i < len(data) - 3:
        # Detect UTF-16LE pattern: ASCII char followed by null byte
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

    # ── Pass 2: ASCII strings (older style, layer names, file metadata) ──
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

    # ── Filter: keep only engineering-relevant strings ──
    all_strings = list(dict.fromkeys(utf16_strings + ascii_strings))  # deduplicate, preserve order

    # Layer names: usually short, uppercase, no spaces
    layer_patterns = re.compile(
        r'^[A-Z0-9_\-\.]{2,30}$'
    )
    # Engineering text patterns
    eng_patterns = re.compile(
        r'(FOOTING|COLUMN|COL|BEAM|SLAB|RCC|GRID|LEVEL|FLOOR|WALL|'
        r'SECTION|DETAIL|PLAN|ELEVATION|SCHEDULE|REINFORCEMENT|'
        r'STIRRUP|MAIN BAR|TIE|LINK|DIA|THK|WIDTH|DEPTH|HEIGHT|'
        r'FOUNDATION|PILE|RAFT|GRADE|M\d0|Fe\d{3}|'
        r'NOTES?|SPEC|DESCRIPTION|DRAWING|TITLE|PROJECT|'
        r'ITALVA|BHARUCH|SURAT|PMC|'  # project-specific
        r'\d+[xX]\d+|\d+mm|\d+\.\d+m)',
        re.IGNORECASE
    )
    # Dimension values: numbers with units
    dim_patterns = re.compile(r'^\d+(\.\d+)?$|^\d+[xX]\d+$')

    for s in all_strings:
        s = s.strip()
        if not s or len(s) < 2:
            continue
        # Skip obvious binary noise
        if re.search(r'[^\x20-\x7E]', s):
            continue
        # Skip file paths, GUIDs, base64-like noise
        if re.search(r'[\\\/].*[\\\/]|[A-F0-9]{8}-[A-F0-9]{4}', s):
            continue

        if eng_patterns.search(s):
            results["texts"].append({"text": s, "source": "binary_extract"})
        elif layer_patterns.match(s) and len(s) >= 3:
            results["layers"].append(s)
        elif dim_patterns.match(s):
            results["dims"].append(s)

    # Also extract ALL utf16 strings that are >= 3 chars and look readable
    # (catches column IDs like "C1", "F1", text that didn't match above)
    for s in utf16_strings:
        s = s.strip()
        if len(s) >= 2 and not any(t["text"] == s for t in results["texts"]):
            # Include anything that looks like it could be a label
            if re.match(r'^[A-Z][0-9A-Z\-_\.]+$', s) and len(s) <= 20:
                results["texts"].append({"text": s, "source": "utf16_label"})

    # Deduplicate texts
    seen = set()
    unique_texts = []
    for t in results["texts"]:
        if t["text"] not in seen:
            seen.add(t["text"])
            unique_texts.append(t)
    results["texts"] = unique_texts[:300]
    results["layers"] = list(dict.fromkeys(results["layers"]))[:100]
    results["dims"] = list(dict.fromkeys(results["dims"]))[:100]

    return results


def render_dxf_to_png(dxf_path, png_path, dpi=120, tiled=False):
    """Pure-python DXF -> PNG using ezdxf + matplotlib. No external binary."""
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
        doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    fig = plt.figure(figsize=(16, 12), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ctx = RenderContext(doc)
    out = MatplotlibBackend(ax)
    Frontend(ctx, out).draw_layout(msp, finalize=True)
    fig.savefig(png_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)

    if not tiled:
        return os.path.exists(png_path) and os.path.getsize(png_path) > 0, []

    # Generate tiles for detail mode
    tiles = []
    try:
        from PIL import Image
        img = Image.open(png_path)
        W, H = img.size
        for row in range(2):
            for col in range(2):
                x1, y1 = col * W // 2, row * H // 2
                x2, y2 = (col + 1) * W // 2, (row + 1) * H // 2
                tile = img.crop((x1, y1, x2, y2))
                tile_path = png_path.replace(".png", f"_tile_{row}{col}.png")
                tile.save(tile_path)
                tiles.append({"path": tile_path, "row": row, "col": col,
                              "position": f"{'Top' if row==0 else 'Bottom'}-{'Left' if col==0 else 'Right'}"})
    except Exception as e:
        pass

    return os.path.exists(png_path) and os.path.getsize(png_path) > 0, tiles


def libreoffice_to_png(input_path, output_png):
    out_dir = os.path.dirname(output_png) or tempfile.gettempdir()
    base = os.path.splitext(os.path.basename(input_path))[0]
    lo_png = os.path.join(out_dir, base + ".png")
    cmd = ["libreoffice", "--headless", "--convert-to", "png",
           "--outdir", out_dir, input_path]
    subprocess.run(cmd, capture_output=True, text=True, timeout=160)
    if os.path.exists(lo_png):
        if lo_png != output_png:
            shutil.move(lo_png, output_png)
        return True
    return False


def dwg_to_dxf_via_oda(dwg_path):
    """Use ODA File Converter through ezdxf.addons.odafc. Needs ODA installed."""
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

    texts = []
    for e in msp:
        try:
            if e.dxftype() in ("TEXT", "ATTDEF", "ATTRIB"):
                t = e.dxf.text.strip()
                if t:
                    texts.append({"text": t, "layer": e.dxf.layer,
                                  "x": round(e.dxf.insert.x, 3),
                                  "y": round(e.dxf.insert.y, 3)})
            elif e.dxftype() == "MTEXT":
                t = e.plain_mtext().strip()
                if t:
                    texts.append({"text": t, "layer": e.dxf.layer,
                                  "x": round(e.dxf.insert.x, 3),
                                  "y": round(e.dxf.insert.y, 3)})
        except Exception:
            pass

    dims = []
    for e in msp:
        try:
            if e.dxftype() == "DIMENSION":
                val = None
                try:
                    val = round(e.dxf.actual_measurement, 4)
                except:
                    pass
                txt = ""
                try:
                    txt = e.dxf.text.strip()
                except:
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
            "scale": scale, "drawing_type": dtype}


def run(input_path, output_png, dpi=120, tiled=False):
    result = {
        "success": False,
        "png_path": None,
        "tiles": [],
        "texts": [],
        "dimensions": [],
        "layers": [],
        "drawing_type": "UNKNOWN",
        "scale": "Not detected",
        "extents": {},
        "errors": [],
        "binary_extract": None,   # NEW: binary text extraction result
    }

    ext = os.path.splitext(input_path)[1].lower()
    dxf_for_meta = input_path if ext == ".dxf" else None

    # ── [NEW] STEP 0: Binary text extraction — works on ALL DWG versions ──
    # Run this first so Gemini always gets some text context even if PNG fails
    if ext in (".dwg", ".dxf"):
        try:
            bin_result = extract_text_from_dwg_binary(input_path)
            result["binary_extract"] = bin_result
            # Merge into texts/layers for backward compatibility with server.js
            if bin_result.get("texts"):
                result["texts"] = bin_result["texts"]
            if bin_result.get("layers"):
                result["layers"] = bin_result["layers"]
            if bin_result.get("dims"):
                result["dimensions"] = [{"value": None, "text": d, "layer": "binary"} for d in bin_result["dims"]]
            print(f"[binary_extract] {len(result['texts'])} texts, {len(result['layers'])} layers found", file=sys.stderr)
        except Exception as e:
            result["errors"].append(f"Binary extract failed: {e}")

    # ── STEP 1: Try ODA conversion for DWG (only if ODA installed) ──
    if ext == ".dwg":
        conv = dwg_to_dxf_via_oda(input_path)
        if conv:
            dxf_for_meta = conv
            result["errors"].append("ODA conversion successful")

    # ── STEP 2: DXF -> PNG via ezdxf ──
    if dxf_for_meta:
        try:
            ok, tiles = render_dxf_to_png(dxf_for_meta, output_png, dpi=dpi, tiled=tiled)
            if ok:
                result["png_path"] = output_png
                result["tiles"] = tiles
                result["success"] = True
        except Exception as e:
            result["errors"].append(f"ezdxf render failed: {e}")

    # ── STEP 3: LibreOffice fallback ──
    if not result["png_path"]:
        try:
            if libreoffice_to_png(input_path, output_png):
                result["png_path"] = output_png
                result["success"] = True
        except FileNotFoundError:
            result["errors"].append("LibreOffice not installed on server")
        except subprocess.TimeoutExpired:
            result["errors"].append("LibreOffice timed out")
        except Exception as e:
            result["errors"].append(f"LibreOffice failed: {e}")

    # ── STEP 4: ezdxf metadata extraction (DXF only) ──
    if dxf_for_meta:
        try:
            meta = extract_ezdxf_meta(dxf_for_meta)
            # Merge ezdxf texts with binary-extracted texts (ezdxf more accurate for DXF)
            if meta.get("texts"):
                existing_texts = set(t["text"] for t in result["texts"])
                for t in meta["texts"]:
                    if t["text"] not in existing_texts:
                        result["texts"].append(t)
                        existing_texts.add(t["text"])
            if meta.get("layers"):
                existing_layers = set(result["layers"])
                result["layers"] = list(existing_layers | set(meta["layers"]))
            if meta.get("dimensions") and not result["dimensions"]:
                result["dimensions"] = meta["dimensions"]
            result["scale"] = meta.get("scale", result["scale"])
            result["drawing_type"] = meta.get("drawing_type", result["drawing_type"])
            result["success"] = True
        except Exception as e:
            result["errors"].append(f"ezdxf metadata failed: {e}")

    # ── Mark success if we at least got text from binary ──
    if not result["success"] and (result["texts"] or result["layers"]):
        result["success"] = True
        result["errors"].append(
            "No PNG rendered (DWG version not supported by ezdxf without ODA). "
            "Text extracted from binary — Gemini will analyze via text context. "
            "For best results, export drawing to PDF or PNG from AutoCAD/GstarCAD."
        )

    print(json.dumps(result))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: dwg_converter.py <input> <output.png> [dpi] [tiled]",
                          "success": False}))
        sys.exit(1)
    try:
        dpi_arg = int(sys.argv[3]) if len(sys.argv) > 3 else 120
        tiled_arg = sys.argv[4].lower() in ("true", "1", "yes") if len(sys.argv) > 4 else False
        run(sys.argv[1], sys.argv[2], dpi=dpi_arg, tiled=tiled_arg)
    except Exception as e:
        print(json.dumps({"success": False,
                          "errors": [f"Fatal: {e}", traceback.format_exc()[:1000]]}))
