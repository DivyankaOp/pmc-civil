#!/usr/bin/env python3
"""
PMC Civil — DWG/DXF Converter (v3)
Strategy (in order, first that works wins for PNG):
  1. DXF -> PNG via ezdxf + matplotlib  (no external binary required — works on Render)
  2. DXF/DWG -> PNG via LibreOffice --headless  (if installed)
  3. DWG -> DXF via ezdxf.addons.odafc   (if ODA File Converter installed)
Always tries ezdxf text/dim extraction on DXF input.
Usage: python3 dwg_converter.py <input_file> <output_png>
"""
import sys, json, os, traceback, subprocess, tempfile, shutil, re

def render_dxf_to_png(dxf_path, png_path):
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

    fig = plt.figure(figsize=(16, 12), dpi=120)
    ax = fig.add_axes([0, 0, 1, 1])
    ctx = RenderContext(doc)
    out = MatplotlibBackend(ax)
    Frontend(ctx, out).draw_layout(msp, finalize=True)
    fig.savefig(png_path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return os.path.exists(png_path) and os.path.getsize(png_path) > 0


def libreoffice_to_png(input_path, output_png):
    out_dir = os.path.dirname(output_png) or tempfile.gettempdir()
    base    = os.path.splitext(os.path.basename(input_path))[0]
    lo_png  = os.path.join(out_dir, base + ".png")
    cmd = ["libreoffice", "--headless", "--convert-to", "png",
           "--outdir", out_dir, input_path]
    subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if os.path.exists(lo_png):
        if lo_png != output_png:
            shutil.move(lo_png, output_png)
        return True
    return False


def dwg_to_dxf_via_oda(dwg_path):
    """Use ODA File Converter through ezdxf.addons.odafc. Needs ODA installed."""
    from ezdxf.addons import odafc
    out_dxf = os.path.join(tempfile.gettempdir(),
                           os.path.splitext(os.path.basename(dwg_path))[0] + "_conv.dxf")
    odafc.convert(dwg_path, out_dxf, version="R2018")
    return out_dxf if os.path.exists(out_dxf) else None


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
                try: val = round(e.dxf.actual_measurement, 4)
                except: pass
                txt = ""
                try: txt = e.dxf.text.strip()
                except: pass
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
    if any(k in combined for k in ["ROAD","GSB","WMM","PQC","CARRIAGEWAY","KERB"]):
        dtype = "ROAD_PLAN"
    elif any(k in combined for k in ["FOUNDATION","FOOTING","PILE","RAFT"]):
        dtype = "FOUNDATION"
    elif any(k in combined for k in ["SLAB","BEAM","COLUMN","RCC","STRUCTURAL"]):
        dtype = "STRUCTURAL"
    elif any(k in combined for k in ["FLOOR","ROOM","TOILET","KITCHEN","LIVING"]):
        dtype = "FLOOR_PLAN"
    elif any(k in combined for k in ["SITE","LAYOUT","PLOT","BOUNDARY","MASTER"]):
        dtype = "SITE_LAYOUT"
    elif any(k in combined for k in ["SECTION","ELEVATION","ELEV","CROSS"]):
        dtype = "SECTION"

    return {"layers": layers, "texts": texts[:300], "dimensions": dims[:200],
            "scale": scale, "drawing_type": dtype}


def run(input_path, output_png):
    result = {
        "success": False,
        "png_path": None,
        "texts": [],
        "dimensions": [],
        "layers": [],
        "drawing_type": "UNKNOWN",
        "scale": "Not detected",
        "extents": {},
        "errors": [],
    }

    ext = os.path.splitext(input_path)[1].lower()
    dxf_for_meta = input_path if ext == ".dxf" else None

    # If DWG, try ODA conversion so we have a DXF to work with
    if ext == ".dwg":
        try:
            conv = dwg_to_dxf_via_oda(input_path)
            if conv:
                dxf_for_meta = conv
        except Exception as e:
            result["errors"].append(f"ODA convert failed: {e}")

    # STEP 1: DXF -> PNG via ezdxf (works without any external binary)
    if dxf_for_meta:
        try:
            if render_dxf_to_png(dxf_for_meta, output_png):
                result["png_path"] = output_png
                result["success"]  = True
        except Exception as e:
            result["errors"].append(f"ezdxf render failed: {e}")

    # STEP 2: LibreOffice fallback (for DWG without ODA, or DXF if ezdxf render failed)
    if not result["png_path"]:
        try:
            if libreoffice_to_png(input_path, output_png):
                result["png_path"] = output_png
                result["success"]  = True
        except FileNotFoundError:
            result["errors"].append("LibreOffice not installed on server")
        except subprocess.TimeoutExpired:
            result["errors"].append("LibreOffice timed out")
        except Exception as e:
            result["errors"].append(f"LibreOffice failed: {e}")

    # STEP 3: ezdxf metadata extraction (only works on DXF)
    if dxf_for_meta:
        try:
            meta = extract_ezdxf_meta(dxf_for_meta)
            result.update(meta)
            if meta["texts"] or meta["dimensions"]:
                result["success"] = True
        except Exception as e:
            result["errors"].append(f"ezdxf metadata failed: {e}")

    if not result["success"] and ext == ".dwg" and not result["png_path"]:
        result["errors"].append(
            "DWG files need LibreOffice OR ODA File Converter on the server. "
            "Neither found. Please export the drawing to DXF/PDF/PNG and upload that instead."
        )

    print(json.dumps(result))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: dwg_converter.py <input> <output.png>",
                          "success": False}))
        sys.exit(1)
    try:
        run(sys.argv[1], sys.argv[2])
    except Exception as e:
        print(json.dumps({"success": False,
                          "errors": [f"Fatal: {e}", traceback.format_exc()[:1000]]}))
