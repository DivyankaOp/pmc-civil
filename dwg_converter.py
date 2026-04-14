#!/usr/bin/env python3
"""
PMC Civil — DWG/DXF Converter (v2)
Strategy:
  1. LibreOffice --headless converts DWG/DXF -> PNG  (Gemini sees actual drawing)
  2. ezdxf extracts text/dims from DXF (for DWG, skip if ezdxf fails)
  3. JSON result printed to stdout
Usage: python3 dwg_converter.py <input_file> <output_png>
"""
import sys, json, os, traceback, subprocess, tempfile, shutil

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
        "error": None
    }

    # STEP 1: LibreOffice DWG/DXF -> PNG
    try:
        out_dir = os.path.dirname(output_png) or "/tmp"
        base    = os.path.splitext(os.path.basename(input_path))[0]
        lo_png  = os.path.join(out_dir, base + ".png")

        cmd = ["libreoffice", "--headless", "--convert-to", "png", "--outdir", out_dir, input_path]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90)

        if os.path.exists(lo_png):
            if lo_png != output_png:
                shutil.move(lo_png, output_png)
            result["png_path"] = output_png
            result["success"]  = True
        else:
            result["lo_error"] = f"No PNG produced. stdout={proc.stdout[:200]} stderr={proc.stderr[:200]}"

    except subprocess.TimeoutExpired:
        result["lo_error"] = "LibreOffice timed out"
    except FileNotFoundError:
        result["lo_error"] = "LibreOffice not found"
    except Exception as e:
        result["lo_error"] = str(e)

    # STEP 2: ezdxf text/dim extraction
    try:
        import ezdxf
        from ezdxf import recover
        import re

        try:
            doc, _ = recover.readfile(input_path)
        except Exception:
            doc = ezdxf.readfile(input_path)

        msp = doc.modelspace()
        result["layers"] = [l.dxf.name for l in doc.layers if l.dxf.name != "0"]

        texts = []
        for e in msp:
            try:
                if e.dxftype() in ("TEXT", "ATTDEF", "ATTRIB"):
                    t = e.dxf.text.strip()
                    if t:
                        texts.append({"text": t, "layer": e.dxf.layer,
                                      "x": round(e.dxf.insert.x, 3), "y": round(e.dxf.insert.y, 3)})
                elif e.dxftype() == "MTEXT":
                    t = e.plain_mtext().strip()
                    if t:
                        texts.append({"text": t, "layer": e.dxf.layer,
                                      "x": round(e.dxf.insert.x, 3), "y": round(e.dxf.insert.y, 3)})
            except Exception:
                pass
        result["texts"] = texts[:300]

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
        result["dimensions"] = dims[:200]

        all_text = " ".join(t["text"] for t in texts)
        m = re.search(r'(?:scale|sc)[:\s]*1\s*[:/]\s*(\d+)', all_text, re.IGNORECASE)
        if not m:
            m = re.search(r'\b1\s*:\s*(\d+)\b', all_text)
        if m:
            result["scale"] = f"1:{m.group(1)}"

        combined = (" ".join(result["layers"]) + " " + all_text).upper()
        if any(k in combined for k in ["ROAD","GSB","WMM","PQC","CARRIAGEWAY","KERB"]):
            result["drawing_type"] = "ROAD_PLAN"
        elif any(k in combined for k in ["FOUNDATION","FOOTING","PILE","RAFT"]):
            result["drawing_type"] = "FOUNDATION"
        elif any(k in combined for k in ["SLAB","BEAM","COLUMN","RCC","STRUCTURAL"]):
            result["drawing_type"] = "STRUCTURAL"
        elif any(k in combined for k in ["FLOOR","ROOM","TOILET","KITCHEN","LIVING"]):
            result["drawing_type"] = "FLOOR_PLAN"
        elif any(k in combined for k in ["SITE","LAYOUT","PLOT","BOUNDARY","MASTER"]):
            result["drawing_type"] = "SITE_LAYOUT"
        elif any(k in combined for k in ["SECTION","ELEVATION","ELEV","CROSS"]):
            result["drawing_type"] = "SECTION"

        if texts or dims:
            result["success"] = True

    except Exception as e:
        result["ezdxf_error"] = str(e)

    print(json.dumps(result))

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: dwg_converter.py <input> <output.png>", "success": False}))
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
