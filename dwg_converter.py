#!/usr/bin/env python3
"""
PMC Civil — DWG/DXF Converter
Reads DXF (and attempts DWG via ezdxf recover),
renders to PNG, extracts all text + dimensions.
Usage: python3 dwg_converter.py <input_file> <output_png>
Prints JSON to stdout.
"""
import sys, json, os, math, traceback

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

    try:
        import ezdxf
        from ezdxf import recover

        # Try loading — works for DXF directly, attempts DWG recovery
        try:
            doc, auditor = recover.readfile(input_path)
        except Exception:
            try:
                doc = ezdxf.readfile(input_path)
                auditor = None
            except Exception as e:
                result["error"] = f"Cannot open file: {e}"
                print(json.dumps(result))
                return

        msp = doc.modelspace()

        # ── Extract layers
        result["layers"] = [layer.dxf.name for layer in doc.layers if layer.dxf.name != "0"]

        # ── Extract all TEXT and MTEXT
        texts = []
        for entity in msp:
            try:
                if entity.dxftype() in ("TEXT", "ATTDEF", "ATTRIB"):
                    t = entity.dxf.text.strip()
                    if t:
                        texts.append({
                            "text": t,
                            "layer": entity.dxf.layer,
                            "x": round(entity.dxf.insert.x, 3),
                            "y": round(entity.dxf.insert.y, 3),
                            "height": round(entity.dxf.height, 3) if hasattr(entity.dxf, "height") else 0
                        })
                elif entity.dxftype() == "MTEXT":
                    t = entity.plain_mtext().strip()
                    if t:
                        texts.append({
                            "text": t,
                            "layer": entity.dxf.layer,
                            "x": round(entity.dxf.insert.x, 3),
                            "y": round(entity.dxf.insert.y, 3),
                            "height": round(entity.dxf.char_height, 3) if hasattr(entity.dxf, "char_height") else 0
                        })
            except Exception:
                pass
        result["texts"] = texts[:300]

        # ── Extract DIMENSION entities
        dims = []
        for entity in msp:
            try:
                if entity.dxftype() == "DIMENSION":
                    val = None
                    try:
                        val = round(entity.dxf.actual_measurement, 4)
                    except Exception:
                        pass
                    text = ""
                    try:
                        text = entity.dxf.text.strip()
                    except Exception:
                        pass
                    dims.append({
                        "value": val,
                        "text": text,
                        "layer": entity.dxf.layer
                    })
            except Exception:
                pass
        result["dimensions"] = dims[:200]

        # ── Detect scale from text annotations
        import re
        all_text_str = " ".join(t["text"] for t in texts)
        scale_match = re.search(r'(?:scale|sc)[:\s]*1\s*[:/]\s*(\d+)', all_text_str, re.IGNORECASE)
        if not scale_match:
            scale_match = re.search(r'1\s*:\s*(\d+)', all_text_str)
        if scale_match:
            result["scale"] = f"1:{scale_match.group(1)}"

        # ── Detect drawing type from layer names + texts
        layer_str = " ".join(result["layers"]).upper()
        text_upper = all_text_str.upper()
        if any(k in layer_str + text_upper for k in ["ROAD", "GSB", "WMM", "PQC", "CARRIAGEWAY"]):
            result["drawing_type"] = "ROAD_PLAN"
        elif any(k in layer_str + text_upper for k in ["FOUNDATION", "FOOTING", "PILE"]):
            result["drawing_type"] = "FOUNDATION"
        elif any(k in layer_str + text_upper for k in ["SLAB", "BEAM", "COLUMN", "RCC", "STRUCTURAL"]):
            result["drawing_type"] = "STRUCTURAL"
        elif any(k in layer_str + text_upper for k in ["FLOOR", "ROOM", "TOILET", "KITCHEN", "PLAN"]):
            result["drawing_type"] = "FLOOR_PLAN"
        elif any(k in layer_str + text_upper for k in ["SITE", "LAYOUT", "PLOT", "BOUNDARY"]):
            result["drawing_type"] = "SITE_LAYOUT"
        elif any(k in layer_str + text_upper for k in ["SECTION", "ELEVATION", "ELEV"]):
            result["drawing_type"] = "SECTION"

        # ── Render to PNG using matplotlib backend
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            from ezdxf.addons.drawing import RenderContext, Frontend
            from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

            fig = plt.figure(figsize=(16, 12), dpi=150)
            ax = fig.add_axes([0, 0, 1, 1])
            ctx = RenderContext(doc)
            backend = MatplotlibBackend(ax)
            frontend = Frontend(ctx, backend)
            frontend.draw_layout(msp, finalize=True)

            # White background
            fig.patch.set_facecolor("white")
            ax.set_facecolor("white")

            fig.savefig(output_png, dpi=150, bbox_inches="tight",
                        facecolor="white", edgecolor="none")
            plt.close(fig)

            result["png_path"] = output_png
            result["success"] = True

            # Get extents from the figure
            xlim = ax.get_xlim()
            ylim = ax.get_ylim()
            result["extents"] = {
                "xmin": round(xlim[0], 2), "xmax": round(xlim[1], 2),
                "ymin": round(ylim[0], 2), "ymax": round(ylim[1], 2),
                "width": round(xlim[1] - xlim[0], 2),
                "height": round(ylim[1] - ylim[0], 2)
            }

        except Exception as render_err:
            # Render failed but text extraction worked — partial success
            result["success"] = True  # text data still useful
            result["png_path"] = None
            result["render_error"] = str(render_err)

    except Exception as e:
        result["error"] = traceback.format_exc()

    print(json.dumps(result))

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: dwg_converter.py <input> <output.png>"}))
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
