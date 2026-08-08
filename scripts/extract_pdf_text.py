import fitz
import os
import sys

pdf = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\ADMIN\Downloads\BHAGYESHREE WARE HOUSE ITALVA-3 COLUMN FOOTING R.C.C LAYOUT & DETAILS.pdf"
print("exists", os.path.exists(pdf), "size", os.path.getsize(pdf) if os.path.exists(pdf) else 0)
doc = fitz.open(pdf)
print("pages", len(doc))
all_lines = []
for i, page in enumerate(doc):
    print("--- PAGE", i + 1, "size", round(page.rect.width), "x", round(page.rect.height))
    d = page.get_text("dict")
    items = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for line in b.get("lines", []):
            for span in line.get("spans", []):
                t = (span.get("text") or "").strip()
                if not t:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                items.append({"x": x0, "y": y0, "text": t})
    byY = {}
    for t in items:
        row = round(t["y"] / 8) * 8
        byY.setdefault(row, []).append(t)
    lines = []
    for row in sorted(byY):
        line = " ".join(x["text"] for x in sorted(byY[row], key=lambda z: z["x"]))
        if line.strip():
            lines.append(line)
    all_lines.extend([f"=== PAGE {i+1} ==="] + lines)
    print("texts", len(items), "lines", len(lines))

os.makedirs("data/fixtures", exist_ok=True)
out = "data/fixtures/bhagyeshree_extract.txt"
open(out, "w", encoding="utf-8").write("\n".join(all_lines))
print("WROTE", out)
print("=== FULL TEXT ===")
print("\n".join(all_lines))
