#!/usr/bin/env python3
"""
dwg_to_png.py — Smart tiled DWG/DXF renderer
Renders at 300 DPI, cuts into zoomed tiles, skips empty tiles
Each tile sent separately to Gemini Vision for accurate analysis
"""
import sys, os, json
from collections import defaultdict

def render_drawing(input_path, output_dir, dpi=300):
    import ezdxf, matplotlib, math, re
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection
    from PIL import Image
    import numpy as np

    os.makedirs(output_dir, exist_ok=True)

    # Load
    try:
        doc = ezdxf.readfile(input_path)
    except Exception:
        doc, _ = ezdxf.recover.readfile(input_path)

    msp = doc.modelspace()

    lines, texts, circles, dims = [], [], [], []

    def extract_block(entities, offset_x=0, offset_y=0):
        for e in entities:
            try:
                t = e.dxftype()
                if t == 'LINE':
                    s, en = e.dxf.start, e.dxf.end
                    lines.append([(s.x+offset_x, s.y+offset_y), (en.x+offset_x, en.y+offset_y)])
                elif t == 'LWPOLYLINE':
                    pts = list(e.get_points())
                    for i in range(len(pts)-1):
                        lines.append([(pts[i][0]+offset_x, pts[i][1]+offset_y),
                                      (pts[i+1][0]+offset_x, pts[i+1][1]+offset_y)])
                    if e.closed and pts:
                        lines.append([(pts[-1][0]+offset_x, pts[-1][1]+offset_y),
                                      (pts[0][0]+offset_x, pts[0][1]+offset_y)])
                elif t == 'POLYLINE':
                    pts = list(e.points())
                    for i in range(len(pts)-1):
                        lines.append([(pts[i][0]+offset_x, pts[i][1]+offset_y),
                                      (pts[i+1][0]+offset_x, pts[i+1][1]+offset_y)])
                elif t == 'CIRCLE':
                    c = e.dxf.center
                    circles.append((c.x+offset_x, c.y+offset_y, e.dxf.radius))
                elif t == 'ARC':
                    c = e.dxf.center
                    r = e.dxf.radius
                    a1 = math.radians(e.dxf.start_angle)
                    a2 = math.radians(e.dxf.end_angle)
                    if a2 < a1: a2 += 2*math.pi
                    steps = max(12, int((a2-a1)*r/2))
                    arc_pts = [(c.x+offset_x + r*math.cos(a1+(a2-a1)*i/steps),
                                c.y+offset_y + r*math.sin(a1+(a2-a1)*i/steps))
                               for i in range(steps+1)]
                    for i in range(len(arc_pts)-1):
                        lines.append([arc_pts[i], arc_pts[i+1]])
                elif t in ('TEXT','MTEXT'):
                    raw = (e.dxf.text if hasattr(e.dxf,'text') else e.text) or ''
                    raw = re.sub(r'\\[A-Za-z][^;]*;','',raw).replace('\\P','\n').strip()
                    if raw:
                        pos = e.dxf.insert if hasattr(e.dxf,'insert') else None
                        if pos:
                            ht = getattr(e.dxf,'height',2) or 2
                            texts.append((pos.x+offset_x, pos.y+offset_y, raw, ht))
                elif t == 'DIMENSION':
                    try:
                        txt = str(e.dxf.text or '')
                        mp = e.dxf.text_midpoint if hasattr(e.dxf,'text_midpoint') else None
                        if txt and mp:
                            dims.append((mp.x+offset_x, mp.y+offset_y, txt))
                        if hasattr(e.dxf,'defpoint') and hasattr(e.dxf,'defpoint2'):
                            d1,d2 = e.dxf.defpoint, e.dxf.defpoint2
                            lines.append([(d1.x+offset_x,d1.y+offset_y),(d2.x+offset_x,d2.y+offset_y)])
                    except: pass
                elif t == 'INSERT':
                    try:
                        blk = doc.blocks[e.dxf.name]
                        ins = e.dxf.insert
                        extract_block(blk, ins.x+offset_x, ins.y+offset_y)
                    except: pass
            except: continue

    extract_block(msp)

    if not lines and not texts:
        return {'success': False, 'error': 'No geometry found'}

    all_x = [p[0] for l in lines for p in l] + [t[0] for t in texts] + [c[0] for c in circles]
    all_y = [p[1] for l in lines for p in l] + [t[1] for t in texts] + [c[1] for c in circles]

    xmin,xmax = min(all_x),max(all_x)
    ymin,ymax = min(all_y),max(all_y)
    w = xmax-xmin; h = ymax-ymin
    pad = max(w,h)*0.02
    xmin-=pad; xmax+=pad; ymin-=pad; ymax+=pad
    w=xmax-xmin; h=ymax-ymin

    # Render full high-res
    fig_w = 50
    fig_h = min(fig_w*(h/w) if w>0 else fig_w, 70)
    fig,ax = plt.subplots(figsize=(fig_w,fig_h),facecolor='white')
    ax.set_facecolor('white'); ax.set_xlim(xmin,xmax); ax.set_ylim(ymin,ymax)
    ax.set_aspect('equal'); ax.axis('off')

    # Lines
    if lines:
        lc = LineCollection(lines, colors='black', linewidths=0.7, alpha=0.95)
        ax.add_collection(lc)

    # Circles
    for (cx,cy,r) in circles:
        ax.add_patch(plt.Circle((cx,cy),r,fill=False,color='black',linewidth=0.6))

    # Text (scaled)
    base_fs = max(5, min(14, w/60))
    for (tx,ty,txt,ht) in texts:
        fs = max(5, min(18, base_fs*(ht/max(w/80,0.5))))
        ax.text(tx,ty,txt,fontsize=fs,color='navy',ha='left',va='bottom',
                clip_on=True,fontweight='bold' if ht>3 else 'normal',
                bbox=dict(boxstyle='round,pad=0.05',facecolor='white',edgecolor='none',alpha=0.6))

    # Dimensions (red)
    for (dx,dy,dtxt) in dims:
        ax.text(dx,dy,dtxt,fontsize=base_fs*0.85,color='red',ha='center',va='center',
                clip_on=True,bbox=dict(boxstyle='round,pad=0.1',facecolor='lightyellow',edgecolor='red',alpha=0.8,linewidth=0.3))

    full_png = os.path.join(output_dir,'full.png')
    plt.savefig(full_png, dpi=dpi, bbox_inches='tight', facecolor='white', pad_inches=0.05)
    plt.close(fig)

    # Tile the full image
    img = Image.open(full_png)
    img_w, img_h = img.size
    entity_count = len(lines)+len(texts)

    # Grid size based on complexity
    if entity_count > 3000: cols,rows = 4,4
    elif entity_count > 1000: cols,rows = 4,3
    elif entity_count > 300: cols,rows = 3,3
    else: cols,rows = 2,2

    overlap = 0.12
    tile_w = img_w//cols; tile_h = img_h//rows
    ov_x = int(tile_w*overlap); ov_y = int(tile_h*overlap)

    tile_paths = []
    idx = 0
    for row in range(rows):
        for col in range(cols):
            x1=max(0,col*tile_w-ov_x); y1=max(0,row*tile_h-ov_y)
            x2=min(img_w,(col+1)*tile_w+ov_x); y2=min(img_h,(row+1)*tile_h+ov_y)
            tile = img.crop((x1,y1,x2,y2))

            # Skip blank tiles
            arr = np.array(tile.convert('L')); 
            if (arr < 200).sum() < 50:  # skip if fewer than 50 actual dark pixels
                continue

            # Resize to max 2048px for Gemini
            tw,th = tile.size
            if max(tw,th) > 2048:
                scale = 2048/max(tw,th)
                tile = tile.resize((int(tw*scale),int(th*scale)),Image.LANCZOS)

            tp = os.path.join(output_dir, f'tile_{idx:02d}_r{row}c{col}.png')
            tile.convert("RGB").save(tp,"PNG")
            vpos = 'Top' if row==0 else('Bottom' if row==rows-1 else 'Middle')
            hpos = 'Left' if col==0 else('Right' if col==cols-1 else 'Center')
            tile_paths.append({'path':tp,'row':row,'col':col,'position':f'{vpos}-{hpos} section of drawing'})
            idx += 1

    # Overview (small)
    ov = img.copy(); ov.thumbnail((1500,1500),Image.LANCZOS)
    ov_path = os.path.join(output_dir,'overview.png'); ov.convert("RGB").save(ov_path,"PNG")

    return {
        'success': True, 'overview_png': ov_path, 'full_png': full_png,
        'tiles': tile_paths, 'tile_count': len(tile_paths),
        'grid': f'{cols}x{rows}', 'entity_count': entity_count,
        'texts_found': list(set(t[2] for t in texts))[:150],
        'dims_found': list(set(d[2] for d in dims))[:80],
        'drawing_size_units': {'w': round(w,2), 'h': round(h,2)}
    }

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(json.dumps({'success':False,'error':'Usage: dwg_to_png.py <input> <output_dir>'}))
        sys.exit(1)
    print(json.dumps(render_drawing(sys.argv[1], sys.argv[2])))
