'use strict';
/**
 * PMC Drawing Intelligence Engine — v3.0 (AI-First)
 *
 * PHILOSOPHY:
 *   Do NOT depend on layer names. Layer names differ per architect.
 *   Extract EVERYTHING raw from DXF — every text, note, symbol,
 *   dimension, block name, hatch pattern, layer name — and send
 *   it ALL to Gemini AI. AI reads drawing content like a human engineer.
 *
 *   LEARNING: What AI understands gets saved to symbols-learned.json.
 *   Next time same architect's drawing comes → learned data used directly.
 *
 * FLOW:
 *   1. scanDXF()          → extract all raw entities
 *   2. buildRawDump()     → format everything for AI
 *   3. buildAIPrompt()    → create Gemini prompt
 *   4. saveLearnedFromAI()→ store AI's understanding for future
 */

const fs   = require('fs');
const path = require('path');
const LEARNED_FILE = path.join(__dirname, 'symbols-learned.json');

// ─────────────────────────────────────────────────────────────────
// SECTION 1 — LEARNED SYMBOLS
// ─────────────────────────────────────────────────────────────────
function loadLearned() {
  try {
    if (fs.existsSync(LEARNED_FILE))
      return JSON.parse(fs.readFileSync(LEARNED_FILE, 'utf8'));
  } catch(e) {}
  return { layer_meanings:{}, block_meanings:{}, hatch_meanings:{}, text_patterns:[], architect_profiles:{} };
}

function saveLearned(data) {
  try {
    const existing = loadLearned();
    const merged = {
      layer_meanings:     { ...existing.layer_meanings,     ...(data.layer_meanings     || {}) },
      block_meanings:     { ...existing.block_meanings,     ...(data.block_meanings     || {}) },
      hatch_meanings:     { ...existing.hatch_meanings,     ...(data.hatch_meanings     || {}) },
      text_patterns:      [...(existing.text_patterns||[]), ...(data.text_patterns      || [])],
      architect_profiles: { ...existing.architect_profiles, ...(data.architect_profiles || {}) },
      _last_updated: new Date().toISOString(),
      _total_drawings: (existing._total_drawings || 0) + 1
    };
    // Deduplicate text_patterns
    const seen = new Set();
    merged.text_patterns = merged.text_patterns.filter(p => {
      if (seen.has(p.pattern)) return false;
      seen.add(p.pattern); return true;
    });
    fs.writeFileSync(LEARNED_FILE, JSON.stringify(merged, null, 2));
    return merged;
  } catch(e) { console.warn('Save learned failed:', e.message); return data; }
}

function saveLearnedFromAI(aiResult, scanned, filename) {
  if (!aiResult || typeof aiResult !== 'object') return;
  const toLearn = { layer_meanings:{}, block_meanings:{}, hatch_meanings:{}, text_patterns:[], architect_profiles:{} };
  if (aiResult.layer_mappings) {
    for (const [k,v] of Object.entries(aiResult.layer_mappings)) {
      if (k && v && v.category) toLearn.layer_meanings[k] = v;
    }
  }
  if (aiResult.block_mappings) {
    for (const [k,v] of Object.entries(aiResult.block_mappings)) {
      if (k && v && v.category) toLearn.block_meanings[k] = v;
    }
  }
  if (aiResult.architect_name) {
    const key = aiResult.architect_name.toUpperCase().replace(/\s+/g,'_').slice(0,20);
    const layerNames = Object.keys(scanned.layers);
    const prefixes = {};
    for (const n of layerNames) { const m = n.match(/^([A-Z]{1,4}[-_])/); if (m) prefixes[m[1]] = (prefixes[m[1]]||0)+1; }
    const prefix = Object.entries(prefixes).sort((a,b)=>b[1]-a[1])[0]?.[0] || '';
    toLearn.architect_profiles[key] = { name: aiResult.architect_name, layer_prefix: prefix, unit: aiResult.unit||'mm', _from: filename, _at: new Date().toISOString() };
  }
  saveLearned(toLearn);
}

// ─────────────────────────────────────────────────────────────────
// SECTION 2 — RAW DXF SCANNER
// ─────────────────────────────────────────────────────────────────
function scanDXF(content) {
  const lines = content.split(/\r?\n/);
  const result = {
    layers:{}, hatches:[], polylines:[], inserts:[], texts:[], dims:[], lines_:[], blocks:{},
    extents:{ xmin:Infinity, xmax:-Infinity, ymin:Infinity, ymax:-Infinity }
  };
  let i = 0;
  const nxt = () => lines[i++]?.trim() ?? '';
  const pk  = () => lines[i]?.trim()  ?? '';

  function readPairs() {
    const p = {};
    while (i < lines.length) {
      const c = parseInt(pk());
      if (isNaN(c)) { nxt(); nxt(); continue; }
      if (c === 0) break;
      nxt(); const v = nxt();
      if (c in p) p[c] = Array.isArray(p[c]) ? [...p[c],v] : [p[c],v];
      else p[c] = v;
    }
    return p;
  }

  function clean(s) {
    if (!s) return '';
    return s.replace(/\\P/g,' ').replace(/\\p[^;]*;/g,'').replace(/\\f[^;]*;/g,'')
      .replace(/\\[HWhWAaCScs][^;]*;/g,'').replace(/\{\\[^}]*\}/g,m=>m.replace(/\{\\[^;]*;/g,'').replace(/[{}]/g,''))
      .replace(/[{}]/g,'').replace(/%%[cCdDpP]/g,'°').replace(/\\[CcLlOoUu]\d*;?/g,'').trim();
  }

  function shoelace(pts) {
    const n=pts.length; if(n<3) return 0;
    let a=0; for(let j=0;j<n;j++){const k=(j+1)%n; a+=pts[j][0]*pts[k][1]-pts[k][0]*pts[j][1];} return Math.abs(a)/2;
  }
  function perimeter(pts) {
    let p=0; for(let j=0;j<pts.length;j++){const k=(j+1)%pts.length; p+=Math.sqrt((pts[k][0]-pts[j][0])**2+(pts[k][1]-pts[j][1])**2);} return p;
  }
  function upd(x,y) {
    if(!isNaN(x)&&!isNaN(y)&&isFinite(x)&&isFinite(y)) {
      if(x<result.extents.xmin) result.extents.xmin=x; if(x>result.extents.xmax) result.extents.xmax=x;
      if(y<result.extents.ymin) result.extents.ymin=y; if(y>result.extents.ymax) result.extents.ymax=y;
    }
  }

  function parseTables() {
    while(i<lines.length) {
      const c=parseInt(pk()); if(isNaN(c)){nxt();nxt();continue;}
      if(c===0){nxt();const v=nxt(); if(v==='ENDSEC') break;
        if(v==='LAYER'){const p=readPairs();const name=p[2]||'';if(name) result.layers[name]={color:Math.abs(parseInt(p[62])||7),linetype:p[6]||'Continuous',entity_count:0};}
      } else {nxt();nxt();}
    }
  }

  function parseEntities() {
    while(i<lines.length) {
      const c=parseInt(pk()); if(isNaN(c)){nxt();nxt();continue;} if(c!==0){nxt();nxt();continue;}
      nxt(); const etype=nxt(); if(!etype) continue; const up=etype.toUpperCase();
      if(up==='ENDSEC'||up==='EOF') break;

      if(up==='TEXT'||up==='ATTDEF'||up==='ATTRIB') {
        const p=readPairs(); const layer=p[8]||'0'; const txt=clean(p[1]||'');
        const x=parseFloat(p[10])||0,y=parseFloat(p[20])||0;
        if(txt){result.texts.push({text:txt,layer,x,y,height:parseFloat(p[40])||2.5}); if(result.layers[layer]) result.layers[layer].entity_count++; upd(x,y);}

      } else if(up==='MTEXT') {
        const p=readPairs(); const layer=p[8]||'0';
        let txt=clean(Array.isArray(p[1])?p[1].join(''):p[1]||'');
        if(!txt) txt=clean(Array.isArray(p[3])?p[3].join(''):p[3]||'');
        const x=parseFloat(p[10])||0,y=parseFloat(p[20])||0;
        if(txt){result.texts.push({text:txt,layer,x,y,height:parseFloat(p[40])||2.5}); if(result.layers[layer]) result.layers[layer].entity_count++; upd(x,y);}

      } else if(up==='DIMENSION') {
        const p=readPairs(); const layer=p[8]||'0'; const measured=parseFloat(p[42]); const txtOver=clean(p[1]||'');
        const x1=parseFloat(p[13])||0,y1=parseFloat(p[23])||0,x2=parseFloat(p[14])||0,y2=parseFloat(p[24])||0;
        const geom=Math.sqrt((x2-x1)**2+(y2-y1)**2); const val=(!isNaN(measured)&&measured>0)?measured:geom;
        if(val>1){result.dims.push({value_mm:Math.round(val),value_m:Math.round(val)/1000,text_override:txtOver,layer,x:x1,y:y1}); if(result.layers[layer]) result.layers[layer].entity_count++;}

      } else if(up==='HATCH') {
        const p=readPairs(); result.hatches.push({layer:p[8]||'0',pattern:p[2]||'SOLID',color:p[62]||''});
        if(result.layers[p[8]||'0']) result.layers[p[8]||'0'].entity_count++;

      } else if(up==='LWPOLYLINE') {
        const p=readPairs(); const layer=p[8]||'0'; const flags=parseInt(p[70])||0; const closed=!!(flags&1);
        const xs=Array.isArray(p[10])?p[10]:(p[10]?[p[10]]:[]);
        const ys=Array.isArray(p[20])?p[20]:(p[20]?[p[20]]:[]);
        if(xs.length>=2) {
          const pts=xs.map((x,idx)=>[parseFloat(x)||0,parseFloat(ys[idx]||0)]);
          const areaMm2=shoelace(pts); const periMm=perimeter(pts);
          pts.forEach(([x,y])=>upd(x,y));
          const isClosed=closed||(pts.length>2&&Math.abs(pts[0][0]-pts[pts.length-1][0])<10&&Math.abs(pts[0][1]-pts[pts.length-1][1])<10);
          result.polylines.push({layer,pts,area_m2:Math.round(areaMm2/1e6*100)/100,area_sqft:Math.round(areaMm2/92903*100)/100,perimeter_m:Math.round(periMm/1000*100)/100,vertices:pts.length,is_closed:isClosed});
          if(result.layers[layer]) result.layers[layer].entity_count++;
        }

      } else if(up==='INSERT') {
        const p=readPairs(); const block=p[2]||'',layer=p[8]||'0';
        const x=parseFloat(p[10])||0,y=parseFloat(p[20])||0;
        result.inserts.push({block,layer,x,y,scale_x:parseFloat(p[41])||1,scale_y:parseFloat(p[42])||1,rotation:parseFloat(p[50])||0});
        if(result.layers[layer]) result.layers[layer].entity_count++; upd(x,y);

      } else if(up==='LINE') {
        const p=readPairs(); const layer=p[8]||'0';
        const x1=parseFloat(p[10])||0,y1=parseFloat(p[20])||0,x2=parseFloat(p[11])||0,y2=parseFloat(p[21])||0;
        const len=Math.sqrt((x2-x1)**2+(y2-y1)**2);
        if(len>0) result.lines_.push({layer,x1,y1,x2,y2,length_mm:Math.round(len)});
        if(result.layers[layer]) result.layers[layer].entity_count++; upd(x1,y1); upd(x2,y2);
      } else { readPairs(); }
    }
  }

  function parseBlocks() {
    let cur=null;
    while(i<lines.length) {
      const c=parseInt(pk()); if(isNaN(c)){nxt();nxt();continue;} if(c!==0){nxt();nxt();continue;}
      nxt(); const v=nxt(); if(v==='ENDSEC') break;
      if(v==='BLOCK'){const p=readPairs();cur=p[2]||'';if(cur&&!cur.startsWith('*')) result.blocks[cur]={texts:[],inserts:[]};}
      else if(v==='ENDBLK'){cur=null;}
      else if(cur&&result.blocks[cur]) {
        if(v==='TEXT'||v==='MTEXT'){const p=readPairs();const txt=clean(p[1]||p[3]||'');if(txt) result.blocks[cur].texts.push(txt);}
        else if(v==='INSERT'){const p=readPairs();result.blocks[cur].inserts.push(p[2]||'');}
        else readPairs();
      } else readPairs();
    }
  }

  while(i<lines.length) {
    const c=parseInt(pk()); if(isNaN(c)){nxt();nxt();continue;} if(c!==0){nxt();nxt();continue;}
    nxt(); const v=nxt();
    if(v==='SECTION'){nxt();const sec=nxt();
      if(sec==='TABLES') parseTables(); else if(sec==='ENTITIES') parseEntities(); else if(sec==='BLOCKS') parseBlocks();
    }
  }
  return result;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 3 — BUILD RAW DUMP FOR AI
// ─────────────────────────────────────────────────────────────────
function buildRawDump(scanned, filename) {
  const uniqueTexts = [...new Map(scanned.texts.map(t=>[t.text,t])).values()];
  const blockCounts = {};
  for (const ins of scanned.inserts) blockCounts[ins.block]=(blockCounts[ins.block]||0)+1;

  const layerStats = Object.entries(scanned.layers)
    .sort((a,b)=>b[1].entity_count-a[1].entity_count).slice(0,60)
    .map(([n,info])=>`${n}(${info.entity_count})`);

  const hatchByLayer = {};
  for (const h of scanned.hatches) {
    if(!hatchByLayer[h.layer]) hatchByLayer[h.layer]={};
    hatchByLayer[h.layer][h.pattern]=(hatchByLayer[h.layer][h.pattern]||0)+1;
  }

  const closedPolys = scanned.polylines.filter(p=>p.is_closed&&p.area_m2>0.1).sort((a,b)=>b.area_m2-a.area_m2);
  const polyByLayer = {};
  for (const p of closedPolys) { if(!polyByLayer[p.layer]) polyByLayer[p.layer]=[]; polyByLayer[p.layer].push(p.area_m2); }

  const dimCounts = {};
  for (const d of scanned.dims) dimCounts[d.value_mm]=(dimCounts[d.value_mm]||0)+1;
  const topDims = Object.entries(dimCounts).sort((a,b)=>b[1]-a[1]).slice(0,50).map(([v,c])=>`${v}mm${c>1?'×'+c:''}`);

  const ext=scanned.extents;
  const wMm=isFinite(ext.xmax-ext.xmin)?Math.round(ext.xmax-ext.xmin):0;
  const hMm=isFinite(ext.ymax-ext.ymin)?Math.round(ext.ymax-ext.ymin):0;

  const textByLayer={};
  for (const t of scanned.texts) { if(!textByLayer[t.layer]) textByLayer[t.layer]=[]; if(!textByLayer[t.layer].includes(t.text)) textByLayer[t.layer].push(t.text); }

  const linesByLayer={};
  for (const l of scanned.lines_) { if(!linesByLayer[l.layer]) linesByLayer[l.layer]={count:0,total:0}; linesByLayer[l.layer].count++; linesByLayer[l.layer].total+=l.length_mm; }

  let d='';
  d+=`FILE: ${filename}\n`;
  d+=`EXTENTS: ${wMm}mm × ${hMm}mm  (${(wMm/1000).toFixed(2)}m × ${(hMm/1000).toFixed(2)}m)\n`;
  d+=`COUNTS: Texts=${scanned.texts.length} | Dims=${scanned.dims.length} | Polylines=${scanned.polylines.length} | Inserts=${scanned.inserts.length} | Lines=${scanned.lines_.length} | Hatches=${scanned.hatches.length}\n\n`;

  d+=`═══ ALL TEXT ANNOTATIONS (${uniqueTexts.length} unique) ═══\n`;
  for (const t of uniqueTexts.slice(0,300)) d+=`  [${t.layer}] "${t.text}"\n`;

  d+=`\n═══ LAYERS (name:entity_count) ═══\n`;
  d+=layerStats.join(' | ')+'\n';

  d+=`\n═══ TEXTS PER LAYER ═══\n`;
  for (const [layer,txts] of Object.entries(textByLayer).slice(0,40))
    d+=`  ${layer}: ${txts.slice(0,10).map(t=>`"${t}"`).join(', ')}\n`;

  d+=`\n═══ BLOCK INSTANCES ═══\n`;
  for (const [name,count] of Object.entries(blockCounts).sort((a,b)=>b[1]-a[1]).slice(0,80)) {
    const btexts=scanned.blocks[name]?.texts?.slice(0,3).join(',')||'';
    d+=`  ${name} × ${count}${btexts?` [contains: ${btexts}]`:''}\n`;
  }

  d+=`\n═══ HATCH PATTERNS BY LAYER ═══\n`;
  for (const [layer,patterns] of Object.entries(hatchByLayer).slice(0,30))
    d+=`  ${layer}: ${Object.entries(patterns).map(([p,c])=>`${p}×${c}`).join(', ')}\n`;

  d+=`\n═══ CLOSED POLYLINE AREAS BY LAYER ═══\n`;
  for (const [layer,areas] of Object.entries(polyByLayer).slice(0,30)) {
    const total=areas.reduce((s,a)=>s+a,0).toFixed(2);
    d+=`  ${layer}: total=${total}m², [${areas.slice(0,5).map(a=>a.toFixed(2)+'m²').join(', ')}${areas.length>5?'...':''}] count=${areas.length}\n`;
  }

  d+=`\n═══ DIMENSION VALUES (top 50) ═══\n`;
  d+=topDims.join(', ')+'\n';

  d+=`\n═══ LINES PER LAYER ═══\n`;
  for (const [layer,info] of Object.entries(linesByLayer).slice(0,20))
    d+=`  ${layer}: ${info.count} lines, total=${(info.total/1000).toFixed(1)}m\n`;

  return d;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 4 — BUILD GEMINI PROMPT
// ─────────────────────────────────────────────────────────────────
function buildAIPrompt(rawDump, learned, ratesJson) {
  const learnedLayers = Object.entries(learned.layer_meanings||{}).slice(0,30)
    .map(([k,v])=>`  "${k}" → ${v.category}${v.thk_mm?' '+v.thk_mm+'mm':''}`).join('\n')||'  none yet';
  const learnedBlocks = Object.entries(learned.block_meanings||{}).slice(0,30)
    .map(([k,v])=>`  "${k}" → ${v.category} (${v.type||''})`).join('\n')||'  none yet';

  const ratesSummary = Object.entries(ratesJson).filter(([k])=>!k.startsWith('_'))
    .map(([cat,items])=>typeof items==='object'?Object.entries(items).slice(0,4).map(([,v])=>`${v.description}:₹${v.rate}/${v.unit}`).join(' | '):'')
    .filter(Boolean).join(' | ').slice(0,800);

  return `You are a senior PMC civil engineer and drawing expert. Analyze this DXF drawing data extracted directly from the file.

READ CAREFULLY: Every text, note, symbol, block name, dimension and layer tells you what the drawing contains.
Your job: understand the drawing like a human engineer — read all notes, understand symbols, figure out every element.

═══ PREVIOUSLY LEARNED (from past drawings by same firm) ═══
LAYER MEANINGS:
${learnedLayers}

BLOCK MEANINGS:
${learnedBlocks}

═══ THIS DRAWING'S RAW DATA ═══
${rawDump}

═══ RATES (Gujarat DSR 2025) ═══
${ratesSummary}

═══ YOUR TASK ═══
1. Read ALL text annotations — they tell you room names, material specs, notes, levels
2. Understand each BLOCK name — SW=sliding window, SLD=sliding door, D=door, W=window, COL=column
3. Read LAYER names as hints (but don't depend on them — read content instead)
4. Identify HATCH patterns — ANSI31=brick, ANSI37=concrete, SOLID=RCC, AR-CONC=concrete
5. Read DIMENSION values — sizes of rooms, walls, openings, floor heights
6. POLYLINE areas on each layer = actual floor areas, wall plan areas etc.
7. Match text near polylines to identify room/space names
8. Extract ALL material notes ("12MM THK TOUGHENED GLASS", "230MM BRICK WALL", "M30 CONCRETE" etc.)
9. If floor levels are mentioned (+7590MM LEVEL etc.) extract them and calculate floor heights

Return ONLY raw JSON (no markdown, no backticks):
{
  "project_name": "",
  "architect_name": "",
  "drawing_type": "FLOOR_PLAN|SECTION_ELEVATION|STRUCTURAL|FOUNDATION|SITE_PLAN|MEP|DETAIL|GENERAL",
  "drawing_title": "",
  "scale": "",
  "unit": "mm",
  "floors_shown": [],
  "floor_levels": [{"name":"","level_mm":0,"level_m":0}],
  "floor_heights": [{"from":"","to":"","height_m":0}],
  "spaces": [{"name":"","area_sqm":0,"floor":"","dimensions":""}],
  "wall_schedule": [{"description":"","thickness_mm":0,"layer":"","area_m2":0,"length_m":0}],
  "opening_schedule": [{"type":"door|window|sliding_door|sliding_window|ventilator","tag":"","count":0,"size":"","remarks":""}],
  "material_notes": [{"item":"","specification":"","layer":""}],
  "structural_elements": [{"type":"column|beam|slab|footing","tag":"","size":"","reinforcement":"","count":0,"concrete_grade":""}],
  "dimension_summary": {"typical_room_width_mm":0,"typical_wall_thk_mm":0,"slab_thk_mm":0,"floor_to_floor_mm":0},
  "boq": [{"sr":1,"description":"","unit":"sqmt|cum|rmt|nos|kg","qty":0,"rate":0,"amount":0}],
  "total_bua_sqm": 0,
  "layer_mappings": {"LAYER_NAME":{"category":"wall|column|slab|door|window|text|dim|grid|hatch|annotation|ignore","thk_mm":null,"notes":""}},
  "block_mappings": {"BLOCK_NAME":{"category":"opening|column|stair|lift|furniture","type":"door|window|sliding_door|sliding_window|column","remarks":""}},
  "observations": [],
  "pmc_recommendation": ""
}

RULES:
- Use ONLY data from the drawing. Do NOT invent numbers.
- If something is unclear, note it in observations — don't guess.
- BOQ must use actual quantities from polyline areas and dimensions.
- layer_mappings and block_mappings will be SAVED for future drawings — fill them accurately.
- For wall BOQ: area_m2 from polylines ÷ thickness = length, length × floor_height = face_area.`;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 5 — APPLY LEARNED (pre-fill known meanings)
// ─────────────────────────────────────────────────────────────────
function applyLearned(scanned, learned) {
  const result = { known_layers:{}, unknown_layers:[], known_blocks:{}, unknown_blocks:[] };
  for (const layerName of Object.keys(scanned.layers)) {
    if (layerName.startsWith('*')||layerName==='Defpoints') continue;
    if (learned.layer_meanings[layerName]) result.known_layers[layerName]=learned.layer_meanings[layerName];
    else result.unknown_layers.push({name:layerName,entity_count:scanned.layers[layerName].entity_count});
  }
  const usedBlocks=[...new Set(scanned.inserts.map(i=>i.block))];
  for (const blockName of usedBlocks) {
    if (learned.block_meanings[blockName]) result.known_blocks[blockName]=learned.block_meanings[blockName];
    else result.unknown_blocks.push({name:blockName,count:scanned.inserts.filter(i=>i.block===blockName).length});
  }
  return result;
}

// ─────────────────────────────────────────────────────────────────
// SECTION 6 — MAIN EXPORT
// ─────────────────────────────────────────────────────────────────
function analyzeDrawing(dxfContent, filename) {
  const scanned = scanDXF(dxfContent);
  const learned = loadLearned();
  const learnedResult = applyLearned(scanned, learned);
  const rawDump = buildRawDump(scanned, filename);

  // Block counts
  const blockCounts = {};
  for (const ins of scanned.inserts) blockCounts[ins.block]=(blockCounts[ins.block]||0)+1;

  // Layer groups for Excel
  const layerGroups = {};
  function ensureLayer(l) { if(!layerGroups[l]) layerGroups[l]={texts:[],lines:[],dims:[],polylines:[],count:0}; }
  for (const t of scanned.texts)     { ensureLayer(t.layer); layerGroups[t.layer].texts.push(t.text); layerGroups[t.layer].count++; }
  for (const d of scanned.dims)      { ensureLayer(d.layer); layerGroups[d.layer].dims.push(d); layerGroups[d.layer].count++; }
  for (const p of scanned.polylines) { ensureLayer(p.layer); layerGroups[p.layer].polylines.push(p); layerGroups[p.layer].count++; }
  for (const l of scanned.lines_)    { ensureLayer(l.layer); layerGroups[l.layer].lines.push(l); layerGroups[l.layer].count++; }

  // Polyline areas (closed only, area > 0.1m2)
  const polylineAreas = scanned.polylines
    .filter(p=>p.is_closed&&p.area_m2>0.1)
    .sort((a,b)=>b.area_m2-a.area_m2)
    .map(p=>({layer:p.layer,area_sqm:p.area_m2,area_sqft:p.area_sqft,perimeter_m:p.perimeter_m,vertices:p.vertices}));

  const ext=scanned.extents;
  const wMm=isFinite(ext.xmax-ext.xmin)?ext.xmax-ext.xmin:0;
  const hMm=isFinite(ext.ymax-ext.ymin)?ext.ymax-ext.ymin:0;

  // Room annotations — texts that look like room labels
  const roomAnnotations = scanned.texts.filter(t =>
    /BEDROOM|TOILET|KITCHEN|LIVING|DINING|HALL|STORE|LOBBY|OFFICE|LIFT|STAIR|DECK|WASH|BATH|ROOM|CORRIDOR|PASSAGE/i.test(t.text)
  );

  // Inline dims — texts like "3660 X 5285" or "3000x4500"
  const inlineDims = scanned.texts
    .filter(t => /^\d{3,5}\s*[xX×]\s*\d{3,5}/.test(t.text))
    .map(t => {
      const m = t.text.match(/(\d{3,5})\s*[xX×]\s*(\d{3,5})/);
      if (!m) return null;
      const l=parseInt(m[1]),w=parseInt(m[2]);
      return { label:t.text, layer:t.layer, length_mm:l, width_mm:w, area_sqm:Math.round(l*w/1e6*100)/100 };
    }).filter(Boolean);

  return {
    filename,
    drawing_extents:{ width_m:Math.round(wMm/10)/100, height_m:Math.round(hMm/10)/100 },
    stats:{
      total_layers:Object.keys(scanned.layers).length, total_texts:scanned.texts.length,
      total_dims:scanned.dims.length, total_polylines:scanned.polylines.length,
      total_lines:scanned.lines_.length, total_inserts:scanned.inserts.length,
      total_hatches:scanned.hatches.length, unique_blocks:Object.keys(scanned.blocks).length
    },
    // For Excel
    all_texts:       [...new Set(scanned.texts.map(t=>t.text))],
    dimension_values: scanned.dims.map(d=>({value_mm:d.value_mm,value_m:d.value_m,text:d.text_override,layer:d.layer})),
    polyline_areas:  polylineAreas,
    block_counts:    blockCounts,
    layer_groups:    layerGroups,
    room_annotations: roomAnnotations,
    inline_dims:     inlineDims,
    // For AI
    raw_dump:  rawDump,
    learned:   learned,
    // Pre-classified
    known_layers:   learnedResult.known_layers,
    unknown_layers: learnedResult.unknown_layers,
    known_blocks:   learnedResult.known_blocks,
    unknown_blocks: learnedResult.unknown_blocks,
    // Pass to saveLearnedFromAI
    _scanned: scanned,
  };
}

module.exports = { analyzeDrawing, buildAIPrompt, buildRawDump, scanDXF, loadLearned, saveLearned, saveLearnedFromAI };
