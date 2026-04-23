/**
 * PMC Drawing Analyzer v5
 * Supports ALL civil drawing types:
 * FLOOR_PLAN | BASEMENT | PARKING | LIFT_SHAFT | STAIRCASE | STRUCTURAL_SECTION
 * FOUNDATION | SITE_LAYOUT | ROAD_PLAN | MEP_PLUMBING | MEP_ELECTRICAL
 * MEP_HVAC | ELEVATION | DETAIL_DRAWING | GENERAL
 */
const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ── Rates: loaded from Rates.json, never hardcoded ─────────────────────────
let RATES = {};
try {
  const ratesRaw = JSON.parse(fs.readFileSync(path.join(__dirname, 'Rates.json'), 'utf8'));
  for (const grp of Object.values(ratesRaw)) {
    if (typeof grp === 'object' && !Array.isArray(grp)) {
      for (const [k, v] of Object.entries(grp)) {
        if (v && typeof v.rate === 'number') RATES[k] = v.rate;
      }
    }
  }
} catch(e) {
  console.warn('Rates.json load failed:', e.message);
}

// ── Drawing type catalogue ──────────────────────────────────────────────────
const DRAWING_TYPES = {
  FLOOR_PLAN:           'Architectural Floor Plan (rooms, walls, doors, windows)',
  BASEMENT:             'Basement Plan (substructure, raft/pile cap, retaining walls)',
  PARKING:              'Parking Layout (car/two-wheeler bays, ramps, column grid)',
  LIFT_SHAFT:           'Lift / Elevator Shaft & Machine Room Detail',
  STAIRCASE:            'Staircase Detail (flight, landing, railing, nosing)',
  STRUCTURAL_SECTION:   'Structural Section / Detail (column, beam, slab, joint)',
  FOUNDATION:           'Foundation Plan (footings, pile caps, raft, grade beam)',
  SITE_LAYOUT:          'Site / Master Plan (plot boundary, roads, open areas)',
  ROAD_PLAN:            'Road / Infrastructure Layout (carriageway, GSB, WMM, PQC)',
  MEP_PLUMBING:         'MEP — Plumbing & Drainage Layout',
  MEP_ELECTRICAL:       'MEP — Electrical / ELV Layout',
  MEP_HVAC:             'MEP — HVAC / Firefighting Layout',
  ELEVATION:            'Building Elevation / Facade Drawing',
  DETAIL_DRAWING:       'Detail Drawing (joint, section cut, connection detail)',
  GENERAL:              'General / Unknown Drawing Type',
};

// ── Drawing-type → BOQ template mapping ────────────────────────────────────
// Each entry: { description, unit, factor_fn(area, counts, dims) }
// factor_fn receives (totalArea_sqm, elementCounts, parsedDims)
const BOQ_TEMPLATES = {

  FLOOR_PLAN: [
    { desc: 'RCC Slab M25 @ 150mm thick',              unit:'sqmt', factor:(a)=>a,          rk:'rcc_m25_cum' },
    { desc: 'RCC Beams M25',                           unit:'cum',  factor:(a)=>a*0.04,     rk:'rcc_m25_cum' },
    { desc: 'RCC Columns M30',                         unit:'cum',  factor:(a)=>a*0.02,     rk:'rcc_m30_cum' },
    { desc: 'Steel Reinforcement Fe500 (slab+beams)',  unit:'kg',   factor:(a)=>a*14,       rk:'steel_fe500_kg' },
    { desc: 'Brickwork 230mm (external walls)',        unit:'cum',  factor:(a)=>a*0.04,     rk:'brickwork_230mm_cum' },
    { desc: 'Brickwork 115mm (internal partitions)',   unit:'cum',  factor:(a)=>a*0.05,     rk:'brickwork_115mm_cum' },
    { desc: 'Plaster 12mm internal (both faces)',      unit:'sqmt', factor:(a)=>a*2.8,      rk:'plaster_15mm_sqmt' },
    { desc: 'Plaster 20mm external',                  unit:'sqmt', factor:(a)=>a*0.6,      rk:'plaster_15mm_sqmt' },
    { desc: 'Vitrified tile flooring 600×600',        unit:'sqmt', factor:(a)=>a*0.80,     rk:'flooring_vitrified_sqmt' },
    { desc: 'Ceramic/Anti-skid tile (toilets)',        unit:'sqmt', factor:(a)=>a*0.12,     rk:'flooring_ceramic_sqmt' },
    { desc: 'Dado tile (bathrooms) up to 2.1m ht',    unit:'sqmt', factor:(a)=>a*0.30,     rk:'tile_dado_sqmt' },
    { desc: 'OBD Internal paint (2 coats)',            unit:'sqmt', factor:(a)=>a*2.8,      rk:'painting_sqmt' },
    { desc: 'Texture paint external',                 unit:'sqmt', factor:(a)=>a*0.6,      rk:'painting_ext_sqmt' },
    { desc: 'UPVC Windows',                           unit:'sqmt', factor:(a)=>a*0.12,     rk:'window_upvc_sqmt' },
    { desc: 'Main entrance door (teak/flush)',         unit:'nos',  factor:(a,c)=>Math.max(c.door_count||1,Math.round(a/120)), rk:'door_main_nos' },
    { desc: 'Internal doors flush',                   unit:'nos',  factor:(a,c)=>Math.max(c.door_count||1,Math.round(a/30))*3, rk:'door_internal_nos' },
    { desc: 'Waterproofing (terrace+bathrooms)',       unit:'sqmt', factor:(a)=>a*0.20,     rk:'waterproofing_sqmt' },
    { desc: 'Formwork (slab+beams+columns)',           unit:'sqmt', factor:(a)=>a*2.5,      rk:'formwork_sqmt' },
    { desc: 'Electrical rough-in (per sqmt BUA)',      unit:'sqmt', factor:(a)=>a,          rk:'electrical_sqmt' },
    { desc: 'Plumbing rough-in (per sqmt BUA)',        unit:'sqmt', factor:(a)=>a,          rk:'plumbing_sqmt' },
  ],

  BASEMENT: [
    { desc: 'Excavation in all soils',                unit:'cum',  factor:(a)=>a*3.5,      rk:'excavation_cum' },
    { desc: 'PCC M10 blinding 75mm',                  unit:'cum',  factor:(a)=>a*0.075,    rk:'pcc_m10_cum' },
    { desc: 'Raft / Mat foundation RCC M30',          unit:'cum',  factor:(a)=>a*0.6,      rk:'rcc_m30_cum' },
    { desc: 'Retaining wall RCC M30 (perimeter)',     unit:'cum',  factor:(a)=>a*0.08,     rk:'rcc_m30_cum' },
    { desc: 'Basement columns RCC M30',               unit:'cum',  factor:(a)=>a*0.025,    rk:'rcc_m30_cum' },
    { desc: 'Steel reinforcement Fe500 (raft)',       unit:'kg',   factor:(a)=>a*0.6*120,  rk:'steel_fe500_kg' },
    { desc: 'Formwork (all RCC elements)',            unit:'sqmt', factor:(a)=>a*2.0,      rk:'formwork_sqmt' },
    { desc: 'Backfilling with excavated soil',        unit:'cum',  factor:(a)=>a*0.8,      rk:'backfilling_cum' },
    { desc: 'Waterproofing — basement raft & walls',  unit:'sqmt', factor:(a)=>a*1.5,      rk:'waterproofing_sqmt' },
    { desc: 'Drainage board + geocomposite membrane', unit:'sqmt', factor:(a)=>a*0.6,      rk:'waterproofing_sqmt' },
    { desc: 'Sump pit RCC M30',                       unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/1000)), rk:'sump_nos' },
    { desc: 'Dewatering during construction',         unit:'lump', factor:(a)=>1,          rk:'dewatering_ls' },
    { desc: 'Podium slab (top of basement) RCC M25', unit:'sqmt', factor:(a)=>a,          rk:'rcc_m25_cum' },
  ],

  PARKING: [
    { desc: 'PCC M15 flooring 100mm (parking)',       unit:'sqmt', factor:(a)=>a,          rk:'pcc_m10_cum' },
    { desc: 'RCC Column grid M30 (parking level)',    unit:'cum',  factor:(a)=>a*0.018,    rk:'rcc_m30_cum' },
    { desc: 'Overhead beam M25 (grid beams)',         unit:'cum',  factor:(a)=>a*0.03,     rk:'rcc_m25_cum' },
    { desc: 'Steel reinforcement Fe500 (columns+beams)',unit:'kg', factor:(a)=>a*6,        rk:'steel_fe500_kg' },
    { desc: 'Formwork — columns + beams',             unit:'sqmt', factor:(a)=>a*1.2,      rk:'formwork_sqmt' },
    { desc: 'Line markings — car bays (paint)',       unit:'rmt',  factor:(a)=>a*0.08,     rk:'painting_sqmt' },
    { desc: 'Speed bump — RCC @ 500×150mm',           unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/500)), rk:'rcc_m25_cum' },
    { desc: 'Wheel stops — precast concrete',         unit:'nos',  factor:(a)=>Math.round(a/15), rk:'rcc_m25_cum' },
    { desc: 'Signage boards (parking directions)',    unit:'nos',  factor:(a)=>Math.max(4,Math.round(a/300)), rk:'streetlight_nos' },
    { desc: 'Ventilation ducts rough-in (parking)',   unit:'rmt',  factor:(a)=>a*0.05,     rk:'pipeline_rmt' },
    { desc: 'Fire sprinkler provision (stub-ins)',    unit:'nos',  factor:(a)=>Math.round(a/9), rk:'pipeline_rmt' },
    { desc: 'Parking ramp RCC M25 (if ramp visible)',unit:'sqmt', factor:(a)=>a*0.06,     rk:'rcc_m25_cum' },
  ],

  LIFT_SHAFT: [
    { desc: 'Lift pit excavation & PCC blinding',    unit:'cum',  factor:(a,c)=>Math.max(c.lift_count||1,1)*3, rk:'excavation_cum' },
    { desc: 'Lift pit RCC M30 (walls+slab)',          unit:'cum',  factor:(a,c)=>Math.max(c.lift_count||1,1)*4, rk:'rcc_m30_cum' },
    { desc: 'Lift shaft walls RCC M25 (per floor)',  unit:'cum',  factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.lift_count||1,1)*0.8, rk:'rcc_m25_cum' },
    { desc: 'Machine room slab + walls RCC M25',     unit:'cum',  factor:(a,c)=>Math.max(c.lift_count||1,1)*3, rk:'rcc_m25_cum' },
    { desc: 'Steel reinforcement Fe500 (shaft)',     unit:'kg',   factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.lift_count||1,1)*100, rk:'steel_fe500_kg' },
    { desc: 'Formwork — shaft walls (all floors)',   unit:'sqmt', factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.lift_count||1,1)*12, rk:'formwork_sqmt' },
    { desc: 'Lift door opening lintel RCC M25',      unit:'nos',  factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.lift_count||1,1), rk:'rcc_m25_cum' },
    { desc: 'Grouting — lift guide rail pockets',   unit:'nos',  factor:(a,c)=>Math.max(c.lift_count||1,1)*20, rk:'grouting_nos' },
    { desc: 'Plaster 12mm — shaft interior walls',  unit:'sqmt', factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.lift_count||1,1)*12, rk:'plaster_15mm_sqmt' },
    { desc: 'Waterproofing — lift pit',              unit:'sqmt', factor:(a,c)=>Math.max(c.lift_count||1,1)*12, rk:'waterproofing_sqmt' },
  ],

  STAIRCASE: [
    { desc: 'RCC Staircase M25 (flights + landings)',unit:'cum',  factor:(a,c)=>Math.max(c.staircase_count||1,1)*2.5, rk:'rcc_m25_cum' },
    { desc: 'Steel Fe500 in staircase',              unit:'kg',   factor:(a,c)=>Math.max(c.staircase_count||1,1)*300, rk:'steel_fe500_kg' },
    { desc: 'Formwork — staircase soffit + sides',  unit:'sqmt', factor:(a,c)=>Math.max(c.staircase_count||1,1)*18, rk:'formwork_sqmt' },
    { desc: 'Kota / Granite nosing on steps',        unit:'rmt',  factor:(a,c)=>Math.max(c.staircase_count||1,1)*12*1.2, rk:'flooring_vitrified_sqmt' },
    { desc: 'Landing tile flooring',                 unit:'sqmt', factor:(a,c)=>Math.max(c.staircase_count||1,1)*5, rk:'flooring_vitrified_sqmt' },
    { desc: 'MS railing with SS handrail',           unit:'rmt',  factor:(a,c)=>Math.max(c.staircase_count||1,1)*24, rk:'railing_ms_rmt' },
    { desc: 'Plaster 12mm — staircase walls',        unit:'sqmt', factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.staircase_count||1,1)*8, rk:'plaster_15mm_sqmt' },
    { desc: 'OBD paint — staircase walls',           unit:'sqmt', factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.staircase_count||1,1)*8, rk:'painting_sqmt' },
    { desc: 'Emergency lighting provision',          unit:'nos',  factor:(a,c,d)=>(d.floor_count||10)*Math.max(c.staircase_count||1,1)*2, rk:'streetlight_nos' },
  ],

  STRUCTURAL_SECTION: [
    { desc: 'RCC M30 (columns as per BBS)',          unit:'cum',  factor:(a,c)=>a*0.025,  rk:'rcc_m30_cum' },
    { desc: 'RCC M25 (beams as per BBS)',             unit:'cum',  factor:(a,c)=>a*0.04,   rk:'rcc_m25_cum' },
    { desc: 'RCC M25 (slab as per BBS)',              unit:'sqmt', factor:(a,c)=>a,        rk:'rcc_m25_cum' },
    { desc: 'Steel Fe500 main bars',                 unit:'kg',   factor:(a,c)=>a*20,     rk:'steel_fe500_kg' },
    { desc: 'Steel Fe500 stirrups/ties',             unit:'kg',   factor:(a,c)=>a*5,      rk:'steel_fe500_kg' },
    { desc: 'Formwork — columns (4-side)',            unit:'sqmt', factor:(a,c)=>a*0.8,    rk:'formwork_sqmt' },
    { desc: 'Formwork — beams (3-side)',              unit:'sqmt', factor:(a,c)=>a*1.2,    rk:'formwork_sqmt' },
    { desc: 'Formwork — slab soffit',                unit:'sqmt', factor:(a,c)=>a,        rk:'formwork_sqmt' },
    { desc: 'Rebar binding wire @ 8 kg/tonne',       unit:'kg',   factor:(a,c)=>a*0.2,    rk:'steel_fe500_kg' },
    { desc: 'Cover blocks (concrete spacers)',        unit:'nos',  factor:(a,c)=>a*8,      rk:'rcc_m25_cum' },
  ],

  FOUNDATION: [
    { desc: 'Excavation for individual footings',    unit:'cum',  factor:(a)=>a*1.5,      rk:'excavation_cum' },
    { desc: 'PCC M10 blinding 75mm',                 unit:'cum',  factor:(a)=>a*0.075,    rk:'pcc_m10_cum' },
    { desc: 'RCC M25 — isolated footings',           unit:'cum',  factor:(a)=>a*0.4,      rk:'rcc_m25_cum' },
    { desc: 'RCC M25 — grade beams / tie beams',     unit:'cum',  factor:(a)=>a*0.06,     rk:'rcc_m25_cum' },
    { desc: 'RCC M30 — pile cap (if piled fdn)',     unit:'cum',  factor:(a)=>a*0.15,     rk:'rcc_m30_cum' },
    { desc: 'Steel Fe500 (footings + grade beams)',  unit:'kg',   factor:(a)=>a*55,       rk:'steel_fe500_kg' },
    { desc: 'Formwork — footings + grade beams',     unit:'sqmt', factor:(a)=>a*1.2,      rk:'formwork_sqmt' },
    { desc: 'Backfilling — compacted in layers',     unit:'cum',  factor:(a)=>a*0.8,      rk:'backfilling_cum' },
    { desc: 'Waterproofing — foundation raft',       unit:'sqmt', factor:(a)=>a,          rk:'waterproofing_sqmt' },
    { desc: 'Anti-termite treatment (soil)',         unit:'sqmt', factor:(a)=>a,          rk:'waterproofing_sqmt' },
  ],

  SITE_LAYOUT: [
    { desc: 'Site clearing + grubbing',              unit:'sqmt', factor:(a)=>a,          rk:'excavation_cum' },
    { desc: 'Compound wall 230mm brick + plastered', unit:'rmt',  factor:(a)=>Math.round(Math.sqrt(a)*4), rk:'compound_wall_rmt' },
    { desc: 'Security cabin RCC M20',                unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/5000)), rk:'rcc_m20_cum' },
    { desc: 'Internal roads — PQC 200mm',            unit:'sqmt', factor:(a)=>a*0.15,     rk:'pqc_road_250mm_sqmt' },
    { desc: 'Footpath — paver blocks 60mm',          unit:'sqmt', factor:(a)=>a*0.08,     rk:'paver_block_80mm_sqmt' },
    { desc: 'Landscaping & turf (open areas)',       unit:'sqmt', factor:(a)=>a*0.30,     rk:'paver_block_80mm_sqmt' },
    { desc: 'Storm water drain RCC / NP3 pipes',     unit:'rmt',  factor:(a)=>Math.round(Math.sqrt(a)*4), rk:'pipeline_rmt' },
    { desc: 'Electrical — street lights on roads',   unit:'nos',  factor:(a)=>Math.round(Math.sqrt(a)*4/25), rk:'streetlight_nos' },
    { desc: 'Water supply — main line HDPE',         unit:'rmt',  factor:(a)=>Math.round(Math.sqrt(a)*4*0.6), rk:'pipeline_rmt' },
    { desc: 'UGT / Sump (underground tank)',         unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/10000)), rk:'sump_nos' },
    { desc: 'Entry gate — mild steel',               unit:'nos',  factor:(a)=>1,          rk:'railing_ms_rmt' },
  ],

  ROAD_PLAN: [
    { desc: 'Box cutting (avg 450mm depth)',         unit:'sqmt', factor:(a)=>a*1.05,     rk:'excavation_cum' },
    { desc: 'GSB 300mm thick, crushed stone',        unit:'sqmt', factor:(a)=>a,          rk:'gsb_300mm_sqmt' },
    { desc: 'WMM 200mm thick, graded metal',         unit:'sqmt', factor:(a)=>a,          rk:'wmm_200mm_sqmt' },
    { desc: 'PQC M30 250mm thick, DLC below',        unit:'sqmt', factor:(a)=>a,          rk:'pqc_road_250mm_sqmt' },
    { desc: 'Dowel bars Fe500 @ 3.87 kg/sqmt',       unit:'kg',   factor:(a)=>a*3.87,     rk:'steel_fe500_kg' },
    { desc: 'Tie bars Fe415 (longitudinal)',          unit:'kg',   factor:(a)=>a*0.5,      rk:'steel_fe415_kg' },
    { desc: 'Kerbing — M30 precast 300×150',         unit:'rmt',  factor:(a)=>Math.round(a/7)*2, rk:'kerbing_rmt' },
    { desc: 'Road marking paint (thermoplastic)',    unit:'rmt',  factor:(a)=>Math.round(a/7)*0.3, rk:'painting_sqmt' },
    { desc: 'Footpath — paver blocks 80mm',          unit:'sqmt', factor:(a)=>a*0.25,     rk:'paver_block_80mm_sqmt' },
    { desc: 'Storm water drain NP3 / NP4',           unit:'rmt',  factor:(a)=>Math.round(a/7), rk:'pipeline_rmt' },
    { desc: 'Median — precast concrete blocks',      unit:'rmt',  factor:(a)=>Math.round(a/7*0.5), rk:'kerbing_rmt' },
    { desc: 'Street lighting (30m spacing)',         unit:'nos',  factor:(a)=>Math.round(a/7*2/30), rk:'streetlight_nos' },
  ],

  MEP_PLUMBING: [
    { desc: 'CPVC water supply pipes (15–32mm)',     unit:'rmt',  factor:(a)=>a*0.18,     rk:'pipeline_rmt' },
    { desc: 'UPVC drainage pipes (75–150mm)',        unit:'rmt',  factor:(a)=>a*0.15,     rk:'pipeline_rmt' },
    { desc: 'Sanitary ware — EWC+wash basin set',   unit:'nos',  factor:(a,c)=>Math.max(c.toilet_count||1,Math.round(a/80)), rk:'streetlight_nos' },
    { desc: 'Concealed stop cocks + valves',         unit:'nos',  factor:(a,c)=>Math.max(c.toilet_count||1,Math.round(a/80))*3, rk:'streetlight_nos' },
    { desc: 'GI pipes (50–100mm) — riser/stack',    unit:'rmt',  factor:(a,c)=>a*0.03,   rk:'pipeline_rmt' },
    { desc: 'Water tank — RCC overhead',             unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/2000)), rk:'rcc_m25_cum' },
    { desc: 'Pressure pump set',                    unit:'nos',  factor:(a)=>1,          rk:'streetlight_nos' },
    { desc: 'Water meter + PRV assembly',            unit:'nos',  factor:(a)=>Math.round(a/120), rk:'streetlight_nos' },
    { desc: 'Manhole RCC (SWD + sewage)',            unit:'nos',  factor:(a)=>Math.round(a/500), rk:'rcc_m25_cum' },
    { desc: 'Chasing + grouting for concealed pipes',unit:'rmt',  factor:(a)=>a*0.12,     rk:'grouting_nos' },
    { desc: 'STP (Sewage Treatment Plant) — civil', unit:'lump', factor:(a)=>1,          rk:'rcc_m30_cum' },
  ],

  MEP_ELECTRICAL: [
    { desc: 'PVC conduit 20–32mm concealed in slab',unit:'rmt',  factor:(a)=>a*0.25,     rk:'pipeline_rmt' },
    { desc: 'DB (distribution board) provision',    unit:'nos',  factor:(a)=>Math.round(a/120), rk:'streetlight_nos' },
    { desc: 'Light point (concealed wiring)',        unit:'nos',  factor:(a)=>Math.round(a/8), rk:'streetlight_nos' },
    { desc: 'Power point (16A socket)',              unit:'nos',  factor:(a)=>Math.round(a/12), rk:'streetlight_nos' },
    { desc: 'Main LT cable (XLPE armoured)',         unit:'rmt',  factor:(a)=>a*0.02,     rk:'pipeline_rmt' },
    { desc: 'Earthing station (GI pipe)',            unit:'nos',  factor:(a)=>Math.max(2,Math.round(a/2000)), rk:'streetlight_nos' },
    { desc: 'DG set provision (civil foundation)',   unit:'nos',  factor:(a)=>1,          rk:'rcc_m25_cum' },
    { desc: 'Lightning arrestor system',             unit:'nos',  factor:(a)=>1,          rk:'streetlight_nos' },
    { desc: 'Solar PV provision — terrace roughing',unit:'sqmt', factor:(a)=>a*0.10,     rk:'electrical_sqmt' },
    { desc: 'EV charging provision (civil stub)',    unit:'nos',  factor:(a)=>Math.round(a/200), rk:'streetlight_nos' },
  ],

  MEP_HVAC: [
    { desc: 'AHU room — RCC provision',             unit:'nos',  factor:(a)=>Math.max(1,Math.round(a/1500)), rk:'rcc_m25_cum' },
    { desc: 'Chilled water pipe sleeve in slab',    unit:'nos',  factor:(a)=>Math.round(a/100), rk:'pipeline_rmt' },
    { desc: 'Duct shaft opening in slab',           unit:'nos',  factor:(a)=>Math.round(a/400), rk:'rcc_m25_cum' },
    { desc: 'Fire damper provision (civil block-outs)',unit:'nos',factor:(a)=>Math.round(a/200), rk:'rcc_m25_cum' },
    { desc: 'Terrace unit base — PCC pad 100mm',    unit:'nos',  factor:(a)=>Math.round(a/200), rk:'pcc_m10_cum' },
    { desc: 'Cooling tower — RCC foundation',       unit:'nos',  factor:(a)=>1,          rk:'rcc_m25_cum' },
    { desc: 'Sprinkler head provision (stub in RCC)',unit:'nos', factor:(a)=>Math.round(a/9), rk:'pipeline_rmt' },
    { desc: 'Fire hose cabinet recess in wall',     unit:'nos',  factor:(a)=>Math.round(a/500), rk:'rcc_m25_cum' },
  ],

  ELEVATION: [
    { desc: 'External plaster 20mm cement sand 1:4',unit:'sqmt', factor:(a)=>a,          rk:'plaster_15mm_sqmt' },
    { desc: 'Texture paint / weather coat',         unit:'sqmt', factor:(a)=>a,          rk:'painting_ext_sqmt' },
    { desc: 'Cladding — stone/vitrified facade',   unit:'sqmt', factor:(a)=>a*0.30,     rk:'flooring_vitrified_sqmt' },
    { desc: 'Aluminum composite panel cladding',   unit:'sqmt', factor:(a)=>a*0.15,     rk:'window_upvc_sqmt' },
    { desc: 'UPVC / Aluminium windows (elevation)', unit:'sqmt', factor:(a)=>a*0.18,     rk:'window_upvc_sqmt' },
    { desc: 'Aluminium structural glazing',         unit:'sqmt', factor:(a)=>a*0.10,     rk:'window_upvc_sqmt' },
    { desc: 'Chajja / sunshade RCC M25',            unit:'rmt',  factor:(a)=>Math.round(a/3)*0.5, rk:'rcc_m25_cum' },
    { desc: 'Parapet wall — RCC / masonry',         unit:'rmt',  factor:(a)=>Math.round(Math.sqrt(a)*2), rk:'brickwork_230mm_cum' },
    { desc: 'Waterproofing — parapet top & sill',  unit:'rmt',  factor:(a)=>Math.round(Math.sqrt(a)*2)*0.3, rk:'waterproofing_sqmt' },
    { desc: 'External scaffolding — rental cost',  unit:'sqmt', factor:(a)=>a,          rk:'formwork_sqmt' },
  ],

  DETAIL_DRAWING: [
    { desc: 'RCC M30 (as per section detail)',      unit:'cum',  factor:(a,c)=>a*0.03,   rk:'rcc_m30_cum' },
    { desc: 'Steel Fe500 (main bars per BBS)',       unit:'kg',   factor:(a,c)=>a*25,     rk:'steel_fe500_kg' },
    { desc: 'Formwork — as per detail section',     unit:'sqmt', factor:(a,c)=>a*1.5,    rk:'formwork_sqmt' },
    { desc: 'Non-shrink grout (connections)',        unit:'nos',  factor:(a,c)=>Math.round(a*0.5), rk:'grouting_nos' },
    { desc: 'Anchor bolts M20/M24 (structural)',    unit:'nos',  factor:(a,c)=>Math.round(a*2), rk:'rcc_m25_cum' },
    { desc: 'Waterproofing membrane (joint)',        unit:'rmt',  factor:(a,c)=>a*0.3,    rk:'waterproofing_sqmt' },
  ],

  GENERAL: [
    { desc: 'Civil works — general (thumb-rule)',   unit:'sqmt', factor:(a)=>a,          rk:'rcc_m25_cum' },
    { desc: 'Steel reinforcement (estimated)',       unit:'kg',   factor:(a)=>a*14,       rk:'steel_fe500_kg' },
    { desc: 'Formwork (estimated)',                 unit:'sqmt', factor:(a)=>a*2.0,      rk:'formwork_sqmt' },
    { desc: 'Plaster (internal)',                   unit:'sqmt', factor:(a)=>a*2.5,      rk:'plaster_15mm_sqmt' },
    { desc: 'Flooring (vitrified tiles)',            unit:'sqmt', factor:(a)=>a*0.85,     rk:'flooring_vitrified_sqmt' },
    { desc: 'Painting (internal, 2 coats)',          unit:'sqmt', factor:(a)=>a*2.5,      rk:'painting_sqmt' },
  ],
};

// ── Get BOQ for any drawing type ────────────────────────────────────────────
function getBOQForDrawingType(drawingType, totalArea, elementCounts, parsedDims) {
  const template = BOQ_TEMPLATES[drawingType] || BOQ_TEMPLATES.GENERAL;
  const ec = elementCounts || {};
  const pd = parsedDims || {};

  return template.map(item => {
    const qty  = Math.round(item.factor(totalArea, ec, pd) * 100) / 100;
    const rate = RATES[item.rk] || 0;
    return { desc: item.desc, unit: item.unit, qty, rate, amount: Math.round(qty * rate), rateKey: item.rk };
  }).filter(item => item.qty > 0 && item.rate > 0);
}

// ── Detect drawing type from text/layer keywords ────────────────────────────
function detectDrawingType(texts, layers, filename) {
  const combined = ([...texts, ...layers, filename || '']).join(' ').toUpperCase();

  if (/BASEMENT|SUBSTRUCTURE|RAFT|RETAINING WALL|DEWATER/i.test(combined)) return 'BASEMENT';
  if (/PARKING|CAR PARK|VEHICLE|RAMP.*PARK|PARK.*RAMP|LEVEL.*PARK|TWO WHEEL/i.test(combined)) return 'PARKING';
  if (/LIFT|ELEVATOR|MACHINE ROOM|MR.*LIFT|LIFT.*SHAFT|LIFT.*PIT|SHAFT/i.test(combined)) return 'LIFT_SHAFT';
  if (/STAIRCASE|STAIR CASE|STAIR.*FLIGHT|FLIGHT.*STAIR|STAIRWELL|NOSING/i.test(combined)) return 'STAIRCASE';
  if (/FOUNDATION|FOOTING|PILE CAP|PILE.*PLAN|GRADE BEAM|STRAP BEAM/i.test(combined)) return 'FOUNDATION';
  if (/ROAD|CARRIAGEWAY|GSB|WMM|PQC|KERB|HIGHWAY|IRC|TRAFFIC|PAVEMENT/i.test(combined)) return 'ROAD_PLAN';
  if (/SITE PLAN|MASTER PLAN|SITE LAYOUT|PLOT BOUNDARY|COMPOUND WALL|SETBACK/i.test(combined)) return 'SITE_LAYOUT';
  if (/PLUMBING|DRAINAGE|SANITARY|SEWER|STP|SWD|WATER SUPPLY|MANHOLE/i.test(combined)) return 'MEP_PLUMBING';
  if (/ELECTRICAL|ELV|EARTHING|LT PANEL|DB.*LAYOUT|CONDUIT.*LAYOUT/i.test(combined)) return 'MEP_ELECTRICAL';
  if (/HVAC|AIR CONDITION|DUCT|AHU|COOLING TOWER|FIREFIGHTING|SPRINKLER/i.test(combined)) return 'MEP_HVAC';
  if (/ELEVATION|FACADE|CLADDING|EXTERIOR.*VIEW|FRONT.*VIEW/i.test(combined)) return 'ELEVATION';
  if (/SECTION|DETAIL|JUNCTION|JOINT DETAIL|CONNECTION DETAIL|BBS|BAR BENDING/i.test(combined)) return 'STRUCTURAL_SECTION';
  if (/STRUCTURAL|COLUMN DETAIL|BEAM DETAIL|SLAB DETAIL|FOOTING DETAIL/i.test(combined)) return 'STRUCTURAL_SECTION';
  if (/FLOOR PLAN|ROOM|BED ROOM|BEDROOM|KITCHEN|TOILET|LIVING|DRAWING ROOM|FLAT|UNIT/i.test(combined)) return 'FLOOR_PLAN';
  if (/PLAN/i.test(combined)) return 'FLOOR_PLAN';
  return 'GENERAL';
}

// ── Quantity helper functions ────────────────────────────────────────────────
function calcRoadQuantities(length_m, width_m) {
  const carriageWidth = Math.max(width_m - 3, width_m * 0.65);
  const area = length_m * carriageWidth;
  return {
    area_sqmt:        Math.round(area * 100) / 100,
    box_cutting_sqmt: Math.round(area * 1.05 * 100) / 100,
    gsb_300mm_ton:    Math.round(area * 1.15 * 0.300 * 1.8 * 100) / 100,
    wmm_200mm_ton:    Math.round(area * 1.15 * 0.200 * 2.1 * 100) / 100,
    pqc_250mm_cum:    Math.round(area * 1.05 * 0.250 * 100) / 100,
    steel_dowel_kg:   Math.round(area * 3.87),
    cost_estimate: {
      gsb:       Math.round(area * (RATES.gsb_300mm_sqmt || 655)),
      wmm:       Math.round(area * (RATES.wmm_200mm_sqmt || 515)),
      pqc:       Math.round(area * (RATES.pqc_road_250mm_sqmt || 1800)),
      total_sqmt:Math.round(area * ((RATES.gsb_300mm_sqmt||655) + (RATES.wmm_200mm_sqmt||515) + (RATES.pqc_road_250mm_sqmt||1800))),
    }
  };
}

function calcStructureQuantities(dims) {
  const { length=0, width=0, height=0, nos=1 } = dims;
  const volume = length * width * height * nos;
  const area   = length * width * nos;
  return {
    volume_cum:    Math.round(volume * 1000) / 1000,
    area_sqmt:     Math.round(area * 100) / 100,
    steel_kg:      Math.round(volume * 120),
    formwork_sqmt: Math.round((2*(length+width)*height + area) * nos * 100) / 100,
  };
}

// ── CV Analysis ─────────────────────────────────────────────────────────────
function runCVAnalysis(b64Image) {
  try {
    const tmpFile = path.join(os.tmpdir(), `drawing_cv_${Date.now()}.txt`);
    fs.writeFileSync(tmpFile, b64Image);
    const result = execSync(`python3 ${path.join(__dirname,'drawing_cv.py')} ${tmpFile}`, { timeout: 30000 });
    fs.unlinkSync(tmpFile);
    return JSON.parse(result.toString());
  } catch(e) { return { error: e.message }; }
}

// ── Gemini Drawing Analysis — adaptive prompt per drawing type ───────────────
async function geminiAnalyzeDrawing(key, files, cvData, fetch) {
  const GEMINI_URL = k => `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;

  const parts = [];
  for (const f of (files || [])) {
    if (f.type === 'application/pdf' || f.name?.match(/\.pdf$/i))
      parts.push({ inline_data: { mime_type: 'application/pdf', data: f.b64 } });
    else if (f.type?.startsWith('image/'))
      parts.push({ inline_data: { mime_type: f.type || 'image/png', data: f.b64 } });
  }

  const cvHints = cvData && !cvData.error ? `\nCV PRE-ANALYSIS:\n- Size: ${cvData.image_dimensions?.width_px}×${cvData.image_dimensions?.height_px}px\n- Spaces detected: ${cvData.detected_spaces?.length || 0}\n- Horizontal dims: ${cvData.dimension_lines?.horizontal?.length || 0} | Vertical: ${cvData.dimension_lines?.vertical?.length || 0}\n- Scale hints: ${cvData.scale_interpretation_hints?.join(' | ')}\n` : '';

  // Build rate hint from loaded RATES
  const rateHint = Object.entries(RATES).slice(0,25).map(([k,v])=>`${k}:₹${v}`).join(' | ');

  const drawingTypesStr = Object.entries(DRAWING_TYPES).map(([k,v])=>`${k}=${v}`).join(' | ');

  const prompt = `You are a SENIOR PMC CIVIL ENGINEER (20+ years, India). Analyze THIS drawing ONLY.

${cvHints}

CRITICAL RULES:
1. DO NOT invent or guess any value. Read from THIS drawing only.
2. Every number must trace back to a visible annotation or measured line.
3. If a value is not visible, use 0 / "" / [].
4. Set confidence=LOW if title block / scale / dimensions are not readable.

DRAWING TYPE DETECTION — pick the SINGLE best match from:
${drawingTypesStr}

EXTRACTION CHECKLIST:
1. Scale (1:50 / 1:100 / 1:200 / scale bar)
2. All dimension annotations
3. Title block (project, drawing no, date, floor name, block/building name)
4. Element counts (doors, windows, lifts, staircases, columns, rooms, parking bays, car count, toilet count, kitchen count, bedroom count)
5. Floor level annotations (GF/FF/B1/B2/+xyz mm)
6. Special elements per type:
   - BASEMENT/PARKING: bay sizes, ramp grade, column grid spacing
   - LIFT_SHAFT: car size, shaft size, pit depth, overhead clearance, number of lifts
   - STAIRCASE: rise/tread, flight width, number of flights, railing type
   - STRUCTURAL_SECTION: member size (b×d), main bar dia, stirrup spacing, cover
   - ROAD_PLAN: road name, carriageway width, total length, cross-section layers
   - MEP_*: pipe dia, pipe material, invert levels, fixture counts
   - FOUNDATION: footing size (L×B×D), grade, pile dia/depth if piled
   - ELEVATION: floor heights, window sizes, cladding material, parapet height
   - SITE_LAYOUT: plot area, setbacks (front/side/rear), road widths, open area %

FORMULAS (apply only to values actually read):
ROAD: area=L×Wcarriage | GSB(t)=area×1.15×0.3×1.8 | WMM(t)=area×1.15×0.2×2.1 | PQC(cum)=area×1.05×0.25
BUILDING slab: area×thickness=cum | steel=cum×120 kg/cum | beam: cum×160 kg/cum
FOOTING: L×B×D=cum | Column: S²×H=cum
Lift shaft: perimeter×floor_height×thickness=cum per floor

RATES (from Gujarat DSR 2025): ${rateHint}

Return ONLY raw JSON (no markdown):
{
  "project_name": "",
  "drawing_no": "",
  "drawing_type": "FLOOR_PLAN|BASEMENT|PARKING|LIFT_SHAFT|STAIRCASE|STRUCTURAL_SECTION|FOUNDATION|SITE_LAYOUT|ROAD_PLAN|MEP_PLUMBING|MEP_ELECTRICAL|MEP_HVAC|ELEVATION|DETAIL_DRAWING|GENERAL",
  "floor_name": "",
  "building_name": "",
  "scale": "",
  "date": "",
  "north_direction": "",

  "elements": [
    {
      "id": "",
      "type": "ROAD|ROOM|WALL|COLUMN|BEAM|SLAB|FOOTING|STAIRCASE|LIFT_SHAFT|PARKING_BAY|RAMP|SUMP|PIPE|etc",
      "name": "",
      "dimensions": { "length_m": 0, "width_m": 0, "height_m": 0, "thickness_m": 0, "diameter_m": 0, "nos": 0, "note": "" },
      "quantities": { "area_sqmt": 0, "volume_cum": 0, "length_rmt": 0, "steel_kg": 0, "formwork_sqmt": 0 },
      "cost_inr": { "total": 0 },
      "confidence": "LOW|MEDIUM|HIGH",
      "annotation_found": ""
    }
  ],

  "element_counts": {
    "door_count": 0, "window_count": 0, "lift_count": 0, "staircase_count": 0,
    "column_count": 0, "footing_count": 0, "bedroom_count": 0, "toilet_count": 0,
    "kitchen_count": 0, "floor_count": 0, "parking_bay_count": 0,
    "ramp_count": 0, "sump_count": 0, "balcony_count": 0
  },

  "special_data": {
    "lift_car_size": "", "lift_shaft_size": "", "lift_pit_depth_m": 0, "lift_overhead_m": 0,
    "stair_rise_mm": 0, "stair_tread_mm": 0, "stair_flight_width_m": 0,
    "column_size": "", "beam_size": "", "main_bar_dia_mm": 0, "stirrup_spacing_mm": 0, "cover_mm": 0,
    "parking_bay_size": "", "ramp_grade_pct": 0,
    "plot_area_sqm": 0, "setback_front_m": 0, "setback_side_m": 0,
    "road_carriageway_width_m": 0, "road_total_length_m": 0,
    "pile_dia_mm": 0, "pile_depth_m": 0,
    "pipe_dia_mm": 0, "pipe_material": "", "invert_level_m": 0,
    "floor_height_m": 0, "parapet_height_m": 0, "cladding_material": ""
  },

  "floor_levels": [{ "name": "", "level_m": 0 }],
  "spaces": [{ "name": "", "area_sqm": 0 }],

  "total_quantities": {
    "total_area_sqmt": 0, "total_road_rmt": 0, "rcc_total_cum": 0,
    "steel_total_kg": 0, "brickwork_total_cum": 0, "formwork_total_sqmt": 0,
    "excavation_total_cum": 0
  },

  "cost_summary": { "civil_total_inr": 0, "civil_total_lacs": 0, "civil_total_crores": 0 },
  "bbs_data": [],
  "observations": [],
  "pmc_recommendation": "",
  "extraction_confidence": "LOW|MEDIUM|HIGH",
  "missing_info": []
}`;

  parts.push({ text: prompt });

  const r = await fetch(GEMINI_URL(key), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ role: 'user', parts }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.0, responseMimeType: 'application/json' }
    })
  });

  let raw = (await r.json())?.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const fb = raw.indexOf('{'), lb = raw.lastIndexOf('}');
  if (fb !== -1 && lb !== -1) raw = raw.slice(fb, lb + 1);

  try {
    return enrichWithCalculations(JSON.parse(raw.replace(/```json|```/g, '').trim()));
  } catch(e) {
    console.error('Parse fail:', e.message);
    return null;
  }
}

// ── Enrich with formula-based calculations ───────────────────────────────────
function enrichWithCalculations(data) {
  if (!data?.elements) return data;
  let totArea=0, totRoad=0, totGSB=0, totWMM=0, totPQC=0, totSteel=0, totCost=0;

  data.elements = data.elements.map(el => {
    const d = el.dimensions || {};
    if (el.type === 'ROAD') {
      const q = calcRoadQuantities(d.length_m||0, d.width_m||0);
      el.quantities = { ...el.quantities, ...q };
      el.cost_inr   = { ...el.cost_inr, ...q.cost_estimate };
      totArea += q.area_sqmt; totRoad += d.length_m||0;
      totGSB += q.gsb_300mm_ton; totWMM += q.wmm_200mm_ton;
      totPQC += q.pqc_250mm_cum; totSteel += q.steel_dowel_kg;
      totCost += q.cost_estimate?.total_sqmt || 0;
    }
    return el;
  });

  data.total_quantities = {
    ...data.total_quantities,
    total_area_sqmt: Math.round(totArea*100)/100,
    total_road_rmt:  Math.round(totRoad*100)/100,
    gsb_total_ton:   Math.round(totGSB*100)/100,
    wmm_total_ton:   Math.round(totWMM*100)/100,
    pqc_total_cum:   Math.round(totPQC*100)/100,
    steel_total_kg:  totSteel,
    calc_note: 'Re-calculated by PMC formula engine',
  };
  const lacs = Math.round(totCost/100000*100)/100;
  data.cost_summary = { ...data.cost_summary, civil_total_inr: totCost, civil_total_lacs: lacs, civil_total_crores: Math.round(lacs/100*100)/100 };
  data.rates_applied = RATES;
  return data;
}

module.exports = { geminiAnalyzeDrawing, runCVAnalysis, calcRoadQuantities, calcStructureQuantities, getBOQForDrawingType, detectDrawingType, DRAWING_TYPES, BOQ_TEMPLATES, RATES };
