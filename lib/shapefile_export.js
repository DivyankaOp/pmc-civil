'use strict';
/**
 * Minimal Shapefile export (Point + Polygon) as ZIP — no npm zip deps.
 * Open in QGIS. Companion to GeoJSON.
 */

const { buildGeoJson } = require('./geojson_export');

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return ~c >>> 0;
}

function u16(n) {
  const b = Buffer.alloc(2);
  b.writeUInt16LE(n, 0);
  return b;
}
function u32(n) {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n, 0);
  return b;
}
function i32be(n) {
  const b = Buffer.alloc(4);
  b.writeInt32BE(n, 0);
  return b;
}
function i32le(n) {
  const b = Buffer.alloc(4);
  b.writeInt32LE(n, 0);
  return b;
}
function f64le(n) {
  const b = Buffer.alloc(8);
  b.writeDoubleLE(n, 0);
  return b;
}

/** Store-only ZIP (no compression) */
function zipStore(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8');
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data);
    const crc = crc32(data);
    const local = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(name.length),
      u16(0),
      name,
      data,
    ]);
    const central = Buffer.concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(name.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      name,
    ]);
    localParts.push(local);
    centralParts.push(central);
    offset += local.length;
  }
  const central = Buffer.concat(centralParts);
  const end = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(central.length),
    u32(offset),
    u16(0),
  ]);
  return Buffer.concat([...localParts, central, end]);
}

function dbfField(name, type, size, dec = 0) {
  const buf = Buffer.alloc(32);
  buf.write(name.slice(0, 11), 0, 'ascii');
  buf.write(type, 11, 'ascii');
  buf[16] = size;
  buf[17] = dec;
  return buf;
}

function buildDbf(records, fields) {
  const headerLen = 32 + fields.length * 32 + 1;
  const recLen = 1 + fields.reduce((s, f) => s + f.size, 0);
  const buf = Buffer.alloc(headerLen + records.length * recLen + 1);
  buf[0] = 0x03;
  const now = new Date();
  buf[1] = now.getFullYear() - 1900;
  buf[2] = now.getMonth() + 1;
  buf[3] = now.getDate();
  buf.writeUInt32LE(records.length, 4);
  buf.writeUInt16LE(headerLen, 8);
  buf.writeUInt16LE(recLen, 10);
  let off = 32;
  for (const f of fields) {
    dbfField(f.name, f.type, f.size, f.dec || 0).copy(buf, off);
    off += 32;
  }
  buf[off] = 0x0d;
  off = headerLen;
  for (const rec of records) {
    buf[off++] = 0x20;
    for (const f of fields) {
      let v = rec[f.name] == null ? '' : String(rec[f.name]);
      if (f.type === 'N') {
        v = v.slice(0, f.size).padStart(f.size, ' ');
      } else {
        v = v.slice(0, f.size).padEnd(f.size, ' ');
      }
      buf.write(v, off, f.size, 'ascii');
      off += f.size;
    }
  }
  buf[off] = 0x1a;
  return buf.slice(0, off + 1);
}

function writeShpPoints(points) {
  // points: [{x,y, props}]
  if (!points.length) return null;
  let xmin = Infinity, ymin = Infinity, xmax = -Infinity, ymax = -Infinity;
  for (const p of points) {
    xmin = Math.min(xmin, p.x); ymin = Math.min(ymin, p.y);
    xmax = Math.max(xmax, p.x); ymax = Math.max(ymax, p.y);
  }
  const fileLenWords = 50 + points.length * 14; // 16-bit words
  const shp = Buffer.alloc(100 + points.length * 28);
  // header
  i32be(9994).copy(shp, 0);
  shp.writeInt32BE(fileLenWords, 24);
  shp.writeInt32LE(1000, 28);
  shp.writeInt32LE(1, 32); // Point
  f64le(xmin).copy(shp, 36);
  f64le(ymin).copy(shp, 44);
  f64le(xmax).copy(shp, 52);
  f64le(ymax).copy(shp, 60);
  let o = 100;
  const shx = Buffer.alloc(100 + points.length * 8);
  i32be(9994).copy(shx, 0);
  shx.writeInt32BE(50 + points.length * 4, 24);
  shx.writeInt32LE(1000, 28);
  shx.writeInt32LE(1, 32);
  f64le(xmin).copy(shx, 36);
  f64le(ymin).copy(shx, 44);
  f64le(xmax).copy(shx, 52);
  f64le(ymax).copy(shx, 60);
  let contentWord = 50;
  for (let i = 0; i < points.length; i++) {
    const p = points[i];
    shx.writeInt32BE(contentWord, 100 + i * 8);
    shx.writeInt32BE(10, 100 + i * 8 + 4);
    shp.writeInt32BE(i + 1, o); o += 4;
    shp.writeInt32BE(10, o); o += 4;
    shp.writeInt32LE(1, o); o += 4;
    f64le(p.x).copy(shp, o); o += 8;
    f64le(p.y).copy(shp, o); o += 8;
    contentWord += 14;
  }
  const fields = [
    { name: 'MARK', type: 'C', size: 20 },
    { name: 'KIND', type: 'C', size: 24 },
    { name: 'GL', type: 'N', size: 12, dec: 3 },
    { name: 'SPT', type: 'N', size: 8, dec: 0 },
  ];
  const dbf = buildDbf(points.map(p => ({
    MARK: (p.props.mark || '').slice(0, 20),
    KIND: (p.props.kind || 'point').slice(0, 24),
    GL: p.props.ground_level_m ?? '',
    SPT: p.props.avg_spt ?? '',
  })), fields);
  const prj = 'LOCAL_CS["PMC_Local",UNIT["metre",1.0]]\n';
  return { shp, shx, dbf, prj, count: points.length };
}

function writeShpPolygons(polys) {
  // polys: [{rings: [[[x,y],...]], props}]
  if (!polys.length) return null;
  let xmin = Infinity, ymin = Infinity, xmax = -Infinity, ymax = -Infinity;
  const recBuffers = [];
  for (const poly of polys) {
    const ring = poly.rings[0] || [];
    if (ring.length < 4) continue;
    for (const [x, y] of ring) {
      xmin = Math.min(xmin, x); ymin = Math.min(ymin, y);
      xmax = Math.max(xmax, x); ymax = Math.max(ymax, y);
    }
    const n = ring.length;
    const contentLen = 44 + 4 + n * 16; // bytes after record header, in... wait shapefile content length in 16-bit words
    const contentBytes = 44 + 4 + n * 16;
    const rec = Buffer.alloc(8 + contentBytes);
    rec.writeInt32LE(5, 8); // Polygon
    f64le(xmin).copy(rec, 12); // will fix bbox per poly
    let bx0 = Infinity, by0 = Infinity, bx1 = -Infinity, by1 = -Infinity;
    for (const [x, y] of ring) {
      bx0 = Math.min(bx0, x); by0 = Math.min(by0, y);
      bx1 = Math.max(bx1, x); by1 = Math.max(by1, y);
    }
    f64le(bx0).copy(rec, 12);
    f64le(by0).copy(rec, 20);
    f64le(bx1).copy(rec, 28);
    f64le(by1).copy(rec, 36);
    rec.writeInt32LE(1, 52); // num parts
    rec.writeInt32LE(n, 56); // num points
    rec.writeInt32LE(0, 60); // parts[0]
    let o = 64;
    for (const [x, y] of ring) {
      f64le(x).copy(rec, o); o += 8;
      f64le(y).copy(rec, o); o += 8;
    }
    recBuffers.push({ rec, contentWords: contentBytes / 2, props: poly.props || {} });
  }
  if (!recBuffers.length) return null;
  let fileWords = 50;
  for (const r of recBuffers) fileWords += 4 + r.contentWords;
  const shp = Buffer.alloc(fileWords * 2);
  i32be(9994).copy(shp, 0);
  shp.writeInt32BE(fileWords, 24);
  shp.writeInt32LE(1000, 28);
  shp.writeInt32LE(5, 32);
  f64le(xmin).copy(shp, 36);
  f64le(ymin).copy(shp, 44);
  f64le(xmax).copy(shp, 52);
  f64le(ymax).copy(shp, 60);
  const shx = Buffer.alloc(100 + recBuffers.length * 8);
  i32be(9994).copy(shx, 0);
  shx.writeInt32BE(50 + recBuffers.length * 4, 24);
  shx.writeInt32LE(1000, 28);
  shx.writeInt32LE(5, 32);
  f64le(xmin).copy(shx, 36);
  f64le(ymin).copy(shx, 44);
  f64le(xmax).copy(shx, 52);
  f64le(ymax).copy(shx, 60);
  let o = 100;
  let contentWord = 50;
  for (let i = 0; i < recBuffers.length; i++) {
    const { rec, contentWords, props } = recBuffers[i];
    shx.writeInt32BE(contentWord, 100 + i * 8);
    shx.writeInt32BE(contentWords, 100 + i * 8 + 4);
    shp.writeInt32BE(i + 1, o); o += 4;
    shp.writeInt32BE(contentWords, o); o += 4;
    rec.slice(8).copy(shp, o); o += contentWords * 2;
    contentWord += 4 + contentWords;
    recBuffers[i].props = props;
  }
  const fields = [
    { name: 'KIND', type: 'C', size: 24 },
    { name: 'QTY', type: 'N', size: 12, dec: 3 },
    { name: 'UNIT', type: 'C', size: 8 },
    { name: 'DESC', type: 'C', size: 40 },
  ];
  const dbf = buildDbf(recBuffers.map(r => ({
    KIND: (r.props.kind || 'area').slice(0, 24),
    QTY: r.props.qty ?? r.props.area_sqm ?? '',
    UNIT: (r.props.unit || 'sqm').slice(0, 8),
    DESC: (r.props.description || '').slice(0, 40),
  })), fields);
  const prj = 'LOCAL_CS["PMC_Local",UNIT["metre",1.0]]\n';
  return { shp: shp.slice(0, o), shx, dbf, prj, count: recBuffers.length };
}

function featuresToShapefileZip(geojson) {
  const features = geojson?.features || [];
  const points = [];
  const polys = [];
  for (const f of features) {
    const g = f.geometry;
    const props = f.properties || {};
    if (!g) continue;
    if (g.type === 'Point') {
      points.push({ x: g.coordinates[0], y: g.coordinates[1], props });
    } else if (g.type === 'Polygon') {
      polys.push({ rings: g.coordinates, props });
    }
  }
  const files = [];
  const pt = writeShpPoints(points);
  if (pt) {
    files.push({ name: 'boreholes.shp', data: pt.shp });
    files.push({ name: 'boreholes.shx', data: pt.shx });
    files.push({ name: 'boreholes.dbf', data: pt.dbf });
    files.push({ name: 'boreholes.prj', data: pt.prj });
  }
  const py = writeShpPolygons(polys);
  if (py) {
    files.push({ name: 'areas.shp', data: py.shp });
    files.push({ name: 'areas.shx', data: py.shx });
    files.push({ name: 'areas.dbf', data: py.dbf });
    files.push({ name: 'areas.prj', data: py.prj });
  }
  if (!files.length) {
    return { error: 'No point/polygon features to export', bytes: null };
  }
  files.push({
    name: 'readme.txt',
    data: 'PMC Civil AI shapefile v1. Local/schematic CRS unless survey coords present. Verify in QGIS.\n',
  });
  return {
    bytes: zipStore(files),
    files: files.map(f => f.name),
    points: points.length,
    polygons: polys.length,
  };
}

function buildShapefileZip(opts = {}) {
  const gj = buildGeoJson(opts);
  return { ...featuresToShapefileZip(gj), geojson: gj };
}

module.exports = { buildShapefileZip, featuresToShapefileZip, zipStore };
