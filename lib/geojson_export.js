'use strict';
/**
 * GeoJSON export — boreholes + plan measure geometry hints (QGIS / map view).
 */

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function boreholeFeatures(geotech = {}) {
  const holes = geotech.boreholes || [];
  const withCoords = holes.filter(b => num(b.easting) != null && num(b.northing) != null);
  const features = [];

  if (withCoords.length) {
    for (const b of withCoords) {
      features.push({
        type: 'Feature',
        geometry: { type: 'Point', coordinates: [num(b.easting), num(b.northing)] },
        properties: {
          kind: 'borehole',
          mark: b.mark,
          ground_level_m: b.ground_level_m,
          water_level_m: b.water_level_m,
          avg_spt: b.avg_spt,
          strata: (b.strata || []).join(', '),
          crs_hint: 'projected_local_or_utm',
        },
      });
    }
    return { features, mode: 'geo' };
  }

  // Schematic layout if no easting/northing
  holes.forEach((b, i) => {
    const x = (i % 5) * 50;
    const y = Math.floor(i / 5) * 50;
    features.push({
      type: 'Feature',
      geometry: { type: 'Point', coordinates: [x, y] },
      properties: {
        kind: 'borehole_schematic',
        mark: b.mark,
        ground_level_m: b.ground_level_m,
        water_level_m: b.water_level_m,
        avg_spt: b.avg_spt,
        strata: (b.strata || []).join(', '),
        crs_hint: 'schematic_local',
      },
    });
  });
  return { features, mode: holes.length ? 'schematic' : 'empty' };
}

function measureFeatures(measure = {}) {
  const features = [];
  const geom = measure.geometry || [];
  for (const g of geom) {
    if (g.type === 'Polygon' && g.coordinates) {
      features.push({
        type: 'Feature',
        geometry: { type: 'Polygon', coordinates: g.coordinates },
        properties: { kind: 'measure_area', ...(g.properties || {}) },
      });
    } else if (g.type === 'LineString' && g.coordinates) {
      features.push({
        type: 'Feature',
        geometry: { type: 'LineString', coordinates: g.coordinates },
        properties: { kind: 'measure_length', ...(g.properties || {}) },
      });
    }
  }
  // Fallback: represent printed areas as square centroids on a schematic grid
  if (!features.length && measure.items?.length) {
    measure.items.filter(i => i.type === 'area').slice(0, 12).forEach((item, i) => {
      const side = Math.sqrt(Math.max(item.qty, 1));
      const ox = (i % 4) * (side + 10);
      const oy = Math.floor(i / 4) * (side + 10);
      features.push({
        type: 'Feature',
        geometry: {
          type: 'Polygon',
          coordinates: [[
            [ox, oy], [ox + side, oy], [ox + side, oy + side], [ox, oy + side], [ox, oy],
          ]],
        },
        properties: {
          kind: 'measure_area_schematic',
          qty: item.qty,
          unit: item.unit,
          description: item.description,
          crs_hint: 'schematic_local_m',
        },
      });
    });
  }
  return features;
}

function buildGeoJson({ geotech = null, measure = null, title = 'PMC export' } = {}) {
  const bh = boreholeFeatures(geotech || {});
  const meas = measureFeatures(measure || {});
  const features = [...bh.features, ...meas];
  return {
    type: 'FeatureCollection',
    name: title,
    features,
    meta: {
      borehole_mode: bh.mode,
      feature_count: features.length,
      note: 'PMC GeoJSON v1 — open in QGIS. Schematic coords are not survey control.',
    },
  };
}

module.exports = { buildGeoJson, boreholeFeatures, measureFeatures };
