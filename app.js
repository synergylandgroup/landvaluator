
// =========================================================
// STORAGE ADAPTER
// ---------------------------------------------------------
// All app persistence flows through these functions.
// Currently backed by localStorage.
//
// SUPABASE MIGRATION: Replace the body of each function
// below with Supabase client calls. The rest of the app
// does not need to change.
//
// Supabase tables needed:
//   zones        (id, user_id, state_abbr, county_name, data jsonb)
//   sheet_configs(id, user_id, state_abbr, county_name, config jsonb)
//   app_state    (id, user_id, state_abbr, county_name)
//
// Example Supabase swap for DB_saveZones:
//   const { error } = await supabase.from('zones').upsert({ user_id: currentUser.id, data: zones });
// =========================================================

const DB = {
  // -- Zones ------------------------------------------
  saveZones(zones) {
    // SUPABASE: await supabase.from('zones').upsert({ user_id, data: zones })
    try { localStorage.setItem('geofence_zones', JSON.stringify(zones)); } catch(e) {}
  },

  loadZones() {
    // SUPABASE: const { data } = await supabase.from('zones').select('data').eq('user_id', user_id).single()
    try { const r = localStorage.getItem('geofence_zones'); return r ? JSON.parse(r) : null; } catch(e) { return null; }
  },

  // -- Sheet Configs -----------------------------------
  saveSheetConfigs(configs) {
    // SUPABASE: await supabase.from('sheet_configs').upsert({ user_id, configs })
    try { localStorage.setItem('geofence_sheet_configs', JSON.stringify(configs)); } catch(e) {}
  },

  loadSheetConfigs() {
    // SUPABASE: const { data } = await supabase.from('sheet_configs').select('configs').eq('user_id', user_id).single()
    try { const r = localStorage.getItem('geofence_sheet_configs'); return r ? JSON.parse(r) : null; } catch(e) { return null; }
  },

  // -- App State ---------------------------------------
  saveAppState(state) {
    // SUPABASE: await supabase.from('app_state').upsert({ user_id, ...state })
    try { localStorage.setItem('geofence_app_state', JSON.stringify(state)); } catch(e) {}
  },

  loadAppState() {
    // SUPABASE: const { data } = await supabase.from('app_state').select('*').eq('user_id', user_id).single()
    try { const r = localStorage.getItem('geofence_app_state'); return r ? JSON.parse(r) : null; } catch(e) { return null; }
  },

  // -- County List Cache -------------------------------
  // (keep in localStorage even after Supabase migration — this is just a UI cache)
  saveCountyCache(abbr, counties) {
    try { localStorage.setItem('counties_'+abbr, JSON.stringify({counties, abbr, ts: Date.now()})); } catch(e) {}
  },

  loadCountyCache(abbr) {
    try { const s = localStorage.getItem('counties_'+abbr); return s ? JSON.parse(s) : null; } catch(e) { return null; }
  },

  clearCountyCache(abbr) {
    try { localStorage.removeItem('counties_'+abbr); } catch(e) {}
  },

  // -- UI State (collapse/expand) ----------------------
  // SUPABASE: await supabase.from('ui_state').upsert({ user_id, key, value })
  saveUIState(key, value) {
    try { localStorage.setItem('czp_ui_'+key, JSON.stringify(value)); } catch(e) {}
  },

  loadUIState(key, fallback = null) {
    try { const r = localStorage.getItem('czp_ui_'+key); return r !== null ? JSON.parse(r) : fallback; } catch(e) { return fallback; }
  },
};

// =========================================================
// STATES DATA
// =========================================================
const STATES = [["Alabama","AL"],["Alaska","AK"],["Arizona","AZ"],["Arkansas","AR"],["California","CA"],["Colorado","CO"],["Connecticut","CT"],["Delaware","DE"],["Florida","FL"],["Georgia","GA"],["Hawaii","HI"],["Idaho","ID"],["Illinois","IL"],["Indiana","IN"],["Iowa","IA"],["Kansas","KS"],["Kentucky","KY"],["Louisiana","LA"],["Maine","ME"],["Maryland","MD"],["Massachusetts","MA"],["Michigan","MI"],["Minnesota","MN"],["Mississippi","MS"],["Missouri","MO"],["Montana","MT"],["Nebraska","NE"],["Nevada","NV"],["New Hampshire","NH"],["New Jersey","NJ"],["New Mexico","NM"],["New York","NY"],["North Carolina","NC"],["North Dakota","ND"],["Ohio","OH"],["Oklahoma","OK"],["Oregon","OR"],["Pennsylvania","PA"],["Rhode Island","RI"],["South Carolina","SC"],["South Dakota","SD"],["Tennessee","TN"],["Texas","TX"],["Utah","UT"],["Vermont","VT"],["Virginia","VA"],["Washington","WA"],["West Virginia","WV"],["Wisconsin","WI"],["Wyoming","WY"]];
const STATE_FIPS = {"AL":"01","AK":"02","AZ":"04","AR":"05","CA":"06","CO":"08","CT":"09","DE":"10","FL":"12","GA":"13","HI":"15","ID":"16","IL":"17","IN":"18","IA":"19","KS":"20","KY":"21","LA":"22","ME":"23","MD":"24","MA":"25","MI":"26","MN":"27","MS":"28","MO":"29","MT":"30","NE":"31","NV":"32","NH":"33","NJ":"34","NM":"35","NY":"36","NC":"37","ND":"38","OH":"39","OK":"40","OR":"41","PA":"42","RI":"44","SC":"45","SD":"46","TN":"47","TX":"48","UT":"49","VT":"50","VA":"51","WA":"53","WV":"54","WI":"55","WY":"56"};

function abbrToFullName(abbr) {
  const s = STATES.find(s => s[1] === abbr);
  return s ? s[0] : abbr;
}

// =========================================================
// CONFIG & STATE
// =========================================================
const SERVICE_ACCOUNT_EMAIL = [99,111,117,110,116,121,122,111,110,101,45,115,104,101,101,116,115,64,99,111,117,110,116,121,122,111,110,101,45,112,114,111,45,52,56,57,48,48,53,46,105,97,109,46,103,115,101,114,118,105,99,101,97,99,99,111,117,110,116,46,99,111,109].map(c=>String.fromCharCode(c)).join('');
// Fill email display via JS to prevent Cloudflare obfuscation
document.addEventListener('DOMContentLoaded', () => { const el = document.getElementById('serviceEmailEl'); if(el) el.textContent = SERVICE_ACCOUNT_EMAIL; });
const MAPBOX_TOKEN = 'pk.eyJ1Ijoic3luZXJneWxhbmRncm91cCIsImEiOiJjbW02MjI5dTEwY2xtMnFuMGs2Y3Y2OWlwIn0.O7gX97oTNFUw9HooOheq6w';
mapboxgl.accessToken = MAPBOX_TOKEN;

// 8 distinct zone colors
const COLORS = ['#5b7fa6','#e07b39','#3dab6a','#dba63a','#9b6bc7','#d96a8a','#38b4c4','#7ab03c'];
let selectedColor = COLORS[0];

// polygons: { id, name, letter, stateAbbr, countyName, color, points, description, labelMarker, handles[], _isRect, _bounds }
let polygons = [];
let properties = [];
// Per-county sheet configs — keyed by "stateAbbr|countyName"
let sheetConfigs = {};
let sheetConfig  = {};  // active config for currently selected county

function _countyKey(stateAbbr, countyName) {
  return (stateAbbr || '').trim() + '|' + (countyName || '').trim();
}

function _getSheetConfig(stateAbbr, countyName) {
  return sheetConfigs[_countyKey(stateAbbr, countyName)] || null;
}

function _setSheetConfig(stateAbbr, countyName, cfg) {
  sheetConfigs[_countyKey(stateAbbr, countyName)] = cfg;
  DB.saveSheetConfigs(sheetConfigs);
}
let countySourceId = null;
const _countyLayers = {}; // key -> sourceId for all counties with zones
let _pendingCountyGeoJSON = null;
let _countyGeoJSONCache = {};

// Draw state
let drawMode = null;
let polyState = 'idle';
let drawPoints = [];

// Zone desc modal
let _editingDescId = null;

// Mapbox draw layer IDs
const SRC_FILL    = '__draw_fill';
const SRC_LINE    = '__draw_line';
const SRC_PREVIEW = '__draw_preview';
const SRC_VERTS   = '__draw_verts';

// =========================================================
// BUILD COLOR SWATCHES — runs immediately on parse
// =========================================================
(function buildSwatches() {
  const row = document.getElementById('colorRow');
  COLORS.forEach((c, i) => {
    const s = document.createElement('div');
    s.className = 'color-swatch' + (i === 0 ? ' sel' : '');
    s.style.background = c;
    s.title = c;
    s.onclick = () => {
      document.querySelectorAll('.color-swatch').forEach(el => el.classList.remove('sel'));
      s.classList.add('sel');
      selectedColor = c;
    };
    row.appendChild(s);
  });
})();

// =========================================================
// MAP STYLES
// =========================================================
const MAP_STYLES = [
  { id: 'mapbox://styles/mapbox/streets-v12',           label: '🗺 Streets'  },
  { id: 'mapbox://styles/mapbox/outdoors-v12',          label: '🌲 Outdoors' },
  { id: 'mapbox://styles/mapbox/satellite-v9',          label: '🛰 Satellite'},
  { id: 'mapbox://styles/mapbox/satellite-streets-v12', label: '🛰 Hybrid'   },
  { id: 'mapbox://styles/mapbox/light-v11',             label: '☁️ Light'    },
  { id: 'mapbox://styles/mapbox/dark-v11',              label: '🌙 Dark'     },
];
const HYBRID_IDX = 3;
(function buildStyleSelect() {
  const el = document.getElementById('styleSelect');
  MAP_STYLES.forEach((s, i) => {
    const o = document.createElement('option');
    o.value = i; o.textContent = s.label;
    if (i === HYBRID_IDX) o.selected = true;
    el.appendChild(o);
  });
})();

// =========================================================
// MAP INIT
// =========================================================
const map = new mapboxgl.Map({
  container: 'map',
  style: MAP_STYLES[HYBRID_IDX].id,
  center: [-98.35, 39.5],
  zoom: 4,
  projection: 'mercator',
  doubleClickZoom: false,
  boxZoom: false,
});
map.addControl(new mapboxgl.NavigationControl(), 'bottom-right');
map.addControl(new mapboxgl.ScaleControl({ maxWidth: 120, unit: 'imperial' }), 'bottom-left');

function resetNorth() { map.easeTo({ bearing: 0, pitch: 0, duration: 500 }); }
map.on('rotate', () => {
  document.getElementById('northBtn').style.opacity = Math.abs(map.getBearing()) > 1 ? '1' : '0.6';
});

function changeMapStyle(idx) {
  polygons.forEach(p => _removeZoneLabel(p));
  map.setStyle(MAP_STYLES[parseInt(idx)].id);
}

map.on('style.load', () => {
  _initDrawLayers();
  _restoreAllZoneLayers();
  polygons.forEach(p => _addZoneLabel(p));
  if (_pendingCountyGeoJSON) _readdCountyLayer(_pendingCountyGeoJSON);
  _rebuildAllLabels();
});

// Re-evaluate clusters on zoom/pan
map.on('zoomend', () => { if (polygons.length) { _refreshLabelMode(); } });
map.on('moveend', () => { if (polygons.length) { _refreshLabelMode(); } });

// =========================================================
// DRAW LAYERS
// =========================================================
function _initDrawLayers() {
  function addSrc(id, type) {
    if (map.getSource(id)) return;
    const empty = type === 'Polygon'
      ? { type:'Feature', properties:{}, geometry:{ type:'Polygon', coordinates:[[]] } }
      : type === 'LineString'
      ? { type:'Feature', properties:{}, geometry:{ type:'LineString', coordinates:[] } }
      : { type:'FeatureCollection', features:[] };
    map.addSource(id, { type:'geojson', data: empty });
  }
  addSrc(SRC_FILL, 'Polygon');
  addSrc(SRC_LINE, 'LineString');
  addSrc(SRC_PREVIEW, 'LineString');
  addSrc(SRC_VERTS, 'FeatureCollection');

  if (!map.getLayer(SRC_FILL))
    map.addLayer({ id:SRC_FILL, type:'fill', source:SRC_FILL, paint:{'fill-color':['get','color'],'fill-opacity':0.12} });
  if (!map.getLayer(SRC_LINE))
    map.addLayer({ id:SRC_LINE, type:'line', source:SRC_LINE, paint:{'line-color':['get','color'],'line-width':2.5,'line-dasharray':[2,2]} });
  if (!map.getLayer(SRC_PREVIEW))
    map.addLayer({ id:SRC_PREVIEW, type:'line', source:SRC_PREVIEW, paint:{'line-color':['get','color'],'line-width':1.5,'line-dasharray':[2,3]} });
  if (!map.getLayer(SRC_VERTS))
    map.addLayer({ id:SRC_VERTS, type:'circle', source:SRC_VERTS, paint:{'circle-radius':4,'circle-color':['get','color'],'circle-stroke-color':'#fff','circle-stroke-width':1.5} });
}

function _setDrawSrc(id, data) { if (map.getSource(id)) map.getSource(id).setData(data); }
function _emptyPoly()   { return { type:'Feature', properties:{}, geometry:{ type:'Polygon', coordinates:[[]] } }; }
function _emptyLine()   { return { type:'Feature', properties:{}, geometry:{ type:'LineString', coordinates:[] } }; }
function _emptyPts()    { return { type:'FeatureCollection', features:[] }; }
function _clearPreviews() {
  _setDrawSrc(SRC_FILL, _emptyPoly());
  _setDrawSrc(SRC_LINE, _emptyLine());
  _setDrawSrc(SRC_PREVIEW, _emptyLine());
  _setDrawSrc(SRC_VERTS, _emptyPts());
}

// =========================================================
// ZONE LAYER MANAGEMENT
// =========================================================
const _srcId  = id => 'zone-src-'  + id;
const _fillId = id => 'zone-fill-' + id;
const _lineId = id => 'zone-line-' + id;

function _addZoneLayers(poly) {
  const { id, color, points } = poly;
  const gj = { type:'Feature', properties:{ color }, geometry:{ type:'Polygon', coordinates:[[...points, points[0]]] } };
  if (!map.getSource(_srcId(id))) map.addSource(_srcId(id), { type:'geojson', data: gj });
  else map.getSource(_srcId(id)).setData(gj);
  if (!map.getLayer(_fillId(id)))
    map.addLayer({ id:_fillId(id), type:'fill', source:_srcId(id), paint:{'fill-color':color,'fill-opacity':0.2} });
  if (!map.getLayer(_lineId(id)))
    map.addLayer({ id:_lineId(id), type:'line', source:_srcId(id), paint:{'line-color':color,'line-width':2} });
  // Polygon fill click intentionally disabled — use zone label to open notes
  map.on('mouseenter', _fillId(id), () => { if (!drawMode) map.getCanvas().style.cursor = 'grab'; });
  map.on('mouseleave', _fillId(id), () => { if (!drawMode) map.getCanvas().style.cursor = ''; });
}
function _removeZoneLayers(id) {
  if (map.getLayer(_fillId(id))) map.removeLayer(_fillId(id));
  if (map.getLayer(_lineId(id))) map.removeLayer(_lineId(id));
  if (map.getSource(_srcId(id))) map.removeSource(_srcId(id));
}
function _restoreAllZoneLayers() { polygons.forEach(p => _addZoneLayers(p)); }

// =========================================================
// ZONE LABELS
// Hierarchy: zoomed-out → county pill → zoom in → zone labels → click → notes
// =========================================================

// Zoom threshold: below this = show county pills, at/above = show zone labels
const COUNTY_PILL_ZOOM = 9;

function _polyCenter(poly) {
  // True centroid using polygon area formula (handles non-convex shapes)
  const pts = poly.points;
  let cx = 0, cy = 0, area = 0;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const xi = pts[i][0], yi = pts[i][1];
    const xj = pts[j][0], yj = pts[j][1];
    const cross = xi * yj - xj * yi;
    area += cross;
    cx += (xi + xj) * cross;
    cy += (yi + yj) * cross;
  }
  area /= 2;
  if (Math.abs(area) < 1e-10) {
    // Fallback to bounding box center for degenerate polygons
    const lngs = pts.map(p=>p[0]), lats = pts.map(p=>p[1]);
    return [(Math.min(...lngs)+Math.max(...lngs))/2, (Math.min(...lats)+Math.max(...lats))/2];
  }
  cx /= (6 * area);
  cy /= (6 * area);
  return [cx, cy];
}

// County pill markers — keyed by "stateAbbr|countyName"
let _countyPillMarkers = {};

// Add/refresh a single zone's individual label (hidden until zoomed in)
function _addZoneLabel(poly) {
  _removeZoneLabel(poly);
  const el = document.createElement('div');
  el.className = 'zone-label';
  el.innerHTML = `<span class="zl-letter" style="color:var(--zone-blue,#2c5282)">ZONE ${poly.letter||''}</span><span class="zl-name">${poly.name||''}</span>`;
  el.setAttribute('data-tip', 'Open pricing panel');
  // Single click on zone label = open Notes & Pricing
  // Guard: do nothing if a county button was just clicked
  el.addEventListener('click', (e) => {
    e.stopPropagation();
    if (window._countyClickActive) return;
    openZoneDescModal(poly.id);
  });
  const center = _polyCenter(poly);
  el.style.borderColor = poly.color + '66';
  // Start hidden — _refreshLabelMode will show/hide based on zoom
  el.style.display = 'none';
  poly.labelMarker = new mapboxgl.Marker({ element:el, anchor:'center' }).setLngLat(center).addTo(map);
  // Do NOT call _refreshLabelMode here — caller is responsible for calling _rebuildAllLabels
  // after all zones are added to avoid thrashing
}

function _removeZoneLabel(poly) {
  if (poly.labelMarker) { poly.labelMarker.remove(); poly.labelMarker = null; }
}

// Build/rebuild all county pills
function _buildCountyPills() {
  // Remove old pills
  Object.values(_countyPillMarkers).forEach(m => m.remove());
  _countyPillMarkers = {};

  // Group by county
  const groups = {};
  polygons.forEach(p => {
    const key = (p.stateAbbr||'?') + '|' + (p.countyName||'?');
    if (!groups[key]) groups[key] = [];
    groups[key].push(p);
  });

  Object.entries(groups).forEach(([key, polys]) => {
    // Centroid of all zone centers in county
    const centers = polys.map(_polyCenter);
    const lng = centers.reduce((s,c) => s+c[0], 0) / centers.length;
    const lat = centers.reduce((s,c) => s+c[1], 0) / centers.length;

    const county = polys[0].countyName || 'County';
    const st = polys[0].stateAbbr || '';
    const count = polys.length;

    const el = document.createElement('div');
    el.className = 'zone-cluster';
    el.innerHTML = `${county} County, ${st}&nbsp;<span class="zc-count">${count}</span>`;
    el.title = `Click to zoom into ${county} County`;

    // Single click on county pill = zoom into county only, never open notes
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      window._countyClickActive = true;
      setTimeout(() => { window._countyClickActive = false; }, 400);
      const b = new mapboxgl.LngLatBounds();
      polys.forEach(p => p.points.forEach(pt => b.extend(pt)));
      map.fitBounds(b, { padding: 100 });
      // Load county boundary after map finishes animating
      const sa = polys[0].stateAbbr, cn = polys[0].countyName;
      map.once('moveend', () => loadCountyBoundaryOnly(sa, cn));
    });

    const marker = new mapboxgl.Marker({ element:el, anchor:'center' })
      .setLngLat([lng, lat]).addTo(map);
    _countyPillMarkers[key] = marker;
  });
}

// Master switch: called on every zoomend/moveend and after label changes
function _refreshLabelMode() {
  const zoom = map.getZoom();
  const zoomed = zoom >= COUNTY_PILL_ZOOM;

  // County pills: visible when zoomed out
  Object.values(_countyPillMarkers).forEach(m => {
    m.getElement().style.display = zoomed ? 'none' : '';
  });

  // Zone labels: visible when zoomed in
  polygons.forEach(p => {
    if (p.labelMarker) p.labelMarker.getElement().style.display = zoomed ? '' : 'none';
  });
}

// Full rebuild — call after any polygon add/remove
function _rebuildAllLabels() {
  _buildCountyPills();
  _refreshLabelMode();
}

// =========================================================
// DRAWING
// =========================================================
function pixelDist(a, b) {
  const pa = map.project(a), pb = map.project(b);
  return Math.hypot(pa.x-pb.x, pa.y-pb.y);
}

function cancelDraw() {
  drawMode = null; polyState = 'idle'; drawPoints = [];
  _clearPreviews();
  map.dragPan.enable(); map.boxZoom.enable();
  map.getCanvas().style.cursor = '';
  document.getElementById('btnPolygon').classList.remove('active');
  document.getElementById('btnCancel').style.display = 'none';
  document.getElementById('drawHint').style.display = 'none';
  document.getElementById('drawHelp').textContent = 'Select a state & county above, then click Draw Polygon.';
}

function startDraw() {
  cancelDraw();
  drawMode = 'polygon'; polyState = 'drawing';
  map.getCanvas().style.cursor = 'crosshair';
  document.getElementById('btnPolygon').classList.add('active');
  document.getElementById('btnCancel').style.display = 'block';
  const hint = document.getElementById('drawHint');
  hint.style.display = 'block';
  hint.textContent = '📍 Click to add vertices — click first point to close';
  document.getElementById('drawHelp').textContent = 'Click to place vertices. Click the first point to close.';
}

map.on('click', function(e) {
  if (window.innerWidth <= 700) document.getElementById('sidebar').classList.remove('open');
  if (drawMode !== 'polygon' || polyState !== 'drawing') return;
  const pt = [e.lngLat.lng, e.lngLat.lat];
  if (drawPoints.length >= 3 && pixelDist(pt, drawPoints[0]) <= 10) { _finishPolygon(); return; }
  drawPoints.push(pt);
  _refreshPolyPreviews(null);
});

map.on('mousemove', function(e) {
  if (drawMode !== 'polygon' || polyState !== 'drawing' || !drawPoints.length) return;
  const cursor = [e.lngLat.lng, e.lngLat.lat];
  const nearClose = drawPoints.length >= 3 && pixelDist(cursor, drawPoints[0]) <= 10;
  map.getCanvas().style.cursor = nearClose ? 'pointer' : 'crosshair';
  _refreshPolyPreviews(cursor, nearClose);
});

function _refreshPolyPreviews(cur, nearClose) {
  const color = selectedColor;
  _setDrawSrc(SRC_LINE, drawPoints.length >= 2
    ? { type:'Feature', properties:{color}, geometry:{ type:'LineString', coordinates:drawPoints } }
    : _emptyLine());
  _setDrawSrc(SRC_FILL, drawPoints.length >= 3
    ? { type:'Feature', properties:{color}, geometry:{ type:'Polygon', coordinates:[[...drawPoints, drawPoints[0]]] } }
    : _emptyPoly());
  if (cur && drawPoints.length >= 1) {
    const pc = nearClose ? '#5b7fa6' : color;
    _setDrawSrc(SRC_PREVIEW, { type:'Feature', properties:{color:pc}, geometry:{ type:'LineString', coordinates:[drawPoints[drawPoints.length-1], cur] } });
    if (map.getLayer(SRC_PREVIEW)) {
      try { map.setPaintProperty(SRC_PREVIEW, 'line-color', pc); } catch(e) {}
      map.setPaintProperty(SRC_PREVIEW, 'line-width', nearClose ? 2.5 : 1.5);
      map.setPaintProperty(SRC_PREVIEW, 'line-dasharray', nearClose ? [1] : [2,3]);
    }
  } else { _setDrawSrc(SRC_PREVIEW, _emptyLine()); }
  _setDrawSrc(SRC_VERTS, {
    type:'FeatureCollection',
    features: drawPoints.map((p, i) => ({ type:'Feature', properties:{color: i===0 ? '#5b7fa6' : color}, geometry:{ type:'Point', coordinates:p } }))
  });
}

function _finishPolygon() {
  if (drawPoints.length < 3) { showToast('Need at least 3 points', 'error'); return; }
  const pts = drawPoints.slice();
  const color = selectedColor;

  // Validate polygon is within selected county boundary
  if (_pendingCountyGeoJSON) {
    // Check if centroid of drawn polygon is within county
    const avgLng = pts.reduce((s,p) => s+p[0], 0) / pts.length;
    const avgLat = pts.reduce((s,p) => s+p[1], 0) / pts.length;
    const inCounty = _pendingCountyGeoJSON.features.some(f => {
      const coords = f.geometry.type === 'Polygon' ? [f.geometry.coordinates]
                   : f.geometry.type === 'MultiPolygon' ? f.geometry.coordinates : [];
      return coords.some(poly => pointInPolygon(avgLat, avgLng, poly[0].map(c => [c[0], c[1]])));
    });
    if (!inCounty) {
      showToast('Zone must be drawn within the selected county boundary', 'error');
      cancelDraw();
      return;
    }
  }

  cancelDraw();
  createPolygonAuto(pts, color);
}

// =========================================================
// LETTER MANAGEMENT — per county
// =========================================================
function _nextLetterForCounty(stateAbbr, countyName) {
  const used = new Set(
    polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName === countyName).map(p => p.letter).filter(Boolean)
  );
  for (let i = 0; i < 26; i++) {
    const l = String.fromCharCode(65 + i);
    if (!used.has(l)) return l;
  }
  return null;
}

// =========================================================
// CREATE POLYGON (auto-named, no modal)
// =========================================================
function createPolygonAuto(pts, color) {
  const stateAbbr  = document.getElementById('stateSelect').value || '';
  const countyName = document.getElementById('countySelect').value || '';
  if (!stateAbbr || !countyName) {
    showToast('Please select a State and County before drawing a zone.', 'error');
    return;
  }
  const letter = _nextLetterForCounty(stateAbbr, countyName);
  if (!letter) { showToast('Max 26 zones per county (A–Z). Delete a zone first.', 'error'); return; }
  const name = `${countyName} County, ${stateAbbr}`;
  const poly = {
    id: 'poly_' + Date.now(), name, letter, stateAbbr, countyName,
    color, points: pts, description: '', labelMarker: null, handles: [],
    _isRect: false, _bounds: null,
  };
  polygons.push(poly);
  _addZoneLayers(poly);
  _addZoneLabel(poly);
  renderPolygonList();
  _rebuildAllLabels();
  persistZones();
  showToast(`Zone ${letter} created — ${name}`, 'success');
}

// =========================================================
// ZONE PRICING EDITOR
// =========================================================
function openZoneDescModal(polyId) { openZoneEditor(polyId); } // alias for existing callers

function openZoneEditor(polyId) {
  const p = polygons.find(p => p.id === polyId);
  if (!p) return;
  _editingDescId = polyId;

  // Header
  document.getElementById('zeBadge').textContent = `ZONE ${p.letter}`;
  document.getElementById('zeTitle').textContent = p.name;
  // zeAllZones checkbox removed from UI


  // Notes
  document.getElementById('zeNotes').value = p.description || '';

  // Pricing rows
  const rows = (p.pricingTiers && p.pricingTiers.length)
    ? p.pricingTiers
    : [{ minAcres: '', maxAcres: '', pricePerAcre: '' }];
  zeRenderRows(rows, p.letter);

  document.getElementById('zoneEditorModal').classList.add('open');
  setTimeout(() => document.getElementById('zeNotes').focus(), 80);
}

function closeZoneEditor() {
  document.getElementById('zoneEditorModal').classList.remove('open');
  _editingDescId = null;
}


// -- Fetch live spreadsheet name from Google Sheets API --
async function _fetchSheetName(sheetId) {
  // Only updates the modal's read-only sheet name field — never the sidebar
  try {
    const r = await fetch('/.netlify/functions/sheets-read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetId, sheetName: 'LI Raw Dataset', metaOnly: true }),
    });
    const data = await r.json();
    const name = data.spreadsheetTitle || data.sheetTitle || '';
    const el = document.getElementById('smSheetNameField');
    if (el) { el.value = name; }
  } catch(e) {}
}

// -- Open sheets modal pre-set to a specific county --
function openSheetsModalForCounty(stateAbbr, countyName, e) {
  if (e) e.stopPropagation();
  // Switch to this county first
  const ss = document.getElementById('stateSelect');
  const cs = document.getElementById('countySelect');
  if (ss.value !== stateAbbr) {
    ss.value = stateAbbr;
    loadCounties(true).then(() => {
      cs.value = countyName;
      openSheetsModal();
    });
  } else {
    cs.value = countyName;
    openSheetsModal();
  }
}

// -- Share a single county via short URL --
async function shareCounty(stateAbbr, countyName, e) {
  if (e) e.stopPropagation();
  const countyPolys = polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName === countyName);
  if (!countyPolys.length) { showToast('No zones to share for this county', 'error'); return; }
  showToast('Generating share link...', 'info');
  try {
    const payload = { version: 2, stateAbbr, countyName, zones: countyPolys.map(_polyToJSON) };
    const r = await fetch('/.netlify/functions/share-save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await r.json();
    if (!data.id) throw new Error(data.error || 'No ID returned');

    let url;
    if (data.fallback) {
      // Blob storage not configured — fall back to encoded URL
      const encoded = btoa(encodeURIComponent(JSON.stringify(data.data)));
      url = `${window.location.origin}${window.location.pathname}?zones=${encoded}`;
    } else {
      url = `${window.location.origin}${window.location.pathname}?share=${data.id}`;
    }

    navigator.clipboard.writeText(url)
      .then(() => showToast('Share link copied!', 'success'))
      .catch(() => {
        const b = document.getElementById('shareBanner');
        b.textContent = '🔗 ' + url; b.style.display = 'block';
        setTimeout(() => b.style.display = 'none', 12000);
      });
  } catch(err) { showToast('Share failed: ' + err.message, 'error'); }
}

// -- Load zones from short share URL --
async function loadZonesFromShareId(shareId) {
  try {
    const r = await fetch('/.netlify/functions/share-load?id=' + shareId);
    const data = await r.json();
    if (!data.zones) throw new Error('No zones in share');
    data.zones.forEach(d => _loadZone(d));
    renderPolygonList(); persistZones(); _rebuildAllLabels();
    if (polygons.length) {
      const b = new mapboxgl.LngLatBounds();
      polygons.forEach(p => p.points.forEach(pt => b.extend(pt)));
      map.fitBounds(b, { padding:60 });
      showToast(`Loaded shared zones for ${data.countyName} County, ${data.stateAbbr}`, 'success');
    }
    return true;
  } catch(e) { showToast('Could not load shared zones', 'error'); return false; }
}

// -- Save & Sync: save pricing + assign + write zones + sync pricing --
async function saveAndSyncZone() {
  if (!_editingDescId) return;
  const p = polygons.find(p => p.id === _editingDescId);
  if (!p) return;

  const btn = document.getElementById('zeSaveBtn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Saving...'; }

  // 1. Save locally — capture state BEFORE closing modal
  p.description = document.getElementById('zeNotes').value.trim();
  p.pricingTiers = zeCollectRows();
  p.allZones = false;
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  persistZones();

  const cfg = (sa && cn) ? (_getSheetConfig(sa, cn) || sheetConfig) : sheetConfig;

  if (!cfg || !cfg.sheetId) {
    closeZoneEditor();
    showToast(`Zone ${p.letter} saved locally (no sheet connected)`, 'success');
    if (btn) { btn.disabled = false; btn.innerHTML = '💾 Save &amp; Sync'; }
    return;
  }

  // If properties aren't loaded yet, fetch them first then proceed
  if (!properties || !properties.length) {
    showToast('Loading properties from sheet...', 'info');
    try {
      const pr = await fetch('/.netlify/functions/sheets-read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId: cfg.sheetId, sheetName: cfg.sheetName || 'LI Raw Dataset' }),
      });
      const pd = await pr.json();
      if (pd.properties && pd.properties.length) {
        loadPropertiesFromFunction(pd.properties);
        document.getElementById('statProperties').textContent = pd.properties.length;
      }
    } catch(e) { console.warn('Could not prefetch properties:', e); }
  }

  // Now safe to close the modal — state/county captured above
  closeZoneEditor();
  showToast(`Zone ${p.letter} saved — syncing...`, 'info');

  try {
    // 2. Run zone assignment
    const assignments = [];
    let assigned = 0;
    properties.forEach(prop => {
      prop.zone = null;
      for (const poly of polygons) {
        if (pointInPolygon(prop.lat, prop.lng, poly.points)) {
          prop.zone = poly.letter;
          poly.propCount = (poly.propCount || 0) + 1;
          assigned++;
          if (prop.rowIndex) assignments.push({ rowIndex: prop.rowIndex, zone: poly.letter });
          break;
        }
      }
    });
    document.getElementById('statAssigned').textContent = assigned;
    renderPolygonList();

    // 3. Write zone letters to sheet
    if (assignments.length) {
      const zr = await fetch('/.netlify/functions/sheets-write-zones', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId: cfg.sheetId, sheetName: cfg.sheetName || 'LI Raw Dataset', assignments }),
      });
      const zd = await zr.json();
      if (!zr.ok) throw new Error(zd.error || 'Zone write failed');
    }

    // 4. Sync all pricing tiers
    const countyPolys = polygons.filter(poly => !poly.stateAbbr || (poly.stateAbbr === sa && poly.countyName === cn));
    const allTiers = [];
    countyPolys.slice().sort((a,b) => (a.letter||'').localeCompare(b.letter||'')).forEach(poly => {
      (poly.pricingTiers || [])
        .filter(t => t.pricePerAcre !== '' && t.pricePerAcre !== undefined && t.pricePerAcre !== null)
        .sort((a,b) => parseFloat(a.minAcres||0) - parseFloat(b.minAcres||0))
        .forEach(t => allTiers.push({ zone: poly.letter, minAcres: t.minAcres, maxAcres: t.maxAcres, pricePerAcre: t.pricePerAcre }));
    });

    if (allTiers.length) {
      const pr = await fetch('/.netlify/functions/sheets-write-pricing', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId: cfg.sheetId, tiers: allTiers }),
      });
      const pd = await pr.json();
      if (!pd.success) throw new Error(pd.error || 'Pricing sync failed');
    }

    showToast(`Zone ${p.letter} saved — ${assigned} assigned, pricing synced ✓`, 'success');
  } catch(err) {
    showToast('Sync error: ' + err.message, 'error');
  }
  if (btn) { btn.disabled = false; btn.innerHTML = '💾 Save &amp; Sync'; }
}

function saveZoneEditor() {
  if (!_editingDescId) return;
  const p = polygons.find(p => p.id === _editingDescId);
  if (!p) return;

  p.description = document.getElementById('zeNotes').value.trim();
  p.pricingTiers = zeCollectRows();
  p.allZones = document.getElementById('zeAllZones').checked;

  persistZones();
  closeZoneEditor();
  const label = p.allZones ? `Zone ${p.letter} saved (ALL zones pricing)` : `Zone ${p.letter} saved`;
  showToast(label, 'success');

}

// -- Sync ALL zones pricing to sheet in one batch --
async function syncAllPricingToSheet() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  // Try current county first, then fall back to any connected config
  let cfg = (sa && cn) ? (_getSheetConfig(sa, cn) || sheetConfig) : sheetConfig;
  // If still no config, try to find one from the polygons being synced
  if (!cfg || !cfg.sheetId) {
    const anyPoly = polygons.find(p => p.stateAbbr && p.countyName && _getSheetConfig(p.stateAbbr, p.countyName));
    if (anyPoly) cfg = _getSheetConfig(anyPoly.stateAbbr, anyPoly.countyName);
  }
  if (!cfg || !cfg.sheetId) { showToast('Connect a Google Sheet first', 'error'); return; }

  // Collect tiers from ALL polygons for this county
  // Filter to only polygons matching current state+county
  const countyPolys = polygons.filter(p => {
    return !p.stateAbbr || (p.stateAbbr === sa && p.countyName === cn);
  });

  if (!countyPolys.length) { showToast('No zones with pricing to sync', 'error'); return; }

  const allTiers = [];
  countyPolys
    .slice() // don't mutate
    .sort((a, b) => (a.letter || '').localeCompare(b.letter || '')) // A→Z
    .forEach(poly => {
      const zoneLabel = poly.allZones ? 'ALL' : poly.letter;
      (poly.pricingTiers || [])
        .filter(t => t.pricePerAcre !== '' && t.pricePerAcre !== undefined && t.pricePerAcre !== null)
        .sort((a, b) => parseFloat(a.minAcres || 0) - parseFloat(b.minAcres || 0)) // low→high acreage
        .forEach(t => {
          allTiers.push({
            zone: zoneLabel,
            minAcres: t.minAcres,
            maxAcres: t.maxAcres,
            pricePerAcre: t.pricePerAcre,
          });
        });
    });

  if (!allTiers.length) { showToast('No pricing tiers to sync — add pricing to your zones first', 'error'); return; }

  showToast('Syncing all pricing to sheet...', 'info');
  try {
    const r = await fetch('/.netlify/functions/sheets-write-pricing', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetId: cfg.sheetId, tiers: allTiers }),
    });
    const data = await r.json();
    if (data.success) showToast(`${allTiers.length} pricing rows synced to sheet ✓`, 'success');
    else showToast('Sync failed: ' + data.error, 'error');
  } catch(e) { showToast('Sync failed: ' + e.message, 'error'); }
}

// -- Table helpers --------------------------------------
function zeRenderRows(rows, defaultLetter) {
  const tbody = document.getElementById('zeTbody');
  tbody.innerHTML = '';
  rows.forEach((r, i) => zeAppendRow(tbody, r, defaultLetter, i));
}

function zeAppendRow(tbody, r, defaultLetter, idx) {
  const tr = document.createElement('tr');
  tr.dataset.idx = idx;
  tr.innerHTML = `
    <td><input class="ze-cell num-col" type="number" min="0" step="any" value="${r.minAcres}" placeholder="0" data-col="minAcres"></td>
    <td><input class="ze-cell num-col" type="number" min="0" step="any" value="${r.maxAcres}" placeholder="∞" data-col="maxAcres"></td>
    <td><div style="display:flex;align-items:center;background:transparent;border:1px solid transparent;border-radius:4px;transition:border-color 0.12s" class="ze-price-wrap"><span style="color:var(--muted);font-size:12px;padding-left:7px;flex-shrink:0">$</span><input class="ze-cell price-col" type="number" min="0" step="0.01" value="${r.pricePerAcre}" placeholder="0.00" data-col="pricePerAcre" style="border:none;padding-left:3px;flex:1"></div></td>
    <td><button class="ze-del-row" onclick="zeDeleteRow(this)" title="Remove row">✕</button></td>
  `;
  tbody.appendChild(tr);
}

function zeAddRow() {
  const p = polygons.find(p => p.id === _editingDescId);
  const defaultLetter = p ? p.letter : 'A';
  const tbody = document.getElementById('zeTbody');
  const newIdx = tbody.rows.length;
  zeAppendRow(tbody, { minAcres: '', maxAcres: '', pricePerAcre: '' }, defaultLetter, newIdx);
  // Focus the min acres cell of new row
  tbody.rows[tbody.rows.length - 1].querySelector('[data-col="minAcres"]').focus();
}

function zeDeleteRow(btn) {
  const tr = btn.closest('tr');
  tr.remove();
}

function zeCollectRows() {
  const rows = [];
  document.querySelectorAll('#zeTbody tr').forEach(tr => {
    const get = col => tr.querySelector(`[data-col="${col}"]`)?.value?.trim() || '';
    rows.push({ minAcres: get('minAcres'), maxAcres: get('maxAcres'), pricePerAcre: get('pricePerAcre') });
  });
  return rows;
}

document.getElementById('zoneEditorModal').addEventListener('click', e => { if (e.target === e.currentTarget) closeZoneEditor(); });

// =========================================================
// RENDER POLYGON LIST — Full State Name → County hierarchy
// =========================================================
function renderPolygonList() {
  document.getElementById('zoneCount').textContent = polygons.length;
  document.getElementById('statPolygons').textContent = polygons.length;
  const stateSet = new Set(polygons.map(p => p.stateAbbr).filter(Boolean));
  document.getElementById('statStates').textContent = stateSet.size;

  const list = document.getElementById('polygonsList');
  if (!polygons.length) {
    list.innerHTML = '<div class="empty-state">No zones yet.<br>Select a state &amp; county,<br>then draw on the map.</div>';
    return;
  }

  // Group by stateAbbr → countyName
  const byState = {};
  polygons.forEach(p => {
    const sa = p.stateAbbr || 'Unknown';
    const cn = p.countyName || 'Unknown';
    if (!byState[sa]) byState[sa] = {};
    if (!byState[sa][cn]) byState[sa][cn] = [];
    byState[sa][cn].push(p);
  });

  list.innerHTML = '';
  Object.keys(byState).sort().forEach(stateAbbr => {
    const fullName = abbrToFullName(stateAbbr);
    const totalZones = Object.values(byState[stateAbbr]).reduce((a,b) => a+b.length, 0);

    const stateDiv = document.createElement('div');
    stateDiv.className = 'state-group';

    const stateOpenKey = 'state_open_' + stateAbbr;
    const isStateOpen = DB.loadUIState(stateOpenKey, true); // default open

    const hdr = document.createElement('div');
    hdr.className = 'state-header' + (isStateOpen ? ' open' : '');
    hdr.innerHTML = `<span class="sg-arrow">▶</span><span class="sg-name">${fullName}</span><span class="sg-count">${totalZones}</span>`;
    hdr.onclick = () => {
      hdr.classList.toggle('open');
      countiesDiv.classList.toggle('open');
      DB.saveUIState(stateOpenKey, hdr.classList.contains('open'));
    };

    const countiesDiv = document.createElement('div');
    countiesDiv.className = 'state-counties' + (isStateOpen ? ' open' : '');

    Object.keys(byState[stateAbbr]).sort().forEach(countyName => {
      const cPolys = byState[stateAbbr][countyName];
      const cGroup = document.createElement('div');
      cGroup.className = 'county-group';

      const cCfg = _getSheetConfig(stateAbbr, countyName);
      const isConnected = !!(cCfg && cCfg.sheetId);

      const cHdr = document.createElement('div');
      cHdr.className = 'county-header';
      cHdr.innerHTML = `
        <div class="county-header-pill">
          <span class="county-name-text">${countyName} County</span>
          <span class="county-zone-count">${cPolys.length} zone${cPolys.length!==1?"s":""}</span>
          <span class="tip-wrap"><button class="county-action-btn" onclick="shareCounty('${stateAbbr}','${CSS.escape(countyName)}',event)">🔗</button><span class="tip-box tip-box-up tip-right" style="white-space:normal;width:190px;">Copy and paste a shareable link to ${countyName} County's page</span></span>
          <span class="tip-wrap"><button class="county-action-btn" onclick="deleteCounty('${stateAbbr}','${CSS.escape(countyName)}',event)">🗑</button><span class="tip-box tip-box-up tip-right">Delete saved zones in ${countyName} County</span></span>
        </div>
      `;

      // Sheet status row — clean, no gray background
      const cStatus = document.createElement('div');
      cStatus.className = 'county-sheet-status';
      cStatus.dataset.state = stateAbbr;
      cStatus.dataset.county = countyName;
      if (isConnected) {
        cStatus.innerHTML = `<span class="tip-wrap"><button class="spill connected" onclick="openSheetsModalForCounty('${stateAbbr}','${CSS.escape(countyName)}',event)"><span class="spill-dot"></span>Sheet Connected</button><span class="tip-box tip-box-up tip-right">Manage sheet connected to ${countyName} County</span></span>`;
      } else {
        cStatus.innerHTML = `<span class="tip-wrap"><button class="spill not-connected" onclick="openSheetsModalForCounty('${stateAbbr}','${CSS.escape(countyName)}',event)"><span class="spill-dot"></span>No Sheet Connected</button><span class="tip-box tip-box-up tip-right">Connect a sheet to ${countyName} County</span></span>`;
      }

      cHdr.onclick = e => {
        if (e.target.closest('.county-action-btn') || e.target.closest('.spill')) return;
        navigateToCounty(stateAbbr, countyName);
      };
      cStatus.onclick = e => { e.stopPropagation(); };

      const polyDiv = document.createElement('div');
      polyDiv.className = 'county-polys';

      cPolys.forEach(p => {
        const div = document.createElement('div');
        div.className = 'polygon-item';
        div.innerHTML = `
          <div style="width:10px;height:10px;border-radius:50%;background:${p.color};flex-shrink:0"></div>
          <div class="poly-info">
            <div class="poly-name">ZONE ${p.letter}</div>
            <div class="poly-count">${p.countyName ? p.countyName+' County, '+p.stateAbbr : ''}</div>
          </div>
          <span class="tip-wrap"><button class="poly-btn notes-btn" onclick="openZoneDescModal('${p.id}')">⚙</button><span class="tip-box tip-box-up tip-right">Open pricing panel</span></span>
          <span class="tip-wrap"><button class="poly-btn delete-btn">✕</button><span class="tip-box tip-box-up tip-right">Delete Zone ${p.letter}</span></span>
        `;
        div.querySelector('.notes-btn').addEventListener('click', e => { e.stopPropagation(); openZoneDescModal(p.id); });
        div.querySelector('.delete-btn').addEventListener('click', e => { e.stopPropagation(); deletePoly(p.id); });
        div.onclick = () => zoomToZoneAndCounty(p);
        polyDiv.appendChild(div);
      });

      cGroup.appendChild(cHdr);
      cGroup.appendChild(cStatus);
      cGroup.appendChild(polyDiv);
      countiesDiv.appendChild(cGroup);
    });

    const hdrWrap = document.createElement('div');
    hdrWrap.className = 'state-header-wrap';
    hdrWrap.appendChild(hdr);
    stateDiv.appendChild(hdrWrap);
    stateDiv.appendChild(countiesDiv);
    list.appendChild(stateDiv);
  });
}

// =========================================================
// COUNTY NAVIGATION
// =========================================================
async function navigateToCounty(stateAbbr, countyName) {
  const ss = document.getElementById('stateSelect');
  const cs = document.getElementById('countySelect');
  ss.value = stateAbbr;
  await loadCounties(true);
  cs.value = countyName;
  saveAppState();
  await loadCounty(); // also fits bounds
}

async function zoomToZoneAndCounty(poly) {
  // Zoom to zone polygon
  const b = new mapboxgl.LngLatBounds();
  poly.points.forEach(pt => b.extend(pt));
  map.fitBounds(b, { padding: 80 });

  // Update sidebar selectors
  if (poly.stateAbbr && poly.countyName) {
    const ss = document.getElementById('stateSelect');
    const cs = document.getElementById('countySelect');
    ss.value = poly.stateAbbr;
    await loadCounties(true);
    cs.value = poly.countyName;
    saveAppState();
    // Show county boundary without refitting the map view
    await loadCountyBoundaryOnly(poly.stateAbbr, poly.countyName);
  }
}

async function loadCountyBoundaryOnly(stateAbbr, countyName) {
  try {
    const key = _countyKey(stateAbbr, countyName);
    if (_countyLayers[key]) return; // already loaded, don't reload
    const fips = STATE_FIPS[stateAbbr];
    if (!fips) return;
    const geojson = await _fetchCountyGeoJSON(fips, countyName);
    if (!geojson) return;
    // Cache for validation use
    _countyGeoJSONCache[key] = geojson;
    _pendingCountyGeoJSON = geojson;
    // Use persistent keyed layer — doesn't remove other county boundaries
    _addCountyBoundaryForKey(key, geojson);
  } catch(e) {}
}


// =========================================================
// CUSTOM CONFIRM DIALOGS
// =========================================================
let _confirmResolve = null;

function _showConfirm({ title, sub, okLabel = 'Delete', typePhrase = null }) {
  return new Promise(resolve => {
    _confirmResolve = resolve;
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmSub').textContent = sub;
    document.getElementById('confirmOkBtn').textContent = okLabel;

    const typeWrap = document.getElementById('confirmTypeWrap');
    const typeInput = document.getElementById('confirmTypeInput');

    if (typePhrase) {
      typeWrap.style.display = '';
      document.getElementById('confirmTypePhrase').textContent = '"' + typePhrase + '"';
      typeInput.value = '';
      typeInput.classList.remove('valid');
      // Enable/disable OK based on match
      _updateConfirmOk(typePhrase);
      typeInput.oninput = () => _updateConfirmOk(typePhrase);
      document.getElementById('confirmOkBtn').disabled = true;
      document.getElementById('confirmOkBtn').style.opacity = '0.4';
    } else {
      typeWrap.style.display = 'none';
      typeInput.oninput = null;
      document.getElementById('confirmOkBtn').disabled = false;
      document.getElementById('confirmOkBtn').style.opacity = '1';
    }

    document.getElementById('confirmModal').classList.add('open');
    if (typePhrase) setTimeout(() => typeInput.focus(), 120);
  });
}

function _updateConfirmOk(phrase) {
  const input = document.getElementById('confirmTypeInput');
  const ok = document.getElementById('confirmOkBtn');
  const match = input.value.trim().toLowerCase() === phrase.toLowerCase();
  input.classList.toggle('valid', match);
  ok.disabled = !match;
  ok.style.opacity = match ? '1' : '0.4';
}

function _closeConfirm(result) {
  document.getElementById('confirmModal').classList.remove('open');
  document.getElementById('confirmTypeInput').oninput = null;
  if (_confirmResolve) { _confirmResolve(result); _confirmResolve = null; }
}

document.getElementById('confirmCancelBtn').addEventListener('click', () => _closeConfirm(false));
document.getElementById('confirmOkBtn').addEventListener('click', () => {
  if (!document.getElementById('confirmOkBtn').disabled) _closeConfirm(true);
});
// Close on overlay click
document.getElementById('confirmModal').addEventListener('click', e => {
  if (e.target === e.currentTarget) _closeConfirm(false);
});
// Enter key submits if OK is enabled
document.getElementById('confirmTypeInput').addEventListener('keydown', e => {
  if (e.key === 'Enter' && !document.getElementById('confirmOkBtn').disabled) _closeConfirm(true);
});

// =========================================================
// DELETE
// =========================================================
async function deletePoly(id, skipConfirm) {
  const i = polygons.findIndex(p => p.id === id);
  if (i === -1) return;
  const p = polygons[i];
  if (!skipConfirm) {
    const fullState = abbrToFullName(p.stateAbbr) || p.stateAbbr;
    const confirmed = await _showConfirm({
      title: `Delete Zone ${p.letter} in ${p.countyName} County, ${fullState}?`,
      sub: 'This action cannot be undone.',
      okLabel: 'Delete'
    });
    if (!confirmed) return;
  }
  _removeZoneLabel(p);
  if (p.handles) p.handles.forEach(h => { if (h && h.remove) h.remove(); });
  _removeZoneLayers(id);
  properties.forEach(prop => { if (prop.zone === p.name) { _markerColor(prop.marker, '#f7c948'); prop.zone = null; } });
  polygons.splice(i, 1);
  renderPolygonList(); persistZones(); _rebuildAllLabels();
}

async function deleteCounty(stateAbbr, countyName, evt) {
  if (evt) evt.stopPropagation();
  const toDelete = polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName === countyName);
  if (!toDelete.length) return;
  const fullState = abbrToFullName(stateAbbr) || stateAbbr;
  const multi = toDelete.length > 1;
  const letters = toDelete.map(p => p.letter);
  const zoneList = letters.length > 2
    ? 'Zones ' + letters.slice(0, -1).join(', ') + ' and ' + letters[letters.length - 1]
    : letters.length === 2
    ? 'Zones ' + letters[0] + ' and ' + letters[1]
    : 'Zone ' + letters[0];
  const title = multi
    ? `Delete ${zoneList} in ${countyName} County, ${fullState}?`
    : `Delete Zone ${toDelete[0].letter} in ${countyName} County, ${fullState}?`;
  const sub = multi
    ? `This will permanently delete ${zoneList}. This cannot be undone.`
    : 'This action cannot be undone.';
  const confirmed = await _showConfirm({ title, sub, okLabel: 'Delete' });
  if (!confirmed) return;
  toDelete.forEach(p => {
    _removeZoneLabel(p);
    if (p.handles) p.handles.forEach(h => { if (h && h.remove) h.remove(); });
    _removeZoneLayers(p.id);
    properties.forEach(prop => { if (prop.zone === p.name) { _markerColor(prop.marker, '#f7c948'); prop.zone = null; } });
  });
  polygons = polygons.filter(p => !(p.stateAbbr === stateAbbr && p.countyName === countyName));
  renderPolygonList(); persistZones(); _rebuildAllLabels();
}

async function clearAllZones() {
  if (!polygons.length) return;
  // Step 1 — initial warning
  const step1 = await _showConfirm({
    title: `Delete all ${polygons.length} saved zone${polygons.length !== 1 ? 's' : ''}?`,
    sub: `This will permanently delete all ${polygons.length} zone${polygons.length !== 1 ? 's' : ''}. This cannot be undone.`,
    okLabel: 'OK'
  });
  if (!step1) return;
  // Step 2 — type to confirm
  const step2 = await _showConfirm({
    title: 'Type to confirm deletion',
    sub: 'This is permanent and cannot be undone.',
    okLabel: 'Delete',
    typePhrase: 'Delete all saved zones'
  });
  if (!step2) return;
  polygons.forEach(p => { _removeZoneLabel(p); if (p.handles) p.handles.forEach(h=>{if(h&&h.remove)h.remove();}); _removeZoneLayers(p.id); });
  polygons = [];
  properties.forEach(prop => { if (prop.zone) { _markerColor(prop.marker, '#f7c948'); prop.zone = null; } });
  renderPolygonList(); persistZones(); _rebuildAllLabels();
  showToast('All zones cleared', 'info');
}

// =========================================================
// PERSIST ZONES
// =========================================================
function _polyToJSON(p) {
  return { id:p.id, name:p.name, letter:p.letter||'', stateAbbr:p.stateAbbr||'', countyName:p.countyName||'',
           color:p.color, points:p.points, description:p.description||'', pricingTiers:p.pricingTiers||[], isRect:!!p._isRect, bounds:p._bounds||null };
}
function persistZones() {
  DB.saveZones(polygons.map(_polyToJSON));
}
async function _loadAllCountyBoundaries() {
  // Find all unique state+county combos that have zones
  const combos = [...new Set(polygons.filter(p => p.stateAbbr && p.countyName).map(p => _countyKey(p.stateAbbr, p.countyName)))];
  for (const key of combos) {
    if (_countyLayers[key]) continue; // already loaded
    const [sa, cn] = key.split('|');
    const fips = STATE_FIPS[sa]; if (!fips) continue;
    try {
      const geojson = await _fetchCountyGeoJSON(fips, cn);
      if (geojson) { _addCountyBoundaryForKey(key, geojson); }
    } catch(e) {}
  }
}

function restoreZones() {
  try {
    const data = DB.loadZones();
    if (!data || !Array.isArray(data) || !data.length) return;
    data.forEach(d => _loadZone(d));
    _rebuildAllLabels();
    showToast(`Restored ${data.length} zone${data.length>1?'s':''}`, 'success');
    setTimeout(() => _loadAllCountyBoundaries(), 500);
  } catch(e) { console.error('restoreZones error:', e); }
}
function _loadZone(d) {
  const poly = { id:d.id, name:d.name, letter:d.letter||'', stateAbbr:d.stateAbbr||'', countyName:d.countyName||'',
    color:d.color, points:d.points, description:d.description||'', pricingTiers:d.pricingTiers||[], labelMarker:null, handles:[], _isRect:d.isRect||false, _bounds:d.bounds||null };
  // Back-compat: derive stateAbbr/countyName from name if missing
  if (!poly.stateAbbr || !poly.countyName) {
    const m = poly.name.match(/^(.+) County,\s*([A-Z]{2})$/);
    if (m) { poly.countyName = m[1]; poly.stateAbbr = m[2]; }
    else { const m2 = poly.name.match(/^(.+),\s*([A-Z]{2})$/); if(m2){poly.countyName=m2[1];poly.stateAbbr=m2[2];} }
  }
  polygons.push(poly);
  _addZoneLayers(poly);
  _addZoneLabel(poly);
  return poly;
}

// =========================================================
// SAVE / LOAD ZONES FILE
// =========================================================
function saveZonesFile() {
  if (!polygons.length) { showToast('No zones to save', 'error'); return; }
  const blob = new Blob([JSON.stringify({ version:2, saved:new Date().toISOString(), zones:polygons.map(_polyToJSON) }, null, 2)], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `countyzone-${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  showToast('Zones file downloaded!', 'success');
}
function loadZonesFile() { document.getElementById('zoneFileInput').click(); }
function importZonesFile(e) {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      const zones = data.zones || data;
      if (!Array.isArray(zones)) throw new Error();
      polygons.forEach(p => { _removeZoneLabel(p); if(p.handles)p.handles.forEach(h=>{if(h&&h.remove)h.remove();}); _removeZoneLayers(p.id); });
      polygons = [];
      zones.forEach(d => _loadZone(d));
      renderPolygonList(); persistZones();
      _rebuildAllLabels();
      showToast(`Loaded ${zones.length} zone${zones.length!==1?'s':''}`, 'success');
      if (polygons.length) {
        const b = new mapboxgl.LngLatBounds();
        polygons.forEach(p => p.points.forEach(pt => b.extend(pt)));
        map.fitBounds(b, { padding:60 });
      }
    } catch(err) { showToast('Could not read zones file', 'error'); }
    e.target.value = '';
  };
  reader.readAsText(file);
}

// =========================================================
// SHARE URL
// =========================================================
function shareURL() {
  if (!polygons.length) { showToast('Draw some zones first', 'error'); return; }
  const encoded = btoa(encodeURIComponent(JSON.stringify(polygons.map(_polyToJSON))));
  const url = window.location.origin + window.location.pathname + '?zones=' + encoded;
  if (url.length > 8000) { showToast('Too many zones — use Save Zones file instead', 'error'); return; }
  navigator.clipboard.writeText(url).then(() => showToast('Share link copied!', 'success')).catch(() => {
    const b = document.getElementById('shareBanner');
    b.textContent = '🔗 ' + url; b.style.display = 'block';
    setTimeout(() => b.style.display = 'none', 8000);
  });
}
function copyShareURL() {
  const b = document.getElementById('shareBanner');
  navigator.clipboard.writeText(b.textContent.replace('🔗 ','')).then(() => { showToast('Copied!','success'); b.style.display='none'; });
}
function loadZonesFromURL() {
  try {
    const shareId = new URLSearchParams(window.location.search).get('share');
    if (shareId) {
      // Load from short share ID after map loads
      setTimeout(() => loadZonesFromShareId(shareId), 200);
      return true;
    }
    const enc = new URLSearchParams(window.location.search).get('zones');
    if (!enc) return false;
    JSON.parse(decodeURIComponent(atob(enc))).forEach(d => _loadZone(d));
    renderPolygonList();
    _rebuildAllLabels();
    if (polygons.length) {
      const b = new mapboxgl.LngLatBounds();
      polygons.forEach(p => p.points.forEach(pt => b.extend(pt)));
      map.fitBounds(b, { padding:60 });
      showToast(`Loaded ${polygons.length} shared zone${polygons.length!==1?'s':''}`, 'success');
    }
    return true;
  } catch(e) { return false; }
}

// =========================================================
// SIDEBAR TOGGLE
// =========================================================
function toggleSidebar() { document.getElementById('sidebar').classList.toggle('open'); }

// =========================================================
// GOOGLE SHEETS
// =========================================================
function openSheetsModal() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  const badge = document.getElementById('sheetsModalCounty');
  const existing = (sa && cn) ? _getSheetConfig(sa, cn) : null;

  // County badge
  if (sa && cn) {
    badge.textContent = `${cn} County, ${sa}`;
    badge.style.display = '';
  } else {
    badge.style.display = 'none';
  }

  // Connection status row
  const dot = document.getElementById('smDot');
  const statusText = document.getElementById('smStatusText');
  const disconnectBtn = document.getElementById('smDisconnectBtn');
  if (existing && existing.sheetId) {
    dot.className = 'spill-dot';
    dot.style.background = 'var(--green)';
    dot.style.boxShadow = '0 0 4px rgba(46,138,90,0.4)';
    statusText.textContent = 'Sheet Connected';
    statusText.style.color = 'var(--green)';
    disconnectBtn.style.display = '';
  } else {
    dot.className = 'spill-dot';
    dot.style.background = 'var(--red)';
    dot.style.boxShadow = '';
    statusText.textContent = 'No Sheet Connected';
    statusText.style.color = '#b94040';
    statusText.style.color = 'var(--muted)';
    disconnectBtn.style.display = 'none';
  }

  // Populate fields
  document.getElementById('smSheetNameField').value = '';
  if (existing) {
    document.getElementById('sheetId').value    = existing.sheetId    || '';
    document.getElementById('sheetName').value  = existing.sheetName  || 'LI Raw Dataset';
    // Fetch live sheet name for read-only field
    if (existing.sheetId) setTimeout(() => _fetchSheetName(existing.sheetId), 100);
    document.getElementById('colLat').value     = existing.colLat     || 'Latitude';
    document.getElementById('colLng').value     = existing.colLng     || 'Longitude';
    document.getElementById('colAPN').value     = existing.colAPN     || 'APN';
    document.getElementById('colCity').value    = existing.colCity    || 'City';
    document.getElementById('colState').value   = existing.colState   || 'State';
    document.getElementById('colZip').value     = existing.colZip     || 'ZIP';
    document.getElementById('colZone').value    = existing.colZone    || 'County Zone';
  } else {
    document.getElementById('sheetId').value = '';
  }

  document.getElementById('sheetsModal').classList.add('open');
}
function closeSheetsModal() { document.getElementById('sheetsModal').classList.remove('open'); }

function disconnectSheet() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  if (!sa || !cn) return;
  delete sheetConfigs[_countyKey(sa, cn)];
  DB.saveSheetConfigs(sheetConfigs);
  if (sheetConfig && sheetConfig.stateAbbr === sa && sheetConfig.countyName === cn) {
    sheetConfig = null;
    setConnected(false);
  }
  // Reset modal status
  const _d = document.getElementById('smDot');
  _d.style.background = '#b94040'; _d.style.boxShadow = '';
  const _st = document.getElementById('smStatusText');
  _st.textContent = 'No Sheet Connected'; _st.style.color = '#b94040';
  document.getElementById('smDisconnectBtn').style.display = 'none';
  document.getElementById('sheetId').value = '';
  renderPolygonList();
  showToast('Sheet disconnected', 'info');
}

function disconnectSheetForCounty(stateAbbr, countyName, evt) {
  if (evt) evt.stopPropagation();
  const key = _countyKey(stateAbbr, countyName);
  delete sheetConfigs[key];
  DB.saveSheetConfigs(sheetConfigs);
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  if (sa === stateAbbr && cn === countyName) {
    sheetConfig = null;
    setConnected(false);
  }
  renderPolygonList();
  showToast(`Sheet disconnected from ${countyName} County`, 'info');
}

async function connectSheets() {
  const sheetId = document.getElementById('sheetId').value.trim();
  if (!sheetId) { showToast('Please enter a Spreadsheet ID', 'error'); return; }
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  if (!sa || !cn) { showToast('Please select a State and County first', 'error'); return; }
  // Check if this sheet ID is already used by a different county
  const existingKey = Object.entries(sheetConfigs).find(([key, cfg]) => {
    return cfg.sheetId === sheetId && key !== _countyKey(sa, cn);
  });
  if (existingKey) {
    const [existKey] = existingKey;
    const [existState, existCounty] = existKey.split('|');
    showToast(`This sheet is already connected to ${existCounty} County, ${existState}`, 'error');
    return;
  }

  sheetConfig = {
    sheetId,
    stateAbbr:  sa,
    countyName: cn,
    sheetName:  document.getElementById('sheetName').value.trim()  || 'LI Raw Dataset',
    colLat:     document.getElementById('colLat').value.trim()     || 'Latitude',
    colLng:     document.getElementById('colLng').value.trim()     || 'Longitude',
    colAPN:     document.getElementById('colAPN').value.trim()     || 'APN',
    colCity:    document.getElementById('colCity').value.trim()    || 'City',
    colState:   document.getElementById('colState').value.trim()   || 'State',
    colZip:     document.getElementById('colZip').value.trim()     || 'ZIP',
    colZone:    document.getElementById('colZone').value.trim()    || 'County Zone',
  };
  // Save per-county
  _setSheetConfig(sa, cn, sheetConfig);
  showToast('Connecting to Google Sheets...', 'info');
  try {
    const r = await fetch('/.netlify/functions/sheets-read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetId: sheetConfig.sheetId, sheetName: sheetConfig.sheetName }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.error || 'HTTP ' + r.status);
    loadPropertiesFromFunction(data.properties);
    setConnected(true);
    // Fetch sheet name for modal display
    setTimeout(() => _fetchSheetName(sheetConfig.sheetId), 200);
    closeSheetsModal();
    showToast(`Connected: ${cn} County, ${sa} — ${data.properties.length} properties loaded`, 'success');
  } catch(e) { showToast('Connection failed: ' + e.message, 'error'); }
}
function loadPropertiesFromFunction(props) {
  properties.forEach(p => { if (p.marker) p.marker.remove(); });
  properties = [];
  let skipped = 0;
  props.forEach(prop => {
    let { lat, lng, apn, address, city, state, zip, county, acreage, zone, rowIndex } = prop;

    // Validate coordinates
    // Auto-swap if clearly reversed
    if (lat < -90 && lng > 0) { [lat, lng] = [lng, lat]; }
    else if (lng > 0 && lat > 0 && lat < 90) { lng = -lng; }
    // Must be real decimal coordinates (not row numbers or integers)
    // Real coords have decimal precision; row numbers are integers
    const latStr = String(prop.lat || '');
    const lngStr = String(prop.lng || '');
    const hasDecimal = latStr.includes('.') && lngStr.includes('.');
    if (!hasDecimal) { skipped++; return; }
    // US bounds: lat 18–72, lng -180 to -60
    if (lat < 18 || lat > 72 || lng < -180 || lng > -60) { skipped++; return; }

    // Filter to selected county only
    const _selCounty = document.getElementById('countySelect').value;
    if (_selCounty && prop.county) {
      const propCounty = prop.county.toLowerCase().replace(' county','').trim();
      const selCounty  = _selCounty.toLowerCase().trim();
      if (propCounty !== selCounty) { skipped++; return; }
    }

    // Store property data without map marker (pins disabled pending better implementation)
    properties.push({ lat, lng, apn, address, city, state, zip, county, acreage, zone: zone || null, rowIndex, marker: null });
  });
  document.getElementById('statProps').textContent = properties.length;
  if (skipped) showToast(`${skipped} properties skipped — coordinates out of range`, 'info');
}

// =========================================================
// ZONE ASSIGNMENT
// =========================================================
function pointInPolygon(lat, lng, pts) {
  const x=lng,y=lat; let inside=false;
  for(let i=0,j=pts.length-1;i<pts.length;j=i++){
    const xi=pts[i][0],yi=pts[i][1],xj=pts[j][0],yj=pts[j][1];
    if(((yi>y)!==(yj>y))&&(x<(xj-xi)*(y-yi)/(yj-yi)+xi)) inside=!inside;
  }
  return inside;
}
async function runAssignment() {
  if (!polygons.length) { showToast('Draw at least one zone first', 'error'); return; }
  if (!properties.length) { showToast('No properties loaded. Connect Google Sheets first.', 'error'); return; }
  polygons.forEach(p => p.propCount = 0);
  let assigned = 0;

  // Step 1 — compute assignments in memory
  const assignments = [];
  properties.forEach(prop => {
    prop.zone = null;
    for (const poly of polygons) {
      if (pointInPolygon(prop.lat, prop.lng, poly.points)) {
        prop.zone = poly.letter; // write letter only (A, B, C...)
        poly.propCount++;
        // marker coloring disabled (pins not shown)
        assigned++;
        if (prop.rowIndex) assignments.push({ rowIndex: prop.rowIndex, zone: poly.letter });
        break;
      }
    }
    // if (!prop.zone) _markerColor(prop.marker, '#f7c948');
  });

  document.getElementById('statAssigned').textContent = assigned;
  renderPolygonList();
  showToast(`${assigned} properties assigned — writing to sheet...`, 'info');

  // Step 2 — write zone letters back to sheet via Netlify function
  // Always get freshest config from the per-county map
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  const activeCfg = (sa && cn) ? (_getSheetConfig(sa, cn) || sheetConfig) : sheetConfig;

  if (assignments.length && activeCfg && activeCfg.sheetId) {
    try {
      const r = await fetch('/.netlify/functions/sheets-write-zones', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sheetId:   activeCfg.sheetId,
          sheetName: activeCfg.sheetName || 'LI Raw Dataset',
          assignments,
        }),
      });
      const data = await r.json();
      if (!r.ok) throw new Error(data.error || 'HTTP ' + r.status);
      showToast(`${assignments.length} zone assignments written to sheet ✓`, 'success');
    } catch(e) {
      showToast('Zone write failed: ' + e.message, 'error');
      console.error('sheets-write-zones error:', e);
    }
  } else {
    showToast(`${assigned} of ${properties.length} properties assigned`, 'success');
  }
}
function _markerColor(m, c) { if (m) m.getElement().style.background = c; }

// Pin toggle removed — pins disabled pending better implementation

// =========================================================
// EXPORT CSV
// =========================================================
function exportCSV() {
  if (!properties.length) { showToast('No properties to export', 'error'); return; }
  const zoneCol = sheetConfig.colZone || 'Zone';
  const headers = [...properties[0].headers];
  if (!headers.includes(zoneCol)) headers.push(zoneCol);
  const zi = headers.indexOf(zoneCol);
  const rows = [headers.map(ce).join(',')];
  properties.forEach(prop => {
    const row = [...prop.row];
    while (row.length < headers.length) row.push('');
    row[zi] = prop.zone || '';
    rows.push(row.map(ce).join(','));
  });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([rows.join('\n')], {type:'text/csv'}));
  a.download = 'properties_with_zones.csv'; a.click();
  showToast('CSV downloaded!', 'success');
}
function ce(v) {
  if (v==null) return '';
  const s = String(v);
  return (s.includes(',')||s.includes('"')||s.includes('\n')) ? `"${s.replace(/"/g,'""')}"` : s;
}
function parseCSV(text) {
  const rows = [];
  text.split('\n').forEach(line => {
    if (!line.trim()) return;
    const row=[]; let cur='',inQ=false;
    for(let i=0;i<line.length;i++){
      if(line[i]==='"'){if(inQ&&line[i+1]==='"'){cur+='"';i++;}else inQ=!inQ;}
      else if(line[i]===','&&!inQ){row.push(cur.trim());cur='';}
      else cur+=line[i];
    }
    row.push(cur.trim()); rows.push(row);
  });
  return rows;
}

// =========================================================
// COUNTY BOUNDARY
// =========================================================
const stateSelect = document.getElementById('stateSelect');
STATES.forEach(([name, abbr]) => {
  const o = document.createElement('option'); o.value = abbr; o.textContent = name;
  stateSelect.appendChild(o);
});

async function loadCounties(silent) {
  const abbr = stateSelect.value; if (!abbr) return;
  const cs = document.getElementById('countySelect');
  if (!silent) cs.innerHTML = '<option value="">Loading counties...</option>';
  function fill(counties) {
    cs.innerHTML = '<option value="">— Select County —</option>';
    counties.forEach(n => { const o=document.createElement('option'); o.value=n; o.textContent=n+' County'; cs.appendChild(o); });
    if (!silent) { const s=loadAppState(); if(s&&s.state===abbr&&s.county) cs.value=s.county; }
  }
  try {
    const cached = DB.loadCountyCache(abbr);
    if (cached) {
      const p = typeof cached === 'string' ? JSON.parse(cached) : cached;
      if (p&&p.counties&&p.abbr===abbr&&Date.now()-p.ts<30*24*60*60*1000) { fill(p.counties); return; }
      DB.clearCountyCache(abbr);
    }
  } catch(e) {}
  try {
    const fips = STATE_FIPS[abbr];
    const data = await (await fetch(`https://api.census.gov/data/2020/dec/pl?get=NAME&for=county:*&in=state:${fips}`)).json();
    const stateName = STATES.find(s=>s[1]===abbr)[0];
    const counties = data.slice(1).map(row=>row[0].replace(`, ${stateName}`,'').replace(/ County$/,'').trim()).sort();
    DB.saveCountyCache(abbr, counties);
    fill(counties);
  } catch(e) {
    try { const _cc=DB.loadCountyCache(abbr); if(_cc){const _ccp=typeof _cc==='string'?JSON.parse(_cc):_cc; fill(_ccp.counties||_ccp);return;} } catch(e2){}
    if (!silent) cs.innerHTML='<option value="">Failed to load</option>';
  }
}

function _removeCountyLayer() {
  if (!countySourceId) return;
  if (map.getLayer(countySourceId+'-fill')) map.removeLayer(countySourceId+'-fill');
  if (map.getLayer(countySourceId+'-line')) map.removeLayer(countySourceId+'-line');
  if (map.getSource(countySourceId))        map.removeSource(countySourceId);
  countySourceId = null;
}
function _addCountyBoundaryForKey(key, geojson) {
  // Add/update a named county boundary layer
  const existing = _countyLayers[key];
  if (existing) {
    if (map.getLayer(existing+'-fill')) map.removeLayer(existing+'-fill');
    if (map.getLayer(existing+'-line')) map.removeLayer(existing+'-line');
    if (map.getSource(existing))        map.removeSource(existing);
  }
  const sid = 'county-' + key.replace(/[^a-zA-Z0-9]/g,'-') + '-' + Date.now();
  _countyLayers[key] = sid;
  map.addSource(sid, { type:'geojson', data:geojson });
  map.addLayer({ id:sid+'-fill', type:'fill', source:sid, paint:{'fill-color':'#000000','fill-opacity':0.08} });
  map.addLayer({ id:sid+'-line', type:'line', source:sid, paint:{'line-color':'#6600cc','line-width':2} });
  // Click boundary to switch to that county
  map.on('click', sid+'-fill', (e) => {
    if (drawMode === 'polygon') return; // don't switch while drawing
    const [sa, cn] = key.split('|');
    if (stateSelect.value === sa && document.getElementById('countySelect').value === cn) return;
    stateSelect.value = sa;
    loadCounties().then(() => {
      document.getElementById('countySelect').value = cn;
      loadCounty();
      // Restore sheet config for this county
      const saved = _getSheetConfig(sa, cn);
      if (saved) { sheetConfig = saved; setConnected(true); }
      else { sheetConfig = null; setConnected(false); }
      showToast(`Switched to ${cn} County, ${sa}`, 'success');
    });
  });
  map.on('mouseenter', sid+'-fill', () => { if (drawMode !== 'polygon') map.getCanvas().style.cursor = 'pointer'; });
  map.on('mouseleave', sid+'-fill', () => { if (drawMode !== 'polygon') map.getCanvas().style.cursor = ''; });
}

function _readdCountyLayer(geojson) {
  const _apply = () => {
    try {
      if (countySourceId) {
        if (map.getLayer(countySourceId+'-fill')) map.removeLayer(countySourceId+'-fill');
        if (map.getLayer(countySourceId+'-line')) map.removeLayer(countySourceId+'-line');
        if (map.getSource(countySourceId))        map.removeSource(countySourceId);
      }
      countySourceId = 'county-' + Date.now();
      map.addSource(countySourceId, { type:'geojson', data:geojson });
      map.addLayer({ id:countySourceId+'-fill', type:'fill', source:countySourceId, paint:{'fill-color':'#000000','fill-opacity':0.12} });
      map.addLayer({ id:countySourceId+'-line', type:'line', source:countySourceId, paint:{'line-color':'#6600cc','line-width':3} });
    } catch(e) {
      // If map isn't ready, retry once on next idle
      map.once('idle', () => _readdCountyLayer(geojson));
    }
  };
  // If map is currently moving, wait for it to settle first
  if (map.isMoving() || map.isZooming()) {
    map.once('idle', _apply);
  } else {
    _apply();
  }
}

async function _fetchCountyGeoJSON(fips, countyName) {
  const name = encodeURIComponent(countyName);

  // 1. Census2020 — stable versioned service, most reliable
  const tigerUrls = [
    `https://tigerweb.geo.census.gov/arcgis/rest/services/Census2020/tigerWMS_Census2020/MapServer/82/query?where=STATE%3D'${fips}'%20AND%20BASENAME%3D'${name}'&outFields=*&f=geojson`,
    `https://tigerweb.geo.census.gov/arcgis/rest/services/Census2020/State_County/MapServer/1/query?where=STATE%3D'${fips}'%20AND%20BASENAME%3D'${name}'&outFields=*&f=geojson`,
    `https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/tigerWMS_Current/MapServer/82/query?where=STATE%3D'${fips}'%20AND%20BASENAME%3D'${name}'&outFields=*&f=geojson`,
  ];
  for (const url of tigerUrls) {
    try {
      const r = await fetch(url);
      if (!r.ok) continue;
      const data = await r.json();
      if (data.features && data.features.length) return data;
    } catch(e) { continue; }
  }

  // 2. Nominatim (OpenStreetMap) — completely independent fallback
  try {
    const stateNames = {"01":"Alabama","02":"Alaska","04":"Arizona","05":"Arkansas","06":"California","08":"Colorado","09":"Connecticut","10":"Delaware","12":"Florida","13":"Georgia","15":"Hawaii","16":"Idaho","17":"Illinois","18":"Indiana","19":"Iowa","20":"Kansas","21":"Kentucky","22":"Louisiana","23":"Maine","24":"Maryland","25":"Massachusetts","26":"Michigan","27":"Minnesota","28":"Mississippi","29":"Missouri","30":"Montana","31":"Nebraska","32":"Nevada","33":"New Hampshire","34":"New Jersey","35":"New Mexico","36":"New York","37":"North Carolina","38":"North Dakota","39":"Ohio","40":"Oklahoma","41":"Oregon","42":"Pennsylvania","44":"Rhode Island","45":"South Carolina","46":"South Dakota","47":"Tennessee","48":"Texas","49":"Utah","50":"Vermont","51":"Virginia","53":"Washington","54":"West Virginia","55":"Wisconsin","56":"Wyoming"};
    const stateName = stateNames[fips];
    if (stateName) {
      const url = `https://nominatim.openstreetmap.org/search?county=${encodeURIComponent(countyName)}&state=${encodeURIComponent(stateName)}&country=USA&format=geojson&polygon_geojson=1&limit=1`;
      const r = await fetch(url, { headers: { 'User-Agent': 'CountyZonePro/1.0' } });
      if (r.ok) {
        const data = await r.json();
        if (data.features && data.features.length) {
          // Nominatim returns slightly different structure — normalize it
          return { type: 'FeatureCollection', features: data.features };
        }
      }
    }
  } catch(e) {}

  return null;
}

async function loadCounty() {
  const abbr = stateSelect.value, county = document.getElementById('countySelect').value;
  saveAppState(); if (!abbr||!county) return;
  _removeCountyLayer();
  showToast('Loading county boundary...', 'info');
  try {
    const fips = STATE_FIPS[abbr];
    const geojson = await _fetchCountyGeoJSON(fips, county);
    if (!geojson) { showToast('County boundary not found','error'); return; }
    _pendingCountyGeoJSON = geojson;
    // Register in persistent keyed layers so it survives county switching
    const key = _countyKey(abbr, county);
    _countyGeoJSONCache[key] = geojson;
    _addCountyBoundaryForKey(key, geojson);
    _readdCountyLayer(geojson); // also set countySourceId for validation
    const bounds = new mapboxgl.LngLatBounds();
    geojson.features.forEach(f => {
      const coords = f.geometry.type==='Polygon' ? f.geometry.coordinates.flat()
                   : f.geometry.type==='MultiPolygon' ? f.geometry.coordinates.flat(2) : [];
      coords.forEach(c => bounds.extend(c));
    });
    map.fitBounds(bounds, { padding:60 });
    showToast(`${county} County loaded`, 'success');
  } catch(e) { showToast('Could not load county boundary: ' + e.message, 'error'); console.error('loadCounty error:', e); }
}

// =========================================================
// HELPERS
// =========================================================
function setConnected(v) {
  const dot = document.getElementById('statusDot');
  const txt = document.getElementById('statusText');
  if (dot) dot.className = 'status-dot'+(v?' on':'');
  if (txt) txt.textContent = v ? 'SHEET CONNECTED' : 'NOT CONNECTED';
  // Refresh county status rows to reflect new connection state
  renderPolygonList();
}
let toastTimer;
function showToast(msg, type='info') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast '+type+' show';
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 4500);
}
function saveAppState() {
  DB.saveAppState({ state: stateSelect.value, county: document.getElementById('countySelect').value });
}
function loadAppState() {
  return DB.loadAppState();
}

// ── TOOLTIP TOGGLE ──────────────────────────────────────
function toggleTooltips() {
  const isOff = document.body.classList.toggle('tooltips-off');
  DB.saveUIState('tooltips_off', isOff);
  _updateTooltipBtn(isOff);
}
function _updateTooltipBtn(isOff) {
  const btn = document.getElementById('tooltipToggleBtn');
  if (btn) btn.textContent = isOff ? '💬 Tooltips: Off' : '💬 Tooltips: On';
}
function _initTooltipToggle() {
  const isOff = DB.loadUIState('tooltips_off', false);
  if (isOff) document.body.classList.add('tooltips-off');
  _updateTooltipBtn(isOff);
}

document.getElementById('sheetsModal').addEventListener('click', e => { if (e.target===e.currentTarget) closeSheetsModal(); });

// =========================================================
// INIT
// =========================================================
map.on('load', () => {
  _initDrawLayers();
  _initTooltipToggle(); // sets body class and updates button label

  // Load sheet configs first so they're available during zone restore
  const _savedCfgs = DB.loadSheetConfigs();
  if (_savedCfgs) { sheetConfigs = _savedCfgs; }

  // Restore zones from localStorage
  const fromURL = loadZonesFromURL();
  if (!fromURL) restoreZones();

  // Safety net: rebuild labels/boundaries after map is fully ready
  setTimeout(() => {
    if (polygons.length) { _rebuildAllLabels(); _loadAllCountyBoundaries(); }
  }, 500);

  // Restore state/county dropdowns and reconnect sheet
  const appState = loadAppState();
  if (appState && appState.state) {
    stateSelect.value = appState.state;
    loadCounties().then(() => {
      const countySelect = document.getElementById('countySelect');
      if (appState.county) {
        countySelect.value = appState.county;

        // Restore active sheet config for this county
        const saved = _getSheetConfig(appState.state, appState.county);
        if (saved) {
          sheetConfig = saved;
          setConnected(true);
          renderPolygonList(); // show sidebar with connected status

          // Reconnect to sheet and reload properties
          showToast('Reconnecting to sheet...', 'info');
          fetch('/.netlify/functions/sheets-read', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sheetId: saved.sheetId, sheetName: saved.sheetName || 'LI Raw Dataset' }),
          }).then(r => r.json()).then(data => {
            if (data.properties && data.properties.length) {
              loadPropertiesFromFunction(data.properties);
              document.getElementById('statProperties').textContent = data.properties.length;

              // Assign properties to zones — wait until polygons are on the map
              const _doAssign = () => {
                let assigned = 0;
                properties.forEach(prop => {
                  prop.zone = null;
                  for (const poly of polygons) {
                    if (pointInPolygon(prop.lat, prop.lng, poly.points)) {
                      prop.zone = poly.letter;
                      assigned++;
                      break;
                    }
                  }
                });
                document.getElementById('statAssigned').textContent = assigned;
                renderPolygonList();
                persistZones(); // save with updated assignments
                showToast(`Sheet reconnected — ${assigned} properties assigned`, 'success');
              };

              if (polygons.length) { _doAssign(); }
              else { map.once('idle', _doAssign); }
            }
          }).catch(err => {
            console.warn('Sheet reconnect failed:', err);
          });
        } else {
          renderPolygonList(); // show sidebar even without sheet
        }
      }
    });
  } else if (polygons.length) {
    // Zones exist but no saved app state — still render the list
    renderPolygonList();
  }
});

