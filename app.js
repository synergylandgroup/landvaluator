// =========================================================
// SUPABASE CLIENT
// =========================================================
const SUPABASE_URL  = 'https://dcrxczsgcuiwimwpokxo.supabase.co';
const SUPABASE_KEY  = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRjcnhjenNnY3Vpd2ltd3Bva3hvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ3NDU4MDAsImV4cCI6MjA5MDMyMTgwMH0.BFNKnN5mzaGLazQQTNhl8TytA5JW5IQxa5ouFg4-KB4';
const _supa = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
let _currentUser = null;

// =========================================================
// AUTH FUNCTIONS
// =========================================================
function _authSwitchTab(tab) {
  const isLogin = tab === 'login';
  document.getElementById('authTabLogin').classList.toggle('active', isLogin);
  document.getElementById('authTabSignup').classList.toggle('active', !isLogin);
  document.getElementById('authFormLogin').style.display = isLogin ? '' : 'none';
  document.getElementById('authFormSignup').style.display = isLogin ? 'none' : '';
  document.getElementById('authError').textContent = '';
}

async function _authLogin() {
  const email = document.getElementById('authLoginEmail').value.trim();
  const password = document.getElementById('authLoginPassword').value;
  const btn = document.getElementById('authLoginBtn');
  const err = document.getElementById('authError');
  if (!email || !password) { err.textContent = 'Please enter your email and password.'; return; }
  btn.disabled = true; btn.textContent = 'Signing in...'; err.textContent = '';
  const { error } = await _supa.auth.signInWithPassword({ email, password });
  if (error) { err.textContent = error.message; btn.disabled = false; btn.textContent = 'Sign In'; }
}

async function _authSignup() {
  const email = document.getElementById('authSignupEmail').value.trim();
  const password = document.getElementById('authSignupPassword').value;
  const confirm = document.getElementById('authSignupConfirm').value;
  const btn = document.getElementById('authSignupBtn');
  const err = document.getElementById('authError');
  if (!email || !password) { err.textContent = 'Please fill in all fields.'; return; }
  if (password.length < 8) { err.textContent = 'Password must be at least 8 characters.'; return; }
  if (password !== confirm) { err.textContent = 'Passwords do not match.'; return; }
  btn.disabled = true; btn.textContent = 'Creating account...'; err.textContent = '';
  const { error } = await _supa.auth.signUp({ email, password });
  if (error) { err.textContent = error.message; btn.disabled = false; btn.textContent = 'Create Account'; }
  else { err.style.color = 'var(--green)'; err.textContent = 'Account created! You are now signed in.'; }
}

async function _authSignOut() {
  _toggleUserMenu();
  await _supa.auth.signOut();
}

function _authShowReset() {
  document.getElementById('authModal').classList.remove('open');
  document.getElementById('resetModal').classList.add('open');
  document.getElementById('resetError').textContent = '';
  document.getElementById('resetEmail').value = document.getElementById('authLoginEmail').value || '';
}

async function _authSendReset() {
  const email = document.getElementById('resetEmail').value.trim();
  const btn = document.getElementById('resetBtn');
  const err = document.getElementById('resetError');
  if (!email) { err.textContent = 'Please enter your email.'; return; }
  btn.disabled = true; btn.textContent = 'Sending...'; err.textContent = '';
  const { error } = await _supa.auth.resetPasswordForEmail(email, {
    redirectTo: 'https://landvaluator.app',
  });
  if (error) { err.textContent = error.message; btn.disabled = false; btn.textContent = 'Send Reset Link'; }
  else {
    err.style.color = 'var(--green)';
    err.textContent = 'Check your email for a reset link.';
    btn.disabled = false; btn.textContent = 'Send Reset Link';
  }
}

function _toggleUserMenu() {
  document.getElementById('userDropdown').classList.toggle('open');
}

function _updateUserUI(user) {
  const wrap = document.getElementById('userMenuWrap');
  const avatar = document.getElementById('userAvatar');
  const label = document.getElementById('userMenuLabel');
  const emailEl = document.getElementById('userDropdownEmail');
  if (user) {
    const initial = (user.email || '?')[0].toUpperCase();
    avatar.textContent = initial;
    label.textContent = user.email.split('@')[0];
    emailEl.textContent = user.email;
    wrap.style.display = '';
  } else {
    wrap.style.display = 'none';
  }
}

// Close user menu on outside click
document.addEventListener('click', (e) => {
  if (!e.target.closest('.user-menu-wrap')) {
    document.getElementById('userDropdown')?.classList.remove('open');
  }
});

// Enter key support for auth forms
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter') return;
  const authModal = document.getElementById('authModal');
  if (!authModal?.classList.contains('open')) return;
  const loginVisible = document.getElementById('authFormLogin')?.style.display !== 'none';
  if (loginVisible) _authLogin();
  else _authSignup();
});

// =========================================================
// AUTH STATE LISTENER — central hub for login/logout
// =========================================================
_supa.auth.onAuthStateChange(async (event, session) => {
  _currentUser = session?.user || null;
  _updateUserUI(_currentUser);

  if (_currentUser) {
    // User just logged in — hide auth modal, show app
    document.getElementById('authModal').classList.remove('open');
    document.getElementById('authError').textContent = '';
    // Disable auth inputs so browser stops offering password autofill
    document.querySelectorAll('#authModal input').forEach(el => { el.disabled = true; el.value = ''; });
    // If app not yet initialized, trigger init
    if (!_authAppReady) {
      _authAppReady = true;
      // Map may already be loaded — if so run init now, else wait for map.on('load')
      if (_mapLoadFired) _initAppAfterAuth();
    }
  } else {
    // User logged out — show auth modal, hide user menu
    document.getElementById('authModal').classList.add('open');
    // Re-enable auth inputs
    document.querySelectorAll('#authModal input').forEach(el => { el.disabled = false; });
    // Clear app state
    polygons.forEach(p => { _removeZoneLabel(p); _removeZoneLayers(p.id); });
    polygons = [];
    properties = [];
    sheetConfigs = {};
    sheetConfig = {};
    renderPolygonList();
    document.getElementById('statProps').textContent = '0';
    document.getElementById('statAssigned').textContent = '0';
  }
});

let _authAppReady = false;
let _mapLoadFired = false;


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
//   const { error } = await supabase.from('zones').// =========================================================

// In-memory UI state cache — loaded from Supabase after login
// Allows renderPolygonList() to read UI state synchronously
const _uiStateCache = {};

const DB = {
  // -- Zones ------------------------------------------
  async saveZones(zones) {
    if (!_currentUser) return;
    try {
      await _supa.from('zones').upsert({ user_id: _currentUser.id, data: zones, updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
    } catch(e) { console.warn('DB.saveZones error:', e); }
  },

  async loadZones() {
    if (!_currentUser) return null;
    try {
      const { data } = await _supa.from('zones').select('data').eq('user_id', _currentUser.id).maybeSingle();
      return data?.data || null;
    } catch(e) { return null; }
  },

  // -- Sheet Configs -----------------------------------
  async saveSheetConfigs(configs) {
    if (!_currentUser) return;
    try {
      await _supa.from('sheet_configs').upsert({ user_id: _currentUser.id, configs, updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
    } catch(e) { console.warn('DB.saveSheetConfigs error:', e); }
  },

  async loadSheetConfigs() {
    if (!_currentUser) return null;
    try {
      const { data } = await _supa.from('sheet_configs').select('configs').eq('user_id', _currentUser.id).maybeSingle();
      return data?.configs || null;
    } catch(e) { return null; }
  },

  // -- App State ---------------------------------------
  async saveAppState(state) {
    if (!_currentUser) return;
    try {
      await _supa.from('app_state').upsert({ user_id: _currentUser.id, state: state.state || '', county: state.county || '', updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
    } catch(e) { console.warn('DB.saveAppState error:', e); }
  },

  async loadAppState() {
    if (!_currentUser) return null;
    try {
      const { data } = await _supa.from('app_state').select('state,county').eq('user_id', _currentUser.id).maybeSingle();
      return data ? { state: data.state, county: data.county } : null;
    } catch(e) { return null; }
  },

  // -- County List Cache (stays in localStorage — pure UI cache) ──
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
  // saveUIState: writes to in-memory cache immediately, syncs to Supabase in background
  saveUIState(key, value) {
    _uiStateCache[key] = value; // instant synchronous update
    if (!_currentUser) return;
    // fire-and-forget to Supabase
    _supa.from('ui_state').upsert({ user_id: _currentUser.id, key, value, updated_at: new Date().toISOString() }, { onConflict: 'user_id,key' })
      .then(() => {}).catch(e => console.warn('DB.saveUIState error:', e));
  },

  // loadUIState: reads from in-memory cache synchronously (populated after login)
  loadUIState(key, fallback = null) {
    return key in _uiStateCache ? _uiStateCache[key] : fallback;
  },

  // loadAllUIState: fetches all UI state from Supabase and populates cache
  // Called once after login in _initAppAfterAuth
  async loadAllUIState() {
    if (!_currentUser) return;
    try {
      const { data } = await _supa.from('ui_state').select('key,value').eq('user_id', _currentUser.id);
      if (data) data.forEach(row => { _uiStateCache[row.key] = row.value; });
    } catch(e) { console.warn('DB.loadAllUIState error:', e); }
  },

  // -- Unassigned Zone Pricing -------------------------
  async saveUnassigned(entries) {
    if (!_currentUser) return;
    try {
      await _supa.from('unassigned_zones').upsert({ user_id: _currentUser.id, data: entries, updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
    } catch(e) { console.warn('DB.saveUnassigned error:', e); }
  },

  async loadUnassigned() {
    if (!_currentUser) return [];
    try {
      const { data } = await _supa.from('unassigned_zones').select('data').eq('user_id', _currentUser.id).maybeSingle();
      return data?.data || [];
    } catch(e) { return []; }
  },
};

// =========================================================
// FIXED TOOLTIP HELPER
// Used for sidebar elements inside overflow:hidden containers
// where position:absolute tooltips get clipped.
// =========================================================
const _ftip = {
  _el: null,
  show(text, triggerEl, align = 'center') {
    if (!this._el) this._el = document.getElementById('fixedTip');
    if (!this._el) return;
    this._el.textContent = text;
    this._el.style.visibility = 'hidden';
    this._el.style.opacity = '0';
    this._el.style.left = '0px';
    this._el.style.top = '0px';
    this._el.classList.add('show');
    requestAnimationFrame(() => {
      const r = triggerEl.getBoundingClientRect();
      const tw = this._el.offsetWidth;
      const th = this._el.offsetHeight;
      let left;
      if (align === 'left') {
        // Right-align tooltip to trigger (tooltip ends at trigger's right edge)
        left = Math.max(4, r.right - tw);
      } else {
        left = Math.max(4, Math.min(r.left + r.width / 2 - tw / 2, window.innerWidth - tw - 4));
      }
      const top = r.top - th - 10;
      const arrowLeft = Math.round((r.left + r.width / 2) - left);
      this._el.style.left = left + 'px';
      this._el.style.top = top + 'px';
      this._el.style.setProperty('--arrow-left', arrowLeft + 'px');
      this._el.style.visibility = '';
      this._el.style.opacity = '';
    });
  },
  hide() {
    if (!this._el) this._el = document.getElementById('fixedTip');
    if (!this._el) return;
    this._el.classList.remove('show');
    this._el.style.visibility = '';
    this._el.style.opacity = '';
  }
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
const SERVICE_ACCOUNT_EMAIL = [108,97,110,100,118,97,108,117,97,116,111,114,45,115,104,101,101,116,115,64,108,97,110,100,118,97,108,117,97,116,111,114,46,105,97,109,46,103,115,101,114,118,105,99,101,97,99,99,111,117,110,116,46,99,111,109].map(c=>String.fromCharCode(c)).join('');
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
let _lastEscapeTime = 0;

// Zone desc modal
let _editingDescId = null;

// Mapbox draw layer IDs
const SRC_FILL    = '__draw_fill';
const SRC_LINE    = '__draw_line';
const SRC_PREVIEW = '__draw_preview';
const SRC_VERTS   = '__draw_verts';
const SRC_PINS    = '__property_pins';
const LAYER_PINS  = '__property_pins_layer';
let   _pinsVisible = false;

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
  { id: 'mapbox://styles/mapbox/streets-v12',           label: 'Streets'   },
  { id: 'mapbox://styles/mapbox/outdoors-v12',          label: 'Outdoors'  },
  { id: 'mapbox://styles/mapbox/satellite-v9',          label: 'Satellite' },
  { id: 'mapbox://styles/mapbox/satellite-streets-v12', label: 'Hybrid'    },
];
const HYBRID_IDX = 3;
(function buildStyleSelect() {
  // Build style dropdown
  let _activeStyleIdx = HYBRID_IDX;
  const _styleDropdown = document.getElementById('styleDropdown');
  const _stylePicker = document.getElementById('stylePicker');

  const STYLE_ICONS = [
    `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>`,
    `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 17l4-8 4 4 4-6 4 10"/></svg>`,
    `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`,
    `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`,
  ];

  function _buildStyleDropdown() {
    _styleDropdown.innerHTML = '';
    MAP_STYLES.forEach((s, i) => {
      const btn = document.createElement('button');
      btn.className = 'style-option' + (i === _activeStyleIdx ? ' active' : '');
      btn.innerHTML = STYLE_ICONS[i] + ' ' + s.label + (i === _activeStyleIdx ? ' ✓' : '');
      btn.onclick = (e) => {
        e.stopPropagation();
        _activeStyleIdx = i;
        changeMapStyle(i);
        _styleDropdown.classList.remove('open');
        _buildStyleDropdown();
      };
      _styleDropdown.appendChild(btn);
    });
  }
  _buildStyleDropdown();

  _stylePicker.addEventListener('click', (e) => {
    e.stopPropagation();
    _styleDropdown.classList.toggle('open');
  });
  document.addEventListener('click', () => _styleDropdown.classList.remove('open'));
})();

// =========================================================
// MAP INIT
// =========================================================
const map = new mapboxgl.Map({
  container: 'map',
  style: MAP_STYLES[HYBRID_IDX].id,
  center: [-95.21117808929444, 38.70748954884343],
  zoom: 4.575130398788689,
  projection: 'mercator',
  doubleClickZoom: false,
  boxZoom: false,
});
map.addControl(new mapboxgl.NavigationControl(), 'bottom-right');
map.addControl(new mapboxgl.ScaleControl({ maxWidth: 120, unit: 'imperial' }), 'bottom-left');

// #10 — remove native browser tooltips from all map control buttons
map.once('load', () => {
  document.querySelectorAll('.mapboxgl-ctrl button[title]').forEach(btn => btn.removeAttribute('title'));
  document.querySelectorAll('.mapboxgl-ctrl button[aria-label]').forEach(btn => btn.removeAttribute('aria-label'));
});

// Custom North button control — sits above NavigationControl in bottom-right
class NorthControl {
  onAdd(map) {
    this._map = map;
    this._container = document.createElement('div');
    this._container.className = 'mapboxgl-ctrl mapboxgl-ctrl-group';
    this._btn = document.createElement('button');
    this._btn.className = 'north-ctrl-btn';
    this._btn.title = '';
    this._btn.innerHTML = `<svg width="18" height="18" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"><path d="M47.655 1.634l-35 95c-.828 2.24 1.659 4.255 3.68 2.98l33.667-21.228l33.666 21.228c2.02 1.271 4.503-.74 3.678-2.98l-35-95C51.907.514 51.163.006 50 .008c-1.163.001-1.99.65-2.345 1.626zm-.155 14.88v57.54L19.89 91.461z" fill="#b94040" fill-rule="evenodd"/></svg>`;
    this._btn.onclick = () => map.easeTo({ bearing: 0, pitch: 0, duration: 500 });
    this._container.appendChild(this._btn);
    return this._container;
  }
  onRemove() { this._container.parentNode.removeChild(this._container); this._map = undefined; }
}
map.addControl(new NorthControl(), 'bottom-right');


function changeMapStyle(idx) {
  const center  = map.getCenter();
  const zoom    = map.getZoom();
  const bearing = map.getBearing();
  const pitch   = map.getPitch();
  polygons.forEach(p => _removeZoneLabel(p));
  map.setStyle(MAP_STYLES[parseInt(idx)].id);
  map.once('style.load', () => {
    map.jumpTo({ center, zoom, bearing, pitch });
  });
}

// 1.2 — track whether map has fully initialised (style.load fires on init AND style changes)
// County boundary layers should only re-draw on style change, not on first page load
let _mapInitComplete = false;

map.on('style.load', () => {
  _initDrawLayers();
  _initPinLayer();
  if (_pinsVisible) _rebuildPins();
  // 1.2 — on initial page load, skip zone polygon fills/lines and county boundaries
  // They only appear after the user actively selects a county.
  // On style switch (_mapInitComplete=true), restore everything normally.
  if (_mapInitComplete) {
    _restoreAllZoneLayers();
    if (_pendingCountyGeoJSON) _readdCountyLayer(_pendingCountyGeoJSON);
  }
  polygons.forEach(p => _addZoneLabel(p));
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
// PROPERTY PIN LAYER
// =========================================================
let _pinPopup = null;

function _initPinLayer() {
  if (!map.getSource(SRC_PINS)) {
    map.addSource(SRC_PINS, { type: 'geojson', data: { type: 'FeatureCollection', features: [] } });
  }
  if (!map.getLayer(LAYER_PINS)) {
    map.addLayer({
      id: LAYER_PINS,
      type: 'circle',
      source: SRC_PINS,
      paint: {
        'circle-radius': 5,
        'circle-color': '#4a90d9',
        'circle-stroke-color': '#ffffff',
        'circle-stroke-width': 2,
        'circle-opacity': _pinsVisible ? 1 : 0,
        'circle-stroke-opacity': _pinsVisible ? 1 : 0,
      },
    });

    map.on('click', LAYER_PINS, (e) => {
      if (!e.features.length) return;
      const f = e.features[0];
      const p = f.properties;
      const coords = e.features[0].geometry.coordinates.slice();

      if (_pinPopup) { _pinPopup.remove(); _pinPopup = null; }

      const zoneLabel = (p.zone && p.zone !== 'null') ? `Zone ${p.zone}` : 'Unassigned';
      const acreage   = p.acreage   ? `${p.acreage} ac`   : '—';
      const liAcreage = p.liAcreage ? `${p.liAcreage} ac` : '—';
      const ownerDisplay = p.ownerName ? (p.ownerName.length > 40 ? p.ownerName.slice(0, 40) + '…' : p.ownerName) : null;
      const ownerRow  = ownerDisplay ? `<div style="font-size:11px;color:#6b7d95;border-bottom:1px solid #eee;padding-bottom:3px;margin-bottom:3px;display:flex;justify-content:space-between"><span>Owner</span><span style="color:#1a2332;font-weight:500;text-align:right;margin-left:8px">${ownerDisplay}</span></div>` : '';
      const linkHtml  = p.parcelLink
        ? `<a href="${p.parcelLink}" target="_blank" rel="noopener" style="display:block;margin-top:9px;text-align:center;font-size:11px;font-weight:700;color:#5b7fa6;background:#edf2f8;border-radius:6px;padding:5px 0;text-decoration:none;letter-spacing:0.03em;">View property page ↗</a>`
        : '';

      const html = `
        <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;min-width:240px;max-width:280px;background:#ffffff;">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px">
            <span style="font-size:11px;color:#6b7d95;font-weight:500;min-width:0">APN: <span style="color:#1a2332;font-weight:700">${p.apn || '—'}</span></span>
            <span style="font-size:10px;font-weight:700;background:#edf2f8;color:#2c5282;border-radius:4px;padding:2px 8px;flex-shrink:0;white-space:nowrap;margin-right:20px">${zoneLabel}</span>
          </div>
          ${ownerRow}
          <div style="font-size:11px;color:#6b7d95;border-bottom:1px solid #eee;padding-bottom:3px;margin-bottom:3px;display:flex;justify-content:space-between"><span>County</span><span style="color:#1a2332;font-weight:500">${p.county ? p.county + ', ' + p.state : '—'}</span></div>
          <div style="font-size:11px;color:#6b7d95;border-bottom:1px solid #eee;padding-bottom:3px;margin-bottom:3px;display:flex;justify-content:space-between"><span>Acreage</span><span style="color:#1a2332;font-weight:500">${acreage}</span></div>
          <div style="font-size:11px;color:#6b7d95;padding-bottom:3px;display:flex;justify-content:space-between"><span>Calc. acreage</span><span style="color:#1a2332;font-weight:500">${liAcreage}</span></div>
          ${linkHtml}
        </div>`;

      _pinPopup = new mapboxgl.Popup({ offset: 14, closeButton: true, maxWidth: '300px' })
        .setLngLat(coords)
        .setHTML(html)
        .addTo(map);
    });

    map.on('mouseenter', LAYER_PINS, () => { if (_pinsVisible) map.getCanvas().style.cursor = 'pointer'; });
    map.on('mouseleave', LAYER_PINS, () => { if (!drawMode) map.getCanvas().style.cursor = ''; });
  }
}

function _rebuildPins() {
  if (!map.getSource(SRC_PINS)) return;
  const features = properties.map(p => ({
    type: 'Feature',
    geometry: { type: 'Point', coordinates: [p.lng, p.lat] },
    properties: {
      apn:        p.apn || '',
      county:     p.county || '',
      state:      p.state || '',
      acreage:    p.acreage || '',
      liAcreage:  p.liAcreage || '',
      parcelLink: p.parcelLink || '',
      ownerName:  p.ownerName || '',
      zone:       p.zone || null,
    },
  }));
  map.getSource(SRC_PINS).setData({ type: 'FeatureCollection', features });
}

function _togglePins(force) {
  _pinsVisible = (force !== undefined) ? force : !_pinsVisible;
  const el = document.getElementById('pinToggle');
  if (el) el.classList.toggle('on', _pinsVisible);
  if (!map.getLayer(LAYER_PINS)) return;
  const opacity = _pinsVisible ? 1 : 0;
  map.setPaintProperty(LAYER_PINS, 'circle-opacity', opacity);
  map.setPaintProperty(LAYER_PINS, 'circle-stroke-opacity', opacity);
  if (!_pinsVisible && _pinPopup) { _pinPopup.remove(); _pinPopup = null; }
  if (_pinsVisible) _rebuildPins();
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
function _restoreAllZoneLayers() { polygons.forEach(p => { if (!p._isUnassigned) _addZoneLayers(p); }); }

// =========================================================
// ZONE LABELS
// Hierarchy: zoomed-out → county pill → zoom in → zone labels → click → notes
// =========================================================

// Zoom threshold: below this = show county pills, at/above = show zone labels
const COUNTY_PILL_ZOOM = 7;

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
  // Tooltip on zone label — use title attr to avoid disrupting Mapbox marker layout
  el.title = `Open pricing panel for Zone ${poly.letter||''}`;
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
    // Use county boundary centroid if available, else fall back to zone centroids
    let lng, lat;
    const cachedGeoJSON = _countyGeoJSONCache && _countyGeoJSONCache[key];

    if (cachedGeoJSON && cachedGeoJSON.features && cachedGeoJSON.features.length) {
      const bounds = new mapboxgl.LngLatBounds();
      const extendB = (coords) => {
        if (!Array.isArray(coords)) return;
        if (typeof coords[0] === 'number') { bounds.extend(coords); return; }
        coords.forEach(c => extendB(c));
      };
      cachedGeoJSON.features.forEach(f => extendB(f.geometry.coordinates));
      if (!bounds.isEmpty()) {
        const ctr = bounds.getCenter();
        lng = ctr.lng; lat = ctr.lat;
      }
    }
    if (lng === undefined) {
      const centers = polys.map(_polyCenter);
      lng = centers.reduce((s,c) => s+c[0], 0) / centers.length;
      lat = centers.reduce((s,c) => s+c[1], 0) / centers.length;
    }

    const county = polys[0].countyName || 'County';
    const st = polys[0].stateAbbr || '';
    const count = polys.length;

    const el = document.createElement('div');
    el.className = 'zone-cluster';
    el.innerHTML = `${county} County, ${st}&nbsp;<span class="zc-count" style="pointer-events:none">${count}</span>`;
    el.title = `Click to zoom into ${county} County`;

    // Single click on county pill = zoom in AND update dropdowns
    el.addEventListener('click', async (e) => {
      e.stopPropagation();
      e.preventDefault();
      window._countyClickActive = true;
      setTimeout(() => { window._countyClickActive = false; }, 400);
      const sa = polys[0].stateAbbr, cn = polys[0].countyName;

      // Update dropdowns immediately
      stateSelect.value = sa;
      _syncStateTrigger(sa);
      const cs = document.getElementById('countySelect');
      await loadCounties(true);
      cs.value = cn;
      if (cs.value !== cn) {
        const o = document.createElement('option');
        o.value = cn; o.textContent = cn + ' County';
        cs.appendChild(o);
        cs.value = cn;
      }
      _syncCountyTrigger(cn);
      saveAppState();
      const saved = _getSheetConfig(sa, cn);
      if (saved) { sheetConfig = saved; setConnected(true); }
      else { sheetConfig = null; setConnected(false); }
      renderPolygonList();

      // Zoom to county bounds from GeoJSON cache if available, else use zone bounds
      const _ck = _countyKey(sa, cn);
      const _cg = _countyGeoJSONCache && _countyGeoJSONCache[_ck];
      const b = new mapboxgl.LngLatBounds();
      if (_cg && _cg.features && _cg.features.length) {
        _cg.features.forEach(f => {
          const coords = f.geometry.type==='Polygon' ? f.geometry.coordinates.flat()
                       : f.geometry.type==='MultiPolygon' ? f.geometry.coordinates.flat(2) : [];
          coords.forEach(c => b.extend(c));
        });
      } else {
        polys.forEach(p => p.points.forEach(pt => b.extend(pt)));
      }
      // 1.1 — clear state boundary when county pill clicked on map
      if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
      if (map.getSource('state-boundary')) map.removeSource('state-boundary');
      map.fitBounds(b, { padding: 60 });
      map.once('moveend', () => { _restoreAllZoneLayers(); loadCountyBoundaryOnly(sa, cn); });
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

  // Zone fill/line layers: visible when zoomed in only
  polygons.forEach(p => {
    if (p._isUnassigned) return;
    const fid = _fillId(p.id), lid = _lineId(p.id);
    const vis = zoomed ? 'visible' : 'none';
    if (map.getLayer(fid)) map.setLayoutProperty(fid, 'visibility', vis);
    if (map.getLayer(lid)) map.setLayoutProperty(lid, 'visibility', vis);
  });

  // County boundaries: show when zoomed in, hide when zoomed out
  // Exception: active (selected) county always stays visible at any zoom
  const activeSA = (document.getElementById('stateSelect') || {}).value || '';
  const activeCN = (document.getElementById('countySelect') || {}).value || '';
  const activeKey = _countyKey(activeSA, activeCN);
  Object.entries(_countyLayers).forEach(([key, sid]) => {
    const visible = zoomed || key === activeKey;
    const vis = visible ? 'visible' : 'none';
    if (map.getLayer(sid+'-fill')) map.setLayoutProperty(sid+'-fill', 'visibility', vis);
    if (map.getLayer(sid+'-line')) map.setLayoutProperty(sid+'-line', 'visibility', vis);
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
}

function undoLastDrawPoint() {
  if (!drawPoints.length) { cancelDraw(); return; }
  drawPoints.pop();
  if (!drawPoints.length) {
    cancelDraw();
    return;
  }
  _refreshPolyPreviews(null);
  const hint = document.getElementById('drawHint');
  if (hint) hint.textContent = `📍 ${drawPoints.length} point${drawPoints.length !== 1 ? 's' : ''} — Esc to undo, double-Esc to cancel`;
}

function startDraw() {
  cancelDraw();
  drawMode = 'polygon'; polyState = 'drawing';
  map.getCanvas().style.cursor = 'crosshair';
  document.getElementById('btnPolygon').classList.add('active');
  document.getElementById('btnCancel').style.display = 'block';
  const hint = document.getElementById('drawHint');
  hint.style.display = 'block';
  hint.textContent = '📍 Click to add vertices — Esc to undo, double-Esc to cancel, click first point to close';
}

map.on('click', function(e) {
  if (window.innerWidth <= 700) document.getElementById('sidebar').classList.remove('open');
  if (drawMode !== 'polygon' || polyState !== 'drawing') return;
  const pt = [e.lngLat.lng, e.lngLat.lat];
  if (drawPoints.length >= 3 && pixelDist(pt, drawPoints[0]) <= 10) { _finishPolygon(); return; }
  // #1 — enforce county boundary on every vertex placement
  if (_pendingCountyGeoJSON) {
    const inCounty = _pendingCountyGeoJSON.features.some(f => {
      const coords = f.geometry.type === 'Polygon' ? [f.geometry.coordinates]
                   : f.geometry.type === 'MultiPolygon' ? f.geometry.coordinates : [];
      return coords.some(poly => pointInPolygon(pt[1], pt[0], poly[0].map(c => [c[0], c[1]])));
    });
    if (!inCounty) {
      showToast('Zone must be drawn within the selected county boundary.', 'error');
      return;
    }
  }
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

  // Overlap detection — sample interior points of new polygon against existing zones in same county
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  const _cnNorm = (cn || '').toLowerCase().trim();
  const existingPolys = polygons.filter(p =>
    p.stateAbbr === sa && (p.countyName || '').toLowerCase().trim() === _cnNorm && p.points && p.points.length >= 3 && !p._isUnassigned
  );
  if (existingPolys.length) {
    // Sample a grid of ~25 interior points from the new polygon's bounding box
    const lngs = pts.map(p => p[0]), lats = pts.map(p => p[1]);
    const minLng = Math.min(...lngs), maxLng = Math.max(...lngs);
    const minLat = Math.min(...lats), maxLat = Math.max(...lats);
    const steps = 5;
    const newPtsAsLngLat = pts; // already [lng, lat]
    let overlappingZone = null;
    outer: for (let i = 0; i <= steps; i++) {
      for (let j = 0; j <= steps; j++) {
        const sLng = minLng + (maxLng - minLng) * (i / steps);
        const sLat = minLat + (maxLat - minLat) * (j / steps);
        // Check sample point is inside new polygon
        if (!pointInPolygon(sLat, sLng, newPtsAsLngLat)) continue;
        // Check if it's inside any existing polygon
        for (const ep of existingPolys) {
          if (pointInPolygon(sLat, sLng, ep.points)) {
            overlappingZone = ep;
            break outer;
          }
        }
      }
    }
    if (overlappingZone) {
      showToast(`Zone overlaps with Zone ${overlappingZone.letter} — adjust your polygon to avoid overlap`, 'error');
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
    polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName === countyName && !p._isUnassigned).map(p => p.letter).filter(Boolean)
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
  document.getElementById('zeBadge').textContent = p._isUnassigned ? 'UNASSIGNED PRICING PANEL' : `ZONE ${p.letter} PRICING PANEL`;
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
  // Fetches live spreadsheet title and updates sheetConfig + modal connected box
  try {
    const r = await fetch('/.netlify/functions/sheets-read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetId, sheetName: 'LI Raw Dataset', metaOnly: true }),
    });
    const data = await r.json();
    const name = data.spreadsheetTitle || data.sheetTitle || '';
    if (!name) return;
    // Update the stored config so renames persist across sessions
    const sa = document.getElementById('stateSelect').value;
    const cn = document.getElementById('countySelect').value;
    if (sa && cn) {
      const cfg = _getSheetConfig(sa, cn);
      if (cfg && cfg.sheetId === sheetId) {
        cfg.sheetTitle = name;
        _setSheetConfig(sa, cn, cfg);
        if (sheetConfig && sheetConfig.sheetId === sheetId) sheetConfig.sheetTitle = name;
      }
    }
    // Update the connected status box title if modal is open
    const titleEl = document.getElementById('smStatusTitle');
    if (titleEl) titleEl.textContent = name;
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
    _syncStateTrigger(stateAbbr);
    loadCounties(true).then(() => {
      cs.value = countyName;
      _syncCountyTrigger(countyName);
      openSheetsModal();
    });
  } else {
    cs.value = countyName;
    _syncCountyTrigger(countyName);
    openSheetsModal();
  }
}

// -- Share a single county via short URL --
function shareCounty(stateAbbr, countyName, e) {
  if (e) e.stopPropagation();
  // Strip " County" suffix if present, to match the format LandValuator expects
  const cleanCounty = (countyName || '').replace(/\s+county$/i, '').trim();
  const url = `${window.location.origin}${window.location.pathname}?state=${encodeURIComponent(stateAbbr)}&county=${encodeURIComponent(cleanCounty)}`;
  navigator.clipboard.writeText(url)
    .then(() => showToast('County link copied!', 'success'))
    .catch(() => {
      const b = document.getElementById('shareBanner');
      b.textContent = '🔗 ' + url; b.style.display = 'block';
      setTimeout(() => b.style.display = 'none', 12000);
    });
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
// ── Acreage overlap validation (item 4) ─────────────────────────────────────
// Returns array of { rowA, rowB, overlapMin, overlapMax } for any conflicting pairs.
// Rows with blank min/max are skipped — incomplete rows don't block save.
function _checkAcreageOverlaps(rows) {
  const conflicts = [];
  const valid = rows.map((r, i) => ({
    idx: i,
    min: r.minAcres !== '' ? parseFloat(r.minAcres) : null,
    max: r.maxAcres !== '' ? parseFloat(r.maxAcres) : null,
  })).filter(r => r.min !== null && r.max !== null && !isNaN(r.min) && !isNaN(r.max));

  for (let i = 0; i < valid.length; i++) {
    for (let j = i + 1; j < valid.length; j++) {
      const a = valid[i], b = valid[j];
      const overlapMin = Math.max(a.min, b.min);
      const overlapMax = Math.min(a.max, b.max);
      if (overlapMin < overlapMax) {
        conflicts.push({ rowA: a.idx, rowB: b.idx, overlapMin, overlapMax });
      }
    }
  }
  return conflicts;
}

// =========================================================
// COUNTY BOUNDARY VALIDATION
// =========================================================
async function _validatePropertiesInCounty(props, fips, countyName) {
  // Reuse cached GeoJSON if available, else fetch
  const key = `${fips}|${countyName}`;
  let geojson = (_countyGeoJSONCache && _countyGeoJSONCache[key]) || null;
  if (!geojson) {
    try { geojson = await _fetchCountyGeoJSON(fips, countyName); }
    catch(e) { return null; } // network failure — skip validation
  }
  if (!geojson || !geojson.features || !geojson.features.length) return null;

  // Extract all polygon rings from GeoJSON features
  const rings = [];
  geojson.features.forEach(f => {
    const g = f.geometry;
    if (!g) return;
    if (g.type === 'Polygon') {
      g.coordinates.forEach(ring => rings.push(ring));
    } else if (g.type === 'MultiPolygon') {
      g.coordinates.forEach(poly => poly.forEach(ring => rings.push(ring)));
    }
  });
  if (!rings.length) return null;

  // GeoJSON rings are [lng, lat] — convert for pointInPolygon which takes (lat, lng, pts as [lng,lat])
  const inAnyRing = (lat, lng) => rings.some(ring => pointInPolygon(lat, lng, ring));

  const outsideProps = [];
  props.forEach(prop => {
    if (!inAnyRing(prop.lat, prop.lng)) {
      outsideProps.push({ apn: prop.apn || '(no APN)', lat: prop.lat, lng: prop.lng });
    }
  });

  // Always log to console for inspection
  if (outsideProps.length) {
    console.group(`[LandValuator] ${outsideProps.length} properties outside ${countyName} County boundary`);
    console.table(outsideProps);
    console.groupEnd();
  }

  return { outsideProps, total: props.length };
}

function _showBoundaryModal({ title, sub, icon = '🚫', outsideProps, onProceed = null }) {
  document.getElementById('boundaryModalIcon').textContent = icon;
  document.getElementById('boundaryModalTitle').textContent = title;
  document.getElementById('boundaryModalSub').textContent = sub;

  // Build copyable APN / lat / lng list
  const listEl = document.getElementById('boundaryModalListWrap');
  const header = 'APN                          Latitude       Longitude';
  const divider = '─'.repeat(55);
  const rows = outsideProps.map(p => {
    const apn = (p.apn || '').padEnd(30);
    const lat = String(p.lat).padEnd(15);
    const lng = String(p.lng);
    return `${apn}${lat}${lng}`;
  }).join('\n');
  listEl.textContent = `${header}\n${divider}\n${rows}`;

  const cancelBtn   = document.getElementById('boundaryModalCancelBtn');
  const proceedBtn  = document.getElementById('boundaryModalProceedBtn');
  const dismissBtn  = document.getElementById('boundaryModalDismissBtn');

  if (onProceed) {
    // Warning mode — show Proceed Anyway + Cancel
    cancelBtn.style.display  = '';
    proceedBtn.style.display = '';
    dismissBtn.style.display = 'none';
  } else {
    // Block mode — dismiss only
    cancelBtn.style.display  = 'none';
    proceedBtn.style.display = 'none';
    dismissBtn.style.display = '';
  }

  // Wire buttons (clone to remove stale listeners)
  const newProceed = proceedBtn.cloneNode(true);
  const newCancel  = cancelBtn.cloneNode(true);
  const newDismiss = dismissBtn.cloneNode(true);
  proceedBtn.replaceWith(newProceed);
  cancelBtn.replaceWith(newCancel);
  dismissBtn.replaceWith(newDismiss);

  const close = () => document.getElementById('boundaryModal').classList.remove('open');

  if (onProceed) {
    document.getElementById('boundaryModalProceedBtn').addEventListener('click', () => { close(); onProceed(); });
    document.getElementById('boundaryModalCancelBtn').addEventListener('click', close);
  } else {
    document.getElementById('boundaryModalDismissBtn').addEventListener('click', close);
  }

  // Overlay click closes
  document.getElementById('boundaryModal').addEventListener('click', function _oc(e) {
    if (e.target === e.currentTarget) { close(); this.removeEventListener('click', _oc); }
  });

  document.getElementById('boundaryModal').classList.add('open');
}

function _showOverlapError(conflicts) {
  // Highlight conflicting rows in red
  const rows = document.querySelectorAll('#zeTbody tr');
  rows.forEach(tr => tr.style.background = '');
  const conflictIdxs = new Set();
  conflicts.forEach(c => { conflictIdxs.add(c.rowA); conflictIdxs.add(c.rowB); });
  conflictIdxs.forEach(i => { if (rows[i]) rows[i].style.background = '#fff5f5'; });

  // Build detail lines for the modal
  const details = conflicts.map(c => {
    const a = c.rowA + 1, b = c.rowB + 1;
    return `Row ${a} and Row ${b} overlap between ${c.overlapMin}–${c.overlapMax} acres`;
  }).join('\n');

  _showConfirm({
    title: 'Acreage Range Overlap Detected',
    sub: 'Please adjust the ranges so they don\'t overlap before saving.\n\n' + details,
    okLabel: 'Go Back & Fix',
    _overlapOnly: true,
  });
}

async function saveAndSyncZone() {
  if (!_editingDescId) return;
  const p = polygons.find(p => p.id === _editingDescId);
  if (!p) return;

  const btn = document.getElementById('zeSaveBtn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Saving...'; }

  // 1. Save locally — capture state BEFORE closing modal
  p.description = document.getElementById('zeNotes').value.trim();
  p.pricingTiers = zeCollectRows();

  // 4 — Acreage overlap validation: block save if any ranges conflict
  const _overlapConflicts = _checkAcreageOverlaps(p.pricingTiers);
  if (_overlapConflicts.length) {
    if (btn) { btn.disabled = false; btn.innerHTML = '💾 Save &amp; Sync'; }
    _showOverlapError(_overlapConflicts);
    return;
  }
  p.allZones = false;
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  persistZones();

  const cfg = (sa && cn) ? (_getSheetConfig(sa, cn) || sheetConfig) : sheetConfig;

  if (!cfg || !cfg.sheetId) {
    closeZoneEditor();
    showToast(p._isUnassigned ? 'Unassigned Zone saved locally' : `Zone ${p.letter} saved locally (no sheet connected)`, 'success');
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
        body: JSON.stringify({ sheetId: cfg.sheetId, sheetName: cfg.sheetName || 'LI Raw Dataset', colCounty: cfg.colCounty || 'County', colAPN: cfg.colAPN || 'APN' }),
      });
      const pd = await pr.json();
      if (pd.properties && pd.properties.length) {
        // Pass county explicitly so filter uses captured cn, not live dropdown
        loadPropertiesFromFunction(pd.properties, cn, pd.scrubbedApns, pd.ownerMap);
        document.getElementById('statProps').textContent = properties.length;
      }
    } catch(e) { console.warn('Could not prefetch properties:', e); }
  }

  // Now safe to close the modal — state/county captured above
  closeZoneEditor();
  showToast(p._isUnassigned ? 'Unassigned Zone saved — syncing...' : `Zone ${p.letter} saved — syncing...`, 'info');

  try {
    // 2. Run zone assignment — scoped to THIS county's polygons only (case-insensitive)
    const _cnNorm = cn.toLowerCase().trim();
    const countyPolygons = polygons.filter(poly => poly.stateAbbr === sa && (poly.countyName||'').toLowerCase().trim() === _cnNorm && !poly._isUnassigned);
    const assignments = [];
    let assigned = 0;
    properties.forEach(prop => {
      prop.zone = null;
      for (const poly of countyPolygons) {
        if (pointInPolygon(prop.lat, prop.lng, poly.points)) {
          prop.zone = poly.letter;
          poly.propCount = (poly.propCount || 0) + 1;
          assigned++;
          if (prop.apn) assignments.push({ apn: prop.apn, zone: poly.letter });
          break;
        }
      }
      // Write UNASSIGNED for properties with no zone
      if (!prop.zone && prop.apn) assignments.push({ apn: prop.apn, zone: 'UNASSIGNED' });
    });
    document.getElementById('statAssigned').textContent = assigned;
    if (_pinsVisible) _rebuildPins();
    // Persist assigned count so it survives refresh
    countyPolygons.forEach(poly => {
      const matched = assignments.filter(a => a.zone === poly.letter).length;
      poly.propCount = matched;
    });
    persistZones();
    renderPolygonList();

    // 3. Write zone letters to sheet
    if (assignments.length) {
      const zr = await fetch('/.netlify/functions/sheets-write-zones', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId: cfg.sheetId, sheetName: 'Scrubbed and Priced', colAPN: cfg.colAPN || 'APN', assignments }),
      });
      const zd = await zr.json();
      if (!zr.ok) throw new Error(zd.error || 'Zone write failed');
    }

    // 4. Sync all pricing tiers (include Unassigned virtual polygon)
    const countyPolys = polygons.filter(poly => !poly.stateAbbr || (poly.stateAbbr === sa && poly.countyName === cn));
    const allTiers = [];
    const _sortZone = p => p._isUnassigned ? 'ZZZZZ' : (p.letter || '');
    countyPolys.slice().sort((a,b) => _sortZone(a).localeCompare(_sortZone(b))).forEach(poly => {
      const zoneLabel = poly._isUnassigned ? 'UNASSIGNED' : poly.letter;
      (poly.pricingTiers || [])
        .filter(t => t.pricePerAcre !== '' && t.pricePerAcre !== undefined && t.pricePerAcre !== null)
        .sort((a,b) => parseFloat(a.minAcres||0) - parseFloat(b.minAcres||0))
        .forEach(t => allTiers.push({ zone: zoneLabel, minAcres: t.minAcres, maxAcres: t.maxAcres, pricePerAcre: t.pricePerAcre }));
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

    showToast(p._isUnassigned ? `Unassigned Zone saved — pricing synced ✓` : `Zone ${p.letter} saved — ${assigned} assigned, pricing synced ✓`, 'success');
  } catch(err) {
    showToast('Sync error: ' + err.message, 'error');
  }
  if (btn) { btn.disabled = false; btn.innerHTML = '💾 Save &amp; Sync'; }
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
  const _realPolyCount = polygons.filter(p => !p._isUnassigned).length;
  document.getElementById('zoneCount').textContent = _realPolyCount;
  document.getElementById('statPolygons').textContent = _realPolyCount;
  const stateSet = new Set(polygons.map(p => p.stateAbbr).filter(Boolean));
  document.getElementById('statStates').textContent = stateSet.size;

  const list = document.getElementById('polygonsList');
  if (!polygons.length) {
    list.innerHTML = '<div class="empty-state">No zones yet.<br>Select a state &amp; county,<br>then draw on the map.</div>';
    return;
  }

  // Group by stateAbbr → countyName (exclude virtual unassigned entries)
  const byState = {};
  polygons.forEach(p => {
    if (p._isUnassigned) return;
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
    const _stateZoneTip = totalZones === 1 ? `1 zone in ${fullName}` : `${totalZones} zones in ${fullName}`;
    hdr.innerHTML = `<span class="state-arrow-zone"><svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 2L8 6L4 10" stroke="#a8bcd4" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg></span><span class="sg-name">${fullName}</span><span class="sg-count">${totalZones}</span>`;
    // Wire fixed tooltips — avoids overflow:hidden clipping in sidebar
    hdr.querySelector('.sg-name').addEventListener('mouseenter', e => _ftip.show('Zoom map into ' + fullName, e.currentTarget));
    hdr.querySelector('.sg-name').addEventListener('mouseleave', () => _ftip.hide());
    hdr.querySelector('.sg-count').addEventListener('mouseenter', e => _ftip.show(_stateZoneTip, e.currentTarget, 'left'));
    hdr.querySelector('.sg-count').addEventListener('mouseleave', () => _ftip.hide());
    hdr.onclick = e => {
      if (e.target.closest('.state-arrow-zone')) {
        const isOpen = hdr.classList.toggle('open');
        countiesDiv.classList.toggle('ac-collapsed', !isOpen);
        DB.saveUIState(stateOpenKey, isOpen);
      } else if (e.target.closest('.sg-name')) {
        // 2.3 — only zoom when clicking the name text, not the badge or header background
        stateSelect.value = stateAbbr;
        _syncStateTrigger(stateAbbr);
        _syncCountyTrigger('');
        // Load counties so the dropdown remains interactive after state zoom
        loadCounties(true);
        navigateToState(stateAbbr);
      }
      // clicks on .sg-count or header background do nothing
    };

    const countiesDiv = document.createElement('div');
    countiesDiv.className = 'state-counties' + (isStateOpen ? '' : ' ac-collapsed');

    Object.keys(byState[stateAbbr]).sort().forEach(countyName => {
      const cPolys = byState[stateAbbr][countyName];
      const cGroup = document.createElement('div');
      cGroup.className = 'county-group';

      const cCfg = _getSheetConfig(stateAbbr, countyName);
      const isConnected = !!(cCfg && cCfg.sheetId);

      const countyOpenKey = 'county_open_' + stateAbbr + '_' + countyName;
      const isCountyOpen = DB.loadUIState(countyOpenKey, true); // default open

      // Sheet connection icon SVG (green=connected, red=not connected)
      const sheetIconSVG = isConnected
        ? `<svg width="17" height="17" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 2h7l5 5v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1z" stroke="#2e8a5a" stroke-width="1.6" stroke-linejoin="round"/><path d="M11 2v5h5" stroke="#2e8a5a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/><line x1="6" y1="10" x2="14" y2="10" stroke="#2e8a5a" stroke-width="1.4" stroke-linecap="round"/><line x1="6" y1="13" x2="14" y2="13" stroke="#2e8a5a" stroke-width="1.4" stroke-linecap="round"/><line x1="6" y1="16" x2="11" y2="16" stroke="#2e8a5a" stroke-width="1.4" stroke-linecap="round"/></svg>`
        : `<svg width="17" height="17" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 2h7l5 5v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1z" stroke="#b94040" stroke-width="1.6" stroke-linejoin="round"/><path d="M11 2v5h5" stroke="#b94040" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/><line x1="6" y1="10" x2="14" y2="10" stroke="#b94040" stroke-width="1.4" stroke-linecap="round"/><line x1="6" y1="13" x2="14" y2="13" stroke="#b94040" stroke-width="1.4" stroke-linecap="round"/><line x1="6" y1="16" x2="11" y2="16" stroke="#b94040" stroke-width="1.4" stroke-linecap="round"/></svg>`;
      const sheetIconTooltip = isConnected
        ? `Manage sheet connected to ${countyName} County`
        : `Connect a sheet to ${countyName} County`;

      const cHdr = document.createElement('div');
      cHdr.className = 'county-header';
      cHdr.innerHTML = `
        <div class="county-header-pill${isCountyOpen ? ' open' : ''}">
          <span class="county-arrow-zone"><svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 2L8 6L4 10" stroke="#a8bcd4" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg></span>
          <span class="county-name-text">${countyName} County</span>
          <span class="county-zone-pill" id="cpill-${stateAbbr}-${countyName.replace(/\s+/g,'_')}">—</span>
          <span class="tip-wrap"><button class="county-action-btn sheet-icon-btn" onclick="openSheetsModalForCounty('${stateAbbr}','${CSS.escape(countyName)}',event)">${sheetIconSVG}</button><span class="tip-box tip-sidebar">${sheetIconTooltip}</span></span>
          <span class="tip-wrap"><button class="county-action-btn sheet-icon-btn" onclick="shareCounty('${stateAbbr}','${CSS.escape(countyName)}',event)"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="#6b7d95" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg></button><span class="tip-box tip-sidebar">Copy link to open ${countyName} County in LandValuator.</span></span>
          <span class="tip-wrap"><button class="county-action-btn sheet-icon-btn" onclick="deleteCounty('${stateAbbr}','${CSS.escape(countyName)}',event)"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="#6b7d95" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg></button><span class="tip-box tip-sidebar">Delete saved zones in ${countyName} County</span></span>
        </div>
      `;
      // Wire fixed tooltips for county name and count badge
      const _cnPropTotal = properties.filter(p => {
        const pc = (p.county||'').toLowerCase().replace(' county','').trim();
        const cc = countyName.toLowerCase().trim();
        return pc === cc && (p.state||'').toUpperCase() === stateAbbr;
      }).length;
      const _cnZoneTip = _cnPropTotal + ' propert' + (_cnPropTotal === 1 ? 'y' : 'ies') + ' connected to ' + countyName + ' County';
      // Update county pill now that total is computed
      const _pillEl2 = cHdr.querySelector('.county-zone-pill');
      if (_pillEl2) _pillEl2.textContent = _cnPropTotal;
      const _cnNameEl = cHdr.querySelector('.county-name-text');
      const _cnPillEl = cHdr.querySelector('.county-zone-pill');
      if (_cnNameEl) { _cnNameEl.addEventListener('mouseenter', e => _ftip.show('Zoom map into ' + countyName + ' County', e.currentTarget)); _cnNameEl.addEventListener('mouseleave', () => _ftip.hide()); }
      if (_cnPillEl) { _cnPillEl.addEventListener('mouseenter', e => _ftip.show(_cnZoneTip, e.currentTarget, 'left')); _cnPillEl.addEventListener('mouseleave', () => _ftip.hide()); }

      const cContent = document.createElement('div');
      cContent.className = 'county-content' + (isCountyOpen ? '' : ' ac-collapsed');

      cHdr.onclick = e => {
        if (e.target.closest('.county-action-btn')) return;
        if (e.target.closest('.sheet-icon-btn')) return;
        const pill = cHdr.querySelector('.county-header-pill');
        if (e.target.closest('.county-arrow-zone')) {
          const isOpen = pill.classList.toggle('open');
          cContent.classList.toggle('ac-collapsed', !isOpen);
          DB.saveUIState(countyOpenKey, isOpen);
        } else if (e.target.closest('.county-name-text')) {
          // 2.3 — only zoom when clicking the name text, not the badge or pill background
          navigateToCounty(stateAbbr, countyName);
        }
        // clicks on .county-zone-pill or pill background do nothing
      };

      const polyDiv = document.createElement('div');
      polyDiv.className = 'county-polys';

      cPolys.sort((a, b) => (a.letter || '').localeCompare(b.letter || '')).forEach(p => {
        const div = document.createElement('div');
        div.className = 'polygon-item';
        div.innerHTML = `
          <div style="width:10px;height:10px;border-radius:50%;background:${p.color};flex-shrink:0"></div>
          <div class="poly-info">
            <div class="poly-name">ZONE ${p.letter}</div>
            <div class="poly-count">${p.countyName ? p.countyName+' County, '+p.stateAbbr : ''}</div>
          </div>
          <div style="display:flex;align-items:center;gap:4px;flex-shrink:0;margin-left:auto"><span class="tip-wrap"><span class="zone-prop-count tip-anchor" style="cursor:default">${p.propCount||0}</span><span class="tip-box tip-sidebar">${p.propCount||0} propert${(p.propCount||0)===1?'y':'ies'} assigned to Zone ${p.letter}</span></span><span class="tip-wrap"><button class="poly-btn notes-btn" onclick="openZoneDescModal('${p.id}')">⚙</button><span class="tip-box tip-sidebar">Open pricing panel for Zone ${p.letter}</span></span>
          <span class="tip-wrap"><button class="poly-btn delete-btn">✕</button><span class="tip-box tip-sidebar">Delete Zone ${p.letter}</span></span></div>
        `;
        div.querySelector('.notes-btn').addEventListener('click', e => { e.stopPropagation(); openZoneDescModal(p.id); });
        div.querySelector('.delete-btn').addEventListener('click', e => { e.stopPropagation(); deletePoly(p.id); });
        div.onclick = () => zoomToZoneAndCounty(p);
        polyDiv.appendChild(div);
      });

      // Unassigned properties row — shown when properties loaded but some have no zone
      const _unassignedCount = properties.filter(p => {
        const pc = (p.county||'').toLowerCase().replace(' county','').trim();
        const cc = countyName.toLowerCase().trim();
        const sc = (p.state||'').toUpperCase();
        return (pc === cc && sc === stateAbbr) && !p.zone;
      }).length;
      if (_unassignedCount > 0) {
        // Get or create a virtual polygon for unassigned properties pricing
        const _uId = `__unassigned__${stateAbbr}|${countyName}`;
        let _uPoly = polygons.find(p => p.id === _uId);
        if (!_uPoly) {
          _uPoly = {
            id: _uId, letter: '?', name: 'Unassigned', color: '#a0aec0',
            stateAbbr, countyName, points: [], propCount: _unassignedCount,
            pricingTiers: [], description: '', _isUnassigned: true,
          };
          polygons.push(_uPoly);
        } else {
          _uPoly.propCount = _unassignedCount;
        }

        const uDiv = document.createElement('div');
        uDiv.className = 'polygon-item';
        uDiv.style.cssText = 'border-style:dashed;';
        uDiv.innerHTML = `
          <div style="width:10px;height:10px;border-radius:50%;background:#a0aec0;flex-shrink:0;border:2px dashed #718096"></div>
          <div class="poly-info">
            <div class="poly-name" style="color:#718096">UNASSIGNED</div>
            <div class="poly-count">${countyName} County, ${stateAbbr}</div>
          </div>
          <div style="display:flex;align-items:center;gap:4px;flex-shrink:0;margin-left:auto"><span class="tip-wrap"><span class="zone-prop-count tip-anchor" style="background:#e8ebef;color:#718096;cursor:default">${_unassignedCount}</span><span class="tip-box tip-sidebar">${_unassignedCount} unassigned propert${_unassignedCount===1?'y':'ies'} in ${countyName} County</span></span>
          <span class="tip-wrap"><button class="poly-btn notes-btn" data-uid="${_uId}">⚙</button><span class="tip-box tip-sidebar">Open pricing panel for Unassigned properties</span></span>
          <span class="tip-wrap"><button class="poly-btn delete-btn" data-uid="${_uId}">✕</button><span class="tip-box tip-sidebar">Remove Unassigned pricing data</span></span></div>
        `;
        uDiv.querySelector('.notes-btn').addEventListener('click', e => {
          e.stopPropagation();
          const up = polygons.find(p => p.id === _uId);
          if (up) openZoneEditor(_uId);
        });
        uDiv.querySelector('.delete-btn').addEventListener('click', e => {
          e.stopPropagation();
          const idx = polygons.findIndex(p => p.id === _uId);
          if (idx !== -1) { polygons.splice(idx, 1); persistZones(); renderPolygonList(); }
        });
        polyDiv.appendChild(uDiv);
      }

      cContent.appendChild(polyDiv);
      cGroup.appendChild(cHdr);
      cGroup.appendChild(cContent);
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
async function navigateToState(stateAbbr) {
  // Do NOT saveAppState here — state zoom is visual only, don't clobber persisted county

  // Remove existing state bbox layer if present
  if (map.getLayer('state-boundary-fill')) map.removeLayer('state-boundary-fill');
  if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
  if (map.getSource('state-boundary')) map.removeSource('state-boundary');

  try {
    // Fetch state boundary from Census TIGER API
    const fips = STATE_FIPS[stateAbbr];
    if (!fips) return;
    const url = `https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/State_County/MapServer/0/query?where=STATE='${fips}'&outFields=*&outSR=4326&f=geojson`;
    const res = await fetch(url);
    const geojson = await res.json();
    if (!geojson.features || !geojson.features.length) return;

    // Add state boundary line — solid 2px, renders on top of county
    map.addSource('state-boundary', { type: 'geojson', data: geojson });
    map.addLayer({ id: 'state-boundary-line', type: 'line', source: 'state-boundary',
      paint: { 'line-color': '#ffffff', 'line-width': 3 }
    });

    // Fit map to state bounds
    const bounds = new mapboxgl.LngLatBounds();
    const extendBounds = (coords) => {
      if (!Array.isArray(coords)) return;
      if (typeof coords[0] === 'number') { bounds.extend(coords); return; }
      coords.forEach(c => extendBounds(c));
    };
    geojson.features.forEach(f => extendBounds(f.geometry.coordinates));
    if (!bounds.isEmpty()) map.fitBounds(bounds, { padding: 60 });

    // New requirement — show county boundary lines for counties that have zones in this state
    const countiesWithZones = [...new Set(
      polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName).map(p => p.countyName)
    )];
    for (const cn of countiesWithZones) {
      const key = _countyKey(stateAbbr, cn);
      if (_countyLayers[key]) continue; // already drawn
      if (_countyGeoJSONCache[key]) {
        _addCountyBoundaryForKey(key, _countyGeoJSONCache[key]);
      } else {
        // fetch and draw in background
        _fetchCountyGeoJSON(fips, cn).then(cGeo => {
          if (cGeo) { _countyGeoJSONCache[key] = cGeo; _addCountyBoundaryForKey(key, cGeo); }
        }).catch(() => {});
      }
    }

  } catch(e) {
    console.error('navigateToState error:', e);
    showToast('Could not load state boundary', 'error');
  }
}

async function navigateToCounty(stateAbbr, countyName) {
  // Remove state bbox layer when navigating to county
  if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
  if (map.getSource('state-boundary')) map.removeSource('state-boundary');

  const ss = document.getElementById('stateSelect');
  const cs = document.getElementById('countySelect');
  ss.value = stateAbbr;
  _syncStateTrigger(stateAbbr);
  await loadCounties(true);
  cs.value = countyName;
  _syncCountyTrigger(countyName);
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
    _syncStateTrigger(poly.stateAbbr);
    await loadCounties(true);
    cs.value = poly.countyName;
    _syncCountyTrigger(poly.countyName);
    saveAppState();
    // Show county boundary without refitting the map view
    await loadCountyBoundaryOnly(poly.stateAbbr, poly.countyName);
  }
}

async function loadCountyBoundaryOnly(stateAbbr, countyName, cacheOnly) {
  try {
    const key = _countyKey(stateAbbr, countyName);
    // Always update _pendingCountyGeoJSON for draw validation
    if (_countyGeoJSONCache[key]) {
      _pendingCountyGeoJSON = _countyGeoJSONCache[key];
      if (cacheOnly) return; // 1.2 — cache hit, no visual layer needed
    }
    if (!cacheOnly && _countyLayers[key]) return; // layer already on map
    const fips = STATE_FIPS[stateAbbr];
    if (!fips) return;
    const geojson = await _fetchCountyGeoJSON(fips, countyName);
    if (!geojson) return;
    _countyGeoJSONCache[key] = geojson;
    _pendingCountyGeoJSON = geojson;
    if (!cacheOnly) _addCountyBoundaryForKey(key, geojson); // 1.2 — skip visual layer on init
  } catch(e) {}
}

// =========================================================
// CUSTOM CONFIRM DIALOGS
// =========================================================
let _confirmResolve = null;

function _showConfirm({ title, sub, okLabel = 'Delete', typePhrase = null, _overlapOnly = false }) {
  return new Promise(resolve => {
    _confirmResolve = resolve;
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmSub').textContent = sub;

    const cancelBtn = document.getElementById('confirmCancelBtn');
    const okBtn = document.getElementById('confirmOkBtn');
    const typeWrap = document.getElementById('confirmTypeWrap');
    const typeInput = document.getElementById('confirmTypeInput');

    if (_overlapOnly) {
      // 4 — Overlap error: single dismiss button only, no cancel
      cancelBtn.style.display = 'none';
      okBtn.textContent = okLabel;
      okBtn.disabled = false;
      okBtn.style.opacity = '1';
      // Style as ghost (not destructive red) since this is informational
      okBtn.style.background = '#fff';
      okBtn.style.color = 'var(--muted)';
      okBtn.style.border = '1px solid var(--border)';
      typeWrap.style.display = 'none';
      typeInput.oninput = null;
    } else {
      cancelBtn.style.display = '';
      okBtn.style.background = '';
      okBtn.style.color = '';
      okBtn.style.border = '';
      okBtn.textContent = okLabel;
      if (typePhrase) {
        typeWrap.style.display = '';
        document.getElementById('confirmTypePhrase').textContent = '"' + typePhrase + '"';
        typeInput.value = '';
        typeInput.classList.remove('valid');
        _updateConfirmOk(typePhrase);
        typeInput.oninput = () => _updateConfirmOk(typePhrase);
        okBtn.disabled = true;
        okBtn.style.opacity = '0.4';
      } else {
        typeWrap.style.display = 'none';
        typeInput.oninput = null;
        okBtn.disabled = false;
        okBtn.style.opacity = '1';
      }
    }

    document.getElementById('confirmModal').classList.add('open');
    if (typePhrase && !_overlapOnly) setTimeout(() => typeInput.focus(), 120);
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
  // Remove this county's boundary layer if no zones remain for it
  const key = _countyKey(stateAbbr, countyName);
  const remainingForCounty = polygons.filter(p => p.stateAbbr === stateAbbr && p.countyName === countyName);
  if (!remainingForCounty.length && _countyLayers[key]) {
    const sid = _countyLayers[key];
    if (map.getLayer(sid+'-fill')) map.removeLayer(sid+'-fill');
    if (map.getLayer(sid+'-line')) map.removeLayer(sid+'-line');
    if (map.getSource(sid)) map.removeSource(sid);
    delete _countyLayers[key];
  }
  // Remove state boundary if no zones remain at all
  if (!polygons.length) {
    _removeCountyLayer();
    if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
    if (map.getSource('state-boundary')) map.removeSource('state-boundary');
  }
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
  // Remove all county boundary layers
  Object.keys(_countyLayers).forEach(key => {
    const sid = _countyLayers[key];
    if (map.getLayer(sid+'-fill')) map.removeLayer(sid+'-fill');
    if (map.getLayer(sid+'-line')) map.removeLayer(sid+'-line');
    if (map.getSource(sid)) map.removeSource(sid);
    delete _countyLayers[key];
  });
  _removeCountyLayer();
  // Remove state boundary layer
  if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
  if (map.getSource('state-boundary')) map.removeSource('state-boundary');
  properties.forEach(prop => { if (prop.zone) { _markerColor(prop.marker, '#f7c948'); prop.zone = null; } });
  renderPolygonList(); persistZones(); _rebuildAllLabels();
  showToast('All zones cleared', 'info');
}

// =========================================================
// PERSIST ZONES
// =========================================================
function _polyToJSON(p) {
  return { id:p.id, name:p.name, letter:p.letter||'', stateAbbr:p.stateAbbr||'', countyName:p.countyName||'',
           color:p.color, points:p.points, description:p.description||'', pricingTiers:p.pricingTiers||[], isRect:!!p._isRect, bounds:p._bounds||null, propCount:p.propCount||0 };
}
async function persistZones() {
  await DB.saveZones(polygons.filter(p => !p._isUnassigned).map(_polyToJSON));
  // Save unassigned virtual polygons separately
  const unassignedEntries = polygons
    .filter(p => p._isUnassigned)
    .map(p => ({ id: p.id, stateAbbr: p.stateAbbr, countyName: p.countyName, pricingTiers: p.pricingTiers || [], description: p.description || '' }));
  await DB.saveUnassigned(unassignedEntries);
}
async function _loadAllCountyBoundaries(cacheOnly) {
  // Find all unique state+county combos that have zones
  const combos = [...new Set(polygons.filter(p => p.stateAbbr && p.countyName).map(p => _countyKey(p.stateAbbr, p.countyName)))];
  for (const key of combos) {
    if (!cacheOnly && _countyLayers[key]) continue; // already loaded (visual layers)
    const [sa, cn] = key.split('|');
    const fips = STATE_FIPS[sa]; if (!fips) continue;
    try {
      const geojson = await _fetchCountyGeoJSON(fips, cn);
      if (geojson) {
        _countyGeoJSONCache[key] = geojson;
        // 1.2 — on page refresh, only cache GeoJSON for draw enforcement; don't draw boundary layers
        if (!cacheOnly) _addCountyBoundaryForKey(key, geojson);
        // If this is the currently selected county, set as active boundary for draw enforcement
        const selState = document.getElementById('stateSelect').value;
        const selCounty = document.getElementById('countySelect').value;
        if (sa === selState && cn === selCounty) {
          _pendingCountyGeoJSON = geojson;
        }
      }
    } catch(e) {}
  }
  // Rebuild pills now that all county GeoJSON is cached — ensures centroids use county bounds
  if (polygons.length) _rebuildAllLabels();
}

async function restoreZones() {
  try {
    const data = await DB.loadZones();
    if (!data || !Array.isArray(data) || !data.length) return;
    // 1.2 — skipLayers=true: restore zone data + labels but don't draw polygon fills/outlines
    // Polygon layers appear only after user actively selects a county
    data.forEach(d => _loadZone(d, true));
    _rebuildAllLabels();
    renderPolygonList(); // always render sidebar immediately after zone data is loaded
    showToast(`Restored ${data.length} zone${data.length>1?'s':''}`, 'success');
  } catch(e) { console.error('restoreZones error:', e); }
  // Draw zone fill/line layers and county boundaries outside try block
  // so any map layer errors don't silently abort zone data restoration
  try { _restoreAllZoneLayers(); } catch(e) { console.warn('restoreZoneLayers error:', e); }
  setTimeout(() => { try { _loadAllCountyBoundaries(false); } catch(e) { console.warn('loadAllCountyBoundaries error:', e); } }, 500);

  // Restore unassigned virtual polygons (pricing data only, no map geometry)
  try {
    const unassigned = await DB.loadUnassigned();
    if (unassigned && unassigned.length) {
      unassigned.forEach(u => {
        const existing = polygons.find(p => p.id === u.id);
        if (!existing) {
          polygons.push({
            id: u.id, letter: '?', name: 'Unassigned', color: '#a0aec0',
            stateAbbr: u.stateAbbr, countyName: u.countyName,
            points: [], propCount: 0, pricingTiers: u.pricingTiers || [],
            description: u.description || '', _isUnassigned: true,
            labelMarker: null, handles: [],
          });
        } else {
          existing.pricingTiers = u.pricingTiers || [];
          existing.description  = u.description  || '';
        }
      });
    }
  } catch(e) { console.error('restoreUnassigned error:', e); }
}
function _loadZone(d, skipLayers) {
  const poly = { id:d.id, name:d.name, letter:d.letter||'', stateAbbr:d.stateAbbr||'', countyName:d.countyName||'',
    color:d.color, points:d.points, description:d.description||'', pricingTiers:d.pricingTiers||[], labelMarker:null, handles:[], _isRect:d.isRect||false, _bounds:d.bounds||null, propCount:d.propCount||0 };
  // Back-compat: derive stateAbbr/countyName from name if missing
  if (!poly.stateAbbr || !poly.countyName) {
    const m = poly.name.match(/^(.+) County,\s*([A-Z]{2})$/);
    if (m) { poly.countyName = m[1]; poly.stateAbbr = m[2]; }
    else { const m2 = poly.name.match(/^(.+),\s*([A-Z]{2})$/); if(m2){poly.countyName=m2[1];poly.stateAbbr=m2[2];} }
  }
  polygons.push(poly);
  if (!skipLayers) _addZoneLayers(poly); // 1.2 — skip polygon fill/line on page restore
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
  a.download = `landvaluator-${new Date().toISOString().slice(0,10)}.json`;
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
// 5.2 — extract spreadsheet ID from a full URL or raw ID string
function _parseSheetId(input) {
  const s = (input || '').trim();
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : s;
}

// 5.2 — update sheets modal status box between connected / not-connected
function _smSetConnected(isConnected, sheetName, sheetId, lastUrl) {
  const box   = document.getElementById('smStatusBox');
  const title = document.getElementById('smStatusTitle');
  const sub   = document.getElementById('smStatusSub');
  const openBtn = document.getElementById('smOpenSheetBtn');
  const discRow = document.getElementById('smDisconnectRow');
  const urlField = document.getElementById('smUrlField');
  const connectBtn = document.getElementById('smConnectBtn');

  if (isConnected) {
    box.className = 'sm-status-box connected';
    title.textContent = sheetName || 'Connected';
    sub.textContent = 'Connected';
    openBtn.style.display = '';
    openBtn.onclick = () => window.open('https://docs.google.com/spreadsheets/d/' + sheetId + '/edit', '_blank');
    discRow.style.display = '';
    urlField.style.display = 'none';
    connectBtn.textContent = 'Refresh & Sync';
  } else {
    box.className = 'sm-status-box not-connected';
    title.textContent = 'Sheet Not Connected';
    sub.textContent = lastUrl
      ? 'Previously connected URL restored below'
      : 'Enter your Google Sheets URL below to connect';
    openBtn.style.display = 'none';
    discRow.style.display = 'none';
    urlField.style.display = '';
    // Pre-fill with last used URL if available
    if (lastUrl) document.getElementById('sheetId').value = lastUrl;
    connectBtn.textContent = 'Connect & Load';
  }
}

function openSheetsModal() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  const badge = document.getElementById('sheetsModalCounty');
  const existing = (sa && cn) ? _getSheetConfig(sa, cn) : null;

  // County badge
  if (sa && cn) { badge.textContent = cn + ' County, ' + sa; badge.style.display = ''; }
  else { badge.style.display = 'none'; }

  // 5.2 — status box + URL field visibility
  const lastUrl = existing && existing.sheetUrl ? existing.sheetUrl : '';
  if (existing && existing.sheetId) {
    _smSetConnected(true, existing.sheetTitle || existing.sheetId, existing.sheetId, lastUrl);
    // Refresh title in background — picks up renames since last connect
    setTimeout(() => _fetchSheetName(existing.sheetId), 150);
  } else {
    _smSetConnected(false, '', '', lastUrl);
  }

  // Populate editable fields
  if (existing) {
    document.getElementById('sheetName').value  = existing.sheetName  || 'LI Raw Dataset';
    document.getElementById('colLat').value     = existing.colLat     || 'Latitude';
    document.getElementById('colLng').value     = existing.colLng     || 'Longitude';
    document.getElementById('colAPN').value     = existing.colAPN     || 'APN';
    document.getElementById('colCity').value    = existing.colCity    || 'City';
    document.getElementById('colCounty').value  = existing.colCounty  || 'County';
    document.getElementById('colState').value   = existing.colState   || 'State';
    document.getElementById('colZip').value     = existing.colZip     || 'ZIP';
    document.getElementById('colZone').value    = existing.colZone    || 'County Zone';
    // Show raw ID in URL field when not connected (pre-fill)
    if (!existing.sheetId) document.getElementById('sheetId').value = existing.sheetUrl || '';
  }

  document.getElementById('sheetsModal').classList.add('open');
}
function closeSheetsModal() { document.getElementById('sheetsModal').classList.remove('open'); }

function disconnectSheet() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  if (!sa || !cn) return;
  const key = _countyKey(sa, cn);
  const lastUrl = (sheetConfigs[key] && sheetConfigs[key].sheetUrl) ? sheetConfigs[key].sheetUrl : '';
  delete sheetConfigs[key];
  DB.saveSheetConfigs(sheetConfigs);
  if (sheetConfig && sheetConfig.stateAbbr === sa && sheetConfig.countyName === cn) {
    sheetConfig = null;
    setConnected(false);
  }
  // 5.2 — restore not-connected state, pre-fill last URL
  _smSetConnected(false, '', '', lastUrl);
  renderPolygonList();
  showToast('Sheet disconnected', 'info');
}


async function connectSheets() {
  const sa = document.getElementById('stateSelect').value;
  const cn = document.getElementById('countySelect').value;
  if (!sa || !cn) { showToast('Please select a State and County first', 'error'); return; }

  // When already connected, the URL field is hidden — fall back to the saved sheetId
  const activeCfg = _getSheetConfig(sa, cn) || sheetConfig;
  const urlField  = document.getElementById('sheetId');
  const rawInput  = urlField.value.trim() || (activeCfg && activeCfg.sheetUrl) || (activeCfg && activeCfg.sheetId) || '';
  if (!rawInput) { showToast('Please enter a Google Sheets URL or ID', 'error'); return; }

  // 5.2 — accept full URL or raw ID
  const sheetId = _parseSheetId(rawInput);
  if (!sheetId) { showToast('Could not parse a Sheet ID from that URL', 'error'); return; }

  // Check if this sheet ID is already used by a different county
  const existingKey = Object.entries(sheetConfigs).find(([key, cfg]) => {
    return cfg.sheetId === sheetId && key !== _countyKey(sa, cn);
  });
  if (existingKey) {
    const [existKey] = existingKey;
    const [existState, existCounty] = existKey.split('|');
    showToast('This sheet is already connected to ' + existCounty + ' County, ' + existState, 'error');
    return;
  }

  sheetConfig = {
    sheetId,
    sheetUrl:   rawInput,   // 5.2 — store original input for pre-fill on disconnect
    sheetTitle: '',         // 5.2 — populated after API call
    stateAbbr:  sa,
    countyName: cn,
    sheetName:  document.getElementById('sheetName').value.trim()   || 'LI Raw Dataset',
    colLat:     document.getElementById('colLat').value.trim()      || 'Latitude',
    colLng:     document.getElementById('colLng').value.trim()      || 'Longitude',
    colAPN:     document.getElementById('colAPN').value.trim()      || 'APN',
    colCity:    document.getElementById('colCity').value.trim()     || 'City',
    colCounty:  document.getElementById('colCounty').value.trim()   || 'County',
    colState:   document.getElementById('colState').value.trim()    || 'State',
    colZip:     document.getElementById('colZip').value.trim()      || 'ZIP',
    colZone:    document.getElementById('colZone').value.trim()     || 'County Zone',
  };
  // Save per-county
  _setSheetConfig(sa, cn, sheetConfig);
  showToast('Connecting to Google Sheets...', 'info');
  try {
    const r = await fetch('/.netlify/functions/sheets-read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetId: sheetConfig.sheetId, sheetName: sheetConfig.sheetName, colCounty: sheetConfig.colCounty || 'County', colAPN: sheetConfig.colAPN || 'APN' }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.error || 'HTTP ' + r.status);
    loadPropertiesFromFunction(data.properties, cn, data.scrubbedApns, data.ownerMap);
    const _sheetTitle = data.spreadsheetTitle || data.sheetTitle || sheetId;

    // 5.3 — County boundary validation
    const _fips = STATE_FIPS[sa];
    const _validResult = _fips ? await _validatePropertiesInCounty(properties, _fips, cn) : null;
    if (_validResult) {
      const { outsideProps, total } = _validResult;
      const pct = total > 0 ? outsideProps.length / total : 0;
      if (outsideProps.length > 0) {
        const countMsg = `${outsideProps.length} of ${total} propert${outsideProps.length === 1 ? 'y' : 'ies'}`;
        if (pct > 0.01) {
          // Block import
          _showBoundaryModal({
            title: `Import Blocked — ${countMsg} outside ${cn} County`,
            sub: 'Resolve the properties below before connecting. Select and copy the list for review.',
            icon: '🚫',
            outsideProps,
          });
          // Roll back — clear properties and connected state
          properties.forEach(p => { if (p.marker) p.marker.remove(); });
          properties = [];
          document.getElementById('statProps').textContent = '0';
          return;
        } else {
          // Warn but allow
          _showBoundaryModal({
            title: `Warning — ${countMsg} outside ${cn} County`,
            sub: `Less than 1% of properties fall outside the county boundary. You may proceed or cancel to investigate.`,
            icon: '⚠️',
            outsideProps,
            onProceed: () => {
              _finishSheetConnect({ sa, cn, sheetConfig, sheetId, rawInput, sheetTitle: _sheetTitle });
            },
          });
          return; // pause here — _finishSheetConnect called on Proceed
        }
      }
    }

    _finishSheetConnect({ sa, cn, sheetConfig, sheetId, rawInput, sheetTitle: _sheetTitle });

  } catch(e) { showToast('Connection failed: ' + e.message, 'error'); }
}

function _finishSheetConnect({ sa, cn, sheetConfig, sheetId, rawInput, sheetTitle }) {
  setConnected(true);

  // Store sheet title and update modal connected state
  sheetTitle = sheetTitle || sheetId;
  sheetConfig.sheetTitle = sheetTitle;
  _setSheetConfig(sa, cn, sheetConfig);
  _smSetConnected(true, sheetTitle, sheetId, rawInput);

  // Auto-assign to existing zones immediately
  const _cnNorm = cn.toLowerCase().trim();
  const countyPolys = polygons.filter(p => p.stateAbbr === sa && (p.countyName||'').toLowerCase().trim() === _cnNorm && !p._isUnassigned);
  if (countyPolys.length && properties.length) {
    let assigned = 0;
    properties.forEach(prop => {
      prop.zone = null;
      for (const poly of countyPolys) {
        if (pointInPolygon(prop.lat, prop.lng, poly.points)) {
          prop.zone = poly.letter;
          assigned++;
          break;
        }
      }
    });
    document.getElementById('statAssigned').textContent = assigned;
    if (_pinsVisible) _rebuildPins();
    renderPolygonList();
    persistZones();
  }

  const _assigned = properties.filter(p => p.zone).length;
  showToast('Connected: ' + cn + ' County — ' + properties.length + ' properties, ' + _assigned + ' assigned', 'success');
  closeSheetsModal();
}
function loadPropertiesFromFunction(props, countyOverride, scrubbedApns, ownerMap) {
  properties.forEach(p => { if (p.marker) p.marker.remove(); });
  properties = [];
  let skipped = 0;

  // Build APN whitelist set if provided
  const apnWhitelist = (scrubbedApns && scrubbedApns.length)
    ? new Set(scrubbedApns.map(a => a.trim().toLowerCase()))
    : null;
  const _ownerMap = ownerMap || {};

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

    // Filter to selected county only — use explicit override if provided
    const _selCounty = countyOverride || document.getElementById('countySelect').value;
    if (_selCounty && prop.county) {
      const propCounty = prop.county.toLowerCase().replace(' county','').trim();
      const selCounty  = _selCounty.toLowerCase().trim();
      if (propCounty !== selCounty) { skipped++; return; }
    }

    // APN whitelist filter — only retain properties present in Scrubbed and Priced
    if (apnWhitelist && apn) {
      if (!apnWhitelist.has(apn.trim().toLowerCase())) { skipped++; return; }
    }

    // Store property data without map marker (pins disabled pending better implementation)
    const ownerName = _ownerMap[apn ? apn.trim().toLowerCase() : ''] || '';
    properties.push({ lat, lng, apn, address, city, state, zip, county, acreage, liAcreage: prop.liAcreage || '', parcelLink: prop.parcelLink || '', ownerName, zone: zone || null, rowIndex, marker: null });
  });
  document.getElementById('statProps').textContent = properties.length;
  if (skipped) showToast(`${skipped} properties skipped — coordinates out of range or not in scrubbed list`, 'info');
  if (_pinsVisible) _rebuildPins();
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
function _markerColor(m, c) { if (m) m.getElement().style.background = c; }

// Pin toggle removed — pins disabled pending better implementation


// =========================================================
// COUNTY BOUNDARY
// =========================================================
const stateSelect = document.getElementById('stateSelect');
STATES.forEach(([name, abbr]) => {
  const o = document.createElement('option'); o.value = abbr; o.textContent = name;
  stateSelect.appendChild(o);
});

// =========================================================
// CUSTOM DROPDOWNS
// =========================================================
let _dropKeyBuffer = '';
let _dropKeyTimer = null;
let _dropFocusIdx = -1;
let _dropActiveWhich = null; // 'state' | 'county'

function _buildCustomList(listEl, options, currentValue, onSelect) {
  listEl.innerHTML = '';
  options.forEach(({ value, label }, idx) => {
    const div = document.createElement('div');
    div.className = 'custom-select-option' + (value === '' ? ' placeholder-opt' : '') + (value === currentValue ? ' selected' : '');
    div.textContent = label;
    div.dataset.value = value;
    div.dataset.idx = idx;
    div.addEventListener('click', (e) => { e.stopPropagation(); onSelect(value, label); });
    listEl.appendChild(div);
  });
}

function _dropSetFocus(listEl, idx) {
  const items = listEl.querySelectorAll('.custom-select-option:not(.placeholder-opt)');
  items.forEach(el => el.classList.remove('kbd-focus'));
  if (idx < 0 || idx >= items.length) { _dropFocusIdx = -1; return; }
  _dropFocusIdx = idx;
  const el = items[idx];
  el.classList.add('kbd-focus');
  el.scrollIntoView({ block: 'start' });
}

function _dropGetItems(listEl) {
  return Array.from(listEl.querySelectorAll('.custom-select-option:not(.placeholder-opt)'));
}

function _syncStateTrigger(value) {
  const trigger = document.getElementById('stateTrigger');
  const label = document.getElementById('stateLabel');
  if (!trigger || !label) return;
  const found = STATES.find(([, abbr]) => abbr === value);
  label.textContent = found ? found[0] : '— Select State —';
  trigger.classList.toggle('placeholder', !found);
}

function _syncCountyTrigger(value) {
  const trigger = document.getElementById('countyTrigger');
  const label = document.getElementById('countyLabel');
  if (!trigger || !label) return;
  const cs = document.getElementById('countySelect');
  const opt = cs ? Array.from(cs.options).find(o => o.value === value) : null;
  label.textContent = opt && value ? opt.textContent : '— Select County —';
  trigger.classList.toggle('placeholder', !value);
}

function _closeAllDropdowns() {
  ['stateTrigger','countyTrigger'].forEach(id => document.getElementById(id)?.classList.remove('open'));
  ['stateList','countyList'].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.classList.remove('open'); el.style.maxHeight = ''; }
  });
  _dropFocusIdx = -1;
  _dropActiveWhich = null;
}

function _openDropdown(which, trigger, list, opts, currentValue, onSelect) {
  _buildCustomList(list, opts, currentValue, onSelect);

  // For county: auto-size to content (no scroll needed for small lists)
  if (which === 'county') {
    // Let CSS handle max-height but shrink if fewer items fit
    const itemH = 34; // approx px per item
    const naturalH = opts.length * itemH;
    list.style.maxHeight = naturalH < 800 ? naturalH + 'px' : '800px';
  } else {
    list.style.maxHeight = '800px';
  }

  trigger.classList.add('open');
  list.classList.add('open');
  _dropFocusIdx = -1;
  _dropActiveWhich = which;

  // Scroll selected item into view
  const selected = list.querySelector('.selected');
  if (selected) selected.scrollIntoView({ block: 'nearest' });
}

function _toggleDropdown(which) {
  const isState = which === 'state';
  const triggerId = isState ? 'stateTrigger' : 'countyTrigger';
  const listId = isState ? 'stateList' : 'countyList';
  const trigger = document.getElementById(triggerId);
  const list = document.getElementById(listId);
  if (!trigger || !list) return;

  const isOpen = list.classList.contains('open');
  _closeAllDropdowns();

  if (!isOpen) {
    if (isState) {
      const opts = [{ value: '', label: '— Select State —' }, ...STATES.map(([name, abbr]) => ({ value: abbr, label: name }))];
      _openDropdown('state', trigger, list, opts, stateSelect.value, (val) => {
        stateSelect.value = val;
        _syncStateTrigger(val);
        _closeAllDropdowns();
        const cs = document.getElementById('countySelect');
        cs.innerHTML = '<option value="">— Select County —</option>';
        _syncCountyTrigger('');
        stateSelect.dispatchEvent(new Event('change'));
      });
    } else {
      const cs = document.getElementById('countySelect');
      if (!cs.options.length || (cs.options.length === 1 && cs.options[0].value === '')) return; // no counties loaded yet
      const opts = Array.from(cs.options).map(o => ({ value: o.value, label: o.textContent }));
      _openDropdown('county', trigger, list, opts, cs.value, (val) => {
        cs.value = val;
        _syncCountyTrigger(val);
        _closeAllDropdowns();
        cs.dispatchEvent(new Event('change'));
      });
    }
  }
}

// Keyboard navigation for open dropdown
// Escape key — undo last draw point (single) or cancel draw (double-tap)
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return;
  if (drawMode !== 'polygon' || polyState !== 'drawing') return;
  e.preventDefault();
  e.stopPropagation();
  const now = Date.now();
  const isDoubleTap = (now - _lastEscapeTime) < 350;
  _lastEscapeTime = now;
  if (isDoubleTap) {
    cancelDraw();
    showToast('Draw cancelled', 'info');
  } else {
    undoLastDrawPoint();
  }
});

document.addEventListener('keydown', (e) => {
  if (!_dropActiveWhich) return;
  const listId = _dropActiveWhich === 'state' ? 'stateList' : 'countyList';
  const list = document.getElementById(listId);
  if (!list || !list.classList.contains('open')) return;
  const items = _dropGetItems(list);
  if (!items.length) return;

  if (e.key === 'Escape') {
    e.preventDefault();
    _closeAllDropdowns();
    return;
  }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    _dropSetFocus(list, Math.min(_dropFocusIdx + 1, items.length - 1));
    return;
  }
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    _dropSetFocus(list, Math.max(_dropFocusIdx - 1, 0));
    return;
  }
  if (e.key === 'Enter') {
    e.preventDefault();
    if (_dropFocusIdx >= 0 && items[_dropFocusIdx]) items[_dropFocusIdx].click();
    return;
  }

  // Type-ahead: buffer key presses, jump to first match
  if (e.key.length === 1) {
    clearTimeout(_dropKeyTimer);
    _dropKeyBuffer += e.key.toLowerCase();
    _dropKeyTimer = setTimeout(() => { _dropKeyBuffer = ''; }, 800);
    const match = items.findIndex(el => el.textContent.toLowerCase().startsWith(_dropKeyBuffer));
    if (match >= 0) _dropSetFocus(list, match);
  }
});

// Close dropdowns on outside click
document.addEventListener('click', (e) => {
  if (!e.target.closest('#stateDropdown') && !e.target.closest('#countyDropdown')) {
    _closeAllDropdowns();
  }
});


async function loadCounties(silent) {
  const abbr = stateSelect.value; if (!abbr) return;
  const cs = document.getElementById('countySelect');
  if (!silent) { cs.innerHTML = '<option value="">Loading counties...</option>'; _syncCountyTrigger(''); }
  function fill(counties) {
    cs.innerHTML = '<option value="">— Select County —</option>';
    counties.forEach(n => { const o=document.createElement('option'); o.value=n; o.textContent=n+' County'; cs.appendChild(o); });
    if (!silent) { const s=loadAppState(); if(s&&s.state===abbr&&s.county) { cs.value=s.county; } }
    _syncCountyTrigger(cs.value);
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
    if (!silent) { cs.innerHTML='<option value="">Failed to load</option>'; _syncCountyTrigger(''); }
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
  const _initVis = map.getZoom() >= COUNTY_PILL_ZOOM ? 'visible' : 'none';
  map.addLayer({ id:sid+'-fill', type:'fill', source:sid, paint:{'fill-color':'#000000','fill-opacity':0.08}, layout:{'visibility':_initVis} });
  map.addLayer({ id:sid+'-line', type:'line', source:sid, paint:{'line-color':'#6600cc','line-width':3}, layout:{'visibility':_initVis} });
  // Click on county fill area (when no zone layer is on top)
  map.on('click', sid+'-fill', async (e) => {
    if (drawMode === 'polygon') return;
    if (map.queryRenderedFeatures(e.point, { layers: [LAYER_PINS] }).length) return;
    const [sa, cn] = key.split('|');
    const alreadySelected = stateSelect.value === sa && document.getElementById('countySelect').value === cn;
    if (!alreadySelected) {
      stateSelect.value = sa;
      _syncStateTrigger(sa);
      const cs = document.getElementById('countySelect');
      await loadCounties(true);
      cs.value = cn;
      if (cs.value !== cn) {
        const o = document.createElement('option');
        o.value = cn; o.textContent = cn + ' County';
        cs.appendChild(o);
        cs.value = cn;
      }
      _syncCountyTrigger(cn);
      saveAppState();
      const saved = _getSheetConfig(sa, cn);
      if (saved) { sheetConfig = saved; setConnected(true); }
      else { sheetConfig = null; setConnected(false); }
      renderPolygonList();
    }
    // loadCounty() intentionally removed — use toolbar county name to zoom
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
    const stateName = STATES.find(([,abbr]) => STATE_FIPS[abbr] === fips)?.[0];
    if (stateName) {
      const url = `https://nominatim.openstreetmap.org/search?county=${encodeURIComponent(countyName)}&state=${encodeURIComponent(stateName)}&country=USA&format=geojson&polygon_geojson=1&limit=1`;
      const r = await fetch(url, { headers: { 'User-Agent': 'LandValuator/1.0' } });
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
  // 1.1 — Always clear state boundary when a county is selected
  if (map.getLayer('state-boundary-line')) map.removeLayer('state-boundary-line');
  if (map.getSource('state-boundary')) map.removeSource('state-boundary');
  // 1.2 — Ensure zone polygon layers are drawn when user actively selects a county
  // (they are skipped on page restore; this is the first chance to draw them)
  _restoreAllZoneLayers();
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
    // Rebuild county pills now that GeoJSON is cached — ensures centroid uses county bounds
    if (polygons.length) _rebuildAllLabels();
    // Remove keyed layers for counties with no zones (run AFTER new county is registered)
    Object.keys(_countyLayers).forEach(k => {
      if (k === key) return; // keep the current county
      const hasZones = polygons.some(p => _countyKey(p.stateAbbr, p.countyName) === k);
      if (!hasZones) {
        const sid = _countyLayers[k];
        if (map.getLayer(sid+'-fill')) map.removeLayer(sid+'-fill');
        if (map.getLayer(sid+'-line')) map.removeLayer(sid+'-line');
        if (map.getSource(sid)) map.removeSource(sid);
        delete _countyLayers[k];
      }
    });
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
  DB.saveAppState({ state: stateSelect.value, county: document.getElementById('countySelect').value }); // fire-and-forget async
}
function loadAppState() {
  return DB.loadAppState(); // returns Promise
}

// ── TOOLTIP TOGGLE ──────────────────────────────────────
function toggleTooltips() {
  const isOff = document.body.classList.toggle('tooltips-off');
  DB.saveUIState('tooltips_off', isOff);
  _updateTooltipBtn(isOff);
}
function _updateTooltipBtn(isOff) {
  const btn = document.getElementById('tooltipToggleBtn');
  if (btn) btn.classList.toggle('active', !isOff);
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
  _initPinLayer();
  _mapLoadFired = true;

  // If user is already logged in (session restored), run app init now
  // Otherwise wait for onAuthStateChange to fire with a valid session
  if (_currentUser) {
    _authAppReady = true;
    _initAppAfterAuth();
  } else {
    // Show auth modal while we wait
    document.getElementById('authModal').classList.add('open');
    // Safety net timer still needed for _mapInitComplete
    setTimeout(() => { _mapInitComplete = true; }, 600);
  }

});

// =========================================================
// LOCALSTORAGE → SUPABASE MIGRATION
// On first login, check if localStorage has existing data
// and offer to migrate it to the user's Supabase account
// =========================================================
async function _checkAndMigrateLocalData() {
  try {
    const lsZones = localStorage.getItem('lv_zones');
    const lsConfigs = localStorage.getItem('lv_sheet_configs');
    const lsAppState = localStorage.getItem('lv_app_state');
    const lsUnassigned = localStorage.getItem('lv_unassigned');

    // Check if there's anything worth migrating
    const hasZones = lsZones && JSON.parse(lsZones).length > 0;
    if (!hasZones) return; // nothing to migrate

    // Check if Supabase already has data for this user
    const { data: existingZones } = await _supa.from('zones').select('data').eq('user_id', _currentUser.id).maybeSingle();
    const alreadyHasData = existingZones?.data && existingZones.data.length > 0;
    if (alreadyHasData) return; // already migrated, skip

    // Prompt user
    const confirmed = await _showConfirm({
      title: 'Migrate Your Saved Zones?',
      sub: `We found ${JSON.parse(lsZones).length} zone(s) saved locally on this device. Would you like to import them into your account so they're available everywhere?`,
      okLabel: 'Import Zones',
    });
    if (!confirmed) return;

    // Migrate all data
    if (lsZones)      await DB.saveZones(JSON.parse(lsZones));
    if (lsConfigs)    await DB.saveSheetConfigs(JSON.parse(lsConfigs));
    if (lsAppState)   await DB.saveAppState(JSON.parse(lsAppState));
    if (lsUnassigned) await DB.saveUnassigned(JSON.parse(lsUnassigned));

    // Clear localStorage now that data is in Supabase
    ['lv_zones','lv_sheet_configs','lv_app_state','lv_unassigned'].forEach(k => localStorage.removeItem(k));

    showToast('Zones migrated to your account ✓', 'success');
  } catch(e) {
    console.warn('Migration check error:', e);
  }
}

// =========================================================
// APP INIT — runs after auth confirmed
// =========================================================
async function _initAppAfterAuth() {
  // Load ALL UI state from Supabase into memory cache first
  // This allows renderPolygonList() to read accordion state synchronously
  await DB.loadAllUIState();

  // Init tooltip state (must be after auth so DB.loadUIState has a user)
  _initTooltipToggle();

  // Check for localStorage data to migrate on first login
  await _checkAndMigrateLocalData();

  // Load sheet configs first
  const _savedCfgs = await DB.loadSheetConfigs();
  if (_savedCfgs) { sheetConfigs = _savedCfgs; }

  // Restore zones or load from URL
  const fromURL = loadZonesFromURL();
  if (!fromURL) await restoreZones();

  // Safety net AFTER zones are loaded — rebuild labels/boundaries
  setTimeout(() => {
    if (polygons.length) { _rebuildAllLabels(); _loadAllCountyBoundaries(true); }
    _mapInitComplete = true;
  }, 600);

  // Check for ?state=XX&county=Name deep-link params
  const _urlParams   = new URLSearchParams(window.location.search);
  const _urlState    = (_urlParams.get('state')  || '').trim().toUpperCase();
  const _urlCounty   = (_urlParams.get('county') || '').trim();
  const _hasDeepLink = !!(_urlState && _urlCounty);

  if (_hasDeepLink) {
    _urlParams.delete('state');
    _urlParams.delete('county');
    const _cleanSearch = _urlParams.toString();
    history.replaceState(null, '', window.location.pathname + (_cleanSearch ? '?' + _cleanSearch : ''));
  }

  const appState = await DB.loadAppState();
  const _initState  = _hasDeepLink ? _urlState  : (appState && appState.state);
  const _initCounty = _hasDeepLink ? _urlCounty : (appState && appState.county);

  if (_initState) {
    stateSelect.value = _initState;
    _syncStateTrigger(_initState);
    loadCounties().then(() => {
      const countySelect = document.getElementById('countySelect');
      if (_initCounty) {
        countySelect.value = _initCounty;
        _syncCountyTrigger(_initCounty);
        const saved = _getSheetConfig(_initState, _initCounty);
        if (saved) { sheetConfig = saved; setConnected(true); }
      }

      renderPolygonList();

      if (_initCounty) {
        loadCountyBoundaryOnly(_initState, _initCounty, true);

        if (_hasDeepLink) {
          loadCounty();
          const _allStateCombos = [...new Set(polygons.map(p => p.stateAbbr).filter(Boolean))];
          _allStateCombos.forEach(sa => {
            DB.saveUIState('state_open_' + sa, sa === _initState);
            const _countiesInState = [...new Set(polygons.filter(p => p.stateAbbr === sa).map(p => p.countyName).filter(Boolean))];
            _countiesInState.forEach(cn => {
              DB.saveUIState('county_open_' + sa + '_' + cn, sa === _initState && cn === _initCounty);
            });
          });
          renderPolygonList();
        }

        const _allConfigs = Object.entries(sheetConfigs || {});
        if (_allConfigs.length) {
          showToast('Reconnecting sheets...', 'info');
          let _reconnected = 0;
          const _reconnectAll = async () => {
            for (const [key, cfg] of _allConfigs) {
              if (!cfg || !cfg.sheetId) continue;
              const [_sa, _cn] = key.split('|');
              try {
                const r = await fetch('/.netlify/functions/sheets-read', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ sheetId: cfg.sheetId, sheetName: cfg.sheetName || 'LI Raw Dataset', colCounty: cfg.colCounty || 'County', colAPN: cfg.colAPN || 'APN' }),
                });
                const data = await r.json();
                if (!data.properties || !data.properties.length) continue;
                const _cnNorm = _cn.toLowerCase().trim();
                const isCurrentCounty = _sa === _initState && _cnNorm === (_initCounty||'').toLowerCase().trim();
                if (isCurrentCounty || !properties.length) {
                  loadPropertiesFromFunction(data.properties, _cn, data.scrubbedApns, data.ownerMap);
                  document.getElementById('statProps').textContent = properties.length;
                  if (isCurrentCounty) { sheetConfig = cfg; setConnected(true); }
                }
                const _doAssign = () => {
                  const _rPolys = polygons.filter(p => p.stateAbbr === _sa && (p.countyName||'').toLowerCase().trim() === _cnNorm);
                  if (!_rPolys.length) return;
                  const _cnProps = properties.filter(p => {
                    if (!p.county) return isCurrentCounty;
                    return p.county.toLowerCase().replace(' county','').trim() === _cnNorm;
                  });
                  let assigned = 0;
                  _cnProps.forEach(prop => {
                    prop.zone = null;
                    for (const poly of _rPolys) {
                      if (pointInPolygon(prop.lat, prop.lng, poly.points)) { prop.zone = poly.letter; assigned++; break; }
                    }
                  });
                  _reconnected += assigned;
                  document.getElementById('statAssigned').textContent =
                    parseInt(document.getElementById('statAssigned').textContent || '0') + assigned;
                  renderPolygonList();
                  persistZones();
                };
                if (polygons.length) { _doAssign(); } else { map.once('idle', _doAssign); }
              } catch(e) { console.warn('Reconnect failed for', key, e); }
            }
            if (_reconnected > 0) showToast(`Sheets reconnected — ${_reconnected} properties assigned`, 'success');
            else showToast('Sheets reconnected', 'success');
          };
          _reconnectAll();
        }
      }
    });
  } else if (polygons.length) {
    renderPolygonList();
  }
}

