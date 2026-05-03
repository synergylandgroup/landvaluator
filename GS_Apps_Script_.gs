// ============================================================
//  MAILING CAMPAIGN TEMPLATE — Apps Script
//  Paste this entire script into Extensions > Apps Script
// ============================================================
//
//  ⚠️  PRICING ACREAGE VERSION NOTE
//  This script uses "LI Calculated Acreage" for all pricing
//  calculations. Falls back to standard Acreage if LI cell is blank.
// ============================================================


// ============================================================
//  LICENSE VERIFICATION
//  Checks once per 24 hours that this sheet is registered to
//  a LandValuator account. Result cached in Script Properties.
// ============================================================

const LV_VERIFY_URL = 'https://landvaluator.app/.netlify/functions/verify-sheet';
const LV_AUTH_KEY   = 'lv_auth_time';

function _isAuthorized() {
  const props = PropertiesService.getScriptProperties();
  const lastCheck = Number(props.getProperty(LV_AUTH_KEY) || 0);
  const hoursSince = (Date.now() - lastCheck) / 3600000;
  if (hoursSince < 24) return true; // cached — skip network call

  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  try {
    const resp = UrlFetchApp.fetch(LV_VERIFY_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ sheetId }),
      muteHttpExceptions: true,
    });
    const result = JSON.parse(resp.getContentText());
    if (result.authorized) {
      props.setProperty(LV_AUTH_KEY, String(Date.now()));
      return true;
    }
  } catch (e) {
    // Network error — fail closed
  }
  return false;
}

function _requireAuth() {
  if (_isAuthorized()) return true;
  SpreadsheetApp.getUi().alert(
    '⚠️ Not Authorized',
    'This spreadsheet is not connected to a LandValuator account.\n\n' +
    'To use these features, connect this sheet at landvaluator.app.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return false;
}

//  Column name mappings
//  All column lookups are name-based and resolved dynamically at runtime
//  using the header row of each sheet. No column positions are hardcoded.
//
//  If you rename a column header in any sheet, update the matching
//  string here to keep lookups in sync.

// LI Raw Dataset → Scrubbed and Priced
//  Key   = header name in LI Raw Dataset
//  Value = header name in Scrubbed and Priced
//  Resolved at runtime via getHeaderMap() — immune to column reordering.
const RAW_TO_SP_HEADERS = {
  // 'Owner Name(s)' is intentionally excluded here — it is built from
  // individual Owner 1 / Owner 2 name parts via buildOwnerName() below.
  // 'County Zone' is intentionally excluded here — it is assigned internally
  // by the script and written only to Scrubbed and Priced, never stored in LI Raw Dataset.
  'Mail Address'               : 'Mail Address',
  'Mail City'                  : 'Mail City',
  'Mail State'                 : 'Mail State',
  'Mail ZIP'                   : 'Mail ZIP',
  'Parcel Address'             : 'Parcel Address',
  'City'                       : 'Parcel City',
  'State'                      : 'Parcel State',
  'ZIP'                        : 'Parcel ZIP',
  'County'                     : 'Parcel County',
  'Legal Description'          : 'Legal Description',
  'APN'                        : 'APN',
  'Acreage'                    : 'Acreage',
  'Calculated Acreage'         : 'LI Calculated Acreage',
  'Market Value Estimate'      : 'LI Market Value Estimate',
  'Market Value Estimate PPA'  : 'LI Market Value Est PPA',
  'Parcel Link'                : 'Parcel Link',
  'Comping Link'               : 'Comping Link',
  'SellerIQ'                   : 'Seller IQ',
};

// Scrubbed and Priced — header names for key columns
//  Used as fallback defaults. Tier and range offer column headers are dynamic —
//  they update automatically when Pricing Settings percentages change.
//  All lookups use prefixCol() or flexCol() so exact percentage suffixes
//  (e.g. "50%") don't need to match — only the base name matters.
const SP_HEADERS = {
  POLYGON      : 'County Zone',
  OWNER_NAME   : 'Owner Name(s)',
  ACREAGE      : 'Acreage',
  MANUAL_PPA   : 'Manually Calculated PPA',
  MANUAL_MV    : 'Manually Calculated Market Value',
  BLIND_OFFER  : 'Blind Offer Price',
  RANGE_LOW    : 'Range Offer Low (50%)',
  RANGE_HIGH   : 'Range Offer High (65%)',
  SELLER_IQ    : 'Seller IQ',
};

// Blind Offer Mail Ready ✉️ — ordered list of headers to sync (must match that sheet's row 1)
//  Range offer columns use flexCol() matching so percentage suffixes are ignored.
const BLIND_HEADERS = [
  'Owner Name(s)', 'Mail Address', 'Mail City', 'Mail State', 'Mail ZIP',
  'Parcel Address', 'Parcel City', 'Parcel State', 'Parcel ZIP', 'Parcel County',
  'Legal Description', 'APN', 'LI Calculated Acreage', 'Blind Offer Price', 'Mailer Code',
];

// Range Offer Mail Ready ✉️ — ordered list of headers to sync (must match that sheet's row 1)
const RANGE_HEADERS = [
  'Owner Name(s)', 'Mail Address', 'Mail City', 'Mail State', 'Mail ZIP',
  'Parcel Address', 'Parcel City', 'Parcel State', 'Parcel ZIP', 'Parcel County',
  'Legal Description', 'APN', 'LI Calculated Acreage', 'Range Offer Low (50%)', 'Range Offer High (65%)', 'Mailer Code',
];

// Pricing Settings — labels used to identify tier rows when scanning the sheet.
//  Range offer multipliers are read directly from fixed cells L4:L5.
//  Tier multipliers, min/max values, and PPA data are found by scanning for these labels.
const PRICING_LABELS = {
  BLIND_PREFIXES    : ['T —', 'U —', 'V —', 'W —', 'X —', 'Y —'],
  MULTIPLIER_HEADER : 'Multiplier (%)',
  MIN_HEADER        : 'Min Value ($)',
  MAX_HEADER        : 'Max Value ($)',
  POLYGON_HEADER    : 'County Zone',
  MIN_ACRES_HEADER  : 'Min Acres',
  MAX_ACRES_HEADER  : 'Max Acres',
  PPA_HEADER        : 'Retail PPA ($)',
};


// ============================================================
//  UTILITY FUNCTIONS
// ============================================================

/**
 * Build a { headerName: columnNumber } map from row 1.
 * Keys are stored in lowercase so lookups are case-insensitive.
 */
function getHeaderMap(sheet) {
  if (!sheet) throw new Error('Sheet not found — check that the tab exists and the name matches exactly.');
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h).trim().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').toLowerCase();
    if (key) map[key] = i + 1; // 1-based
  });
  return map;
}

/** Get column number or throw a clear error. */
function requireCol(map, header, sheetName) {
  const col = map[String(header).trim().toLowerCase()];
  if (!col) throw new Error('Column "' + header + '" not found in "' + sheetName + '" (check row 1 headers).');
  return col;
}

/** Safe case-insensitive lookup (returns undefined if not found). */
function mapCol(map, header) {
  return map[String(header).trim().toLowerCase()];
}

/** Find column by prefix match — for dynamic headers like "Range Offer Low (50%)". */
function prefixCol(map, prefix) {
  const p = prefix.trim().toLowerCase();
  for (const [key, col] of Object.entries(map)) {
    if (key.startsWith(p)) return col;
  }
  return undefined;
}

/** Exact match first, prefix match fallback — strips trailing "(50%)" etc. */
function flexCol(map, header) {
  return mapCol(map, header) || prefixCol(map, header.replace(/\s*\(.*\)$/, '').trim());
}


// ============================================================
//  PRICING SETTINGS
// ============================================================

/**
 * Read pricing settings.
 * Returns: { blind: [{mult, min, max}×5], range: [{mult}×2], polygonPPA: [{polygon, minAcres, maxAcres, ppa}] }
 */
function getPricingSettings() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pricing Settings');

  const defaultBlind = [
    { mult: 0.50, min: 25000,  max: 70000  },
    { mult: 0.55, min: 70000,  max: 95000  },
    { mult: 0.60, min: 95000,  max: 125000 },
    { mult: 0.65, min: 125000, max: 250000 },
    { mult: 0.70, min: 250000, max: 999999 },
  ];
  const defaultRange = [{ mult: 0.50 }, { mult: 0.65 }];

  if (!sheet) return { blind: defaultBlind, range: defaultRange, polygonPPA: [] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  let headerRowIdx = -1;
  let polygonCol = -1, minAcresCol = -1, maxAcresCol = -1, ppaCol = -1;
  let blindMultCol = -1, minCol = -1, maxCol = -1;
  let blindMultFound = false;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = String(data[r][c]).trim().toLowerCase();
      if (val === PRICING_LABELS.POLYGON_HEADER.toLowerCase())    { polygonCol = c; headerRowIdx = r; }
      if (val === PRICING_LABELS.MIN_ACRES_HEADER.toLowerCase())  { minAcresCol = c; }
      if (val === PRICING_LABELS.MAX_ACRES_HEADER.toLowerCase())  { maxAcresCol = c; }
      if (val === PRICING_LABELS.PPA_HEADER.toLowerCase())        { ppaCol = c; }
      if (val === PRICING_LABELS.MULTIPLIER_HEADER.toLowerCase()) {
        if (!blindMultFound) { blindMultCol = c; blindMultFound = true; }
      }
      if (val === PRICING_LABELS.MIN_HEADER.toLowerCase() && minCol === -1) minCol = c;
      if (val === PRICING_LABELS.MAX_HEADER.toLowerCase() && maxCol === -1) maxCol = c;
    }
  }

  // Read polygon PPA rows
  const polygonPPA = [];
  if (polygonCol >= 0 && headerRowIdx >= 0) {
    for (let r = headerRowIdx + 2; r < data.length; r++) { // +2 to skip header + hint row
      const row     = data[r];
      const polygon = String(row[polygonCol] || '').trim().toUpperCase();
      if (!polygon) continue;
      const minAcres = minAcresCol >= 0 ? Number(row[minAcresCol]) : 0;
      const maxAcres = maxAcresCol >= 0 ? Number(row[maxAcresCol]) : 999999;
      const ppa      = ppaCol >= 0      ? Number(row[ppaCol])      : 0;
      if (ppa > 0) polygonPPA.push({ polygon, minAcres, maxAcres, ppa });
    }
  }

  // Read blind offer tiers
  const blind = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    let rowMatched = false;
    for (let c = 0; c < row.length; c++) {
      const label = String(row[c]).trim();
      if (!rowMatched && PRICING_LABELS.BLIND_PREFIXES.some(p => label.startsWith(p))) {
        const mult = blindMultCol >= 0 ? Number(row[blindMultCol]) : 0;
        const min  = minCol >= 0       ? Number(row[minCol])       : 0;
        const max  = maxCol >= 0       ? Number(row[maxCol])       : 999999;
        blind.push({ mult, min, max });
        rowMatched = true;
      }
    }
  }

  if (blind.length > 5) blind.splice(5);

  // Read range multipliers directly from L4:L5 — same cells used by updatePricingHeaders_
  // This avoids fragile label scanning and always reflects the current live values
  const rangeLowVal  = data.length > 3 ? Number(data[3][11]) : 0;  // L4 = row index 3, col index 11
  const rangeHighVal = data.length > 4 ? Number(data[4][11]) : 0;  // L5 = row index 4, col index 11
  const parseMult = v => v > 0 ? (v > 1 ? v / 100 : v) : 0;
  const rangeLow  = parseMult(rangeLowVal);
  const rangeHigh = parseMult(rangeHighVal);
  const rangeFromSheet = rangeLow > 0 && rangeHigh > 0
    ? [{ mult: rangeLow }, { mult: rangeHigh }]
    : defaultRange;
  return {
    blind      : blind.length > 0 ? blind : defaultBlind,
    range      : rangeFromSheet,
    polygonPPA : polygonPPA,
  };
}

/**
 * Look up PPA for a given polygon + acreage.
 * Falls back to 'ALL' polygon if no exact match found.
 */
function lookupPPA(polygonPPA, polygon, acreage) {
  const poly = String(polygon || '').trim().toUpperCase();
  const ac   = Number(acreage) || 0;
  for (const entry of polygonPPA) {
    if (entry.polygon === poly && ac >= entry.minAcres && ac < entry.maxAcres) return entry.ppa;
  }
  for (const entry of polygonPPA) {
    if (entry.polygon === 'ALL' && ac >= entry.minAcres && ac < entry.maxAcres) return entry.ppa;
  }
  return 0;
}


// ============================================================
//  CORE FUNCTIONS
// ============================================================

/** Refresh pricing only (no data reload). Pass silent=true to suppress the confirmation alert. */
function refreshOfferPrices(silent) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const spSheet = ss.getSheetByName('Scrubbed and Priced');
  if (!spSheet) { SpreadsheetApp.getUi().alert('Could not find "Scrubbed and Priced" sheet.'); return; }
  const lastRow = spSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data in Scrubbed and Priced. Run "Load data" first.'); return; }

  try {
    const spHeaders  = getHeaderMap(spSheet);
    const polygonCol   = requireCol(spHeaders, SP_HEADERS.POLYGON,     'Scrubbed and Priced');
    const acreageCol   = requireCol(spHeaders, SP_HEADERS.ACREAGE,     'Scrubbed and Priced');
    const liAcreageCol = mapCol(spHeaders, 'LI Calculated Acreage');
    const manPPACol  = requireCol(spHeaders, SP_HEADERS.MANUAL_PPA,  'Scrubbed and Priced');
    const manMVCol   = requireCol(spHeaders, SP_HEADERS.MANUAL_MV,   'Scrubbed and Priced');
    const blindCol   = requireCol(spHeaders, SP_HEADERS.BLIND_OFFER, 'Scrubbed and Priced');
    const rlCol      = prefixCol(spHeaders, 'Range Offer Low')  || requireCol(spHeaders, SP_HEADERS.RANGE_LOW,  'Scrubbed and Priced');
    const rhCol      = prefixCol(spHeaders, 'Range Offer High') || requireCol(spHeaders, SP_HEADERS.RANGE_HIGH, 'Scrubbed and Priced');
    const { blind: blindTiers, range: rangeTiers, polygonPPA } = getPricingSettings();

    // Find tier columns using the actual percentages from Pricing Settings,
    // not hardcoded values — so if G4 changes from 50% to 20%, we find "20% Offer Price"
    const tierCols = blindTiers.map(tier => {
      const pct = Math.round(tier.mult * 100) + '%';
      return prefixCol(spHeaders, pct + ' Offer') || null;
    });
    const dataRows = lastRow - 1;
    const lastCol  = spSheet.getLastColumn();
    const allData  = spSheet.getRange(2, 1, dataRows, lastCol).getValues();

    const ppaOut = [], mvOut = [], blindOut = [], rlOut = [], rhOut = [];
    const tierOut = tierCols.map(() => []);

    for (let i = 0; i < allData.length; i++) {
      const row     = allData[i];
      const polygon   = row[polygonCol - 1];
      const liRaw     = liAcreageCol ? row[liAcreageCol - 1] : '';
      const acreage   = (liRaw !== '' && liRaw !== null) ? (Number(liRaw) || 0) : (Number(row[acreageCol - 1]) || 0);
      const ppa       = lookupPPA(polygonPPA, polygon, acreage);
      const mv        = ppa > 0 && acreage > 0 ? ppa * acreage : 0;

      ppaOut.push([ppa > 0 ? ppa : '']);
      mvOut.push( [mv  > 0 ? mv  : '']);

      let maxTier = '';
      tierCols.forEach((col, t) => {
        if (!col) return;
        const tier = blindTiers[t];
        const val  = (tier && mv >= tier.min && mv < tier.max) ? mv * tier.mult : '';
        tierOut[t].push([val]);
        if (val !== '' && (maxTier === '' || val > maxTier)) maxTier = val;
      });

      blindOut.push([maxTier]);
      rlOut.push([mv > 0 ? mv * rangeTiers[0].mult : '']);
      rhOut.push([mv > 0 ? mv * rangeTiers[1].mult : '']);
    }

    spSheet.getRange(2, manPPACol, dataRows, 1).setValues(ppaOut);
    spSheet.getRange(2, manMVCol,  dataRows, 1).setValues(mvOut);
    tierCols.forEach((col, t) => { if (col) spSheet.getRange(2, col, dataRows, 1).setValues(tierOut[t]); });
    spSheet.getRange(2, blindCol, dataRows, 1).setValues(blindOut);
    spSheet.getRange(2, rlCol,    dataRows, 1).setValues(rlOut);
    spSheet.getRange(2, rhCol,    dataRows, 1).setValues(rhOut);
    SpreadsheetApp.flush();

    // Re-sync Mail Ready tabs if they already have data
    let mailReadyMsg = '';
    const ownerCol  = mapCol(spHeaders, SP_HEADERS.OWNER_NAME) || 1;
    const validRows = allData.filter(r => r[ownerCol - 1] !== '' && r[ownerCol - 1] !== null);

    const blindSheet = ss.getSheetByName('Blind Offer Mail Ready ✉️');
    if (blindSheet && blindSheet.getLastRow() > 1) {
      const colMap = BLIND_HEADERS.map(h => { const c = flexCol(spHeaders, h); return c ? c - 1 : null; });
      blindSheet.getRange(2, 1, blindSheet.getLastRow() - 1, BLIND_HEADERS.length).clearContent();
      const out = validRows.map(row => colMap.map(i => i !== null ? row[i] : ''));
      if (out.length > 0) {
        blindSheet.getRange(2, 1, out.length, BLIND_HEADERS.length).setValues(out);
        blindSheet.getRange(2, 13, out.length, 1).setNumberFormat('0.#');
      }
      mailReadyMsg += '\n• Blind Offer Mail Ready ✉️ updated (' + out.length + ' rows)';
    }

    const rangeSheet = ss.getSheetByName('Range Offer Mail Ready ✉️');
    if (rangeSheet && rangeSheet.getLastRow() > 1) {
      const colMap = RANGE_HEADERS.map(h => { const c = flexCol(spHeaders, h); return c ? c - 1 : null; });
      rangeSheet.getRange(2, 1, rangeSheet.getLastRow() - 1, RANGE_HEADERS.length).clearContent();
      const out = validRows.map(row => colMap.map(i => i !== null ? row[i] : ''));
      if (out.length > 0) {
        rangeSheet.getRange(2, 1, out.length, RANGE_HEADERS.length).setValues(out);
        rangeSheet.getRange(2, 13, out.length, 1).setNumberFormat('0.#');
      }
      mailReadyMsg += '\n• Range Offer Mail Ready ✉️ updated (' + out.length + ' rows)';
    }

    if (!silent) SpreadsheetApp.getUi().alert('✓ Offer prices refreshed in Scrubbed and Priced!' + mailReadyMsg);
  } catch (e) {
    SpreadsheetApp.getUi().alert('✗ Error: ' + e.message);
  }
}

/** Build Owner Name(s) from individual name parts.
 *  Rules:
 *  1. Only Owner 1 exists                          → "Owner1First Owner1Last"
 *  2. Owner 1 and Owner 2 share the same last name → "Owner1First and Owner2First SharedLast"
 *  3. Owner 1 and Owner 2 have different last names → "Owner1First Owner1Last and Owner2First Owner2Last"
 *  4. Owner 1 is blank but Owner 2 has values      → "Owner2First Owner2Last"
 *  5. All first/last names are blank               → fall back to Owner 1 Full Name
 */
function buildOwnerName(o1First, o1Last, o2First, o2Last, o1Full) {
  const f1   = String(o1First || '').trim();
  const l1   = String(o1Last  || '').trim();
  const f2   = String(o2First || '').trim();
  const l2   = String(o2Last  || '').trim();
  const full = String(o1Full  || '').trim();
  const has1 = f1 || l1;
  const has2 = f2 || l2;

  if (!has1 && !has2) return full;                                               // Rule 5
  if (has1 && !has2)  return [f1, l1].filter(Boolean).join(' ');                 // Rule 1
  if (!has1 && has2)  return [f2, l2].filter(Boolean).join(' ');                 // Rule 4
  if (l1 && l2 && l1.toLowerCase() === l2.toLowerCase())
    return [f1, 'and', f2, l1].filter(Boolean).join(' ');                        // Rule 2
  const name1 = [f1, l1].filter(Boolean).join(' ');
  const name2 = [f2, l2].filter(Boolean).join(' ');
  return [name1, 'and', name2].filter(Boolean).join(' ');                        // Rule 3
}

/**
 * Convert an address string to proper case following USPS conventions.
 * - Directionals (N, S, E, W, NW, NE, SW, SE) stay fully uppercase
 * - Common street suffixes are title-cased (St, Rd, Ave, etc.)
 * - All other words are title-cased
 */
function toProperAddress(str) {
  if (!str) return str;
  const UPPERCASE_WORDS = new Set([
    'N','S','E','W','NW','NE','SW','SE','NNW','NNE','SSW','SSE',
  ]);
  const TITLE_CASE_SUFFIXES = new Set([
    'St','Rd','Ave','Blvd','Dr','Ln','Ct','Pl','Cir','Hwy','Pkwy',
    'Way','Ter','Trl','Run','Loop','Path','Pass','Pike','Xing',
  ]);
  return String(str).trim().toLowerCase().split(/\s+/).map(word => {
    const upper = word.toUpperCase();
    if (UPPERCASE_WORDS.has(upper)) return upper;
    const titled = word.charAt(0).toUpperCase() + word.slice(1);
    // Check if this matches a known suffix (case-insensitive)
    for (const suffix of TITLE_CASE_SUFFIXES) {
      if (suffix.toLowerCase() === word) return suffix;
    }
    return titled;
  }).join(' ');
}

/** STEP 1a: Show filter modal before loading data. */
function populateScrubbed() {
  if (!_requireAuth()) return;
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('LI Raw Dataset');
  if (!rawSheet || rawSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No data found in LI Raw Dataset. Paste your data first.');
    return;
  }

  // Load saved filter prefs, default all included
  const props    = PropertiesService.getScriptProperties();
  const savedKey = 'filterPrefs_' + ss.getId();
  let prefs;
  try { prefs = JSON.parse(props.getProperty(savedKey) || 'null'); } catch(e) { prefs = null; }
  if (!prefs) prefs = { landlocked: true, nonLandlocked: true, individual: true, corporate: true };

  const html = HtmlService.createHtmlOutput(getFilterModalHtml_(prefs))
    .setWidth(460)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'Load Data — Filter Options');
}

/** STEP 1b: Called from filter modal — applies filters and loads data. */
function applyFiltersAndLoad(prefs) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('LI Raw Dataset');
  const spSheet  = ss.getSheetByName('Scrubbed and Priced');
  const kwSheet  = ss.getSheetByName('Keyword Scrub List');

  // Save prefs for next time
  PropertiesService.getScriptProperties()
    .setProperty('filterPrefs_' + ss.getId(), JSON.stringify(prefs));

  try {
    const rawHeaders = getHeaderMap(rawSheet);
    const spHeaders  = getHeaderMap(spSheet);

    // Raw filter columns (0-based)
    const rawLandlockedCol = (mapCol(rawHeaders, 'land locked')  || 0) - 1;
    const rawOwnerTypeCol  = (mapCol(rawHeaders, 'owner type')   || 0) - 1;

    // Build rawColIndex → spColIndex pairs (0-based)
    const colMap = [];
    for (const [rawH, spH] of Object.entries(RAW_TO_SP_HEADERS)) {
      const rc = mapCol(rawHeaders, rawH), sc = mapCol(spHeaders, spH);
      if (rc && sc) colMap.push([rc - 1, sc - 1]);
    }

    const polygonCol  = requireCol(spHeaders, SP_HEADERS.POLYGON,     'Scrubbed and Priced');
    const acreageCol  = requireCol(spHeaders, SP_HEADERS.ACREAGE,     'Scrubbed and Priced');
    const liAcreageCol = mapCol(spHeaders, 'LI Calculated Acreage');
    const varPctCol   = mapCol(spHeaders, 'Acreage Variance %');
    const manPPACol   = requireCol(spHeaders, SP_HEADERS.MANUAL_PPA,  'Scrubbed and Priced');
    const manMVCol    = requireCol(spHeaders, SP_HEADERS.MANUAL_MV,   'Scrubbed and Priced');
    const blindCol    = requireCol(spHeaders, SP_HEADERS.BLIND_OFFER, 'Scrubbed and Priced');
    const rlCol       = prefixCol(spHeaders, 'Range Offer Low')  || requireCol(spHeaders, SP_HEADERS.RANGE_LOW,  'Scrubbed and Priced');
    const rhCol       = prefixCol(spHeaders, 'Range Offer High') || requireCol(spHeaders, SP_HEADERS.RANGE_HIGH, 'Scrubbed and Priced');
    const ownerCol    = requireCol(spHeaders, SP_HEADERS.OWNER_NAME,  'Scrubbed and Priced');

    const rawO1First = (mapCol(rawHeaders, 'Owner 1 First Name') || 0) - 1;
    const rawO1Last  = (mapCol(rawHeaders, 'Owner 1 Last Name')  || 0) - 1;
    const rawO2First = (mapCol(rawHeaders, 'Owner 2 First Name') || 0) - 1;
    const rawO2Last  = (mapCol(rawHeaders, 'Owner 2 Last Name')  || 0) - 1;
    const rawO1Full  = (mapCol(rawHeaders, 'Owner 1 Full Name')  || 0) - 1;

    const totalCols = spSheet.getLastColumn();
    const { blind: blindTiers, range: rangeTiers, polygonPPA } = getPricingSettings();

    // Find tier columns by actual current percentages from Pricing Settings
    const tierCols = blindTiers.map(tier => {
      const pct = Math.round(tier.mult * 100) + '%';
      return prefixCol(spHeaders, pct + ' Offer') || null;
    });

    const keywords   = getKeywords(kwSheet);
    const lastRawRow = rawSheet.getLastRow();
    const rawData    = rawSheet.getRange(2, 1, lastRawRow - 1, rawSheet.getLastColumn()).getValues();

    // Clear SP data rows (keep header)
    const spLastRow = spSheet.getLastRow();
    if (spLastRow > 1) {
      spSheet.getRange(2, 1, spLastRow - 1, totalCols).clearContent();
    }

    const outputRows = [];
    for (const raw of rawData) {
      if (raw.every(c => c === '' || c === null)) continue;

      // Apply filters
      const landlocked = rawLandlockedCol >= 0 ? String(raw[rawLandlockedCol] || '').trim().toUpperCase() : '';
      const ownerType  = rawOwnerTypeCol  >= 0 ? String(raw[rawOwnerTypeCol]  || '').trim().toLowerCase() : '';
      if (landlocked === 'Y' && !prefs.landlocked)    continue;
      if (landlocked !== 'Y' && !prefs.nonLandlocked) continue;
      if ((ownerType === 'individual' || ownerType === 'private') && !prefs.individual) continue;
      if ((ownerType === 'corporation' || ownerType === 'corporate') && !prefs.corporate) continue;

      const row = new Array(totalCols).fill('');
      for (const [ri, si] of colMap) row[si] = raw[ri] ?? '';

      // Convert address fields to proper case
      const addrFields = ['Mail Address', 'Mail City', 'Parcel Address', 'Parcel City'];
      addrFields.forEach(h => {
        const c = mapCol(spHeaders, h);
        if (c) row[c - 1] = toProperAddress(row[c - 1]);
      });

      row[ownerCol - 1] = buildOwnerName(
        rawO1First >= 0 ? raw[rawO1First] : '',
        rawO1Last  >= 0 ? raw[rawO1Last]  : '',
        rawO2First >= 0 ? raw[rawO2First] : '',
        rawO2Last  >= 0 ? raw[rawO2Last]  : '',
        rawO1Full  >= 0 ? raw[rawO1Full]  : ''
      );

      // Acreage Variance %
      const rawAcreage = Math.round((Number(row[acreageCol - 1])  || 0) * 10) / 10;
      const liAcreage  = liAcreageCol ? Math.round((Number(row[liAcreageCol - 1]) || 0) * 10) / 10 : 0;
      let variance = -1;
      if (rawAcreage > 0 && liAcreage > 0) {
        variance = Math.abs(rawAcreage - liAcreage) / Math.max(rawAcreage, liAcreage);
        if (varPctCol) row[varPctCol - 1] = Math.round(variance * 100);
      }

      // Pricing — use LI Calculated Acreage; fall back to Acreage if LI cell is blank
      const liRawVal       = liAcreageCol ? row[liAcreageCol - 1] : '';
      const pricingAcreage = (liRawVal !== '' && liRawVal !== null) ? liAcreage : rawAcreage;
      const polygon        = row[polygonCol - 1];
      const ppa            = lookupPPA(polygonPPA, polygon, pricingAcreage);
      const mv             = ppa > 0 && pricingAcreage > 0 ? ppa * pricingAcreage : 0;
      row[manPPACol - 1] = ppa > 0 ? ppa : '';
      row[manMVCol  - 1] = mv  > 0 ? mv  : '';

      let maxTier = '';
      tierCols.forEach((col, t) => {
        if (!col) return;
        const tier = blindTiers[t];
        const val  = (tier && mv >= tier.min && mv < tier.max) ? mv * tier.mult : '';
        row[col - 1] = val;
        if (val !== '' && (maxTier === '' || val > maxTier)) maxTier = val;
      });
      row[blindCol - 1] = maxTier;
      row[rlCol    - 1] = mv > 0 ? mv * rangeTiers[0].mult : '';
      row[rhCol    - 1] = mv > 0 ? mv * rangeTiers[1].mult : '';

      outputRows.push(row);
    }

    if (outputRows.length === 0) {
      SpreadsheetApp.getUi().alert('No rows match the selected filters.');
      return;
    }

    spSheet.getRange(2, 1, outputRows.length, totalCols).setValues(outputRows);
    SpreadsheetApp.flush();

    // Format Acreage Variance % column
    if (varPctCol) {
      spSheet.getRange(2, varPctCol, outputRows.length, 1).setNumberFormat('0"%"');
    }

    // Format Acreage and LI Calculated Acreage as numbers (0.# hides trailing .0)
    if (acreageCol)   spSheet.getRange(2, acreageCol,   outputRows.length, 1).setNumberFormat('0.#');
    if (liAcreageCol) spSheet.getRange(2, liAcreageCol, outputRows.length, 1).setNumberFormat('0.#');

    applySheetFormatting_(spSheet, outputRows.length, totalCols);
    highlightKeywords(spSheet, keywords, outputRows.length, ownerCol);

    SpreadsheetApp.getUi().alert('✓ Loaded ' + outputRows.length + ' properties into Scrubbed & Priced.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('✗ Error: ' + e.message);
  }
}

/** Generate filter modal HTML. */
function getFilterModalHtml_(prefs) {
  const chk = (val) => val ? 'checked' : '';
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  :root{--blue:#4a6e93;--border:#dde1e9;--bg:#f6f7f9;--panel:#fff;--text:#1a2332;--muted:#6b7d95;
    --font:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:var(--font);background:var(--bg);color:var(--text);padding:20px;}
  h3{font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
     color:var(--muted);margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--border);}
  .group{background:var(--panel);border:1px solid var(--border);border-radius:8px;padding:14px 16px;margin-bottom:14px;}
  label{display:flex;align-items:center;gap:10px;font-size:12px;padding:5px 0;cursor:pointer;}
  input[type=checkbox]{width:15px;height:15px;accent-color:var(--blue);cursor:pointer;flex-shrink:0;}
  .footer{display:flex;justify-content:flex-end;gap:10px;padding-top:4px;}
  button{padding:8px 22px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;
         font-family:var(--font);border:1px solid var(--border);}
  .btn-cancel{background:var(--panel);color:var(--muted);}
  .btn-apply{background:var(--blue);color:#fff;border-color:var(--blue);}
</style>
</head>
<body>
  <div class="group">
    <h3>Landlocked Status</h3>
    <label><input type="checkbox" id="nonLandlocked" ${chk(prefs.nonLandlocked)}> Include Non-Landlocked Properties</label>
    <label><input type="checkbox" id="landlocked"    ${chk(prefs.landlocked)}>    Include Landlocked Properties</label>
  </div>
  <div class="group">
    <h3>Owner Type</h3>
    <label><input type="checkbox" id="individual" ${chk(prefs.individual)}> Include Individually Owned</label>
    <label><input type="checkbox" id="corporate"  ${chk(prefs.corporate)}>  Include Corporate Owned</label>
  </div>
  <div class="footer">
    <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
    <button class="btn-apply"  onclick="apply()">Apply & Load →</button>
  </div>
<script>
function apply() {
  const prefs = {
    nonLandlocked : document.getElementById('nonLandlocked').checked,
    landlocked    : document.getElementById('landlocked').checked,
    individual    : document.getElementById('individual').checked,
    corporate     : document.getElementById('corporate').checked,
  };
  google.script.run
    .withSuccessHandler(() => google.script.host.close())
    .withFailureHandler(e => { alert('Error: ' + e.message); })
    .applyFiltersAndLoad(prefs);
}
<\/script>
</body>
</html>`;
}

function getKeywords(kwSheet) {
  if (!kwSheet) return [];
  const lastRow = kwSheet.getLastRow();
  if (lastRow < 2) return [];
  return kwSheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim().toUpperCase()).filter(k => k !== '');
}

/** Highlight keyword matches in Owner Name column. */
function highlightKeywords(spSheet, keywords, dataRows, ownerCol) {
  if (dataRows === 0 || keywords.length === 0) return;
  ownerCol = ownerCol || 1;
  const nameRange = spSheet.getRange(2, ownerCol, dataRows, 1);
  const names = nameRange.getValues();
  const bg = [], fc = [], fw = [];
  for (const [name] of names) {
    const upper = String(name).toUpperCase();
    const hit   = keywords.some(kw => kw && new RegExp('\\b' + kw.replace(/[.*+?^$()|[\]{}]/g, '$&') + '\\b').test(upper));
    bg.push([hit ? '#FFE066' : null]);
    fc.push([hit ? '#7B4F00' : null]);
    fw.push([hit ? 'bold'    : 'normal']);
  }
  nameRange.setBackgrounds(bg).setFontColors(fc).setFontWeights(fw);
}

/** Remove all flagged (yellow) rows. */
function removeFlaggedRows() {
  if (!_requireAuth()) return;
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const spSheet = ss.getSheetByName('Scrubbed and Priced');
  const lastRow = spSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data in Scrubbed and Priced.'); return; }

  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Remove flagged rows?',
               'This will permanently delete all yellow-highlighted rows. Continue?',
               ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    const ownerCol = (mapCol(getHeaderMap(spSheet), SP_HEADERS.OWNER_NAME) || 1);
    const bgs      = spSheet.getRange(2, ownerCol, lastRow - 1, 1).getBackgrounds();
    let count = 0;
    for (let i = bgs.length - 1; i >= 0; i--) {
      const bg = (bgs[i][0] || '').toLowerCase();
      if (bg === '#ffe066' || bg === '#ffff00') { spSheet.deleteRow(i + 2); count++; }
    }
    ui.alert(count === 0
      ? 'No flagged rows found. Run "Load data" first to apply highlighting.'
      : '✓ Done! Removed ' + count + ' flagged row(s).');
  } catch (e) { ui.alert('✗ Error: ' + e.message); }
}

/** Remove all Low Seller IQ rows. */
function removeLowSellerIQ() {
  if (!_requireAuth()) return;
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const spSheet = ss.getSheetByName('Scrubbed and Priced');
  const lastRow = spSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data in Scrubbed and Priced.'); return; }

  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Remove Low likelihood sellers?',
               'This will permanently delete all rows where Seller IQ = Low. Continue?',
               ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    const iqCol = requireCol(getHeaderMap(spSheet), SP_HEADERS.SELLER_IQ, 'Scrubbed and Priced');
    const vals  = spSheet.getRange(2, iqCol, lastRow - 1, 1).getValues();
    let count = 0;
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim().toLowerCase() === 'low') { spSheet.deleteRow(i + 2); count++; }
    }
    ui.alert(count === 0
      ? 'No rows with Seller IQ = Low found.'
      : '✓ Done! Removed ' + count + ' Low likelihood seller row(s).');
  } catch (e) { ui.alert('✗ Error: ' + e.message); }
}

/** Sync Blind Offer Mail Ready ✉️ Tab. */
function syncBlindOfferTab() {
  if (!_requireAuth()) return;
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const spSheet    = ss.getSheetByName('Scrubbed and Priced');
  const blindSheet = ss.getSheetByName('Blind Offer Mail Ready ✉️');
  if (!blindSheet) { SpreadsheetApp.getUi().alert('Could not find "Blind Offer Mail Ready ✉️" tab.'); return; }
  const lastRow = spSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data in Scrubbed and Priced to sync.'); return; }

  try {
    const spHeaders  = getHeaderMap(spSheet);
    const ownerCol   = mapCol(spHeaders, SP_HEADERS.OWNER_NAME) || 1;
    const dataRows   = lastRow - 1;
    const spData     = spSheet.getRange(2, 1, dataRows, spSheet.getLastColumn()).getValues();
    const validRows  = spData.filter(r => r[ownerCol - 1] !== '' && r[ownerCol - 1] !== null);
    const colMap     = BLIND_HEADERS.map(h => { const c = flexCol(spHeaders, h); return c ? c - 1 : null; });
    const blindLastRow = blindSheet.getLastRow();
    if (blindLastRow > 1) blindSheet.getRange(2, 1, blindLastRow - 1, BLIND_HEADERS.length).clearContent();
    const out = validRows.map(row => colMap.map(i => i !== null ? row[i] : ''));
    if (out.length > 0) {
      blindSheet.getRange(2, 1, out.length, BLIND_HEADERS.length).setValues(out);
      blindSheet.getRange(2, 13, out.length, 1).setNumberFormat('0.#');
    }
    SpreadsheetApp.getUi().alert('✓ Blind Offer tab synced!\n• ' + validRows.length + ' rows pushed to Blind Offer Mail Ready ✉️.');
  } catch (e) { SpreadsheetApp.getUi().alert('✗ Error: ' + e.message); }
}

/** Sync Range Offer Mail Ready ✉️ Tab. */
function syncRangeOfferTab() {
  if (!_requireAuth()) return;
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const spSheet    = ss.getSheetByName('Scrubbed and Priced');
  const rangeSheet = ss.getSheetByName('Range Offer Mail Ready ✉️');
  if (!rangeSheet) { SpreadsheetApp.getUi().alert('Could not find "Range Offer Mail Ready ✉️" tab.'); return; }
  const lastRow = spSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data in Scrubbed and Priced to sync.'); return; }

  try {
    const spHeaders  = getHeaderMap(spSheet);
    const ownerCol   = mapCol(spHeaders, SP_HEADERS.OWNER_NAME) || 1;
    const dataRows   = lastRow - 1;
    const spData     = spSheet.getRange(2, 1, dataRows, spSheet.getLastColumn()).getValues();
    const validRows  = spData.filter(r => r[ownerCol - 1] !== '' && r[ownerCol - 1] !== null);
    const colMap     = RANGE_HEADERS.map(h => { const c = flexCol(spHeaders, h); return c ? c - 1 : null; });
    const rangeLastRow = rangeSheet.getLastRow();
    if (rangeLastRow > 1) rangeSheet.getRange(2, 1, rangeLastRow - 1, RANGE_HEADERS.length).clearContent();
    const out = validRows.map(row => colMap.map(i => i !== null ? row[i] : ''));
    if (out.length > 0) {
      rangeSheet.getRange(2, 1, out.length, RANGE_HEADERS.length).setValues(out);
      rangeSheet.getRange(2, 13, out.length, 1).setNumberFormat('0.#');
    }
    SpreadsheetApp.getUi().alert('✓ Range Offer tab synced!\n• ' + validRows.length + ' rows pushed to Range Offer Mail Ready ✉️.');
  } catch (e) { SpreadsheetApp.getUi().alert('✗ Error: ' + e.message); }
}


// ============================================================
//  EDIT TRIGGERS + STYLING
// ============================================================

/**
 * Reads tier multipliers from Pricing Settings and updates:
 *  - Column F labels (F4:F8) e.g. "U — 0.52 (52%)"
 *  - SP tier column headers e.g. "52% Offer Price"
 *  - SP range column headers e.g. "Range Offer Low (52%)"
 *  - SP_HEADERS constants and RANGE_HEADERS are updated in memory for this run only;
 *    the sheet headers are the source of truth after this writes them.
 */
function updatePricingHeaders_() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const pricing = ss.getSheetByName('Pricing Settings');
  const sp      = ss.getSheetByName('Scrubbed and Priced');
  if (!pricing || !sp) return;

  // Read G4:G8 (blind tier multipliers) and L4:L5 (range multipliers)
  const tierVals  = pricing.getRange('G4:G8').getValues().map(r => Number(r[0]) || 0);
  const rangeVals = pricing.getRange('L4:L5').getValues().map(r => Number(r[0]) || 0);

  // Tier row labels in Pricing Settings col F: "U — 0.52 (52%)"
  const tierLetters = ['U', 'V', 'W', 'X', 'Y'];
  const fLabels = tierVals.map((v, i) => {
    if (!v) return '';
    const dec = v > 1 ? (v / 100).toFixed(2) : v.toFixed(2);
    const pct = Math.round(v > 1 ? v : v * 100);
    return tierLetters[i] + ' \u2014 ' + dec + ' (' + pct + '%)';
  });
  pricing.getRange('F4:F8').setValues(fLabels.map(l => [l]));

  // SP tier column headers (columns U–Y = cols 21–25)
  const spTierCols   = [21, 22, 23, 24, 25];
  const spTierLabels = tierVals.map(v => {
    if (!v) return null;
    return Math.round(v > 1 ? v : v * 100) + '% Offer Price';
  });
  // Batch all SP header writes (U-AB) in one range operation, then set wrap once
  const spHeaders8 = sp.getRange(1, 21, 1, 8).getValues()[0];
  spTierCols.forEach((col, i) => {
    if (spTierLabels[i]) spHeaders8[col - 21] = spTierLabels[i];
  });
  const rlPct = rangeVals[0] ? Math.round(rangeVals[0] > 1 ? rangeVals[0] : rangeVals[0] * 100) : null;
  const rhPct = rangeVals[1] ? Math.round(rangeVals[1] > 1 ? rangeVals[1] : rangeVals[1] * 100) : null;
  if (rlPct) spHeaders8[6] = 'Range Offer Low ('  + rlPct + '%)';
  if (rhPct) spHeaders8[7] = 'Range Offer High (' + rhPct + '%)';
  sp.getRange(1, 21, 1, 8).setValues([spHeaders8]);

  // Recalculate all offer prices using the updated percentages
  if (sp.getLastRow() > 1) refreshOfferPrices(true);
}


/**
 * onEdit trigger:
 *  - Re-applies LandValuator styling after a paste into LI Raw Dataset
 *  - Auto-uppercases County Zone in Pricing Settings
 *  - Updates SP pricing headers when Pricing Settings multipliers change
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const name  = sheet.getName();
  const col   = e.range.getColumn();
  const row   = e.range.getRow();

  // ── Re-apply styling after paste into LI Raw Dataset ──
  if (name === 'LI Raw Dataset') {
    const numRows = e.range.getNumRows();
    const numCols = e.range.getNumColumns();
    if ((numRows > 1 || numCols > 1) && row <= 1) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) applySheetFormatting_(sheet, lastRow > 1 ? lastRow - 1 : 0, lastCol);
      return;
    }
  }

  // ── Auto-capitalize County Zone in Pricing Settings (col A, skip rows 1-3) ──
  if (name === 'Pricing Settings' && col === 1) {
    if (row <= 3) return;
    const val = e.range.getValue();
    if (val) e.range.setValue(String(val).toUpperCase().trim());
    return;
  }

  // ── Dynamic pricing header updates when Pricing Settings G or L columns change ──
  if (name === 'Pricing Settings') {
    // G4:G8 = blind tier multipliers, L4:L5 = range offer multipliers
    if ((col === 7 && row >= 4 && row <= 8) || (col === 12 && row >= 4 && row <= 5)) {
      updatePricingHeaders_();
    }
  }
}

/**
 * Apply standard LandValuator formatting to a sheet.
 * Header row: #4a6e93 bg, white font, size 11, 45px height.
 * Data rows: alternating white/#eef0f4, font size 11, 20px height.
 * numDataRows: if provided, only formats that many data rows (faster after load).
 *              If omitted, formats all rows currently in the sheet.
 */
function applySheetFormatting_(sheet, numDataRows, numCols) {
  const lastCol = numCols || sheet.getLastColumn();
  if (lastCol < 1) return;

  // Header row — style only, never touch banding
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#4a6e93')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 45);


  // Data rows — font/size/height only, never set backgrounds (let native banding handle colors)
  const dataRows = numDataRows || Math.max(sheet.getLastRow() - 1, 0);
  if (dataRows < 1) return;
  sheet.getRange(2, 1, dataRows, lastCol)
    .setFontColor('#1a2332')
    .setFontWeight('normal')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  if (dataRows > 0) sheet.setRowHeightsForced(2, dataRows, 26);
}


// ============================================================
//  PHASE 4 — MARGIN REFERENCE CHART (HTML Modal)
// ============================================================

/** Show the Margin Reference Chart modal. */
function showMarginReferenceChart() {
  if (!_requireAuth()) return;
  const html = HtmlService.createHtmlOutput(getMarginChartHtml())
    .setWidth(1060)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Margin Reference Chart');
}

function getMarginChartHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Margin Reference Chart</title>
<style>
  :root {
    --bg:#f6f7f9;--panel:#ffffff;--panel2:#eef0f4;--border:#dde1e9;
    --accent:#5b7fa6;--blue:#4a6e93;--accent-light:#edf2f8;
    --text:#1a2332;--muted:#6b7d95;
    --font:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);font-family:var(--font);color:var(--text);display:flex;flex-direction:column;height:100vh;overflow:hidden;}
  .titlebar{background:var(--blue);padding:12px 18px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;}
  .title-group{display:flex;align-items:baseline;gap:12px;}
  .title{font-size:14px;font-weight:700;color:#fff;letter-spacing:.03em;}
  .subtitle{font-size:10px;color:rgba(255,255,255,.6);letter-spacing:.08em;text-transform:uppercase;}
  .close-x{width:26px;height:26px;border-radius:50%;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;}
  .body{padding:14px 18px 16px;display:flex;flex-direction:column;flex:1;min-height:0;}
  .inputs-row{display:grid;grid-template-columns:1fr auto 1fr auto 1fr auto 1fr;gap:0;margin-bottom:10px;background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:10px 16px;align-items:center;flex-shrink:0;}
  .input-group{display:flex;flex-direction:column;gap:6px;align-items:center;}
  .input-label{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);text-align:center;line-height:1.4;}
  .input-pill{display:inline-flex;align-items:center;gap:2px;background:var(--panel2);border:1px solid var(--border);border-radius:20px;padding:5px 10px;font-size:13px;font-weight:700;color:var(--text);}
  .input-pill:focus-within{border-color:var(--accent);}
  .affix{color:var(--muted);font-size:12px;font-weight:600;}
  .input-pill input{background:none;border:none;outline:none;font-family:var(--font);font-size:13px;font-weight:700;color:var(--text);}
  input.w60{width:60px;} input.w40{width:40px;} input.w24{width:24px;text-align:right;}
  .divider{width:1px;height:38px;background:var(--border);margin:0 2px;flex-shrink:0;}
  .input-note{font-size:10px;color:var(--muted);margin-left:6px;font-style:italic;white-space:nowrap;}
  .tiers-row{display:flex;gap:6px;margin-bottom:10px;background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:8px 14px;align-items:center;flex-shrink:0;flex-wrap:wrap;}
  .tier-group{display:flex;flex-direction:column;gap:3px;align-items:center;}
  .tier-label{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);}
  .tier-pill{display:inline-flex;align-items:center;gap:1px;background:var(--panel2);border:1.5px solid var(--border);border-radius:20px;padding:4px 9px;}
  .tier-pill:focus-within{border-color:var(--accent);}
  .tier-pill input{background:none;border:none;outline:none;font-family:var(--font);font-size:12px;font-weight:700;width:24px;text-align:right;color:var(--text);}
  .tiers-note{font-size:10px;color:var(--muted);font-style:italic;margin-left:auto;white-space:nowrap;}
  .legend{display:flex;gap:10px;margin-bottom:10px;flex-wrap:wrap;align-items:center;background:var(--panel);border:1px solid var(--border);border-radius:8px;padding:7px 12px;flex-shrink:0;}
  .legend-item{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--text);}
  .legend-dot{width:10px;height:10px;border-radius:2px;flex-shrink:0;}
  .legend-note{margin-left:auto;font-size:10px;color:var(--muted);font-style:italic;}
  .chart-wrap{background:var(--panel);border:1px solid var(--border);border-radius:8px;flex:1;position:relative;min-height:0;overflow:hidden;}
  canvas{position:absolute;top:0;left:0;}
  .tooltip{position:fixed;background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:12px 15px;box-shadow:0 8px 28px rgba(44,82,130,.16);pointer-events:none;opacity:0;transition:opacity .12s;min-width:210px;z-index:100;}
  .tooltip.visible{opacity:1;}
  .tt-header{display:flex;align-items:baseline;justify-content:space-between;padding-bottom:8px;margin-bottom:8px;border-bottom:1px solid var(--border);}
  .tt-price{font-size:14px;font-weight:800;color:var(--blue);}
  .tt-label{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);}
  .tt-row{display:flex;align-items:center;justify-content:space-between;padding:3px 0;gap:14px;}
  .tt-name{display:flex;align-items:center;gap:7px;font-size:11px;color:var(--muted);white-space:nowrap;}
  .tt-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0;}
  .tt-right{text-align:right;}
  .tt-val{font-size:12px;font-weight:700;color:var(--text);}
  .tt-sub{font-size:10px;color:var(--muted);}
  .x-axis-name{text-align:center;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);padding:6px 0 0;flex-shrink:0;}
  .footer{display:flex;align-items:center;justify-content:space-between;padding-top:10px;border-top:1px solid var(--border);flex-shrink:0;flex-wrap:wrap;gap:8px;margin-top:10px;}
  .footer-note{font-size:10px;color:var(--muted);font-style:italic;}
  .close-btn{padding:7px 22px;border-radius:20px;border:1px solid var(--border);background:var(--panel);color:var(--muted);font-size:11px;font-weight:700;cursor:pointer;letter-spacing:.04em;font-family:var(--font);}
  .close-btn:hover{border-color:var(--accent);color:var(--accent);}
</style>
</head>
<body>
<div class="titlebar">
  <div class="title-group">
    <span class="title">Margin Reference Chart</span>
    <span class="subtitle">Profit at Purchase · Editable</span>
  </div>
</div>
<div class="body">

  <div class="inputs-row">
    <div class="input-group">
      <div class="input-label">Closing<br>Costs</div>
      <div class="input-pill"><span class="affix">$</span><input class="w60" id="closing" type="text" value="1200" oninput="draw()"/></div>
    </div>
    <div class="divider"></div>
    <div class="input-group">
      <div class="input-label">Realtor<br>Commission</div>
      <div class="input-pill"><input class="w24" id="commission" type="text" value="3" oninput="draw()"/><span class="affix">%</span></div>
    </div>
    <div class="divider"></div>
    <div class="input-group">
      <div class="input-label">Cost of<br>Capital</div>
      <div class="input-pill"><span class="affix">$</span><input class="w60" id="capital" type="text" value="4000" oninput="draw()"/></div>
    </div>
    <div class="divider"></div>
    <div class="input-group">
      <div class="input-label">Value Add /<br>Improvements</div>
      <div class="input-pill"><span class="affix">$</span><input class="w60" id="valueadd" type="text" value="0" oninput="draw()"/></div>
    </div>
  </div>

  <div class="tiers-row" id="tiersRow">
    <span class="tiers-note">Edit buy % tiers · chart only</span>
  </div>

  <div class="legend" id="legend"></div>

  <div class="chart-wrap" id="chartWrap">
    <canvas id="chart"></canvas>
    <div class="tooltip" id="tooltip">
      <div class="tt-header">
        <span class="tt-price" id="tt-price"></span>
        <span class="tt-label">Purchase Price</span>
      </div>
      <div id="tt-rows"></div>
    </div>
  </div>

  <div class="x-axis-name">Purchase Price</div>
  <div class="footer">
    <span class="footer-note">Hover over the chart to see profit margin potential. Changes to costs or offer percentages are for reference only, and do not affect app or spreadsheet pricing.</span>
    <button class="close-btn" onclick="google.script.host.close()">Close</button>
  </div>
</div>

<script>
const COLORS = ['#6ecf8a','#4db870','#3a9e5f','#2a7a48','#1a5a32'];
const DASHES = [[],[6,3],[6,3],[6,3],[]];
const WIDTHS = [2.5,1.8,1.8,1.8,2.5];
const PRICES = [];
for (let p = 25000; p <= 250000; p += 2500) PRICES.push(p);

let tiers = [50, 55, 60, 65, 70];

const fmt = v => {
  const abs = Math.abs(Math.round(v));
  const s = '$' + abs.toLocaleString();
  return v < 0 ? '(' + s + ')' : s;
};

function getInputs() {
  const n = id => parseFloat(String(document.getElementById(id).value).replace(/,/g,'')) || 0;
  return { closing: n('closing'), capital: n('capital'), valueadd: n('valueadd'), commission: n('commission') / 100 };
}

function netProfit(mv, buyPct, inp) {
  return mv - (mv * buyPct) - inp.closing - inp.capital - inp.valueadd - (mv * inp.commission);
}

function buildTierInputs() {
  const row  = document.getElementById('tiersRow');
  const note = row.querySelector('.tiers-note');
  row.querySelectorAll('.tier-group, .tier-divider').forEach(el => el.remove());
  tiers.forEach((pct, i) => {
    if (i > 0) {
      const d = document.createElement('div');
      d.className = 'tier-divider';
      d.style.cssText = 'width:1px;height:32px;background:var(--border);margin:0 2px;flex-shrink:0;';
      row.insertBefore(d, note);
    }
    const g = document.createElement('div');
    g.className = 'tier-group';
    g.innerHTML = \`<div class="tier-label" style="color:\${COLORS[i]}">Tier \${i+1}</div>
      <div class="tier-pill" style="border-color:\${COLORS[i]}60">
        <input type="text" value="\${pct}" style="color:\${COLORS[i]}" oninput="tierChanged(\${i},this.value)"/>
        <span class="affix" style="color:\${COLORS[i]}">%</span>
      </div>\`;
    row.insertBefore(g, note);
  });
}

function buildLegend() {
  document.getElementById('legend').innerHTML =
    tiers.map((pct, i) => \`<div class="legend-item"><div class="legend-dot" style="background:\${COLORS[i]}"></div>Buy at \${pct}%</div>\`).join('') +
    '<span class="legend-note">All deductions applied</span>';
}

function tierChanged(i, val) {
  const v = parseFloat(val);
  if (!isNaN(v) && v > 0 && v < 100) { tiers[i] = v; buildLegend(); draw(); }
}

const canvas  = document.getElementById('chart');
const ctx     = canvas.getContext('2d');
const wrap    = document.getElementById('chartWrap');
const tooltip = document.getElementById('tooltip');
const PAD     = { l:70, r:52, t:14, b:26 };
let W, H;

function resize() {
  const dpr = window.devicePixelRatio || 1;
  W = wrap.clientWidth; H = wrap.clientHeight;
  canvas.width = W * dpr; canvas.height = H * dpr;
  canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
  ctx.scale(dpr, dpr); draw();
}

function xPos(p) { return PAD.l + (p - 25000) / (250000 - 25000) * (W - PAD.l - PAD.r); }
function yPos(v, mn, mx) { return PAD.t + (1 - (v - mn) / (mx - mn)) * (H - PAD.t - PAD.b); }

function draw() {
  if (!W || !H) return;
  const inp = getInputs();
  ctx.clearRect(0, 0, W, H);
  const allVals = tiers.flatMap(pct => PRICES.map(p => netProfit(p, pct / 100, inp)));
  const minY = Math.min(0, ...allVals), maxY = Math.max(...allVals) * 1.08;
  const step = maxY > 80000 ? 20000 : maxY > 40000 ? 10000 : 5000;
  const gridVals = [0];
  for (let v = step; v <= maxY; v += step) gridVals.push(v);
  if (minY < 0) gridVals.push(minY);
  gridVals.forEach(v => {
    const y = yPos(v, minY, maxY);
    ctx.strokeStyle = '#dde1e9'; ctx.lineWidth = 1; ctx.setLineDash([3,3]);
    ctx.beginPath(); ctx.moveTo(PAD.l, y); ctx.lineTo(W - PAD.r, y); ctx.stroke(); ctx.setLineDash([]);
    ctx.fillStyle = '#6b7d95'; ctx.font = '11px Arial'; ctx.textAlign = 'right';
    ctx.fillText(fmt(v), PAD.l - 7, y + 4);
  });
  ctx.save(); ctx.translate(14, H / 2); ctx.rotate(-Math.PI / 2);
  ctx.fillStyle = '#6b7d95'; ctx.font = 'bold 10px Arial'; ctx.textAlign = 'center';
  ctx.fillText('PROFIT MARGIN', 0, 0); ctx.restore();
  ctx.strokeStyle = '#dde1e9'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(PAD.l, PAD.t); ctx.lineTo(PAD.l, H - PAD.b); ctx.stroke();
  const xTicks = [25000,50000,75000,100000,125000,150000,175000,200000,225000,250000];
  ctx.fillStyle = '#6b7d95'; ctx.font = '10px Arial'; ctx.textAlign = 'center';
  xTicks.forEach(p => ctx.fillText('$' + (p/1000) + 'K', xPos(p), H - 4));
  tiers.forEach((pct, i) => {
    ctx.beginPath(); ctx.strokeStyle = COLORS[i]; ctx.lineWidth = WIDTHS[i]; ctx.setLineDash(DASHES[i]);
    PRICES.forEach((p, j) => {
      const x = xPos(p), y = yPos(netProfit(p, pct / 100, inp), minY, maxY);
      j === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
    });
    ctx.stroke(); ctx.setLineDash([]);
    const endY = yPos(netProfit(250000, pct / 100, inp), minY, maxY);
    const ep   = Math.round(netProfit(250000, pct / 100, inp) / 250000 * 100);
    ctx.fillStyle = COLORS[i]; ctx.font = 'bold 11px Arial'; ctx.textAlign = 'left';
    ctx.fillText(ep + '%', W - PAD.r + 4, endY + 4);
  });
}

canvas.addEventListener('mousemove', e => {
  const rect  = canvas.getBoundingClientRect();
  const mx    = e.clientX - rect.left;
  const rawP  = 25000 + (mx - PAD.l) / (W - PAD.l - PAD.r) * (250000 - 25000);
  if (rawP < 25000 || rawP > 250000) { tooltip.classList.remove('visible'); return; }
  const snapped = Math.round(rawP / 2500) * 2500;
  const inp     = getInputs();
  document.getElementById('tt-price').textContent = fmt(snapped);
  document.getElementById('tt-rows').innerHTML = tiers.map((pct, i) => {
    const np = netProfit(snapped, pct / 100, inp);
    return \`<div class="tt-row">
      <span class="tt-name"><span class="tt-dot" style="background:\${COLORS[i]}"></span>Buy at \${pct}%</span>
      <div class="tt-right"><div class="tt-val">\${fmt(np)}</div><div class="tt-sub">net profit</div></div>
    </div>\`;
  }).join('');

  // Position tooltip — always keep fully on screen using fixed positioning
  const TT_W = 230, TT_H = tiers.length * 28 + 60;
  let tx = e.clientX + 14;
  let ty = e.clientY - 80;
  if (tx + TT_W > window.innerWidth  - 8) tx = e.clientX - TT_W - 8;
  if (tx < 8) tx = 8;
  if (ty < 8) ty = 8;
  if (ty + TT_H > window.innerHeight - 8) ty = window.innerHeight - TT_H - 8;
  tooltip.style.left = tx + 'px';
  tooltip.style.top  = ty + 'px';
  tooltip.classList.add('visible');
});

canvas.addEventListener('mouseleave', () => tooltip.classList.remove('visible'));
window.addEventListener('resize', resize);
window.addEventListener('load', () => { buildTierInputs(); buildLegend(); resize(); });
</script>
</body>
</html>`;
}


// ============================================================
//  PHASE 5 — OPEN COUNTY IN APP
// ============================================================

/**
 * Opens the county page in LandValuator for the active sheet.
 *
 * Usage: Assign this function to an in-sheet image or button.
 *
 * Reads Parcel State and Parcel County from the first data row of
 * "Scrubbed and Priced" and opens LandValuator with ?state=XX&county=Name
 *
 * URL structure: https://landvaluator.app/?state=XX&county=CountyName
 */
function openCountyInApp() {
  if (!_requireAuth()) return;
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  // Read Parcel State + Parcel County from Scrubbed and Priced
  const spSheet = ss.getSheetByName('Scrubbed and Priced');
  let stateAbbr  = '';
  let countyName = '';

  if (spSheet && spSheet.getLastRow() > 1) {
    const spHeaders = getHeaderMap(spSheet);
    const stateCol  = mapCol(spHeaders, 'Parcel State')  || mapCol(spHeaders, 'State');
    const countyCol = mapCol(spHeaders, 'Parcel County') || mapCol(spHeaders, 'County');

    if (stateCol && countyCol) {
      const lastRow = spSheet.getLastRow();
      const vals = spSheet.getRange(2, 1, lastRow - 1, spSheet.getLastColumn()).getValues();
      for (const row of vals) {
        const s = String(row[stateCol  - 1] || '').trim();
        const c = String(row[countyCol - 1] || '').trim();
        if (s && c) { stateAbbr = s; countyName = c; break; }
      }
    }
  }

  // Fallback: try LI Raw Dataset
  if (!stateAbbr || !countyName) {
    const rawSheet = ss.getSheetByName('LI Raw Dataset');
    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rawHeaders = getHeaderMap(rawSheet);
      const stateCol   = mapCol(rawHeaders, 'State');
      const countyCol  = mapCol(rawHeaders, 'County');
      if (stateCol && countyCol) {
        const vals = rawSheet.getRange(2, 1, 1, rawSheet.getLastColumn()).getValues()[0];
        stateAbbr  = String(vals[stateCol  - 1] || '').trim();
        countyName = String(vals[countyCol - 1] || '').trim();
      }
    }
  }

  if (!stateAbbr || !countyName) {
    ui.alert(
      'County not found.',
      'Make sure "Scrubbed and Priced" has data with Parcel State and Parcel County columns.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Strip " County" suffix if present (LandValuator uses bare name e.g. "Eagle" not "Eagle County")
  countyName = countyName.replace(/\s+county$/i, '').trim();

  const url = 'https://landvaluator.app/?state=' + encodeURIComponent(stateAbbr)
            + '&county=' + encodeURIComponent(countyName);

  // Open in a new browser tab via modal dialog (Apps Script can't open tabs directly)
  const html = HtmlService.createHtmlOutput(
    '<script>window.open(' + JSON.stringify(url) + ', "_blank"); google.script.host.close();<\/script>'
  ).setWidth(10).setHeight(10);
  ui.showModalDialog(html, 'Opening LandValuator…');
}


// ============================================================
//  CHART TRIGGER + MENU
// ============================================================


/** Add custom menu when sheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✉️ Mailing Campaign Commands')
    .addItem('1. Load data into Scrubbed & Priced',   'populateScrubbed')
    .addSeparator()
    .addItem('2. Remove flagged rows (yellow)',        'removeFlaggedRows')
    .addItem('3. Remove all Low likelihood sellers',  'removeLowSellerIQ')
    .addSeparator()
    .addItem('4. Sync Blind Offer Mail Ready Tab',    'syncBlindOfferTab')
    .addItem('5. Sync Range Offer Mail Ready Tab',    'syncRangeOfferTab')
    .addSeparator()
    .addItem('6. Refresh offer prices only',          'refreshOfferPrices')
    .addSeparator()
    .addItem('7. Open county in LandValuator',        'openCountyInApp')
    .addToUi();
}

/**
 * Web app endpoint — called by LandValuator after Save & Sync completes.
 * Triggers refreshOfferPrices() so pricing columns update automatically
 * without requiring a manual click of menu item 6.
 *
 * Deploy steps (one-time):
 *   1. Extensions → Apps Script → Deploy → New deployment
 *   2. Type: Web app | Execute as: Me | Who has access: Anyone
 *   3. Copy the exec URL → add to Netlify as GAS_REFRESH_URL env var
 */
function doPost(e) {
  try {
    refreshOfferPrices(true);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
