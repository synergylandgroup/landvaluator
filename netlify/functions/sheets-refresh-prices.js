// netlify/functions/sheets-refresh-prices.js
// Recalculates all offer price columns in Scrubbed and Priced using
// pricing tiers passed from the app + blind/range multipliers read from Pricing Settings.
// Called fire-and-forget after Save & Sync completes in saveAndSyncZone().

const BLIND_PREFIXES  = ['T —', 'U —', 'V —', 'W —', 'X —', 'Y —'];
const DEFAULT_BLIND   = [
  { mult: 0.50, min: 25000,  max: 70000  },
  { mult: 0.55, min: 70000,  max: 95000  },
  { mult: 0.60, min: 95000,  max: 125000 },
  { mult: 0.65, min: 125000, max: 250000 },
  { mult: 0.70, min: 250000, max: 999999 },
];
const DEFAULT_RANGE = [{ mult: 0.50 }, { mult: 0.65 }];

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST')    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const { sheetId, tiers } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };

    const token = await getAccessToken();

    // 1. Read Pricing Settings for blind/range multipliers
    const { blindTiers, rangeTiers } = await readPricingSettings(sheetId, token);

    // 2. Read Scrubbed and Priced
    const spData = await readSheet(sheetId, 'Scrubbed and Priced', token);
    if (!spData || spData.length < 2) return { statusCode: 200, headers, body: JSON.stringify({ success: true, updated: 0 }) };

    const headerRow = spData[0];
    const dataRows  = spData.slice(1);

    // Build header map (lowercase, 0-based)
    const hMap = {};
    headerRow.forEach((h, i) => { if (h) hMap[String(h).trim().toLowerCase()] = i; });

    const col       = name   => hMap[name.toLowerCase()];
    const prefixCol = prefix => { const p = prefix.toLowerCase(); const k = Object.keys(hMap).find(k => k.startsWith(p)); return k !== undefined ? hMap[k] : undefined; };

    const polygonCol   = col('county zone');
    const acreageCol   = col('acreage');
    const liAcreageCol = col('li calculated acreage');
    const manPPACol    = col('manually calculated ppa');
    const manMVCol     = col('manually calculated market value');
    const blindCol     = col('blind offer price');
    const rlCol        = prefixCol('range offer low');
    const rhCol        = prefixCol('range offer high');

    if (polygonCol === undefined) return { statusCode: 400, headers, body: JSON.stringify({ error: 'County Zone column not found' }) };

    // Find tier columns by matching the blind tier percentages to column headers
    const tierCols = blindTiers.map(tier => {
      const pct = Math.round(tier.mult * 100) + '%';
      const k = Object.keys(hMap).find(k => k.startsWith(pct + ' offer'));
      return k !== undefined ? hMap[k] : undefined;
    });

    // 3. Calculate offer prices for every data row
    const ppaOut = [], mvOut = [], blindOut = [], rlOut = [], rhOut = [];
    const tierOut = tierCols.map(() => []);

    for (const row of dataRows) {
      const polygon = String(row[polygonCol] || '').trim();
      const liRaw   = liAcreageCol !== undefined ? row[liAcreageCol] : '';
      const acreage = (liRaw !== '' && liRaw !== null && liRaw !== undefined)
        ? (parseFloat(liRaw) || 0)
        : (parseFloat(row[acreageCol] || 0) || 0);

      const ppa = lookupPPA(tiers || [], polygon, acreage);
      const mv  = ppa > 0 && acreage > 0 ? ppa * acreage : 0;

      ppaOut.push(ppa > 0 ? ppa : '');
      mvOut.push(mv  > 0 ? mv  : '');

      let maxTier = '';
      tierCols.forEach((tCol, t) => {
        const tier = blindTiers[t];
        const val  = (tier && mv >= tier.min && mv < tier.max) ? mv * tier.mult : '';
        tierOut[t].push(val);
        if (val !== '' && (maxTier === '' || val > maxTier)) maxTier = val;
      });

      blindOut.push(maxTier);
      rlOut.push(mv > 0 ? mv * rangeTiers[0].mult : '');
      rhOut.push(mv > 0 ? mv * rangeTiers[1].mult : '');
    }

    // 4. Write back only the calculated columns (batch update)
    const startRow = 2;
    const writeData = [];
    const addRange = (colIdx, values) => {
      if (colIdx === undefined) return;
      const letter = colToLetter(colIdx + 1);
      writeData.push({
        range:  `Scrubbed and Priced!${letter}${startRow}:${letter}${startRow + values.length - 1}`,
        values: values.map(v => [v]),
      });
    };

    addRange(manPPACol, ppaOut);
    addRange(manMVCol,  mvOut);
    addRange(blindCol,  blindOut);
    addRange(rlCol,     rlOut);
    addRange(rhCol,     rhOut);
    tierCols.forEach((tCol, t) => addRange(tCol, tierOut[t]));

    if (writeData.length > 0) {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ valueInputOption: 'RAW', data: writeData }),
      });
      if (!res.ok) throw new Error(`Sheets batchUpdate failed: ${res.status}`);
    }

    return { statusCode: 200, headers, body: JSON.stringify({ success: true, updated: dataRows.length }) };
  } catch (err) {
    console.error('sheets-refresh-prices error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};

// ── Read Pricing Settings to extract blind/range multipliers ──────────────────

async function readPricingSettings(sheetId, token) {
  const data = await readSheet(sheetId, 'Pricing Settings', token);
  if (!data || data.length < 2) return { blindTiers: DEFAULT_BLIND, rangeTiers: DEFAULT_RANGE };

  let multCol = -1, minCol = -1, maxCol = -1;
  for (const row of data) {
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '').trim().toLowerCase();
      if (v === 'multiplier (%)' && multCol < 0) multCol = c;
      if (v === 'min value ($)'  && minCol  < 0) minCol  = c;
      if (v === 'max value ($)'  && maxCol  < 0) maxCol  = c;
    }
    if (multCol >= 0 && minCol >= 0 && maxCol >= 0) break;
  }

  const blindTiers = [];
  for (const row of data) {
    let matched = false;
    for (let c = 0; c < row.length; c++) {
      const label = String(row[c] || '').trim();
      if (!matched && BLIND_PREFIXES.some(p => label.startsWith(p))) {
        const rawMult = multCol >= 0 ? parseFloat(row[multCol]) || 0 : 0;
        const mult    = rawMult > 1 ? rawMult / 100 : rawMult;
        const min     = minCol  >= 0 ? parseFloat(row[minCol]) || 0      : 0;
        const max     = maxCol  >= 0 ? parseFloat(row[maxCol]) || 999999 : 999999;
        if (mult > 0) blindTiers.push({ mult, min, max });
        matched = true;
      }
    }
  }

  // Range multipliers live in L4:L5 (row index 3 & 4, col index 11)
  const parseMult = v => { const n = parseFloat(v) || 0; return n > 0 ? (n > 1 ? n / 100 : n) : 0; };
  const rangeLow  = data.length > 3 ? parseMult(data[3][11]) : 0;
  const rangeHigh = data.length > 4 ? parseMult(data[4][11]) : 0;

  // If min/max columns weren't found, all tiers default to (0, 999999) which
  // causes every tier to match every row. Fall back to DEFAULT_BLIND in that case.
  const hasValidRanges = blindTiers.some(t => t.min > 0 || t.max < 999999);
  return {
    blindTiers: (blindTiers.length > 0 && hasValidRanges) ? blindTiers.slice(0, 5) : DEFAULT_BLIND,
    rangeTiers: rangeLow > 0 && rangeHigh > 0
      ? [{ mult: rangeLow }, { mult: rangeHigh }]
      : DEFAULT_RANGE,
  };
}

// ── Look up price per acre for a given zone + acreage ────────────────────────

function lookupPPA(tiers, polygon, acreage) {
  const poly = String(polygon || '').trim().toUpperCase();
  const ac   = parseFloat(acreage) || 0;
  for (const t of tiers) {
    if (String(t.zone || '').toUpperCase() === poly &&
        ac >= (parseFloat(t.minAcres) || 0) &&
        ac <  (parseFloat(t.maxAcres) || 999999)) return parseFloat(t.pricePerAcre) || 0;
  }
  for (const t of tiers) {
    if (String(t.zone || '').toUpperCase() === 'ALL' &&
        ac >= (parseFloat(t.minAcres) || 0) &&
        ac <  (parseFloat(t.maxAcres) || 999999)) return parseFloat(t.pricePerAcre) || 0;
  }
  return 0;
}

// ── Sheets API helpers ───────────────────────────────────────────────────────

async function readSheet(sheetId, sheetName, token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetName)}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return null;
  const data = await res.json();
  return data.values || [];
}

function colToLetter(col) {
  let result = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

async function getAccessToken() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const now   = Math.floor(Date.now() / 1000);
  const claim = { iss: creds.client_email, scope: 'https://www.googleapis.com/auth/spreadsheets', aud: 'https://oauth2.googleapis.com/token', exp: now + 3600, iat: now };
  const header  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = b64url(JSON.stringify(claim));
  const sig     = await signRS256(`${header}.${payload}`, creds.private_key);
  const jwt     = `${header}.${payload}.${sig}`;
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Auth failed: ' + JSON.stringify(data));
  return data.access_token;
}

function b64url(str) {
  return Buffer.from(str).toString('base64').replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
}

async function signRS256(data, pemKey) {
  const { createSign } = require('crypto');
  const sign = createSign('RSA-SHA256');
  sign.update(data); sign.end();
  return sign.sign(pemKey).toString('base64').replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
}
