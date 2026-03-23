// netlify/functions/sheets-read.js
// Reads property data from "LI Raw Dataset" tab
// Uses Google Sheets REST API via fetch — no googleapis package needed

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };

  try {
    const { sheetId, sheetName = 'LI Raw Dataset', colCounty = 'County', colAPN = 'APN', metaOnly } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };

    const token = await getAccessToken();

    // Always fetch spreadsheet metadata for title
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`;
    const metaRes = await fetch(metaUrl, { headers: { Authorization: `Bearer ${token}` } });
    let spreadsheetTitle = sheetId;
    if (metaRes.ok) {
      const metaData = await metaRes.json();
      spreadsheetTitle = metaData.properties?.title || sheetId;
    }

    if (metaOnly) {
      return { statusCode: 200, headers, body: JSON.stringify({ spreadsheetTitle, totalRows: 0, properties: [], scrubbedApns: [] }) };
    }

    // Fetch LI Raw Dataset and Scrubbed and Priced in parallel
    const scrubbedTabName = 'Scrubbed and Priced';
    const [rawRes, scrubbedRes] = await Promise.all([
      fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetName)}`, { headers: { Authorization: `Bearer ${token}` } }),
      fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(scrubbedTabName)}`, { headers: { Authorization: `Bearer ${token}` } }),
    ]);
    if (!rawRes.ok) throw new Error(`Sheets API (raw) ${rawRes.status}: ${await rawRes.text()}`);

    // --- LI Raw Dataset ---
    const rawData = await rawRes.json();
    const rows = rawData.values || [];
    if (rows.length < 2) return { statusCode: 200, headers, body: JSON.stringify({ spreadsheetTitle, properties: [], totalRows: 0, scrubbedApns: [] }) };

    const headerRow = rows[0];
    const dataRows = rows.slice(1);
    const colIndex = {};
    headerRow.forEach((h, i) => { if (h) colIndex[h.trim()] = i; });

    const latCol = colIndex['Latitude'];
    const lngCol = colIndex['Longitude'];
    if (latCol === undefined || lngCol === undefined) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'Latitude or Longitude column not found' }) };
    }

    const get = (row, col) => (col !== undefined && row[col]) ? row[col].trim() : '';
    const apnColIndex = colIndex[colAPN] !== undefined ? colIndex[colAPN] : colIndex['APN'];
    const countyColIndex = colIndex[colCounty] !== undefined ? colIndex[colCounty] : colIndex['County'];

    const properties = [];
    dataRows.forEach((row, i) => {
      const lat = parseFloat(get(row, latCol));
      const lng = parseFloat(get(row, lngCol));
      if (isNaN(lat) || isNaN(lng)) return;
      properties.push({
        rowIndex: i + 2, lat, lng,
        apn:     get(row, apnColIndex),
        address: get(row, colIndex['Parcel Address']),
        city:    get(row, colIndex['City']),
        state:   get(row, colIndex['State']),
        zip:     get(row, colIndex['ZIP']),
        county:  get(row, countyColIndex),
        acreage:     get(row, colIndex['Acreage']),
        liAcreage:   get(row, colIndex['LI Calculated Acreage']),
        parcelLink:  get(row, colIndex['Parcel Link']),
        zone:        get(row, colIndex['County Zone']),
      });
    });

    // --- Scrubbed and Priced: extract APN whitelist ---
    let scrubbedApns = [];
    if (scrubbedRes.ok) {
      const scrubbedData = await scrubbedRes.json();
      const sRows = scrubbedData.values || [];
      if (sRows.length >= 1) {
        const sHeader = sRows[0];
        const sColIndex = {};
        sHeader.forEach((h, i) => { if (h) sColIndex[h.trim()] = i; });
        const sApnCol = sColIndex[colAPN] !== undefined ? sColIndex[colAPN] : sColIndex['APN'];
        if (sApnCol !== undefined) {
          scrubbedApns = sRows.slice(1)
            .map(row => (row[sApnCol] || '').trim())
            .filter(Boolean);
        }
      }
    }

    return { statusCode: 200, headers, body: JSON.stringify({ spreadsheetTitle, properties, totalRows: dataRows.length, scrubbedApns }) };
  } catch (err) {
    console.error('sheets-read error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};

async function getAccessToken() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const now = Math.floor(Date.now() / 1000);
  const claim = { iss: creds.client_email, scope: 'https://www.googleapis.com/auth/spreadsheets', aud: 'https://oauth2.googleapis.com/token', exp: now + 3600, iat: now };
  const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = b64url(JSON.stringify(claim));
  const sig = await signRS256(`${header}.${payload}`, creds.private_key);
  const jwt = `${header}.${payload}.${sig}`;
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
