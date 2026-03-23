// netlify/functions/sheets-write-zones.js
// Writes zone letters back to "County Zone" column in "LI Raw Dataset"

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const { sheetId, sheetName = 'Scrubbed and Priced', assignments } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };
    if (!assignments || !assignments.length) return { statusCode: 400, headers, body: JSON.stringify({ error: 'assignments required' }) };

    const token = await getAccessToken();

    // Find County Zone column by header name
    const headerUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(sheetName + '!1:1')}`;
    const headerRes = await fetch(headerUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!headerRes.ok) throw new Error(`Sheets API ${headerRes.status}: ${await headerRes.text()}`);
    const headerData = await headerRes.json();
    const headerRow = (headerData.values || [[]])[0] || [];
    const zoneColIndex = headerRow.findIndex(h => h && h.trim() === 'County Zone');
    if (zoneColIndex === -1) return { statusCode: 400, headers, body: JSON.stringify({ error: '"County Zone" column not found' }) };

    const zoneColLetter = colToLetter(zoneColIndex);

    // Batch update
    const data = assignments.map(({ rowIndex, zone }) => ({
      range: `${sheetName}!${zoneColLetter}${rowIndex}`,
      values: [[zone.toUpperCase()]],
    }));

    const writeUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`;
    const writeRes = await fetch(writeUrl, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ valueInputOption: 'RAW', data }),
    });
    if (!writeRes.ok) throw new Error(`Sheets API ${writeRes.status}: ${await writeRes.text()}`);

    return { statusCode: 200, headers, body: JSON.stringify({ success: true, updated: assignments.length }) };
  } catch (err) {
    console.error('sheets-write-zones error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};

function colToLetter(col) {
  let letter = ''; col++;
  while (col > 0) { const mod = (col-1)%26; letter = String.fromCharCode(65+mod)+letter; col = Math.floor((col-mod)/26); }
  return letter;
}

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
