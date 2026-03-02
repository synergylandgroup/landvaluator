// netlify/functions/sheets-write-pricing.js
// Writes all zone pricing tiers to "Pricing Settings" tab

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const { sheetId, pricingSheetName = 'Pricing Settings', tiers } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };
    if (!tiers || !tiers.length) return { statusCode: 400, headers, body: JSON.stringify({ error: 'tiers required' }) };

    const token = await getAccessToken();

    // Clear existing data rows A4:D103
    const clearUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(pricingSheetName + '!A4:D103')}:clear`;
    const clearRes = await fetch(clearUrl, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: '{}',
    });
    if (!clearRes.ok) throw new Error(`Clear failed ${clearRes.status}: ${await clearRes.text()}`);

    // Write new tiers starting at A4
    const writeData = tiers.map(t => [
      String(t.zone || '').toUpperCase(),
      t.minAcres !== '' && t.minAcres !== undefined ? Number(t.minAcres) : '',
      t.maxAcres !== '' && t.maxAcres !== undefined ? Number(t.maxAcres) : '',
      t.pricePerAcre !== '' && t.pricePerAcre !== undefined ? Number(t.pricePerAcre) : '',
    ]);

    const writeUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(pricingSheetName + '!A4')}?valueInputOption=RAW`;
    const writeRes = await fetch(writeUrl, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: writeData }),
    });
    if (!writeRes.ok) throw new Error(`Write failed ${writeRes.status}: ${await writeRes.text()}`);

    return { statusCode: 200, headers, body: JSON.stringify({ success: true, tiersWritten: tiers.length }) };
  } catch (err) {
    console.error('sheets-write-pricing error:', err);
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
