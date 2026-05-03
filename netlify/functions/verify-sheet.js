// netlify/functions/verify-sheet.js
// Checks if a Google Sheet ID is registered to any LandValuator user account.
// Called by Apps Script on first function use (cached for 24 hours).

const SUPABASE_URL = 'https://dcrxczsgcuiwimwpokxo.supabase.co';

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const { sheetId } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };

    const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
    if (!serviceKey) return { statusCode: 500, headers, body: JSON.stringify({ error: 'Server configuration error' }) };

    // Fetch all sheet_configs rows — check if any contains this sheetId
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/sheet_configs?select=configs`, {
      headers: {
        'apikey': serviceKey,
        'Authorization': `Bearer ${serviceKey}`,
        'Content-Type': 'application/json',
      },
    });

    if (!resp.ok) throw new Error(`Supabase error: ${resp.status}`);

    const rows = await resp.json();
    const authorized = rows.some(row => {
      const configs = row.configs || {};
      return Object.values(configs).some(c => c && c.sheetId === sheetId);
    });

    return { statusCode: 200, headers, body: JSON.stringify({ authorized }) };
  } catch (err) {
    console.error('verify-sheet error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
