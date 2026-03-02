// netlify/functions/sheets-write-pricing.js
// Writes zone pricing tiers to the "Pricing Settings" tab
// Preserves the existing header structure (rows 1-3)
// Replaces all data rows starting at row 4
// All column lookups by header name

const { google } = require('googleapis');

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  try {
    const {
      sheetId,
      pricingSheetName = 'Pricing Settings',
      tiers, // [{ zone: 'A', minAcres: 0, maxAcres: 5, pricePerAcre: 1200 }, ...]
    } = JSON.parse(event.body || '{}');

    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };
    if (!tiers || !tiers.length) return { statusCode: 400, headers, body: JSON.stringify({ error: 'tiers array required' }) };

    // Auth
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    // Read current header rows (rows 1-3) to find column positions by name
    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `${pricingSheetName}!2:2`, // row 2 has the actual column labels
    });

    const headerRow = (headerRes.data.values || [[]])[0] || [];
    const colIndex  = {};
    headerRow.forEach((h, i) => { if (h) colIndex[h.trim()] = i; });

    // Find required columns by name
    const zoneCol  = colIndex['County Zone'];
    const minCol   = colIndex['Min Acres'];
    const maxCol   = colIndex['Max Acres'];
    const ppaCol   = colIndex['Price Per Acre ($)'];

    if (zoneCol === undefined || minCol === undefined || maxCol === undefined || ppaCol === undefined) {
      return {
        statusCode: 400, headers,
        body: JSON.stringify({
          error: 'Required columns not found in Pricing Settings row 2',
          found: Object.keys(colIndex),
          needed: ['County Zone', 'Min Acres', 'Max Acres', 'Price Per Acre ($)']
        })
      };
    }

    // Determine how many columns wide the sheet is
    const numCols = headerRow.length;

    // Build output rows — one per tier
    // Each row is sparse: only fill the 4 pricing columns, leave others blank
    const outputRows = tiers.map(tier => {
      const row = new Array(numCols).fill('');
      row[zoneCol] = String(tier.zone || '').toUpperCase();
      row[minCol]  = tier.minAcres  !== '' && tier.minAcres  !== undefined ? Number(tier.minAcres)  : '';
      row[maxCol]  = tier.maxAcres  !== '' && tier.maxAcres  !== undefined ? Number(tier.maxAcres)  : '';
      row[ppaCol]  = tier.pricePerAcre !== '' && tier.pricePerAcre !== undefined ? Number(tier.pricePerAcre) : '';
      return row;
    });

    // Clear existing data rows (row 4 onwards) in the polygon PPA section only
    // We only clear columns A-D (the 4 pricing columns) to avoid touching Blind/Range sections
    const lastDataRow = 3 + 100; // clear up to 100 rows of old data
    await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId,
      range: `${pricingSheetName}!A4:D${lastDataRow}`,
    });

    // Write new pricing rows starting at row 4
    if (outputRows.length > 0) {
      // Only write columns A-D
      const writeData = outputRows.map(row => [row[zoneCol], row[minCol], row[maxCol], row[ppaCol]]);

      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `${pricingSheetName}!A4`,
        valueInputOption: 'RAW',
        requestBody: { values: writeData },
      });
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ success: true, tiersWritten: tiers.length }),
    };

  } catch (err) {
    console.error('sheets-write-pricing error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
