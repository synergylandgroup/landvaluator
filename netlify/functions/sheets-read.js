// netlify/functions/sheets-read.js
// Reads property data from "LI Raw Dataset" tab
// All column lookups by header name — never by position

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

  try {
    const { sheetId, sheetName = 'LI Raw Dataset' } = JSON.parse(event.body || '{}');
    if (!sheetId) return { statusCode: 400, headers, body: JSON.stringify({ error: 'sheetId required' }) };

    // Auth via service account
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    // Fetch all data from the tab
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: sheetName,
    });

    const rows = res.data.values;
    if (!rows || rows.length < 2) {
      return { statusCode: 200, headers, body: JSON.stringify({ properties: [], headers: [] }) };
    }

    const headerRow = rows[0];
    const dataRows  = rows.slice(1);

    // Build header → index map (by name, never by position)
    const colIndex = {};
    headerRow.forEach((h, i) => { if (h) colIndex[h.trim()] = i; });

    // Required columns
    const latCol  = colIndex['Latitude'];
    const lngCol  = colIndex['Longitude'];

    if (latCol === undefined || lngCol === undefined) {
      return {
        statusCode: 400, headers,
        body: JSON.stringify({ error: 'Latitude or Longitude column not found in sheet headers' })
      };
    }

    // Optional columns — all by header name
    const apnCol     = colIndex['APN'];
    const addressCol = colIndex['Parcel Address'];
    const cityCol    = colIndex['City'];
    const stateCol   = colIndex['State'];
    const zipCol     = colIndex['ZIP'];
    const countyCol  = colIndex['County'];
    const acreageCol = colIndex['Acreage'];
    const zoneCol    = colIndex['County Zone'];

    const properties = [];
    dataRows.forEach((row, i) => {
      const lat = parseFloat(row[latCol]);
      const lng = parseFloat(row[lngCol]);
      if (isNaN(lat) || isNaN(lng)) return;

      properties.push({
        rowIndex: i + 2, // 1-based, accounting for header row
        lat,
        lng,
        apn:     apnCol     !== undefined ? (row[apnCol]     || '') : '',
        address: addressCol !== undefined ? (row[addressCol] || '') : '',
        city:    cityCol    !== undefined ? (row[cityCol]    || '') : '',
        state:   stateCol   !== undefined ? (row[stateCol]   || '') : '',
        zip:     zipCol     !== undefined ? (row[zipCol]     || '') : '',
        county:  countyCol  !== undefined ? (row[countyCol]  || '') : '',
        acreage: acreageCol !== undefined ? (row[acreageCol] || '') : '',
        zone:    zoneCol    !== undefined ? (row[zoneCol]    || '') : '',
      });
    });

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ properties, totalRows: dataRows.length }),
    };

  } catch (err) {
    console.error('sheets-read error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
