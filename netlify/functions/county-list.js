// netlify/functions/county-list.js
// Proxies Census Bureau API requests for county lists to avoid CORS restrictions.

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json',
  };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };

  const fips = event.queryStringParameters?.fips;
  if (!fips) return { statusCode: 400, headers, body: JSON.stringify({ error: 'fips required' }) };

  try {
    // Try Census Bureau API
    const res = await fetch(
      `https://api.census.gov/data/2020/dec/pl?get=NAME&for=county:*&in=state:${fips}`,
      { headers: { 'User-Agent': 'LandValuator/1.0' } }
    );
    if (!res.ok) throw new Error(`Census API ${res.status}: ${await res.text()}`);
    const data = await res.json();
    return { statusCode: 200, headers, body: JSON.stringify(data) };
  } catch (err) {
    console.error('county-list error:', err.message);
    // Fallback: try TIGERweb REST API
    try {
      const res2 = await fetch(
        `https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/State_County/MapServer/1/query?where=STATE=${fips}&outFields=NAME,COUNTY&f=json`
      );
      if (!res2.ok) throw new Error(`TIGERweb ${res2.status}`);
      const data2 = await res2.json();
      if (data2.features && data2.features.length) {
        // Convert to same format as Census API: [[name, state, county], ...]
        const rows = [['NAME', 'state', 'county']];
        data2.features.forEach(f => {
          rows.push([`${f.attributes.NAME} County`, String(fips), f.attributes.COUNTY]);
        });
        return { statusCode: 200, headers, body: JSON.stringify(rows) };
      }
    } catch(e2) {
      console.error('TIGERweb fallback error:', e2.message);
    }
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
