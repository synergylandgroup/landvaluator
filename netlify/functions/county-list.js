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
    const res = await fetch(`https://api.census.gov/data/2020/dec/pl?get=NAME&for=county:*&in=state:${fips}`);
    if (!res.ok) throw new Error(`Census API ${res.status}`);
    const data = await res.json();
    return { statusCode: 200, headers, body: JSON.stringify(data) };
  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
