// netlify/functions/share-load.js
// Loads zone data from Netlify Blobs REST API (no SDK needed)

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };

  try {
    const id = event.queryStringParameters?.id;
    if (!id) return { statusCode: 400, headers, body: JSON.stringify({ error: 'id required' }) };

    const siteId = process.env.NETLIFY_SITE_ID;
    const token = process.env.NETLIFY_TOKEN || process.env.NETLIFY_BLOBS_TOKEN;

    if (!siteId || !token) {
      return { statusCode: 503, headers, body: JSON.stringify({ error: 'Blob storage not configured' }) };
    }

    const blobUrl = `https://api.netlify.com/api/v1/blobs/${siteId}/zone-shares/${id}`;
    const res = await fetch(blobUrl, {
      headers: { 'Authorization': `Bearer ${token}` },
    });

    if (!res.ok) {
      return { statusCode: 404, headers, body: JSON.stringify({ error: 'Share not found or expired' }) };
    }

    const data = await res.json();
    return { statusCode: 200, headers, body: JSON.stringify(data) };
  } catch (err) {
    console.error('share-load error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
