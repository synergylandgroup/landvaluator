// netlify/functions/share-save.js
// Saves zone data using Netlify Blobs REST API (no SDK needed)

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };

  try {
    const payload = JSON.parse(event.body || '{}');
    if (!payload.zones || !payload.zones.length) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'zones required' }) };
    }

    // Generate short ID
    const id = Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);

    // Use Netlify Blobs REST API
    const siteId = process.env.NETLIFY_SITE_ID;
    const token = process.env.NETLIFY_TOKEN || process.env.NETLIFY_BLOBS_TOKEN;

    if (!siteId || !token) {
      // Fallback: encode in response and let client use URL params
      return { statusCode: 200, headers, body: JSON.stringify({ id, fallback: true, data: payload }) };
    }

    const blobUrl = `https://api.netlify.com/api/v1/blobs/${siteId}/zone-shares/${id}`;
    const res = await fetch(blobUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ ...payload, created: new Date().toISOString() }),
    });

    if (!res.ok) {
      const err = await res.text();
      console.error('Blob PUT error:', res.status, err);
      // Fallback if blob fails
      return { statusCode: 200, headers, body: JSON.stringify({ id, fallback: true, data: payload }) };
    }

    return { statusCode: 200, headers, body: JSON.stringify({ id }) };
  } catch (err) {
    console.error('share-save error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
