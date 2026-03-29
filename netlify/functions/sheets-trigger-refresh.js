// netlify/functions/sheets-trigger-refresh.js
// Called by LandValuator after Save & Sync to auto-refresh offer prices
// in "Scrubbed and Priced" via the Apps Script web app endpoint.
//
// Required env var: GAS_REFRESH_URL
//   Set in Netlify → Site configuration → Environment variables
//   Value: the exec URL from your Apps Script deployment (see doPost comments in GS)

exports.handler = async () => {
  const url = process.env.GAS_REFRESH_URL;
  if (!url) {
    // No URL configured — silently succeed so sync still completes
    return { statusCode: 200, body: JSON.stringify({ success: true, skipped: true }) };
  }
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source: 'landvaluator' }),
      redirect: 'follow',
    });
    const data = await res.json().catch(() => ({}));
    return {
      statusCode: 200,
      body: JSON.stringify({ success: data.success !== false }),
    };
  } catch (err) {
    // Non-fatal — sync already succeeded; log and move on
    console.error('sheets-trigger-refresh error:', err.message);
    return { statusCode: 200, body: JSON.stringify({ success: false, error: err.message }) };
  }
};
