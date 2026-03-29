// netlify/functions/auth-callback.js
// Handles Supabase auth redirects for password recovery (PKCE flow)
// Supabase sends ?code= to this endpoint after the user clicks the reset link
// We exchange the code and redirect back to the app with a clean ?type=recovery param

exports.handler = async (event) => {
  const params = new URLSearchParams(event.queryStringParameters || {});
  const code = params.get('code');
  const type = params.get('type') || 'recovery';
  const next = params.get('next') || '/';

  const SUPABASE_URL = 'https://dcrxczsgcuiwimwpokxo.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRjcnhjenNnY3Vpd2ltd3Bva3hvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ3NDU4MDAsImV4cCI6MjA5MDMyMTgwMH0.BFNKnN5mzaGLazQQTNhl8TytA5JW5IQxa5ouFg4-KB4';

  if (!code) {
    // No code — redirect to home
    return {
      statusCode: 302,
      headers: { Location: 'https://landvaluator.app' },
      body: '',
    };
  }

  try {
    // Exchange the code for a session via Supabase REST API
    const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=pkce`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
      },
      body: JSON.stringify({ auth_code: code }),
    });

    const data = await res.json();

    if (!res.ok || !data.access_token) {
      console.error('Token exchange failed:', data);
      return {
        statusCode: 302,
        headers: { Location: 'https://landvaluator.app?auth_error=1' },
        body: '',
      };
    }

    // Build redirect URL — pass tokens in hash so Supabase client can pick them up
    // and add type=recovery as a query param we control
    const redirectUrl = new URL('https://landvaluator.app');
    redirectUrl.searchParams.set('type', type);
    redirectUrl.hash = `access_token=${data.access_token}&refresh_token=${data.refresh_token}&token_type=bearer&type=${type}`;

    return {
      statusCode: 302,
      headers: { Location: redirectUrl.toString() },
      body: '',
    };
  } catch (err) {
    console.error('auth-callback error:', err);
    return {
      statusCode: 302,
      headers: { Location: 'https://landvaluator.app?auth_error=1' },
      body: '',
    };
  }
};
