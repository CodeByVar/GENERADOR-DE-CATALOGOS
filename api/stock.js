export const config = {
  runtime: 'edge',
};

export default async function handler(request) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
    'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0'
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxrXCYxH9JX-uO2rw5Wg7XY5PnbKso50ugmpkTnrPacwy12GoMpxn-AvlbRZ_m0a9k45w/exec";

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 28000);

    const res = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
      },
      redirect: 'follow',
      signal: controller.signal,
      cache: 'no-store'
    });

    clearTimeout(timeoutId);

    if (!res.ok) {
      return new Response(JSON.stringify({ error: `Google HTTP ${res.status}` }), {
        status: res.status,
        headers: { 'Content-Type': 'application/json', ...corsHeaders }
      });
    }

    const data = await res.json();
    return new Response(JSON.stringify(data), {
      status: 200,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        ...corsHeaders
      }
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message || 'Error al conectar con Google Sheets' }), {
      status: 500,
      headers: { 'Content-Type': 'application/json', ...corsHeaders }
    });
  }
}
