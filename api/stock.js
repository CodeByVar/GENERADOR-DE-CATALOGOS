export default async function handler(req, res) {
  // Permitir solicitudes desde cualquier origen (CORS)
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxrXCYxH9JX-uO2rw5Wg7XY5PnbKso50ugmpkTnrPacwy12GoMpxn-AvlbRZ_m0a9k45w/exec";

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 28000); // 28 segundos para asegurar respuesta de Google Drive

    const response = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
      },
      redirect: 'follow',
      signal: controller.signal
    });

    clearTimeout(timeout);

    if (!response.ok) {
      return res.status(response.status).json({ error: `HTTP ${response.status} de Google Apps Script` });
    }

    const data = await response.json();

    // Cache inteligente en Vercel CDN: 60s fresco, hasta 120s mientras revalida en fondo
    res.setHeader('Cache-Control', 's-maxage=60, stale-while-revalidate=120');
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ error: error.message || 'Error al conectar con Google Sheets/Apps Script' });
  }
}
