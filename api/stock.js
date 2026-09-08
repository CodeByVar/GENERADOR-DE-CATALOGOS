export const config = {
  runtime: 'edge',
};

// Micro-caché en memoria en la instancia Edge (solo para proteger cuotas simultáneas de Google)
let inMemoryCache = null;
let inMemoryCacheTime = 0;
let pendingFetchPromise = null;

const CACHE_TTL_MS = 6 * 1000; // 6 segundos de micro-caché en servidor Edge
const GOOGLE_SCRIPT_URLS = [
  "https://script.google.com/macros/s/AKfycbz3pjscUdPvuSLWgTA1KugkoffYyWw9zJRqrg22eJCK-by3aTHLF2oZ7t0S3SwmOnwS/exec", // UYUS
  "https://script.google.com/macros/s/AKfycbw5rOmXaEKusH_PYZAG2r0OpybEqqfGlrZsQRQdeiJtJXbCsJsW-oxjQK8q690s8No/exec"  // VARIOS
];

async function fetchOneUrl(url) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 15000);
  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      },
      redirect: 'follow',
      signal: controller.signal,
      cache: 'no-store'
    });
    clearTimeout(timeoutId);
    if (!res.ok) return {};
    const data = await res.json();
    return (data && !data.error) ? data : {};
  } catch (err) {
    clearTimeout(timeoutId);
    return {};
  }
}

async function fetchFromGoogle() {
  const results = await Promise.all(GOOGLE_SCRIPT_URLS.map(fetchOneUrl));
  const merged = Object.assign({}, ...results);
  if (Object.keys(merged).length > 0) {
    inMemoryCache = merged;
    inMemoryCacheTime = Date.now();
    return merged;
  } else {
    throw new Error('No se pudo obtener datos de ninguna de las hojas de Google Sheets');
  }
}

export default async function handler(request) {
  const url = new URL(request.url);
  const isForced = url.searchParams.has('force') || url.searchParams.has('fresh');

  const baseCorsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: baseCorsHeaders });
  }

  const now = Date.now();

  // Si la micro-caché es ultra reciente (< 6 segundos) y no se forzó actualización, servirla al instante
  if (!isForced && inMemoryCache && (now - inMemoryCacheTime < CACHE_TTL_MS)) {
    return new Response(JSON.stringify(inMemoryCache), {
      status: 200,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0, s-maxage=6',
        'X-Stock-Cache': 'HIT-MEMORY',
        ...baseCorsHeaders
      }
    });
  }

  // Si ya hay una consulta en curso a Google, reutilizar la misma promesa (Deduplicación de peticiones)
  if (!pendingFetchPromise) {
    pendingFetchPromise = fetchFromGoogle()
      .finally(() => {
        pendingFetchPromise = null;
      });
  }

  try {
    const data = await pendingFetchPromise;

    return new Response(JSON.stringify(data), {
      status: 200,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0, s-maxage=6',
        'X-Stock-Cache': isForced ? 'BYPASS' : 'MISS',
        ...baseCorsHeaders
      }
    });
  } catch (err) {
    // Si Google falla o tarda mucho, pero tenemos una copia previa en memoria, entregarla
    if (inMemoryCache) {
      return new Response(JSON.stringify(inMemoryCache), {
        status: 200,
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Cache-Control': 'no-cache, no-store, must-revalidate, max-age=0',
          'X-Stock-Cache': 'STALE-FALLBACK',
          ...baseCorsHeaders
        }
      });
    }

    return new Response(JSON.stringify({ error: err.message || 'Error al sincronizar con Google Sheets' }), {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        ...baseCorsHeaders
      }
    });
  }
}
