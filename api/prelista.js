export const config = {
  runtime: 'edge',
};

// URL de tu Google Apps Script implementado en el Google Sheet "I. RIVERO 2026"
// Pega aquí la URL de tu Web App (terminada en /exec) una vez la implementes:
let APPS_SCRIPT_PRELISTA_URL = process.env.APPS_SCRIPT_PRELISTA_URL || "https://script.google.com/macros/s/AKfycbwZluWpdw3riIr35gdrHhgdn7cRLcoNOuqQffxbnPIFqFbu2EgsxZffipAs0c4_gDpbKg/exec";

// Códigos que el administrador excluyó para no mostrar ni vender a clientes
const EXCLUDED_CODES_LIST = [
  // EXCLUDED_CODES_START
  "SUM022",
  "DC-LLAVERO"
  // EXCLUDED_CODES_END
];
const EXCLUDED_CODES = new Set(EXCLUDED_CODES_LIST.map(c => String(c).toUpperCase().trim()));

// Micro-caché en servidor Edge por hoja
const cacheMap = new Map();
const CACHE_TTL_MS = 10 * 1000; // 10 segundos

// Datos de demostración inicial basados en las hojas más recientes de tu Google Sheet
// (Garantiza que la web funcione de inmediato mientras conecta tu script)
// DEMO_PRELISTA_START
const DEMO_PRELISTA = {
  success: true,
  sheet: 'Prelista',
  availableSheets: ['Prelista'],
  totalProductos: 0,
  actualizadoEn: new Date().toISOString(),
  productos: []
};
// DEMO_PRELISTA_END

export default async function handler(request) {
  const url = new URL(request.url);
  const isForced = url.searchParams.has('force') || url.searchParams.has('fresh');
  const sheetParam = url.searchParams.get('sheet') || '';
  const scriptUrlParam = url.searchParams.get('script_url') || '';

  const targetScriptUrl = scriptUrlParam || APPS_SCRIPT_PRELISTA_URL;

  const baseCorsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
    'Content-Type': 'application/json; charset=utf-8'
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: baseCorsHeaders });
  }

  const now = Date.now();

  // Si no hay URL configurada todavía, retornar datos de demostración con aviso
  if (!targetScriptUrl) {
    let responseData = { ...DEMO_PRELISTA };
    if (sheetParam) responseData.sheet = sheetParam;
    return new Response(JSON.stringify(responseData), {
      status: 200,
      headers: {
        ...baseCorsHeaders,
        'Cache-Control': 'no-store',
        'X-Prelista-Source': 'Demo-Fallback'
      }
    });
  }

  const cacheKey = (sheetParam || '__DEFAULT__').trim().toUpperCase();

  // Si la caché es reciente y no se forzó actualización
  const cached = cacheMap.get(cacheKey);
  if (!isForced && cached && (now - cached.time < CACHE_TTL_MS)) {
    return new Response(JSON.stringify(cached.data), {
      status: 200,
      headers: {
        ...baseCorsHeaders,
        'Cache-Control': 'public, max-age=6, s-maxage=10, stale-while-revalidate=30',
        'X-Prelista-Source': 'Edge-Cache'
      }
    });
  }

  try {
    const sheetNames = sheetParam ? sheetParam.split(/[,+]/).map(s => s.trim()).filter(Boolean) : [];

    // Si se piden múltiples hojas
    if (sheetNames.length > 1) {
      const fetchPromises = sheetNames.map(async s => {
        try {
          const sUrl = new URL(targetScriptUrl);
          sUrl.searchParams.set('sheet', s);
          if (isForced) sUrl.searchParams.set('force', '1');
          const resp = await fetch(sUrl.toString(), {
            method: 'GET',
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
            redirect: 'follow',
            cache: 'no-store'
          });
          if (!resp.ok) return null;
          return await resp.json();
        } catch(e) {
          return null;
        }
      });

      const results = await Promise.all(fetchPromises);
      const prodMap = new Map();
      const mergedList = [];
      let availableSheets = [];

      results.forEach((r, idx) => {
        if (r && r.success && Array.isArray(r.productos)) {
          if (Array.isArray(r.availableSheets) && r.availableSheets.length > availableSheets.length) {
            availableSheets = r.availableSheets;
          }
          const sName = sheetNames[idx];
          r.productos.forEach(p => {
            const cod = String(p.codigo || '').toUpperCase().trim();
            if (!cod || EXCLUDED_CODES.has(cod)) return;
            if (prodMap.has(cod)) {
              const ex = prodMap.get(cod);
              ex.cajasVienen = (ex.cajasVienen || 0) + (p.cajasVienen || 0);
              ex.stockReserva = (ex.stockReserva || 0) + (p.stockReserva || 0);
              if (!ex._origenHojas) ex._origenHojas = [];
              if (!ex._origenHojas.includes(sName)) ex._origenHojas.push(sName);
            } else {
              const item = { ...p, _origenHojas: [sName] };
              prodMap.set(cod, item);
              mergedList.push(item);
            }
          });
        }
      });

      const multiData = {
        success: true,
        sheet: sheetNames.join(' + '),
        availableSheets: availableSheets.length > 0 ? availableSheets : sheetNames,
        totalProductos: mergedList.length,
        productos: mergedList
      };

      cacheMap.set(cacheKey, { data: multiData, time: now });
      return new Response(JSON.stringify(multiData), {
        status: 200,
        headers: {
          ...baseCorsHeaders,
          'Cache-Control': 'public, max-age=6, s-maxage=10',
          'X-Prelista-Source': 'Live-Google-Sheets-Multi'
        }
      });
    }

    const fetchUrl = new URL(targetScriptUrl);
    if (sheetParam) fetchUrl.searchParams.set('sheet', sheetParam);
    if (isForced) fetchUrl.searchParams.set('force', '1');

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000);

    const res = await fetch(fetchUrl.toString(), {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      },
      redirect: 'follow',
      signal: controller.signal,
      cache: 'no-store'
    });
    clearTimeout(timeoutId);

    if (!res.ok) {
      throw new Error(`Google Apps Script devolvió código ${res.status}`);
    }

    const data = await res.json();
    if (data && data.success) {
      if (Array.isArray(data.productos) && EXCLUDED_CODES.size > 0) {
        data.productos = data.productos.filter(p => !EXCLUDED_CODES.has(String(p.codigo || '').toUpperCase().trim()));
        data.totalProductos = data.productos.length;
      }
      cacheMap.set(cacheKey, { data, time: now });
      if (data.sheet) {
        cacheMap.set(data.sheet.trim().toUpperCase(), { data, time: now });
      }
      return new Response(JSON.stringify(data), {
        status: 200,
        headers: {
          ...baseCorsHeaders,
          'Cache-Control': 'public, max-age=6, s-maxage=10',
          'X-Prelista-Source': 'Live-Google-Sheets'
        }
      });
    } else {
      throw new Error(data.error || 'Respuesta inválida de Google Apps Script');
    }

  } catch (err) {
    // Si falla Google Apps Script por red o cuota, responder con los datos en memoria o demo
    const cachedAny = cacheMap.get(cacheKey) || cacheMap.get('__DEFAULT__');
    const fallback = cachedAny ? cachedAny.data : { ...DEMO_PRELISTA };
    if (sheetParam && fallback) fallback.sheet = sheetParam;
    return new Response(JSON.stringify(fallback), {
      status: 200,
      headers: {
        ...baseCorsHeaders,
        'Cache-Control': 'no-store',
        'X-Prelista-Warning': err.message || 'Error de conexión'
      }
    });
  }
}
