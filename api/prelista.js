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
const DEMO_PRELISTA = {
  success: true,
  sheet: "Hoja 49",
  availableSheets: ["Hoja 50", "Hoja 49", "Hoja 48", "Hoja 47", "Hoja 46", "Hoja 45"],
  tipoCambio: 11.5,
  totalProductos: 26,
  actualizadoEn: new Date().toISOString(),
  productos: [
    {
      id: "P_1",
      codigo: "THWS030301",
      detalle: "ACOPLE RAPIDO 1/2\" PLASTICO SET/3PZS TOTAL",
      marca: "TOTAL",
      cajasVienen: 2,
      cantPorCaja: 48,
      unidad: "SET",
      precioMayor: 18.00,
      precioCaja: 17.00,
      precioEspecial: 15.00,
      precioRefUni: 17.00,
      precioCajaTotal: 816.00,
      stockReserva: 2,
      agotado: false
    },
    {
      id: "P_2",
      codigo: "TSH501420",
      detalle: "ARNES DE SEGURIDAD 14MM*1.95MTS(3PTS. D/FIJACION 4PTS. D/AJUSTE) TOTAL",
      marca: "TOTAL",
      cajasVienen: 3,
      cantPorCaja: 6,
      unidad: "UNI",
      precioMayor: 290.00,
      precioCaja: 285.00,
      precioEspecial: 280.00,
      precioRefUni: 285.00,
      precioCajaTotal: 1710.00,
      stockReserva: 3,
      agotado: false
    },
    {
      id: "P_3",
      codigo: "THRRT32712",
      detalle: "BOLSA DE HERRAMIENTAS C/RUEDAS 27\"/25KGS/13BOLSILLOS TOTAL",
      marca: "TOTAL",
      cajasVienen: 5,
      cantPorCaja: 2,
      unidad: "UNI",
      precioMayor: 340.00,
      precioCaja: 335.00,
      precioEspecial: 330.00,
      precioRefUni: 335.00,
      precioCajaTotal: 670.00,
      stockReserva: 5,
      agotado: false
    },
    {
      id: "P_4",
      codigo: "TPBX0141",
      detalle: "CAJA HERRAMIENTAS PLASTICA 14\" TOTAL",
      marca: "TOTAL",
      cajasVienen: 25,
      cantPorCaja: 8,
      unidad: "UNI",
      precioMayor: 50.00,
      precioCaja: 48.00,
      precioEspecial: 46.00,
      precioRefUni: 48.00,
      precioCajaTotal: 384.00,
      stockReserva: 25,
      agotado: false
    },
    {
      id: "P_5",
      codigo: "TPBX0171",
      detalle: "CAJA HERRAMIENTAS PLASTICA 17\" TOTAL",
      marca: "TOTAL",
      cajasVienen: 7,
      cantPorCaja: 8,
      unidad: "UNI",
      precioMayor: 80.00,
      precioCaja: 75.00,
      precioEspecial: 70.00,
      precioRefUni: 75.00,
      precioCajaTotal: 600.00,
      stockReserva: 7,
      agotado: false
    },
    {
      id: "P_6",
      codigo: "TPBX0201",
      detalle: "CAJA HERRAMIENTAS PLASTICA 20\" TOTAL",
      marca: "TOTAL",
      cajasVienen: 7,
      cantPorCaja: 8,
      unidad: "UNI",
      precioMayor: 130.00,
      precioCaja: 125.00,
      precioEspecial: 120.00,
      precioRefUni: 125.00,
      precioCajaTotal: 1000.00,
      stockReserva: 7,
      agotado: false
    },
    {
      id: "P_7",
      codigo: "TP3XS102",
      detalle: "CAJA DE HERRAMIENTA APILABLE DE 19\"*5.8\"*12.5\" TOTAL",
      marca: "TOTAL",
      cajasVienen: 1,
      cantPorCaja: 4,
      unidad: "UNI",
      precioMayor: 100.00,
      precioCaja: 95.00,
      precioEspecial: 90.00,
      precioRefUni: 95.00,
      precioCajaTotal: 380.00,
      stockReserva: 0,
      agotado: true
    },
    {
      id: "P_8",
      codigo: "TPBXK0021",
      detalle: "CAJA HERRAMIENTAS PLASTICA SET/2EN1(14\"-17\") SS TOTAL",
      marca: "TOTAL",
      cajasVienen: 5,
      cantPorCaja: 6,
      unidad: "SET",
      precioMayor: 110.00,
      precioCaja: 105.00,
      precioEspecial: 100.00,
      precioRefUni: 105.00,
      precioCajaTotal: 630.00,
      stockReserva: 3,
      agotado: false
    },
    {
      id: "P_9",
      codigo: "TPBXK0032",
      detalle: "CAJA HERRAMIENTAS PLASTICA SET/3EN1(14\"-17\"-20\") REFORZADO TOTAL",
      marca: "TOTAL",
      cajasVienen: 8,
      cantPorCaja: 4,
      unidad: "SET",
      precioMayor: 240.00,
      precioCaja: 235.00,
      precioEspecial: 230.00,
      precioRefUni: 235.00,
      precioCajaTotal: 940.00,
      stockReserva: 3,
      agotado: false
    },
    {
      id: "P_10",
      codigo: "TPBXK0031",
      detalle: "CAJA HERRAMIENTAS PLASTICA SET/3EN1(14\"-17\"-20\") SS TOTAL",
      marca: "TOTAL",
      cajasVienen: 8,
      cantPorCaja: 4,
      unidad: "SET",
      precioMayor: 200.00,
      precioCaja: 195.00,
      precioEspecial: 190.00,
      precioRefUni: 195.00,
      precioCajaTotal: 780.00,
      stockReserva: 5,
      agotado: false
    },
    {
      id: "P_11",
      codigo: "THT571001",
      detalle: "CORTACERAMICA 1000MM C/ESTUCHE TOTAL",
      marca: "TOTAL",
      cajasVienen: 7,
      cantPorCaja: 2,
      unidad: "UNI",
      precioMayor: 830.00,
      precioCaja: 825.00,
      precioEspecial: 820.00,
      precioRefUni: 825.00,
      precioCajaTotal: 1650.00,
      stockReserva: 5,
      agotado: false
    },
    {
      id: "P_12",
      codigo: "THTC12008",
      detalle: "CORTACERAMICA 16MM/1200MM C/ESTUCHE TOTAL",
      marca: "TOTAL",
      cajasVienen: 5,
      cantPorCaja: 2,
      unidad: "UNI",
      precioMayor: 880.00,
      precioCaja: 870.00,
      precioEspecial: 860.00,
      precioRefUni: 870.00,
      precioCajaTotal: 1740.00,
      stockReserva: 3,
      agotado: false
    },
    {
      id: "P_13",
      codigo: "THT576003",
      detalle: "CORTACERAMICA 600MM C/ESTUCHE TOTAL",
      marca: "TOTAL",
      cajasVienen: 7,
      cantPorCaja: 2,
      unidad: "UNI",
      precioMayor: 635.00,
      precioCaja: 630.00,
      precioEspecial: 625.00,
      precioRefUni: 630.00,
      precioCajaTotal: 1260.00,
      stockReserva: 5,
      agotado: false
    },
    {
      id: "P_14",
      codigo: "THT511826",
      detalle: "CUCHILLO CARTONERO MB ECONOMICO TOTAL",
      marca: "TOTAL",
      cajasVienen: 14,
      cantPorCaja: 96,
      unidad: "UNI",
      precioMayor: 12.00,
      precioCaja: 11.00,
      precioEspecial: 9.00,
      precioRefUni: 11.00,
      precioCajaTotal: 1056.00,
      stockReserva: 14,
      agotado: false
    },
    {
      id: "P_15",
      codigo: "THKISD12203L",
      detalle: "DADO DE IMPACTO CORTO Y LARGO 1/2\" SET/20PZS TOTAL",
      marca: "TOTAL",
      cajasVienen: 12,
      cantPorCaja: 6,
      unidad: "SET",
      precioMayor: 290.00,
      precioCaja: 285.00,
      precioEspecial: 280.00,
      precioRefUni: 285.00,
      precioCajaTotal: 1710.00,
      stockReserva: 12,
      agotado: false
    },
    {
      id: "P_16",
      codigo: "THT121201",
      detalle: "DADOS CON CHICHARRA 1/2\"*SET/20PZS CR-V MALETA TOTAL",
      marca: "TOTAL",
      cajasVienen: 6,
      cantPorCaja: 6,
      unidad: "SET",
      precioMayor: 265.00,
      precioCaja: 260.00,
      precioEspecial: 255.00,
      precioRefUni: 260.00,
      precioCajaTotal: 1560.00,
      stockReserva: 6,
      agotado: false
    },
    {
      id: "P_17",
      codigo: "THT421942",
      detalle: "DADOS CON CHICHARRA 1/4\"-1/2\"*SET/94PZS CR-V MALETA TOTAL",
      marca: "TOTAL",
      cajasVienen: 23,
      cantPorCaja: 2,
      unidad: "SET",
      precioMayor: 610.00,
      precioCaja: 605.00,
      precioEspecial: 600.00,
      precioRefUni: 605.00,
      precioCajaTotal: 1210.00,
      stockReserva: 23,
      agotado: false
    },
    {
      id: "P_18",
      codigo: "THTS082016",
      detalle: "DADOS CON CHICHARRA 1/4\"-3/8\"-1/2\"*SET/201PZS CR-V MALETA TOTAL",
      marca: "TOTAL",
      cajasVienen: 17,
      cantPorCaja: 1,
      unidad: "SET",
      precioMayor: 1070.00,
      precioCaja: 1065.00,
      precioEspecial: 1060.00,
      precioRefUni: 1065.00,
      precioCajaTotal: 1065.00,
      stockReserva: 17,
      agotado: false
    },
    {
      id: "P_19",
      codigo: "THT421802",
      detalle: "DADOS CON CHICHARRA Y LLAVES COMB. 1/4\"-1/2\"*SET/82PZS TOTAL",
      marca: "TOTAL",
      cajasVienen: 2,
      cantPorCaja: 2,
      unidad: "SET",
      precioMayor: 555.00,
      precioCaja: 550.00,
      precioEspecial: 545.00,
      precioRefUni: 550.00,
      precioCajaTotal: 1100.00,
      stockReserva: 2,
      agotado: false
    },
    {
      id: "P_20",
      codigo: "THKTHP21396",
      detalle: "DADOS CON CHICHARRA Y LLAVES COMB. 1/4\"-3/8\"-1/2\"*SET/139PZS TOTAL",
      marca: "TOTAL",
      cajasVienen: 22,
      cantPorCaja: 1,
      unidad: "SET",
      precioMayor: 995.00,
      precioCaja: 990.00,
      precioEspecial: 985.00,
      precioRefUni: 990.00,
      precioCajaTotal: 990.00,
      stockReserva: 22,
      agotado: false
    },
    {
      id: "P_21",
      codigo: "THKISD12143L",
      detalle: "DADOS DE IMPACTO LARGO 1/2\" SET/14PZS TOTAL",
      marca: "TOTAL",
      cajasVienen: 2,
      cantPorCaja: 6,
      unidad: "SET",
      precioMayor: 280.00,
      precioCaja: 275.00,
      precioEspecial: 270.00,
      precioRefUni: 275.00,
      precioCajaTotal: 1650.00,
      stockReserva: 0,
      agotado: true
    }
  ]
};

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
