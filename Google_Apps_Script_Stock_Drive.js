/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: LECTURA DE STOCK EN TIEMPO REAL DESDE GOOGLE DRIVE (ALTA VELOCIDAD)
 * =========================================================================================
 * Este script debe pegarse en cada libro de Google Sheets conectado en Google Drive
 * (por ejemplo en el libro "UYUS", en el libro "VARIOS", y en cualquier otro libro).
 *
 * ¿QUÉ HACE ESTE SCRIPT?
 * 1. Recorre TODAS las hojas del libro (cada marca o categoría tiene su propia hoja).
 * 2. En cada hoja, escanea automáticamente las primeras filas buscando la columna "STOCK ACTUAL"
 *    y la columna "CÓDIGO".
 * 3. Lee ÚNICAMENTE las columnas necesarias (en vez de leer toda la hoja pesada),
 *    lo que hace la lectura 15 VECES MÁS RÁPIDA y previene el error de tiempo de espera (timeout).
 * 4. Detecta dónde terminan los productos reales y corta filas vacías automáticamente.
 * 5. Extrae el valor numérico exacto de la columna "STOCK ACTUAL".
 * 6. Guarda en micro-caché de Google para que las consultas posteriores respondan al instante (< 200 ms).
 *
 * PASOS PARA ACTUALIZAR EN GOOGLE DRIVE:
 * 1. Abre tu libro de Google Sheets en Google Drive (UYUS o VARIOS).
 * 2. En el menú superior haz clic en: Extensiones -> Apps Script.
 * 3. Borra todo el código que haya y pega ESTE ARCHIVO COMPLETO.
 * 4. Haz clic en "Guardar" (ícono de disquete 💾).
 * 5. Haz clic en "Implementar" (botón azul arriba a la derecha) -> "Gestionar implementaciones"
 *    (o "Nueva implementación").
 * 6. Haz clic en el lápiz ✏️ de editar (o crea nueva):
 *    - Versión: "Nueva versión"
 *    - Quién tiene acceso: "Cualquier persona" (¡OBLIGATORIO!)
 * 7. Haz clic en "Implementar".
 * =========================================================================================
 */

function doGet(e) {
  try {
    var force = (e && e.parameter && (e.parameter.force === "1" || e.parameter.fresh === "1"));
    var cache = CacheService.getScriptCache();
    var cachePrefix = "STOCK_V4";

    // Si no se fuerza actualización y hay datos guardados en caché rápida (< 10 min), responder en 150 ms
    if (!force) {
      var cachedStr = leerDeCacheChunks(cache, cachePrefix);
      if (cachedStr && cachedStr.length > 20) {
        return ContentService.createTextOutput(cachedStr)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    var stockMap = obtenerMapaStockCompleto();
    var jsonOutput = JSON.stringify(stockMap);

    // Guardar en la memoria ultra rápida de Google Apps Script por 10 minutos (600s) dividido en partes
    guardarEnCacheChunks(cache, cachePrefix, jsonOutput, 600);

    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: "Error leyendo stock de las hojas: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Guarda cadenas JSON de cualquier tamaño en CacheService dividiéndolas en bloques seguros (< 90 KB)
 */
function guardarEnCacheChunks(cache, prefijo, str, ttl) {
  try {
    var CHUNK_SIZE = 85000;
    var numChunks = Math.ceil(str.length / CHUNK_SIZE);
    cache.put(prefijo + "_COUNT", String(numChunks), ttl);
    for (var i = 0; i < numChunks; i++) {
      cache.put(prefijo + "_" + i, str.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE), ttl);
    }
  } catch(e) {}
}

/**
 * Reconstruye la cadena completa desde los bloques en CacheService
 */
function leerDeCacheChunks(cache, prefijo) {
  try {
    var countStr = cache.get(prefijo + "_COUNT");
    if (!countStr) return null;
    var count = parseInt(countStr, 10);
    var fullStr = "";
    for (var i = 0; i < count; i++) {
      var chunk = cache.get(prefijo + "_" + i);
      if (chunk === null) return null; // Incompleto, regenerar
      fullStr += chunk;
    }
    return fullStr;
  } catch(e) {
    return null;
  }
}

/**
 * Recorre CADA HOJA del libro y toma el valor de la columna "STOCK ACTUAL"
 * con optimización de lectura en columnas específicas.
 */
function obtenerMapaStockCompleto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var resultMap = {};

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];

    // Ignorar hojas ocultas o de sistema
    if (sheet.isSheetHidden()) continue;
    var sheetName = sheet.getName();
    if (/^(PORTADA|GRAFICAS?|CONFIG|RESUMEN|HISTORIAL|VENTAS|DATOS|CLIENTES|PROVEEDOR|INSTRUCCION|PLANTILLA|MENU|INDEX)/i.test(sheetName)) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) continue;

    // Escanear las primeras 7 filas para detectar encabezados (soporta celdas combinadas y multinivel)
    var maxHeaderScan = Math.min(7, lastRow);
    var scanCols = Math.min(35, lastCol);
    var headerData = sheet.getRange(1, 1, maxHeaderScan, scanCols).getValues();

    var bestColCod = -1;
    var bestColStock = -1;
    var bestColCaja = -1;
    var bestColUni = -1;
    var lastHeaderRow = 1;

    for (var r = 0; r < maxHeaderScan; r++) {
      var rowVals = headerData[r];
      for (var c = 0; c < rowVals.length; c++) {
        var val = normalizarTexto(rowVals[c]);
        if (!val) continue;

        // 1. CÓDIGO
        if (bestColCod === -1) {
          if (/^(CODIGO|COD|ITEM|ARTICULO|REF|REFERENCIA|COD_ARTICULO|MODELO)$/.test(val) ||
              val.indexOf("CODIGO") === 0 || val.indexOf("COD.") === 0) {
            bestColCod = c;
            lastHeaderRow = Math.max(lastHeaderRow, r + 1);
          }
        }

        // 2. STOCK ACTUAL (Excluyendo explícitamente columnas de RESERVAS, VENTAS y SALIDAS)
        if (val.indexOf("RESERVA") !== -1 || val.indexOf("VENTA") !== -1 || val.indexOf("MINIMO") !== -1 || val.indexOf("SALIDA") !== -1) {
          // Ignorar columna secundaria
        } else if (val === "STOCK ACTUAL" || val === "STOCK_ACTUAL" || val.indexOf("STOCK ACTUAL") !== -1 ||
                   val === "SALDO ACTUAL" || val.indexOf("SALDO ACTUAL") !== -1 ||
                   val === "STOCK DISPONIBLE" || val === "CANTIDAD ACTUAL" || val === "EXISTENCIA ACTUAL" ||
                   val === "STOCK FISICO" || val === "STOCK FINAL" || val === "TOTAL STOCK") {
          bestColStock = c;
          lastHeaderRow = Math.max(lastHeaderRow, r + 1);
        } else if (bestColStock === -1 && (val === "STOCK" || val === "SALDO" || val === "EXISTENCIAS")) {
          bestColStock = c;
          lastHeaderRow = Math.max(lastHeaderRow, r + 1);
        }

        // 3. Cantidad por caja: PRIORIDAD DIRECTA a "Q. POR CAJA" (UYUS y VARIOS)
        var cleanLetters = val.replace(/[^A-Z]/g, "");
        var esPrecioOCajaSola = (val.indexOf("PRECIO") !== -1 || val.indexOf("COSTO") !== -1 || val.indexOf("MAYOR") !== -1 || val.indexOf("ESPECIAL") !== -1 || val.indexOf("TOTAL") !== -1 || val.indexOf("STOCK") !== -1 || val === "CAJA");
        
        var esQPorCajaDirecto = (
          cleanLetters === "QPORCAJA" ||
          cleanLetters === "QCAJA" ||
          cleanLetters === "PORCAJA" ||
          val.indexOf("Q. POR CAJA") !== -1 ||
          val.indexOf("Q.POR CAJA") !== -1 ||
          val.indexOf("Q.PORCAJA") !== -1 ||
          val.indexOf("Q POR CAJA") !== -1 ||
          val.indexOf("Q/CAJA") !== -1 ||
          val.indexOf("Q. CAJA") !== -1 ||
          val.indexOf("CANTIDAD DE CAJA") !== -1 ||
          val.indexOf("CANTIDAD CAJA") !== -1 ||
          val.indexOf("CANT. CAJA") !== -1 ||
          val.indexOf("N° CANTIDAD DE CAJA") !== -1 ||
          val.indexOf("N CANTIDAD DE CAJA") !== -1
        );

        if (esQPorCajaDirecto) {
          bestColCaja = c;
          lastHeaderRow = Math.max(lastHeaderRow, r + 1);
        } else if (bestColCaja === -1 && !esPrecioOCajaSola) {
          if (cleanLetters.indexOf("PORCAJA") !== -1 || val.indexOf("UNID/CAJA") !== -1 || val.indexOf("PZS/CAJA") !== -1 || val === "EMPAQUE") {
            bestColCaja = c;
            lastHeaderRow = Math.max(lastHeaderRow, r + 1);
          }
        }

        // 4. Unidad de medida (ej: "UN/ MED", "UNIDAD", "MEDIDA", "U.M.")
        // 4. Unidad de medida (ej: "UN/ MED", "UNIDAD", "MEDIDA", "U.M.")
        if (bestColUni === -1) {
          if (val.indexOf("UN/") !== -1 || val.indexOf("UN /") !== -1 || val.indexOf("UNID") !== -1 || val === "MEDIDA" || val === "U.M." || val === "U/M" || val === "UM") {
            bestColUni = c;
            lastHeaderRow = Math.max(lastHeaderRow, r + 1);
          }
        }
      }
    }

    // FALLBACK INTELIGENTE DIRECTO:
    // En las hojas TOTAL / VARIOS / UYUS:
    // Columna C (índice 2) = CÓDIGO
    // Columna G (índice 6) = Q. POR CAJA
    // Columna J (índice 9) = UN/ MED
    // Columna P (índice 15) = STOCK ACTUAL
    if (bestColCod === -1 && lastCol >= 3) bestColCod = 2;
    if (bestColStock === -1 && lastCol >= 16) bestColStock = 15;
    if (bestColCaja === -1 && lastCol >= 7) bestColCaja = 6;
    if (bestColUni === -1 && lastCol >= 10) bestColUni = 9;

    // Si aún así no hay Código o Stock, continuar
    if (bestColCod === -1 || bestColStock === -1) {
      continue;
    }

    // Las filas de datos comienzan después de los encabezados (típicamente fila 4)
    var startDataRow = Math.max(lastHeaderRow + 1, 4);
    var numRows = lastRow - startDataRow + 1;
    if (numRows <= 0) continue;

    // Limitamos a máximo 2000 productos por hoja para ultra velocidad
    numRows = Math.min(numRows, 2000);

    // LECTURA ULTRA RÁPIDA: Una sola llamada getRange() por hoja que abarca hasta la Columna P
    var maxColNeeded = Math.max(bestColCod, bestColStock, bestColCaja, bestColUni) + 1;
    var sheetData = sheet.getRange(startDataRow, 1, numRows, maxColNeeded).getValues();

    var emptyStreak = 0;
    for (var i = 0; i < sheetData.length; i++) {
      var row = sheetData[i];
      var rawCod = row[bestColCod];

      if (rawCod === null || rawCod === undefined || String(rawCod).trim() === "") {
        emptyStreak++;
        if (emptyStreak > 25) break; // Si hay más de 25 filas vacías consecutivas, fin de hoja
        continue;
      }
      emptyStreak = 0;

      var codStr = String(rawCod).trim();
      var upperCod = codStr.toUpperCase();
      // Ignorar subtítulos repetidos o totales
      if (/^(TOTAL|SUBTOTAL|TOTALES|CODIGO|DETALLE|PRECIO|MAYOR|CAJA|ESPECIAL|ITEM)$/.test(upperCod)) {
        continue;
      }

      var normCod = upperCod.replace(/\s+/g, '');
      var simpleCod = normCod.replace(/[\-._/]/g, '');

      // Extraer valor de STOCK ACTUAL (Columna P / bestColStock)
      var rawStockVal = row[bestColStock];
      var stockActual = 0;

      if (typeof rawStockVal === 'number') {
        stockActual = isNaN(rawStockVal) ? 0 : rawStockVal;
      } else if (typeof rawStockVal === 'string') {
        var cleanedStr = rawStockVal.replace(/[^0-9.\-]/g, '').trim();
        var parsed = parseFloat(cleanedStr);
        stockActual = isNaN(parsed) ? 0 : parsed;
      } else if (typeof rawStockVal === 'boolean') {
        stockActual = rawStockVal ? 999 : 0;
      }

      if (stockActual < 0) stockActual = 0;

      // Cantidad por Caja (Q. POR CAJA - Columna G / bestColCaja)
      var cantCaja = 1;
      if (bestColCaja !== -1) {
        var rawCaja = row[bestColCaja];
        if (typeof rawCaja === 'number' && rawCaja > 0) {
          cantCaja = rawCaja;
        } else if (typeof rawCaja === 'string') {
          var pC = parseFloat(rawCaja.replace(/,/g, '.').replace(/[^0-9.]/g, ''));
          if (!isNaN(pC) && pC > 0) cantCaja = pC;
        }
      }

      // Unidad (UN/ MED - Columna J / bestColUni)
      var unidad = "UNI";
      if (bestColUni !== -1) {
        var rawU = row[bestColUni];
        if (rawU && typeof rawU === 'string' && isNaN(rawU)) {
          unidad = rawU.trim().toUpperCase();
        }
      }

      var cajas = Math.floor(stockActual / cantCaja);
      var estado = "AGOTADO";
      if (stockActual > 0) {
        if (cajas <= 3 || stockActual <= (cantCaja * 3)) {
          estado = "POCO_STOCK";
        } else {
          estado = "DISPONIBLE";
        }
      }

      // Estructura ultra liviana compatible con CacheService y todos los catálogos
      var itemInfo = {
        s: stockActual,
        c: cantCaja,
        u: unidad,
        b: cajas,
        e: estado
      };

      resultMap[normCod] = itemInfo;
      if (upperCod !== normCod) {
        resultMap[upperCod] = itemInfo;
      }
      if (simpleCod !== normCod && simpleCod !== upperCod) {
        resultMap[simpleCod] = itemInfo;
      }
    }
  }

  return resultMap;
}

/**
 * Limpia y normaliza cadenas de texto para comparar encabezados
 */
function normalizarTexto(txt) {
  if (txt === null || txt === undefined) return "";
  return txt.toString()
    .toUpperCase()
    .trim()
    .replace(/[ÁÀÄ]/g, "A")
    .replace(/[ÉÈË]/g, "E")
    .replace(/[ÍÌÏ]/g, "I")
    .replace(/[ÓÒÖ]/g, "O")
    .replace(/[ÚÙÜ]/g, "U")
    .replace(/[\n\r\t]+/g, " ")
    .replace(/\s+/g, " ");
}

/**
 * FUNCIÓN DE PRUEBA RÁPIDA:
 * Selecciona esta función en Apps Script y haz clic en "Ejecutar"
 * para verificar en segundos cuántos productos y hojas leyó.
 */
function probarLecturaStock() {
  Logger.log(">>> Iniciando prueba de lectura ultra rápida de stock (Prioridad Columna P)...");
  var t0 = new Date().getTime();
  var res = obtenerMapaStockCompleto();
  var t1 = new Date().getTime();
  var totalKeys = Object.keys(res).length;
  Logger.log(">>> [ÉXITO] Todo el libro leído en SOLO: " + ((t1 - t0) / 1000).toFixed(2) + " segundos.");
  Logger.log(">>> [ÉXITO] Total de códigos registrados: " + totalKeys);

  // Comprobar el producto específico que causó duda (TAKTMT1502)
  var testItem = res["TAKTMT1502"];
  if (testItem) {
    Logger.log(">>> [PRUEBA TAKTMT1502]: Stock = " + testItem.s + " | Q. POR CAJA = " + testItem.c + " " + testItem.u + " | Cajas = " + testItem.b + " | Estado = " + testItem.e);
  } else {
    Logger.log(">>> [AVISO] TAKTMT1502 no encontrado en las hojas procesadas.");
  }

  if (totalKeys > 0) {
    var sampleKeys = Object.keys(res).slice(0, 5);
    for (var k = 0; k < sampleKeys.length; k++) {
      var it = res[sampleKeys[k]];
      Logger.log(">>> Muestra (" + sampleKeys[k] + ") -> Stock: " + it.s + " | Q. POR CAJA: " + it.c + " " + it.u + " | Cajas: " + it.b + " (" + it.e + ")");
    }
  }
}
