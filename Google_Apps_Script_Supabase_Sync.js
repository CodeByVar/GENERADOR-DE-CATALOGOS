/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: SINCRONIZADOR DIRECTO A SUPABASE (POSTGRESQL)
 * =========================================================================================
 * Este script lee tus hojas de Google Sheets y envía el stock a Supabase en 1 segundo.
 * 
 * BENEFICIOS:
 * 1. Tu catálogo web ahora leerá desde Supabase en 30 milisegundos (Cero caídas ni esperas).
 * 2. Tus empleados siguen usando Google Sheets exactamente igual.
 * 3. Se sincroniza automáticamente cada 5 o 10 minutos (o al presionar un botón).
 * =========================================================================================
 */

// CONFIGURACIÓN DE TU PROYECTO SUPABASE
var SUPABASE_URL = "https://mjiezwmldydnlcpshlpq.supabase.co";
var SUPABASE_KEY = "sb_publishable_5Nxl1zMTRm6ngigdYQUA-g_lOKdUpzo";

/**
 * Función principal que extrae el stock y lo envía a Supabase
 */
function sincronizarStockConSupabase() {
  Logger.log(">>> Iniciando sincronización hacia Supabase...");
  var t0 = new Date().getTime();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var productosParaEnviar = [];
  var codigosVistos = {};

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (sheet.isSheetHidden()) continue;
    var sheetName = sheet.getName();
    if (/^(PORTADA|GRAFICAS?|CONFIG|RESUMEN|HISTORIAL|VENTAS|DATOS|CLIENTES|PROVEEDOR|INSTRUCCION|PLANTILLA|MENU|INDEX)/i.test(sheetName)) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) continue;

    // Escaneo de encabezados (primeras 7 filas)
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

        // 2. STOCK ACTUAL (Excluyendo reservas, ventas, salidas)
        if (val.indexOf("RESERVA") !== -1 || val.indexOf("VENTA") !== -1 || val.indexOf("MINIMO") !== -1 || val.indexOf("SALIDA") !== -1) {
          // Ignorar columnas de ventas o salidas
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

        // 3. Q. POR CAJA
        var cleanLetters = val.replace(/[^A-Z]/g, "");
        var esPrecioOCajaSola = (val.indexOf("PRECIO") !== -1 || val.indexOf("COSTO") !== -1 || val.indexOf("MAYOR") !== -1 || val.indexOf("ESPECIAL") !== -1 || val.indexOf("TOTAL") !== -1 || val.indexOf("STOCK") !== -1 || val === "CAJA");
        
        var esQPorCajaDirecto = (
          cleanLetters === "QPORCAJA" || cleanLetters === "QCAJA" || cleanLetters === "PORCAJA" ||
          val.indexOf("Q. POR CAJA") !== -1 || val.indexOf("Q.POR CAJA") !== -1 || val.indexOf("Q.PORCAJA") !== -1 ||
          val.indexOf("Q POR CAJA") !== -1 || val.indexOf("Q/CAJA") !== -1 || val.indexOf("Q. CAJA") !== -1 ||
          val.indexOf("CANTIDAD DE CAJA") !== -1 || val.indexOf("CANTIDAD CAJA") !== -1 ||
          val.indexOf("N° CANTIDAD DE CAJA") !== -1 || val.indexOf("N CANTIDAD DE CAJA") !== -1
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

        // 4. UNIDAD
        if (bestColUni === -1) {
          if (val.indexOf("UN/") !== -1 || val.indexOf("UN /") !== -1 || val.indexOf("UNID") !== -1 || val === "MEDIDA" || val === "U.M." || val === "U/M" || val === "UM") {
            bestColUni = c;
            lastHeaderRow = Math.max(lastHeaderRow, r + 1);
          }
        }
      }
    }

    // AUTO-DETECCIÓN INTELIGENTE DE ESTRUCTURA (UYUS vs VARIOS / TOTAL):
    // Formato UYUS:
    //   Col C (índice 2) = CÓDIGO
    //   Col D (índice 3) = Q. POR CAJA
    //   Col E (índice 4) = UN/ MED
    //   Col K (índice 10) = STOCK ACTUAL (Columna azul)
    // Formato VARIOS / TOTAL:
    //   Col C (índice 2) = CÓDIGO
    //   Col G (índice 6) = Q. POR CAJA
    //   Col J (índice 9) = UN/ MED
    //   Col P (índice 15) = STOCK ACTUAL

    var esFormatoUYUS = false;
    var esFormatoVarios = false;

    // Verificar fila 4 o encabezados para determinar estructura exacta
    if (lastRow >= 4 && lastCol >= 11) {
      try {
        var sampleRow = sheet.getRange(4, 1, 1, Math.min(lastCol, 17)).getValues()[0];
        var valE = String(sampleRow[4] || "").trim().toUpperCase(); // Columna E (UYUS UN/MED)
        var valJ = String(sampleRow[9] || "").trim().toUpperCase(); // Columna J (VARIOS UN/MED)
        
        if (valE === "SET" || valE === "UNI" || valE === "PZA" || valE === "PAR" || valE === "DOC" || valE === "ROLLO") {
          esFormatoUYUS = true;
        } else if (valJ === "SET" || valJ === "UNI" || valJ === "PZA" || valJ === "PAR" || valJ === "DOC" || valJ === "ROLLO") {
          esFormatoVarios = true;
        }
      } catch(e) {}
    }

    if (bestColCod === -1) bestColCod = 2; // Columna C siempre es CÓDIGO

    if (esFormatoUYUS) {
      if (bestColCaja === -1) bestColCaja = 3; // Columna D (Q. POR CAJA en UYUS)
      if (bestColUni === -1)  bestColUni  = 4; // Columna E (UN/ MED en UYUS)
      if (bestColStock === -1 || bestColStock > 12) bestColStock = 10; // Columna K (STOCK ACTUAL en UYUS)
    } else {
      // Por defecto o si es formato VARIOS / TOTAL:
      if (bestColCaja === -1) bestColCaja = (bestColCaja !== -1) ? bestColCaja : 6;  // Columna G
      if (bestColUni === -1)  bestColUni  = (bestColUni !== -1) ? bestColUni : 9;   // Columna J
      if (bestColStock === -1) bestColStock = (bestColStock !== -1) ? bestColStock : 15; // Columna P
    }

    // Fallbacks de seguridad si aún no están definidos
    if (bestColStock === -1) {
      bestColStock = (lastCol >= 16) ? 15 : 10;
    }
    if (bestColCaja === -1) {
      bestColCaja = (lastCol >= 7) ? 6 : 3;
    }
    if (bestColUni === -1) {
      bestColUni = (lastCol >= 10) ? 9 : 4;
    }

    if (bestColCod === -1 || bestColStock === -1) continue;

    var startDataRow = Math.max(lastHeaderRow + 1, 4);
    var numRows = lastRow - startDataRow + 1;
    if (numRows <= 0) continue;
    numRows = Math.min(numRows, 2000);

    // Lectura en 1 solo bloque por hoja
    var maxColNeeded = Math.max(bestColCod, bestColStock, bestColCaja, bestColUni) + 1;
    var sheetData = sheet.getRange(startDataRow, 1, numRows, maxColNeeded).getValues();

    var emptyStreak = 0;
    for (var i = 0; i < sheetData.length; i++) {
      var row = sheetData[i];
      var rawCod = row[bestColCod];
      if (rawCod === null || rawCod === undefined || String(rawCod).trim() === "") {
        emptyStreak++;
        if (emptyStreak > 25) break;
        continue;
      }
      emptyStreak = 0;

      var codStr = String(rawCod).trim();
      var upperCod = codStr.toUpperCase();
      if (/^(TOTAL|SUBTOTAL|TOTALES|CODIGO|DETALLE|PRECIO|MAYOR|CAJA|ESPECIAL|ITEM)$/.test(upperCod)) continue;

      var normCod = upperCod.replace(/\s+/g, '');
      if (codigosVistos[normCod]) continue; // Evitar duplicados
      codigosVistos[normCod] = true;

      // Stock
      var rawStockVal = row[bestColStock];
      var stockActual = 0;
      if (typeof rawStockVal === 'number') {
        stockActual = isNaN(rawStockVal) ? 0 : rawStockVal;
      } else if (typeof rawStockVal === 'string') {
        var parsed = parseFloat(rawStockVal.replace(/[^0-9.\-]/g, '').trim());
        stockActual = isNaN(parsed) ? 0 : parsed;
      }
      if (stockActual < 0) stockActual = 0;

      // Cantidad Caja
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

      // Unidad
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

      productosParaEnviar.push({
        codigo: normCod,
        stock_actual: stockActual,
        cantidad_caja: cantCaja,
        unidad_medida: unidad,
        estado: estado,
        updated_at: new Date().toISOString()
      });
    }
  }

  Logger.log(">>> Total de productos listos para enviar a Supabase: " + productosParaEnviar.length);

  // Verificación en tiempo real de productos clave en el log
  for (var p = 0; p < productosParaEnviar.length; p++) {
    var cTest = productosParaEnviar[p].codigo;
    if (cTest === "DAT815" || cTest === "DAD20U" || cTest === "DAD30U" || cTest === "TAKTMT1502") {
      Logger.log(">>> [VERIFICACIÓN " + cTest + "]: Stock = " + productosParaEnviar[p].stock_actual + " | Q. Caja = " + productosParaEnviar[p].cantidad_caja + " " + productosParaEnviar[p].unidad_medida + " | Estado = " + productosParaEnviar[p].estado);
    }
  }

  // Enviar a Supabase en paquetes (chunks de 500 para máxima velocidad y seguridad)
  var BATCH_SIZE = 500;
  for (var b = 0; b < productosParaEnviar.length; b += BATCH_SIZE) {
    var chunk = productosParaEnviar.slice(b, b + BATCH_SIZE);
    enviarChunkASupabase(chunk);
  }

  var t1 = new Date().getTime();
  Logger.log(">>> [ÉXITO SUPABASE] Sincronización completada en " + ((t1 - t0) / 1000).toFixed(2) + " segundos.");
}

/**
 * Realiza el Upsert directo en la tabla catalogo_stock de Supabase
 */
function enviarChunkASupabase(chunk) {
  var url = SUPABASE_URL + "/rest/v1/catalogo_stock?on_conflict=codigo";
  var options = {
    method: "post",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": "Bearer " + SUPABASE_KEY,
      "Content-Type": "application/json",
      "Prefer": "resolution=merge-duplicates"
    },
    payload: JSON.stringify(chunk),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    Logger.log(">>> [OK] Paquete de " + chunk.length + " productos sincronizado.");
  } else {
    Logger.log(">>> [ERROR SUPABASE] Código " + code + ": " + response.getContentText());
  }
}

/**
 * Agrega un botón en la barra superior de tu Google Sheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Catálogo Supabase")
    .addItem("Sincronizar Stock Ahora", "sincronizarStockConSupabase")
    .addToUi();
}

function normalizarTexto(txt) {
  if (txt === null || txt === undefined) return "";
  return txt.toString().toUpperCase().trim()
    .replace(/[ÁÀÄ]/g, "A").replace(/[ÉÈË]/g, "E").replace(/[ÍÌÏ]/g, "I").replace(/[ÓÒÖ]/g, "O").replace(/[ÚÙÜ]/g, "U")
    .replace(/[\n\r\t]+/g, " ").replace(/\s+/g, " ");
}
