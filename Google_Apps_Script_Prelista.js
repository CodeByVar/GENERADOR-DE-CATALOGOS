/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: LECTURA SEGURA DE PRELISTA (MERCADERÍA EN CAMINO / RESERVAS)
 * =========================================================================================
 * IMPORTADORA RIVERO - BOLIVIA
 * 
 * ¿QUÉ HACE ESTE SCRIPT?
 * 1. Lee tu hoja activa (ej: "Hoja 48", "Hoja 47", etc.) del libro "I. RIVERO 2026".
 * 2. Extrae ÚNICAMENTE los datos públicos para los clientes:
 *    - Código de producto
 *    - Detalle / Descripción
 *    - Cantidad total de cajas que vienen (Q. DE CAJAS)
 *    - Cantidad de unidades por caja (Q. POR CAJA / PACKING)
 *    - Unidad de medida (UN/MED: SET, UNI, etc.)
 *    - Precios de venta (Precio Caja, Precio Mayor, Precio Especial)
 *    - Cajas libres para reserva (STOCK DE RESERVAS)
 *    - Tipo de cambio
 * 
 * 3. 🔒 SEGURIDAD TOTAL:
 *    - OMITE y BLOQUEA al 100% la columna de COSTO y porcentajes de ganancia.
 *    - OMITE y BLOQUEA los nombres de los clientes internos (Karen, Rubi, Marcelo, etc.).
 *    - Nadie puede ver quién más está reservando ni cuánto te costó la mercadería.
 * 
 * 4. VELOCIDAD:
 *    - Guarda en caché rápida (CacheService) para responder en menos de 100 milisegundos.
 * 
 * INSTRUCCIONES DE INSTALACIÓN (En 2 minutos):
 * 1. Abre tu Google Sheets "I. RIVERO 2026".
 * 2. En el menú superior haz clic en: Extensiones -> Apps Script.
 * 3. Crea un archivo nuevo o reemplaza el código con TODO este archivo.
 * 4. Guarda (ícono de disquete 💾).
 * 5. Haz clic en "Implementar" (botón azul arriba a la derecha) -> "Nueva implementación".
 *    - Tipo: Aplicación web
 *    - Descripción: Prelista Rivero V1
 *    - Ejecutar como: Yo (tu cuenta de Google)
 *    - Quién tiene acceso: Cualquier persona
 * 6. Haz clic en "Implementar" y copia la URL que te da (termina en /exec).
 * =========================================================================================
 */

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var sheetParam = params.sheet || "";
    var force = (params.force === "1" || params.fresh === "1");
    var listOnly = (params.listSheets === "1");

    var cache = CacheService.getScriptCache();
    var cacheKey = "PRELISTA_" + (sheetParam ? sheetParam.replace(/\s+/g, '_') : "DEFAULT");

    // Si se pide lista de hojas disponibles
    if (listOnly) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheets = ss.getSheets();
      var sheetNames = [];
      for (var s = 0; s < sheets.length; s++) {
        var sh = sheets[s];
        if (!sh.isSheetHidden() && /^Hoja\s*\d+/i.test(sh.getName())) {
          sheetNames.push(sh.getName());
        }
      }
      sheetNames.sort(function(a, b) {
        var matchA = a.match(/\d+/);
        var matchB = b.match(/\d+/);
        var numA = matchA ? parseInt(matchA[0], 10) : 0;
        var numB = matchB ? parseInt(matchB[0], 10) : 0;
        if (numA !== numB) return numB - numA;
        return a.localeCompare(b);
      });
      return jsonResponse({ success: true, sheets: sheetNames });
    }

    // Si está en caché y no se forza actualización, responder al instante
    if (!force) {
      var cached = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    var result = extraerDatosPrelista(sheetParam);
    var jsonStr = JSON.stringify(result);

    // Guardar en caché por 60 segundos
    try {
      cache.put(cacheKey, jsonStr, 60);
    } catch(cErr) {}

    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return jsonResponse({
      success: false,
      error: "Error leyendo prelista: " + err.toString()
    });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizarTexto(txt) {
  if (!txt) return "";
  return String(txt).trim().toUpperCase()
    .replace(/[ÁÀÄÂ]/g, "A")
    .replace(/[ÉÈËÊ]/g, "E")
    .replace(/[ÍÌÏÎ]/g, "I")
    .replace(/[ÓÒÖÔ]/g, "O")
    .replace(/[ÚÙÜÛ]/g, "U");
}

function extraerDatosPrelista(targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = null;

  var sheets = ss.getSheets();
  var availableSheets = [];
  for (var s = 0; s < sheets.length; s++) {
    var curName = sheets[s].getName();
    if (!sheets[s].isSheetHidden() && (/^Hoja\s*\d+/i.test(curName) || /PRE/i.test(curName) || /LLEGA/i.test(curName))) {
      availableSheets.push(curName);
    }
  }

  // Ordenar de forma natural descendente por número (ej: Hoja 50, Hoja 49, Hoja 48...)
  availableSheets.sort(function(a, b) {
    var matchA = a.match(/\d+/);
    var matchB = b.match(/\d+/);
    var numA = matchA ? parseInt(matchA[0], 10) : 0;
    var numB = matchB ? parseInt(matchB[0], 10) : 0;
    if (numA !== numB) return numB - numA;
    return a.localeCompare(b);
  });

  // Si se solicitaron múltiples hojas separadas por coma o '+'
  if (targetSheetName && (targetSheetName.indexOf(',') !== -1 || targetSheetName.indexOf('+') !== -1)) {
    var targetList = targetSheetName.split(/[,+]/).map(function(item) { return item.trim(); }).filter(Boolean);
    var prodsMap = {};
    var mergedList = [];
    for (var i = 0; i < targetList.length; i++) {
      var singleRes = extraerDatosPrelista(targetList[i]);
      if (singleRes && singleRes.success && Array.isArray(singleRes.productos)) {
        for (var p = 0; p < singleRes.productos.length; p++) {
          var prod = singleRes.productos[p];
          var cod = String(prod.codigo || '').toUpperCase().trim();
          if (prodsMap[cod]) {
            prodsMap[cod].cajasVienen = (prodsMap[cod].cajasVienen || 0) + (prod.cajasVienen || 0);
            prodsMap[cod].stockReserva = (prodsMap[cod].stockReserva || 0) + (prod.stockReserva || 0);
            if (!prodsMap[cod]._origenHojas) prodsMap[cod]._origenHojas = [];
            if (prodsMap[cod]._origenHojas.indexOf(targetList[i]) === -1) {
              prodsMap[cod]._origenHojas.push(targetList[i]);
            }
          } else {
            prod._origenHojas = [targetList[i]];
            prodsMap[cod] = prod;
            mergedList.push(prod);
          }
        }
      }
    }
    return {
      success: true,
      sheet: targetList.join(' + '),
      availableSheets: availableSheets,
      totalProductos: mergedList.length,
      productos: mergedList
    };
  }

  // 1. Seleccionar la hoja indicada o buscar la más reciente (ej: Hoja 48)
  if (targetSheetName) {
    sheet = ss.getSheetByName(targetSheetName);
  }

  if (!sheet) {
    // Si no se especificó hoja, buscar la más reciente que tenga datos reales (>= 3 filas)
    for (var k = 0; k < availableSheets.length; k++) {
      var candidate = ss.getSheetByName(availableSheets[k]);
      if (candidate && candidate.getLastRow() >= 3) {
        sheet = candidate;
        break;
      }
    }
    // Si ninguna tiene datos o todas están vacías, tomar la primera disponible
    if (!sheet && availableSheets.length > 0) {
      sheet = ss.getSheetByName(availableSheets[0]);
    } else if (!sheet) {
      sheet = ss.getActiveSheet();
    }
  }

  if (!sheet) {
    return { success: false, error: "No se encontró ninguna hoja válida en el documento." };
  }

  var actualSheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 3) {
    return {
      success: true,
      sheet: actualSheetName,
      availableSheets: availableSheets,
      productos: []
    };
  }

  // 2. Escanear encabezados en las primeras 5 filas
  var headerScanRows = Math.min(5, lastRow);
  var headerData = sheet.getRange(1, 1, headerScanRows, lastCol).getValues();

  var colDetalle = -1;
  var colCodigo = -1;
  var colQCajas = -1;
  var colQPorCaja = -1;
  var colUnMed = -1;
  var colPrecioMayor = -1;
  var colPrecioCaja = -1;
  var colPrecioEspecial = -1;
  var colTipoCambio = -1;
  var colStockReservas = -1;
  var headerRowIdx = 1; // 0-based

  for (var r = 0; r < headerScanRows; r++) {
    var row = headerData[r];
    for (var c = 0; c < row.length; c++) {
      var cellVal = normalizarTexto(row[c]);
      if (!cellVal) continue;

      if (cellVal === "DETALLE" || cellVal === "DESCRIPCION" || cellVal === "PRODUCTO" || cellVal === "ARTICULO") {
        colDetalle = c;
        headerRowIdx = r;
      } else if (cellVal === "CODIGO" || cellVal === "COD" || cellVal === "COD.") {
        colCodigo = c;
        headerRowIdx = r;
      } else if (((cellVal.indexOf("CAJA") >= 0 && (cellVal.indexOf("Q") >= 0 || cellVal.indexOf("CANT") >= 0 || cellVal.indexOf("TOT") >= 0)) || cellVal === "CAJAS" || cellVal === "CJS" || cellVal === "Q. CAJAS" || cellVal === "Q. DE CAJAS") && cellVal.indexOf("POR") === -1 && cellVal.indexOf("X") === -1 && cellVal.indexOf("PRECIO") === -1) {
        colQCajas = c;
      } else if (cellVal.indexOf("POR CAJA") >= 0 || cellVal.indexOf("X CAJA") >= 0 || cellVal === "PACKING" || cellVal === "EMPAQUE" || cellVal === "UNID/CAJA" || (cellVal.indexOf("Q.") >= 0 && cellVal.indexOf("POR") >= 0)) {
        colQPorCaja = c;
      } else if (cellVal === "UN/MED" || cellVal === "UN/ MED" || cellVal === "UNIDAD" || cellVal === "U.M." || cellVal === "MEDIDA") {
        colUnMed = c;
      } else if (cellVal.indexOf("PRECIO") >= 0 && cellVal.indexOf("MAYOR") >= 0) {
        colPrecioMayor = c;
      } else if (cellVal.indexOf("PRECIO") >= 0 && cellVal.indexOf("CAJA") >= 0) {
        colPrecioCaja = c;
      } else if (cellVal.indexOf("PRECIO") >= 0 && cellVal.indexOf("ESPECIAL") >= 0) {
        colPrecioEspecial = c;
      } else if (cellVal.indexOf("TIPO DE CAMBIO") >= 0 || cellVal.indexOf("CAMBIO $") >= 0) {
        colTipoCambio = c;
      } else if (cellVal.indexOf("STOCK") >= 0 && cellVal.indexOf("RESERVA") >= 0 || cellVal === "STOC DE RESEF AS" || cellVal.indexOf("RESEF") >= 0) {
        colStockReservas = c;
      }
    }
    if (colCodigo !== -1 && colDetalle !== -1) {
      break;
    }
  }

  // Si no se detectó por nombre exacto, usar columnas estándar según la estructura vista:
  // Col C (index 2) = Detalle
  // Col D (index 3) = Codigo
  // Col E/F (index 4) = Q. De Cajas
  // Col G (index 6) = Q. Por Caja
  // Col H (index 7) = UN/MED
  // Col K (index 10) = Precio Mayor
  // Col L (index 11) = Precio Caja
  // Col M (index 12) = Precio Especial
  // Col N (index 13) = Tipo de Cambio
  // Col AE (index 30) = Stock Reservas
  if (colDetalle === -1) colDetalle = 2;
  if (colCodigo === -1) colCodigo = 3;
  if (colQCajas === -1) colQCajas = 4;
  if (colQPorCaja === -1) colQPorCaja = 6;
  if (colUnMed === -1) colUnMed = 7;
  if (colPrecioMayor === -1) colPrecioMayor = 10;
  if (colPrecioCaja === -1) colPrecioCaja = 11;
  if (colPrecioEspecial === -1) colPrecioEspecial = 12;
  if (colTipoCambio === -1) colTipoCambio = 13;
  if (colStockReservas === -1) {
    // Buscar la última columna con encabezado de stock o usar penúltima/última
    colStockReservas = Math.max(0, lastCol - 1);
  }

  // 3. Leer los datos de filas
  var startRow = headerRowIdx + 2; // fila 1-based donde empiezan los productos
  var numRows = lastRow - startRow + 1;
  if (numRows < 1) {
    return {
      success: true,
      sheet: actualSheetName,
      availableSheets: availableSheets,
      productos: []
    };
  }

  var fullData = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  var productos = [];
  var tipoCambioDetectado = 11.5;

  for (var i = 0; i < fullData.length; i++) {
    var rData = fullData[i];
    var codRaw = rData[colCodigo];
    if (!codRaw) continue;

    var codigo = String(codRaw).trim();
    if (!codigo || codigo.length < 2 || /^TOTAL/i.test(codigo)) continue;

    var detalle = rData[colDetalle] ? String(rData[colDetalle]).trim() : "";
    if (!detalle) continue;

    var qCajas = parseNumero(rData[colQCajas], 0);
    var qPorCaja = parseNumero(rData[colQPorCaja], 1);
    var unMed = rData[colUnMed] ? String(rData[colUnMed]).trim().toUpperCase() : "UNI";

    var precioMayor = parseNumero(rData[colPrecioMayor], 0);
    var precioCaja = parseNumero(rData[colPrecioCaja], 0);
    var precioEspecial = parseNumero(rData[colPrecioEspecial], 0);

    if (colTipoCambio !== -1 && rData[colTipoCambio]) {
      var tc = parseNumero(rData[colTipoCambio], 0);
      if (tc > 1) tipoCambioDetectado = tc;
    }

    var stockReservas = (colStockReservas !== -1 && rData[colStockReservas] !== "") ? parseNumero(rData[colStockReservas], 0) : qCajas;

    // Detectar marca
    var marca = "TOTAL";
    if (detalle.indexOf("TOTAL") >= 0 || codigo.indexOf("TH") === 0 || codigo.indexOf("TS") === 0 || codigo.indexOf("TP") === 0) {
      marca = "TOTAL";
    } else if (detalle.indexOf("DONG CHENG") >= 0 || detalle.indexOf("DONGCHENG") >= 0 || detalle.indexOf("DC") >= 0) {
      marca = "DONG CHENG";
    } else if (detalle.indexOf("UYUSTOOLS") >= 0 || detalle.indexOf("UYUS") >= 0) {
      marca = "UYUSTOOLS";
    } else if (detalle.indexOf("CROWN") >= 0) {
      marca = "CROWN";
    } else if (detalle.indexOf("AQUASTRONG") >= 0) {
      marca = "AQUASTRONG";
    }

    // Calcular precio de caja completa
    var precioRefUni = precioCaja > 0 ? precioCaja : (precioMayor > 0 ? precioMayor : 0);
    var subtotalCaja = Math.round((precioRefUni * qPorCaja) * 100) / 100;

    productos.push({
      id: "P_" + i,
      codigo: codigo,
      detalle: detalle,
      marca: marca,
      cajasVienen: qCajas,
      cantPorCaja: qPorCaja,
      unidad: unMed,
      precioMayor: precioMayor,
      precioCaja: precioCaja,
      precioEspecial: precioEspecial,
      precioRefUni: precioRefUni,
      precioCajaTotal: subtotalCaja,
      stockReserva: stockReservas, // Cajas que aún quedan disponibles para reservar
      agotado: (stockReservas <= 0)
    });
  }

  return {
    success: true,
    sheet: actualSheetName,
    availableSheets: availableSheets,
    tipoCambio: tipoCambioDetectado,
    totalProductos: productos.length,
    actualizadoEn: new Date().toISOString(),
    productos: productos
  };
}

function parseNumero(val, defecto) {
  if (val === undefined || val === null || val === "") return defecto;
  if (val instanceof Date) return defecto;
  if (typeof val === "number") return val;
  var str = String(val).trim();
  // Si parece una fecha (ej: 2026-01-01 o 1/1/2026) ignorarla
  if (/^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(str)) return defecto;
  str = str.replace(/[^0-9.,-]/g, '').trim();
  // Manejo de comas y puntos
  if (str.indexOf(',') >= 0 && str.indexOf('.') >= 0) {
    if (str.indexOf(',') > str.indexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.indexOf(',') >= 0) {
    str = str.replace(',', '.');
  }
  var n = parseFloat(str);
  return isNaN(n) ? defecto : n;
}
