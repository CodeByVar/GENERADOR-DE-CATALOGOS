# -*- coding: utf-8 -*-
import os

files_to_update = [
    "index.html",
    "catalogos.html",
    "catalogos_desktop.html",
    "catalogos_mobile.html"
]

target_text = '''        const rawCode = card.getAttribute('data-code') || '';
        const normCode = rawCode.toUpperCase().replace(/\\s+/g, '');
        const info = stockMap[normCode] || stockMap[rawCode.toUpperCase()];
        
        const pillEl = card.querySelector('.stock-status-pill') || document.getElementById(`stock_pill_${rawCode}`);
        const pkgEl = card.querySelector('.packaging-info') || document.getElementById(`pkg_info_${rawCode}`);

        if (info) {
          const cantCaja = info.c || info.cantPorCaja || 1;
          let unMed = info.u || info.unidadMedida || "";
          
          const inputEl = card.querySelector('.input-qty');
          const cardUnit = (inputEl && inputEl.getAttribute('data-unit')) || card.getAttribute('data-unit') || "UNI";
          if (!unMed || !isNaN(unMed) || /^\\d+$/.test(String(unMed).trim())) {
            unMed = cardUnit;
          }

          if (pkgEl) {
            if (cantCaja && cantCaja > 1) {
              pkgEl.innerHTML = `📦 ${cantCaja} ${unMed} / Caja`;
              pkgEl.setAttribute('title', `Viene ${cantCaja} ${unMed} por caja`);
            } else {
              pkgEl.innerHTML = `📦 1 ${unMed} / Caja`;
              pkgEl.setAttribute('title', `Viene 1 ${unMed} por caja`);
            }
          }
          
          const stock = typeof info.s === 'number' ? info.s : (typeof info.stockActual === 'number' ? info.stockActual : (typeof info.stock === 'number' ? info.stock : (info.stock === true ? 999 : 0)));
          const cajas = typeof info.b === 'number' ? info.b : (typeof info.cajas === 'number' ? info.cajas : Math.floor(stock / cantCaja));
          const estado = info.e || info.estado || (stock > 0 && info.stock !== false ? "DISPONIBLE" : "AGOTADO");'''

replacement_text = '''        const rawCode = card.getAttribute('data-code') || '';
        const normCode = rawCode.toUpperCase().replace(/\\s+/g, '');
        const cleanCode = normCode.replace(/[\\-._/]/g, '');
        const info = stockMap[normCode] || stockMap[rawCode.toUpperCase()] || stockMap[cleanCode] || stockMap[rawCode.trim()];
        
        const pillEl = card.querySelector('.stock-status-pill') || document.getElementById(`stock_pill_${rawCode}`);
        const pkgEl = card.querySelector('.packaging-info') || document.getElementById(`pkg_info_${rawCode}`);

        if (info) {
          let cantCaja = 1;
          const rawC = info.c !== undefined ? info.c : info.cantPorCaja;
          if (typeof rawC === 'number' && rawC > 0) {
            cantCaja = rawC;
          } else if (typeof rawC === 'string') {
            const parsedC = parseFloat(rawC.replace(/,/g, '').trim());
            if (!isNaN(parsedC) && parsedC > 0) cantCaja = parsedC;
          }

          let unMed = info.u || info.unidadMedida || "";
          const inputEl = card.querySelector('.input-qty');
          const cardUnit = (inputEl && inputEl.getAttribute('data-unit')) || card.getAttribute('data-unit') || "UNI";
          if (!unMed || !isNaN(unMed) || /^\\d+$/.test(String(unMed).trim())) {
            unMed = cardUnit;
          }

          if (pkgEl) {
            if (cantCaja && cantCaja > 1) {
              pkgEl.innerHTML = `📦 ${cantCaja} ${unMed} / Caja`;
              pkgEl.setAttribute('title', `Viene ${cantCaja} ${unMed} por caja`);
            } else {
              pkgEl.innerHTML = `📦 1 ${unMed} / Caja`;
              pkgEl.setAttribute('title', `Viene 1 ${unMed} por caja`);
            }
          }
          
          let stock = 0;
          const rawS = (info.s !== undefined && info.s !== null && info.s !== "") ? info.s :
                       ((info.stockActual !== undefined && info.stockActual !== null && info.stockActual !== "") ? info.stockActual : info.stock);
          if (typeof rawS === 'number') {
            stock = isNaN(rawS) ? 0 : rawS;
          } else if (typeof rawS === 'string') {
            const sParsed = parseFloat(rawS.replace(/,/g, '').trim());
            stock = isNaN(sParsed) ? 0 : sParsed;
          } else if (rawS === true) {
            stock = 999;
          } else if (rawS === false) {
            stock = 0;
          }

          let cajas = 0;
          const rawB = (info.b !== undefined && info.b !== null && info.b !== "") ? info.b : info.cajas;
          if (typeof rawB === 'number') {
            cajas = isNaN(rawB) ? Math.floor(stock / cantCaja) : rawB;
          } else if (typeof rawB === 'string') {
            const bParsed = parseFloat(rawB.replace(/,/g, '').trim());
            cajas = isNaN(bParsed) ? Math.floor(stock / cantCaja) : bParsed;
          } else {
            cajas = Math.floor(stock / cantCaja);
          }

          const estado = info.e || info.estado || (stock > 0 && info.stock !== false ? "DISPONIBLE" : "AGOTADO");'''

for fname in files_to_update:
    if os.path.exists(fname):
        print(f"Actualizando {fname}...")
        with open(fname, "r", encoding="utf-8") as f:
            content = f.read()
        if target_text in content:
            content = content.replace(target_text, replacement_text)
            with open(fname, "w", encoding="utf-8") as f:
                f.write(content)
            print(f"  [OK] {fname} actualizado con éxito.")
        else:
            print(f"  [AVISO] No se encontró target_text en {fname}.")

if __name__ == "__main__":
    pass
