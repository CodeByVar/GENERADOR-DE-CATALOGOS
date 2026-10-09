# -*- coding: utf-8 -*-
"""
Importadora Rivero - Servidor del Generador de Catálogos Web
===========================================================
Reemplaza la GUI clásica de Tkinter por una aplicación web interactiva local
ejecutándose en tu navegador de forma ultrarrápida y sin dependencias externas.
"""

import http.server
import socketserver
import webbrowser
import os
import sys
import re
import urllib.parse
import urllib.request
import ssl
import json
import subprocess
import generar_catalogo
from datetime import date, datetime

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

PORT = 5000

class SSEStdoutWriter:
    def __init__(self, handler):
        self.handler = handler
    def write(self, text):
        if not text:
            return
        for line in text.splitlines(keepends=False):
            # Formatear el texto de la consola para EventSource (SSE)
            # quitamos caracteres de retorno de carro
            line_clean = line.replace('\r', '').strip()
            if line_clean:
                try:
                    data = f"data: {line_clean}\n\n"
                    self.handler.wfile.write(data.encode('utf-8'))
                    self.handler.wfile.flush()
                except Exception:
                    pass
    def flush(self):
        try:
            self.handler.wfile.flush()
        except Exception:
            pass

class RedirectStdout:
    def __init__(self, new_stdout):
        self.new_stdout = new_stdout
        self.old_stdout = None
        self.old_stderr = None
        
    def __enter__(self):
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        sys.stdout = self.new_stdout
        sys.stderr = self.new_stdout
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr

def normalizar_marca_nombre(b):
    if not b: return ''
    b_u = str(b).upper().strip()
    if 'UYUS' in b_u or 'UYU' in b_u: return 'UYUSTOOLS'
    if 'TOTAL' in b_u: return 'TOTAL'
    if 'DONG' in b_u or 'DONGCHENG' in b_u: return 'DONG CHENG'
    if 'CROWN' in b_u: return 'CROWN'
    if 'AQUA' in b_u: return 'AQUASTRONG'
    if 'WADFOW' in b_u: return 'WADFOW'
    if 'LUTIAN' in b_u: return 'LUTIAN'
    if 'MAKAWA' in b_u: return 'MAKAWA'
    if 'MASTER' in b_u: return 'MASTERMAQ'
    if 'POWER' in b_u: return 'POWERMAQ'
    if 'TOYAKI' in b_u: return 'TOYAKI'
    if 'FERTON' in b_u: return 'FERTON'
    if 'FERR' in b_u: return 'FERRAWYY'
    if 'KAILI' in b_u: return 'KAILI'
    if 'KAMASA' in b_u: return 'KAMASA'
    if 'ASAKI' in b_u: return 'ASAKI'
    if 'DWT' in b_u: return 'DWT'
    if 'NEVA' in b_u: return 'NEVA'
    if 'OMEGA' in b_u: return 'OMEGA'
    if 'RIO' in b_u: return 'RIO'
    if 'PEGASUS' in b_u: return 'PEGASUS'
    return b

def cargar_db_marcas_excel(excel_path="catalogos.xlsx"):
    from generar_catalogo import load_workbook
    code_to_brand = {}
    if not os.path.exists(excel_path):
        return code_to_brand
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                row_str = [str(c).strip() if c is not None else '' for c in row]
                if sheet == 'FORMATO INVENTARIO' and len(row_str) > 4:
                    c = row_str[2].upper().strip()
                    b = row_str[4].upper().strip()
                    if c and b: code_to_brand[c] = b
                elif 'INVENTARIO' in sheet.upper() and len(row_str) > 3:
                    c = row_str[1].upper().strip()
                    b = row_str[2].upper().strip() if len(row_str) > 2 else ''
                    if c and b and len(c) > 2: code_to_brand[c] = b
        wb.close()
    except Exception as e:
        print(f"[AVISO] Error cargando marcas de Excel: {e}")
    return code_to_brand

def resolver_marca_prelista(c_raw, d_raw, code_to_brand=None):
    c = str(c_raw or '').upper().strip()
    d = str(d_raw or '').upper().strip()
    if code_to_brand and c in code_to_brand:
        nb = normalizar_marca_nombre(code_to_brand[c])
        if nb: return nb
    if any(k in d for k in ['UYUSTOOLS', 'UYUS', ' UYU', '-UYU', 'UYU ']) or d.endswith('UYU') or 'UYU' in c:
        return 'UYUSTOOLS'
    if 'MAKAWA' in d or c.startswith('MK-') or c.startswith('MK'):
        return 'MAKAWA'
    if 'LUTIAN' in d or c.startswith('LT') or ('HIDROLAVADORA' in d and 'LUTIAN' in d):
        return 'LUTIAN'
    if 'MASTERMAQ' in d or 'MASTER' in d or c.startswith('MAX-'):
        return 'MASTERMAQ'
    if 'POWERMAQ' in d:
        return 'POWERMAQ'
    if 'TOYAKI' in d or c.startswith('TK-'):
        return 'TOYAKI'
    if 'DONG CHENG' in d or 'DONGCHENG' in d or 'DCA' in d or c.startswith('DC'):
        return 'DONG CHENG'
    if 'CROWN' in d or (c.startswith('CT') and len(c) >= 5 and c[2].isdigit()):
        return 'CROWN'
    if 'AQUASTRONG' in d or 'AQUAS' in d:
        return 'AQUASTRONG'
    if 'WADFOW' in d or c.startswith('WDF') or c.startswith('WSS') or (c.startswith('W') and len(c) >= 4 and any(c.startswith(p) for p in ['WTB','WAG','WWH','WDT','WPL','WCS','WFS','WCD','WES'])):
        return 'WADFOW'
    if 'FERTON' in d or c.startswith('FT'):
        return 'FERTON'
    if 'FERRAWYY' in d or 'FERRA' in d:
        return 'FERRAWYY'
    if 'KAILI' in d or c.startswith('KL'):
        return 'KAILI'
    if 'KAMASA' in d or c.startswith('KM'):
        return 'KAMASA'
    if 'ASAKI' in d or c.startswith('AK'):
        return 'ASAKI'
    if 'DWT' in d:
        return 'DWT'
    if 'NEVA' in d:
        return 'NEVA'
    if 'OMEGA' in d:
        return 'OMEGA'
    if 'RIO' in d:
        return 'RIO'
    if 'PEGASUS' in d:
        return 'PEGASUS'
    if 'TOTAL' in d or 'TOTA' in d or any(c.startswith(p) for p in ['TH', 'TS', 'TP', 'TG', 'TAC', 'TL', 'THT', 'TV', 'TOS', 'TIDLI', 'TIWLI', 'TMT', 'TBC', 'TW', 'TB', 'TCKLI', 'PMST', 'PMTS']):
        return 'TOTAL'
    return 'VARIOS'

def sincronizar_imagenes_para_prelista(productos, excel_path="catalogos.xlsx"):
    """
    Sincroniza y extrae las fotos de cada producto de la prelista directamente
    desde 'catalogos.xlsx' y la caché local en 'temp_imgs/'.
    Convierte cada foto a un Data URI en Base64 (data:image/webp;base64,...)
    para que se sirva al 100% en Vercel sin depender de rutas locales ni dar 404.
    Además verifica y asigna la marca exacta de cada producto basándose en 'catalogos.xlsx'.
    """
    import io, base64
    from PIL import Image as PILImage
    from generar_catalogo import (
        load_workbook, detectar_hojas_inventario, detectar_columnas,
        extraer_bytes_de_imagen, autocrop_image, normalizar_codigo
    )

    if not os.path.exists("temp_imgs"):
        os.makedirs("temp_imgs")

    # Resolver marcas reales de cada producto
    code_to_brand = cargar_db_marcas_excel(excel_path)
    for p in productos:
        m_curr = p.get("marca")
        m_res = resolver_marca_prelista(p.get("codigo"), p.get("detalle"), code_to_brand)
        if m_res and m_res != "VARIOS":
            p["marca"] = m_res
        elif not m_curr or m_curr == "TOTAL":
            p["marca"] = m_res

    # 1. Identificar productos que ya tienen imagen en disco vs los que necesitan extraerse de Excel
    codigos_pendientes = {} # { norm_code: [producto_dicts] }
    
    for p in productos:
        cod = str(p.get("codigo", "")).strip()
        if not cod:
            continue
        clean_cod = re.sub(r'[\\/*?:"<>| ]', "_", cod)
        norm_cod = normalizar_codigo(cod).upper()
        
        # Si ya tiene una imagen Base64 válida, mantenerla
        if p.get("imagen") and str(p["imagen"]).startswith("data:image/"):
            continue

        # Verificar si ya existe en disco
        img_disk = os.path.join("temp_imgs", f"prod_{clean_cod}.webp")
        if not os.path.exists(img_disk):
            img_disk_png = os.path.join("temp_imgs", f"prod_{clean_cod}.png")
            if os.path.exists(img_disk_png):
                img_disk = img_disk_png
                
        if os.path.exists(img_disk) and os.path.getsize(img_disk) > 400:
            try:
                with open(img_disk, "rb") as f_i:
                    raw_b = f_i.read()
                mime = "image/webp" if img_disk.endswith(".webp") else "image/png"
                p["imagen"] = f"data:{mime};base64,{base64.b64encode(raw_b).decode('utf-8')}"
                continue
            except Exception:
                pass
                
        # Si no tiene imagen en disco, registrarlo para buscar en catalogos.xlsx
        if norm_cod not in codigos_pendientes:
            codigos_pendientes[norm_cod] = []
        codigos_pendientes[norm_cod].append(p)

    # 2. Si hay códigos pendientes y existe catalogos.xlsx, extraer directo de Excel
    if codigos_pendientes and os.path.exists(excel_path):
        print(f"\n[EXCEL PRELISTA] Buscando fotos para {len(codigos_pendientes)} productos pendientes en '{excel_path}'...")
        try:
            wb = load_workbook(excel_path, data_only=True)
            hojas = detectar_hojas_inventario(wb)
            
            # Buscar en cada hoja de inventario
            for ws in hojas:
                if not hasattr(ws, '_images') or not ws._images:
                    continue
                cols, start_row = detectar_columnas(ws)
                col_c = cols.get("codigo", 2)
                
                # Mapear filas que corresponden a nuestros códigos en esta hoja
                filas_target = {} # { row_number: norm_code }
                max_r = min(ws.max_row + 10, 20000) if ws.max_row else 5000
                consecutive_empty = 0
                
                for r in range(start_row, max_r):
                    val = ws.cell(row=r, column=col_c).value
                    if val and str(val).strip():
                        norm_v = normalizar_codigo(val).upper()
                        if norm_v in codigos_pendientes:
                            filas_target[r] = norm_v
                        consecutive_empty = 0
                    else:
                        consecutive_empty += 1
                        if consecutive_empty >= 80:
                            break
                            
                if not filas_target:
                    continue
                    
                print(f"  [EXCEL PRELISTA] Hoja '{ws.title}': detectados {len(filas_target)} productos coincidentes. Extrayendo fotos...")
                
                # Extraer imágenes de las filas encontradas
                for img in ws._images:
                    try:
                        ancla = getattr(img, 'anchor', None)
                        fila = None
                        if ancla is not None:
                            if hasattr(ancla, '_from') and hasattr(ancla._from, 'row'):
                                fila = ancla._from.row + 1
                            elif hasattr(ancla, 'from_row'):
                                fila = ancla.from_row + 1
                            elif hasattr(ancla, 'row'):
                                fila = ancla.row + 1
                            elif isinstance(ancla, str):
                                m = re.search(r'\d+', ancla)
                                if m: fila = int(m.group())
                                
                        if fila is None:
                            continue
                            
                        # Verificar si coincide con una fila de interés (+- 2 de tolerancia)
                        norm_match = None
                        for off in [0, -1, 1, -2, 2]:
                            if (fila + off) in filas_target:
                                norm_match = filas_target[fila + off]
                                break
                                
                        if not norm_match or norm_match not in codigos_pendientes:
                            continue
                            
                        raw_bytes = extraer_bytes_de_imagen(img)
                        if raw_bytes and len(raw_bytes) > 200:
                            img_pil = PILImage.open(io.BytesIO(raw_bytes))
                            img_cropped = autocrop_image(img_pil)
                            if img_cropped.width > 350 or img_cropped.height > 350:
                                resample_filter = getattr(PILImage, "Resampling", None)
                                filter_type = resample_filter.LANCZOS if resample_filter else getattr(PILImage, "ANTIALIAS", 3)
                                img_cropped.thumbnail((350, 350), filter_type)
                                
                            clean_cod = re.sub(r'[\\/*?:"<>| ]', "_", norm_match)
                            dest_webp = os.path.join("temp_imgs", f"prod_{clean_cod}.webp")
                            img_cropped.save(dest_webp, "WEBP", quality=65)
                            
                            with open(dest_webp, "rb") as f_w:
                                webp_raw = f_w.read()
                            b64_str = f"data:image/webp;base64,{base64.b64encode(webp_raw).decode('utf-8')}"
                            
                            # Asignar a todos los productos con este código
                            for p_item in codigos_pendientes[norm_match]:
                                p_item["imagen"] = b64_str
                                
                            del codigos_pendientes[norm_match]
                    except Exception:
                        pass
                        
            wb.close()
            print(f"[EXCEL PRELISTA] Extracción finalizada. Pendientes sin foto: {len(codigos_pendientes)}.")
        except Exception as e_wb:
            print(f"[EXCEL PRELISTA ERROR] No se pudo leer {excel_path}: {e_wb}")

    return productos

def sincronizar_imagenes_prelista_existente():
    """
    Revisa prelista_data.json en el arranque. Si hay productos sin imagen, extrae de catalogos.xlsx
    y actualiza tanto prelista_data.json como prelista.html con las fotos en Base64.
    """
    if not os.path.exists("prelista_data.json"):
        return
    try:
        with open("prelista_data.json", "r", encoding="utf-8") as f_in:
            data = json.load(f_in)
        productos = data.get("productos", [])
        if not productos:
            return
        
        # Verificar si falta imagen o si las marcas necesitan sincronización
        sin_foto = sum(1 for p in productos if not p.get("imagen"))
        marcas_total_excesivas = sum(1 for p in productos if p.get("marca") == "TOTAL" and "TOTAL" not in str(p.get("detalle", "")).upper() and not str(p.get("codigo", "")).upper().startswith("T"))
        if sin_foto > 0 or marcas_total_excesivas > 0:
            print(f"[INICIO PRELISTA] Sincronizando fotos y marcas de prelista con Excel...")
            productos = sincronizar_imagenes_para_prelista(productos)
            data["productos"] = productos
            data["actualizadoEn"] = datetime.now().isoformat()
            with open("prelista_data.json", "w", encoding="utf-8") as f_out:
                json.dump(data, f_out, indent=2, ensure_ascii=False)
            
            # Hornear en prelista.html
            if os.path.exists("prelista.html"):
                with open("prelista.html", "r", encoding="utf-8") as f_h:
                    html_content = f_h.read()
                baked_json = json.dumps(productos, ensure_ascii=False, indent=2)
                new_baked_block = f"/* BAKED_PRELISTA_DATA_START */\n    const BAKED_PRELISTA_DATA = {baked_json};\n    /* BAKED_PRELISTA_DATA_END */"
                if "/* BAKED_PRELISTA_DATA_START */" in html_content:
                    html_content = re.sub(
                        r'/\* BAKED_PRELISTA_DATA_START \*/.*?/\* BAKED_PRELISTA_DATA_END \*/',
                        lambda m: new_baked_block,
                        html_content,
                        flags=re.DOTALL
                    )
                with open("prelista.html", "w", encoding="utf-8") as f_hw:
                    f_hw.write(html_content)

            # También sincronizar api/prelista.js para Vercel
            if os.path.exists("api/prelista.js"):
                try:
                    with open("api/prelista.js", "r", encoding="utf-8") as f_api:
                        api_content = f_api.read()
                    demo_payload = {
                        "success": True,
                        "sheet": data.get("sheet", "Prelista"),
                        "availableSheets": data.get("availableSheets", ["Prelista"]),
                        "totalProductos": len(productos),
                        "actualizadoEn": datetime.now().isoformat(),
                        "productos": productos
                    }
                    demo_json_str = json.dumps(demo_payload, ensure_ascii=False, indent=2)
                    new_demo_block = f"// DEMO_PRELISTA_START\nconst DEMO_PRELISTA = {demo_json_str};\n// DEMO_PRELISTA_END"
                    if "// DEMO_PRELISTA_START" in api_content:
                        api_content = re.sub(
                            r'// DEMO_PRELISTA_START.*?// DEMO_PRELISTA_END',
                            lambda m: new_demo_block,
                            api_content,
                            flags=re.DOTALL
                        )
                        with open("api/prelista.js", "w", encoding="utf-8") as f_api_w:
                            f_api_w.write(api_content)
                except Exception as e_api_err:
                    print(f"[INICIO PRELISTA AVISO] No se pudo actualizar api/prelista.js: {e_api_err}")

            print(f"[INICIO PRELISTA OK] ¡Fotos de prelista horneadas exitosamente!")
    except Exception as e_sync:
        print(f"[INICIO PRELISTA AVISO] No se pudo sincronizar fotos iniciales: {e_sync}")

class CatalogWebHandler(http.server.BaseHTTPRequestHandler):
    
    # Registrar conexiones de red entrantes para ver quién se conecta al panel
    def log_message(self, format, *args):
        log_line = format % args
        # Filtrar peticiones secundarias para no inundar la consola y mostrar solo accesos importantes
        if any(x in log_line for x in ["GET / ", "GET /index.html", "GET /generar", "POST /"]):
            import datetime
            hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{hora}] [CONEXIÓN] Cliente {self.client_address[0]} accedió a: {log_line.strip()}")

    def serve_file(self, file_path, content_type):
        if not os.path.exists(file_path):
            self.send_error(404, "File not found")
            return
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        size = os.path.getsize(file_path)
        self.send_header('Content-Length', str(size))
        self.end_headers()
        with open(file_path, 'rb') as f:
            self.wfile.write(f.read())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.end_headers()

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        if parsed_url.path == "/api/prelista/guardar":
            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode('utf-8'))
                excluidos = data.get("excluidos", [])
                sheet = data.get("sheet", "Hoja 49")
                productos = data.get("productos", [])

                # Sincronizar y extraer imágenes directamente desde catalogos.xlsx y temp_imgs
                productos = sincronizar_imagenes_para_prelista(productos)

                # 1. Guardar prelista_data.json con todos los productos de prelista
                with open("prelista_data.json", "w", encoding="utf-8") as f_data:
                    json.dump({
                        "success": True,
                        "sheet": sheet,
                        "totalProductos": len(productos),
                        "actualizadoEn": datetime.now().isoformat(),
                        "productos": productos
                    }, f_data, indent=2, ensure_ascii=False)

                # 2. Guardar prelista_excluidos.json por compatibilidad
                with open("prelista_excluidos.json", "w", encoding="utf-8") as f_ex:
                    json.dump({"sheet": sheet, "excluidos": excluidos}, f_ex, indent=2, ensure_ascii=False)

                # 3. Grabar directamente los productos en prelista.html (Bake estático instantáneo)
                if os.path.exists("prelista.html"):
                    with open("prelista.html", "r", encoding="utf-8") as f_html:
                        html_content = f_html.read()

                    baked_json = json.dumps(productos, ensure_ascii=False, indent=2)
                    new_baked_block = f"/* BAKED_PRELISTA_DATA_START */\n    const BAKED_PRELISTA_DATA = {baked_json};\n    /* BAKED_PRELISTA_DATA_END */"
                    if "/* BAKED_PRELISTA_DATA_START */" in html_content:
                        html_content = re.sub(
                            r'/\* BAKED_PRELISTA_DATA_START \*/.*?/\* BAKED_PRELISTA_DATA_END \*/',
                            lambda m: new_baked_block,
                            html_content,
                            flags=re.DOTALL
                        )
                    with open("prelista.html", "w", encoding="utf-8") as f_html_w:
                        f_html_w.write(html_content)

                # 4. Actualizar DEMO_PRELISTA en api/prelista.js para Vercel Edge Runtime
                if os.path.exists("api/prelista.js"):
                    with open("api/prelista.js", "r", encoding="utf-8") as f_api:
                        api_content = f_api.read()

                    sheet_list = [s.strip() for s in re.split(r'[,+]', str(sheet)) if s.strip()]
                    demo_payload = {
                        "success": True,
                        "sheet": sheet,
                        "availableSheets": sheet_list if sheet_list else [sheet],
                        "totalProductos": len(productos),
                        "actualizadoEn": datetime.now().isoformat(),
                        "productos": productos
                    }
                    demo_json_str = json.dumps(demo_payload, ensure_ascii=False, indent=2)
                    new_demo_block = f"// DEMO_PRELISTA_START\nconst DEMO_PRELISTA = {demo_json_str};\n// DEMO_PRELISTA_END"
                    if "// DEMO_PRELISTA_START" in api_content:
                        api_content = re.sub(
                            r'// DEMO_PRELISTA_START.*?// DEMO_PRELISTA_END',
                            lambda m: new_demo_block,
                            api_content,
                            flags=re.DOTALL
                        )
                        with open("api/prelista.js", "w", encoding="utf-8") as f_api_w:
                            f_api_w.write(api_content)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "success": True, 
                    "total": len(productos),
                    "sheet": sheet,
                    "mensaje": f"Se grabaron {len(productos)} productos en prelista.html con éxito."
                }).encode('utf-8'))
            except Exception as e_save:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(e_save)}).encode('utf-8'))
            return

        self.send_error(404, "Not found")

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        
        # 1. Endpoint SSE para generación en tiempo real con streaming de consola
        if parsed_url.path == "/generar":
            self.send_response(200)
            self.send_header('Content-Type', 'text/event-stream')
            self.send_header('Cache-Control', 'no-cache')
            self.send_header('Connection', 'keep-alive')
            self.end_headers()
            
            # Recuperar parámetros
            codes_raw = query_params.get('codes', [''])[0]
            sync_raw = query_params.get('sync', ['false'])[0]
            layout_raw = query_params.get('layout', ['desktop'])[0]
            force_images_raw = query_params.get('force_images', ['false'])[0]
            whatsapp_raw = query_params.get('whatsapp', [''])[0]
            filter_stock_raw = query_params.get('filter_stock', ['false'])[0]
            
            descargar_nube = (sync_raw.lower() == 'true')
            forzar_imagenes = (force_images_raw.lower() == 'true')
            filtrar_agotados = (filter_stock_raw.lower() == 'true')
            if codes_raw and codes_raw.strip():
                tokens = re.split(r'[\r\n,;\t]+', codes_raw)
                codigos_custom = [t.strip().strip('"\'') for t in tokens if t.strip().strip('"\'')]
            else:
                codigos_custom = None
            
            writer = SSEStdoutWriter(self)
            
            writer.write(">>> Iniciando generación de catálogo desde el servidor web...\n")
            if codigos_custom:
                writer.write(f">>> Códigos recibidos: {len(codigos_custom)} ítems para el catálogo.\n")
            else:
                writer.write(">>> Leyendo códigos desde el Excel local (hoja Vista_Catalogo)...\n")
            if descargar_nube:
                writer.write(">>> [NUBE] Sincronización con Google Drive activada.\n")
            else:
                writer.write(">>> [LOCAL DIRECTO] Usando 'catalogos.xlsx' local (rápido, sin límites de peso de Drive).\n")
            if filtrar_agotados:
                writer.write(">>> [FILTRO] Activado: Se omitirán productos agotados según el stock en tiempo real.\n")
                
            with RedirectStdout(writer):
                try:
                    import importlib
                    importlib.reload(generar_catalogo)
                    generar_catalogo.generar(descargar_nube=descargar_nube, codigos_custom=codigos_custom, layout=layout_raw, forzar_imagenes=forzar_imagenes, whatsapp_phone=whatsapp_raw, filtrar_agotados=filtrar_agotados)
                    # Enviar señal de éxito final
                    writer.write("EVENT_SUCCESS: Proceso finalizado con éxito.\n")
                except BaseException as e:
                    import traceback
                    traceback.print_exc(file=writer)
                    writer.write("EVENT_ERROR: Ocurrió un error al procesar el catálogo.\n")
            return

        # 1.5 Endpoint SSE para Publicar en Vercel vía Git Push
        elif parsed_url.path == "/publicar_vercel":
            self.send_response(200)
            self.send_header('Content-Type', 'text/event-stream')
            self.send_header('Cache-Control', 'no-cache')
            self.send_header('Connection', 'keep-alive')
            self.end_headers()
            
            writer = SSEStdoutWriter(self)
            writer.write(">>> [VERCEL] Preparando despliegue de catálogo online...\n")
            try:
                import shutil
                if os.path.exists("catalogos.html"):
                    shutil.copyfile("catalogos.html", "index.html")
                    writer.write(">>> [VERCEL] Sincronizado catalogos.html con index.html.\n")
                
                import glob
                add_files = ["index.html", "catalogos.html", "catalogos_desktop.html", "catalogos_mobile.html", "prelista.html", "prelista_data.json", "prelista_excluidos.json", "vercel.json", "generar_catalogo.py", "web_generator.py", "Publicar_en_Vercel.bat", "api/stock.js", "api/prelista.js"]
                for img_pat in ["*.png", "*.jpg", "*.jpeg", "*.webp"]:
                    add_files.extend(glob.glob(img_pat))
                subprocess.run(["git", "add"] + add_files, capture_output=True)
                subprocess.run(["git", "add", "-u"], capture_output=True)
                
                writer.write(">>> [VERCEL] Creando punto de actualización en historial...\n")
                subprocess.run(["git", "commit", "-m", "Actualizacion del catalogo online para clientes"], capture_output=True)
                
                writer.write(">>> [VERCEL] Subiendo cambios a GitHub / Vercel en la nube...\n")
                try:
                    res = subprocess.run(["git", "push", "origin", "main"], capture_output=True, text=True, timeout=60)
                    if res.returncode == 0:
                        writer.write(">>> [VERCEL] [OK] Subida completada con éxito!\n")
                        writer.write(">>> [VERCEL] Vercel se está actualizando en vivo en tu enlace web.\n")
                        writer.write("EVENT_SUCCESS: Catálogo publicado con éxito en Vercel.\n")
                    else:
                        err_msg = res.stderr or res.stdout
                        writer.write(f">>> [VERCEL AVISO] {err_msg.strip()}\n")
                        writer.write(">>> [CONSEJO] Si requiere inicio de sesión en GitHub, ejecuta 'Publicar_en_Vercel.bat' en la carpeta.\n")
                        writer.write("EVENT_ERROR: Error al subir cambios a GitHub / Vercel.\n")
                except subprocess.TimeoutExpired:
                    writer.write(">>> [VERCEL AVISO] Git tardó demasiado (puede requerir inicio de sesión en GitHub).\n")
                    writer.write(">>> [SOLUCIÓN] Haz doble clic en 'Publicar_en_Vercel.bat' para iniciar sesión en GitHub con ventana visible.\n")
                    writer.write("EVENT_ERROR: Tiempo de espera agotado al conectar con GitHub.\n")
            except Exception as ex:
                writer.write(f">>> [VERCEL ERROR] {ex}\n")
                writer.write("EVENT_ERROR: Ocurrió una excepción al publicar.\n")
            return

        # 1.6 Endpoint API para obtener el resumen de inventario (Buscador, Marcas y Plantillas)
        elif parsed_url.path == "/api/productos":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            try:
                import importlib
                importlib.reload(generar_catalogo)
                data = generar_catalogo.obtener_resumen_inventario()
                self.wfile.write(json.dumps(data).encode('utf-8'))
            except Exception as e:
                self.wfile.write(json.dumps({"error": str(e), "productos": [], "marcas": {}, "categorias": {}}).encode('utf-8'))
            return

        # 1.7 Endpoint API para sincronizar y probar stock en vivo desde Google Apps Script
        elif parsed_url.path == "/api/stock":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate, max-age=0')
            self.send_header('Pragma', 'no-cache')
            self.send_header('Expires', '0')
            self.end_headers()
            try:
                # Comprobar si existe micro-caché reciente en memoria del servidor local (< 6s)
                import time
                now = time.time()
                cache_data = getattr(self.server, '_stock_cache_data', None)
                cache_time = getattr(self.server, '_stock_cache_time', 0)
                is_fresh = "fresh" in query_params or "force" in query_params

                if not is_fresh and cache_data and (now - cache_time < 6):
                    self.wfile.write(cache_data)
                    return

                # 1. Intentar primero con Supabase (Paginación paralela para traer TODOS los 5,300+ productos en < 1s)
                supabase_base_url = "https://mjiezwmldydnlcpshlpq.supabase.co/rest/v1/catalogo_stock?select=codigo,stock_actual,cantidad_caja,unidad_medida,cajas_disponibles,estado&limit=1000&offset="
                supabase_key = "sb_publishable_5Nxl1zMTRm6ngigdYQUA-g_lOKdUpzo"
                merged = {}

                try:
                    import concurrent.futures
                    def fetch_sb_page(offset):
                        try:
                            req_sb = urllib.request.Request(f"{supabase_base_url}{offset}", headers={
                                'apikey': supabase_key,
                                'Authorization': f'Bearer {supabase_key}',
                                'Content-Type': 'application/json'
                            })
                            ctx_sb = ssl.create_default_context()
                            ctx_sb.check_hostname = False
                            ctx_sb.verify_mode = ssl.CERT_NONE
                            with urllib.request.urlopen(req_sb, context=ctx_sb, timeout=6) as resp_sb:
                                data_p = json.loads(resp_sb.read().decode('utf-8'))
                                return data_p if isinstance(data_p, list) else []
                        except Exception:
                            return []

                    offsets = [0, 1000, 2000, 3000, 4000, 5000, 6000]
                    all_sb_rows = []
                    with concurrent.futures.ThreadPoolExecutor(max_workers=7) as ex_sb:
                        for p_rows in ex_sb.map(fetch_sb_page, offsets):
                            all_sb_rows.extend(p_rows)

                    if len(all_sb_rows) > 0:
                        for row in all_sb_rows:
                            raw_c = row.get("codigo")
                            if not raw_c:
                                continue
                            raw_k = str(raw_c).strip()
                            upper_k = raw_k.upper()
                            norm_k = upper_k.replace(" ", "")
                            simple_k = re.sub(r'[\-._/]', '', norm_k)

                            def _safe_float(v, d=0.0):
                                if v is None or v == "": return d
                                try: return float(str(v).replace(",", "").strip())
                                except Exception: return d
                            def _safe_int(v, d=0):
                                if v is None or v == "": return d
                                try: return int(float(str(v).replace(",", "").strip()))
                                except Exception: return d

                            item_info = {
                                "s": _safe_float(row.get("stock_actual"), 0.0),
                                "c": _safe_float(row.get("cantidad_caja"), 1.0) or 1.0,
                                "u": str(row.get("unidad_medida") or "UNI"),
                                "b": _safe_int(row.get("cajas_disponibles"), 0),
                                "e": str(row.get("estado") or "DISPONIBLE")
                            }

                            merged[raw_k] = item_info
                            merged[upper_k] = item_info
                            merged[norm_k] = item_info
                            if simple_k != norm_k:
                                merged[simple_k] = item_info
                        print(f">>> [SUPABASE OK] ¡ÉXITO! {len(all_sb_rows)} productos leídos instantáneamente desde Supabase.")
                except Exception as err_sb:
                    print(f">>> [SUPABASE AVISO] No se pudo leer Supabase ({err_sb}). Intentando con Google Sheets...")

                # 2. Respaldo secundario: Si Supabase no devolvió productos, consultar Google Sheets
                if not merged:
                    stock_urls = [
                        "https://script.google.com/macros/s/AKfycbz3pjscUdPvuSLWgTA1KugkoffYyWw9zJRqrg22eJCK-by3aTHLF2oZ7t0S3SwmOnwS/exec",
                        "https://script.google.com/macros/s/AKfycbw5rOmXaEKusH_PYZAG2r0OpybEqqfGlrZsQRQdeiJtJXbCsJsW-oxjQK8q690s8No/exec"
                    ]
                    import concurrent.futures

                    def fetch_url(u):
                        try:
                            r = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
                            c = ssl.create_default_context()
                            c.check_hostname = False
                            c.verify_mode = ssl.CERT_NONE
                            with urllib.request.urlopen(r, context=c, timeout=18) as resp:
                                data_parsed = json.loads(resp.read().decode('utf-8'))
                                if isinstance(data_parsed, dict) and not data_parsed.get("error"):
                                    print(f">>> [STOCK API OK] {len(data_parsed)} productos obtenidos de Google Drive ({u[:50]}...)")
                                    return data_parsed
                                else:
                                    print(f">>> [STOCK API AVISO] Respuesta vacía o con error: {data_parsed}")
                                    return {}
                        except Exception as err:
                            print(f">>> [STOCK API ERROR] No se pudo leer {u[:50]}... -> {err}")
                            return {}

                    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                        futures = [executor.submit(fetch_url, u) for u in stock_urls]
                        for fut in concurrent.futures.as_completed(futures):
                            res = fut.result()
                            if isinstance(res, dict):
                                for k, v in res.items():
                                    if not k or not v:
                                        continue
                                    raw_k = str(k).strip()
                                    upper_k = raw_k.upper()
                                    norm_k = upper_k.replace(" ", "")
                                    simple_k = re.sub(r'[\-._/]', '', norm_k)
                                    merged[raw_k] = v
                                    merged[upper_k] = v
                                    merged[norm_k] = v
                                    if simple_k != norm_k:
                                        merged[simple_k] = v
                                    continue
                                raw_k = str(k).strip()
                                upper_k = raw_k.upper()
                                norm_k = upper_k.replace(" ", "")
                                simple_k = re.sub(r'[\-._/]', '', norm_k)
                                merged[raw_k] = v
                                merged[upper_k] = v
                                merged[norm_k] = v
                                if simple_k != norm_k:
                                    merged[simple_k] = v

                content = json.dumps(merged).encode('utf-8')
                if len(merged) > 0:
                    self.server._stock_cache_data = content
                    self.server._stock_cache_time = now
                    print(f">>> [STOCK TOTAL] Total sincronizado y unificado en catálogo: {len(merged)} claves de productos.")
                    self.wfile.write(content)
                else:
                    # Si Google tardó o ambas fallaron, usar caché previa en memoria si existe
                    fallback = getattr(self.server, '_stock_cache_data', None)
                    if fallback:
                        print(">>> [STOCK] Sirviendo última copia previa de stock guardada en memoria.")
                        self.wfile.write(fallback)
                    else:
                        self.wfile.write(content)
            except Exception as e:
                fallback = getattr(self.server, '_stock_cache_data', None)
                if fallback:
                    self.wfile.write(fallback)
                else:
                    self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))
            return

        # 1.8 Endpoint API para obtener los códigos del último catálogo generado
        elif parsed_url.path == "/api/ultimos_codigos":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            codigos_res = []
            if os.path.exists("ultimos_codigos.json"):
                try:
                    with open("ultimos_codigos.json", "r", encoding="utf-8") as f_u:
                        data_u = json.load(f_u)
                        codigos_res = data_u.get("codigos", [])
                except Exception:
                    pass
            # Si no existe archivo de historial, extraer automáticamente de catalogos.html o index.html (excluyendo scripts de JS)
            if not codigos_res:
                ref_file = "catalogos.html" if os.path.exists("catalogos.html") else ("index.html" if os.path.exists("index.html") else None)
                if ref_file:
                    try:
                        with open(ref_file, "r", encoding="utf-8", errors="ignore") as f_ref:
                            content = f_ref.read()
                            content_clean = re.sub(r'<script.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
                            matches = re.findall(r'data-code="([^"]+)"', content_clean)
                            seen = set()
                            for m in matches:
                                mu = m.strip().upper()
                                if mu and mu not in seen and "{" not in mu and "$" not in mu:
                                    seen.add(mu)
                                    codigos_res.append(mu)
                    except Exception:
                        pass
            else:
                codigos_res = [c for c in codigos_res if "{" not in c and "$" not in c]
            self.wfile.write(json.dumps({"codigos": codigos_res, "total": len(codigos_res)}).encode('utf-8'))
            return

        # 1.9 Endpoint API para prelista (proxy a Google Apps Script)
        elif parsed_url.path == "/api/prelista":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            
            script_url = "https://script.google.com/macros/s/AKfycbwZluWpdw3riIr35gdrHhgdn7cRLcoNOuqQffxbnPIFqFbu2EgsxZffipAs0c4_gDpbKg/exec"
            query_str = parsed_url.query
            params = urllib.parse.parse_qs(query_str) if query_str else {}
            sheet_val = params.get('sheet', [''])[0]
            force_val = params.get('force', [''])[0]

            # Si se piden múltiples hojas separadas por coma o '+'
            if sheet_val and (',' in sheet_val or '+' in sheet_val):
                sheet_names = [s.strip() for s in re.split(r'[,+]', sheet_val) if s.strip()]
                merged_prods = []
                prods_map = {}
                available_sheets = []
                c = ssl.create_default_context()
                c.check_hostname = False
                c.verify_mode = ssl.CERT_NONE
                
                for s_name in sheet_names:
                    s_qs = f"sheet={urllib.parse.quote(s_name)}"
                    if force_val:
                        s_qs += "&force=1"
                    try:
                        r = urllib.request.Request(f"{script_url}?{s_qs}", headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(r, context=c, timeout=25) as resp:
                            s_data = json.loads(resp.read().decode('utf-8'))
                            if s_data.get("success") and isinstance(s_data.get("productos"), list):
                                if s_data.get("availableSheets") and len(s_data["availableSheets"]) > len(available_sheets):
                                    available_sheets = s_data["availableSheets"]
                                for prod in s_data["productos"]:
                                    cod = str(prod.get("codigo", "")).strip().upper()
                                    if not cod:
                                        continue
                                    if cod in prods_map:
                                        ex = prods_map[cod]
                                        ex["cajasVienen"] = int(ex.get("cajasVienen", 0)) + int(prod.get("cajasVienen", 0))
                                        ex["stockReserva"] = int(ex.get("stockReserva", 0)) + int(prod.get("stockReserva", 0))
                                        if "_origenHojas" not in ex:
                                            ex["_origenHojas"] = []
                                        if s_name not in ex["_origenHojas"]:
                                            ex["_origenHojas"].append(s_name)
                                    else:
                                        prod["_origenHojas"] = [s_name]
                                        prods_map[cod] = prod
                                        merged_prods.append(prod)
                    except Exception:
                        pass
                
                resp_payload = {
                    "success": True,
                    "sheet": " + ".join(sheet_names),
                    "availableSheets": available_sheets,
                    "totalProductos": len(merged_prods),
                    "productos": merged_prods
                }
                self.wfile.write(json.dumps(resp_payload, ensure_ascii=False).encode('utf-8'))
                return

            target_url = f"{script_url}?{query_str}" if query_str else script_url
            try:
                r = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
                c = ssl.create_default_context()
                c.check_hostname = False
                c.verify_mode = ssl.CERT_NONE
                with urllib.request.urlopen(r, context=c, timeout=25) as resp:
                    self.wfile.write(resp.read())
            except Exception as e_script:
                self.wfile.write(json.dumps({"success": False, "error": str(e_script)}).encode('utf-8'))
            return

        # 1.10 Endpoint API para sincronizar fotos de la prelista directamente desde Excel
        elif parsed_url.path == "/api/prelista/sincronizar_fotos":
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            try:
                with open("prelista_data.json", "r", encoding="utf-8") as f_d:
                    data = json.load(f_d)
                prods = data.get("productos", [])
                prods = sincronizar_imagenes_para_prelista(prods)
                data["productos"] = prods
                data["actualizadoEn"] = datetime.now().isoformat()
                with open("prelista_data.json", "w", encoding="utf-8") as f_d_w:
                    json.dump(data, f_d_w, indent=2, ensure_ascii=False)

                if os.path.exists("prelista.html"):
                    with open("prelista.html", "r", encoding="utf-8") as f_h:
                        html_c = f_h.read()
                    baked_json = json.dumps(prods, ensure_ascii=False, indent=2)
                    new_baked = f"/* BAKED_PRELISTA_DATA_START */\n    const BAKED_PRELISTA_DATA = {baked_json};\n    /* BAKED_PRELISTA_DATA_END */"
                    if "/* BAKED_PRELISTA_DATA_START */" in html_c:
                        html_c = re.sub(
                            r'/\* BAKED_PRELISTA_DATA_START \*/.*?/\* BAKED_PRELISTA_DATA_END \*/',
                            lambda m: new_baked,
                            html_c,
                            flags=re.DOTALL
                        )
                    with open("prelista.html", "w", encoding="utf-8") as f_hw:
                        f_hw.write(html_c)

                # También actualizar api/prelista.js para Vercel
                if os.path.exists("api/prelista.js"):
                    try:
                        with open("api/prelista.js", "r", encoding="utf-8") as f_api:
                            api_c = f_api.read()
                        demo_payload = {
                            "success": True,
                            "sheet": data.get("sheet", "Prelista"),
                            "availableSheets": data.get("availableSheets", ["Prelista"]),
                            "totalProductos": len(prods),
                            "actualizadoEn": data.get("actualizadoEn", datetime.now().isoformat()),
                            "productos": prods
                        }
                        demo_str = json.dumps(demo_payload, ensure_ascii=False, indent=2)
                        new_demo = f"// DEMO_PRELISTA_START\nconst DEMO_PRELISTA = {demo_str};\n// DEMO_PRELISTA_END"
                        if "// DEMO_PRELISTA_START" in api_c:
                            api_c = re.sub(
                                r'// DEMO_PRELISTA_START.*?// DEMO_PRELISTA_END',
                                lambda m: new_demo,
                                api_c,
                                flags=re.DOTALL
                            )
                            with open("api/prelista.js", "w", encoding="utf-8") as f_api_w:
                                f_api_w.write(api_c)
                    except Exception:
                        pass
                con_foto = sum(1 for p in prods if p.get("imagen") and str(p["imagen"]).startswith("data:image/"))
                resp = {"success": True, "total": len(prods), "conFoto": con_foto, "mensaje": f"Se sincronizaron {con_foto} fotos desde catalogos.xlsx y caché."}
            except Exception as e_s:
                resp = {"success": False, "error": str(e_s)}
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode('utf-8'))
            return

        # 2. Servir el PDF de catálogo
        elif parsed_url.path == "/catalogos.pdf":
            self.serve_file("catalogos.pdf", "application/pdf")
            return
            
        # 3. Servir el catálogo o prelista en HTML
        elif parsed_url.path in ("/catalogos.html", "/catalogos_desktop.html", "/catalogos_mobile.html", "/prelista.html", "/prelista"):
            filename = "prelista.html" if parsed_url.path in ("/prelista", "/prelista.html") else parsed_url.path[1:]
            filename = urllib.parse.unquote(filename)
            self.serve_file(filename, "text/html; charset=utf-8")
            return
            
        # 4. Servir el logotipo
        elif parsed_url.path == "/Logo%20Impor.png" or parsed_url.path == "/Logo Impor.png":
            self.serve_file("Logo Impor.png", "image/png")
            return

        # 5. Servir imágenes de productos dinámicas de la base de datos
        elif parsed_url.path.startswith("/temp_imgs/"):
            img_path = parsed_url.path[1:] # quitar la barra inicial
            img_path = urllib.parse.unquote(img_path)
            ext = os.path.splitext(img_path)[1].lower()
            mime = "image/png"
            if ext == ".webp":
                mime = "image/webp"
            elif ext in (".jpg", ".jpeg"):
                mime = "image/jpeg"
            self.serve_file(img_path, mime)
            return

        # 6. Servir otros logotipos corporativos si son requeridos por la vista previa
        elif parsed_url.path.startswith("/Logo") or parsed_url.path.endswith((".png", ".webp", ".jpg", ".jpeg")):
            filename = parsed_url.path[1:]
            filename = urllib.parse.unquote(filename)
            if os.path.exists(filename):
                ext = os.path.splitext(filename)[1].lower()
                mime = "image/png"
                if ext == ".webp":
                    mime = "image/webp"
                elif ext in (".jpg", ".jpeg"):
                    mime = "image/jpeg"
                self.serve_file(filename, mime)
                return
            else:
                self.send_error(404, "Logo not found")
                return

        # 7. Ruta raíz: Servir el panel de control web principal
        elif parsed_url.path == "/" or parsed_url.path == "/index.html":
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            
            # Generar catálogo año y mes
            meses = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                     7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
            hoy = date.today()
            mes_año_actual = f"{meses[hoy.month]} {hoy.year}"
            
            # Cargar estado de vista previa si ya existen archivos
            preview_available = "true" if os.path.exists("catalogos.html") else "false"
            pdf_available = "true" if os.path.exists("catalogos.pdf") else "false"
            desktop_available = "true" if os.path.exists("catalogos_desktop.html") else "false"
            mobile_available = "true" if os.path.exists("catalogos_mobile.html") else "false"
            
            html_ui = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Panel de Control - Importadora Rivero</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap');
    
    :root {{
      --bg-dark: #080C14;
      --bg-panel: rgba(13, 18, 30, 0.85);
      --bg-console: #030509;
      --border-glow: rgba(245, 158, 11, 0.3);
      --primary: #F59E0B;
      --primary-hover: #D97706;
      --accent: #F59E0B;
      --success: #10B981;
      --success-bg: rgba(16, 185, 129, 0.12);
      --text-main: #F8FAFC;
      --text-muted: #94A3B8;
      --border-panel: rgba(255, 255, 255, 0.08);
    }}

    body {{
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: var(--bg-dark);
      color: var(--text-main);
      margin: 0;
      padding: 0;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      overflow-x: hidden;
    }}

    header {{
      background: rgba(8, 12, 20, 0.95);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--border-panel);
      padding: 12px 25px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
    }}

    .header-left {{
      display: flex;
      align-items: center;
      gap: 15px;
    }}

    .header-logo {{
      max-height: 44px;
      object-fit: contain;
      border-radius: 4px;
      background: white;
      padding: 3px 6px;
    }}

    .header-title-container h1 {{
      font-size: 15pt;
      font-weight: 800;
      margin: 0;
      letter-spacing: 0.5px;
      color: var(--text-main);
    }}

    .header-title-container p {{
      font-size: 8pt;
      color: var(--accent);
      margin: 1px 0 0 0;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 1px;
    }}

    .status-badge {{
      display: flex;
      align-items: center;
      gap: 8px;
      background: var(--success-bg);
      color: var(--success);
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 8.5pt;
      font-weight: 700;
      border: 1px solid rgba(16, 185, 129, 0.2);
    }}

    .status-dot {{
      width: 8px;
      height: 8px;
      background-color: var(--success);
      border-radius: 50%;
      animation: pulse 1.8s infinite;
    }}

    @keyframes pulse {{
      0% {{ transform: scale(0.95); opacity: 0.5; }}
      50% {{ transform: scale(1.15); opacity: 1; }}
      100% {{ transform: scale(0.95); opacity: 0.5; }}
    }}

    .dashboard-container {{
      display: grid;
      grid-template-columns: 490px 1fr;
      gap: 18px;
      padding: 16px;
      flex-grow: 1;
      height: calc(100vh - 78px);
      box-sizing: border-box;
    }}

    .glass-panel {{
      background: var(--bg-panel);
      backdrop-filter: blur(16px);
      border: 1px solid var(--border-panel);
      border-radius: 14px;
      padding: 18px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.35);
      display: flex;
      flex-direction: column;
      height: 100%;
      box-sizing: border-box;
      overflow: hidden;
    }}

    .control-panel {{
      display: flex;
      flex-direction: column;
      gap: 12px;
      overflow-y: auto;
    }}

    .control-panel::-webkit-scrollbar {{
      width: 6px;
    }}
    .control-panel::-webkit-scrollbar-thumb {{
      background: #334155;
      border-radius: 3px;
    }}

    .section-title {{
      font-size: 10pt;
      font-weight: 800;
      color: var(--primary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin: 0;
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .smart-tabs {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 5px;
      background: rgba(15, 23, 42, 0.6);
      padding: 4px;
      border-radius: 10px;
      border: 1px solid var(--border-panel);
    }}

    .smart-tab-btn {{
      background: transparent;
      border: none;
      color: var(--text-muted);
      padding: 8px 4px;
      border-radius: 7px;
      font-family: inherit;
      font-size: 8pt;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 3px;
      transition: all 0.2s ease;
    }}

    .smart-tab-btn:hover {{
      color: var(--text-main);
      background: rgba(255, 255, 255, 0.04);
    }}

    .smart-tab-btn.active {{
      background: var(--primary);
      color: #0F172A;
      box-shadow: 0 2px 8px rgba(245, 158, 11, 0.3);
    }}

    .tab-content {{
      display: none;
      flex-direction: column;
      gap: 10px;
    }}
    .tab-content.active {{
      display: flex;
    }}

    .active-banner {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: rgba(245, 158, 11, 0.1);
      border: 1px solid rgba(245, 158, 11, 0.25);
      border-radius: 8px;
      padding: 8px 12px;
    }}
    .active-badge {{
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 9pt;
      font-weight: 800;
      color: var(--primary);
    }}
    .active-actions {{
      display: flex;
      gap: 6px;
    }}
    .btn-chip {{
      background: rgba(255, 255, 255, 0.08);
      border: 1px solid rgba(255, 255, 255, 0.1);
      color: var(--text-main);
      padding: 3px 8px;
      border-radius: 6px;
      font-size: 7.5pt;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
    }}
    .btn-chip:hover {{
      background: rgba(255, 255, 255, 0.16);
    }}
    .btn-chip.danger:hover {{
      background: rgba(239, 68, 68, 0.3);
      color: #F87171;
    }}

    .search-input-wrapper {{
      position: relative;
      display: flex;
      align-items: center;
    }}
    .search-input {{
      width: 100%;
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 8px;
      color: var(--text-main);
      padding: 10px 32px 10px 12px;
      font-family: inherit;
      font-size: 9pt;
      outline: none;
      box-sizing: border-box;
      transition: border 0.2s;
    }}
    .search-input:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 10px var(--border-glow);
    }}
    .search-clear-btn {{
      position: absolute;
      right: 10px;
      background: transparent;
      border: none;
      color: var(--text-muted);
      cursor: pointer;
      font-size: 11pt;
      display: none;
    }}

    .product-results-list {{
      max-height: 180px;
      overflow-y: auto;
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 8px;
      padding: 6px;
      display: flex;
      flex-direction: column;
      gap: 4px;
    }}
    .product-results-list::-webkit-scrollbar {{
      width: 5px;
    }}
    .product-results-list::-webkit-scrollbar-thumb {{
      background: #334155;
      border-radius: 3px;
    }}

    .product-item-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 6px 8px;
      border-radius: 6px;
      background: rgba(255, 255, 255, 0.02);
      border: 1px solid rgba(255, 255, 255, 0.03);
      transition: background 0.15s;
    }}
    .product-item-row:hover {{
      background: rgba(255, 255, 255, 0.06);
    }}
    .product-item-info {{
      display: flex;
      flex-direction: column;
      gap: 2px;
      overflow: hidden;
    }}
    .product-item-code {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 8pt;
      font-weight: 700;
      color: var(--primary);
    }}
    .product-item-name {{
      font-size: 8pt;
      color: var(--text-main);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 320px;
    }}
    .btn-item-add {{
      background: rgba(245, 158, 11, 0.15);
      border: 1px solid rgba(245, 158, 11, 0.3);
      color: var(--primary);
      padding: 4px 8px;
      border-radius: 6px;
      font-size: 7.5pt;
      font-weight: 800;
      cursor: pointer;
      white-space: nowrap;
      transition: all 0.15s;
    }}
    .btn-item-add:hover {{
      background: var(--primary);
      color: #0F172A;
    }}
    .btn-item-add.added {{
      background: rgba(16, 185, 129, 0.2);
      border-color: rgba(16, 185, 129, 0.4);
      color: #34D399;
    }}

    .brands-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 6px;
      max-height: 180px;
      overflow-y: auto;
      padding-right: 4px;
    }}

    .brand-card-btn {{
      background: rgba(15, 23, 42, 0.5);
      border: 1px solid var(--border-panel);
      padding: 8px 10px;
      border-radius: 8px;
      color: var(--text-main);
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: center;
      text-align: left;
      font-family: inherit;
      transition: all 0.2s;
    }}
    .brand-card-btn:hover {{
      border-color: var(--primary);
      background: rgba(245, 158, 11, 0.08);
    }}
    .brand-card-name {{
      font-size: 8pt;
      font-weight: 700;
    }}
    .brand-card-count {{
      font-size: 7.5pt;
      background: rgba(255, 255, 255, 0.1);
      padding: 2px 6px;
      border-radius: 10px;
      color: var(--text-muted);
      font-weight: 700;
    }}

    .templates-list {{
      display: flex;
      flex-direction: column;
      gap: 6px;
      max-height: 180px;
      overflow-y: auto;
    }}
    .template-item {{
      background: rgba(15, 23, 42, 0.6);
      border: 1px solid var(--border-panel);
      border-radius: 8px;
      padding: 8px 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .template-info {{
      display: flex;
      flex-direction: column;
      gap: 2px;
    }}
    .template-name {{
      font-size: 8.5pt;
      font-weight: 700;
      color: var(--text-main);
    }}
    .template-count {{
      font-size: 7.5pt;
      color: var(--text-muted);
    }}
    .template-actions {{
      display: flex;
      gap: 4px;
    }}

    textarea {{
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 8px;
      color: var(--text-main);
      padding: 10px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 9pt;
      resize: none;
      height: 85px;
      outline: none;
      transition: all 0.2s ease;
      box-sizing: border-box;
      width: 100%;
    }}
    textarea:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 10px var(--border-glow);
    }}

    .toggle-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: rgba(15, 23, 42, 0.4);
      padding: 8px 12px;
      border-radius: 8px;
      border: 1px solid rgba(255, 255, 255, 0.03);
    }}

    .switch {{
      position: relative;
      display: inline-block;
      width: 40px;
      height: 22px;
    }}
    .switch input {{
      opacity: 0;
      width: 0;
      height: 0;
    }}
    .slider {{
      position: absolute;
      cursor: pointer;
      top: 0; left: 0; right: 0; bottom: 0;
      background-color: #334155;
      transition: .3s;
      border-radius: 22px;
    }}
    .slider:before {{
      position: absolute;
      content: "";
      height: 14px;
      width: 14px;
      left: 4px;
      bottom: 4px;
      background-color: white;
      transition: .3s;
      border-radius: 50%;
    }}
    input:checked + .slider {{
      background-color: var(--primary);
    }}
    input:checked + .slider:before {{
      transform: translateX(18px);
    }}

    .btn-generate {{
      background: linear-gradient(135deg, var(--primary) 0%, #D97706 100%);
      color: #0F172A;
      border: none;
      border-radius: 9px;
      padding: 12px 18px;
      font-family: inherit;
      font-size: 10.5pt;
      font-weight: 800;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 8px;
      box-shadow: 0 4px 15px rgba(245, 158, 11, 0.25);
    }}
    .btn-generate:hover {{
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(245, 158, 11, 0.4);
      background: linear-gradient(135deg, #FBBF24 0%, #D97706 100%);
    }}
    .btn-generate:disabled {{
      background: #334155;
      color: var(--text-muted);
      cursor: not-allowed;
      transform: none;
      box-shadow: none;
    }}

    .console-panel {{
      flex-grow: 1;
      display: flex;
      flex-direction: column;
      min-height: 110px;
      overflow: hidden;
    }}

    .console-output {{
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 8px;
      padding: 10px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 8.5pt;
      color: #CBD5E1;
      overflow-y: auto;
      flex-grow: 1;
      white-space: pre-wrap;
      box-shadow: inset 0 2px 6px rgba(0, 0, 0, 0.7);
      line-height: 1.4;
    }}

    .log-line {{ margin-bottom: 3px; }}
    .log-success {{ color: var(--success); font-weight: 700; }}
    .log-error {{ color: #EF4444; font-weight: 700; }}
    .log-info {{ color: var(--primary); }}
    .log-warning {{ color: #F59E0B; }}

    .preview-panel {{
      display: flex;
      flex-direction: column;
      position: relative;
    }}
    .preview-header-bar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }}
    .preview-title {{
      font-size: 11pt;
      font-weight: 800;
      color: var(--text-main);
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .device-btn {{
      background: transparent;
      border: none;
      color: var(--text-muted);
      padding: 5px 10px;
      border-radius: 6px;
      font-family: inherit;
      font-size: 8.5pt;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 5px;
      transition: all 0.2s;
    }}
    .device-btn:hover {{
      color: var(--text-main);
    }}
    .device-btn.active {{
      background: rgba(255, 255, 255, 0.12);
      color: var(--text-main);
    }}

    .preview-viewport-wrapper {{
      flex-grow: 1;
      display: flex;
      justify-content: center;
      align-items: center;
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 10px;
      overflow: hidden;
      position: relative;
    }}

    iframe {{
      width: 100%;
      height: 100%;
      border: none;
      background: white;
      transition: width 0.3s ease;
    }}

    .view-mobile {{
      width: 400px;
      height: 92%;
      border: 8px solid #334155;
      border-radius: 18px;
      box-shadow: 0 15px 35px rgba(0,0,0,0.6);
    }}

    .no-preview {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 10px;
      color: var(--text-muted);
      text-align: center;
      padding: 30px;
      position: absolute;
      width: 100%;
      height: 100%;
      box-sizing: border-box;
      z-index: 10;
    }}

    .spinner {{
      border: 3px solid rgba(255, 255, 255, 0.1);
      width: 18px;
      height: 18px;
      border-radius: 50%;
      border-left-color: white;
      animation: spin 0.8s linear infinite;
      display: none;
    }}
    @keyframes spin {{
      0% {{ transform: rotate(0deg); }}
      100% {{ transform: rotate(360deg); }}
    }}
  </style>
</head>
<body>

  <!-- Header Superior -->
  <header>
    <div class="header-left">
      <img class="header-logo" src="Logo Impor.png" alt="Importadora Rivero" onerror="this.style.display='none'">
      <div class="header-title-container">
        <h1>IMPORTADORA RIVERO</h1>
        <p>Generador de Catálogos Inteligente</p>
      </div>
    </div>
    
    <div class="status-badge">
      <div class="status-dot"></div>
      Servidor Conectado
    </div>
  </header>

  <!-- Dashboard Principal -->
  <main class="dashboard-container">
    
    <!-- Lado Izquierdo: Controles Inteligentes -->
    <div class="glass-panel control-panel">
      
      <!-- Pestañas Inteligentes -->
      <div class="smart-tabs">
        <button class="smart-tab-btn active" onclick="switchSmartTab('manual')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
          <span>Pegado</span>
        </button>
        <button class="smart-tab-btn" onclick="switchSmartTab('search')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
          <span>Buscador</span>
        </button>
        <button class="smart-tab-btn" onclick="switchSmartTab('brands')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"></path><line x1="7" y1="7" x2="7.01" y2="7"></line></svg>
          <span>Marcas</span>
        </button>
        <button class="smart-tab-btn" onclick="switchSmartTab('templates')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21l-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"></path></svg>
          <span>Plantillas</span>
        </button>
        <button class="smart-tab-btn" onclick="switchSmartTab('order')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
          <span>Procesar Pedido</span>
        </button>
        <button class="smart-tab-btn" onclick="switchSmartTab('prelista')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 17l10 5 10-5M2 12l10 5 10-5M2 7l10 5 10-5"></path></svg>
          <span style="color: #38BDF8; font-weight: 800;">🚢 Prelista</span>
        </button>
      </div>

      <!-- Barra de Estado de Selección Activa -->
      <div class="active-banner">
        <div class="active-badge">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>
          <span id="badge-count-text">0 códigos listos</span>
        </div>
        <div class="active-actions">
          <button class="btn-chip" id="btn-purge-stock" onclick="depurarProductosAgotados()" title="Eliminar productos agotados usando el stock en vivo de Google Sheets" style="display: flex; align-items: center; gap: 4px; background: rgba(239, 68, 68, 0.16); color: #FCA5A5; border-color: rgba(239, 68, 68, 0.35);">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line></svg>
            <span id="btn-purge-stock-text">Quitar Agotados</span>
          </button>
          <button class="btn-chip" onclick="copiarListaSeleccionada()" title="Copiar códigos" style="display: flex; align-items: center; gap: 4px;">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
            <span>Copiar</span>
          </button>
          <button class="btn-chip" onclick="guardarComoPlantillaPrompt()" title="Guardar plantilla" style="display: flex; align-items: center; gap: 4px;">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline></svg>
            <span>Guardar</span>
          </button>
          <button class="btn-chip danger" onclick="limpiarSeleccion()" title="Vaciar selección" style="display: flex; align-items: center; gap: 4px;">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
            <span>Limpiar</span>
          </button>
        </div>
      </div>

      <!-- TAB 1: Pegado Manual -->
      <div class="tab-content active" id="tab-manual">
        <div style="display: flex; justify-content: space-between; align-items: center; gap: 6px; margin-bottom: 5px;">
          <div style="display: flex; gap: 6px; align-items: center;">
            <button class="btn-chip" id="btn-load-last" onclick="cargarUltimosCodigosGenerados()" title="Cargar los códigos del último catálogo generado para seguir trabajando sobre ellos" style="display: flex; align-items: center; gap: 4px; background: rgba(245, 158, 11, 0.2); color: var(--primary); border-color: rgba(245, 158, 11, 0.4); font-weight: 800; cursor: pointer;">
              <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>
              <span id="btn-load-last-text">🔄 Cargar Último Catálogo</span>
            </button>
            <button class="btn-chip" id="btn-append-mode" onclick="toggleAppendMode()" title="Al pegar o agregar nuevos códigos, anexar al final en vez de reemplazar" style="display: flex; align-items: center; gap: 4px; font-size: 7.5pt; background: rgba(16, 185, 129, 0.15); color: #34D399; border-color: rgba(16, 185, 129, 0.35); cursor: pointer;">
              <span id="append-mode-label">➕ Modo Anexar: ACTIVO</span>
            </button>
          </div>
          <span style="font-size: 7.5pt; color: var(--text-muted);" id="last-catalog-info-label"></span>
        </div>
        <textarea id="codes" placeholder="Pega los códigos aquí (uno por línea o separados por comas)...&#10;Ejemplo:&#10;DSM02-100&#10;FF02-100&#10;Deja vacío para procesar todo el inventario." oninput="onTextareaChanged()"></textarea>
      </div>

      <!-- TAB 2: Buscador Visual Predictivo -->
      <div class="tab-content" id="tab-search">
        <div class="search-input-wrapper" style="position: relative;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="position: absolute; left: 10px; color: var(--text-muted);"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
          <input type="text" class="search-input" id="search-input" placeholder="Buscar por código, nombre o medida..." style="padding-left: 32px;" oninput="onSearchInput(this.value)">
          <button class="search-clear-btn" id="search-clear" onclick="clearSearch()" style="display: none; align-items: center; justify-content: center;">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
          </button>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 7.5pt; color: var(--text-muted);">
          <span id="search-results-count">Cargando inventario...</span>
          <button class="btn-chip" id="btn-add-all-search" onclick="addAllSearchResults()" style="display: none; align-items: center; gap: 4px;">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
            <span>Agregar visibles</span>
          </button>
        </div>
        <div class="product-results-list" id="search-results-list">
          <div style="text-align: center; padding: 20px; color: var(--text-muted); font-size: 8.5pt;">Escribe para buscar productos al instante...</div>
        </div>
      </div>

      <!-- TAB 3: Filtro Rápido por Marcas y Categorías -->
      <div class="tab-content" id="tab-brands">
        <div style="font-size: 8pt; color: var(--text-muted);">Selecciona una marca para agregar todos sus productos o categorías:</div>
        <div class="brands-grid" id="brands-grid-container">
          <div style="grid-column: span 2; text-align: center; padding: 15px; color: var(--text-muted); font-size: 8.5pt;">Cargando marcas del inventario...</div>
        </div>
        <div id="brand-categories-wrapper" style="display: none; background: rgba(15, 23, 42, 0.6); padding: 8px; border-radius: 8px; border: 1px solid var(--border-panel); flex-direction: column; gap: 6px;">
          <div style="display: flex; justify-content: space-between; align-items: center;">
            <span id="selected-brand-title" style="font-size: 8.5pt; font-weight: 800; color: var(--primary);"></span>
            <button class="btn-chip" id="btn-add-entire-brand" onclick="addEntireBrand()" style="display: flex; align-items: center; gap: 4px;">
              <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
              <span>Agregar toda la marca</span>
            </button>
          </div>
          <div id="brand-categories-chips" style="display: flex; flex-wrap: wrap; gap: 4px; max-height: 80px; overflow-y: auto;"></div>
        </div>
      </div>

      <!-- TAB 4: Plantillas Guardadas -->
      <div class="tab-content" id="tab-templates">
        <!-- Opción Permanente: Catálogo Completo (Todo el Inventario) -->
        <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.18) 0%, rgba(245, 158, 11, 0.06) 100%); border: 1.5px solid rgba(245, 158, 11, 0.45); border-radius: 10px; padding: 10px 14px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; box-shadow: 0 4px 14px rgba(245, 158, 11, 0.15);">
          <div>
            <div style="font-weight: 800; font-size: 9pt; color: var(--primary); display: flex; align-items: center; gap: 6px;">
              <span>⭐ Catálogo Completo</span>
            </div>
            <div style="font-size: 7.5pt; color: #CBD5E1; margin-top: 2px;">Cargar todos los códigos y productos del inventario</div>
          </div>
          <button class="btn-chip" onclick="cargarTodoElInventario()" style="background: var(--primary); color: #0F172A; border: none; padding: 7px 14px; font-weight: 800; font-size: 8.5pt; display: flex; align-items: center; gap: 5px; cursor: pointer; box-shadow: 0 2px 8px rgba(245, 158, 11, 0.3);">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 11 12 14 22 4"></polyline><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path></svg>
            <span>Cargar Todo</span>
          </button>
        </div>

        <div style="display: flex; gap: 6px; margin-bottom: 8px;">
          <input type="text" id="new-template-name" placeholder="Guardar selección actual como..." style="flex-grow: 1; background: var(--bg-console); border: 1px solid var(--border-panel); color: var(--text-main); padding: 6px 10px; border-radius: 6px; font-size: 8.5pt; outline: none;">
          <button class="btn-chip" onclick="guardarPlantillaDesdeInput()" style="background: var(--primary); color: #0F172A; border: none; padding: 6px 12px; font-weight: 800;">Guardar</button>
        </div>
        <div class="templates-list" id="templates-list-container"></div>
      </div>

      <!-- TAB 5: Procesar Pedido de WhatsApp para Google Sheets -->
      <div class="tab-content" id="tab-order">
        <div style="font-size: 8pt; color: var(--text-muted); margin-bottom: 6px;">
          Pega el texto del pedido recibido por WhatsApp para generar las filas de Google Sheets:
        </div>
        <textarea id="order-raw-input" placeholder="Pega aquí el mensaje del cliente recibido en WhatsApp...&#10;&#10;Ejemplo:&#10;1. [BOM6044] BOMBIN TUBO METAL - Cantidad: 4 Cajas&#10;2. [MSS011] MASCARA SOLDAR - Cantidad: 1 Cajas" style="height: 90px;" oninput="parseWhatsAppOrder(this.value)"></textarea>
        
        <div id="order-parsed-result" style="display: none; flex-direction: column; gap: 8px; margin-top: 8px;">
          <div style="background: rgba(37, 211, 102, 0.1); border: 1px solid rgba(37, 211, 102, 0.25); border-radius: 6px; padding: 6px 10px; font-size: 8pt; color: #F8FAFC; display: flex; justify-content: space-between; align-items: center;">
            <span id="order-parsed-client-info" style="font-weight: 600;">Cliente</span>
            <button class="btn-chip" onclick="copiarInfoCliente()" title="Copiar datos del cliente" style="display: flex; align-items: center; gap: 3px;">
              <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
              <span>Copiar Cliente</span>
            </button>
          </div>

          <div style="max-height: 120px; overflow-y: auto; background: var(--bg-console); border: 1px solid var(--border-panel); border-radius: 6px; padding: 4px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 7.5pt; text-align: left;">
              <thead>
                <tr style="color: var(--text-muted); border-bottom: 1px solid var(--border-panel);">
                  <th style="padding: 3px 6px;">PEDIDO / CANT</th>
                  <th style="padding: 3px 6px;">UN/MED</th>
                  <th style="padding: 3px 6px;">DETALLE</th>
                  <th style="padding: 3px 6px;">CÓDIGO</th>
                </tr>
              </thead>
              <tbody id="order-parsed-table-body"></tbody>
            </table>
          </div>

          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 6px;">
            <button class="btn-chip" onclick="copiarFormatoCantidad()" style="background: #2563EB; color: white; border: none; padding: 8px 10px; font-weight: 800; font-size: 8pt; display: flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 4px 12px rgba(37, 99, 235, 0.4);" title="Formato IR01XX (4 columnas: CANTIDAD | UN/MED | DETALLE | CODIGO)">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
              <span>1. Por Cantidad (4 Col)</span>
            </button>
            <button class="btn-chip" onclick="copiarFormatoCajas()" style="background: #D97706; color: white; border: none; padding: 8px 10px; font-weight: 800; font-size: 8pt; display: flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 4px 12px rgba(217, 119, 6, 0.4);" title="Formato IR01ML (5 columnas: CAJAS | CANT. UNI | UN/MED | DETALLE | CODIGO)">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
              <span>2. Por Cajas (5 Col)</span>
            </button>
          </div>
          <div style="display: flex; gap: 6px;">
            <button class="btn-chip" onclick="cargarPedidoAlGenerador()" title="Cargar códigos a la lista para generar catálogo" style="flex-grow: 1; background: rgba(255, 255, 255, 0.08); display: flex; align-items: center; justify-content: center; gap: 4px; padding: 6px 10px;">
              <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>
              <span>Cargar estos códigos al Generador</span>
            </button>
          </div>
        </div>
      </div>

      <!-- TAB 6: Prelista / Mercadería en Tránsito (Google Sheets) -->
      <div class="tab-content" id="tab-prelista">
        <!-- 1. Fila de Hoja / Contenedor Google Sheets -->
        <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid var(--border-panel); border-radius: 8px; padding: 10px; margin-bottom: 8px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; flex-wrap: wrap; gap: 6px;">
            <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
              <span style="font-size: 8pt; font-weight: 800; color: #38BDF8;">📦 Hojas Google Sheets:</span>
              <select id="prelista-panel-sheet-select" style="background: var(--bg-console); border: 1px solid var(--border-panel); color: #FFFFFF; font-size: 8pt; font-weight: 700; padding: 5px 8px; border-radius: 6px; outline: none; cursor: pointer;">
                <option value="">(Detectando hojas en vivo...)</option>
              </select>
              
              <button type="button" class="btn-chip" onclick="cargarHojaSeleccionada(false)" title="Cargar solo esta hoja reemplazando la lista" style="background: rgba(56, 189, 248, 0.15); color: #38BDF8; border-color: rgba(56, 189, 248, 0.35); font-weight: 700; font-size: 7.5pt; padding: 5px 8px; cursor: pointer;">
                🔄 Cargar
              </button>

              <button type="button" class="btn-chip" onclick="cargarHojaSeleccionada(true)" title="Sumar / anexar los productos de esta hoja a la lista sin borrar los existentes" style="background: linear-gradient(135deg, #10B981, #059669); color: white; border: none; font-weight: 800; font-size: 7.5pt; padding: 5px 10px; cursor: pointer; box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3);">
                ➕ Sumar Hoja
              </button>

              <button type="button" class="btn-chip" id="btn-toggle-multi-sheets" onclick="toggleMultiSheetsPanel()" title="Marcar varias hojas con casillas y agregarlas todas juntas" style="background: rgba(245, 158, 11, 0.18); color: #FBBF24; border: 1px solid rgba(245, 158, 11, 0.4); font-weight: 700; font-size: 7.5pt; padding: 5px 9px; cursor: pointer;">
                📑 Elegir Varias...
              </button>
            </div>
            
            <div style="display: flex; gap: 6px;">
              <button type="button" class="btn-chip" onclick="consultarPrelistaPanel(null, true)" title="Refrescar datos en vivo desde Google Sheets" style="display: flex; align-items: center; gap: 4px; background: rgba(56, 189, 248, 0.15); color: #38BDF8; border-color: rgba(56, 189, 248, 0.35); cursor: pointer; font-size: 7.5pt;">
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>
                <span>Sincronizar</span>
              </button>
              <a href="/prelista.html" target="_blank" class="btn-chip" style="display: flex; align-items: center; gap: 4px; background: rgba(16, 185, 129, 0.15); color: #34D399; border-color: rgba(16, 185, 129, 0.35); text-decoration: none; font-size: 7.5pt;">
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
                <span>Ver Web Prelista</span>
              </a>
            </div>
          </div>

          <!-- Panel Desplegable: Selección de Múltiples Hojas -->
          <div id="multi-sheets-panel" style="display: none; background: rgba(9, 13, 22, 0.95); border: 1px solid rgba(56, 189, 248, 0.35); border-radius: 6px; padding: 10px; margin-bottom: 8px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; border-bottom: 1px solid rgba(255, 255, 255, 0.08); padding-bottom: 6px;">
              <span style="font-size: 7.5pt; font-weight: 800; color: #38BDF8;">📑 MARCA LAS HOJAS QUE DESEAS COMBINAR:</span>
              <div style="display: flex; gap: 4px;">
                <button type="button" class="btn-chip" onclick="marcarRecientesHojas(2)" style="font-size: 6.8pt; padding: 2px 6px;">Últimas 2</button>
                <button type="button" class="btn-chip" onclick="marcarRecientesHojas(3)" style="font-size: 6.8pt; padding: 2px 6px;">Últimas 3</button>
                <button type="button" class="btn-chip" onclick="desmarcarTodasHojas()" style="font-size: 6.8pt; padding: 2px 6px;">Desmarcar</button>
                <button type="button" onclick="toggleMultiSheetsPanel()" style="background: none; border: none; color: #94A3B8; font-size: 9pt; cursor: pointer; padding: 0 4px;" title="Cerrar panel">✕</button>
              </div>
            </div>
            
            <div id="multi-sheets-checkboxes-grid" style="display: flex; flex-wrap: wrap; gap: 6px; max-height: 110px; overflow-y: auto; padding: 4px; margin-bottom: 8px;">
              <!-- Se generará dinámicamente -->
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 6px;">
              <button type="button" class="btn-chip" onclick="cargarHojasSeleccionadasMultiple(false)" style="background: rgba(56, 189, 248, 0.2); color: #38BDF8; border-color: rgba(56, 189, 248, 0.4); font-size: 7.5pt; font-weight: 700; padding: 5px 10px; cursor: pointer;">
                🔄 Cargar Marcadas (Reemplazar)
              </button>
              <button type="button" class="btn-chip" onclick="cargarHojasSeleccionadasMultiple(true)" style="background: linear-gradient(135deg, #10B981, #059669); color: white; border: none; font-size: 7.5pt; font-weight: 800; padding: 5px 12px; cursor: pointer; box-shadow: 0 2px 8px rgba(16, 185, 129, 0.35);">
                ➕ Sumar Marcadas a la Lista
              </button>
            </div>
          </div>

          <!-- Fila Informativa de Hojas Incluidas Actualmente -->
          <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap; font-size: 7.5pt; padding-top: 4px; border-top: 1px solid rgba(255, 255, 255, 0.05);">
            <span style="color: #94A3B8; font-weight: 700;">Hojas en esta prelista:</span>
            <div id="loaded-sheets-badges" style="display: flex; gap: 4px; flex-wrap: wrap;">
              <span style="background: rgba(148, 163, 184, 0.15); color: #94A3B8; border: 1px solid rgba(148, 163, 184, 0.3); border-radius: 4px; padding: 2px 6px; font-weight: 600;">(Ninguna hoja cargada aún)</span>
            </div>
          </div>
        </div>

        <!-- 2. Entrada Manual de Códigos y Número de Cajas -->
        <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid var(--border-panel); border-radius: 8px; padding: 8px; margin-bottom: 8px;">
          <div style="font-size: 7.5pt; font-weight: 800; color: #38BDF8; margin-bottom: 6px;">
            <span>➕ AGREGAR CÓDIGOS Y NÚMERO DE CAJAS A LA PRELISTA:</span>
          </div>
          <div style="display: flex; gap: 6px; margin-bottom: 6px;">
            <input type="text" id="prelista-input-code" placeholder="Código (ej: TPBX0171)" style="flex: 2; background: var(--bg-console); border: 1px solid var(--border-panel); color: #FFFFFF; padding: 6px 10px; border-radius: 6px; font-size: 8pt; outline: none;">
            <input type="number" id="prelista-input-boxes" placeholder="Cajas" min="1" value="10" style="width: 75px; background: var(--bg-console); border: 1px solid var(--border-panel); color: #34D399; font-weight: bold; padding: 6px 8px; border-radius: 6px; font-size: 8pt; text-align: center; outline: none;">
            <button class="btn-chip" onclick="agregarItemManualPrelista()" style="background: var(--primary); color: #0F172A; border: none; padding: 6px 12px; font-weight: 800; font-size: 8pt; white-space: nowrap; cursor: pointer;">
              ➕ Agregar
            </button>
          </div>
          <div style="display: flex; gap: 6px; align-items: center;">
            <input type="text" id="prelista-multi-input" placeholder="O pega varios (ej: DSM07-115 10, DZR110 5, DMY02-235 8)..." style="flex-grow: 1; background: var(--bg-console); border: 1px solid var(--border-panel); color: #FFFFFF; padding: 5px 8px; border-radius: 6px; font-size: 7.5pt; outline: none;">
            <button class="btn-chip" onclick="agregarMultiplesCodigosPrelista()" style="background: rgba(56, 189, 248, 0.2); color: #38BDF8; border-color: rgba(56, 189, 248, 0.4); font-size: 7.5pt; padding: 5px 8px; white-space: nowrap;">
              📥 Cargar Varios
            </button>
          </div>
        </div>

        <!-- 3. BOTÓN PRINCIPAL DE GENERAR / GRABAR DIRECTO EN LA WEB -->
        <div style="margin-bottom: 8px;">
          <button type="button" id="btn-save-prelista-web" onclick="guardarYGrabarWebPrelista()" style="width: 100%; background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); color: #FFFFFF; border: 1.5px solid #38BDF8; border-radius: 8px; padding: 10px 14px; font-weight: 800; font-size: 9pt; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; box-shadow: 0 4px 15px rgba(2, 132, 199, 0.35); transition: all 0.2s;">
            <span style="font-size: 13pt;">🚀</span>
            <span>GUARDAR Y GRABAR EN WEB PRELISTA (<span id="lbl-prelista-count">0</span> productos)</span>
          </button>
          <div style="font-size: 7pt; color: #94A3B8; text-align: center; margin-top: 4px;">
            Graba directamente los productos en <b>prelista.html</b> para que tus clientes los vean al instante sin demoras.
          </div>
        </div>

        <div style="display: flex; gap: 6px; margin-bottom: 8px;">
          <button class="btn-chip" onclick="sincronizarFotosPrelistaExcel()" style="flex: 1; background: rgba(56, 189, 248, 0.15); color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.35); font-weight: 800; font-size: 7.5pt; padding: 6px 8px; display: flex; align-items: center; justify-content: center; gap: 4px;" title="Extraer y sincronizar fotos directo desde catalogos.xlsx">
            <span>🖼️ Sincronizar Fotos Excel</span>
          </button>
          <button class="btn-chip" onclick="publicarVercelDesdePrelista()" style="flex: 1; background: linear-gradient(135deg, #10B981, #059669); color: white; border: none; font-weight: 800; font-size: 7.5pt; padding: 6px 8px; display: flex; align-items: center; justify-content: center; gap: 4px;" title="Subir los cambios de prelista.html a Vercel en la nube">
            <span>☁️ Publicar en Vercel Ahora</span>
          </button>
          <button class="btn-chip" onclick="cargarCodigosPrelistaAlGenerador()" style="flex: 1; background: rgba(255, 255, 255, 0.08); font-size: 7.5pt; padding: 6px 8px; display: flex; align-items: center; justify-content: center; gap: 4px;" title="Copiar códigos de esta prelista al catálogo principal">
            <span>📋 Cargar al Catálogo</span>
          </button>
          <button class="btn-chip danger" onclick="vaciarPrelistaCache()" style="font-size: 7.5pt; padding: 6px 8px;" title="Limpiar la lista previa">
            <span>🗑️ Limpiar</span>
          </button>
        </div>

        <!-- 4. Tabla Previa de Mercadería en Tránsito -->
        <div style="max-height: 220px; overflow-y: auto; background: var(--bg-console); border: 1px solid var(--border-panel); border-radius: 6px; padding: 4px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 7.5pt; text-align: left;">
            <thead>
              <tr style="color: var(--text-muted); border-bottom: 1px solid var(--border-panel);">
                <th style="padding: 4px 6px;">CÓDIGO</th>
                <th style="padding: 4px 6px;">DETALLE</th>
                <th style="padding: 4px 6px; text-align: center;">CAJAS</th>
                <th style="padding: 4px 6px; text-align: center;">LIBRES</th>
                <th style="padding: 4px 6px;">PRECIO CAJA</th>
                <th style="padding: 4px 6px; text-align: center;">ACCIÓN</th>
              </tr>
            </thead>
            <tbody id="prelista-panel-table-body">
              <tr>
                <td colspan="6" style="text-align: center; padding: 15px; color: var(--text-dim);">Abre esta pestaña o haz clic en "Sincronizar Hoja" para consultar Google Sheets...</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Ajustes de Generación -->
      <div class="toggle-row">
        <div>
          <span class="label-text" style="font-weight: 600; font-size: 8.5pt;">Descargar base desde Google Drive</span>
          <div style="font-size: 7.5pt; color: var(--text-dim); margin-top: 2px;">(Por defecto desactivado: se usa <b>catalogos.xlsx</b> local sin límite de peso)</div>
        </div>
        <label class="switch">
          <input type="checkbox" id="sync">
          <span class="slider"></span>
        </label>
      </div>
      
      <div class="toggle-row">
        <span class="label-text" style="color: var(--accent); font-weight: 700; font-size: 8.5pt;">Forzar regeneración de imágenes</span>
        <label class="switch">
          <input type="checkbox" id="force_images">
          <span class="slider"></span>
        </label>
      </div>

      <div class="toggle-row" style="background: rgba(239, 68, 68, 0.08); border: 1px solid rgba(239, 68, 68, 0.25);">
        <div style="display: flex; align-items: center; gap: 6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#F87171" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line></svg>
          <span class="label-text" style="color: #F87171; font-weight: 700; font-size: 8.5pt;">Omitir productos agotados (solo con stock)</span>
        </div>
        <label class="switch">
          <input type="checkbox" id="filter_stock">
          <span class="slider"></span>
        </label>
      </div>

      <!-- Teléfono de WhatsApp para pedidos -->
      <div class="toggle-row">
        <div style="display: flex; align-items: center; gap: 6px;">
          <svg viewBox="0 0 24 24" width="15" height="15" fill="#25D366"><path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946.003-6.556 5.338-11.891 11.893-11.891 3.181.001 6.167 1.24 8.413 3.488 2.245 2.248 3.481 5.236 3.48 8.414-.003 6.557-5.338 11.892-11.893 11.892-1.99-.001-3.951-.5-5.688-1.448l-6.305 1.654zm6.597-3.807c1.676.995 3.276 1.591 5.392 1.592 5.448 0 9.886-4.434 9.889-9.885.002-5.462-4.415-9.89-9.881-9.892-5.452 0-9.887 4.434-9.889 9.884-.001 2.225.651 3.891 1.746 5.634l-.999 3.648 3.742-.981zm11.387-5.464c-.074-.124-.272-.198-.57-.347-.297-.149-1.758-.868-2.031-.967-.272-.099-.47-.149-.669.149-.198.297-.768.967-.941 1.165-.173.198-.347.223-.644.074-.297-.149-1.255-.462-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.297-.347.446-.521.151-.172.2-.296.3-.495.099-.198.05-.372-.025-.521-.075-.148-.669-1.611-.916-2.206-.242-.579-.487-.501-.669-.51l-.57-.01c-.198 0-.52.074-.792.372s-1.04 1.016-1.04 2.479 1.065 2.876 1.213 3.074c.149.198 2.095 3.2 5.076 4.487.709.306 1.263.489 1.694.626.712.226 1.36.194 1.872.118.571-.085 1.758-.719 2.006-1.413.248-.695.248-1.29.173-1.414z"/></svg>
          <span class="label-text" style="font-weight: 600; font-size: 8.5pt;">WhatsApp Pedidos</span>
        </div>
        <input type="text" id="whatsapp" placeholder="Ej: +59170000000" style="background: var(--bg-console); border: 1px solid var(--border-panel); color: #25D366; padding: 5px 8px; border-radius: 6px; font-family: 'JetBrains Mono', monospace; font-size: 8.5pt; width: 140px; outline: none; font-weight: 700;">
      </div>

      <!-- Control de Stock en Vivo desde Google Drive -->
      <div class="toggle-row" style="background: rgba(34, 197, 94, 0.08); border: 1px solid rgba(34, 197, 94, 0.22); padding: 8px 10px; border-radius: 8px; margin-top: 4px;">
        <div style="display: flex; flex-direction: column; gap: 2px;">
          <div style="display: flex; align-items: center; gap: 6px;">
            <span style="width: 8px; height: 8px; border-radius: 50%; background: #22C55E; box-shadow: 0 0 6px #22C55E; display: inline-block;"></span>
            <span style="font-weight: 700; font-size: 8.5pt; color: #86EFAC;">Stock en Vivo (Google Drive)</span>
          </div>
          <span style="font-size: 7.5pt; color: #94A3B8;">Sincronizado: Uyus + Varios</span>
        </div>
        <button type="button" onclick="testStockConnection(event)" class="btn-chip" style="font-size: 7.5pt; padding: 4px 8px; background: rgba(34, 197, 94, 0.15); color: #86EFAC; border: 1px solid rgba(34, 197, 94, 0.35); cursor: pointer;">Probar API</button>
      </div>
      
      <!-- Botón de Generar -->
      <button class="btn-generate" id="btn-run" onclick="iniciarGeneracion()">
        <span class="spinner" id="btn-spinner"></span>
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
        <span id="btn-text">Empezar Generación</span>
      </button>
      
      <!-- Actividad y Logs (Consola) -->
      <div class="console-panel">
        <div class="section-title" style="margin-bottom: 6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 17 10 11 4 5"></polyline><line x1="12" y1="19" x2="20" y2="19"></line></svg>
          <span>Actividad del Servidor</span>
        </div>
        <div class="console-output" id="console">Panel listo. Selecciona tus productos y presiona 'Empezar Generación'...</div>
      </div>

    </div>

    <!-- Lado Derecho: Vista Previa Interactiva -->
    <div class="glass-panel preview-panel">
      
      <div class="preview-header-bar">
        <div class="preview-title">
          <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
          Vista Previa del Catálogo ({mes_año_actual})
        </div>
        
        <div class="device-selectors" style="display: flex; gap: 6px;">
          <div style="display: flex; gap: 3px; background: rgba(15, 23, 42, 0.5); padding: 3px; border-radius: 8px; border: 1px solid var(--border-panel);">
            <button class="device-btn active" id="btn-device-desktop" onclick="setDevice('desktop')" style="display: flex; align-items: center; gap: 5px;">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line><line x1="12" y1="17" x2="12" y2="21"></line></svg>
              <span>Escritorio</span>
            </button>
            <button class="device-btn" id="btn-device-mobile" onclick="setDevice('mobile')" style="display: flex; align-items: center; gap: 5px;">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"></rect><line x1="12" y1="18" x2="12.01" y2="18"></line></svg>
              <span>Celular</span>
            </button>
          </div>
          <button class="device-btn" id="btn-full-preview" onclick="verCompleto()" style="background-color: rgba(245, 158, 11, 0.15); color: var(--accent); border: 1px solid rgba(245, 158, 11, 0.3); display: flex; align-items: center; gap: 5px;" {"" if preview_available == "true" else "disabled"}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
            <span>Ver Completo</span>
          </button>
          <button class="device-btn" id="btn-purge-preview" onclick="depurarAgotadosEnVistaPrevia()" style="background-color: rgba(239, 68, 68, 0.18); color: #FCA5A5; border: 1px solid rgba(239, 68, 68, 0.35); display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight: 700;" title="Quitar productos agotados directamente de la vista previa">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line></svg>
            <span>Quitar Agotados</span>
          </button>
          <a id="btn-download-html" href="#" download="catalogos_desktop.html" class="device-btn" style="background-color: rgba(16, 185, 129, 0.15); color: var(--success); border: 1px solid rgba(16, 185, 129, 0.3); text-decoration: none; display: flex; align-items: center; gap: 5px; pointer-events: none; opacity: 0.5;" onclick="return document.getElementById('btn-download-html').getAttribute('href') !== '#'">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            <span>Descargar HTML</span>
          </a>
          <button class="device-btn" id="btn-publish-vercel" onclick="publicarEnVercel()" style="background-color: rgba(99, 102, 241, 0.18); color: #818CF8; border: 1px solid rgba(99, 102, 241, 0.4); display: flex; align-items: center; gap: 5px; cursor: pointer; font-weight: 700;" title="Subir catálogo a GitHub y actualizar en Vercel">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor"><path d="M12 1L24 22H0L12 1Z"/></svg>
            <span>Publicar en Vercel</span>
          </button>
        </div>
      </div>
      
      <!-- Contenedor del Iframe -->
      <div class="preview-viewport-wrapper">
        <div class="no-preview" id="no-preview" style="display: {'none' if preview_available == 'true' else 'flex'};">
          <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom: 8px; opacity: 0.3;"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
          <h3 style="margin: 0;">El folleto aún no ha sido generado</h3>
          <p style="margin: 0; font-size: 9pt;">Los resultados aparecerán aquí una vez inicies la generación.</p>
        </div>
        
        <iframe id="preview-iframe" onload="onIframeLoaded()" src="{"catalogos_desktop.html" if desktop_available == "true" else ("catalogos.html" if preview_available == "true" else "about:blank")}" style="display: {'block' if preview_available == 'true' else 'none'};"></iframe>
      </div>

    </div>

  </main>

  <script>
    const consoleDiv = document.getElementById('console');
    const btnRun = document.getElementById('btn-run');
    const btnText = document.getElementById('btn-text');
    const btnSpinner = document.getElementById('btn-spinner');
    const noPreview = document.getElementById('no-preview');
    const iframe = document.getElementById('preview-iframe');
    const textareaCodes = document.getElementById('codes');
    
    let hasDesktopFile = {desktop_available};
    let hasMobileFile = {mobile_available};
    let currentDevice = 'desktop';

    // ─── ESTADO GLOBAL DE PRODUCTOS Y PLANTILLAS ───
    let allInventoryProducts = [];
    let inventoryBrands = {{}};
    let inventoryCategories = {{}};
    let selectedCodesSet = new Set();
    let currentSearchResults = [];
    let activeBrandSelected = null;
    let appendModeActive = true;
    let ultimosCodigosGeneradosCache = [];

    // ─── 0. FUNCIONES DE HISTORIAL, ANEXAR Y DEPURACIÓN DE STOCK ───
    function toggleAppendMode() {{
      appendModeActive = !appendModeActive;
      const btn = document.getElementById('btn-append-mode');
      const lbl = document.getElementById('append-mode-label');
      if (appendModeActive) {{
        btn.style.background = 'rgba(16, 185, 129, 0.15)';
        btn.style.color = '#34D399';
        btn.style.borderColor = 'rgba(16, 185, 129, 0.35)';
        lbl.innerText = '➕ Modo Anexar: ACTIVO';
        log('[MODO] Modo Anexar activado: los nuevos códigos se sumarán a la lista existente.');
      }} else {{
        btn.style.background = 'rgba(255, 255, 255, 0.08)';
        btn.style.color = 'var(--text-muted)';
        btn.style.borderColor = 'rgba(255, 255, 255, 0.15)';
        lbl.innerText = '🔄 Modo Reemplazar';
        log('[MODO] Modo Reemplazar activado: al cargar o pegar se sustituirá la lista.');
      }}
    }}

    async function cargarUltimosCodigosGenerados() {{
      const btnText = document.getElementById('btn-load-last-text');
      const originalText = btnText ? btnText.innerText : '🔄 Cargar Último Catálogo';
      if (btnText) btnText.innerText = 'Cargando...';
      
      try {{
        let codes = [];
        const res = await fetch('/api/ultimos_codigos');
        if (res.ok) {{
          const data = await res.json();
          codes = data.codigos || [];
        }}
        
        if (codes.length === 0) {{
          try {{
            const localSaved = localStorage.getItem('rivero_last_generated_codes');
            if (localSaved) {{
              codes = JSON.parse(localSaved);
            }}
          }} catch(e) {{}}
        }}

        // Filtrar cualquier texto inválido o variable residual
        codes = (codes || []).filter(c => c && typeof c === 'string' && !c.includes('$') && !c.includes('{') && !c.includes('}'));
        
        if (codes.length === 0) {{
          alert('No se encontraron códigos del último catálogo generado todavía. Genera un catálogo primero.');
          if (btnText) btnText.innerText = originalText;
          return;
        }}

        ultimosCodigosGeneradosCache = codes;

        if (appendModeActive && selectedCodesSet.size > 0) {{
          let agregados = 0;
          codes.forEach(c => {{
            const u = (c || '').toUpperCase().trim();
            if (u && !selectedCodesSet.has(u)) {{
              selectedCodesSet.add(u);
              agregados++;
            }}
          }});
          syncSetToTextarea();
          log(`[HISTORIAL] Se anexaron ${{agregados}} códigos del último catálogo (Total ahora: ${{selectedCodesSet.size}}).`, 'success');
          alert(`¡Listo! Se anexaron ${{agregados}} códigos sobre lo que ya tenías.\\nTotal actual: ${{selectedCodesSet.size}} códigos listos.`);
        }} else {{
          selectedCodesSet.clear();
          codes.forEach(c => {{
            const u = (c || '').toUpperCase().trim();
            if (u) selectedCodesSet.add(u);
          }});
          syncSetToTextarea();
          log(`[HISTORIAL] Se cargaron exitosamente ${{selectedCodesSet.size}} códigos del último catálogo generado.`, 'success');
          alert(`🎉 ¡Listo! Se cargaron los ${{selectedCodesSet.size}} productos del último catálogo.`);
        }}
        
        const infoLbl = document.getElementById('last-catalog-info-label');
        if (infoLbl) infoLbl.innerText = `(${{codes.length}} en historial)`;
        switchSmartTab('manual');
      }} catch(err) {{
        log(`[ERROR] No se pudo cargar el historial: ${{err.message}}`, 'error');
      }} finally {{
        if (btnText) btnText.innerText = originalText;
      }}
    }}

    async function consultarUltimosCodigosSilencioso() {{
      try {{
        const res = await fetch('/api/ultimos_codigos');
        if (res.ok) {{
          const data = await res.json();
          let codes = (data.codigos || []).filter(c => c && typeof c === 'string' && !c.includes('$') && !c.includes('{') && !c.includes('}'));
          if (codes.length > 0) {{
            ultimosCodigosGeneradosCache = codes;
            const infoLbl = document.getElementById('last-catalog-info-label');
            if (infoLbl) infoLbl.innerText = `(${{codes.length}} en historial)`;
            const btnText = document.getElementById('btn-load-last-text');
            if (btnText) btnText.innerText = `🔄 Cargar Último (${{codes.length}})`;
          }}
        }}
      }} catch(e) {{}}
    }}

    async function depurarProductosAgotados() {{
      const btn = document.getElementById('btn-purge-stock');
      const btnText = document.getElementById('btn-purge-stock-text');
      const originalText = btnText ? btnText.innerText : 'Quitar Agotados';
      if (btnText) btnText.innerText = 'Consultando stock...';
      if (btn) btn.disabled = true;

      log(">>> [STOCK] Consultando stock en tiempo real desde Google Sheets para depurar agotados...");

      try {{
        const res = await fetch('/api/stock', {{ cache: 'no-store' }});
        if (!res.ok) throw new Error("HTTP " + res.status);
        const stockMap = await res.json();
        if (stockMap.error) throw new Error(stockMap.error);

        if (selectedCodesSet.size === 0 && allInventoryProducts.length > 0) {{
          allInventoryProducts.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
        }}

        if (selectedCodesSet.size === 0) {{
          alert("No hay códigos seleccionados para depurar.");
          if (btnText) btnText.innerText = originalText;
          if (btn) btn.disabled = false;
          return;
        }}

        const codigosAEliminar = [];
        selectedCodesSet.forEach(rawCode => {{
          const norm = rawCode.toUpperCase().replace(/\\s+/g, '');
          const info = stockMap[norm] || stockMap[rawCode.toUpperCase()];
          if (!info) {{
            codigosAEliminar.push(rawCode);
          }} else {{
            const stk = typeof info.s === 'number' ? info.s : (typeof info.stockActual === 'number' ? info.stockActual : (typeof info.stock === 'number' ? info.stock : (info.stock === true ? 999 : 0)));
            const est = info.e || info.estado || "";
            if (stk <= 0 || est === "AGOTADO" || info.stock === false) {{
              codigosAEliminar.push(rawCode);
            }}
          }}
        }});

        codigosAEliminar.forEach(c => selectedCodesSet.delete(c));
        syncSetToTextarea();

        const eliminados = codigosAEliminar.length;
        const restantes = selectedCodesSet.size;

        log(`>>> [STOCK OK] Depuración completada: se quitaron ${{eliminados}} productos agotados. Quedan ${{restantes}} disponibles con stock.`, 'success');
        alert(`⚡ ¡Depuración de stock completada!\\n\\n• Productos agotados eliminados: ${{eliminados}}\\n• Productos disponibles con stock: ${{restantes}}`);
      }} catch(err) {{
        log(`>>> [STOCK ERROR] No se pudo consultar el stock: ${{err.message}}`, 'error');
        alert("❌ Ocurrió un inconveniente al consultar el stock en Google Sheets. Revisa la consola.");
      }} finally {{
        if (btnText) btnText.innerText = originalText;
        if (btn) btn.disabled = false;
      }}
    }}

    function depurarAgotadosEnVistaPrevia() {{
      if (!confirm("¿Deseas quitar todos los productos agotados del catálogo en vista previa y de la lista del generador?")) return;
      try {{
        if (iframe && iframe.contentWindow) {{
          iframe.contentWindow.postMessage({{ type: 'PURGE_OUT_OF_STOCK' }}, '*');
        }}
      }} catch(e) {{}}
      depurarProductosAgotados();
    }}

    function onIframeLoaded() {{
      try {{
        if (iframe && iframe.contentDocument && iframe.contentDocument.body) {{
          iframe.contentDocument.body.classList.add('is-generator-iframe');
        }}
      }} catch(e) {{}}
    }}

    // Cargar inventario desde API al inicio
    async function cargarInventarioAPI() {{
      try {{
        const res = await fetch('/api/productos');
        const data = await res.json();
        allInventoryProducts = data.productos || [];
        inventoryBrands = data.marcas || {{}};
        inventoryCategories = data.categorias || {{}};
        
        renderBrandsGrid();
        document.getElementById('search-results-count').innerText = `${{allInventoryProducts.length}} productos disponibles`;
        initSelectedFromTextarea();
      }} catch(err) {{
        console.error("Error al cargar inventario:", err);
      }}
    }}

    // Sincronizar códigos activos desde el textarea
    function initSelectedFromTextarea() {{
      selectedCodesSet.clear();
      const raw = textareaCodes.value;
      const tokens = raw.split(/[\\r\\n,;\\t]+/).map(s => s.trim()).filter(Boolean);
      tokens.forEach(c => selectedCodesSet.add(c.toUpperCase()));
      updateActiveCounter();
    }}

    function onTextareaChanged() {{
      initSelectedFromTextarea();
      refreshVisibleSearchButtons();
    }}

    function syncSetToTextarea() {{
      textareaCodes.value = Array.from(selectedCodesSet).join('\\n');
      updateActiveCounter();
      refreshVisibleSearchButtons();
    }}

    function updateActiveCounter() {{
      const count = selectedCodesSet.size;
      document.getElementById('badge-count-text').innerText = count === 0 ? 'Todo el inventario (0 seleccionados)' : `${{count}} código(s) listos`;
    }}

    function switchSmartTab(tabId) {{
      const tabs = ['manual', 'search', 'brands', 'templates', 'order', 'prelista'];
      tabs.forEach(t => {{
        const el = document.getElementById('tab-' + t);
        if (el) el.classList.toggle('active', t === tabId);
      }});
      
      const buttons = document.querySelectorAll('.smart-tab-btn');
      buttons.forEach((btn, idx) => {{
        btn.classList.toggle('active', tabs[idx] === tabId);
      }});

      if (tabId === 'templates') renderTemplatesList();
      if (tabId === 'search') refreshVisibleSearchButtons();
      if (tabId === 'order') {{
        const raw = document.getElementById('order-raw-input')?.value;
        if (raw) parseWhatsAppOrder(raw);
      }}
      if (tabId === 'prelista') {{
        if (!prelistaProductsCache || prelistaProductsCache.length === 0) {{
          consultarPrelistaPanel();
        }}
      }}
    }}

    // ─── 0.5. PRELISTA EN EL GENERADOR (MULTIPLE HOJAS) ───
    let prelistaProductsCache = [];
    let prelistaLoadedSheets = new Set();
    let availableSheetsCache = [];

    function updateLoadedSheetsBadges() {{
      const container = document.getElementById('loaded-sheets-badges');
      if (!container) return;
      if (prelistaLoadedSheets.size === 0) {{
        container.innerHTML = `<span style="background: rgba(148, 163, 184, 0.15); color: #94A3B8; border: 1px solid rgba(148, 163, 184, 0.3); border-radius: 4px; padding: 2px 6px; font-weight: 600;">(Ninguna hoja cargada aún)</span>`;
        return;
      }}
      container.innerHTML = Array.from(prelistaLoadedSheets).map(s => {{
        const count = prelistaProductsCache.filter(p => (p._origenHojas || []).includes(s)).length;
        const countTxt = count > 0 ? ` (${{count}})` : '';
        return `
          <span style="display: inline-flex; align-items: center; gap: 4px; background: rgba(56, 189, 248, 0.18); color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.4); border-radius: 4px; padding: 2px 6px; font-weight: 700;">
            <span>📦 ${{s}}${{countTxt}}</span>
            <button type="button" onclick="quitarHojaDePrelista('${{s}}')" title="Quitar productos de ${{s}}" style="background: none; border: none; color: #F87171; font-weight: bold; cursor: pointer; padding: 0 2px; font-size: 8pt; line-height: 1;">✕</button>
          </span>
        `;
      }}).join('');
    }}

    function updatePrelistaBadge() {{
      const lbl = document.getElementById('lbl-prelista-count');
      if (lbl) lbl.textContent = prelistaProductsCache.length;
      const btnLbl = document.getElementById('btn-load-prelista-label');
      if (btnLbl) {{
        btnLbl.innerText = `📋 Cargar Códigos al Catálogo (${{prelistaProductsCache.length}} items)`;
      }}
      updateLoadedSheetsBadges();
    }}

    function fusionarProductosPrelista(nuevosProductos, nombreHoja, modoSumar) {{
      if (!modoSumar) {{
        prelistaProductsCache = [];
        prelistaLoadedSheets.clear();
      }}
      
      if (nombreHoja) {{
        prelistaLoadedSheets.add(nombreHoja);
      }}

      let agregadosNuevos = 0;
      let cajasSumadas = 0;

      nuevosProductos.forEach(np => {{
        const cod = String(np.codigo || '').toUpperCase().trim();
        if (!cod) return;
        
        const existing = prelistaProductsCache.find(p => String(p.codigo || '').toUpperCase().trim() === cod);
        if (existing) {{
          const addCajas = parseInt(np.cajasVienen) || 0;
          const addReserva = parseInt(np.stockReserva) || addCajas;
          existing.cajasVienen = (parseInt(existing.cajasVienen) || 0) + addCajas;
          existing.stockReserva = (parseInt(existing.stockReserva) || 0) + addReserva;
          if (!existing._origenHojas) existing._origenHojas = [];
          if (nombreHoja && !existing._origenHojas.includes(nombreHoja)) {{
            existing._origenHojas.push(nombreHoja);
          }}
          cajasSumadas++;
        }} else {{
          const cloned = {{ ...np }};
          cloned._origenHojas = nombreHoja ? [nombreHoja] : ['General'];
          prelistaProductsCache.push(cloned);
          agregadosNuevos++;
        }}
      }});

      return {{ agregadosNuevos, cajasSumadas }};
    }}

    function actualizarSelectYCasillasHojas(sheets, currentSelectedSheet) {{
      const sel = document.getElementById('prelista-panel-sheet-select');
      if (sel) {{
        const prevVal = sel.value;
        sel.innerHTML = '';
        sheets.forEach(s => {{
          const opt = document.createElement('option');
          opt.value = s;
          opt.textContent = `📦 ${{s}}`;
          if (s === currentSelectedSheet || (!currentSelectedSheet && s === prevVal)) {{
            opt.selected = true;
          }}
          sel.appendChild(opt);
        }});
        if (currentSelectedSheet) sel.value = currentSelectedSheet;
      }}

      const grid = document.getElementById('multi-sheets-checkboxes-grid');
      if (grid) {{
        grid.innerHTML = sheets.map(s => {{
          const isAlreadyLoaded = prelistaLoadedSheets.has(s);
          return `
            <label style="display: inline-flex; align-items: center; gap: 4px; background: rgba(15, 23, 42, 0.85); border: 1px solid ${{isAlreadyLoaded ? '#10B981' : 'rgba(56, 189, 248, 0.3)'}}; border-radius: 5px; padding: 3px 7px; cursor: pointer; user-select: none; font-size: 7.5pt; font-weight: 700; color: #FFFFFF;">
              <input type="checkbox" value="${{s}}" class="chk-multi-sheet" ${{isAlreadyLoaded ? 'checked' : ''}} style="accent-color: #0284C7; cursor: pointer;">
              <span>📦 ${{s}}</span>
            </label>
          `;
        }}).join('');
      }}
    }}

    async function consultarPrelistaPanel(sheetName, isForced, appendMode = false) {{
      const tbody = document.getElementById('prelista-panel-table-body');
      const sel = document.getElementById('prelista-panel-sheet-select');
      
      const targetSheet = (sheetName !== undefined && sheetName !== null && sheetName !== '')
        ? sheetName 
        : (sel && sel.value ? sel.value : '');
      
      if (tbody) {{
        const accionMsg = appendMode ? 'Sumando datos' : 'Consultando datos';
        const sheetMsg = targetSheet || 'la hoja más reciente';
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 15px; color: var(--text-muted);">${{accionMsg}} de ${{sheetMsg}} en Google Sheets...</td></tr>`;
      }}

      try {{
        const queryParts = [];
        if (targetSheet) queryParts.push(`sheet=${{encodeURIComponent(targetSheet)}}`);
        if (isForced) queryParts.push('force=1');
        const qs = queryParts.join('&');

        const res = await fetch(`/api/prelista${{qs ? '?' + qs : ''}}`);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        
        if (data && Array.isArray(data.productos)) {{
          const sheetNameLoaded = data.sheet || targetSheet || 'Hoja';
          
          if (Array.isArray(data.availableSheets) && data.availableSheets.length > 0) {{
            availableSheetsCache = data.availableSheets;
            actualizarSelectYCasillasHojas(data.availableSheets, sheetNameLoaded);
          }}

          const stats = fusionarProductosPrelista(data.productos, sheetNameLoaded, appendMode);

          updatePrelistaBadge();
          renderPrelistaPanelTable(prelistaProductsCache);

          if (appendMode) {{
            log(`[PRELISTA] ¡Hoja ${{sheetNameLoaded}} sumada! (${{stats.agregadosNuevos}} nuevos, ${{stats.cajasSumadas}} coincidencias). Total: ${{prelistaProductsCache.length}} productos.`, 'success');
          }} else {{
            log(`[PRELISTA] ¡Éxito! ${{data.productos.length}} productos en tránsito obtenidos de ${{sheetNameLoaded}}.`, 'success');
          }}
        }} else {{
          throw new Error(data.error || "Formato de datos no reconocido");
        }}
      }} catch(err) {{
        if (tbody) {{
          tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 15px; color: #FCA5A5;">Error al consultar Google Sheets: ${{err.message}}.</td></tr>`;
        }}
        log(`[PRELISTA ERROR] ${{err.message}}`, 'error');
      }}
    }}

    function cargarHojaSeleccionada(esSumar) {{
      const sel = document.getElementById('prelista-panel-sheet-select');
      const targetSheet = sel ? sel.value : '';
      if (!targetSheet) {{
        alert("Por favor selecciona una hoja de la lista.");
        return;
      }}
      if (!esSumar && prelistaProductsCache.length > 0 && prelistaLoadedSheets.size > 1) {{
        if (!confirm(`Tienes ${{prelistaProductsCache.length}} productos combinados de varias hojas.\\n\\n¿Deseas REEMPLAZAR todo y cargar únicamente "${{targetSheet}}"?\\n(Si deseas conservar los actuales y añadir esta hoja, haz clic en Cancelar y luego en "➕ Sumar Hoja")`)) {{
          return;
        }}
      }}
      consultarPrelistaPanel(targetSheet, true, esSumar);
    }}

    function toggleMultiSheetsPanel() {{
      const panel = document.getElementById('multi-sheets-panel');
      if (!panel) return;
      const isHidden = panel.style.display === 'none' || panel.style.display === '';
      panel.style.display = isHidden ? 'block' : 'none';
      if (isHidden && availableSheetsCache.length > 0) {{
        actualizarSelectYCasillasHojas(availableSheetsCache, null);
      }}
    }}

    function marcarRecientesHojas(count) {{
      const checks = document.querySelectorAll('.chk-multi-sheet');
      checks.forEach((chk, idx) => {{
        chk.checked = idx < count;
      }});
    }}

    function desmarcarTodasHojas() {{
      const checks = document.querySelectorAll('.chk-multi-sheet');
      checks.forEach(chk => {{ chk.checked = false; }});
    }}

    async function cargarHojasSeleccionadasMultiple(esSumar) {{
      const checks = Array.from(document.querySelectorAll('.chk-multi-sheet:checked')).map(c => c.value);
      if (checks.length === 0) {{
        alert("Marca al menos una casilla de hoja para cargar.");
        return;
      }}

      if (!esSumar && prelistaProductsCache.length > 0) {{
        if (!confirm(`¿Deseas reemplazar la prelista actual y cargar únicamente las ${{checks.length}} hojas marcadas (${{checks.join(', ')}})?`)) {{
          return;
        }}
      }}

      const tbody = document.getElementById('prelista-panel-table-body');
      if (tbody) {{
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 15px; color: #38BDF8;">Cargando y combinando ${{checks.length}} hojas: ${{checks.join(', ')}}...</td></tr>`;
      }}

      if (!esSumar) {{
        prelistaProductsCache = [];
        prelistaLoadedSheets.clear();
      }}

      let errores = [];
      let hojasCargadas = 0;

      for (const s of checks) {{
        try {{
          const res = await fetch(`/api/prelista?sheet=${{encodeURIComponent(s)}}&force=1`);
          if (!res.ok) throw new Error("HTTP " + res.status);
          const data = await res.json();
          if (data && Array.isArray(data.productos)) {{
            fusionarProductosPrelista(data.productos, s, true);
            hojasCargadas++;
            log(`[PRELISTA] Hoja ${{s}} procesada (${{data.productos.length}} productos).`, 'info');
          }} else {{
            errores.push(`${{s}}: ${{data.error || 'Sin productos'}}`);
          }}
        }} catch(err) {{
          errores.push(`${{s}}: ${{err.message}}`);
        }}
      }}

      updatePrelistaBadge();
      renderPrelistaPanelTable(prelistaProductsCache);

      const panel = document.getElementById('multi-sheets-panel');
      if (panel) panel.style.display = 'none';

      if (hojasCargadas > 0) {{
        log(`[PRELISTA ÉXITO] ¡${{hojasCargadas}} hojas combinadas! Total en prelista: ${{prelistaProductsCache.length}} productos.`, 'success');
        alert(`🎉 ¡ÉXITO!\\n\\nSe combinaron ${{hojasCargadas}} hojas:\\n• ${{checks.join('\\n• ')}}\\n\\nTotal acumulado: ${{prelistaProductsCache.length}} productos listos en la prelista.`);
      }}
      if (errores.length > 0) {{
        alert(`⚠️ Hubo advertencias al cargar algunas hojas:\\n${{errores.join('\\n')}}`);
      }}
    }}

    function quitarHojaDePrelista(sheetName) {{
      if (!confirm(`¿Deseas quitar los productos de "${{sheetName}}" de la prelista actual?`)) return;
      prelistaLoadedSheets.delete(sheetName);
      
      prelistaProductsCache = prelistaProductsCache.filter(p => {{
        if (!p._origenHojas || p._origenHojas.length === 0) return true;
        if (p._origenHojas.length === 1 && p._origenHojas[0] === sheetName) {{
          return false;
        }}
        p._origenHojas = p._origenHojas.filter(h => h !== sheetName);
        return true;
      }});

      renderPrelistaPanelTable(prelistaProductsCache);
      updatePrelistaBadge();
      log(`[PRELISTA] Hoja ${{sheetName}} retirada de la lista. Quedan ${{prelistaProductsCache.length}} productos.`, 'info');
    }}

    function renderPrelistaPanelTable(prods) {{
      const tbody = document.getElementById('prelista-panel-table-body');
      if (!tbody) return;

      if (!prods || prods.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 15px; color: var(--text-muted);">No hay productos en esta prelista. Agrega códigos arriba o sincroniza una hoja.</td></tr>`;
        return;
      }}

      tbody.innerHTML = prods.map(p => {{
        const origenBadge = (p._origenHojas && p._origenHojas.length > 0)
          ? `<div style="font-size: 6.5pt; color: #38BDF8; font-weight: 600; margin-top: 1px;">📦 ${{p._origenHojas.join(' + ')}}</div>`
          : '';
        return `
          <tr style="border-bottom: 1px solid rgba(255, 255, 255, 0.05);">
            <td style="padding: 4px 6px; font-weight: 800; color: #FFFFFF;">
              <div>${{p.codigo}}</div>
              ${{origenBadge}}
            </td>
            <td style="padding: 4px 6px; color: var(--text-main); max-width: 170px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${{p.detalle}}">${{p.detalle}}</td>
            <td style="padding: 4px 6px; text-align: center;">
              <input type="number" min="1" value="${{p.cajasVienen}}" onchange="cambiarCajasPrelista('${{p.codigo}}', this.value)" style="width: 52px; background: #090D16; border: 1px solid #38BDF8; color: #34D399; font-weight: bold; border-radius: 4px; padding: 2px 4px; text-align: center; font-size: 7.5pt;">
            </td>
            <td style="padding: 4px 6px; text-align: center; font-weight: 700; color: ${{p.stockReserva > 0 ? '#34D399' : '#F87171'}};">${{p.stockReserva}} cjs</td>
            <td style="padding: 4px 6px; color: var(--primary); font-weight: 700;">US$ ${{Number(p.precioRefUni || 0).toFixed(2)}}</td>
            <td style="padding: 4px 6px; text-align: center;">
              <button type="button" class="btn-chip danger" onclick="eliminarItemPrelista('${{p.codigo}}')" style="font-size: 7pt; padding: 2px 6px; cursor: pointer;" title="Quitar este producto de la prelista">
                ❌ Quitar
              </button>
            </td>
          </tr>
        `;
      }}).join('');
    }}

    function eliminarItemPrelista(codigo) {{
      const u = codigo.toUpperCase().trim();
      const idx = prelistaProductsCache.findIndex(p => p.codigo.toUpperCase().trim() === u);
      if (idx !== -1) {{
        prelistaProductsCache.splice(idx, 1);
        renderPrelistaPanelTable(prelistaProductsCache);
        updatePrelistaBadge();
        log(`[PRELISTA] Producto ${{u}} quitado de la lista previa.`, 'info');
      }}
    }}

    function cambiarCajasPrelista(codigo, val) {{
      const u = codigo.toUpperCase().trim();
      const item = prelistaProductsCache.find(p => p.codigo.toUpperCase().trim() === u);
      if (item) {{
        const n = parseInt(val) || 1;
        item.cajasVienen = n;
        item.stockReserva = n;
        renderPrelistaPanelTable(prelistaProductsCache);
      }}
    }}

    function agregarItemManualPrelista() {{
      const codeInput = document.getElementById('prelista-input-code');
      const boxInput = document.getElementById('prelista-input-boxes');
      const cod = (codeInput ? codeInput.value : '').trim().toUpperCase();
      const boxes = parseInt(boxInput ? boxInput.value : 10) || 10;
      if (!cod) {{
        alert("Escribe el código del producto que deseas agregar a la prelista.");
        if (codeInput) codeInput.focus();
        return;
      }}
      
      const invProd = allInventoryProducts.find(p => (p.cod || '').toUpperCase() === cod);
      const detalle = invProd ? `${{invProd.nombre}} ${{invProd.marca}}` : `PRODUCTO ${{cod}}`;
      const marca = invProd ? invProd.marca : 'GENERAL';
      const unidad = invProd ? (invProd.size || 'UNI') : 'UNI';
      const cantCaja = invProd ? (invProd.inner || 1) : 1;
      const precio = invProd ? (invProd.precioMayor || 0) : 0;

      const existing = prelistaProductsCache.find(p => p.codigo.toUpperCase() === cod);
      if (existing) {{
        existing.cajasVienen = boxes;
        existing.stockReserva = boxes;
      }} else {{
        prelistaProductsCache.unshift({{
          id: "P_MANUAL_" + Date.now(),
          codigo: cod,
          detalle: detalle,
          marca: marca,
          cajasVienen: boxes,
          cantPorCaja: cantCaja,
          unidad: unidad,
          precioMayor: precio,
          precioCaja: precio,
          precioEspecial: precio,
          precioRefUni: precio,
          precioCajaTotal: precio * cantCaja,
          stockReserva: boxes,
          agotado: false,
          _origenHojas: ['Manual']
        }});
      }}
      if (codeInput) codeInput.value = '';
      renderPrelistaPanelTable(prelistaProductsCache);
      updatePrelistaBadge();
      log(`[PRELISTA] Producto ${{cod}} (${{boxes}} cajas) agregado.`, 'success');
    }}

    function agregarMultiplesCodigosPrelista() {{
      const multiInput = document.getElementById('prelista-multi-input');
      const val = multiInput ? multiInput.value.trim() : '';
      if (!val) {{
        alert("Pega o escribe códigos con sus cajas (ej: DSM07-115 10, DZR110 5, DMY02-235 8).");
        return;
      }}
      const tokens = val.split(/[\\r\\n,]+/);
      let count = 0;
      tokens.forEach(tok => {{
        const parts = tok.trim().split(/\\s+/);
        if (parts.length > 0 && parts[0]) {{
          const cod = parts[0].toUpperCase().trim();
          const boxes = parts.length > 1 ? (parseInt(parts[1]) || 10) : 10;
          if (cod.length >= 2) {{
            const invProd = allInventoryProducts.find(p => (p.cod || '').toUpperCase() === cod);
            const detalle = invProd ? `${{invProd.nombre}} ${{invProd.marca}}` : `PRODUCTO ${{cod}}`;
            const marca = invProd ? invProd.marca : 'GENERAL';
            const cantCaja = invProd ? (invProd.inner || 1) : 1;
            const precio = invProd ? (invProd.precioMayor || 0) : 0;

            const existing = prelistaProductsCache.find(p => p.codigo.toUpperCase() === cod);
            if (existing) {{
              existing.cajasVienen = boxes;
              existing.stockReserva = boxes;
            }} else {{
              prelistaProductsCache.push({{
                id: "P_MANUAL_" + Math.random().toString(36).substr(2, 9),
                codigo: cod,
                detalle: detalle,
                marca: marca,
                cajasVienen: boxes,
                cantPorCaja: cantCaja,
                unidad: invProd ? (invProd.size || 'UNI') : 'UNI',
                precioMayor: precio,
                precioCaja: precio,
                precioEspecial: precio,
                precioRefUni: precio,
                precioCajaTotal: precio * cantCaja,
                stockReserva: boxes,
                agotado: false,
                _origenHojas: ['Manual']
              }});
            }}
            count++;
          }}
        }}
      }});
      if (multiInput) multiInput.value = '';
      renderPrelistaPanelTable(prelistaProductsCache);
      updatePrelistaBadge();
      log(`[PRELISTA] Se cargaron ${{count}} productos a la prelista.`, 'success');
      alert(`🎉 ¡Listo! Se cargaron ${{count}} producto(s) a la prelista.`);
    }}

    function vaciarPrelistaCache() {{
      if (confirm("¿Estás seguro de vaciar la lista de prelista?")) {{
        prelistaProductsCache = [];
        prelistaLoadedSheets.clear();
        renderPrelistaPanelTable(prelistaProductsCache);
        updatePrelistaBadge();
      }}
    }}

    function guardarYGrabarWebPrelista() {{
      if (!prelistaProductsCache || prelistaProductsCache.length === 0) {{
        alert("Primero sincroniza una o más hojas o agrega códigos para guardar.");
        return;
      }}
      const sel = document.getElementById('prelista-panel-sheet-select');
      const sheetNames = Array.from(prelistaLoadedSheets);
      const sheet = sheetNames.length > 0 
        ? sheetNames.join(' + ') 
        : (sel && sel.value ? sel.value : 'Prelista General');
      
      log(`[PRELISTA] Guardando y grabando ${{prelistaProductsCache.length}} productos (${{sheet}}) directamente en prelista.html...`, 'info');
      
      fetch('/api/prelista/guardar', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify({{
          sheet: sheet,
          productos: prelistaProductsCache,
          excluidos: []
        }})
      }})
      .then(r => r.json())
      .then(res => {{
        if (res.success) {{
          log(`[PRELISTA ÉXITO] ${{res.mensaje || 'Grabado con éxito.'}}`, 'success');
          alert(`🎉 ¡WEB PRELISTA GRABADA CON ÉXITO!\\n\\n` +
                `• Hojas combinadas: ${{sheet}}\\n` +
                `• Se grabaron ${{prelistaProductsCache.length}} productos en 'prelista.html'.\\n` +
                `• Ya no requiere conexión lenta ni selección de hoja para tus clientes.\\n\\n` +
                `🚀 SIGUIENTE PASO:\\n` +
                `Haz clic en '☁️ Publicar en Vercel Ahora' para subirla en vivo.`);
        }} else {{
          alert("Error al grabar: " + (res.error || "Desconocido"));
        }}
      }})
      .catch(err => {{
        alert("Error de conexión al guardar: " + err.message);
      }});
    }}

    function sincronizarFotosPrelistaExcel() {{
      log(">>> [PRELISTA FOTOS] Buscando y extrayendo fotos desde catalogos.xlsx...", "info");
      fetch('/api/prelista/sincronizar_fotos')
        .then(r => r.json())
        .then(res => {{
          if (res.success) {{
            log(`[PRELISTA FOTOS OK] ${{res.mensaje}}`, "success");
            alert(`🎉 ¡Fotos sincronizadas con éxito!\\n\\n${{res.mensaje}}\\nTotal productos con foto: ${{res.conFoto}}`);
            consultarPrelistaPanel();
          }} else {{
            log(`[PRELISTA FOTOS ERROR] ${{res.error}}`, "error");
            alert(`Aviso: ${{res.error}}`);
          }}
        }})
        .catch(err => {{
          log(`[PRELISTA FOTOS ERROR] ${{err.message}}`, "error");
          alert("Error al conectar: " + err.message);
        }});
    }}

    function publicarVercelDesdePrelista() {{
      publicarEnVercel();
    }}

    function cargarCodigosPrelistaAlGenerador() {{
      if (prelistaProductsCache.length === 0) {{
        alert("Primero sincroniza o agrega productos a la prelista.");
        return;
      }}
      let agregados = 0;
      prelistaProductsCache.forEach(p => {{
        const u = p.codigo.toUpperCase().trim();
        if (u) {{
          selectedCodesSet.add(u);
          agregados++;
        }}
      }});
      syncSetToTextarea();
      log(`[PRELISTA] Se copiaron ${{agregados}} códigos de prelista a la selección activa.`, 'success');
      alert(`🎉 ¡Listo! Se copiaron ${{agregados}} códigos de la prelista al Catálogo Principal.\\nTotal códigos listos: ${{selectedCodesSet.size}}`);
    }}

    // ─── 1. BÚSQUEDA PREDICTIVA ───
    function onSearchInput(query) {{
      const q = query.trim().toUpperCase();
      const clearBtn = document.getElementById('search-clear');
      const addAllBtn = document.getElementById('btn-add-all-search');
      clearBtn.style.display = q ? 'block' : 'none';

      if (!q) {{
        currentSearchResults = [];
        document.getElementById('search-results-count').innerText = `${{allInventoryProducts.length}} productos disponibles`;
        document.getElementById('search-results-list').innerHTML = '<div style="text-align: center; padding: 20px; color: var(--text-muted); font-size: 8.5pt;">Escribe para buscar por código, nombre o medida...</div>';
        addAllBtn.style.display = 'none';
        return;
      }}

      // Búsqueda multi-palabra
      const words = q.split(/\\s+/);
      currentSearchResults = allInventoryProducts.filter(p => {{
        const target = `${{p.cod}} ${{p.nombre}} ${{p.marca}} ${{p.categoria}} ${{p.size}}`.toUpperCase();
        return words.every(w => target.includes(w));
      }}).slice(0, 40); // Mostrar máximo 40 resultados por fluidez

      document.getElementById('search-results-count').innerText = `${{currentSearchResults.length}} resultado(s) encontrado(s)`;
      addAllBtn.style.display = currentSearchResults.length > 0 ? 'inline-block' : 'none';

      renderSearchResultsList(currentSearchResults);
    }}

    function renderSearchResultsList(items) {{
      const container = document.getElementById('search-results-list');
      if (items.length === 0) {{
        container.innerHTML = '<div style="text-align: center; padding: 20px; color: #EF4444; font-size: 8.5pt;">No se encontraron productos coincidentes.</div>';
        return;
      }}

      let html = '';
      items.forEach(p => {{
        const isAdded = selectedCodesSet.has(p.cod.toUpperCase());
        html += `
          <div class="product-item-row">
            <div class="product-item-info">
              <div style="display: flex; align-items: center; gap: 6px;">
                <span class="product-item-code">${{p.cod}}</span>
                <span style="font-size: 7pt; background: rgba(245,158,11,0.12); color: var(--primary); padding: 1px 5px; border-radius: 4px;">${{p.marca}}</span>
                ${{p.size ? `<span style="font-size: 7pt; color: #94A3B8;">${{p.size}}</span>` : ''}}
              </div>
              <div class="product-item-name" title="${{p.nombre}}">${{p.nombre}}</div>
            </div>
            <button class="btn-item-add ${{isAdded ? 'added' : ''}}" onclick="toggleProductCode('${{p.cod}}', this)" style="display: flex; align-items: center; gap: 4px;">
              ${{isAdded 
                ? '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg><span>Agregado</span>' 
                : '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg><span>Agregar</span>'}}
            </button>
          </div>
        `;
      }});
      container.innerHTML = html;
    }}

    function toggleProductCode(code, btn) {{
      const upper = code.toUpperCase();
      if (selectedCodesSet.has(upper)) {{
        selectedCodesSet.delete(upper);
        if (btn) {{
          btn.classList.remove('added');
          btn.innerHTML = '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg><span>Agregar</span>';
        }}
      }} else {{
        selectedCodesSet.add(upper);
        if (btn) {{
          btn.classList.add('added');
          btn.innerHTML = '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg><span>Agregado</span>';
        }}
      }}
      syncSetToTextarea();
    }}

    function addAllSearchResults() {{
      currentSearchResults.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
      syncSetToTextarea();
      renderSearchResultsList(currentSearchResults);
    }}

    function clearSearch() {{
      const input = document.getElementById('search-input');
      input.value = '';
      onSearchInput('');
      input.focus();
    }}

    function refreshVisibleSearchButtons() {{
      if (currentSearchResults.length > 0) {{
        renderSearchResultsList(currentSearchResults);
      }}
    }}

    // ─── 2. FILTRO POR MARCAS Y CATEGORÍAS ───
    function renderBrandsGrid() {{
      const container = document.getElementById('brands-grid-container');
      const brandKeys = Object.keys(inventoryBrands);
      if (brandKeys.length === 0) {{
        container.innerHTML = '<div style="grid-column: span 2; text-align: center; padding: 15px; color: var(--text-muted); font-size: 8.5pt;">No se detectaron marcas en el Excel.</div>';
        return;
      }}

      let html = '';
      brandKeys.forEach(brand => {{
        const count = inventoryBrands[brand];
        html += `
          <button class="brand-card-btn" onclick="selectBrandFilter('${{brand}}')">
            <span class="brand-card-name">${{brand}}</span>
            <span class="brand-card-count">${{count}}</span>
          </button>
        `;
      }});
      container.innerHTML = html;
    }}

    function selectBrandFilter(brand) {{
      activeBrandSelected = brand;
      const wrapper = document.getElementById('brand-categories-wrapper');
      wrapper.style.display = 'flex';
      document.getElementById('selected-brand-title').innerText = `${{brand}} (${{inventoryBrands[brand]}} productos)`;

      const catsObj = inventoryCategories[brand] || {{}};
      let chipsHtml = '';
      Object.keys(catsObj).forEach(cat => {{
        const cnt = catsObj[cat];
        chipsHtml += `
          <button class="btn-chip" onclick="addCategoryProducts('${{brand}}', '${{cat}}')" style="display: flex; align-items: center; gap: 4px;">
            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
            <span>${{cat}} (${{cnt}})</span>
          </button>
        `;
      }});
      document.getElementById('brand-categories-chips').innerHTML = chipsHtml;
    }}

    function addEntireBrand() {{
      if (!activeBrandSelected) return;
      const prods = allInventoryProducts.filter(p => p.marca.toUpperCase() === activeBrandSelected.toUpperCase());
      prods.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
      syncSetToTextarea();
      log(`[INFO] Se agregaron ${{prods.length}} productos de ${{activeBrandSelected}} a la selección.`);
    }}

    function addCategoryProducts(brand, cat) {{
      const prods = allInventoryProducts.filter(p => p.marca.toUpperCase() === brand.toUpperCase() && p.categoria.toUpperCase() === cat.toUpperCase());
      prods.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
      syncSetToTextarea();
      log(`[INFO] Se agregaron ${{prods.length}} productos de ${{brand}} > ${{cat}}.`);
    }}

    // ─── 3. GESTOR DE PLANTILLAS ───
    function getStoredTemplates() {{
      const defaultTemplates = [
        {{ name: '⭐ Catálogo Completo (Todo el Inventario)', codes: [], is_all: true }},
        {{ name: 'Top Ventas General', codes: ['ACC014', 'ACC017', 'ACT080', 'DSM02-100'] }},
        {{ name: 'Herramientas DongCheng & Crown', codes: ['DSM02-100', 'FF02-100', 'CT10128'] }}
      ];
      try {{
        const stored = localStorage.getItem('rivero_catalog_templates_v3');
        if (stored) {{
          const parsed = JSON.parse(stored);
          if (!parsed.some(t => t.is_all || t.name.includes('Todo el Inventario'))) {{
            parsed.unshift({{ name: '⭐ Catálogo Completo (Todo el Inventario)', codes: [], is_all: true }});
          }}
          return parsed;
        }}
        return defaultTemplates;
      }} catch(e) {{
        return defaultTemplates;
      }}
    }}

    function saveStoredTemplates(templates) {{
      localStorage.setItem('rivero_catalog_templates_v3', JSON.stringify(templates));
      renderTemplatesList();
    }}

    function renderTemplatesList() {{
      const container = document.getElementById('templates-list-container');
      const templates = getStoredTemplates();
      if (templates.length === 0) {{
        container.innerHTML = '<div style="text-align: center; padding: 15px; color: var(--text-muted); font-size: 8.5pt;">No tienes plantillas guardadas.</div>';
        return;
      }}

      let html = '';
      templates.forEach((t, idx) => {{
        const isAll = !!t.is_all;
        const countDisplay = isAll ? `${{allInventoryProducts.length || 'Todos los'}} producto(s)` : `${{t.codes.length}} producto(s)`;
        const itemStyle = isAll ? 'border: 1px solid rgba(245, 158, 11, 0.4); background: rgba(245, 158, 11, 0.08);' : '';
        const nameStyle = isAll ? 'font-weight: 800; color: var(--primary);' : '';

        html += `
          <div class="template-item" style="${{itemStyle}}">
            <div class="template-info">
              <span class="template-name" style="${{nameStyle}}">${{t.name}}</span>
              <span class="template-count">${{countDisplay}}</span>
            </div>
            <div class="template-actions">
              <button class="btn-chip" onclick="loadTemplateByIndex(${{idx}})" style="background: rgba(245, 158, 11, 0.2); color: var(--primary); border-color: rgba(245, 158, 11, 0.3); display: flex; align-items: center; gap: 4px;">
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 11 12 14 22 4"></polyline><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path></svg>
                <span>Cargar</span>
              </button>
              ${{isAll ? '' : `
              <button class="btn-chip danger" onclick="deleteTemplateByIndex(${{idx}})" title="Eliminar" style="display: flex; align-items: center; justify-content: center; padding: 4px 6px;">
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
              </button>
              `}}
            </div>
          </div>
        `;
      }});
      container.innerHTML = html;
    }}

    function guardarPlantillaDesdeInput() {{
      const input = document.getElementById('new-template-name');
      const name = input.value.trim();
      if (!name) {{
        alert("Escribe un nombre para la plantilla.");
        return;
      }}
      if (selectedCodesSet.size === 0) {{
        alert("Primero selecciona o pega códigos para guardar en la plantilla.");
        return;
      }}

      const templates = getStoredTemplates();
      templates.unshift({{ name: name, codes: Array.from(selectedCodesSet) }});
      saveStoredTemplates(templates);
      input.value = '';
      log(`[OK] Plantilla '${{name}}' guardada exitosamente con ${{selectedCodesSet.size}} productos.`);
    }}

    function cargarTodoElInventario() {{
      if (allInventoryProducts.length === 0) {{
        alert("El inventario aún se está cargando o no contiene productos.");
        return;
      }}
      selectedCodesSet.clear();
      allInventoryProducts.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
      syncSetToTextarea();
      switchSmartTab('manual');
      log(`[OK] ¡Cargados exitosamente los ${{allInventoryProducts.length}} productos del inventario al generador!`, 'success');
      alert(`🎉 ¡Listo! Se seleccionaron los ${{allInventoryProducts.length}} productos de todo el inventario.`);
    }}

    function guardarComoPlantillaPrompt() {{
      if (selectedCodesSet.size === 0) {{
        alert("No tienes códigos seleccionados para guardar.");
        return;
      }}
      const name = prompt("Escribe el nombre de la nueva plantilla:", "Catálogo " + new Date().toLocaleDateString('es-ES'));
      if (name && name.trim()) {{
        const templates = getStoredTemplates();
        templates.unshift({{ name: name.trim(), codes: Array.from(selectedCodesSet) }});
        saveStoredTemplates(templates);
        log(`[OK] Plantilla '${{name.trim()}}' guardada exitosamente.`);
      }}
    }}

    function loadTemplateByIndex(idx) {{
      const templates = getStoredTemplates();
      const t = templates[idx];
      if (t) {{
        selectedCodesSet.clear();
        if (t.is_all || t.name.includes('Todo el Inventario') || t.name.includes('Catálogo Completo')) {{
          allInventoryProducts.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
          syncSetToTextarea();
          log(`[OK] ¡Plantilla de Catálogo Completo cargada con los ${{allInventoryProducts.length}} productos del inventario!`, 'success');
        }} else if (t.codes) {{
          t.codes.forEach(c => selectedCodesSet.add(c.toUpperCase()));
          syncSetToTextarea();
          log(`[OK] Plantilla '${{t.name}}' cargada con ${{t.codes.length}} productos.`);
        }}
        switchSmartTab('manual');
      }}
    }}

    function deleteTemplateByIndex(idx) {{
      if (!confirm("¿Deseas eliminar esta plantilla guardada?")) return;
      const templates = getStoredTemplates();
      templates.splice(idx, 1);
      saveStoredTemplates(templates);
    }}

    // ─── 4. PROCESADOR DE PEDIDOS DE WHATSAPP PARA GOOGLE SHEETS ───
    let currentParsedOrder = {{ client: {{}}, items: [] }};

    function parseWhatsAppOrder(raw) {{
      const container = document.getElementById('order-parsed-result');
      if (!raw || !raw.trim()) {{
        if (container) container.style.display = 'none';
        currentParsedOrder = {{ client: {{}}, items: [] }};
        return;
      }}

      const lines = raw.split(/\\r?\\n/);
      let clientName = '';
      let clientAddress = '';
      let clientPhone = '';
      const items = [];

      lines.forEach(line => {{
        const cleanLine = line.replace(/[*_~`]/g, '').trim();
        const mName = cleanLine.match(/(?:Cliente|Nombre|👤)\\s*[:]?\\s*(.+)/i);
        if (mName && !clientName && mName[1].trim()) clientName = mName[1].trim();

        const mAddr = cleanLine.match(/(?:Dirección|Direccion|Zona|Destino|📍)\\s*[:]?\\s*(.+)/i);
        if (mAddr && !clientAddress && mAddr[1].trim()) clientAddress = mAddr[1].trim();

        const mPhone = cleanLine.match(/(?:Teléfono|Telefono|Celular|WhatsApp|Telf|📱)\\s*[:]?\\s*(.+)/i);
        if (mPhone && !clientPhone && mPhone[1].trim()) clientPhone = mPhone[1].trim();
      }});

      let currentItem = null;

      for (let i = 0; i < lines.length; i++) {{
        const rawLine = lines[i].trim();
        if (!rawLine) continue;

        const cleanLine = rawLine.replace(/[*_~`]/g, '').trim();

        // Si llegamos al resumen general o despedida, cerramos el último item y salimos del bucle
        if (cleanLine.match(/(?:RESUMEN GENERAL|Total Cajas cerradas|TOTAL MERCADER|Por favor confirmar|Muchas gracias)/i)) {{
          if (currentItem) {{
            items.push(currentItem);
            currentItem = null;
          }}
          break;
        }}

        // Detectar código de producto tipo [TIWLI201351] o 1. [CODIGO] Nombre
        const mCode = cleanLine.match(/\\[([A-Za-z0-9_\\-\\./]+)\\]/);
        if (mCode) {{
          if (currentItem) {{
            items.push(currentItem);
          }}
          const codeVal = mCode[1].trim().toUpperCase();
          const cleanName = cleanLine.replace(/^[0-9\\u20E3\\uFE0F\\.\\)\\-\\s▪•■]+/, '').replace(/\\[[A-Za-z0-9_\\-\\./]+\\]/, '').trim();
          const prodInfo = allInventoryProducts.find(p => p.cod.toUpperCase() === codeVal) || {{}};

          currentItem = {{
            code: codeVal,
            name: prodInfo.nombre || cleanName || codeVal,
            brand: prodInfo.marca || '',
            unitType: prodInfo.uni || 'UNI',
            cajas: 0,
            uni: 0,
            cantCaja: prodInfo.cant_caja || 1,
            totalUnits: 0,
            qty: 0
          }};
          continue;
        }}

        // Si tenemos un item activo, extraer cantidades y empaque de forma precisa
        if (currentItem) {{
          // 1. Empaque: "Empaque: 6 SET/caja | Marca: TOTAL"
          if (cleanLine.match(/(?:Empaque|Viene)/i)) {{
            const mPkg = cleanLine.match(/(?:Empaque|Viene)\\s*[:•\\-]?\\s*(\\d+)\\s*([A-Za-z]+)?\\/caja/i);
            if (mPkg) {{
              currentItem.cantCaja = parseInt(mPkg[1]) || currentItem.cantCaja || 1;
              if (mPkg[2]) currentItem.unitType = mPkg[2].toUpperCase();
            }}
            continue;
          }}

          // 2. Línea de Pedido: "Pedido: 1 Cajas + 2 SET → Total: 8 SET"
          if (cleanLine.match(/Pedido\\s*[:•\\-]/i)) {{
            // Extraer Cajas
            const mCajas = cleanLine.match(/(\\d+)\\s*(?:Cajas?|Cj|Cjas?)/i);
            if (mCajas) {{
              currentItem.cajas = parseInt(mCajas[1]);
            }}

            // Extraer Unidades Sueltas (después del signo +)
            const mPlusUni = cleanLine.match(/\\+\\s*(\\d+)\\s*([A-Za-z]+)?/i);
            if (mPlusUni) {{
              currentItem.uni = parseInt(mPlusUni[1]);
              if (mPlusUni[2] && !mPlusUni[2].toUpperCase().startsWith('TOT')) {{
                currentItem.unitType = mPlusUni[2].toUpperCase();
              }}
            }} else if (!mCajas) {{
              // Si no pidió cajas, solo pidió sueltas (ej: "Pedido: 2 SET")
              const mOnlyUni = cleanLine.match(/Pedido\\s*[:•\\-]?\\s*(\\d+)\\s*([A-Za-z]+)?/i);
              if (mOnlyUni) {{
                currentItem.uni = parseInt(mOnlyUni[1]);
                if (mOnlyUni[2] && !mOnlyUni[2].toUpperCase().startsWith('TOT')) {{
                  currentItem.unitType = mOnlyUni[2].toUpperCase();
                }}
              }}
            }}

            // Total de piezas explícito (ej: "Total: 8 SET")
            const mTotal = cleanLine.match(/Total\\s*[:=]\\s*(\\d+)/i);
            if (mTotal) {{
              currentItem.totalUnits = parseInt(mTotal[1]);
            }}
            continue;
          }}
        }} else if (cleanLine.match(/^[A-Za-z0-9_\\-\\./]+\\s+\\d+/)) {{
          // Formato simple por línea: CODIGO CANTIDAD
          const parts = cleanLine.split(/\\s+/);
          const simpleCode = parts[0].toUpperCase();
          const simpleQty = parseInt(parts[1]) || 1;
          const prodInfo = allInventoryProducts.find(p => p.cod.toUpperCase() === simpleCode) || {{}};
          const pkg = prodInfo.cant_caja || 1;
          items.push({{
            code: simpleCode,
            cajas: simpleQty,
            uni: 0,
            cantCaja: pkg,
            totalUnits: simpleQty * pkg,
            qty: simpleQty,
            unitType: prodInfo.uni || 'UNI',
            name: prodInfo.nombre || simpleCode,
            brand: prodInfo.marca || ''
          }});
        }}
      }}

      if (currentItem) {{
        items.push(currentItem);
      }}

      // Calcular totales por item si no vinieron explícitos
      items.forEach(it => {{
        if (!it.totalUnits || it.totalUnits === 0) {{
          it.totalUnits = (it.cajas * (it.cantCaja || 1)) + (it.uni || 0);
          if (it.totalUnits === 0) it.totalUnits = it.qty || 1;
        }}
      }});

      currentParsedOrder = {{
        client: {{ name: clientName, address: clientAddress, phone: clientPhone }},
        items: items
      }};

      if (items.length === 0 && !clientName) {{
        if (container) container.style.display = 'none';
        return;
      }}

      if (container) container.style.display = 'flex';
      
      const infoParts = [];
      if (clientName) infoParts.push(`👤 <strong>${{clientName}}</strong>`);
      if (clientAddress) infoParts.push(`📍 ${{clientAddress}}`);
      if (clientPhone) infoParts.push(`📱 ${{clientPhone}}`);
      const clientInfoEl = document.getElementById('order-parsed-client-info');
      if (clientInfoEl) {{
        clientInfoEl.innerHTML = infoParts.length > 0 ? infoParts.join(' • ') : 'Pedido recibido';
      }}

      const tbody = document.getElementById('order-parsed-table-body');
      if (tbody) {{
        let tbodyHtml = '';
        items.forEach(it => {{
          if (it.cajas > 0 && it.uni > 0) {{
            // Fila 1: Cajas cerradas
            tbodyHtml += `
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
                <td style="padding: 3px 6px; font-weight: 800; color: #D97706;">📦 ${{it.cajas}} Cj</td>
                <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
                <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
                <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
              </tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.08); background: rgba(37, 211, 102, 0.03);">
                <td style="padding: 3px 6px; font-weight: 800; color: #22C55E;">↳ ${{it.uni}} Uni</td>
                <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
                <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
                <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
              </tr>
            `;
          }} else if (it.cajas > 0) {{
            tbodyHtml += `
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
                <td style="padding: 3px 6px; font-weight: 800; color: #D97706;">📦 ${{it.cajas}} Cj</td>
                <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
                <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
                <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
              </tr>
            `;
          }} else if (it.uni > 0) {{
            tbodyHtml += `
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
                <td style="padding: 3px 6px; font-weight: 800; color: #22C55E;">${{it.uni}} Uni</td>
                <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
                <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
                <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
              </tr>
            `;
          }} else {{
            tbodyHtml += `
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
                <td style="padding: 3px 6px; font-weight: 800; color: #25D366;">${{it.qty || 1}}</td>
                <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
                <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
                <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
              </tr>
            `;
          }}
        }});
        tbody.innerHTML = tbodyHtml;
      }}
    }}

    function copiarInfoCliente() {{
      const c = currentParsedOrder.client;
      const text = `${{c.name || ''}}\\t${{c.address || ''}}\\t${{c.phone || ''}}`;
      navigator.clipboard.writeText(text).then(() => {{
        log(`[OK] Datos del cliente copiados: ${{c.name || ''}} (${{c.address || ''}})`);
        alert(`¡Datos del cliente copiados!\\n\\nNombre: ${{c.name || ''}}\\nDirección: ${{c.address || ''}}\\nTeléfono: ${{c.phone || ''}}`);
      }});
    }}

    function copiarFormatoCantidad() {{
      if (!currentParsedOrder.items || currentParsedOrder.items.length === 0) {{
        alert("No hay productos detectados en el texto para copiar.");
        return;
      }}

      // FOTO 2: Formato IR01XX (4 Columnas):
      // CANTIDAD (A) \t UN/MED (B) \t DETALLE (C) \t CODIGO (D)
      const rows = currentParsedOrder.items.map(it => {{
        const qtyVal = it.totalUnits || (it.cajas * (it.cantCaja || 1) + it.uni) || it.qty || 1;
        return `${{qtyVal}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`;
      }});

      const tsv = rows.join('\\n');
      navigator.clipboard.writeText(tsv).then(() => {{
        alert(`¡${{currentParsedOrder.items.length}} productos copiados para Formato por Cantidad (4 Columnas)!\\n\\n1. Ve a tu Google Sheets (Formato IR01XX).\\n2. Clic en celda A5 (CANTIDAD).\\n3. Presiona Ctrl + V para pegar.`);
        log(`[OK] ${{currentParsedOrder.items.length}} filas copiadas para Formato por Cantidad (4 col: CANTIDAD | UN/MED | DETALLE | CODIGO).`);
      }}).catch(() => {{
        const ta = document.createElement("textarea");
        ta.value = tsv;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        alert(`¡${{currentParsedOrder.items.length}} productos copiados!\\n\\nVe a tu Google Sheets en celda A5 y presiona Ctrl + V.`);
      }});
    }}

    function copiarFormatoCajas() {{
      if (!currentParsedOrder.items || currentParsedOrder.items.length === 0) {{
        alert("No hay productos detectados en el texto para copiar.");
        return;
      }}

      // FOTO 1: Formato IR01ML (5 Columnas en filas separadas para Cajas y Unidades sueltas):
      // CANT. CAJAS (A) \t CANT. UNI. (B) \t UN/MED (C) \t DETALLE (D) \t CODIGO (E)
      const rows = [];
      currentParsedOrder.items.forEach(it => {{
        const pkg = it.cantCaja || 1;
        let unitsInBoxes = it.cajas * pkg;
        if (it.totalUnits && it.totalUnits > it.uni && it.cajas > 0) {{
          unitsInBoxes = it.totalUnits - it.uni;
        }}

        if (it.cajas > 0 && it.uni > 0) {{
          // Fila 1: Caja cerrada con la cantidad de unidades que trae la caja (para el subtotal del Excel)
          rows.push(`${{it.cajas}}\\t${{unitsInBoxes}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`);
          // Fila 2: Unidades sueltas debajo (Cajas queda vacío, Uni lleva la cantidad suelta)
          rows.push(`\\t${{it.uni}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`);
        }} else if (it.cajas > 0) {{
          rows.push(`${{it.cajas}}\\t${{unitsInBoxes}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`);
        }} else if (it.uni > 0) {{
          rows.push(`\\t${{it.uni}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`);
        }} else {{
          rows.push(`${{it.qty || 1}}\\t${{it.qty || 1}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`);
        }}
      }});

      const tsv = rows.join('\\n');
      navigator.clipboard.writeText(tsv).then(() => {{
        alert(`¡${{currentParsedOrder.items.length}} productos copiados para Formato por Cajas (5 Columnas)!\\n\\n1. Ve a tu Google Sheets (Formato IR01ML).\\n2. Clic en celda A5 (CANT. CAJAS).\\n3. Presiona Ctrl + V para pegar.`);
        log(`[OK] ${{currentParsedOrder.items.length}} filas copiadas para Formato por Cajas (5 col: CANT. CAJAS | CANT. UNI. | UN/MED | DETALLE | CODIGO).`);
      }}).catch(() => {{
        const ta = document.createElement("textarea");
        ta.value = tsv;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        alert(`¡${{currentParsedOrder.items.length}} productos copiados!\\n\\nVe a tu Google Sheets en celda A5 y presiona Ctrl + V.`);
      }});
    }}

    function cargarPedidoAlGenerador() {{
      if (!currentParsedOrder.items || currentParsedOrder.items.length === 0) return;
      selectedCodesSet.clear();
      currentParsedOrder.items.forEach(it => selectedCodesSet.add(it.code.toUpperCase()));
      syncSetToTextarea();
      switchSmartTab('manual');
      log(`[OK] ${{currentParsedOrder.items.length}} códigos del pedido cargados al generador.`);
    }}

    function limpiarSeleccion() {{
      selectedCodesSet.clear();
      textareaCodes.value = '';
      updateActiveCounter();
      refreshVisibleSearchButtons();
      log("[INFO] Selección de productos vaciada.");
    }}

    function copiarListaSeleccionada() {{
      const text = Array.from(selectedCodesSet).join('\\n');
      if (!text) {{
        alert("No hay códigos para copiar.");
        return;
      }}
      navigator.clipboard.writeText(text).then(() => {{
        alert(`¡Copiados ${{selectedCodesSet.size}} códigos al portapapeles!`);
      }});
    }}

    // ─── GENERACIÓN Y VISTA PREVIA ───
    function getHtmlUrl(device) {{
      if (device === 'mobile' && hasMobileFile) return 'catalogos_mobile.html';
      if (device === 'desktop' && hasDesktopFile) return 'catalogos_desktop.html';
      return 'catalogos.html';
    }}

    function updateDownloadHtmlLink() {{
      const btnDownloadHtml = document.getElementById('btn-download-html');
      const fileUrl = getHtmlUrl(currentDevice);
      const isAvailable = (fileUrl === 'catalogos.html') ? {preview_available} : (currentDevice === 'desktop' ? hasDesktopFile : hasMobileFile);
      
      if (isAvailable) {{
        const finalUrl = fileUrl + '?t=' + Date.now();
        btnDownloadHtml.href = finalUrl;
        btnDownloadHtml.setAttribute('download', fileUrl);
        btnDownloadHtml.style.pointerEvents = 'auto';
        btnDownloadHtml.style.opacity = '1';
      }} else {{
        btnDownloadHtml.href = '#';
        btnDownloadHtml.style.pointerEvents = 'none';
        btnDownloadHtml.style.opacity = '0.5';
      }}
    }}
    
    function log(message, type = '') {{
      let styleClass = '';
      if (type === 'success' || message.includes('[OK]') || message.includes('exitosamente')) styleClass = 'class="log-success"';
      else if (type === 'error' || message.includes('[ERROR]') || message.includes('Traceback')) styleClass = 'class="log-error"';
      else if (message.includes('[NUBE]') || message.includes('>>>')) styleClass = 'class="log-info"';
      else if (message.includes('[AVISO]') || message.includes('⚠️')) styleClass = 'class="log-warning"';
      
      consoleDiv.innerHTML += `<div class="log-line"><span ${{styleClass}}>${{message}}</span></div>`;
      consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }}
    
    function setDevice(device) {{
      currentDevice = device;
      document.getElementById('btn-device-desktop').classList.toggle('active', device === 'desktop');
      document.getElementById('btn-device-mobile').classList.toggle('active', device === 'mobile');
      
      const fileUrl = getHtmlUrl(device);
      const isAvailable = (fileUrl === 'catalogos.html') ? {preview_available} : (device === 'desktop' ? hasDesktopFile : hasMobileFile);
      if (isAvailable) {{
        iframe.src = fileUrl + '?t=' + Date.now();
      }}
      
      iframe.className = (device === 'mobile') ? 'view-mobile' : '';
      updateDownloadHtmlLink();
    }}
    
    function iniciarGeneracion() {{
      const codes = textareaCodes.value;
      const sync = document.getElementById('sync').checked;
      const layout = document.getElementById('layout')?.value || 'desktop';
      const forceImages = document.getElementById('force_images').checked;
      const whatsapp = document.getElementById('whatsapp').value;
      const filterStock = document.getElementById('filter_stock')?.checked || false;
      
      consoleDiv.innerHTML = '';
      log(">>> Iniciando petición al backend...");
      if (filterStock) {{
        log(">>> [FILTRO ACTIVO] Se filtrarán los productos agotados según el stock en vivo de Google Sheets.");
      }}
      
      btnRun.disabled = true;
      btnText.innerText = "Procesando...";
      btnSpinner.style.display = "block";
      
      const url = `/generar?codes=${{encodeURIComponent(codes)}}&sync=${{sync}}&layout=${{layout}}&force_images=${{forceImages}}&whatsapp=${{encodeURIComponent(whatsapp)}}&filter_stock=${{filterStock}}`;
      const source = new EventSource(url);
      
      source.onmessage = function(event) {{
        const msg = event.data;
        if (msg.startsWith("EVENT_SUCCESS:")) {{
          log(msg.replace("EVENT_SUCCESS:", ""), "success");
          source.close();
          finalizarGeneracion(true);
        }} else if (msg.startsWith("EVENT_ERROR:")) {{
          log(msg.replace("EVENT_ERROR:", ""), "error");
          source.close();
          finalizarGeneracion(false);
        }} else {{
          log(msg);
        }}
      }};
      
      source.onerror = function() {{
        log("[ERROR] La conexión con el servidor se interrumpió de forma inesperada.", "error");
        source.close();
        finalizarGeneracion(false);
      }};
    }}
    
    function finalizarGeneracion(success) {{
      btnRun.disabled = false;
      btnText.innerText = "Empezar Generación";
      btnSpinner.style.display = "none";
      document.getElementById('force_images').checked = false;
      
      if (success) {{
        hasDesktopFile = true;
        hasMobileFile = true;
        
        noPreview.style.display = 'none';
        iframe.style.display = 'block';
        setDevice(currentDevice);
        document.getElementById('btn-full-preview').disabled = false;

        // Asegurar que el iframe tenga la clase de administrador
        onIframeLoaded();

        // Guardar códigos en memoria local para rápida recuperación
        try {{
          const list = Array.from(selectedCodesSet).filter(c => c && !c.includes('$') && !c.includes('{') && !c.includes('}'));
          if (list.length > 0) {{
            localStorage.setItem('rivero_last_generated_codes', JSON.stringify(list));
            ultimosCodigosGeneradosCache = list;
            const infoLbl = document.getElementById('last-catalog-info-label');
            if (infoLbl) infoLbl.innerText = `(${{list.length}} en historial)`;
            const btnText = document.getElementById('btn-load-last-text');
            if (btnText) btnText.innerText = `🔄 Cargar Último (${{list.length}})`;
          }}
        }} catch(e) {{}}
      }}
    }}
    
    async function testStockConnection(e) {{
      const btn = e ? e.target : null;
      const originalText = btn ? btn.innerText : 'Probar API';
      if (btn) btn.innerText = 'Conectando...';
      log(">>> [STOCK] Probando conexión con Supabase / Google Drive...");
      try {{
        const res = await fetch("/api/stock?force=1&_t=" + Date.now(), {{ cache: 'no-store' }});
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        const total = Object.keys(data).length;
        log(`>>> [STOCK OK] ¡Conexión exitosa! Se encontraron ${{total}} productos sincronizados en tiempo real.`, "success");
        if (btn) btn.innerText = `OK (${{total}} items)`;
      }} catch (err) {{
        log(">>> [STOCK ERROR] " + err.message, "error");
        if (btn) btn.innerText = 'Error';
      }}
      if (btn) {{
        setTimeout(() => {{ btn.innerText = originalText; }}, 4000);
      }}
    }}

    // Iniciar carga del inventario y estado
    cargarInventarioAPI();
    consultarUltimosCodigosSilencioso();
    function publicarEnVercel() {{
      if (!confirm("¿Deseas publicar y actualizar el catálogo online en Vercel ahora mismo?")) return;
      
      const btnPublish = document.getElementById('btn-publish-vercel');
      btnPublish.disabled = true;
      btnPublish.style.opacity = '0.6';
      
      log(">>> [VERCEL] Conectando con GitHub y Vercel...");
      
      const evtSource = new EventSource('/publicar_vercel');
      evtSource.onmessage = function(e) {{
        if (e.data.startsWith('EVENT_SUCCESS')) {{
          evtSource.close();
          btnPublish.disabled = false;
          btnPublish.style.opacity = '1';
          log(">>> [VERCEL] ¡PUBLICACIÓN EXITOSA! Tu catálogo ya está en la nube.", 'success');
          alert("🎉 ¡Catálogo publicado con éxito en Vercel!\\nEn unos 15 segundos estará disponible en vivo en tu enlace web.");
        }} else if (e.data.startsWith('EVENT_ERROR')) {{
          evtSource.close();
          btnPublish.disabled = false;
          btnPublish.style.opacity = '1';
          log(">>> [VERCEL] Falló la publicación.", 'error');
          alert("❌ Ocurrió un inconveniente al subir a Vercel/GitHub. Revisa la consola del panel.");
        }} else {{
          log(e.data);
        }}
      }};
      evtSource.onerror = function() {{
        evtSource.close();
        btnPublish.disabled = false;
        btnPublish.style.opacity = '1';
      }};
    }}

    // Escuchar eliminaciones en vivo y depuraciones desde la vista previa interactiva
    window.addEventListener('message', function(event) {{
      if (event.data && event.data.type === 'REMOVE_CATALOG_ITEM') {{
        const code = (event.data.code || '').toUpperCase();
        if (selectedCodesSet.size === 0 && allInventoryProducts.length > 0) {{
          allInventoryProducts.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
        }}
        if (selectedCodesSet.has(code)) {{
          selectedCodesSet.delete(code);
          syncSetToTextarea();
          log(`[EDITOR EN VIVO] Producto '${{code}}' quitado de la selección directamente desde el catálogo.`, 'success');
        }}
      }} else if (event.data && event.data.type === 'PURGED_OUT_OF_STOCK_RESULT') {{
        const purged = event.data.codes || [];
        if (purged.length > 0) {{
          if (selectedCodesSet.size === 0 && allInventoryProducts.length > 0) {{
            allInventoryProducts.forEach(p => selectedCodesSet.add(p.cod.toUpperCase()));
          }}
          purged.forEach(c => selectedCodesSet.delete(c.toUpperCase()));
          syncSetToTextarea();
          log(`[DEPURACIÓN] Se quitaron ${{purged.length}} productos agotados del catálogo visual.`, 'success');
        }}
      }}
    }});

    updateDownloadHtmlLink();
  </script>
</body>
</html>
"""
            self.wfile.write(html_ui.encode('utf-8'))
            return
            
        else:
            self.send_error(404, "Not found")

def obtener_ip_local():
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # Intento de conexión simulado para determinar la interfaz activa
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except Exception:
            ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def iniciar_tunel_ssh(port):
    import subprocess
    import platform
    if platform.system() == "Windows":
        # Usamos localhost.run con IP directa 127.0.0.1, que resolvió el problema del usuario
        comando = f"ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=NUL -R 80:127.0.0.1:{port} nokey@localhost.run"
        try:
            print("[TÚNEL] Iniciando túnel web público en una ventana nueva...")
            # Popen con start cmd /k abre una nueva consola independiente de Windows que no bloquea este script
            subprocess.Popen(f'start cmd /k "title Tunel de Red Local - Importadora Rivero && echo ======================================================= && echo   TÚNEL DE INTERNET ACTIVO PARA COLABORADORES && echo ======================================================= && echo Copia el enlace que termina en \'.lhr.life\' que aparezca abajo && echo y mandaselo a tus companeros de la oficina por WhatsApp. && echo ======================================================= && echo. && {comando}"', shell=True)
        except Exception as e:
            print(f"[TÚNEL] [AVISO] No se pudo iniciar el túnel automático: {e}")

def start_server():
    # Configurar para que el servidor maneje solicitudes concurrentes (para no trabar SSE)
    class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
        daemon_threads = True
        allow_reuse_address = True

    initial_port = int(os.environ.get("PORT", PORT))
    httpd = None
    port = initial_port
    
    # Intentar en el puerto configurado o buscar el siguiente libre si está ocupado
    try:
        for p in range(initial_port, initial_port + 10):
            try:
                server_address = ('', p)
                httpd = ThreadingHTTPServer(server_address, CatalogWebHandler)
                port = p
                break
            except OSError as ex_bind:
                if p == initial_port + 9:
                    raise ex_bind
                continue
        ip_local = obtener_ip_local()
        print(f"\n=======================================================")
        print(f"  SERVIDOR DEL GENERADOR DE CATÁLOGOS CORRIENDO (OFICINA)")
        print(f"  Acceso Local (Tú):     http://localhost:{port}")
        if ip_local and ip_local != '127.0.0.1':
            print(f"  Acceso Oficina (Compas): http://{ip_local}:{port}")
        print(f"=======================================================\n")
        
        # Iniciar túnel de internet automático para colaboradores en red local / Wi-Fi
        iniciar_tunel_ssh(port)
        
        # Abrir navegador automáticamente solo localmente
        webbrowser.open(f"http://localhost:{port}")
        
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nApagando servidor web...")
            httpd.server_close()
    except Exception as e:
        print(f"[ERROR] No se pudo iniciar el servidor web en el puerto {port}: {e}")
        print("Asegúrate de que no haya otra instancia corriendo.")
        input("\nPresiona Enter para salir...")

if __name__ == "__main__":
    try:
        sincronizar_imagenes_prelista_existente()
    except Exception as e_sync_init:
        print(f"[PRELISTA AVISO] Error al sincronizar imágenes iniciales: {e_sync_init}")
    start_server()
