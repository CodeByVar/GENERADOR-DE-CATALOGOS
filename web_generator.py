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
import urllib.parse
import urllib.request
import ssl
import json
import subprocess
import generar_catalogo
from datetime import date

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

    def do_POST(self):
        # Desactivado por seguridad en red local (intranet) para evitar que clientes
        # externos abran programas o carpetas en la máquina principal.
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
            sync_raw = query_params.get('sync', ['true'])[0]
            layout_raw = query_params.get('layout', ['desktop'])[0]
            force_images_raw = query_params.get('force_images', ['false'])[0]
            whatsapp_raw = query_params.get('whatsapp', [''])[0]
            
            descargar_nube = (sync_raw.lower() == 'true')
            forzar_imagenes = (force_images_raw.lower() == 'true')
            if codes_raw and codes_raw.strip():
                import re
                tokens = re.split(r'[\r\n,;\t]+', codes_raw)
                codigos_custom = [t.strip().strip('"\'') for t in tokens if t.strip().strip('"\'')]
            else:
                codigos_custom = None
            
            writer = SSEStdoutWriter(self)
            
            writer.write(">>> Iniciando generación de catálogo desde el servidor web...\n")
            if codigos_custom:
                writer.write(f">>> Códigos recibidos: {len(codigos_custom)} ítems.\n")
            else:
                writer.write(">>> Leyendo códigos desde la hoja Vista_Catalogo en Excel...\n")
                
            with RedirectStdout(writer):
                try:
                    import importlib
                    importlib.reload(generar_catalogo)
                    generar_catalogo.generar(descargar_nube=descargar_nube, codigos_custom=codigos_custom, layout=layout_raw, forzar_imagenes=forzar_imagenes, whatsapp_phone=whatsapp_raw)
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
                
                writer.write(">>> [VERCEL] Registrando archivos en Git...\n")
                subprocess.run(["git", "add", "index.html", "catalogos.html", "catalogos_desktop.html", "catalogos_mobile.html", "vercel.json", "generar_catalogo.py", "web_generator.py", "Publicar_en_Vercel.bat"], capture_output=True)
                subprocess.run(["git", "add", "-u"], capture_output=True)
                
                writer.write(">>> [VERCEL] Creando punto de actualización en historial...\n")
                subprocess.run(["git", "commit", "-m", "Actualizacion del catalogo online para clientes"], capture_output=True)
                
                writer.write(">>> [VERCEL] Subiendo cambios a GitHub / Vercel en la nube...\n")
                res = subprocess.run(["git", "push", "origin", "main"], capture_output=True, text=True)
                if res.returncode == 0:
                    writer.write(">>> [VERCEL] [OK] Subida completada con éxito!\n")
                    writer.write(">>> [VERCEL] Vercel se está actualizando en vivo en tu enlace web.\n")
                    writer.write("EVENT_SUCCESS: Catálogo publicado con éxito en Vercel.\n")
                else:
                    err_msg = res.stderr or res.stdout
                    writer.write(f">>> [VERCEL AVISO] {err_msg.strip()}\n")
                    writer.write("EVENT_ERROR: Error al subir cambios a GitHub / Vercel.\n")
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
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            try:
                stock_url = getattr(generar_catalogo, 'URL_STOCK_API', "https://script.google.com/macros/s/AKfycbxrXCYxH9JX-uO2rw5Wg7XY5PnbKso50ugmpkTnrPacwy12GoMpxn-AvlbRZ_m0a9k45w/exec")
                req = urllib.request.Request(stock_url, headers={'User-Agent': 'Mozilla/5.0'})
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                with urllib.request.urlopen(req, context=ctx, timeout=90) as response:
                    content = response.read()
                    self.wfile.write(content)
            except Exception as e:
                self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))
            return

        # 2. Servir el PDF de catálogo
        elif parsed_url.path == "/catalogos.pdf":
            self.serve_file("catalogos.pdf", "application/pdf")
            return
            
        # 3. Servir el catálogo en HTML
        elif parsed_url.path in ("/catalogos.html", "/catalogos_desktop.html", "/catalogos_mobile.html"):
            filename = parsed_url.path[1:] # quitar la barra inicial
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
      </div>

      <!-- Barra de Estado de Selección Activa -->
      <div class="active-banner">
        <div class="active-badge">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>
          <span id="badge-count-text">0 códigos listos</span>
        </div>
        <div class="active-actions">
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
                  <th style="padding: 3px 6px;">CAJAS</th>
                  <th style="padding: 3px 6px;">UNI</th>
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

      <!-- Ajustes de Generación -->
      <div class="toggle-row">
        <span class="label-text" style="font-weight: 600; font-size: 8.5pt;">Sincronizar base de datos desde Google Drive</span>
        <label class="switch">
          <input type="checkbox" id="sync" checked>
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
      
      <div class="toggle-row">
        <span class="label-text" style="font-weight: 600; font-size: 8.5pt;">Diseño del Folleto</span>
        <select id="layout" style="background: var(--bg-console); border: 1px solid var(--border-panel); color: var(--text-main); padding: 5px 8px; border-radius: 6px; font-family: inherit; font-size: 8.5pt; outline: none; cursor: pointer; font-weight: 700;">
          <option value="desktop" selected>A4 Impresora (2 Columnas)</option>
          <option value="mobile">Celular / WhatsApp (1 Columna)</option>
        </select>
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
        
        <iframe id="preview-iframe" src="{"catalogos_desktop.html" if desktop_available == "true" else ("catalogos.html" if preview_available == "true" else "about:blank")}" style="display: {'block' if preview_available == 'true' else 'none'};"></iframe>
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
      const tabs = ['manual', 'search', 'brands', 'templates', 'order'];
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
        const mName = line.match(/(?:Cliente|Nombre|Cliente:)\\s*[:\\*]?\\s*([^\\n\\*_]+)/i);
        if (mName && !clientName) clientName = mName[1].trim();

        const mAddr = line.match(/(?:Dirección|Direccion|Zona|Destino)\\s*[:\\*]?\\s*([^\\n\\*_]+)/i);
        if (mAddr && !clientAddress) clientAddress = mAddr[1].trim();

        const mPhone = line.match(/(?:Teléfono|Telefono|Celular|WhatsApp|Telf)\\s*[:\\*]?\\s*([^\\n\\*_]+)/i);
        if (mPhone && !clientPhone) clientPhone = mPhone[1].trim();
      }});

      let currentItemCode = null;
      let currentItemName = '';

      for (let i = 0; i < lines.length; i++) {{
        const line = lines[i].trim();
        if (!line) continue;

        const mCode = line.match(/\\[([A-Za-z0-9_\\-\\./]+)\\]/);
        if (mCode) {{
          currentItemCode = mCode[1].trim().toUpperCase();
          currentItemName = line.replace(/^[0-9\\u20E3\\uFE0F\\.\\)\\-\\s]+/, '').replace(/\\[[A-Za-z0-9_\\-\\./]+\\]/, '').trim();
        }}

        const mCajas = line.match(/(\\d+)\\s*Caja/i);
        const mUni = line.match(/(\\d+)\\s*Unid/i);
        const mGeneral = line.match(/(?:Cantidad|Cant|Pedir)\\s*[:•\\-]?\\s*\\*?(\\d+)\\*?/i);

        if ((mCajas || mUni || mGeneral) && currentItemCode) {{
          const cajas = mCajas ? parseInt(mCajas[1]) : 0;
          const uni = mUni ? parseInt(mUni[1]) : 0;
          const generalQty = mGeneral ? parseInt(mGeneral[1]) : (cajas || uni || 1);
          
          const prodInfo = allInventoryProducts.find(p => p.cod.toUpperCase() === currentItemCode) || {{}};
          items.push({{
            code: currentItemCode,
            cajas: cajas,
            uni: uni,
            qty: generalQty,
            unitType: prodInfo.uni || 'UNI',
            name: prodInfo.nombre || currentItemName || currentItemCode,
            brand: prodInfo.marca || ''
          }});
          currentItemCode = null;
          currentItemName = '';
        }} else if (!mCode && line.match(/^[A-Za-z0-9_\\-\\./]+\\s+\\d+/)) {{
          const parts = line.split(/\\s+/);
          const simpleCode = parts[0].toUpperCase();
          const simpleQty = parseInt(parts[1]) || 1;
          const prodInfo = allInventoryProducts.find(p => p.cod.toUpperCase() === simpleCode) || {{}};
          items.push({{
            code: simpleCode,
            cajas: simpleQty,
            uni: 0,
            qty: simpleQty,
            unitType: prodInfo.uni || 'UNI',
            name: prodInfo.nombre || simpleCode,
            brand: prodInfo.marca || ''
          }});
        }}
      }}

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
      if (clientName) infoParts.push(`<strong>${{clientName}}</strong>`);
      if (clientAddress) infoParts.push(`📍 ${{clientAddress}}`);
      if (clientPhone) infoParts.push(`📞 ${{clientPhone}}`);
      const clientInfoEl = document.getElementById('order-parsed-client-info');
      if (clientInfoEl) {{
        clientInfoEl.innerHTML = infoParts.length > 0 ? infoParts.join(' • ') : 'Pedido recibido';
      }}

      const tbody = document.getElementById('order-parsed-table-body');
      if (tbody) {{
        let tbodyHtml = '';
        items.forEach(it => {{
          let cantDisplay = [];
          if (it.cajas > 0) cantDisplay.push(it.cajas + ' Cj');
          if (it.uni > 0) cantDisplay.push(it.uni + ' Uni');
          const finalCant = cantDisplay.length > 0 ? cantDisplay.join(' + ') : it.qty;

          tbodyHtml += `
            <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
              <td style="padding: 3px 6px; font-weight: 800; color: #25D366;">${{finalCant}}</td>
              <td style="padding: 3px 6px; color: var(--text-muted);">${{it.unitType}}</td>
              <td style="padding: 3px 6px; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${{it.name}}">${{it.name}}</td>
              <td style="padding: 3px 6px; font-family: 'JetBrains Mono', monospace; color: var(--primary); font-weight: 700;">${{it.code}}</td>
            </tr>
          `;
        }});
        tbody.innerHTML = tbodyHtml;
      }}
    }}

    function copiarInfoCliente() {{
      const c = currentParsedOrder.client;
      const text = `CLIENTE: ${{c.name || ''}}\\tDIRECCION: ${{c.address || ''}}\\tTELF: ${{c.phone || ''}}`;
      navigator.clipboard.writeText(text).then(() => {{
        log(`[OK] Datos del cliente copiados: ${{c.name || ''}} (${{c.address || ''}})`);
      }});
    }}

    function copiarFormatoCantidad() {{
      if (!currentParsedOrder.items || currentParsedOrder.items.length === 0) {{
        alert("No hay productos detectados en el texto para copiar.");
        return;
      }}

      // Formato IR01XX (4 Columnas):
      // CANTIDAD (A) \t UN/MED (B) \t DETALLE (C) \t CODIGO (D)
      const rows = currentParsedOrder.items.map(it => {{
        const qtyVal = it.cajas > 0 ? it.cajas : (it.uni > 0 ? it.uni : it.qty);
        return `${{qtyVal}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`;
      }});

      const tsv = rows.join('\\n');
      navigator.clipboard.writeText(tsv).then(() => {{
        alert(`¡${{currentParsedOrder.items.length}} productos copiados para Formato por Cantidad!\\n\\n1. Ve a tu Google Sheets (formato IR01XX).\\n2. Haz clic en la celda A5 (CANTIDAD).\\n3. Presiona Ctrl + V para pegar.`);
        log(`[OK] ${{currentParsedOrder.items.length}} filas copiadas para Formato por Cantidad (4 col).`);
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

      // Formato IR01ML (5 Columnas):
      // CANT. CAJAS (A) \t CANT. UNI. (B) \t UN/MED (C) \t DETALLE (D) \t CODIGO (E)
      const rows = currentParsedOrder.items.map(it => {{
        const cajasVal = it.cajas > 0 ? it.cajas : (it.uni === 0 ? it.qty : '');
        const uniVal = it.uni > 0 ? it.uni : '';
        return `${{cajasVal}}\\t${{uniVal}}\\t${{it.unitType || 'UNI'}}\\t${{it.name}}\\t${{it.code}}`;
      }});

      const tsv = rows.join('\\n');
      navigator.clipboard.writeText(tsv).then(() => {{
        alert(`¡${{currentParsedOrder.items.length}} productos copiados para Formato por Cajas!\\n\\n1. Ve a tu Google Sheets (formato IR01ML).\\n2. Haz clic en la celda A5 (CANT. CAJAS).\\n3. Presiona Ctrl + V para pegar.`);
        log(`[OK] ${{currentParsedOrder.items.length}} filas copiadas para Formato por Cajas (5 col).`);
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
      const layout = document.getElementById('layout').value;
      const forceImages = document.getElementById('force_images').checked;
      const whatsapp = document.getElementById('whatsapp').value;
      
      consoleDiv.innerHTML = '';
      log(">>> Iniciando petición al backend...");
      
      btnRun.disabled = true;
      btnText.innerText = "Procesando...";
      btnSpinner.style.display = "block";
      
      const url = `/generar?codes=${{encodeURIComponent(codes)}}&sync=${{sync}}&layout=${{layout}}&force_images=${{forceImages}}&whatsapp=${{encodeURIComponent(whatsapp)}}`;
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
      }}
    }}
    
    async function testStockConnection(e) {{
      const btn = e ? e.target : null;
      const originalText = btn ? btn.innerText : 'Probar API';
      if (btn) btn.innerText = 'Conectando...';
      log(">>> [STOCK] Probando conexión con Google Apps Script (Uyus + Varios)...");
      try {{
        const res = await fetch("/api/stock", {{ cache: 'no-store' }});
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        const total = Object.keys(data).length;
        log(`>>> [STOCK OK] ¡Conexión exitosa con Google Drive! Se encontraron ${{total}} productos sincronizados en tiempo real.`, "success");
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

    // Escuchar eliminaciones en vivo desde la vista previa interactiva
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
        comando = f"ssh -R 80:127.0.0.1:{port} nokey@localhost.run"
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

    # Permitir puerto dinámico mediante variable de entorno para mayor compatibilidad
    port = int(os.environ.get("PORT", PORT))
    server_address = ('', port)
    
    # Intentar liberar el puerto si se quedó colgado de una ejecución anterior
    try:
        httpd = ThreadingHTTPServer(server_address, CatalogWebHandler)
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
    start_server()
