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
                    generar_catalogo.generar(descargar_nube=descargar_nube, codigos_custom=codigos_custom, layout=layout_raw, forzar_imagenes=forzar_imagenes)
                    # Enviar señal de éxito final
                    writer.write("EVENT_SUCCESS: Proceso finalizado con éxito.\n")
                except BaseException as e:
                    import traceback
                    traceback.print_exc(file=writer)
                    writer.write("EVENT_ERROR: Ocurrió un error al procesar el catálogo.\n")
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
      --bg-panel: rgba(13, 18, 30, 0.75);
      --bg-console: #030509;
      --border-glow: rgba(245, 158, 11, 0.3);
      --primary: #F59E0B;
      --primary-hover: #D97706;
      --accent: #F59E0B;
      --success: #10B981;
      --success-bg: rgba(16, 185, 129, 0.1);
      --text-main: #F8FAFC;
      --text-muted: #94A3B8;
      --border-panel: rgba(255, 255, 255, 0.06);
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

    /* Header Bar */
    header {{
      background: rgba(8, 12, 20, 0.9);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--border-panel);
      padding: 15px 30px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
    }}

    .header-left {{
      display: flex;
      align-items: center;
      gap: 15px;
    }}

    .header-logo {{
      max-height: 50px;
      object-fit: contain;
      border-radius: 4px;
      background: white;
      padding: 4px 8px;
    }}

    .header-title-container h1 {{
      font-size: 16pt;
      font-weight: 800;
      margin: 0;
      letter-spacing: 0.5px;
      color: var(--text-main);
    }}

    .header-title-container p {{
      font-size: 8.5pt;
      color: var(--accent);
      margin: 2px 0 0 0;
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
      font-size: 9pt;
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

    /* Main Grid Layout */
    .dashboard-container {{
      display: grid;
      grid-template-columns: 460px 1fr;
      gap: 20px;
      padding: 20px;
      flex-grow: 1;
      height: calc(100vh - 90px);
      box-sizing: border-box;
    }}

    .glass-panel {{
      background: var(--bg-panel);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid var(--border-panel);
      border-radius: 16px;
      padding: 22px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
      display: flex;
      flex-direction: column;
      height: 100%;
      box-sizing: border-box;
      overflow: hidden;
    }}

    /* Control Panel Form (Left) */
    .control-panel {{
      display: flex;
      flex-direction: column;
      gap: 18px;
    }}

    .section-title {{
      font-size: 11pt;
      font-weight: 800;
      color: var(--primary);
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin: 0 0 10px 0;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .textarea-container {{
      display: flex;
      flex-direction: column;
      gap: 6px;
    }}

    .label-text {{
      font-size: 9pt;
      font-weight: 600;
      color: var(--text-muted);
    }}

    textarea {{
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 10px;
      color: var(--text-main);
      padding: 12px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 9.5pt;
      resize: none;
      height: 90px;
      outline: none;
      transition: all 0.3s ease;
    }}

    textarea:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 10px var(--border-glow);
    }}

    /* Toggle Switch */
    .toggle-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: rgba(15, 23, 42, 0.4);
      padding: 10px 14px;
      border-radius: 10px;
      border: 1px solid rgba(255, 255, 255, 0.03);
    }}

    .switch {{
      position: relative;
      display: inline-block;
      width: 44px;
      height: 24px;
    }}

    .switch input {{
      opacity: 0;
      width: 0;
      height: 0;
    }}

    .slider {{
      position: absolute;
      cursor: pointer;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: #334155;
      transition: .3s;
      border-radius: 24px;
    }}

    .slider:before {{
      position: absolute;
      content: "";
      height: 16px;
      width: 16px;
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
      transform: translateX(20px);
    }}

    /* Start Button */
    .btn-generate {{
      background: linear-gradient(135deg, var(--primary) 0%, #D97706 100%);
      color: #0F172A;
      border: none;
      border-radius: 10px;
      padding: 14px 20px;
      font-family: inherit;
      font-size: 11pt;
      font-weight: 800;
      cursor: pointer;
      transition: all 0.25s ease;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 10px;
      box-shadow: 0 4px 15px rgba(245, 158, 11, 0.2);
    }}

    .btn-generate:hover {{
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(245, 158, 11, 0.35);
      background: linear-gradient(135deg, #FBBF24 0%, #D97706 100%);
    }}

    .btn-generate:active {{
      transform: translateY(1px);
    }}

    .btn-generate:disabled {{
      background: #334155;
      color: var(--text-muted);
      cursor: not-allowed;
      transform: none;
      box-shadow: none;
    }}

    /* Activity Console Box */
    .console-panel {{
      flex-grow: 1;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }}

    .console-output {{
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 10px;
      padding: 15px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 9pt;
      color: #CBD5E1;
      overflow-y: auto;
      flex-grow: 1;
      white-space: pre-wrap;
      box-shadow: inset 0 2px 8px rgba(0, 0, 0, 0.8);
      line-height: 1.5;
    }}

    .console-output::-webkit-scrollbar {{
      width: 8px;
    }}

    .console-output::-webkit-scrollbar-track {{
      background: var(--bg-console);
    }}

    .console-output::-webkit-scrollbar-thumb {{
      background: #334155;
      border-radius: 4px;
    }}

    .console-output::-webkit-scrollbar-thumb:hover {{
      background: #475569;
    }}

    .log-line {{
      margin-bottom: 4px;
    }}

    .log-success {{ color: var(--success); font-weight: 700; }}
    .log-error {{ color: #EF4444; font-weight: 700; }}
    .log-info {{ color: var(--primary); }}
    .log-warning {{ color: #F59E0B; }}

    /* Quick Action Buttons (Results) */
    .results-row {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-top: 5px;
    }}

    .btn-action {{
      border: none;
      border-radius: 8px;
      padding: 11px;
      font-family: inherit;
      font-size: 9.5pt;
      font-weight: 700;
      color: white;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: all 0.2s ease;
    }}

    .btn-action-pdf {{
      background-color: #EF4444;
    }}
    .btn-action-pdf:hover {{
      background-color: #DC2626;
      box-shadow: 0 4px 12px rgba(239, 68, 68, 0.2);
    }}

    .btn-action-folder {{
      background-color: #475569;
    }}
    .btn-action-folder:hover {{
      background-color: #334155;
      box-shadow: 0 4px 12px rgba(71, 85, 105, 0.2);
    }}

    .btn-action:disabled {{
      background: #1E293B;
      color: #475569;
      cursor: not-allowed;
      box-shadow: none;
    }}

    /* Preview Section (Right) */
    .preview-panel {{
      display: flex;
      flex-direction: column;
      position: relative;
    }}

    .preview-header-bar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }}

    .preview-title {{
      font-size: 12pt;
      font-weight: 800;
      color: var(--text-main);
      display: flex;
      align-items: center;
      gap: 10px;
    }}

    .device-selectors {{
      display: flex;
      gap: 6px;
      background: rgba(15, 23, 42, 0.5);
      padding: 3px;
      border-radius: 8px;
      border: 1px solid var(--border-panel);
    }}

    .device-btn {{
      background: transparent;
      border: none;
      color: var(--text-muted);
      border-radius: 6px;
      padding: 4px 10px;
      font-size: 8.5pt;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s;
    }}

    .device-btn.active {{
      background: var(--primary);
      color: #0F172A;
    }}
    
    .mini-icon {{
      display: inline-block;
      vertical-align: middle;
      width: 1.1em;
      height: 1.1em;
      fill: none;
      stroke: currentColor;
      stroke-width: 2;
      stroke-linecap: round;
      stroke-linejoin: round;
      flex-shrink: 0;
    }}

    .preview-viewport-wrapper {{
      flex-grow: 1;
      display: flex;
      justify-content: center;
      align-items: center;
      background: var(--bg-console);
      border: 1px solid var(--border-panel);
      border-radius: 12px;
      overflow: hidden;
      position: relative;
    }}

    iframe {{
      width: 100%;
      height: 100%;
      border: none;
      background: white;
      transition: width 0.4s ease;
    }}

    /* Mobile view styling */
    .view-mobile {{
      width: 420px;
      height: 90%;
      border: 10px solid #334155;
      border-radius: 20px;
      box-shadow: 0 20px 40px rgba(0,0,0,0.6);
    }}

    .no-preview {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 15px;
      color: var(--text-muted);
      text-align: center;
      padding: 40px;
      position: absolute;
      width: 100%;
      height: 100%;
      box-sizing: border-box;
      z-index: 10;
    }}

    .no-preview-icon {{
      font-size: 40pt;
      color: #334155;
    }}

    .spinner {{
      border: 3px solid rgba(255, 255, 255, 0.1);
      width: 20px;
      height: 20px;
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
        <p>Generador de Catálogos Premium</p>
      </div>
    </div>
    
    <div class="status-badge">
      <div class="status-dot"></div>
      Servidor Activo (Puerto {PORT})
    </div>
  </header>

  <!-- Dashboard Principal -->
  <main class="dashboard-container">
    
    <!-- Lado Izquierdo: Controles y Consola -->
    <div class="glass-panel control-panel">
      
      <div class="section-title">
        <svg class="mini-icon" viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
        Configuración de Generación
      </div>
      
      <!-- Entrada de códigos -->
      <div class="textarea-container">
        <label class="label-text" for="codes">Códigos de productos a incluir:</label>
        <textarea id="codes" placeholder="Pega los códigos aquí (uno por línea)...&#10;Ejemplo:&#10;ACC014&#10;ACC017&#10;ACT080&#10;Deja vacío para procesar todo lo de la nube."></textarea>
      </div>
      
      <!-- Sincronización en la Nube -->
      <div class="toggle-row">
        <span class="label-text" style="color: var(--text-main); font-weight: 600;">Sincronizar base de datos desde Google Drive</span>
        <label class="switch">
          <input type="checkbox" id="sync" checked>
          <span class="slider"></span>
        </label>
      </div>
      
      <!-- Forzar imágenes -->
      <div class="toggle-row">
        <span class="label-text" style="color: var(--accent); font-weight: 700;">Forzar regeneración de imágenes</span>
        <label class="switch">
          <input type="checkbox" id="force_images">
          <span class="slider"></span>
        </label>
      </div>
      
      <!-- Formato de Diseño PDF -->
      <div class="toggle-row">
        <span class="label-text" style="color: var(--text-main); font-weight: 600;">Diseño del PDF / Folleto</span>
        <select id="layout" style="background: var(--bg-console); border: 1px solid var(--border-panel); color: var(--text-main); padding: 6px 10px; border-radius: 8px; font-family: inherit; font-size: 9.5pt; outline: none; cursor: pointer; font-weight: 700;">
          <option value="desktop" selected>A4 Impresora (2 Columnas)</option>
          <option value="mobile">Celular / WhatsApp (1 Columna)</option>
        </select>
      </div>
      
      <!-- Botón de Generar -->
      <button class="btn-generate" id="btn-run" onclick="iniciarGeneracion()">
        <span class="spinner" id="btn-spinner"></span>
        <svg class="mini-icon" viewBox="0 0 24 24" style="stroke-width: 2.5; width: 1.2em; height: 1.2em; fill: currentColor; stroke: none; margin-right: 2px;"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>
        <span id="btn-text">Empezar Generación</span>
      </button>
      
      <!-- Actividad y Logs (Consola) -->
      <div class="console-panel">
        <div class="section-title" style="margin-top: 10px;">
          <svg class="mini-icon" viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="2" y1="20" x2="22" y2="20"></line><line x1="12" y1="17" x2="12" y2="20"></line></svg>
          Actividad del Servidor
        </div>
        <div class="console-output" id="console">Log de actividad listo. Presiona 'Empezar Generación' para iniciar el proceso...</div>
      </div>
      

    </div>

    <!-- Lado Derecho: Vista Previa Interactiva -->
    <div class="glass-panel preview-panel">
      
      <div class="preview-header-bar">
        <div class="preview-title">
          <svg class="mini-icon" viewBox="0 0 24 24" style="width: 1.15em; height: 1.15em;"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
          Vista Previa del Catálogo ({mes_año_actual})
        </div>
        
        <!-- Selectores de pantalla desktop/mobile y Ver Completo -->
        <div class="device-selectors" style="display: flex; gap: 6px;">
          <div style="display: flex; gap: 3px; background: rgba(15, 23, 42, 0.5); padding: 3px; border-radius: 8px; border: 1px solid var(--border-panel);">
            <button class="device-btn active" id="btn-device-desktop" onclick="setDevice('desktop')">
              <svg class="mini-icon" viewBox="0 0 24 24" style="width: 1.1em; height: 1.1em;"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line><line x1="12" y1="17" x2="12" y2="21"></line></svg>
              Escritorio
            </button>
            <button class="device-btn" id="btn-device-mobile" onclick="setDevice('mobile')">
              <svg class="mini-icon" viewBox="0 0 24 24" style="width: 1.1em; height: 1.1em;"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"></rect><line x1="12" y1="18" x2="12.01" y2="18"></line></svg>
              Celular
            </button>
          </div>
          <button class="device-btn" id="btn-full-preview" onclick="verCompleto()" style="background-color: rgba(245, 158, 11, 0.15); color: var(--accent); border: 1px solid rgba(245, 158, 11, 0.3);" {"" if preview_available == "true" else "disabled"}>
            <svg class="mini-icon" viewBox="0 0 24 24" style="width: 1.1em; height: 1.1em;"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
            Ver Completo
          </button>
          <a id="btn-download-html" href="#" download="catalogos_desktop.html" class="device-btn" style="background-color: rgba(16, 185, 129, 0.15); color: var(--success); border: 1px solid rgba(16, 185, 129, 0.3); text-decoration: none; display: flex; align-items: center; gap: 6px; pointer-events: none; opacity: 0.5;" onclick="return document.getElementById('btn-download-html').getAttribute('href') !== '#'">
            <svg class="mini-icon" viewBox="0 0 24 24" style="width: 1.1em; height: 1.1em;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            Descargar HTML
          </a>
        </div>
      </div>
      
      <!-- Contenedor del Iframe -->
      <div class="preview-viewport-wrapper">
        <div class="no-preview" id="no-preview" style="display: {'none' if preview_available == 'true' else 'flex'};">
          <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom: 10px; opacity: 0.3;"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
          <h3>El folleto aún no ha sido generado</h3>
          <p>Los resultados de la vista previa aparecerán en este panel una vez se complete la generación de los productos.</p>
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
    
    let hasDesktopFile = {desktop_available};
    let hasMobileFile = {mobile_available};
    let currentDevice = 'desktop';

    function getHtmlUrl(device) {{
      if (device === 'mobile' && hasMobileFile) {{
        return 'catalogos_mobile.html';
      }}
      if (device === 'desktop' && hasDesktopFile) {{
        return 'catalogos_desktop.html';
      }}
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
      
      if (device === 'mobile') {{
        iframe.className = 'view-mobile';
      }} else {{
        iframe.className = '';
      }}
      
      updateDownloadHtmlLink();
    }}
    
    function iniciarGeneracion() {{
      const codes = document.getElementById('codes').value;
      const sync = document.getElementById('sync').checked;
      const layout = document.getElementById('layout').value;
      const forceImages = document.getElementById('force_images').checked;
      
      // Limpiar consola
      consoleDiv.innerHTML = '';
      log(">>> Iniciando petición al backend...");
      
      // Bloquear controles
      btnRun.disabled = true;
      btnText.innerText = "Procesando...";
      btnSpinner.style.display = "block";
      
      // Conectar mediante EventSource (SSE)
      const url = `/generar?codes=${{encodeURIComponent(codes)}}&sync=${{sync}}&layout=${{layout}}&force_images=${{forceImages}}`;
      const source = new EventSource(url);
      
      source.onmessage = function(event) {{
        const msg = event.data;
        
        // Manejar señales especiales de fin
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
        
        // Cargar según el dispositivo seleccionado actualmente
        setDevice(currentDevice);
        
        // Habilitar botón de pantalla completa
        document.getElementById('btn-full-preview').disabled = false;
      }}
    }}
    
    function verCompleto() {{
      const fileUrl = getHtmlUrl(currentDevice);
      window.open(fileUrl + '?t=' + Date.now(), '_blank');
    }}
    
    // Funciones de escritorio desactivadas por seguridad en red local
    
    // Inicializar el link de descarga al cargar la página
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
