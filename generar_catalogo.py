# -*- coding: utf-8 -*-
"""
DongCheng Herramientas - Generador de Catálogo con HTML + CSS
===========================================================
REQUISITO: pip install openpyxl pillow
USO:
  1. Pon este script en la misma carpeta que catalogos.xlsx
  2. En Vista_Catalogo columna J desde fila 4 escribe los códigos
  3. Guarda y cierra el Excel
  4. Ejecuta el servidor web mediante Generar_Catalogo.bat
"""

from openpyxl import load_workbook
from datetime import date
from io import BytesIO
from PIL import Image as PILImage
import os, sys, math, re
import urllib.request
import subprocess
import hashlib
import json
import time

# ─── CONFIGURACIÓN ────────────────────────────────────────────
# La base de datos Excel se guarda como caché interna en temp_imgs
# para evitar tener el archivo catalogos.xlsx visible en la carpeta principal.
ARCHIVO_EXCEL       = os.path.join("temp_imgs", "catalogos_db_cache.xlsx")
URL_GOOGLE_SHEETS   = "https://docs.google.com/spreadsheets/d/181FkDYPFME5Fx75og4tNO3mvBMSGDY-M9IxGFcR28SI/edit?usp=sharing"
HOJA_DB             = "FORMATO INVENTARIO"
HOJA_VISTA          = "Vista_Catalogo"
HOJA_CATALOGO       = "CATALOGO"
COLUMNA_CODIGOS     = 10   # columna J en Vista_Catalogo
FILA_INICIO_CODIGOS = 4

# Columnas en DC MANU
COL_CATEGORIA = 2   # B
COL_CODIGO    = 3   # C
COL_IMAGEN    = 4   # D  ← imágenes están en esta columna
COL_TIPO      = 5   # E
COL_NOMBRE    = 6   # F
COL_SIZE      = 7   # G
COL_DETALLE   = 8   # H
COL_UNI       = 9   # I

FILA_INICIO_DB = 4
# ──────────────────────────────────────────────────────────────

def extraer_id_google_sheets(url_or_id):
    if not url_or_id:
        return None
    if "/" not in url_or_id and len(url_or_id) > 20:
        return url_or_id
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_or_id)
    if match:
        return match.group(1)
    return None

def descargar_base_de_datos_nube(url=None, destino=ARCHIVO_EXCEL):
    """
    Descarga la versión más reciente del Excel desde Google Sheets / Google Drive.
    """
    input_url = url or URL_GOOGLE_SHEETS
    if not input_url:
        return False
        
    doc_id = extraer_id_google_sheets(input_url)
    if doc_id:
        download_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx&t={int(time.time())}"
    else:
        download_url = input_url
 
    print(f"\n[NUBE] Sincronizando base de datos desde Google Drive...")
    try:
        req = urllib.request.Request(
            download_url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            data = response.read()
            if len(data) > 1000:  # Verificación mínima de tamaño
                dir_name = os.path.dirname(destino)
                if dir_name and not os.path.exists(dir_name):
                    os.makedirs(dir_name)
                with open(destino, "wb") as out_file:
                    out_file.write(data)
                print(f"[NUBE] [OK] Base de datos actualizada con éxito desde Google Sheets ({len(data)} bytes)")
                return True
            else:
                print(f"[NUBE] [AVISO] La respuesta descargada fue muy pequeña ({len(data)} bytes).")
                return False
    except urllib.error.HTTPError as he:
        if he.code == 404 or he.code == 403:
            print(f"[NUBE] [AVISO] El documento en Google Drive está privado o restringido (HTTP {he.code}).")
            print(f"[NUBE] Para activar la sincronización automática, haz clic en 'Compartir' en Google Sheets y selecciona 'Cualquier persona con el enlace'.")
        else:
            print(f"[NUBE] [AVISO] Error HTTP {he.code} al descargar de la nube: {he}")
        print(f"[NUBE] Se utilizará la copia local existente de '{destino}'.")
        return False
    except Exception as e:
        print(f"[NUBE] [AVISO] No se pudo descargar desde la nube: {e}")
        print(f"[NUBE] Se utilizará la copia local existente de '{destino}'.")
        return False

# Temas y logotipos para las marcas
BRAND_THEMES = {
    "DONGCHENG ELECT.": {
        "logo": "Logo DonElec.webp",
        "display_name": "DONGCHENG ELECTRICO",
        "header_bg": "FFFFFF",
        "subtitle_color": "005BAC",
        "category_bg": "DBEAFE",
        "category_fg": "1E40AF",
        "card_header_bg": "005BAC",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E2E8F0",
        "card_measure_fg": "000000",
    },
    "DONGCHENG MANUAL": {
        "logo": "Logo DonManu.jpg",
        "display_name": "DONGCHENG MANUAL",
        "header_bg": "FFFFFF",
        "subtitle_color": "334155",
        "category_bg": "E2E8F0",
        "category_fg": "334155",
        "card_header_bg": "005BAC",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E2E8F0",
        "card_measure_fg": "000000",
    },
    "FERRAWYY": {
        "logo": "Logo Ferr.png",
        "display_name": "FERRAWYY",
        "header_bg": "F97316",
        "subtitle_color": "FFFFFF",
        "category_bg": "FFEDD5",
        "category_fg": "C2410C",
        "card_header_bg": "EA580C",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FFEDD5",
        "card_measure_fg": "C2410C",
    },
    "UYUSTOOLS ELECT.": {
        "logo": "Logo UyuElec.jpg",
        "display_name": "UYUSTOOLS ELECTRICO",
        "header_bg": "FDC800",
        "subtitle_color": "000000",
        "category_bg": "1E293B",
        "category_fg": "FFFFFF",
        "card_header_bg": "FDC800",
        "card_header_fg": "000000",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FEF3C7",
        "card_measure_fg": "92400E",
    },
    "UYUSTOOLS MANUAL": {
        "logo": "Logo UyuManu.png",
        "display_name": "UYUSTOOLS MANUAL",
        "header_bg": "FDC800",
        "subtitle_color": "000000",
        "category_bg": "334155",
        "category_fg": "FFFFFF",
        "card_header_bg": "FDC800",
        "card_header_fg": "000000",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FEF3C7",
        "card_measure_fg": "92400E",
    },
    "PEGASUS": {
        "logo": "Logo Pega.png",
        "display_name": "PEGASUS",
        "header_bg": "FFFFFF",
        "subtitle_color": "1E293B",
        "category_bg": "F1F5F9",
        "category_fg": "0F172A",
        "card_header_bg": "334155",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E2E8F0",
        "card_measure_fg": "0F172A",
    },
    "FERTON": {
        "logo": "Logo Fer.jpg",
        "display_name": "FERTON",
        "header_bg": "000000",
        "subtitle_color": "FBBF24",
        "category_bg": "FEE2E2",
        "category_fg": "991B1B",
        "card_header_bg": "1E293B",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FEF3C7",
        "card_measure_fg": "78350F",
    },
    "AQUASTRONG": {
        "logo": "Logo Aquas.webp",
        "display_name": "AQUASTRONG",
        "header_bg": "FFFFFF",
        "subtitle_color": "0284C7",
        "category_bg": "E0F2FE",
        "category_fg": "0369A1",
        "card_header_bg": "0284C7",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E0F2FE",
        "card_measure_fg": "0369A1",
    },
    "GALAXIA": {
        "logo": "Logo Galaxia.png",
        "display_name": "GALAXIA",
        "header_bg": "FFFFFF",
        "subtitle_color": "16355F",
        "category_bg": "E0E7FF",
        "category_fg": "16355F",
        "card_header_bg": "16355F",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E0E7FF",
        "card_measure_fg": "16355F",
    },
    "GATE": {
        "logo": "Logo Gate.png",
        "display_name": "GATE",
        "header_bg": "FFFFFF",
        "subtitle_color": "EC403E",
        "category_bg": "FFECEC",
        "category_fg": "991B1B",
        "card_header_bg": "EC403E",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FEE2E2",
        "card_measure_fg": "991B1B",
    },
    "CROWN": {
        "logo": "Logo Crown.png",
        "display_name": "CROWN",
        "header_bg": "FFFFFF",
        "subtitle_color": "CA0A10", # Official Crown red
        "category_bg": "FFECEC", # Light red/rose matching brand
        "category_fg": "CA0A10", # Official Crown red
        "card_header_bg": "CA0A10", # Official Crown red
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E5E7EB", # Light gray matching screenshot
        "card_measure_fg": "CA0A10", # Official Crown red
    },
    "RIO": {
        "logo": "Logo Rio.png",
        "display_name": "RIO",
        "header_bg": "FFFFFF",
        "subtitle_color": "9B2226", # True tinto red
        "category_bg": "FEE2E2",
        "category_fg": "9B2226", # True tinto red
        "card_header_bg": "9B2226", # True tinto red
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FFE4E6",
        "card_measure_fg": "9B2226", # True tinto red
    },
    "OMEGA": {
        "logo": "Logo Omega.png",
        "display_name": "OMEGA",
        "header_bg": "FFFFFF",
        "subtitle_color": "7A1C1C", # Burgundy red (rojo tinto)
        "category_bg": "FEE2E2", # Light burgundy pink
        "category_fg": "7A1C1C", # Burgundy red (rojo tinto)
        "card_header_bg": "7A1C1C", # Burgundy red (rojo tinto)
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "E0F2FE", # Light blue
        "card_measure_fg": "0284C7", # Blue
    },
    "TOTAL": {
        "logo": "Logo total.png",
        "display_name": "TOTAL",
        "header_bg": "FFFFFF",
        "subtitle_color": "186E6F", # Characteristic green/teal
        "category_bg": "E6F3F3", # Light teal/green background matching brand
        "category_fg": "186E6F", # Characteristic green/teal
        "card_header_bg": "186E6F", # Characteristic green/teal
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FFECEC", # Light red/rose background
        "card_measure_fg": "E2231A", # Brand red text
    },
    "WADFOW": {
        "logo": "Logo Wadfow.png",
        "display_name": "WADFOW",
        "header_bg": "FFFFFF",
        "subtitle_color": "0A4E9D", # Wadfow blue
        "category_bg": "EFF6FF", # Very light blue background
        "category_fg": "0A4E9D", # Wadfow blue
        "card_header_bg": "0A4E9D", # Wadfow blue
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FFEDD5", # Light orange background
        "card_measure_fg": "FD5F00", # Wadfow orange text
    },
    "NEVA": {
        "logo": "LogoNeva.png",
        "display_name": "NEVA",
        "header_bg": "FFFFFF",
        "subtitle_color": "000000",
        "category_bg": "FEF08A", # Light yellow (Yellow 200)
        "category_fg": "000000",
        "card_header_bg": "000000", # Black
        "card_header_fg": "FACC15", # Yellow (Yellow 400)
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "FEF9C3", # Very light yellow (Yellow 100)
        "card_measure_fg": "000000",
    },
    "LUTIAN": {
        "logo": "LogoLutian.png",
        "display_name": "LUTIAN",
        "header_bg": "FFFFFF",
        "subtitle_color": "4F8A10", # Dark green text
        "category_bg": "F7FEE7", # Light lime green background (Lime 50)
        "category_fg": "3F6212", # Dark lime green text (Lime 800)
        "card_header_bg": "84CC16", # Lime green (Lime 500)
        "card_header_fg": "FFFFFF",
        "card_detail_bg": "FFFFFF",
        "card_photo_bg": "FFFFFF",
        "card_measure_bg": "F1FEE7", # Very light lime green
        "card_measure_fg": "4D7C0F", # Lime green text (Lime 700)
    }
}

def convert_svg_to_png(svg_path, png_path):
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
    except ImportError:
        try:
            print("  Instalando librerías necesarias para procesar archivos SVG (svglib y reportlab)...")
            subprocess.run([sys.executable, "-m", "pip", "install", "svglib", "reportlab"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            from svglib.svglib import svg2rlg
            from reportlab.graphics import renderPM
        except Exception as e:
            print(f"  [AVISO] No se pudo instalar svglib/reportlab: {e}.")
            return False
    try:
        drawing = svg2rlg(svg_path)
        renderPM.drawToFile(drawing, png_path, fmt="PNG")
        return True
    except Exception as e:
        print(f"  [AVISO] Error al convertir SVG a PNG: {e}")
        return False

def get_brand_theme(tipo):
    tipo_norm = str(tipo).upper().strip()
    for key, theme in BRAND_THEMES.items():
        if key.upper().replace(" ", "").replace(".", "") in tipo_norm.replace(" ", "").replace(".", ""):
            logo_path = theme["logo"]
            if logo_path:
                if os.path.exists(logo_path):
                    if logo_path.lower().endswith(".svg"):
                        png_path = logo_path[:-4] + ".png"
                        if convert_svg_to_png(logo_path, png_path):
                            theme_copy = theme.copy()
                            theme_copy["logo"] = png_path
                            return theme_copy
                    return theme
                base_name = os.path.splitext(logo_path)[0]
                for f in os.listdir("."):
                    if f.lower().startswith(base_name.lower()):
                        target_logo = f
                        if f.lower().endswith(".svg"):
                            png_path = base_name + ".png"
                            if convert_svg_to_png(f, png_path):
                                target_logo = png_path
                            else:
                                continue
                        theme_copy = theme.copy()
                        theme_copy["logo"] = target_logo
                        return theme_copy
            return theme
            
    clean_tipo = tipo_norm.replace(" ", "").replace(".", "").replace("-", "")
    for f in os.listdir("."):
        if f.lower().startswith("logo"):
            name_part = os.path.splitext(f.lower()[4:])[0].strip().replace(" ", "").replace("-", "").replace(".", "")
            if name_part and (name_part in clean_tipo.lower() or clean_tipo.lower() in name_part):
                if "ferr" in name_part and "ferton" in clean_tipo.lower():
                    continue
                target_logo = f
                if f.lower().endswith(".svg"):
                    base_name = os.path.splitext(f)[0]
                    png_path = base_name + ".png"
                    if convert_svg_to_png(f, png_path):
                        target_logo = png_path
                    else:
                        continue
                return {
                    "logo": target_logo,
                    "display_name": tipo_norm,
                    "header_bg": "FFFFFF",
                    "subtitle_color": "0F172A",
                    "category_bg": "F1F5F9",
                    "category_fg": "0F172A",
                    "card_header_bg": "334155",
                }
                
    return {
        "logo": None,
        "display_name": tipo_norm,
        "header_bg": "1E293B",
        "subtitle_color": "FFFFFF",
        "category_bg": "0F172A",
        "category_fg": "FFFFFF",
        "card_header_bg": "334155",
    }

def autocrop_image(img, tolerance=20):
    from PIL import ImageChops
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    bg_color = img.getpixel((0, 0))
    bg = PILImage.new("RGBA", img.size, bg_color)
    diff = ImageChops.difference(img, bg)
    diff_gray = diff.convert("L")
    threshold_img = diff_gray.point(lambda p: 255 if p > tolerance else 0)
    bbox = threshold_img.getbbox()
    if bbox:
        left, top, right, bottom = bbox
        left = max(0, left - 6)
        top = max(0, top - 6)
        right = min(img.width, right + 6)
        bottom = min(img.height, bottom + 6)
        return img.crop((left, top, right, bottom))
    return img

def generar_logo_blanco_empresa(input_path, output_path):
    """
    Convierte el logo de la empresa a una silueta puramente blanca con fondo transparente.
    """
    if os.path.exists(output_path) and os.path.getmtime(input_path) <= os.path.getmtime(output_path):
        return output_path
    try:
        img = PILImage.open(input_path).convert("RGBA")
        
        # Recortar márgenes iniciales
        bg_color = img.getpixel((0, 0))
        from PIL import ImageChops
        bg = PILImage.new("RGBA", img.size, bg_color)
        diff = ImageChops.difference(img, bg)
        diff_gray = diff.convert("L")
        threshold_img = diff_gray.point(lambda p: 255 if p > 30 else 0)
        bbox = threshold_img.getbbox()
        if bbox:
            img = img.crop(bbox)
            
        # Convertir elementos de color a blanco y hacer el fondo negro transparente
        datas = img.getdata()
        new_data = []
        for item in datas:
            r, g, b, a = item
            brightness = (r + g + b) / 3
            if brightness < 65:  # Umbral para el fondo oscuro
                new_data.append((0, 0, 0, 0))
            else:
                new_data.append((255, 255, 255, 255))
        img.putdata(new_data)
        
        # Volver a recortar los márgenes transparentes sobrantes
        bg_trans = PILImage.new("RGBA", img.size, (0, 0, 0, 0))
        diff_trans = ImageChops.difference(img, bg_trans)
        diff_trans_gray = diff_trans.convert("L")
        bbox_trans = diff_trans_gray.getbbox()
        if bbox_trans:
            img = img.crop(bbox_trans)
            
        ext = os.path.splitext(output_path)[1].lower()
        if ext == ".webp":
            img.save(output_path, "WEBP", quality=80)
        else:
            img.save(output_path, "PNG")
        return output_path
    except Exception as e:
        print(f"  [AVISO] No se pudo generar logo blanco de importadora: {e}")
def normalizar_codigo(cod):
    if cod is None:
        return ""
    # Convertir a texto y limpiar espacios normales e invisibles (\xa0, \u200b, \ufeff)
    s = str(cod).replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "").strip()
    # Si viene como flotante de Excel tipo "1024.0", dejarlo en "1024"
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    return s

def clave_busqueda(cod):
    # Clave de búsqueda limpia: sólo alfanuméricos en mayúsculas
    s = normalizar_codigo(cod)
    return re.sub(r'[^A-Za-z0-9]', '', s).upper()

def buscar_producto_en_db(cod, db, db_norm=None, db_clean=None):
    if not cod:
        return None
    raw = normalizar_codigo(cod)
    if raw in db:
        return db[raw]
    if db_norm is not None:
        norm = raw.upper()
        if norm in db_norm:
            return db_norm[norm]
    if db_clean is not None:
        clean = clave_busqueda(raw)
        if clean in db_clean:
            return db_clean[clean]
    return None

def detectar_hojas_inventario(wb):
    """
    Identifica las hojas que contienen inventario/productos.
    Prioriza 'FORMATO INVENTARIO' o nombres similares, o todas las hojas excepto vistas/catálogos.
    """
    candidatas = []
    # 1. Coincidencia por nombre de inventario o formato
    for name in wb.sheetnames:
        name_clean = name.strip().upper()
        if "FORMATO" in name_clean or "INVENTARIO" in name_clean or "INVENT" in name_clean:
            candidatas.append(wb[name])
            
    if candidatas:
        return candidatas
        
    # 2. Si no hay con esas palabras, tomar todas excepto las de vista o catálogo generado
    for name in wb.sheetnames:
        name_clean = name.strip().upper()
        if not any(k in name_clean for k in ["VISTA", "CATALOGO", "CONFIG", "RESUMEN", "PORTADA"]):
            candidatas.append(wb[name])
            
    if not candidatas:
        candidatas = [wb.active or wb[wb.sheetnames[0]]]
        
    return candidatas

def detectar_columnas(ws):
    """
    Detecta automáticamente las posiciones de las columnas de producto leyendo las primeras filas.
    Retorna diccionario con índices (1-based) y la fila de inicio de datos.
    """
    cols = {
        "categoria": COL_CATEGORIA,
        "codigo": COL_CODIGO,
        "imagen": COL_IMAGEN,
        "tipo": COL_TIPO,
        "nombre": COL_NOMBRE,
        "size": COL_SIZE,
        "detalle": COL_DETALLE,
        "uni": COL_UNI,
    }
    start_row = FILA_INICIO_DB
    
    max_scan_header = min(10, ws.max_row) if ws.max_row else 10
    found_headers = False
    
    for r in range(1, max_scan_header + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip().upper() for c in range(1, 20)]
        for idx, val in enumerate(row_vals, start=1):
            if not val:
                continue
            if any(k in val for k in ["CODIGO", "COD.", "ITEM", "REFERENCIA"]) and "BARRAS" not in val:
                cols["codigo"] = idx
                found_headers = True
            elif any(k in val for k in ["CATEGORIA", "RUBRO", "FAMILIA"]):
                cols["categoria"] = idx
            elif any(k in val for k in ["TIPO", "MARCA", "LINEA"]):
                cols["tipo"] = idx
            elif any(k in val for k in ["NOMBRE", "DESCRIPCION", "DESCRIP", "PRODUCTO"]) and "DETALLE" not in val:
                cols["nombre"] = idx
            elif any(k in val for k in ["MEDIDA", "SIZE", "TAMANO", "TAMAÑO", "DIMENSION"]):
                cols["size"] = idx
            elif any(k in val for k in ["DETALLE", "ESPECIFICACION", "CARACTERISTICA", "OBSERVACION"]):
                cols["detalle"] = idx
            elif any(k in val for k in ["UNI", "UNIDAD", "U.M.", "EMPAQUE", "PRESENTACION"]):
                cols["uni"] = idx
            elif any(k in val for k in ["IMAGEN", "FOTO", "IMG"]):
                cols["imagen"] = idx
                
        if found_headers:
            start_row = r + 1
            break
            
    return cols, start_row

def extraer_imagenes_db(ws_db, filas_interes=None):
    """
    Extrae las imágenes de la hoja de base de datos directamente de los objetos en memoria.
    Devuelve dict: { fila: bytes_imagen }
    Si se pasa filas_interes, solo extrae las imágenes que estén en esas filas específicas.
    """
    imagenes_por_fila = {}
    if not hasattr(ws_db, '_images'):
        return imagenes_por_fila
    print(f"  Imágenes detectadas en hoja '{ws_db.title}': {len(ws_db._images)}")
    for img in ws_db._images:
        try:
            ancla = img.anchor
            if hasattr(ancla, '_from'):
                fila = ancla._from.row + 1
            elif hasattr(ancla, 'row'):
                fila = ancla.row + 1
            else:
                continue
            
            # OPTIMIZACIÓN: Omitir la extracción de bytes si la fila no nos interesa
            if filas_interes is not None and fila not in filas_interes:
                continue
                
            if hasattr(img, 'ref') and img.ref:
                img.ref.seek(0)
                img_bytes = img.ref.read()
                img.ref.seek(0)
                imagenes_por_fila[fila] = img_bytes
        except Exception as e:
            pass
    return imagenes_por_fila

def get_color(color_str, default="000000"):
    """
    Formatea un color hex o palabra clave de color para CSS.
    """
    if not color_str:
        return f"#{default}"
    color_str = str(color_str).strip()
    if color_str.startswith("#"):
        return color_str
    if re.match(r'^[a-fA-F0-9]{3,8}$', color_str):
        return f"#{color_str}"
    return color_str

def exportar_a_pdf_edge(html_path, pdf_path):
    """
    Exporta un archivo HTML a PDF utilizando Microsoft Edge headless.
    """
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    
    edge_executable = None
    for path in edge_paths:
        if os.path.exists(path):
            edge_executable = path
            break
            
    if not edge_executable:
        print("\n[ERROR] No se encontró Microsoft Edge o Google Chrome en las rutas estándar.")
        return False
        
    abs_html = os.path.abspath(html_path)
    abs_pdf = os.path.abspath(pdf_path)
    
    if os.path.exists(abs_pdf):
        try:
            os.remove(abs_pdf)
        except Exception as e:
            print(f"[ERROR] No se pudo eliminar el archivo PDF existente: {e}")
            return False
            
    print(f"  Generando PDF con: {edge_executable}")
    cmd = [
        edge_executable,
        "--headless",
        "--no-header-footer",
        f"--print-to-pdf={abs_pdf}",
        abs_html
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if os.path.exists(abs_pdf):
            print(f"  [OK] PDF generado con éxito en: {abs_pdf}")
            return True
        else:
            print(f"  [ERROR] El navegador no generó el PDF. Detalle: {result.stdout} {result.stderr}")
            return False
    except Exception as e:
        print(f"  [ERROR] Ocurrió un error al convertir HTML a PDF: {e}")
        return False

def exportar_a_pdf(excel_path, pdf_path, sheet_name, forzar_imagenes=False):
    """
    Método para compatibilidad con el backend web.
    Redirige a la generación de HTML y conversión a PDF.
    """
    print("\n>>> Iniciando exportación rápida a PDF desde base de datos...")
    return generar(descargar_nube=False, forzar_imagenes=forzar_imagenes)

def generar_html_y_imagenes(db, codigos, imagenes_por_fila, layout="desktop", output_filename="catalogos.html", forzar_imagenes=False, db_norm=None, db_clean=None):
    import os
    import re
    def to_base64_src(path_or_bytes, default_mime="image/png"):
        if not path_or_bytes:
            return ""
        try:
            if isinstance(path_or_bytes, bytes):
                import base64
                encoded = base64.b64encode(path_or_bytes).decode('utf-8')
                return f"data:{default_mime};base64,{encoded}"
            elif isinstance(path_or_bytes, str) and os.path.exists(path_or_bytes):
                import base64
                ext = os.path.splitext(path_or_bytes)[1].lower()
                mime = "image/png"
                if ext == ".webp":
                    mime = "image/webp"
                elif ext in (".jpg", ".jpeg"):
                    mime = "image/jpeg"
                with open(path_or_bytes, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode('utf-8')
                return f"data:{mime};base64,{encoded}"
        except Exception as e:
            print(f"  [AVISO] No se pudo codificar imagen a base64: {e}")
        return path_or_bytes
    temp_dir = "temp_imgs"
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
        
    # Cargar base de hashes para invalidación de caché de imágenes de productos
    hashes_path = os.path.join(temp_dir, "image_hashes.json")
    image_hashes = {}
    if not forzar_imagenes and os.path.exists(hashes_path):
        try:
            with open(hashes_path, "r", encoding="utf-8") as f_h:
                image_hashes = json.load(f_h)
        except Exception:
            pass
    if forzar_imagenes:
        print("  [INFO] Forzando regeneración de todas las imágenes. Ignorando caché previa...")
    hashes_updated = forzar_imagenes
        
    # Variables de diseño condicional según el layout (Desktop A4 vs Celular)
    page_size_css = """
    @page {
      size: A4 portrait;
      margin: 12mm 10mm 12mm 10mm;
    }
    """ if layout != "mobile" else """
    @page {
      size: 108mm 192mm;
      margin: 8mm 6mm 8mm 6mm;
    }
    """
    
    grid_cols_css = "grid-template-columns: 1fr 1fr;" if layout != "mobile" else "grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));"
    min_height_css = "min-height: 272mm;" if layout != "mobile" else "min-height: 176mm;"
    cover_title_size = "30pt" if layout != "mobile" else "22pt"
    brand_banner_height = "80px" if layout != "mobile" else "55px"
    brand_logo_height = "65px" if layout != "mobile" else "45px"
    
    # ─── GENERACIÓN Y RECORTE DE LOGO BLANCO DE LA EMPRESA ───
    impor_logo_cropped = "Logo Impor.png"
    if os.path.exists("Logo Impor.png"):
        impor_logo_cropped = os.path.join(temp_dir, "logo_impor_blanco.webp")
        generar_logo_blanco_empresa("Logo Impor.png", impor_logo_cropped)
        
    # Organizar estructura
    brands_prods = {}       # { brand: { category: [products] } }
    brands_orden = []       # Orden de las marcas
    brands_cats_orden = {}  # { brand: [categories] }
    no_encontrados = []
    
    if db_norm is None:
        db_norm = {normalizar_codigo(k).upper(): v for k, v in db.items()}
    if db_clean is None:
        db_clean = {clave_busqueda(k): v for k, v in db.items()}

    rows_with_products = {p["fila_db"] for p in db.values()}
    
    print("\nProcesando y guardando imágenes de productos...")
    for cod in codigos:
        prod = buscar_producto_en_db(cod, db, db_norm, db_clean)
        if not prod:
            no_encontrados.append(cod)
            continue
        
        brand_raw = prod.get("tipo", "OTRO") or "OTRO"
        brand_theme = get_brand_theme(brand_raw)
        brand_name = brand_theme["display_name"]
        
        cat = prod["categoria"] or "Sin Categoría"
        
        if brand_name not in brands_prods:
            brands_prods[brand_name] = {}
            brands_orden.append(brand_name)
            brands_cats_orden[brand_name] = []
            
        if cat not in brands_prods[brand_name]:
            brands_prods[brand_name][cat] = []
            brands_cats_orden[brand_name].append(cat)
            
        # Extraer imagen
        fila_db = prod["fila_db"]
        ws_title = prod.get("ws_title", "")
        img_bytes = imagenes_por_fila.get((ws_title, fila_db)) or imagenes_por_fila.get(fila_db)
        if not img_bytes:
            img_bytes = imagenes_por_fila.get((ws_title, fila_db - 1)) or imagenes_por_fila.get(fila_db - 1)
            if not img_bytes:
                img_bytes = imagenes_por_fila.get((ws_title, fila_db + 1)) or imagenes_por_fila.get(fila_db + 1)
                
        img_relative_path = ""
        if img_bytes:
            try:
                clean_cod = re.sub(r'[\\/*?:"<>| ]', "_", prod["cod"])
                img_filename = f"prod_{clean_cod}.webp"
                img_path = os.path.join(temp_dir, img_filename)
                
                # Calcular el hash MD5 de los bytes de imagen originales del Excel
                img_hash = hashlib.md5(img_bytes).hexdigest()
                
                # Si el archivo no existe o el hash cambió, procesamos e invalidamos la caché
                if not os.path.exists(img_path) or image_hashes.get(clean_cod) != img_hash:
                    with open(img_path, "wb") as f_img:
                        f_img.write(img_bytes)
                    
                    # Auto recortar márgenes en blanco de la imagen y optimizar tamaño
                    try:
                        img_pil = PILImage.open(img_path)
                        img_cropped = autocrop_image(img_pil)
                        
                        # Redimensionar si es muy grande para reducir peso manteniendo calidad
                        max_dim = 400
                        if img_cropped.width > max_dim or img_cropped.height > max_dim:
                            resample_filter = getattr(PILImage, "Resampling", None)
                            filter_type = resample_filter.LANCZOS if resample_filter else getattr(PILImage, "ANTIALIAS", 3)
                            img_cropped.thumbnail((max_dim, max_dim), filter_type)
                            
                        img_cropped.save(img_path, "WEBP", quality=60)
                    except Exception:
                        pass
                        
                    image_hashes[clean_cod] = img_hash
                    hashes_updated = True
                
                img_relative_path = f"{temp_dir}/{img_filename}"
            except Exception as e:
                print(f"  [AVISO] No se pudo guardar imagen para {prod['cod']}: {e}")
                
        prod_copy = prod.copy()
        prod_copy["img_path"] = img_relative_path
        brands_prods[brand_name][cat].append(prod_copy)
        
    total_prods = sum(len(brands_prods[b][c]) for b in brands_prods for c in brands_prods[b])
    total_marcas = len(brands_orden)
    
    # Función auxiliar para determinar la plantilla de tarjeta según la marca
    def get_card_layout_class(brand_name_str):
        bname = str(brand_name_str).upper()
        if "UYUSTOOLS" in bname:
            return "card-bold"
        elif "FERRAWYY" in bname or "GATE" in bname:
            return "card-vibrant"
        elif any(x in bname for x in ["FERTON", "OMEGA"]):
            return "card-dark-luxury"
        else:
            return "card-tech"
            
    # Función auxiliar para saber si es tema oscuro
    def is_dark_theme(brand_name_str):
        bname = str(brand_name_str).upper()
        return any(x in bname for x in ["FERTON", "OMEGA"])
    
    # ─── CONSTRUIR HTML ───
    html_out = []
    
    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Catálogo de Productos - Importadora Rivero</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;600;700;800&display=swap');
    
    /* PLACEHOLDER_PAGE_SIZE */
    body {
      font-family: 'Plus Jakarta Sans', 'Inter', system-ui, sans-serif;
      background-color: #F1F5F9;
      color: #0F172A;
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
      /* PLACEHOLDER_BODY_STYLE */
    }
    
    /* Portada Estilo Revista con Texturas y Formas (Oscura como la Foto 1) */
    .cover-page {
      background: radial-gradient(circle at 80% 20%, #1E2530 0%, #0B0E14 100%);
      border-radius: 20px;
      padding: 30px 45px 25px 45px;
      box-sizing: border-box;
      min-height: auto;
      display: flex;
      flex-direction: column;
      justify-content: flex-start;
      page-break-after: always;
      break-after: page;
      position: relative;
      overflow: hidden;
      box-shadow: 0 10px 40px rgba(0,0,0,0.03);
    }
    /* Doble marco de lujo translúcido */
    .cover-page::before {
      content: "";
      position: absolute;
      top: 18px;
      left: 18px;
      right: 18px;
      bottom: 18px;
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 16px;
      pointer-events: none;
      z-index: 5;
    }
    .cover-page::after {
      content: "";
      position: absolute;
      top: 23px;
      left: 23px;
      right: 23px;
      bottom: 23px;
      border: 1px solid rgba(255, 255, 255, 0.03);
      border-radius: 14px;
      pointer-events: none;
      z-index: 5;
    }
    
    /* Formas decorativas desenfocadas de fondo */
    .decor-circle {
      position: absolute;
      border-radius: 50%;
      filter: blur(90px);
      z-index: 1;
      opacity: 0.15;
    }
    .decor-1 {
      width: 400px;
      height: 400px;
      background: #F97316; /* Naranja */
      top: -150px;
      right: -100px;
    }
    .decor-2 {
      width: 500px;
      height: 500px;
      background: #005BAC; /* Azul */
      bottom: -200px;
      left: -150px;
    }
    
    /* Agrupado hacia arriba para reducir espacios vacíos innecesarios */
    .cover-content {
      position: relative;
      z-index: 10;
      display: flex;
      flex-direction: column;
      justify-content: flex-start;
      height: 100%;
      flex-grow: 1;
    }
    
    /* Cabecera ultra-ajustada arriba y abajo (lo más corta posible) */
    .cover-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      width: 100%;
      margin-top: 5px;
      margin-bottom: 0px;
      padding: 0;
    }
    .company-logo {
      max-height: 42px;
      max-width: 200px;
      object-fit: contain;
    }
    .company-name-fallback {
      font-size: 16pt;
      font-weight: 800;
      color: #FFFFFF;
      margin: 0;
      letter-spacing: 1px;
    }
    .header-date-box {
      border: 1px solid rgba(255, 255, 255, 0.25);
      padding: 4px 10px;
      font-size: 8.5pt;
      font-weight: 700;
      letter-spacing: 2px;
      text-transform: uppercase;
      color: rgba(255, 255, 255, 0.85);
      border-radius: 2px;
    }
    
    /* Traemos el contenido más cerca de la cabecera eliminando la holgura central */
    .cover-main-content {
      margin-top: 35px;
      margin-bottom: 20px;
      text-align: left;
      padding-left: 20px;
      z-index: 10;
      position: relative;
    }
    .cover-label {
      font-size: 10pt;
      font-weight: 800;
      color: #F97316; /* Naranja */
      letter-spacing: 3px;
      text-transform: uppercase;
      margin-bottom: 12px;
    }
    .cover-main-title {
      font-size: /* PLACEHOLDER_COVER_TITLE_SIZE */;
      font-weight: 800;
      line-height: 1.15;
      color: #FFFFFF;
      margin: 0 0 15px 0;
      letter-spacing: 1px;
      text-transform: uppercase;
    }
    .title-highlight {
      color: #F97316; /* Naranja */
    }
    .cover-description {
      font-size: 10.5pt;
      line-height: 1.6;
      color: #94A3B8; /* Gris azulado suave */
      max-width: 580px;
      margin: 0;
    }
    
    /* El pie de página se empuja al final de forma natural mediante margin-top: auto */
    .cover-footer {
      width: 100%;
      display: flex;
      justify-content: flex-start;
      margin-top: auto;
      margin-bottom: 5px;
      padding-left: 20px;
    }
    .cover-footer-text {
      font-size: 9pt;
      color: rgba(255, 255, 255, 0.35);
      letter-spacing: 1px;
      text-transform: uppercase;
    }
    
    /* Secciones de Marca y Producto con Fondos Adaptados */
    .brand-section {
      page-break-before: always;
      break-before: page;
      padding: 15px 25px;
      margin-bottom: 0;
      /* PLACEHOLDER_MIN_HEIGHT */
      box-sizing: border-box;
    }
    
    /* Encabezado de marcas ajustado para que no sea espacioso */
    .brand-banner {
      width: 100%;
      height: /* PLACEHOLDER_BRAND_BANNER_HEIGHT */;
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 6px 15px;
      border-radius: 12px;
      margin-bottom: 5px;
      box-sizing: border-box;
      page-break-after: avoid;
      break-after: avoid;
      box-shadow: 0 4px 15px rgba(0,0,0,0.02);
    }
    .brand-logo {
      max-height: /* PLACEHOLDER_BRAND_LOGO_HEIGHT */;
      max-width: 95%;
      object-fit: contain;
    }
    .brand-logo-fallback {
      font-size: 20pt;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    .brand-subtitle {
      text-align: center;
      font-size: 10pt;
      font-weight: 800;
      margin-bottom: 15px;
      text-transform: uppercase;
      letter-spacing: 1.5px;
      page-break-after: avoid;
      break-after: avoid;
    }
    .category-section {
      margin-bottom: 25px;
    }
    .category-header {
      font-size: 11pt;
      font-weight: 800;
      padding: 8px 14px;
      border-radius: 8px;
      margin: 15px 0 10px 0;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.01);
      page-break-after: avoid;
      break-after: avoid;
    }
    .products-grid {
      display: grid;
      /* PLACEHOLDER_GRID_COLS */
      gap: 18px;
      margin-bottom: 20px;
    }
    
    /* ──── 1. TECH CARD (DongCheng, etc.) ──── */
    .card-tech {
      background: #FFFFFF;
      border: 1px solid rgba(15, 23, 42, 0.08);
      border-radius: 14px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 0 4px 15px rgba(0, 0, 0, 0.02);
      border-top: 4px solid var(--brand-header);
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .card-tech .card-header {
      background: rgba(15, 23, 42, 0.01);
      color: var(--brand-header);
      padding: 12px 14px 4px 14px;
      font-size: 9.5pt;
      font-weight: 800;
      text-align: left;
    }
    .card-tech .card-body {
      padding: 12px;
      display: flex;
      flex-direction: column;
      flex-grow: 1;
    }
    .card-tech .product-name {
      font-size: 10.5pt;
      font-weight: 700;
      color: #1E293B;
      margin: 0 0 12px 0;
      line-height: 1.3;
      height: 2.6em;
      overflow: hidden;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
    }
    .card-tech .image-container {
      background: #FFFFFF;
      height: 180px;
      display: flex;
      justify-content: center;
      align-items: center;
      margin-bottom: 12px;
      padding: 8px;
      border-radius: 8px;
      border: 1px solid rgba(15, 23, 42, 0.03);
    }
    .card-tech .product-img {
      max-width: 100%;
      max-height: 100%;
      object-fit: contain;
    }
    .card-tech .measure-pill {
      text-align: center;
      font-size: 9.5pt;
      font-weight: 700;
      padding: 6px 12px;
      border-radius: 20px;
      margin-bottom: 12px;
      background: var(--brand-measure-bg);
      color: var(--brand-measure-fg);
      border: 1px solid rgba(15, 23, 42, 0.03);
    }
    .card-tech .card-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: auto;
      padding-top: 8px;
      border-top: 1px dashed rgba(15, 23, 42, 0.08);
    }
    .card-tech .packaging-info {
      font-size: 9pt;
      font-weight: 600;
      color: #64748B;
    }
    .card-tech .availability-pill {
      background: #DCFCE7;
      color: #15803D;
      font-size: 8.5pt;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 12px;
    }
    
    /* ──── 2. BOLD CARD (Uyustools) ──── */
    .card-bold {
      background: #FFFFFF;
      border: 2px solid #0F172A;
      border-radius: 14px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 5px 5px 0px #0F172A;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .card-bold .card-header {
      background: #0F172A;
      color: var(--brand-header);
      padding: 9px 12px;
      font-size: 9.5pt;
      font-weight: 800;
      text-align: center;
      letter-spacing: 0.5px;
    }
    .card-bold .card-body {
      padding: 12px;
      display: flex;
      flex-direction: column;
      flex-grow: 1;
    }
    .card-bold .product-name {
      font-size: 10.5pt;
      font-weight: 800;
      color: #0F172A;
      margin: 0 0 12px 0;
      line-height: 1.3;
      height: 2.6em;
      overflow: hidden;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      text-transform: uppercase;
    }
    .card-bold .image-container {
      background: #FFFFFF;
      height: 180px;
      display: flex;
      justify-content: center;
      align-items: center;
      margin-bottom: 12px;
      padding: 8px;
      border: 2px solid #0F172A;
      border-radius: 8px;
    }
    .card-bold .product-img {
      max-width: 100%;
      max-height: 100%;
      object-fit: contain;
    }
    .card-bold .measure-pill {
      text-align: center;
      font-size: 9.5pt;
      font-weight: 800;
      padding: 6px 12px;
      border-radius: 6px;
      margin-bottom: 12px;
      background: var(--brand-header);
      color: #0F172A;
      border: 2px solid #0F172A;
    }
    .card-bold .card-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: auto;
      padding-top: 8px;
      border-top: 2px solid #0F172A;
    }
    .card-bold .packaging-info {
      font-size: 9pt;
      font-weight: 700;
      color: #0F172A;
    }
    .card-bold .availability-pill {
      background: #0F172A;
      color: var(--brand-header);
      font-size: 8.5pt;
      font-weight: 800;
      padding: 4px 10px;
      border-radius: 6px;
    }
    
    /* ──── 3. VIBRANT CARD (Ferrawyy, etc.) ──── */
    .card-vibrant {
      background: linear-gradient(180deg, #FFFFFF 0%, #FFFDFB 100%);
      border: 1px solid var(--brand-measure-bg);
      border-radius: 18px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 0 8px 25px rgba(249, 115, 22, 0.06);
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .card-vibrant .card-header {
      background: linear-gradient(135deg, var(--brand-header) 0%, var(--brand-measure-fg) 100%);
      color: #FFFFFF;
      padding: 11px 12px;
      font-size: 9pt;
      font-weight: 700;
      text-align: center;
      border-bottom-left-radius: 12px;
      border-bottom-right-radius: 12px;
      margin: 0 12px;
      box-shadow: 0 4px 10px rgba(249, 115, 22, 0.15);
    }
    .card-vibrant .card-body {
      padding: 12px;
      display: flex;
      flex-direction: column;
      flex-grow: 1;
    }
    .card-vibrant .product-name {
      font-size: 10.5pt;
      font-weight: 700;
      color: #2D3748;
      margin: 0 0 12px 0;
      line-height: 1.3;
      height: 2.6em;
      overflow: hidden;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
    }
    .card-vibrant .image-container {
      background: #FFFFFF;
      height: 180px;
      display: flex;
      justify-content: center;
      align-items: center;
      margin-bottom: 12px;
      padding: 8px;
      border-radius: 12px;
      border: 1px solid rgba(249, 115, 22, 0.08);
    }
    .card-vibrant .product-img {
      max-width: 100%;
      max-height: 100%;
      object-fit: contain;
    }
    .card-vibrant .measure-pill {
      text-align: center;
      font-size: 9.5pt;
      font-weight: 700;
      padding: 6px 12px;
      border-radius: 30px;
      margin-bottom: 12px;
      background: var(--brand-measure-bg);
      color: var(--brand-measure-fg);
    }
    .card-vibrant .card-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: auto;
      padding-top: 8px;
      border-top: 1px dashed var(--brand-measure-bg);
    }
    .card-vibrant .packaging-info {
      font-size: 9pt;
      font-weight: 600;
      color: var(--brand-measure-fg);
    }
    .card-vibrant .availability-pill {
      background: #FFEDD5;
      color: var(--brand-measure-fg);
      font-size: 8.5pt;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 12px;
    }
    
    /* ──── 4. LUXURY DARK CARD (Ferton, Omega, Rio, Crown) ──── */
    .card-dark-luxury {
      background: #1E293B;
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 16px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 0 12px 25px rgba(0, 0, 0, 0.35);
      border-top: 4px solid var(--brand-header);
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .card-dark-luxury .card-header {
      background: rgba(255, 255, 255, 0.02);
      color: var(--brand-header);
      padding: 12px 14px 4px 14px;
      font-size: 9.5pt;
      font-weight: 800;
      text-align: left;
    }
    .card-dark-luxury .card-body {
      padding: 12px;
      display: flex;
      flex-direction: column;
      flex-grow: 1;
    }
    .card-dark-luxury .product-name {
      font-size: 10.5pt;
      font-weight: 700;
      color: #FFFFFF;
      margin: 0 0 12px 0;
      line-height: 1.3;
      height: 2.6em;
      overflow: hidden;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
    }
    .card-dark-luxury .image-container {
      background: #FFFFFF;
      height: 180px;
      display: flex;
      justify-content: center;
      align-items: center;
      margin-bottom: 12px;
      padding: 8px;
      border-radius: 10px;
    }
    .card-dark-luxury .product-img {
      max-width: 100%;
      max-height: 100%;
      object-fit: contain;
    }
    .card-dark-luxury .measure-pill {
      text-align: center;
      font-size: 9.5pt;
      font-weight: 700;
      padding: 6px 12px;
      border-radius: 12px;
      margin-bottom: 12px;
      background: rgba(255, 255, 255, 0.06);
      color: var(--brand-header);
      border: 1px solid rgba(255, 255, 255, 0.08);
    }
    .card-dark-luxury .card-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: auto;
      padding-top: 8px;
      border-top: 1px dashed rgba(255, 255, 255, 0.1);
    }
    .card-dark-luxury .packaging-info {
      font-size: 9pt;
      font-weight: 600;
      color: #94A3B8;
    }
    .card-dark-luxury .availability-pill {
      background: rgba(16, 185, 129, 0.2);
      color: #34D399;
      font-size: 8.5pt;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 12px;
    }
    
    /* OMEGA Custom Branding - Black box background with light blue accents */
    .brand-section-omega .card-dark-luxury {
      background: #080C14;
      border: 1px solid rgba(56, 189, 248, 0.25);
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.55);
    }
    .brand-section-omega .card-dark-luxury .product-name {
      color: #F8FAFC;
    }
    .brand-section-omega .card-dark-luxury .measure-pill {
      background: rgba(56, 189, 248, 0.1);
      color: #38BDF8;
      border: 1px solid rgba(56, 189, 248, 0.25);
    }
    .brand-section-omega .card-dark-luxury .card-footer {
      border-top: 1px dashed rgba(56, 189, 248, 0.15);
    }
    .brand-section-omega .card-dark-luxury .packaging-info {
      color: #94A3B8;
    }
    
    /* RIO Custom Branding - Royal blue border matching packaging */
    .brand-section-rio .card-tech {
      border: 1.5px solid rgba(37, 99, 235, 0.4);
      box-shadow: 0 6px 18px rgba(37, 99, 235, 0.05);
    }
    .brand-section-rio .card-tech .card-header {
      border-bottom: 1.5px solid rgba(37, 99, 235, 0.15);
    }
    
    /* CROWN Custom Branding - Vibrant red border matching brand */
    .brand-section-crown .card-tech {
      border: 1.5px solid #CA0A10;
      border-top: 1.5px solid #CA0A10; /* uniform solid border all around */
      box-shadow: 0 6px 18px rgba(202, 10, 16, 0.05);
    }
    .brand-section-crown .card-tech .card-header {
      background: #CA0A10;
      color: #FFFFFF;
      border-bottom: none;
      padding: 10px 14px;
      font-weight: 700;
    }
    
    /* TOTAL Custom Branding - Teal green border and solid green header matching brand */
    .brand-section-total .card-tech {
      border: 1.5px solid #186E6F;
      border-top: 1.5px solid #186E6F; /* uniform solid border all around */
      box-shadow: 0 6px 18px rgba(24, 110, 111, 0.05);
    }
    .brand-section-total .card-tech .card-header {
      background: #186E6F;
      color: #FFFFFF;
      border-bottom: none;
      padding: 10px 14px;
      font-weight: 700;
    }
    
    /* WADFOW Custom Branding - Blue border and solid blue header matching brand */
    .brand-section-wadfow .card-tech {
      border: 1.5px solid #0A4E9D;
      border-top: 1.5px solid #0A4E9D; /* uniform solid border all around */
      box-shadow: 0 6px 18px rgba(10, 78, 157, 0.05);
    }
    .brand-section-wadfow .card-tech .card-header {
      background: #0A4E9D;
      color: #FFFFFF;
      border-bottom: none;
      padding: 10px 14px;
      font-weight: 700;
    }
    
    /* NEVA Custom Branding - Black border and solid black header with yellow text */
    .brand-section-neva .card-tech {
      border: 1.5px solid #000000;
      border-top: 1.5px solid #000000;
      box-shadow: 0 6px 18px rgba(0, 0, 0, 0.05);
    }
    .brand-section-neva .card-tech .card-header {
      background: #000000;
      color: #FACC15;
      border-bottom: none;
      padding: 10px 14px;
      font-weight: 700;
    }
    
    /* LUTIAN Custom Branding - Lime green border and solid green header matching brand */
    .brand-section-lutian .card-tech {
      border: 1.5px solid #84CC16;
      border-top: 1.5px solid #84CC16;
      box-shadow: 0 6px 18px rgba(132, 204, 22, 0.05);
    }
    .brand-section-lutian .card-tech .card-header {
      background: #84CC16;
      color: #FFFFFF;
      border-bottom: none;
      padding: 10px 14px;
      font-weight: 700;
    }
  </style>
</head>
<body>
"""
    
    body_style_css = """
      max-width: 1200px;
      margin: 0 auto;
      box-shadow: 0 0 30px rgba(0,0,0,0.05);
    """ if layout == "mobile" else ""
    
    html_template_processed = html_template.replace("/* PLACEHOLDER_PAGE_SIZE */", page_size_css)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_BODY_STYLE */", body_style_css)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_GRID_COLS */", grid_cols_css)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_MIN_HEIGHT */", min_height_css)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_COVER_TITLE_SIZE */", cover_title_size)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_BRAND_BANNER_HEIGHT */", brand_banner_height)
    html_template_processed = html_template_processed.replace("/* PLACEHOLDER_BRAND_LOGO_HEIGHT */", brand_logo_height)
    
    html_out.append(html_template_processed)
    
    # ─── AGREGAR PORTADA ───
    # (Página 1: Portada Limpia estilo Foto 1)
    hoy = date.today()
    mes_nombres = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
                   7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}
    titulo_fecha = f"{mes_nombres[hoy.month]} {hoy.year}"
    
    html_out.append('  <div class="cover-page">')
    # Añadir formas geométricas de fondo para diseño premium
    html_out.append('    <div class="decor-circle decor-1"></div>')
    html_out.append('    <div class="decor-circle decor-2"></div>')
    
    # Contenido flotando sobre el fondo
    html_out.append('    <div class="cover-content">')
    
    # Header Logo de Importadora Blanco (Izquierda) e Información de Fecha (Derecha) - Muy compacto
    html_out.append('      <div class="cover-header">')
    if os.path.exists(impor_logo_cropped):
        company_logo_b64 = to_base64_src(impor_logo_cropped)
        html_out.append(f'        <img class="company-logo" src="{company_logo_b64}" alt="Logo Importadora Rivero" />')
    else:
        html_out.append('        <span class="company-name-fallback">IMPORTADORA RIVERO</span>')
    
    # Caja de fecha estilo Foto 1
    html_out.append(f'        <div class="header-date-box">{titulo_fecha}</div>')
    html_out.append('      </div>')
    
    # Contenido Principal (Centro-Izquierda) - Traído hacia arriba
    html_out.append('      <div class="cover-main-content">')
    html_out.append('        <div class="cover-label">Catálogo General</div>')
    html_out.append('        <h1 class="cover-main-title">')
    html_out.append('          Herramientas &<br>Equipos <span class="title-highlight">Profesionales</span>')
    html_out.append('        </h1>')
    html_out.append('        <p class="cover-description">')
    html_out.append('          Selección de herramientas eléctricas, equipo industrial y accesorios eléctricos, respaldados por marcas de calidad comprobada para el profesional y el distribuidor.')
    html_out.append('        </p>')
    html_out.append('      </div>')
    
    # Slogan inferior estilo Foto 1
    html_out.append('      <div class="cover-footer">')
    html_out.append('        <div class="cover-footer-text">')
    html_out.append(f'          ✦ Resumen: {total_prods} productos en {total_marcas} marcas • simplificamos tu esfuerzo ✦')
    html_out.append('        </div>')
    html_out.append('      </div>')
    
    html_out.append('    </div>') # cover-content
    html_out.append('  </div>') # cover-page
    
    # (NOTA: El Índice General / Página 2 ha sido eliminado completamente según la petición del usuario)
    
    # ─── AGREGAR SECCIONES DE MARCAS ───
    for b_name in brands_orden:
        b_theme = get_brand_theme(b_name)
        bg_hdr = get_color(b_theme["header_bg"], "FFFFFF")
        card_hdr_bg = get_color(b_theme["card_header_bg"], "1E293B")
        card_hdr_fg = get_color(b_theme.get("card_header_fg", "FFFFFF"), "FFFFFF")
        card_meas_bg = get_color(b_theme.get("card_measure_bg", "F1F5F9"))
        card_meas_fg = get_color(b_theme.get("card_measure_fg", "0F172A"))
        
        # Determinar estilo de tarjeta y si es tema oscuro
        card_class = get_card_layout_class(b_name)
        is_dark = is_dark_theme(b_name)
        
        # Color de fondo de sección dinámica según branding
        bname_upper = b_name.upper()
        if "UYUSTOOLS" in bname_upper:
            sec_bg = "#FAF7F2"
            fg_sub = "#475569"
        elif "FERRAWYY" in bname_upper or "GATE" in bname_upper:
            sec_bg = "#FFF9F5"
            fg_sub = "#C2410C"
        elif "LUTIAN" in bname_upper:
            sec_bg = "#F7FAF4" # Soft green-tinted light background
            fg_sub = get_color(b_theme["subtitle_color"], "3F6212")
        elif "NEVA" in bname_upper:
            sec_bg = "#FAF9F5" # Warm ivory/linen background
            fg_sub = get_color(b_theme["subtitle_color"], "000000")
        elif is_dark:
            sec_bg = "#0B0F19"
            fg_sub = get_color(b_theme.get("subtitle_color"), "FBBF24")
        else:
            sec_bg = "#F3F7FC"
            fg_sub = get_color(b_theme["subtitle_color"], "0F172A")
            
        clean_brand_class = f"brand-section-{re.sub(r'[^a-zA-Z0-9]', '-', b_name.lower())}"
        html_out.append(f'  <div class="brand-section {clean_brand_class}" style="background-color: {sec_bg};">')
        
        # Crop del logo para el banner de la marca
        logo_path = b_theme.get("logo")
        logo_src = ""
        if logo_path and os.path.exists(logo_path):
            try:
                clean_bname = re.sub(r'[\\/*?:"<>| ]', "_", b_name)
                dest_logo_path = os.path.join(temp_dir, f"cropped_logo_{clean_bname}.webp")
                if not os.path.exists(dest_logo_path) or os.path.getmtime(logo_path) > os.path.getmtime(dest_logo_path):
                    logo_img = PILImage.open(logo_path)
                    logo_img = autocrop_image(logo_img)
                    if logo_img.width > 300 or logo_img.height > 300:
                        resample_filter = getattr(PILImage, "Resampling", None)
                        filter_type = resample_filter.LANCZOS if resample_filter else getattr(PILImage, "ANTIALIAS", 3)
                        logo_img.thumbnail((300, 300), filter_type)
                    logo_img.save(dest_logo_path, "WEBP", quality=70)
                logo_src = dest_logo_path
            except Exception as e:
                print(f"  [AVISO] No se pudo recortar el logo de la marca {b_name}: {e}")
                logo_src = logo_path

        # Cabecera de marca muy ajustada arriba y abajo
        html_out.append(f'    <div class="brand-banner" style="background-color: {bg_hdr}; border: 1px solid #E2E8F0;">')
        if logo_src:
            brand_logo_b64 = to_base64_src(logo_src)
            html_out.append(f'      <img class="brand-logo" src="{brand_logo_b64}" alt="{b_name}" />')
        else:
            html_out.append(f'      <span class="brand-logo-fallback" style="color: {fg_sub};">{b_name}</span>')
        html_out.append('    </div>')
        
        html_out.append(f'    <div class="brand-subtitle" style="color: {fg_sub};">MARCA: {b_name}</div>')
        
        # Categorías de la Marca
        for cat in brands_cats_orden[b_name]:
            prods = brands_prods[b_name][cat]
            
            if is_dark:
                cat_bg = "rgba(255, 255, 255, 0.05)"
                cat_fg = "#FFFFFF"
            else:
                cat_bg = get_color(b_theme["category_bg"], "F1F5F9")
                cat_fg = get_color(b_theme["category_fg"], "0F172A")
            
            html_out.append(f'    <div class="category-section">')
            # Encabezado de línea de categoría más estrecho
            html_out.append(f'      <div class="category-header" style="background-color: {cat_bg}; color: {cat_fg}; border-left: 4px solid {card_hdr_bg};">')
            html_out.append(f'        ✦ LÍNEA DE {cat.upper()}')
            html_out.append('      </div>')
            
            # Variables de estilo inyectadas en CSS inline para las tarjetas
            html_out.append(f'      <div class="products-grid" style="--brand-header: {card_hdr_bg}; --brand-header-fg: {card_hdr_fg}; --brand-measure-bg: {card_meas_bg}; --brand-measure-fg: {card_meas_fg};">')
            
            for prod in prods:
                html_out.append(f'        <div class="{card_class}">')
                html_out.append('          <div class="card-header">')
                html_out.append(f'            CÓDIGO: {prod["cod"]}')
                html_out.append('          </div>')
                
                html_out.append('          <div class="card-body">')
                html_out.append(f'            <div class="product-name">{prod["nombre"]}</div>')
                
                html_out.append('            <div class="image-container">')
                if prod.get("img_path"):
                    img_b64 = to_base64_src(prod["img_path"])
                    html_out.append(f'              <img class="product-img" src="{img_b64}" alt="{prod["nombre"]}" />')
                else:
                    html_out.append('              <div class="no-photo">SIN FOTO</div>')
                html_out.append('            </div>')
                
                html_out.append('            <div class="measure-pill">')
                html_out.append(f'              Medida: {prod["size"]}')
                html_out.append('            </div>')
                
                html_out.append('            <div class="card-footer">')
                html_out.append(f'              <span class="packaging-info">Caja: -- {prod["uni"]}</span>')
                html_out.append('              <span class="availability-pill">Disponible</span>')
                html_out.append('            </div>')
                
                html_out.append('          </div>') # card-body
                html_out.append('        </div>') # product-card
            html_out.append('      </div>') # products-grid
            html_out.append('    </div>') # category-section
            
        html_out.append('  </div>') # brand-section
        
    html_out.append("""</body>
</html>""")
    
    html_path = output_filename
    with open(html_path, "w", encoding="utf-8") as f_html:
        f_html.write("\n".join(html_out))
        
    if hashes_updated:
        try:
            with open(hashes_path, "w", encoding="utf-8") as f_h:
                json.dump(image_hashes, f_h, indent=2)
            print(f"  [CACHE] Guardados nuevos hashes de imágenes en '{hashes_path}'")
        except Exception as e:
            print(f"  [AVISO] No se pudo guardar el archivo de hashes de imágenes: {e}")
        
    if no_encontrados:
        print(f"\n  AVISO: Códigos no encontrados en la base de datos: {no_encontrados}")
        
    return html_path, total_prods

def generar(descargar_nube=True, codigos_custom=None, layout="desktop", forzar_imagenes=False):
    # Copiar Excel local de la raíz si existe y es más nuevo que la caché
    local_root_excel = "catalogos.xlsx"
    if os.path.exists(local_root_excel):
        if not os.path.exists(ARCHIVO_EXCEL) or os.path.getmtime(local_root_excel) > os.path.getmtime(ARCHIVO_EXCEL):
            print(f"\n[LOCAL] Detectado '{local_root_excel}' en la raíz más reciente que la caché. Copiando...")
            import shutil
            try:
                dest_dir = os.path.dirname(ARCHIVO_EXCEL)
                if dest_dir and not os.path.exists(dest_dir):
                    os.makedirs(dest_dir)
                shutil.copy2(local_root_excel, ARCHIVO_EXCEL)
                print(f"[LOCAL] [OK] Excel de la raíz copiado con éxito a la caché.")
            except Exception as e:
                print(f"[LOCAL] [AVISO] No se pudo copiar '{local_root_excel}': {e}")

    if descargar_nube and URL_GOOGLE_SHEETS:
        descargar_base_de_datos_nube(URL_GOOGLE_SHEETS, ARCHIVO_EXCEL)

    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"\nERROR: No se encontró '{ARCHIVO_EXCEL}'")
        raise FileNotFoundError(f"No se encontró el archivo base de datos Excel: {ARCHIVO_EXCEL}")

    print(f"\nAbriendo {ARCHIVO_EXCEL}...")
    wb = load_workbook(ARCHIVO_EXCEL, data_only=True)

    # 1. Leer códigos a procesar
    codigos = []
    if codigos_custom:
        if isinstance(codigos_custom, str):
            tokens = re.split(r'[\r\n,;\t]+', codigos_custom)
            codigos = [normalizar_codigo(c) for c in tokens if normalizar_codigo(c)]
        elif isinstance(codigos_custom, (list, tuple)):
            for item in codigos_custom:
                if isinstance(item, str):
                    tokens = re.split(r'[\r\n,;\t]+', item)
                    for t in tokens:
                        c_clean = normalizar_codigo(t)
                        if c_clean:
                            codigos.append(c_clean)
                else:
                    c_clean = normalizar_codigo(item)
                    if c_clean:
                        codigos.append(c_clean)

    if not codigos:
        # Buscar en hoja Vista_Catalogo si existe
        ws_vista = None
        for name in wb.sheetnames:
            if "VISTA" in name.upper():
                ws_vista = wb[name]
                break
        if not ws_vista and HOJA_VISTA in wb.sheetnames:
            ws_vista = wb[HOJA_VISTA]
            
        if ws_vista:
            empty_count = 0
            max_vista_rows = min(ws_vista.max_row + 10, 5000) if ws_vista.max_row else 1000
            for row in range(FILA_INICIO_CODIGOS, max_vista_rows):
                val = ws_vista.cell(row=row, column=COLUMNA_CODIGOS).value
                if val and str(val).strip():
                    codigos.append(normalizar_codigo(val))
                    empty_count = 0
                else:
                    empty_count += 1
                    if empty_count >= 50:
                        break

    if not codigos:
        print(f"No se encontraron códigos de producto a procesar.")
        raise ValueError(f"No se encontraron códigos de producto a procesar.")

    print(f"Códigos a procesar: {len(codigos)}")

    # 2. Identificar hojas de inventario y leer base de datos de productos de forma rápida y completa
    db = {}
    db_norm = {}
    db_clean = {}
    hojas_inv = detectar_hojas_inventario(wb)

    print(f"Hojas de inventario identificadas: {[ws.title for ws in hojas_inv]}")

    for ws_cur in hojas_inv:
        cols_cfg, fila_inicio = detectar_columnas(ws_cur)
        col_c = cols_cfg["codigo"]
        col_cat = cols_cfg["categoria"]
        col_tip = cols_cfg["tipo"]
        col_nom = cols_cfg["nombre"]
        col_siz = cols_cfg["size"]
        col_det = cols_cfg["detalle"]
        col_uni = cols_cfg["uni"]
        
        consecutive_empty = 0
        max_r = min(ws_cur.max_row + 50, 30000) if ws_cur.max_row else 15000
        
        for row in range(fila_inicio, max_r):
            cod_val = ws_cur.cell(row=row, column=col_c).value
            if cod_val is None or str(cod_val).strip() == "":
                consecutive_empty += 1
                if consecutive_empty >= 150:
                    break
                continue
            
            consecutive_empty = 0
            raw_clave = normalizar_codigo(cod_val)
            if not raw_clave:
                continue
                
            norm_k = raw_clave.upper()
            clean_k = clave_busqueda(raw_clave)
            
            prod_info = {
                "categoria": str(ws_cur.cell(row=row, column=col_cat).value or "").strip(),
                "cod":       raw_clave,
                "tipo":      str(ws_cur.cell(row=row, column=col_tip).value or "").strip(),
                "nombre":    str(ws_cur.cell(row=row, column=col_nom).value or "").strip(),
                "size":      str(ws_cur.cell(row=row, column=col_siz).value or "").strip(),
                "detalle":   str(ws_cur.cell(row=row, column=col_det).value or "").strip(),
                "uni":       str(ws_cur.cell(row=row, column=col_uni).value or "pcs").strip(),
                "fila_db":   row,
                "ws_title":  ws_cur.title,
            }
            
            if raw_clave not in db:
                db[raw_clave] = prod_info
            if norm_k not in db_norm:
                db_norm[norm_k] = prod_info
            if clean_k and clean_k not in db_clean:
                db_clean[clean_k] = prod_info

    print(f"Productos cargados en base de datos: {len(db)}")

    # 3. Determinar las filas de interés por hoja y extraer imágenes de forma optimizada
    filas_por_hoja = {ws_cur.title: set() for ws_cur in hojas_inv}
    
    for cod in codigos:
        prod = buscar_producto_en_db(cod, db, db_norm, db_clean)
        if prod:
            ws_name = prod.get("ws_title", hojas_inv[0].title)
            r = prod["fila_db"]
            if ws_name in filas_por_hoja:
                filas_por_hoja[ws_name].add(r)
                filas_por_hoja[ws_name].add(r - 1)
                filas_por_hoja[ws_name].add(r + 1)

    imagenes_por_fila = {}
    for ws_cur in hojas_inv:
        filas_int = filas_por_hoja.get(ws_cur.title, set())
        if filas_int:
            imgs = extraer_imagenes_db(ws_cur, filas_interes=filas_int)
            for r, img_data in imgs.items():
                imagenes_por_fila[(ws_cur.title, r)] = img_data
                if r not in imagenes_por_fila:
                    imagenes_por_fila[r] = img_data

    # 4. Eliminar hoja CATALOGO del Excel para reducir drásticamente el tamaño del archivo
    if HOJA_CATALOGO in wb.sheetnames:
        print(f"Limpiando hoja de catálogo antigua del Excel para reducir espacio...")
        del wb[HOJA_CATALOGO]
        try:
            wb.save(ARCHIVO_EXCEL)
            print(f"Excel guardado y optimizado ({ARCHIVO_EXCEL})")
        except Exception as e:
            print(f"  [AVISO] No se pudo optimizar el tamaño de '{ARCHIVO_EXCEL}': {e}")

    # 5. Generar HTML y extraer fotos a disco (primero el layout seleccionado)
    html_file, total_prods = generar_html_y_imagenes(db, codigos, imagenes_por_fila, layout=layout, output_filename="catalogos.html", forzar_imagenes=forzar_imagenes, db_norm=db_norm, db_clean=db_clean)

    # También generar copias de ambos layouts
    try:
        otro_layout = "mobile" if layout == "desktop" else "desktop"
        otro_filename = f"catalogos_{otro_layout}.html"
        generar_html_y_imagenes(db, codigos, imagenes_por_fila, layout=otro_layout, output_filename=otro_filename, forzar_imagenes=forzar_imagenes, db_norm=db_norm, db_clean=db_clean)
        # Guardar el diseño actual con su nombre específico también
        generar_html_y_imagenes(db, codigos, imagenes_por_fila, layout=layout, output_filename=f"catalogos_{layout}.html", forzar_imagenes=forzar_imagenes, db_norm=db_norm, db_clean=db_clean)
        print(f"  [HTML] Generadas copias específicas: 'catalogos_desktop.html' y 'catalogos_mobile.html'")
    except Exception as e:
        print(f"  [AVISO] No se pudo generar la copia del diseño alternativo: {e}")

    # 6. Proceso completado exitosamente (Solo HTML)
    print(f"\n[OK] ¡Proceso completado exitosamente!")
    print(f"  Catálogo HTML: {html_file}")
    print(f"  Total productos: {total_prods}")

# Meses diccionario global para generar_html_y_imagenes
mes_nombres = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
               7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}

if __name__ == "__main__":
    generar()