"""
Organizador de Pólizas de Seguro
---------------------------------
Extrae ZIPs, combina PDFs por titular, anexa EndosoBen,
comprime y guarda en carpeta "POLIZA OUTPUT".

Requiere: pip install pypdf pdfplumber
Ghostscript (opcional, para comprimir): https://www.ghostscript.com/releases/
"""

import os
import re
import shutil
import subprocess
import tempfile
import threading
import zipfile
from collections import defaultdict
from datetime import datetime
from tkinter import (
    Tk, Frame, Label, Button, Text, Scrollbar,
    filedialog, messagebox, StringVar, ttk
)

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


# ──────────────────────────────────────────────
#  Lógica principal
# ──────────────────────────────────────────────

def detectar_ghostscript():
    """Devuelve el ejecutable de Ghostscript si está disponible."""
    for cmd in ["gs", "gswin64c", "gswin32c"]:
        try:
            r = subprocess.run([cmd, "--version"], capture_output=True, timeout=5)
            if r.returncode == 0:
                return cmd
        except Exception:
            pass
    return None


def comprimir_pdf(src, gs_cmd):
    """Comprime src con Ghostscript. Devuelve ruta al comprimido (en /tmp)."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, dir=tempfile.gettempdir()) as tf:
        tmp_out = tf.name
    res = subprocess.run([
        gs_cmd, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/ebook", "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={tmp_out}", src
    ], timeout=60)
    if res.returncode == 0 and os.path.getsize(tmp_out) > 0:
        return tmp_out
    os.remove(tmp_out)
    return None


def _extraer_de_texto(text):
    """Aplica todos los patrones de regex sobre un bloque de texto. Devuelve (nombre, cedula) o (None, None)."""
    # Aseguradora del Sur
    m_as = re.search(r'ASEGURADO:\s*([A-ZÁÉÍÓÚÑ ]+?)\s+TELF.*?RUC/CED:\s*(\d{7,})', text, re.DOTALL)
    def _norm(s):
        return re.sub(r'\s+', ' ', s).strip()

    if m_as:
        return _norm(m_as.group(1)), m_as.group(2)

    # Equisuiza
    m_eq_n = re.search(r'Asegurado\s*:\s*\d+\s*-\s*([A-ZÁÉÍÓÚÑ ]+)', text)
    m_eq_c = re.search(r'RUC o Documento\s*:\s*(\d+)', text)
    if m_eq_n and m_eq_c:
        return _norm(m_eq_n.group(1)), m_eq_c.group(1)

    # Patrones estándar
    m_cedula = (
        re.search(r'CEDULA/RUC\s*:\s*(\d+)', text) or
        re.search(r'CEDULA\s*:\s*(\d+)', text) or
        re.search(r'RUC/CEDULA\s+ID\.?\s*[:\.]?\s*(\d+)', text) or   # Alianza imagen
        re.search(r'R\.U\.C\s*:\s*(\d+)', text) or
        re.search(r'CI\s*/\s*RUC\s+(\d+)', text) or
        re.search(r'RUC/CI\.?\s*:\s*(\d+)', text) or
        re.search(r'IDENTIFICACION\s*:\s*(\d+)', text)
    )
    m_nombre = (
        re.search(r'ASEGURADO\s*[:\s]+([A-ZÁÉÍÓÚÑ ]+)\s*\(', text) or
        re.search(r'ASEGURADO\s*[:\s]+([A-ZÁÉÍÓÚÑ ]+)\s*CEDULA', text) or
        re.search(r'ASEGURADO\s*[:\s]+([A-ZÁÉÍÓÚÑ ]+)\s*R\.U\.C', text) or
        re.search(r'ASEGURADO\s+([A-ZÁÉÍÓÚÑ ]+)\s+CI\s*/\s*RUC', text) or
        re.search(r'ASEGURADO\s*[:\s]+([A-ZÁÉÍÓÚÑ ]+)\s*IDENTIFICACION', text) or
        re.search(r'ASEGURADO\s*[:\s]+([A-ZÁÉÍÓÚÑ ]+)\r?\n', text)
    )
    if m_cedula and m_nombre:
        return _norm(m_nombre.group(1)), m_cedula.group(1)
    return None, None


def _ocr_pdf(pdf_path):
    """OCR de todas las páginas usando PyMuPDF + Tesseract. Devuelve texto completo."""
    if fitz is None:
        return ""
    try:
        doc = fitz.open(pdf_path)
        texto_total = ""
        for page in doc:
            pix = page.get_pixmap(dpi=250)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False,
                                             dir=tempfile.gettempdir()) as tf:
                img_path = tf.name
            pix.save(img_path)
            res = subprocess.run(
                ["tesseract", img_path, "stdout", "-l", "eng"],
                capture_output=True, text=True, timeout=30
            )
            os.remove(img_path)
            texto_total += res.stdout + "\n"
        return texto_total
    except Exception:
        return ""


def extraer_datos_pdf(pdf_path):
    """Extrae nombre y cédula. Primero intenta pdfplumber; si no hay texto, usa OCR."""
    # 1. Intento con pdfplumber (PDFs con texto)
    if pdfplumber is not None:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                texto_total = ""
                for page in pdf.pages:
                    texto_total += (page.extract_text() or "") + "\n"
                if texto_total.strip():
                    nombre, cedula = _extraer_de_texto(texto_total)
                    if nombre and cedula:
                        return nombre, cedula
        except Exception:
            pass

    # 2. Fallback OCR (PDFs de imagen, ej. Seguros Alianza)
    texto_ocr = _ocr_pdf(pdf_path)
    if texto_ocr.strip():
        nombre, cedula = _extraer_de_texto(texto_ocr)
        if nombre and cedula:
            return nombre, cedula

    return None, None


def guardar_pdf(src_bytes_path, output_dir, nombre_archivo):
    """Guarda el PDF en output_dir con el nombre dado, manejando archivos existentes."""
    destino = os.path.join(output_dir, nombre_archivo)
    if not os.path.exists(destino):
        try:
            open(destino, 'wb').close()
        except (FileNotFoundError, PermissionError):
            # Crear en raíz del output_dir y mover
            tmp_root = os.path.join(output_dir, "_placeholder.pdf")
            open(tmp_root, 'wb').close()
            os.rename(tmp_root, destino)
    with open(src_bytes_path, "rb") as fin, open(destino, "r+b") as fout:
        fout.seek(0)
        fout.write(fin.read())
        fout.truncate()
    return destino


_SUFIJOS_CONOCIDOS = {
    # Vehículo frente/póliza
    "FRENTE VH", "FRENTE", "POLIZA VH", "POLIZA", "CARATULA", "O",
    # Formato compuesto VH (NOMBRE VH POL / VH END / etc.)
    "VH POL", "VH FRENTE", "VH COND", "VH FC", "VH END",
    # Condiciones / carta de pago vehículo
    "COND VH", "CONDICIONES VH", "CONDICIONES", "COND", "VH", "PC", "CP VH", "CP VD", "CP",
    # Factura vehículo
    "FAC VH", "FACTURA VH", "FC VH", "FC VD", "FC", "FACT",
    # Endoso
    "ENDOSO", "EB", "ENDOSOBEN",
    # AP
    "AP", "POLIZA AP", "AP POL", "AP FRENTE", "AP FC",
    # Vida/VD
    "VIDA", "POLIZA VD", "COND VD", "VD POL", "VD FRENTE", "VD FC",
    # Factura VD
    "FACTURA VD", "FC VD",
    # Cuotas
    "CUOTAS", "CUOTAS VH", "CUOTAS VD",
    # Hogar / Incendio (pólizas standalone)
    "POLIZA MH", "POLIZA IN",
    # Formato compuesto IN/MH (al final para evitar colisiones con apellidos)
    "MH", "IN",
    # Documentos adicionales
    "CERTIFICADO BANCARIO", "CI",
}

# Sufijos que van a la carpeta PRODUBANCO (docs del vehículo + endoso)
_SUFIJOS_VH = {
    "FRENTE VH", "FRENTE", "POLIZA VH", "POLIZA", "CARATULA", "O",
    "VH POL", "VH FRENTE", "VH COND", "VH FC", "VH END",
    "COND VH", "CONDICIONES VH", "CONDICIONES", "COND", "VH", "PC", "CP VH", "CP VD", "CP",
    "FAC VH", "FACTURA VH", "FC VH", "FC VD", "FC", "FACT",
    "ENDOSO", "EB",
    "IN", "MH", "POLIZA IN", "POLIZA MH",
}

def _detectar_sufijo(nombre_sin_ext, sufijos_dict):
    """
    Detecta sufijo en el nombre de archivo. Soporta:
    - 'NOMBRE SUFIJO' y 'NOMBRE. SUFIJO' (con o sin punto)
    - Sufijos con paréntesis: 'NOMBRE endoso (2)' → sufijo = ENDOSO
    - Case-insensitive
    - Aseguradora del Sur: 'NOMBRE_NNNNN_NNN SUFIJO'
    """
    n_orig = nombre_sin_ext.strip()

    # Aseguradora del Sur: NOMBRE_POLIZA_CERT SUFIJO
    m = re.match(r'^(.+?)_\d+_\d+\s+(EB|O|PC)$', n_orig, re.IGNORECASE)
    if m:
        return m.group(1).replace('_', ' ').strip(), m.group(2).upper()

    # Latina formato: NOMBRE_CON_GUIONES_NUMEROO
    m = re.match(r'^([A-ZÁÉÍÓÚÑ_]+[A-ZÁÉÍÓÚÑ])_(\d{10,})O?$', n_orig, re.IGNORECASE)
    if m:
        return m.group(1).replace('_', ' ').strip(), "POLIZA"

    # Quitar paréntesis al final: "sufijo (N)" → "sufijo"
    n_clean = re.sub(r'\s*\(\d+\)\s*$', '', n_orig).strip()

    for suf in sorted(sufijos_dict.keys(), key=len, reverse=True):
        for sep in ['. ', ' ', '_']:
            if n_clean.upper().endswith(sep + suf.upper()):
                cliente = re.sub(r'\s+', ' ', n_clean[:-(len(suf) + len(sep))].rstrip('.')).strip()
                return cliente, suf.upper()
    return None, None

def _tiene_sufijos_conocidos(archivos):
    """Verifica si algún archivo tiene un sufijo conocido."""
    dummy = {s: 0 for s in _SUFIJOS_CONOCIDOS}
    for f in archivos:
        cliente, _ = _detectar_sufijo(f.replace(".pdf", ""), dummy)
        if cliente:
            return True
    return False


def organizar(base, log, gs_cmd):
    """Función principal de organización."""
    # ── Captura de log para reporte ──
    _log_lines = []
    _inicio = datetime.now()
    _log_orig = log
    def log(msg):
        _log_lines.append(msg)
        _log_orig(msg)

    if PdfReader is None or PdfWriter is None:
        log("❌ Faltan dependencias. Instala: pip install pypdf pdfplumber")
        return
    if pdfplumber is None:
        log("⚠️  pdfplumber no instalado. No se extraerá nombre/cédula automáticamente.")

    output_dir   = os.path.join(base, "POLIZA OUTPUT")
    clientes_dir = os.path.join(output_dir, "CLIENTES")
    produbanco_dir_out = os.path.join(output_dir, "PRODUBANCO")
    os.makedirs(clientes_dir, exist_ok=True)
    os.makedirs(produbanco_dir_out, exist_ok=True)
    _sin_endoso_global = []   # clientes VH sin endoso Produbanco
    tmp_work = tempfile.mkdtemp(prefix="polizas_")

    try:
        archivos = os.listdir(base)

        # ── Detectar tipo de entrada ──────────────────────────
        zips = [f for f in archivos if f.endswith(".zip")]
        pdfs_sueltos_base = [f for f in archivos if f.endswith(".pdf")]
        tiene_prefijos_alianza = any(
            f.startswith(("car_", "cnp_", "b_1_")) for f in pdfs_sueltos_base
        )

        # ── CASO A: ZIPs ──────────────────────────────────────
        if zips:
            log(f"📦 Encontrados {len(zips)} ZIPs")
            produbanco_dir = os.path.join(tmp_work, "PRODUBANCO")
            os.makedirs(produbanco_dir, exist_ok=True)

            for archivo in zips:
                nombre = archivo
                for prefijo in ["Polizas_Generadas_", "Polizas_"]:
                    if nombre.startswith(prefijo):
                        nombre = nombre[len(prefijo):]
                        break
                nombre = re.sub(r'\s*\(\d+\)\s*', '', nombre).replace(".zip", "").strip()
                es_banco = "PRODUBANCO" in nombre or "BANCO_DE_LA_PRODUCCION" in nombre
                carpeta = "PRODUBANCO" if es_banco else nombre
                destino_zip = os.path.join(tmp_work, carpeta)
                os.makedirs(destino_zip, exist_ok=True)
                try:
                    subprocess.run(
                        ["unzip", "-o", os.path.join(base, archivo), "-d", destino_zip],
                        capture_output=True, timeout=30
                    )
                except Exception:
                    # Fallback Python
                    try:
                        with zipfile.ZipFile(os.path.join(base, archivo)) as z:
                            z.extractall(destino_zip)
                    except Exception as e:
                        log(f"  ⚠️  No se pudo extraer {archivo}: {e}")

            # Copiar EndosoBen sueltos de la carpeta raíz
            for f in pdfs_sueltos_base:
                if "EndosoBen" in f:
                    shutil.copy2(os.path.join(base, f), produbanco_dir)

            # Recopilar todos los PDFs extraídos recursivamente (excluye PRODUBANCO)
            pdfs_extraidos = {}  # nombre → ruta completa
            for root_dir, dirs, files in os.walk(tmp_work):
                if "PRODUBANCO" in root_dir:
                    continue
                for f in files:
                    if f.endswith(".pdf"):
                        pdfs_extraidos[f] = os.path.join(root_dir, f)

            hay_sufijos = _tiene_sufijos_conocidos(list(pdfs_extraidos.keys()))

            if hay_sufijos:
                # Formato NOMBRE SUFIJO dentro del ZIP (ej: POLIZAS CASH)
                SUFIJOS = {
                    "FRENTE VH": 0, "FRENTE": 0, "POLIZA VH": 0, "POLIZA": 0,
                    "CARATULA": 0, "O": 0, "POLIZA MH": 0, "POLIZA IN": 0, "MH": 0, "IN": 0,
                    "VH POL": 0, "VH FRENTE": 0,
                    "COND VH": 1, "CONDICIONES VH": 1, "CONDICIONES": 1, "COND": 1,
                    "VH COND": 1, "VH": 1, "PC": 1, "CP VH": 1, "CP VD": 1, "CP": 1, "CERTIFICADO BANCARIO": 1,
                    "VH FC": 2, "FAC VH": 2, "FACTURA VH": 2, "FC VH": 2, "FC VD": 2, "FC": 2, "FACT": 2,
                    "VH END": 3, "ENDOSO": 3, "EB": 3,
                    "AP": 4, "POLIZA AP": 4, "AP POL": 4, "AP FRENTE": 4,
                    "VIDA": 5, "POLIZA VD": 5, "COND VD": 5, "VD POL": 5, "VD FRENTE": 5,
                    "FACTURA VD": 6, "FC VD": 6, "AP FC": 6, "VD FC": 6,
                    "CUOTAS": 7, "CUOTAS VH": 7, "CUOTAS VD": 7,
                    "ENDOSOBEN": 8, "CI": 9,
                }
                grupos_tipo = defaultdict(list)
                for nombre_f, ruta_f in pdfs_extraidos.items():
                    cliente, sufijo = _detectar_sufijo(nombre_f.replace(".pdf", ""), SUFIJOS)
                    if cliente:
                        grupos_tipo[cliente].append((sufijo, ruta_f))

                log(f"📄 Formato NOMBRE TIPO en ZIP — {len(grupos_tipo)} cliente(s)")
                sin_endoso_zip = []
                for cliente, archivos_c in sorted(grupos_tipo.items()):
                    archivos_ord = sorted(archivos_c, key=lambda x: SUFIJOS.get(x[0].upper(), 99))

                    # ── Archivo completo → CLIENTES ──
                    writer_full = PdfWriter()
                    for _, pdf_path in archivos_ord:
                        for page in PdfReader(pdf_path).pages:
                            writer_full.add_page(page)
                    _guardar_titular(writer_full, base, clientes_dir, cliente, gs_cmd, log)

                    # ── PRODUBANCO: VH-only si tiene VH, completo si no ──
                    archivos_vh = [(s, f) for s, f in archivos_ord if s.upper() in _SUFIJOS_VH]
                    tiene_poliza_vh = any(
                        s.upper() in {"FRENTE VH","FRENTE","POLIZA VH","POLIZA","CARATULA","O",
                                      "VH POL","VH FRENTE"}
                        and s.upper() not in {"POLIZA MH","POLIZA IN","MH","IN"}
                        for s, _ in archivos_ord
                    )
                    tiene_endoso = any(s.upper() in {"ENDOSO","EB","ENDOSOBEN","VH END"} for s, _ in archivos_c)
                    if archivos_vh and tiene_poliza_vh:
                        writer_vh = PdfWriter()
                        for _, pdf_path in archivos_vh:
                            for page in PdfReader(pdf_path).pages:
                                writer_vh.add_page(page)
                        _guardar_titular(writer_vh, base, produbanco_dir_out, cliente, gs_cmd, log)
                        if not tiene_endoso:
                            sin_endoso_zip.append(cliente)
                    else:
                        # Sin VH → archivo completo a PRODUBANCO
                        writer_cp = PdfWriter()
                        for _, pdf_path in archivos_ord:
                            for page in PdfReader(pdf_path).pages:
                                writer_cp.add_page(page)
                        _guardar_titular(writer_cp, base, produbanco_dir_out, cliente, gs_cmd, log)

                _sin_endoso_global.extend(sin_endoso_zip)
                if sin_endoso_zip:
                    reporte_path = os.path.join(output_dir, "SIN ENDOSO.txt")
                    with open(reporte_path, "w", encoding="utf-8") as rf:
                        rf.write("CLIENTES CON POLIZA VH SIN ENDOSO PRODUBANCO\n")
                        rf.write("=" * 45 + "\n")
                        for c in sin_endoso_zip:
                            rf.write(f"- {c}\n")
                    log(f"⚠️  {len(sin_endoso_zip)} cliente(s) sin endoso → SIN ENDOSO.txt")

            else:
                # Formato Produbanco clásico (PDFs numerados en subcarpetas)
                EXCLUIDAS = {"PRODUBANCO"}
                mapa_polizas = {}
                for carpeta in os.listdir(tmp_work):
                    ruta = os.path.join(tmp_work, carpeta)
                    if not os.path.isdir(ruta) or carpeta in EXCLUIDAS:
                        continue
                    for pdf in os.listdir(ruta):
                        m = re.search(r'_(\d+)\.pdf$', pdf)
                        if m:
                            mapa_polizas[int(m.group(1))] = carpeta
                            break

                endosos = {}
                for pdf in os.listdir(produbanco_dir):
                    if "EndosoBen" not in pdf:
                        continue
                    m = re.search(r'_(\d+)\.pdf$', pdf)
                    if not m:
                        continue
                    num = int(m.group(1))
                    titular = mapa_polizas.get(num - 1)
                    if titular:
                        endosos[titular] = os.path.join(produbanco_dir, pdf)

                for carpeta in sorted(os.listdir(tmp_work)):
                    ruta = os.path.join(tmp_work, carpeta)
                    if not os.path.isdir(ruta) or carpeta in EXCLUIDAS:
                        continue
                    pdfs_titular = sorted([f for f in os.listdir(ruta) if f.endswith(".pdf")])
                    if not pdfs_titular:
                        continue

                    writer = PdfWriter()
                    for pdf in pdfs_titular:
                        for page in PdfReader(os.path.join(ruta, pdf)).pages:
                            writer.add_page(page)
                    if carpeta in endosos:
                        for page in PdfReader(endosos[carpeta]).pages:
                            writer.add_page(page)
                        log(f"  📎 EndosoBen anexado a {carpeta}")

                    _guardar_titular(writer, base, clientes_dir, carpeta, gs_cmd, log)

        # ── CASO B: PDFs sueltos Seguros Alianza (car_, cnp_, b_1_) ──
        elif tiene_prefijos_alianza:
            log("📄 Formato Seguros Alianza detectado")
            orden = {"car": 0, "cnp": 1, "b_1": 2}
            grupos = defaultdict(list)
            for f in pdfs_sueltos_base:
                num = re.search(r'_(\d+)\.pdf$', f)
                if num:
                    grupos[num.group(1)].append(f)

            for num_poliza, archivos_grupo in sorted(grupos.items()):
                archivos_ord = sorted(archivos_grupo,
                    key=lambda f: orden.get(f.split("_")[0], 99))
                writer = PdfWriter()
                for pdf in archivos_ord:
                    for page in PdfReader(os.path.join(base, pdf)).pages:
                        writer.add_page(page)
                carpeta = f"poliza_{num_poliza}"
                _guardar_titular(writer, base, clientes_dir, carpeta, gs_cmd, log)

        # ── CASO C: PDFs sueltos con sufijo tipo (NOMBRE FRENTE.pdf, AP, VH, etc.) ──
        elif _tiene_sufijos_conocidos(pdfs_sueltos_base):
            SUFIJOS = {
                "FRENTE VH": 0, "FRENTE": 0, "POLIZA VH": 0, "POLIZA": 0,
                "CARATULA": 0, "O": 0, "POLIZA MH": 0, "POLIZA IN": 0, "MH": 0, "IN": 0,
                "VH POL": 0, "VH FRENTE": 0,
                "COND VH": 1, "CONDICIONES VH": 1, "CONDICIONES": 1, "COND": 1,
                "VH COND": 1, "VH": 1, "PC": 1, "CP VH": 1, "CP VD": 1, "CP": 1, "CERTIFICADO BANCARIO": 1,
                "VH FC": 2, "FAC VH": 2, "FACTURA VH": 2, "FC VH": 2, "FC VD": 2, "FC": 2, "FACT": 2,
                "VH END": 3, "ENDOSO": 3, "EB": 3,
                "AP": 4, "POLIZA AP": 4, "AP POL": 4, "AP FRENTE": 4,
                "VIDA": 5, "POLIZA VD": 5, "COND VD": 5, "VD POL": 5, "VD FRENTE": 5,
                "FACTURA VD": 6, "FC VD": 6, "AP FC": 6, "VD FC": 6,
                "CUOTAS": 7, "CUOTAS VH": 7, "CUOTAS VD": 7,
                "ENDOSOBEN": 8, "CI": 9,
            }

            # Separar EndosoBen sueltos de Produbanco
            endosoben_sueltos = {}
            pdfs_para_grupo = []
            for f in pdfs_sueltos_base:
                m_banco = re.match(r'^BANCO_DE_LA_PRODUCCION_S_A__PRODUBANCO\s+(.+)\.pdf$', f, re.IGNORECASE)
                if m_banco:
                    endosoben_sueltos[f] = m_banco.group(1).upper().split()
                else:
                    pdfs_para_grupo.append(f)

            grupos_tipo = defaultdict(list)
            for f in pdfs_para_grupo:
                cliente, sufijo = _detectar_sufijo(f.replace(".pdf", ""), SUFIJOS)
                if cliente:
                    grupos_tipo[cliente].append((sufijo, f))

            # Asociar EndosoBen al cliente por nombre
            for f_endoso, palabras in endosoben_sueltos.items():
                asociado = False
                for cliente in grupos_tipo:
                    if all(p in cliente.upper() for p in palabras):
                        grupos_tipo[cliente].append(("ENDOSO", f_endoso))
                        log(f"  📎 EndosoBen '{f_endoso}' → {cliente}")
                        asociado = True
                        break
                if not asociado:
                    log(f"  ⚠️  EndosoBen '{f_endoso}' sin cliente — se procesa solo")
                    grupos_tipo[f_endoso.replace('.pdf', '')].append(("ENDOSO", f_endoso))

            log(f"📄 Formato NOMBRE TIPO — {len(grupos_tipo)} cliente(s)")
            sin_endoso = []

            for cliente, archivos_c in sorted(grupos_tipo.items()):
                archivos_ord = sorted(archivos_c, key=lambda x: SUFIJOS.get(x[0].upper(), 99))

                # ── Archivo completo → CLIENTES ──
                writer_full = PdfWriter()
                for _, pdf in archivos_ord:
                    for page in PdfReader(os.path.join(base, pdf)).pages:
                        writer_full.add_page(page)
                nombre_archivo = _guardar_titular(writer_full, base, clientes_dir, cliente, gs_cmd, log)

                # ── Archivo VH → PRODUBANCO (solo si tiene docs de vehículo) ──
                archivos_vh = [(s, f) for s, f in archivos_ord if s.upper() in _SUFIJOS_VH]
                tiene_poliza_vh = any(
                    s.upper() in {"FRENTE VH","FRENTE","POLIZA VH","POLIZA","CARATULA","O",
                                  "VH POL","VH FRENTE"}
                    and s.upper() not in {"POLIZA MH","POLIZA IN","MH","IN"}
                    for s, _ in archivos_ord
                )
                tiene_endoso = any(s.upper() in {"ENDOSO","EB","ENDOSOBEN","VH END"} for s, _ in archivos_c)

                if archivos_vh and tiene_poliza_vh:
                    writer_vh = PdfWriter()
                    for _, pdf in archivos_vh:
                        for page in PdfReader(os.path.join(base, pdf)).pages:
                            writer_vh.add_page(page)
                    _guardar_titular(writer_vh, base, produbanco_dir_out, cliente, gs_cmd, log)
                    if not tiene_endoso:
                        sin_endoso.append(cliente)
                else:
                    # Sin VH → archivo completo a PRODUBANCO
                    writer_cp = PdfWriter()
                    for _, pdf in archivos_ord:
                        for page in PdfReader(os.path.join(base, pdf)).pages:
                            writer_cp.add_page(page)
                    _guardar_titular(writer_cp, base, produbanco_dir_out, cliente, gs_cmd, log)

            # ── Reporte SIN ENDOSO ──
            _sin_endoso_global.extend(sin_endoso)
            if sin_endoso:
                reporte_path = os.path.join(output_dir, "SIN ENDOSO.txt")
                with open(reporte_path, "w", encoding="utf-8") as rf:
                    rf.write("CLIENTES CON POLIZA VH SIN ENDOSO PRODUBANCO\n")
                    rf.write("=" * 45 + "\n")
                    for c in sin_endoso:
                        rf.write(f"- {c}\n")
                log(f"⚠️  {len(sin_endoso)} cliente(s) sin endoso → SIN ENDOSO.txt")

        # ── CASO D: PDFs sueltos ya nombrados (con o sin cédula) ──
        elif pdfs_sueltos_base:
            log(f"📄 {len(pdfs_sueltos_base)} PDFs sueltos")
            for nombre_archivo in sorted(pdfs_sueltos_base):
                src_pdf = os.path.join(base, nombre_archivo)
                orig_kb = os.path.getsize(src_pdf) // 1024
                # Si ya tiene " - CEDULA" en el nombre, usar tal cual
                # Strip prefijo numérico tipo Equisuiza: "1786475370771_NOMBRE - CEDULA.pdf"
                if re.search(r' - \d{7,}\.pdf$', nombre_archivo):
                    nombre_final = re.sub(r'^\d+_', '', nombre_archivo)
                else:
                    # Intentar extraer cédula del contenido
                    nombre, cedula = extraer_datos_pdf(src_pdf)
                    if nombre and cedula:
                        nombre_final = f"{nombre} - {cedula}.pdf"
                    else:
                        # Usar nombre del archivo quitando números entre paréntesis
                        nombre_final = re.sub(r'\s*\(\d+\)', '', nombre_archivo).strip()
                        log(f"  ⚠️ Sin datos en {nombre_archivo}")
                _guardar_comprimido(src_pdf, clientes_dir, nombre_final,
                                    orig_kb, base, gs_cmd, log)
                # Todos los PDFs sueltos también van a PRODUBANCO
                _guardar_comprimido(src_pdf, produbanco_dir_out, nombre_final,
                                    orig_kb, base, gs_cmd, log)
        else:
            log("⚠️  No se encontraron archivos para procesar.")

    finally:
        shutil.rmtree(tmp_work, ignore_errors=True)

    log("\n✅ ¡Listo! Los PDFs están en la carpeta POLIZA OUTPUT.")

    # ── Escribir reporte ──
    try:
        _fin = datetime.now()
        _duracion = int((_fin - _inicio).total_seconds())
        _ok    = sum(1 for l in _log_lines if l.strip().startswith("✓"))
        _warn  = sum(1 for l in _log_lines if "⚠" in l)
        _err   = sum(1 for l in _log_lines if l.strip().startswith("❌"))

        _c_files  = set(f for f in os.listdir(clientes_dir) if f.endswith(".pdf")) \
                    if os.path.isdir(clientes_dir) else set()
        _pb_files = set(f for f in os.listdir(produbanco_dir_out) if f.endswith(".pdf")) \
                    if os.path.isdir(produbanco_dir_out) else set()
        _clientes_out = len(_c_files)
        _pb_out       = len(_pb_files)
        _solo_clientes = sorted(_c_files - _pb_files)   # fueron a CLIENTES pero NO a PRODUBANCO

        # Log en pantalla: resumen de carpetas
        _log_orig(f"\n📊 RESUMEN FINAL:")
        _log_orig(f"   CLIENTES   : {_clientes_out} PDFs")
        _log_orig(f"   PRODUBANCO : {_pb_out} PDFs")
        if _solo_clientes:
            _log_orig(f"   Solo en CLIENTES ({len(_solo_clientes)}, sin póliza VH):")
            for f in _solo_clientes:
                _log_orig(f"     · {f}")
        if _sin_endoso_global:
            _log_orig(f"   En PRODUBANCO sin endoso ({len(_sin_endoso_global)}):")
            for c in _sin_endoso_global:
                _log_orig(f"     ⚠ {c}")

        reporte_path = os.path.join(output_dir, "REPORTE PROCESO.txt")
        with open(reporte_path, "w", encoding="utf-8") as rf:
            rf.write("=" * 55 + "\n")
            rf.write("   REPORTE DE PROCESO — ORGANIZADOR DE PÓLIZAS\n")
            rf.write("=" * 55 + "\n")
            rf.write(f"  Fecha       : {_inicio.strftime('%d/%m/%Y %H:%M:%S')}\n")
            rf.write(f"  Duración    : {_duracion} segundos\n")
            rf.write(f"  Carpeta     : {base}\n")
            rf.write(f"  Ghostscript : {'Sí (' + gs_cmd + ')' if gs_cmd else 'No (sin compresión)'}\n")
            rf.write("-" * 55 + "\n")
            rf.write("  RESUMEN\n")
            rf.write("-" * 55 + "\n")
            rf.write(f"  PDFs en CLIENTES              : {_clientes_out}\n")
            rf.write(f"  PDFs en PRODUBANCO            : {_pb_out}\n")
            rf.write(f"  Diferencia                    : {_clientes_out - _pb_out} (solo CLIENTES, sin póliza VH)\n")
            rf.write(f"  Archivos OK (✓)               : {_ok}\n")
            rf.write(f"  Advertencias (⚠)              : {_warn}\n")
            rf.write(f"  Errores (❌)                  : {_err}\n")
            # ── Solo en CLIENTES (no VH) ──
            rf.write("-" * 55 + "\n")
            if _solo_clientes:
                rf.write(f"  SOLO EN CLIENTES — sin póliza VH ({len(_solo_clientes)})\n")
                rf.write("-" * 55 + "\n")
                for f in _solo_clientes:
                    rf.write(f"  · {f}\n")
            else:
                rf.write("  SOLO EN CLIENTES              : ninguno\n")
            # ── Sin endoso ──
            rf.write("-" * 55 + "\n")
            if _sin_endoso_global:
                rf.write(f"  EN PRODUBANCO SIN ENDOSO ({len(_sin_endoso_global)})\n")
                rf.write("-" * 55 + "\n")
                for c in _sin_endoso_global:
                    rf.write(f"  ⚠  {c}\n")
            else:
                rf.write("  EN PRODUBANCO SIN ENDOSO      : ninguno\n")
            rf.write("=" * 55 + "\n")
            rf.write("  DETALLE DEL PROCESO\n")
            rf.write("=" * 55 + "\n\n")
            for linea in _log_lines:
                rf.write(linea + "\n")
        _log_orig(f"📄 Reporte guardado → POLIZA OUTPUT/REPORTE PROCESO.txt")
    except Exception as e:
        _log_orig(f"⚠️  No se pudo escribir el reporte: {e}")


def _guardar_titular(writer, base, output_dir, carpeta, gs_cmd, log):
    """Guarda el PDF combinado de un titular, extrae nombre/cédula y comprime. Devuelve nombre final."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False,
                                      dir=tempfile.gettempdir()) as tf:
        tmp_pdf = tf.name
    with open(tmp_pdf, "wb") as f:
        writer.write(f)

    orig_kb = os.path.getsize(tmp_pdf) // 1024
    nombre, cedula = extraer_datos_pdf(tmp_pdf)

    if nombre and cedula:
        nombre_archivo = f"{nombre} - {cedula}.pdf"
    else:
        nombre_archivo = f"{carpeta.replace('_', ' ')}.pdf"
        log(f"  ⚠️  Sin datos en {carpeta}, usando nombre de carpeta")

    _guardar_comprimido(tmp_pdf, output_dir, nombre_archivo, orig_kb, base, gs_cmd, log)
    os.remove(tmp_pdf)
    return nombre_archivo


def _guardar_comprimido(src_pdf, output_dir, nombre_archivo, orig_kb, base, gs_cmd, log):
    """Comprime (si hay GS) y guarda en output_dir."""
    src = src_pdf
    tmp_compressed = None

    if gs_cmd:
        tmp_compressed = comprimir_pdf(src_pdf, gs_cmd)
        if tmp_compressed:
            src = tmp_compressed

    comp_kb = os.path.getsize(src) // 1024
    reduccion = 100 - comp_kb * 100 // orig_kb if orig_kb > 0 else 0

    try:
        destino = guardar_pdf(src, output_dir, nombre_archivo)
        log(f"  ✓ {nombre_archivo} ({orig_kb} KB → {comp_kb} KB, {reduccion}% reducción)")
    except Exception as e:
        log(f"  ❌ Error guardando {nombre_archivo}: {e}")
    finally:
        if tmp_compressed and os.path.exists(tmp_compressed):
            os.remove(tmp_compressed)


# ──────────────────────────────────────────────
#  Interfaz gráfica
# ──────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de Pólizas")
        self.root.geometry("680x480")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        self.carpeta = StringVar(value="(ninguna seleccionada)")
        self.gs_cmd = detectar_ghostscript()

        # ── Título ──
        Label(root, text="Organizador de Pólizas", font=("Segoe UI", 16, "bold"),
              bg="#f0f0f0", fg="#1a1a2e").pack(pady=(20, 4))
        Label(root, text="Extrae · Combina · Comprime · Renombra",
              font=("Segoe UI", 9), bg="#f0f0f0", fg="#666").pack()

        # ── Selección de carpeta ──
        frame_carpeta = Frame(root, bg="#f0f0f0")
        frame_carpeta.pack(pady=14, padx=20, fill="x")

        Label(frame_carpeta, text="Carpeta:", font=("Segoe UI", 10),
              bg="#f0f0f0").pack(side="left")
        Label(frame_carpeta, textvariable=self.carpeta, font=("Segoe UI", 9),
              bg="#e8e8e8", fg="#333", padx=8, pady=4, relief="sunken",
              width=45, anchor="w").pack(side="left", padx=8)
        Button(frame_carpeta, text="Seleccionar", command=self.seleccionar_carpeta,
               font=("Segoe UI", 9), bg="#4a90d9", fg="white",
               relief="flat", padx=10, cursor="hand2").pack(side="left")

        # ── Estado GS ──
        gs_texto = f"✓ Ghostscript disponible ({self.gs_cmd})" if self.gs_cmd \
            else "⚠ Ghostscript no encontrado — se omitirá la compresión"
        gs_color = "#2e7d32" if self.gs_cmd else "#e65100"
        Label(root, text=gs_texto, font=("Segoe UI", 8),
              bg="#f0f0f0", fg=gs_color).pack()

        # ── Botón principal ──
        self.btn = Button(root, text="▶  Organizar Pólizas",
                          command=self.iniciar,
                          font=("Segoe UI", 12, "bold"),
                          bg="#2e7d32", fg="white", relief="flat",
                          padx=20, pady=10, cursor="hand2")
        self.btn.pack(pady=14)

        # ── Barra de progreso ──
        self.progress = ttk.Progressbar(root, mode="indeterminate", length=400)
        self.progress.pack()

        # ── Log ──
        frame_log = Frame(root, bg="#f0f0f0")
        frame_log.pack(padx=20, pady=10, fill="both", expand=True)

        scroll = Scrollbar(frame_log)
        scroll.pack(side="right", fill="y")
        self.log_box = Text(frame_log, height=10, font=("Consolas", 9),
                            bg="#1e1e1e", fg="#d4d4d4", relief="flat",
                            yscrollcommand=scroll.set, state="disabled")
        self.log_box.pack(fill="both", expand=True)
        scroll.config(command=self.log_box.yview)

        # Colores de log
        self.log_box.tag_config("ok", foreground="#4ec9b0")
        self.log_box.tag_config("err", foreground="#f44747")
        self.log_box.tag_config("info", foreground="#dcdcaa")

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta con las pólizas")
        if carpeta:
            self.carpeta.set(carpeta)

    def log(self, msg):
        tag = "ok" if msg.strip().startswith("✓") else \
              "err" if msg.strip().startswith("❌") else "info"
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n", tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.root.update_idletasks()

    def iniciar(self):
        carpeta = self.carpeta.get()
        if carpeta == "(ninguna seleccionada)" or not os.path.isdir(carpeta):
            messagebox.showwarning("Carpeta no válida",
                                   "Selecciona una carpeta antes de continuar.")
            return

        self.btn.configure(state="disabled", text="Procesando...")
        self.progress.start(10)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.log(f"📁 Carpeta: {carpeta}\n")

        def tarea():
            try:
                organizar(carpeta, self.log, self.gs_cmd)
            except Exception as e:
                self.log(f"❌ Error inesperado: {e}")
            finally:
                self.progress.stop()
                self.btn.configure(state="normal", text="▶  Organizar Pólizas")

        threading.Thread(target=tarea, daemon=True).start()


if __name__ == "__main__":
    root = Tk()
    app = App(root)
    root.mainloop()
