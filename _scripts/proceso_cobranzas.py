# -*- coding: utf-8 -*-
"""
PROCESO 2 - COBRANZAS
Genera el CONSOLIDADO DE CUOTAS INICIALES desde los archivos CASH del dia.
Regla de asignacion: COSTA -> RODOLFO MORENO / SIERRA -> HERNAN VELEZ
La asignacion es por cliente (todas las polizas del mismo CI van al mismo).
"""
import sys
import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from unificar import unificar

BASE            = Path(__file__).parent.parent
DIR_ENTRADA     = BASE / "COBRANZAS" / "entrada"
DIR_SALIDA      = BASE / "COBRANZAS" / "salida"
TABLA_PB_PATH   = BASE / "catalogos" / "TABLA CLIENTES CUENTAS PB.xlsx"

CUENTA_EMPRESA_PB = "06001645011"
BANCO_PB_CODE     = "0036"

COLS_CM = [
    "Orientacion de Transacción",
    "NUMERO DE CUENTA DE LA EMPRESA",
    "NUMERO SECUENCIAL",
    "NUMERO DE COMPROBANTE DE COBRO",
    "CONTRAPARTIDA: CODIGO DE DEUDOR",
    "MONEDA",
    "VALOR",
    "FORMA DE COBRO",
    "CODIGO DE BANCO",
    "TIPO DE CUENTA",
    "NUMERO DE CUENTA COBRO",
    "TIPO DE DOCUMENTO DEL DEUDOR",
    "NUMERO DE CEDULA DE DEUDOR",
    "NOMBRES DE DEUDOR",
    "DIRECCION DEUDOR",
    "CIUDAD DEUDOR",
    "TELEFONO DEUDOR",
    "LOCALIDAD DE COBRO",
    "REFERENCIA COBRO",
    "- -",
    "BASE IMPONIBLE - RETENCIÓN FUENTE (SERVICIOS)",
    "Codigo Retención Renta (SERVICIOS)",
    "BASE IMPONIBLE - RETENCIÓN FUENTE (BIENES)",
    "Codigo Retención Renta (BIENES)",
    "BASE IMPONIBLE - RETENCIÓN IVA (SERVICIOS)",
    "Codigo Retención IVA (SERVICIOS)",
    "BASE IMPONIBLE - RETENCIÓN IVA (BIENES)",
    "Codigo Retención IVA (BIENES)",
    "Base Iva 0%",
    "Base ICE",
    "COMPROMISO DE PAGO",
]

COLOR_HEADER      = "20584E"
WARN_FILL         = PatternFill("solid", fgColor="FFFF99")   # amarillo: factura invalida
CUENTA_PB_FILL    = PatternFill("solid", fgColor="CCE5FF")   # azul: cuenta del catalogo PB
MISSING_PB_FILL   = PatternFill("solid", fgColor="FF9999")   # rojo: cuenta PB faltante

_ASIGNACION = {
    "SIERRA": "HERNAN VELEZ",
    "COSTA":  "RODOLFO MORENO",
}

SUCURSAL_POR_CIUDAD = {
    "QUITO": "SIERRA", "AMBATO": "SIERRA", "LOJA": "SIERRA", "RIOBAMBA": "SIERRA",
    "CUENCA": "SIERRA", "IBARRA": "SIERRA", "LATACUNGA": "SIERRA", "TULCAN": "SIERRA",
    "SANTO DOMINGO DE LOS COLORADOS": "SIERRA",
    "GUAYAQUIL": "COSTA", "MANTA": "COSTA", "ESMERALDAS": "COSTA",
    "MACHALA": "COSTA", "PORTOVIEJO": "COSTA", "BAHIA DE CARAQUEZ": "COSTA",
}

COLS_SALIDA = [
    "ASIGNACION COBRANZAS",
    "FECHA DE REPORTE",
    "SUCURSAL",
    "IDENTIFICACION",
    "CLIENTE",
    "RAMO",
    "VIGENCIA",
    "CELULAR",
    "EMAIL",
    "ASEGURADORA",
    "VALOR CUOTA",
    "FORMA DE PAGO",
    "CUENTA",
    "COMPROMISO DE PAGO",
    "OBS_COMPROMISO DE PAGO",   # valor original del CASH, a la derecha de COMPROMISO
    "FECHA DE COBRO CF",
    "EJECUTIVO",
    "AREA",
    "ESTADO DE PAGO",
    "FECHA DE PAGO",
    "PRIMA NETA",
    "COMISIONES REPORTADA",
    "COMISION FACTURADA",
    "DIFERENCIA",
    "# DE PÓLIZA",
    "ANEXO",
    "FACTURA",
]

COLS_NUM   = ["VALOR CUOTA", "PRIMA NETA", "COMISIONES REPORTADA"]
COLS_TEXTO = [
    "IDENTIFICACION", "# DE PÓLIZA", "FACTURA",
    "CELULAR", "CUENTA", "ANEXO", "OBS_COMPROMISO DE PAGO",
]


# ---------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------
def _get(row, col):
    v = row.get(col)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none", "nat") else None


def _num(v):
    if v is None:
        return None
    try:
        return round(float(str(v).replace(",", ".").strip()), 2)
    except (ValueError, TypeError):
        return None


def _fecha(raw):
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None
    try:
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            return pd.to_datetime(s).date()
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


_MESES_ES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}


def _parsear_fecha_texto(s, anio):
    """
    Parsea texto en espanol como fecha usando el anio indicado.
    Soporta:
      'DD DE MES'        ->  '30 DE AGOSTO'      -> 30/08/anio
      'PRIMER DEB DD MES' -> 'PRIMER DEB 5 AGOSTO' -> 05/08/anio
    """
    su = s.strip().upper()

    # Patron: DD DE MES
    m = re.search(r"\b(\d{1,2})\s+DE\s+([A-Z]+)\b", su)
    if m:
        mes = _MESES_ES.get(m.group(2))
        if mes:
            try:
                return date(anio, mes, int(m.group(1)))
            except ValueError:
                pass

    # Patron: ...DEB DD MES  (cubre 'PRIMER DEB 5 AGOSTO', etc.)
    m = re.search(r"\bDEB\s+(\d{1,2})\s+([A-Z]+)\b", su)
    if m:
        mes = _MESES_ES.get(m.group(2))
        if mes:
            try:
                return date(anio, mes, int(m.group(1)))
            except ValueError:
                pass

    return None


def _compromiso(raw, hoy):
    """
    Procesa FECHA_COBRO_CI.
    Retorna (compromiso_date, obs_original).
    Prioridad:
      1. INMEDIAT*           -> hoy
      2. texto espanol       -> _parsear_fecha_texto
      3. parseo estandar     -> _fecha
      4. no reconocido       -> None (OBS conserva el texto)
    """
    if raw is None:
        return None, None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None, None
    obs = s

    if re.search(r"inmediat", s, re.IGNORECASE):
        return hoy, obs

    fecha_texto = _parsear_fecha_texto(s, hoy.year)
    if fecha_texto:
        return fecha_texto, obs

    return _fecha(s), obs


def _extraer_numero_cuenta(raw):
    """
    Extrae solo el numero de cuenta de un texto mixto.
    Ej: 'PB-AH-12014116060 (COBRO 15 C/MES)' -> '12014116060'
    Toma la secuencia de digitos mas larga (>= 8 caracteres).
    """
    if not raw:
        return ""
    secuencias = re.findall(r"\d+", str(raw))
    if not secuencias:
        return ""
    cuenta = max(secuencias, key=len)
    return cuenta if len(cuenta) >= 8 else ""


def _extraer_anexo(poliza_norm):
    """Extrae el ultimo segmento del numero de poliza normalizado como ANEXO."""
    if not poliza_norm:
        return None
    partes = str(poliza_norm).split("-")
    return partes[-1] if len(partes) >= 2 else None


def _motivo_factura(v):
    """Retorna motivo de advertencia SRI, o None si la factura es valida."""
    if v is None or str(v).strip() == "":
        return "FACTURA vacia"
    s = str(v).strip()
    if len(s) != 15:
        return f"FACTURA con {len(s)} digitos (SRI exige 15)"
    return None


def _region_cliente(row):
    r = (_get(row, "REGION") or "").strip().upper()
    if r in ("SIERRA", "COSTA"):
        return r
    ciudad = (_get(row, "CIUDAD") or "").strip().upper()
    return SUCURSAL_POR_CIUDAD.get(ciudad, "SIERRA")


# ---------------------------------------------------------------
# Formato openpyxl
# ---------------------------------------------------------------
def _aplicar_formato(wb):
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(name="Arial", bold=True, color="FFFFFF", size=8)
            c.fill = PatternFill("solid", fgColor=COLOR_HEADER)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=8)
                if isinstance(c.value, (datetime, date)):
                    c.number_format = "DD/MM/YYYY"


def _formato_decimal(ws, nombres):
    cab = [c.value for c in ws[1]]
    for nombre in nombres:
        if nombre in cab:
            idx = cab.index(nombre) + 1
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=idx)
                if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    cell.number_format = "0.00"


def _formato_texto(ws, nombres):
    cab = [c.value for c in ws[1]]
    for nombre in nombres:
        if nombre in cab:
            idx = cab.index(nombre) + 1
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=idx).number_format = "@"


def _colorear_cuenta_pb(ws, df_sheet):
    """Marca en azul toda la fila donde la cuenta fue completada desde catalogo PB."""
    for excel_row, (_, fila) in enumerate(df_sheet.iterrows(), start=2):
        if fila.get("_CUENTA_SRC") == "PB":
            for cell in ws[excel_row]:
                cell.fill = CUENTA_PB_FILL


def _colorear_pb_faltante(ws, df_sheet):
    """Marca en rojo las filas DEBITO PRODUBANCO sin numero de cuenta."""
    for excel_row, (_, fila) in enumerate(df_sheet.iterrows(), start=2):
        if fila.get("_CUENTA_FALTANTE"):
            for cell in ws[excel_row]:
                cell.fill = MISSING_PB_FILL


def _colorear_advertencias(ws, df_sheet):
    """Marca de amarillo las filas con factura invalida en la hoja."""
    for excel_row, (_, fila) in enumerate(df_sheet.iterrows(), start=2):
        if _motivo_factura(fila.get("FACTURA")) is not None:
            for cell in ws[excel_row]:
                cell.fill = WARN_FILL


# ---------------------------------------------------------------
# Catalogo TABLA CLIENTES CUENTAS PB
# ---------------------------------------------------------------
def _inferir_tipo_cuenta(cuenta):
    """01/02 -> CTE   |   06/12/20 -> AHO   |   otro -> ''"""
    s = str(cuenta).strip()
    if s.startswith(("01", "02")):
        return "CTE"
    if s.startswith(("06", "12", "20")):
        return "AHO"
    return ""


def _cargar_tabla_pb():
    """
    Carga el catalogo PB como dict {CI: cuenta_str}.
    Las cuentas se normalizan a string sin decimales.
    """
    if not TABLA_PB_PATH.exists():
        return {}
    try:
        df = pd.read_excel(TABLA_PB_PATH, sheet_name="Hoja1", dtype=str)
        tabla = {}
        for _, row in df.iterrows():
            ci = str(row.get("IDENTIFICACION") or "").strip()
            raw = str(row.get("CUENTA") or "").strip()
            if not ci or not raw or raw in ("nan", ""):
                continue
            # Normalizar: quitar parte decimal si viene como float string
            try:
                raw = str(int(float(raw)))
            except (ValueError, TypeError):
                pass
            tabla[ci] = raw
        return tabla
    except Exception as e:
        print(f"       AVISO: no se pudo cargar catalogo PB: {e}")
        return {}


def _actualizar_tabla_pb(df_out, tabla_pb, hoy):
    """
    Agrega al archivo PB los clientes de DEBITO PRODUBANCO que:
      - Tienen CUENTA definida en el CASH
      - No estan todavia en el catalogo
    Columna E = fecha de creacion del registro.
    Retorna cantidad de registros nuevos agregados.
    """
    if not TABLA_PB_PATH.exists():
        return 0

    mask_prod = df_out["FORMA DE PAGO"].str.upper().str.contains("PRODUBANCO", na=False)
    nuevos = []
    for _, row in df_out[mask_prod].iterrows():
        ci     = str(row.get("IDENTIFICACION") or "").strip()
        cuenta = str(row.get("CUENTA") or "").strip()
        if not ci or not cuenta or ci in tabla_pb:
            continue
        nuevos.append([
            ci,
            row.get("CLIENTE") or "",
            cuenta,
            _inferir_tipo_cuenta(cuenta),
            hoy,                            # col E: fecha de creacion
        ])
        tabla_pb[ci] = cuenta   # evitar duplicados dentro de la misma ejecucion

    if not nuevos:
        return 0

    wb_pb = load_workbook(TABLA_PB_PATH)
    ws_pb = wb_pb["Hoja1"]
    for fila in nuevos:
        ws_pb.append(fila)
    wb_pb.save(TABLA_PB_PATH)
    return len(nuevos)


# ---------------------------------------------------------------
# CASH MANAGEMENT PRODUBANCO
# ---------------------------------------------------------------
def _fmt_fecha_col(v):
    """Formatea date/datetime a DD/MM/YYYY; None o sin valor -> ''."""
    if v is None:
        return ""
    if isinstance(v, (date, datetime)):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none", "nat") else ""


def _extraer_total_cuotas(forma_pago):
    """Extrae el primer numero entero del texto de FORMA DE PAGO (total de cuotas)."""
    if not forma_pago:
        return "0"
    m = re.search(r'\b(\d+)\b', str(forma_pago))
    return m.group(1) if m else "0"


def _valor_centavos(v):
    """Convierte monto a centavos zero-padded a 13 digitos."""
    if v is None:
        return "0" * 13
    try:
        centavos = round(float(str(v).replace(",", ".")) * 100)
        return str(int(centavos)).zfill(13)
    except (ValueError, TypeError):
        return "0" * 13


def _referencia_pb(pol_original, total_cuotas):
    """
    Construye REFERENCIA COBRO: POLIZA-ANEXO_PB-1/TOTAL
    usando el valor original del CASH antes de normalizacion.
    Ej: '0394139-000425-0' -> '0394139-000425-1/9'
    """
    if not pol_original:
        return ""
    s = re.sub(r"[/\s]+", "-", str(pol_original).strip())
    partes = [p for p in s.split("-") if p]
    if len(partes) >= 2:
        poliza, anexo_pb = partes[0], partes[1]
    else:
        poliza, anexo_pb = (partes[0] if partes else ""), "0"
    return f"{poliza}-{anexo_pb}-1/{total_cuotas}"


def _tipo_doc(ci):
    """C = cedula (10 digitos), R = RUC (13 digitos)."""
    return "R" if len(str(ci).strip()) == 13 else "C"


def _generar_cash_management_pb(df_prod, hoy, out_path):
    """
    Genera CASH_MANAGEMENT_PB_<fecha>.xlsx con una hoja por aseguradora.
    Cada hoja lleva el nombre de la aseguradora y las 29 columnas del
    formato ProduBanco, pobladas desde los registros DEBITO PRODUBANCO.
    """
    if df_prod.empty:
        return 0

    aseguradoras = sorted(df_prod["ASEGURADORA"].fillna("SIN_ASEGURADORA").unique())

    total_filas = 0
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        for aseg in aseguradoras:
            df_aseg = df_prod[
                df_prod["ASEGURADORA"].fillna("SIN_ASEGURADORA") == aseg
            ].reset_index(drop=True)

            filas_cm = []
            for seq, (_, row) in enumerate(df_aseg.iterrows(), start=1):
                ci       = str(row.get("IDENTIFICACION") or "").strip()
                total_c  = _extraer_total_cuotas(row.get("FORMA DE PAGO"))
                pol_orig = str(row.get("_POL_ORIGINAL") or "").strip()

                filas_cm.append({
                    "Orientacion de Transacción":                     "CO",
                    "NUMERO DE CUENTA DE LA EMPRESA":                CUENTA_EMPRESA_PB,
                    "NUMERO SECUENCIAL":                             seq,
                    "NUMERO DE COMPROBANTE DE COBRO":                row.get("FACTURA") or "",
                    "CONTRAPARTIDA: CODIGO DE DEUDOR":               ci,
                    "MONEDA":                                        "USD",
                    "VALOR":                                         _valor_centavos(row.get("VALOR CUOTA")),
                    "FORMA DE COBRO":                                "CTA",
                    "CODIGO DE BANCO":                               BANCO_PB_CODE,
                    "TIPO DE CUENTA":                                _inferir_tipo_cuenta(row.get("CUENTA") or ""),
                    "NUMERO DE CUENTA COBRO":                        row.get("CUENTA") or "",
                    "TIPO DE DOCUMENTO DEL DEUDOR":                  _tipo_doc(ci),
                    "NUMERO DE CEDULA DE DEUDOR":                    ci,
                    "NOMBRES DE DEUDOR":                             row.get("CLIENTE") or "",
                    "DIRECCION DEUDOR":                              "",
                    "CIUDAD DEUDOR":                                 "",
                    "TELEFONO DEUDOR":                               row.get("CELULAR") or "",
                    "LOCALIDAD DE COBRO":                            "",
                    "REFERENCIA COBRO":                              _referencia_pb(pol_orig, total_c),
                    "- -":                                           "RELIANCE",
                    "BASE IMPONIBLE - RETENCIÓN FUENTE (SERVICIOS)": "00000000000000",
                    "Codigo Retención Renta (SERVICIOS)":            "NA",
                    "BASE IMPONIBLE - RETENCIÓN FUENTE (BIENES)":    "00000000000000",
                    "Codigo Retención Renta (BIENES)":               "NA",
                    "BASE IMPONIBLE - RETENCIÓN IVA (SERVICIOS)":    "00000000000000",
                    "Codigo Retención IVA (SERVICIOS)":              "NA",
                    "BASE IMPONIBLE - RETENCIÓN IVA (BIENES)":       "00000000000000",
                    "Codigo Retención IVA (BIENES)":                 "NA",
                    "Base Iva 0%":                                   "0",
                    "Base ICE":                                      "0",
                    "COMPROMISO DE PAGO":                            _fmt_fecha_col(row.get("COMPROMISO DE PAGO")),
                })

            df_cm = pd.DataFrame(filas_cm, columns=COLS_CM)
            sheet_name = str(aseg)[:31]   # Excel limita nombres de hoja a 31 chars
            df_cm.to_excel(xw, sheet_name=sheet_name, index=False)
            total_filas += len(df_cm)

    # Formato Arial 8 + cabecera verde
    wb = load_workbook(out_path)
    _aplicar_formato(wb)
    wb.save(out_path)
    return total_filas


# ---------------------------------------------------------------
# Proceso principal
# ---------------------------------------------------------------
def run():
    hoy       = date.today()
    fecha_str = hoy.strftime("%Y-%m-%d")

    print("=" * 60)
    print(f"COBRANZAS DBroker  |  {hoy.strftime('%d/%m/%Y')}")
    print("=" * 60)

    archivos_cash = sorted(DIR_ENTRADA.glob("*.xlsx"))
    if not archivos_cash:
        print("\nERROR: no se encontraron archivos .xlsx en la carpeta 'entrada'.")
        return

    print(f"\n[1/3] Archivos CASH en entrada: {len(archivos_cash)}")
    for f in archivos_cash:
        print(f"       - {f.name}")

    print("\n[2/3] Leyendo archivos CASH...")
    df, log_unif = unificar(archivos_cash, hoy=hoy)
    for linea in log_unif:
        print(f"       {linea}")

    if df.empty:
        print("\nERROR: no se pudo extraer ninguna fila.")
        return

    print(f"\n       Filas leidas      : {len(df)}")
    print(f"       Clientes unicos   : {df['CI'].nunique()}")

    # Cargar catalogo PB para lookup de cuentas
    tabla_pb = _cargar_tabla_pb()
    print(f"       Catalogo PB       : {len(tabla_pb)} registros cargados")

    print("\n[3/3] Generando COBRANZAS...")

    # --- Asignacion balanceada por cliente ---
    ci_region = {}
    for _, row in df.iterrows():
        ci = str(row.get("CI") or "").strip()
        if ci and ci not in ci_region:
            ci_region[ci] = _region_cliente(row)

    sierra_cis = sorted(ci for ci, r in ci_region.items() if r == "SIERRA")
    costa_cis  = sorted(ci for ci, r in ci_region.items() if r == "COSTA")

    total  = len(sierra_cis) + len(costa_cis)
    target = total // 2

    hernan_cis  = set(sierra_cis)
    rodolfo_cis = set(costa_cis)

    n_hernan_base  = len(hernan_cis)
    n_rodolfo_base = len(rodolfo_cis)

    movidos_a_rodolfo, movidos_a_hernan = [], []

    if n_hernan_base > target:
        n_mover = n_hernan_base - target
        for ci in sierra_cis[:n_mover]:
            hernan_cis.discard(ci)
            rodolfo_cis.add(ci)
            movidos_a_rodolfo.append(ci)
    elif n_rodolfo_base > target:
        n_mover = n_rodolfo_base - target
        for ci in costa_cis[:n_mover]:
            rodolfo_cis.discard(ci)
            hernan_cis.add(ci)
            movidos_a_hernan.append(ci)

    ci_asignado = {ci: "HERNAN VELEZ"   for ci in hernan_cis}
    ci_asignado.update({ci: "RODOLFO MORENO" for ci in rodolfo_cis})

    filas = []
    for _, row in df.iterrows():
        ci     = str(row.get("CI") or "").strip()
        region = ci_region.get(ci, "SIERRA")
        asig   = ci_asignado.get(ci, "HERNAN VELEZ")

        raw_comp               = _get(row, "FECHA_COBRO_CI")
        compromiso, obs_comp   = _compromiso(raw_comp, hoy)
        pol_norm               = _get(row, "POLIZA_EMITIDA") or ""
        if pol_norm and "-" not in pol_norm:
            pol_norm = pol_norm + "-0"

        # CUENTA: extraer solo el numero del CASH (filtra texto extra)
        forma_pago  = (_get(row, "NUM_CUOTAS") or "").upper()
        es_produbanco = "PRODUBANCO" in forma_pago
        cuenta = _extraer_numero_cuenta(_get(row, "CUENTA"))
        cuenta_src = "CASH"
        if not cuenta:
            cuenta_pb = tabla_pb.get(ci, "")
            if cuenta_pb:
                cuenta = cuenta_pb
                cuenta_src = "PB"
        cuenta_faltante = es_produbanco and not cuenta

        filas.append({
            "ASIGNACION COBRANZAS":   asig,
            "FECHA DE REPORTE":       hoy,
            "SUCURSAL":               region,
            "IDENTIFICACION":         ci,
            "CLIENTE":                _get(row, "NOMBRE_CLIENTE") or "",
            "RAMO":                   _get(row, "RAMO") or "",
            "VIGENCIA":               _fecha(_get(row, "VIGENCIA_DESDE")),
            "CELULAR":                _get(row, "CELULAR_1") or "",
            "EMAIL":                  _get(row, "CORREO") or "",
            "ASEGURADORA":            _get(row, "ASEGURADORA_EMISORA") or "",
            "VALOR CUOTA":            _num(_get(row, "CUOTA_INICIAL")),
            "FORMA DE PAGO":          _get(row, "NUM_CUOTAS") or "",
            "CUENTA":                 cuenta,
            "COMPROMISO DE PAGO":     compromiso,
            "OBS_COMPROMISO DE PAGO": obs_comp or "",
            "FECHA DE COBRO CF":      _get(row, "FECHA_COBRO_CF") or "",
            "EJECUTIVO":              _get(row, "EJECUTIVO") or "",
            "AREA":                   None,
            "ESTADO DE PAGO":         "IMPAGO",
            "FECHA DE PAGO":          None,
            "PRIMA NETA":             _num(_get(row, "PRIMA_NETA")),
            "COMISIONES REPORTADA":   _num(_get(row, "COMISION")),
            "COMISION FACTURADA":     None,
            "DIFERENCIA":             None,
            "# DE PÓLIZA":            pol_norm,
            "ANEXO":                  _extraer_anexo(pol_norm),
            "FACTURA":                _get(row, "FACTURA_ASEGURADORA") or "",
            "_CUENTA_SRC":            cuenta_src,
            "_CUENTA_FALTANTE":       cuenta_faltante,
            "_POL_ORIGINAL":          _get(row, "POLIZA_EMITIDA_ORIGINAL") or "",
        })

    df_out = pd.DataFrame(filas)

    ci_hernan  = len(hernan_cis)
    ci_rodolfo = len(rodolfo_cis)
    pol_hernan  = (df_out["ASIGNACION COBRANZAS"] == "HERNAN VELEZ").sum()
    pol_rodolfo = (df_out["ASIGNACION COBRANZAS"] == "RODOLFO MORENO").sum()

    print(f"       Total filas       : {len(df_out)}")
    print(f"       Total clientes    : {total}  (target por asignado: {target})")
    print(f"       HERNAN VELEZ      : {ci_hernan} clientes / {pol_hernan} polizas")
    print(f"       RODOLFO MORENO    : {ci_rodolfo} clientes / {pol_rodolfo} polizas")

    if movidos_a_rodolfo:
        print(f"\n       Reasignados SIERRA -> RODOLFO ({len(movidos_a_rodolfo)} clientes):")
        for ci in movidos_a_rodolfo:
            print(f"         CI {ci}")
    if movidos_a_hernan:
        print(f"\n       Reasignados COSTA -> HERNAN ({len(movidos_a_hernan)} clientes):")
        for ci in movidos_a_hernan:
            print(f"         CI {ci}")

    ramos_dist = df_out["RAMO"].value_counts()
    print(f"\n       Distribucion por ramo:")
    for ramo, cnt in ramos_dist.items():
        print(f"         {ramo}: {cnt}")

    # --- Sub-hojas por forma de pago ---
    mask_prod = df_out["FORMA DE PAGO"].str.upper().str.contains("PRODUBANCO", na=False)
    df_prod   = df_out[mask_prod].reset_index(drop=True)
    df_otras  = df_out[~mask_prod].reset_index(drop=True)

    # --- Hoja de advertencias: factura invalida + cuenta PB + cuenta faltante ---
    mask_fact     = df_out["FACTURA"].apply(lambda v: _motivo_factura(v) is not None)
    mask_pb       = df_out["_CUENTA_SRC"] == "PB"
    mask_faltante = df_out["_CUENTA_FALTANTE"] == True
    df_advert = df_out[mask_fact | mask_pb | mask_faltante].copy().reset_index(drop=True)

    def _motivo_row(row):
        partes = []
        m = _motivo_factura(row.get("FACTURA"))
        if m:
            partes.append(m)
        if row.get("_CUENTA_SRC") == "PB":
            partes.append("CUENTA COMPLETADA DESDE CATALOGO PB")
        if row.get("_CUENTA_FALTANTE"):
            partes.append("CUENTA DEBITO PRODUBANCO NO ENCONTRADA")
        return " | ".join(partes)

    df_advert["MOTIVO ADVERTENCIA"] = df_advert.apply(_motivo_row, axis=1)
    n_advert   = len(df_advert)

    # --- Hoja CUENTAS PB FALTANTES ---
    COLS_FALTANTES = ["IDENTIFICACION", "CLIENTE", "EJECUTIVO", "ASEGURADORA"]
    df_faltantes = (
        df_out[mask_faltante][COLS_FALTANTES]
        .drop_duplicates(subset="IDENTIFICACION")
        .reset_index(drop=True)
    )
    n_faltantes = len(df_faltantes)

    # Enriquecer catalogo PB con clientes nuevos
    n_nuevos_pb = _actualizar_tabla_pb(df_out, tabla_pb, hoy)

    print(f"\n       DEBITO PRODUBANCO : {len(df_prod)} polizas")
    print(f"       OTRAS FORMAS PAGO : {len(df_otras)} polizas")
    print(f"       Advertencias      : {n_advert} registros")
    print(f"       Cuentas PB falt.  : {n_faltantes} clientes sin cuenta")
    if n_nuevos_pb:
        print(f"       Catalogo PB +{n_nuevos_pb} clientes nuevos agregados")

    out_path = DIR_SALIDA / f"COBRANZAS_{fecha_str}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        df_out.to_excel(xw,   sheet_name="COBRANZAS",              index=False, columns=COLS_SALIDA)
        df_prod.to_excel(xw,  sheet_name="DEBITO PRODUBANCO",      index=False, columns=COLS_SALIDA)
        df_otras.to_excel(xw, sheet_name="OTRAS FORMAS DE PAGO",   index=False, columns=COLS_SALIDA)
        if n_advert > 0:
            df_advert.to_excel(xw, sheet_name="ADVERTENCIAS",         index=False, columns=COLS_SALIDA + ["MOTIVO ADVERTENCIA"])
        if n_faltantes > 0:
            df_faltantes.to_excel(xw, sheet_name="CUENTAS PB FALTANTES", index=False)

    wb = load_workbook(out_path)
    _aplicar_formato(wb)

    # Formato + color advertencias en las 3 hojas principales
    for nombre, df_sheet in [
        ("COBRANZAS",            df_out),
        ("DEBITO PRODUBANCO",    df_prod),
        ("OTRAS FORMAS DE PAGO", df_otras),
    ]:
        if nombre in wb.sheetnames:
            ws = wb[nombre]
            _formato_decimal(ws, COLS_NUM)
            _formato_texto(ws, COLS_TEXTO)
            _colorear_cuenta_pb(ws, df_sheet)     # azul: cuenta del catalogo PB
            _colorear_advertencias(ws, df_sheet)  # amarillo: factura invalida
            _colorear_pb_faltante(ws, df_sheet)   # rojo: cuenta PB no encontrada (maxima prioridad)

    if "ADVERTENCIAS" in wb.sheetnames:
        _formato_texto(wb["ADVERTENCIAS"], COLS_TEXTO)

    wb.save(out_path)

    hojas = "COBRANZAS | DEBITO PRODUBANCO | OTRAS FORMAS DE PAGO"
    if n_advert > 0:
        hojas += " | ADVERTENCIAS"
    if n_faltantes > 0:
        hojas += " | CUENTAS PB FALTANTES"
    print(f"\n       Guardado: salida/{out_path.name}")
    print(f"         Hojas: {hojas}")

    # Generar CASH MANAGEMENT PRODUBANCO
    cm_path = DIR_SALIDA / f"CASH_MANAGEMENT_PB_{fecha_str}.xlsx"
    n_cm = _generar_cash_management_pb(df_prod, hoy, cm_path)
    if n_cm > 0:
        n_aseg_cm = df_prod["ASEGURADORA"].nunique()
        print(f"\n       Cash Management PB: {n_cm} registros / {n_aseg_cm} aseguradora(s)")
        print(f"       Guardado: salida/{cm_path.name}")
    else:
        print("\n       Cash Management PB: sin registros DEBITO PRODUBANCO, archivo no generado")

    print("\n" + "=" * 60)
    print("PROCESO COBRANZAS COMPLETADO")
    print(f"  salida/{out_path.name}")
    if n_cm > 0:
        print(f"  salida/{cm_path.name}")
    print("=" * 60)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    run()
