# -*- coding: utf-8 -*-
"""
Modulo: gen_polizas_ap_vida.py
Genera el archivo POLIZAS AP / VIDA en formato DBroker (34 cols)
desde el DataFrame maestro filtrado a ramos AP y VIDA.
"""
import re
from datetime import date, datetime
from pathlib import Path
import pandas as pd

COLS_DBROKER_AP_VIDA = [
    "SECUENCIA", "RUC_CED", "NM_CONTRATANTE", "ITEM_OBJ", "DESCRIPCION",
    "FC_DESDE", "FC_HASTA", "NUM_OPERACION",
    "VAL_ASEGURADO", "TASA", "FACTOR", "PRIMA_NETA",
    "UBICACIÓN", "FORMA_PAGO", "OTRO_VALOR", "APLICA_IVA", "VAL_FINANCIAMIENTO",
    "DERECHOS_EMISION", "SUPER_BANCOS", "SEGURO_CAMPESINO",
    "PCT_IVA", "VALOR_IVA", "PRIMA_TOTAL",
    "NUM_CUOTAS", "PCT_CUOTA_INI",
    "POLIZA", "FACTURA", "ANEXO", "NOTAS_FP",
    "TIPO_POLIZA", "NUM_RENOVACION",
    "TIPO_COMISION",
    "COD_GRUPO_CONTRAT", "COD_CORTO_EJECUTIVO",
]

RAMOS_AP_VIDA = {"AP", "VIDA"}

SUCURSAL = {"SIERRA": "1", "COSTA": "2"}
SUCURSAL_POR_CIUDAD = {
    "QUITO": "1", "AMBATO": "1", "LOJA": "1", "RIOBAMBA": "1",
    "CUENCA": "1", "IBARRA": "1", "LATACUNGA": "1", "TULCAN": "1",
    "SANTO DOMINGO DE LOS COLORADOS": "1",
    "GUAYAQUIL": "2", "MANTA": "2", "ESMERALDAS": "2",
    "MACHALA": "2", "PORTOVIEJO": "2", "BAHIA DE CARAQUEZ": "2",
}

# Suma asegurada fija por aseguradora (aplica a AP y VIDA por igual)
_VAL_ASEG = {"LATINA": 20000, "MAPFRE": 5000, "ALIANZA": 5000}

_GARANTIA_DIRECTO = {"SIN CREDITO", "POST CREDITO", "SEGURO CHECK", "WEB", "RRSS", "NUEVO"}


def _cargar_catalogos_avida(catalogo_path: Path):
    cat_fp = pd.read_excel(catalogo_path, sheet_name="FORMAS DE PAGO",
                           engine="xlrd", dtype=str)
    cat_fp.columns = ["FORMA_CASH", "FORMA_DBROKER"]
    forma_cod = {
        r["FORMA_CASH"].strip().upper(): r["FORMA_DBROKER"].strip()
        for _, r in cat_fp.iterrows()
    }
    cat_ej = pd.read_excel(catalogo_path, sheet_name="EJECUTIVO",
                           engine="xlrd", dtype=str)
    cat_ej.columns = ["EJECUTIVO", "CODIGO_UIO", "CODIGO_GYE",
                      "COD_CORTO_UIO", "COD_CORTO_GYE"]
    ejecutivo_corto = {
        r["EJECUTIVO"].strip().upper(): {
            "1": r["COD_CORTO_UIO"].strip(),
            "2": r["COD_CORTO_GYE"].strip(),
        }
        for _, r in cat_ej.iterrows()
    }
    return forma_cod, ejecutivo_corto


def _get(row, col):
    v = row.get(col)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none") else None


def _num(v, decimals=2):
    if v is None:
        return None
    try:
        return round(float(str(v).replace("%", "").strip()), decimals)
    except (ValueError, TypeError):
        return None


def _fecha(raw):
    if raw is None:
        return None
    try:
        if pd.isna(raw):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    if isinstance(raw, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel
            d = from_excel(int(raw))
            return d.date() if isinstance(d, datetime) else d
        except Exception:
            return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


def _split_poliza(v):
    """POLIZA = string completo; ANEXO = parte tras el ultimo guion (string).
    Si contiene ';' (poliza VH + AP juntos), toma solo la parte de AP/VIDA.
    Sin guion: POLIZA queda como 'NUMERO-0' para mantener formato uniforme."""
    if not v:
        return "", "0"
    s = str(v).strip()
    if ";" in s:
        s = s.split(";")[-1].strip()
    idx = s.rfind("-")
    if idx >= 0:
        return s, s[idx + 1:] or "0"
    return s + "-0", "0"


def _cod_grupo(garantia):
    """COD_GRUPO_CONTRAT desde GARANTIA — misma logica que VH e INCENDIO."""
    if not garantia or str(garantia).strip() == "":
        return None, "GARANTIA vacia: COD_GRUPO_CONTRAT no determinado"
    g = str(garantia).strip().upper()
    if "DIRECTORES" in g:
        return "18", None
    if "FUNCIONARIOS" in g:
        return "20", None
    if "DIRECTO" in g:
        return "21", None
    for kw in _GARANTIA_DIRECTO:
        if kw in g:
            return "21", None
    return "14", None


def _tipo_poliza(tipo_raw):
    t = (tipo_raw or "").strip().upper()
    if "NUEVO" in t or t == "N":
        return "N", 0
    if "RENOV" in t or t == "R":
        return "R", 1
    return t[:1] if t else "N", 0


def _extraer_num_cuotas(raw):
    if not raw:
        return 1
    if str(raw).strip().upper().startswith("TC"):
        return 1
    m = re.search(r"\d+", str(raw))
    return int(m.group()) if m else 1


def _suc_key(region, ciudad):
    r = (region or "").strip().upper()
    if r in SUCURSAL:
        return SUCURSAL[r]
    return SUCURSAL_POR_CIUDAD.get((ciudad or "").strip().upper(), "1")


def _val_asegurado(aseguradora):
    k = (aseguradora or "").strip().upper()
    for nombre, val in _VAL_ASEG.items():
        if nombre in k:
            return val
    return None


def generar_polizas_ap_vida(df_otros: pd.DataFrame, catalogo_path: Path):
    """
    Filtra df_otros a ramos AP y VIDA.
    Retorna (df_ap, df_vida, df_alertas, ap_alertas_flags, vida_alertas_flags).
    """
    forma_cod, ejecutivo_corto = _cargar_catalogos_avida(catalogo_path)

    df_av = df_otros[
        df_otros["RAMO"].fillna("").str.strip().str.upper().isin(RAMOS_AP_VIDA)
    ].reset_index(drop=True)

    filas_ap, filas_vida = [], []
    alertas_rep = []
    ap_alertas_flags, vida_alertas_flags = [], []
    seq_ap = seq_vida = 1

    for _, r in df_av.iterrows():
        al = []

        ci        = _get(r, "CI") or ""
        nombre    = _get(r, "NOMBRE_CLIENTE") or ""
        ejecutivo = (_get(r, "EJECUTIVO") or "").upper()
        region    = _get(r, "REGION") or ""
        ciudad    = _get(r, "CIUDAD") or ""
        suc       = _suc_key(region, ciudad)
        ramo      = (_get(r, "RAMO") or "").strip().upper()

        tipo_raw = _get(r, "TIPO") or ""
        tipo_poliza, num_renov = _tipo_poliza(tipo_raw)
        if tipo_poliza not in ("N", "R"):
            al.append(f"TIPO_POLIZA: valor '{tipo_raw}' no reconocido")

        cuotas_raw = (_get(r, "NUM_CUOTAS") or "").strip().upper()
        forma_pago = forma_cod.get(cuotas_raw, "")
        if not forma_pago:
            al.append(f"FORMA_PAGO: '{cuotas_raw}' sin equivalente en catalogo")
        num_cuotas = _extraer_num_cuotas(cuotas_raw)

        poliza_full, anexo = _split_poliza(_get(r, "POLIZA_EMITIDA"))
        if not poliza_full:
            al.append("POLIZA_EMITIDA vacia")

        codigos_ej = ejecutivo_corto.get(ejecutivo)
        if codigos_ej:
            cod_corto = codigos_ej[suc]
        else:
            cod_corto = ""
            al.append(f"EJECUTIVO '{ejecutivo}' sin COD_CORTO en catalogo")

        cod_grupo, al_grupo = _cod_grupo(_get(r, "GARANTIA"))
        if al_grupo:
            al.append(al_grupo)

        aseguradora = _get(r, "ASEGURADORA_EMISORA") or ""
        val_aseg = _val_asegurado(aseguradora)
        if val_aseg is None:
            al.append(f"ASEGURADORA '{aseguradora}': VAL_ASEGURADO no definido")

        ubicacion = ciudad if ciudad else None
        if not ubicacion:
            al.append("CIUDAD vacia: UBICACION quedara en blanco")

        prima_total = _num(_get(r, "PRIMA_TOTAL"))
        de  = _num(_get(r, "DERECHOS_EMISION")) or 0
        sbs = _num(_get(r, "SUPER_BANCOS"))     or 0
        sc  = _num(_get(r, "SEGURO_CAMPESINO")) or 0
        prima_neta = _num(_get(r, "PRIMA_NETA"))
        if prima_neta is None and prima_total is not None:
            prima_neta = round(prima_total - de - sbs - sc, 2)

        seq = seq_ap if ramo == "AP" else seq_vida
        fila = {
            "SECUENCIA":           seq,
            "RUC_CED":             ci,
            "NM_CONTRATANTE":      nombre,
            "ITEM_OBJ":            1,
            "DESCRIPCION":         ramo,
            "FC_DESDE":            _fecha(r.get("VIGENCIA_DESDE")),
            "FC_HASTA":            _fecha(r.get("VIGENCIA_HASTA")),
            "NUM_OPERACION":       _get(r, "NUM_PRESTAMO"),
            "VAL_ASEGURADO":       val_aseg,
            "TASA":                0,
            "FACTOR":              100,
            "PRIMA_NETA":          prima_neta,
            "UBICACIÓN":           ubicacion,
            "FORMA_PAGO":          forma_pago,
            "OTRO_VALOR":          0,
            "APLICA_IVA":          "NO",
            "VAL_FINANCIAMIENTO":  0,
            "DERECHOS_EMISION":    de,
            "SUPER_BANCOS":        sbs,
            "SEGURO_CAMPESINO":    sc,
            "PCT_IVA":             0,
            "VALOR_IVA":           0,
            "PRIMA_TOTAL":         prima_total,
            "NUM_CUOTAS":          num_cuotas,
            "PCT_CUOTA_INI":       0,
            "POLIZA":              poliza_full,
            "FACTURA":             _get(r, "FACTURA_ASEGURADORA"),
            "ANEXO":               anexo,
            "NOTAS_FP":            None,
            "TIPO_POLIZA":         tipo_poliza,
            "NUM_RENOVACION":      num_renov,
            "TIPO_COMISION":       0,
            "COD_GRUPO_CONTRAT":   cod_grupo,
            "COD_CORTO_EJECUTIVO": cod_corto,
        }

        if ramo == "AP":
            filas_ap.append(fila)
            ap_alertas_flags.append(bool(al))
            seq_ap += 1
        else:
            filas_vida.append(fila)
            vida_alertas_flags.append(bool(al))
            seq_vida += 1

        if al:
            alertas_rep.append({
                "CI":        ci,
                "CLIENTE":   nombre,
                "EJECUTIVO": ejecutivo,
                "POLIZA":    poliza_full,
                "RAMO":      ramo,
                "ALERTAS":   " | ".join(al),
            })

    df_ap    = pd.DataFrame(filas_ap,   columns=COLS_DBROKER_AP_VIDA)
    df_vida  = pd.DataFrame(filas_vida, columns=COLS_DBROKER_AP_VIDA)
    df_alert = pd.DataFrame(alertas_rep)
    return df_ap, df_vida, df_alert, ap_alertas_flags, vida_alertas_flags
