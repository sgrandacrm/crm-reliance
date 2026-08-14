# -*- coding: utf-8 -*-
"""
PROCESO 4 - ENVÍO DE PÓLIZAS
Crea borradores en Outlook para cada cliente de la carpeta CLIENTES.
"""
import sys, io, re
from pathlib import Path
import pandas as pd

BASE         = Path(__file__).parent.parent
DIR_POLIZAS  = BASE / "POLIZAS" / "salida"
DIR_OP       = BASE / "OPERACIONES" / "salida"
DIR_TARJETAS = BASE / "TARJETAS ASISTENCIA"
CAT_PATH     = BASE / "catalogos" / "CATALOGO CONTRATANTES.xlsx"

RAMO_CODIGO = {
    "VEHICULOS":              "VH",
    "INCENDIO":               "IN",
    "MULTIHOGAR":             "MH",
    "VIDA":                   "VD",
    "AP":                     "AP",
    "ACCIDENTES PERSONALES":  "AP",
    "DESGRAVAMEN":            "VD",
    "DESGRAVAMEN - VD":       "VD",
}

ASUNTO_RAMO = {
    "VH": "Renovación de Póliza de Seguro de Vehículo",
    "IN": "Renovación de Póliza de Seguro de Incendio",
    "MH": "Renovación de Póliza de Seguro de Hogar",
    "AP": "Renovación de Póliza de Seguro de Accidentes Personales",
    "VD": "Renovación de Póliza de Seguro de Vida",
}

FIJOS_SINIESTROS = [
    ("Siniestros Sierra", "Victor Trujillo",  "0984242511", "vtrujillo@reliance.ec"),
    ("Cobranzas Sierra",  "Daniela Peña",     "0979965842", "dpena@reliance.ec"),
    ("Cobranzas Costa",   "Andrea Veliz",     "0979965826", "aveliz@reliance.ec"),
    ("Siniestros Costa",  "Freddy Falcones",  "0996701405", "siniestrosgye@reliance.ec"),
]


# ─── helpers ─────────────────────────────────────────────────────────────────

def _ci_de_nombre(nombre_pdf):
    m = re.search(r'-\s*(\d{7,13})\.pdf$', nombre_pdf, re.IGNORECASE)
    return m.group(1).lstrip("0") if m else None

def _norm_ci(ci):
    return str(ci).strip().lstrip("0")

def _ramo_a_codigo(ramo_str):
    r = str(ramo_str).strip().upper()
    for k, v in RAMO_CODIGO.items():
        if k in r:
            return v
    return "VH"

def _prioridad_ramo(codigo):
    return {"VH": 0, "IN": 1, "MH": 2, "AP": 3, "VD": 4}.get(codigo, 99)

def _escanear_tarjetas():
    excluidos = {"QUIENES SOMOS"}
    resultado = []
    for f in sorted(DIR_TARJETAS.glob("*.pdf")):
        if any(e in f.name.upper() for e in excluidos):
            continue
        name = re.sub(r'(\.pdf)+$', '', f.name, flags=re.IGNORECASE)
        m = re.search(
            r'\s+((?:(?:VH|IN|MH|VD|AP|VIDA)(?:-(?:VH|IN|MH|VD|AP|VIDA))*))\s*$',
            name, re.IGNORECASE
        )
        if m:
            aseg   = name[:m.start()].strip().upper()
            codigos = [c.upper() for c in m.group(1).split('-')]
            if 'VIDA' in codigos:
                codigos = [c for c in codigos if c != 'VIDA'] + ['VD', 'AP']
            resultado.append((aseg, codigos, f))
    return resultado

def _match_tarjeta(tarjetas, aseg_cliente, codigo_ramo):
    aseg_cliente = aseg_cliente.strip().upper()
    for (t_aseg, t_codigos, t_path) in tarjetas:
        if codigo_ramo in t_codigos:
            if t_aseg in aseg_cliente or aseg_cliente in t_aseg:
                return t_path
    return None

def _buscar_tarjetas_cliente(tarjetas, filas_cliente):
    adjuntos = []
    for _, fila in filas_cliente.iterrows():
        aseg   = str(fila.get("ASEGURADORA_EMISORA", "") or "").strip().upper()
        codigo = _ramo_a_codigo(str(fila.get("RAMO", "") or ""))
        t = _match_tarjeta(tarjetas, aseg, codigo)
        if t and t not in adjuntos:
            adjuntos.append(t)
    return adjuntos

def _parsear_pago(num_cuotas_raw, cuenta_raw):
    s = str(num_cuotas_raw or "").upper()
    es_produbanco = "PRODUBANCO" in s
    es_debito     = "DEBITO" in s or "DÉBITO" in s
    if es_debito:
        m = re.search(r'\b(\d+)\b', s)
        n = m.group(1) if m else "N/A"
        cuenta = str(cuenta_raw or "").strip() if es_produbanco else "N/A"
        return "DÉBITO BANCARIO", cuenta or "N/A", n
    return "PAGO DIRECTO", "N/A", "N/A"

def _fmt_prima(v):
    try:
        return f"{float(str(v).replace(',', '.')):.2f}"
    except Exception:
        return str(v or "N/A").strip() or "N/A"

def _celular_wa(cel):
    return re.sub(r'[^\d]', '', str(cel or ""))

def _asunto(codigo_ramo, nombre):
    prefijo = ASUNTO_RAMO.get(codigo_ramo, "Renovación de Póliza de Seguro")
    return f"{prefijo} / {nombre}"

def _parrafo_poliza(fila_principal, codigo_ramo):
    poliza  = str(fila_principal.get("POLIZA_EMITIDA", "") or "").strip() or "N/A"
    if codigo_ramo == "VH":
        marca  = str(fila_principal.get("MARCA", "")          or "").strip()
        modelo = str(fila_principal.get("MODELO", "")         or "").strip()
        anio   = str(fila_principal.get("ANIO_VEHICULO", "")  or "").strip()
        placa  = str(fila_principal.get("PLACA", "")          or "").strip()
        return (
            f"Adjunto encontrará la póliza de su seguro de vehículo No. <strong>{poliza}</strong>, "
            f"emitida para su <strong>{marca} {modelo} {anio}</strong>, con placa <strong>{placa}</strong>. "
            f"Con esta cobertura, garantizamos su protección y tranquilidad en cada recorrido."
        )
    etiqueta = {"IN": "incendio", "MH": "hogar", "AP": "accidentes personales", "VD": "vida"}.get(codigo_ramo, "seguro")
    return (
        f"Adjunto encontrará la póliza de su seguro de {etiqueta} No. <strong>{poliza}</strong>. "
        f"Con esta cobertura, garantizamos su protección y tranquilidad."
    )

def _html_body(fila, codigo_ramo, forma_pago, num_cuenta, cuotas_str, ejec_info, ejec_nombre_raw, prima_total_sum):
    nombre   = str(fila.get("NOMBRE_CLIENTE", "") or "").strip()
    ci       = str(fila.get("CI", "")             or "").strip()
    celular  = str(fila.get("CELULAR_1", "")      or "").strip()
    correo   = str(fila.get("CORREO", "")         or "").strip()
    prestamo = str(fila.get("NUM_PRESTAMO", "")   or "").strip()
    prima    = prima_total_sum

    ejec_nombre = ejec_nombre_raw.title()
    ejec_mail   = ejec_info.get("mail", "")
    ejec_cel    = ejec_info.get("celular", "")
    ejec_wa     = _celular_wa(ejec_cel)
    p_poliza    = _parrafo_poliza(fila, codigo_ramo)

    filas_canales = "".join(
        f"<tr><td>{dep}</td><td>{ejec}</td><td>{cel}</td>"
        f"<td><a href='mailto:{mail}'>{mail}</a></td></tr>"
        for dep, ejec, cel, mail in FIJOS_SINIESTROS
    )

    return f"""<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">
<p>Señor(a) <strong>{nombre}</strong>,</p>
<p>Agradecemos su confianza en nosotros como su asesor de seguros.</p>
<p>{p_poliza}</p>
<p>Esta renovación corresponde al <strong>CRÉDITO No. {prestamo}</strong> otorgado por <strong>PRODUBANCO</strong>.
Es importante que mantenga su póliza vigente mientras su obligación crediticia siga activa.</p>
<p>Si tiene dudas o necesita asistencia, contáctenos a través de nuestros canales oficiales.</p>
<br>
<p><strong>Información del Cliente</strong></p>
<table border="1" cellpadding="5" cellspacing="0"
       style="border-collapse:collapse;width:60%;font-size:10pt;">
  <tr style="background:#f2f2f2;">
    <td><strong>Campo</strong></td><td><strong>Información</strong></td>
  </tr>
  <tr><td>CI</td><td>{ci}</td></tr>
  <tr><td>Nombre Cliente</td><td>{nombre}</td></tr>
  <tr><td>Celular</td><td>{celular}</td></tr>
  <tr><td>Correo Electrónico</td>
      <td><a href="mailto:{correo}">{correo}</a></td></tr>
  <tr><td>Forma de Pago</td><td>{forma_pago}</td></tr>
  <tr><td># de Cuenta</td><td>{num_cuenta}</td></tr>
  <tr><td>Cuotas</td><td>{cuotas_str}</td></tr>
  <tr><td>Monto Total</td><td>{prima}</td></tr>
</table>
<br>
<p><strong>Canales de Comunicación</strong></p>
<table border="1" cellpadding="5" cellspacing="0"
       style="border-collapse:collapse;width:70%;font-size:10pt;">
  <tr style="background:#f2f2f2;">
    <td><strong>Departamento</strong></td><td><strong>Ejecutivo</strong></td>
    <td><strong>Celular</strong></td><td><strong>Correo</strong></td>
  </tr>
  <tr>
    <td>Comercial</td><td>{ejec_nombre}</td><td>{ejec_cel}</td>
    <td><a href="mailto:{ejec_mail}">{ejec_mail}</a></td>
  </tr>
  {filas_canales}
</table>
<br>
<p>
  <strong>{ejec_nombre}</strong><br>
  Ejecutivo Comercial<br>
  Reliance S.A.<br><br>
  <a href="https://wa.me/{ejec_wa}">Contáctame por WhatsApp</a>
</p>
</div>"""


# ─── proceso principal ────────────────────────────────────────────────────────

def run():
    print("=" * 60)
    print("ENVÍO DE PÓLIZAS  |  Reliance Asesores de Seguros")
    print("=" * 60)
    print()

    # 1. Carpeta CLIENTES más reciente
    carpetas = sorted(
        [d for d in DIR_POLIZAS.iterdir() if d.is_dir() and "POLIZA OUTPUT" in d.name],
        reverse=True
    )
    if not carpetas:
        print("  ERROR: No hay carpetas POLIZA OUTPUT en POLIZAS/salida/")
        return

    clientes_dir = carpetas[0] / "CLIENTES"
    if not clientes_dir.exists():
        print(f"  ERROR: No existe CLIENTES en {carpetas[0].name}")
        return

    pdfs = sorted(clientes_dir.glob("*.pdf"))
    if not pdfs:
        print("  ERROR: No hay PDFs en CLIENTES/")
        return

    print(f"  Carpeta  : {carpetas[0].name}")
    print(f"  Clientes : {len(pdfs)} PDFs")
    print()

    # 2. UNIFICADO
    vh_files = sorted(DIR_OP.glob("UNIFICADO_VEHICULOS_*.xlsx"), reverse=True)
    ot_files = sorted(DIR_OP.glob("UNIFICADO_OTROS_*.xlsx"),     reverse=True)
    dfs = []
    if vh_files: dfs.append(pd.read_excel(vh_files[0], dtype=str))
    if ot_files: dfs.append(pd.read_excel(ot_files[0], dtype=str))
    if not dfs:
        print("  ERROR: No hay archivos UNIFICADO en OPERACIONES/salida/")
        return
    df_uni = pd.concat(dfs, ignore_index=True)
    df_uni["_CI"] = df_uni["CI"].astype(str).apply(_norm_ci)

    # 3. Ejecutivos
    df_ejec = pd.read_excel(CAT_PATH, sheet_name="EJECUTIVO CONTACTOS")
    ejec_map = {
        str(r["EJECUTIVO"]).strip().upper(): {
            "mail":    str(r["MAIL"]).strip(),
            "celular": str(r["CELULAR"]).strip(),
        }
        for _, r in df_ejec.iterrows()
    }

    # 4. Tarjetas
    tarjetas = _escanear_tarjetas()

    # 5. Outlook
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"  ERROR al conectar con Outlook: {e}")
        return

    ok, errores = 0, []

    for pdf_path in pdfs:
        try:
            ci = _ci_de_nombre(pdf_path.name)
            if not ci:
                errores.append(f"{pdf_path.name}: no se pudo extraer CI")
                continue

            filas = df_uni[df_uni["_CI"] == ci]
            if filas.empty:
                errores.append(f"{pdf_path.name}: CI {ci} no encontrado en UNIFICADO")
                continue


            # Fila principal (mayor prioridad de ramo)
            filas_ord = filas.copy()
            filas_ord["_PRI"] = filas_ord["RAMO"].apply(
                lambda r: _prioridad_ramo(_ramo_a_codigo(str(r or "")))
            )
            fila = filas_ord.sort_values("_PRI").iloc[0]

            ejec_raw  = str(fila.get("EJECUTIVO", "") or "").strip().upper()
            ejec_info = ejec_map.get(ejec_raw, {"mail": "", "celular": ""})
            correo    = str(fila.get("CORREO", "") or "").strip()

            if not correo:
                errores.append(f"{pdf_path.name}: sin correo en UNIFICADO")
                continue

            codigo_ramo = _ramo_a_codigo(str(fila.get("RAMO", "") or ""))
            forma_pago, num_cuenta, cuotas_str = _parsear_pago(
                fila.get("NUM_CUOTAS"), fila.get("CUENTA")
            )

            prima_sum = _fmt_prima(sum(
                float(str(r.get("PRIMA_TOTAL", "0") or "0").replace(",", "."))
                for _, r in filas.iterrows()
                if r.get("PRIMA_TOTAL") not in (None, "", "nan")
            ))
            html    = _html_body(fila, codigo_ramo, forma_pago, num_cuenta, cuotas_str, ejec_info, ejec_raw, prima_sum)
            asunto  = _asunto(codigo_ramo, str(fila.get("NOMBRE_CLIENTE", "") or "").strip())
            adjuntos_tarjeta = _buscar_tarjetas_cliente(tarjetas, filas)

            mail           = outlook.CreateItem(0)
            mail.To        = correo
            mail.CC        = ejec_info["mail"]
            mail.Subject   = asunto
            mail.HTMLBody  = html
            mail.Attachments.Add(str(pdf_path.resolve()))
            for t in adjuntos_tarjeta:
                mail.Attachments.Add(str(t.resolve()))
            mail.Save()

            nombre = str(fila.get("NOMBRE_CLIENTE", "") or "").strip()
            print(f"  ✓  {nombre[:40]:<40}  →  {correo}")
            ok += 1

        except Exception as e:
            errores.append(f"{pdf_path.name}: {e}")

    print()
    print(f"  Borradores creados : {ok}")
    if errores:
        print(f"  Errores            : {len(errores)}")
        for err in errores:
            print(f"    - {err}")
    print()
    print("  Revisa los borradores en Outlook antes de enviar.")


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    run()
