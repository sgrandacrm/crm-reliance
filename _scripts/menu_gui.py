# -*- coding: utf-8 -*-
"""
DBroker GUI — Reliance Asesores de Seguros
"""
import sys, io, os, json, queue, threading, importlib, winsound, ctypes, shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# ── Rutas base ─────────────────────────────────────────────────────────────────
BASE         = Path(__file__).parent.parent
HISTORY_FILE = Path(__file__).parent / "_history.json"
LOGS_DIR     = Path(__file__).parent / "logs"
BACKUPS_DIR  = BASE / "_backups"

# ── Paleta de marca ────────────────────────────────────────────────────────────
C_VERDE     = "#20584e"
C_VERDE_HV  = "#28786b"
C_VERDE_DIM = "#153d36"
C_DORADO    = "#b6985d"
C_BG        = "#111816"
C_CARD      = "#1a2624"
C_LOG_BG    = "#0b1513"
C_WHITE     = "#e8e8e8"
C_DIM       = "#7a9a95"
C_OK        = "#4cca8a"
C_ERR       = "#e05050"
C_WARN      = "#e0a030"

PROCESOS = [
    ("OPERACIONES", "Genera UNIFICADO, CONTRATANTES y MATRICES",  "proceso_operaciones"),
    ("COBRANZAS",   "Genera archivos de cobranza desde el CASH",   "proceso_cobranzas"),
    ("PÓLIZAS",     "Organiza, combina y comprime PDFs de pólizas","proceso_polizas"),
    ("ENVÍO",       "Crea borradores de correo en Outlook",        "proceso_envio_polizas"),
]

SALIDA_DIRS = {
    "proceso_operaciones": BASE / "OPERACIONES" / "salida",
    "proceso_cobranzas":   BASE / "COBRANZAS"   / "salida",
    "proceso_polizas":     BASE / "POLIZAS"      / "salida",
}

ENTRADA_CASH = {
    "proceso_operaciones": BASE / "OPERACIONES" / "entrada",
    "proceso_cobranzas":   BASE / "COBRANZAS"   / "entrada",
}


# ── Validaciones + advertencias ────────────────────────────────────────────────
def _validar(modulo: str) -> tuple:
    """Retorna (errores_bloqueantes, advertencias_con_confirmacion)."""
    errores, advertencias = [], []

    if modulo == "proceso_operaciones":
        entrada = ENTRADA_CASH[modulo]
        cash = list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx"))
        if not cash:
            errores.append("No hay archivos CASH en OPERACIONES/entrada/")
        elif len(cash) > 1:
            nombres = "\n".join(f"  • {f.name}" for f in cash)
            advertencias.append(
                f"Hay {len(cash)} archivos CASH en entrada/.\n"
                f"Se procesarán juntos:\n{nombres}\n\n¿Continuar?"
            )

    elif modulo == "proceso_cobranzas":
        entrada = ENTRADA_CASH[modulo]
        cash = list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx"))
        if not cash:
            errores.append("No hay archivos CASH en COBRANZAS/entrada/")
        elif len(cash) > 1:
            nombres = "\n".join(f"  • {f.name}" for f in cash)
            advertencias.append(
                f"Hay {len(cash)} archivos CASH en entrada/.\n"
                f"Se procesarán juntos:\n{nombres}\n\n¿Continuar?"
            )

    elif modulo == "proceso_polizas":
        entrada = BASE / "POLIZAS" / "entrada"
        if not list(entrada.rglob("*.pdf")) and not list(entrada.glob("*.zip")):
            errores.append("No hay PDFs ni ZIPs en POLIZAS/entrada/")

    elif modulo == "proceso_envio_polizas":
        op_sal = BASE / "OPERACIONES" / "salida"
        if not list(op_sal.glob("UNIFICADO_*.xlsx")):
            errores.append("No hay UNIFICADO en OPERACIONES/salida/\n→ Ejecuta OPERACIONES primero")
        carpetas = sorted((BASE / "POLIZAS" / "salida").glob("POLIZA OUTPUT*"), reverse=True)
        if not carpetas:
            errores.append("No hay POLIZA OUTPUT en POLIZAS/salida/\n→ Ejecuta PÓLIZAS primero")
        else:
            clientes = carpetas[0] / "CLIENTES"
            pdfs = list(clientes.glob("*.pdf")) if clientes.exists() else []
            if not pdfs:
                errores.append(f"No hay PDFs en {carpetas[0].name}/CLIENTES/")
            else:
                advertencias.append(
                    f"Se generarán {len(pdfs)} borradores de correo.\n¿Continuar?"
                )

    return errores, advertencias


# ── Backup automático del CASH ─────────────────────────────────────────────────
def _backup_cash(modulo: str):
    entrada = ENTRADA_CASH.get(modulo)
    if not entrada:
        return
    BACKUPS_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    for f in list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx")):
        shutil.copy2(f, BACKUPS_DIR / f"{ts}_{f.name}")


# ── Historial de ejecuciones ───────────────────────────────────────────────────
def _cargar_historial() -> dict:
    try:
        with open(HISTORY_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_historial(modulo: str, ok: bool):
    h = _cargar_historial()
    h[modulo] = {"fecha": datetime.now().strftime("%d/%m  %H:%M"), "ok": ok}
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(h, f, ensure_ascii=False)


# ── Resumen de archivos generados ──────────────────────────────────────────────
def _resumen_salida(modulo: str) -> str:
    folder = SALIDA_DIRS.get(modulo)
    if not folder or not folder.exists():
        return ""
    today = datetime.now().strftime("%Y-%m-%d")

    if modulo == "proceso_polizas":
        carpetas = [f for f in folder.iterdir() if today in f.name and f.is_dir()]
        if not carpetas:
            return ""
        clientes = carpetas[0] / "CLIENTES"
        pdfs = list(clientes.glob("*.pdf")) if clientes.exists() else []
        return f"  {len(pdfs)} PDFs combinados en CLIENTES/"

    items = [f for f in folder.iterdir() if today in f.name and f.is_file()]
    if not items:
        return ""
    n = len(items)
    return f"  {n} archivo{'s' if n != 1 else ''} generado{'s' if n != 1 else ''} en salida/"


# ── Streams auxiliares ─────────────────────────────────────────────────────────
class TextQueue(io.TextIOBase):
    def __init__(self, q, log_file=None):
        self.q = q
        self.log_file = log_file

    def write(self, s):
        if s:
            self.q.put(("txt", s))
            if self.log_file:
                try:
                    self.log_file.write(s)
                    self.log_file.flush()
                except Exception:
                    pass
        return len(s)

    def flush(self):
        pass


class AutoEnter(io.StringIO):
    """Stdin falso: cualquier input() retorna vacío de inmediato (simula Enter)."""
    def readline(self):
        return "\n"


# ── Aplicación principal ───────────────────────────────────────────────────────
class DBrokerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DBroker — Reliance Asesores de Seguros")
        self.configure(bg=C_BG)
        self.geometry("780x620")
        self.minsize(680, 520)
        self._q           = queue.Queue()
        self._busy        = False
        self._thread      = None
        self._cards       = []
        self._hist_lbls   = {}
        self._btn_cancel  = None
        self._lbl_abrir   = None
        self._build_ui()
        self._refrescar_historial()
        self._poll()

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=C_VERDE, pady=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="DBROKER  ·  RELIANCE ASESORES DE SEGUROS",
                 font=("Segoe UI", 13, "bold"), bg=C_VERDE, fg=C_DORADO).pack()

        # Grid 2×2 de procesos
        btn_area = tk.Frame(self, bg=C_BG, padx=20, pady=14)
        btn_area.pack(fill="x")
        btn_area.columnconfigure(0, weight=1)
        btn_area.columnconfigure(1, weight=1)

        for i, (nombre, desc, modulo) in enumerate(PROCESOS):
            wrap = tk.Frame(btn_area, bg=C_BG)
            wrap.grid(row=i // 2, column=i % 2, padx=10, pady=4, sticky="nsew")
            wrap.columnconfigure(0, weight=1)

            card = self._make_card(wrap, nombre, desc, modulo)
            card.grid(row=0, column=0, sticky="nsew")
            self._cards.append(card)

            hist_lbl = tk.Label(wrap, text="", font=("Segoe UI", 7),
                                bg=C_BG, fg=C_DIM, anchor="w")
            hist_lbl.grid(row=1, column=0, sticky="w", padx=4, pady=(2, 0))
            self._hist_lbls[modulo] = hist_lbl

        # Separador
        tk.Frame(self, bg=C_VERDE, height=2).pack(fill="x")

        # Log
        log_outer = tk.Frame(self, bg=C_CARD, padx=10, pady=6)
        log_outer.pack(fill="both", expand=True)

        log_hdr = tk.Frame(log_outer, bg=C_CARD)
        log_hdr.pack(fill="x", pady=(0, 4))
        tk.Label(log_hdr, text="▶  Salida del proceso",
                 font=("Segoe UI", 8, "bold"), bg=C_CARD, fg=C_DIM,
                 anchor="w").pack(side="left")
        self._lbl_abrir = tk.Label(log_hdr, text="",
                                   font=("Segoe UI", 8, "underline"),
                                   bg=C_CARD, fg=C_DORADO, cursor="hand2")
        self._lbl_abrir.pack(side="right", padx=(0, 4))

        log_row = tk.Frame(log_outer, bg=C_LOG_BG)
        log_row.pack(fill="both", expand=True)

        self._txt = tk.Text(log_row, bg=C_LOG_BG, fg=C_WHITE,
                            font=("Consolas", 9), relief="flat", bd=0,
                            state="disabled", wrap="word", insertbackground=C_WHITE)
        self._txt.tag_config("ok",   foreground=C_OK)
        self._txt.tag_config("err",  foreground=C_ERR)
        self._txt.tag_config("hdr",  foreground=C_DORADO, font=("Consolas", 9, "bold"))
        self._txt.tag_config("warn", foreground=C_WARN)

        sb = tk.Scrollbar(log_row, command=self._txt.yview, bg=C_CARD, troughcolor=C_BG)
        self._txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._txt.pack(side="left", fill="both", expand=True)

        # Pie
        foot = tk.Frame(self, bg=C_BG, pady=10, padx=20)
        foot.pack(fill="x")

        tk.Button(foot, text="Limpiar", command=self._limpiar,
                  bg=C_CARD, fg=C_DIM, relief="flat",
                  font=("Segoe UI", 9), padx=12, pady=5, cursor="hand2",
                  activebackground=C_VERDE, activeforeground=C_WHITE).pack(side="left")

        self._btn_cancel = tk.Button(foot, text="✕  Cancelar", command=self._cancelar,
                                     bg="#5a1a1a", fg="#ff9999", relief="flat",
                                     font=("Segoe UI", 9), padx=12, pady=5, cursor="hand2",
                                     activebackground="#7a2a2a", activeforeground=C_WHITE)
        # Se muestra solo mientras hay proceso corriendo

        tk.Button(foot, text="  Salir  ", command=self.destroy,
                  bg=C_VERDE, fg=C_WHITE, relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=5, cursor="hand2",
                  activebackground=C_VERDE_HV, activeforeground=C_WHITE).pack(side="right")

    def _make_card(self, parent, nombre, desc, modulo):
        card = tk.Frame(parent, bg=C_VERDE, cursor="hand2", pady=14, padx=16)
        lbl_n = tk.Label(card, text=nombre, font=("Segoe UI", 11, "bold"),
                         bg=C_VERDE, fg=C_DORADO)
        lbl_n.pack()
        lbl_d = tk.Label(card, text=desc, font=("Segoe UI", 8),
                         bg=C_VERDE, fg=C_WHITE, wraplength=300, justify="center")
        lbl_d.pack(pady=(4, 0))

        widgets = [card, lbl_n, lbl_d]

        def on_enter(_):
            if not self._busy:
                for w in widgets: w.config(bg=C_VERDE_HV)

        def on_leave(_):
            bg = C_VERDE_DIM if self._busy else C_VERDE
            for w in widgets: w.config(bg=bg)

        def on_click(_):
            self._launch(modulo, nombre)

        for w in widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)

        return card

    # ── Lanzar proceso ────────────────────────────────────────────────────────
    def _launch(self, modulo, nombre):
        if self._busy:
            return

        # Validación bloqueante
        errores, advertencias = _validar(modulo)
        if errores:
            messagebox.showwarning(f"{nombre} — Verificación previa",
                                   "\n".join(f"• {e}" for e in errores), parent=self)
            return

        # Confirmaciones (CASH duplicado, conteo de borradores, etc.)
        for adv in advertencias:
            if not messagebox.askyesno(f"{nombre} — Confirmación", adv, parent=self):
                return

        # Backup silencioso del CASH
        if modulo in ENTRADA_CASH:
            _backup_cash(modulo)
            self._log(f"  Backup CASH guardado en _backups/\n", "warn")

        self._lbl_abrir.config(text="")
        self._busy = True
        self._set_dim(True)
        self._btn_cancel.pack(side="left", padx=(8, 0))
        self._log(f"\n{'─' * 54}\n  {nombre}\n{'─' * 54}\n", "hdr")

        self._thread = threading.Thread(
            target=self._worker, args=(modulo, nombre), daemon=True
        )
        self._thread.start()

    # ── Worker (hilo) ─────────────────────────────────────────────────────────
    def _worker(self, modulo, nombre):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        LOGS_DIR.mkdir(exist_ok=True)
        ts       = datetime.now().strftime("%Y-%m-%d_%H-%M")
        log_path = LOGS_DIR / f"{ts}_{modulo}.txt"

        ok = False
        try:
            with open(log_path, "w", encoding="utf-8") as lf:
                stream                   = TextQueue(self._q, lf)
                old_out, old_err, old_in = sys.stdout, sys.stderr, sys.stdin
                sys.stdout = stream
                sys.stderr = stream
                sys.stdin  = AutoEnter()
                try:
                    mod = importlib.import_module(modulo)
                    importlib.reload(mod)
                    mod.run()
                    ok = True
                    self._q.put(("ok", f"\n✓  {nombre} completado.\n"))
                except SystemExit:
                    self._q.put(("warn", "\n⚠  Proceso cancelado por el usuario.\n"))
                except Exception as exc:
                    self._q.put(("err", f"\n✗  ERROR: {exc}\n"))
                finally:
                    sys.stdout = old_out
                    sys.stderr = old_err
                    sys.stdin  = old_in
        except Exception as exc:
            self._q.put(("err", f"\n✗  ERROR al abrir log: {exc}\n"))

        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

        _guardar_historial(modulo, ok)
        self._q.put(("done", (modulo, ok)))

    # ── Poll cola → UI ────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "done":
                    modulo, ok = data
                    self._busy  = False
                    self._thread = None
                    self._set_dim(False)
                    self._btn_cancel.pack_forget()
                    self._refrescar_historial()
                    self._notificar(ok)
                    if ok:
                        res = _resumen_salida(modulo)
                        if res:
                            self._log(res + "\n", "ok")
                        if modulo in SALIDA_DIRS:
                            self._mostrar_abrir(modulo)
                        if modulo in ENTRADA_CASH:
                            self.after(400, lambda m=modulo: self._ofrecer_limpieza(m))
                elif kind in ("ok", "err", "warn"):
                    self._log(data, kind)
                else:
                    self._log(data)
        except queue.Empty:
            pass
        self.after(80, self._poll)

    # ── Helpers UI ────────────────────────────────────────────────────────────
    def _log(self, txt, tag=""):
        self._txt.config(state="normal")
        if tag:
            self._txt.insert("end", txt, tag)
        else:
            self._txt.insert("end", txt)
        self._txt.see("end")
        self._txt.config(state="disabled")

    def _limpiar(self):
        self._txt.config(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.config(state="disabled")
        self._lbl_abrir.config(text="")

    def _cancelar(self):
        if not (self._thread and self._thread.is_alive()):
            return
        if not messagebox.askyesno(
            "Cancelar proceso",
            "¿Cancelar el proceso en curso?\nPueden quedar archivos incompletos.",
            parent=self, icon="warning",
        ):
            return
        try:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(
                ctypes.c_ulong(self._thread.ident),
                ctypes.py_object(SystemExit),
            )
        except Exception:
            pass

    def _set_dim(self, dim: bool):
        bg  = C_VERDE_DIM if dim else C_VERDE
        cur = "arrow"     if dim else "hand2"
        for card in self._cards:
            card.config(bg=bg, cursor=cur)
            for child in card.winfo_children():
                child.config(bg=bg)

    def _mostrar_abrir(self, modulo: str):
        folder = SALIDA_DIRS.get(modulo)
        if not folder:
            return
        self._lbl_abrir.config(text="▸ Abrir carpeta de resultados")
        self._lbl_abrir.bind("<Button-1>", lambda _: os.startfile(str(folder)))

    def _ofrecer_limpieza(self, modulo: str):
        entrada = ENTRADA_CASH.get(modulo)
        if not entrada:
            return
        cash = list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx"))
        if not cash:
            return
        nombres = "\n".join(f"  • {f.name}" for f in cash)
        if messagebox.askyesno(
            "Limpiar entrada",
            f"El CASH ya fue respaldado en _backups/\n"
            f"¿Eliminar de entrada/ los archivos procesados?\n\n{nombres}",
            parent=self,
        ):
            for f in cash:
                f.unlink()
            self._log("  Archivos eliminados de entrada/\n", "warn")

    def _refrescar_historial(self):
        h = _cargar_historial()
        for _, _, modulo in PROCESOS:
            lbl  = self._hist_lbls.get(modulo)
            info = h.get(modulo)
            if not lbl:
                continue
            if info:
                icono = "✓" if info["ok"] else "✗"
                color = C_OK  if info["ok"] else C_ERR
                lbl.config(text=f"{icono}  Última vez: {info['fecha']}", fg=color)
            else:
                lbl.config(text="Sin ejecuciones registradas", fg=C_DIM)

    def _notificar(self, ok: bool):
        try:
            winsound.MessageBeep(winsound.MB_OK if ok else winsound.MB_ICONHAND)
        except Exception:
            pass


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    DBrokerApp().mainloop()
