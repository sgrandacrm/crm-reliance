# -*- coding: utf-8 -*-
"""
DBroker GUI — Reliance Asesores de Seguros
"""
import sys, io, os, json, queue, threading, importlib, winsound
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# ── Rutas base ─────────────────────────────────────────────────────────────────
BASE         = Path(__file__).parent.parent
HISTORY_FILE = Path(__file__).parent / "_history.json"

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

PROCESOS = [
    ("OPERACIONES", "Genera UNIFICADO, CONTRATANTES y MATRICES",  "proceso_operaciones"),
    ("COBRANZAS",   "Genera archivos de cobranza desde el CASH",   "proceso_cobranzas"),
    ("PÓLIZAS",     "Organiza, combina y comprime PDFs de pólizas","proceso_polizas"),
    ("ENVÍO",       "Crea borradores de correo en Outlook",        "proceso_envio_polizas"),
]


# ── Validaciones pre-proceso ───────────────────────────────────────────────────
def _validar(modulo: str) -> list:
    errores = []

    if modulo == "proceso_operaciones":
        entrada = BASE / "OPERACIONES" / "entrada"
        cash = list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx"))
        if not cash:
            errores.append("No hay archivos CASH en OPERACIONES/entrada/")

    elif modulo == "proceso_cobranzas":
        entrada = BASE / "COBRANZAS" / "entrada"
        cash = list(entrada.glob("*.xls")) + list(entrada.glob("*.xlsx"))
        if not cash:
            errores.append("No hay archivos CASH en COBRANZAS/entrada/")

    elif modulo == "proceso_polizas":
        entrada = BASE / "POLIZAS" / "entrada"
        archivos = list(entrada.rglob("*.pdf")) + list(entrada.glob("*.zip"))
        if not archivos:
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
            if not clientes.exists() or not list(clientes.glob("*.pdf")):
                errores.append(f"No hay PDFs en {carpetas[0].name}/CLIENTES/")

    return errores


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


# ── Streams auxiliares ────────────────────────────────────────────────────────
class TextQueue(io.TextIOBase):
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("txt", s))
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
        self.geometry("780x610")
        self.minsize(680, 500)
        self._q         = queue.Queue()
        self._busy      = False
        self._cards     = []
        self._hist_lbls = {}
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
        tk.Label(log_outer, text="▶  Salida del proceso",
                 font=("Segoe UI", 8, "bold"), bg=C_CARD, fg=C_DIM, anchor="w"
                 ).pack(fill="x", pady=(0, 4))

        log_row = tk.Frame(log_outer, bg=C_LOG_BG)
        log_row.pack(fill="both", expand=True)

        self._txt = tk.Text(log_row, bg=C_LOG_BG, fg=C_WHITE,
                            font=("Consolas", 9), relief="flat", bd=0,
                            state="disabled", wrap="word", insertbackground=C_WHITE)
        self._txt.tag_config("ok",  foreground=C_OK)
        self._txt.tag_config("err", foreground=C_ERR)
        self._txt.tag_config("hdr", foreground=C_DORADO, font=("Consolas", 9, "bold"))

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

        errores = _validar(modulo)
        if errores:
            messagebox.showwarning(
                f"{nombre} — Verificación previa",
                "\n".join(f"• {e}" for e in errores),
                parent=self,
            )
            return

        self._busy = True
        self._set_dim(True)
        self._log(f"\n{'─' * 54}\n  {nombre}\n{'─' * 54}\n", "hdr")
        threading.Thread(target=self._worker, args=(modulo, nombre), daemon=True).start()

    def _worker(self, modulo, nombre):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        stream = TextQueue(self._q)
        old_out, old_err, old_in = sys.stdout, sys.stderr, sys.stdin
        sys.stdout = stream
        sys.stderr = stream
        sys.stdin  = AutoEnter()
        ok = False
        try:
            mod = importlib.import_module(modulo)
            importlib.reload(mod)
            mod.run()
            ok = True
            self._q.put(("ok", f"\n✓  {nombre} completado.\n"))
        except Exception as exc:
            self._q.put(("err", f"\n✗  ERROR: {exc}\n"))
        finally:
            sys.stdout = old_out
            sys.stderr = old_err
            sys.stdin  = old_in
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            _guardar_historial(modulo, ok)
            self._q.put(("done", ok))

    # ── Poll cola → UI ────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "done":
                    self._busy = False
                    self._set_dim(False)
                    self._refrescar_historial()
                    self._notificar(data)
                elif kind in ("ok", "err"):
                    self._log(data, kind)
                else:
                    self._log(data)
        except queue.Empty:
            pass
        self.after(80, self._poll)

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

    def _set_dim(self, dim: bool):
        bg  = C_VERDE_DIM if dim else C_VERDE
        cur = "arrow"     if dim else "hand2"
        for card in self._cards:
            card.config(bg=bg, cursor=cur)
            for child in card.winfo_children():
                child.config(bg=bg)

    def _refrescar_historial(self):
        h = _cargar_historial()
        for _, _, modulo in PROCESOS:
            lbl  = self._hist_lbls.get(modulo)
            info = h.get(modulo)
            if not lbl:
                continue
            if info:
                icono = "✓" if info["ok"] else "✗"
                color = C_OK if info["ok"] else C_ERR
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
