# -*- coding: utf-8 -*-
"""
DBroker GUI — Reliance Asesores de Seguros
"""
import sys, io, os, queue, threading, importlib
import tkinter as tk

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


class TextQueue(io.TextIOBase):
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("txt", s))
        return len(s)

    def flush(self):
        pass


class DBrokerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DBroker — Reliance Asesores de Seguros")
        self.configure(bg=C_BG)
        self.geometry("780x580")
        self.minsize(680, 480)
        self._q     = queue.Queue()
        self._busy  = False
        self._cards = []
        self._build_ui()
        self._poll()

    # ── Construcción de la UI ──────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=C_VERDE, pady=16)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text="DBROKER  ·  RELIANCE ASESORES DE SEGUROS",
            font=("Segoe UI", 13, "bold"),
            bg=C_VERDE, fg=C_DORADO,
        ).pack()

        # Grid 2×2 de procesos
        btn_area = tk.Frame(self, bg=C_BG, padx=20, pady=16)
        btn_area.pack(fill="x")
        btn_area.columnconfigure(0, weight=1)
        btn_area.columnconfigure(1, weight=1)

        for i, (nombre, desc, modulo) in enumerate(PROCESOS):
            card = self._make_card(btn_area, nombre, desc, modulo)
            card.grid(row=i // 2, column=i % 2, padx=10, pady=8, sticky="nsew")
            self._cards.append(card)

        # Separador
        tk.Frame(self, bg=C_VERDE, height=2).pack(fill="x")

        # Área de log
        log_outer = tk.Frame(self, bg=C_CARD, padx=10, pady=6)
        log_outer.pack(fill="both", expand=True)

        tk.Label(
            log_outer, text="▶  Salida del proceso",
            font=("Segoe UI", 8, "bold"),
            bg=C_CARD, fg=C_DIM, anchor="w",
        ).pack(fill="x", pady=(0, 4))

        log_row = tk.Frame(log_outer, bg=C_LOG_BG)
        log_row.pack(fill="both", expand=True)

        self._txt = tk.Text(
            log_row, bg=C_LOG_BG, fg=C_WHITE,
            font=("Consolas", 9), relief="flat", bd=0,
            state="disabled", wrap="word", insertbackground=C_WHITE,
        )
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

        tk.Button(
            foot, text="Limpiar", command=self._limpiar,
            bg=C_CARD, fg=C_DIM, relief="flat",
            font=("Segoe UI", 9), padx=12, pady=5, cursor="hand2",
            activebackground=C_VERDE, activeforeground=C_WHITE,
        ).pack(side="left")

        tk.Button(
            foot, text="  Salir  ", command=self.destroy,
            bg=C_VERDE, fg=C_WHITE, relief="flat",
            font=("Segoe UI", 9, "bold"), padx=12, pady=5, cursor="hand2",
            activebackground=C_VERDE_HV, activeforeground=C_WHITE,
        ).pack(side="right")

    def _make_card(self, parent, nombre, desc, modulo):
        card = tk.Frame(parent, bg=C_VERDE, cursor="hand2", pady=14, padx=16)
        lbl_n = tk.Label(
            card, text=nombre,
            font=("Segoe UI", 11, "bold"),
            bg=C_VERDE, fg=C_DORADO,
        )
        lbl_n.pack()
        lbl_d = tk.Label(
            card, text=desc,
            font=("Segoe UI", 8),
            bg=C_VERDE, fg=C_WHITE,
            wraplength=300, justify="center",
        )
        lbl_d.pack(pady=(4, 0))

        widgets = [card, lbl_n, lbl_d]

        def on_enter(_):
            if not self._busy:
                for w in widgets:
                    w.config(bg=C_VERDE_HV)

        def on_leave(_):
            bg = C_VERDE_DIM if self._busy else C_VERDE
            for w in widgets:
                w.config(bg=bg)

        def on_click(_):
            self._launch(modulo, nombre)

        for w in widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)

        return card

    # ── Ejecución de procesos ──────────────────────────────────────────────────
    def _launch(self, modulo, nombre):
        if self._busy:
            return
        self._busy = True
        self._set_dim(True)
        self._log(f"\n{'─' * 54}\n  {nombre}\n{'─' * 54}\n", "hdr")
        threading.Thread(target=self._worker, args=(modulo, nombre), daemon=True).start()

    def _worker(self, modulo, nombre):
        # Inicializar COM en el hilo para win32com (proceso ENVÍO)
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        stream = TextQueue(self._q)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = stream
        sys.stderr = stream
        try:
            mod = importlib.import_module(modulo)
            importlib.reload(mod)
            mod.run()
            self._q.put(("ok", f"\n✓  {nombre} completado.\n"))
        except Exception as exc:
            self._q.put(("err", f"\n✗  ERROR: {exc}\n"))
        finally:
            sys.stdout = old_out
            sys.stderr = old_err
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._q.put(("done", None))

    # ── Poll de cola → widget ──────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "done":
                    self._busy = False
                    self._set_dim(False)
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
        bg     = C_VERDE_DIM if dim else C_VERDE
        cursor = "arrow"    if dim else "hand2"
        for card in self._cards:
            card.config(bg=bg, cursor=cursor)
            for child in card.winfo_children():
                child.config(bg=bg)


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    DBrokerApp().mainloop()
