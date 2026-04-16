"""
Generador de Layout de Dispersión Omonel  v1.1
Procesa archivos de nómina y cuentas vales para generar dispersión bancaria.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import sys
import datetime
import threading
import queue
import subprocess
import platform

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

APP_VERSION = "1.1"

# ─────────────────────────── Temas ────────────────────────────

THEMES = {
    "dark": dict(
        bg="#161C2C", surface="#1E2A3C", card="#243046",
        border="#334560", accent="#5088D8", accent2="#3DAFA0",
        text="#E0E8F4", muted="#7888A0", success="#42B254",
        error="#C8564D", warn="#D4992B", btn_fg="#FFFFFF",
        warn_fg="#1A1A1A", scrollbar="#334560", card_ok_bg="#183530",
    ),
    "light": dict(
        bg="#DDE2ED", surface="#FFFFFF", card="#EDF1FA",
        border="#A8B6CC", accent="#3D6CC0", accent2="#2E9888",
        text="#182038", muted="#4E607A", success="#1E8834",
        error="#B83838", warn="#B87800", btn_fg="#FFFFFF",
        warn_fg="#FFFFFF", scrollbar="#A8B6CC", card_ok_bg="#C8E4DC",
    ),
}

FONT_TITLE = ("Consolas", 22, "bold")
FONT_LABEL = ("Consolas", 10)
FONT_SMALL = ("Consolas", 9)
FONT_MONO  = ("Courier New", 9)
FONT_BTN   = ("Consolas", 11, "bold")

LOG_ICONS  = {"ok": "✓ ", "err": "✗ ", "warn": "⚠ ", "info": "ℹ ", "": "  "}

HELP_TEXT = """\
ESTRUCTURA ESPERADA DE ARCHIVOS

── Archivo People ──────────────────────────────
Columnas fijas (no son conceptos):
  CLAVE EMPLEADO · NOMBRE · AP. PATERNO · AP. MATERNO
  PUESTO · DEPARTAMENTO · SUCURSAL · TIPO DE PAGO
  FECHA INGRESO · NSS · RFC · CURP · EMISOR

Columnas de conceptos (dinámicas):
  Cualquier otra columna se trata como concepto.
  Ej: SUELDO, P2AH, BONO, etc.
  El importe es la SUMA de los conceptos seleccionados.

── Archivo Cuenta Vales ────────────────────────
Columnas requeridas:
  CLAVE EMPLEADO · CUENTA VALE

La unión se hace por CLAVE EMPLEADO (inner join).

── Formatos de dispersión TXT ──────────────────
Por Tarjeta  — Header 21 dígitos / Body 26 dígitos
  Header : cliente(7) + 0000(4) + total_centavos(10)
  Body   : cuenta_vale(16) + importe_centavos(10)

Por Empleado — Header 21 dígitos / Body 27 dígitos
  Header : empresa(7) + 0000(4) + total_centavos(10)
  Body   : departamento(7) + clave_empleado(10) + importe_centavos(10)
"""


# ─────────────────────────── Utilidades ────────────────────────────

def detect_system_theme() -> str:
    """Detecta el tema del sistema operativo."""
    try:
        if platform.system() == "Darwin":
            r = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True, text=True, timeout=2,
            )
            return "dark" if "Dark" in r.stdout else "light"
        elif platform.system() == "Windows":
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if val == 1 else "dark"
    except Exception:
        pass
    return "dark"


def validate_excel_file(path: str, header_row: int) -> dict:
    """Valida un archivo Excel y retorna metadata básica."""
    try:
        df = pd.read_excel(path, dtype=str, header=header_row - 1)
        orig_cols = [str(c) for c in df.columns]
        norm_cols = [c.strip().lower().replace(" ", "_") for c in orig_cols]
        has_key   = "clave_empleado" in norm_cols
        # Columnas que no son fijas → son conceptos; solo incluir las que tienen algún valor > 0
        concept_cols = []
        for orig, norm in zip(orig_cols, norm_cols):
            if norm not in FIXED_PEOPLE_COLS:
                col_vals = pd.to_numeric(df[orig], errors="coerce")
                if (col_vals > 0).any():
                    concept_cols.append(orig)
        return {
            "ok":           True,
            "rows":         len(df),
            "has_key":      has_key,
            "concept_cols": concept_cols,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def open_file(path: str):
    """Abre un archivo con la aplicación predeterminada del sistema."""
    try:
        if platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        elif platform.system() == "Windows":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")


# ─────────────────────────── Tooltip ────────────────────────────

class Tooltip:
    def __init__(self, widget, text: str):
        self._widget = widget
        self._text   = text
        self._win    = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, event=None):
        if self._win:
            return
        x = self._widget.winfo_rootx() + 10
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 6
        self._win = tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self._text, bg="#2D3748", fg="#FFFFFF",
            font=FONT_SMALL, relief="flat", bd=1,
            padx=8, pady=5, wraplength=340, justify="left",
        ).pack()

    def _hide(self, event=None):
        if self._win:
            self._win.destroy()
            self._win = None


# ─────────────────────────── TagInput ────────────────────────────

class TagInput(tk.Frame):
    """
    Combobox multi-select con dropdown filtrable.
    Chips de tags seleccionados inline; dropdown tipo label-picker al enfocar.
    """

    def __init__(self, parent, T: dict, initial_tags: list | None = None, **kw):
        super().__init__(parent, bg=T["card"], **kw)
        self._T                = T
        self.tags: list[str]         = list(initial_tags or [])
        self._suggestions: list[str] = []
        self._dropdown: tk.Toplevel | None = None
        self._outside_click_id: str | None = None
        self._build()

    # ── Construcción ──────────────────────────────────────────────

    def _build(self):
        T = self._T

        # Caja principal (chips + entry + botones)
        self._box = tk.Frame(
            self, bg=T["surface"],
            highlightthickness=1, highlightbackground=T["border"],
        )
        self._box.pack(fill="x", padx=6, pady=6)

        # Zona izquierda: chips + entry
        self._inner = tk.Frame(self._box, bg=T["surface"])
        self._inner.pack(side="left", fill="x", expand=True, padx=4, pady=4)

        # Entry (siempre al final de los chips)
        self._entry = tk.Entry(
            self._inner, bg=T["surface"], fg=T["text"],
            insertbackground=T["accent"],
            relief="flat", bd=0, font=FONT_MONO,
            highlightthickness=0,
        )
        self._entry.pack(side="left", fill="x", expand=True, ipady=2)
        self._entry.bind("<KeyRelease>", self._on_keyrelease)
        self._entry.bind("<FocusIn>",    self._on_focus_in)
        self._entry.bind("<FocusOut>",   lambda e: self.after(150, self._maybe_hide))
        self._entry.bind("<Return>",     self._on_return)
        self._entry.bind("<Escape>",     lambda e: self._hide_dropdown())
        self._entry.bind("<BackSpace>",  self._on_backspace)

        # Botones derecha
        btn_box = tk.Frame(self._box, bg=T["surface"])
        btn_box.pack(side="right", padx=(0, 4))
        tk.Button(btn_box, text="⊗", command=self._clear_all,
                  bg=T["surface"], fg=T["muted"], font=("Consolas", 11),
                  relief="flat", bd=0, cursor="hand2").pack(side="left")
        tk.Button(btn_box, text="⌄", command=self._toggle_dropdown,
                  bg=T["surface"], fg=T["muted"], font=("Consolas", 11),
                  relief="flat", bd=0, cursor="hand2").pack(side="left")

        # Hint
        tk.Label(self,
                 text="Escribe o despliega para filtrar  ·  Enter para agregar manual  "
                      "·  sin selección = todos los conceptos",
                 bg=T["card"], fg=T["muted"], font=FONT_SMALL,
                 ).pack(anchor="w", padx=6, pady=(0, 4))

        self._rebuild_chips()

    # ── Chips ─────────────────────────────────────────────────────

    def _rebuild_chips(self):
        T = self._T
        for w in self._inner.winfo_children():
            if w is not self._entry:
                w.destroy()
        for tag in self.tags:
            chip = tk.Frame(self._inner, bg=T["accent"])
            chip.pack(side="left", padx=(0, 3), pady=1)
            tk.Label(chip, text=tag, bg=T["accent"], fg=T["btn_fg"],
                     font=FONT_SMALL, padx=5, pady=1).pack(side="left")
            tk.Button(chip, text="×", command=lambda t=tag: self._remove_tag(t),
                      bg=T["accent"], fg=T["btn_fg"],
                      font=("Consolas", 9), relief="flat", bd=0,
                      cursor="hand2", padx=3).pack(side="left")
        # Mantener entry al final
        self._entry.pack_forget()
        self._entry.pack(side="left", fill="x", expand=True, ipady=2)

    def _remove_tag(self, tag):
        if tag in self.tags:
            self.tags.remove(tag)
        self._rebuild_chips()
        self._repopulate()

    def _clear_all(self):
        self.tags.clear()
        self._rebuild_chips()
        self._repopulate()

    # ── Entrada de teclado ────────────────────────────────────────

    def _on_return(self, event=None):
        text = self._entry.get().strip().upper()
        if text:
            if text not in self.tags:
                self.tags.append(text)
            self._entry.delete(0, "end")
            self._rebuild_chips()
            self._repopulate()

    def _on_backspace(self, event=None):
        if not self._entry.get() and self.tags:
            self.tags.pop()
            self._rebuild_chips()
            self._repopulate()

    def _on_keyrelease(self, event=None):
        if event and event.keysym in ("Return", "Escape", "Tab",
                                       "Up", "Down", "Left", "Right"):
            return
        self._open_or_repopulate()

    def _on_focus_in(self, event=None):
        T = self._T
        self._box.configure(highlightbackground=T["accent"])
        self._open_or_repopulate()

    # ── Dropdown ──────────────────────────────────────────────────

    def _toggle_dropdown(self):
        if self._dropdown and self._dropdown.winfo_exists():
            self._hide_dropdown()
        else:
            self._entry.focus_set()

    def _open_or_repopulate(self):
        if not self._suggestions:
            return
        if self._dropdown and self._dropdown.winfo_exists():
            self._repopulate()
        else:
            self._open_dropdown()

    def _on_outside_click(self, event):
        """Cierra el dropdown si el clic fue fuera del TagInput o del dropdown."""
        w = event.widget
        while w is not None:
            if w is self or w is self._box or \
               (self._dropdown and self._dropdown.winfo_exists() and w is self._dropdown):
                return
            try:
                w = w.master
            except Exception:
                break
        self._hide_dropdown()

    def _open_dropdown(self):
        if not self._suggestions:
            return
        self.update_idletasks()
        T = self._T

        self._dropdown = dd = tk.Toplevel(self)
        dd.wm_overrideredirect(True)
        dd.configure(bg=T["border"])
        dd.bind("<FocusOut>", lambda e: self.after(150, self._maybe_hide))

        outer = tk.Frame(dd, bg=T["surface"])
        outer.pack(fill="both", expand=True, padx=1, pady=1)

        # Header
        hdr = tk.Frame(outer, bg=T["surface"])
        hdr.pack(fill="x", padx=8, pady=(6, 4))
        self._hdr_lbl = tk.Label(hdr, text="All labels",
                                  bg=T["surface"], fg=T["muted"], font=FONT_SMALL)
        self._hdr_lbl.pack(side="left")

        tk.Frame(outer, bg=T["border"], height=1).pack(fill="x")

        # Lista con scrollbar
        list_container = tk.Frame(outer, bg=T["surface"])
        list_container.pack(fill="both", expand=True)

        self._canvas_dd = tk.Canvas(list_container, bg=T["surface"],
                                     highlightthickness=0)
        vsb = tk.Scrollbar(list_container, orient="vertical",
                            command=self._canvas_dd.yview)
        self._canvas_dd.configure(yscrollcommand=vsb.set)

        self._list_frame = tk.Frame(self._canvas_dd, bg=T["surface"])
        self._list_win   = self._canvas_dd.create_window(
            (0, 0), window=self._list_frame, anchor="nw")

        self._list_frame.bind("<Configure>", lambda e: (
            self._canvas_dd.configure(scrollregion=self._canvas_dd.bbox("all")),
            self._canvas_dd.itemconfig(self._list_win,
                                        width=self._canvas_dd.winfo_width()),
        ))
        self._canvas_dd.bind("<Configure>", lambda e: self._canvas_dd.itemconfig(
            self._list_win, width=e.width))

        self._canvas_dd.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._repopulate()
        self._reposition()
        # Registrar handler global para cerrar al clicar fuera
        if not self._outside_click_id:
            self._outside_click_id = self.winfo_toplevel().bind(
                "<Button-1>", self._on_outside_click, add="+"
            )

    def _repopulate(self):
        if not self._dropdown or not self._dropdown.winfo_exists():
            return
        T          = self._T
        query      = self._entry.get().strip().upper()
        tags_upper = {t.upper() for t in self.tags}

        # Filtrar
        visible = [s for s in self._suggestions
                   if not query or query in s.upper()]

        # Header dinámico
        self._hdr_lbl.config(
            text=f"All labels  ({len(visible)})" if not query
            else f'"{self._entry.get().strip()}"  —  {len(visible)} resultado(s)'
        )

        for w in self._list_frame.winfo_children():
            w.destroy()

        if not visible:
            tk.Label(self._list_frame, text="Sin resultados",
                     bg=T["surface"], fg=T["muted"], font=FONT_SMALL,
                     padx=12, pady=8).pack(anchor="w")
        else:
            for item in visible:
                selected = item.upper() in tags_upper
                self._make_row(item, selected, T)

        self._reposition()

    def _make_row(self, item: str, selected: bool, T: dict):
        bg_normal  = T["surface"]
        bg_hover   = T["card"]
        fg_lbl     = T["muted"] if selected else T["text"]
        bar_color  = T["accent"] if selected else bg_normal

        row = tk.Frame(self._list_frame, bg=bg_normal, cursor="hand2")
        row.pack(fill="x")

        # Barra lateral de selección
        bar = tk.Frame(row, bg=bar_color, width=3)
        bar.pack(side="left", fill="y")

        lbl = tk.Label(row, text=item, bg=bg_normal, fg=fg_lbl,
                       font=FONT_SMALL, anchor="w", padx=10, pady=6)
        lbl.pack(side="left", fill="x", expand=True)

        if not selected:
            def _enter(e, r=row, l=lbl, b=bar):
                r.configure(bg=bg_hover)
                l.configure(bg=bg_hover)
                b.configure(bg=bg_hover)

            def _leave(e, r=row, l=lbl, b=bar):
                r.configure(bg=bg_normal)
                l.configure(bg=bg_normal)
                b.configure(bg=bg_normal)

            def _click(e=None, i=item):
                self._pick(i)

            row.bind("<Enter>", _enter); row.bind("<Leave>", _leave)
            lbl.bind("<Enter>", _enter); lbl.bind("<Leave>", _leave)
            row.bind("<Button-1>", _click)
            lbl.bind("<Button-1>", _click)

    def _pick(self, item: str):
        tag = item.strip().upper()
        if tag and tag not in self.tags:
            self.tags.append(tag)
        self._entry.delete(0, "end")
        self._rebuild_chips()
        self._hide_dropdown()
        self.winfo_toplevel().focus_set()

    def _reposition(self):
        if not self._dropdown or not self._dropdown.winfo_exists():
            return
        self.update_idletasks()
        x = self._box.winfo_rootx()
        y = self._box.winfo_rooty() + self._box.winfo_height()
        w = self._box.winfo_width()
        # Altura máxima: min(250, contenido real)
        self._list_frame.update_idletasks()
        content_h = self._list_frame.winfo_reqheight()
        h = min(250, content_h) + 42   # 42 ≈ header + separador
        self._canvas_dd.configure(height=max(h - 42, 40))
        self._dropdown.geometry(f"{w}x{h}+{x}+{y}")

    def _hide_dropdown(self):
        T = self._T
        self._box.configure(highlightbackground=T["border"])
        if self._dropdown and self._dropdown.winfo_exists():
            self._dropdown.destroy()
        self._dropdown = None
        # Desregistrar el handler global de clic
        if self._outside_click_id:
            try:
                self.winfo_toplevel().unbind("<Button-1>", self._outside_click_id)
            except Exception:
                pass
            self._outside_click_id = None

    def _maybe_hide(self):
        try:
            focused = str(self.focus_get())
            if focused == str(self._entry):
                return
            if self._dropdown and self._dropdown.winfo_exists():
                if focused.startswith(str(self._dropdown)):
                    return
        except Exception:
            pass
        self._hide_dropdown()

    # ── API pública ───────────────────────────────────────────────

    def set_suggestions(self, concepts: list[str]):
        """Recibe los conceptos detectados en el archivo People."""
        self._suggestions = list(concepts)
        self._repopulate()

    def get_tags(self) -> list[str]:
        return list(self.tags)


# ─────────────────────────── FilePickerCard ────────────────────────────

class FilePickerCard(tk.Frame):
    """Tarjeta grande de selección de archivo xlsx con diseño centrado."""

    def __init__(self, parent, label: str, icon_idle: str, icon_ok: str,
                 T: dict, path_var: tk.StringVar,
                 header_row_var: tk.StringVar | None = None,
                 on_concepts_found=None, **kw):
        super().__init__(parent, bg=T["bg"], height=190, **kw)
        self.pack_propagate(False)
        self._T                 = T
        self._label             = label
        self._icon_idle         = icon_idle
        self._icon_ok           = icon_ok
        self.path_var           = path_var
        self._header_row_var    = header_row_var
        self._on_concepts_found = on_concepts_found
        self._loaded            = False
        self._build()

    def _build(self):
        T  = self._T
        bg = T["surface"]

        self._card = tk.Frame(
            self, bg=bg,
            highlightthickness=2, highlightbackground=T["border"],
            cursor="hand2",
        )
        self._card.pack(fill="both", expand=True)

        # Status dot
        self._dot = tk.Label(self._card, text="●", bg=bg, fg=T["muted"],
                             font=("Consolas", 10))
        self._dot.place(relx=1.0, rely=0.0, x=-12, y=10)

        # Bottom strip: fila de encabezado
        if self._header_row_var is not None:
            bottom = tk.Frame(self._card, bg=bg,
                              highlightthickness=1, highlightbackground=T["border"])
            bottom.pack(side="bottom", fill="x")
            tk.Label(bottom, text="Fila encabezado:",
                     bg=bg, fg=T["muted"], font=FONT_SMALL,
                     padx=8, pady=4).pack(side="left")
            spb = tk.Spinbox(
                bottom, textvariable=self._header_row_var,
                from_=1, to=100, width=4,
                bg=bg, fg=T["text"], insertbackground=T["accent"],
                relief="flat", bd=0, font=FONT_MONO,
                buttonbackground=bg,
                highlightthickness=0,
            )
            spb.pack(side="left", pady=4)
            # clic en label/frame del strip no abre el file picker
            bottom.bind("<Button-1>", lambda e: "break")

        # Centered content
        center = tk.Frame(self._card, bg=bg)
        center.place(relx=0.5, rely=0.45, anchor="center")

        self._icon_lbl = tk.Label(center, text=self._icon_idle,
                                   bg=bg, fg=T["muted"],
                                   font=("Consolas", 30))
        self._icon_lbl.pack(pady=(0, 8))

        self._title_lbl = tk.Label(center, text=self._label,
                                    bg=bg, fg=T["muted"],
                                    font=("Consolas", 8, "bold"))
        self._title_lbl.pack()

        self._name_lbl = tk.Label(center, text="Arrastra archivo o exam...",
                                   bg=bg, fg=T["muted"],
                                   font=FONT_SMALL, wraplength=180)
        self._name_lbl.pack(pady=(4, 0))

        for w in (self._card, center, self._icon_lbl,
                  self._title_lbl, self._name_lbl):
            w.bind("<Button-1>", lambda e: self._pick())

        if HAS_DND:
            self._card.drop_target_register(DND_FILES)
            self._card.dnd_bind("<<Drop>>",      self._on_drop)
            self._card.dnd_bind("<<DragEnter>>",  self._on_drag_enter)
            self._card.dnd_bind("<<DragLeave>>",  self._on_drag_leave)

    # ── DnD ──────────────────────────────────────────────────────────

    def _on_drag_enter(self, event=None):
        self._card.configure(highlightbackground=self._T["accent2"])

    def _on_drag_leave(self, event=None):
        T = self._T
        self._card.configure(
            highlightbackground=T["accent2"] if self._loaded else T["border"])

    def _on_drop(self, event):
        self._on_drag_leave()
        path = event.data.strip().strip("{}")
        if os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
            self._set_path(path)

    # ── File selection ────────────────────────────────────────────────

    def _pick(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if path:
            self._set_path(path)

    def _set_path(self, path):
        self.path_var.set(path)
        self._name_lbl.config(text="Validando…", fg=self._T["muted"])
        try:
            header_row = int(self._header_row_var.get()) if self._header_row_var else 1
        except (ValueError, AttributeError):
            header_row = 1

        def worker():
            result = validate_excel_file(path, header_row)
            self.after(0, lambda: self._show_validation(result, path))
        threading.Thread(target=worker, daemon=True).start()

    def _show_validation(self, result: dict, path: str):
        T       = self._T
        fname   = os.path.basename(path)
        display = (fname[:22] + "…") if len(fname) > 22 else fname

        if not result["ok"]:
            self._name_lbl.config(text=result["error"][:32], fg=T["error"])
            self._dot.config(fg=T["error"])
            return

        self._loaded = True
        ok_bg = T["card_ok_bg"]

        for w in self._all_widgets(self._card):
            try:
                w.configure(bg=ok_bg)
            except tk.TclError:
                pass

        self._card.configure(highlightbackground=T["accent2"])
        self._dot.configure(fg=T["success"])
        self._icon_lbl.configure(text=self._icon_ok, fg=T["accent2"])
        self._name_lbl.configure(text=display, fg=T["text"])

        if self._on_concepts_found and result.get("concept_cols"):
            self._on_concepts_found(result["concept_cols"])

    # ── Utilities ─────────────────────────────────────────────────────

    def _all_widgets(self, w):
        yield w
        for child in w.winfo_children():
            yield from self._all_widgets(child)

    def get_path(self) -> str:
        return self.path_var.get()


# ─────────────────────────── LogPanel ────────────────────────────

class LogPanel(tk.Frame):
    """Panel de log con scroll, iconos por tipo y botón de copiar."""
    def __init__(self, parent, T: dict, **kw):
        super().__init__(parent, bg=T["surface"],
                         highlightbackground=T["border"],
                         highlightthickness=1, **kw)
        self._T = T
        self._build()

    def _build(self):
        T = self._T
        header = tk.Frame(self, bg=T["surface"])
        header.pack(fill="x", padx=10, pady=(8, 0))

        tk.Label(header, text="⬡  LOG DE PROCESO", bg=T["surface"],
                 fg=T["muted"], font=("Consolas", 9, "bold")).pack(side="left")

        tk.Button(header, text="⎘ Copiar", bg=T["surface"], fg=T["muted"],
                  font=FONT_SMALL, relief="flat", bd=0, cursor="hand2",
                  command=self._copy).pack(side="right", padx=(6, 0))

        tk.Button(header, text="✕ Limpiar", bg=T["surface"], fg=T["muted"],
                  font=FONT_SMALL, relief="flat", bd=0, cursor="hand2",
                  command=self.clear).pack(side="right")

        self.text = tk.Text(
            self, bg=T["surface"], fg=T["text"],
            font=FONT_MONO, relief="flat", bd=0,
            state="disabled", wrap="word",
            selectbackground=T["accent"],
        )
        self.text.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.text.tag_config("ok",   foreground=T["success"])
        self.text.tag_config("err",  foreground=T["error"])
        self.text.tag_config("warn", foreground=T["warn"])
        self.text.tag_config("info", foreground=T["accent2"])
        self.text.tag_config("ts",   foreground=T["muted"])

    def log(self, msg: str, kind: str = ""):
        self.text.config(state="normal")
        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        icon = LOG_ICONS.get(kind, "  ")
        self.text.insert("end", f"[{ts}] ", "ts")
        self.text.insert("end", icon + msg + "\n", kind)
        self.text.see("end")
        self.text.config(state="disabled")

    def clear(self):
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.config(state="disabled")

    def _copy(self):
        content = self.text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(content)


# ─────────────────────────── ProgressStepper ────────────────────────────

class ProgressStepper(tk.Frame):
    """Barra de progreso con etiquetas de pasos."""
    STEPS = ["Leyendo", "Cruzando", "Generando", "Listo"]

    def __init__(self, parent, T: dict, **kw):
        super().__init__(parent, bg=T["bg"], **kw)
        self._T = T
        self._build()

    def _build(self):
        T = self._T

        # ttk style para la barra
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Omonel.Horizontal.TProgressbar",
                        background=T["accent"],
                        troughcolor=T["surface"],
                        bordercolor=T["border"],
                        lightcolor=T["accent"],
                        darkcolor=T["accent"])

        self.pbar = ttk.Progressbar(
            self, style="Omonel.Horizontal.TProgressbar",
            mode="determinate", maximum=len(self.STEPS),
        )
        self.pbar.pack(fill="x", padx=0, pady=(0, 4))

        lf = tk.Frame(self, bg=T["bg"])
        lf.pack(fill="x")
        self.step_labels = []
        for name in self.STEPS:
            lbl = tk.Label(lf, text=name, bg=T["bg"],
                           fg=T["muted"], font=("Consolas", 8))
            lbl.pack(side="left", expand=True)
            self.step_labels.append(lbl)

    def set_step(self, step: int):
        T = self._T
        self.pbar["value"] = step + 1
        for i, lbl in enumerate(self.step_labels):
            if i < step:
                lbl.config(fg=T["success"])
            elif i == step:
                lbl.config(fg=T["accent"], font=("Consolas", 8, "bold"))
            else:
                lbl.config(fg=T["muted"], font=("Consolas", 8))

    def reset(self):
        T = self._T
        self.pbar["value"] = 0
        for lbl in self.step_labels:
            lbl.config(fg=T["muted"], font=("Consolas", 8))


# ─────────────────────────── Generadores Omonel ────────────────────────────

def generate_omonel_layout(result_df: pd.DataFrame, output_path: str) -> str:
    """Genera el layout de dispersión bancaria Omonel en formato xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Dispersión Omonel"

    thin       = Side(style="thin", color="D0D7DE")
    border     = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill   = PatternFill("solid", fgColor="2D3748")
    hdr_font   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font  = Font(name="Arial", size=10)
    data_align = Alignment(horizontal="left", vertical="center")
    num_align  = Alignment(horizontal="right", vertical="center")
    total_fill = PatternFill("solid", fgColor="D5EDE9")
    total_font = Font(name="Arial", bold=True, size=10)

    ws.merge_cells("A1:H1")
    ws["A1"] = "OMONEL  |  LAYOUT DE DISPERSIÓN BANCARIA"
    ws["A1"].font      = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill      = PatternFill("solid", fgColor="2D3748")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ts = datetime.datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
    ws.merge_cells("A2:H2")
    ws["A2"] = f"Generado: {ts}"
    ws["A2"].font      = Font(name="Arial", italic=True, size=9, color="888888")
    ws["A2"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6

    HEADERS = ["NO.", "CLAVE EMPLEADO", "RFC", "NOMBRE COMPLETO",
               "CLABE INTERBANCARIA", "BANCO", "CUENTA VALE", "IMPORTE ($)"]
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = border
    ws.row_dimensions[4].height = 28

    alt_fill = PatternFill("solid", fgColor="EBF5F3")
    row_num  = 5
    for i, rec in enumerate(result_df.itertuples(index=False), 1):
        fill = alt_fill if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        values = [i,
                  getattr(rec, "clave_empleado", ""),
                  getattr(rec, "rfc", ""),
                  getattr(rec, "nombre", ""),
                  getattr(rec, "clabe", ""),
                  getattr(rec, "banco", ""),
                  getattr(rec, "cuenta_vale", ""),
                  getattr(rec, "importe", 0)]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.font = data_font; cell.fill = fill; cell.border = border
            if col == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 8:
                cell.alignment = num_align
                cell.number_format = '"$"#,##0.00'
            else:
                cell.alignment = data_align
        ws.row_dimensions[row_num].height = 18
        row_num += 1

    ws.cell(row=row_num, column=7, value="TOTAL DISPERSIÓN:").font  = total_font
    ws.cell(row=row_num, column=7).fill      = total_fill
    ws.cell(row=row_num, column=7).alignment = Alignment(horizontal="right")
    ws.cell(row=row_num, column=7).border    = border
    total_cell = ws.cell(row=row_num, column=8, value=f"=SUM(H5:H{row_num-1})")
    total_cell.font = total_font; total_cell.fill = total_fill
    total_cell.alignment = num_align
    total_cell.number_format = '"$"#,##0.00'; total_cell.border = border
    ws.row_dimensions[row_num].height = 22

    for i, w in enumerate([6, 18, 16, 32, 22, 18, 18, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("Resumen")
    ws2["A1"] = "RESUMEN DE DISPERSIÓN"
    ws2["A1"].font = Font(name="Arial", bold=True, size=13, color="2D3748")
    ws2.append([])
    ws2.append(["Total de registros", len(result_df)])
    total_imp = result_df["importe"].sum() if "importe" in result_df.columns else 0
    ws2.append(["Importe total a dispersar", f"${total_imp:,.2f}"])
    ws2.append(["Fecha de generación", ts])
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 22

    wb.save(output_path)
    return output_path


def generate_omonel_txt(result_df: pd.DataFrame, output_path: str,
                        numero_cliente: str) -> str:
    """Header 21 / Body 26  — dispersión por tarjeta."""
    num    = numero_cliente.strip().zfill(7)[:7]
    total  = int(round(result_df["importe"].sum() * 100))
    lines  = [f"{num}0000{total:010d}"]
    for rec in result_df.itertuples(index=False):
        cuenta = str(getattr(rec, "cuenta_vale", "")).strip().zfill(16)[:16]
        imp    = int(round(getattr(rec, "importe", 0) * 100))
        lines.append(f"{cuenta}{imp:010d}")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return output_path


def generate_omonel_txt_empleado(result_df: pd.DataFrame, output_path: str,
                                  numero_empresa: str) -> str:
    """Header 21 / Body 27  — dispersión por número de empleado."""
    num    = numero_empresa.strip().zfill(7)[:7]
    total  = int(round(result_df["importe"].sum() * 100))
    lines  = [f"{num}0000{total:010d}"]
    for rec in result_df.itertuples(index=False):
        depto_raw    = str(getattr(rec, "departamento", "")).strip()
        depto_digits = "".join(filter(str.isdigit, depto_raw))
        depto        = depto_digits.zfill(7)[:7] if depto_digits else "0000000"
        empleado     = str(getattr(rec, "clave_empleado", "")).strip().zfill(10)[:10]
        imp          = int(round(getattr(rec, "importe", 0) * 100))
        lines.append(f"{depto}{empleado}{imp:010d}")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return output_path


# ──────────────────────────── Lógica de negocio ───────────────────────────

FIXED_PEOPLE_COLS = {
    "clave_empleado", "nombre", "ap._paterno", "ap._materno",
    "puesto", "departamento", "sucursal", "tipo_de_pago",
    "fecha_ingreso", "nss", "rfc", "curp", "emisor",
}


def process_files(
    path_people:       str,
    path_vales:        str,
    conceptos:         list[str],
    log_fn,
    progress_fn,
    header_row_people: int = 1,
    header_row_vales:  int = 1,
) -> dict | None:
    """
    Lee ambos archivos, cruza por clave_empleado y calcula importes.
    Retorna {"df": DataFrame, "unmatched": int} o None si hay error.
    """
    try:
        progress_fn(0)   # Leyendo
        log_fn(f"Leyendo People… (fila de encabezados: {header_row_people})", "info")
        df_people = pd.read_excel(path_people, dtype=str, header=header_row_people - 1)
        df_people.columns = [c.strip().lower().replace(" ", "_") for c in df_people.columns]
        log_fn(f"  ↳ {len(df_people):,} filas | columnas: {list(df_people.columns)}", "")

        log_fn(f"Leyendo Cuenta Vales… (fila de encabezados: {header_row_vales})", "info")
        df_vales = pd.read_excel(path_vales, dtype=str, header=header_row_vales - 1)
        df_vales.columns = [c.strip().lower().replace(" ", "_") for c in df_vales.columns]
        log_fn(f"  ↳ {len(df_vales):,} filas | columnas: {list(df_vales.columns)}", "")

        # Validar clave_empleado
        for label, df in [("People", df_people), ("Cuenta Vales", df_vales)]:
            if "clave_empleado" not in df.columns:
                log_fn(f"ERROR: '{label}' no tiene columna 'clave_empleado'", "err")
                return None

        df_people["clave_empleado"] = df_people["clave_empleado"].str.strip().str.upper()
        df_vales["clave_empleado"]  = df_vales["clave_empleado"].str.strip().str.upper()

        # Columnas de concepto
        concept_cols = [c for c in df_people.columns if c not in FIXED_PEOPLE_COLS]
        log_fn(f"Conceptos detectados: {concept_cols}", "info")

        if conceptos:
            conceptos_norm = [c.strip().lower().replace(" ", "_") for c in conceptos]
            selected_cols  = [c for c in concept_cols if c in conceptos_norm]
            missing        = [c for c in conceptos_norm if c not in selected_cols]
            if missing:
                log_fn(f"Sin columna para: {missing}", "warn")
            if not selected_cols:
                log_fn("ERROR: Ningún concepto encontrado como columna.", "err")
                return None
            log_fn(f"Conceptos seleccionados: {selected_cols}", "info")
        else:
            selected_cols = concept_cols
            log_fn("Sin filtro: se suman todas las columnas de concepto.", "warn")

        df_people["importe"] = (
            df_people[selected_cols]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1)
        )

        progress_fn(1)   # Cruzando
        total_people = len(df_people)
        col_cv = "cuenta_vale" if "cuenta_vale" in df_vales.columns else df_vales.columns[-1]
        df_merge = df_people.merge(
            df_vales[["clave_empleado", col_cv]],
            on="clave_empleado", how="inner",
        )
        unmatched = total_people - len(df_merge)
        log_fn(f"Coincidencias: {len(df_merge):,}  |  Sin coincidencia: {unmatched:,}", "ok")

        if len(df_merge) == 0:
            log_fn("No se encontraron coincidencias.", "warn")
            return None

        def _col(df, *names, default=""):
            for n in names:
                if n in df.columns:
                    return df[n].fillna("").str.strip()
            return pd.Series([default] * len(df), index=df.index)

        nombre_completo = (
            _col(df_merge, "nombre") + " " +
            _col(df_merge, "ap._paterno") + " " +
            _col(df_merge, "ap._materno")
        ).str.strip()

        out = pd.DataFrame()
        out["clave_empleado"] = df_merge["clave_empleado"]
        out["rfc"]            = _col(df_merge, "rfc")
        out["nombre"]         = nombre_completo
        out["clabe"]          = _col(df_merge, "clabe", "clabe_interbancaria")
        out["banco"]          = _col(df_merge, "banco")
        out["departamento"]   = _col(df_merge, "departamento")
        out["cuenta_vale"]    = df_merge[col_cv]
        out["importe"]        = df_merge["importe"]

        return {"df": out, "unmatched": unmatched}

    except Exception as e:
        log_fn(f"Error procesando archivos: {e}", "err")
        return None


# ─────────────────────────── Ventana principal ────────────────────────────

if HAS_DND:
    _BaseApp = TkinterDnD.Tk
else:
    _BaseApp = tk.Tk


class OmonelApp(_BaseApp):
    def __init__(self):
        super().__init__()
        self.title("Dispersión Omonel  |  Generador de Layout")
        self.geometry("1060x820")
        self.minsize(900, 680)

        self._theme = detect_system_theme()
        self._T     = THEMES[self._theme]
        self.configure(bg=self._T["bg"])

        self._people_path = tk.StringVar()
        self._vales_path  = tk.StringVar()
        self._header_row       = tk.StringVar(value="7")
        self._header_row_vales = tk.StringVar(value="1")
        self._tipo        = tk.StringVar(value="empleado")
        self._cliente     = tk.StringVar()
        self._out_path    = tk.StringVar(value=os.path.join(
            os.path.expanduser("~"), "layout_omonel.txt"))
        self._tags: list[str] = []

        self._queue: queue.Queue = queue.Queue()
        self._working = False

        self._build_ui()

    # ──────────────────────────────────────────────────────────────────
    # UI BUILD
    # ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        T = self._T

        # ── HEADER ───────────────────────────────────────────────────
        header = tk.Frame(self, bg=T["surface"],
                          highlightbackground=T["border"], highlightthickness=1)
        header.pack(fill="x", side="top")

        inner_h = tk.Frame(header, bg=T["surface"])
        inner_h.pack(fill="x", padx=20, pady=10)

        logo = tk.Frame(inner_h, bg=T["accent"], width=44, height=44)
        logo.pack(side="left")
        logo.pack_propagate(False)
        tk.Label(logo, text="≡", bg=T["accent"], fg=T["btn_fg"],
                 font=("Consolas", 20, "bold")).place(relx=0.5, rely=0.5, anchor="center")

        title_block = tk.Frame(inner_h, bg=T["surface"])
        title_block.pack(side="left", padx=14)
        tk.Label(title_block, text="DISPERSIÓN OMONEL",
                 bg=T["surface"], fg=T["text"],
                 font=("Consolas", 14, "bold")).pack(anchor="w")
        tk.Label(title_block,
                 text=f"Generador de Layout Bancario  //  v{APP_VERSION}",
                 bg=T["surface"], fg=T["muted"], font=FONT_SMALL).pack(anchor="w")

        tk.Button(inner_h, text="Manual de Ayuda",
                  bg=T["text"], fg=T["surface"],
                  font=("Consolas", 10, "bold"), relief="flat", bd=0,
                  cursor="hand2", padx=16, pady=7,
                  command=self._show_help).pack(side="right", padx=(8, 0))

        theme_icon = "☾" if self._theme == "light" else "☀"
        tk.Button(inner_h, text=theme_icon,
                  bg=T["surface"], fg=T["muted"],
                  font=("Consolas", 16), relief="flat", bd=0, cursor="hand2",
                  command=self._toggle_theme).pack(side="right", padx=(0, 8))

        # ── TWO-COLUMN BODY ───────────────────────────────────────────
        body = tk.Frame(self, bg=T["bg"])
        body.pack(fill="both", expand=True, padx=20, pady=16)
        body.columnconfigure(0, weight=3, minsize=480)
        body.columnconfigure(1, weight=2, minsize=320)
        body.rowconfigure(0, weight=1)

        # LEFT: scrollable canvas ─────────────────────────────────────
        left_wrap = tk.Frame(body, bg=T["bg"])
        left_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_wrap.rowconfigure(0, weight=1)
        left_wrap.columnconfigure(0, weight=1)

        lcv  = tk.Canvas(left_wrap, bg=T["bg"], highlightthickness=0)
        lvsb = tk.Scrollbar(left_wrap, orient="vertical", command=lcv.yview)
        lcv.configure(yscrollcommand=lvsb.set)

        left_body = tk.Frame(lcv, bg=T["bg"])
        lbwin = lcv.create_window((0, 0), window=left_body, anchor="nw")

        def _lconf(e):
            lcv.configure(scrollregion=lcv.bbox("all"))
            lcv.itemconfig(lbwin, width=lcv.winfo_width())
        left_body.bind("<Configure>", _lconf)
        lcv.bind("<Configure>", lambda e: lcv.itemconfig(lbwin, width=e.width))
        lcv.bind_all("<MouseWheel>",
                     lambda e: lcv.yview_scroll(int(-1*(e.delta/120)), "units"))

        lcv.grid(row=0, column=0, sticky="nsew")
        lvsb.grid(row=0, column=1, sticky="ns")

        # RIGHT: fixed column ─────────────────────────────────────────
        right_col = tk.Frame(body, bg=T["bg"])
        right_col.grid(row=0, column=1, sticky="nsew")
        right_col.columnconfigure(0, weight=1)
        right_col.rowconfigure(1, weight=1)

        # ── SECTION 01 ────────────────────────────────────────────────
        self._section_badge(left_body, "01", "ARCHIVOS DE ENTRADA", pady_top=0)

        cards_row = tk.Frame(left_body, bg=T["bg"])
        cards_row.pack(fill="x", pady=(0, 8))
        cards_row.columnconfigure(0, weight=1)
        cards_row.columnconfigure(1, weight=1)

        self.picker_people = FilePickerCard(
            cards_row, label="ARCHIVO PEOPLE",
            icon_idle="⬇", icon_ok="✓", T=T,
            path_var=self._people_path,
            header_row_var=self._header_row,
            on_concepts_found=lambda cols: self.tag_input.set_suggestions(cols),
        )
        self.picker_people.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self.picker_vales = FilePickerCard(
            cards_row, label="ARCHIVO CUENTA VALES",
            icon_idle="⬇", icon_ok="✓", T=T,
            path_var=self._vales_path,
            header_row_var=self._header_row_vales,
        )
        self.picker_vales.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # ── SECTION 02 ────────────────────────────────────────────────
        self._section_badge(left_body, "02", "CLAVES DE CONCEPTOS")

        concept_card = tk.Frame(left_body, bg=T["surface"],
                                highlightbackground=T["border"], highlightthickness=1)
        concept_card.pack(fill="x")

        tk.Label(
            concept_card,
            text="Ingresa las claves correspondientes a columnas del archivo People "
                 "(ej: SUELDO, P2AH).",
            bg=T["surface"], fg=T["muted"], font=FONT_SMALL,
            justify="left", wraplength=480,
        ).pack(anchor="w", padx=16, pady=(12, 6))

        self.tag_input = TagInput(concept_card, T=T, initial_tags=self._tags)
        self.tag_input.pack(fill="x", padx=8, pady=(0, 10))

        # ── SECTION 03 ────────────────────────────────────────────────
        self._sec03_frame = tk.Frame(right_col, bg=T["surface"],
                                     highlightbackground=T["border"],
                                     highlightthickness=1)
        self._sec03_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        s3 = tk.Frame(self._sec03_frame, bg=T["surface"])
        s3.pack(fill="x", padx=16, pady=(14, 16))

        hdr03 = tk.Frame(s3, bg=T["surface"])
        hdr03.pack(fill="x", pady=(0, 16))
        self._badge_label(hdr03, "03")
        tk.Label(hdr03, text="ARCHIVO DE SALIDA", bg=T["surface"], fg=T["text"],
                 font=("Consolas", 11, "bold")).pack(side="left", padx=10)

        # Tipo toggle
        tk.Label(s3, text="TIPO DE DISPERSIÓN", bg=T["surface"], fg=T["muted"],
                 font=("Consolas", 8, "bold")).pack(anchor="w", pady=(0, 6))

        toggle_row = tk.Frame(s3, bg=T["surface"])
        toggle_row.pack(fill="x", pady=(0, 14))
        self._tipo_btns = {}
        for val, lbl in [("empleado", "Por No. Empleado"), ("tarjeta", "Por Tarjeta")]:
            b = tk.Button(
                toggle_row, text=lbl,
                command=lambda v=val: self._select_tipo(v),
                relief="flat", bd=0, cursor="hand2",
                font=("Consolas", 9, "bold"), padx=12, pady=8,
                highlightthickness=1,
            )
            b.pack(side="left", fill="x", expand=True,
                   padx=(0, 4) if val == "empleado" else (0, 0))
            self._tipo_btns[val] = b

        # Número
        self.cli_label = tk.Label(s3, text="NÚMERO DE EMPRESA", bg=T["surface"],
                                   fg=T["muted"], font=("Consolas", 8, "bold"))
        self.cli_label.pack(anchor="w", pady=(0, 6))

        num_box = tk.Frame(s3, bg=T["surface"],
                           highlightbackground=T["border"], highlightthickness=1)
        num_box.pack(fill="x", pady=(0, 4))
        tk.Label(num_box, text="⚙", bg=T["surface"], fg=T["muted"],
                 font=("Consolas", 12), padx=8).pack(side="left")
        tk.Frame(num_box, bg=T["border"], width=1).pack(side="left", fill="y", pady=4)
        tk.Entry(num_box, textvariable=self._cliente, width=14,
                 bg=T["surface"], fg=T["text"], insertbackground=T["accent"],
                 relief="flat", bd=0, font=FONT_MONO,
                 highlightthickness=0).pack(
            side="left", fill="x", expand=True, padx=8, ipady=6)

        tk.Label(s3, text="Proporcionado por OMONEL", bg=T["surface"],
                 fg=T["muted"], font=FONT_SMALL).pack(anchor="w", pady=(0, 14))

        # Destino
        tk.Label(s3, text="DESTINO DEL ARCHIVO", bg=T["surface"], fg=T["muted"],
                 font=("Consolas", 8, "bold")).pack(anchor="w", pady=(0, 6))

        dest_box = tk.Frame(s3, bg=T["surface"],
                            highlightbackground=T["border"], highlightthickness=1)
        dest_box.pack(fill="x", pady=(0, 14))
        tk.Entry(dest_box, textvariable=self._out_path,
                 bg=T["surface"], fg=T["muted"], insertbackground=T["accent"],
                 relief="flat", bd=0, font=FONT_MONO,
                 highlightthickness=0).pack(
            side="left", fill="x", expand=True, padx=8, ipady=6)
        tk.Frame(dest_box, bg=T["border"], width=1).pack(side="right", fill="y", pady=4)
        tk.Button(dest_box, text="⬇", command=self._pick_output,
                  bg=T["surface"], fg=T["muted"],
                  relief="flat", bd=0, cursor="hand2",
                  font=("Consolas", 12), padx=10, pady=6).pack(side="right")

        # Generate button
        self.btn_run = tk.Button(
            s3, text="GENERAR LAYOUT OMONEL  →",
            command=self._run,
            bg=T["warn"], fg=T["btn_fg"],
            font=("Consolas", 11, "bold"), relief="flat", bd=0, cursor="hand2",
        )
        self.btn_run.pack(fill="x", ipady=12)

        # Summary panel (oculto hasta que haya resultado)
        self.summary_frame = tk.Frame(self._sec03_frame, bg=T["surface"])

        # ── SECTION 04 ────────────────────────────────────────────────
        sec04 = tk.Frame(right_col, bg=T["surface"],
                         highlightbackground=T["border"], highlightthickness=1)
        sec04.grid(row=1, column=0, sticky="nsew")
        sec04.rowconfigure(1, weight=1)
        sec04.columnconfigure(0, weight=1)

        hdr04 = tk.Frame(sec04, bg=T["surface"])
        hdr04.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 4))
        self._badge_label(hdr04, "04")
        tk.Label(hdr04, text="RESULTADO", bg=T["surface"], fg=T["text"],
                 font=("Consolas", 11, "bold")).pack(side="left", padx=10)
        tk.Button(hdr04, text="LIMPIAR", bg=T["surface"], fg=T["muted"],
                  font=FONT_SMALL, relief="flat", bd=0, cursor="hand2",
                  command=lambda: self.log_panel.clear()).pack(side="right")

        self.log_panel = LogPanel(sec04, T=T)
        self.log_panel.grid(row=1, column=0, sticky="nsew")
        self.log_panel.log("Sistema listo. Carga los archivos y presiona Generar.", "info")

        self._select_tipo(self._tipo.get())

    # ──────────────────────────── HELPERS UI ─────────────────────────

    def _badge_label(self, parent, num: str):
        T = self._T
        tk.Label(parent, text=num, bg=T["accent"], fg=T["btn_fg"],
                 font=("Consolas", 11, "bold"),
                 padx=10, pady=4).pack(side="left")

    def _section_badge(self, parent, num: str, title: str, pady_top: int = 16):
        T = self._T
        frame = tk.Frame(parent, bg=T["bg"])
        frame.pack(fill="x", pady=(pady_top, 12))
        self._badge_label(frame, num)
        tk.Label(frame, text=title, bg=T["bg"], fg=T["text"],
                 font=("Consolas", 11, "bold")).pack(side="left", padx=12)

    def _select_tipo(self, value: str):
        T = self._T
        self._tipo.set(value)
        for val, btn in self._tipo_btns.items():
            if val == value:
                btn.config(bg=T["accent"], fg=T["btn_fg"],
                           highlightbackground=T["accent"], highlightthickness=2)
            else:
                btn.config(bg=T["surface"], fg=T["muted"],
                           highlightbackground=T["border"], highlightthickness=1)
        self.cli_label.config(
            text="NÚMERO DE CLIENTE" if value == "tarjeta" else "NÚMERO DE EMPRESA")

    def _toggle_theme(self):
        self._tags = self.tag_input.get_tags() if hasattr(self, "tag_input") else self._tags
        self._theme = "light" if self._theme == "dark" else "dark"
        self._T = THEMES[self._theme]
        self.configure(bg=self._T["bg"])
        for w in self.winfo_children():
            w.destroy()
        self._build_ui()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile="layout_omonel.txt",
        )
        if path:
            self._out_path.set(path)

    def _set_btn_state(self, state: str):
        T = self._T
        cfg = {
            "normal":  (T["warn"],    T["warn_fg"], "GENERAR LAYOUT OMONEL  →", "normal"),
            "loading": (T["border"],  T["muted"],   "⏳  Procesando…",          "disabled"),
            "success": (T["success"], T["btn_fg"],  "✓  GENERADO EXITOSAMENTE", "disabled"),
            "error":   (T["error"],   T["btn_fg"],  "✗  ERROR — reintentar",    "normal"),
        }[state]
        self.btn_run.config(bg=cfg[0], fg=cfg[1], text=cfg[2], state=cfg[3])

    def _show_summary(self, df: pd.DataFrame, unmatched: int,
                      txt_path: str, xlsx_path: str):
        T  = self._T
        sf = self.summary_frame
        for w in sf.winfo_children():
            w.destroy()

        total = df["importe"].sum()

        tk.Frame(sf, bg=T["border"], height=1).pack(fill="x")

        inner = tk.Frame(sf, bg=T["surface"])
        inner.pack(fill="x", padx=16, pady=8)
        for col, (lbl, val) in enumerate([
            ("Registros", f"{len(df):,}"),
            ("Importe",   f"${total:,.2f}"),
            ("Sin match", f"{unmatched:,}"),
        ]):
            tk.Label(inner, text=lbl, bg=T["surface"],
                     fg=T["muted"], font=FONT_SMALL).grid(
                row=0, column=col, sticky="w", padx=(0, 16))
            tk.Label(inner, text=val, bg=T["surface"],
                     fg=T["text"], font=("Consolas", 9, "bold")).grid(
                row=1, column=col, sticky="w", padx=(0, 16))

        btn_row = tk.Frame(sf, bg=T["surface"])
        btn_row.pack(fill="x", padx=16, pady=(4, 12))
        tk.Button(btn_row, text="Abrir Excel",
                  command=lambda: open_file(xlsx_path),
                  bg=T["accent"], fg=T["btn_fg"], font=FONT_SMALL,
                  relief="flat", bd=0, cursor="hand2",
                  ).pack(side="left", ipadx=10, ipady=4, padx=(0, 6))
        tk.Button(btn_row, text="Abrir TXT",
                  command=lambda: open_file(txt_path),
                  bg=T["accent2"], fg=T["btn_fg"], font=FONT_SMALL,
                  relief="flat", bd=0, cursor="hand2",
                  ).pack(side="left", ipadx=10, ipady=4)

        sf.pack(fill="x", before=self.btn_run)

    def _show_help(self):
        T   = self._T
        win = tk.Toplevel(self)
        win.title("Manual de Ayuda")
        win.configure(bg=T["bg"])
        win.geometry("560x480")
        win.resizable(False, False)

        tk.Label(win, text="AYUDA  —  DISPERSIÓN OMONEL",
                 bg=T["bg"], fg=T["accent"],
                 font=("Consolas", 11, "bold")).pack(padx=20, pady=(16, 8))

        txt = tk.Text(win, bg=T["surface"], fg=T["text"],
                      font=FONT_MONO, relief="flat", bd=0, wrap="word",
                      state="normal",
                      highlightthickness=1, highlightbackground=T["border"])
        txt.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        txt.insert("end", HELP_TEXT)
        txt.config(state="disabled")

        tk.Button(win, text="Cerrar", command=win.destroy,
                  bg=T["accent"], fg=T["btn_fg"], font=FONT_SMALL,
                  relief="flat", bd=0, cursor="hand2",
                  ).pack(pady=(0, 16), ipadx=20, ipady=6)

    # ──────────────────────────── RUN / THREADING ────────────────────

    def _run(self):
        if self._working:
            return

        p_people  = self._people_path.get().strip()
        p_vales   = self._vales_path.get().strip()
        conceptos = self.tag_input.get_tags()
        out_path  = self._out_path.get().strip()
        cliente   = self._cliente.get().strip()
        hdr_str_people = self._header_row.get().strip()
        hdr_str_vales  = self._header_row_vales.get().strip()

        errors = []
        if not p_people:  errors.append("Selecciona el archivo People.")
        if not p_vales:   errors.append("Selecciona el archivo Cuenta Vales.")
        if not out_path:  errors.append("Define la ruta del archivo de salida.")
        if not cliente:   errors.append("Ingresa el Número de cliente / empresa.")
        elif not cliente.isdigit():
            errors.append("El Número de cliente debe ser solo dígitos.")
        if not hdr_str_people.isdigit() or int(hdr_str_people) < 1:
            errors.append("La fila de encabezados (People) debe ser un número ≥ 1.")
        if not hdr_str_vales.isdigit() or int(hdr_str_vales) < 1:
            errors.append("La fila de encabezados (Cuenta Vales) debe ser un número ≥ 1.")
        if errors:
            messagebox.showerror("Validación", "\n".join(errors))
            return

        self.summary_frame.pack_forget()
        self._working = True
        self._set_btn_state("loading")

        log = self.log_panel.log
        log("═" * 48, "")
        log("Iniciando proceso…", "info")
        if conceptos:
            log(f"Conceptos a filtrar: {', '.join(conceptos)}", "info")

        params = {
            "p_people":          p_people,
            "p_vales":           p_vales,
            "conceptos":         conceptos,
            "out_path":          out_path,
            "cliente":           cliente,
            "header_row_people": int(hdr_str_people),
            "header_row_vales":  int(hdr_str_vales),
            "tipo":              self._tipo.get(),
        }
        threading.Thread(target=self._worker, args=(params,), daemon=True).start()
        self.after(50, self._poll_queue)

    def _worker(self, params):
        q = self._queue

        def log(msg, kind=""):
            q.put({"type": "log", "msg": msg, "kind": kind})

        def progress(step):
            q.put({"type": "progress", "step": step})

        result = process_files(
            params["p_people"], params["p_vales"],
            params["conceptos"], log, progress,
            params["header_row_people"],
            params["header_row_vales"],
        )

        if result is None:
            q.put({"type": "done", "success": False})
            return

        df        = result["df"]
        unmatched = result["unmatched"]
        txt_path  = params["out_path"]
        xlsx_path = os.path.splitext(txt_path)[0] + ".xlsx"
        errors    = []

        progress(2)

        try:
            if params["tipo"] == "tarjeta":
                generate_omonel_txt(df, txt_path, params["cliente"])
                log(f"TXT (Por Tarjeta) → {txt_path}", "ok")
            else:
                generate_omonel_txt_empleado(df, txt_path, params["cliente"])
                log(f"TXT (Por Empleado) → {txt_path}", "ok")
        except Exception as e:
            log(f"Error al guardar TXT: {e}", "err")
            errors.append(str(e))

        try:
            generate_omonel_layout(df, xlsx_path)
            log(f"Excel (referencia) → {xlsx_path}", "ok")
        except Exception as e:
            log(f"Error al guardar Excel: {e}", "err")
            errors.append(str(e))

        total = df["importe"].sum()
        log(f"Total: {len(df):,} registros  |  ${total:,.2f}", "ok")
        progress(3)

        q.put({
            "type":      "done",
            "success":   len(errors) == 0,
            "df":        df,
            "unmatched": unmatched,
            "txt_path":  txt_path,
            "xlsx_path": xlsx_path,
            "errors":    errors,
        })

    def _poll_queue(self):
        try:
            while True:
                msg   = self._queue.get_nowait()
                mtype = msg["type"]

                if mtype == "log":
                    self.log_panel.log(msg["msg"], msg["kind"])

                elif mtype == "progress":
                    pass   # no stepper en esta UI

                elif mtype == "done":
                    self._working = False
                    log = self.log_panel.log
                    log("═" * 48, "")

                    if msg["success"]:
                        self._set_btn_state("success")
                        self.after(3000, lambda: self._set_btn_state("normal"))
                        self._show_summary(
                            msg["df"], msg["unmatched"],
                            msg["txt_path"], msg["xlsx_path"],
                        )
                    else:
                        self._set_btn_state("error")
                        if msg.get("errors"):
                            messagebox.showerror(
                                "Errores al guardar",
                                "\n".join(msg["errors"]),
                            )
                    return

        except queue.Empty:
            pass

        if self._working:
            self.after(50, self._poll_queue)


def main():
    app = OmonelApp()
    app.mainloop()


if __name__ == "__main__":
    main()
