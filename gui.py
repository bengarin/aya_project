"""
gui.py
------
Interface graphique du logiciel.

Elle utilise CustomTkinter s'il est installe (look moderne), sinon Tkinter/ttk
standard. Le traitement est lance dans un thread separe pour que la fenetre ne
se fige jamais : le thread envoie ses messages dans une file d'attente, et
l'interface la consulte toutes les 100 ms.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))

from excel_reader import list_sheet_names                     # noqa: E402
from logger import CHECK, CREATE, ERROR, KAM_NOT_FOUND, MODIF, RED, STORE_NOT_FOUND  # noqa: E402
from pipeline import PipelineError, Selection, Session        # noqa: E402
from processor import Options                                 # noqa: E402

try:                                                          # look moderne si disponible
    import customtkinter as ctk

    HAS_CTK = True
except Exception:                                             # pragma: no cover
    ctk = None
    HAS_CTK = False

APP_TITLE = "AUTOMATISATION EXCEL - AYA"
AUTO = "(detection automatique)"
FONT = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_TITLE = ("Segoe UI", 18, "bold")

COLORS = {
    "bg": "#F2F4F7",
    "card": "#FFFFFF",
    "primary": "#1F4E79",
    "primary_dark": "#163A5A",
    "accent": "#2E7D32",
    "text": "#1B1F24",
    "muted": "#5A6270",
    "danger": "#C62828",
}

LEVEL_COLORS = {
    ERROR: "#B71C1C",
    CHECK: "#E65100",
    RED: "#C62828",
    STORE_NOT_FOUND: "#AD1457",
    KAM_NOT_FOUND: "#6A1B9A",
    CREATE: "#1B5E20",
    MODIF: "#1F4E79",
}


# ---------------------------------------------------------------------------
# Petite couche de compatibilite CustomTkinter / Tkinter
# ---------------------------------------------------------------------------
def _frame(parent, **kw):
    if HAS_CTK:
        return ctk.CTkFrame(parent, fg_color=kw.pop("color", COLORS["card"]), corner_radius=kw.pop("radius", 10))
    kw.pop("color", None)
    kw.pop("radius", None)
    return tk.Frame(parent, bg=COLORS["card"], highlightbackground="#D8DCE3", highlightthickness=1)


def _label(parent, text, font=FONT, color=None, **kw):
    if HAS_CTK:
        return ctk.CTkLabel(parent, text=text, font=font, text_color=color or COLORS["text"], **kw)
    kw.pop("anchor", None)
    return tk.Label(parent, text=text, font=font, fg=color or COLORS["text"],
                    bg=parent["bg"] if isinstance(parent, tk.Frame) else COLORS["card"], anchor="w", **kw)


def _button(parent, text, command, kind="normal", width=None):
    if HAS_CTK:
        colors = {
            "primary": (COLORS["primary"], COLORS["primary_dark"]),
            "accent": (COLORS["accent"], "#1B5E20"),
            "normal": ("#4A5568", "#2D3748"),
        }[kind]
        return ctk.CTkButton(parent, text=text, command=command, font=FONT_BOLD,
                             fg_color=colors[0], hover_color=colors[1],
                             width=width or 160, height=36, corner_radius=8)
    bg = {"primary": COLORS["primary"], "accent": COLORS["accent"], "normal": "#4A5568"}[kind]
    return tk.Button(parent, text=text, command=command, font=FONT_BOLD, bg=bg, fg="white",
                     activebackground=bg, relief="flat", padx=12, pady=6,
                     cursor="hand2", disabledforeground="#C8CDD4")


def _entry(parent, textvariable):
    if HAS_CTK:
        return ctk.CTkEntry(parent, textvariable=textvariable, font=FONT, height=32)
    return tk.Entry(parent, textvariable=textvariable, font=FONT, relief="solid", bd=1, bg="#FBFCFD")


def _check(parent, text, variable):
    if HAS_CTK:
        return ctk.CTkCheckBox(parent, text=text, variable=variable, font=FONT, checkbox_width=18, checkbox_height=18)
    return tk.Checkbutton(parent, text=text, variable=variable, font=FONT, bg=COLORS["card"],
                          anchor="w", highlightthickness=0)


def _progressbar(parent):
    if HAS_CTK:
        bar = ctk.CTkProgressBar(parent, height=16, corner_radius=8)
        bar.set(0)
        return bar
    return ttk.Progressbar(parent, mode="determinate", maximum=100)


def _set_progress(bar, value: float) -> None:
    """value entre 0 et 1."""
    if HAS_CTK:
        bar.set(max(0.0, min(1.0, value)))
    else:
        bar["value"] = max(0.0, min(1.0, value)) * 100


def _set_state(widget, enabled: bool) -> None:
    try:
        widget.configure(state="normal" if enabled else "disabled")
    except Exception:                                         # pragma: no cover
        pass


def open_in_explorer(path: Path) -> None:
    """Ouvre un fichier avec l'application par defaut (Excel en general)."""
    path = Path(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))                           # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:                                  # pragma: no cover
        messagebox.showwarning("Ouverture impossible", f"Ouvrez le fichier manuellement :\n{path}\n\n({exc})")


# ---------------------------------------------------------------------------
# Fenetre principale
# ---------------------------------------------------------------------------
class App:
    def __init__(self) -> None:
        if HAS_CTK:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")
            self.root = ctk.CTk()
        else:
            self.root = tk.Tk()
            style = ttk.Style(self.root)
            if "clam" in style.theme_names():
                style.theme_use("clam")
            style.configure("TCombobox", fieldbackground="white")
        self.root.title(APP_TITLE)
        # On s'adapte a la taille de l'ecran (utile sur les petits portables).
        width = min(1200, self.root.winfo_screenwidth() - 60)
        height = min(920, self.root.winfo_screenheight() - 70)
        self.root.geometry(f"{max(width, 900)}x{max(height, 620)}")
        self.root.minsize(1000, 620)
        self.root.configure(bg=COLORS["bg"])

        # Etat
        self.queue: queue.Queue = queue.Queue()
        self.session: Session | None = None
        self.output_path: Path | None = None
        self.report_paths: list[Path] = []
        self.busy = False
        self.all_entries: list = []

        self.var_ref = tk.StringVar()
        self.var_target = tk.StringVar()
        self.var_kam = tk.StringVar()
        self.var_ref_sheet = tk.StringVar(value=AUTO)
        self.var_target_sheet = tk.StringVar(value=AUTO)
        self.var_kam_sheet = tk.StringVar(value=AUTO)
        self.var_status = tk.StringVar(value="Selectionnez les 3 fichiers Excel puis lancez l'analyse.")
        self.var_filter = tk.StringVar(value="Tout afficher")
        self.opt_city = tk.BooleanVar(value=False)
        self.opt_unknown_div = tk.BooleanVar(value=False)
        self.opt_idaya = tk.BooleanVar(value=False)
        self.opt_missing_stores = tk.BooleanVar(value=False)
        self.opt_highlight = tk.BooleanVar(value=False)
        self.stat_vars: dict[str, tk.StringVar] = {}

        self._build_ui()
        self.root.after(100, self._poll_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Construction de l'interface
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=COLORS["bg"])
        container.pack(fill="both", expand=True, padx=14, pady=10)

        header = tk.Frame(container, bg=COLORS["bg"])
        header.pack(fill="x", pady=(0, 8))
        tk.Label(header, text=APP_TITLE, font=FONT_TITLE, bg=COLORS["bg"], fg=COLORS["primary"]).pack(side="left")
        tk.Label(header, text="Reference (lecture seule)  +  BDD a traiter  +  Affectation KAM",
                 font=FONT, bg=COLORS["bg"], fg=COLORS["muted"]).pack(side="left", padx=14)

        self._build_files(container)
        self._build_options(container)
        self._build_actions(container)
        self._build_results(container)

    # ------------------------------------------------------------------
    def _build_files(self, parent) -> None:
        card = _frame(parent)
        card.pack(fill="x", pady=4)
        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill="x", padx=14, pady=8)
        tk.Label(inner, text="1. Fichiers", font=FONT_BOLD, bg=COLORS["card"], fg=COLORS["primary"]).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        rows = (
            ("Fichier Reference (jamais modifie)", self.var_ref, self.var_ref_sheet, "reference"),
            ("Fichier a traiter (le seul modifie)", self.var_target, self.var_target_sheet, "target"),
            ("Fichier Affectation KAM", self.var_kam, self.var_kam_sheet, "kam"),
        )
        self.sheet_combos: dict[str, ttk.Combobox] = {}
        self.path_entries: dict[str, object] = {}
        for index, (label, path_var, sheet_var, kind) in enumerate(rows, start=1):
            tk.Label(inner, text=label, font=FONT, bg=COLORS["card"], fg=COLORS["text"], anchor="w", width=34).grid(
                row=index, column=0, sticky="w", pady=3)
            entry = _entry(inner, path_var)
            entry.grid(row=index, column=1, sticky="ew", padx=8, pady=4)
            self.path_entries[kind] = entry
            _button(inner, "Parcourir", lambda k=kind: self.choose_file(k), kind="normal", width=110).grid(
                row=index, column=2, padx=4, pady=4)
            combo = ttk.Combobox(inner, textvariable=sheet_var, values=[AUTO], state="readonly", width=26, font=FONT)
            combo.grid(row=index, column=3, padx=6, pady=4)
            combo.bind("<<ComboboxSelected>>", lambda _e: self._reset_analysis())
            self.sheet_combos[kind] = combo
        inner.grid_columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    def _build_options(self, parent) -> None:
        card = _frame(parent)
        card.pack(fill="x", pady=4)
        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill="x", padx=14, pady=8)
        tk.Label(inner, text="2. Options (par defaut : regles metier strictes)", font=FONT_BOLD,
                 bg=COLORS["card"], fg=COLORS["primary"]).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        options = (
            (self.opt_city, "Mettre a jour la colonne City depuis la Reference"),
            (self.opt_unknown_div, "Traiter les divisions non prevues (RAC...) comme des divisions normales"),
            (self.opt_idaya, "Numeroter automatiquement IDAYA sur les lignes creees"),
            (self.opt_missing_stores, "Ajouter les magasins de la Reference absents du fichier 2"),
            (self.opt_highlight, "Surligner en orange les magasins absents de la Reference"),
        )
        for index, (var, text) in enumerate(options):
            row, column = 1 + index // 2, index % 2
            check = _check(inner, text, var)
            check.grid(row=row, column=column, sticky="w", padx=(0, 24), pady=2)
            var.trace_add("write", lambda *_a: self._reset_analysis())
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    def _build_actions(self, parent) -> None:
        card = _frame(parent)
        card.pack(fill="x", pady=4)
        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill="x", padx=14, pady=8)

        self.btn_analyze = _button(inner, "Lancer l'analyse (apercu)", self.on_analyze, kind="primary", width=220)
        self.btn_analyze.pack(side="left")
        self.btn_generate = _button(inner, "Confirmer et generer le fichier", self.on_generate, kind="accent", width=250)
        self.btn_generate.pack(side="left", padx=10)
        self.btn_open = _button(inner, "Ouvrir le fichier resultat", self.on_open_result, width=210)
        self.btn_open.pack(side="left")
        self.btn_report = _button(inner, "Ouvrir le rapport", self.on_open_report, width=170)
        self.btn_report.pack(side="left", padx=10)
        _set_state(self.btn_generate, False)
        _set_state(self.btn_open, False)
        _set_state(self.btn_report, False)

        bar_frame = tk.Frame(card, bg=COLORS["card"])
        bar_frame.pack(fill="x", padx=14, pady=(0, 8))
        self.progress = _progressbar(bar_frame)
        self.progress.pack(fill="x")
        tk.Label(bar_frame, textvariable=self.var_status, font=FONT, bg=COLORS["card"],
                 fg=COLORS["muted"], anchor="w", justify="left").pack(fill="x", pady=(6, 0))

    # ------------------------------------------------------------------
    def _build_results(self, parent) -> None:
        card = _frame(parent)
        card.pack(fill="both", expand=True, pady=4)
        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        # --- colonne gauche : statistiques
        left = tk.Frame(inner, bg=COLORS["card"], width=330)
        left.pack(side="left", fill="y", padx=(0, 14))
        left.pack_propagate(False)
        tk.Label(left, text="3. Statistiques", font=FONT_BOLD, bg=COLORS["card"], fg=COLORS["primary"]).pack(anchor="w")
        labels = (
            "Lignes analysees", "Stores traites", "Lignes modifiees", "Lignes deja conformes",
            "Lignes ajoutees", "Lignes rouges / Store vide", "Stores non trouves",
            "KAM trouves", "KAM non trouves", "Duplicates evites", "A verifier",
        )
        for name in labels:
            row = tk.Frame(left, bg=COLORS["card"])
            row.pack(fill="x", pady=0)
            tk.Label(row, text=name, font=FONT, bg=COLORS["card"], fg=COLORS["text"], anchor="w").pack(side="left")
            var = tk.StringVar(value="-")
            self.stat_vars[name] = var
            tk.Label(row, textvariable=var, font=FONT_BOLD, bg=COLORS["card"], fg=COLORS["primary"]).pack(side="right")

        tk.Label(left, text="Colonnes detectees", font=FONT_BOLD, bg=COLORS["card"],
                 fg=COLORS["primary"]).pack(anchor="w", pady=(8, 2))
        self.columns_text = tk.Text(left, height=5, font=("Consolas", 8), wrap="word",
                                    relief="solid", bd=1, bg="#FBFCFD")
        self.columns_text.pack(fill="both", expand=True)
        self.columns_text.insert("1.0", "En attente des fichiers...")
        self.columns_text.configure(state="disabled")

        # --- colonne droite : rapport
        right = tk.Frame(inner, bg=COLORS["card"])
        right.pack(side="left", fill="both", expand=True)
        head = tk.Frame(right, bg=COLORS["card"])
        head.pack(fill="x")
        tk.Label(head, text="4. Rapport de traitement", font=FONT_BOLD, bg=COLORS["card"],
                 fg=COLORS["primary"]).pack(side="left")
        self.filter_combo = ttk.Combobox(head, textvariable=self.var_filter, state="readonly", width=24,
                                         font=FONT, values=["Tout afficher"])
        self.filter_combo.pack(side="right")
        self.filter_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_report())

        text_frame = tk.Frame(right, bg=COLORS["card"])
        text_frame.pack(fill="both", expand=True, pady=(6, 0))
        self.report_text = tk.Text(text_frame, font=("Consolas", 9), wrap="none", relief="solid", bd=1, bg="#FBFCFD")
        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.report_text.yview)
        xscroll = ttk.Scrollbar(text_frame, orient="horizontal", command=self.report_text.xview)
        self.report_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set, state="disabled")
        self.report_text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        for level, color in LEVEL_COLORS.items():
            self.report_text.tag_configure(level, foreground=color)

    # ------------------------------------------------------------------
    # Actions utilisateur
    # ------------------------------------------------------------------
    def choose_file(self, kind: str) -> None:
        titles = {
            "reference": "Fichier 1 - Reference (lecture seule)",
            "target": "Fichier 2 - BDD a traiter",
            "kam": "Fichier 3 - Affectation des KAM",
        }
        path = filedialog.askopenfilename(
            title=titles[kind], filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")]
        )
        if not path:
            return
        {"reference": self.var_ref, "target": self.var_target, "kam": self.var_kam}[kind].set(path)
        entry = self.path_entries.get(kind)
        if entry is not None:
            try:
                entry.xview_moveto(1.0)                       # afficher la fin du chemin (le nom du fichier)
            except Exception:
                pass
        sheets = list_sheet_names(path)
        combo = self.sheet_combos[kind]
        combo["values"] = [AUTO] + sheets
        {"reference": self.var_ref_sheet, "target": self.var_target_sheet, "kam": self.var_kam_sheet}[kind].set(AUTO)
        self._reset_analysis()
        self.var_status.set(f"{titles[kind]} : {Path(path).name} ({len(sheets)} feuille(s) detectee(s)).")

    # ------------------------------------------------------------------
    def _selection(self) -> Selection:
        def sheet(var: tk.StringVar) -> str | None:
            value = var.get()
            return None if value in ("", AUTO) else value

        kam_sheet = sheet(self.var_kam_sheet)
        return Selection(
            reference=self.var_ref.get().strip() or None,
            target=self.var_target.get().strip() or None,
            kam=self.var_kam.get().strip() or None,
            reference_sheet=sheet(self.var_ref_sheet),
            target_sheet=sheet(self.var_target_sheet),
            kam_sheets=[kam_sheet] if kam_sheet else None,
        )

    def _options(self) -> Options:
        return Options(
            update_city=self.opt_city.get(),
            process_unknown_divisions=self.opt_unknown_div.get(),
            auto_number_idaya=self.opt_idaya.get(),
            add_missing_reference_stores=self.opt_missing_stores.get(),
            highlight_unknown_stores=self.opt_highlight.get(),
        )

    def _reset_analysis(self) -> None:
        """Toute modification d'un fichier/option invalide l'apercu precedent."""
        if self.busy:
            return
        _set_state(self.btn_generate, False)
        self.output_path = None
        self.report_paths = []
        _set_state(self.btn_open, False)
        _set_state(self.btn_report, False)

    # ------------------------------------------------------------------
    def on_analyze(self) -> None:
        if self.busy:
            return
        try:
            selection = self._selection()
            selection.validate()
        except PipelineError as exc:
            messagebox.showerror("Fichiers manquants", str(exc))
            return

        self._set_busy(True)
        self._clear_report()
        if self.session is not None:
            self.session.close()
        self.session = Session(selection, self._options())
        threading.Thread(target=self._worker_analyze, daemon=True).start()

    def _worker_analyze(self) -> None:
        assert self.session is not None
        try:
            data = self.session.load(lambda m, v: self.queue.put(("progress", m, v)))
            self.queue.put(("columns", data.describe(), data.warnings))
            plan = self.session.analyze(lambda m, v: self.queue.put(("progress", m, v)))
            self.queue.put(("analyzed", plan))
        except PipelineError as exc:
            self.queue.put(("error", str(exc)))
        except Exception as exc:                              # pragma: no cover
            self.queue.put(("error", f"Erreur inattendue :\n{exc}\n\n{traceback.format_exc(limit=3)}"))

    # ------------------------------------------------------------------
    def on_generate(self) -> None:
        if self.busy or self.session is None or self.session.plan is None:
            return
        default = self.session.default_output()
        path = filedialog.asksaveasfilename(
            title="Enregistrer le fichier resultat",
            defaultextension=".xlsx",
            initialfile=default.name,
            initialdir=str(default.parent),
            filetypes=[("Classeur Excel", "*.xlsx")],
        )
        if not path:
            return
        self._set_busy(True)
        threading.Thread(target=self._worker_generate, args=(path,), daemon=True).start()

    def _worker_generate(self, path: str) -> None:
        assert self.session is not None
        try:
            result, reports = self.session.apply(
                path, write_report=True, progress=lambda m, v: self.queue.put(("progress", m, v))
            )
            self.queue.put(("written", result, reports))
        except PipelineError as exc:
            self.queue.put(("error", str(exc)))
        except Exception as exc:                              # pragma: no cover
            self.queue.put(("error", f"Erreur inattendue :\n{exc}\n\n{traceback.format_exc(limit=3)}"))

    # ------------------------------------------------------------------
    def on_open_result(self) -> None:
        if self.output_path and Path(self.output_path).exists():
            open_in_explorer(self.output_path)

    def on_open_report(self) -> None:
        for report in self.report_paths:
            if str(report).endswith(".xlsx") and Path(report).exists():
                open_in_explorer(report)
                return
        if self.report_paths:
            open_in_explorer(self.report_paths[0])

    # ------------------------------------------------------------------
    # File d'attente : le thread parle a l'interface uniquement par ici
    # ------------------------------------------------------------------
    def _poll_queue(self) -> None:
        try:
            while True:
                message = self.queue.get_nowait()
                kind = message[0]
                if kind == "progress":
                    self.var_status.set(message[1])
                    _set_progress(self.progress, message[2])
                elif kind == "columns":
                    self._show_columns(message[1], message[2])
                elif kind == "analyzed":
                    self._show_plan(message[1])
                    self._set_busy(False)
                elif kind == "written":
                    self._show_written(message[1], message[2])
                    self._set_busy(False)
                elif kind == "error":
                    self._set_busy(False)
                    _set_progress(self.progress, 0)
                    self.var_status.set("Traitement interrompu.")
                    messagebox.showerror("Erreur", message[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        _set_state(self.btn_analyze, not busy)
        if busy:
            _set_state(self.btn_generate, False)

    def _show_columns(self, description: str, warnings: list[str]) -> None:
        self.columns_text.configure(state="normal")
        self.columns_text.delete("1.0", "end")
        text = description
        if warnings:
            text += "\n\nATTENTION :\n- " + "\n- ".join(warnings)
        self.columns_text.insert("1.0", text)
        self.columns_text.configure(state="disabled")

    def _clear_report(self) -> None:
        self.all_entries = []
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        self.report_text.configure(state="disabled")
        for var in self.stat_vars.values():
            var.set("-")

    def _show_plan(self, plan) -> None:
        for label, value in plan.stats.as_pairs():
            if label in self.stat_vars:
                self.stat_vars[label].set(str(value))
        self.all_entries = plan.logger.sorted_entries()
        levels = ["Tout afficher"] + sorted({entry.level for entry in self.all_entries})
        self.filter_combo["values"] = levels
        if self.var_filter.get() not in levels:
            self.var_filter.set("Tout afficher")
        self._refresh_report()
        _set_progress(self.progress, 1.0)
        stats = plan.stats
        self.var_status.set(
            f"Apercu pret : {stats.updated_rows} ligne(s) a modifier, {stats.created_rows} a creer, "
            f"{stats.red_rows} ligne(s) rouge(s), {stats.duplicates_avoided} duplicate(s) evite(s), "
            f"{stats.to_check} point(s) a verifier. Cliquez sur 'Confirmer et generer le fichier'."
        )
        _set_state(self.btn_generate, True)   # meme sans modification, l'utilisateur peut generer une copie

    def _refresh_report(self) -> None:
        wanted = self.var_filter.get()
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        shown = 0
        for entry in self.all_entries:
            if wanted != "Tout afficher" and entry.level != wanted:
                continue
            self.report_text.insert("end", entry.as_text() + "\n", entry.level)
            shown += 1
        if not shown:
            self.report_text.insert("end", "Aucun evenement a afficher.")
        self.report_text.configure(state="disabled")

    def _show_written(self, result, reports: list[Path]) -> None:
        self.output_path = Path(result.output_path)
        self.report_paths = [Path(p) for p in reports]
        _set_state(self.btn_open, True)
        _set_state(self.btn_report, bool(reports))
        _set_progress(self.progress, 1.0)
        message = (
            f"Fichier genere :\n{result.output_path}\n\n"
            f"Cellules modifiees : {result.updated_cells}\n"
            f"Lignes creees      : {result.created_rows}\n"
            f"Lignes colorees    : {result.colored_rows}\n"
        )
        if reports:
            message += "\nRapport :\n" + "\n".join(str(p) for p in reports)
        if result.warnings:
            message += "\n\nATTENTION :\n- " + "\n- ".join(result.warnings)
        self.var_status.set(f"Termine. Fichier resultat : {result.output_path}")
        messagebox.showinfo("Traitement termine", message)

    # ------------------------------------------------------------------
    def _on_close(self) -> None:
        if self.busy and not messagebox.askokcancel("Quitter", "Un traitement est en cours. Quitter quand meme ?"):
            return
        if self.session is not None:
            self.session.close()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch() -> None:
    App().run()


if __name__ == "__main__":
    launch()
