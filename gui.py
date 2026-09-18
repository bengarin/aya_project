"""
gui.py
------
Interface graphique du logiciel.

Elle est construite autour de 3 etapes numerotees, pour qu'un utilisateur qui
ne connait pas le programme sache toujours quoi faire ensuite :

    1. choisir les 3 fichiers
    2. lancer l'analyse (rien n'est ecrit)
    3. lire le resultat, puis generer le fichier

Elle utilise CustomTkinter s'il est installe, sinon Tkinter standard.
Le traitement tourne dans un thread separe : il envoie ses messages dans une
file d'attente que l'interface consulte toutes les 100 ms, donc la fenetre ne
se fige jamais.
"""

from __future__ import annotations

import json
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

from excel_reader import list_sheet_names                       # noqa: E402
from logger import (                                            # noqa: E402
    A_VERIFIER, CONFORME, CREEE, DUPLICATE, ERREUR, IGNOREE,
    KAM_NON_TROUVE, MODIFIEE, ROUGE, STORE_NON_TROUVE,
)
from pipeline import PipelineError, Selection, Session          # noqa: E402
from processor import Options                                   # noqa: E402

try:                                                            # look moderne si disponible
    import customtkinter as ctk

    HAS_CTK = True
except Exception:                                               # pragma: no cover
    ctk = None
    HAS_CTK = False

APP_TITLE = "AUTOMATISATION EXCEL - AYA"
AUTO = "(detection automatique)"
CONFIG_FILE = Path.home() / ".aya_excel.json"

FONT = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 9)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_STEP = ("Segoe UI", 13, "bold")
FONT_TITLE = ("Segoe UI", 19, "bold")
FONT_BIG = ("Segoe UI", 22, "bold")

C = {
    "bg": "#EEF2F1",
    "card": "#FFFFFF",
    "soft": "#F5F8F7",
    "line": "#D7DFDD",
    "ink": "#182225",
    "muted": "#5B6B6C",
    "primary": "#1F7A6C", "primary_bg": "#E6F4F0",
    "primary_dark": "#15564C",
    "red": "#B3392C", "red_bg": "#FBEAE7",
    "gray": "#6B7280", "gray_bg": "#EEF0F1",
    "blue": "#2F6FA8", "blue_bg": "#E8F1F9",
    "green": "#2F7D52", "green_bg": "#E7F6EE",
    "amber": "#B9791A", "amber_bg": "#FBF1DE",
}

# Couleur de chaque decision, la meme que dans le schema explicatif
LEVEL_COLORS = {
    ERREUR: C["red"],
    A_VERIFIER: C["amber"],
    ROUGE: C["red"],
    STORE_NON_TROUVE: C["gray"],
    KAM_NON_TROUVE: C["amber"],
    CREEE: C["primary"],
    MODIFIEE: C["blue"],
    CONFORME: C["green"],
    IGNOREE: C["muted"],
    DUPLICATE: C["amber"],
}

# Les 5 chiffres qui comptent vraiment, affiches en grand
CARTES = [
    ("Lignes modifiees", "Modifiees", "blue"),
    ("Lignes deja conformes", "Conformes", "green"),
    ("Lignes ajoutees", "Ajoutees", "primary"),
    ("Lignes Store vide", "Rouges", "red"),
    ("A verifier", "A verifier", "amber"),
]
# Le reste, en une ligne discrete
DETAILS = ["Lignes analysees", "Stores non trouves", "KAM trouves",
           "KAM non trouves", "Duplicates evites"]


# ---------------------------------------------------------------------------
# Petite couche de compatibilite CustomTkinter / Tkinter
# ---------------------------------------------------------------------------
def _button(parent, text, command, kind="normal", width=170):
    colors = {
        "primary": (C["primary"], C["primary_dark"]),
        "accent": (C["blue"], "#24567F"),
        "normal": ("#5A6A6B", "#3F4C4D"),
    }[kind]
    if HAS_CTK:
        return ctk.CTkButton(parent, text=text, command=command, font=FONT_BOLD,
                             fg_color=colors[0], hover_color=colors[1],
                             width=width, height=38, corner_radius=8)
    return tk.Button(parent, text=text, command=command, font=FONT_BOLD, bg=colors[0],
                     fg="white", activebackground=colors[1], activeforeground="white",
                     relief="flat", padx=14, pady=8, cursor="hand2",
                     disabledforeground="#C8CDD4")


def _entry(parent, textvariable):
    if HAS_CTK:
        return ctk.CTkEntry(parent, textvariable=textvariable, font=FONT, height=32)
    return tk.Entry(parent, textvariable=textvariable, font=FONT, relief="solid", bd=1, bg=C["soft"])


def _check(parent, text, variable):
    if HAS_CTK:
        return ctk.CTkCheckBox(parent, text=text, variable=variable, font=FONT,
                               checkbox_width=18, checkbox_height=18)
    return tk.Checkbutton(parent, text=text, variable=variable, font=FONT, bg=C["card"],
                          anchor="w", highlightthickness=0)


def _progressbar(parent):
    if HAS_CTK:
        bar = ctk.CTkProgressBar(parent, height=14, corner_radius=7,
                                 progress_color=C["primary"])
        bar.set(0)
        return bar
    return ttk.Progressbar(parent, mode="determinate", maximum=100)


def _set_progress(bar, value: float) -> None:
    value = max(0.0, min(1.0, value))
    if HAS_CTK:
        bar.set(value)
    else:
        bar["value"] = value * 100


def _set_state(widget, enabled: bool) -> None:
    try:
        widget.configure(state="normal" if enabled else "disabled")
    except Exception:                                           # pragma: no cover
        pass


def open_in_explorer(path: Path) -> None:
    """Ouvre un fichier avec l'application par defaut (Excel en general)."""
    path = Path(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))                             # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:                                    # pragma: no cover
        messagebox.showwarning("Ouverture impossible",
                               f"Ouvrez le fichier manuellement :\n{path}\n\n({exc})")


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
        self.root.title(APP_TITLE)
        width = min(1180, self.root.winfo_screenwidth() - 60)
        height = min(900, self.root.winfo_screenheight() - 70)
        self.root.geometry(f"{max(width, 960)}x{max(height, 640)}")
        self.root.minsize(960, 640)
        self.root.configure(bg=C["bg"])

        # Etat
        self.queue: queue.Queue = queue.Queue()
        self.session: Session | None = None
        self.output_path: Path | None = None
        self.report_paths: list[Path] = []
        self.busy = False
        self.rows: list[tuple[str, str]] = []
        self.active_filter = "Tout"

        self.var_ref = tk.StringVar()
        self.var_target = tk.StringVar()
        self.var_kam = tk.StringVar()
        self.var_ref_sheet = tk.StringVar(value=AUTO)
        self.var_target_sheet = tk.StringVar(value=AUTO)
        self.var_kam_sheet = tk.StringVar(value=AUTO)
        self.var_status = tk.StringVar(value="Choisissez les 3 fichiers Excel pour commencer.")
        self.opt_idaya = tk.BooleanVar(value=False)
        self.stat_vars: dict[str, tk.StringVar] = {}
        self.var_details = tk.StringVar(value="")
        self.var_alert = tk.StringVar(value="")
        self.path_entries: dict[str, object] = {}
        self.sheet_combos: dict[str, ttk.Combobox] = {}
        self.step_badges: dict[int, tk.Label] = {}

        self._build_ui()
        self._load_config()
        self.root.after(100, self._poll_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Construction de l'interface
    # ------------------------------------------------------------------
    def _card(self, parent) -> tk.Frame:
        outer = tk.Frame(parent, bg=C["line"], bd=0)
        outer.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(outer, bg=C["card"])
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return inner

    def _step_header(self, parent, number: int, title: str, hint: str) -> None:
        head = tk.Frame(parent, bg=C["card"])
        head.pack(fill="x", padx=16, pady=(12, 6))
        badge = tk.Label(head, text=str(number), font=FONT_BOLD, bg=C["gray_bg"],
                         fg=C["muted"], width=3, pady=2)
        badge.pack(side="left")
        self.step_badges[number] = badge
        tk.Label(head, text=title, font=FONT_STEP, bg=C["card"], fg=C["ink"]).pack(side="left", padx=9)
        tk.Label(head, text=hint, font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(side="left")

    def _mark_step(self, number: int, done: bool) -> None:
        badge = self.step_badges.get(number)
        if badge is None:
            return
        if done:
            badge.configure(text="OK", bg=C["green_bg"], fg=C["green"])
        else:
            badge.configure(text=str(number), bg=C["gray_bg"], fg=C["muted"])

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=C["bg"])
        container.pack(fill="both", expand=True, padx=14, pady=10)

        header = tk.Frame(container, bg=C["bg"])
        header.pack(fill="x", pady=(0, 10))
        tk.Label(header, text=APP_TITLE, font=FONT_TITLE, bg=C["bg"],
                 fg=C["primary"]).pack(side="left")
        tk.Label(header, text="  Reference (jamais modifiee)  +  BDD a traiter  +  Affectation KAM",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(side="left")

        self._build_step1(container)
        self._build_step2(container)
        self._build_step3(container)

    # ------------------------------------------------------------------
    def _build_step1(self, parent) -> None:
        card = self._card(parent)
        self._step_header(card, 1, "Choisir les 3 fichiers",
                          "la feuille est detectee toute seule")
        grid = tk.Frame(card, bg=C["card"])
        grid.pack(fill="x", padx=16, pady=(0, 14))

        rows = (
            ("Fichier Reference", "jamais modifie", self.var_ref, self.var_ref_sheet, "reference"),
            ("Fichier a traiter", "le seul modifie", self.var_target, self.var_target_sheet, "target"),
            ("Affectation KAM", "feuilles VD / DA", self.var_kam, self.var_kam_sheet, "kam"),
        )
        for index, (label, hint, path_var, sheet_var, kind) in enumerate(rows):
            tk.Label(grid, text=label, font=FONT_BOLD, bg=C["card"], fg=C["ink"],
                     anchor="w", width=17).grid(row=index, column=0, sticky="w", pady=4)
            tk.Label(grid, text=hint, font=FONT_SMALL, bg=C["card"], fg=C["muted"],
                     anchor="w", width=15).grid(row=index, column=1, sticky="w")
            entry = _entry(grid, path_var)
            entry.grid(row=index, column=2, sticky="ew", padx=8, pady=4)
            self.path_entries[kind] = entry
            _button(grid, "Parcourir", lambda k=kind: self.choose_file(k),
                    kind="normal", width=110).grid(row=index, column=3, padx=4)
            combo = ttk.Combobox(grid, textvariable=sheet_var, values=[AUTO],
                                 state="readonly", width=24, font=FONT)
            combo.grid(row=index, column=4, padx=(6, 0))
            combo.bind("<<ComboboxSelected>>", lambda _e: self._reset_analysis())
            self.sheet_combos[kind] = combo
        grid.grid_columnconfigure(2, weight=1)

    # ------------------------------------------------------------------
    def _build_step2(self, parent) -> None:
        card = self._card(parent)
        self._step_header(card, 2, "Analyser", "aucun fichier n'est ecrit a cette etape")
        body = tk.Frame(card, bg=C["card"])
        body.pack(fill="x", padx=16, pady=(0, 14))

        line = tk.Frame(body, bg=C["card"])
        line.pack(fill="x")
        self.btn_analyze = _button(line, "Lancer l'analyse", self.on_analyze,
                                   kind="primary", width=190)
        self.btn_analyze.pack(side="left")
        _check(line, "Numeroter IDAYA sur les lignes creees", self.opt_idaya).pack(side="left", padx=16)
        self.opt_idaya.trace_add("write", lambda *_a: self._reset_analysis())

        self.progress = _progressbar(body)
        self.progress.pack(fill="x", pady=(12, 6))
        tk.Label(body, textvariable=self.var_status, font=FONT, bg=C["card"],
                 fg=C["muted"], anchor="w", justify="left").pack(fill="x")

    # ------------------------------------------------------------------
    def _build_step3(self, parent) -> None:
        card = self._card(parent)
        card.master.pack_configure(fill="both", expand=True)
        card.pack_configure(fill="both", expand=True)
        self._step_header(card, 3, "Resultat", "verifiez, puis generez le fichier")
        body = tk.Frame(card, bg=C["card"])
        body.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        # --- cartes de chiffres
        cards = tk.Frame(body, bg=C["card"])
        cards.pack(fill="x")
        for label, court, couleur in CARTES:
            box = tk.Frame(cards, bg=C[f"{couleur}_bg"], highlightbackground=C["line"],
                           highlightthickness=1)
            box.pack(side="left", expand=True, fill="x", padx=(0, 8))
            var = tk.StringVar(value="-")
            self.stat_vars[label] = var
            tk.Label(box, textvariable=var, font=FONT_BIG, bg=C[f"{couleur}_bg"],
                     fg=C[couleur]).pack(pady=(8, 0))
            tk.Label(box, text=court, font=FONT_SMALL, bg=C[f"{couleur}_bg"],
                     fg=C[couleur]).pack(pady=(0, 8))

        tk.Label(body, textvariable=self.var_details, font=FONT_SMALL, bg=C["card"],
                 fg=C["muted"], anchor="w").pack(fill="x", pady=(8, 0))

        # --- bandeau d'alerte (a verifier)
        self.alert = tk.Label(body, textvariable=self.var_alert, font=FONT_BOLD,
                              bg=C["amber_bg"], fg=C["amber"], anchor="w", padx=12, pady=7)

        # --- boutons d'action
        actions = tk.Frame(body, bg=C["card"])
        actions.pack(fill="x", pady=10)
        self.actions = actions
        self.btn_generate = _button(actions, "Generer le fichier", self.on_generate,
                                    kind="accent", width=200)
        self.btn_generate.pack(side="left")
        self.btn_open = _button(actions, "Ouvrir le resultat", self.on_open_result, width=170)
        self.btn_open.pack(side="left", padx=8)
        self.btn_report = _button(actions, "Ouvrir le rapport", self.on_open_report, width=170)
        self.btn_report.pack(side="left")
        for btn in (self.btn_generate, self.btn_open, self.btn_report):
            _set_state(btn, False)

        # --- filtres du rapport
        self.filters = tk.Frame(body, bg=C["card"])
        self.filters.pack(fill="x", pady=(4, 6))
        self.filter_buttons: dict[str, tk.Button] = {}

        # --- resume des colonnes detectees : place reservee en bas
        self.var_columns = tk.StringVar(value="")
        tk.Label(body, textvariable=self.var_columns, font=("Consolas", 8), bg=C["card"],
                 fg=C["muted"], anchor="w", justify="left").pack(side="bottom", fill="x", pady=(6, 0))

        # --- rapport (prend toute la place restante)
        text_frame = tk.Frame(body, bg=C["card"])
        text_frame.pack(fill="both", expand=True)
        self.report_text = tk.Text(text_frame, font=("Consolas", 9), wrap="none",
                                   relief="solid", bd=1, bg=C["soft"], height=10)
        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.report_text.yview)
        xscroll = ttk.Scrollbar(text_frame, orient="horizontal", command=self.report_text.xview)
        self.report_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set,
                                   state="disabled")
        self.report_text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        for level, color in LEVEL_COLORS.items():
            self.report_text.tag_configure(level, foreground=color)

    # ------------------------------------------------------------------
    # Memoire des derniers fichiers utilises
    # ------------------------------------------------------------------
    def _load_config(self) -> None:
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        for key, var in (("reference", self.var_ref), ("target", self.var_target),
                         ("kam", self.var_kam)):
            path = data.get(key, "")
            if path and Path(path).exists():
                var.set(path)
                self._fill_sheets(key, path)
        if any(v.get() for v in (self.var_ref, self.var_target, self.var_kam)):
            self.var_status.set("Derniers fichiers utilises recharges. Verifiez-les puis lancez l'analyse.")
        self._refresh_step1()

    def _save_config(self) -> None:
        try:
            CONFIG_FILE.write_text(json.dumps({
                "reference": self.var_ref.get(),
                "target": self.var_target.get(),
                "kam": self.var_kam.get(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:                                       # pragma: no cover
            pass

    # ------------------------------------------------------------------
    # Actions utilisateur
    # ------------------------------------------------------------------
    def _fill_sheets(self, kind: str, path: str) -> None:
        sheets = list_sheet_names(path)
        combo = self.sheet_combos.get(kind)
        if combo is not None:
            combo["values"] = [AUTO] + sheets
        entry = self.path_entries.get(kind)
        if entry is not None:
            try:
                entry.xview_moveto(1.0)                         # montrer le nom du fichier
            except Exception:
                pass

    def choose_file(self, kind: str) -> None:
        titles = {
            "reference": "Fichier 1 - Reference (jamais modifie)",
            "target": "Fichier 2 - BDD a traiter",
            "kam": "Fichier 3 - Affectation des KAM",
        }
        variables = {"reference": self.var_ref, "target": self.var_target, "kam": self.var_kam}
        current = variables[kind].get()
        path = filedialog.askopenfilename(
            title=titles[kind],
            initialdir=str(Path(current).parent) if current else None,
            filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        variables[kind].set(path)
        {"reference": self.var_ref_sheet, "target": self.var_target_sheet,
         "kam": self.var_kam_sheet}[kind].set(AUTO)
        self._fill_sheets(kind, path)
        self._save_config()
        self._reset_analysis()
        self.var_status.set(f"{titles[kind]} : {Path(path).name}")
        self._refresh_step1()

    def _refresh_step1(self) -> None:
        pret = all(v.get().strip() for v in (self.var_ref, self.var_target, self.var_kam))
        self._mark_step(1, pret)

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

    def _reset_analysis(self) -> None:
        """Changer un fichier ou une option invalide l'apercu precedent."""
        if self.busy:
            return
        self._mark_step(2, False)
        self._mark_step(3, False)
        for btn in (self.btn_generate, self.btn_open, self.btn_report):
            _set_state(btn, False)
        self.output_path = None
        self.report_paths = []
        self.alert.pack_forget()

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
        self.session = Session(selection, Options(auto_number_idaya=self.opt_idaya.get()))
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
        except Exception as exc:                                # pragma: no cover
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
                path, write_report=True,
                progress=lambda m, v: self.queue.put(("progress", m, v)),
            )
            self.queue.put(("written", result, reports))
        except PipelineError as exc:
            self.queue.put(("error", str(exc)))
        except Exception as exc:                                # pragma: no cover
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
        """Resume court : feuille retenue pour chaque fichier (le detail est dans le rapport)."""
        courtes = []
        for ligne in description.splitlines():
            titre, _, reste = ligne.partition(":")
            morceaux = [m.strip() for m in reste.split("|")]
            courtes.append(f"{titre.strip()} -> {' | '.join(morceaux[:2])}")
        text = "   ".join(courtes)
        if warnings:
            text += "\nATTENTION : " + " | ".join(warnings)
        self.var_columns.set(text)

    def _clear_report(self) -> None:
        self.rows = []
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        self.report_text.configure(state="disabled")
        for var in self.stat_vars.values():
            var.set("-")
        self.var_details.set("")
        self.var_alert.set("")
        self.alert.pack_forget()

    def _show_plan(self, plan) -> None:
        stats = dict(plan.stats.as_pairs())
        for label, var in self.stat_vars.items():
            var.set(str(stats.get(label, 0)))
        self.var_details.set("   ".join(f"{name} : {stats.get(name, 0)}" for name in DETAILS))

        self.rows = plan.logger.display_rows()
        self._build_filters()
        self._refresh_report()

        a_verifier = stats.get("A verifier", 0)
        if a_verifier:
            self.var_alert.set(
                f"  {a_verifier} point(s) a verifier : le logiciel n'a rien decide tout seul. "
                "Filtrez sur 'A verifier' ci-dessous avant de generer."
            )
            self.alert.pack(fill="x", pady=(8, 0), before=self.actions)
        else:
            self.alert.pack_forget()

        _set_progress(self.progress, 1.0)
        self.var_status.set(
            f"Analyse terminee : {stats.get('Lignes modifiees', 0)} a modifier, "
            f"{stats.get('Lignes ajoutees', 0)} a creer, {stats.get('Lignes Store vide', 0)} en rouge. "
            "Cliquez sur 'Generer le fichier'."
        )
        self._mark_step(2, True)
        _set_state(self.btn_generate, True)

    def _build_filters(self) -> None:
        for widget in self.filters.winfo_children():
            widget.destroy()
        self.filter_buttons = {}
        niveaux = ["Tout"] + sorted({level for level, _ in self.rows})
        tk.Label(self.filters, text="Filtrer :", font=FONT_SMALL, bg=C["card"],
                 fg=C["muted"]).pack(side="left", padx=(0, 6))
        for niveau in niveaux:
            nombre = len(self.rows) if niveau == "Tout" else sum(1 for l, _ in self.rows if l == niveau)
            couleur = LEVEL_COLORS.get(niveau, C["ink"])
            btn = tk.Button(self.filters, text=f"{niveau} ({nombre})", font=FONT_SMALL,
                            bg=C["soft"], fg=couleur, relief="solid", bd=1, padx=8, pady=2,
                            cursor="hand2", command=lambda n=niveau: self._set_filter(n))
            btn.pack(side="left", padx=2)
            self.filter_buttons[niveau] = btn
        self.active_filter = "Tout"
        self._highlight_filter()

    def _set_filter(self, niveau: str) -> None:
        self.active_filter = niveau
        self._highlight_filter()
        self._refresh_report()

    def _highlight_filter(self) -> None:
        for niveau, btn in self.filter_buttons.items():
            actif = niveau == self.active_filter
            btn.configure(bg=C["primary"] if actif else C["soft"],
                          fg="white" if actif else LEVEL_COLORS.get(niveau, C["ink"]))

    def _refresh_report(self) -> None:
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        shown = 0
        for level, text in self.rows:
            if self.active_filter != "Tout" and level != self.active_filter:
                continue
            self.report_text.insert("end", text + "\n", level)
            shown += 1
        if not shown:
            self.report_text.insert("end", "Aucune ligne dans ce filtre.")
        self.report_text.configure(state="disabled")

    def _show_written(self, result, reports: list[Path]) -> None:
        self.output_path = Path(result.output_path)
        self.report_paths = [Path(p) for p in reports]
        _set_state(self.btn_open, True)
        _set_state(self.btn_report, bool(reports))
        _set_progress(self.progress, 1.0)
        self._mark_step(3, True)
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
        self._save_config()
        if self.session is not None:
            self.session.close()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch() -> None:
    App().run()


if __name__ == "__main__":
    launch()
