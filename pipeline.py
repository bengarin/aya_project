"""
pipeline.py
-----------
Le "chef d'orchestre" : il enchaine lecture -> analyse -> ecriture.

L'interface graphique (gui.py) et la ligne de commande (main.py --cli)
utilisent toutes les deux ce fichier, pour garantir exactement le meme
comportement metier dans les deux cas.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import excel_reader as reader
from excel_reader import ExcelReadError, ReferenceData, TargetData
from excel_writer import ExcelWriteError, WriteResult, apply_plan, suggest_output_path
from kam_service import KamService
from logger import ProcessLogger
from processor import Options, Plan, Processor

Progress = Callable[[str, float], None]        # (message, avancement 0..1)


class PipelineError(Exception):
    """Erreur metier destinee a etre affichee a l'utilisateur."""


@dataclass
class Selection:
    """Les 3 fichiers choisis par l'utilisateur (+ feuilles forcees si besoin)."""

    reference: str | Path | None = None
    target: str | Path | None = None
    kam: str | Path | None = None
    reference_sheet: str | None = None
    target_sheet: str | None = None
    kam_sheets: list[str] | None = None

    def validate(self) -> None:
        missing = [
            label for label, value in (
                ("Fichier 1 (Reference)", self.reference),
                ("Fichier 2 (BDD a traiter)", self.target),
                ("Fichier 3 (Affectation KAM)", self.kam),
            ) if not value
        ]
        if missing:
            raise PipelineError("Fichier(s) manquant(s) : " + ", ".join(missing))
        for label, value in (
            ("Fichier 1", self.reference), ("Fichier 2", self.target), ("Fichier 3", self.kam)
        ):
            path = Path(str(value))
            if not path.exists():
                raise PipelineError(f"{label} introuvable : {path}")
            if path.suffix.lower() not in (".xlsx", ".xlsm"):
                raise PipelineError(f"{label} : format non supporte ({path.suffix}). Utilisez .xlsx ou .xlsm")


@dataclass
class LoadedData:
    reference: ReferenceData
    target: TargetData
    kam: KamService
    warnings: list[str] = field(default_factory=list)

    def describe(self) -> str:
        lines = [
            "Fichier 1 (Reference)      : " + self.reference.layout.describe(),
            "Fichier 2 (BDD a traiter)  : " + self.target.layout.describe(),
            "Fichier 3 (Affectation KAM): "
            + ", ".join(f"feuille '{l.title}' ({l.data_rows} lignes)" for l in self.kam.data.layouts)
            + f" | {self.kam.summary()}",
        ]
        return "\n".join(lines)


class Session:
    """Une session de travail : charge les fichiers, analyse, puis ecrit le resultat."""

    def __init__(self, selection: Selection, options: Options | None = None) -> None:
        self.selection = selection
        self.options = options or Options()
        self.data: LoadedData | None = None
        self.plan: Plan | None = None
        self.logger = ProcessLogger()
        self._applied = False

    # ------------------------------------------------------------------
    def load(self, progress: Progress | None = None) -> LoadedData:
        """Ouvre les 3 fichiers et detecte feuilles + colonnes."""
        self.selection.validate()
        self.close()
        self._applied = False
        try:
            if progress:
                progress("Lecture du fichier Reference...", 0.05)
            reference = reader.load_reference(self.selection.reference, self.selection.reference_sheet)

            if progress:
                progress("Lecture du fichier Affectation KAM...", 0.15)
            kam = KamService(reader.load_kam(self.selection.kam, self.selection.kam_sheets))

            if progress:
                progress("Lecture du fichier a traiter (peut prendre quelques secondes)...", 0.25)
            target = reader.load_target(self.selection.target, self.selection.target_sheet)
        except ExcelReadError as exc:
            raise PipelineError(str(exc)) from exc
        except Exception as exc:                            # pragma: no cover - fichier corrompu
            raise PipelineError(f"Lecture impossible : {exc}") from exc

        warnings = list(target.warnings)
        lost = reader.detect_unsupported_parts(self.selection.target)
        if lost:
            warnings.append(
                "Elements du fichier 2 non recopiables par openpyxl (ils seront absents du resultat) : "
                + ", ".join(lost) + ". Les donnees, formules, couleurs, tableaux et filtres sont conserves."
            )
        if not kam.divisions:
            warnings.append(
                "Le fichier 3 ne contient pas de feuille 'VD' / 'DA' : le KAM sera cherche "
                "uniquement par nom de magasin (toutes divisions confondues)."
            )
        elif len(kam.divisions) == 1:
            warnings.append(
                f"Le fichier 3 ne contient qu'une feuille de division ({kam.divisions[0]}) : "
                "les autres divisions n'auront pas de KAM."
            )
        self.data = LoadedData(reference=reference, target=target, kam=kam, warnings=warnings)
        if progress:
            progress("Fichiers charges.", 0.35)
        return self.data

    # ------------------------------------------------------------------
    def analyze(self, progress: Progress | None = None) -> Plan:
        """Applique les regles metier SANS rien ecrire (mode Preview)."""
        if self.data is None or self._applied:
            self.load(progress)
        assert self.data is not None
        self.logger = ProcessLogger()
        processor = Processor(
            reference=self.data.reference, target=self.data.target,
            kam=self.data.kam, options=self.options, logger=self.logger,
        )

        def on_row(done: int, total: int) -> None:
            if progress and (done % 25 == 0 or done == total):
                progress(f"Analyse des lignes {done}/{total}...", 0.35 + 0.55 * done / max(total, 1))

        self.plan = processor.run(progress=on_row)
        if progress:
            progress("Analyse terminee.", 0.95)
        return self.plan

    # ------------------------------------------------------------------
    def apply(self, output_path: str | Path, write_report: bool = True,
              progress: Progress | None = None) -> tuple[WriteResult, list[Path]]:
        """Ecrit le fichier resultat (+ le rapport) sans toucher aux fichiers sources."""
        if self.plan is None or self.data is None:
            raise PipelineError("Lancez d'abord l'analyse (Preview) avant de generer le fichier.")
        if self._applied:
            raise PipelineError("Ce plan a deja ete ecrit. Relancez l'analyse avant de regenerer un fichier.")

        if progress:
            progress("Ecriture du fichier resultat...", 0.1)
        sources = [Path(str(p)) for p in (self.selection.reference, self.selection.target, self.selection.kam)]
        try:
            result = apply_plan(
                target=self.data.target, plan=self.plan, output_path=output_path,
                sources=sources, logger=self.logger,
            )
        except ExcelWriteError as exc:
            raise PipelineError(str(exc)) from exc
        self._applied = True

        reports: list[Path] = []
        if write_report:
            if progress:
                progress("Ecriture du rapport...", 0.8)
            header = self.report_header(result)
            base = Path(output_path)
            try:
                reports.append(self.logger.save_text(base.with_name(base.stem + "_RAPPORT.txt"), header))
                reports.append(self.logger.save_xlsx(base.with_name(base.stem + "_RAPPORT.xlsx"), header))
            except Exception as exc:                        # pragma: no cover
                self.logger.error(f"Rapport non enregistre : {exc}")
        if progress:
            progress("Termine.", 1.0)
        return result, reports

    # ------------------------------------------------------------------
    def report_header(self, result: WriteResult | None = None) -> str:
        lines = [
            "AUTOMATISATION EXCEL - AYA",
            f"Fichier 1 (Reference, lecture seule) : {self.selection.reference}",
            f"Fichier 2 (BDD a traiter)            : {self.selection.target}",
            f"Fichier 3 (Affectation KAM)          : {self.selection.kam}",
        ]
        if self.data:
            lines.append(self.data.describe())
        if self.plan:
            lines.append("-" * 100)
            lines += [f"{label:<32}: {value}" for label, value in self.plan.stats.as_pairs()]
        if result:
            lines.append(f"Fichier resultat                : {result.output_path}")
        return "\n".join(lines)

    # ------------------------------------------------------------------
    def default_output(self, suffix: str = "TRAITE") -> Path:
        return suggest_output_path(self.selection.target or "resultat.xlsx", suffix)

    def close(self) -> None:
        if self.data is not None:
            self.data.target.close()
            self.data = None
        self.plan = None
