"""
excel_writer.py
---------------
Ecriture du resultat.

Regle d'or : on n'ecrase JAMAIS les fichiers d'origine.
On part du classeur du fichier 2 (deja charge avec ses couleurs, ses tableaux,
ses filtres et ses formules), on y applique le plan, puis on enregistre sous un
NOUVEAU nom.
"""

from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from excel_reader import TargetData
from logger import ERREUR, INFO, ProcessLogger
from processor import Plan

# Rouge : seule couleur prevue par les regles (ligne dont le Store est vide).
RED_FILL = PatternFill("solid", start_color="FFFF0000", end_color="FFFF0000")

_INT_RE = re.compile(r"^-?\d+$")
_FLOAT_RE = re.compile(r"^-?\d+[.,]\d+$")


class ExcelWriteError(Exception):
    """Erreur d'ecriture affichable telle quelle dans l'interface."""


@dataclass
class WriteResult:
    output_path: Path
    updated_cells: int = 0
    created_rows: int = 0
    colored_rows: int = 0
    warnings: list[str] = field(default_factory=list)


# Seules ces colonnes sont ecrites comme des NOMBRES (pour les sommes et les filtres).
# Tout le reste reste du texte : un code magasin ou un ID promoteur ne doit jamais
# etre transforme en nombre (risque de perdre des zeros ou de casser les recherches).
NUMERIC_COLUMNS = {"DAY", "IDAYA", "Annee"}


def _coerce(text: str, column: str = ""):
    """Convertit la valeur texte en valeur Excel.

    'C000114327' -> texte (toujours)
    '26' en DAY  -> 26 (nombre)
    ''           -> None (cellule vide)
    """
    if text is None or text == "":
        return None
    if column not in NUMERIC_COLUMNS:
        return text
    if _INT_RE.match(text):
        value = int(text)
        if str(value) == text:            # evite de perdre les zeros de gauche ('007')
            return value
        return text
    if _FLOAT_RE.match(text):
        try:
            return float(text.replace(",", "."))
        except ValueError:
            return text
    return text


def _used_columns(target: TargetData) -> tuple[int, int]:
    """Plage de colonnes 'utilisees' de la feuille (pour colorer une ligne entiere)."""
    ws, layout = target.worksheet, target.layout
    indexes = set(layout.columns.values()) | set(layout.header_labels.keys())
    if layout.table_name:
        table = ws.tables.get(layout.table_name)
        if table is not None:
            min_c, _, max_c, _ = range_boundaries(table.ref)
            indexes |= {min_c, max_c}
    if not indexes:
        return 1, max(ws.max_column or 1, 1)
    return min(indexes), max(indexes)


def _check_output_path(output_path: Path, sources: list[Path]) -> None:
    out = output_path.expanduser().resolve()
    for src in sources:
        try:
            if src and out == Path(src).expanduser().resolve():
                raise ExcelWriteError(
                    "Le fichier de sortie porte le meme chemin qu'un fichier source.\n"
                    "Les fichiers d'origine ne doivent jamais etre ecrases : choisissez un autre nom."
                )
        except OSError:                                     # pragma: no cover
            continue
    if not out.parent.exists():
        raise ExcelWriteError(f"Le dossier de destination n'existe pas : {out.parent}")


def apply_plan(
    target: TargetData,
    plan: Plan,
    output_path: str | Path,
    sources: list[Path] | None = None,
    logger: ProcessLogger | None = None,
) -> WriteResult:
    """Applique le plan sur le classeur du fichier 2, puis enregistre le resultat."""
    log = logger or plan.logger
    output_path = Path(output_path)
    _check_output_path(output_path, sources or [])

    ws = target.worksheet
    layout = target.layout
    result = WriteResult(output_path=output_path)
    first_col, last_col = _used_columns(target)

    # ------------------------------------------------------------------
    # 1. Modification des lignes existantes
    # ------------------------------------------------------------------
    for update in plan.updates:
        for change in update.changes:
            col = layout.col(change.column)
            if not col:
                continue
            if _is_formula(ws.cell(update.row, col).value):
                # Securite : une cellule calculee n'est jamais ecrasee.
                log.log(INFO, f"Formule conservee dans la colonne {change.column}",
                        row=update.row, store=update.store, division=update.division)
                continue
            ws.cell(update.row, col).value = _coerce(change.new, change.column)
            result.updated_cells += 1

    # ------------------------------------------------------------------
    # 2. Lignes a colorer en rouge (Store vide, regle 2)
    # ------------------------------------------------------------------
    for mark in plan.marks:
        for col in range(first_col, last_col + 1):
            ws.cell(mark.row, col).fill = RED_FILL
        result.colored_rows += 1

    # ------------------------------------------------------------------
    # 3. Creation des nouvelles lignes (a la fin du tableau)
    # ------------------------------------------------------------------
    if plan.creations:
        start_row = layout.last_data_row + 1
        count = len(plan.creations)

        # Si la zone d'ecriture n'est pas vide, on insere des lignes pour ne rien ecraser.
        occupied = any(
            ws.cell(r, c).value not in (None, "")
            for r in range(start_row, start_row + count)
            for c in range(first_col, last_col + 1)
        )
        if occupied:
            ws.insert_rows(start_row, amount=count)
            result.warnings.append(
                f"{count} ligne(s) inserees en ligne {start_row} : le contenu situe en dessous a ete decale."
            )

        for offset, creation in enumerate(plan.creations):
            row = start_row + offset
            template = creation.template_row or layout.last_data_row
            # 1. La nouvelle ligne est une COPIE de la ligne soeur : mise en forme,
            #    valeurs et formules des colonnes qu'on ne remplit pas (demande
            #    explicite : "les champs que je ne remplis pas -> duplicate").
            _copy_row_style(ws, template, row, first_col, last_col)
            _duplicate_row_values(ws, template, row, first_col, last_col)
            # 2. Puis on ecrit par-dessus ce que la Reference impose, sans jamais
            #    ecraser une cellule qui contient une formule.
            for name, col in layout.columns.items():
                if name not in creation.values:
                    continue
                if _is_formula(ws.cell(row, col).value):
                    log.log(INFO, f"Ligne {row} : formule conservee dans la colonne {name}",
                            row=row, store=creation.store, division=creation.division)
                    continue
                ws.cell(row, col).value = _coerce(creation.values.get(name, ""), name)
            result.created_rows += 1
            log.log(
                INFO,
                f"Ligne {row} ecrite dans le fichier resultat (cle {creation.values.get('Column1', '')})",
                row=row, store=creation.store, division=creation.division,
            )

        _extend_ranges(target, count, log, result)

    # ------------------------------------------------------------------
    # 4. Enregistrement sous un NOUVEAU fichier
    # ------------------------------------------------------------------
    try:
        target.workbook.save(output_path)
    except PermissionError as exc:
        raise ExcelWriteError(
            f"Impossible d'ecrire {output_path.name} : le fichier est peut-etre ouvert dans Excel.\n({exc})"
        ) from exc
    except Exception as exc:                                # pragma: no cover - cas rare
        log.log(ERREUR, f"Echec de l'enregistrement : {exc}")
        raise ExcelWriteError(f"Echec de l'enregistrement : {exc}") from exc

    return result


def _is_formula(value) -> bool:
    """Vrai si la cellule contient une formule Excel (=...)."""
    return isinstance(value, str) and value.startswith("=")


def _duplicate_row_values(ws, source_row: int, dest_row: int, first_col: int, last_col: int) -> None:
    """Recopie les VALEURS d'une ligne modele vers une nouvelle ligne.

    Les formules sont recopiees en ajustant leurs references de ligne
    (=A2*B2 en ligne 2 devient =A3*B3 en ligne 3), exactement comme un
    copier-coller Excel.
    """
    from openpyxl.formula.translate import Translator

    for col in range(first_col, last_col + 1):
        source = ws.cell(source_row, col)
        value = source.value
        if _is_formula(value):
            try:
                value = Translator(value, origin=source.coordinate).translate_formula(
                    ws.cell(dest_row, col).coordinate
                )
            except Exception:                                # pragma: no cover - formule exotique
                pass
        ws.cell(dest_row, col).value = value


def _copy_row_style(ws, source_row: int, dest_row: int, first_col: int, last_col: int) -> None:
    """Recopie la mise en forme d'une ligne modele vers une nouvelle ligne."""
    for col in range(first_col, last_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(dest_row, col)
        try:
            dst._style = copy(src._style)                   # police, fond, bordures, format nombre
        except Exception:                                   # pragma: no cover - securite
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
    height = ws.row_dimensions[source_row].height
    if height:
        ws.row_dimensions[dest_row].height = height


def _extend_ranges(target: TargetData, added: int, log: ProcessLogger, result: WriteResult) -> None:
    """Agrandit le tableau Excel et le filtre automatique pour inclure les nouvelles lignes."""
    ws, layout = target.worksheet, target.layout

    if layout.table_name and layout.table_name in ws.tables:
        table = ws.tables[layout.table_name]
        min_c, min_r, max_c, max_r = range_boundaries(table.ref)
        new_ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r + added}"
        table.ref = new_ref
        if table.autoFilter is not None:
            table.autoFilter.ref = new_ref
        log.info(f"Tableau '{layout.table_name}' agrandi : {new_ref}")

    if ws.auto_filter and ws.auto_filter.ref:
        min_c, min_r, max_c, max_r = range_boundaries(ws.auto_filter.ref)
        ws.auto_filter.ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r + added}"
        log.info(f"Filtre automatique agrandi : {ws.auto_filter.ref}")

    if ws.data_validations.dataValidation:
        result.warnings.append(
            "Cette feuille contient des listes deroulantes (validation de donnees) : "
            "verifiez qu'elles couvrent bien les nouvelles lignes."
        )


def suggest_output_path(target_path: str | Path, suffix: str = "TRAITE") -> Path:
    """Propose un nom de fichier resultat : 'BDD_AOUT.xlsx' -> 'BDD_AOUT_TRAITE.xlsx'."""
    path = Path(target_path)
    return path.with_name(f"{path.stem}_{suffix}{path.suffix or '.xlsx'}")
