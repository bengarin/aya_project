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
import zipfile
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
    # 3. Creation des nouvelles lignes
    #    Comme "dupliquer la ligne" dans Excel : la nouvelle ligne est inseree
    #    JUSTE SOUS la ligne du meme Store (VD sous DA, ou DA sous VD).
    # ------------------------------------------------------------------
    if plan.creations:
        # Les traces Power Query sont retirees d'abord : leurs noms caches ne
        # doivent pas empecher l'insertion.
        _neutraliser_tableaux_de_requete(target.workbook, log, result)
        positions = _ecrire_creations(target, plan, first_col, last_col, log, result)
        _extend_ranges(target, len(plan.creations), log, result)
        _renumeroter_rapport(plan, log, positions)

    # ------------------------------------------------------------------
    # 4. Nettoyage avant enregistrement (evite le message de reparation d'Excel)
    # ------------------------------------------------------------------
    _neutraliser_tableaux_de_requete(target.workbook, log, result)
    try:
        # Excel recalcule tout a l'ouverture : les formules affichent des valeurs
        # a jour meme si leurs cellules sources viennent d'etre corrigees.
        target.workbook.calculation.fullCalcOnLoad = True
    except Exception:                                       # pragma: no cover
        pass
    valeurs_calculees = _valeurs_calculees_source(Path(target.path))

    # ------------------------------------------------------------------
    # 5. Enregistrement sous un NOUVEAU fichier
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

    # ------------------------------------------------------------------
    # 6. Apres sauvegarde : remettre les resultats de formules, puis verifier
    #    que le fichier ne contient aucune reference cassee (sinon Excel
    #    afficherait son message de reparation a l'ouverture).
    # ------------------------------------------------------------------
    _restaurer_valeurs_calculees(output_path, valeurs_calculees, log, result)
    for probleme in verifier_fichier(output_path):
        result.warnings.append(f"Verification du fichier : {probleme}")
        log.log(ERREUR, f"Verification du fichier resultat : {probleme}")

    return result


_CELL_RE = re.compile(
    rb'<c\b(?P<attrs>[^>]*?)\br="(?P<ref>[A-Z]+\d+)"(?P<rest>[^>]*?)(?:/>|>(?P<body>.*?)</c>)',
    re.S,
)
_T_ATTR_RE = re.compile(rb'\bt="([^"]*)"')
_V_VIDE_RE = re.compile(rb"<v\s*/>|<v></v>")


def _feuilles_vers_parties(archive: zipfile.ZipFile) -> dict[str, str]:
    """Associe le nom de chaque feuille au fichier XML qui la contient."""
    classeur = archive.read("xl/workbook.xml").decode("utf-8", "ignore")
    relations = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    cibles: dict[str, str] = {}
    for bloc in re.findall(r"<Relationship\b[^>]*>", relations):
        identifiant = re.search(r'Id="([^"]+)"', bloc)
        cible = re.search(r'Target="([^"]+)"', bloc)
        if identifiant and cible:
            cibles[identifiant.group(1)] = cible.group(1)
    parties: dict[str, str] = {}
    for bloc in re.findall(r"<sheet\b[^>]*>", classeur):
        nom = re.search(r'name="([^"]+)"', bloc)
        identifiant = re.search(r'r:id="([^"]+)"', bloc) or re.search(r'\bid="([^"]+)"', bloc)
        if not (nom and identifiant):
            continue
        cible = (cibles.get(identifiant.group(1), "") or "").lstrip("/")
        parties[nom.group(1)] = cible if cible.startswith("xl/") else "xl/" + cible
    return parties


def _valeurs_calculees_source(chemin: Path) -> dict[str, dict[bytes, tuple[bytes | None, bytes]]]:
    """Releve, dans le fichier d'origine, le resultat de chaque formule.

    openpyxl ecrit la formule mais pas son resultat : a l'ouverture, Excel
    affiche des cellules vides le temps de recalculer, et signale parfois une
    reparation sur les formules qui pointent vers un autre classeur. On releve
    donc les resultats ici pour les remettre apres la sauvegarde.

    Resultat : {nom de feuille : {reference cellule : (type, <v>...</v>)}}
    """
    valeurs: dict[str, dict[bytes, tuple[bytes | None, bytes]]] = {}
    try:
        with zipfile.ZipFile(chemin) as archive:
            for feuille, partie in _feuilles_vers_parties(archive).items():
                if partie not in archive.namelist():
                    continue
                xml = archive.read(partie)
                if b"<f" not in xml:
                    continue
                cellules: dict[bytes, tuple[bytes | None, bytes]] = {}
                for trouve in _CELL_RE.finditer(xml):
                    corps = trouve.group("body")
                    if not corps or b"<f" not in corps:
                        continue
                    valeur = re.search(rb"<v[^>]*>.*?</v>", corps, re.S)
                    if not valeur or valeur.group(0) in (b"<v></v>",):
                        continue
                    attributs = trouve.group("attrs") + trouve.group("rest")
                    type_cellule = _T_ATTR_RE.search(attributs)
                    cellules[trouve.group("ref")] = (
                        type_cellule.group(1) if type_cellule else None,
                        valeur.group(0),
                    )
                if cellules:
                    valeurs[feuille] = cellules
    except Exception:                                           # pragma: no cover - securite
        return {}
    return valeurs


def _restaurer_valeurs_calculees(sortie: Path, valeurs: dict[str, dict[bytes, tuple[bytes | None, bytes]]],
                                 log: ProcessLogger, result: WriteResult) -> None:
    """Remet dans le fichier resultat les resultats de formules releves plus haut.

    Seules les cellules qui contiennent une formule sans resultat sont touchees.
    Les cellules ecrites par le logiciel ne sont jamais concernees : il n'ecrit
    jamais dans une cellule contenant une formule.
    """
    if not valeurs:
        return
    try:
        with zipfile.ZipFile(sortie) as archive:
            ordre = archive.namelist()
            parties = {nom: archive.read(nom) for nom in ordre}
            feuilles = _feuilles_vers_parties(archive)

        restaurees = 0

        def reparer(trouve, table):
            nonlocal restaurees
            corps = trouve.group("body")
            if not corps or b"<f" not in corps:
                return trouve.group(0)
            attendu = table.get(trouve.group("ref"))
            if not attendu:
                return trouve.group(0)
            type_cellule, valeur = attendu
            corps_repare, remplacements = _V_VIDE_RE.subn(valeur, corps, count=1)
            if remplacements == 0:
                if b"<v" in corps:
                    return trouve.group(0)                      # resultat deja present
                corps_repare = corps + valeur
            entier = trouve.group(0)
            ouvrant = entier[: entier.index(b">") + 1]
            if type_cellule and not _T_ATTR_RE.search(ouvrant):
                ouvrant = ouvrant[:-1] + b' t="' + type_cellule + b'">'
            restaurees += 1
            return ouvrant + corps_repare + b"</c>"

        for feuille, table in valeurs.items():
            partie = feuilles.get(feuille)
            xml = parties.get(partie) if partie else None
            if xml is None or b"<f" not in xml:
                continue
            parties[partie] = _CELL_RE.sub(lambda m, t=table: reparer(m, t), xml)

        if not restaurees:
            return
        with zipfile.ZipFile(sortie, "w", zipfile.ZIP_DEFLATED) as archive:
            for nom in ordre:
                archive.writestr(nom, parties[nom])
        log.info(f"{restaurees} resultat(s) de formule recopie(s) depuis le fichier d'origine.")
    except Exception as exc:                                    # pragma: no cover - securite
        result.warnings.append(
            "Les resultats des formules n'ont pas pu etre recopies "
            f"({exc}) : Excel les recalculera a l'ouverture."
        )


def _neutraliser_tableaux_de_requete(workbook, log: ProcessLogger, result: WriteResult) -> None:
    """Enleve du classeur TOUTES les traces de Power Query qu'openpyxl ne recopie pas.

    Pourquoi : le tableau de la feuille traitee est un tableau Power Query.
    openpyxl ne sait pas recopier la requete ni sa connexion, mais il gardait
    3 traces qui pointent vers elles. Excel les trouve orphelines et affiche,
    a chaque ouverture, "Excel a pu ouvrir le fichier en supprimant ou en
    reparant le contenu illisible" :

      1. l'etiquette du tableau (tableType="queryTable")
         -> "Enregistrements supprimes : Tableau dans /xl/tables/tableN.xml"
      2. sur chaque colonne du tableau, le lien vers le champ de la requete
         (queryTableFieldId, uniqueName) -> meme message
      3. les noms caches "ExternalData_1, _2..." qui designent la zone remplie
         par la requete -> "Partie supprimee : Plage de donnees externes"

    Une fois ces traces retirees, le tableau devient un tableau Excel normal
    (meme plage, meme style, memes filtres) et le fichier s'ouvre sans message.
    Seul le lien vers la requete, deja perdu, disparait.
    """
    transformes: list[str] = []
    for sheet in workbook.worksheets:
        for table in getattr(sheet, "tables", {}).values():
            touche = False
            if getattr(table, "tableType", None):
                table.tableType = None
                touche = True
            for colonne in getattr(table, "tableColumns", None) or []:
                if getattr(colonne, "queryTableFieldId", None) is not None:
                    colonne.queryTableFieldId = None
                    touche = True
                if getattr(colonne, "uniqueName", None) is not None:
                    colonne.uniqueName = None
                    touche = True
            if touche:
                transformes.append(f"{sheet.title}!{table.name}")

    noms_retires: list[str] = []
    conteneurs = [("", getattr(workbook, "defined_names", None))]
    conteneurs += [(f"{ws.title}!", getattr(ws, "defined_names", None)) for ws in workbook.worksheets]
    for prefixe, noms in conteneurs:
        if not noms:
            continue
        for nom in [n for n in list(noms.keys()) if str(n).lower().startswith("externaldata_")]:
            try:
                del noms[nom]
                noms_retires.append(prefixe + str(nom))
            except Exception:                                   # pragma: no cover
                pass

    if transformes or noms_retires:
        details = []
        if transformes:
            details.append("tableau(x) Power Query converti(s) en tableau Excel normal : "
                           + ", ".join(transformes))
        if noms_retires:
            details.append(f"{len(noms_retires)} plage(s) de donnees externes orpheline(s) retiree(s)")
        message = ("Nettoyage Power Query : " + " ; ".join(details)
                   + " (la connexion a la requete n'est pas recopiable ; "
                     "le fichier s'ouvre ainsi sans message de reparation).")
        log.info(message)
        result.warnings.append(message)


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
        if getattr(table, "sortState", None) is not None:
            # Le dernier tri memorise ne couvre plus toute la plage : on l'oublie
            # (les donnees ne bougent pas, Excel ne garde simplement plus ce tri).
            table.sortState = None
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


# ---------------------------------------------------------------------------
# Verification du fichier produit
# ---------------------------------------------------------------------------
def verifier_fichier(chemin: str | Path) -> list[str]:
    """Cherche les defauts qui font apparaitre le message de reparation d'Excel.

    Excel refuse silencieusement un classeur dont une partie est annoncee mais
    absente : lien casse, tableau de requete sans requete, reference externe
    declaree sans son fichier. On verifie tout cela ici, avant que l'utilisateur
    n'ouvre le fichier.
    """
    chemin = Path(chemin)
    problemes: list[str] = []
    try:
        with zipfile.ZipFile(chemin) as archive:
            parties = set(archive.namelist())

            # 1. chaque cible de relation existe-t-elle ?
            for rels in [n for n in parties if n.endswith(".rels")]:
                base = "" if rels == "_rels/.rels" else rels.rsplit("/_rels/", 1)[0]
                xml = archive.read(rels).decode("utf-8", "ignore")
                for cible, mode in re.findall(r'Target="([^"]+)"(?:\s+TargetMode="([^"]+)")?', xml):
                    if mode == "External" or cible.startswith(("http", "file:", "\\\\")):
                        continue
                    brut = cible[1:] if cible.startswith("/") else (f"{base}/{cible}" if base else cible)
                    morceaux: list[str] = []
                    for bout in brut.split("/"):
                        if bout == "..":
                            if morceaux:
                                morceaux.pop()
                        elif bout not in ("", "."):
                            morceaux.append(bout)
                    if "/".join(morceaux) not in parties:
                        problemes.append(f"lien casse dans {rels} vers {cible}")

            # 2. tableau declare comme tableau de requete sans la requete
            for table in [n for n in parties if re.match(r"xl/tables/table\d+\.xml$", n)]:
                xml = archive.read(table).decode("utf-8", "ignore")
                rel = table.replace("xl/tables/", "xl/tables/_rels/") + ".rels"
                if 'tableType="queryTable"' in xml and rel not in parties:
                    problemes.append(f"tableau de requete sans requete : {table}")

            # 2b. colonnes encore reliees a un champ de requete dans un tableau normal
            for table in [n for n in parties if re.match(r"xl/tables/table\d+\.xml$", n)]:
                xml = archive.read(table).decode("utf-8", "ignore")
                if 'tableType="queryTable"' not in xml and "queryTableFieldId" in xml:
                    problemes.append(f"colonne reliee a une requete disparue : {table}")

            # 2c. plage de donnees externes (nom ExternalData_*) sans requete
            classeur = archive.read("xl/workbook.xml").decode("utf-8", "ignore")
            if re.search(r'<definedName[^>]*name="ExternalData_', classeur) and \
                    not any(n.startswith("xl/queryTables/") for n in parties):
                problemes.append("plage de donnees externes orpheline (nom ExternalData_*)")

            # 3. reference externe declaree mais absente
            if "<externalReference" in classeur and not any("externalLink" in n for n in parties):
                problemes.append("reference externe declaree dans workbook.xml mais absente du fichier")

            # 4. partie annoncee dans [Content_Types] mais absente
            types = archive.read("[Content_Types].xml").decode("utf-8", "ignore")
            for annoncee in re.findall(r'PartName="/([^"]+)"', types):
                if annoncee not in parties:
                    problemes.append(f"partie annoncee mais absente : {annoncee}")
    except Exception as exc:                                    # pragma: no cover
        problemes.append(f"fichier illisible ({exc})")
    return problemes


# ---------------------------------------------------------------------------
# Creation des lignes : sous la ligne soeur, ou a la fin si c'est plus sur
# ---------------------------------------------------------------------------
@dataclass
class _Positions:
    """Ou se trouvent les lignes apres les insertions (pour le rapport)."""

    insertions: list[tuple[int, int]] = field(default_factory=list)   # (ligne d'origine, nombre insere)
    nouvelles: dict[str, int] = field(default_factory=dict)            # cle -> ligne finale

    def finale(self, ligne: int) -> int:
        """Ligne du fichier resultat correspondant a une ligne du fichier d'origine."""
        return ligne + sum(nombre for point, nombre in self.insertions if point <= ligne)


def _insertion_impossible(workbook, ws, premiere_ligne: int) -> str | None:
    """Dit si inserer des lignes au milieu de la feuille peut casser quelque chose.

    openpyxl deplace les cellules mais ne met pas a jour les references qui
    pointent vers des numeros de ligne precis. On n'insere donc au milieu que
    si rien ne depend de ces numeros ; sinon on ajoute a la fin (sans risque).
    """
    if list(ws.conditional_formatting):
        return "la feuille a des mises en forme conditionnelles"
    if ws.data_validations.dataValidation:
        return "la feuille a des listes deroulantes (validations)"
    if ws.merged_cells.ranges:
        return "la feuille a des cellules fusionnees"

    for cellule in list(getattr(ws, "_cells", {}).values()):
        if cellule.row >= premiere_ligne and _is_formula(cellule.value):
            return f"la feuille contient des formules (ex. {cellule.coordinate})"

    nom = ws.title
    motif = re.compile(r"(?:'" + re.escape(nom.replace("'", "''")) + r"'|\b" + re.escape(nom) + r")!([$A-Z0-9:]+)")
    for autre in workbook.worksheets:
        for cellule in list(getattr(autre, "_cells", {}).values()):
            valeur = cellule.value
            if not _is_formula(valeur) or nom not in valeur:
                continue
            for plage in motif.findall(valeur):
                if re.search(r"\d", plage):
                    return f"la formule {autre.title}!{cellule.coordinate} vise des lignes precises ({plage})"

    noms = list(getattr(workbook, "defined_names", {}).items())
    for feuille in workbook.worksheets:
        noms += list(getattr(feuille, "defined_names", {}).items())
    for cle, defini in noms:
        if str(cle).startswith("_xlnm._FilterDatabase"):
            continue
        texte = str(getattr(defini, "attr_text", "") or "")
        for plage in motif.findall(texte):
            if re.search(r"\d", plage):
                return f"le nom defini {cle} vise des lignes precises ({plage})"
    return None


def _decaler_hauteurs(ws, depuis: int, nombre: int) -> None:
    """Deplace les hauteurs / lignes masquees, qu'openpyxl ne deplace pas tout seul."""
    a_deplacer = {i: ws.row_dimensions[i] for i in list(ws.row_dimensions.keys()) if i >= depuis}
    for i in a_deplacer:
        del ws.row_dimensions[i]
    for i, dimension in a_deplacer.items():
        nouvelle = copy(dimension)
        nouvelle.index = i + nombre
        ws.row_dimensions[i + nombre] = nouvelle


def _remplir_ligne(ws, layout, creation, ligne: int, modele: int, first_col: int, last_col: int,
                   log: ProcessLogger) -> None:
    """La nouvelle ligne = copie de la ligne modele, puis les valeurs de la Reference par-dessus."""
    _copy_row_style(ws, modele, ligne, first_col, last_col)
    _duplicate_row_values(ws, modele, ligne, first_col, last_col)
    for nom, col in layout.columns.items():
        if nom not in creation.values:
            continue
        if _is_formula(ws.cell(ligne, col).value):
            log.log(INFO, f"Ligne {ligne} : formule conservee dans la colonne {nom}",
                    store=creation.store, division=creation.division)
            continue
        ws.cell(ligne, col).value = _coerce(creation.values.get(nom, ""), nom)


def _ecrire_creations(target: TargetData, plan: Plan, first_col: int, last_col: int,
                      log: ProcessLogger, result: WriteResult) -> _Positions:
    ws, layout = target.worksheet, target.layout
    positions = _Positions()

    sous_soeur = [c for c in plan.creations if c.template_row]
    a_la_fin = [c for c in plan.creations if not c.template_row]

    if sous_soeur:
        premiere = min(c.template_row for c in sous_soeur) + 1
        raison = _insertion_impossible(target.workbook, ws, premiere)
        if raison:
            message = (f"Nouvelles lignes ajoutees a la FIN du tableau (et non sous la ligne du meme Store) : "
                       f"{raison}. Les inserer au milieu risquerait de casser ces references.")
            result.warnings.append(message)
            log.info(message)
            a_la_fin = sous_soeur + a_la_fin
            sous_soeur = []

    # 1. Insertions sous la ligne soeur, du BAS vers le HAUT : chaque insertion
    #    ne decale ainsi que des lignes deja traitees.
    groupes: dict[int, list] = {}
    for creation in sous_soeur:
        groupes.setdefault(creation.template_row + 1, []).append(creation)
    for point in sorted(groupes, reverse=True):
        lot = groupes[point]
        ws.insert_rows(point, amount=len(lot))
        _decaler_hauteurs(ws, point, len(lot))
        for decalage, creation in enumerate(lot):
            _remplir_ligne(ws, layout, creation, point + decalage, point - 1, first_col, last_col, log)
            result.created_rows += 1
        positions.insertions.append((point, len(lot)))

    for point in sorted(groupes):
        for decalage, creation in enumerate(groupes[point]):
            positions.nouvelles[creation.key] = positions.finale(point - 1) + 1 + decalage

    # 2. Ajouts a la fin du tableau (pas de ligne soeur, ou insertion jugee risquee)
    if a_la_fin:
        debut = positions.finale(layout.last_data_row) + 1
        occupe = any(
            ws.cell(r, c).value not in (None, "")
            for r in range(debut, debut + len(a_la_fin))
            for c in range(first_col, last_col + 1)
        )
        if occupe:
            ws.insert_rows(debut, amount=len(a_la_fin))
            _decaler_hauteurs(ws, debut, len(a_la_fin))
            result.warnings.append(
                f"{len(a_la_fin)} ligne(s) inserees en ligne {debut} : le contenu situe en dessous a ete decale.")
        for decalage, creation in enumerate(a_la_fin):
            ligne = debut + decalage
            modele = positions.finale(creation.template_row) if creation.template_row \
                else positions.finale(layout.last_data_row)
            _remplir_ligne(ws, layout, creation, ligne, modele, first_col, last_col, log)
            positions.nouvelles[creation.key] = ligne
            result.created_rows += 1

    for creation in plan.creations:
        ligne = positions.nouvelles.get(creation.key)
        log.log(INFO, f"Ligne {ligne} creee dans le fichier resultat (cle {creation.key})",
                row=ligne, store=creation.store, division=creation.division)
    return positions


def _renumeroter_rapport(plan: Plan, log: ProcessLogger, positions: _Positions) -> None:
    """Apres insertion, le rapport donne les numeros de ligne du FICHIER RESULTAT."""
    for decision in log.decisions:
        if decision.row:
            decision.row = positions.finale(decision.row)
        elif decision.key in positions.nouvelles and decision.decision in ("CREEE", "A VERIFIER"):
            decision.row = positions.nouvelles[decision.key]
    for entree in log.entries:
        if entree.row and not entree.message.startswith("Ligne "):
            entree.row = positions.finale(entree.row)
    for update in plan.updates:
        update.row = positions.finale(update.row)
    for mark in plan.marks:
        mark.row = positions.finale(mark.row)
