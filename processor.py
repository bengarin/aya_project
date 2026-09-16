"""
processor.py
------------
Le CERVEAU du logiciel : toutes les regles metier sont ici.

Tres important : ce fichier ne touche PAS a Excel. Il lit des donnees deja
chargees et il fabrique un "plan de travail" :
   - quelles lignes modifier (et quelles cellules exactement),
   - quelles lignes creer,
   - quelles lignes colorer en rouge.

C'est ce plan qui permet l'apercu (mode Preview) AVANT d'ecrire le fichier.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Callable

from excel_reader import ReferenceData, ReferenceRow, TargetData, TargetRow
from kam_service import AMBIGUOUS, KamService, NOT_FOUND, WRONG_DIVISION
from logger import (
    CHECK, CREATE, DUPLICATE, KAM_NOT_FOUND, MODIF,
    ProcessLogger, RED, STORE_NOT_FOUND,
)
from matcher import KNOWN_DIVISIONS, build_key, norm_division, norm_key, norm_store

# Colonnes du fichier 2 alimentees depuis la Reference (regle 4).
FIELDS_FROM_REFERENCE = ("ID Promoter", "Code Store", "Promoter", "Division", "Division1", "DAY", "Column1")


@dataclass
class Options:
    """Options du traitement. Par defaut = regles metier strictes."""

    update_city: bool = False                   # mettre a jour City depuis la Reference
    process_unknown_divisions: bool = False     # traiter RAC / autres comme une division normale
    auto_number_idaya: bool = False             # numeroter IDAYA sur les lignes creees
    add_missing_reference_stores: bool = False  # ajouter les stores de la Reference absents du fichier 2
    highlight_unknown_stores: bool = False      # surligner en orange les stores absents de la Reference


@dataclass
class CellChange:
    column: str
    old: str
    new: str


@dataclass
class RowUpdate:
    row: int
    store: str
    division: str
    changes: list[CellChange]
    values: dict[str, str]


@dataclass
class RowCreation:
    store: str
    division: str
    values: dict[str, str]
    template_row: int | None
    origin: str = ""                            # d'ou vient la creation (info rapport)


@dataclass
class RowMark:
    row: int
    reason: str
    kind: str = "red"                           # "red" ou "orange"


@dataclass
class Stats:
    analyzed_rows: int = 0
    stores_processed: int = 0
    updated_rows: int = 0
    unchanged_rows: int = 0
    created_rows: int = 0
    red_rows: int = 0
    store_not_found: int = 0
    kam_found: int = 0
    kam_not_found: int = 0
    duplicates_avoided: int = 0
    to_check: int = 0

    def as_pairs(self) -> list[tuple[str, int]]:
        return [
            ("Lignes analysees", self.analyzed_rows),
            ("Stores traites", self.stores_processed),
            ("Lignes modifiees", self.updated_rows),
            ("Lignes deja conformes", self.unchanged_rows),
            ("Lignes ajoutees", self.created_rows),
            ("Lignes rouges / Store vide", self.red_rows),
            ("Stores non trouves", self.store_not_found),
            ("KAM trouves", self.kam_found),
            ("KAM non trouves", self.kam_not_found),
            ("Duplicates evites", self.duplicates_avoided),
            ("A verifier", self.to_check),
        ]


@dataclass
class Plan:
    updates: list[RowUpdate] = field(default_factory=list)
    creations: list[RowCreation] = field(default_factory=list)
    marks: list[RowMark] = field(default_factory=list)
    stats: Stats = field(default_factory=Stats)
    logger: ProcessLogger = field(default_factory=ProcessLogger)

    @property
    def has_work(self) -> bool:
        return bool(self.updates or self.creations or self.marks)


class Processor:
    """Applique les regles metier et produit un Plan."""

    def __init__(
        self,
        reference: ReferenceData,
        target: TargetData,
        kam: KamService,
        options: Options | None = None,
        logger: ProcessLogger | None = None,
    ) -> None:
        self.reference = reference
        self.target = target
        self.kam = kam
        self.options = options or Options()
        self.log = logger or ProcessLogger()
        self.plan = Plan(logger=self.log)

        # Index de travail
        self._key_index: dict[str, list[int]] = {}
        self._rows_by_number: dict[int, TargetRow] = {}
        self._rows_by_store: dict[str, list[TargetRow]] = {}
        self._consumed_rows: set[int] = set()
        self._planned_keys: set[str] = set()
        self._build_indexes()

    # ------------------------------------------------------------------
    # Indexation du fichier 2
    # ------------------------------------------------------------------
    def _build_indexes(self) -> None:
        for trow in self.target.rows:
            self._rows_by_number[trow.row] = trow
            store_key = norm_store(trow.get("Store"))
            if store_key:
                self._rows_by_store.setdefault(store_key, []).append(trow)

            # Cle 1 : valeur de Column1 telle qu'elle existe dans le fichier
            column1 = norm_key(trow.get("Column1"))
            if column1:
                self._key_index.setdefault(column1, []).append(trow.row)
            # Cle 2 : reconstruite depuis les cellules Code Store + Division1
            rebuilt = build_key(trow.get("Code Store"), trow.get("Division1"))
            if rebuilt and rebuilt != column1:
                self._key_index.setdefault(rebuilt, []).append(trow.row)

        # Valeurs les plus frequentes : servent de valeur par defaut pour les lignes creees
        self._dominant: dict[str, str] = {}
        for name in ("Annee", "Mois", "STATUT"):
            values = [r.get(name) for r in self.target.rows if r.get(name)]
            if values:
                self._dominant[name] = Counter(values).most_common(1)[0][0]

        idaya_numbers = []
        for r in self.target.rows:
            raw = r.get("IDAYA").replace(" ", "")
            if raw.isdigit():
                idaya_numbers.append(int(raw))
        self._next_idaya = (max(idaya_numbers) + 1) if idaya_numbers else 1

    # ------------------------------------------------------------------
    # Traitement principal
    # ------------------------------------------------------------------
    def run(self, progress: Callable[[int, int], None] | None = None) -> Plan:
        rows = self.target.rows
        total = max(len(rows), 1)
        processed_stores: set[str] = set()

        for index, trow in enumerate(rows, start=1):
            self.plan.stats.analyzed_rows += 1

            # ---- ETAPE 1 : Store vide -> ligne rouge, rien d'autre (regle 2)
            if not trow.get("Store").strip():
                self.plan.marks.append(RowMark(row=trow.row, reason="Store vide", kind="red"))
                self.plan.stats.red_rows += 1
                self.log.log(RED, "Store vide -> ligne coloree en rouge, aucune autre modification", row=trow.row)
                self._consumed_rows.add(trow.row)
                if progress:
                    progress(index, total)
                continue

            store_key = norm_store(trow.get("Store"))
            if store_key not in processed_stores:
                processed_stores.add(store_key)
                self._process_store(store_key, trow.get("Store"), trow.row)
            if progress:
                progress(index, total)

        # ---- Option : ajouter les magasins de la Reference absents du fichier 2
        if self.options.add_missing_reference_stores:
            self._process_missing_reference_stores()

        self._report_orphan_rows()
        self.plan.stats.stores_processed = len(processed_stores)
        return self.plan

    # ------------------------------------------------------------------
    def _process_store(self, store_key: str, store_label: str, first_row: int) -> None:
        """Traite toutes les affectations (divisions) d'un magasin."""
        entries = self.reference.by_store.get(store_key, [])
        if not entries:
            # Regle 16 : on signale, on ne modifie rien.
            rows = self._rows_by_store.get(store_key, [])
            for trow in rows:
                self.plan.stats.store_not_found += 1
                self.log.log(
                    STORE_NOT_FOUND,
                    "Store absent du fichier Reference -> ligne non modifiee",
                    row=trow.row, store=store_label,
                )
                self._consumed_rows.add(trow.row)
                if self.options.highlight_unknown_stores:
                    self.plan.marks.append(RowMark(row=trow.row, reason="Store absent de la Reference", kind="orange"))
            return

        # Regroupement par division : 'VD+DA' compte pour VD et pour DA (regle 6)
        by_division: dict[str, list[ReferenceRow]] = {}
        for entry in entries:
            if not entry.divisions:
                self.plan.stats.to_check += 1
                self.log.check(
                    f"Colonne 'Divisions' vide dans la Reference (ligne {entry.source_row}) -> affectation ignoree",
                    row=first_row, store=store_label,
                )
                continue
            for division in entry.divisions:
                by_division.setdefault(division, []).append(entry)

        for division in sorted(by_division):
            candidates = by_division[division]

            # Division non prevue par les regles (ex: 'RAC')
            if division not in KNOWN_DIVISIONS and not self.options.process_unknown_divisions:
                self.plan.stats.to_check += 1
                self.log.check(
                    f"Division '{division}' non prevue par les regles metier "
                    f"(Reference ligne(s) {', '.join(str(e.source_row) for e in candidates)}) -> non traitee",
                    row=first_row, store=store_label, division=division,
                )
                for ref_row in candidates:
                    self._skip_rows(build_key(ref_row.store_code, division), store_key, division)
                continue

            # Plusieurs lignes Reference pour la MEME division : on ne choisit pas au hasard (regle 20)
            unique = {
                (e.id_promoter, e.store_code, e.promoter, e.working_days, e.divisions_raw): e
                for e in candidates
            }
            if len(unique) > 1:
                self.plan.stats.to_check += 1
                detail = " | ".join(
                    f"ligne {e.source_row}: {e.id_promoter} {e.promoter} ({e.divisions_raw}, {e.working_days}j)"
                    for e in candidates
                )
                self.log.check(
                    f"{len(candidates)} promoteurs differents pour {store_label} / {division} dans la Reference "
                    f"-> aucune modification automatique. Details : {detail}",
                    row=first_row, store=store_label, division=division,
                )
                for ref_row in candidates:
                    self._skip_rows(build_key(ref_row.store_code, division), store_key, division)
                continue

            entry = next(iter(unique.values()))
            self._apply_affectation(entry, division, store_label)

    # ------------------------------------------------------------------
    def _apply_affectation(self, entry: ReferenceRow, division: str, store_label: str) -> None:
        """Modifie la ligne existante, ou cree la ligne manquante, pour une cle donnee."""
        key = build_key(entry.store_code, division)
        target_row = self._find_row_for_key(key, entry, division)

        values = self._values_from_reference(entry, division, key)
        kam_result = self.kam.lookup(store_label or entry.store, division)
        row_number = target_row.row if target_row else None

        if kam_result.ok:
            values["KAM"] = kam_result.kam
            self.plan.stats.kam_found += 1
        else:
            self.plan.stats.kam_not_found += 1
            level = KAM_NOT_FOUND
            message = {
                NOT_FOUND: "KAM introuvable dans le fichier 3",
                WRONG_DIVISION: "KAM introuvable pour cette division",
                AMBIGUOUS: "KAM ambigu dans le fichier 3",
            }.get(kam_result.status, "KAM introuvable")
            if kam_result.status == AMBIGUOUS:
                level = CHECK
                self.plan.stats.to_check += 1
            self.log.log(
                level, f"{message} ({kam_result.detail}) -> valeur KAM existante conservee",
                row=row_number, store=store_label, division=division,
            )

        if target_row is not None:
            self._plan_update(target_row, values, store_label, division, key)
        else:
            self._plan_creation(entry, division, values, store_label, key)

    # ------------------------------------------------------------------
    def _candidate_rows(self, key: str, store_key: str, division: str) -> list[TargetRow]:
        """Toutes les lignes du fichier 2 qui correspondent a une cle, non deja utilisees."""
        found = [self._rows_by_number[r] for r in self._key_index.get(key, []) if r not in self._consumed_rows]
        if found:
            return found
        return [
            trow for trow in self._rows_by_store.get(store_key, [])
            if trow.row not in self._consumed_rows and norm_division(trow.get("Division1")) == division
        ]

    def _skip_rows(self, key: str, store_key: str, division: str) -> None:
        """Marque les lignes concernees comme 'vues' (deja signalees dans le rapport).

        Evite qu'une situation deja signalee (conflit, division inconnue) soit
        re-signalee une deuxieme fois comme 'affectation absente de la Reference'.
        """
        for trow in self._candidate_rows(key, store_key, division):
            self._consumed_rows.add(trow.row)

    def _find_row_for_key(self, key: str, entry: ReferenceRow, division: str) -> TargetRow | None:
        """Cherche la ligne existante correspondant a 'Code Store + Division' (regle 8)."""
        rows = [r for r in self._key_index.get(key, []) if r not in self._consumed_rows]
        if rows:
            if len(rows) > 1:
                for extra in rows[1:]:
                    self.plan.stats.to_check += 1
                    self.log.check(
                        f"Plusieurs lignes portent deja la cle {key} dans le fichier 2 "
                        "-> seule la premiere est mise a jour, les autres sont a verifier",
                        row=extra, store=entry.store, division=division,
                    )
            chosen = self._rows_by_number[rows[0]]
            self._consumed_rows.add(chosen.row)
            return chosen

        # Repli : meme magasin + meme Division1, mais Code Store / Column1 incoherents dans le fichier 2
        for trow in self._rows_by_store.get(norm_store(entry.store), []):
            if trow.row in self._consumed_rows:
                continue
            if norm_division(trow.get("Division1")) == division:
                self.log.info(
                    f"Ligne rapprochee par Store + Division1 (Code Store/Column1 incoherents : "
                    f"'{trow.get('Code Store')}' / '{trow.get('Column1')}' -> cle attendue {key})",
                    row=trow.row, store=entry.store, division=division,
                )
                self._consumed_rows.add(trow.row)
                return trow
        return None

    # ------------------------------------------------------------------
    def _values_from_reference(self, entry: ReferenceRow, division: str, key: str) -> dict[str, str]:
        values = {
            "ID Promoter": entry.id_promoter,
            "Code Store": entry.store_code,
            "Promoter": entry.promoter,
            "Division": entry.divisions_raw,
            "Division1": division,
            "DAY": entry.working_days,
            "Column1": key,
        }
        if self.options.update_city and entry.city:
            values["City"] = entry.city
        return values

    # ------------------------------------------------------------------
    def _plan_update(self, trow: TargetRow, values: dict[str, str], store_label: str, division: str, key: str) -> None:
        # La cle existait deja : on modifie la ligne, on ne cree pas de doublon (regles 8 et 12).
        self.plan.stats.duplicates_avoided += 1
        self._planned_keys.add(key)

        changes = [
            CellChange(column=name, old=trow.get(name), new=new)
            for name, new in values.items()
            if name in self.target.layout.columns and trow.get(name) != new
        ]
        if not changes:
            self.plan.stats.unchanged_rows += 1
            return
        self.plan.updates.append(
            RowUpdate(row=trow.row, store=store_label, division=division, changes=changes, values=values)
        )
        self.plan.stats.updated_rows += 1
        detail = "; ".join(f"{c.column}: '{c.old}' -> '{c.new}'" for c in changes)
        self.log.log(MODIF,
                     f"Cle {key} deja presente -> ligne existante mise a jour (duplicate evite). "
                     f"{len(changes)} cellule(s) : {detail}",
                     row=trow.row, store=store_label, division=division)

    # ------------------------------------------------------------------
    def _plan_creation(
        self, entry: ReferenceRow, division: str, values: dict[str, str],
        store_label: str, key: str, origin: str = "",
    ) -> None:
        if key in self._planned_keys:
            self.plan.stats.duplicates_avoided += 1
            self.log.log(DUPLICATE, f"Cle {key} deja traitee -> creation annulee",
                         store=store_label, division=division)
            return

        sibling = self._sibling_row(entry)
        full = dict(values)
        full["Store"] = (sibling.get("Store") if sibling else "") or entry.store or store_label

        # Colonnes que la Reference ne fournit pas : on recopie la ligne soeur du meme magasin.
        for name, fallback in (("Annee", None), ("Mois", None), ("STATUT", None)):
            if name not in self.target.layout.columns:
                continue
            value = sibling.get(name) if sibling else ""
            if not value:
                value = self._dominant.get(name, "")
                if value:
                    self.log.check(
                        f"Nouvelle ligne {key} : '{name}' non fourni par la Reference "
                        f"-> valeur la plus frequente du fichier 2 utilisee ('{value}') - a verifier",
                        store=store_label, division=division,
                    )
                    self.plan.stats.to_check += 1
            full[name] = value

        if "City" in self.target.layout.columns and not full.get("City"):
            full["City"] = (sibling.get("City") if sibling else "") or entry.city

        if "IDAYA" in self.target.layout.columns:
            if self.options.auto_number_idaya:
                full["IDAYA"] = str(self._next_idaya)
                self._next_idaya += 1
            else:
                full["IDAYA"] = ""
                self.plan.stats.to_check += 1
                self.log.check(
                    f"Nouvelle ligne {key} : colonne IDAYA laissee vide (aucune regle definie) - a completer",
                    store=store_label, division=division,
                )

        self.plan.creations.append(
            RowCreation(
                store=full["Store"], division=division, values=full,
                template_row=sibling.row if sibling else None, origin=origin,
            )
        )
        self._planned_keys.add(key)
        self.plan.stats.created_rows += 1
        self.log.log(
            CREATE,
            f"Nouvelle ligne creee pour la cle {key} "
            f"({full.get('ID Promoter', '')} - {full.get('Promoter', '')}, {full.get('DAY', '')} jours)"
            + (f" [{origin}]" if origin else ""),
            row=sibling.row if sibling else None, store=full["Store"], division=division,
        )

    # ------------------------------------------------------------------
    def _sibling_row(self, entry: ReferenceRow) -> TargetRow | None:
        """Ligne deja presente pour le meme magasin (sert de modele de style et de valeurs)."""
        rows = self._rows_by_store.get(norm_store(entry.store), [])
        return rows[0] if rows else None

    # ------------------------------------------------------------------
    def _process_missing_reference_stores(self) -> None:
        """Option : cree les affectations de la Reference dont le magasin est totalement absent."""
        for store_key, entries in self.reference.by_store.items():
            if not store_key or store_key in self._rows_by_store:
                continue
            for entry in entries:
                for division in entry.divisions:
                    if division not in KNOWN_DIVISIONS and not self.options.process_unknown_divisions:
                        continue
                    key = build_key(entry.store_code, division)
                    values = self._values_from_reference(entry, division, key)
                    kam_result = self.kam.lookup(entry.store, division)
                    if kam_result.ok:
                        values["KAM"] = kam_result.kam
                        self.plan.stats.kam_found += 1
                    else:
                        self.plan.stats.kam_not_found += 1
                        self.log.log(KAM_NOT_FOUND,
                                     f"KAM introuvable ({kam_result.detail})",
                                     store=entry.store, division=division)
                    self._plan_creation(entry, division, values, entry.store, key,
                                        origin="option: store present dans la Reference, absent du fichier 2")

    # ------------------------------------------------------------------
    def _report_orphan_rows(self) -> None:
        """Signale les lignes du fichier 2 qui ne correspondent a aucune affectation de la Reference."""
        for trow in self.target.rows:
            if trow.row in self._consumed_rows:
                continue
            self.plan.stats.to_check += 1
            self.log.check(
                f"Affectation '{trow.get('Store')} / {trow.get('Division1')}' absente de la Reference "
                "-> ligne conservee telle quelle",
                row=trow.row, store=trow.get("Store"), division=trow.get("Division1"),
            )
