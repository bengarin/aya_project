"""
processor.py
------------
LES REGLES METIER, et rien d'autre.

Ce fichier ne touche jamais a Excel : il lit les donnees deja chargees et
fabrique un PLAN (quoi modifier, quoi creer, quoi colorer en rouge) ainsi que
la decision + la raison pour CHAQUE ligne du fichier 2.

Ordre de traitement impose (regle 17) :
  1. Store vide            -> rouge uniquement, STOP
  2. Store introuvable     -> signaler, STOP
  3. Analyser Divisions de la Reference (VD / DA / VD+DA ; RAC = a verifier)
  4. Pour chaque division : cle = Code Store + Division
  5. Cle presente -> UPDATE ; cle absente -> CREATE
  6. KAM cherche avec Store + Division
  7. Column1 recalculee = Code Store + Division1
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from excel_reader import ReferenceData, ReferenceRow, TargetData, TargetRow
from kam_service import AMBIGUOUS, KamService, NOT_FOUND, WRONG_DIVISION
from logger import (
    A_VERIFIER, CONFORME, CREEE, DUPLICATE, IGNOREE, KAM_NON_TROUVE,
    MODIFIEE, ProcessLogger, ROUGE, STORE_NON_TROUVE,
)
from matcher import build_key, norm_division, norm_store, split_divisions

# Les seules divisions traitees automatiquement (regle 3).
DIVISIONS_TRAITEES = ("VD", "DA")

# Marqueur ecrit dans KAM quand aucun KAM n'est trouve dans le fichier 3
# (decision explicite de l'utilisateur : ni invente, ni laisse a l'ancienne valeur).
KAM_PLACEHOLDER = "#"

# Colonnes du fichier 2 alimentees depuis la Reference (regles 4 et 5).
#   Division (C), Division1 (D), KAM (E), ID Promoter (F), Code Store (G),
#   Promoter (H), Store (I), City (J), DAY (L), Column1 (N)
COLONNES_ECRITES = ("Division", "Division1", "KAM", "ID Promoter", "Code Store",
                    "Promoter", "Store", "City", "DAY", "Column1")

# Colonnes jamais modifiees sur une ligne existante : Annee (A), Mois (B),
# STATUT (K), IDAYA (M) -> regle 13 et "ne rien inventer".
COLONNES_NON_TOUCHEES = ("Annee", "Mois", "STATUT", "IDAYA")


@dataclass
class Options:
    """Seule option prevue par les regles (regle 13)."""

    auto_number_idaya: bool = False


@dataclass
class CellChange:
    column: str
    old: str
    new: str

    def as_text(self) -> str:
        return f"{self.column}: '{self.old}' -> '{self.new}'"


@dataclass
class RowUpdate:
    row: int
    store: str
    division: str
    key: str
    changes: list[CellChange]


@dataclass
class RowCreation:
    store: str
    division: str
    key: str
    values: dict[str, str]
    template_row: int | None


@dataclass
class RowMark:
    row: int
    reason: str


@dataclass
class Stats:
    lignes_analysees: int = 0
    lignes_modifiees: int = 0
    lignes_conformes: int = 0
    lignes_ajoutees: int = 0
    lignes_store_vide: int = 0
    stores_non_trouves: int = 0
    kam_trouves: int = 0
    kam_non_trouves: int = 0
    duplicates_evites: int = 0
    a_verifier: int = 0

    def as_pairs(self) -> list[tuple[str, int]]:
        return [
            ("Lignes analysees", self.lignes_analysees),
            ("Lignes modifiees", self.lignes_modifiees),
            ("Lignes deja conformes", self.lignes_conformes),
            ("Lignes ajoutees", self.lignes_ajoutees),
            ("Lignes Store vide", self.lignes_store_vide),
            ("Stores non trouves", self.stores_non_trouves),
            ("KAM trouves", self.kam_trouves),
            ("KAM non trouves", self.kam_non_trouves),
            ("Duplicates evites", self.duplicates_evites),
            ("A verifier", self.a_verifier),
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

        self._rows_by_number: dict[int, TargetRow] = {}
        self._rows_by_store: dict[str, list[TargetRow]] = {}
        self._rows_by_cells: dict[str, list[TargetRow]] = {}     # Code Store (G) + Division1 (D)
        self._rows_by_column1: dict[str, list[TargetRow]] = {}   # valeur ecrite dans Column1 (N)
        self._decided_rows: set[int] = set()     # lignes qui ont deja leur decision
        self._used_rows: set[int] = set()        # lignes deja rattachees a une cle
        self._keys_traitees: set[str] = set()    # cles deja traitees (update ou create)
        self._usage_cles: dict[str, list[str]] = {}   # cle -> Stores qui l'utilisent
        self._build_indexes()

    # ------------------------------------------------------------------
    # Index du fichier 2
    # ------------------------------------------------------------------
    def _build_indexes(self) -> None:
        """Indexe les lignes du fichier 2 par Store et par cle Code Store + Division.

        La cle d'une ligne existante est reconnue de 2 facons (regles 8 et 12),
        et dans CET ordre de confiance :
          1. Code Store (G) + Division1 (D)  <- la vraie cle metier,
          2. la valeur ecrite dans Column1 (N) <- peut etre ancienne/fausse.
        """
        for trow in self.target.rows:
            self._rows_by_number[trow.row] = trow
            store_key = norm_store(trow.get("Store"))
            if store_key:
                self._rows_by_store.setdefault(store_key, []).append(trow)
            cells_key = build_key(trow.get("Code Store"), trow.get("Division1"))
            if cells_key:
                self._rows_by_cells.setdefault(cells_key, []).append(trow)
            column1 = build_key(trow.get("Column1"), "")
            if column1 and column1 != cells_key:
                self._rows_by_column1.setdefault(column1, []).append(trow)

        numbers = [int(r.get("IDAYA")) for r in self.target.rows if r.get("IDAYA").strip().isdigit()]
        self._next_idaya = (max(numbers) + 1) if numbers else 1

    def _candidates_for_key(self, key: str, store_key: str = "", division: str = "") -> list[TargetRow]:
        """Lignes libres qui correspondent a cette affectation, par ordre de confiance :

        1. **meme Store + meme Division1** : le plus sur, parce que le nom du Store
           est ce qui identifie vraiment le magasin (c'est ainsi qu'on le cherche
           dans la Reference). Indispensable quand deux Stores differents partagent
           le meme Store Code dans la Reference : chacun retrouve SA ligne.
        2. Code Store (G) + Division1 (D) : la cle metier reconstruite.
        3. la valeur ecrite dans Column1 (N), qui peut etre ancienne ou fausse.
        """
        if store_key and division:
            rows = [
                t for t in self._rows_by_store.get(store_key, [])
                if t.row not in self._used_rows and norm_division(t.get("Division1")) == division
            ]
            if rows:
                return rows
        for index in (self._rows_by_cells, self._rows_by_column1):
            rows = [t for t in index.get(key, []) if t.row not in self._used_rows]
            if rows:
                return rows
        return []

    # ------------------------------------------------------------------
    # Boucle principale : une ligne du fichier 2 a la fois (regle 17)
    # ------------------------------------------------------------------
    def run(self, progress: Callable[[int, int], None] | None = None) -> Plan:
        rows = self.target.rows
        total = max(len(rows), 1)
        stores_traites: set[str] = set()

        for index, trow in enumerate(rows, start=1):
            self.plan.stats.lignes_analysees += 1

            # STEP 1 : Store vide -> ROUGE uniquement, aucun traitement (regle 2)
            if not trow.get("Store").strip():
                self.plan.marks.append(RowMark(row=trow.row, reason="Store vide"))
                self.plan.stats.lignes_store_vide += 1
                self._decide(
                    ROUGE, trow.row,
                    reason="Colonne Store (I) vide : aucune recherche, aucune modification, aucune creation.",
                    details="Seule action autorisee : coloration de toute la ligne en rouge.",
                )
                if progress:
                    progress(index, total)
                continue

            store_key = norm_store(trow.get("Store"))
            if store_key not in stores_traites:
                stores_traites.add(store_key)
                self._traiter_store(store_key, trow.get("Store").strip())
            if progress:
                progress(index, total)

        self._decider_lignes_restantes()
        self._verifier_cles_en_double()
        return self.plan

    # ------------------------------------------------------------------
    def _traiter_store(self, store_key: str, store_label: str) -> None:
        lignes_du_store = self._rows_by_store.get(store_key, [])

        # STEP 2 : Store introuvable dans la Reference (regle 16)
        refs = self.reference.by_store.get(store_key, [])
        if not refs:
            self.plan.stats.stores_non_trouves += 1
            for trow in lignes_du_store:
                self._decide(
                    STORE_NON_TROUVE, trow.row, store=store_label,
                    division=trow.get("Division1"),
                    reason=f"Store '{store_label}' absent du fichier Reference : "
                           "aucune donnee inventee, aucune recherche approchante, ligne laissee intacte.",
                )
                self._used_rows.add(trow.row)
            return

        # STEP 3 et 4 : eclatement des Divisions de la Reference
        par_division: dict[str, list[ReferenceRow]] = {}
        for entry in refs:
            divisions = split_divisions(entry.divisions_raw)
            if not divisions:
                self._verifier(
                    store_label, "", "",
                    f"Colonne Divisions vide dans la Reference (ligne {entry.source_row}) : "
                    "affectation non traitee.",
                )
                continue
            for division in divisions:
                par_division.setdefault(division, []).append(entry)

        for division in sorted(par_division):
            entries = par_division[division]

            # Regle 3 : RAC (ou toute autre valeur) n'est jamais traite comme VD ou DA
            if division not in DIVISIONS_TRAITEES:
                lignes = self._lignes_candidates(entries, division, store_key)
                self._verifier(
                    store_label, division,
                    build_key(entries[0].store_code, division),
                    f"Division '{division}' non prevue par les regles (seules VD et DA sont traitees) : "
                    "aucune donnee ecrite, aucune transformation en VD ou DA.",
                    rows=[t.row for t in lignes],
                    details="Reference ligne(s) " + ", ".join(str(e.source_row) for e in entries),
                )
                continue

            # Regle 7 : meme Store + meme Division avec plusieurs promoteurs -> A VERIFIER
            distinctes = {
                (e.id_promoter, e.store_code, e.promoter, e.working_days, e.city, e.divisions_raw): e
                for e in entries
            }
            if len(distinctes) > 1:
                lignes = self._lignes_candidates(entries, division, store_key)
                detail = " | ".join(
                    f"Reference ligne {e.source_row}: {e.id_promoter} / {e.promoter} / "
                    f"{e.working_days} jours"
                    for e in entries
                )
                self._verifier(
                    store_label, division, build_key(entries[0].store_code, division),
                    f"{len(entries)} promoteurs differents pour {store_label} + {division} dans la Reference : "
                    "aucun choix automatique, aucune ligne ecrasee, donnees laissees intactes.",
                    rows=[t.row for t in lignes],
                    details=detail,
                )
                continue

            self._traiter_affectation(next(iter(distinctes.values())), division, store_label, store_key)

    # ------------------------------------------------------------------
    def _traiter_affectation(self, entry: ReferenceRow, division: str, store_label: str,
                             store_key: str = "") -> None:
        """STEP 5 a 8 pour une combinaison Code Store + Division."""
        key = build_key(entry.store_code, division)
        ligne = self._trouver_ligne(key, store_key or norm_store(entry.store), division)

        # Regles 8 et 15 : un vrai duplicate evite = une CREATION annulee parce que
        # la combinaison Code Store + Division a deja ete traitee dans ce passage.
        if ligne is None and key in self._keys_traitees:
            self.plan.stats.duplicates_evites += 1
            self.log.log(DUPLICATE, f"Cle {key} deja traitee dans ce passage -> creation annulee",
                         store=store_label, division=division)
            return

        # STEP 7 : KAM cherche avec Store + Division (regle 11)
        kam_result = self.kam.lookup(store_label or entry.store, division)
        if kam_result.ok:
            kam_value = kam_result.kam
            self.plan.stats.kam_trouves += 1
            kam_reason = f"KAM '{kam_value}' trouve dans la feuille {division} du fichier 3."
        else:
            # KAM introuvable : on ne devine rien. Decision explicite de l'utilisateur :
            # on ecrit le marqueur "#" (placeholder deja utilise dans son fichier),
            # que la ligne soit existante ou creee.
            kam_value = KAM_PLACEHOLDER
            self.plan.stats.kam_non_trouves += 1
            motif = {
                NOT_FOUND: "Store absent du fichier 3",
                WRONG_DIVISION: f"Store absent de la feuille {division} du fichier 3",
                AMBIGUOUS: "plusieurs KAM differents pour ce Store dans le fichier 3",
            }.get(kam_result.status, "KAM introuvable")
            action = f"KAM ecrit comme '{KAM_PLACEHOLDER}'"
            kam_reason = f"KAM non trouve ({motif}) : {action}."
            self.log.log(KAM_NON_TROUVE, f"{motif} -> {action}",
                         row=ligne.row if ligne else None, store=store_label, division=division)

        valeurs = self._valeurs_reference(entry, division, key, kam_value)

        if ligne is not None:
            self._planifier_update(ligne, valeurs, entry, division, store_label, key, kam_reason)
        else:
            self._planifier_creation(entry, division, store_label, key, valeurs, kam_reason)
        self._keys_traitees.add(key)

    # ------------------------------------------------------------------
    def _trouver_ligne(self, key: str, store_key: str = "", division: str = "") -> TargetRow | None:
        """STEP 5 : la combinaison Code Store + Division existe-t-elle deja ? (regles 8, 9, 10)"""
        candidates = self._candidates_for_key(key, store_key, division)
        if not candidates:
            return None
        choisie = candidates[0]
        self._used_rows.add(choisie.row)
        for autre in candidates[1:]:
            self._verifier(
                autre.get("Store"), norm_division(autre.get("Division1")), key,
                f"La cle {key} apparait sur plusieurs lignes du fichier 2 : "
                f"seule la ligne {choisie.row} est mise a jour, celle-ci est laissee intacte.",
                rows=[autre.row],
            )
            self._used_rows.add(autre.row)
        return choisie

    # ------------------------------------------------------------------
    def _valeurs_reference(self, entry: ReferenceRow, division: str, key: str, kam: str) -> dict[str, str]:
        """Regles 4, 5 et 12 : ce que la Reference impose dans le fichier 2."""
        return {
            "Division": entry.divisions_raw,     # valeur brute de la Reference (VD, DA ou VD+DA)
            "Division1": division,               # division reellement traitee
            "KAM": kam,
            "ID Promoter": entry.id_promoter,
            "Code Store": entry.store_code,
            "Promoter": entry.promoter,
            "Store": entry.store,
            "City": entry.city,
            "DAY": entry.working_days,
            "Column1": key,                      # Code Store + Division1
        }

    # ------------------------------------------------------------------
    def _planifier_update(self, trow: TargetRow, valeurs: dict[str, str], entry: ReferenceRow,
                          division: str, store_label: str, key: str, kam_reason: str) -> None:
        self._usage_cles.setdefault(key, []).append(store_label or entry.store)
        changes: list[CellChange] = []
        formules: list[str] = []
        for name, valeur in valeurs.items():
            if name not in self.target.layout.columns or trow.get(name) == valeur:
                continue
            if trow.is_formula(name):
                # La cellule contient une FORMULE : on n'y touche jamais, elle se
                # recalcule toute seule (demande explicite de l'utilisateur).
                formules.append(name)
                continue
            changes.append(CellChange(column=name, old=trow.get(name), new=valeur))
        source = f"Reference ligne {entry.source_row} (Divisions={entry.divisions_raw})"
        if formules:
            source += f" | formule conservee : {', '.join(formules)}"
        if not changes:
            self.plan.stats.lignes_conformes += 1
            self._decide(
                CONFORME, trow.row, store=store_label, division=division, key=key,
                reason=f"Cle {key} presente dans le fichier 2 et deja identique a la Reference : "
                       "aucune cellule modifiee.",
                details=f"{source}. {kam_reason}",
            )
            return
        self.plan.updates.append(
            RowUpdate(row=trow.row, store=store_label, division=division, key=key, changes=changes)
        )
        self.plan.stats.lignes_modifiees += 1
        self._decide(
            MODIFIEE, trow.row, store=store_label, division=division, key=key,
            reason=f"Cle {key} deja presente dans le fichier 2 : la ligne existante est mise a jour "
                   f"({len(changes)} cellule(s)), aucune ligne creee.",
            details=f"{source}. {kam_reason} Cellules : " + " ; ".join(c.as_text() for c in changes),
        )

    # ------------------------------------------------------------------
    def _planifier_creation(self, entry: ReferenceRow, division: str, store_label: str,
                            key: str, valeurs: dict[str, str], kam_reason: str) -> None:
        soeur = self._ligne_soeur(entry)
        complet = dict(valeurs)
        notes = []

        # Colonnes que la Reference ne fournit pas : recopiees depuis la ligne
        # du MEME Code Store (autre division) si elle existe, sinon laissees vides.
        for name in ("Annee", "Mois", "STATUT"):
            if name not in self.target.layout.columns:
                continue
            valeur = soeur.get(name) if soeur else ""
            complet[name] = valeur
            if not valeur:
                notes.append(f"{name} laisse vide (absent de la Reference et aucune ligne du meme Code Store)")

        if "IDAYA" in self.target.layout.columns:
            if self.options.auto_number_idaya:
                complet["IDAYA"] = str(self._next_idaya)
                notes.append(f"IDAYA={self._next_idaya} (numerotation automatique des lignes creees)")
                self._next_idaya += 1
            else:
                complet["IDAYA"] = ""
                notes.append("IDAYA laisse vide (option de numerotation desactivee)")

        self._usage_cles.setdefault(key, []).append(complet.get("Store", store_label))
        self.plan.creations.append(
            RowCreation(store=complet.get("Store", store_label), division=division,
                        key=key, values=complet, template_row=soeur.row if soeur else None)
        )
        self.plan.stats.lignes_ajoutees += 1
        if notes:
            self._verifier(
                complet.get("Store", store_label), division, key,
                "Ligne creee avec des colonnes non fournies par la Reference : " + " ; ".join(notes),
            )
        origine = (f"Annee/Mois/STATUT recopies depuis la ligne {soeur.row} (meme Code Store)"
                   if soeur else "aucune ligne du meme Code Store pour recopier Annee/Mois/STATUT")
        self._decide(
            CREEE, None, store=complet.get("Store", store_label), division=division, key=key,
            reason=f"Cle {key} absente du fichier 2 alors que la Reference (ligne {entry.source_row}, "
                   f"Divisions={entry.divisions_raw}) l'exige : nouvelle ligne creee.",
            details=f"{kam_reason} {origine}. Valeurs : ID Promoter={complet.get('ID Promoter')}, "
                    f"Promoter={complet.get('Promoter')}, DAY={complet.get('DAY')}, "
                    f"City={complet.get('City')}.",
        )

    # ------------------------------------------------------------------
    def _ligne_soeur(self, entry: ReferenceRow) -> TargetRow | None:
        """Ligne du fichier 2 portant le MEME Code Store (autre division).

        Elle sert uniquement a recopier Annee / Mois / STATUT (que la Reference
        ne contient pas) et a reprendre la mise en forme. Rien d'autre.
        """
        code = build_key(entry.store_code, "")
        for trow in self.target.rows:
            if code and build_key(trow.get("Code Store"), "") == code:
                return trow
        rows = self._rows_by_store.get(norm_store(entry.store), [])
        return rows[0] if rows else None

    # ------------------------------------------------------------------
    def _lignes_candidates(self, entries: list[ReferenceRow], division: str, store_key: str) -> list[TargetRow]:
        """Lignes du fichier 2 concernees par une division non traitee (RAC, conflit)."""
        trouvees: list[TargetRow] = []
        for entry in entries:
            for trow in self._candidates_for_key(build_key(entry.store_code, division), store_key, division):
                if trow not in trouvees:
                    trouvees.append(trow)
        if not trouvees:
            for trow in self._rows_by_store.get(store_key, []):
                if trow.row not in self._used_rows and norm_division(trow.get("Division1")) == division:
                    trouvees.append(trow)
        for trow in trouvees:
            self._used_rows.add(trow.row)
        return trouvees

    # ------------------------------------------------------------------
    def _verifier_cles_en_double(self) -> None:
        """Signale les cles Code Store + Division utilisees par PLUSIEURS Stores.

        Cela arrive quand la Reference donne le meme Store Code a deux magasins
        differents : chacun garde bien ses infos, mais les deux lignes finissent
        avec le meme Column1. Le logiciel ne choisit pas : il le signale.
        """
        for key, stores in self._usage_cles.items():
            distincts = sorted({s for s in stores if s})
            if len(distincts) > 1:
                self._verifier(
                    " / ".join(distincts), "", key,
                    f"La cle {key} est utilisee par {len(distincts)} Stores differents "
                    f"({', '.join(distincts)}) : ils ont le meme Store Code dans la Reference. "
                    "Chaque ligne garde ses propres infos, mais Column1 sera identique.",
                )

    # ------------------------------------------------------------------
    def _decider_lignes_restantes(self) -> None:
        """Lignes du fichier 2 sans decision : aucune affectation correspondante."""
        for trow in self.target.rows:
            if trow.row in self._decided_rows:
                continue
            self._decide(
                IGNOREE, trow.row, store=trow.get("Store"), division=trow.get("Division1"),
                key=build_key(trow.get("Code Store"), trow.get("Division1")),
                reason="Aucune affectation correspondante dans la Reference pour ce Code Store + Division : "
                       "ligne laissee intacte (aucune suppression, aucune modification).",
            )

    # ------------------------------------------------------------------
    # Helpers de journalisation
    # ------------------------------------------------------------------
    def _decide(self, decision: str, row: int | None, reason: str, store: str = "",
                division: str = "", key: str = "", details: str = "") -> None:
        if row is not None:
            if row in self._decided_rows:
                return
            self._decided_rows.add(row)
        self.log.decide(decision, reason, row=row, store=store,
                        division=norm_division(division) if division else "", key=key, details=details)

    def _verifier(self, store: str, division: str, key: str, reason: str,
                  rows: list[int] | None = None, details: str = "") -> None:
        """Enregistre un cas 'A VERIFIER' (le logiciel ne decide pas a la place de l'utilisateur)."""
        self.plan.stats.a_verifier += 1
        if rows:
            for row in rows:
                self._decide(A_VERIFIER, row, reason=reason, store=store, division=division,
                             key=key, details=details)
        else:
            self.log.decide(A_VERIFIER, reason, row=None, store=store,
                            division=division, key=key, details=details)
