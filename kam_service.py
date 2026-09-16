"""
kam_service.py
--------------
Service de recherche du KAM (fichier 3).

Explication simple :
le fichier 3 est un annuaire. On lui donne un magasin (et sa division), il rend
le nom du KAM responsable. Le fichier fourni contient DEUX feuilles : 'VD' et 'DA'.
Pour un meme magasin, le KAM VD et le KAM DA sont differents. Le service cherche
donc d'abord dans la feuille de la bonne division.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from excel_reader import KamData, KamEntry
from matcher import norm_division, norm_store

# Resultats possibles d'une recherche
FOUND = "FOUND"                    # KAM trouve dans la bonne feuille / feuille generique
NOT_FOUND = "NOT_FOUND"            # magasin absent du fichier 3
WRONG_DIVISION = "WRONG_DIVISION"  # magasin present, mais pas pour cette division
AMBIGUOUS = "AMBIGUOUS"            # plusieurs KAM differents pour le meme magasin+division


@dataclass
class KamResult:
    kam: str
    status: str
    detail: str = ""

    @property
    def ok(self) -> bool:
        return self.status == FOUND


@dataclass
class KamService:
    data: KamData
    _by_division: dict[str, dict[str, list[KamEntry]]] = field(default_factory=dict, init=False)
    _generic: dict[str, list[KamEntry]] = field(default_factory=dict, init=False)

    def __post_init__(self) -> None:
        for entry in self.data.entries:
            key = norm_store(entry.store)
            if entry.division:
                self._by_division.setdefault(entry.division, {}).setdefault(key, []).append(entry)
            else:
                self._generic.setdefault(key, []).append(entry)

    # -- informations pour l'interface --------------------------------------
    @property
    def divisions(self) -> list[str]:
        return sorted(self._by_division)

    @property
    def has_generic_sheet(self) -> bool:
        return bool(self._generic)

    def summary(self) -> str:
        parts = [f"{d}: {len(v)} magasins" for d, v in sorted(self._by_division.items())]
        if self._generic:
            parts.append(f"generique: {len(self._generic)} magasins")
        return " | ".join(parts) if parts else "aucune donnee"

    # -- recherche ----------------------------------------------------------
    def lookup(self, store: str, division: str) -> KamResult:
        """Cherche le KAM d'un magasin pour une division donnee."""
        key = norm_store(store)
        div = norm_division(division)
        if not key:
            return KamResult("", NOT_FOUND, "Store vide")

        # 1. feuille de la division demandee
        entries = self._by_division.get(div, {}).get(key, [])
        source = f"feuille {div}"
        # 2. sinon feuille generique (fichier a une seule feuille sans division)
        if not entries and self._generic.get(key):
            entries = self._generic[key]
            source = "feuille generique"

        if entries:
            names = {e.kam for e in entries}
            if len(names) > 1:
                return KamResult(
                    "", AMBIGUOUS,
                    f"{len(names)} KAM differents pour ce magasin ({source}) : " + ", ".join(sorted(names)),
                )
            return KamResult(entries[0].kam, FOUND, source)

        # 3. present ailleurs ? on ne recopie PAS le KAM d'une autre division.
        others = [d for d, table in self._by_division.items() if key in table]
        if others:
            return KamResult(
                "", WRONG_DIVISION,
                f"Magasin present uniquement dans la/les feuille(s) {', '.join(sorted(others))} "
                f"- aucun KAM {div} defini (valeur existante conservee).",
            )
        return KamResult("", NOT_FOUND, "Magasin absent du fichier 3")
