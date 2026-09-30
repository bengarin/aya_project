"""
id_service.py
-------------
IDAYA (colonne M) lu dans les NOMS de fichiers d'un dossier (ex. "Etat de vente").

Format d'un nom de fichier :   <numero>-<Store> [<Division(s)>].<extension>
    35-Ashiama panoramique DA.jpeg        -> Store Ashiama panoramique, DA seule      -> ID 35
    39-Carrefour Ain Sebaa VD+DA.jpg      -> un fichier pour VD ET DA                 -> 39DA / 39VD
    40-CARREFOUR BENI MELLAL.jpg          -> pas de division = tout le Store          -> 40DA / 40VD
                                             (ou 40 si le Store n'a qu'une division)

Regle de l'ID (constatee dans la BDD d'aout) :
  - le fichier couvre UNE seule division de la ligne  -> numero seul      (35)
  - le fichier couvre PLUSIEURS divisions             -> numero + division (39DA, 39VD)

Recherche du Store :
  1. nom identique (majuscules, accents, espaces, lettres doublees ignores)
     "ELECTROBOUSFIHA MEKNES" = "Electro Boussfiha Meknes"
  2. sinon nom tres proche (>= 85 %, et nettement meilleur que le 2e candidat)
     "Carrefour Berrachid" ~ "Carrefour Berrechid"  -> utilise MAIS signale dans le rapport.
  Jamais de choix au hasard : 2 fichiers differents pour le meme Store + Division -> rien ecrit.
"""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

SEUIL_APPROCHE = 0.85        # ressemblance minimale pour un nom approche
ECART_MINIMUM = 0.05         # avance minimale sur le 2e meilleur Store

FOUND = "FOUND"
APPROX = "APPROX"
NOT_FOUND = "NOT_FOUND"
AMBIGUOUS = "AMBIGUOUS"

_NOM_RE = re.compile(r"^\s*(\d+)\s*[-_ ]\s*(.+?)\s*$")
_DIV_RE = re.compile(r"(?:^|[\s\-_])((?:VD|DA)(?:\s*[+&/,\-]\s*(?:VD|DA))*)$", re.IGNORECASE)


def _squash(text: str) -> str:
    """Nom de Store compare : sans accents, ni espaces/ponctuation, ni lettres doublees."""
    text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode().lower()
    text = re.sub(r"[^a-z0-9]", "", text)
    return re.sub(r"(.)\1+", r"\1", text)


@dataclass
class IdFile:
    number: str
    store: str                       # nom du Store tel qu'ecrit dans le nom de fichier
    divisions: frozenset[str]        # vide = pas de division dans le nom (tout le Store)
    filename: str
    squashed: str = field(init=False)

    def __post_init__(self) -> None:
        self.squashed = _squash(self.store)


@dataclass
class IdResult:
    status: str
    value: str = ""
    reason: str = ""


def parse_filename(filename: str) -> IdFile | None:
    """'39-Carrefour Ain Sebaa VD+DA.jpg' -> IdFile(39, 'Carrefour Ain Sebaa', {VD, DA})."""
    stem = Path(filename).stem
    match = _NOM_RE.match(stem)
    if not match:
        return None
    number, reste = match.group(1), match.group(2)
    divisions: frozenset[str] = frozenset()
    div = _DIV_RE.search(reste)
    if div:
        divisions = frozenset(re.findall(r"VD|DA", div.group(1).upper()))
        reste = reste[:div.start(1)].rstrip(" -_")
    if not _squash(reste):
        return None
    return IdFile(number=str(int(number)), store=reste.strip(), divisions=divisions, filename=filename)


class IdService:
    def __init__(self, files: list[IdFile], ignored: list[str] | None = None) -> None:
        self.files = files
        self.ignored = ignored or []                  # fichiers au nom non reconnu
        self._by_store: dict[str, list[IdFile]] = {}
        for f in files:
            self._by_store.setdefault(f.squashed, []).append(f)
        self.used: set[str] = set()                   # fichiers effectivement utilises

    @classmethod
    def from_folder(cls, folder: str | Path) -> "IdService":
        folder = Path(folder)
        if not folder.is_dir():
            raise FileNotFoundError(f"Dossier des IDs introuvable : {folder}")
        files, ignored = [], []
        for root, _dirs, names in os.walk(folder):
            for name in sorted(names):
                if name.startswith(("~$", ".")):
                    continue
                parsed = parse_filename(name)
                if parsed is None:
                    ignored.append(str(Path(root, name).relative_to(folder)))
                else:
                    files.append(parsed)
        return cls(files, ignored)

    def describe(self) -> str:
        texte = f"{len(self.files)} fichier(s) avec un numero"
        if self.ignored:
            texte += f", {len(self.ignored)} ignore(s) (nom sans 'numero-Store')"
        return texte

    # ------------------------------------------------------------------
    def _store_files(self, noms: list[str]) -> tuple[list[IdFile], str, str]:
        """Fichiers du Store : (fichiers, 'exact'|'approche'|'', nom du fichier retenu)."""
        cles = [_squash(n) for n in noms if _squash(n)]
        for cle in cles:
            if cle in self._by_store:
                return self._by_store[cle], "exact", self._by_store[cle][0].store
        meilleur: tuple[float, str] | None = None
        second = 0.0
        for cle_fichier in self._by_store:
            score = max(SequenceMatcher(None, cle, cle_fichier).ratio() for cle in cles) if cles else 0
            if meilleur is None or score > meilleur[0]:
                second = meilleur[0] if meilleur else 0.0
                meilleur = (score, cle_fichier)
            elif score > second:
                second = score
        if meilleur and meilleur[0] >= SEUIL_APPROCHE and meilleur[0] - second >= ECART_MINIMUM:
            fichiers = self._by_store[meilleur[1]]
            return fichiers, f"approche {meilleur[0]:.0%}", fichiers[0].store
        return [], "", ""

    def lookup(self, noms_store: list[str], division: str, divisions_du_store: set[str]) -> IdResult:
        """ID de la ligne Store + Division.

        noms_store         : noms possibles du Store (Reference puis fichier 2)
        divisions_du_store : divisions VD/DA du Store dans la Reference (pour les fichiers sans division)
        """
        fichiers, mode, nom_fichier = self._store_files(noms_store)
        if not fichiers:
            return IdResult(NOT_FOUND, reason="aucun fichier du dossier des IDs pour ce Store")

        candidats = []
        for f in fichiers:
            couvre = set(f.divisions) if f.divisions else set(divisions_du_store) or {division}
            if division in couvre:
                valeur = f.number + division if len(couvre) > 1 else f.number
                candidats.append((valeur, f))
        if not candidats:
            return IdResult(NOT_FOUND, reason=f"fichier(s) '{nom_fichier}' trouve(s) mais aucun pour la division {division}")

        valeurs = sorted({v for v, _ in candidats})
        noms = ", ".join(f.filename for _, f in candidats)
        if len(valeurs) > 1:
            return IdResult(AMBIGUOUS, reason=f"plusieurs fichiers donnent des IDs differents ({noms}) : ID non ecrit")
        self.used.update(f.filename for _, f in candidats)
        if mode == "exact":
            return IdResult(FOUND, valeurs[0], f"ID {valeurs[0]} lu dans le fichier '{noms}'")
        return IdResult(APPROX, valeurs[0],
                        f"ID {valeurs[0]} lu dans '{noms}' : nom de Store APPROCHE ({mode}), a controler")

    def unused(self) -> list[str]:
        return [f.filename for f in self.files if f.filename not in self.used]
