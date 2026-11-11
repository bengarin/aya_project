"""
id_service.py
-------------
IDAYA (colonne M) lu dans les NOMS de fichiers d'un dossier (ex. "Etat de vente").

Formats de noms acceptes :   <numero><separateur><Store>[<separateur><Division(s)>][<separateur>PartN]
    35-Ashiama panoramique DA.jpeg              -> Store Ashiama panoramique, DA seule  -> ID 35
    39-Carrefour Ain Sebaa VD+DA.jpg            -> un fichier pour VD ET DA             -> 39DA / 39VD
    40-CARREFOUR BENI MELLAL.jpg                -> pas de division = tout le Store      -> 40DA / 40VD
    81_ASWAK-ASSALAM-KENITRA_VD_Part1.jpg       -> separateurs _ et -, "Part1" ignore   -> 81
    22_ASWAK-ASSALAM-OUJDA_VD-DA_Part2.jpg      -> VD-DA = VD+DA                        -> 22DA / 22VD

Regle de l'ID (constatee dans la BDD d'aout) :
  - le fichier couvre UNE seule division de la ligne  -> numero seul      (35)
  - le fichier couvre PLUSIEURS divisions             -> numero + division (39DA, 39VD)
  Plusieurs pages du meme document (Part1, Part2...) = meme ID, aucun conflit.

Recherche du Store, dans cet ordre :
  1. nom identique (majuscules, accents, espaces, tirets, lettres doublees ignores)
     "ELECTRO-BOUSFIHA-DAR-BOUAAZAA" = "Electro Boussfiha Dar Bouazza", "CARREFOURFES" = "CARREFOUR FES"
  2. nom du Store CONTENU dans le nom du fichier, avec des mots en plus
     "ASHIAMA-PANORAMIQUE-BRANDSHOP" contient "Ashiama Panoramique"
     -> utilise MAIS signale dans le rapport (sauf si un autre Store colle mieux).
  3. nom ecrit autrement, mais sans doute possible (signale "a controler") :
     a. mots communs avec UN SEUL Store, dont un mot propre a ce Store
        "CARREFOUR-MARRAKECH-MAZAR" -> Carrefour Al Mazar (MAZAR), "AERIA-MALL-BRANDSHOP" -> Ashiama Aeria mall
     b. nom raccourci : "ESPACE-BADR" est dans "Espace Badr Bourgogne" ET "Espace Badr Bouskoura" ;
        Bouskoura a deja ses propres fichiers -> c'est Bourgogne. Sinon : rien.
  4. nom tres proche (>= 85 %, nettement meilleur que le 2e candidat)
     "Carrefour Berrachid" ~ "Carrefour Berrechid" -> utilise MAIS signale dans le rapport.
  Jamais de choix au hasard : 2 fichiers qui donnent 2 IDs differents -> rien ecrit.
"""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

SEUIL_APPROCHE = 0.85        # ressemblance minimale pour un nom approche
ECART_MINIMUM = 0.05         # avance minimale sur le 2e meilleur Store
MOTS_MINIMUM = 2             # un Store d'un seul mot n'est jamais cherche "contenu" (trop risque)

FOUND = "FOUND"
APPROX = "APPROX"
NOT_FOUND = "NOT_FOUND"
AMBIGUOUS = "AMBIGUOUS"

_NOM_RE = re.compile(r"^\s*(\d+)\s*(?:[-_ .]\s*)?([^\W\d_].*?|[-_ .].+?)\s*$")   # "218ASSWAK..." accepte
_COPIE_RE = re.compile(r"\s*\(\d+\)\s*$")                                   # "nom (1)" (copie Windows)
_PART_RE = re.compile(r"[\s_\-.]*(?:part|partie|page)\s*\d+\s*$", re.IGNORECASE)
_DIV_RE = re.compile(r"(?:^|[\s\-_.])((?:VD|DA)(?:\s*[+&/,\-_]\s*(?:VD|DA))*)$", re.IGNORECASE)


def _ascii(text: str) -> str:
    return unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode().lower()


def _squash(text: str) -> str:
    """Nom de Store compare : sans accents, ni espaces/ponctuation, ni lettres doublees."""
    return re.sub(r"(.)\1+", r"\1", re.sub(r"[^a-z0-9]", "", _ascii(text)))


def _mots(text: str) -> frozenset[str]:
    """Mots du nom, chacun 'ecrase' comme _squash ('BOUAAZAA' = 'Bouazza')."""
    return frozenset(m for m in (_squash(w) for w in re.split(r"[^a-z0-9]+", _ascii(text))) if m)


@dataclass
class IdFile:
    number: str
    store: str                       # nom du Store tel qu'ecrit dans le nom de fichier
    divisions: frozenset[str]        # vide = pas de division dans le nom (tout le Store)
    filename: str
    squashed: str = field(init=False)
    mots: frozenset[str] = field(init=False)

    def __post_init__(self) -> None:
        self.squashed = _squash(self.store)
        self.mots = _mots(self.store)

    def describe(self) -> str:
        div = "+".join(sorted(self.divisions)) if self.divisions else "toutes divisions"
        return f"'{self.filename}' lu comme : ID {self.number}, Store '{self.store}', {div}"


@dataclass
class IdResult:
    status: str
    value: str = ""
    reason: str = ""


def parse_filename(filename: str) -> IdFile | None:
    """'81_ASWAK-ASSALAM-KENITRA_VD_Part1.jpg' -> IdFile(81, 'ASWAK ASSALAM KENITRA', {VD})."""
    stem = _COPIE_RE.sub("", Path(filename).stem)
    stem = _PART_RE.sub("", stem)
    match = _NOM_RE.match(stem)
    if not match:
        return None
    number, reste = match.group(1), match.group(2)
    divisions: frozenset[str] = frozenset()
    div = _DIV_RE.search(reste)
    if div:
        divisions = frozenset(re.findall(r"VD|DA", div.group(1).upper()))
        reste = reste[:div.start(1)]
    store = " ".join(re.sub(r"[_\-]+", " ", reste).split())
    if not _squash(store):
        return None
    return IdFile(number=str(int(number)), store=store, divisions=divisions, filename=filename)


class IdService:
    def __init__(self, files: list[IdFile], ignored: list[str] | None = None) -> None:
        self.files = files
        self.ignored = ignored or []                  # fichiers au nom non reconnu
        self._by_store: dict[str, list[IdFile]] = {}
        for f in files:
            self._by_store.setdefault(f.squashed, []).append(f)
        self.used: set[str] = set()                   # fichiers effectivement utilises
        self._connus: dict[str, frozenset[str]] = {}  # tous les Stores (Reference + fichier 2)
        self._attribution: dict[str, tuple[str, str]] | None = None   # fichier -> (Store, explication)

    @classmethod
    def from_folder(cls, folder: str | Path) -> "IdService":
        folder = Path(folder)
        if not folder.is_dir():
            raise FileNotFoundError(f"Dossier des IDs introuvable : {folder}")
        files, ignored = [], []
        for root, _dirs, names in os.walk(folder):
            for name in sorted(names):
                if name.startswith(("~$", ".")) or name.lower() in ("thumbs.db", "desktop.ini"):
                    continue
                parsed = parse_filename(name)
                if parsed is None:
                    ignored.append(str(Path(root, name).relative_to(folder)))
                else:
                    files.append(parsed)
        return cls(files, ignored)

    def set_known_stores(self, names: Iterable[str]) -> None:
        """Tous les Stores du traitement : evite d'attribuer a un Store le fichier d'un autre."""
        for name in names:
            cle = _squash(name)
            if cle:
                self._connus[cle] = _mots(name)
        self._attribution = None

    def describe(self) -> str:
        texte = f"{len(self.files)} fichier(s) avec un numero, {len(self._by_store)} Store(s)"
        if self.ignored:
            texte += f", {len(self.ignored)} ignore(s) (nom sans 'numero-Store')"
        return texte

    # ------------------------------------------------------------------
    def _autre_store_prioritaire(self, cles: set[str], mots: frozenset[str], cle_fichier: str,
                                 mots_fichier: frozenset[str]) -> bool:
        """Un AUTRE Store connu colle-t-il au moins aussi bien a ce fichier ?"""
        for cle, mots_autre in self._connus.items():
            if cle in cles:
                continue
            if cle == cle_fichier:                                    # nom identique : c'est le sien
                return True
            if len(mots_autre) >= MOTS_MINIMUM and mots_autre <= mots_fichier \
                    and len(mots_autre) >= len(mots) and mots_autre != mots:
                return True
        return False

    def _a_son_fichier(self, cle_store: str) -> bool:
        """Ce Store a-t-il deja un fichier a son nom (identique ou contenu) ?"""
        mots = self._connus[cle_store]
        return any(cle == cle_store or (len(mots) >= MOTS_MINIMUM and mots <= fichiers[0].mots)
                   for cle, fichiers in self._by_store.items())

    def _attribuer(self) -> dict[str, tuple[str, str]]:
        """Fichiers au nom ecrit autrement -> Store, seulement quand il n'y a aucun doute (regles 3a / 3b)."""
        if self._attribution is not None:
            return self._attribution
        frequence: dict[str, int] = {}
        for mots in self._connus.values():
            for mot in mots:
                frequence[mot] = frequence.get(mot, 0) + 1
        attribution: dict[str, tuple[str, str]] = {}
        for cle_fichier, fichiers in self._by_store.items():
            if cle_fichier in self._connus:
                continue                                              # nom identique : deja gere
            mots_f = fichiers[0].mots
            # 3a. mots communs avec un seul Store, dont au moins un mot propre a ce Store
            meilleurs, record = [], 0
            for cle, mots in self._connus.items():
                communs = mots & mots_f
                if len(communs) < MOTS_MINIMUM or 2 * len(communs) < len(mots):
                    continue
                if len(communs) > record:
                    meilleurs, record = [(cle, communs)], len(communs)
                elif len(communs) == record:
                    meilleurs.append((cle, communs))
            if len(meilleurs) == 1:
                cle, communs = meilleurs[0]
                propres = sorted(w.upper() for w in communs if frequence.get(w) == 1)
                if propres:
                    attribution[cle_fichier] = (cle, "nom du fichier different, mot propre a ce Store : "
                                                + ", ".join(propres))
                    continue
            # 3b. nom raccourci contenu dans plusieurs Stores : on ecarte ceux qui ont deja leur fichier
            if len(mots_f) >= MOTS_MINIMUM:
                contenants = [cle for cle, mots in self._connus.items() if mots_f <= mots]
                restants = [cle for cle in contenants if not self._a_son_fichier(cle)]
                if len(contenants) >= 1 and len(restants) == 1:
                    autres = [c for c in contenants if c not in restants]
                    detail = "nom du fichier raccourci" + (
                        f" (les autres Stores possibles ont deja leur fichier)" if autres else "")
                    attribution[cle_fichier] = (restants[0], detail)
        self._attribution = attribution
        return attribution

    def _store_files(self, noms: list[str]) -> tuple[list[IdFile], str]:
        """Fichiers du Store : (fichiers, '' | 'exact' | description de l'approximation)."""
        cles = {_squash(n) for n in noms if _squash(n)}
        for cle in cles:
            if cle in self._by_store:
                return self._by_store[cle], "exact"

        # 2. nom du Store contenu dans le nom du fichier (mots en plus)
        contenus: dict[str, str] = {}
        for nom in noms:
            mots = _mots(nom)
            if len(mots) < MOTS_MINIMUM:
                continue
            for cle_fichier, fichiers in self._by_store.items():
                mots_fichier = fichiers[0].mots
                if mots <= mots_fichier and not self._autre_store_prioritaire(cles, mots, cle_fichier,
                                                                              mots_fichier):
                    en_plus = sorted(w.upper() for w in mots_fichier - mots)
                    contenus[cle_fichier] = "mot(s) en plus dans le fichier : " + ", ".join(en_plus)
        if len(contenus) == 1:
            cle_fichier, detail = next(iter(contenus.items()))
            return self._by_store[cle_fichier], detail
        if len(contenus) > 1:
            return [], "plusieurs fichiers de Stores differents contiennent ce nom : " + ", ".join(
                self._by_store[c][0].store for c in sorted(contenus))

        # 3. nom ecrit autrement mais sans doute possible (mots propres / nom raccourci)
        attribues = [(cle_fichier, detail) for cle_fichier, (cle_store, detail) in self._attribuer().items()
                     if cle_store in cles]
        if attribues:
            fichiers = [f for cle_fichier, _ in attribues for f in self._by_store[cle_fichier]]
            return fichiers, attribues[0][1]

        # 4. nom tres proche (faute de frappe)
        meilleur: tuple[float, str] | None = None
        second = 0.0
        nb_mots = {len(_mots(n)) for n in noms}
        for cle_fichier, fichiers_store in self._by_store.items():
            if cle_fichier in self._connus and cle_fichier not in cles:
                continue                                              # fichier d'un autre Store
            if len(fichiers_store[0].mots) not in nb_mots:
                continue          # "Carrefour Al Mazar" n'est pas "CARREFOUR ALMAZ" (autre magasin)
            score = max((SequenceMatcher(None, cle, cle_fichier).ratio() for cle in cles), default=0)
            if meilleur is None or score > meilleur[0]:
                second = meilleur[0] if meilleur else 0.0
                meilleur = (score, cle_fichier)
            elif score > second:
                second = score
        if meilleur and meilleur[0] >= SEUIL_APPROCHE and meilleur[0] - second >= ECART_MINIMUM:
            return self._by_store[meilleur[1]], f"nom approche {meilleur[0]:.0%}"
        return [], ""

    def lookup(self, noms_store: list[str], division: str, divisions_du_store: set[str]) -> IdResult:
        """ID de la ligne Store + Division.

        noms_store         : noms possibles du Store (Reference puis fichier 2)
        divisions_du_store : divisions VD/DA du Store dans la Reference (pour les fichiers sans division)
        """
        fichiers, mode = self._store_files(noms_store)
        if not fichiers:
            if mode:
                return IdResult(AMBIGUOUS, reason=f"{mode} : ID non ecrit")
            return IdResult(NOT_FOUND, reason=f"aucun fichier du dossier des IDs pour '{noms_store[0]}'")

        candidats = []
        for f in fichiers:
            couvre = set(f.divisions) if f.divisions else set(divisions_du_store) or {division}
            if division in couvre:
                valeur = f.number + division if len(couvre) > 1 else f.number
                candidats.append((valeur, f))
        if not candidats:
            presentes = ", ".join(sorted({"+".join(sorted(f.divisions)) for f in fichiers if f.divisions}))
            return IdResult(NOT_FOUND, reason=f"fichier(s) de '{fichiers[0].store}' trouve(s) ({presentes}) "
                                              f"mais aucun pour la division {division}")

        valeurs = sorted({v for v, _ in candidats})
        noms = ", ".join(f.filename for _, f in candidats)
        if len(valeurs) > 1:
            return IdResult(AMBIGUOUS, reason=f"plusieurs fichiers donnent des IDs differents ({noms}) : ID non ecrit")
        self.used.update(f.filename for _, f in candidats)
        if mode == "exact":
            return IdResult(FOUND, valeurs[0], f"ID {valeurs[0]} lu dans le fichier '{noms}'")
        return IdResult(APPROX, valeurs[0], f"ID {valeurs[0]} lu dans '{noms}' : {mode}, a controler")

    def unused(self) -> list[IdFile]:
        return [f for f in self.files if f.filename not in self.used]
