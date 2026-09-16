"""
matcher.py
----------
Tout ce qui sert a COMPARER des textes entre les 3 fichiers Excel.

Pourquoi ce fichier existe :
dans la vraie vie, "Aswak  Assalam Mohammedia " et "aswak assalam mohammedia"
sont le meme magasin, mais pour l'ordinateur ce sont 2 textes differents.
Ici on "nettoie" les textes avant de les comparer.
"""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata

# Divisions connues et explicitement decrites dans les regles metier.
KNOWN_DIVISIONS = ("VD", "DA")

# Separateurs possibles d'une division combinee : "VD+DA", "VD/DA", "VD & DA"...
_DIVISION_SPLIT = re.compile(r"\s*[+/,;&]\s*|\s+ET\s+")

# Espaces "exotiques" (insecables) que Excel ramene souvent depuis des copier/coller.
_WEIRD_SPACES = dict.fromkeys(map(ord, "   ​\t\r\n"), " ")

_QUOTES = {ord("’"): "'", ord("‘"): "'", ord("“"): '"', ord("”"): '"'}


def cell_to_text(value) -> str:
    """Transforme la valeur brute d'une cellule Excel en texte propre.

    - None                -> ""
    - 26.0 (float entier) -> "26"   (sinon on ecrirait "26.0" dans le fichier)
    - date                -> format ISO court
    - texte               -> texte sans espaces inutiles
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "VRAI" if value else "FAUX"
    if isinstance(value, float):
        # Excel stocke souvent les entiers en float : 26.0 -> "26"
        if value.is_integer():
            return str(int(value))
        return repr(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()[:10]
    text = str(value).translate(_WEIRD_SPACES).translate(_QUOTES)
    return " ".join(text.split()).strip()


def clean_text(value) -> str:
    """Alias lisible de cell_to_text (texte visible, sans transformation de casse)."""
    return cell_to_text(value)


def strip_accents(text: str) -> str:
    """Enleve les accents : 'Mohammédia' -> 'Mohammedia'."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def norm_text(value) -> str:
    """Forme normalisee pour COMPARER (minuscules, sans accent, espaces simples)."""
    return strip_accents(cell_to_text(value)).casefold().strip()


def norm_store(value) -> str:
    """Normalisation dediee aux noms de magasins.

    En plus de norm_text, on neutralise les ponctuations de separation
    ('-', '_', '.') qui varient d'un fichier a l'autre.
    """
    text = norm_text(value)
    text = re.sub(r"[\-_.]+", " ", text)
    return " ".join(text.split())


def norm_code(value) -> str:
    """Normalisation des codes magasin : majuscules, sans espace."""
    return re.sub(r"\s+", "", cell_to_text(value)).upper()


def norm_division(value) -> str:
    """Normalisation d'une division : 'vd ' -> 'VD'."""
    return re.sub(r"\s+", "", strip_accents(cell_to_text(value))).upper()


def split_divisions(value) -> list[str]:
    """Transforme la valeur 'Divisions' de la Reference en liste de divisions.

    'VD'     -> ['VD']
    'VD+DA'  -> ['VD', 'DA']
    'RAC'    -> ['RAC']            (division inconnue, geree plus loin)
    ''       -> []
    """
    raw = norm_division(value)
    if not raw:
        return []
    parts = [p for p in _DIVISION_SPLIT.split(raw) if p]
    # On garde l'ordre d'apparition sans doublon.
    seen: list[str] = []
    for part in parts:
        if part not in seen:
            seen.append(part)
    return seen


def is_known_division(division: str) -> bool:
    return norm_division(division) in KNOWN_DIVISIONS


def build_key(store_code, division) -> str:
    """Construit la cle metier 'Code Store + Division1'.

    Exemple : ('C003470765', 'VD') -> 'C003470765VD'
    C'est la valeur attendue dans la colonne Column1 du fichier 2.
    """
    return f"{norm_code(store_code)}{norm_division(division)}"


def norm_key(value) -> str:
    """Normalise une cle deja existante (ex: contenu de Column1) pour comparaison."""
    return norm_code(value)


def is_blank(value) -> bool:
    """Vrai si la cellule est vide (None, '', espaces uniquement)."""
    return cell_to_text(value) == ""
