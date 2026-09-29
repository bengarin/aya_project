"""
app_paths.py
------------
Ou le logiciel range SES petits fichiers (jamais les fichiers Excel de l'utilisateur).

  - config (derniers fichiers utilises) :
        Windows : %APPDATA%\\AYA Excel\\config.json
        autres  : ~/.config/aya-excel/config.json
  - journaux d'erreurs :
        Windows : %LOCALAPPDATA%\\AYA Excel\\logs\\aya_excel.log
        autres  : ~/.cache/aya-excel/logs/aya_excel.log

Ces dossiers sont dans le profil de l'utilisateur : ecriture possible sans droits
administrateur, et une mise a jour du logiciel (nouveau dossier / nouvel EXE)
ne les touche jamais.

resource_path() retrouve les fichiers livres avec le logiciel (icone...), que
l'on tourne depuis les sources Python ou depuis l'EXE PyInstaller.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import sys
from pathlib import Path

APP_DIR_NAME = "AYA Excel"
LEGACY_CONFIG = Path.home() / ".aya_excel.json"        # emplacement des versions < 1.1


def is_frozen() -> bool:
    """True quand le code tourne depuis l'EXE (PyInstaller)."""
    return bool(getattr(sys, "frozen", False))


def resource_path(name: str) -> Path:
    """Chemin d'un fichier livre avec le logiciel (dossier 'assets')."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "assets" / name


def _windows_dir(env: str, fallback: Path) -> Path:
    value = os.environ.get(env)
    return Path(value) if value else fallback


def config_dir() -> Path:
    if sys.platform.startswith("win"):
        return _windows_dir("APPDATA", Path.home() / "AppData" / "Roaming") / APP_DIR_NAME
    return Path.home() / ".config" / "aya-excel"


def log_dir() -> Path:
    if sys.platform.startswith("win"):
        return _windows_dir("LOCALAPPDATA", Path.home() / "AppData" / "Local") / APP_DIR_NAME / "logs"
    return Path.home() / ".cache" / "aya-excel" / "logs"


def config_file() -> Path:
    return config_dir() / "config.json"


def read_config_text() -> str | None:
    """Contenu de la config ; reprend l'ancien fichier ~/.aya_excel.json si besoin."""
    for path in (config_file(), LEGACY_CONFIG):
        try:
            return path.read_text(encoding="utf-8")
        except OSError:
            continue
    return None


def write_config_text(text: str) -> None:
    """Ecriture atomique : un crash pendant l'ecriture ne corrompt pas la config."""
    path = config_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def setup_logging() -> Path | None:
    """Journal des erreurs (1 Mo x 3 fichiers). Ne bloque jamais le demarrage."""
    try:
        folder = log_dir()
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / "aya_excel.log"
        handler = logging.handlers.RotatingFileHandler(
            path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        root = logging.getLogger()
        root.addHandler(handler)
        root.setLevel(logging.INFO)
        return path
    except Exception:                                   # pragma: no cover
        return None
