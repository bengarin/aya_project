# -*- mode: python ; coding: utf-8 -*-
"""
Recette PyInstaller du logiciel "Automatisation Excel AYA".

Produit UN dossier  dist/AYA_Excel/  qui contient :
  - AYA_Excel.exe       -> l'application (fenetre, sans console)  : c'est celle de l'utilisateur
  - AYA_Excel_CLI.exe   -> la meme application en ligne de commande (tests, automatisation)
  - _internal/          -> Python + openpyxl + Tkinter + CustomTkinter (rien a installer)

Mode "dossier" (onedir) plutot qu'un EXE unique : demarrage rapide, moins de
faux positifs antivirus, et les deux EXE partagent les memes fichiers.

Construction : voir build_windows.ps1 (ne pas lancer a la main).
"""

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo, StringFileInfo, StringStruct, StringTable, VarFileInfo, VarStruct, VSVersionInfo,
)

ROOT = Path(SPECPATH).resolve().parent
sys.path.insert(0, str(ROOT))
from version import APP_NAME, AUTHOR, COPYRIGHT, VERSION  # noqa: E402

ICON = str(ROOT / "assets" / "aya.ico")
_v = tuple(int(x) for x in VERSION.split(".")) + (0,) * (4 - len(VERSION.split(".")))


def version_info(internal_name: str) -> VSVersionInfo:
    return VSVersionInfo(
        ffi=FixedFileInfo(filevers=_v, prodvers=_v),
        kids=[
            StringFileInfo([StringTable("040C04B0", [
                StringStruct("CompanyName", AUTHOR),
                StringStruct("FileDescription", APP_NAME),
                StringStruct("LegalCopyright", COPYRIGHT),
                StringStruct("FileVersion", VERSION),
                StringStruct("InternalName", internal_name),
                StringStruct("OriginalFilename", internal_name + ".exe"),
                StringStruct("ProductName", APP_NAME),
                StringStruct("ProductVersion", VERSION),
            ])]),
            VarFileInfo([VarStruct("Translation", [0x040C, 1200])]),
        ],
    )


a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    binaries=[],
    # themes / polices de CustomTkinter + icone du logiciel
    datas=collect_data_files("customtkinter") + [(ICON, "assets")],
    hiddenimports=["gui", "darkdetect"],     # gui est importe dans main() : on le declare
    hookspath=[],
    runtime_hooks=[],
    # modules inutiles au logiciel (reduit la taille, rien de fonctionnel en moins)
    excludes=["PIL", "numpy", "pandas", "matplotlib", "pytest", "tests",
              "setuptools", "pkg_resources", "distutils", "unittest", "pydoc_data"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe_gui = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="AYA_Excel",
    console=False,                 # pas de fenetre noire
    icon=ICON,
    version=version_info("AYA_Excel"),
    upx=False,                     # UPX declenche des faux positifs antivirus
)
exe_cli = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="AYA_Excel_CLI",
    console=True,
    icon=ICON,
    version=version_info("AYA_Excel_CLI"),
    upx=False,
)
coll = COLLECT(
    exe_gui, exe_cli,
    a.binaries, a.datas,
    upx=False,
    name="AYA_Excel",
)
