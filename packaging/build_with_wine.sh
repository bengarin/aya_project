#!/usr/bin/env bash
# build_with_wine.sh - Construit l'EXE Windows depuis LINUX (Ubuntu 24.04) avec Wine.
#
# Meme resultat que build_windows.ps1, pour construire sans PC Windows.
#   sudo apt install wine64 wine32:i386 xvfb      (une fois)
#   bash packaging/build_with_wine.sh
# Derriere un proxy HTTPS avec certificat maison : AYA_PIP_CERT=/chemin/ca.crt bash ...
#
# Outils telecharges (versions et empreintes SHA-256 FIXES = build reproductible) :
#   - Python 3.11 Windows "python-build-standalone" (avec Tkinter)
#   - Inno Setup 6.5.4 (installateur)
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="${AYA_BUILD_DIR:-$ROOT/packaging/build}"
export WINEPREFIX="$WORK/wineprefix" WINEARCH=win64 WINEDEBUG=-all

PY_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20250317/cpython-3.11.11+20250317-x86_64-pc-windows-msvc-install_only.tar.gz"
PY_SHA="1cf5760eea0a9df3308ca2c4111b5cc18fd638b2a912dbe07606193e3f9aa123"
IS_URL="https://github.com/jrsoftware/issrc/releases/download/is-6_5_4/innosetup-6.5.4.exe"
IS_SHA="fa73bf47a4da250d185d07561c2bfda387e5e20db77e4570004cf6a133cc10b1"

VERSION="$(python3 -c "import sys; sys.path.insert(0, '$ROOT'); from version import VERSION; print(VERSION)")"
win() { echo "Z:${1//\//\\}"; }                          # chemin Linux -> chemin Wine
fetch() {                                                # telecharge + verifie l'empreinte
    [ -f "$2" ] || curl -sSL -o "$2" "$1"
    echo "$3  $2" | sha256sum -c --quiet || { echo "Empreinte invalide : $2"; rm -f "$2"; exit 1; }
}

mkdir -p "$WORK" "$ROOT/release"
cd "$WORK"

echo "=== 1/6 Outils ==="
WINE64="$(command -v wine64 || echo /usr/lib/wine/wine64)"       # prefixe 64 bits obligatoire
[ -d wineprefix ] || xvfb-run -a "$WINE64" wineboot -i >/dev/null 2>&1
fetch "$PY_URL" python.tar.gz "$PY_SHA"
[ -d python ] || tar xzf python.tar.gz
wine python/python.exe -m pip install -q --disable-pip-version-check --no-warn-script-location \
    ${AYA_PIP_CERT:+--cert "$(win "$AYA_PIP_CERT")"} -r "$(win "$ROOT/packaging/requirements-build.txt")" </dev/null

echo "=== 2/6 Tests metier ==="
(cd "$ROOT" && python3 -m unittest discover -s tests -q)

echo "=== 3/6 PyInstaller ==="
rm -rf dist pyi
wine python/python.exe -m PyInstaller --noconfirm --clean --distpath "$(win "$WORK/dist")" \
    --workpath "$(win "$WORK/pyi")" "$(win "$ROOT/packaging/aya_excel.spec")" </dev/null 2>&1 | tail -3

echo "=== 4/6 Verification ==="
python3 "$ROOT/packaging/verify_package.py" wine "$WORK/dist/AYA_Excel/AYA_Excel_CLI.exe" </dev/null | tail -2
xvfb-run -a wine "$WORK/dist/AYA_Excel/AYA_Excel.exe" --smoke-test </dev/null && echo "Fenetre : OK"

echo "=== 5/6 ZIP portable ==="
cp "$ROOT/packaging/LISEZMOI.txt" dist/AYA_Excel/LISEZMOI.txt
ZIP="$ROOT/release/AYA_Excel_Portable_$VERSION.zip"
rm -f "$ZIP"; (cd dist && python3 -m zipfile -c "$ZIP" AYA_Excel)
echo "Cree : $ZIP"

echo "=== 6/6 Installateur ==="
fetch "$IS_URL" innosetup.exe "$IS_SHA"
[ -f "$WINEPREFIX/drive_c/InnoSetup6/ISCC.exe" ] || \
    xvfb-run -a wine innosetup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- '/DIR=C:\InnoSetup6' </dev/null >/dev/null 2>&1
wine 'C:\InnoSetup6\ISCC.exe' /Q "/DAppVersion=$VERSION" "/DSourceDir=$(win "$WORK/dist/AYA_Excel")" \
    "/DOutputDir=$(win "$ROOT/release")" "$(win "$ROOT/packaging/installer.iss")" </dev/null
echo "Cree : $ROOT/release/AYA_Excel_Setup_$VERSION.exe"
