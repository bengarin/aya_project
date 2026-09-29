<#
build_windows.ps1 - Construit le logiciel Windows "Automatisation Excel AYA"
============================================================================

A lancer sur un PC Windows 10/11 (developpeur), PAS chez l'utilisateur final :

    powershell -ExecutionPolicy Bypass -File build_windows.ps1

Pre-requis (une seule fois, sur le PC de construction uniquement) :
  - Python 3.11 64 bits (python.org)  -> commande "py -3.11" disponible
  - optionnel : Inno Setup 6 (jrsoftware.org) pour produire l'installateur

Etapes :
  1. environnement Python isole (.venv-build) + dependances aux versions exactes
  2. tests metier (35 tests)
  3. PyInstaller  -> dist\AYA_Excel\  (AYA_Excel.exe + AYA_Excel_CLI.exe)
  4. verification : l'EXE donne le meme resultat que le code source,
     et la fenetre s'ouvre/se ferme sans erreur (--smoke-test)
  5. release\AYA_Excel_Portable_<version>.zip
  6. release\AYA_Excel_Setup_<version>.exe  (si Inno Setup est installe)
#>

param(
    [string]$Python = "py",          # lanceur Python de Windows
    [string]$PythonVersion = "-3.11",
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Step($text) { Write-Host "`n=== $text ===" -ForegroundColor Cyan }
function Run($exe, [string[]]$argsList) {
    & $exe @argsList
    if ($LASTEXITCODE -ne 0) { throw "Echec : $exe $($argsList -join ' ') (code $LASTEXITCODE)" }
}

$version = (Select-String -Path version.py -Pattern 'VERSION = "(.+)"').Matches[0].Groups[1].Value
Write-Host "Version : $version"

Step "1/6 Environnement de construction"
if (-not (Test-Path .venv-build)) {
    & $Python $PythonVersion -m venv .venv-build
    if ($LASTEXITCODE -ne 0) { throw "Python 3.11 introuvable. Installez-le depuis python.org." }
}
$py = Join-Path $PSScriptRoot ".venv-build\Scripts\python.exe"
Run $py @("-m", "pip", "install", "--disable-pip-version-check", "-q", "-r", "packaging\requirements-build.txt")

Step "2/6 Tests metier"
Run $py @("-m", "unittest", "discover", "-s", "tests", "-q")

Step "3/6 PyInstaller"
Run $py @("-m", "PyInstaller", "--noconfirm", "--clean",
          "--distpath", "dist", "--workpath", "build", "packaging\aya_excel.spec")

Step "4/6 Verification de l'EXE"
Run $py @("packaging\verify_package.py", "dist\AYA_Excel\AYA_Excel_CLI.exe")
$gui = Start-Process -FilePath "dist\AYA_Excel\AYA_Excel.exe" -ArgumentList "--smoke-test" -PassThru -Wait
if ($gui.ExitCode -ne 0) { throw "La fenetre ne demarre pas (code $($gui.ExitCode))" }
Write-Host "Fenetre : OK"

Step "5/6 ZIP portable"
New-Item -ItemType Directory -Force -Path release | Out-Null
Copy-Item packaging\LISEZMOI.txt dist\AYA_Excel\LISEZMOI.txt -Force
$zip = "release\AYA_Excel_Portable_$version.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path dist\AYA_Excel -DestinationPath $zip
Write-Host "Cree : $zip"

Step "6/6 Installateur"
$iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
          "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($SkipInstaller) {
    Write-Host "Installateur ignore (-SkipInstaller)."
} elseif (-not $iscc) {
    Write-Host "Inno Setup 6 non trouve : installateur non construit (le ZIP portable suffit)." -ForegroundColor Yellow
} else {
    Run $iscc @("/Q", "/DAppVersion=$version", "/DSourceDir=$PSScriptRoot\dist\AYA_Excel",
                "/DOutputDir=$PSScriptRoot\release", "packaging\installer.iss")
    Write-Host "Cree : release\AYA_Excel_Setup_$version.exe"
}

Write-Host "`nTermine. Fichiers a distribuer : dossier release\" -ForegroundColor Green
