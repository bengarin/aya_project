# Distribuer le logiciel sous Windows (sans Python)

## 1. Ce que reçoit l'utilisateur

Deux choix, **le même logiciel** dedans :

| Fichier | Pour qui | Comment |
|---|---|---|
| `AYA_Excel_Setup_<version>.exe` (≈ 11 Mo) | utilisateur normal (recommandé) | double-clic → Suivant → Installer. Raccourcis Bureau + menu Démarrer, désinstallation dans « Applications installées ». **Pas besoin d'être administrateur.** |
| `AYA_Excel_Portable_<version>.zip` (≈ 15 Mo) | PC verrouillé, clé USB | extraire le ZIP → double-clic sur `AYA_Excel.exe` |

Rien d'autre à installer : Python, openpyxl, Tkinter, CustomTkinter sont **dans** le logiciel.
Windows 10 ou 11, 64 bits.

**Premier lancement** : Windows SmartScreen peut dire « Windows a protégé votre ordinateur »
car le logiciel n'est pas signé → « Informations complémentaires » → « Exécuter quand même ».
Pour supprimer ce message il faut un certificat de signature de code (payant, ~200-400 €/an) [NON VÉRIFIÉ : prix à comparer].

## 2. Où sont les données

| Quoi | Où | Touché par une mise à jour ? |
|---|---|---|
| le programme | `%LOCALAPPDATA%\Programs\AYA Excel\` (installateur) ou le dossier extrait (ZIP) | oui, remplacé |
| derniers fichiers choisis | `%APPDATA%\AYA Excel\config.json` | **jamais** |
| journal d'erreurs | `%LOCALAPPDATA%\AYA Excel\logs\aya_excel.log` | **jamais** |
| fichiers Excel de l'utilisateur | là où il les a mis | **jamais** (le résultat est toujours un NOUVEAU fichier) |

Il n'y a **pas de base de données**, pas de mot de passe, pas de clé API, pas de serveur :
le logiciel lit 3 fichiers Excel et écrit un nouveau fichier. Rien à protéger ou à cacher dans l'EXE.

## 3. Mettre à jour chez l'utilisateur

1. Changer le numéro dans `version.py` (ex. `1.1.0` → `1.2.0`).
2. Construire (section 4).
3. Envoyer le nouvel `AYA_Excel_Setup_1.2.0.exe`.
4. L'utilisateur ferme le logiciel et lance le nouvel installateur : il remplace l'ancienne version
   au même endroit (même `AppId` dans `packaging/installer.iss` — **ne jamais le changer**).
   Réglages et fichiers Excel conservés.

Version ZIP : supprimer l'ancien dossier `AYA_Excel`, extraire le nouveau.

## 4. Construire (développeur uniquement)

### Option A — sur un PC Windows
Une seule fois : installer Python 3.11 64 bits (python.org) et, pour l'installateur, Inno Setup 6.
Puis :
```powershell
powershell -ExecutionPolicy Bypass -File build_windows.ps1
```
Le script fait tout : environnement isolé, versions exactes (`packaging/requirements-build.txt`),
tests métier, PyInstaller, vérification que l'EXE donne **le même résultat** que le code source,
test d'ouverture de la fenêtre, ZIP portable, installateur. Résultat dans `release\`.

### Option B — automatique sur GitHub (vrai Windows)
`.github/workflows/build-windows.yml` : à chaque push, construit sur Windows Server 2022,
puis teste sur une 2e machine **sans Python dans le PATH** : installation sans admin, raccourci,
fenêtre, traitement des 10 cas de test, désinstallation (config conservée), ZIP portable.
Télécharger le résultat : onglet **Actions** → dernier run → artefact `AYA_Excel_Windows`.

### Option C — depuis Linux (Wine)
```bash
sudo apt install wine64 wine32:i386 xvfb
bash packaging/build_with_wine.sh
```
Python Windows et Inno Setup sont téléchargés avec une empreinte SHA-256 vérifiée.

## 5. Architecture du paquet

```
AYA_Excel\
  AYA_Excel.exe        application fenêtre (pas de console noire)
  AYA_Excel_CLI.exe    même moteur en ligne de commande (support, automatisation)
  _internal\           Python 3.11 + openpyxl + Tkinter/Tcl-Tk + CustomTkinter + icône
```
- PyInstaller en mode **dossier** (pas « un seul EXE ») : démarrage rapide (pas de décompression
  à chaque lancement) et moins d'alertes antivirus. UPX désactivé pour la même raison.
- Recette : `packaging/aya_excel.spec`. Icône : `assets/aya.ico` (`packaging/make_icon.py`).

## 6. Erreurs

- Erreur pendant un traitement → message clair dans la fenêtre + détail dans le journal.
- Erreur inattendue (bouton, démarrage) → boîte de dialogue + journal ; le logiciel ne se ferme pas en silence.
- Support : demander le fichier `%LOCALAPPDATA%\AYA Excel\logs\aya_excel.log`.

## 7. Commandes utiles (support)

```
AYA_Excel_CLI.exe --version
AYA_Excel_CLI.exe --cli -r Reference.xlsx -t BDD.xlsx -k KAM.xlsx -o Resultat.xlsx
AYA_Excel_CLI.exe --cli ... --preview        (analyse sans rien écrire)
AYA_Excel.exe --smoke-test                   (ouvre la fenêtre 1,5 s puis la ferme : test)
```
