"""
tests/make_fixtures.py
----------------------
Fabrique 3 fichiers Excel de test qui reproduisent EXACTEMENT les 10 cas a tester :

 1. Store vide                         -> ligne 6
 2. Store non trouve                   -> ligne 7
 3. Store + VD existant                -> ligne 2
 4. Store + DA existant                -> ligne 3
 5. Store VD existant mais DA absent   -> ligne 4 (+ creation)
 6. Store VD+DA                        -> ligne 4
 7. Meme Store + meme Division, 2 promoteurs -> ligne 9
 8. KAM different entre VD et DA       -> lignes 2 et 3
 9. Column1 incorrect                  -> ligne 10
10. Deuxieme execution sans duplicate  -> test dedie

Structure utilisee (celle imposee par les regles) :
  Fichier 1 : feuille 'Feuil2', en-tetes ligne 1,
              C=Divisions F=ID Promoter G=Store Code H=PROMOTER I=STORE J=CITY L=Working Days
  Fichier 2 : feuille 'BDD PROMOTERS MONTH', en-tetes ligne 1, colonnes A..N
  Fichier 3 : feuilles 'VD' et 'DA'
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- Fichier 1 : Reference (colonne -> lettre)
REF_COLUMNS = {"Divisions": 3, "ID Promoter": 6, "Store Code": 7, "PROMOTER": 8,
               "STORE": 9, "CITY": 10, "Working Days": 12}

# Divisions, ID Promoter, Store Code, PROMOTER, STORE, CITY, Working Days
REF_ROWS = [
    ("VD", "IDVD123", "C123", "PROMO VD KENITRA", "Aswak Assalam Kenitra", "Kenitra", 26),
    ("DA", "IDDA123", "C123", "PROMO DA KENITRA", "Aswak Assalam Kenitra", "Kenitra", 24),
    ("VD+DA", "IDVDDA200", "C200", "PROMO VDDA", "Store VDDA", "Casablanca", 26),
    ("VD", "IDSOLO300", "C300", "PROMO SOLO", "  Store   Solo ", "Fes", 26),
    ("RAC", "IDRAC400", "C400", "PROMO RAC", "Store Rac", "Fes", 17),
    ("VD", "IDCONF1", "C500", "PROMO CONFLIT 1", "Store Conflit", "Oujda", 26),
    ("VD", "IDCONF2", "C500", "PROMO CONFLIT 2", "Store Conflit", "Oujda", 26),
    ("VD", "IDNOKAM600", "C600", "PROMO SANS KAM", "Store NoKam", "Safi", 26),
    ("VD", "IDALIASA", "C700", "PROMO ALIAS A", "Store Alias A", "Rabat", 26),
    ("VD", "IDALIASB", "C700", "PROMO ALIAS B", "Store Alias B", "Rabat", 26),
]

# --- Fichier 2 : BDD a traiter
TARGET_HEADERS = ["Année", "Mois", "Division", "Division1", "KAM", "ID Promoter", "Code Store",
                  "Promoter", "Store", "City", "STATUT", "DAY", "IDAYA", "Column1"]

TARGET_ROWS = [
    # 2 : VD existant, donnees fausses -> MODIFIEE (KAM VD attendu)
    [2026, "Juillet", "VD", "VD", "ANCIEN KAM", "ANCIEN_ID", "C123", "ANCIEN PROMO",
     "Aswak Assalam Kenitra", "Kenitra", "Actif", 20, 1, "C123VD"],
    # 3 : DA existant, deja exact -> CONFORME (KAM DA different du VD)
    [2026, "Juillet", "DA", "DA", "KAM DA KENITRA", "IDDA123", "C123", "PROMO DA KENITRA",
     "Aswak Assalam Kenitra", "Kenitra", "Actif", 24, 2, "C123DA"],
    # 4 : VD+DA -> VD existe (a corriger), DA absent -> sera CREE
    [2026, "Juillet", "VD+DA", "VD", "KAM VD VDDA", "IDVDDA200", "C200", "PROMO VDDA",
     "Store VDDA", "Casablanca", "Actif", 20, 3, "C200VD"],
    # 5 : Store VIDE -> ROUGE uniquement
    [2026, "Juillet", "VD", "VD", "KAM X", "ID_ORPHELIN", "C900", "PROMO ORPHELIN",
     "", "Casablanca", "Actif", 26, 4, "C900VD"],
    # 6 : Store absent de la Reference -> signale, intact
    [2026, "Juillet", "DA", "DA", "KAM Y", "ID_INCONNU", "C999", "PROMO INCONNU",
     "Store Inconnu", "Oujda", "Actif", 26, 5, "C999DA"],
    # 7 : division RAC -> A VERIFIER, intacte
    [2026, "Juillet", "RAC", "RAC", "KAM Z", "IDRAC400", "C400", "PROMO RAC",
     "Store Rac", "Fes", "Actif", 17, 6, "C400RAC"],
    # 8 : 2 promoteurs pour Store Conflit + VD -> A VERIFIER, intacte
    [2026, "Juillet", "VD", "VD", "KAM W", "ANCIEN_CONF", "C500", "ANCIEN PROMO CONF",
     "Store Conflit", "Oujda", "Actif", 26, 7, "C500VD"],
    # 9 : Column1 faux mais Code Store correct -> MODIFIEE, Column1 corrigee
    [2026, "Juillet", "VD", "VD", "KAM VD SOLO", "IDSOLO300", "C300", "PROMO SOLO",
     "Store Solo", "Fes", "Actif", 26, 8, "ANCIENNE_CLE_FAUSSE"],
    # 10 : Store Solo + DA n'existe pas dans la Reference -> IGNOREE
    [2026, "Juillet", "DA", "DA", "KAM ??", "IDSOLO300", "C300", "PROMO SOLO",
     "Store Solo", "Fes", "Actif", 26, 9, "C300DA"],
    # 11 : aucun KAM dans le fichier 3 -> KAM vide (rien d'invente)
    [2026, "Juillet", "VD", "VD", "ANCIEN KAM NOKAM", "IDNOKAM600", "C600", "PROMO SANS KAM",
     "Store NoKam", "Safi", "Actif", 26, 10, "C600VD"],
    # 12 et 13 : meme cle C700VD sur 2 lignes (2 noms de Store, meme Code Store)
    [2026, "Juillet", "VD", "VD", "KAM VD ALIAS", "ANCIEN_ALIAS", "C700", "ANCIEN PROMO ALIAS",
     "Store Alias A", "Rabat", "Actif", 26, 11, "C700VD"],
    [2026, "Juillet", "VD", "VD", "KAM VD ALIAS", "ANCIEN_ALIAS", "C700", "ANCIEN PROMO ALIAS",
     "Store Alias B", "Rabat", "Actif", 26, 12, "C700VD"],
]

# --- Fichier 3 : Affectation KAM (City, Code store, Store, KAM) en B..E, en-tetes ligne 2
KAM_VD = [
    ["Kenitra", "C123", "Aswak Assalam Kenitra", "KAM VD KENITRA"],
    ["Casablanca", "C200", "Store VDDA", "KAM VD VDDA"],
    ["Fes", "C300", "Store Solo", "KAM VD SOLO"],
    ["Rabat", "C700", "Store Alias A", "KAM VD ALIAS"],
    ["Rabat", "C700", "Store Alias B", "KAM VD ALIAS"],
]
KAM_DA = [
    ["Kenitra", "C123", "Aswak Assalam Kenitra", "KAM DA KENITRA"],
    ["Casablanca", "C200", "Store VDDA", "KAM DA VDDA"],
]


def build_reference(path: Path) -> Path:
    wb = Workbook()
    autre = wb.active
    autre.title = "Feuil1"                                  # feuille parasite
    autre["A1"] = "ancienne version"
    ws = wb.create_sheet("Feuil2")
    for name, column in REF_COLUMNS.items():
        ws.cell(1, column, name).font = Font(bold=True)
    for row_index, values in enumerate(REF_ROWS, start=2):
        for name, value in zip(REF_COLUMNS, values):
            ws.cell(row_index, REF_COLUMNS[name], value)
    ws.auto_filter.ref = f"C1:L{1 + len(REF_ROWS)}"
    wb.save(path)
    return path


def build_target(path: Path) -> Path:
    wb = Workbook()
    autre = wb.active
    autre.title = "Notes"                                   # feuille parasite
    autre["A1"] = "feuille sans rapport"

    ws = wb.create_sheet("BDD PROMOTERS MONTH")
    for index, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(1, index, header)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", start_color="FF1F4E79")
    for row_index, row in enumerate(TARGET_ROWS, start=2):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row_index, col_index, value).font = Font(name="Calibri", size=11)
    last = 1 + len(TARGET_ROWS)
    table = Table(displayName="BDD_TEST", ref=f"A1:N{last}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def build_kam(path: Path) -> Path:
    wb = Workbook()
    ws_vd = wb.active
    ws_vd.title = "VD"
    ws_da = wb.create_sheet("DA")
    for sheet, rows in ((ws_vd, KAM_VD), (ws_da, KAM_DA)):
        for index, header in enumerate(["CITY", "Code store", "Store", "KAM"], start=2):
            sheet.cell(2, index, header).font = Font(bold=True)
        for row_index, row in enumerate(rows, start=3):
            for col_index, value in enumerate(row, start=2):
                sheet.cell(row_index, col_index, value)
        sheet.auto_filter.ref = f"B2:E{2 + len(rows)}"
    wb.save(path)
    return path


def build_all(folder: Path) -> dict[str, Path]:
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    return {
        "reference": build_reference(folder / "TEST_Reference.xlsx"),
        "target": build_target(folder / "TEST_BDD.xlsx"),
        "kam": build_kam(folder / "TEST_KAM.xlsx"),
    }


if __name__ == "__main__":
    import sys

    destination = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("fixtures")
    for name, path in build_all(destination).items():
        print(f"{name:10} -> {path}")
