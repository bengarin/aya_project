"""
tests/make_fixtures.py
----------------------
Fabrique 3 petits fichiers Excel de test qui reproduisent TOUS les cas des
regles metier : VD+DA, meme store sur 2 lignes, Store vide, store inconnu,
KAM manquant, conflit de promoteurs, division RAC, cle Column1 incoherente.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

REF_HEADERS = ["ID Promoter", "Store Code", "GRADE", "PROMOTER", "STORE",
               "Family", "Sub-channel", "CITY", "Divisions", "Working Days"]

REF_ROWS = [
    # VD+DA sur une seule ligne -> doit produire 2 affectations
    ["HT204806", "C000114373", "Grade A", "ROUAK MOULAY HICHAM", "Aswak Assalam Mohammedia",
     "Aswak", "Hyper", "Mohammedia", "VD+DA", 26],
    # meme store, 2 lignes, promoteurs differents
    ["AAD250045", "C003470765", "Grade A", "HASBI ABDELLAH", "Aswak Assalam Hay Riad",
     "Aswak", "Hyper", "Rabat", "VD", 24],
    ["BA331963", "C003470765", "Grade A", "BAGHDADI ABDELMOUIJIB", "Aswak Assalam Hay Riad",
     "Aswak", "Hyper", "Rabat", "DA", 26],
    # store normal, mais Column1 incoherent dans le fichier 2
    ["XX111111", "C999999999", "Grade A", "PROMO SOLO", "  Store   Solo ",
     "Autre", "Retailer", "Fes", "VD", 26],
    # division non prevue par les regles
    ["RAC11111", "C888888888", "Grade A", "PROMO RAC", "Store Rac", "Autre", "Retailer", "Fes", "RAC", 17],
    # conflit : 2 promoteurs differents pour la meme division
    ["CONF1111", "C777777777", "Grade A", "PROMO A", "Store Conflit", "Autre", "Retailer", "Fes", "VD", 26],
    ["CONF2222", "C777777777", "Grade A", "PROMO B", "Store Conflit", "Autre", "Retailer", "Fes", "VD", 26],
    # store absent du fichier 2 (utilise seulement avec l'option dediee)
    ["NEW11111", "C555555555", "Grade A", "PROMO ABSENT", "Store Absent BDD", "Autre", "Retailer", "Fes", "VD", 26],
]

TARGET_HEADERS = ["Annee", "Mois", "Division", "Division1", "KAM", "ID Promoter", "Code Store",
                  "Promoter", "Store", "City", "STATUT", "DAY", "IDAYA", "Column1"]

TARGET_ROWS = [
    # 2 : VD existant a corriger (promoteur + jours faux, KAM vide) ; le DA devra etre cree
    [2026, "Juillet", "VD", "VD", "", "ANCIEN_ID", "C000114373", "ANCIEN PROMO",
     "Aswak Assalam Mohammedia", "Mohammedia", "Actif", 20, 1, "C000114373VD"],
    # 3 : VD deja bon sauf le KAM
    [2026, "Juillet", "VD", "VD", "MAUVAIS KAM", "AAD250045", "C003470765", "HASBI ABDELLAH",
     "Aswak Assalam Hay Riad", "Rabat", "Actif", 24, 2, "C003470765VD"],
    # 4 : deja parfaitement conforme -> aucune modification attendue
    [2026, "Juillet", "DA", "DA", "KAM DA2", "BA331963", "C003470765", "BAGHDADI ABDELMOUIJIB",
     "Aswak Assalam Hay Riad", "Rabat", "Actif", 26, 3, "C003470765DA"],
    # 5 : Store VIDE -> ligne rouge, aucune autre modification
    [2026, "Juillet", "VD", "VD", "KAM X", "ID_ORPHELIN", "C123456789", "PROMO ORPHELIN",
     "", "Casablanca", "Actif", 26, 4, "C123456789VD"],
    # 6 : store absent de la Reference -> signale, non modifie
    [2026, "Juillet", "DA", "DA", "KAM Y", "ID_INCONNU", "C321", "PROMO INCONNU",
     "Store Inconnu", "Oujda", "Actif", 26, 5, "C321DA"],
    # 7 : division RAC -> non traitee (aucune regle definie)
    [2026, "Juillet", "RAC", "RAC", "KAM Z", "RAC11111", "C888888888", "PROMO RAC",
     "Store Rac", "Fes", "Actif", 17, 6, "C888888888RAC"],
    # 8 : store en conflit dans la Reference -> non modifie
    [2026, "Juillet", "VD", "VD", "KAM W", "CONF0000", "C777777777", "PROMO INCONNU",
     "Store Conflit", "Fes", "Actif", 26, 7, "C777777777VD"],
    # 9 : Code Store ET Column1 incoherents -> rapprochement par Store + Division1
    [2026, "Juillet", "VD", "VD", "", "XX111111", "C000000000", "PROMO SOLO",
     "Store Solo", "Fes", "Actif", 26, 8, "CLE_FAUSSE"],
]

KAM_VD = [
    ["Mohammedia", "C000114373", "Aswak Assalam Mohammedia", "KAM VD1"],
    ["Rabat", "C003470765", "Aswak Assalam Hay Riad", "KAM VD2"],
    ["Fes", "C999999999", "Store Solo", "KAM VD3"],
    ["Fes", "C555555555", "Store Absent BDD", "KAM VD4"],
]
KAM_DA = [
    ["Mohammedia", "C000114373", "Aswak Assalam Mohammedia", "KAM DA1"],
    ["Rabat", "C003470765", "Aswak Assalam Hay Riad", "KAM DA2"],
]


def build_reference(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    # en-tetes volontairement en ligne 2 et colonne B (comme le vrai fichier)
    for index, header in enumerate(REF_HEADERS, start=2):
        cell = ws.cell(2, index, header)
        cell.font = Font(bold=True)
    for row_index, row in enumerate(REF_ROWS, start=3):
        for col_index, value in enumerate(row, start=2):
            ws.cell(row_index, col_index, value)
    ws.auto_filter.ref = f"B2:K{2 + len(REF_ROWS)}"
    wb.save(path)
    return path


def build_target(path: Path) -> Path:
    wb = Workbook()
    other = wb.active
    other.title = "Notes"                                   # feuille parasite : ne doit pas etre choisie
    other["A1"] = "feuille sans rapport"

    ws = wb.create_sheet("BDD PROMOTERS MONTH")
    for index, header in enumerate(TARGET_HEADERS, start=1):
        cell = ws.cell(1, index, header)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", start_color="FF1F4E79")
    for row_index, row in enumerate(TARGET_ROWS, start=2):
        for col_index, value in enumerate(row, start=1):
            cell = ws.cell(row_index, col_index, value)
            cell.font = Font(name="Calibri", size=11)
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
        for index, header in enumerate(["City", "Code store", "Store", "KAM"], start=2):
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

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("fixtures")
    for name, path in build_all(target).items():
        print(f"{name:10} -> {path}")
