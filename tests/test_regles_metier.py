"""
tests/test_regles_metier.py
---------------------------
Un test par regle, et les 10 cas obligatoires demandes.

    python -m unittest discover -s tests -v
"""

from __future__ import annotations

import hashlib
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import openpyxl                                                    # noqa: E402

from logger import (                                               # noqa: E402
    A_VERIFIER, CONFORME, CREEE, IGNOREE, MODIFIEE, ROUGE, STORE_NON_TROUVE,
)
from make_fixtures import build_all                                # noqa: E402
from matcher import build_key, norm_store, split_divisions         # noqa: E402
from pipeline import PipelineError, Selection, Session             # noqa: E402
from processor import Options                                      # noqa: E402

SHEET = "BDD PROMOTERS MONTH"
COL = {name: index for index, name in enumerate(
    ["Annee", "Mois", "Division", "Division1", "KAM", "ID Promoter", "Code Store",
     "Promoter", "Store", "City", "STATUT", "DAY", "IDAYA", "Column1"], start=1)}

# Lignes du fichier 2 de test (voir tests/make_fixtures.py)
L_VD_EXISTANT = 2
L_DA_EXISTANT = 3
L_VDDA_VD = 4
L_STORE_VIDE = 5
L_STORE_INCONNU = 6
L_RAC = 7
L_CONFLIT = 8
L_COLUMN1_FAUX = 9
L_SANS_REFERENCE = 10
L_SANS_KAM = 11
L_ALIAS_A = 12
L_ALIAS_B = 13
L_CREEE = 14


def md5(path: Path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


class BaseCase(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp(prefix="aya_test_"))
        self.files = build_all(self.folder)
        self.output = self.folder / "RESULTAT.xlsx"
        self.hashes = {name: md5(path) for name, path in self.files.items()}
        self.plan = None

    def tearDown(self) -> None:
        shutil.rmtree(self.folder, ignore_errors=True)

    # -- helpers -------------------------------------------------------
    def run_pipeline(self, options: Options | None = None, output: Path | None = None,
                     target: Path | None = None):
        session = Session(
            Selection(reference=self.files["reference"], target=target or self.files["target"],
                      kam=self.files["kam"]),
            options or Options(),
        )
        try:
            session.load()
            plan = session.analyze()
            result, _ = session.apply(output or self.output, write_report=False)
            self.plan = plan
            return plan, result
        finally:
            session.close()

    def sheet(self, path: Path | None = None):
        self.wb = openpyxl.load_workbook(path or self.output)
        return self.wb[SHEET]

    def cell(self, ws, row: int, column: str):
        return ws.cell(row, COL[column]).value

    def decision_of(self, plan, row: int) -> str:
        found = [d.decision for d in plan.logger.decisions if d.row == row]
        self.assertEqual(len(found), 1, f"la ligne {row} doit avoir exactement 1 decision, trouve {found}")
        return found[0]

    def reason_of(self, plan, row: int) -> str:
        return next(d.reason + " " + d.details for d in plan.logger.decisions if d.row == row)


class TestCasObligatoires(BaseCase):
    """Les 10 cas exiges avant de considerer le projet termine."""

    def test_01_store_vide(self):
        plan, result = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_STORE_VIDE), ROUGE)
        self.assertEqual(plan.stats.lignes_store_vide, 1)
        self.assertEqual(result.colored_rows, 1)
        for column in COL:
            self.assertEqual(ws.cell(L_STORE_VIDE, COL[column]).fill.start_color.rgb, "FFFF0000",
                             f"colonne {column} non coloree")
        # aucune valeur touchee
        self.assertEqual(self.cell(ws, L_STORE_VIDE, "ID Promoter"), "ID_ORPHELIN")
        self.assertEqual(self.cell(ws, L_STORE_VIDE, "Promoter"), "PROMO ORPHELIN")
        self.assertEqual(self.cell(ws, L_STORE_VIDE, "KAM"), "KAM X")
        self.assertIsNone(self.cell(ws, L_STORE_VIDE, "Store"))

    def test_02_store_non_trouve(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_STORE_INCONNU), STORE_NON_TROUVE)
        self.assertEqual(plan.stats.stores_non_trouves, 1)
        self.assertEqual(self.cell(ws, L_STORE_INCONNU, "ID Promoter"), "ID_INCONNU")
        self.assertEqual(self.cell(ws, L_STORE_INCONNU, "KAM"), "KAM Y")

    def test_03_store_vd_existant(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_VD_EXISTANT), MODIFIEE)
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "ID Promoter"), "IDVD123")
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "Promoter"), "PROMO VD KENITRA")
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "DAY"), 26)
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "City"), "Kenitra")
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "Division"), "VD")
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "Division1"), "VD")
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "Column1"), "C123VD")

    def test_04_store_da_existant(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_DA_EXISTANT), CONFORME)
        self.assertEqual(self.cell(ws, L_DA_EXISTANT, "Promoter"), "PROMO DA KENITRA")
        self.assertEqual(self.cell(ws, L_DA_EXISTANT, "DAY"), 24)
        self.assertEqual(self.cell(ws, L_DA_EXISTANT, "Column1"), "C123DA")

    def test_05_vd_existant_da_absent(self):
        plan, result = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(plan.stats.lignes_ajoutees, 1)
        self.assertEqual(result.created_rows, 1)
        self.assertEqual(ws.max_row, L_CREEE)
        self.assertEqual(self.cell(ws, L_CREEE, "Column1"), "C200DA")
        self.assertEqual(self.cell(ws, L_CREEE, "Division1"), "DA")
        self.assertEqual(self.cell(ws, L_CREEE, "ID Promoter"), "IDVDDA200")
        self.assertEqual(self.cell(ws, L_CREEE, "DAY"), 26)
        self.assertEqual(self.cell(ws, L_CREEE, "City"), "Casablanca")
        # la ligne VD existante a ete mise a jour, pas dupliquee
        self.assertEqual(self.decision_of(plan, L_VDDA_VD), MODIFIEE)

    def test_06_store_vd_plus_da(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        cles = [self.cell(ws, row, "Column1") for row in range(2, ws.max_row + 1)]
        self.assertEqual(cles.count("C200VD"), 1)
        self.assertEqual(cles.count("C200DA"), 1)
        creees = [d.key for d in plan.logger.by_decision(CREEE)]
        self.assertEqual(creees, ["C200DA"])

    def test_07_meme_store_meme_division_deux_promoteurs(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_CONFLIT), A_VERIFIER)
        self.assertIn("2 promoteurs differents", self.reason_of(plan, L_CONFLIT))
        # donnees intactes : aucun des deux promoteurs n'a ete choisi
        self.assertEqual(self.cell(ws, L_CONFLIT, "ID Promoter"), "ANCIEN_CONF")
        self.assertEqual(self.cell(ws, L_CONFLIT, "Promoter"), "ANCIEN PROMO CONF")
        self.assertEqual(self.cell(ws, L_CONFLIT, "KAM"), "KAM W")

    def test_08_kam_different_entre_vd_et_da(self):
        _, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "KAM"), "KAM VD KENITRA")
        self.assertEqual(self.cell(ws, L_DA_EXISTANT, "KAM"), "KAM DA KENITRA")
        self.assertEqual(self.cell(ws, L_VDDA_VD, "KAM"), "KAM VD VDDA")
        self.assertEqual(self.cell(ws, L_CREEE, "KAM"), "KAM DA VDDA")

    def test_09_column1_incorrect_corrige(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_COLUMN1_FAUX), MODIFIEE)
        self.assertEqual(self.cell(ws, L_COLUMN1_FAUX, "Column1"), "C300VD")
        self.assertEqual(self.cell(ws, L_COLUMN1_FAUX, "Code Store"), "C300")
        # aucune ligne creee a cause de la cle fausse
        self.assertEqual(plan.stats.lignes_ajoutees, 1)

    def test_10_deuxieme_execution_sans_duplicate(self):
        premier, _ = self.run_pipeline()
        lignes_apres_1 = self.sheet().max_row
        second_output = self.folder / "RESULTAT2.xlsx"
        plan, _ = self.run_pipeline(target=self.output, output=second_output)
        ws = self.sheet(second_output)
        self.assertEqual(plan.stats.lignes_modifiees, 0)
        self.assertEqual(plan.stats.lignes_ajoutees, 0)
        self.assertEqual(ws.max_row, lignes_apres_1)
        cles = [self.cell(ws, row, "Column1") for row in range(2, ws.max_row + 1)]
        self.assertEqual(len(cles), len(set(cles)) + 1)   # seul le doublon deja present reste
        self.assertEqual(premier.stats.lignes_analysees + 1, plan.stats.lignes_analysees)


class TestAutresRegles(BaseCase):
    def test_rac_jamais_traite_comme_vd_ou_da(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_RAC), A_VERIFIER)
        self.assertIn("RAC", self.reason_of(plan, L_RAC))
        self.assertEqual(self.cell(ws, L_RAC, "Division1"), "RAC")
        self.assertEqual(self.cell(ws, L_RAC, "ID Promoter"), "IDRAC400")
        self.assertEqual(self.cell(ws, L_RAC, "KAM"), "KAM Z")

    def test_kam_non_trouve_laisse_vide(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(plan.stats.kam_non_trouves, 1)
        self.assertIsNone(self.cell(ws, L_SANS_KAM, "KAM"))

    def test_ligne_sans_correspondance_ignoree(self):
        plan, _ = self.run_pipeline()
        ws = self.sheet()
        self.assertEqual(self.decision_of(plan, L_SANS_REFERENCE), IGNOREE)
        self.assertEqual(self.cell(ws, L_SANS_REFERENCE, "KAM"), "KAM ??")

    def test_duplicate_compte_uniquement_les_creations_annulees(self):
        plan, _ = self.run_pipeline()
        self.assertEqual(plan.stats.duplicates_evites, 1)
        self.assertLess(plan.stats.duplicates_evites, plan.stats.lignes_analysees)
        self.assertEqual(self.decision_of(plan, L_ALIAS_A), MODIFIEE)
        self.assertEqual(self.decision_of(plan, L_ALIAS_B), A_VERIFIER)

    def test_colonnes_jamais_touchees(self):
        _, _ = self.run_pipeline()
        ws = self.sheet()
        for row, idaya in ((L_VD_EXISTANT, 1), (L_VDDA_VD, 3), (L_COLUMN1_FAUX, 8)):
            self.assertEqual(self.cell(ws, row, "Annee"), 2026)
            self.assertEqual(self.cell(ws, row, "Mois"), "Juillet")
            self.assertEqual(self.cell(ws, row, "STATUT"), "Actif")
            self.assertEqual(self.cell(ws, row, "IDAYA"), idaya)

    def test_statistiques_exactes(self):
        plan, _ = self.run_pipeline()
        attendu = {
            "Lignes analysees": 12, "Lignes modifiees": 5, "Lignes deja conformes": 1,
            "Lignes ajoutees": 1, "Lignes Store vide": 1, "Stores non trouves": 1,
            "KAM trouves": 6, "KAM non trouves": 1, "Duplicates evites": 1, "A verifier": 4,
        }
        self.assertEqual(dict(plan.stats.as_pairs()), attendu)

    def test_rapport_explique_chaque_ligne(self):
        plan, _ = self.run_pipeline()
        lignes = {d.row for d in plan.logger.decisions if d.row}
        self.assertEqual(lignes, set(range(2, 14)))          # les 12 lignes du fichier 2
        for decision in plan.logger.decisions:
            self.assertTrue(decision.reason.strip(), "chaque decision doit avoir une raison")
        self.assertEqual(len(plan.logger.by_decision(CREEE)), 1)

    def test_option_idaya_uniquement_sur_les_creations(self):
        _, _ = self.run_pipeline(Options(auto_number_idaya=True))
        ws = self.sheet()
        self.assertEqual(self.cell(ws, L_CREEE, "IDAYA"), 13)
        self.assertEqual(self.cell(ws, L_VD_EXISTANT, "IDAYA"), 1)


class TestFichiers(BaseCase):
    def test_reference_feuille_feuil2(self):
        session = Session(Selection(reference=self.files["reference"], target=self.files["target"],
                                    kam=self.files["kam"]))
        try:
            data = session.load()
            self.assertEqual(data.reference.layout.title, "Feuil2")
            self.assertEqual(data.reference.layout.header_row, 1)
            self.assertEqual(data.reference.layout.col("Divisions"), 3)
            self.assertEqual(data.reference.layout.col("Working Days"), 12)
            self.assertEqual(data.target.layout.title, SHEET)
            self.assertEqual(data.kam.divisions, ["DA", "VD"])
        finally:
            session.close()

    def test_repli_par_position_si_entetes_illisibles(self):
        wb = openpyxl.load_workbook(self.files["reference"])
        ws = wb["Feuil2"]
        for column in (3, 6, 7, 8, 9, 10, 12):
            ws.cell(1, column, f"colonne {column}")          # en-tetes rendus illisibles
        cassee = self.folder / "REF_SANS_ENTETES.xlsx"
        wb.save(cassee)
        self.files["reference"] = cassee
        plan, _ = self.run_pipeline()
        ws2 = self.sheet()
        self.assertEqual(self.cell(ws2, L_VD_EXISTANT, "Promoter"), "PROMO VD KENITRA")
        self.assertEqual(plan.stats.lignes_ajoutees, 1)

    def test_sources_jamais_modifiees(self):
        self.run_pipeline()
        for name, path in self.files.items():
            self.assertEqual(md5(path), self.hashes[name], f"{name} a ete modifie")

    def test_sortie_ne_peut_pas_ecraser_une_source(self):
        with self.assertRaises(PipelineError):
            self.run_pipeline(output=Path(self.files["target"]))

    def test_mise_en_forme_conservee(self):
        self.run_pipeline()
        ws = self.sheet()
        self.assertIn("Notes", self.wb.sheetnames)
        self.assertEqual(ws.freeze_panes, "A2")
        self.assertEqual(list(ws.tables.values())[0].ref, f"A1:N{L_CREEE}")
        self.assertTrue(ws.cell(1, 1).font.b)
        self.assertEqual(ws.cell(L_CREEE, 1).font.name, ws.cell(L_CREEE - 1, 1).font.name)

    def test_rapport_genere(self):
        session = Session(Selection(reference=self.files["reference"], target=self.files["target"],
                                    kam=self.files["kam"]))
        try:
            session.load()
            session.analyze()
            _, reports = session.apply(self.output, write_report=True)
        finally:
            session.close()
        self.assertEqual(len(reports), 2)
        texte = Path(reports[0]).read_text(encoding="utf-8")
        for mot in ("ROUGE", "MODIFIEE", "CREEE", "CONFORME", "A VERIFIER", "IGNOREE", "STORE NON TROUVE"):
            self.assertIn(mot, texte)


class TestMatcher(unittest.TestCase):
    def test_normalisation(self):
        self.assertEqual(norm_store("  Aswak   Assalam  Mohammédia "), "aswak assalam mohammedia")
        self.assertEqual(build_key(" c003470765 ", "vd"), "C003470765VD")
        self.assertNotEqual(build_key("C003470765", "VD"), build_key("C003470765", "DA"))
        self.assertEqual(split_divisions("VD+DA"), ["VD", "DA"])
        self.assertEqual(split_divisions("RAC"), ["RAC"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
