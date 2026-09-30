"""
tests/test_ids.py
-----------------
IDAYA (colonne M) lu dans les noms de fichiers d'un dossier ("Etat de vente").

    35-Ashiama panoramique DA.jpeg   -> 35       (fichier pour UNE division)
    39-Carrefour Ain Sebaa VD+DA.jpg -> 39DA / 39VD
    40-CARREFOUR BENI MELLAL.jpg     -> 40DA / 40VD (pas de division = tout le Store)
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import openpyxl                                                    # noqa: E402

from id_service import (                                           # noqa: E402
    AMBIGUOUS, APPROX, FOUND, NOT_FOUND, IdService, parse_filename,
)
from logger import ID_APPROCHE, ID_NON_TROUVE, MODIFIEE            # noqa: E402
from make_fixtures import build_all                                # noqa: E402
from pipeline import PipelineError, Selection, Session             # noqa: E402
from test_regles_metier import (                                   # noqa: E402
    COL, L_CREEE, L_DA_EXISTANT, L_STORE_INCONNU, L_VD_EXISTANT, L_VDDA_VD,
    L_COLUMN1_FAUX, L_SANS_KAM, SHEET,
)


def service(*names: str) -> IdService:
    return IdService([parse_filename(n) for n in names])


class NomDeFichierTest(unittest.TestCase):
    def test_formats_reels(self):
        cas = {
            "35-Ashiama panoramique DA.jpeg": ("35", "Ashiama panoramique", {"DA"}),
            "39-Carrefour Ain Sebaa VD+DA.jpg": ("39", "Carrefour Ain Sebaa", {"VD", "DA"}),
            "42-ELECTRO SERGHINI MEKNES DA+VD.pdf": ("42", "ELECTRO SERGHINI MEKNES", {"VD", "DA"}),
            "40-CARREFOUR BENI MELLAL.jpg": ("40", "CARREFOUR BENI MELLAL", set()),
            "007 - Store Solo vd.pdf": ("7", "Store Solo", {"VD"}),
        }
        for nom, (numero, store, divisions) in cas.items():
            f = parse_filename(nom)
            self.assertEqual((f.number, f.store, set(f.divisions)), (numero, store, divisions), nom)

    def test_nom_sans_numero_ignore(self):
        self.assertIsNone(parse_filename("Etat de vente.xlsx"))
        self.assertIsNone(parse_filename("12-.pdf"))

    def test_da_dans_le_nom_du_store_n_est_pas_une_division(self):
        f = parse_filename("5-Dar Bouazza.pdf")
        self.assertEqual((f.store, set(f.divisions)), ("Dar Bouazza", set()))


class RegleIdTest(unittest.TestCase):
    def test_fichier_une_division_numero_seul(self):
        ids = service("35-Ashiama panoramique DA.jpeg", "36-Ashiama panoramique VD.jpeg")
        self.assertEqual(ids.lookup(["Ashiama Panoramique"], "DA", {"VD", "DA"}).value, "35")
        self.assertEqual(ids.lookup(["Ashiama Panoramique"], "VD", {"VD", "DA"}).value, "36")

    def test_fichier_vd_da_numero_plus_division(self):
        ids = service("39-Carrefour Ain Sebaa VD+DA.jpg")
        self.assertEqual(ids.lookup(["Carrefour Ain Sebaa"], "DA", {"VD", "DA"}).value, "39DA")
        self.assertEqual(ids.lookup(["Carrefour Ain Sebaa"], "VD", {"VD", "DA"}).value, "39VD")

    def test_fichier_sans_division(self):
        ids = service("40-CARREFOUR BENI MELLAL.jpg")
        self.assertEqual(ids.lookup(["Carrefour Beni Mellal"], "VD", {"VD", "DA"}).value, "40VD")
        self.assertEqual(ids.lookup(["Carrefour Beni Mellal"], "VD", {"VD"}).value, "40")

    def test_nom_ecrit_autrement_meme_store(self):
        ids = service("43-ELECTROBOUSFIHA MEKNES DA.pdf")
        result = ids.lookup(["Electro Boussfiha Meknes"], "DA", {"DA"})
        self.assertEqual((result.status, result.value), (FOUND, "43"))

    def test_nom_approche_utilise_mais_signale(self):
        ids = service("41-Carrefour Berrachid VD+DA.jpeg", "39-Carrefour Ain Sebaa VD+DA.jpg")
        result = ids.lookup(["Carrefour Berrechid"], "DA", {"VD", "DA"})
        self.assertEqual((result.status, result.value), (APPROX, "41DA"))

    def test_store_proche_mais_different_refuse(self):
        ids = service("14-Carrefour Targa VD+DA.jpg")
        self.assertEqual(ids.lookup(["Carrefour Tanger"], "VD", {"VD", "DA"}).status, NOT_FOUND)

    def test_deux_fichiers_contradictoires(self):
        ids = service("35-Store X DA.pdf", "36-Store X DA.pdf")
        self.assertEqual(ids.lookup(["Store X"], "DA", {"DA"}).status, AMBIGUOUS)

    def test_division_absente_du_fichier(self):
        ids = service("35-Store X DA.pdf")
        self.assertEqual(ids.lookup(["Store X"], "VD", {"VD", "DA"}).status, NOT_FOUND)


class DossierIdsPipelineTest(unittest.TestCase):
    """Le dossier des IDs dans le vrai traitement (fichiers de test des 10 cas)."""

    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp(prefix="aya_ids_"))
        self.files = build_all(self.folder)
        self.ids = self.folder / "Etat de vente" / "Part2"
        self.ids.mkdir(parents=True)
        for name in ("101-ASWAK ASSALAM KENITRA VD.pdf",   # ligne VD existante (IDAYA 1 -> 101)
                     "102-aswak assalam kenitra DA.jpeg",  # ligne DA conforme  (IDAYA 2 -> 102)
                     "200-Store VDDA VD+DA.jpg",           # VD existante + DA creee
                     "300-Store Solo.pdf",                 # pas de division, Store en VD seule
                     "600-Store NoKan VD.pdf",             # nom approche de 'Store NoKam'
                     "999-Store Inconnu DA.pdf",           # Store absent de la Reference
                     "notes.txt"):                         # nom non reconnu -> ignore
            (self.ids / name).write_bytes(b"")
        self.output = self.folder / "RESULTAT.xlsx"

    def tearDown(self) -> None:
        shutil.rmtree(self.folder, ignore_errors=True)

    def run_ids(self, folder):
        session = Session(Selection(reference=self.files["reference"], target=self.files["target"],
                                    kam=self.files["kam"], ids_folder=folder))
        try:
            data = session.load()
            plan = session.analyze()
            session.apply(self.output, write_report=False)
            return data, plan
        finally:
            session.close()

    def idaya(self, row: int):
        wb = openpyxl.load_workbook(self.output)
        try:
            return wb[SHEET].cell(row=row, column=COL["IDAYA"]).value
        finally:
            wb.close()

    def test_ids_ecrits_depuis_le_dossier(self):
        _data, plan = self.run_ids(self.folder / "Etat de vente")    # sous-dossiers lus aussi
        attendus = {L_VD_EXISTANT: 101, L_DA_EXISTANT: 102, L_VDDA_VD: "200VD", L_CREEE: "200DA",
                    L_COLUMN1_FAUX: 300, L_SANS_KAM: 600}
        for row, valeur in attendus.items():
            self.assertEqual(self.idaya(row), valeur, f"ligne {row}")
        self.assertEqual(plan.stats.ids_trouves, 6)
        self.assertEqual(plan.stats.ids_approches, 1)

    def test_store_hors_reference_intact(self):
        self.run_ids(self.folder / "Etat de vente")
        self.assertEqual(self.idaya(L_STORE_INCONNU), 5)

    def test_ligne_conforme_devient_modifiee_si_id_change(self):
        _data, plan = self.run_ids(self.folder / "Etat de vente")
        decision = next(d for d in plan.logger.decisions if d.row == L_DA_EXISTANT)
        self.assertEqual(decision.decision, MODIFIEE)
        self.assertIn("IDAYA: '2' -> '102'", decision.details)

    def test_rapport_signale_approche_et_fichiers_inutilises(self):
        data, plan = self.run_ids(self.folder / "Etat de vente")
        messages = [(e.level, e.message) for e in plan.logger.entries]
        self.assertTrue(any(l == ID_APPROCHE and "600-Store NoKan VD.pdf" in m for l, m in messages))
        self.assertTrue(any(l == ID_NON_TROUVE and "999-Store Inconnu DA.pdf" in m for l, m in messages))
        self.assertTrue(any("notes.txt" in w for w in data.warnings))

    def test_sans_dossier_idaya_jamais_touchee(self):
        self.run_ids(None)
        self.assertEqual(self.idaya(L_VD_EXISTANT), 1)
        self.assertIsNone(self.idaya(L_CREEE))

    def test_dossier_inexistant(self):
        with self.assertRaises(PipelineError):
            self.run_ids(self.folder / "nexiste_pas")


if __name__ == "__main__":
    unittest.main()
