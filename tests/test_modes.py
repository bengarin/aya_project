"""
tests/test_modes.py
-------------------
Seul le fichier a traiter (BDD) est obligatoire. Reference, KAM et dossier des IDs
sont optionnels : seules les colonnes qu'ils alimentent changent.

Sans Reference il n'y a aucune creation : les numeros de ligne du resultat sont
ceux du fichier d'origine (tests/make_fixtures.py) :
  2 Kenitra VD | 3 Kenitra DA | 4 VDDA VD | 5 Store vide | 6 Store Inconnu DA | 7 RAC
  8 Conflit VD | 9 Solo VD | 10 Solo DA | 11 NoKam VD | 12 Alias A VD | 13 Alias B DA
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

from logger import A_VERIFIER, CONFORME, MODIFIEE, ROUGE           # noqa: E402
from make_fixtures import build_all                                # noqa: E402
from pipeline import PipelineError, Selection, Session             # noqa: E402
from test_regles_metier import COL, L_CREEE, SHEET                 # noqa: E402

DERNIERE_LIGNE_SOURCE = 13


class ModesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp(prefix="aya_modes_"))
        self.files = build_all(self.folder)
        self.ids = self.folder / "ids"
        self.ids.mkdir()
        for name in ("101-ASWAK ASSALAM KENITRA VD.pdf", "200-Store VDDA VD+DA.jpg"):
            (self.ids / name).write_bytes(b"")
        self.output = self.folder / "RESULTAT.xlsx"

    def tearDown(self) -> None:
        shutil.rmtree(self.folder, ignore_errors=True)

    def run_mode(self, reference=False, kam=False, ids=False, target=None):
        selection = Selection(
            target=target or self.files["target"],
            reference=self.files["reference"] if reference else None,
            kam=self.files["kam"] if kam else None,
            ids_folder=self.ids if ids else None,
        )
        session = Session(selection)
        try:
            session.load()
            plan = session.analyze()
            session.apply(self.output, write_report=False)
            return plan, selection
        finally:
            session.close()

    def grid(self, path: Path) -> dict[tuple[int, str], object]:
        wb = openpyxl.load_workbook(path)
        try:
            ws = wb[SHEET]
            return {(r, name): ws.cell(r, c).value
                    for r in range(1, ws.max_row + 1) for name, c in COL.items()}
        finally:
            wb.close()

    def changed_columns(self) -> set[str]:
        avant, apres = self.grid(self.files["target"]), self.grid(self.output)
        self.assertEqual(set(avant), set(apres), "aucune ligne ne doit etre creee")
        return {col for (row, col), value in avant.items() if apres[(row, col)] != value}

    def decision(self, plan, row: int) -> str:
        return next(d.decision for d in plan.logger.decisions if d.row == row)

    # ------------------------------------------------------------------
    def test_bdd_seule_refusee(self):
        with self.assertRaises(PipelineError):
            self.run_mode()

    def test_bdd_obligatoire(self):
        with self.assertRaises(PipelineError):
            Selection(reference=self.files["reference"], kam=self.files["kam"]).validate()

    def test_bdd_plus_kam_seulement_colonne_kam(self):
        plan, selection = self.run_mode(kam=True)
        self.assertEqual(self.changed_columns(), {"KAM"})
        g = self.grid(self.output)
        self.assertEqual(g[(2, "KAM")], "KAM VD KENITRA")         # Store + Division1 de la ligne
        self.assertEqual(g[(2, "Promoter")], "ANCIEN PROMO")      # pas de Reference : intact
        self.assertEqual(g[(6, "KAM")], "#")                      # absent du fichier KAM
        self.assertEqual(g[(7, "KAM")], "KAM Z")                  # RAC : jamais traite
        self.assertEqual(self.decision(plan, 3), CONFORME)
        self.assertEqual(self.decision(plan, 5), ROUGE)
        self.assertEqual(self.decision(plan, 7), A_VERIFIER)
        self.assertEqual(plan.stats.lignes_ajoutees, 0)
        self.assertNotIn("Stores non trouves", dict(plan.stats.as_pairs()))
        self.assertIn("Non touche : ID", selection.mode())

    def test_bdd_plus_ids_seulement_colonne_id(self):
        plan, _ = self.run_mode(ids=True)
        self.assertEqual(self.changed_columns(), {"IDAYA"})
        g = self.grid(self.output)
        self.assertEqual(g[(2, "IDAYA")], 101)
        self.assertEqual(g[(4, "IDAYA")], "200VD")
        self.assertEqual(g[(3, "IDAYA")], 2)                      # pas de fichier : conserve
        self.assertEqual(self.decision(plan, 2), MODIFIEE)
        self.assertNotIn("KAM trouves", dict(plan.stats.as_pairs()))

    def test_bdd_plus_kam_plus_ids(self):
        self.run_mode(kam=True, ids=True)
        self.assertEqual(self.changed_columns(), {"KAM", "IDAYA"})

    def test_bdd_plus_reference_kam_jamais_touche(self):
        plan, selection = self.run_mode(reference=True)
        g = self.grid(self.output)
        self.assertEqual(g[(2, "KAM")], "ANCIEN KAM")             # pas de fichier KAM
        self.assertEqual(g[(2, "Promoter")], "PROMO VD KENITRA")  # la Reference s'applique
        self.assertEqual(g[(L_CREEE, "KAM")], "#")                # jamais le KAM de la ligne soeur
        self.assertEqual(plan.stats.lignes_ajoutees, 1)
        self.assertIn("Non touche : KAM, ID", selection.mode())

    def test_sans_reference_doublon_jamais_choisi(self):
        wb = openpyxl.load_workbook(self.files["target"])
        ws = wb[SHEET]
        ws.append([ws.cell(9, c).value for c in range(1, 15)])   # 2e ligne Store Solo + VD
        doublon = self.folder / "BDD_DOUBLON.xlsx"
        wb.save(doublon)
        plan, _ = self.run_mode(kam=True, target=doublon)
        self.assertEqual(self.decision(plan, 9), A_VERIFIER)
        self.assertEqual(self.decision(plan, DERNIERE_LIGNE_SOURCE + 1), A_VERIFIER)


if __name__ == "__main__":
    unittest.main()
