"""Tests des emplacements des fichiers du logiciel (config, journaux)."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app_paths                                                   # noqa: E402
import main                                                        # noqa: E402


class AppPathsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.home = Path(self.tmp.name)
        env = {"APPDATA": str(self.home / "Roaming"), "LOCALAPPDATA": str(self.home / "Local")}
        self.patches = [
            mock.patch.object(app_paths.sys, "platform", "win32"),
            mock.patch.dict(os.environ, env),
            mock.patch.object(app_paths, "LEGACY_CONFIG", self.home / ".aya_excel.json"),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self) -> None:
        for p in reversed(self.patches):
            p.stop()
        self.tmp.cleanup()

    def test_dossiers_windows_dans_le_profil(self):
        self.assertEqual(app_paths.config_file(), self.home / "Roaming" / "AYA Excel" / "config.json")
        self.assertEqual(app_paths.log_dir(), self.home / "Local" / "AYA Excel" / "logs")

    def test_config_ecrite_puis_relue(self):
        self.assertIsNone(app_paths.read_config_text())
        app_paths.write_config_text('{"target": "BDD.xlsx"}')
        self.assertEqual(app_paths.read_config_text(), '{"target": "BDD.xlsx"}')
        self.assertFalse(app_paths.config_file().with_suffix(".tmp").exists())

    def test_ancienne_config_reprise(self):
        app_paths.LEGACY_CONFIG.write_text('{"kam": "KAM.xlsx"}', encoding="utf-8")
        self.assertEqual(app_paths.read_config_text(), '{"kam": "KAM.xlsx"}')
        app_paths.write_config_text('{"kam": "NOUVEAU.xlsx"}')
        self.assertEqual(app_paths.read_config_text(), '{"kam": "NOUVEAU.xlsx"}')


class SansConsoleTest(unittest.TestCase):
    def test_print_ne_plante_pas_sans_console(self):
        """EXE fenetre : stdout/stderr valent None -> main les remplace."""
        with mock.patch.object(sys, "stdout", None), mock.patch.object(sys, "stderr", None):
            main._fix_missing_console()
            print("ok")
            print("ok", file=sys.stderr)
            self.assertIsNotNone(sys.stdout)


if __name__ == "__main__":
    unittest.main()
