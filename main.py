"""
main.py
-------
Point d'entree du logiciel.

  python main.py                -> ouvre l'interface graphique
  python main.py --cli ...      -> traitement sans interface (tests, automatisation)

Le mode --cli fait EXACTEMENT le meme traitement que l'interface.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Permet de lancer le logiciel depuis n'importe quel dossier.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pipeline import PipelineError, Selection, Session       # noqa: E402
from processor import Options                                # noqa: E402

APP_NAME = "Automatisation Excel AYA"
VERSION = "1.0.0"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="aya-excel",
        description=f"{APP_NAME} - traitement BDD commissions (Reference + BDD + Affectation KAM)",
    )
    parser.add_argument("--cli", action="store_true", help="traitement en ligne de commande (sans fenetre)")
    parser.add_argument("-r", "--reference", help="Fichier 1 : Reference (lecture seule)")
    parser.add_argument("-t", "--target", help="Fichier 2 : BDD a traiter")
    parser.add_argument("-k", "--kam", help="Fichier 3 : Affectation des KAM")
    parser.add_argument("-o", "--output", help="Fichier resultat (cree, jamais ecrase sur les sources)")
    parser.add_argument("--reference-sheet", help="forcer la feuille du fichier 1")
    parser.add_argument("--target-sheet", help="forcer la feuille du fichier 2")
    parser.add_argument("--kam-sheets", nargs="*", help="forcer les feuilles du fichier 3")
    parser.add_argument("--preview", action="store_true", help="analyser seulement, sans generer le fichier")
    parser.add_argument("--no-report", action="store_true", help="ne pas generer le rapport a cote du resultat")
    parser.add_argument("--auto-idaya", action="store_true",
                        help="numeroter automatiquement IDAYA sur les lignes CREEES uniquement")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {VERSION}")
    return parser


def options_from_args(args) -> Options:
    return Options(auto_number_idaya=args.auto_idaya)


def run_cli(args) -> int:
    selection = Selection(
        reference=args.reference, target=args.target, kam=args.kam,
        reference_sheet=args.reference_sheet, target_sheet=args.target_sheet,
        kam_sheets=args.kam_sheets,
    )
    session = Session(selection, options_from_args(args))

    def progress(message: str, value: float) -> None:
        print(f"[{value * 100:5.1f}%] {message}")

    try:
        data = session.load(progress)
        print("\n--- Colonnes detectees ---")
        print(data.describe())
        for warning in data.warnings:
            print(f"[ATTENTION] {warning}")

        plan = session.analyze(progress)
        print("\n--- Apercu (Preview) ---")
        for label, value in plan.stats.as_pairs():
            print(f"{label:<32}: {value}")

        if args.preview:
            print("\nMode preview : aucun fichier ecrit.")
            print(plan.logger.as_text())
            return 0

        output = Path(args.output) if args.output else session.default_output()
        result, reports = session.apply(output, write_report=not args.no_report, progress=progress)
        print(f"\nFichier resultat : {result.output_path}")
        print(f"  cellules modifiees : {result.updated_cells}")
        print(f"  lignes creees      : {result.created_rows}")
        print(f"  lignes colorees    : {result.colored_rows}")
        for warning in result.warnings:
            print(f"[ATTENTION] {warning}")
        for report in reports:
            print(f"Rapport : {report}")
        return 0
    except PipelineError as exc:
        print(f"\nERREUR : {exc}", file=sys.stderr)
        return 2
    except Exception as exc:                                # pragma: no cover
        print(f"\nERREUR INATTENDUE : {exc}", file=sys.stderr)
        return 3
    finally:
        session.close()


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.cli or any([args.reference, args.target, args.kam]):
        return run_cli(args)
    try:
        from gui import launch
    except Exception as exc:                                # pragma: no cover - environnement sans Tkinter
        print(
            "Impossible de demarrer l'interface graphique : " + str(exc) + "\n\n"
            "Installez Tkinter (Windows/Mac : inclus avec Python ; Linux : 'sudo apt install python3-tk')\n"
            "ou utilisez le mode ligne de commande :\n"
            "  python main.py --cli -r Reference.xlsx -t BDD.xlsx -k KAM.xlsx -o Resultat.xlsx",
            file=sys.stderr,
        )
        return 4
    launch()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
