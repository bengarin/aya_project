"""
logger.py
---------
Le rapport de traitement.

Regle imposee : le rapport doit dire, pour CHAQUE ligne du fichier 2,
POURQUOI elle a ete modifiee / creee / laissee conforme / mise en rouge /
mise "A VERIFIER" / ignoree.

Le logiciel enregistre donc UNE decision par ligne, avec sa raison,
plus des messages techniques a part.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# Les 7 decisions possibles pour une ligne (rien d'autre n'est autorise)
# ---------------------------------------------------------------------------
ROUGE = "ROUGE"                        # Store vide -> coloree en rouge, aucun traitement
MODIFIEE = "MODIFIEE"                  # ligne existante mise a jour depuis la Reference
CONFORME = "CONFORME"                  # ligne existante deja identique a la Reference
CREEE = "CREEE"                        # nouvelle ligne ajoutee (cle absente du fichier 2)
STORE_NON_TROUVE = "STORE NON TROUVE"  # Store rempli mais absent de la Reference
A_VERIFIER = "A VERIFIER"              # cas ambigu : le logiciel ne decide pas
IGNOREE = "IGNOREE"                    # aucune affectation correspondante dans la Reference

DECISIONS = (ROUGE, A_VERIFIER, STORE_NON_TROUVE, CREEE, MODIFIEE, IGNOREE, CONFORME)

# Messages techniques (hors decisions de ligne)
INFO = "INFO"
KAM_NON_TROUVE = "KAM NON TROUVE"
DUPLICATE = "DUPLICATE EVITE"
ERREUR = "ERREUR"

NIVEAUX = DECISIONS + (KAM_NON_TROUVE, DUPLICATE, INFO, ERREUR)


@dataclass
class Decision:
    """Ce qui est arrive a UNE ligne, et pourquoi."""

    decision: str
    reason: str                        # POURQUOI (phrase claire)
    row: int | None = None             # ligne dans le fichier 2 ('-' pour une ligne creee)
    store: str = ""
    division: str = ""
    key: str = ""                      # cle Code Store + Division
    details: str = ""                  # cellules changees, valeurs sources...

    def as_text(self) -> str:
        where = f"ligne {self.row}" if self.row else "nouvelle"
        who = f"{self.store}" + (f" / {self.division}" if self.division else "")
        text = f"{self.decision:<16} | {where:<12} | {who:<34} | {self.reason}"
        if self.details:
            text += f" || {self.details}"
        return text


@dataclass
class LogEntry:
    """Message technique (KAM manquant, duplicate evite, info d'ecriture...)."""

    level: str
    message: str
    row: int | None = None
    store: str = ""
    division: str = ""

    def as_text(self) -> str:
        where = f"ligne {self.row}" if self.row else "-"
        who = f"{self.store}" + (f" / {self.division}" if self.division else "")
        return f"{self.level:<16} | {where:<12} | {who:<34} | {self.message}"


@dataclass
class ProcessLogger:
    decisions: list[Decision] = field(default_factory=list)
    entries: list[LogEntry] = field(default_factory=list)
    started_at: _dt.datetime = field(default_factory=_dt.datetime.now)

    # -- ecriture -----------------------------------------------------------
    def decide(self, decision: str, reason: str, **kw) -> Decision:
        record = Decision(decision=decision, reason=reason, **kw)
        self.decisions.append(record)
        return record

    def log(self, level: str, message: str, row: int | None = None, store: str = "", division: str = "") -> LogEntry:
        entry = LogEntry(level=level, message=message, row=row, store=store, division=division)
        self.entries.append(entry)
        return entry

    def info(self, message: str, **kw) -> LogEntry:
        return self.log(INFO, message, **kw)

    def error(self, message: str, **kw) -> LogEntry:
        return self.log(ERREUR, message, **kw)

    # -- lecture ------------------------------------------------------------
    def by_decision(self, decision: str) -> list[Decision]:
        return [d for d in self.decisions if d.decision == decision]

    def counts(self) -> dict[str, int]:
        result: dict[str, int] = {}
        for record in self.decisions:
            result[record.decision] = result.get(record.decision, 0) + 1
        for entry in self.entries:
            if entry.level in (KAM_NON_TROUVE, DUPLICATE, ERREUR):
                result[entry.level] = result.get(entry.level, 0) + 1
        return result

    def display_rows(self) -> list[tuple[str, str]]:
        """Lignes affichables dans l'interface : (niveau, texte)."""
        rank = {name: index for index, name in enumerate(DECISIONS)}
        ordered = sorted(self.decisions, key=lambda d: (rank.get(d.decision, 99), d.row or 10**9))
        rows = [(d.decision, d.as_text()) for d in ordered]
        rows += [(e.level, e.as_text()) for e in self.entries if e.level != INFO]
        return rows

    def as_text(self, header: str = "") -> str:
        lines = []
        if header:
            lines.append(header)
        lines.append(f"Rapport de traitement - {self.started_at:%d/%m/%Y %H:%M:%S}")
        lines.append("=" * 120)
        counts = self.counts()
        for level in NIVEAUX:
            if level in counts:
                lines.append(f"{level:<18} : {counts[level]}")
        lines.append("=" * 120)
        lines.append("DECISION PAR LIGNE (pourquoi chaque ligne a ete traitee de cette facon)")
        lines.append("-" * 120)
        for decision, text in self.display_rows():
            if decision in DECISIONS:
                lines.append(text)
        lines.append("")
        lines.append("MESSAGES TECHNIQUES")
        lines.append("-" * 120)
        for entry in self.entries:
            lines.append(entry.as_text())
        return "\n".join(lines)

    # -- export -------------------------------------------------------------
    def save_text(self, path: str | Path, header: str = "") -> Path:
        path = Path(path)
        path.write_text(self.as_text(header), encoding="utf-8")
        return path

    def save_xlsx(self, path: str | Path, header: str = "") -> Path:
        """Rapport Excel : une ligne = une decision, filtrable."""
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill

        path = Path(path)
        wb = Workbook()
        ws = wb.active
        ws.title = "Decisions"
        ws.append(["Decision", "Ligne fichier 2", "Store", "Division", "Cle (Code Store + Division)",
                   "Pourquoi", "Details"])
        for record in sorted(self.decisions, key=lambda d: (d.row or 10**9, d.division)):
            ws.append([record.decision, record.row, record.store, record.division,
                       record.key, record.reason, record.details])

        ws2 = wb.create_sheet("Messages techniques")
        ws2.append(["Niveau", "Ligne", "Store", "Division", "Message"])
        for entry in self.entries:
            ws2.append([entry.level, entry.row, entry.store, entry.division, entry.message])

        for sheet, widths in ((ws, (18, 14, 32, 10, 26, 70, 90)), (ws2, (18, 10, 32, 10, 110))):
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFFFF")
                cell.fill = PatternFill("solid", start_color="FF1F4E79")
                cell.alignment = Alignment(horizontal="center")
            for index, width in enumerate(widths, start=1):
                sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
            sheet.auto_filter.ref = f"A1:{sheet.cell(1, len(widths)).column_letter}{max(sheet.max_row, 1)}"
            sheet.freeze_panes = "A2"

        if header:
            ws3 = wb.create_sheet("Contexte")
            for index, line in enumerate(header.splitlines(), start=1):
                ws3.cell(index, 1, line)
            ws3.column_dimensions["A"].width = 130
        wb.save(path)
        return path
