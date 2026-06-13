from pathlib import Path
import sys

from .gui import BewerbungsverwaltungApp


def main() -> None:
    if getattr(sys, "frozen", False):
        root = Path(sys.executable).resolve().parent
    else:
        root = Path(__file__).resolve().parents[2]
    workbook_path = root / "data" / "Bewerbungsaktivitäten mit Erinnerungen.xlsx"
    app = BewerbungsverwaltungApp(workbook_path=workbook_path)
    app.run()
