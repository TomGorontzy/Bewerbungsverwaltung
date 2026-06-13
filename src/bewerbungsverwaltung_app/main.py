from pathlib import Path

from .gui import BewerbungsverwaltungApp


def main() -> None:
    root = Path(__file__).resolve().parents[2]
    workbook_path = root / "Bewerbungsaktivitäten mit Erinnerungen.xlsx"
    app = BewerbungsverwaltungApp(workbook_path=workbook_path)
    app.run()
