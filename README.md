# Bewerbungsverwaltung (Python + Excel + GUI)

Dieses Projekt erweitert die vorhandene Excel-Datei `Bewerbungsaktivitäten mit Erinnerungen.xlsx` um eine geführte Desktop-GUI.

## Rubriken in der App

- **Neue Bewerbungen**: strukturierte Erfassung mit Pflichtfeldern
- **Nachzufassende Bewerbungen**: fällige Aktivitäten, direkte Status-Updates
- **Erledigte Bewerbungen (Archiv)**: abgeschlossene Vorgänge

## Voraussetzungen

- Windows mit Python 3.13 (virtuelle Umgebung `.venv`)
- Vorhandene Excel-Datei im Projektordner:
  - `Bewerbungsaktivitäten mit Erinnerungen.xlsx`

## Starten (Entwicklung)

- `python run.py`

## EXE bauen

- `./build_exe.ps1`
- optional sauber neu: `./build_exe.ps1 -Clean`

Danach liegt die Anwendung als `dist/Bewerbungsverwaltung.exe` vor.

## Hinweise zur Excel-Logik

Die App schreibt in das Blatt **Bewerbungsübersicht** und übernimmt die bestehenden Formeln für:

- Bewerbungs-ID
- Erinnerungsdatum
- Heute erledigen?

Lookup-Werte (Dropdowns) werden aus `Hilfstabellen` gelesen.
