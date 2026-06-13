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

Bei jedem Build wird automatisch eine Markdown-Lint/Fix-Routine ausgeführt.

## Release erstellen

- `./release.ps1 -Version v0.1.2`
- optional sauber neu: `./release.ps1 -Version v0.1.2 -Clean`

Der Release-Prozess führt automatisch aus:

- Markdown-Lint/Fix
- optionales Auto-Commit von Markdown-Korrekturen
- EXE-Build
- Tag-Erstellung und Push
- GitHub Release mit EXE-Asset

## Dokumentation

- Anwenderdokumentation: `docs/DOKUMENTATION_ANWENDER.md`
- Technische Dokumentation: `docs/DOKUMENTATION_TECHNIK.md`
- FAQ: `docs/FAQ.md`

## Hinweise zur Excel-Logik

Die App schreibt in das Blatt **Bewerbungsübersicht** und übernimmt die bestehenden Formeln für:

- Bewerbungs-ID
- Erinnerungsdatum
- Heute erledigen?

Lookup-Werte (Dropdowns) werden aus `Hilfstabellen` gelesen.
