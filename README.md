# Bewerbungsverwaltung (Python + Excel + GUI)

Desktop-Anwendung zur geführten Pflege einer bestehenden Excel-Arbeitsmappe für Bewerbungsprozesse.

Die App arbeitet auf Basis von `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx` und ergänzt eine strukturierte GUI für Erfassung, Nachfassen und Archivierung.

> **Migrationshinweis:** Die Excel-Datei wurde vom Root-Verzeichnis in den Unterordner `data/` verschoben. Falls Sie ein bestehendes Setup aktualisieren, legen Sie die Arbeitsmappe bitte unter `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx` ab.

## Funktionsüberblick

- **Neue Bewerbungen**
  - strukturierte Erfassung mit Pflichtfeldern
  - Datumsauswahl mit Picker
  - konsistente Dropdown-Werte aus Excel-`Hilfstabellen`
- **Nachzufassende Bewerbungen**
  - fällige/überfällige Einträge mit direkter Aktualisierung
  - farbliche Hervorhebung je Überfälligkeit
  - Legende als Statusleiste im Tab
- **Erledigte Bewerbungen (Archiv)**
  - abgeschlossene Vorgänge (z. B. Absage/Zusage) mit Filter/Sortierung

## Voraussetzungen

- Windows
- Python 3.13 (für Entwicklung)
- Datei `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx` im Projektordner bzw. neben der EXE

## Start (Entwicklung)

- `python src/run.py`

## EXE-Build

- `./src/build_exe.ps1`
- optional Clean-Build: `./src/build_exe.ps1 -Clean`

Ergebnis: `dist/Bewerbungsverwaltung.exe`

## Release-Erstellung

- `./src/release.ps1 -ReleaseVersion vX.Y.Z`
- optional Clean-Release: `./src/release.ps1 -ReleaseVersion vX.Y.Z -Clean`

Der Release-Prozess umfasst u. a.:

- Markdown-Lint/Fix
- EXE-Build
- Tag-Erstellung/Push
- GitHub Release inkl. EXE-Asset

## Excel-Integration

Schreibziel ist `Bewerbungsübersicht`. Wichtige Excel-Mechaniken:

- Formel-basierte Felder (u. a. Bewerbungs-ID, Erinnerungsdatum, „Heute erledigen?“)
- Lookup-/Dropdown-Werte aus `Hilfstabellen`
- Auto-Reparatur beim Start für:
  - Summary-Formeln (`Heute erledigen`, `Diese Woche erledigen`)
  - Datenvalidierungen
  - bedingte Formatierungen

Zusätzlich verfügbar:

- Verwaltung des Katalogs **„Nächster Schritt“** direkt in der GUI
  - Änderungen werden in `Hilfstabellen` gespeichert
  - Dropdowns in GUI und Excel werden synchron gehalten

## Dokumentation

- Anwenderdokumentation: `docs/DOKUMENTATION_ANWENDER.md`
- Technische Dokumentation: `docs/DOKUMENTATION_TECHNIK.md`
- FAQ: `docs/FAQ.md`

### Dokumentation als PDF exportieren

- Konvertierungsskript: `src/scripts/convert_docs_to_pdf.ps1`
- Erstellt modern gestaltete PDFs für:
  - `docs/DOKUMENTATION_ANWENDER.md`
  - `docs/DOKUMENTATION_TECHNIK.md`
  - `docs/SCHNELLSTART.md`
- Ausgabeordner: `docs/pdf/`
- Feste PDF-Regeln:
  - Seite 1 = Titelblatt
  - Seite 2 = Inhaltsverzeichnis mit klickbaren Hyperlinks
  - ab Seite 3 = eigentlicher Inhalt
  - Überschriften bleiben nach Möglichkeit mit dem direkt folgenden Inhalt zusammen

- In der GUI sind die PDFs direkt über die Header-Schaltflächen `Schnellstart`, `Anwender-Doku` und `Technik-Doku` erreichbar.

Farbakzent im Layout basiert auf der definierten CI-Farbe (CMYK `0 / 60 / 100 / 0`).
