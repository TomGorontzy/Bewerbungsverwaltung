# DOKUMENTATION TECHNIK

## Inhaltsverzeichnis

- Projektüberblick
- Repository-Struktur
- Datenquelle und Blätter
- Fachlogik (MVP)
	- Neue Bewerbungen
	- Nachzufassen
	- Archiv
- Entwicklung lokal
- Build (EXE)
- CI
- Bekannte technische Hinweise
- Erweiterungsvorschläge

## Projektüberblick

Dieses Projekt stellt eine Desktop-GUI für eine bestehende Excel-basierte Bewerbungsverwaltung bereit.

- Sprache: **Python 3.13**
- GUI: **tkinter/ttk**
- Excel-Zugriff: **openpyxl**
- EXE-Build: **PyInstaller**

## Repository-Struktur

- `run.py` – Einstiegspunkt
- `src/bewerbungsverwaltung_app/main.py` – App-Start
- `src/bewerbungsverwaltung_app/gui.py` – GUI (Tabs, Events, Bedienlogik)
- `src/bewerbungsverwaltung_app/excel_repository.py` – Excel-Lese-/Schreiblogik
- `src/bewerbungsverwaltung_app/constants.py` – Feld-/Spaltenzuordnung, Kategorien
- `src/bewerbungsverwaltung_app/models.py` – Datenmodelle
- `src/bewerbungsverwaltung_app/utils.py` – Konvertierung (String/Datum)
- `build_exe.ps1` – Build-Helfer für EXE
- `.github/workflows/ci.yml` – CI Smoke-Checks

## Datenquelle und Blätter

Primäre Datei:

- `Bewerbungsaktivitäten mit Erinnerungen.xlsx`

Relevante Blätter:

- `Bewerbungsübersicht` (Master-Daten)
- `Hilfstabellen` (Lookup-/Dropdown-Werte)

Andere Blätter wie `Heute erledigen` und `Diese Woche erledigen` werden als abgeleitete Sichten verstanden.

## Fachlogik (MVP)

### Neue Bewerbungen

Schreibt neue Zeilen in `Bewerbungsübersicht`.

Setzt/übernimmt Formeln für:

- Bewerbungs-ID (Spalte A)
- Erinnerungsdatum (Spalte R)
- Heute erledigen? (Spalte V)

### Nachzufassen

Kategorisierung auf Basis von:

- Erinnerungsdatum fällig
- Follow-up-Status älter als 14 Tage
- Feld „Heute erledigen?“ = `ja`

### Archiv

Archiviert (in GUI-Sicht), wenn:

- Endergebnis gesetzt ist oder
- Status in finaler Menge (`Absage`, `Zusage`) liegt

## Entwicklung lokal

Voraussetzungen:

- Python 3.13
- virtuelle Umgebung `.venv`

Abhängigkeiten installieren:

- `pip install -r requirements.txt`

Starten:

- `python run.py`

## Build (EXE)

Skript:

- `./build_exe.ps1`

Optional clean:

- `./build_exe.ps1 -Clean`

Ausgabe:

- `dist/Bewerbungsverwaltung.exe`

## CI

Workflow:

- `.github/workflows/ci.yml`

Checks:

- Dependency-Installation
- Syntaxprüfung (`py_compile`)
- Excel-Smoke-Test gegen die mitgelieferte Arbeitsmappe

## Bekannte technische Hinweise

- `openpyxl` kann Warnungen zu Data Validation Extensions ausgeben.
- Bei OneDrive kann es temporär Dateisperren geben (insb. während Sync).
- Bei EXE-Ausführung ist entscheidend, dass Excel-Datei und EXE im selben Verzeichnis liegen.

## Erweiterungsvorschläge

- Feldvalidierung weiter ausbauen (Domänenregeln je Status)
- Such-/Filterfunktionen im Archiv
- Optionaler Wechsel auf Excel-Automation (z. B. `xlwings`) für komplexere Excel-Features
- Logging/Diagnosemodus für Supportfälle
