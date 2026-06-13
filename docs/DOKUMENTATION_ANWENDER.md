# DOKUMENTATION ANWENDER

## Inhaltsverzeichnis

- [Zweck](#zweck)
- [Voraussetzungen](#voraussetzungen)
- [Start der Anwendung](#start-der-anwendung)
- [Bedienung](#bedienung)
  - [Neue Bewerbungen](#1-neue-bewerbungen)
  - [Nachzufassende Bewerbungen](#2-nachzufassende-bewerbungen)
  - [Erledigte Bewerbungen (Archiv)](#3-erledigte-bewerbungen-archiv)
- [Datensicherung](#datensicherung)
- [Häufige Probleme](#häufige-probleme)
- [Support-Hinweis](#support-hinweis)

## Zweck

Die Anwendung **Bewerbungsverwaltung** unterstützt bei der strukturierten Pflege von Bewerbungsaktivitäten auf Basis der Datei `Bewerbungsaktivitäten mit Erinnerungen.xlsx`.

Die GUI bietet drei Hauptbereiche:

- **Neue Bewerbungen**
- **Nachzufassende Bewerbungen**
- **Erledigte Bewerbungen (Archiv)**

## Voraussetzungen

- Windows-PC
- Die Datei `Bewerbungsverwaltung.exe`
- Die Excel-Datei `Bewerbungsaktivitäten mit Erinnerungen.xlsx`

> Wichtig: EXE und Excel-Datei müssen im selben Ordner liegen.

## Start der Anwendung

1. `Bewerbungsverwaltung.exe` per Doppelklick starten.
2. Falls Windows SmartScreen warnt:
   - `Weitere Informationen`
   - `Trotzdem ausführen`

## Bedienung

### 1) Neue Bewerbungen

Im Tab **Neue Bewerbungen** neue Einträge erfassen.

Pflichtfelder:

- Unternehmen
- Position
- Datum der Bewerbung
- Aktueller Status

Ablauf:

1. Felder ausfüllen
2. Dropdowns verwenden (einheitliche Werte)
3. `Speichern` klicken

Automatisch gesetzt durch Excel-Logik:

- Bewerbungs-ID
- Erinnerungsdatum
- Kennzeichnung „Heute erledigen?“

### 2) Nachzufassende Bewerbungen

Im Tab **Nachzufassende Bewerbungen** werden fällige Aktivitäten angezeigt.

Typische Aktion:

1. Eintrag in der Liste auswählen
2. Status und Kontaktinformationen aktualisieren
3. `Änderung speichern` klicken

Empfehlung: Nach Telefonat/E-Mail direkt aktualisieren, damit Erinnerungen korrekt bleiben.

### 3) Erledigte Bewerbungen (Archiv)

Hier erscheinen abgeschlossene Bewerbungen (z. B. Absage/Zusage oder gesetztes Endergebnis).

## Datensicherung

- Vor größeren Änderungen eine Kopie der Excel-Datei anlegen.
- Bei OneDrive-Synchronisierung nach dem Speichern kurz warten.

## Häufige Probleme

### Excel-Datei wird nicht gefunden

- Prüfen, ob `Bewerbungsaktivitäten mit Erinnerungen.xlsx` im gleichen Ordner wie die EXE liegt.

### Speichern schlägt fehl / Datei gesperrt

- Excel-Datei in Microsoft Excel schließen
- OneDrive-Synchronisierung abwarten
- Erneut speichern

### SmartScreen blockiert Start

- `Weitere Informationen` → `Trotzdem ausführen`

## Support-Hinweis

Bei technischen Rückfragen bitte angeben:

- Datum/Uhrzeit des Fehlers
- Kurze Beschreibung des letzten Arbeitsschritts
- Wenn möglich Screenshot der Meldung
