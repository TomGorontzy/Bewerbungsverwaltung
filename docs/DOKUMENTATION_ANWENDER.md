# DOKUMENTATION ANWENDER

## Inhaltsverzeichnis

- [Zweck](#zweck)
- [Voraussetzungen](#voraussetzungen)
- [Start der Anwendung](#start-der-anwendung)
- [Bedienung](#bedienung)
  - [Neue Bewerbungen](#1-neue-bewerbungen)
  - [Nachzufassende Bewerbungen](#2-nachzufassende-bewerbungen)
  - [Erledigte Bewerbungen (Archiv)](#3-erledigte-bewerbungen-archiv)
- [Farblogik in Nachzufassen](#farblogik-in-nachzufassen)
- [Datensicherung](#datensicherung)
- [Häufige Probleme](#häufige-probleme)
- [Support-Hinweis](#support-hinweis)

## Zweck

Die Anwendung **Bewerbungsverwaltung** unterstützt bei der strukturierten Pflege von Bewerbungsaktivitäten auf Basis der Datei `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx`.

Die GUI bietet drei Hauptbereiche:

- **Neue Bewerbungen**
- **Nachzufassende Bewerbungen**
- **Erledigte Bewerbungen (Archiv)**

Zusätzlich enthält die Anwendung:

- zentrale Schaltfläche `Daten neu laden`
- verwaltbaren Dropdown-Katalog für **Nächster Schritt**

## Voraussetzungen

- Windows-PC
- Die Datei `Bewerbungsverwaltung.exe`
- Die Excel-Datei `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx`

> Wichtig: Die Excel-Datei liegt im Unterordner `data` neben der EXE.

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

Hinweis zu **Nächster Schritt**:

- Das Feld ist als Dropdown hinterlegt.
- Über den Button `Schritt-Katalog verwalten` kann die Liste gepflegt werden (hinzufügen, löschen, Reihenfolge ändern).
- Änderungen wirken direkt in GUI und Excel-Dropdowns.

Automatisch gesetzt durch Excel-Logik:

- Bewerbungs-ID
- Erinnerungsdatum
- Kennzeichnung „Heute erledigen?“

### 2) Nachzufassende Bewerbungen

Im Tab **Nachzufassende Bewerbungen** werden fällige Aktivitäten angezeigt.

Möglichkeiten:

- Filtern
- Spaltenweise sortieren
- Datensatz auswählen und unten aktualisieren (`Status`, `Kontaktart`, `Ergebnis Nachfrage`, `Nächster Schritt`, `Endergebnis`)

Typische Aktion:

1. Eintrag in der Liste auswählen
2. Status und Kontaktinformationen aktualisieren
3. `Änderung speichern` klicken

Empfehlung: Nach Telefonat/E-Mail direkt aktualisieren, damit Erinnerungen korrekt bleiben.

Hinweis für die direkte Arbeit in Excel:

- Nachzufassende Bewerbungen werden automatisch in den Tabellenblättern **Heute erledigen** und **Diese Woche erledigen** angezeigt.

### Farblogik in Nachzufassen

Überfällige Einträge werden in der Liste farblich hervorgehoben. Eine Legende ist als Statusleiste direkt im Tab sichtbar:

- **1–2 Tage überfällig**: hellgelb
- **3–6 Tage überfällig**: orange-beige
- **7–13 Tage überfällig**: orange
- **ab 14 Tagen überfällig**: hellrot

Kein Farbhintergrund bedeutet: aktuell noch nicht überfällig.

### 3) Erledigte Bewerbungen (Archiv)

Hier erscheinen abgeschlossene Bewerbungen (z. B. Absage/Zusage oder gesetztes Endergebnis).

## Datensicherung

- Vor größeren Änderungen eine Kopie der Excel-Datei anlegen.
- Bei OneDrive-Synchronisierung nach dem Speichern kurz warten.

## Häufige Probleme

### Excel-Datei wird nicht gefunden

- Prüfen, ob `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx` vorhanden ist.

### Speichern schlägt fehl / Datei gesperrt

- Excel-Datei in Microsoft Excel schließen
- OneDrive-Synchronisierung abwarten
- Erneut speichern

### Dropdown enthält erwarteten Wert nicht

- In `Neue Bewerbungen` auf `Schritt-Katalog verwalten` klicken und den Wert ergänzen.
- Alternativ prüfen, ob in Excel `Hilfstabellen` der Eintrag vorhanden ist.
- Danach `Daten neu laden` nutzen.

### Erinnerungsdatum wirkt leer/unerwartet

Die Anwendung berücksichtigt Excel-Formeln und nutzt bei fehlendem Formel-Cache eine interne Fallback-Berechnung für Follow-up-Fälle. Nach Änderungen kann `Daten neu laden` helfen.

### SmartScreen blockiert Start

- `Weitere Informationen` → `Trotzdem ausführen`

## Support-Hinweis

Bei technischen Rückfragen bitte angeben:

- Datum/Uhrzeit des Fehlers
- Kurze Beschreibung des letzten Arbeitsschritts
- Wenn möglich Screenshot der Meldung
