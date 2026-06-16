# DOKUMENTATION ANWENDER

## Inhaltsverzeichnis

- [Zweck](#zweck)
- [Voraussetzungen](#voraussetzungen)
- [Start der Anwendung](#start-der-anwendung)
- [Bedienung](#bedienung)
  - [Neue Bewerbungen](#1-neue-bewerbungen)
  - [Laufende Bewerbungen](#2-laufende-bewerbungen)
  - [Nachzufassende Bewerbungen](#3-nachzufassende-bewerbungen)
  - [Erledigte Bewerbungen (Archiv)](#4-erledigte-bewerbungen-archiv)
  - [Detailbearbeitung per Doppelklick](#detailbearbeitung-per-doppelklick)
- [Farblogik in Nachzufassen](#farblogik-in-nachzufassen)
- [Datensicherung](#datensicherung)
- [Häufige Probleme](#häufige-probleme)
- [Support-Hinweis](#support-hinweis)

## Zweck

Die Anwendung **Bewerbungsverwaltung** unterstützt bei der strukturierten Pflege von Bewerbungsaktivitäten auf Basis der Datei `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx`.

Die GUI bietet drei Hauptbereiche:

- **Neue Bewerbungen**
- **Laufende Bewerbungen**
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

Fenstermodus/Vollbild:

- Die Anwendung startet standardmäßig maximiert.
- Mit `F11` kann der echte Vollbildmodus ein- oder ausgeschaltet werden.
- Mit `Esc` wird der echte Vollbildmodus beendet.

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

### 2) Laufende Bewerbungen

Im Tab **Laufende Bewerbungen** werden alle offenen Bewerbungen ohne Endergebnis angezeigt.

Möglichkeiten:

- Filtern
- Spaltenweise sortieren
- Datensatz auswählen und unten aktualisieren (`Status`, `Kontaktart`, `Ergebnis Nachfrage`, `Nächster Schritt`, `Endergebnis`, `Datum Bewerbung`)

Typische Aktion:

1. Eintrag in der Liste auswählen
2. Änderungen unten eintragen
3. `Änderung speichern` klicken

### 3) Nachzufassende Bewerbungen

Im Tab **Nachzufassende Bewerbungen** werden fällige Aktivitäten angezeigt.

Kriterien für die Listenführung:

- Einträge mit **Endergebnis** oder Status **Absage/Zusage** werden nicht hier geführt, sondern im Archiv.
- Ein Eintrag erscheint in der Liste, wenn **mindestens eines** der folgenden Kriterien erfüllt ist:
  - `Erinnerungsdatum` ist erreicht oder überschritten (≤ heute),
  - Status gehört zu den Follow-up-Status und das `Status-Datum` ist mindestens 14 Tage alt,
  - Kennzeichnung `Heute erledigen?` steht auf `ja`.
- Sortierung in der Nachfassliste:
  - primär nach `Erinnerungsdatum` (frühere zuerst),
  - bei Gleichstand werden hohe Prioritäten bevorzugt.

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

### 4) Erledigte Bewerbungen (Archiv)

Hier erscheinen abgeschlossene Bewerbungen (z. B. Absage/Zusage oder gesetztes Endergebnis).

### Detailbearbeitung per Doppelklick

In allen drei Listen (`Laufende Bewerbungen`, `Nachzufassende Bewerbungen`, `Erledigte Bewerbungen`) öffnet ein **Doppelklick** auf einen Eintrag ein separates Detailfenster.

Im Detailfenster:

- werden alle Felder aus `Bewerbungsübersicht` angezeigt,
- können Felder direkt bearbeitet werden,
- werden Änderungen erst nach Bestätigung (`Speichern`) in Excel geschrieben.

Schließen/Abbrechen:

- `Abbrechen` oder Schließen über `X` fragt bei ungespeicherten Änderungen nach,
- Änderungen können gespeichert oder verworfen werden,
- bei Auswahl `Abbrechen` bleibt das Fenster geöffnet.

Hinweis zur Anzeige ungespeicherter Änderungen:

- `*` im Fenstertitel,
- Statushinweis im Fenster (`Ungespeicherte Änderungen`).

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
