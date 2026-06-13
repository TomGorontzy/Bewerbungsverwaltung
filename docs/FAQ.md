# FAQ

## Inhaltsverzeichnis

- [Allgemein](#allgemein)
- [Installation & Start](#installation--start)
- [Dateneingabe](#dateneingabe)
- [Probleme & Lösungen](#probleme--lösungen)
- [Versionen & Updates](#versionen--updates)
- [Support](#support)

## Allgemein

### Was macht die Anwendung?
Die Anwendung unterstützt die Pflege von Bewerbungsaktivitäten in drei Bereichen:

- **Neue Bewerbungen**
- **Nachzufassende Bewerbungen**
- **Erledigte Bewerbungen (Archiv)**

Die Daten werden in der Excel-Datei `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx` gespeichert.

### Muss ich weiterhin mit Excel arbeiten?
Ja. Die GUI ist ein geführter Zugang zur bestehenden Excel-Arbeitsmappe. Excel bleibt die zentrale Datenbasis.

## Installation & Start

### Wo müssen die Dateien liegen?
`Bewerbungsverwaltung.exe` muss neben dem Unterordner `data` liegen; die Excel-Datei befindet sich in `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx`.

### Wie starte ich die Anwendung?
Per Doppelklick auf `Bewerbungsverwaltung.exe`.

### Windows zeigt „Der PC wurde geschützt“ (SmartScreen) – was tun?
1. `Weitere Informationen` anklicken
2. `Trotzdem ausführen` wählen

## Dateneingabe

### Welche Felder sind bei neuen Bewerbungen Pflicht?
- Unternehmen
- Position
- Datum der Bewerbung
- Aktueller Status

### Warum werden manche Felder automatisch befüllt?
Die Excel-Logik setzt automatisch:

- Bewerbungs-ID
- Erinnerungsdatum
- Kennzeichnung „Heute erledigen?“

### Wie pflege ich die Liste für „Nächster Schritt“?
Im Tab **Neue Bewerbungen** auf `Schritt-Katalog verwalten` klicken.

Dort können Sie:

- neue Einträge hinzufügen
- vorhandene Einträge löschen
- Reihenfolge ändern

Nach `Speichern` stehen die Werte sofort in GUI und Excel-Dropdowns zur Verfügung.

## Probleme & Lösungen

### Die App findet die Excel-Datei nicht.
Bitte prüfen:

- Dateiname exakt `Bewerbungsaktivitäten mit Erinnerungen.xlsx`
- Datei liegt unter `data/Bewerbungsaktivitäten mit Erinnerungen.xlsx`

### Speichern funktioniert nicht / Datei ist gesperrt.
Mögliche Ursache: Datei ist in Excel geöffnet oder durch OneDrive-Sync gesperrt.

Lösung:
- Excel schließen
- kurz auf Abschluss der OneDrive-Synchronisierung warten
- erneut speichern

### Nach dem Speichern sehe ich die Änderung nicht sofort.
Bitte auf `Daten neu laden` klicken, um die Ansicht in der GUI zu aktualisieren.

### Warum taucht ein Datensatz nicht in „Nachzufassende Bewerbungen“ auf?
Die Anzeige basiert auf Fälligkeitslogik (z. B. Erinnerungsdatum, Status-Alter, „Heute erledigen?“). Nicht-fällige Einträge erscheinen dort nicht.

Hinweis für die direkte Bearbeitung in Excel:

- Fällige Nachfassungen werden automatisch in den Blättern **Heute erledigen** und **Diese Woche erledigen** angezeigt.

### Warum sind Einträge in „Nachzufassende Bewerbungen“ farbig?
Die Farben zeigen die Überfälligkeit des Nachfassungstermins:

- 1–2 Tage: hellgelb
- 3–6 Tage: orange-beige
- 7–13 Tage: orange
- ab 14 Tagen: hellrot

Die Legende ist direkt im Tab als Statusleiste sichtbar.

### Warum wirkt das Erinnerungsdatum manchmal leer?
Wenn Excel Formelwerte noch nicht gecacht hat, kann ein Datum leer erscheinen. Die Anwendung nutzt für relevante Follow-up-Fälle eine Fallback-Berechnung.

## Versionen & Updates

### Wo finde ich neue Versionen?
Über die GitHub Releases des Projekts (z. B. `v0.1.13`).

### Kann ich einfach eine neue EXE über die alte kopieren?
Ja. Vorher am besten die laufende App schließen und die Excel-Datei sichern.

## Support

### Welche Infos soll ich bei einem Fehler mitgeben?
- Kurze Beschreibung des Problems
- Zeitpunkt (Datum/Uhrzeit)
- Letzter Klick/Arbeitsschritt vor dem Fehler
- Screenshot der Fehlermeldung (falls vorhanden)
