# FAQ

## Allgemein

### Was macht die Anwendung?
Die Anwendung unterstützt die Pflege von Bewerbungsaktivitäten in drei Bereichen:

- **Neue Bewerbungen**
- **Nachzufassende Bewerbungen**
- **Erledigte Bewerbungen (Archiv)**

Die Daten werden in der Excel-Datei `Bewerbungsaktivitäten mit Erinnerungen.xlsx` gespeichert.

### Muss ich weiterhin mit Excel arbeiten?
Ja. Die GUI ist ein geführter Zugang zur bestehenden Excel-Arbeitsmappe. Excel bleibt die zentrale Datenbasis.

## Installation & Start

### Wo müssen die Dateien liegen?
`Bewerbungsverwaltung.exe` und `Bewerbungsaktivitäten mit Erinnerungen.xlsx` müssen im **gleichen Ordner** liegen.

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

## Probleme & Lösungen

### Die App findet die Excel-Datei nicht.
Bitte prüfen:

- Dateiname exakt `Bewerbungsaktivitäten mit Erinnerungen.xlsx`
- Datei liegt im selben Ordner wie die EXE

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

## Versionen & Updates

### Wo finde ich neue Versionen?
Über die GitHub Releases des Projekts (z. B. `v0.1.0`).

### Kann ich einfach eine neue EXE über die alte kopieren?
Ja. Vorher am besten die laufende App schließen und die Excel-Datei sichern.

## Support

### Welche Infos soll ich bei einem Fehler mitgeben?
- Kurze Beschreibung des Problems
- Zeitpunkt (Datum/Uhrzeit)
- Letzter Klick/Arbeitsschritt vor dem Fehler
- Screenshot der Fehlermeldung (falls vorhanden)
