from __future__ import annotations

from collections import OrderedDict

SHEET_OVERVIEW = "Bewerbungsübersicht"
SHEET_LOOKUPS = "Hilfstabellen"

FIELD_ORDER = OrderedDict(
    [
        ("bewerbungs_id", 1),
        ("unternehmen", 2),
        ("ort", 3),
        ("position", 4),
        ("referenznummer", 5),
        ("ansprechpartner", 6),
        ("kontakt", 7),
        ("quelle", 8),
        ("art_bewerbung", 9),
        ("datum_bewerbung", 10),
        ("unterlagen", 11),
        ("status", 12),
        ("status_datum", 13),
        ("letzte_kontaktaufnahme", 14),
        ("art_kontaktaufnahme", 15),
        ("ergebnis_nachfrage", 16),
        ("naechster_schritt", 17),
        ("erinnerungsdatum", 18),
        ("prioritaet", 19),
        ("endergebnis", 20),
        ("notizen", 21),
        ("heute_erledigen", 22),
    ]
)

LOOKUP_COLUMNS = {
    "unterlagen": "A",
    "art_bewerbung": "C",
    "status": "E",
    "art_kontaktaufnahme": "G",
    "prioritaet": "I",
    "endergebnis": "K",
    "quelle": "M",
    "naechster_schritt": "O",
}

DEFAULT_NAECHSTER_SCHRITT_CATALOG = [
    "Unterlagen prüfen und ergänzen",
    "Eingang der Bewerbung bestätigen lassen",
    "Freundliche Nachfrage per E-Mail senden",
    "Telefonisches Follow-up durchführen",
    "Interviewtermin abstimmen",
    "Interview vorbereiten",
    "Interview nachbereiten und bedanken",
    "Zusätzliche Unterlagen nachreichen",
    "Finale Rückmeldung anfordern",
    "Bewerbung abschließen und dokumentieren",
]

ARCHIVE_STATUS = {"Absage", "Zusage"}
FOLLOW_UP_STATUS = {"Bewerbung versendet", "Eingangsbestätigung", "Rückmeldung ausstehend"}
