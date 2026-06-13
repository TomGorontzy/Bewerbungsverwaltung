from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Optional


@dataclass(slots=True)
class ApplicationRecord:
    row: int
    bewerbungs_id: str
    unternehmen: str
    ort: str
    position: str
    referenznummer: str
    ansprechpartner: str
    kontakt: str
    quelle: str
    art_bewerbung: str
    datum_bewerbung: Optional[date]
    unterlagen: str
    status: str
    status_datum: Optional[date]
    letzte_kontaktaufnahme: Optional[date]
    art_kontaktaufnahme: str
    ergebnis_nachfrage: str
    naechster_schritt: str
    erinnerungsdatum: Optional[date]
    prioritaet: str
    endergebnis: str
    notizen: str
    heute_erledigen: str


@dataclass(slots=True)
class LookupValues:
    unterlagen: list[str]
    art_bewerbung: list[str]
    status: list[str]
    art_kontaktaufnahme: list[str]
    prioritaet: list[str]
    endergebnis: list[str]
    quelle: list[str]
