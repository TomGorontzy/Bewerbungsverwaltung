from __future__ import annotations

import warnings
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .constants import ARCHIVE_STATUS, FIELD_ORDER, FOLLOW_UP_STATUS, LOOKUP_COLUMNS, SHEET_LOOKUPS, SHEET_OVERVIEW
from .models import ApplicationRecord, LookupValues
from .utils import as_date, as_str, today


class ExcelRepository:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path

    @staticmethod
    def _load_workbook(*args, **kwargs):
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message="Data Validation extension is not supported and will be removed",
                category=UserWarning,
                module=r"openpyxl\..*",
            )
            return load_workbook(*args, **kwargs)

    def load_lookups(self) -> LookupValues:
        wb = self._load_workbook(self.workbook_path, data_only=True)
        ws = wb[SHEET_LOOKUPS]
        values: dict[str, list[str]] = {}

        for key, col in LOOKUP_COLUMNS.items():
            items: list[str] = []
            for row in range(2, 200):
                value = as_str(ws[f"{col}{row}"].value)
                if value and value not in items:
                    items.append(value)
            values[key] = items

        wb.close()
        return LookupValues(**values)

    def load_records(self) -> list[ApplicationRecord]:
        wb = self._load_workbook(self.workbook_path, data_only=True)
        ws = wb[SHEET_OVERVIEW]
        records: list[ApplicationRecord] = []

        for row in range(2, ws.max_row + 1):
            unternehmen = as_str(ws.cell(row=row, column=FIELD_ORDER["unternehmen"]).value)
            position = as_str(ws.cell(row=row, column=FIELD_ORDER["position"]).value)
            if not unternehmen and not position:
                continue

            values: dict[str, Any] = {"row": row}
            for field, col in FIELD_ORDER.items():
                raw = ws.cell(row=row, column=col).value
                if field in {"datum_bewerbung", "status_datum", "letzte_kontaktaufnahme", "erinnerungsdatum"}:
                    values[field] = as_date(raw)
                else:
                    values[field] = as_str(raw)

            records.append(ApplicationRecord(**values))

        wb.close()
        return records

    def add_record(self, payload: dict[str, Any]) -> None:
        wb = self._load_workbook(self.workbook_path, data_only=False)
        ws = wb[SHEET_OVERVIEW]

        row = ws.max_row + 1
        self._write_payload(ws, row, payload)
        self._apply_default_formulas(ws, row)

        wb.save(self.workbook_path)
        wb.close()

    def update_record(self, row: int, updates: dict[str, Any]) -> None:
        wb = self._load_workbook(self.workbook_path, data_only=False)
        ws = wb[SHEET_OVERVIEW]

        self._write_payload(ws, row, updates, partial=True)
        self._apply_default_formulas(ws, row)

        wb.save(self.workbook_path)
        wb.close()

    def categorize_records(self, records: list[ApplicationRecord]) -> tuple[list[ApplicationRecord], list[ApplicationRecord]]:
        follow_up: list[ApplicationRecord] = []
        archive: list[ApplicationRecord] = []

        now = today()
        for record in records:
            archived = bool(record.endergebnis) or record.status in ARCHIVE_STATUS
            if archived:
                archive.append(record)
                continue

            due_by_reminder = record.erinnerungsdatum is not None and record.erinnerungsdatum <= now
            due_by_status_age = (
                record.status in FOLLOW_UP_STATUS
                and record.status_datum is not None
                and (now - record.status_datum).days >= 14
            )
            flagged_today = record.heute_erledigen.lower() == "ja"

            if due_by_reminder or due_by_status_age or flagged_today:
                follow_up.append(record)

        follow_up.sort(key=lambda x: (x.erinnerungsdatum or date.max, x.prioritaet != "Hoch"))
        archive.sort(key=lambda x: (x.status_datum or date.min), reverse=True)
        return follow_up, archive

    def _write_payload(self, ws, row: int, payload: dict[str, Any], partial: bool = False) -> None:
        for field, value in payload.items():
            if field not in FIELD_ORDER:
                continue
            col = FIELD_ORDER[field]
            if value is None and partial:
                continue
            ws.cell(row=row, column=col).value = value

    @staticmethod
    def _apply_default_formulas(ws, row: int) -> None:
        ws.cell(row=row, column=FIELD_ORDER["bewerbungs_id"]).value = (
            f'=IF(B{row}="","","BEW-"&YEAR(J{row})&"-"&TEXT(ROW()-1,"000"))'
        )
        ws.cell(row=row, column=FIELD_ORDER["erinnerungsdatum"]).value = (
            f'=IF(OR(L{row}="Bewerbung versendet",L{row}="Eingangsbestätigung",L{row}="Rückmeldung ausstehend"),M{row}+14,"")'
        )
        ws.cell(row=row, column=FIELD_ORDER["heute_erledigen"]).value = (
            f'=IF(AND(R{row}<>"",R{row}<=TODAY(),T{row}=""),"ja","")'
        )
