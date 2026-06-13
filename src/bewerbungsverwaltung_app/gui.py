from __future__ import annotations

import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import messagebox, ttk

from .excel_repository import ExcelRepository
from .models import ApplicationRecord, LookupValues
from .utils import as_date


class BewerbungsverwaltungApp:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self.repo = ExcelRepository(workbook_path)
        self.lookups: LookupValues = self.repo.load_lookups()

        self.root = tk.Tk()
        self.root.title("Bewerbungsverwaltung")
        self.root.geometry("1250x780")

        self.records: list[ApplicationRecord] = []
        self.follow_up_records: list[ApplicationRecord] = []
        self.archive_records: list[ApplicationRecord] = []

        self.form_vars: dict[str, tk.StringVar] = {}
        self.update_vars: dict[str, tk.StringVar] = {}

        self._build_ui()
        self.refresh_data()

    def run(self) -> None:
        self.root.mainloop()

    def _build_ui(self) -> None:
        header = ttk.Frame(self.root, padding=12)
        header.pack(fill="x")

        ttk.Label(
            header,
            text="Bewerbungsverwaltung – geführte Pflege",
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left")

        ttk.Button(header, text="Daten neu laden", command=self.refresh_data).pack(side="right", padx=4)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab_new = ttk.Frame(notebook, padding=12)
        self.tab_follow_up = ttk.Frame(notebook, padding=12)
        self.tab_archive = ttk.Frame(notebook, padding=12)

        notebook.add(self.tab_new, text="Neue Bewerbungen")
        notebook.add(self.tab_follow_up, text="Nachzufassende Bewerbungen")
        notebook.add(self.tab_archive, text="Erledigte Bewerbungen (Archiv)")

        self._build_new_tab()
        self._build_follow_up_tab()
        self._build_archive_tab()

    def _build_new_tab(self) -> None:
        info = (
            "Pflichtfelder: Unternehmen, Position, Datum der Bewerbung, Status. "
            "Die Bewerbungs-ID und das Erinnerungsdatum werden automatisch gesetzt."
        )
        ttk.Label(self.tab_new, text=info).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        fields = [
            ("unternehmen", "Unternehmen *", "entry"),
            ("ort", "Ort", "entry"),
            ("position", "Position *", "entry"),
            ("referenznummer", "Referenznummer", "entry"),
            ("ansprechpartner", "Ansprechpartner", "entry"),
            ("kontakt", "Kontakt (E-Mail/Telefon)", "entry"),
            ("quelle", "Quelle", "combo", self.lookups.quelle),
            ("art_bewerbung", "Art der Bewerbung", "combo", self.lookups.art_bewerbung),
            ("datum_bewerbung", "Datum Bewerbung (TT.MM.JJJJ)", "entry"),
            ("unterlagen", "Versendete Unterlagen", "combo", self.lookups.unterlagen),
            ("status", "Aktueller Status *", "combo", self.lookups.status),
            ("status_datum", "Status-Datum (TT.MM.JJJJ)", "entry"),
            ("prioritaet", "Priorität", "combo", self.lookups.prioritaet),
            ("naechster_schritt", "Nächster Schritt", "entry"),
            ("notizen", "Notizen", "entry"),
        ]

        for idx, definition in enumerate(fields, start=1):
            key = definition[0]
            label = definition[1]
            widget_type = definition[2]
            values = definition[3] if len(definition) > 3 else None

            row = idx
            ttk.Label(self.tab_new, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            var = tk.StringVar()
            self.form_vars[key] = var

            if widget_type == "combo":
                widget = ttk.Combobox(self.tab_new, textvariable=var, values=values or [], state="readonly", width=45)
                if values:
                    var.set(values[0])
            else:
                widget = ttk.Entry(self.tab_new, textvariable=var, width=48)

            widget.grid(row=row, column=1, sticky="w", pady=4)

        buttons = ttk.Frame(self.tab_new)
        buttons.grid(row=len(fields) + 2, column=0, columnspan=2, pady=16, sticky="w")

        ttk.Button(buttons, text="Speichern", command=self._save_new_record).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Formular leeren", command=self._reset_form).pack(side="left")

    def _build_follow_up_tab(self) -> None:
        ttk.Label(
            self.tab_follow_up,
            text="Fällige Bewerbungen: Erinnerungsdatum erreicht oder Follow-up überfällig.",
        ).pack(anchor="w", pady=(0, 8))

        columns = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt")
        self.follow_tree = ttk.Treeview(self.tab_follow_up, columns=columns, show="headings", height=16)
        self.follow_tree.pack(fill="x", pady=(0, 10))

        headings = {
            "id": "ID",
            "unternehmen": "Unternehmen",
            "position": "Position",
            "status": "Status",
            "erinnerung": "Erinnerungsdatum",
            "prioritaet": "Priorität",
            "schritt": "Nächster Schritt",
        }
        widths = {"id": 120, "unternehmen": 210, "position": 220, "status": 170, "erinnerung": 130, "prioritaet": 90, "schritt": 280}

        for col in columns:
            self.follow_tree.heading(col, text=headings[col])
            self.follow_tree.column(col, width=widths[col], anchor="w")

        update_box = ttk.LabelFrame(self.tab_follow_up, text="Ausgewählte Bewerbung aktualisieren", padding=10)
        update_box.pack(fill="x")

        mapping = [
            ("status", "Neuer Status", self.lookups.status),
            ("art_kontaktaufnahme", "Kontaktart", self.lookups.art_kontaktaufnahme),
            ("ergebnis_nachfrage", "Ergebnis Nachfrage", None),
            ("naechster_schritt", "Nächster Schritt", None),
            ("endergebnis", "Endergebnis", self.lookups.endergebnis),
        ]

        for i, (key, label, values) in enumerate(mapping):
            ttk.Label(update_box, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=4)
            var = tk.StringVar()
            self.update_vars[key] = var
            if values is not None:
                widget = ttk.Combobox(update_box, textvariable=var, values=values, state="readonly", width=45)
                if values:
                    var.set(values[0])
            else:
                widget = ttk.Entry(update_box, textvariable=var, width=48)
            widget.grid(row=i, column=1, sticky="w", pady=4)

        ttk.Button(update_box, text="Änderung speichern", command=self._save_follow_up_update).grid(
            row=len(mapping), column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

    def _build_archive_tab(self) -> None:
        ttk.Label(
            self.tab_archive,
            text="Archiv: abgeschlossene Bewerbungen (Absage/Zusage oder mit Endergebnis).",
        ).pack(anchor="w", pady=(0, 8))

        columns = ("id", "unternehmen", "position", "status", "endergebnis", "status_datum")
        self.archive_tree = ttk.Treeview(self.tab_archive, columns=columns, show="headings", height=22)
        self.archive_tree.pack(fill="both", expand=True)

        headings = {
            "id": "ID",
            "unternehmen": "Unternehmen",
            "position": "Position",
            "status": "Status",
            "endergebnis": "Endergebnis",
            "status_datum": "Status-Datum",
        }
        widths = {"id": 120, "unternehmen": 230, "position": 260, "status": 180, "endergebnis": 180, "status_datum": 120}

        for col in columns:
            self.archive_tree.heading(col, text=headings[col])
            self.archive_tree.column(col, width=widths[col], anchor="w")

    def _save_new_record(self) -> None:
        payload = {
            "unternehmen": self.form_vars["unternehmen"].get().strip(),
            "ort": self.form_vars["ort"].get().strip(),
            "position": self.form_vars["position"].get().strip(),
            "referenznummer": self.form_vars["referenznummer"].get().strip(),
            "ansprechpartner": self.form_vars["ansprechpartner"].get().strip(),
            "kontakt": self.form_vars["kontakt"].get().strip(),
            "quelle": self.form_vars["quelle"].get().strip(),
            "art_bewerbung": self.form_vars["art_bewerbung"].get().strip(),
            "datum_bewerbung": as_date(self.form_vars["datum_bewerbung"].get().strip()),
            "unterlagen": self.form_vars["unterlagen"].get().strip(),
            "status": self.form_vars["status"].get().strip(),
            "status_datum": as_date(self.form_vars["status_datum"].get().strip()) or date.today(),
            "prioritaet": self.form_vars["prioritaet"].get().strip(),
            "naechster_schritt": self.form_vars["naechster_schritt"].get().strip(),
            "notizen": self.form_vars["notizen"].get().strip(),
        }

        missing = [name for name in ("unternehmen", "position", "datum_bewerbung", "status") if not payload[name]]
        if missing:
            messagebox.showwarning("Pflichtfelder", "Bitte Unternehmen, Position, Datum Bewerbung und Status ausfüllen.")
            return

        self.repo.add_record(payload)
        self.refresh_data()
        self._reset_form()
        messagebox.showinfo("Gespeichert", "Neue Bewerbung wurde in Excel gespeichert.")

    def _save_follow_up_update(self) -> None:
        selected = self.follow_tree.selection()
        if not selected:
            messagebox.showwarning("Auswahl fehlt", "Bitte zuerst eine Bewerbung in der Liste auswählen.")
            return

        item = selected[0]
        row = int(self.follow_tree.set(item, "_row"))

        updates = {
            "status": self.update_vars["status"].get().strip(),
            "art_kontaktaufnahme": self.update_vars["art_kontaktaufnahme"].get().strip(),
            "ergebnis_nachfrage": self.update_vars["ergebnis_nachfrage"].get().strip(),
            "naechster_schritt": self.update_vars["naechster_schritt"].get().strip(),
            "endergebnis": self.update_vars["endergebnis"].get().strip(),
            "letzte_kontaktaufnahme": date.today(),
            "status_datum": date.today(),
        }

        self.repo.update_record(row=row, updates=updates)
        self.refresh_data()
        messagebox.showinfo("Aktualisiert", "Bewerbung wurde aktualisiert.")

    def _reset_form(self) -> None:
        for key, var in self.form_vars.items():
            if key in {"quelle", "art_bewerbung", "unterlagen", "status", "prioritaet"}:
                choices = getattr(self.lookups, key)
                var.set(choices[0] if choices else "")
            else:
                var.set("")

    def refresh_data(self) -> None:
        self.records = self.repo.load_records()
        self.follow_up_records, self.archive_records = self.repo.categorize_records(self.records)
        self._render_follow_up_records()
        self._render_archive_records()

    def _render_follow_up_records(self) -> None:
        for item in self.follow_tree.get_children():
            self.follow_tree.delete(item)

        self.follow_tree["displaycolumns"] = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt")
        self.follow_tree["columns"] = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt", "_row")

        for rec in self.follow_up_records:
            self.follow_tree.insert(
                "",
                "end",
                values=(
                    rec.bewerbungs_id,
                    rec.unternehmen,
                    rec.position,
                    rec.status,
                    rec.erinnerungsdatum.strftime("%d.%m.%Y") if rec.erinnerungsdatum else "",
                    rec.prioritaet,
                    rec.naechster_schritt,
                    rec.row,
                ),
            )

    def _render_archive_records(self) -> None:
        for item in self.archive_tree.get_children():
            self.archive_tree.delete(item)

        for rec in self.archive_records:
            self.archive_tree.insert(
                "",
                "end",
                values=(
                    rec.bewerbungs_id,
                    rec.unternehmen,
                    rec.position,
                    rec.status,
                    rec.endergebnis,
                    rec.status_datum.strftime("%d.%m.%Y") if rec.status_datum else "",
                ),
            )
