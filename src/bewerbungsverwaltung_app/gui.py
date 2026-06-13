from __future__ import annotations

import os
import tkinter as tk
import webbrowser
from datetime import date
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from .excel_repository import ExcelRepository
from .models import ApplicationRecord, LookupValues
from .utils import as_date


class BewerbungsverwaltungApp:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self.app_dir = workbook_path.parent
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
        self.update_widgets: dict[str, tk.Widget] = {}

        self._build_ui()
        self._reset_form()
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

        self._build_documentation_buttons(header)
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
        self.tab_new.columnconfigure(0, weight=0)
        self.tab_new.columnconfigure(1, weight=1)

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
            ("datum_bewerbung", "Datum Bewerbung", "date"),
            ("unterlagen", "Versendete Unterlagen", "combo", self.lookups.unterlagen),
            ("status", "Aktueller Status *", "combo", self.lookups.status),
            ("status_datum", "Status-Datum", "date"),
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
            ttk.Label(self.tab_new, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=6)
            var = tk.StringVar()
            self.form_vars[key] = var

            if widget_type == "combo":
                widget = ttk.Combobox(self.tab_new, textvariable=var, values=values or [], state="readonly", width=44)
            elif widget_type == "date":
                widget = self._build_date_input(self.tab_new, var, field_key=key)
            else:
                widget = ttk.Entry(self.tab_new, textvariable=var, width=47)

            widget.grid(row=row, column=1, sticky="ew", pady=6)

        buttons = ttk.Frame(self.tab_new)
        buttons.grid(row=len(fields) + 2, column=0, columnspan=2, pady=(18, 0), sticky="w")

        ttk.Button(buttons, text="Speichern", command=self._save_new_record, width=18).pack(side="left", padx=(0, 10))
        ttk.Button(buttons, text="Formular leeren", command=self._reset_form, width=18).pack(side="left")

    def _build_documentation_buttons(self, parent: ttk.Frame) -> None:
        docs_frame = ttk.Frame(parent)
        docs_frame.pack(side="right", padx=(0, 12))
        ttk.Button(
            docs_frame,
            text="Anwender-Doku",
            command=lambda: self._open_documentation(Path("docs") / "DOKUMENTATION_ANWENDER.md"),
            width=16,
        ).pack(side="left", padx=4)
        ttk.Button(
            docs_frame,
            text="FAQ",
            command=lambda: self._open_documentation(Path("docs") / "FAQ.md"),
            width=10,
        ).pack(side="left", padx=4)
        ttk.Button(
            docs_frame,
            text="Technik-Doku",
            command=lambda: self._open_documentation(Path("docs") / "DOKUMENTATION_TECHNIK.md"),
            width=16,
        ).pack(side="left", padx=4)

    def _build_date_input(self, parent: ttk.Frame, var: tk.StringVar, field_key: str) -> ttk.Frame:
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=0)
        entry = ttk.Entry(frame, textvariable=var, width=16, state="readonly", cursor="hand2")
        entry.grid(row=0, column=0, sticky="w")
        entry.bind("<Button-1>", lambda _event: self._open_date_picker(field_key))
        entry.bind("<Return>", lambda _event: self._open_date_picker(field_key))
        entry.bind("<space>", lambda _event: self._open_date_picker(field_key))
        ttk.Button(
            frame,
            text="Wählen",
            command=lambda: self._open_date_picker(field_key),
            width=10,
        ).grid(row=0, column=1, padx=(8, 0), sticky="w")
        ttk.Button(
            frame,
            text="Heute",
            command=lambda: self._set_form_date(field_key, date.today()),
            width=8,
        ).grid(row=0, column=2, padx=(8, 0), sticky="w")
        ttk.Button(frame, text="Leeren", command=lambda: var.set(""), width=8).grid(row=0, column=3, padx=(8, 0), sticky="w")
        return frame

    def _set_form_date(self, field_key: str, selected_date: date) -> None:
        formatted = selected_date.strftime("%d.%m.%Y")
        self.form_vars[field_key].set(formatted)

        if field_key == "datum_bewerbung" and not self.form_vars["status_datum"].get().strip():
            self.form_vars["status_datum"].set(formatted)

    def _open_date_picker(self, field_key: str) -> None:
        current_value = as_date(self.form_vars[field_key].get().strip()) or date.today()

        dialog = tk.Toplevel(self.root)
        dialog.title("Datum auswählen")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Datum auswählen", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        day_values = [f"{day:02d}" for day in range(1, 32)]
        month_values = [f"{month:02d}" for month in range(1, 13)]
        year_values = [str(year) for year in range(current_value.year - 5, current_value.year + 6)]

        day_var = tk.StringVar(value=f"{current_value.day:02d}")
        month_var = tk.StringVar(value=f"{current_value.month:02d}")
        year_var = tk.StringVar(value=str(current_value.year))

        ttk.Combobox(container, textvariable=day_var, values=day_values, state="readonly", width=4).grid(
            row=1, column=0, sticky="w", padx=(0, 8)
        )
        ttk.Combobox(container, textvariable=month_var, values=month_values, state="readonly", width=4).grid(
            row=1, column=1, sticky="w", padx=(0, 8)
        )
        ttk.Combobox(container, textvariable=year_var, values=year_values, state="readonly", width=7).grid(
            row=1, column=2, sticky="w"
        )

        def apply_date() -> None:
            try:
                selected = date(int(year_var.get()), int(month_var.get()), int(day_var.get()))
            except ValueError:
                messagebox.showerror("Ungültiges Datum", "Bitte ein gültiges Datum auswählen.", parent=dialog)
                return

            self._set_form_date(field_key, selected)
            dialog.destroy()

        buttons = ttk.Frame(container)
        buttons.grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Button(buttons, text="Übernehmen", command=apply_date).pack(side="left", padx=(0, 8))
        ttk.Button(
            buttons,
            text="Heute",
            command=lambda: (
                day_var.set(f"{date.today().day:02d}"),
                month_var.set(f"{date.today().month:02d}"),
                year_var.set(str(date.today().year)),
            ),
        ).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Abbrechen", command=dialog.destroy).pack(side="left")

        dialog.bind("<Return>", lambda _event: apply_date())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _open_documentation(self, relative_path: Path) -> None:
        target = self.app_dir / relative_path
        if not target.exists():
            messagebox.showwarning("Datei fehlt", f"Dokumentation nicht gefunden:\n{target}")
            return

        try:
            os.startfile(target)  # type: ignore[attr-defined]
        except AttributeError:
            webbrowser.open(target.resolve().as_uri())
        except OSError as exc:
            messagebox.showerror("Öffnen fehlgeschlagen", f"Dokumentation konnte nicht geöffnet werden:\n{exc}")

    def _build_follow_up_tab(self) -> None:
        self.tab_follow_up.columnconfigure(0, weight=1)

        ttk.Label(
            self.tab_follow_up,
            text="Fällige Bewerbungen: Erinnerungsdatum erreicht oder Follow-up überfällig.",
        ).pack(anchor="w", pady=(0, 8))

        columns = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt")
        self.follow_tree = ttk.Treeview(self.tab_follow_up, columns=columns, show="headings", height=16)
        self.follow_tree.pack(fill="both", expand=True, pady=(0, 10))
        self.follow_tree.bind("<<TreeviewSelect>>", self._on_follow_tree_select)
        self.follow_tree.bind("<Double-1>", self._on_follow_tree_double_click)

        headings = {
            "id": "ID",
            "unternehmen": "Unternehmen",
            "position": "Position",
            "status": "Status",
            "erinnerung": "Erinnerungsdatum",
            "prioritaet": "Priorität",
            "schritt": "Nächster Schritt",
        }
        widths = {"id": 120, "unternehmen": 240, "position": 250, "status": 200, "erinnerung": 150, "prioritaet": 110, "schritt": 320}
        self.follow_headings = headings
        self.follow_min_widths = widths
        self.follow_max_width = 460

        for col in columns:
            self.follow_tree.heading(col, text=headings[col])
            self.follow_tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=True)

        update_box = ttk.LabelFrame(self.tab_follow_up, text="Ausgewählte Bewerbung aktualisieren", padding=10)
        update_box.pack(fill="x")
        update_box.columnconfigure(1, weight=1)

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
            else:
                widget = ttk.Entry(update_box, textvariable=var, width=48)
            self.update_widgets[key] = widget
            widget.grid(row=i, column=1, sticky="ew", pady=4)

        ttk.Button(update_box, text="Änderung speichern", command=self._save_follow_up_update).grid(
            row=len(mapping), column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

    def _on_follow_tree_select(self, _event: tk.Event | None = None) -> None:
        selected = self.follow_tree.selection()
        if not selected:
            self._clear_follow_up_form()
            return

        item = selected[0]
        row = int(self.follow_tree.set(item, "_row"))
        record = next((rec for rec in self.follow_up_records if rec.row == row), None)
        if record is None:
            self._clear_follow_up_form()
            return

        self.update_vars["status"].set(record.status or "")
        self.update_vars["art_kontaktaufnahme"].set(record.art_kontaktaufnahme or "")
        self.update_vars["ergebnis_nachfrage"].set(record.ergebnis_nachfrage or "")
        self.update_vars["naechster_schritt"].set(record.naechster_schritt or "")
        self.update_vars["endergebnis"].set(record.endergebnis or "")

    def _on_follow_tree_double_click(self, _event: tk.Event | None = None) -> None:
        self._on_follow_tree_select()
        first_widget = self.update_widgets.get("status")
        if first_widget is not None:
            first_widget.focus_set()

    def _clear_follow_up_form(self) -> None:
        for var in self.update_vars.values():
            var.set("")

    def _build_archive_tab(self) -> None:
        self.tab_archive.columnconfigure(0, weight=1)

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
        self.archive_headings = headings
        self.archive_min_widths = widths
        self.archive_max_width = 460

        for col in columns:
            self.archive_tree.heading(col, text=headings[col])
            self.archive_tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=True)

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
            if key in {"datum_bewerbung", "status_datum"}:
                var.set(date.today().strftime("%d.%m.%Y"))
            else:
                var.set("")

    def refresh_data(self) -> None:
        self.records = self.repo.load_records()
        self.follow_up_records, self.archive_records = self.repo.categorize_records(self.records)
        self._render_follow_up_records()
        self._render_archive_records()
        self._clear_follow_up_form()

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

        self._autosize_follow_up_columns()

    def _autosize_follow_up_columns(self) -> None:
        font_name = ttk.Style().lookup("Treeview", "font")
        measure_font = tkfont.Font(font=font_name) if font_name else tkfont.nametofont("TkDefaultFont")
        padding = 24

        for col in ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt"):
            heading_text = self.follow_headings.get(col, col)
            required = measure_font.measure(str(heading_text)) + padding

            for item in self.follow_tree.get_children():
                cell_value = self.follow_tree.set(item, col)
                cell_width = measure_font.measure(str(cell_value)) + padding
                if cell_width > required:
                    required = cell_width

            min_width = self.follow_min_widths.get(col, 100)
            width = max(min_width, min(required, self.follow_max_width))
            self.follow_tree.column(col, width=width, minwidth=min_width)

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

        self._autosize_archive_columns()

    def _autosize_archive_columns(self) -> None:
        font_name = ttk.Style().lookup("Treeview", "font")
        measure_font = tkfont.Font(font=font_name) if font_name else tkfont.nametofont("TkDefaultFont")
        padding = 24

        for col in ("id", "unternehmen", "position", "status", "endergebnis", "status_datum"):
            heading_text = self.archive_headings.get(col, col)
            required = measure_font.measure(str(heading_text)) + padding

            for item in self.archive_tree.get_children():
                cell_value = self.archive_tree.set(item, col)
                cell_width = measure_font.measure(str(cell_value)) + padding
                if cell_width > required:
                    required = cell_width

            min_width = self.archive_min_widths.get(col, 100)
            width = max(min_width, min(required, self.archive_max_width))
            self.archive_tree.column(col, width=width, minwidth=min_width)
