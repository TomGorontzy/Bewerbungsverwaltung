from __future__ import annotations

import os
import tkinter as tk
import webbrowser
from datetime import date
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from .constants import ARCHIVE_STATUS
from .excel_repository import ExcelRepository
from .models import ApplicationRecord, LookupValues
from .utils import as_date


class BewerbungsverwaltungApp:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        if workbook_path.parent.name.lower() == "data":
            self.app_dir = workbook_path.parent.parent
        else:
            self.app_dir = workbook_path.parent
        self.repo = ExcelRepository(workbook_path)
        self.repo.ensure_workbook_health()
        self.lookups: LookupValues = self.repo.load_lookups()

        self.root = tk.Tk()
        self.root.title("Bewerbungsverwaltung")
        self.root.geometry("1250x860")
        self.root.minsize(1100, 820)
        self.fullscreen_var = tk.BooleanVar(value=False)
        self._set_startup_fullscreen()

        self.records: list[ApplicationRecord] = []
        self.follow_up_records: list[ApplicationRecord] = []
        self.archive_records: list[ApplicationRecord] = []
        self.active_records: list[ApplicationRecord] = []

        self.form_vars: dict[str, tk.StringVar] = {}
        self.form_widgets: dict[str, tk.Widget] = {}
        self.update_vars: dict[str, tk.StringVar] = {}
        self.update_widgets: dict[str, tk.Widget] = {}
        self.active_update_vars: dict[str, tk.StringVar] = {}
        self.active_update_widgets: dict[str, tk.Widget] = {}

        self.follow_filter_var = tk.StringVar()
        self.archive_filter_var = tk.StringVar()
        self.active_filter_var = tk.StringVar()
        self.filter_debounce_ms = 150
        self.follow_filter_after_id: str | None = None
        self.archive_filter_after_id: str | None = None
        self.active_filter_after_id: str | None = None
        self.follow_sort_column = "erinnerung"
        self.follow_sort_desc = False
        self.archive_sort_column = "status_datum"
        self.archive_sort_desc = True
        self.active_sort_column = "datum"
        self.active_sort_desc = True
        self.follow_overdue_thresholds = [
            (14, "overdue_critical"),
            (7, "overdue_high"),
            (3, "overdue_medium"),
            (1, "overdue_low"),
        ]
        self.follow_overdue_legend = [
            ("overdue_low", "1–2 Tage überfällig", "#fff7cc"),
            ("overdue_medium", "3–6 Tage überfällig", "#ffe7bf"),
            ("overdue_high", "7–13 Tage überfällig", "#ffd5bf"),
            ("overdue_critical", "ab 14 Tagen überfällig", "#ffc7c7"),
        ]
        self.status_var = tk.StringVar(value="Bereit")
        self.reload_button: ttk.Button | None = None

        self._build_ui()
        self._reset_form()
        self.root.after_idle(self.refresh_data)

    def run(self) -> None:
        self.root.mainloop()

    def _set_startup_fullscreen(self) -> None:
        try:
            self.root.state("zoomed")
            return
        except tk.TclError:
            pass

        try:
            self.root.attributes("-zoomed", True)
            return
        except tk.TclError:
            pass

        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")

    def _set_busy_cursor(self, busy: bool) -> None:
        self.root.config(cursor="watch" if busy else "")
        self.root.update_idletasks()

    def _set_refresh_state(self, busy: bool, status_text: str | None = None) -> None:
        if self.reload_button is not None:
            self.reload_button.configure(state="disabled" if busy else "normal")
        self.status_var.set(status_text or ("Daten werden geladen..." if busy else "Bereit"))
        self._set_busy_cursor(busy)

    def _build_ui(self) -> None:
        self._configure_treeview_styles()
        self._configure_window_mode_controls()

        header = ttk.Frame(self.root, padding=12)
        header.pack(fill="x")

        ttk.Label(
            header,
            text="Bewerbungsverwaltung – geführte Pflege",
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left")

        self._build_documentation_buttons(header)
        self.reload_button = ttk.Button(header, text="Daten neu laden", command=self.refresh_data)
        self.reload_button.pack(side="right", padx=4)

        status_bar = ttk.Frame(self.root, padding=(12, 0, 12, 10))
        status_bar.pack(fill="x")
        ttk.Label(status_bar, textvariable=self.status_var, foreground="#666666").pack(anchor="w")

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab_new = ttk.Frame(notebook, padding=12)
        self.tab_active = ttk.Frame(notebook, padding=12)
        self.tab_follow_up = ttk.Frame(notebook, padding=12)
        self.tab_archive = ttk.Frame(notebook, padding=12)

        notebook.add(self.tab_new, text="Neue Bewerbungen")
        notebook.add(self.tab_active, text="Laufende Bewerbungen")
        notebook.add(self.tab_follow_up, text="Nachzufassende Bewerbungen")
        notebook.add(self.tab_archive, text="Erledigte Bewerbungen (Archiv)")

        self._build_new_tab()
        self._build_active_tab()
        self._build_follow_up_tab()
        self._build_archive_tab()

    def _configure_window_mode_controls(self) -> None:
        menubar = tk.Menu(self.root)
        view_menu = tk.Menu(menubar, tearoff=False)
        view_menu.add_checkbutton(
            label="Echter Vollbildmodus (F11)",
            variable=self.fullscreen_var,
            command=self._toggle_true_fullscreen,
        )
        view_menu.add_command(label="Vollbild verlassen (Esc)", command=lambda: self._set_true_fullscreen(False))
        menubar.add_cascade(label="Ansicht", menu=view_menu)
        self.root.config(menu=menubar)

        self.root.bind("<F11>", self._toggle_true_fullscreen)
        self.root.bind("<Escape>", self._exit_true_fullscreen)

    def _toggle_true_fullscreen(self, _event: tk.Event | None = None) -> str:
        self._set_true_fullscreen(not self.fullscreen_var.get())
        return "break"

    def _exit_true_fullscreen(self, _event: tk.Event | None = None) -> str:
        if self.fullscreen_var.get():
            self._set_true_fullscreen(False)
            return "break"
        return ""

    def _set_true_fullscreen(self, enabled: bool) -> None:
        self.root.attributes("-fullscreen", enabled)
        self.fullscreen_var.set(enabled)
        if not enabled:
            self._set_startup_fullscreen()

    def _configure_treeview_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(
            "TNotebook",
            background="#dde3ea",
            borderwidth=0,
            tabmargins=[2, 4, 0, 0],
        )
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10),
            foreground="#1a1a1a",
            background="#b8c4cf",
            padding=[14, 6],
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#2e75b6"), ("active", "#7fa8c9")],
            foreground=[("selected", "#ffffff"), ("active", "#1a1a1a")],
            font=[("selected", ("Segoe UI", 10, "bold"))],
        )

        style.configure("Readable.Treeview", rowheight=24)
        style.configure(
            "Readable.Treeview.Heading",
            foreground="#1f1f1f",
            background="#f0f0f0",
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Readable.Treeview.Heading",
            foreground=[("active", "#000000"), ("disabled", "#777777")],
            background=[("active", "#e6e6e6")],
        )

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
            ("naechster_schritt", "Nächster Schritt", "combo", [""] + self.lookups.naechster_schritt),
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

            self.form_widgets[key] = widget
            widget.grid(row=row, column=1, sticky="ew", pady=6)

        buttons = ttk.Frame(self.tab_new)
        buttons.grid(row=len(fields) + 2, column=0, columnspan=2, pady=(18, 0), sticky="w")

        ttk.Button(buttons, text="Speichern", command=self._save_new_record, width=18).pack(side="left", padx=(0, 10))
        ttk.Button(buttons, text="Formular leeren", command=self._reset_form, width=18).pack(side="left")
        ttk.Button(
            buttons,
            text="Schritt-Katalog verwalten",
            command=self._open_naechster_schritt_catalog_dialog,
            width=28,
        ).pack(side="left", padx=(10, 0))

    def _build_documentation_buttons(self, parent: ttk.Frame) -> None:
        docs_frame = ttk.Frame(parent)
        docs_frame.pack(side="right", padx=(0, 12))
        header_button_width = 16
        ttk.Button(
            docs_frame,
            text="Excel-Datei",
            command=self._open_release_excel_file,
            width=header_button_width,
        ).pack(side="left", padx=4)
        ttk.Button(
            docs_frame,
            text="Anwender-Doku",
            command=lambda: self._open_documentation(Path("docs") / "DOKUMENTATION_ANWENDER.md"),
            width=header_button_width,
        ).pack(side="left", padx=4)
        ttk.Button(
            docs_frame,
            text="FAQ",
            command=lambda: self._open_documentation(Path("docs") / "FAQ.md"),
            width=header_button_width,
        ).pack(side="left", padx=4)
        ttk.Button(
            docs_frame,
            text="Technik-Doku",
            command=lambda: self._open_documentation(Path("docs") / "DOKUMENTATION_TECHNIK.md"),
            width=header_button_width,
        ).pack(side="left", padx=4)

    def _open_release_excel_file(self) -> None:
        target = self.workbook_path
        if not target.exists():
            messagebox.showwarning("Datei fehlt", f"Excel-Datei nicht gefunden:\n{target}")
            return

        try:
            os.startfile(target)  # type: ignore[attr-defined]
        except AttributeError:
            webbrowser.open(target.resolve().as_uri())
        except OSError as exc:
            messagebox.showerror("Öffnen fehlgeschlagen", f"Excel-Datei konnte nicht geöffnet werden:\n{exc}")

    def _build_date_input(self, parent: ttk.Frame, var: tk.StringVar, field_key: str) -> ttk.Frame:
        return self._build_date_input_with_callback(
            parent=parent,
            var=var,
            on_select=lambda selected_date: self._set_form_date(field_key, selected_date),
        )

    def _build_date_input_with_callback(
        self,
        parent: ttk.Frame,
        var: tk.StringVar,
        on_select,
    ) -> ttk.Frame:
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=0)
        entry = ttk.Entry(frame, textvariable=var, width=16, state="readonly", cursor="hand2")
        entry.grid(row=0, column=0, sticky="w")
        entry.bind("<Button-1>", lambda _event: self._open_date_picker_for_var(var, on_select))
        entry.bind("<Return>", lambda _event: self._open_date_picker_for_var(var, on_select))
        entry.bind("<space>", lambda _event: self._open_date_picker_for_var(var, on_select))
        ttk.Button(
            frame,
            text="Wählen",
            command=lambda: self._open_date_picker_for_var(var, on_select),
            width=10,
        ).grid(row=0, column=1, padx=(8, 0), sticky="w")
        ttk.Button(
            frame,
            text="Heute",
            command=lambda: on_select(date.today()),
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
        self._open_date_picker_for_var(
            self.form_vars[field_key],
            lambda selected_date: self._set_form_date(field_key, selected_date),
        )

    def _open_date_picker_for_var(self, var: tk.StringVar, on_select) -> None:
        current_value = as_date(var.get().strip()) or date.today()

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

            on_select(selected)
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

    def _build_active_tab(self) -> None:
        self.tab_active.columnconfigure(0, weight=1)

        ttk.Label(
            self.tab_active,
            text="Alle laufenden Bewerbungen ohne Endergebnis, sortiert nach Datum der Bewerbung.",
        ).pack(anchor="w", pady=(0, 8))

        active_filter_box = ttk.Frame(self.tab_active)
        active_filter_box.pack(fill="x", pady=(0, 8))
        ttk.Label(active_filter_box, text="Filter:").pack(side="left", padx=(0, 8))
        active_filter_entry = ttk.Entry(active_filter_box, textvariable=self.active_filter_var, width=48)
        active_filter_entry.pack(side="left", padx=(0, 8))
        active_filter_entry.bind("<KeyRelease>", lambda _event: self._schedule_active_filter_render())
        ttk.Button(active_filter_box, text="Zurücksetzen", command=self._reset_active_filter, width=12).pack(side="left")

        columns = ("id", "unternehmen", "position", "datum", "status", "prioritaet", "schritt")
        active_tree_frame = ttk.Frame(self.tab_active)
        active_tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.active_tree = ttk.Treeview(
            active_tree_frame,
            columns=columns,
            show="headings",
            height=10,
            style="Readable.Treeview",
        )
        active_tree_vscroll = ttk.Scrollbar(active_tree_frame, orient="vertical", command=self.active_tree.yview)
        active_tree_hscroll = ttk.Scrollbar(active_tree_frame, orient="horizontal", command=self.active_tree.xview)
        self.active_tree.configure(yscrollcommand=active_tree_vscroll.set, xscrollcommand=active_tree_hscroll.set)

        self.active_tree.grid(row=0, column=0, sticky="nsew")
        active_tree_vscroll.grid(row=0, column=1, sticky="ns")
        active_tree_hscroll.grid(row=1, column=0, sticky="ew")
        active_tree_frame.columnconfigure(0, weight=1)
        active_tree_frame.rowconfigure(0, weight=1)
        self.active_tree.bind("<<TreeviewSelect>>", self._on_active_tree_select)
        self.active_tree.bind("<Double-1>", self._on_active_tree_double_click)

        headings = {
            "id": "ID",
            "unternehmen": "Unternehmen",
            "position": "Position",
            "datum": "Datum Bewerbung",
            "status": "Status",
            "prioritaet": "Priorität",
            "schritt": "Nächster Schritt",
        }
        widths = {"id": 120, "unternehmen": 230, "position": 240, "datum": 130, "status": 200, "prioritaet": 110, "schritt": 280}
        self.active_headings = headings
        self.active_min_widths = widths
        self.active_max_width = 460

        for col in columns:
            self.active_tree.heading(col, text=headings[col], command=lambda c=col: self._toggle_active_sort(c))
            self.active_tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=True)

        update_box = ttk.LabelFrame(self.tab_active, text="Ausgewählte Bewerbung aktualisieren", padding=10)
        update_box.pack(fill="x")
        update_box.columnconfigure(1, weight=1)

        mapping = [
            ("status", "Neuer Status", self.lookups.status),
            ("art_kontaktaufnahme", "Kontaktart", self.lookups.art_kontaktaufnahme),
            ("ergebnis_nachfrage", "Ergebnis Nachfrage", None),
            ("naechster_schritt", "Nächster Schritt", [""] + self.lookups.naechster_schritt),
            ("endergebnis", "Endergebnis", [""] + self.lookups.endergebnis),
        ]

        for i, (key, label, values) in enumerate(mapping):
            ttk.Label(update_box, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=4)
            var = tk.StringVar()
            self.active_update_vars[key] = var
            if values is not None:
                widget = ttk.Combobox(update_box, textvariable=var, values=values, state="readonly", width=45)
            else:
                widget = ttk.Entry(update_box, textvariable=var, width=48)
            self.active_update_widgets[key] = widget
            widget.grid(row=i, column=1, sticky="ew", pady=4)

        date_row = len(mapping)
        ttk.Label(update_box, text="Datum Bewerbung").grid(row=date_row, column=0, sticky="w", padx=(0, 8), pady=4)
        datum_var = tk.StringVar()
        self.active_update_vars["datum_bewerbung"] = datum_var
        datum_widget = self._build_date_input_with_callback(
            parent=update_box,
            var=datum_var,
            on_select=lambda selected_date: datum_var.set(selected_date.strftime("%d.%m.%Y")),
        )
        self.active_update_widgets["datum_bewerbung"] = datum_widget
        datum_widget.grid(row=date_row, column=1, sticky="w", pady=4)

        ttk.Button(update_box, text="Änderung speichern", command=self._save_active_update).grid(
            row=date_row + 1, column=0, columnspan=2, sticky="w", pady=(10, 0)
        )

    def _on_active_tree_select(self, _event: tk.Event | None = None) -> None:
        selected = self.active_tree.selection()
        if not selected:
            self._clear_active_form()
            return

        item = selected[0]
        row = int(self.active_tree.set(item, "_row"))
        record = next((rec for rec in self.active_records if rec.row == row), None)
        if record is None:
            self._clear_active_form()
            return

        self.active_update_vars["status"].set(record.status or "")
        self.active_update_vars["art_kontaktaufnahme"].set(record.art_kontaktaufnahme or "")
        self.active_update_vars["ergebnis_nachfrage"].set(record.ergebnis_nachfrage or "")
        self.active_update_vars["naechster_schritt"].set(record.naechster_schritt or "")
        self.active_update_vars["endergebnis"].set(record.endergebnis or "")
        self.active_update_vars["datum_bewerbung"].set(
            record.datum_bewerbung.strftime("%d.%m.%Y") if record.datum_bewerbung else ""
        )

    def _on_active_tree_double_click(self, _event: tk.Event | None = None) -> None:
        self._on_active_tree_select()
        first_widget = self.active_update_widgets.get("status")
        if first_widget is not None:
            first_widget.focus_set()

    def _clear_active_form(self) -> None:
        for var in self.active_update_vars.values():
            var.set("")

    def _reset_active_filter(self) -> None:
        if self.active_filter_after_id is not None:
            self.root.after_cancel(self.active_filter_after_id)
            self.active_filter_after_id = None
        self.active_filter_var.set("")
        self._render_active_records()

    def _schedule_active_filter_render(self) -> None:
        if self.active_filter_after_id is not None:
            self.root.after_cancel(self.active_filter_after_id)
        self.active_filter_after_id = self.root.after(self.filter_debounce_ms, self._apply_active_filter_render)

    def _apply_active_filter_render(self) -> None:
        self.active_filter_after_id = None
        self._render_active_records()

    def _toggle_active_sort(self, column: str) -> None:
        if self.active_sort_column == column:
            self.active_sort_desc = not self.active_sort_desc
        else:
            self.active_sort_column = column
            self.active_sort_desc = False
        self._render_active_records()

    def _active_sort_key(self, rec: ApplicationRecord, column: str):
        if column == "datum":
            return rec.datum_bewerbung or date.min
        if column == "id":
            return rec.bewerbungs_id.lower()
        if column == "unternehmen":
            return rec.unternehmen.lower()
        if column == "position":
            return rec.position.lower()
        if column == "status":
            return rec.status.lower()
        if column == "prioritaet":
            return rec.prioritaet.lower()
        if column == "schritt":
            return rec.naechster_schritt.lower()
        return ""

    def _save_active_update(self) -> None:
        selected = self.active_tree.selection()
        if not selected:
            messagebox.showwarning("Auswahl fehlt", "Bitte zuerst eine Bewerbung in der Liste auswählen.")
            return

        item = selected[0]
        row = int(self.active_tree.set(item, "_row"))

        datum_bewerbung_text = self.active_update_vars["datum_bewerbung"].get().strip()
        if datum_bewerbung_text:
            datum_bewerbung = as_date(datum_bewerbung_text)
            if datum_bewerbung is None:
                messagebox.showwarning("Ungültiges Datum", "Bitte ein gültiges Datum im Format TT.MM.JJJJ auswählen.")
                return
        else:
            datum_bewerbung = ""

        updates = {
            "status": self.active_update_vars["status"].get().strip(),
            "art_kontaktaufnahme": self.active_update_vars["art_kontaktaufnahme"].get().strip(),
            "ergebnis_nachfrage": self.active_update_vars["ergebnis_nachfrage"].get().strip(),
            "naechster_schritt": self.active_update_vars["naechster_schritt"].get().strip(),
            "endergebnis": self.active_update_vars["endergebnis"].get().strip(),
            "datum_bewerbung": datum_bewerbung,
            "letzte_kontaktaufnahme": date.today(),
            "status_datum": date.today(),
        }

        self.repo.update_record(row=row, updates=updates)
        self.refresh_data()
        messagebox.showinfo("Aktualisiert", "Bewerbung wurde aktualisiert.")

    def _render_active_records(self) -> None:
        for item in self.active_tree.get_children():
            self.active_tree.delete(item)

        self.active_tree["displaycolumns"] = ("id", "unternehmen", "position", "datum", "status", "prioritaet", "schritt")
        self.active_tree["columns"] = ("id", "unternehmen", "position", "datum", "status", "prioritaet", "schritt", "_row")

        filter_text = self.active_filter_var.get().strip().lower()
        records = self.active_records
        if filter_text:
            records = [
                rec
                for rec in records
                if filter_text in " ".join(
                    [
                        rec.bewerbungs_id,
                        rec.unternehmen,
                        rec.position,
                        rec.status,
                        rec.prioritaet,
                        rec.naechster_schritt,
                        rec.datum_bewerbung.strftime("%d.%m.%Y") if rec.datum_bewerbung else "",
                    ]
                ).lower()
            ]

        records = sorted(records, key=lambda rec: self._active_sort_key(rec, self.active_sort_column), reverse=self.active_sort_desc)

        for col, base in self.active_headings.items():
            marker = ""
            if col == self.active_sort_column:
                marker = " ↓" if self.active_sort_desc else " ↑"
            self.active_tree.heading(col, text=f"{base}{marker}", command=lambda c=col: self._toggle_active_sort(c))

        for rec in records:
            self.active_tree.insert(
                "",
                "end",
                values=(
                    rec.bewerbungs_id,
                    rec.unternehmen,
                    rec.position,
                    rec.datum_bewerbung.strftime("%d.%m.%Y") if rec.datum_bewerbung else "",
                    rec.status,
                    rec.prioritaet,
                    rec.naechster_schritt,
                    rec.row,
                ),
            )

        self._autosize_active_columns()

    def _autosize_active_columns(self) -> None:
        font_name = ttk.Style().lookup("Treeview", "font")
        measure_font = tkfont.Font(font=font_name) if font_name else tkfont.nametofont("TkDefaultFont")
        padding = 24

        for col in ("id", "unternehmen", "position", "datum", "status", "prioritaet", "schritt"):
            heading_text = self.active_headings.get(col, col)
            required = measure_font.measure(str(heading_text)) + padding

            for item in self.active_tree.get_children():
                cell_value = self.active_tree.set(item, col)
                cell_width = measure_font.measure(str(cell_value)) + padding
                if cell_width > required:
                    required = cell_width

            min_width = self.active_min_widths.get(col, 100)
            width = max(min_width, min(required, self.active_max_width))
            self.active_tree.column(col, width=width, minwidth=min_width)

    def _build_follow_up_tab(self) -> None:
        self.tab_follow_up.columnconfigure(0, weight=1)

        ttk.Label(
            self.tab_follow_up,
            text="Fällige Bewerbungen: Erinnerungsdatum erreicht oder Follow-up überfällig.",
        ).pack(anchor="w", pady=(0, 8))

        follow_filter_box = ttk.Frame(self.tab_follow_up)
        follow_filter_box.pack(fill="x", pady=(0, 8))
        ttk.Label(follow_filter_box, text="Filter:").pack(side="left", padx=(0, 8))
        follow_filter_entry = ttk.Entry(follow_filter_box, textvariable=self.follow_filter_var, width=48)
        follow_filter_entry.pack(side="left", padx=(0, 8))
        follow_filter_entry.bind("<KeyRelease>", lambda _event: self._schedule_follow_filter_render())
        ttk.Button(follow_filter_box, text="Zurücksetzen", command=self._reset_follow_filter, width=12).pack(side="left")

        self._build_follow_up_legend_bar()

        columns = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt")
        follow_tree_frame = ttk.Frame(self.tab_follow_up)
        follow_tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.follow_tree = ttk.Treeview(
            follow_tree_frame,
            columns=columns,
            show="headings",
            height=10,
            style="Readable.Treeview",
        )
        follow_tree_vscroll = ttk.Scrollbar(follow_tree_frame, orient="vertical", command=self.follow_tree.yview)
        follow_tree_hscroll = ttk.Scrollbar(follow_tree_frame, orient="horizontal", command=self.follow_tree.xview)
        self.follow_tree.configure(yscrollcommand=follow_tree_vscroll.set, xscrollcommand=follow_tree_hscroll.set)

        self.follow_tree.grid(row=0, column=0, sticky="nsew")
        follow_tree_vscroll.grid(row=0, column=1, sticky="ns")
        follow_tree_hscroll.grid(row=1, column=0, sticky="ew")
        follow_tree_frame.columnconfigure(0, weight=1)
        follow_tree_frame.rowconfigure(0, weight=1)
        for tag_name, _label, color in self.follow_overdue_legend:
            self.follow_tree.tag_configure(tag_name, background=color)
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
            self.follow_tree.heading(col, text=headings[col], command=lambda c=col: self._toggle_follow_sort(c))
            self.follow_tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=True)

        update_box = ttk.LabelFrame(self.tab_follow_up, text="Ausgewählte Bewerbung aktualisieren", padding=10)
        update_box.pack(fill="x")
        update_box.columnconfigure(1, weight=1)

        mapping = [
            ("status", "Neuer Status", self.lookups.status),
            ("art_kontaktaufnahme", "Kontaktart", self.lookups.art_kontaktaufnahme),
            ("ergebnis_nachfrage", "Ergebnis Nachfrage", None),
            ("naechster_schritt", "Nächster Schritt", [""] + self.lookups.naechster_schritt),
            ("endergebnis", "Endergebnis", [""] + self.lookups.endergebnis),
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

    def _build_follow_up_legend_bar(self) -> None:
        legend_bar = tk.Frame(self.tab_follow_up, bd=1, relief="groove", bg="#f4f6f8", padx=8, pady=6)
        legend_bar.pack(fill="x", pady=(0, 8))

        tk.Label(
            legend_bar,
            text="Statusleiste – Überfälligkeit:",
            bg="#f4f6f8",
            fg="#2c2c2c",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=(0, 10))

        for _tag_name, label, color in self.follow_overdue_legend:
            tk.Label(
                legend_bar,
                text=f"  {label}  ",
                bg=color,
                fg="#1f1f1f",
                bd=1,
                relief="solid",
                padx=6,
                pady=2,
            ).pack(side="left", padx=(0, 6))

        tk.Label(
            legend_bar,
            text="kein Farbhintergrund = noch nicht überfällig",
            bg="#f4f6f8",
            fg="#555555",
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(6, 0))

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

    def _reset_follow_filter(self) -> None:
        if self.follow_filter_after_id is not None:
            self.root.after_cancel(self.follow_filter_after_id)
            self.follow_filter_after_id = None
        self.follow_filter_var.set("")
        self._render_follow_up_records()

    def _schedule_follow_filter_render(self) -> None:
        if self.follow_filter_after_id is not None:
            self.root.after_cancel(self.follow_filter_after_id)
        self.follow_filter_after_id = self.root.after(self.filter_debounce_ms, self._apply_follow_filter_render)

    def _apply_follow_filter_render(self) -> None:
        self.follow_filter_after_id = None
        self._render_follow_up_records()

    def _toggle_follow_sort(self, column: str) -> None:
        if self.follow_sort_column == column:
            self.follow_sort_desc = not self.follow_sort_desc
        else:
            self.follow_sort_column = column
            self.follow_sort_desc = False
        self._render_follow_up_records()

    def _follow_sort_key(self, rec: ApplicationRecord, column: str):
        if column == "erinnerung":
            return rec.erinnerungsdatum or date.max
        if column == "id":
            return rec.bewerbungs_id.lower()
        if column == "unternehmen":
            return rec.unternehmen.lower()
        if column == "position":
            return rec.position.lower()
        if column == "status":
            return rec.status.lower()
        if column == "prioritaet":
            return rec.prioritaet.lower()
        if column == "schritt":
            return rec.naechster_schritt.lower()
        return ""

    def _get_overdue_days(self, rec: ApplicationRecord) -> int:
        if rec.erinnerungsdatum is None:
            return 0
        return max((date.today() - rec.erinnerungsdatum).days, 0)

    def _get_follow_row_tag(self, rec: ApplicationRecord) -> str:
        overdue_days = self._get_overdue_days(rec)
        for minimum_days, tag in self.follow_overdue_thresholds:
            if overdue_days >= minimum_days:
                return tag
        return ""

    def _build_archive_tab(self) -> None:
        self.tab_archive.columnconfigure(0, weight=1)

        ttk.Label(
            self.tab_archive,
            text="Archiv: abgeschlossene Bewerbungen (Absage/Zusage oder mit Endergebnis).",
        ).pack(anchor="w", pady=(0, 8))

        archive_filter_box = ttk.Frame(self.tab_archive)
        archive_filter_box.pack(fill="x", pady=(0, 8))
        ttk.Label(archive_filter_box, text="Filter:").pack(side="left", padx=(0, 8))
        archive_filter_entry = ttk.Entry(archive_filter_box, textvariable=self.archive_filter_var, width=48)
        archive_filter_entry.pack(side="left", padx=(0, 8))
        archive_filter_entry.bind("<KeyRelease>", lambda _event: self._schedule_archive_filter_render())
        ttk.Button(archive_filter_box, text="Zurücksetzen", command=self._reset_archive_filter, width=12).pack(side="left")

        columns = ("id", "unternehmen", "position", "status", "endergebnis", "status_datum")
        archive_tree_frame = ttk.Frame(self.tab_archive)
        archive_tree_frame.pack(fill="both", expand=True)

        self.archive_tree = ttk.Treeview(
            archive_tree_frame,
            columns=columns,
            show="headings",
            height=10,
            style="Readable.Treeview",
        )
        archive_tree_vscroll = ttk.Scrollbar(archive_tree_frame, orient="vertical", command=self.archive_tree.yview)
        archive_tree_hscroll = ttk.Scrollbar(archive_tree_frame, orient="horizontal", command=self.archive_tree.xview)
        self.archive_tree.configure(yscrollcommand=archive_tree_vscroll.set, xscrollcommand=archive_tree_hscroll.set)

        self.archive_tree.grid(row=0, column=0, sticky="nsew")
        archive_tree_vscroll.grid(row=0, column=1, sticky="ns")
        archive_tree_hscroll.grid(row=1, column=0, sticky="ew")
        archive_tree_frame.columnconfigure(0, weight=1)
        archive_tree_frame.rowconfigure(0, weight=1)

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
            self.archive_tree.heading(col, text=headings[col], command=lambda c=col: self._toggle_archive_sort(c))
            self.archive_tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=True)

    def _reset_archive_filter(self) -> None:
        if self.archive_filter_after_id is not None:
            self.root.after_cancel(self.archive_filter_after_id)
            self.archive_filter_after_id = None
        self.archive_filter_var.set("")
        self._render_archive_records()

    def _schedule_archive_filter_render(self) -> None:
        if self.archive_filter_after_id is not None:
            self.root.after_cancel(self.archive_filter_after_id)
        self.archive_filter_after_id = self.root.after(self.filter_debounce_ms, self._apply_archive_filter_render)

    def _apply_archive_filter_render(self) -> None:
        self.archive_filter_after_id = None
        self._render_archive_records()

    def _toggle_archive_sort(self, column: str) -> None:
        if self.archive_sort_column == column:
            self.archive_sort_desc = not self.archive_sort_desc
        else:
            self.archive_sort_column = column
            self.archive_sort_desc = False
        self._render_archive_records()

    def _archive_sort_key(self, rec: ApplicationRecord, column: str):
        if column == "status_datum":
            return rec.status_datum or date.min
        if column == "id":
            return rec.bewerbungs_id.lower()
        if column == "unternehmen":
            return rec.unternehmen.lower()
        if column == "position":
            return rec.position.lower()
        if column == "status":
            return rec.status.lower()
        if column == "endergebnis":
            return rec.endergebnis.lower()
        return ""

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

    def _open_naechster_schritt_catalog_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Katalog 'Nächster Schritt' verwalten")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("640x480")
        dialog.minsize(560, 420)

        container = ttk.Frame(dialog, padding=12)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text="Pflege den Dropdown-Katalog für 'Nächster Schritt' (GUI + Excel).",
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        list_frame = ttk.Frame(container)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        listbox = tk.Listbox(list_frame, selectmode="browse", activestyle="dotbox", exportselection=False)
        listbox.grid(row=0, column=0, sticky="nsew")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        listbox.configure(yscrollcommand=list_scroll.set)

        for item in self.lookups.naechster_schritt:
            listbox.insert("end", item)

        input_frame = ttk.Frame(container)
        input_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        input_frame.columnconfigure(1, weight=1)

        new_value_var = tk.StringVar()
        ttk.Label(input_frame, text="Neuer Schritt:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        new_value_entry = ttk.Entry(input_frame, textvariable=new_value_var)
        new_value_entry.grid(row=0, column=1, sticky="ew")

        button_row = ttk.Frame(container)
        button_row.grid(row=3, column=0, sticky="w", pady=(10, 0))

        def add_item() -> None:
            value = new_value_var.get().strip()
            if not value:
                return
            existing = [listbox.get(i).strip().casefold() for i in range(listbox.size())]
            if value.casefold() in existing:
                messagebox.showwarning("Bereits vorhanden", "Dieser Schritt ist bereits im Katalog.", parent=dialog)
                return
            listbox.insert("end", value)
            new_value_var.set("")
            new_value_entry.focus_set()

        def remove_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Auswahl fehlt", "Bitte zuerst einen Eintrag auswählen.", parent=dialog)
                return
            listbox.delete(selection[0])

        def move_selected(direction: int) -> None:
            selection = listbox.curselection()
            if not selection:
                return
            idx = selection[0]
            target = idx + direction
            if target < 0 or target >= listbox.size():
                return
            value = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(target, value)
            listbox.selection_clear(0, "end")
            listbox.selection_set(target)
            listbox.activate(target)

        def save_catalog() -> None:
            items = [listbox.get(i).strip() for i in range(listbox.size()) if listbox.get(i).strip()]
            if not items:
                messagebox.showwarning("Leerer Katalog", "Bitte mindestens einen Eintrag hinterlegen.", parent=dialog)
                return
            try:
                self.repo.save_naechster_schritt_catalog(items)
                self.lookups = self.repo.load_lookups()
                self._refresh_lookup_comboboxes()
            except (OSError, PermissionError) as exc:
                messagebox.showerror(
                    "Speichern fehlgeschlagen",
                    f"Katalog konnte nicht gespeichert werden:\n{exc}\n\nIst die Excel-Datei vielleicht gerade geöffnet?",
                    parent=dialog,
                )
                return
            except ValueError as exc:
                messagebox.showwarning("Ungültiger Katalog", str(exc), parent=dialog)
                return

            messagebox.showinfo("Gespeichert", "Katalog wurde aktualisiert.", parent=dialog)
            dialog.destroy()

        ttk.Button(button_row, text="Hinzufügen", command=add_item, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Ausgewählten löschen", command=remove_selected, width=20).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Nach oben", command=lambda: move_selected(-1), width=12).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Nach unten", command=lambda: move_selected(1), width=12).pack(side="left")

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        ttk.Button(footer, text="Speichern", command=save_catalog, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(footer, text="Abbrechen", command=dialog.destroy, width=14).pack(side="left")

        new_value_entry.bind("<Return>", lambda _event: add_item())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        new_value_entry.focus_set()

    def _refresh_lookup_comboboxes(self) -> None:
        form_lookup_values = {
            "quelle": self.lookups.quelle,
            "art_bewerbung": self.lookups.art_bewerbung,
            "unterlagen": self.lookups.unterlagen,
            "status": self.lookups.status,
            "prioritaet": self.lookups.prioritaet,
            "naechster_schritt": [""] + self.lookups.naechster_schritt,
        }
        for key, values in form_lookup_values.items():
            widget = self.form_widgets.get(key)
            if isinstance(widget, ttk.Combobox):
                widget.configure(values=values)
                current_value = widget.get().strip()
                if current_value and current_value not in values:
                    widget.set("")

        update_lookup_values = {
            "status": self.lookups.status,
            "art_kontaktaufnahme": self.lookups.art_kontaktaufnahme,
            "naechster_schritt": [""] + self.lookups.naechster_schritt,
            "endergebnis": [""] + self.lookups.endergebnis,
        }
        for key, values in update_lookup_values.items():
            for widget_dict in (self.update_widgets, self.active_update_widgets):
                widget = widget_dict.get(key)
                if isinstance(widget, ttk.Combobox):
                    widget.configure(values=values)
                    current_value = widget.get().strip()
                    if current_value and current_value not in values:
                        widget.set("")

    def _reset_form(self) -> None:
        for key, var in self.form_vars.items():
            if key in {"datum_bewerbung", "status_datum"}:
                var.set(date.today().strftime("%d.%m.%Y"))
            else:
                var.set("")

    def refresh_data(self) -> None:
        self._set_refresh_state(True, "Daten werden geladen...")
        try:
            self.records = self.repo.load_records()
            self.follow_up_records, self.archive_records = self.repo.categorize_records(self.records)
            self.active_records = [
                rec for rec in self.records
                if not bool(rec.endergebnis) and rec.status not in ARCHIVE_STATUS
            ]
            self._render_active_records()
            self._render_follow_up_records()
            self._render_archive_records()
            self._clear_follow_up_form()
            self._clear_active_form()
            self.status_var.set(f"{len(self.records)} Einträge geladen")
        finally:
            self._set_refresh_state(False, self.status_var.get())

    def _render_follow_up_records(self) -> None:
        for item in self.follow_tree.get_children():
            self.follow_tree.delete(item)

        self.follow_tree["displaycolumns"] = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt")
        self.follow_tree["columns"] = ("id", "unternehmen", "position", "status", "erinnerung", "prioritaet", "schritt", "_row")

        filter_text = self.follow_filter_var.get().strip().lower()
        records = self.follow_up_records
        if filter_text:
            records = [
                rec
                for rec in records
                if filter_text in " ".join(
                    [
                        rec.bewerbungs_id,
                        rec.unternehmen,
                        rec.position,
                        rec.status,
                        rec.prioritaet,
                        rec.naechster_schritt,
                        rec.erinnerungsdatum.strftime("%d.%m.%Y") if rec.erinnerungsdatum else "",
                    ]
                ).lower()
            ]

        records = sorted(records, key=lambda rec: self._follow_sort_key(rec, self.follow_sort_column), reverse=self.follow_sort_desc)

        for col, base in self.follow_headings.items():
            marker = ""
            if col == self.follow_sort_column:
                marker = " ↓" if self.follow_sort_desc else " ↑"
            self.follow_tree.heading(col, text=f"{base}{marker}", command=lambda c=col: self._toggle_follow_sort(c))

        for rec in records:
            row_tag = self._get_follow_row_tag(rec)
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
                tags=(row_tag,) if row_tag else (),
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

        filter_text = self.archive_filter_var.get().strip().lower()
        records = self.archive_records
        if filter_text:
            records = [
                rec
                for rec in records
                if filter_text in " ".join(
                    [
                        rec.bewerbungs_id,
                        rec.unternehmen,
                        rec.position,
                        rec.status,
                        rec.endergebnis,
                        rec.status_datum.strftime("%d.%m.%Y") if rec.status_datum else "",
                    ]
                ).lower()
            ]

        records = sorted(records, key=lambda rec: self._archive_sort_key(rec, self.archive_sort_column), reverse=self.archive_sort_desc)

        for col, base in self.archive_headings.items():
            marker = ""
            if col == self.archive_sort_column:
                marker = " ↓" if self.archive_sort_desc else " ↑"
            self.archive_tree.heading(col, text=f"{base}{marker}", command=lambda c=col: self._toggle_archive_sort(c))

        for rec in records:
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
