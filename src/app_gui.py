from __future__ import annotations

import queue
import tkinter as tk
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from report_service import ExportRequest, ReportMetadata, export_report, validate_export_request


APP_TITLE = "PM Document Converter"


class PMDocumentConverterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("1120x720")
        self.minsize(1040, 680)

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.export_thread: threading.Thread | None = None

        self.zip_path_var = ctk.StringVar()
        self.output_dir_var = ctk.StringVar()
        self.project_name_var = ctk.StringVar()
        self.project_no_var = ctk.StringVar()
        self.db_name_var = ctk.StringVar()
        self.customer_full_var = ctk.StringVar()
        self.customer_abbrev_var = ctk.StringVar()
        self.quarter_var = ctk.StringVar()
        self.report_type_var = ctk.StringVar(value="Preventive Maintenance")
        self.language_var = ctk.StringVar(value="English")
        self.report_date_var = ctk.StringVar()
        self.person_type_var = ctk.StringVar(value="Customer(s)")
        self.person_name_var = ctk.StringVar()
        self.person_phone_var = ctk.StringVar()
        self.person_email_var = ctk.StringVar()
        self.create_pdf_var = ctk.BooleanVar(value=False)
        self.people: dict[str, list[dict[str, str]]] = {
            "customers": [],
            "sales": [],
            "engineers": [],
        }
        self.people_listboxes: dict[str, tk.Listbox] = {}

        self._build_layout()
        self.after(150, self._poll_events)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=230, corner_radius=0, fg_color="#10243E")
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        title = ctk.CTkLabel(
            sidebar,
            text="PM Document\nConverter",
            text_color="white",
            font=ctk.CTkFont(size=22, weight="bold"),
            justify="left",
        )
        title.pack(anchor="w", padx=22, pady=(28, 8))

        subtitle = ctk.CTkLabel(
            sidebar,
            text="Fast Assessment to PM report",
            text_color="#B9C7D8",
            font=ctk.CTkFont(size=12),
        )
        subtitle.pack(anchor="w", padx=22, pady=(0, 24))

        for item in ("1. Project", "2. Input", "3. Export"):
            ctk.CTkLabel(
                sidebar,
                text=item,
                text_color="#E8EEF7",
                anchor="w",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(fill="x", padx=22, pady=9)

        self.status_badge = ctk.CTkLabel(
            sidebar,
            text="Ready",
            fg_color="#1D6F42",
            text_color="white",
            corner_radius=6,
            height=30,
        )
        self.status_badge.pack(side="bottom", fill="x", padx=22, pady=24)

        main = ctk.CTkFrame(self, fg_color="#F4F7FB", corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 10))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Generate Preventive Maintenance Report",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="#172033",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Select a Fast Assessment ZIP, confirm metadata, then export DOCX/PDF.",
            font=ctk.CTkFont(size=13),
            text_color="#526070",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        body = ctk.CTkScrollableFrame(main, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=28, pady=(0, 18))
        body.grid_columnconfigure((0, 1), weight=1, uniform="columns")

        self._build_project_card(body)
        self._build_input_card(body)
        self._build_contacts_card(body)
        self._build_export_card(body)

    def _build_project_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, "Project Information")
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)

        self._entry(card, "Project name", self.project_name_var, 0)
        self._entry(card, "Project no.", self.project_no_var, 1)
        self._entry(card, "Database name", self.db_name_var, 2)
        self._entry(card, "Customer full name", self.customer_full_var, 3)
        self._entry(card, "Customer abbreviation", self.customer_abbrev_var, 4)
        self._entry(card, "Quarter", self.quarter_var, 5)
        self._entry(card, "Report date", self.report_date_var, 6)

        ctk.CTkLabel(card, text="Report type", anchor="w", text_color="#344054").grid(
            row=15, column=0, sticky="ew", padx=18, pady=(8, 2)
        )
        ctk.CTkOptionMenu(
            card,
            values=["Preventive Maintenance", "Installation"],
            variable=self.report_type_var,
        ).grid(row=16, column=0, sticky="ew", padx=18, pady=(0, 12))

        ctk.CTkLabel(card, text="Report language", anchor="w", text_color="#344054").grid(
            row=17, column=0, sticky="ew", padx=18, pady=(8, 2)
        )
        ctk.CTkOptionMenu(
            card,
            values=["English", "Thai"],
            variable=self.language_var,
        ).grid(row=18, column=0, sticky="ew", padx=18, pady=(0, 12))

    def _build_input_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, "Input and Output")
        card.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)

        self._path_row(card, "Fast Assessment ZIP", self.zip_path_var, self._browse_zip, 0)
        self._path_row(card, "Output folder", self.output_dir_var, self._browse_output_dir, 2)

        ctk.CTkLabel(
            card,
            text="The export creates DOCX, parsed JSON, charts, and PDF when Word or LibreOffice is available.",
            text_color="#667085",
            wraplength=390,
            justify="left",
        ).grid(row=5, column=0, sticky="ew", padx=18, pady=(12, 18))

    def _build_contacts_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, "Personal Information")
        card.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=0, pady=10)
        card.grid_columnconfigure(0, weight=1)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=18, pady=(4, 18))
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(2, weight=2)

        left = ctk.CTkFrame(content, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_columnconfigure(1, weight=1)

        for row, (label, key, height) in enumerate(
            (("Customer(s)", "customers", 5), ("Sale", "sales", 2), ("Engineer(s)", "engineers", 5))
        ):
            ctk.CTkLabel(left, text=label, anchor="w", text_color="#344054").grid(
                row=row * 2, column=0, sticky="nw", padx=(0, 8), pady=(2, 4)
            )
            listbox = tk.Listbox(
                left,
                height=height,
                exportselection=False,
                activestyle="none",
                font=("Tahoma", 9),
            )
            listbox.grid(row=row * 2, column=1, sticky="ew", pady=(0, 8))
            listbox.bind("<<ListboxSelect>>", lambda _event, item_key=key: self._load_selected_person(item_key))
            self.people_listboxes[key] = listbox

            ctk.CTkButton(
                left,
                text="Select",
                width=58,
                command=lambda item_label=label: self._select_person_type(item_label),
            ).grid(row=row * 2, column=2, sticky="n", padx=(8, 0), pady=(0, 8))

        separator = ctk.CTkFrame(content, width=1, fg_color="#98A2B3")
        separator.grid(row=0, column=1, sticky="ns", padx=18)

        editor = ctk.CTkFrame(content, fg_color="transparent")
        editor.grid(row=0, column=2, sticky="nsew")
        editor.grid_columnconfigure(1, weight=1)

        fields = (
            ("Select on", self.person_type_var, True),
            ("Name", self.person_name_var, False),
            ("Phone", self.person_phone_var, False),
            ("Email Address", self.person_email_var, False),
        )
        for row, (label, variable, readonly) in enumerate(fields):
            ctk.CTkLabel(editor, text=label, anchor="w", text_color="#344054").grid(
                row=row, column=0, sticky="ew", padx=(0, 10), pady=(0, 10)
            )
            entry = ctk.CTkEntry(editor, textvariable=variable, height=30)
            entry.grid(row=row, column=1, sticky="ew", pady=(0, 10))
            if readonly:
                entry.configure(state="disabled")

        buttons = ctk.CTkFrame(editor, fg_color="transparent")
        buttons.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        buttons.grid_columnconfigure((0, 1, 2), weight=1, uniform="person_buttons")
        for index, (text, command) in enumerate(
            (
                ("Add", self._add_person),
                ("Edit", self._edit_person),
                ("Delete", self._delete_person),
                ("Clear", self._clear_person_inputs),
                ("Load", self._load_active_selected_person),
            )
        ):
            ctk.CTkButton(buttons, text=text, height=32, command=command).grid(
                row=index // 3,
                column=index % 3,
                sticky="ew",
                padx=(0, 8),
                pady=(0, 8),
            )

    def _build_export_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, "Export Status")
        card.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=0, pady=10)
        card.grid_rowconfigure(1, weight=1)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=18, pady=(4, 10))
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkCheckBox(
            actions,
            text="Create PDF",
            variable=self.create_pdf_var,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.export_button = ctk.CTkButton(
            actions,
            text="Export Report",
            height=38,
            command=self._start_export,
        )
        self.export_button.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(actions)
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.progress_bar.set(0)

        self.log_box = ctk.CTkTextbox(card, height=260, wrap="word")
        self.log_box.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self._append_log("Ready.")

    def _card(self, parent: ctk.CTkFrame, title: str) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color="white", corner_radius=8, border_width=1, border_color="#E3E8EF")
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#172033",
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        return card

    def _entry(self, parent: ctk.CTkFrame, label: str, variable: ctk.StringVar, index: int) -> None:
        row = index * 2 + 1
        ctk.CTkLabel(parent, text=label, anchor="w", text_color="#344054").grid(
            row=row, column=0, sticky="ew", padx=18, pady=(5, 2)
        )
        ctk.CTkEntry(parent, textvariable=variable, height=34).grid(
            row=row + 1, column=0, sticky="ew", padx=18, pady=(0, 5)
        )

    def _path_row(
        self,
        parent: ctk.CTkFrame,
        label: str,
        variable: ctk.StringVar,
        command,
        row: int,
    ) -> None:
        ctk.CTkLabel(parent, text=label, anchor="w", text_color="#344054").grid(
            row=row + 1, column=0, sticky="ew", padx=18, pady=(6, 2)
        )
        group = ctk.CTkFrame(parent, fg_color="transparent")
        group.grid(row=row + 2, column=0, sticky="ew", padx=18, pady=(0, 8))
        group.grid_columnconfigure(0, weight=1)

        ctk.CTkEntry(group, textvariable=variable, height=34).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(group, text="Browse", width=92, command=command).grid(row=0, column=1)

    def _browse_zip(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Fast Assessment ZIP",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
        )
        if selected:
            self.zip_path_var.set(selected)

    def _browse_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="Select output folder")
        if selected:
            self.output_dir_var.set(selected)

    def _person_key(self) -> str:
        mapping = {
            "Customer(s)": "customers",
            "Sale": "sales",
            "Engineer(s)": "engineers",
        }
        return mapping.get(self.person_type_var.get(), "customers")

    def _person_label(self, key: str) -> str:
        return {
            "customers": "Customer(s)",
            "sales": "Sale",
            "engineers": "Engineer(s)",
        }.get(key, "Customer(s)")

    def _select_person_type(self, label: str) -> None:
        self.person_type_var.set(label)
        for listbox in self.people_listboxes.values():
            listbox.selection_clear(0, "end")

    def _person_from_inputs(self) -> dict[str, str]:
        return {
            "name": self.person_name_var.get().strip(),
            "phone": self.person_phone_var.get().strip(),
            "email": self.person_email_var.get().strip(),
        }

    def _format_person(self, person: dict[str, str]) -> str:
        return person.get("name") or person.get("email") or person.get("phone") or "(blank)"

    def _refresh_people_list(self, key: str) -> None:
        listbox = self.people_listboxes.get(key)
        if not listbox:
            return
        listbox.delete(0, "end")
        for person in self.people[key]:
            listbox.insert("end", self._format_person(person))

    def _selected_person_index(self, key: str) -> int | None:
        listbox = self.people_listboxes.get(key)
        if not listbox:
            return None
        selection = listbox.curselection()
        return int(selection[0]) if selection else None

    def _load_selected_person(self, key: str) -> None:
        self.person_type_var.set(self._person_label(key))
        index = self._selected_person_index(key)
        if index is None or index >= len(self.people[key]):
            return
        person = self.people[key][index]
        self.person_name_var.set(person.get("name", ""))
        self.person_phone_var.set(person.get("phone", ""))
        self.person_email_var.set(person.get("email", ""))

    def _load_active_selected_person(self) -> None:
        self._load_selected_person(self._person_key())

    def _add_person(self) -> None:
        person = self._person_from_inputs()
        if not any(person.values()):
            return
        key = self._person_key()
        self.people[key].append(person)
        self._refresh_people_list(key)
        self._clear_person_inputs()

    def _edit_person(self) -> None:
        key = self._person_key()
        index = self._selected_person_index(key)
        person = self._person_from_inputs()
        if index is None or index >= len(self.people[key]):
            if any(person.values()):
                self._add_person()
            return
        if not any(person.values()):
            return
        self.people[key][index] = person
        self._refresh_people_list(key)
        self.people_listboxes[key].selection_set(index)

    def _delete_person(self) -> None:
        key = self._person_key()
        index = self._selected_person_index(key)
        if index is None or index >= len(self.people[key]):
            return
        del self.people[key][index]
        self._refresh_people_list(key)
        self._clear_person_inputs()

    def _clear_person_inputs(self) -> None:
        self.person_name_var.set("")
        self.person_phone_var.set("")
        self.person_email_var.set("")

    def _metadata(self) -> ReportMetadata:
        customers = list(self.people["customers"])
        sales = list(self.people["sales"])
        engineers = list(self.people["engineers"])
        sale_names = [person.get("name", "") for person in sales if person.get("name", "")]
        sale = "; ".join(sale_names)
        first_engineer = engineers[0]["name"] if engineers and engineers[0].get("name") else ""

        return ReportMetadata(
            project_name=self.project_name_var.get().strip(),
            project_no=self.project_no_var.get().strip(),
            db_name=self.db_name_var.get().strip(),
            customer_full=self.customer_full_var.get().strip(),
            customer_abbrev=self.customer_abbrev_var.get().strip(),
            quarter=self.quarter_var.get().strip(),
            sale=sale,
            report_type=self.report_type_var.get().strip(),
            language=self.language_var.get().strip(),
            report_date=self.report_date_var.get().strip(),
            customers=customers,
            sales=sales,
            engineers=engineers,
            change_records=[
                {
                    "date": self.report_date_var.get().strip() or datetime.now().strftime("%d-%b-%Y"),
                    "author": first_engineer,
                    "version": "1.0",
                    "ref": "Initial Document",
                }
            ],
            reviewers=[],
        )

    def _request(self) -> ExportRequest:
        output_dir = self.output_dir_var.get().strip()
        return ExportRequest(
            zip_path=Path(self.zip_path_var.get().strip()),
            output_dir=Path(output_dir) if output_dir else None,
            metadata=self._metadata(),
            create_pdf=self.create_pdf_var.get(),
        )

    def _start_export(self) -> None:
        if self.export_thread and self.export_thread.is_alive():
            return

        request = self._request()
        errors = validate_export_request(request)
        if errors:
            messagebox.showerror("Export validation failed", "\n".join(errors))
            self._append_log("Validation failed.")
            return

        self.log_box.delete("1.0", "end")
        self._append_log("Starting export.")
        self.progress_bar.set(0.15)
        self.status_badge.configure(text="Running", fg_color="#B7791F")
        self.export_button.configure(state="disabled", text="Exporting...")

        self.export_thread = threading.Thread(target=self._export_worker, args=(request,), daemon=True)
        self.export_thread.start()

    def _export_worker(self, request: ExportRequest) -> None:
        try:
            result = export_report(request, progress=lambda msg: self.events.put(("progress", msg)))
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("error", exc))

    def _poll_events(self) -> None:
        while True:
            try:
                event_type, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event_type == "progress":
                self._append_log(str(payload))
                self.progress_bar.set(min(self.progress_bar.get() + 0.25, 0.85))
            elif event_type == "success":
                self.progress_bar.set(1)
                self.status_badge.configure(text="Completed", fg_color="#1D6F42")
                self.export_button.configure(state="normal", text="Export Report")
                self._append_log(f"DOCX: {payload.docx_path}")
                self._append_log(f"JSON: {payload.json_path}")
                if payload.pdf_path:
                    self._append_log(f"PDF: {payload.pdf_path}")
                else:
                    self._append_log("PDF: not created")
                messagebox.showinfo("Export completed", f"DOCX:\n{payload.docx_path}\n\nJSON:\n{payload.json_path}")
            elif event_type == "error":
                self.progress_bar.set(0)
                self.status_badge.configure(text="Failed", fg_color="#A52929")
                self.export_button.configure(state="normal", text="Export Report")
                self._append_log(f"Error: {payload}")
                messagebox.showerror("Export failed", str(payload))

        self.after(150, self._poll_events)

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n")
        self.log_box.see("end")


def main() -> None:
    app = PMDocumentConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
