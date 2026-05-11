from __future__ import annotations

import queue
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
        self.output_dir_var = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.project_name_var = ctk.StringVar(value="MA Oracle 5 Years")
        self.project_no_var = ctk.StringVar()
        self.db_name_var = ctk.StringVar()
        self.customer_full_var = ctk.StringVar()
        self.customer_abbrev_var = ctk.StringVar()
        self.quarter_var = ctk.StringVar(value="Q2 2025")
        self.sale_var = ctk.StringVar()
        self.report_type_var = ctk.StringVar(value="Preventive Maintenance")
        self.customer_name_var = ctk.StringVar()
        self.customer_phone_var = ctk.StringVar()
        self.customer_email_var = ctk.StringVar()
        self.engineer_name_var = ctk.StringVar()
        self.engineer_phone_var = ctk.StringVar()
        self.engineer_email_var = ctk.StringVar()

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
        self._entry(card, "Sales name", self.sale_var, 6)

        ctk.CTkLabel(card, text="Report type", anchor="w", text_color="#344054").grid(
            row=15, column=0, sticky="ew", padx=18, pady=(8, 2)
        )
        ctk.CTkOptionMenu(
            card,
            values=["Preventive Maintenance", "Installation"],
            variable=self.report_type_var,
        ).grid(row=16, column=0, sticky="ew", padx=18, pady=(0, 12))

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
        card = self._card(parent, "Contacts")
        card.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=10)

        self._entry(card, "Customer contact", self.customer_name_var, 0)
        self._entry(card, "Customer phone", self.customer_phone_var, 1)
        self._entry(card, "Customer email", self.customer_email_var, 2)
        self._entry(card, "MFEC engineer", self.engineer_name_var, 3)
        self._entry(card, "Engineer phone", self.engineer_phone_var, 4)
        self._entry(card, "Engineer email", self.engineer_email_var, 5)

    def _build_export_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, "Export Status")
        card.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=10)
        card.grid_rowconfigure(1, weight=1)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=18, pady=(4, 10))
        actions.grid_columnconfigure(0, weight=1)

        self.export_button = ctk.CTkButton(
            actions,
            text="Export Report",
            height=38,
            command=self._start_export,
        )
        self.export_button.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(actions)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 16))
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

    def _metadata(self) -> ReportMetadata:
        customers = []
        if any(v.get().strip() for v in (self.customer_name_var, self.customer_phone_var, self.customer_email_var)):
            customers.append(
                {
                    "name": self.customer_name_var.get().strip(),
                    "phone": self.customer_phone_var.get().strip(),
                    "email": self.customer_email_var.get().strip(),
                }
            )

        engineers = []
        if any(v.get().strip() for v in (self.engineer_name_var, self.engineer_phone_var, self.engineer_email_var)):
            engineers.append(
                {
                    "name": self.engineer_name_var.get().strip(),
                    "phone": self.engineer_phone_var.get().strip(),
                    "email": self.engineer_email_var.get().strip(),
                }
            )

        return ReportMetadata(
            project_name=self.project_name_var.get().strip(),
            project_no=self.project_no_var.get().strip(),
            db_name=self.db_name_var.get().strip(),
            customer_full=self.customer_full_var.get().strip(),
            customer_abbrev=self.customer_abbrev_var.get().strip(),
            quarter=self.quarter_var.get().strip(),
            sale=self.sale_var.get().strip(),
            report_type=self.report_type_var.get().strip(),
            customers=customers,
            engineers=engineers,
            change_records=[
                {
                    "date": datetime.now().strftime("%d-%b-%Y"),
                    "author": self.engineer_name_var.get().strip(),
                    "version": "1.0",
                    "ref": "Initial Document",
                }
            ],
            reviewers=[],
        )

    def _request(self) -> ExportRequest:
        return ExportRequest(
            zip_path=Path(self.zip_path_var.get().strip()),
            output_dir=Path(self.output_dir_var.get().strip()),
            metadata=self._metadata(),
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
