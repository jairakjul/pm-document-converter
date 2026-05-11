from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from fa_parser_zip_report import process_zip_to_reports


ProgressCallback = Callable[[str], None]


@dataclass(frozen=True)
class ReportMetadata:
    project_name: str = ""
    project_no: str = ""
    db_name: str = ""
    customer_full: str = ""
    customer_abbrev: str = ""
    quarter: str = ""
    sale: str = ""
    report_type: str = "Preventive Maintenance"
    customers: list[dict] = field(default_factory=list)
    engineers: list[dict] = field(default_factory=list)
    change_records: list[dict] = field(default_factory=list)
    reviewers: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "project_name": self.project_name,
            "project_no": self.project_no,
            "db_name": self.db_name,
            "customer_full": self.customer_full,
            "customer_abbrev": self.customer_abbrev,
            "quarter": self.quarter,
            "sale": self.sale,
            "report_type": self.report_type,
            "customers": self.customers,
            "engineers": self.engineers,
            "change_records": self.change_records,
            "reviewers": self.reviewers,
        }


@dataclass(frozen=True)
class ExportRequest:
    zip_path: Path
    output_dir: Path
    metadata: ReportMetadata = field(default_factory=ReportMetadata)


@dataclass(frozen=True)
class ExportResult:
    docx_path: Path
    pdf_path: Optional[Path]
    json_path: Path


def validate_export_request(request: ExportRequest) -> list[str]:
    errors = []

    if not request.zip_path:
        errors.append("Select a Fast Assessment ZIP file.")
    elif not request.zip_path.exists():
        errors.append(f"ZIP file does not exist: {request.zip_path}")
    elif request.zip_path.suffix.lower() != ".zip":
        errors.append("Input file must be a .zip file.")

    if not request.output_dir:
        errors.append("Select an output folder.")
    elif request.output_dir.exists() and not request.output_dir.is_dir():
        errors.append(f"Output path is not a folder: {request.output_dir}")

    return errors


def export_report(request: ExportRequest, progress: ProgressCallback | None = None) -> ExportResult:
    errors = validate_export_request(request)
    if errors:
        raise ValueError("\n".join(errors))

    def emit(message: str) -> None:
        if progress:
            progress(message)

    emit("Validating input")
    request.output_dir.mkdir(parents=True, exist_ok=True)

    emit("Extracting ZIP and parsing Fast Assessment data")
    docx_path, pdf_path, json_path = process_zip_to_reports(
        request.zip_path,
        request.output_dir,
        metadata=request.metadata.to_dict(),
    )

    emit("Report export completed")
    return ExportResult(docx_path=docx_path, pdf_path=pdf_path, json_path=json_path)

