from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
SAMPLE_ZIP = PROJECT_ROOT / "samples" / "MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip"
OUTPUT_DIR = PROJECT_ROOT / "output"
VALIDATION_RUNNER = PROJECT_ROOT / "scripts" / "run_all_validations.py"
BASE_NAME = "MFEC_PM_MotifPRDDBN_Q2_2025_20260109"


def remove_previous_sample_outputs() -> None:
    patterns = [
        f"{BASE_NAME}_PM_Report.docx",
        f"{BASE_NAME}_PM_Report.pdf",
        f"{BASE_NAME}_parsed_data.json",
        "database_growth_rate.png",
        "database_pga_advisory.png",
        "database_sga_advisory.png",
        "os_runq_perf.png",
        "os_cpu_util_all.png",
        "os_cpu_util_system.png",
        "os_cpu_util_user.png",
        "os_mem_free.png",
        "os_swap_used.png",
        "os_page_in.png",
        "os_page_out.png",
        "os_cpu_wait_io.png",
        "os_disk_service_time.png",
        "os_cpu_perf.png",
        "os_disk_perf.png",
        "os_mem_perf.png",
    ]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for name in patterns:
        path = OUTPUT_DIR / name
        if path.exists():
            path.unlink()


def export_sample() -> None:
    if not SAMPLE_ZIP.exists():
        raise FileNotFoundError(f"Sample ZIP not found: {SAMPLE_ZIP}")

    sys.path.insert(0, str(SRC_DIR))
    from report_service import ExportRequest, ReportMetadata, export_report

    request = ExportRequest(
        zip_path=SAMPLE_ZIP,
        output_dir=OUTPUT_DIR,
        metadata=ReportMetadata(
            project_name="Sample Validation Export",
            db_name="icompdb",
            quarter="Q2 2025",
            report_type="Preventive Maintenance",
        ),
    )

    result = export_report(request, progress=lambda message: print(f"EXPORT: {message}"))
    print(f"EXPORT: DOCX {result.docx_path}")
    print(f"EXPORT: JSON {result.json_path}")
    print(f"EXPORT: PDF {result.pdf_path if result.pdf_path else 'not created'}")


def run_validations() -> int:
    result = subprocess.run([sys.executable, str(VALIDATION_RUNNER)], cwd=str(PROJECT_ROOT))
    return result.returncode


def main() -> int:
    print("Cleaning previous sample outputs")
    remove_previous_sample_outputs()

    print("Exporting sample report")
    export_sample()

    print("Running validations")
    return run_validations()


if __name__ == "__main__":
    sys.exit(main())
