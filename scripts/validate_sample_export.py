from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "output"
BASE_NAME = "MFEC_PM_MotifPRDDBN_Q2_2025_20260109"
JSON_PATH = OUTPUT_DIR / f"{BASE_NAME}_parsed_data.json"
DOCX_PATH = OUTPUT_DIR / f"{BASE_NAME}_PM_Report.docx"


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def check_equal(label: str, actual, expected) -> None:
    if actual != expected:
        fail(f"{label}: expected {expected!r}, got {actual!r}")
    print(f"OK: {label} = {actual!r}")


def main() -> int:
    if not JSON_PATH.exists():
        fail(f"Parsed JSON not found: {JSON_PATH}")
    if not DOCX_PATH.exists():
        fail(f"DOCX not found: {DOCX_PATH}")

    with JSON_PATH.open("r", encoding="utf-8") as file:
        data = json.load(file)

    os_info = data.get("os_info", {})

    check_equal("DOCX exists", DOCX_PATH.exists(), True)
    check_equal("JSON exists", JSON_PATH.exists(), True)
    check_equal("os_info.hostname", os_info.get("hostname"), "MotifPRDDBN")
    check_equal("os_info.ip_address", os_info.get("ip_address"), "192.1.255.190")
    check_equal("os_info.oracle_home", os_info.get("oracle_home"), "/u01/app/oracle/product/19.0.0/db_1")
    check_equal("datafiles count", len(data.get("datafiles", [])), 35)
    check_equal("registry_list count", len(data.get("registry_list", [])), 13)
    check_equal("invalid_objects_list count", len(data.get("invalid_objects_list", [])), 367)
    check_equal("growth_rows_raw count", len(data.get("growth_rows_raw", [])), 5)

    registry_ids = [row.get("comp_id") for row in data.get("registry_list", [])]
    required_registry_ids = {"CATALOG", "CATPROC", "JAVAVM", "XML", "XDB", "SDO", "XOQ"}
    missing_registry_ids = sorted(required_registry_ids.difference(registry_ids))
    if missing_registry_ids:
        fail(f"registry_list missing component ids: {', '.join(missing_registry_ids)}")

    print("OK: required registry component ids are present")
    print("Sample export validation passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

