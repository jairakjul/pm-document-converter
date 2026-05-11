from __future__ import annotations

import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = [
    PROJECT_ROOT / "scripts" / "validate_sample_export.py",
    PROJECT_ROOT / "scripts" / "validate_docx_structure.py",
    PROJECT_ROOT / "scripts" / "validate_report_quality.py",
]


def main() -> int:
    for script in SCRIPTS:
        print(f"\n=== Running {script.name} ===")
        result = subprocess.run([sys.executable, str(script)], cwd=str(PROJECT_ROOT))
        if result.returncode != 0:
            print(f"\nValidation failed: {script.name}")
            return result.returncode

    print("\nAll validations passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

