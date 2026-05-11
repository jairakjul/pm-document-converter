# CircleCI

## What CI Does

The CircleCI pipeline has two jobs:

1. `validate-report`
   - Installs Python dependencies on Linux.
   - Compiles Python files.
   - Runs `scripts/export_and_validate_sample.py`.
   - Stores generated report output as artifacts.

2. `build-windows-exe`
   - Runs after `validate-report`.
   - Uses a Windows runner.
   - Runs the same sample export validation on Windows.
   - Builds the desktop app with `scripts/build_desktop.ps1`.
   - Verifies `PMDocumentConverter.exe` is a Windows GUI executable.
   - Stores `dist/PMDocumentConverter` as the packaged app artifact.

## Artifacts

CircleCI stores:

- `report-output`: generated DOCX, JSON, and chart files from Linux validation.
- `windows-report-output`: generated DOCX, JSON, and chart files from Windows validation.
- `packaged-app`: the built Windows desktop application folder.

## Notes

The Windows job requires CircleCI Windows executor access. If the account does not have Windows runners enabled, keep `validate-report` active and run packaging locally with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -Clean
```
