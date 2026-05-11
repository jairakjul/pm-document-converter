# PM Document Converter Fresh

Clean restart workspace for the PM Document Converter.

## Goal
Generate Preventive Maintenance reports from a Fast Assessment ZIP file.

## Current Baseline
This folder starts from the current working baseline:

- `src/app_gui.py` is the modern CustomTkinter desktop UI.
- `src/report_service.py` validates UI input and calls the export backend.
- `src/fa_parser_zip_report.py` handles ZIP extraction, parsing orchestration, the legacy Tkinter GUI, and PDF conversion.
- `src/report_writer_light.py` handles DOCX creation and report formatting.
- `samples/` stores Fast Assessment ZIP files for validation.
- `output/` stores generated DOCX, PDF, JSON, and chart files.

## Run
The fresh workspace uses a local virtual environment. If it already exists, run:

```powershell
cd C:\mfec\pm_document_converter_fresh
.\.venv\Scripts\Activate.ps1
python .\src\app_gui.py
```

If the environment needs to be recreated, use:

```powershell
cd C:\mfec\pm_document_converter_fresh
uv python install 3.12
uv venv .venv --python 3.12
uv pip install --python .\.venv\Scripts\python.exe -r requirements.txt
```

The old Tkinter UI is still available for fallback:

```powershell
python .\src\fa_parser_zip_report.py
```

Select:

```text
samples\MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip
```

Use `output\` as the output folder.

## Development Rule
Keep the fresh folder clean:

- Do not keep `build/`, `dist/`, `__pycache__/`, or temporary inspection folders here.
- Keep generated reports under `output/`.
- Keep sample input ZIPs under `samples/`.
- Keep implementation code under `src/`.
- Keep notes and design docs under `docs/`.

## Package

Build the Windows desktop executable with PyInstaller:

```powershell
cd C:\mfec\pm_document_converter_fresh
powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -Clean
```

The executable will be created at:

```text
dist\PMDocumentConverter\PMDocumentConverter.exe
```

See `docs\packaging.md` for the packaging workflow and smoke test checklist.

For macOS, build on a Mac:

```bash
cd /path/to/pm_document_converter_fresh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
bash scripts/build_desktop_macos.sh
```

The macOS app bundle will be created at:

```text
dist/PMDocumentConverter.app
```

## CI

CircleCI configuration is available at `.circleci\config.yml`.

It validates report generation, builds the Windows desktop app, and stores generated reports plus the packaged app as CircleCI artifacts.

See `docs\ci.md` for details.

GitHub Actions workflow is also available at `.github\workflows\build.yml`.

It can build the macOS `.app` on Apple-hosted macOS runners and upload it as an artifact.

See `docs\github_actions.md` for details.
