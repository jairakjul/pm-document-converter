# Packaging

## Goal

Build a Windows desktop executable for the CustomTkinter GUI without opening a terminal window.

## Build Command

### Windows

Run from PowerShell:

```powershell
cd C:\mfec\pm_document_converter_fresh
powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -Clean
```

The executable is created at:

```text
dist\PMDocumentConverter\PMDocumentConverter.exe
```

### macOS

A real `.app` bundle must be built on macOS. PyInstaller cannot create a working macOS app bundle from Windows.

On a Mac:

```bash
cd /path/to/pm_document_converter_fresh
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
bash scripts/build_desktop_macos.sh
```

The app bundle is created at:

```text
dist/PMDocumentConverter.app
```

## What Is Included

- `src/app_gui.py` as the GUI entrypoint.
- `src/report_service.py`, `src/fa_parser_zip_report.py`, and `src/report_writer_light.py` through normal Python imports.
- CustomTkinter package data.
- Matplotlib runtime data for chart generation using the `Agg` backend.
- Runtime dependencies installed in `.venv` from `requirements.txt`.

## What Is Not Included

- Sample ZIP files.
- Generated reports from `output/`.
- Build artifacts from previous runs when `-Clean` is used.
- Unused Pythonwin UI files.
- PIL AVIF image support, which is not needed for this app.
- Package metadata folders that are not needed at runtime.

## Manual Smoke Test

After building:

1. Run `dist\PMDocumentConverter\PMDocumentConverter.exe`.
2. Select `samples\MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip`.
3. Select an empty output folder outside `dist`.
4. Click `Export Report`.
5. Confirm a DOCX and parsed JSON are created.

## Notes

PDF conversion still depends on Microsoft Word or LibreOffice being available on the target machine. The executable can create DOCX reports without those tools.

On macOS, DOCX generation works without Microsoft Word. PDF conversion requires LibreOffice in `PATH` unless a macOS-specific Word automation path is added later.
