# Development Notes

## Fresh Start Decision
The original `C:\mfec` root contains stable code, backups, generated reports, old packaged builds, temporary extracted data, and experiment files. A fresh folder is easier to reason about and safer for v2 work.

## Baseline Files
- `src/fa_parser_zip_report.py`
- `src/report_writer_light.py`

## First Validation Target
Run an end-to-end export using:

```text
samples/MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip
```

Confirm the generated JSON includes:

- `os_info.hostname`
- `os_info.ip_address`
- `os_info.os_version`
- `os_info.oracle_home`
- `datafiles`
- `growth_rows_raw`
- `registry_list`
- `section53`

## Next Engineering Step
The Python environment has been created with `uv` and a clean export now succeeds from the fresh folder.

Generated files are under:

```text
output/
```

Current validation results:

- DOCX generated successfully.
- Parsed JSON generated successfully.
- Registry parser returns 13 valid components.
- Invalid objects count is 367.
- PDF was not generated in the automated run.

Next, manually inspect the generated DOCX and continue v2 feature work section by section.

## UI V2
The modern UI is implemented separately from the parser:

- `src/app_gui.py` contains the CustomTkinter interface.
- `src/report_service.py` contains request validation and export orchestration.
- `src/fa_parser_zip_report.py` remains the export backend and legacy GUI fallback.

Run the new UI:

```powershell
cd C:\mfec\pm_document_converter_fresh
.\.venv\Scripts\Activate.ps1
python .\src\app_gui.py
```

The service layer was smoke-tested directly with the sample ZIP and generated DOCX/JSON successfully.

## Validation Scripts
Use these scripts after parser, report writer, or UI export changes:

```powershell
C:\mfec\pm_document_converter_fresh\.venv\Scripts\python.exe C:\mfec\pm_document_converter_fresh\scripts\export_and_validate_sample.py
```

This command regenerates the sample report and then runs all validations.

To validate the current output without regenerating:

```powershell
C:\mfec\pm_document_converter_fresh\.venv\Scripts\python.exe C:\mfec\pm_document_converter_fresh\scripts\run_all_validations.py
```

Or run each validator individually:

```powershell
C:\mfec\pm_document_converter_fresh\.venv\Scripts\python.exe C:\mfec\pm_document_converter_fresh\scripts\validate_sample_export.py
C:\mfec\pm_document_converter_fresh\.venv\Scripts\python.exe C:\mfec\pm_document_converter_fresh\scripts\validate_docx_structure.py
C:\mfec\pm_document_converter_fresh\.venv\Scripts\python.exe C:\mfec\pm_document_converter_fresh\scripts\validate_report_quality.py
```

`validate_sample_export.py` checks parsed JSON values and counts.

`validate_docx_structure.py` checks that the generated Word report contains the required major sections and appendices.

`validate_report_quality.py` checks higher-level content and visual-quality signals: table count, paragraph count, embedded chart count, generated chart files, key real values in the DOCX, registry component IDs, and obvious placeholder counts.

`run_all_validations.py` runs all validators in order and stops on the first failure.

`export_and_validate_sample.py` deletes the previous sample outputs, regenerates DOCX/JSON/charts from the sample ZIP, then runs `run_all_validations.py`.
