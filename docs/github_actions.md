# GitHub Actions

## Workflow

The workflow file is:

```text
.github/workflows/build.yml
```

It runs on:

- Push to `main`
- Pull requests
- Manual `workflow_dispatch`

## Jobs

### `validate-report`

Runs on Ubuntu:

- Installs dependencies.
- Compiles Python files.
- Runs `scripts/export_and_validate_sample.py`.
- Uploads generated report output as `report-output-linux`.

Sample export validation runs only when a ZIP file exists under `samples/`.

### `build-windows`

Runs on Windows:

- Runs sample export validation.
- Builds `dist/PMDocumentConverter/PMDocumentConverter.exe`.
- Verifies the executable uses the Windows GUI subsystem.
- Uploads the packaged Windows app as `PMDocumentConverter-windows`.

### `build-macos`

Runs on Apple-hosted macOS:

- Runs sample export validation.
- Builds `dist/PMDocumentConverter.app`.
- Zips the `.app` bundle with `ditto`.
- Uploads the app bundle as `PMDocumentConverter-macos`.

## Important Notes

PyInstaller cannot cross-compile. The Windows app must be built on Windows, and the macOS `.app` must be built on macOS.

The macOS artifact is zipped because `.app` bundles are directories and should be downloaded as a single archive.

PDF export still depends on LibreOffice or Microsoft Word availability on the runner. The validation expects DOCX generation and does not require PDF output.

The real sample ZIP is intentionally ignored by Git. Add a sanitized sample ZIP under `samples/` if you want GitHub Actions to run the full export validation on every push.
