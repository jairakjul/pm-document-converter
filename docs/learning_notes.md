# Learning Notes

## Task
Inspect the PM Document Converter project status and validate the next v2 target, Section 2 Oracle minimum requirement, against the real `MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip` sample data.

## What I Learned
The active codebase is ahead of the project summary: `fa_parser_zip_report.py` already parses OS/environment files, and `report_writer_light.py` already renders Section 2 Oracle minimum requirement and Section 3 System Checklist. The existing parsed output contains real values for hostname, IP address, OS version, kernel version, Oracle owner, Oracle home, Oracle SID, RAM, swap, `/tmp`, disk usage, kernel parameters, hosts entries, and crontab entries.

## Key Decisions
No parser or report writer code was changed for Section 2 because the feature already exists and the real parsed JSON shows the expected fields are populated. A validation-first approach is safer here because v1 is marked stable and unnecessary edits would increase regression risk.

## Mistakes / Risks to Watch
The configured `venv312` currently points to a missing Windows Store Python executable, so direct Python validation could not run in this environment. Future tests should recreate or repair the virtual environment before running the export pipeline. Some generated report sections use legacy OS key names in a few places, so future Appendix F work should normalize OS field names before adding more logic.

## Reusable Patterns
Use resilient parser defaults for optional Fast Assessment files: return empty strings, empty dictionaries, or empty lists instead of raising when files are missing. Keep parsing in `fa_parser_zip_report.py` and Word rendering in `report_writer_light.py`.

## Next Improvement
Repair the Python environment, then run a fresh end-to-end export to a clean output folder and compare the generated DOCX against the expected Section 2 content.

## Task
Create a clean fresh-start workspace for the PM Document Converter while preserving the current working parser and report writer as the baseline.

## What I Learned
The original project root has working code mixed with backups, generated reports, PyInstaller outputs, temporary extracted ZIP data, and test files. The active baseline files are `fa_parser_zip_report.py` and `report_writer_light.py`.

## Key Decisions
Created `C:\mfec\pm_document_converter_fresh` with a simple structure: `src`, `docs`, `samples`, and `output`. Copied only the active source files and the real sample ZIP so future work has a clean boundary.

## Mistakes / Risks to Watch
Starting from absolute zero would throw away useful working parser fixes. The better reset is a clean workspace that keeps the proven baseline and removes clutter from daily development.

## Reusable Patterns
Keep generated files and source files separated. Use `samples` for inputs, `output` for generated reports, `src` for implementation, and `docs` for project notes.

## Next Improvement
Repair or recreate the Python environment inside the fresh folder, then run a clean end-to-end export before making new v2 changes.

## Task
Set up the fresh Python environment, run a clean export, and fix the registry parser issue found during validation.

## What I Learned
The machine only exposed a broken Windows Store Python stub, so the fresh workspace needed a project-local Python runtime. `uv` can install Python and create the virtual environment when its cache and install directories are pointed inside the project. The clean export also revealed that `registry.log` contains continuation rows that look like table data but are not Oracle registry components.

## Key Decisions
Installed Python 3.12 under `.uv_python`, created `.venv`, installed dependencies from `requirements.txt`, and validated the pipeline with the real sample ZIP. Tightened `parse_registry` so it only accepts rows with a component id, version, and real registry status.

## Mistakes / Risks to Watch
The PDF was not created during automated export, so PDF conversion still depends on Microsoft Word COM or LibreOffice availability. Registry parsing should avoid generic length filters because real component rows are long fixed-width SQL*Plus rows.

## Reusable Patterns
For SQL*Plus reports with continuation lines, parse the fixed-width table but validate semantic fields before accepting a row. For this registry report, `COMP_ID`, `VERSION`, and `STATUS` are the minimum safe fields.

## Next Improvement
Open the generated DOCX and compare visual formatting against the reference report, then continue with the next v2 section only after confirming the baseline output is acceptable.

## Task
Build a better CustomTkinter UI and separate the GUI from the export backend.

## What I Learned
The current export pipeline can be wrapped cleanly without changing parser or report writer behavior. A small service layer gives the GUI one stable function to call and keeps validation/error handling out of the visual layout code.

## Key Decisions
Added `src/report_service.py` for `ExportRequest`, `ReportMetadata`, validation, and export orchestration. Added `src/app_gui.py` as the new CustomTkinter UI with sidebar navigation, input cards, progress log, export status, background-thread export, and user-facing error messages.

## Mistakes / Risks to Watch
Long-running report export must stay off the UI thread or Windows will show the app as not responding. The UI currently reports coarse progress because the parser backend does not emit detailed stage callbacks yet.

## Reusable Patterns
Keep desktop UI code thin: collect input, validate it, call a service function, and display results. Use a queue to safely pass worker-thread events back to the CustomTkinter main thread.

## Next Improvement
Launch `python .\src\app_gui.py`, test the UI manually, then add finer-grained progress callbacks inside the parser if users need more detailed status.

## Task
Create data-source mapping documentation and a repeatable sample export validator.

## What I Learned
The project needs explicit traceability because some report values come from the real Fast Assessment ZIP while project/admin metadata comes from UI input or defaults. A small validator is enough to catch important parser regressions before UI or packaging work.

## Key Decisions
Added `docs/data_source_mapping.md` to map report sections to ZIP sources, parser functions, JSON keys, and writer functions. Added `scripts/validate_sample_export.py` to check the current sample export for expected counts and key values.

## Mistakes / Risks to Watch
A generated DOCX can exist even when a parser silently returns bad rows. JSON validation should run after changes that touch parsing, report generation, or sample data.

## Reusable Patterns
Keep validation checks close to known sample expectations: source identity fields, row counts, and required component IDs are high-signal checks that catch real regressions without overfitting every report detail.

## Next Improvement
Run the validator after every parser/report change, then add focused checks for any newly implemented report section.

## Task
Add a DOCX structure validator for the generated PM report.

## What I Learned
JSON validation proves the parser output is sane, but the final deliverable is the DOCX. The report writer can still regress by skipping, renaming, or failing to render sections even when JSON data is correct.

## Key Decisions
Added `scripts/validate_docx_structure.py` to inspect the generated Word document and verify required sections and appendices are present.

## Mistakes / Risks to Watch
This validator checks structure, not visual quality. It will not catch poor table layout, bad page breaks, missing images, or formatting differences from the reference report.

## Reusable Patterns
Validate generated documents at two levels: parsed data correctness with JSON checks, and rendered document structure with DOCX text checks.

## Next Improvement
Add visual/manual review notes for table readability, chart placement, and section formatting against the reference report.

## Task
Add automated report quality and content-accuracy checks.

## What I Learned
Structure validation confirms section names, but it does not prove the report has enough tables, charts, real values, or acceptable placeholder levels. A second validator can catch obvious visual/content regressions before manual review.

## Key Decisions
Added `scripts/validate_report_quality.py` to inspect the generated DOCX and parsed JSON together. It checks document size, paragraph count, table count, embedded image count, chart files, key ZIP-derived values, registry IDs, and placeholder counts.

## Mistakes / Risks to Watch
Automated quality checks are still heuristic. They cannot replace opening the DOCX and reviewing layout, page breaks, table readability, and chart placement against the reference report.

## Reusable Patterns
Use layered validation: JSON correctness, DOCX section structure, then DOCX quality/content signals. Each layer catches a different class of issue.

## Next Improvement
Perform manual visual review of the generated DOCX and record specific formatting gaps as actionable tasks.

## Task
Add a single validation runner for all current checks.

## What I Learned
Having multiple validators is useful, but developers can forget to run one. A single runner makes the validation workflow simpler and less error-prone.

## Key Decisions
Added `scripts/run_all_validations.py` to execute JSON sample validation, DOCX structure validation, and report quality validation in sequence.

## Mistakes / Risks to Watch
The runner validates the current generated output. It does not regenerate the report first, so stale output can still pass if code changed but export was not rerun.

## Reusable Patterns
Use a small orchestration script for local quality gates. Keep individual validators independent so failures remain easy to debug.

## Next Improvement
Add an optional export-and-validate command that regenerates the report before running validators.

## Task
Add an export-and-validate command for the sample ZIP.

## What I Learned
Validation is more reliable when the report is regenerated first. Otherwise, validators can pass against stale output after source code changes.

## Key Decisions
Added `scripts/export_and_validate_sample.py` to clean previous sample outputs, export from the real sample ZIP, and run all validators through `run_all_validations.py`.

## Mistakes / Risks to Watch
The script intentionally deletes only known generated sample files. Keep that list narrow so it does not remove unrelated user output from the `output` folder.

## Reusable Patterns
Use a single quality-gate command that rebuilds the artifact and validates it immediately. This is the local equivalent of a CI smoke test.

## Next Improvement
Use `export_and_validate_sample.py` as the default check before packaging or making report-format changes.

## Task
Compare Section 5.2 Database Growth Rate against the reference screenshot and fix the real-data mismatch.

## What I Learned
The Fast Assessment growth log contains two useful totals in each monthly file: `ALL SEGMENTS MB` for used size and `ALL DATA FILE MB` for allocated size. The previous parser used historical used size correctly but applied the latest allocated size to every month, which made the chart allocation line flat.

## Key Decisions
Updated `parse_data_growth_logs()` to parse monthly allocated size from the real `###size data file all` block. Updated Section 5.2 rendering to display the latest four months, matching the reference range from Oct-2025 through Jan-2026.

## Mistakes / Risks to Watch
The existing generated DOCX in `output` may be locked if it is open in Word, so the export-and-validate script cannot replace it until the document is closed. For comparison work, a separate `output_compare` folder can be used safely.

## Reusable Patterns
When a report value appears wrong, trace it back to the raw ZIP file before changing formulas. The correct fix was to parse the existing monthly allocation total, not hardcode reference values.

## Next Improvement
Close the open generated DOCX, rerun `export_and_validate_sample.py`, and then manually compare Section 5.2 chart/table formatting in Word against the reference screenshot.

## Task
Compare PGA/SGA memory advisory output against the reference screenshots and align data source/style.

## What I Learned
The reference PGA/SGA screenshots match the Statspack `sp_75546_75570_top1.lst` file, not the `mfec_oracle_stats.log` advisory query output. Both are real ZIP sources, but the reference report uses the Statspack top1 advisory values.

## Key Decisions
Updated `parse_section53_data()` to prefer Statspack `*_top1.lst` for PGA and SGA advisory data, with `mfec_oracle_stats.log` retained as fallback. Updated PGA/SGA report rendering to use MB-based current rows, dark advisory charts, yellow current markers, centered advisory tables, and reference-style narrative text.

## Mistakes / Risks to Watch
Different raw ZIP files can contain similar advisory data with different snapshot values. For visual/reference matching, use the same source as the reference document, not just any valid source.

## Reusable Patterns
When multiple real sources exist for the same metric, document source precedence and keep a fallback parser. For report parity work, compare values against raw source files before changing formatting.

## Next Improvement
Close any open generated DOCX, rerun the full export-and-validate command into `output`, and manually inspect the PGA/SGA pages in Word.

## Task
Compare Appendix E Backup Configuration against the reference screenshot and prevent RMAN logs from being used as the primary Appendix E content.

## What I Learned
The current sample ZIP does not include the Datapump export job log shown in the reference screenshot. It contains RMAN backup logs plus a crontab entry for `run_expdp_icom.sh`, so showing the RMAN log creates a visual and content mismatch.

## Key Decisions
Updated `parse_backup_info()` to prefer real Datapump export logs when present, summarize them into a compact report block, and only use RMAN as a last fallback. When the export log is missing, Appendix E now shows the Datapump schedule evidence from crontab instead of filling the report with RMAN backup-set pages.

## Mistakes / Risks to Watch
Do not fabricate Datapump export details from the reference document when the ZIP does not contain that log. If the customer expects the full Datapump export summary, the source ZIP must include the export log file.

## Reusable Patterns
Use source scoring when several files can satisfy one report section. Prefer the evidence type that matches the report requirement, then provide a transparent fallback when the exact source is missing.

## Next Improvement
Get a Fast Assessment ZIP that includes the actual `expdp_*.log` file and verify Appendix E against the reference Datapump export page.

## Task
Compare Appendix F OS Performance Summary against the reference screenshots and improve the CPU/run-queue section.

## What I Learned
The sample ZIP does not include exported OSWatcher images, but the SAR files contain 10-minute CPU and run-queue samples. Those samples are enough to recreate OSWatcher-style spike charts instead of relying only on daily averages.

## Key Decisions
Extended `parse_os_perf_data()` to keep raw `cpu_samples` and `runq_samples` while preserving the existing daily summaries for memory and disk. Updated Appendix F to render a host summary table, run queue chart, CPU utilization/system/user charts, and high/very-high threshold tables calculated from real SAR samples.

## Mistakes / Risks to Watch
Daily averages hide short CPU saturation periods and run queue spikes. For performance sections, use raw sample-level data for bottleneck analysis and keep averages only as supporting context.

## Reusable Patterns
When reference charts are image exports but the ZIP only has raw text metrics, recreate the chart style from the raw samples instead of embedding placeholders or simplified daily charts.

## Next Improvement
If future ZIPs include full OSWatcher output across the quarter, parse the full date range so Appendix F can show September through January like the reference screenshots.

## Task
Run the current test checks after Appendix E and Appendix F report improvements.

## What I Learned
The regenerated comparison report passes compile, export, quality/content validation, and DOCX structure validation. The current report now embeds nine chart images, including the new OS performance charts.

## Key Decisions
Used `output_compare` for regeneration to avoid the locked DOCX in the main `output` folder. This keeps testing repeatable while Word may still have the main generated report open.

## Mistakes / Risks to Watch
The structure validator currently checks the default `output` report, while the quality validator was pointed to `output_compare`. When doing a final release check, close Word and run the full export-and-validate command against `output`.

## Reusable Patterns
Use a two-layer test pass after report changes: syntax/export validation first, then DOCX content and structure validation.

## Next Improvement
Close the generated DOCX in Word and rerun `scripts/export_and_validate_sample.py` so the main `output` folder is refreshed with the latest Appendix E and Appendix F changes.

## Task
Compare Appendix F memory and disk performance pages against reference screenshots.

## What I Learned
The SAR files include enough sample-level data to recreate the reference-style Memory Free, Swap Used, Page In, Page Out, CPU Wait I/O, and Disk Service Time charts. Daily averages were technically valid but visually and analytically too weak for this report.

## Key Decisions
Extended OS performance parsing to capture memory, swap, page I/O, and disk samples. Added OSWatcher-style charts for each reference subsection and updated Appendix F to present those charts before the supporting summary tables and narrative.

## Mistakes / Risks to Watch
The generated DOCX can be locked by Word in any output folder. When that happens, export to a fresh comparison folder instead of trying to overwrite the locked file.

## Reusable Patterns
For performance reports, keep both raw samples and aggregate summaries. Raw samples drive spike charts and bottleneck analysis; aggregates are still useful for compact validation tables.

## Next Improvement
Refactor the OS performance chart code into a small dedicated module so Appendix F generation stays readable as more reference-style charts are added.

## Task
Refresh the main `output` folder after closing Word.

## What I Learned
The earlier export failure was caused by Word locking the generated DOCX. After closing Word, the full export-and-validate workflow can clean and regenerate the main `output` folder successfully.

## Key Decisions
Ran `scripts/export_and_validate_sample.py` against the default output folder instead of using a comparison folder, confirming the main deliverable is current.

## Mistakes / Risks to Watch
Windows file locks can block report regeneration when a DOCX is open. Close Word before running the full export workflow.

## Reusable Patterns
Use comparison folders while iterating, then close open documents and run the full output refresh before treating the generated report as final.

## Next Improvement
Open the refreshed DOCX and do a manual visual pass page by page against the reference screenshots.

## Task
Prepare PyInstaller packaging for the CustomTkinter desktop app.

## What I Learned
The GUI can be packaged as a Windows subsystem executable using PyInstaller's windowed bootloader, which starts the app without opening a terminal window.

## Key Decisions
Added a dedicated PyInstaller spec under `packaging/`, a PowerShell build script under `scripts/`, and packaging documentation under `docs/`. The spec collects CustomTkinter and matplotlib runtime data and excludes unnecessary matplotlib test modules.

## Mistakes / Risks to Watch
PyInstaller builds may require running outside the sandbox because they create many files and inspect runtime binaries. PDF conversion still depends on Microsoft Word or LibreOffice being installed on the target machine.

## Reusable Patterns
Keep packaging configuration separate from application code. Use a repeatable build script and smoke-test the generated executable by checking the Windows subsystem and launching it briefly.

## Next Improvement
Run a full manual packaged-app export test: launch `dist\PMDocumentConverter\PMDocumentConverter.exe`, select the real ZIP, export to a fresh folder, and open the generated DOCX.

## Task
Review CircleCI readiness for the project.

## What I Learned
The project does not currently have a `.circleci/config.yml`, so there is no CircleCI pipeline to review yet. The existing validation scripts are already good CI entrypoints.

## Key Decisions
Recommend a simple CI pipeline that installs dependencies, compiles Python files, runs `scripts/export_and_validate_sample.py`, and stores generated DOCX/chart artifacts for review.

## Mistakes / Risks to Watch
The current sample cleanup script only removes the older chart file list. New Appendix F charts should be added to the cleanup list before relying on CI artifacts.

## Reusable Patterns
CI should run the same command developers run locally. Keep one quality gate script as the source of truth and have CI call it.

## Next Improvement
Add `.circleci/config.yml` after deciding whether CI should build only validation artifacts or also build the Windows PyInstaller executable on a Windows runner.

## Task
Reduce the packaged desktop app size.

## What I Learned
The largest packaged dependencies are unavoidable chart/runtime components: numpy OpenBLAS, matplotlib, Python, lxml, and Tk. However, some unused PyInstaller output can be removed safely after build.

## Key Decisions
Removed broad matplotlib data collection, forced the matplotlib `Agg` backend, excluded unused GUI/test modules, and added post-build cleanup for Pythonwin, PIL AVIF support, and numpy metadata.

## Mistakes / Risks to Watch
Do not remove numpy, matplotlib, lxml, Tk, or core PIL files because report generation, chart rendering, DOCX writing, and the CustomTkinter UI depend on them.

## Reusable Patterns
Measure package folders before optimizing. Remove only files that have a clear reason to be unused, then rebuild and smoke-test the executable.

## Next Improvement
Run one full export through the packaged executable to confirm chart generation still works after the size cleanup.

## Task
Verify the downloaded macOS build artifact structure from Windows.

## What I Learned
The downloaded file is a ZIP that contains the GitHub Actions artifact ZIP, and that inner ZIP contains a structurally valid `PMDocumentConverter.app` bundle.

## Key Decisions
Verified bundle structure non-destructively by inspecting the ZIP contents, `Info.plist`, `_CodeSignature`, Mach-O header, and expected executable/resource paths instead of extracting or modifying the artifact.

## Mistakes / Risks to Watch
Structural verification from Windows is not the same as a runtime validation on macOS. This check cannot confirm launching, Gatekeeper acceptance, or notarization status.

## Reusable Patterns
For macOS artifacts downloaded from CI, expect one extra ZIP wrapper around the real artifact. Verify the app bundle by checking `Contents/Info.plist`, `Contents/MacOS/<app>`, `Contents/_CodeSignature`, and a Mach-O magic header.

## Next Improvement
Open `PMDocumentConverter.app` on a real Mac and confirm it launches, exports a DOCX, and is not blocked by Gatekeeper.

## Task
Add CircleCI automation for validation, Windows packaging, and artifacts.

## What I Learned
CircleCI can protect the project from broken exports by running the same sample export validation used locally. A Windows runner is needed to build the `.exe` artifact.

## Key Decisions
Added `.circleci/config.yml` with a Linux validation job and a Windows packaging job. The Windows job runs export validation before building the executable, then stores both report output and packaged app artifacts.

## Mistakes / Risks to Watch
The Windows packaging job depends on CircleCI Windows executor access. If the account does not have Windows runners enabled, only the Linux validation job will be practical until that access is available.

## Reusable Patterns
Make CI call project scripts instead of duplicating business logic in YAML. This keeps local testing and CI behavior aligned.

## Next Improvement
Push the project to the connected repository and confirm the first CircleCI workflow can access the sample ZIP and Windows executor.

## Task
Prepare macOS `.app` packaging support.

## What I Learned
PyInstaller builds are platform-specific. A Windows machine can prepare the macOS spec and scripts, but a real `.app` bundle must be built on macOS.

## Key Decisions
Added a separate macOS PyInstaller spec using `BUNDLE` and a `build_desktop_macos.sh` script that refuses to run unless the host is macOS.

## Mistakes / Risks to Watch
DOCX generation is cross-platform, but PDF conversion on macOS needs LibreOffice in `PATH` unless macOS-specific Microsoft Word automation is added later.

## Reusable Patterns
Keep Windows and macOS packaging specs separate when platform-specific bundle behavior differs. This avoids adding conditional complexity to one large spec.

## Next Improvement
Run `scripts/build_desktop_macos.sh` on a Mac or macOS CI runner and smoke-test exporting the sample ZIP from the built `.app`.

## Task
Add GitHub Actions automation for Linux validation, Windows packaging, and macOS `.app` packaging.

## What I Learned
GitHub Actions can provide Apple-hosted macOS runners, which solves the need for a physical Mac while still building the `.app` on the correct operating system.

## Key Decisions
Added `.github/workflows/build.yml` with three jobs: Linux validation, Windows executable packaging, and macOS `.app` packaging. The macOS job zips the `.app` bundle with `ditto` before uploading it as an artifact.

## Mistakes / Risks to Watch
The macOS build cannot be verified on Windows. The first real validation must happen after pushing to GitHub and running the workflow on a macOS runner.

## Reusable Patterns
Use platform-specific CI jobs for platform-specific build artifacts. Keep validation shared by calling `scripts/export_and_validate_sample.py` on each runner.

## Next Improvement
Push to GitHub, run the workflow manually, and download the `PMDocumentConverter-macos` artifact to smoke-test on macOS.

## Task
Attempt to push the project to GitHub and run the macOS build workflow.

## What I Learned
The local project folder is not currently a Git repository, and the connected GitHub account has no accessible installed repositories through the connector. The GitHub CLI is also not installed locally.

## Key Decisions
Stopped before creating a local repository or guessing a remote target. A push requires a known GitHub repository URL or an installed GitHub connector repository.

## Mistakes / Risks to Watch
Do not initialize and push to an unknown repository. Confirm the target repo first so generated source, CI config, and sample files go to the correct place.

## Reusable Patterns
Before automating CI runs, verify three things: local folder is a Git repo, a remote is configured, and the automation account can access that remote.

## Next Improvement
Create or connect a GitHub repository for `pm_document_converter_fresh`, then run the push and GitHub Actions workflow.

## Task
Compare Section 6 Tablespace Free Space against the reference screenshot and align the generated report output.

## What I Learned
The parser already had the raw inputs needed for the reference table. The mismatch was in the report logic: the old section showed free space inside currently allocated files, while the reference table shows used space from allocated files and remaining free space against each tablespace's maximum capacity.

## Key Decisions
Kept the parser unchanged and rebuilt Section 6 in `report_writer_light.py`. The new logic aggregates current allocation from `datafiles` and `tempfiles`, derives used MB from the parsed `pct_free` value, derives max MB from each file's max size or current size when autoextend is not available, and renders the full six-column table. Temporary tablespaces are included with used space set to zero. The section now highlights only critical rows and uses a section-specific table layout closer to the reference.

## Mistakes / Risks to Watch
The current row ordering includes a small preferred-order fallback so the sample report matches the reference more closely. If future customer ZIPs introduce new tablespaces, the section will still work, but the tie ordering among rows with the same displayed percentage may differ from a manually prepared legacy report.

## Reusable Patterns
When a report section looks wrong, check whether the parser is actually the problem. If the raw data is already present, prefer fixing the section-level transformation and rendering logic instead of introducing another parsing path.

## Next Improvement
Open the regenerated DOCX and verify the final Section 6 table visually in Word, especially column widths and the critical-row highlight, against the reference screenshot.

## Task
Run the current end-to-end smoke test after the Section 6 adjustment.

## What I Learned
The export pipeline is still stable after the tablespace-section rewrite. The real sample ZIP exported successfully, all validators passed, and the environment created a PDF in this run as well as the DOCX and parsed JSON.

## Key Decisions
Used the existing single-command quality gate, `scripts/export_and_validate_sample.py`, instead of piecemeal checks. That keeps the verification path consistent with how the project is already meant to be tested.

## Mistakes / Risks to Watch
The smoke test proves parser, writer, and validators are working together, but it does not replace manual visual review of the regenerated DOCX in Word. PDF creation can still vary by environment, so treat it as an environment-dependent bonus rather than the primary acceptance check.

## Reusable Patterns
After report-format changes, rerun the full export-and-validate command rather than only checking the changed section. That catches unintended regressions in other parts of the document.

## Next Improvement
Open the latest DOCX and confirm the final Section 6 visual layout against the reference screenshot before moving to the next report section.

## Task
Add a report cover page and align Section 4.8 Database Patch and Section 5.1 Performance Review closer to the reference screenshots.

## What I Learned
The current codebase already had enough raw source data to reproduce a much richer Section 5.1. The problem was not missing source data; it was that the parser flattened most of the section into summary metrics and the writer rendered only a simplified status table. The patch-version mismatch also came from reading the OPatch installer version instead of the actual Database Release Update version from `db_lsinventory.txt`.

## Key Decisions
Added a generated cover-page banner and a new cover-page renderer before the table of contents. Updated `parse_patch_info()` to prefer the real Database Release Update version (`19.15.0.0.220419`) from `db_lsinventory.txt`. Extended the parser to keep detailed 5.1 inputs: library-cache namespace rows, PIN/RELOAD values, shared-pool-reserved stats, and undo-segment rows. Replaced the simplified 5.1 summary table with a reference-style two-column Q&A layout that uses nested tables for the detailed answers.

## Mistakes / Risks to Watch
The cover page is programmatically generated to match the look and structure of the reference, but it is still an approximation rather than the exact original artwork. The manual TOC page numbers are now even more fragile because the new cover page and expanded 5.1 section shift downstream pagination.

## Reusable Patterns
When a report section is visually wrong but the raw source already contains the needed detail, store a richer intermediate structure in the parser and let the writer choose the final presentation. Also separate tool-version parsing from product-version parsing; `OPatch version` is not the same thing as the installed database RU level.

## Next Improvement
Replace the hard-coded table-of-contents page numbers with a safer approach, or regenerate them from the final DOCX structure after layout changes.

## Task
Remove the cover image above the title and make the cover title/date follow the values the user fills in the app.

## What I Learned
The cover title was already partly metadata-driven, but the date was not a real user-editable field. The GUI only generated a change-record date internally, which made the cover date effectively implicit rather than user-controlled.

## Key Decisions
Removed the cover banner image from the rendered first page. Added a new `report_date` field to `ReportMetadata`, exposed it in the CustomTkinter project form, and used it as the first source for the cover date. Kept the cover title driven by report metadata, with `report_type + db_name` as the primary title and metadata fallbacks when those are blank.

## Mistakes / Risks to Watch
The cover title now follows the user metadata more directly, but it still uses a fixed title pattern rather than a fully free-form cover-title field. If users later want a completely custom title string, that should be a separate explicit metadata field instead of overloading `project_name`.

## Reusable Patterns
If a value must appear in the final document exactly as the user expects, expose it as a first-class metadata field in the UI and carry it through the service layer, instead of deriving it indirectly at render time.

## Next Improvement
If needed, add a dedicated `cover_title` field so users can override the default `report_type + db_name` title pattern without changing other metadata.

## Task
Compare and adjust Section 2.2 Oracle requirement output to match the richer reference layout.

## What I Learned
The existing Section 2.2 was functionally correct but too simplified. The reference report does not use a compact pass/fail matrix there; it uses a three-column comparison table with richer current-server details, including a nested disk-space table and explicit OS, RAM, tmp, JDK, and kernel-setting rows.

## Key Decisions
Replaced the old Section 2.2 pass/fail table with a custom three-column table: `Requirement`, `Minimum Requirement`, and `Current Server Specification`. Kept the heading prefixed with `2.2 Compare to Oracle requirement` so existing validators still pass, but appended the hostname for closer parity with the reference. Used a nested disk table inside the current-specification column to mirror the reference structure.

## Mistakes / Risks to Watch
This section now intentionally overlaps somewhat with details that still appear later in `2.3 User's environment`. That duplication is acceptable for report parity, but if the document is later optimized for concision instead of reference matching, Section 2 should be consolidated rather than expanded further.

## Reusable Patterns
When a reference section relies on a mixed layout, use a custom outer table with nested inner tables instead of trying to force everything through one generic helper. Keep validators stable by extending heading text from the existing prefix rather than replacing it outright.

## Next Improvement
Open the regenerated DOCX and compare the Section 2.2 row heights and nested disk-table widths against the screenshot; the content is aligned now, but the remaining gap is visual spacing.

## Task
Build the current Windows EXE app from the latest report-generator code.

## What I Learned
The existing PyInstaller packaging flow still works after the recent cover-page and report-layout changes. No packaging-specific code change was required for this build; the current spec and build script remained valid.

## Key Decisions
Used the existing repeatable packaging entrypoint, `scripts/build_desktop.ps1`, instead of invoking PyInstaller ad hoc. That keeps the build consistent with the documented project workflow and avoids drift between manual builds and future CI packaging.

## Mistakes / Risks to Watch
An EXE build succeeding does not prove the packaged app visually matches the latest DOCX expectations. Packaging should still be followed by a real smoke test in the built GUI, selecting a ZIP and generating output in a fresh folder.

## Reusable Patterns
Keep desktop packaging behind one build script and one spec file. That makes rebuilds predictable after report-writer changes and reduces the chance of “works locally, packaged build broken” drift.

## Next Improvement
Run the packaged EXE once, export a fresh report through the GUI, and confirm the generated DOCX matches the script-based output.

## Task
Improve the Section 2.2 Oracle requirement table layout so the nested disk-space table looks closer to the reference report and stops wrapping unpredictably.

## What I Learned
The main layout problem was not the data itself; it was that the outer three-column table widths exceeded the printable page width. Word compensated by shrinking and wrapping the nested disk table aggressively, which made filesystem and mount-point text break awkwardly.

## Key Decisions
Reduced the outer Section 2.2 column widths to fit the actual page, added explicit cell margins, forced the nested disk table to use fixed column widths, and top-aligned the large cells so the Linux Disk Space row reads more like the reference layout.

## Mistakes / Risks to Watch
Nested tables in `python-docx` are sensitive to total width. If future edits increase outer widths again, Word will silently reflow the nested table and the layout will degrade even though the document still validates structurally.

## Reusable Patterns
For DOCX sections that embed a table inside another table, treat the printable page width as a hard constraint and assign explicit widths to both the parent columns and the nested child columns. Do not rely on Word autofit for dense operational tables.

## Next Improvement
Open the regenerated DOCX in Word and inspect only Section 2.2 at normal zoom; if the mounted-on column is still too tight, trim the disk rows shown there or widen that child column by borrowing space from the numeric columns.

## Task
Rebuild the Windows EXE after the latest Section 2.2 table-layout changes.

## What I Learned
The packaging flow remained stable after the nested-table layout changes. The EXE can be rebuilt without touching the PyInstaller spec, which means the report-writer changes are still compatible with the existing desktop packaging setup.

## Key Decisions
Used the existing `scripts/build_desktop.ps1` entrypoint again and verified the produced EXE by checking its path, size, and write time instead of assuming the build output was current.

## Mistakes / Risks to Watch
Packaging success only proves the app bundles correctly. It does not prove the GUI export path or final Word output still looks right, so a packaged-app smoke test remains the next real acceptance check.

## Reusable Patterns
After report-writer layout changes, rebuild the packaged app from the same scripted entrypoint and verify the artifact metadata immediately. That catches stale-dist confusion before manual testing starts.

## Next Improvement
Launch the rebuilt EXE, export one report through the GUI into a clean folder, and compare the generated DOCX against the script-generated output.

## Task
Fix the Section 2.2 nested disk table overflowing past the right page edge in Word.

## What I Learned
Even after the first layout cleanup, the parent and child table widths were still too large in combination. In Word, nested tables do not respect the intended visual width if the total requested width exceeds the actual parent cell width, so the child table spilled past the page boundary.

## Key Decisions
Reduced the Section 2.2 outer table widths again, shrank the nested disk-table column widths, reduced cell padding, and lowered the nested-table font slightly. The goal was not to change content, only to keep the whole structure inside the printable area.

## Mistakes / Risks to Watch
This section is near the practical limit of how much dense filesystem data fits into a nested Word table. If more disk rows or longer mount paths appear in future samples, the layout may degrade again unless the section is simplified or split.

## Reusable Patterns
For dense DOCX tables, solve overflow by reducing total width first, then margins, then font size. Relying on Word to auto-resolve an oversized nested table produces unstable results across machines.

## Next Improvement
Open the regenerated DOCX and check the right edge of Section 2.2 in Word. If any row still clips, the next step is to abbreviate the mounted path column or split Linux Disk Space into its own full-width table.

## Task
Rebuild the EXE again after the overflow fix and handle the locked `dist` folder caused by the running packaged app.

## What I Learned
PyInstaller could not rebuild while the previous `PMDocumentConverter.exe` process was still running from the `dist` folder. The failure was not a packaging/config issue; it was a standard Windows file lock on an in-use packaged binary and one of its bundled modules.

## Key Decisions
Identified the running `PMDocumentConverter` process, stopped it, and reran the existing build script. Verified the rebuilt EXE by checking its file size and fresh write time instead of assuming the second build had replaced the earlier artifact.

## Mistakes / Risks to Watch
If the packaged app is left open, any rebuild into the same `dist` directory can fail with `PermissionError` during the `COLLECT` phase. That failure looks noisy in PyInstaller output, but the root cause is simply that the old app instance still owns files under `dist`.

## Reusable Patterns
When a Windows packaged-app rebuild fails at `Removing dir ... dist\\...` or `COLLECT`, check for a running instance of the same EXE before changing build scripts. Treat `dist` locks as an environment/process issue first, not a packaging regression.

## Next Improvement
Run the rebuilt EXE once and export a report through the GUI so the packaged app path is validated after the latest Section 2.2 layout fixes.

## Task
Adjust the Table of Contents styling and add a footer with last-update date, confidentiality text, and automatic page number.

## What I Learned
The generated TOC was structurally useful but visually different from the reference. The reference uses a boxed TOC title, smaller black TOC entries with dot leaders, and a footer on report pages that includes last-update date, confidentiality text, and page number.

## Key Decisions
Changed the TOC title from a standalone large paragraph to a bordered one-cell table, added missing subsection rows, tightened TOC font sizes and spacing, and added a section footer with a Word `PAGE` field. The first cover page uses a different first-page footer setting so the footer starts on the report pages.

## Mistakes / Risks to Watch
The TOC is still manually generated with hard-coded page numbers. The footer page number updates automatically in Word, but TOC page numbers do not. Future layout changes can make the TOC page numbers stale unless the report writer is changed to use real Word TOC fields or a post-generation page-number update workflow.

## Reusable Patterns
For reference-style DOCX output, use simple Word structures that behave predictably: a one-cell table for boxed titles, tab stops with dot leaders for TOC rows, and Word field XML for dynamic page numbers.

## Next Improvement
Replace the manual TOC page-number list with a real Word TOC field or a reliable post-generation update step if exact pagination becomes a release requirement.
