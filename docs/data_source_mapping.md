# Data Source Mapping

This document maps each generated report area to the Fast Assessment source, parser function, JSON key, and writer function. The goal is to make real ZIP-derived data easy to verify and separate it from UI-entered project metadata.

## Pipeline

```text
Fast Assessment ZIP
  -> process_zip_to_reports()
  -> find_fast_assessment_root()
  -> find_os_dir()
  -> build_all_data()
  -> create_docx_report()
```

## Report Sections

| Report area | ZIP source | Parser / service | JSON key | Writer | Fallback behavior |
|---|---|---|---|---|---|
| Project metadata | UI fields | `ReportMetadata.to_dict()` / GUI metadata | `gui_metadata` | `add_general_project_information()` | Uses defaults or blank fields when UI metadata is not provided |
| Defect Classification | Static report standard | N/A | N/A | `add_defect_classification()` | Static table always rendered |
| Suggestion Summary | Parsed data from multiple checks | `build_suggestion_summary()` | Multiple keys | `add_suggestion_summary()` | Status falls back to OK/blank when source data is missing |
| 2. Oracle minimum requirement | `os/oracle/*.txt` | `parse_os_info()` | `os_info` | `add_oracle_minimum_requirement()` | Empty strings/lists when OS files are missing |
| 3. System Checklist | `os/oracle/*.txt` | `parse_os_info()` | `os_info.hosts_entries`, `os_info.crontab_entries`, CPU/RAM fields | `add_system_checklist()` | Empty tables/placeholders when files are missing |
| 4.1 Database Configuration | `auto_collection/mfec_pm.txt` section `4_1` | `parse_mfec_pm()` | `instance_name`, `rdbms_version`, `archive_mode`, `num_datafiles`, `optimizer_mode`, `disk_space` | `add_database_information()` | Blank or `?` values when the section is missing |
| 4.2 Database Parameter | `auto_collection/mfec_pm.txt` section `4_2` | `parse_mfec_pm()` / `_parse_kv_section()` | `db_parameters` | `add_database_information()` | Empty table when no parameters are parsed |
| 4.4 Database Files | `auto_collection/mfec_pm.txt` section `4_4` | `parse_mfec_pm()` | `datafiles` | `add_database_information()` | Empty table when no datafiles are parsed |
| 4.5 Temporary Files | `auto_collection/mfec_pm.txt` section `4_5` | `parse_mfec_pm()` | `tempfiles` | `add_database_information()` | Empty table when no tempfiles are parsed |
| 4.6 Redo Log File | `auto_collection/mfec_pm.txt` section `4_6` | `parse_mfec_pm()` | `redo_logs` | `add_database_information()` | Empty table when no redo rows are parsed |
| 4.7 Control File | `auto_collection/mfec_pm.txt` control-file section | `parse_mfec_pm()` | `control_files` | `add_database_information()` | Empty table when no control files are parsed |
| 4.8 Database Patch | `log/db_lsinventory.txt` and related log files | `parse_patch_info()` | `patch_info` | `add_database_information()` | Empty/no recommendation rows when patch data is missing |
| 5.1 Performance Review | `auto_collection/mfec_pm.txt`, `auto_collection/mfec_oracle_stats.log`, statspack files | `parse_mfec_pm()`, `parse_oracle_stats()`, `parse_section53_data()` | `library_cache_hit_ratio`, `oracle_stats`, `section53`, `archive_mode` | `add_performance_review()` | Uses blank values and conservative status text |
| 5.2 Database Growth Rate | `data_growth_log/*_segment_size.txt`, datafile allocation from `mfec_pm.txt` | `parse_data_growth_logs()`, `_enrich_growth_with_allocated()` | `growth_rows_raw` | `add_database_growth_rate()` | Chart/table skipped or rendered with available rows |
| 5.3 Performance Analysis | Statspack/report files, preferring `report/*_top1.lst` for PGA/SGA advisory; fallback to `auto_collection/mfec_oracle_stats.log` | `parse_section53_data()` | `section53` | `add_performance_analysis()` | Empty sections or placeholders when source data is missing |
| 6. Tablespace Free Space | `auto_collection/mfec_pm.txt` section `6_1` | `parse_mfec_pm()` | `tablespace_free` | `add_tablespace_free_space()` | Empty table when no rows are parsed |
| 7. Default tablespace and temporary tablespace | `auto_collection/mfec_pm.txt` section `7_1` | `parse_mfec_pm()` | `default_tablespaces` | `add_default_tablespace()` | Empty table when no rows are parsed |
| 8. Database Registry | `auto_collection/registry.log` | `parse_registry()` | `registry_list` | `add_database_registry()` | Empty table when registry log is missing or unparsable |
| Appendix A Invalid Object | `auto_collection/invalid_objects.log` | `parse_invalid_objects()` | `invalid_objects_list` | `add_appendix_invalid_objects()` | Empty table when no invalid objects are parsed |
| Appendix B Alert log | `log/alert_*.log` | Existing appendix writer reads from parsed/available data | Multiple keys/source files | `add_appendix_defaults()` | Placeholder text when no alert data is available |
| Appendix C SQL statements to investigate | Top SQL/statspack sources | `parse_section53_data()` | `section53.top_sql_disk`, `section53.top_sql_mem` | `add_appendix_defaults()` | Placeholder text when no SQL data is available |
| Appendix D Operating system log | `os/oracle/*.txt` | `parse_os_info()` and appendix defaults | `os_info` | `add_appendix_defaults()` | Placeholder text when no OS log data is available |
| Appendix E Backup Configuration | Backup/export/datapump logs under Fast Assessment tree | `parse_backup_info()` | `backup_info` | `add_appendix_defaults()` | `No backup configuration data found.` |
| Appendix F OS Performance Summary | `sar`, OSWatcher, and OS performance files | `parse_os_perf_data()` | `os_perf` | `add_appendix_defaults()` | Clear placeholder when OS performance data is missing |

## Sample Validation Expectations

For `samples/MFEC_PM_MotifPRDDBN_Q2_2025_20260109.zip`, the current expected parsed values are:

| Check | Expected |
|---|---:|
| `os_info.hostname` | `MotifPRDDBN` |
| `os_info.ip_address` | `192.1.255.190` |
| `os_info.oracle_home` | `/u01/app/oracle/product/19.0.0/db_1` |
| `datafiles` | `35` |
| `registry_list` | `13` |
| `invalid_objects_list` | `367` |
| `growth_rows_raw` | `5` |

Run:

```powershell
python .\scripts\validate_sample_export.py
```
