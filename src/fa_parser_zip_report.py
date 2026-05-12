"""
FA Parser ZIP Report
รับไฟล์ .zip จาก fast_assessment แล้วสร้างรายงาน Word (.docx) และพยายามแปลงเป็น PDF (.pdf)

วิธีใช้แบบ CLI:
    fa_parser_zip_report.exe input.zip output_folder

ถ้าไม่ใส่ argument โปรแกรมจะเปิดหน้าต่างให้เลือกไฟล์ zip และโฟลเดอร์ปลายทาง

หมายเหตุ:
- DOCX สร้างได้ด้วย python-docx
- PDF จะสร้างได้ถ้าเครื่อง Windows มี Microsoft Word ติดตั้งอยู่ หรือมี LibreOffice อยู่ใน PATH
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

# ================================================================
# Original parser logic from fa_parser.py
# ================================================================


def parse_sqlplus_table(text: str) -> list[dict]:
    lines = [ln.rstrip() for ln in text.splitlines()]
    results = []

    sep_idx = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped and re.match(r"^[-\s]+$", stripped) and "-" in stripped:
            sep_idx = i
            break

    if sep_idx is None or sep_idx == 0:
        return results

    sep_line = lines[sep_idx]
    header_line = None
    for i in range(sep_idx - 1, -1, -1):
        if lines[i].strip():
            header_line = lines[i]
            break

    if header_line is None:
        return results

    col_spans = []
    for m in re.finditer(r"-+", sep_line):
        col_spans.append((m.start(), m.end()))

    if not col_spans:
        return results

    headers = []
    for start, end in col_spans:
        name = header_line[start:end].strip() if len(header_line) >= end else header_line[start:].strip()
        headers.append(name)

    for line in lines[sep_idx + 1 :]:
        if re.match(r"^[-\s]+$", line.strip()) and "-" in line:
            continue
        if re.search(r"\d+ rows? selected", line, re.IGNORECASE):
            break
        if re.match(r"^\*8\*|^>O<|^\^o\^", line.strip()):
            break
        if not line.strip():
            continue

        row = {}
        for i, (start, end) in enumerate(col_spans):
            if i < len(headers):
                val = line[start:end].strip() if len(line) >= end else line[start:].strip()
                row[headers[i]] = val
        if row:
            results.append(row)

    return results


def _split_mfec_pm_sections(path: Path) -> dict:
    content = path.read_text(encoding="utf-8", errors="replace")
    sections = {}
    parts = re.split(r">O<#+(\w+)#+>O<", content)
    for i in range(1, len(parts) - 1, 2):
        sec_id = parts[i].strip()
        sec_body = parts[i + 1] if i + 1 < len(parts) else ""
        sections[sec_id] = sec_body
    return sections


def _parse_kv_section(section_text: str) -> dict:
    rows = parse_sqlplus_table(section_text)
    result = {}
    for row in rows:
        name = row.get("NAME", "").strip()
        value = row.get("VALUE", "").strip()
        if name:
            result[name] = value
    return result


def parse_mfec_pm(auto_collection_dir: Path) -> dict:
    path = auto_collection_dir / "mfec_pm.txt"
    if not path.exists():
        return {}

    sections = _split_mfec_pm_sections(path)
    result = {}

    sec41 = sections.get("4_1", "")

    def _find_subsec(keyword: str) -> str:
        for chunk in re.split(r"\^o\^[^\^]+\^o\^", sec41):
            if keyword.lower() in chunk.lower():
                return chunk
        m = re.search(rf"{keyword}.*?(?=\^o\^|\*8\*|$)", sec41, re.DOTALL | re.IGNORECASE)
        return m.group(0) if m else ""

    inst_rows = parse_sqlplus_table(_find_subsec("Instance Name"))
    result["instance_name"] = inst_rows[0].get("VALUE", "") if inst_rows else ""

    ver_rows = parse_sqlplus_table(_find_subsec("RDBMS Version"))
    if ver_rows:
        banner = ver_rows[0].get("BANNER", "") or ver_rows[0].get("BANNER_FULL", "")
        result["rdbms_version"] = banner.strip()
    else:
        result["rdbms_version"] = ""

    arch_text = _find_subsec("Archive log mode")
    result["archive_mode"] = "ARCHIVELOG" if "archivelog" in arch_text.lower() else "NOARCHIVELOG"

    df_m = re.search(r"(\d+)\s*\n.*(?:data|dbfile)", sec41, re.IGNORECASE)
    result["num_datafiles"] = df_m.group(1) if df_m else "?"

    tbs_rows = parse_sqlplus_table(_find_subsec("Tbs. size"))
    result["tbs_size_mb"] = {}
    for r in tbs_rows:
        cols = list(r.values())
        if len(cols) >= 2:
            try:
                float(cols[0])
                result["tbs_size_mb"][cols[1].strip()] = cols[0].strip()
            except ValueError:
                result["tbs_size_mb"][cols[0].strip()] = cols[-1].strip()

    opt_m = re.search(r"optimizer_mode.*?(\w+_\w+|\w+)\s*$", _find_subsec("Optimizer Mode"), re.IGNORECASE | re.MULTILINE)
    result["optimizer_mode"] = opt_m.group(1) if opt_m else "ALL_ROWS"

    disk_m = re.search(r"(\d[\d,.]+)\s+(?:GB|MB)?", _find_subsec("db.files capa"), re.IGNORECASE)
    result["disk_space"] = disk_m.group(1).replace(",", "") if disk_m else ""

    sec42 = sections.get("4_2", "")
    result["db_parameters"] = _parse_kv_section(sec42)

    sec44 = sections.get("4_4", "")
    df_rows = parse_sqlplus_table(sec44)
    result["datafiles"] = []
    for r in df_rows:
        raw_mb = r.get("MB", r.get("BYTES/1024/1024", r.get("SIZE_MB", "")))
        raw_maxmb = r.get("MAXMB", r.get("MAXBYTES/1024/1024", r.get("MAX_MB", "0")))
        raw_aut = r.get("AUT", r.get("AUTOEXTENSIBLE", r.get("AUTO_EXT", "NO")))
        raw_inc = r.get("INCREMENT_BY", r.get("INC_MB", "0"))
        # Convert INCREMENT_BY from Oracle blocks (8KB each) to MB
        try:
            inc_blocks = float(raw_inc)
            inc_mb_val = round(inc_blocks * 8 / 1024, 4) if inc_blocks > 0 else 0
        except (ValueError, TypeError):
            inc_mb_val = raw_inc
        result["datafiles"].append(
            {
                "tbs_name": r.get("TABLESPACE_NAME", ""),
                "file_name": r.get("FILE_NAME", ""),
                "size_mb": raw_mb,
                "max_mb": raw_maxmb,
                "auto_ext": raw_aut,
                "inc_mb": str(inc_mb_val),
            }
        )

    sec45 = sections.get("4_5", "")
    tmp_rows = parse_sqlplus_table(sec45)
    result["tempfiles"] = []
    for r in tmp_rows:
        raw_mb = r.get("MB", r.get("BYTES/1024/1024", r.get("SIZE_MB", "")))
        raw_maxmb = r.get("MAXMB", r.get("MAXBYTES/1024/1024", r.get("MAX_MB", "0")))
        raw_aut = r.get("AUT", r.get("AUTOEXTENSIBLE", r.get("AUTO_EXT", "NO")))
        raw_inc = r.get("INCREMENT_BY", r.get("INC_MB", "0"))
        try:
            inc_blocks = float(raw_inc)
            inc_mb_val = round(inc_blocks * 8 / 1024, 4) if inc_blocks > 0 else 0
        except (ValueError, TypeError):
            inc_mb_val = raw_inc
        result["tempfiles"].append(
            {
                "tbs_name": r.get("TABLESPACE_NAME", ""),
                "file_name": r.get("FILE_NAME", ""),
                "size_mb": raw_mb,
                "max_mb": raw_maxmb,
                "auto_ext": raw_aut,
                "inc_mb": str(inc_mb_val),
            }
        )

    sec46 = sections.get("4_6", "")
    redo_rows = parse_sqlplus_table(sec46)
    result["redo_logs"] = []
    seen = set()
    for r in redo_rows:
        gn = r.get("GROUP#", "")
        fn = r.get("MEMBER", "")
        if fn and fn not in seen:
            seen.add(fn)
            size_bytes = r.get("BYTES", r.get("SIZE", "0"))
            try:
                size_mb = str(int(size_bytes) // (1024 * 1024))
            except Exception:
                size_mb = size_bytes
            result["redo_logs"].append({"group_no": gn, "member": fn, "size_mb": size_mb})

    sec47 = sections.get("4_7", "")
    ctrl_lines = [
        ln.strip()
        for ln in sec47.splitlines()
        if ln.strip() and not re.match(r"^[-\s]+$|^>O<|^\^o\^|\d+ rows? selected", ln.strip())
    ]
    result["control_files"] = [{"file_name": ln} for ln in ctrl_lines if "/" in ln or "\\" in ln]

    sec51 = sections.get("5_1", "")
    result["perf_stats"] = {}
    result["perf_review"] = {
        "library_cache_namespace": [],
        "pin_reload": {},
        "shared_pool_stats": {},
        "undo_segments": [],
    }

    # Split section 5_1 by ^o^...^o^ markers into named sub-sections
    perf_subsections = {}
    perf_parts = re.split(r"\^o\^[-]+([^-\^]+?)[-]+\^o\^", sec51)
    for i in range(1, len(perf_parts) - 1, 2):
        sub_name = perf_parts[i].strip()
        sub_body = perf_parts[i + 1] if i + 1 < len(perf_parts) else ""
        perf_subsections[sub_name] = sub_body

    # Extract single numeric value from a sub-section body
    def _extract_single_value(body):
        for ln in body.splitlines():
            stripped = ln.strip()
            if stripped and re.match(r"^[\d.]+$", stripped):
                return stripped
        return ""

    metrics_map = {
        "DictionaryCache HitRatio": "DICT_CACHE_HITRATIO",
        "Shared pool statistics": "SHARED_POOL_STATS",
        "Library Cache Hit Ratio": "LIBRARY_CACHE_PIN_RELOAD",
        "Redolog space request": "REDO_LOG_SPACE_REQUEST",
        "DB Block Buffer Cache Hit Ratio": "BUFFER_CACHE_HIT_RATIO",
        "Latch Hit Ratio": "LATCH_HIT_RATIO",
        "Disk Sort Ratio": "DISK_SORT_RATIO",
        "Rollback Segment Waits": "ROLLBACK_SEGMENT_WAITS",
        "Dispatcher Workload": "DISPATCHER_WORKLOAD",
        "PGA Cache Hit Percentage": "PGA_CACHE_HIT_PCT",
    }

    for sub_name, body in perf_subsections.items():
        for kw, metric_name in metrics_map.items():
            if kw.lower() in sub_name.lower():
                val = _extract_single_value(body)
                if val and metric_name not in result["perf_stats"]:
                    result["perf_stats"][metric_name] = val
                break

    # Library Cache HitRatio — parse the table and average GETHITRATIO
    for sub_name, body in perf_subsections.items():
        if "librarycache hitratio" in sub_name.lower():
            rows = parse_sqlplus_table(body)
            if rows:
                namespace_rows = []
                total = 0.0
                count = 0
                for r in rows:
                    namespace = r.get("NAMESPACE", "").strip()
                    val = r.get("GETHITRATIO", "").strip()
                    if namespace:
                        namespace_rows.append({"namespace": namespace, "gethitratio": val})
                    try:
                        total += float(val)
                        count += 1
                    except ValueError:
                        pass
                result["perf_review"]["library_cache_namespace"] = namespace_rows
                if count > 0:
                    avg = total / count
                    result["perf_stats"]["LIBRARY_CACHE_HITRATIO"] = str(round(avg, 6))
            break

    for sub_name, body in perf_subsections.items():
        lower_name = sub_name.lower()
        rows = parse_sqlplus_table(body)

        if "pin/reload ratio" in lower_name and rows:
            row = rows[0]
            result["perf_review"]["pin_reload"] = {
                "executions": row.get("Executions", row.get("EXECUTIONS", "")),
                "cache_misses": row.get("Cache Misses", row.get("CACHE MISSES", "")),
                "ratio": row.get("SUM(RELOADS)/SUM(PINS)", row.get("SUM(RELOADS)/SUM(PINS)+SUM(RELOADS)", "")),
            }
        elif "shared pool statistics" in lower_name and rows:
            row = rows[0]
            result["perf_review"]["shared_pool_stats"] = {
                "free_space": row.get("FREE_SPACE", ""),
                "avg_free_size": row.get("AVG_FREE_SIZE", ""),
                "max_free_size": row.get("MAX_FREE_SIZE", ""),
                "used_space": row.get("USED_SPACE", ""),
                "avg_used_size": row.get("AVG_USED_SIZE", ""),
            }
        elif "no. and size of undo segments" in lower_name and rows:
            undo_rows = []
            for row in rows:
                amount = row.get("COUNT(*)", "").strip()
                segment_type = row.get("SEGMENT_TYPE", "").strip()
                size_mb = row.get("MB", "").strip()
                if amount or segment_type or size_mb:
                    undo_rows.append(
                        {
                            "amount": amount,
                            "segment_type": segment_type,
                            "size_mb": size_mb,
                        }
                    )
            result["perf_review"]["undo_segments"] = undo_rows

    sec61 = sections.get("6_1", "")
    ts_free_rows = parse_sqlplus_table(sec61)
    result["tablespace_free"] = {}
    for r in ts_free_rows:
        name = r.get("TABLESPACE_NAME", "").strip()
        if name:
            try:
                pct_free = float(r.get("PCT_FREE", "0") or "0")
            except ValueError:
                pct_free = 0.0
            result["tablespace_free"][name] = {
                "pct_free": pct_free,
                "max_blocks": r.get("MAX_BLOCKS", "0"),
                "sum_free_blocks": r.get("SUM_FREE_BLOCKS", "0"),
            }

    sec71 = sections.get("7_1", "")
    user_rows = parse_sqlplus_table(sec71)
    result["user_tablespaces"] = []
    for r in user_rows:
        un = r.get("USERNAME", "").strip()
        dt = r.get("DEFAULT_TABLESPACE", "").strip()
        tt = r.get("TEMPORARY_TABLESPACE", "").strip()
        if un:
            result["user_tablespaces"].append({"user_name": un, "default_tbs": dt, "temp_tbs": tt})

    return result


def parse_invalid_objects(auto_collection_dir: Path) -> list[dict]:
    path = auto_collection_dir / "invalid_obj.log"
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8", errors="replace")
    rows = parse_sqlplus_table(text)
    result = []
    for r in rows:
        owner = r.get("OWNER", "").strip()
        obj = r.get("OBJECT", r.get("OBJECT_NAME", "")).strip()
        otype = r.get("TYPE", r.get("OBJECT_TYPE", "")).strip()
        status = r.get("STATUS", "INVALID").strip()
        if owner:
            result.append({"creator": owner, "object_name": obj, "object_type": otype, "status": status})
    return result


def parse_registry(auto_collection_dir: Path) -> list[dict]:
    path = auto_collection_dir / "registry.log"
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = [ln.rstrip() for ln in text.splitlines()]

    sep_idx = None
    for i, ln in enumerate(lines):
        if re.match(r"^-{10,}", ln.strip()):
            sep_idx = i
            break
    if sep_idx is None:
        return []

    sep_line = lines[sep_idx]
    col_spans = [(m.start(), m.end()) for m in re.finditer(r"-+", sep_line)]
    if not col_spans:
        return []

    header_line = ""
    for i in range(sep_idx - 1, -1, -1):
        if lines[i].strip():
            header_line = lines[i]
            break

    headers = [header_line[s:e].strip() for s, e in col_spans]

    results = []
    for line in lines[sep_idx + 1 :]:
        if not line.strip():
            continue
        if re.match(r"^[A-Z_$]+\.\w+\s*$", line.strip()):
            continue
        if re.match(r"^[A-Z_,]+$", line.strip()):
            continue

        row = {}
        for i, (s, e) in enumerate(col_spans):
            if i < len(headers):
                val = line[s:e].strip() if len(line) >= e else line[s:].strip()
                row[headers[i]] = val

        comp_id = row.get("COMP_ID", "").strip()
        version = row.get("VERSION", "").strip()
        status = row.get("STATUS", "").strip()
        valid_statuses = {"VALID", "INVALID", "REMOVED", "OPTION OFF", "LOADING", "LOADED"}
        if comp_id and version and status.upper() in valid_statuses and len(comp_id) < 40:
            results.append(
                {
                    "comp_id": comp_id,
                    "version": version,
                    "status": status,
                    "last_modified": row.get("MODIFIED", "").strip(),
                }
            )

    return results


def parse_data_growth_logs(fa_root: Path) -> list[dict]:
    growth_dir = fa_root / "data_growth_log"
    if not growth_dir.exists():
        return []

    monthly = []
    for txt_file in sorted(growth_dir.glob("*_segment_size.txt")):
        date_m = re.search(r"_(\d{8})_", txt_file.name)
        if not date_m:
            continue
        date_str = date_m.group(1)
        try:
            dt = datetime.strptime(date_str, "%Y%m%d")
            month_label = dt.strftime("%b-%Y")
        except ValueError:
            dt = datetime.now()
            month_label = date_str

        text = txt_file.read_text(encoding="utf-8", errors="replace")

        total_used_mb = _parse_growth_total_mb(
            text,
            marker="segment all",
            header="ALL SEGMENTS MB",
        )
        if total_used_mb == 0.0:
            rows = parse_sqlplus_table(text)
            for r in rows:
                seg_mb = r.get("SEGMENTS MB", r.get("MB", "0")) or "0"
                try:
                    total_used_mb += float(str(seg_mb).replace(",", ""))
                except ValueError:
                    pass

        total_allocated_mb = _parse_growth_total_mb(
            text,
            marker="size data file all",
            header="ALL DATA FILE MB",
        )

        total_used_gb = total_used_mb / 1024
        total_allocated_gb = total_allocated_mb / 1024
        pct_used = (total_used_gb / total_allocated_gb * 100) if total_allocated_gb else 0.0

        monthly.append(
            {
                "month": month_label,
                "date": dt,
                "used_gb": round(total_used_gb, 2),
                "allocated_gb": round(total_allocated_gb, 2),
                "pct_used": round(pct_used, 2),
            }
        )

    return sorted(monthly, key=lambda x: x["date"])


def _parse_growth_total_mb(text: str, marker: str, header: str) -> float:
    pattern = rf"###{re.escape(marker)}.*?{re.escape(header)}\s*-+\s*([\d,.]+)"
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if not match:
        return 0.0
    try:
        return float(match.group(1).replace(",", ""))
    except ValueError:
        return 0.0


def _enrich_growth_with_allocated(growth_rows: list, tbs_size_mb: dict, datafiles: list = None) -> list:
    """Calculate allocated_gb from datafiles (most accurate) or fall back to tbs_size_mb."""
    total_alloc_mb = 0.0

    # Prefer datafiles (direct MB column) as they are most accurate
    if datafiles:
        for f in datafiles:
            sv = str(f.get("size_mb", "0") or "0").replace(",", "")
            try:
                total_alloc_mb += float(sv)
            except ValueError:
                pass

    # Fallback: use tbs_size_mb if datafiles gave nothing
    if total_alloc_mb == 0.0 and tbs_size_mb:
        for v in tbs_size_mb.values():
            sv = str(v).replace(",", "")
            try:
                total_alloc_mb += float(sv)
            except ValueError:
                pass

    total_alloc_gb = total_alloc_mb / 1024

    for row in growth_rows:
        row_alloc_gb = float(row.get("allocated_gb", 0) or 0)
        if row_alloc_gb == 0.0 and total_alloc_gb > 0:
            row_alloc_gb = total_alloc_gb
            row["allocated_gb"] = round(total_alloc_gb, 2)

        if row_alloc_gb > 0:
            row["pct_used"] = round(row["used_gb"] / row_alloc_gb * 100, 2)
        else:
            row["pct_used"] = 0.0

    return growth_rows


def parse_oracle_stats(auto_collection_dir: Path) -> dict:
    path = auto_collection_dir / "mfec_oracle_stats.log"
    if not path.exists():
        return {}

    text = path.read_text(encoding="utf-8", errors="replace")
    stats = {}

    pga_m = re.search(r"aggregate PGA target parameter\s+([\d,]+)\s+MBytes", text, re.IGNORECASE)
    if pga_m:
        stats["PGA_TARGET_MB"] = pga_m.group(1).replace(",", "")

    ch_m = re.search(r"cache hit percentage\s+([\d.]+)\s+percent", text, re.IGNORECASE)
    if ch_m:
        stats["PGA_CACHE_HIT_PCT"] = ch_m.group(1)

    return stats


def _parse_statspack_pga_advisory(text: str) -> list[dict]:
    rows = []
    marker = "PGA Memory Advisory"
    idx = text.find(marker)
    if idx == -1:
        return rows

    section = text[idx:]
    for line in section.splitlines():
        if rows and "-------------------------------------------------------------" in line:
            break
        parts = line.split()
        if len(parts) >= 7 and parts[0].replace(",", "").isdigit():
            try:
                size_mb = float(parts[0].replace(",", ""))
                factor = float(parts[1])
                estd_extra_mb = float(parts[3].replace(",", ""))
                cache_hit = float(parts[5])
                rows.append(
                    {
                        "pga_size_bytes": size_mb * 1024 * 1024,
                        "pga_size_gb": round(size_mb / 1024, 2),
                        "pga_target_factor": factor,
                        "is_current": abs(factor - 1.0) < 0.0001,
                        "estd_extra_bytes_rw": estd_extra_mb * 1024 * 1024,
                        "estd_extra_mb_rw": round(estd_extra_mb),
                        "estd_pga_cache_hit_pct": cache_hit,
                    }
                )
            except ValueError:
                continue

    return rows


def _parse_statspack_sga_advisory(text: str) -> list[dict]:
    rows = []
    marker = "SGA Target Advisory"
    idx = text.find(marker)
    if idx == -1:
        return rows

    section = text[idx:]
    for line in section.splitlines():
        if rows and "-------------------------------------------------------------" in line:
            break
        parts = line.split()
        if len(parts) >= 5 and parts[0].replace(",", "").isdigit():
            try:
                size_mb = float(parts[0].replace(",", ""))
                factor = float(parts[1])
                phys_reads = float(parts[-1].replace(",", ""))
                rows.append(
                    {
                        "sga_size_mb": size_mb,
                        "sga_size_gb": round(size_mb / 1024, 2),
                        "sga_size_factor": factor,
                        "is_current": abs(factor - 1.0) < 0.0001,
                        "estd_physical_reads": phys_reads,
                    }
                )
            except ValueError:
                continue

    return rows


def parse_patch_info(log_dir: Path) -> list[dict]:
    candidates = [
        log_dir / "db_lsinventory.txt",
        log_dir / "DB_OPatch.txt",
        log_dir / "DB_Opatch.txt",
    ]
    path = next((p for p in candidates if p.exists()), None)
    if not path:
        return []

    text = path.read_text(encoding="utf-8", errors="replace")
    current_ver = "Unknown"

    ru_match = re.search(r"Database Release Update\s*:\s*([\d.]+)", text, re.IGNORECASE)
    if ru_match:
        current_ver = ru_match.group(1).strip()
    else:
        product_match = re.search(r"Oracle Database .*?\s(\d+\.\d+\.\d+\.\d+\.\d+)", text, re.IGNORECASE)
        if product_match:
            current_ver = product_match.group(1).strip()
        else:
            opatch_match = re.search(r"OPatch version\s*:\s*([\d.]+)", text, re.IGNORECASE)
            if opatch_match:
                current_ver = opatch_match.group(1).strip()

    return [
        {
            "component": "Oracle Home",
            "current_version": current_ver,
            "recommended_version": "19.29.0.0.0 (OCT 2025)",
            "suggestion": "Recommend applying the latest patch to fix bugs and update database security."
            if current_ver != "19.29.0.0.0"
            else "Patch is up to date.",
        }
    ]


# ================================================================
# Section 2: OS information parser
# ================================================================


def find_os_dir(extracted_root: Path, fa_root: Path) -> Optional[Path]:
    """Find the os/oracle/ directory in the extracted ZIP.
    The os/ folder is typically a sibling or cousin of the fast_assessment root,
    NOT inside it. We search by walking up from fa_root."""
    direct_oracle = fa_root / "os" / "oracle"
    if direct_oracle.is_dir():
        return direct_oracle
    direct_root = fa_root / "os" / "root"
    if direct_root.is_dir():
        return direct_root
    direct_os = fa_root / "os"
    if direct_os.is_dir():
        return direct_os

    # Strategy 1: Check sibling/parent directories walking up from fa_root
    current = fa_root
    for _ in range(5):
        parent = current.parent
        if parent == current:
            break
        os_oracle = parent / "os" / "oracle"
        if os_oracle.is_dir():
            return os_oracle
        os_root = parent / "os" / "root"
        if os_root.is_dir():
            return parent / "os" / "root"
        os_dir = parent / "os"
        if os_dir.is_dir():
            return os_dir
        current = parent

    # Strategy 2: Search entire extracted tree
    for p in extracted_root.rglob("os"):
        if p.is_dir():
            oracle_sub = p / "oracle"
            if oracle_sub.is_dir():
                return oracle_sub
            return p

    return None


def _safe_read(path: Path) -> str:
    """Read a text file safely, returning empty string if not found."""
    if not path or not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace").strip()
    except Exception:
        return ""


def parse_os_info(os_dir: Optional[Path]) -> dict:
    """Parse OS information files collected by pm_linux_tool.txt."""
    info = {
        "hostname": "",
        "ip_address": "",
        "subnet_mask": "",
        "os_version": "",
        "kernel_version": "",
        "architecture": "",
        "oracle_owner": "",
        "oracle_home": "",
        "oracle_sid": "",
        "oracle_base": "",
        "shell": "",
        "oracle_groups": "",
        "cpu_model": "",
        "cpu_count": "",
        "cores_per_socket": "",
        "sockets": "",
        "cpu_mhz": "",
        "cache_size": "",
        "total_ram_mb": "",
        "total_ram_gb": "",
        "swap_total_mb": "",
        "swap_used_mb": "",
        "tmp_size": "",
        "tmp_used": "",
        "tmp_avail": "",
        "jdk_version": "",
        "kernel_params": {},
        "semaphores": "",
        "disk_info": [],
        "user_limits": {},
    }

    if not os_dir or not os_dir.is_dir():
        return info

    # --- Hostname ---
    info["hostname"] = _safe_read(os_dir / "hostname.txt").split("\n")[0].strip()

    # --- OS info (uname -a) ---
    uname_text = _safe_read(os_dir / "os_info1.txt")
    if uname_text:
        parts = uname_text.split()
        if len(parts) >= 3:
            info["kernel_version"] = parts[2]  # e.g. 3.10.0-1160.42.2.el7.x86_64
        if "x86_64" in uname_text:
            info["architecture"] = "x86_64"
        elif "aarch64" in uname_text:
            info["architecture"] = "aarch64"
        elif "sparc" in uname_text.lower():
            info["architecture"] = "SPARC"
        else:
            # fallback
            info["architecture"] = parts[-2] if len(parts) >= 2 else ""

    # --- OS Release ---
    release_text = _safe_read(os_dir / "os_release.txt")
    if release_text:
        # Prefer explicit "release X.Y" line (e.g. "Red Hat Enterprise Linux Server release 7.9 (Maipo)")
        release_line = ""
        for ln in release_text.splitlines():
            stripped = ln.strip()
            if re.match(r"^[A-Z].*release\s+[\d.]+", stripped, re.IGNORECASE) and not stripped.startswith("#"):
                release_line = stripped
                break
        if release_line:
            info["os_version"] = release_line
        else:
            # Fallback to PRETTY_NAME
            m = re.search(r'PRETTY_NAME="([^"]+)"', release_text)
            if m:
                info["os_version"] = m.group(1)

    # --- IP address (ifconfig) ---
    ip_text = _safe_read(os_dir / "IP.txt")
    if ip_text:
        # Find first non-loopback inet address
        for block in re.split(r"\n(?=\S)", ip_text):
            if "LOOPBACK" in block or block.strip().startswith("lo:"):
                continue
            if "virbr" in block:
                continue
            inet_m = re.search(r"inet\s+([\d.]+)", block)
            mask_m = re.search(r"netmask\s+([\d.]+)", block)
            if inet_m:
                info["ip_address"] = inet_m.group(1)
                if mask_m:
                    info["subnet_mask"] = mask_m.group(1)
                break

    # --- lscpu ---
    lscpu_text = _safe_read(os_dir / "lscpu.txt")
    if lscpu_text:
        lscpu_dict = {}
        for ln in lscpu_text.splitlines():
            if ":" in ln:
                k, v = ln.split(":", 1)
                lscpu_dict[k.strip()] = v.strip()
        info["cpu_model"] = lscpu_dict.get("Model name", "")
        info["cpu_count"] = lscpu_dict.get("CPU(s)", "")
        info["cores_per_socket"] = lscpu_dict.get("Core(s) per socket", "")
        info["sockets"] = lscpu_dict.get("Socket(s)", "")
        info["cpu_mhz"] = lscpu_dict.get("CPU MHz", "")
        # Cache: prefer L3, fallback L2
        info["cache_size"] = lscpu_dict.get("L3 cache", lscpu_dict.get("L2 cache", ""))

    # --- Memory (free -m) ---
    swap_text = _safe_read(os_dir / "swap.txt")
    if swap_text:
        for ln in swap_text.splitlines():
            if ln.strip().startswith("Mem:"):
                parts = ln.split()
                if len(parts) >= 2:
                    info["total_ram_mb"] = parts[1]
                    try:
                        info["total_ram_gb"] = str(round(float(parts[1]) / 1024, 2))
                    except ValueError:
                        pass
            elif ln.strip().startswith("Swap:"):
                parts = ln.split()
                if len(parts) >= 2:
                    info["swap_total_mb"] = parts[1]
                if len(parts) >= 3:
                    info["swap_used_mb"] = parts[2]

    # --- /tmp space ---
    tmp_text = _safe_read(os_dir / "tmp_space.txt")
    if tmp_text:
        for ln in tmp_text.splitlines():
            if "/" in ln and not ln.startswith("Filesystem"):
                parts = ln.split()
                if len(parts) >= 4:
                    info["tmp_size"] = parts[1]
                    info["tmp_used"] = parts[2]
                    info["tmp_avail"] = parts[3]
                break

    # --- JDK version ---
    jdk_text = _safe_read(os_dir / "java_version.txt")
    if jdk_text:
        first_line = jdk_text.splitlines()[0].strip()
        info["jdk_version"] = first_line

    # --- Kernel parameters (/etc/sysctl.conf) ---
    kernel_text = _safe_read(os_dir / "kernel.txt")
    if kernel_text:
        for ln in kernel_text.splitlines():
            ln = ln.strip()
            if ln and not ln.startswith("#") and "=" in ln:
                k, v = ln.split("=", 1)
                info["kernel_params"][k.strip()] = v.strip()

    # --- Semaphores ---
    sem_text = _safe_read(os_dir / "semaphore1.txt")
    if sem_text:
        info["semaphores"] = sem_text.splitlines()[0].strip()

    # --- User environment ---
    env_text = _safe_read(os_dir / "user_environments.txt")
    if env_text:
        env_dict = {}
        for ln in env_text.splitlines():
            if "=" in ln:
                k, v = ln.split("=", 1)
                env_dict[k.strip()] = v.strip()
        info["oracle_home"] = env_dict.get("ORACLE_HOME", "")
        info["oracle_sid"] = env_dict.get("ORACLE_SID", "")
        info["oracle_base"] = env_dict.get("ORACLE_BASE", "")
        info["shell"] = env_dict.get("SHELL", "")
        info["oracle_owner"] = env_dict.get("USER", "")

    # --- User id (groups) ---
    uid_text = _safe_read(os_dir / "user_id.txt")
    if uid_text:
        groups_m = re.search(r"groups=(.+)", uid_text)
        if groups_m:
            info["oracle_groups"] = groups_m.group(1).strip()

    # --- Disk info (df -h) ---
    disk_text = _safe_read(os_dir / "disk.txt")
    if disk_text:
        for ln in disk_text.splitlines():
            if ln.startswith("Filesystem") or not ln.strip():
                continue
            parts = ln.split()
            if len(parts) >= 6:
                info["disk_info"].append({
                    "filesystem": parts[0],
                    "size": parts[1],
                    "used": parts[2],
                    "avail": parts[3],
                    "use_pct": parts[4],
                    "mounted_on": parts[5],
                })

    # --- User limits (ulimit) ---
    ulimit_text = _safe_read(os_dir / "ulimit.txt")
    if ulimit_text:
        for ln in ulimit_text.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            m = re.match(r"^(.+?)\s{2,}\(.*?\)\s+(.+)$", ln)
            if m:
                info["user_limits"][m.group(1).strip()] = m.group(2).strip()

    # --- Crontab ---
    crontab_text = _safe_read(os_dir / "crontab.txt")
    info["crontab_entries"] = []
    if crontab_text:
        for ln in crontab_text.splitlines():
            stripped = ln.strip()
            if not stripped:
                continue
            # Section headers like "######## Monitor ########"
            if re.match(r"^#+\s*.*\s*#+$", stripped):
                continue
            # Commented-out entries
            is_active = not stripped.startswith("#")
            # Remove leading # or ## for commented entries
            clean = re.sub(r"^#+\s*", "", stripped)
            # Parse cron schedule + command
            cron_m = re.match(r"^(\S+\s+\S+\s+\S+\s+\S+\s+\S+)\s+(.+)$", clean)
            if cron_m:
                info["crontab_entries"].append({
                    "schedule": cron_m.group(1),
                    "command": cron_m.group(2).strip(),
                    "active": is_active,
                })

    # --- Hosts file ---
    hosts_text = _safe_read(os_dir / "hosts.txt")
    info["hosts_entries"] = []
    if hosts_text:
        for ln in hosts_text.splitlines():
            stripped = ln.strip()
            if not stripped or stripped.startswith("#"):
                continue
            parts = stripped.split()
            if len(parts) >= 2:
                info["hosts_entries"].append({
                    "ip": parts[0],
                    "hostnames": " ".join(parts[1:]),
                })

    return info


def parse_section53_data(auto_coll_dir: Path) -> dict:
    sec53 = {
        "top_sql_disk": [],
        "top_sql_mem": [],
        "open_cursors": {"limit": "", "usage": "", "pct": ""},
        "memory_pools": {"shared_free_pct": "", "java_free_pct": "", "large_free_pct": ""},
        "pga_advisory": [],
        "sga_advisory": [],
        "instance_efficiency": {},
        "top5_events": []
    }
    
    if not auto_coll_dir or not auto_coll_dir.exists():
        return sec53
        
    report_dir = auto_coll_dir.parent / "report"
    if report_dir.exists():
        lst_files = list(report_dir.glob("sp_*.lst"))
        if lst_files:
            advisory_files = sorted([p for p in lst_files if "_top1" in p.name.lower()])
            advisory_file = advisory_files[0] if advisory_files else sorted(lst_files)[-1]
            advisory_content = _safe_read(advisory_file)
            if advisory_content:
                sec53["pga_advisory"] = _parse_statspack_pga_advisory(advisory_content)
                sec53["sga_advisory"] = _parse_statspack_sga_advisory(advisory_content)

            # Just read the first statspack file to get representative data
            content = _safe_read(lst_files[0])
            if content:
                lines = content.splitlines()
                in_inst_eff = False
                in_top5 = False
                for i, ln in enumerate(lines):
                    if "Instance Efficiency Indicators" in ln:
                        in_inst_eff = True
                        continue
                    if in_inst_eff and ln.strip() == "":
                        # Might be multiple blank lines, stop after a few or at next section
                        if i+1 < len(lines) and "Wait Events" in lines[i+1]:
                            in_inst_eff = False
                    if in_inst_eff and ":" in ln and "%" in ln:
                        # Use regex to capture patterns like "Key Name %:  99.55"
                        kv_pairs = re.findall(r'([A-Za-z][A-Za-z %>\d/]*?%)\s*:\s*([\d.]+)', ln)
                        for k, v in kv_pairs:
                            sec53["instance_efficiency"][k.strip()] = v.strip()
                    
                    if "Top 5 Timed Events" in ln or "Top 5 Timed Foreground Events" in ln:
                        in_top5 = True
                        continue
                    if in_top5:
                        if "------------------" in ln or ln.strip() == "" or "Event" in ln:
                            if ln.strip() == "" and len(sec53["top5_events"]) > 0:
                                in_top5 = False
                            continue
                        parts = ln.strip().split()
                        if len(parts) >= 4 and not ln.startswith("-"):
                            # Event name can have spaces, but wait time is usually numeric
                            # E.g.: "CPU time                 72,555             60.2"
                            # Let's just grab the event name which is everything before the first number
                            event_name = []
                            for part in parts:
                                if part.replace(",", "").replace(".", "").isdigit():
                                    break
                                event_name.append(part)
                            if event_name:
                                sec53["top5_events"].append(" ".join(event_name))
                        if len(sec53["top5_events"]) >= 5:
                            in_top5 = False

    # --- Top SQL by Disk ---
    sql_disk_txt = _safe_read(auto_coll_dir / "top25sql_use_disk.txt")
    if sql_disk_txt:
        # Extract the first few lines of the table
        lines = sql_disk_txt.splitlines()
        headers_found = False
        for ln in lines:
            if "sql_id" in ln.lower() and "text" in ln.lower():
                headers_found = True
                continue
            if headers_found and not ln.startswith("-") and ln.strip():
                parts = ln.split()
                if len(parts) >= 3 and parts[0].isdigit():
                    sec53["top_sql_disk"].append({
                        "reads": parts[0],
                        "sql_id": parts[1],
                        "text": " ".join(parts[2:])[:50] + "..."
                    })
                if len(sec53["top_sql_disk"]) >= 5:
                    break

    # --- Top SQL by Mem ---
    sql_mem_txt = _safe_read(auto_coll_dir / "top25sql_use_mem.txt")
    if sql_mem_txt:
        lines = sql_mem_txt.splitlines()
        headers_found = False
        for ln in lines:
            if "sql_id" in ln.lower() and "text" in ln.lower():
                headers_found = True
                continue
            if headers_found and not ln.startswith("-") and ln.strip():
                parts = ln.split()
                if len(parts) >= 3 and parts[0].isdigit():
                    sec53["top_sql_mem"].append({
                        "buffer_gets": parts[0],
                        "sql_id": parts[1],
                        "text": " ".join(parts[2:])[:50] + "..."
                    })
                if len(sec53["top_sql_mem"]) >= 5:
                    break

    # --- Open Cursors (from mfec_oracle_stats.log) ---
    stats_log = _safe_read(auto_coll_dir / "mfec_oracle_stats.log")
    if stats_log:
        lines = stats_log.splitlines()
        
        # 1. Open Cursors limit & usage
        for i, ln in enumerate(lines):
            if "open_cursors" in ln.lower() and "%" in ln:
                parts = ln.strip().split()
                if len(parts) >= 3:
                    sec53["open_cursors"]["limit"] = parts[1]
                    sec53["open_cursors"]["pct"] = parts[-1]
                    # Estimate usage
                    try:
                        limit = float(parts[1])
                        pct = float(parts[-1].replace("%", ""))
                        sec53["open_cursors"]["usage"] = str(int((limit * pct) / 100))
                    except ValueError:
                        pass
                break
                
        # 2. Memory Pool Free %
        for i, ln in enumerate(lines):
            if "Shared Pool Free %" in ln:
                for j in range(i+1, i+5):
                    if lines[j].strip().replace(".", "").isdigit():
                        sec53["memory_pools"]["shared_free_pct"] = lines[j].strip()
                        break
            elif "Java Pool Free %" in ln:
                for j in range(i+1, i+5):
                    if lines[j].strip().replace(".", "").isdigit():
                        sec53["memory_pools"]["java_free_pct"] = lines[j].strip()
                        break
            elif "Large Pool Free %" in ln:
                for j in range(i+1, i+5):
                    if lines[j].strip().replace(".", "").isdigit():
                        sec53["memory_pools"]["large_free_pct"] = lines[j].strip()
                        break
                        
        # 3. SGA Advisory fallback from mfec_oracle_stats.log
        if not sec53["sga_advisory"]:
            in_sga = False
            for ln in lines:
                if "SGA_SIZE" in ln and "ESTD_PHYSICAL_READS" in ln:
                    in_sga = True
                    continue
                if in_sga:
                    if "###" in ln or "Date Run" in ln or "PGA_TARGET" in ln:
                        in_sga = False
                        continue
                    if not ln.strip() or "---" in ln or "CON_ID" in ln:
                        continue
                    parts = ln.split()
                    if len(parts) >= 6 and parts[0].isdigit():
                        try:
                            size_mb = float(parts[1])
                            factor = float(parts[2])
                            phys_reads = float(parts[5])
                            sec53["sga_advisory"].append({
                                "sga_size_mb": size_mb,
                                "sga_size_gb": round(size_mb / 1024, 2),
                                "sga_size_factor": factor,
                                "is_current": abs(factor - 1.0) < 0.0001,
                                "estd_physical_reads": phys_reads
                            })
                        except ValueError:
                            pass

        # 4. PGA Advisory fallback from mfec_oracle_stats.log
        if not sec53["pga_advisory"]:
            in_pga = False
            for ln in lines:
                if "PGA_TARGET_FOR_ESTIMATE" in ln and "ESTD_PGA_CACHE_HIT_PERCENTAGE" in ln:
                    in_pga = True
                    continue
                if in_pga:
                    if "###" in ln or "Date Run" in ln or "SGA_SIZE" in ln:
                        in_pga = False
                        continue
                    if not ln.strip() or "---" in ln:
                        continue
                    parts = ln.split()
                    if len(parts) >= 5 and parts[0].isdigit():
                        try:
                            size_bytes = float(parts[1])
                            factor = float(parts[2])
                            estd_extra_bytes_rw = float(parts[3])
                            cache_hit = float(parts[4])
                            sec53["pga_advisory"].append({
                                "pga_size_bytes": size_bytes,
                                "pga_size_gb": round(size_bytes / (1024**3), 2),
                                "pga_target_factor": factor,
                                "is_current": abs(factor - 1.0) < 0.0001,
                                "estd_extra_bytes_rw": estd_extra_bytes_rw,
                                "estd_extra_mb_rw": round(estd_extra_bytes_rw / (1024 * 1024)),
                                "estd_pga_cache_hit_pct": cache_hit
                            })
                        except ValueError:
                            pass

    return sec53

def _format_oracle_log_date(raw_date: str) -> str:
    for fmt in ("%a %b %d %H:%M:%S %Y", "%b %d %H:%M:%S %Y"):
        try:
            return datetime.strptime(raw_date.strip(), fmt).strftime("%d-%b-%Y")
        except ValueError:
            continue
    return raw_date.strip()


def _extract_collect_date(fa_root: Path) -> str:
    match = re.search(r"(20\d{6})", str(fa_root))
    if not match:
        return ""
    try:
        return datetime.strptime(match.group(1), "%Y%m%d").strftime("%d-%b-%Y")
    except ValueError:
        return ""


def _summarize_datapump_log(content: str, max_exported_rows: int = 7) -> str:
    lines = [line.rstrip() for line in content.splitlines()]
    output = []
    exported_rows = []
    final_rows = []
    seen_processing = 0

    for line in lines:
        stripped = line.strip()
        if not stripped:
            if output and output[-1] != "":
                output.append("")
            continue

        lower = stripped.lower()
        if lower.startswith(". . exported"):
            exported_rows.append(line)
            continue

        if (
            lower.startswith("master table")
            or lower.startswith("dump file set")
            or lower.startswith("job ")
            or "successfully completed" in lower
            or stripped.startswith("*" * 20)
        ):
            final_rows.append(line)
            continue

        if lower.startswith("processing object type"):
            seen_processing += 1
            if seen_processing <= 8:
                output.append(line)
            continue

        if len(output) < 22:
            output.append(line)

    while output and output[-1] == "":
        output.pop()

    if exported_rows:
        output.extend(["", "..."])
        output.extend(exported_rows[:max_exported_rows])
        if len(exported_rows) > max_exported_rows:
            output.append("...")

    if final_rows:
        output.extend([""] if output and output[-1] != "" else [])
        output.extend(final_rows[-6:])

    return "\n".join(output).strip()


def _score_backup_file(path: Path, content: str) -> int:
    name = path.name.lower()
    lower = content.lower()
    score = 0

    if "expdp" in name or "datapump" in name:
        score += 80
    if "export" in name or "dump" in name:
        score += 25
    if "export: release" in lower:
        score += 80
    if "sys_export_schema" in lower:
        score += 70
    if "dumpfile=" in lower or "schemas=" in lower:
        score += 40
    if ". . exported" in lower:
        score += 35
    if "rman" in name or "recovery manager" in lower:
        score -= 80
    if "list of backup sets" in lower or "list of archived logs" in lower:
        score -= 50

    return score


def _build_datapump_schedule_summary(fa_root: Path) -> dict:
    crontab_path = None
    for current in [fa_root, *fa_root.parents[:5]]:
        candidate = current / "os" / "oracle" / "crontab.txt"
        if candidate.exists():
            crontab_path = candidate
            break

    if crontab_path is None:
        return {}

    content = _safe_read(crontab_path)
    if not content:
        return {}

    lines = []
    for line in content.splitlines():
        stripped = line.strip()
        if stripped and "run_expdp" in stripped.lower():
            lines.append(stripped)

    if not lines:
        return {}

    summary = [
        "Datapump export job output was not included in this Fast Assessment ZIP.",
        "",
        "Configured Datapump backup schedule found in crontab:",
        *lines,
        "",
        f"Source: {crontab_path.name}",
    ]
    return {
        "found": True,
        "type": "datapump_schedule",
        "source_file": str(crontab_path),
        "title": "Datapump backup schedule found; export job log not included in ZIP.",
        "log_content": "\n".join(summary),
    }


def parse_backup_info(fa_root: Path) -> dict:
    """Parse APPENDIX E - Backup Configuration"""
    backup_data = {"found": False, "type": "", "source_file": "", "title": "", "log_content": ""}
    
    if not fa_root or not fa_root.exists():
        return backup_data
        
    possible_files = [
        p for p in fa_root.rglob("*")
        if p.is_file() and p.suffix.lower() in {".log", ".txt", ".lst", ".out"}
    ]

    candidates = []
    for path in possible_files:
        content = _safe_read(path)
        if not content or len(content.strip()) < 50:
            continue
        score = _score_backup_file(path, content)
        if score > 0:
            candidates.append((score, path, content))

    if candidates:
        score, path, content = max(candidates, key=lambda item: item[0])
        if score >= 100:
            export_date = ""
            date_match = re.search(r"Export:\s+Release.*?\bon\s+(.+)$", content, re.IGNORECASE | re.MULTILINE)
            if date_match:
                export_date = _format_oracle_log_date(date_match.group(1))

            collect_date = _extract_collect_date(fa_root)
            if export_date and collect_date:
                title = f"Last schema export (Datapump) backup on: {export_date} (Collect data on: {collect_date})"
            elif export_date:
                title = f"Last schema export (Datapump) backup on: {export_date}"
            else:
                title = "Last schema export (Datapump) backup log"

            return {
                "found": True,
                "type": "datapump_log",
                "source_file": str(path),
                "title": title,
                "log_content": _summarize_datapump_log(content),
            }

    schedule_summary = _build_datapump_schedule_summary(fa_root)
    if schedule_summary:
        return schedule_summary

    backup_dir = fa_root / "auto_collection" / "backup"
    rman_files = sorted(backup_dir.glob("*.log")) if backup_dir.exists() else []
    for path in rman_files:
        content = _safe_read(path)
        if content and len(content.strip()) > 50:
            lines = content.splitlines()
            log_content = "\n".join(lines[:80])
            if len(lines) > 80:
                log_content += "\n... (truncated for report) ..."
            return {
                "found": True,
                "type": "rman_log",
                "source_file": str(path),
                "title": "RMAN backup information found; Datapump export log not included in ZIP.",
                "log_content": log_content,
            }

    return backup_data

def parse_os_perf_data(fa_root: Path) -> dict:
    """Parse APPENDIX F - OS Performance Summary from sar logs"""
    perf_data = {
        "found": False,
        "cpu_trend": [],    # { day, usr, sys, iowait, idle }
        "mem_trend": [],    # { day, pswpin, pswpout, pgpgin, pgpgout }
        "disk_trend": [],   # { day, await, svctm, util }
        "cpu_samples": [],  # { day, time, usr, sys, iowait, util }
        "runq_samples": [], # { day, time, runq }
        "mem_free_samples": [], # { day, time, free_gb, inactive_gb }
        "swap_samples": [], # { day, time, used_gb, used_pct }
        "page_io_samples": [], # { day, time, pgpgin, pgpgout }
        "disk_samples": [], # { day, time, device, await, svctm, util }
    }
    
    if not fa_root or not fa_root.exists():
        return perf_data
        
    sar_dir = fa_root / "sar"
    if not sar_dir.exists():
        sar_dir = fa_root / "auto_collection" / "sar"
        if not sar_dir.exists():
            return perf_data
            
    sar_files = list(sar_dir.glob("sar[0-9][0-9]"))
    if not sar_files:
        return perf_data
        
    perf_data["found"] = True
    sar_files.sort(key=lambda x: x.name)
    
    for fpath in sar_files:
        day_str = fpath.name.replace("sar", "")
        try:
            content = _safe_read(fpath)
            if not content: continue
        except:
            continue
            
        lines = content.splitlines()
        day_cpu = {"usr": None, "sys": None, "iowait": None, "idle": None}
        day_mem = {"pswpin": None, "pswpout": None, "pgpgin": None, "pgpgout": None}
        day_disk = {"await": None, "svctm": None, "util": None}
        
        # New Metrics
        day_runq = {"max_runq_sz": None, "avg_runq_sz": None}
        day_free_mem = {"min_kbmemfree": None, "avg_kbmemfree": None}
        day_swap_used = {"max_kbswpused": None, "max_pctswpused": None}
        
        current_block = None
        for ln in lines:
            if not ln.strip(): continue
            if " CPU " in ln and "%usr" in ln: current_block = "CPU"; continue
            elif "runq-sz" in ln: current_block = "QUEUE"; continue
            elif "pswpin/s" in ln: current_block = "SWAP_IO"; continue
            elif "pgpgin/s" in ln: current_block = "PAGE_IO"; continue
            elif " DEV " in ln and "tps" in ln and "svctm" in ln: current_block = "DISK"; continue
            elif "kbmemfree" in ln: current_block = "MEM"; continue
            elif "kbswpfree" in ln and "kbswpused" in ln: current_block = "SWAP_USE"; continue
                
            if ln.startswith("Average:"):
                parts = ln.split()
                if len(parts) < 3: continue
                
                if current_block == "CPU" and parts[1] == "all":
                    try:
                        day_cpu["usr"] = float(parts[2])
                        day_cpu["sys"] = float(parts[4])
                        day_cpu["iowait"] = float(parts[5])
                        day_cpu["idle"] = float(parts[-1])
                    except: pass
                elif current_block == "QUEUE":
                    try:
                        day_runq["avg_runq_sz"] = float(parts[1])
                    except: pass
                elif current_block == "SWAP_IO" and parts[1] != "pswpin/s":
                    try:
                        day_mem["pswpin"] = float(parts[1])
                        day_mem["pswpout"] = float(parts[2])
                    except: pass
                elif current_block == "PAGE_IO" and parts[1] != "pgpgin/s":
                    try:
                        day_mem["pgpgin"] = float(parts[1])
                        day_mem["pgpgout"] = float(parts[2])
                    except: pass
                elif current_block == "MEM" and parts[1] != "kbmemfree":
                    try:
                        day_free_mem["avg_kbmemfree"] = float(parts[1])
                    except: pass
                elif current_block == "SWAP_USE" and parts[1] != "kbswpfree":
                    try:
                        day_swap_used["max_pctswpused"] = float(parts[3])
                    except: pass
                elif current_block == "DISK" and parts[1].startswith("dev"):
                    try:
                        util = float(parts[-1])
                        svctm = float(parts[-2])
                        await_val = float(parts[-3])
                        day_disk["util"] = max(day_disk["util"] or 0, util)
                        day_disk["svctm"] = max(day_disk["svctm"] or 0, svctm)
                        day_disk["await"] = max(day_disk["await"] or 0, await_val)
                    except: pass
            else:
                # Capture Maximums/Minimums during the day (not just Average)
                parts = ln.split()
                if len(parts) > 2 and parts[0] != "Linux" and parts[0] != "Average:":
                    has_ampm = len(parts) > 1 and parts[1] in {"AM", "PM"}
                    sample_time = f"{parts[0]} {parts[1]}" if has_ampm else parts[0]

                    if current_block == "CPU" and "CPU" not in ln:
                        cpu_idx = 2 if has_ampm else 1
                        if len(parts) > cpu_idx and parts[cpu_idx] == "all":
                            try:
                                usr = float(parts[cpu_idx + 1])
                                sys_pct = float(parts[cpu_idx + 3])
                                iowait = float(parts[cpu_idx + 4])
                                idle = float(parts[-1])
                                perf_data["cpu_samples"].append({
                                    "day": day_str,
                                    "time": sample_time,
                                    "usr": usr,
                                    "sys": sys_pct,
                                    "iowait": iowait,
                                    "idle": idle,
                                    "util": max(0.0, min(100.0, 100.0 - idle)),
                                })
                            except Exception:
                                pass

                    if current_block == "QUEUE" and "runq-sz" not in ln:
                        try:
                            val = float(parts[2] if has_ampm else parts[1])
                            day_runq["max_runq_sz"] = max(day_runq["max_runq_sz"] or 0, val)
                            perf_data["runq_samples"].append({
                                "day": day_str,
                                "time": sample_time,
                                "runq": val,
                            })
                        except: pass
                    elif current_block == "MEM" and "kbmemfree" not in ln:
                        try:
                            # Usually kbmemfree is the 2nd column after time
                            base_idx = 2 if has_ampm else 1
                            val = float(parts[base_idx])
                            if day_free_mem["min_kbmemfree"] is None:
                                day_free_mem["min_kbmemfree"] = val
                            else:
                                day_free_mem["min_kbmemfree"] = min(day_free_mem["min_kbmemfree"], val)
                            inactive_kb = float(parts[base_idx + 7]) if len(parts) > base_idx + 7 else 0
                            perf_data["mem_free_samples"].append({
                                "day": day_str,
                                "time": sample_time,
                                "free_gb": round(val / 1024 / 1024, 4),
                                "inactive_gb": round(inactive_kb / 1024 / 1024, 4),
                            })
                        except: pass
                    elif current_block == "SWAP_USE" and "kbswpfree" not in ln:
                        try:
                            # kbswpused is usually 3rd, %swpused is 4th
                            base_idx = 2 if has_ampm else 1
                            used_kb = float(parts[base_idx + 1])
                            idx_pct = base_idx + 2
                            val_pct = float(parts[idx_pct])
                            day_swap_used["max_pctswpused"] = max(day_swap_used["max_pctswpused"] or 0, val_pct)
                            perf_data["swap_samples"].append({
                                "day": day_str,
                                "time": sample_time,
                                "used_gb": round(used_kb / 1024 / 1024, 4),
                                "used_pct": val_pct,
                            })
                        except: pass
                    elif current_block == "PAGE_IO" and "pgpgin/s" not in ln:
                        try:
                            base_idx = 2 if has_ampm else 1
                            perf_data["page_io_samples"].append({
                                "day": day_str,
                                "time": sample_time,
                                "pgpgin": float(parts[base_idx]),
                                "pgpgout": float(parts[base_idx + 1]),
                            })
                        except: pass
                    elif current_block == "DISK" and "DEV" not in ln:
                        try:
                            base_idx = 2 if has_ampm else 1
                            device = parts[base_idx]
                            await_val = float(parts[base_idx + 6])
                            svctm = float(parts[base_idx + 7])
                            util = float(parts[base_idx + 8])
                            perf_data["disk_samples"].append({
                                "day": day_str,
                                "time": sample_time,
                                "device": device,
                                "await": await_val,
                                "svctm": svctm,
                                "util": util,
                            })
                        except: pass

        if day_cpu["usr"] is not None:
            perf_data["cpu_trend"].append({"day": day_str, "usr": day_cpu["usr"], "sys": day_cpu["sys"], "iowait": day_cpu["iowait"], "idle": day_cpu["idle"]})
            perf_data["runq_trend"] = perf_data.get("runq_trend", [])
            perf_data["runq_trend"].append({"day": day_str, "max": day_runq["max_runq_sz"] or 0, "avg": day_runq["avg_runq_sz"] or 0})
        if day_mem["pgpgin"] is not None or day_mem["pswpin"] is not None:
            perf_data["mem_trend"].append({
                "day": day_str, "pgpgin": day_mem["pgpgin"] or 0, "pgpgout": day_mem["pgpgout"] or 0,
                "pswpin": day_mem["pswpin"] or 0, "pswpout": day_mem["pswpout"] or 0
            })
            perf_data["free_mem_trend"] = perf_data.get("free_mem_trend", [])
            perf_data["free_mem_trend"].append({"day": day_str, "min": day_free_mem["min_kbmemfree"] or 0})
            perf_data["swap_trend"] = perf_data.get("swap_trend", [])
            perf_data["swap_trend"].append({"day": day_str, "max_pct": day_swap_used["max_pctswpused"] or 0})
        if day_disk["svctm"] is not None:
            perf_data["disk_trend"].append({"day": day_str, "svctm": day_disk["svctm"], "await": day_disk["await"], "util": day_disk["util"]})

    return perf_data

def build_all_data(fa_root: str, os_dir: Optional[Path] = None) -> dict:
    root = Path(fa_root)
    auto_coll_dir = root / "auto_collection"
    log_dir = root / "log"

    data = {}
    pm_data = parse_mfec_pm(auto_coll_dir)
    data.update(pm_data)

    data["invalid_objects_list"] = parse_invalid_objects(auto_coll_dir)
    data["registry_list"] = parse_registry(auto_coll_dir)
    raw_growth = parse_data_growth_logs(root)
    data["growth_rows_raw"] = _enrich_growth_with_allocated(
        raw_growth,
        pm_data.get("tbs_size_mb", {}),
        datafiles=pm_data.get("datafiles", [])
    )
    data["oracle_stats"] = parse_oracle_stats(auto_coll_dir)
    data["patch_info"] = parse_patch_info(log_dir)
    data["os_info"] = parse_os_info(os_dir)
    data["section53"] = parse_section53_data(auto_coll_dir)
    data["backup_info"] = parse_backup_info(root)
    data["os_perf"] = parse_os_perf_data(root)

    return data


# ================================================================
# ZIP input + report generation
# ================================================================


def find_fast_assessment_root(extracted_root: Path) -> Path:
    candidates = []
    for p in extracted_root.rglob("*"):
        if not p.is_dir():
            continue
        if (p / "auto_collection").is_dir() and (p / "log").is_dir():
            candidates.append(p)
    if not candidates:
        raise FileNotFoundError("Cannot find fast_assessment folder inside ZIP. Expected folder containing auto_collection/ and log/.")
    candidates.sort(key=lambda x: len(str(x)))
    return candidates[0]


def find_fast_assessment_root_or_nested_zip(extracted_root: Path) -> tuple[Path, Optional[Path]]:
    try:
        return find_fast_assessment_root(extracted_root), None
    except FileNotFoundError:
        pass

    nested_zips = sorted(p for p in extracted_root.rglob("*.zip") if p.is_file())
    nested_candidates = []
    for nested_zip in nested_zips:
        nested_extract_dir = extracted_root / f"__nested_{nested_zip.stem}"
        nested_extract_dir.mkdir(parents=True, exist_ok=True)
        try:
            with zipfile.ZipFile(nested_zip, "r") as zf:
                zf.extractall(nested_extract_dir)
            fa_root = find_fast_assessment_root(nested_extract_dir)
        except (zipfile.BadZipFile, FileNotFoundError):
            continue

        auto_collection = fa_root / "auto_collection"
        score = 0
        for required_file in ("mfec_pm.txt", "registry.log", "invalid_obj.log", "mfec_oracle_stats.log"):
            if (auto_collection / required_file).is_file():
                score += 10
        if list((fa_root / "report").glob("*_top*.lst")):
            score += 5
        nested_candidates.append((score, nested_zip.name, fa_root, nested_zip))

    if nested_candidates:
        nested_candidates.sort(key=lambda item: (-item[0], item[1]))
        _, _, fa_root, nested_zip = nested_candidates[0]
        return fa_root, nested_zip

    raise FileNotFoundError(
        "Cannot find fast_assessment folder inside ZIP. Expected folder containing "
        "auto_collection/ and log/, or a ZIP bundle containing a Fast Assessment ZIP."
    )


def discover_fast_assessment_sources(extracted_root: Path) -> list[dict]:
    sources = []
    try:
        fa_root = find_fast_assessment_root(extracted_root)
        sources.append({"root": fa_root, "source_zip": None, "score": 100})
        return sources
    except FileNotFoundError:
        pass

    for nested_zip in sorted(p for p in extracted_root.rglob("*.zip") if p.is_file()):
        nested_extract_dir = extracted_root / f"__nested_{nested_zip.stem}"
        nested_extract_dir.mkdir(parents=True, exist_ok=True)
        try:
            with zipfile.ZipFile(nested_zip, "r") as zf:
                zf.extractall(nested_extract_dir)
            fa_root = find_fast_assessment_root(nested_extract_dir)
        except (zipfile.BadZipFile, FileNotFoundError):
            continue

        auto_collection = fa_root / "auto_collection"
        score = 0
        for required_file in ("mfec_pm.txt", "registry.log", "invalid_obj.log", "mfec_oracle_stats.log"):
            if (auto_collection / required_file).is_file():
                score += 10
        if list((fa_root / "report").glob("*_top*.lst")):
            score += 5
        sources.append({"root": fa_root, "source_zip": nested_zip, "score": score})

    if not sources:
        raise FileNotFoundError(
            "Cannot find fast_assessment folder inside ZIP. Expected folder containing "
            "auto_collection/ and log/, or a ZIP bundle containing a Fast Assessment ZIP."
        )
    sources.sort(key=lambda item: (-item["score"], (item["source_zip"].name if item["source_zip"] else "")))
    return sources


def json_default(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    return str(obj)


def status_from_pct_free(pct_free: float) -> str:
    if pct_free < 10:
        return "Critical"
    if pct_free < 20:
        return "Warning"
    return "OK"


def add_key_value_table(doc, rows):
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for k, v in rows:
        cells = table.add_row().cells
        cells[0].text = str(k)
        cells[1].text = str(v)
    return table


def add_list_table(doc, headers, rows, max_rows=None):
    if max_rows is not None:
        rows = rows[:max_rows]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = str(h)
    for row in rows:
        cells = table.add_row().cells
        for i, h in enumerate(headers):
            cells[i].text = str(row.get(h, ""))
    return table


def create_docx_report(data: dict, output_docx: Path, source_zip: Path):
    try:
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except Exception as e:
        raise RuntimeError("Missing python-docx. Install with: pip install python-docx") from e

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)

    title = doc.add_heading("Preventive Maintenance Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph(f"Source ZIP: {source_zip.name}\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("Suggestion Summary", level=1)
    summary_rows = []

    tablespace_free = data.get("tablespace_free", {})
    critical_tbs = []
    warning_tbs = []
    ok_tbs = []
    for name, info in tablespace_free.items():
        pct = float(info.get("pct_free", 0) or 0)
        st = status_from_pct_free(pct)
        if st == "Critical":
            critical_tbs.append(name)
        elif st == "Warning":
            warning_tbs.append(name)
        else:
            ok_tbs.append(name)

    summary_rows.append({"Topic": "Database Configuration", "Status": "OK", "Remark": "Basic configuration parsed"})
    summary_rows.append({"Topic": "Database Patch", "Status": "Warning" if data.get("patch_info") else "OK", "Remark": "Check patch recommendation"})
    summary_rows.append({"Topic": "Tablespace Free Space", "Status": "Critical" if critical_tbs else "Warning" if warning_tbs else "OK", "Remark": ", ".join(critical_tbs or warning_tbs)})
    summary_rows.append({"Topic": "Invalid Objects", "Status": "Warning" if data.get("invalid_objects_list") else "OK", "Remark": f"{len(data.get('invalid_objects_list', []))} objects"})
    add_list_table(doc, ["Topic", "Status", "Remark"], summary_rows)

    doc.add_heading("1. Database Information", level=1)
    add_key_value_table(
        doc,
        [
            ("Instance name", data.get("instance_name", "")),
            ("RDBMS Version", data.get("rdbms_version", "")),
            ("Archive mode", data.get("archive_mode", "")),
            ("Optimizer mode", data.get("optimizer_mode", "")),
            ("Disk Space", data.get("disk_space", "")),
            ("Datafiles", len(data.get("datafiles", []))),
            ("Tempfiles", len(data.get("tempfiles", []))),
            ("Redo logs", len(data.get("redo_logs", []))),
            ("Control files", len(data.get("control_files", []))),
        ],
    )

    doc.add_heading("2. Database Parameters", level=1)
    param_rows = [{"Name": k, "Value": v} for k, v in data.get("db_parameters", {}).items()]
    if param_rows:
        add_list_table(doc, ["Name", "Value"], param_rows, max_rows=80)
    else:
        doc.add_paragraph("No database parameter data found.")

    doc.add_heading("3. Database Patch", level=1)
    patch_rows = data.get("patch_info", [])
    if patch_rows:
        add_list_table(doc, ["component", "current_version", "recommended_version", "suggestion"], patch_rows)
    else:
        doc.add_paragraph("No patch information found.")

    doc.add_heading("4. Performance Review", level=1)
    perf_rows = [{"Metric": k, "Value": v} for k, v in data.get("perf_stats", {}).items()]
    if perf_rows:
        add_list_table(doc, ["Metric", "Value"], perf_rows)
    else:
        doc.add_paragraph("No performance statistics found.")

    doc.add_heading("5. Database Growth Rate", level=1)
    growth_rows = []
    for r in data.get("growth_rows_raw", []):
        growth_rows.append({"Month": r.get("month"), "Used GB": r.get("used_gb"), "Allocated GB": r.get("allocated_gb"), "Usage %": r.get("pct_used")})
    if growth_rows:
        add_list_table(doc, ["Month", "Used GB", "Allocated GB", "Usage %"], growth_rows)
    else:
        doc.add_paragraph("No database growth data found.")

    doc.add_heading("6. Tablespace Free Space", level=1)
    tbs_rows = []
    for name, info in tablespace_free.items():
        pct = float(info.get("pct_free", 0) or 0)
        tbs_rows.append({"Tablespace": name, "Free %": pct, "Status": status_from_pct_free(pct)})
    tbs_rows.sort(key=lambda x: x["Free %"])
    if tbs_rows:
        add_list_table(doc, ["Tablespace", "Free %", "Status"], tbs_rows)
    else:
        doc.add_paragraph("No tablespace free space data found.")

    doc.add_heading("7. Database Registry", level=1)
    registry_rows = data.get("registry_list", [])
    if registry_rows:
        add_list_table(doc, ["comp_id", "version", "status", "last_modified"], registry_rows)
    else:
        doc.add_paragraph("No registry data found.")

    doc.add_heading("Appendix A – Invalid Objects", level=1)
    invalid_rows = data.get("invalid_objects_list", [])
    if invalid_rows:
        add_list_table(doc, ["creator", "object_name", "object_type", "status"], invalid_rows, max_rows=300)
        if len(invalid_rows) > 300:
            doc.add_paragraph(f"Showing first 300 of {len(invalid_rows)} invalid objects.")
    else:
        doc.add_paragraph("No invalid object data found.")

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_docx)


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    # Method 1: Microsoft Word COM automation on Windows
    if os.name == "nt":
        word = None
        doc = None
        try:
            import win32com.client  # type: ignore

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path), ReadOnly=False)
            doc.Repaginate()
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Save()
            doc.SaveAs(str(pdf_path), FileFormat=17)
            return pdf_path.exists()
        except Exception:
            pass
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

    # Method 2: LibreOffice if installed
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            generated = docx_path.with_suffix(".pdf")
            if generated.exists() and generated != pdf_path:
                generated.rename(pdf_path)
            return pdf_path.exists()
        except Exception:
            return False

    return False


def process_zip_to_reports(zip_path: Path, output_dir: Path, metadata: dict = None, create_pdf: bool = True) -> tuple[Path, Optional[Path], Path]:
    if not zip_path.exists():
        raise FileNotFoundError(f"ZIP file not found: {zip_path}")
    if zip_path.suffix.lower() != ".zip":
        raise ValueError("Input file must be .zip")

    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="fa_zip_") as tmp:
        tmp_root = Path(tmp)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(tmp_root)
        sources = discover_fast_assessment_sources(tmp_root)
        primary_source = sources[0]
        fa_root = primary_source["root"]
        nested_zip = primary_source["source_zip"]
        source_zip = nested_zip or zip_path
        base_name = source_zip.stem
        docx_path = output_dir / f"{base_name}_PM_Report.docx"
        pdf_path = output_dir / f"{base_name}_PM_Report.pdf"
        json_path = output_dir / f"{base_name}_parsed_data.json"
        os_dir = find_os_dir(tmp_root, fa_root)
        data = build_all_data(str(fa_root), os_dir=os_dir)

        bundle_nodes = []
        for source in sources:
            source_root = source["root"]
            node_os_dir = find_os_dir(tmp_root, source_root)
            node_os_info = parse_os_info(node_os_dir)
            source_name = source["source_zip"].name if source["source_zip"] else zip_path.name
            bundle_nodes.append(
                {
                    "source": source_name,
                    "has_database_results": (source_root / "auto_collection" / "mfec_pm.txt").is_file(),
                    "has_registry": (source_root / "auto_collection" / "registry.log").is_file(),
                    "os_info": node_os_info,
                    "patch_info": parse_patch_info(source_root / "log"),
                }
            )
        data["bundle_sources"] = [node["source"] for node in bundle_nodes]
        data["bundle_nodes"] = bundle_nodes
        if not data.get("patch_info"):
            for node in bundle_nodes:
                if node.get("patch_info"):
                    data["patch_info"] = node["patch_info"]
                    break
        
        if metadata:
            data["gui_metadata"] = metadata

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=json_default)

        from report_writer_light import create_docx_report

        create_docx_report(
            data=data,
            output_docx=docx_path,
            source_zip=source_zip,
            output_dir=output_dir
        )
        pdf_ok = convert_docx_to_pdf(docx_path, pdf_path) if create_pdf else False

    return docx_path, pdf_path if pdf_ok else None, json_path


def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from pathlib import Path
    from datetime import datetime

    root = tk.Tk()
    root.title("PM Document Converter")
    root.geometry("950x650")
    root.resizable(False, False)

    # Variables for Project Info
    proj_name = tk.StringVar()
    proj_no = tk.StringVar()
    db_name = tk.StringVar()
    cust_full = tk.StringVar()
    cust_abbrev = tk.StringVar()
    q_var = tk.StringVar()
    sale_var = tk.StringVar()
    
    # Contact Form Variables
    contact_type_var = tk.StringVar(value="Customer")
    name_var = tk.StringVar()
    phone_var = tk.StringVar()
    email_var = tk.StringVar()
    
    _cust_data = []
    _eng_data = []

    # Report Type & Paths
    report_type_var = tk.StringVar(value="Preventive Maintenance")
    os_path_var = tk.StringVar()
    db_path_var = tk.StringVar()
    output_dir = tk.StringVar()

    # Layout: Left Treeview, Right Main Container
    left_frame = tk.Frame(root, width=150, bg="white", relief="sunken", bd=1)
    left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 10), pady=(10, 10))
    
    tree = ttk.Treeview(left_frame, show="tree", selectmode="browse")
    tree.pack(fill=tk.BOTH, expand=True)
    
    main_container = tk.Frame(root)
    main_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10), pady=(10, 10))
    main_container.grid_rowconfigure(0, weight=1)
    main_container.grid_columnconfigure(0, weight=1)
    
    # ------------------------------------------------------------------
    # FRAME 1: Project Information (New Project)
    # ------------------------------------------------------------------
    frame_project = tk.Frame(main_container)
    frame_project.grid(row=0, column=0, sticky="nsew")
    
    proj_lf = ttk.LabelFrame(frame_project, text="Project Information")
    proj_lf.pack(fill=tk.X, pady=5)
    
    tk.Label(proj_lf, text="Project Name").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=proj_name, width=30).grid(row=0, column=1, padx=10, pady=5)
    tk.Label(proj_lf, text="Customer (full)").grid(row=0, column=2, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=cust_full, width=30).grid(row=0, column=3, padx=10, pady=5)
    
    tk.Label(proj_lf, text="Project No.").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=proj_no, width=15).grid(row=1, column=1, sticky="w", padx=10, pady=5)
    tk.Label(proj_lf, text="Customer (abbrev)").grid(row=1, column=2, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=cust_abbrev, width=15).grid(row=1, column=3, sticky="w", padx=10, pady=5)
    
    tk.Label(proj_lf, text="Database Name").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=db_name, width=15).grid(row=2, column=1, sticky="w", padx=10, pady=5)
    tk.Label(proj_lf, text="Q").grid(row=2, column=2, sticky="w", padx=10, pady=5)
    tk.Entry(proj_lf, textvariable=q_var, width=15).grid(row=2, column=3, sticky="w", padx=10, pady=5)

    pers_lf = ttk.LabelFrame(frame_project, text="Personal Information")
    pers_lf.pack(fill=tk.BOTH, expand=True, pady=5)
    
    left_pers = tk.Frame(pers_lf)
    left_pers.grid(row=0, column=0, padx=10, pady=10, sticky="nw")
    
    tk.Label(left_pers, text="Customer(s)").grid(row=0, column=0, sticky="nw")
    cust_list = tk.Listbox(left_pers, height=4, width=35)
    cust_list.grid(row=0, column=1, rowspan=3)
    tk.Button(left_pers, text=">").grid(row=0, column=2, padx=4)
    tk.Button(left_pers, text="^").grid(row=1, column=2, padx=4)
    tk.Button(left_pers, text="v").grid(row=2, column=2, padx=4)
    
    tk.Label(left_pers, text="Sale").grid(row=3, column=0, sticky="w", pady=10)
    tk.Entry(left_pers, textvariable=sale_var, width=35).grid(row=3, column=1, pady=10)
    tk.Button(left_pers, text=">").grid(row=3, column=2, padx=4)
    
    tk.Label(left_pers, text="Engineer(s)").grid(row=4, column=0, sticky="nw")
    eng_list = tk.Listbox(left_pers, height=4, width=35)
    eng_list.grid(row=4, column=1, rowspan=3)
    tk.Button(left_pers, text=">").grid(row=4, column=2, padx=4)
    tk.Button(left_pers, text="^").grid(row=5, column=2, padx=4)
    tk.Button(left_pers, text="v").grid(row=6, column=2, padx=4)
    
    ttk.Separator(pers_lf, orient=tk.VERTICAL).grid(row=0, column=1, sticky="ns", padx=(10, 20))
    form_frame = tk.Frame(pers_lf)
    form_frame.grid(row=0, column=2, padx=10, pady=10, sticky="n")
    
    tk.Label(form_frame, text="Select on").grid(row=0, column=0, sticky="w", pady=2)
    ttk.Combobox(form_frame, textvariable=contact_type_var, values=["Customer", "Engineer"], width=12, state="readonly").grid(row=0, column=1, sticky="w", pady=2)
    tk.Label(form_frame, text="Name").grid(row=1, column=0, sticky="w", pady=2)
    tk.Entry(form_frame, textvariable=name_var, width=30).grid(row=1, column=1, pady=2, sticky="w")
    tk.Label(form_frame, text="Phone").grid(row=2, column=0, sticky="w", pady=2)
    tk.Entry(form_frame, textvariable=phone_var, width=15).grid(row=2, column=1, sticky="w", pady=2)
    tk.Label(form_frame, text="Email Address").grid(row=3, column=0, sticky="w", pady=2)
    tk.Entry(form_frame, textvariable=email_var, width=30).grid(row=3, column=1, pady=2, sticky="w")
    
    btn_frame = tk.Frame(form_frame)
    btn_frame.grid(row=4, column=0, columnspan=2, pady=15)
    
    def add_contact():
        name = name_var.get().strip()
        phone = phone_var.get().strip()
        email = email_var.get().strip()
        if not name: return
        val = f"{name} ({phone}) [{email}]"
        cdict = {"name": name, "phone": phone, "email": email}
        if contact_type_var.get() == "Customer":
            cust_list.insert(tk.END, val)
            _cust_data.append(cdict)
        else:
            eng_list.insert(tk.END, val)
            _eng_data.append(cdict)
        name_var.set(""); phone_var.set(""); email_var.set("")
        
    def del_contact():
        if contact_type_var.get() == "Customer":
            s = cust_list.curselection()
            if s: idx=s[0]; cust_list.delete(idx); _cust_data.pop(idx)
        else:
            s = eng_list.curselection()
            if s: idx=s[0]; eng_list.delete(idx); _eng_data.pop(idx)

    tk.Button(btn_frame, text="Add", width=6, command=add_contact).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text="Edit", width=6).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text="Delete", width=6, command=del_contact).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text="Clear", width=6).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text="Load", width=6).pack(side=tk.LEFT, padx=2)

    # ------------------------------------------------------------------
    # FRAME 2: Version
    # ------------------------------------------------------------------
    frame_version = tk.Frame(main_container)
    frame_version.grid(row=0, column=0, sticky="nsew")
    
    vers_lf = ttk.LabelFrame(frame_version, text="Release Version")
    vers_lf.pack(fill=tk.BOTH, expand=True, pady=5)
    
    # Change Record Table
    cr_frame = tk.Frame(vers_lf)
    cr_frame.pack(fill=tk.X, padx=10, pady=10)
    tk.Label(cr_frame, text="Changed Record", font=("Arial", 9, "bold")).pack(anchor="w")
    
    cr_canvas = tk.Canvas(cr_frame, height=120, bg="lightgray")
    cr_canvas.pack(fill=tk.X)
    
    cr_inner = tk.Frame(cr_canvas, bg="lightgray")
    cr_canvas.create_window((0,0), window=cr_inner, anchor="nw")
    
    # Headers
    tk.Label(cr_inner, text="Date", width=12, bg="white", relief="ridge").grid(row=0, column=0)
    tk.Label(cr_inner, text="Author", width=25, bg="white", relief="ridge").grid(row=0, column=1)
    tk.Label(cr_inner, text="Get", width=4, bg="white", relief="ridge").grid(row=0, column=2)
    tk.Label(cr_inner, text="Version", width=8, bg="white", relief="ridge").grid(row=0, column=3)
    tk.Label(cr_inner, text="Change Reference", width=35, bg="white", relief="ridge").grid(row=0, column=4)
    tk.Label(cr_inner, text="Del", width=4, bg="white", relief="ridge").grid(row=0, column=5)
    
    _cr_rows = []
    
    def add_cr_row():
        r = len(_cr_rows) + 1
        d_var = tk.StringVar()
        a_var = tk.StringVar()
        v_var = tk.StringVar(value="1.0")
        ref_var = tk.StringVar(value="Initial Document")
        
        e_d = tk.Entry(cr_inner, textvariable=d_var, width=12)
        e_d.grid(row=r, column=0)
        e_a = tk.Entry(cr_inner, textvariable=a_var, width=25)
        e_a.grid(row=r, column=1)
        
        def set_today(v=d_var): v.set(datetime.now().strftime("%d-%b-%Y"))
        tk.Button(cr_inner, text="G", width=2, command=set_today).grid(row=r, column=2)
        
        e_v = tk.Entry(cr_inner, textvariable=v_var, width=8)
        e_v.grid(row=r, column=3)
        e_ref = tk.Entry(cr_inner, textvariable=ref_var, width=35)
        e_ref.grid(row=r, column=4)
        
        row_dict = {"d": d_var, "a": a_var, "v": v_var, "ref": ref_var, "widgets": [e_d, e_a, e_v, e_ref]}
        
        def del_row(rd=row_dict):
            for w in rd["widgets"]: w.destroy()
            _cr_rows.remove(rd)
            
        btn_del = tk.Button(cr_inner, text="X", width=2, fg="red", command=del_row)
        btn_del.grid(row=r, column=5)
        row_dict["widgets"].extend([cr_inner.grid_slaves(row=r, column=2)[0], btn_del])
        _cr_rows.append(row_dict)

    add_cr_row()
    tk.Button(cr_frame, text="Add Row", command=add_cr_row).pack(anchor="w", pady=2)
    
    # Reviewers Table
    rev_frame = tk.Frame(vers_lf)
    rev_frame.pack(fill=tk.X, padx=10, pady=10)
    tk.Label(rev_frame, text="Reviewers", font=("Arial", 9, "bold")).pack(anchor="w")
    
    rev_canvas = tk.Canvas(rev_frame, height=120, bg="lightgray")
    rev_canvas.pack(fill=tk.X)
    
    rev_inner = tk.Frame(rev_canvas, bg="lightgray")
    rev_canvas.create_window((0,0), window=rev_inner, anchor="nw")
    
    tk.Label(rev_inner, text="Date", width=12, bg="white", relief="ridge").grid(row=0, column=0)
    tk.Label(rev_inner, text="Name", width=25, bg="white", relief="ridge").grid(row=0, column=1)
    tk.Label(rev_inner, text="Get", width=4, bg="white", relief="ridge").grid(row=0, column=2)
    tk.Label(rev_inner, text="Position", width=35, bg="white", relief="ridge").grid(row=0, column=3)
    tk.Label(rev_inner, text="Del", width=4, bg="white", relief="ridge").grid(row=0, column=4)
    
    _rev_rows = []
    
    def add_rev_row():
        r = len(_rev_rows) + 1
        d_var = tk.StringVar()
        n_var = tk.StringVar()
        p_var = tk.StringVar()
        
        e_d = tk.Entry(rev_inner, textvariable=d_var, width=12)
        e_d.grid(row=r, column=0)
        e_n = tk.Entry(rev_inner, textvariable=n_var, width=25)
        e_n.grid(row=r, column=1)
        
        def set_today(v=d_var): v.set(datetime.now().strftime("%d-%b-%Y"))
        tk.Button(rev_inner, text="G", width=2, command=set_today).grid(row=r, column=2)
        
        e_p = tk.Entry(rev_inner, textvariable=p_var, width=35)
        e_p.grid(row=r, column=3)
        
        row_dict = {"d": d_var, "n": n_var, "p": p_var, "widgets": [e_d, e_n, e_p]}
        
        def del_row(rd=row_dict):
            for w in rd["widgets"]: w.destroy()
            _rev_rows.remove(rd)
            
        btn_del = tk.Button(rev_inner, text="X", width=2, fg="red", command=del_row)
        btn_del.grid(row=r, column=4)
        row_dict["widgets"].extend([rev_inner.grid_slaves(row=r, column=2)[0], btn_del])
        _rev_rows.append(row_dict)

    add_rev_row()
    tk.Button(rev_frame, text="Add Row", command=add_rev_row).pack(anchor="w", pady=2)


    # ------------------------------------------------------------------
    # FRAME 3: Report Type
    # ------------------------------------------------------------------
    frame_report = tk.Frame(main_container)
    frame_report.grid(row=0, column=0, sticky="nsew")
    
    rep_lf = ttk.LabelFrame(frame_report, text="Report")
    rep_lf.pack(fill=tk.X, pady=5)
    tk.Label(rep_lf, text="Report Type").grid(row=0, column=0, padx=10, pady=10)
    ttk.Combobox(rep_lf, textvariable=report_type_var, values=["Preventive Maintenance", "Installation"], width=30, state="readonly").grid(row=0, column=1, pady=10)
    
    os_lf = ttk.LabelFrame(frame_report, text="OS Information")
    os_lf.pack(fill=tk.X, pady=5)
    tk.Label(os_lf, text="OS").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    ttk.Combobox(os_lf, values=["Linux", "Windows"], width=15).grid(row=0, column=1, sticky="w", pady=5)
    tk.Label(os_lf, text="Input Path").grid(row=1, column=0, padx=10, pady=5, sticky="nw")
    tk.Entry(os_lf, textvariable=os_path_var, width=55).grid(row=1, column=1, pady=5, sticky="w")
    tk.Button(os_lf, text="Browse", command=lambda: os_path_var.set(filedialog.askopenfilename(title="Select OS ZIP") or os_path_var.get())).grid(row=1, column=2, padx=5)
    
    db_lf = ttk.LabelFrame(frame_report, text="Database Information")
    db_lf.pack(fill=tk.X, pady=5)
    tk.Label(db_lf, text="Database").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    ttk.Combobox(db_lf, values=["Oracle"], width=15).grid(row=0, column=1, sticky="w", pady=5)
    tk.Label(db_lf, text="Input Path").grid(row=1, column=0, padx=10, pady=5, sticky="nw")
    tk.Entry(db_lf, textvariable=db_path_var, width=55).grid(row=1, column=1, pady=5, sticky="w")
    tk.Button(db_lf, text="Browse", command=lambda: db_path_var.set(filedialog.askopenfilename(title="Select PM ZIP file", filetypes=[("ZIP files", "*.zip")]) or db_path_var.get())).grid(row=1, column=2, padx=5)

    out_lf = ttk.LabelFrame(frame_report, text="Output Folder")
    out_lf.pack(fill=tk.X, pady=5)
    tk.Label(out_lf, text="Output Dir").grid(row=0, column=0, padx=10, pady=5, sticky="nw")
    tk.Entry(out_lf, textvariable=output_dir, width=55).grid(row=0, column=1, pady=5, sticky="w")
    tk.Button(out_lf, text="Browse", command=lambda: output_dir.set(filedialog.askdirectory(title="Select output folder") or output_dir.get())).grid(row=0, column=2, padx=5)

    # ------------------------------------------------------------------
    # Navigation Logic
    # ------------------------------------------------------------------
    node_proj = tree.insert("", "end", text="New Project", open=True)
    node_vers = tree.insert(node_proj, "end", text="Version")
    node_rep = tree.insert(node_proj, "end", text="Report Type")
    
    frame_map = {
        node_proj: frame_project,
        node_vers: frame_version,
        node_rep: frame_report
    }
    
    def on_tree_select(event):
        sel = tree.selection()
        if sel:
            item = sel[0]
            if item in frame_map:
                frame_map[item].tkraise()
                
    tree.bind("<<TreeviewSelect>>", on_tree_select)
    frame_project.tkraise() # Default view

    # ------------------------------------------------------------------
    # Export Logic
    # ------------------------------------------------------------------
    def export_report():
        target_zip = db_path_var.get()
        if not target_zip:
            messagebox.showwarning("Warning", "Please select Database Input Path (ZIP file) on the Report Type page.")
            return

        cr_data = [{"date": r["d"].get(), "author": r["a"].get(), "version": r["v"].get(), "ref": r["ref"].get()} for r in _cr_rows]
        rev_data = [{"date": r["d"].get(), "name": r["n"].get(), "position": r["p"].get()} for r in _rev_rows]

        metadata = {
            "project_name": proj_name.get(),
            "project_no": proj_no.get(),
            "db_name": db_name.get(),
            "customer_full": cust_full.get(),
            "customer_abbrev": cust_abbrev.get(),
            "quarter": q_var.get(),
            "sale": sale_var.get(),
            "customers": _cust_data,
            "engineers": _eng_data,
            "change_records": cr_data,
            "reviewers": rev_data,
            "report_type": report_type_var.get()
        }

        try:
            out = Path(output_dir.get()) if output_dir.get() else Path.cwd() / "report_output"
            docx_path, pdf_path, json_path = process_zip_to_reports(Path(target_zip), out, metadata=metadata)

            msg = f"DOCX: {docx_path}\nJSON: {json_path}"
            if pdf_path: msg += f"\nPDF: {pdf_path}"
            else: msg += "\nPDF: not created"

            messagebox.showinfo("Export completed", msg)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Export failed", str(e))

    export_frame = tk.Frame(main_container)
    export_frame.grid(row=1, column=0, sticky="ew", pady=10)
    export_btn = ttk.Button(export_frame, text="Export DOCX / PDF", command=export_report)
    export_btn.pack(side=tk.RIGHT, padx=20)

    root.mainloop()

if __name__ == "__main__":
    try:
        run_gui()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
