#!/usr/bin/env python3
"""
merge_to_master.py  (hardened)
------------------------------
Ingest Excel/CSV files (single file, multiple files, or whole folders) and
pour all row data into one master file using a raw_column_name →
suggested_canonical mapping.

Same-row rule when multiple raw columns map to one canonical field:
  - Process source columns left-to-right in each sheet.
  - First non-empty value is kept.
  - If a later value is identical (case-insensitive, trimmed) → skip it.
  - If a later value is different → append with " | " separator.
  - Different rows / sheets / files are never merged together.

Only columns present in the mapping are exported — the mapping acts as a
column allowlist; unmapped columns are listed by NAME ONLY in the
Unmapped_Headers sheet, their data never enters the master.

Hardening (vs. the original draft):
  * Every sheet of the output is sanitized (control characters stripped,
    formula-leading values escaped to literal text) — fixes both the
    IllegalCharacterError crash at the final write and formula injection
    via the Unmapped_Headers / log sheets.
  * CSV files are read with encoding fallback (utf-8-sig -> utf-8 ->
    cp1252 -> latin-1) instead of being skipped when not UTF-8.
  * Every failed file/sheet is recorded in the Ingest_Log with
    status=error — nothing can go missing silently.
  * The Excel row limit (1,048,576 rows/sheet) is guarded: a larger master
    is split across Master, Master_2, ... sheets instead of crashing.
  * Preview part files (*_partNNN.xlsx) are skipped by default so sample
    rows are not double-ingested next to their source files; pass
    --include-previews to ingest preview workbooks deliberately.
  * The output file and the mapping file are excluded from ingestion even
    when they sit inside an input folder.
  * Sources are opened read-only and never modified; no network activity.

Usage:
    python merge_to_master.py --input "Category-1" \\
        --mapping "Category-1/column_merge_suggestions.xlsx" \\
        --mapping-sheet "Column_Samples" \\
        --output "Category-1/category1_master_merged.xlsx" \\
        --overwrite

Mapping file must contain columns (names flexible):
    raw_column_name  |  suggested_canonical

Options:
    --input PATH        source file or folder (repeatable)
    --mapping PATH      Excel/CSV mapping file
    --mapping-sheet S   sheet name for Excel mapping (default: auto)
    --output PATH       full path for merged output (.xlsx or .csv)
    --include-source    add source_file + source_sheet columns
    --include-previews  also ingest *_partNNN.xlsx preview workbooks
    --overwrite         replace existing output

With a .csv output, the master rows go to the CSV and the accounting
(Ingest_Log, Run_Summary, Mapping_Used, Unmapped_Headers) is written to
<output>_ingest_log.xlsx alongside it.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import warnings
from collections import OrderedDict

import pandas as pd

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

RESERVED_SHEETS = {"Index", "Report"}
PREVIEW_HEADER_ROW = 8
PREVIEW_META_ROWS = 8
EXCEL_MAX_ROWS = 1_048_576  # hard xlsx limit per sheet, including header row

EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".xlam", ".ods"}
CSV_EXTS = {".csv", ".tsv", ".txt"}

SKIP_NAMES = {
    "column_merge_suggestions.xlsx",
    "category1_unique_columns_merged.xlsx",
    "all_unique_columns_merged.xlsx",
    "cursor_all_cleaned_unique_columns_merged.xlsx",
}
SKIP_SUFFIXES = (
    "_unique_columns.xlsx",
    "_merge_suggestions.xlsx",
    "_master_merged.xlsx",
)
PART_FILE_RE = re.compile(r"_part\d{3}\.(?:xlsx|xlsm)$", re.I)

RAW_COL_ALIASES = {"raw_column_name", "raw column name", "raw_column", "column_name", "header", "source_column"}
CANONICAL_ALIASES = {
    "suggested_canonical", "canonical", "canonical_field", "target_column",
    "merged_column", "destination_column", "display_group",
}

CSV_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]

try:
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE  # type: ignore
except Exception:  # pragma: no cover - fallback if openpyxl relocates it
    ILLEGAL_CHARACTERS_RE = re.compile(
        r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f﷐-﷟￾￿]"
    )

FORMULA_LEADERS = ("=", "+", "-", "@", "\t", "\r")


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and v != v:
        return True
    return str(v).strip().lower() in ("", "nan", "none")


def cell_str(v) -> str:
    return "" if is_blank(v) else str(v).strip()


def safe_cell(v, max_chars=32000) -> str:
    s = cell_str(v)
    s = ILLEGAL_CHARACTERS_RE.sub("", s)
    if len(s) > max_chars:
        s = s[: max_chars - 1] + "…"
    if s.startswith(FORMULA_LEADERS):
        s = "'" + s
    return s


def df_map(df: pd.DataFrame, fn) -> pd.DataFrame:
    try:
        return df.map(fn)  # pandas >= 2.1
    except AttributeError:  # pragma: no cover - older pandas
        return df.applymap(fn)


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Sanitize every string cell of a DataFrame before writing to Excel."""
    if df.empty:
        return df
    return df_map(df, lambda v: safe_cell(v) if isinstance(v, str) else v)


def read_csv_robust(filepath: str, sep: str, nrows: int | None = None):
    """Read a CSV with encoding fallback; returns (df, encoding_used)."""
    last_error: Exception | None = None
    for encoding in CSV_ENCODINGS:
        for extra in ({}, {"engine": "python", "on_bad_lines": "skip"}):
            try:
                df = pd.read_csv(
                    filepath, sep=sep, header=None, dtype=str,
                    encoding=encoding, nrows=nrows, **extra,
                )
                return df, encoding
            except UnicodeDecodeError as e:
                last_error = e
                break  # this encoding cannot decode the file; try the next
            except Exception as e:
                last_error = e
                continue  # parser trouble; retry with the tolerant engine
    raise RuntimeError(f"could not read as CSV: {last_error}")


def normalize_for_compare(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


def excel_engine(ext: str):
    if ext == ".xls":
        return "xlrd"
    if ext == ".xlsb":
        return "pyxlsb"
    if ext == ".ods":
        return "odf"
    return None


def is_supported_data_file(fname: str) -> bool:
    ext = os.path.splitext(fname)[1].lower()
    return ext in EXCEL_EXTS or ext in CSV_EXTS


def should_skip_data_file(fname: str) -> bool:
    if fname.startswith(".") or fname.startswith("~$"):
        return True
    if fname in SKIP_NAMES:
        return True
    return any(fname.endswith(s) for s in SKIP_SUFFIXES)


def is_preview_format(df_raw: pd.DataFrame) -> bool:
    if df_raw.empty or df_raw.shape[1] < 2:
        return False
    return cell_str(df_raw.iloc[0, 0]) == "Field" and cell_str(df_raw.iloc[0, 1]) == "Value"


def parse_preview_metadata(df_raw: pd.DataFrame) -> dict:
    meta = {}
    for i in range(min(PREVIEW_META_ROWS, len(df_raw))):
        key = cell_str(df_raw.iloc[i, 0])
        val = cell_str(df_raw.iloc[i, 1]) if df_raw.shape[1] > 1 else ""
        if key and key != "Field":
            meta[key] = val
    return meta


def find_header_row(df_raw: pd.DataFrame) -> int:
    best_i, best_score = 0, float("-inf")
    for i in range(min(15, len(df_raw))):
        row = [cell_str(v) for v in df_raw.iloc[i].tolist()]
        non_empty = [v for v in row if v]
        if not non_empty:
            continue
        texty = sum(1 for v in non_empty if not re.fullmatch(r"[\d\s\-\+\(\)\.,/]+", v))
        score = texty + len(non_empty) * 0.2
        if score > best_score:
            best_i, best_score = i, score
    return best_i


def extract_sheet_table(df_raw: pd.DataFrame) -> tuple[list[str], list[list[str]], dict, int]:
    """Return headers, all data rows, preview metadata, header row (1-based)."""
    preview_meta = {}
    if is_preview_format(df_raw):
        preview_meta = parse_preview_metadata(df_raw)
        if len(df_raw) <= PREVIEW_HEADER_ROW:
            return [], [], preview_meta, 0
        hrow = PREVIEW_HEADER_ROW
    else:
        hrow = find_header_row(df_raw)

    headers = [cell_str(v) for v in df_raw.iloc[hrow].tolist()]
    while headers and not headers[-1]:
        headers.pop()

    rows: list[list[str]] = []
    for r in range(hrow + 1, len(df_raw)):
        vals = [
            cell_str(df_raw.iloc[r, c]) if c < df_raw.shape[1] else ""
            for c in range(len(headers))
        ]
        if any(vals):
            rows.append(vals)

    return headers, rows, preview_meta, hrow + 1


def merge_same_row_values(values: list[str], separator: str = " | ") -> str:
    """Merge values for one canonical field within a single source row."""
    parts: list[str] = []
    seen: set[str] = set()
    for v in values:
        v = v.strip()
        if not v:
            continue
        key = normalize_for_compare(v)
        if key in seen:
            continue
        seen.add(key)
        parts.append(v)
    return separator.join(parts)


def resolve_column(name: str, aliases: set[str], columns: list[str]) -> str | None:
    for col in columns:
        if col.strip().lower() in aliases:
            return col
    for col in columns:
        low = re.sub(r"\s+", " ", col.strip().lower())
        for alias in aliases:
            if alias in low or low.replace("_", " ") == alias.replace("_", " "):
                return col
    return None


def load_mapping(mapping_path: str, mapping_sheet: str = "") -> tuple[dict[str, str], list[str]]:
    ext = os.path.splitext(mapping_path)[1].lower()
    if ext in CSV_EXTS:
        sep = "\t" if ext in (".tsv", ".txt") else ","
        df, _enc = read_csv_robust(mapping_path, sep)
        if df.empty:
            sys.exit(f"ERROR: mapping file is empty: {mapping_path}")
        df.columns = [cell_str(c) for c in df.iloc[0].tolist()]
        df = df.iloc[1:]
    else:
        if mapping_sheet:
            df = pd.read_excel(mapping_path, sheet_name=mapping_sheet, dtype=str)
        else:
            xl = pd.ExcelFile(mapping_path)
            preferred = [s for s in xl.sheet_names if s.lower() in ("column_samples", "mapping", "map")]
            sheet = preferred[0] if preferred else xl.sheet_names[0]
            df = pd.read_excel(mapping_path, sheet_name=sheet, dtype=str)
            xl.close()

    cols = [str(c) for c in df.columns]
    raw_col = resolve_column("", RAW_COL_ALIASES, cols)
    canon_col = resolve_column("", CANONICAL_ALIASES, cols)

    if not raw_col or not canon_col:
        sys.exit(
            f"ERROR: mapping file must have raw + canonical columns.\n"
            f"  Found columns: {cols}\n"
            f"  Expected something like: raw_column_name, suggested_canonical"
        )

    raw_to_canonical: OrderedDict[str, str] = OrderedDict()
    for _, row in df.iterrows():
        raw = cell_str(row[raw_col])
        canonical = cell_str(row[canon_col])
        if not raw or not canonical:
            continue
        if raw.lower() in ("raw_column_name", "column_name"):
            continue
        raw_to_canonical[raw] = canonical

    if not raw_to_canonical:
        sys.exit("ERROR: mapping file contains no valid raw -> canonical rows.")

    canonical_order: list[str] = []
    seen_canonical: set[str] = set()
    for canonical in raw_to_canonical.values():
        if canonical not in seen_canonical:
            seen_canonical.add(canonical)
            canonical_order.append(canonical)

    return dict(raw_to_canonical), canonical_order


def collect_input_files(
    inputs: list[str],
    exclude_paths: set[str],
    include_previews: bool,
) -> tuple[list[str], list[str]]:
    """Return (data_files, skipped_preview_parts)."""
    files: list[str] = []
    skipped_parts: list[str] = []
    exclude_abs = {os.path.abspath(p) for p in exclude_paths}

    def consider(filepath: str) -> None:
        fname = os.path.basename(filepath)
        if os.path.abspath(filepath) in exclude_abs:
            return
        if not is_supported_data_file(fname) or should_skip_data_file(fname):
            return
        if not include_previews and PART_FILE_RE.search(fname):
            skipped_parts.append(filepath)
            return
        files.append(filepath)

    for path in inputs:
        path = os.path.abspath(path)
        if os.path.isfile(path):
            before = len(files) + len(skipped_parts)
            consider(path)
            if len(files) + len(skipped_parts) == before:
                print(f"  [SKIP] not a supported data file: {path}")
        elif os.path.isdir(path):
            for dirpath, dirnames, filenames in os.walk(path):
                dirnames[:] = [d for d in dirnames if not d.startswith(".") and d != "__MACOSX"]
                for fname in sorted(filenames):
                    consider(os.path.join(dirpath, fname))
        else:
            print(f"  [WARN] path not found: {path}")
    return sorted(set(files)), sorted(set(skipped_parts))


def build_merged_row(
    headers: list[str],
    row: list[str],
    raw_to_canonical: dict[str, str],
    canonical_order: list[str],
) -> dict[str, str]:
    # Group values by canonical, preserving left-to-right sheet order
    canonical_values: dict[str, list[str]] = {c: [] for c in canonical_order}

    for col_idx, header in enumerate(headers):
        if header not in raw_to_canonical:
            continue
        canonical = raw_to_canonical[header]
        val = row[col_idx] if col_idx < len(row) else ""
        if val:
            canonical_values.setdefault(canonical, []).append(val)

    return {c: merge_same_row_values(canonical_values.get(c, [])) for c in canonical_order}


def iter_sheet_rows(filepath: str, root_for_relpath: str):
    """Yield tagged tuples:
       ("ok", rel_path, file_name, output_sheet, source_sheet, headers, rows, note)
       ("error", rel_path, file_name, scope, message)
    so every failure reaches the Ingest_Log, not just the console."""
    fname = os.path.basename(filepath)
    rel_path = os.path.relpath(filepath, root_for_relpath) if root_for_relpath else fname
    ext = os.path.splitext(fname)[1].lower()

    if ext in CSV_EXTS:
        sep = "\t" if ext in (".tsv", ".txt") else ","
        try:
            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                sample = f.read(4096)
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            sep = dialect.delimiter
        except Exception:
            pass
        try:
            df_raw, enc = read_csv_robust(filepath, sep)
        except Exception as e:
            print(f"  [ERROR] {rel_path}: {e}")
            yield ("error", rel_path, fname, "CSV", str(e)[:300])
            return
        headers, rows, meta, _ = extract_sheet_table(df_raw)
        if headers and rows:
            note = f"encoding={enc}" if enc != "utf-8-sig" else ""
            yield ("ok", rel_path, fname, "CSV", meta.get("sheet_name", "CSV"), headers, rows, note)
        return

    engine = excel_engine(ext)
    try:
        kwargs = {"engine": engine} if engine else {}
        xl = pd.ExcelFile(filepath, **kwargs)
    except Exception as e:
        print(f"  [ERROR] {rel_path}: {e}")
        yield ("error", rel_path, fname, "", str(e)[:300])
        return

    for sheet_name in xl.sheet_names:
        if sheet_name in RESERVED_SHEETS:
            continue
        try:
            # NOTE: no engine kwarg here — the ExcelFile already carries it,
            # and some pandas versions reject engine= alongside an ExcelFile.
            df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=str)
        except Exception as e:
            print(f"  [WARN] {rel_path} / {sheet_name}: {e}")
            yield ("error", rel_path, fname, str(sheet_name), str(e)[:300])
            continue

        headers, rows, meta, _ = extract_sheet_table(df_raw)
        if not headers or not rows:
            continue
        source_sheet = meta.get("sheet_name", sheet_name)
        yield ("ok", rel_path, fname, sheet_name, source_sheet, headers, rows, "")

    xl.close()


def run(
    inputs: list[str],
    mapping_path: str,
    output_path: str,
    mapping_sheet: str = "",
    include_source: bool = False,
    overwrite: bool = False,
    include_previews: bool = False,
) -> None:
    mapping_path = os.path.abspath(mapping_path)
    output_path = os.path.abspath(output_path)
    out_ext = os.path.splitext(output_path)[1].lower()
    log_companion = (
        f"{os.path.splitext(output_path)[0]}_ingest_log.xlsx" if out_ext == ".csv" else None
    )

    if not os.path.isfile(mapping_path):
        sys.exit(f"ERROR: mapping file not found: {mapping_path}")
    for existing in filter(None, [output_path, log_companion]):
        if os.path.exists(existing) and not overwrite:
            sys.exit(f"ERROR: output exists: {existing}\nPass --overwrite to replace.")

    raw_to_canonical, canonical_order = load_mapping(mapping_path, mapping_sheet)
    print(f"Loaded mapping: {len(raw_to_canonical)} raw columns -> {len(canonical_order)} canonical fields")
    for raw, can in raw_to_canonical.items():
        print(f"  {raw!r} -> {can}")

    exclude = {mapping_path, output_path}
    if log_companion:
        exclude.add(log_companion)
    data_files, skipped_previews = collect_input_files(inputs, exclude, include_previews)

    # Use first input path as relpath root (or common parent)
    roots = [os.path.abspath(p) for p in inputs]
    root_for_relpath = roots[0] if len(roots) == 1 and os.path.isdir(roots[0]) else os.path.commonpath(roots)

    master_rows: list[dict] = []
    log_rows: list[dict] = []
    unmapped_headers_seen: set[str] = set()
    mapped_headers_seen: set[str] = set()

    def rel_of(path: str) -> str:
        return os.path.relpath(path, root_for_relpath) if os.path.isdir(root_for_relpath) else os.path.basename(path)

    for p in skipped_previews:
        log_rows.append({
            "file_path": rel_of(p),
            "file_name": os.path.basename(p),
            "output_sheet": "",
            "source_sheet": "",
            "status": "skipped",
            "rows_merged": 0,
            "detail": "preview part file (pass --include-previews to ingest preview workbooks)",
        })
    if skipped_previews:
        print(f"Skipping {len(skipped_previews)} preview part file(s) "
              f"(pass --include-previews to ingest them).")

    if not data_files:
        hint = "\nOnly preview part files were found — pass --include-previews to ingest them." if skipped_previews else ""
        sys.exit(f"ERROR: no input data files found.{hint}")

    print(f"\nIngesting {len(data_files)} file(s)...")

    for i, fp in enumerate(data_files, 1):
        rel = rel_of(fp)
        print(f"[{i}/{len(data_files)}] {rel}")
        file_row_count = 0
        got_any = False
        had_error = False

        for item in iter_sheet_rows(fp, root_for_relpath):
            if item[0] == "error":
                _, rel_path, fname, scope, msg = item
                had_error = True
                log_rows.append({
                    "file_path": rel_path,
                    "file_name": fname,
                    "output_sheet": scope,
                    "source_sheet": "",
                    "status": "error",
                    "rows_merged": 0,
                    "detail": msg,
                })
                continue

            _, rel_path, fname, out_sheet, src_sheet, headers, rows, note = item
            got_any = True
            sheet_mapped = [h for h in headers if h in raw_to_canonical]
            for h in headers:
                if h in raw_to_canonical:
                    mapped_headers_seen.add(h)
                elif h:
                    unmapped_headers_seen.add(h)

            if not sheet_mapped:
                log_rows.append({
                    "file_path": rel_path,
                    "file_name": fname,
                    "output_sheet": out_sheet,
                    "source_sheet": src_sheet,
                    "status": "skipped",
                    "rows_merged": 0,
                    "detail": "no mapped columns in sheet headers",
                })
                continue

            sheet_row_count = 0
            for row in rows:
                merged = build_merged_row(headers, row, raw_to_canonical, canonical_order)
                if not any(merged.values()):
                    continue
                out_row = {c: safe_cell(merged[c]) for c in canonical_order}
                if include_source:
                    out_row["source_file"] = fname
                    out_row["source_sheet"] = src_sheet
                master_rows.append(out_row)
                sheet_row_count += 1
            file_row_count += sheet_row_count

            detail = f"mapped {len(sheet_mapped)}/{len(headers)} headers; scanned {len(rows)} rows"
            if note:
                detail += f"; {note}"
            log_rows.append({
                "file_path": rel_path,
                "file_name": fname,
                "output_sheet": out_sheet,
                "source_sheet": src_sheet,
                "status": "ok",
                "rows_merged": sheet_row_count,
                "detail": detail,
            })

        if not got_any and not had_error:
            log_rows.append({
                "file_path": rel,
                "file_name": os.path.basename(fp),
                "output_sheet": "",
                "source_sheet": "",
                "status": "skipped",
                "rows_merged": 0,
                "detail": "no readable data rows",
            })

        print(f"  -> {file_row_count} master row(s) from this file")

    if not master_rows:
        sys.exit("ERROR: no data rows merged. Check mapping vs source column names (case-sensitive).")

    out_cols = list(canonical_order)
    if include_source:
        out_cols += ["source_file", "source_sheet"]

    master_df = pd.DataFrame(master_rows, columns=out_cols)
    master_df.columns = [safe_cell(c, 255) or f"col_{i + 1}" for i, c in enumerate(master_df.columns)]
    log_df = pd.DataFrame(log_rows)
    summary_df = pd.DataFrame([{
        "input_paths": "; ".join(inputs),
        "mapping_file": mapping_path,
        "output_file": output_path,
        "files_processed": len(data_files),
        "preview_parts_skipped": len(skipped_previews),
        "files_or_sheets_failed": sum(1 for r in log_rows if r["status"] == "error"),
        "total_master_rows": len(master_df),
        "canonical_columns": len(canonical_order),
        "raw_columns_mapped": len(raw_to_canonical),
        "unmapped_headers_seen": len(unmapped_headers_seen),
    }])
    mapping_df = pd.DataFrame([
        {"raw_column_name": k, "suggested_canonical": v}
        for k, v in raw_to_canonical.items()
    ])
    unmapped_df = pd.DataFrame(
        sorted(unmapped_headers_seen), columns=["unmapped_header"]
    )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    def write_log_sheets(writer) -> None:
        sanitize_df(log_df).to_excel(writer, sheet_name="Ingest_Log", index=False)
        sanitize_df(summary_df).to_excel(writer, sheet_name="Run_Summary", index=False)
        sanitize_df(mapping_df).to_excel(writer, sheet_name="Mapping_Used", index=False)
        if not unmapped_df.empty:
            sanitize_df(unmapped_df).to_excel(writer, sheet_name="Unmapped_Headers", index=False)

    master_sheets = []
    if out_ext == ".csv":
        master_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        with pd.ExcelWriter(log_companion, engine="openpyxl") as writer:
            write_log_sheets(writer)
    else:
        # Guard the xlsx hard limit: split a large master across sheets
        chunk = EXCEL_MAX_ROWS - 1
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for k in range(0, len(master_df), chunk):
                name = "Master" if k == 0 else f"Master_{k // chunk + 1}"
                master_sheets.append(name)
                master_df.iloc[k:k + chunk].to_excel(writer, sheet_name=name, index=False)
            write_log_sheets(writer)

    print("\nDone.")
    print(f"  Master rows written : {len(master_df)}")
    if len(master_sheets) > 1:
        print(f"  NOTE: master exceeds the Excel sheet limit — split across "
              f"{len(master_sheets)} sheets: {', '.join(master_sheets)}")
    print(f"  Canonical columns   : {', '.join(canonical_order)}")
    print(f"  Output              : {output_path}")
    if log_companion:
        print(f"  Ingest accounting   : {log_companion}")
    errors = sum(1 for r in log_rows if r["status"] == "error")
    if errors:
        print(f"  Failed files/sheets : {errors} (see Ingest_Log, status=error)")
    if unmapped_headers_seen:
        print(f"  Unmapped headers    : {len(unmapped_headers_seen)} (see Unmapped_Headers sheet)")


def main():
    parser = argparse.ArgumentParser(
        description="Merge ingested files into one master file using raw→canonical column mapping."
    )
    parser.add_argument(
        "--input", "-i", action="append", required=True,
        help="source file or folder path (repeat for multiple)",
    )
    parser.add_argument(
        "--mapping", "-m", required=True,
        help="Excel/CSV with raw_column_name and suggested_canonical",
    )
    parser.add_argument(
        "--mapping-sheet", default="",
        help="sheet name in mapping Excel (default: auto-detect Column_Samples)",
    )
    parser.add_argument(
        "--output", "-o", required=True,
        help="full output path (.xlsx or .csv)",
    )
    parser.add_argument(
        "--include-source", action="store_true",
        help="add source_file and source_sheet columns to master output",
    )
    parser.add_argument(
        "--include-previews", action="store_true",
        help="also ingest *_partNNN.xlsx preview workbooks (skipped by default "
             "to avoid double-ingesting sample rows next to their source files)",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    run(
        inputs=args.input,
        mapping_path=args.mapping,
        output_path=args.output,
        mapping_sheet=args.mapping_sheet,
        include_source=args.include_source,
        overwrite=args.overwrite,
        include_previews=args.include_previews,
    )


if __name__ == "__main__":
    main()
