# mysa_DB_Dhruval_Sept2026 — Leads corpus tooling

Tooling for profiling a large corpus of leads spreadsheets (~5,500 Excel/CSV
files, ~250 GB) before consolidating them into a single queryable database
filterable by products, industry, location, area, city, pincode, business
type, and business level.

The pipeline (all stages hardened and test-covered):

1. `extract_headers_and_samples.py` — structure survey: headers + samples.
2. `suggest_column_merges.py` — profiles column *values* and suggests which
   raw columns mean the same canonical field.
3. `merge_to_master.py` — pours row data into one master file using a
   reviewed raw→canonical mapping.

## `extract_headers_and_samples.py`

Scans a folder tree of Excel/CSV files and writes preview workbooks: one
sheet per source sheet with a metadata block, the detected headers, and up
to 10 sample rows.

### Export policy (enforced in code, not just by convention)

| Boundary | Limit | Enforcement |
|---|---|---|
| Rows read per sheet | 25 | `nrows=` on every reader **plus** defensive truncation and `assert` |
| Sample rows exported per sheet | 10 | truncation + `assert` at write time |
| Headers exported | all columns of the sheet | — |
| Sample values | exported **as-is (unmasked)** by the data owner's decision | treat the preview workbooks as containing real lead data; share them only with people who may see it |
| Formulas in output | impossible | every written value is sanitized: control characters stripped, leading `=` `+` `-` `@` escaped to literal text |
| Output formats | Excel only — no CSVs | run accounting is a `Report` sheet inside part 1 |
| Network activity | none | no network code exists in the script |
| Source files | read-only | never written, moved, or deleted |

### Usage

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# Test drive on a small sub-folder first:
python extract_headers_and_samples.py "/path/to/leads DB/some subfolder" preview.xlsx

# Full run (write output OUTSIDE the data folder):
python extract_headers_and_samples.py "/path/to/leads DB" ~/leads_preview/preview.xlsx
```

Options: `--include-txt` (scan `.txt` too — off by default so notes/readme
files don't leak into the output), `--overwrite`, `--sheets-per-file N`
(default 500 preview sheets per output part), `--limit N` (first N files
only — for test runs).

### Outputs (for output name `preview.xlsx`) — all Excel, no CSVs

- `preview_part001.xlsx`, `preview_part002.xlsx`, … — preview workbooks.
  Each has an `Index` sheet (always the first tab) listing **exactly** the
  sheets that were actually written, with `header_detected = no` flagging
  files where the chosen header row is a best guess and may be data.
  Output is split into parts so workbooks stay small enough for Excel and
  a late failure can never lose the whole run.
- `preview_part001.xlsx` additionally carries the **`Report` sheet** (second
  tab): one row per source file/sheet with `ok / skipped / error` and the
  reason (e.g. `encoding=cp1252`, `missing engine: pyxlsb`, `.txt excluded
  by default`). This is the complete accounting of the run — check it to
  see nothing went missing. It is written even if the run is interrupted.

### What is hardened (vs. the original draft script)

1. **No crash on dirty data**: control characters used to raise
   `IllegalCharacterError` during the final write, losing the whole run and
   leaving a partial file whose Index claimed completeness. Values are now
   sanitized before writing, sheets are written incrementally in parts, and
   each Index reflects reality.
2. **No formula injection**: cell values like `=HYPERLINK(...)` from source
   files used to become live formulas in the output workbook (a data-exfil
   vector on open). All formula-leading values are now stored as literal
   text — verified by test.
3. **Encoding fallback** for CSVs (`utf-8-sig → utf-8 → cp1252 → latin-1`);
   non-UTF-8 files were previously skipped silently.
4. **Header detection** scores candidate rows (text-ness, known header
   keywords, uniqueness) instead of accepting the first dense row, and
   flags low-confidence detections instead of presenting data as headers.
5. **Junk excluded**: `~$` lock files, hidden files/folders, and (by
   default) `.txt` files.
6. **Accounting**: the `Report` sheet in part 1, missing-engine detection
   with a single clear warning, overwrite protection on outputs.

Note: sample values are deliberately **not masked** — the data owner chose
full-fidelity previews. If a preview will be shared beyond trusted people,
that decision should be revisited.

### Verify what you run

The audit applies to an exact file. Pin it:

```bash
shasum -a 256 extract_headers_and_samples.py
```

Compare against the hash of the reviewed commit before running on the real
corpus. For a guarantee that is independent of any code review, run inside
a no-network container with the corpus mounted read-only:

```bash
docker run --rm --network none \
  -v "/path/to/leads DB":/data:ro \
  -v "$HOME/leads_preview":/out \
  -v "$PWD":/app -w /app python:3.12-slim \
  sh -c "pip install -r requirements.txt -q && python extract_headers_and_samples.py /data /out/preview.xlsx"
```

With `--network none` and `:ro`, exfiltration and source modification are
impossible regardless of what any script does.

## `suggest_column_merges.py` (stage 2)

Scans a folder (raw files or preview workbooks), reads at most ~21 rows per
sheet, and classifies every raw column by its actual values (emails, phones,
URLs, pincodes, addresses, company suffixes, person names with Indian
honorifics) combined with header-name hints. Handles the ambiguous bare
`Name` column per sheet. Output workbook: `Merge_Sheet_Suggestion`,
`Column_Samples` (the mapping input for stage 3 — **review this by hand
before merging**), `Per_Sheet_Mapping`, `Column_Profiles`,
`Ambiguous_Name_Like`, `Skipped_Files` (every unreadable file/sheet with
the reason), `Run_Summary`.

```bash
python suggest_column_merges.py "/path/to/Category-1" \
    --output "Category-1/column_merge_suggestions.xlsx" --overwrite
```

Hardened: all output cells sanitized (no formula injection, no
control-character crash), CSV encoding fallback, failures recorded in the
workbook instead of console-only.

## `merge_to_master.py` (stage 3)

Ingests **all rows** of the input files and merges them into one master
table using the reviewed mapping. The mapping is a **column allowlist**:
only mapped columns' data enters the master; unmapped columns are listed by
name only in `Unmapped_Headers`. Same-row rule for multiple columns mapping
to one field: first non-empty wins, case-insensitive duplicates dropped,
different values joined with `" | "`. Rows are never merged across files.

```bash
python merge_to_master.py --input "/path/to/Category-1" \
    --mapping "Category-1/column_merge_suggestions.xlsx" \
    --output "Category-1/category1_master_merged.xlsx" \
    --include-source --overwrite
```

Hardened: every output sheet sanitized, CSV encoding fallback, every failed
file/sheet logged in `Ingest_Log` with `status=error`, the Excel row limit
(1,048,576/sheet) guarded by splitting into `Master`, `Master_2`, …,
preview `*_partNNN.xlsx` files skipped by default (`--include-previews` to
ingest them), and the output + mapping files excluded from ingestion.

Operational notes: run per category (everything is held in RAM before the
final write); after each run review `Ingest_Log` rows with
`status=skipped/error` and the `Unmapped_Headers` sheet — that is your
completeness check. A `.csv` output writes the accounting to
`<output>_ingest_log.xlsx` alongside it.

### Tests

```bash
python tests/smoke_test.py            # stage 1 (extractor)
python tests/pipeline_smoke_test.py   # stages 2-3 (suggest + merge)
```

Generates synthetic dummy files (banner rows, headerless data, cp1252
encoding, formula injection, control characters, lock files, stray notes)
in a temp dir, runs the extractor on them, and asserts the export policy
and every hardening fix above — including that values arrive unmasked and
that no CSV files are produced. The run must end with "All checks passed."

## Roadmap

1. **Done — structure survey** (this tool): know every header variant and
   see samples of what columns actually contain.
2. **In repo — canonical schema tooling**: `suggest_column_merges.py` proposes
   the raw→canonical mapping from data evidence; you review and edit it.
3. **In repo — consolidation (v1)**: `merge_to_master.py` merges mapped
   columns into per-category master files with full accounting.
4. **Next — database + query layer**: load the per-category masters into an
   indexed database with dedupe, then filtered extraction for campaigns
   across the eight target filter dimensions.
