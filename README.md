# mysa_DB_Dhruval_Sept2026 — Leads corpus tooling

Tooling for profiling a large corpus of leads spreadsheets (~5,500 Excel/CSV
files, ~250 GB) before consolidating them into a single queryable database
filterable by products, industry, location, area, city, pincode, business
type, and business level.

This repository currently contains **step 1**: a hardened structure-survey
tool that inventories every sheet's column headers plus a small, masked
sample of its data — so the unified schema can be designed from evidence.

## `extract_headers_and_samples.py`

Scans a folder tree of Excel/CSV files and writes preview workbooks: one
sheet per source sheet with a metadata block, the detected headers, and up
to 10 sanitized sample rows.

### Export policy (enforced in code, not just by convention)

| Boundary | Limit | Enforcement |
|---|---|---|
| Rows read per sheet | 25 | `nrows=` on every reader **plus** defensive truncation and `assert` |
| Sample rows exported per sheet | 10 | truncation + `assert` at write time |
| Headers exported | all columns of the sheet | headers only, no cell data |
| Sensitive values in samples | masked by default | phones / 8+ digit runs → `98******10`, emails → `o***@domain.com`, GSTIN/PAN partially starred; 6-digit pincodes stay readable |
| Formulas in output | impossible | every written value is sanitized: control characters stripped, leading `=` `+` `-` `@` escaped to literal text |
| Network activity | none | no network code exists in the script |
| Source files | read-only | never written, moved, or deleted |

Masking exists because the preview workbook is intended to be shared (e.g.
with a developer): structure stays visible, real contact data does not
travel. `--no-mask` disables it — only for internal use.

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
files don't leak into the output), `--no-mask`, `--overwrite`,
`--sheets-per-file N` (default 500 preview sheets per output part),
`--limit N` (first N files only — for test runs).

### Outputs (for output name `preview.xlsx`)

- `preview_part001.xlsx`, `preview_part002.xlsx`, … — preview workbooks.
  Each has an `Index` sheet (always the first tab) listing **exactly** the
  sheets that were actually written, with `header_detected = no` flagging
  files where the chosen header row is a best guess and may be data.
  Output is split into parts so workbooks stay small enough for Excel and
  a late failure can never lose the whole run.
- `preview_report.csv` — one row per source file/sheet with
  `ok / skipped / error` and the reason (e.g. `encoding=cp1252`,
  `missing engine: pyxlsb`, `.txt excluded by default`). This is the
  complete accounting of the run — check it to see nothing went missing.

### What was hardened (vs. the original draft script)

1. **No crash on dirty data**: control characters used to raise
   `IllegalCharacterError` during the final write, losing the whole run and
   leaving a partial file whose Index claimed completeness. Values are now
   sanitized before writing, sheets are written incrementally in parts, and
   each Index reflects reality.
2. **No formula injection**: cell values like `=HYPERLINK(...)` from source
   files used to become live formulas in the output workbook (a data-exfil
   vector on open). All formula-leading values are now stored as literal
   text — verified by test.
3. **Masking** of phones, emails, GSTIN/PAN in sample rows (default on).
4. **Encoding fallback** for CSVs (`utf-8-sig → utf-8 → cp1252 → latin-1`);
   non-UTF-8 files were previously skipped silently.
5. **Header detection** scores candidate rows (text-ness, known header
   keywords, uniqueness) instead of accepting the first dense row, and
   flags low-confidence detections instead of presenting data as headers.
6. **Junk excluded**: `~$` lock files, hidden files/folders, and (by
   default) `.txt` files.
7. **Accounting**: per-file report CSV, missing-engine detection with a
   single clear warning, overwrite protection on outputs.

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

### Tests

```bash
python tests/smoke_test.py
```

Generates synthetic dummy files (banner rows, headerless data, cp1252
encoding, formula injection, control characters, lock files, stray notes)
in a temp dir, runs the extractor on them, and asserts the export policy
and every hardening fix above. 23 checks; all must pass.

## Roadmap

1. **Done — structure survey** (this tool): know every header variant and
   see masked samples of what columns actually contain.
2. **Next — canonical schema**: map raw header variants onto canonical
   fields (company, owner, mobile, email, address, area, city, district,
   state, country, pincode, products, industry, business type, business
   level, turnover, …) using the survey output as evidence.
3. **Then — consolidation ETL**: stream all 5,500 files into a single
   indexed database (one normalized `leads` table + source lineage), with
   dedupe and cleaning.
4. **Finally — query layer**: filtered extraction for sales/marketing
   campaigns across the eight target filter dimensions.
