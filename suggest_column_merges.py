#!/usr/bin/env python3
"""
suggest_column_merges.py  (hardened)
------------------------------------
Read sheet data from Excel/CSV preview workbooks, profile each **raw column**
(case-sensitive, as extracted upstream), and suggest which columns are
semantically the same based on **actual cell values** — not just header
spelling/case.

Outputs an Excel workbook with:
  - Column_Samples        : each unique raw column + 5 sample values for verification
  - Merge_Sheet_Suggestion: canonical groups  e.g.  person_name: Name, name, fullName, FullName
  - Per_Sheet_Mapping     : per sheet, which raw column maps to which canonical field
                            (handles cases where "Name" = person in one sheet but company in another)
  - Column_Profiles       : detailed content-type scores per raw column
  - Skipped_Files         : every file/sheet that could not be read, with the reason
  - Run_Summary           : scan statistics

Hardening (vs. the original draft):
  * Every value written to the output is sanitized: control characters are
    stripped (fixes the IllegalCharacterError crash at the final write) and
    formula-leading values (= + - @) are stored as literal text, so the
    output can never contain live formulas.
  * CSV files are read with encoding fallback (utf-8-sig -> utf-8 -> cp1252
    -> latin-1) instead of being skipped when not UTF-8.
  * Files/sheets that fail to read are recorded in the Skipped_Files sheet,
    not just printed to the console.
  * Reads stay bounded: at most ~21 rows per sheet are ever read.
  * Sources are opened read-only and never modified; no network activity.

Usage:
    python suggest_column_merges.py [folder] [options]

Examples:
    python suggest_column_merges.py "Category-1"
    python suggest_column_merges.py "Category-1" --output "Unique merged output columns/Category-1/column_merge_suggestions.xlsx" --overwrite
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import warnings
from collections import Counter, defaultdict
from dataclasses import dataclass, field

import pandas as pd

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FOLDER = os.path.join(SCRIPT_DIR, "Category-1")
DEFAULT_OUTPUT = os.path.join(
    SCRIPT_DIR,
    "Unique merged output columns",
    "Category-1",
    "column_merge_suggestions.xlsx",
)

RESERVED_SHEETS = {"Index", "Report"}
PREVIEW_HEADER_ROW = 8
PREVIEW_META_ROWS = 8
MAX_DATA_ROWS = 12          # rows read per sheet for profiling
SAMPLE_VALUES_OUT = 5       # sample values shown in output
MAX_SAMPLE_STORE = 30       # max values kept per raw column globally

EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".xlam", ".ods"}
CSV_EXTS = {".csv", ".tsv", ".txt"}

SKIP_SUFFIXES = ("_unique_columns.xlsx", "_merge_suggestions.xlsx")
SKIP_NAMES = {
    "column_merge_suggestions.xlsx",
    "category1_unique_columns_merged.xlsx",
    "all_unique_columns_merged.xlsx",
}

CSV_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]

# ── Output sanitization ──────────────────────────────────────────────────────
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
    """Strip illegal characters and neutralize formula interpretation."""
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


# ── Content detection patterns ───────────────────────────────────────────────
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
PHONE_RE = re.compile(r"^[\d\s\-\+\(\)\./]{7,}$")
URL_RE = re.compile(r"(?:https?://|www\.)", re.I)
IMAGE_URL_RE = re.compile(r"\.(?:jpg|jpeg|png|gif|webp|bmp)(?:\?|$)", re.I)
MEMBER_LINK_RE = re.compile(r"memberdetails|encryptedmemberid|/member/", re.I)
PINCODE_RE = re.compile(r"\b\d{6}\b")
ADDRESS_WORDS = re.compile(
    r"\b(street|st\.|road|rd\.|lane|ln\.|sector|plot|floor|building|bldg|near|nagar|colony|"
    r"pincode|pin|area|district|city|state|highway|market|shop|flat|house|h\.?\s*no)\b",
    re.I,
)
COMPANY_SUFFIX_RE = re.compile(
    r"\b(pvt\.?\s*ltd\.?|ltd\.?|llp|inc\.?|corp\.?|corporation|limited|enterprises|"
    r"industries|solutions|traders|trading|company|co\.|group|services|consulting|"
    r"international|global|works|factory|mills|laboratory|lab|studio|atelier)\b",
    re.I,
)
PERSON_PREFIX_RE = re.compile(r"^(mr\.?|mrs\.?|ms\.?|dr\.?|ca\.?|shri\.?|smt\.?|prof\.?)\s+", re.I)
PROFESSION_CATEGORY_RE = re.compile(r">|\|")
PERSON_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z.'\-]*(?:\s+[A-Za-z][A-Za-z.'\-]*){0,4}$")

CANONICAL_FIELDS = [
    "person_name",
    "company_name",
    "phone",
    "email",
    "address",
    "profession",
    "designation",
    "chapter",
    "photo_url",
    "member_link",
    "website",
    "city",
    "state",
    "pincode",
    "category",
    "serial_number",
    "other_text",
    "unknown",
]

HEADER_TO_CANONICAL = [
    (re.compile(r"^(full\s*name|fullname|member\s*name|contact\s*name|person\s*name|name\s*of\s*person)$", re.I), "person_name", 0.85),
    (re.compile(r"^(name|naam)$", re.I), "person_name", 0.55),  # ambiguous — lowered, data decides
    (re.compile(r"^(company(\s*name)?|firm(\s*name)?|business(\s*name)?|enterprise\s*name|factory\s*name|name\s*of\s*(the\s+)?(firm|company|industr))$", re.I), "company_name", 0.90),
    (re.compile(r"^(co\.?\s*name|organisation|organization|enterprise)$", re.I), "company_name", 0.80),
    (re.compile(r"^(mobile|mob\.?|cell(\s*no\.?)?|phone|tel(?:ephone)?|contact\s*no\.?|billing_phone|msisdn)$", re.I), "phone", 0.88),
    (re.compile(r"^(email|e[\-\s]?mail|mail(\s*id)?|emailid|email\s*id[\-\s]?\d*)$", re.I), "email", 0.90),
    (re.compile(r"^(address|addr\.?|add\.|communication\s*address|registered[\s_]*office[\s_]*address|office\s*address|factory\s*address)$", re.I), "address", 0.88),
    (re.compile(r"^(profession|industry|business\s*category|sector|segment)$", re.I), "profession", 0.85),
    (re.compile(r"^(designation|title|role|position|job\s*title)$", re.I), "designation", 0.85),
    (re.compile(r"^(chapter|region|zone|branch|group\s*name)$", re.I), "chapter", 0.80),
    (re.compile(r"^(photo\s*url|photourl|photo|image(\s*url)?|picture|avatar)$", re.I), "photo_url", 0.88),
    (re.compile(r"^(member\s*link|memberlink|profile\s*link|link|member\s*url|portfolio\s*link)$", re.I), "member_link", 0.85),
    (re.compile(r"^(website|web|url|site)$", re.I), "website", 0.75),
    (re.compile(r"^(city|town|district|dist\.?)$", re.I), "city", 0.82),
    (re.compile(r"^(state|province)$", re.I), "state", 0.82),
    (re.compile(r"^(pin\s*code|pincode|zip\s*code|postal\s*code|zip)$", re.I), "pincode", 0.88),
    (re.compile(r"^(category|type|subcategory)$", re.I), "category", 0.75),
    (re.compile(r"^(sr\.?\s*no\.?|s\.?\s*no\.?|serial(\s*no\.?)?|sl\.?\s*no\.?)$", re.I), "serial_number", 0.85),
]

DISPLAY_NAMES = {
    "person_name": "Person Name",
    "company_name": "Company",
    "phone": "Mobile / Phone",
    "email": "Email",
    "address": "Address",
    "profession": "Profession / Industry",
    "designation": "Designation",
    "chapter": "Chapter / Region",
    "photo_url": "Photo URL",
    "member_link": "Member Link",
    "website": "Website",
    "city": "City",
    "state": "State",
    "pincode": "Pincode",
    "category": "Category",
    "serial_number": "Serial Number",
    "other_text": "Other Text",
    "unknown": "Unknown",
}


def normalize_value(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


def excel_engine(ext: str):
    if ext == ".xls":
        return "xlrd"
    if ext == ".xlsb":
        return "pyxlsb"
    if ext == ".ods":
        return "odf"
    return None


def is_supported(fname: str) -> bool:
    ext = os.path.splitext(fname)[1].lower()
    return ext in EXCEL_EXTS or ext in CSV_EXTS


def should_skip(fname: str) -> bool:
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


def extract_sheet_data(df_raw: pd.DataFrame) -> tuple[list[str], list[list[str]], dict, int]:
    preview_meta = {}
    if is_preview_format(df_raw):
        preview_meta = parse_preview_metadata(df_raw)
        if len(df_raw) <= PREVIEW_HEADER_ROW:
            return [], [], preview_meta, 0
        hrow = PREVIEW_HEADER_ROW
    else:
        hrow = find_header_row(df_raw)

    headers = [cell_str(v) for v in df_raw.iloc[hrow].tolist()]
    # trim trailing empty headers
    while headers and not headers[-1]:
        headers.pop()

    rows: list[list[str]] = []
    for r in range(hrow + 1, min(hrow + 1 + MAX_DATA_ROWS, len(df_raw))):
        vals = []
        for c in range(len(headers)):
            vals.append(cell_str(df_raw.iloc[r, c]) if c < df_raw.shape[1] else "")
        if any(vals):
            rows.append(vals)

    return headers, rows, preview_meta, hrow + 1


def score_value_content(value: str) -> dict[str, float]:
    """Return content-type scores for a single cell value (0..1)."""
    scores = {k: 0.0 for k in CANONICAL_FIELDS}
    if not value:
        return scores

    s = value.strip()
    low = s.lower()

    if EMAIL_RE.match(s):
        scores["email"] = 1.0
        return scores

    digits = sum(ch.isdigit() for ch in s)
    if PHONE_RE.match(s) and digits >= 7:
        scores["phone"] = min(1.0, 0.6 + digits / 15.0)
        return scores

    if MEMBER_LINK_RE.search(s):
        scores["member_link"] = 1.0
        return scores

    if URL_RE.search(s) or IMAGE_URL_RE.search(s):
        if IMAGE_URL_RE.search(s) or "photo" in low or "image" in low or "cms" in low:
            scores["photo_url"] = 0.95
        else:
            scores["website"] = 0.85
        return scores

    if PINCODE_RE.search(s) and len(s) <= 12:
        scores["pincode"] = 0.9
        return scores

    if PROFESSION_CATEGORY_RE.search(s):
        scores["profession"] = 0.9
        return scores

    if COMPANY_SUFFIX_RE.search(s):
        scores["company_name"] = 0.85
        return scores

    if len(s) >= 20 and (PINCODE_RE.search(s) or ADDRESS_WORDS.search(s) or s.count(",") >= 2):
        scores["address"] = min(1.0, 0.5 + len(s) / 120.0)
        return scores

    cleaned = PERSON_PREFIX_RE.sub("", s).strip()
    words = cleaned.split()
    if 1 <= len(words) <= 5 and PERSON_NAME_RE.match(cleaned):
        if not COMPANY_SUFFIX_RE.search(cleaned):
            scores["person_name"] = 0.75 if len(words) >= 2 else 0.55
            return scores

    if len(s) >= 8 and not digits:
        if COMPANY_SUFFIX_RE.search(s):
            scores["company_name"] = 0.8
        elif len(words) >= 2:
            scores["company_name"] = 0.45
            scores["other_text"] = 0.35
        else:
            scores["other_text"] = 0.5
        return scores

    if len(s) <= 40 and not digits:
        scores["designation"] = 0.35
        scores["chapter"] = 0.30
        scores["other_text"] = 0.35

    if not any(scores.values()):
        scores["unknown"] = 0.2
    return scores


def profile_column_values(values: list[str]) -> dict[str, float]:
    """Aggregate content scores across sample values."""
    totals = Counter()
    n = 0
    for v in values:
        if not v:
            continue
        n += 1
        for k, sc in score_value_content(v).items():
            if sc > 0:
                totals[k] += sc
    if n == 0:
        return {k: 0.0 for k in CANONICAL_FIELDS}
    return {k: totals[k] / n for k in CANONICAL_FIELDS}


def header_canonical_hint(header: str) -> tuple[str, float]:
    h = header.strip()
    for pattern, canonical, conf in HEADER_TO_CANONICAL:
        if pattern.match(h):
            return canonical, conf
    norm = re.sub(r"[_\-\./\\]+", " ", h.lower()).strip()
    for pattern, canonical, conf in HEADER_TO_CANONICAL:
        if pattern.match(norm):
            return canonical, conf * 0.95
    return "unknown", 0.0


def choose_canonical(content_scores: dict[str, float], header: str) -> tuple[str, float, str]:
    hint, hint_conf = header_canonical_hint(header)

    # Special disambiguation: bare "Name" / "name"
    if re.match(r"^name$", header.strip(), re.I):
        person_s = content_scores.get("person_name", 0)
        company_s = content_scores.get("company_name", 0)
        if person_s >= company_s + 0.15:
            return "person_name", min(0.98, 0.55 + person_s * 0.4), "data: person-name pattern dominates"
        if company_s >= person_s + 0.15:
            return "company_name", min(0.98, 0.55 + company_s * 0.4), "data: company-name pattern dominates"
        if person_s > 0:
            return "person_name", 0.65, "header: ambiguous Name; slight person-name lean"
        return "unknown", 0.4, "header: ambiguous Name; insufficient data"

    combined = dict(content_scores)
    if hint != "unknown":
        combined[hint] = combined.get(hint, 0) + hint_conf

    best = max(CANONICAL_FIELDS, key=lambda k: combined.get(k, 0))
    best_score = combined.get(best, 0)
    if best_score < 0.25:
        if hint != "unknown":
            return hint, hint_conf * 0.7, f"header hint: {hint}"
        return "unknown", 0.2, "low confidence"

    reasons = []
    if content_scores.get(best, 0) >= 0.35:
        reasons.append(f"data: {best}={content_scores.get(best, 0):.2f}")
    if hint == best and hint_conf >= 0.5:
        reasons.append(f"header: {hint}")
    elif hint != "unknown" and hint != best:
        reasons.append(f"header suggested {hint}, data chose {best}")

    conf = min(0.99, 0.35 + best_score * 0.55 + (0.1 if hint == best else 0))
    return best, conf, "; ".join(reasons) or f"score {best}={best_score:.2f}"


def value_overlap_ratio(a_vals: list[str], b_vals: list[str]) -> float:
    if not a_vals or not b_vals:
        return 0.0
    pairs = min(len(a_vals), len(b_vals))
    if pairs == 0:
        return 0.0
    matches = sum(
        1 for i in range(pairs)
        if a_vals[i] and b_vals[i] and normalize_value(a_vals[i]) == normalize_value(b_vals[i])
    )
    return matches / pairs


def find_source_files(root_folder: str) -> list[str]:
    files = []
    for dirpath, dirnames, filenames in os.walk(root_folder):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d != "__MACOSX"]
        for fname in sorted(filenames):
            if not is_supported(fname) or should_skip(fname):
                continue
            files.append(os.path.join(dirpath, fname))
    return files


@dataclass
class SheetColumnObs:
    file_path: str
    file_name: str
    output_sheet: str
    source_sheet_name: str
    raw_column: str
    col_index: int
    values: list[str] = field(default_factory=list)
    content_scores: dict[str, float] = field(default_factory=dict)
    canonical: str = "unknown"
    confidence: float = 0.0
    reason: str = ""


def read_workbook_sheets(filepath: str, input_folder: str) -> tuple[list[SheetColumnObs], list[dict]]:
    """Return (observations, failures). Failures are recorded, not just printed."""
    rel_path = os.path.relpath(filepath, input_folder)
    fname = os.path.basename(filepath)
    ext = os.path.splitext(fname)[1].lower()
    observations: list[SheetColumnObs] = []
    failures: list[dict] = []

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
            df_raw, _enc = read_csv_robust(
                filepath, sep, nrows=PREVIEW_HEADER_ROW + 1 + MAX_DATA_ROWS
            )
        except Exception as e:
            print(f"  [ERROR] {rel_path}: {e}")
            failures.append({"file_path": rel_path, "scope": "file", "reason": str(e)[:300]})
            return observations, failures
        headers, rows, meta, _ = extract_sheet_data(df_raw)
        source_sheet = meta.get("sheet_name", "CSV")
        for ci, header in enumerate(headers):
            if not header:
                continue
            vals = [row[ci] if ci < len(row) else "" for row in rows]
            observations.append(SheetColumnObs(rel_path, fname, "CSV", source_sheet, header, ci, vals))
        return observations, failures

    engine = excel_engine(ext)
    try:
        kwargs = {"engine": engine} if engine else {}
        xl = pd.ExcelFile(filepath, **kwargs)
    except Exception as e:
        print(f"  [ERROR] {rel_path}: {e}")
        failures.append({"file_path": rel_path, "scope": "file", "reason": str(e)[:300]})
        return observations, failures

    for sheet_name in xl.sheet_names:
        if sheet_name in RESERVED_SHEETS:
            continue
        try:
            # NOTE: no engine kwarg here — the ExcelFile already carries it,
            # and some pandas versions reject engine= alongside an ExcelFile.
            df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=str,
                                   nrows=PREVIEW_HEADER_ROW + 1 + MAX_DATA_ROWS)
        except Exception as e:
            print(f"  [WARN] {rel_path} / {sheet_name}: {e}")
            failures.append({"file_path": rel_path, "scope": str(sheet_name), "reason": str(e)[:300]})
            continue

        headers, rows, meta, _ = extract_sheet_data(df_raw)
        if not headers:
            continue
        source_sheet = meta.get("sheet_name", sheet_name)
        for ci, header in enumerate(headers):
            if not header:
                continue
            vals = [row[ci] if ci < len(row) else "" for row in rows]
            observations.append(
                SheetColumnObs(rel_path, fname, sheet_name, source_sheet, header, ci, vals)
            )
    xl.close()
    return observations, failures


def build_global_samples(observations: list[SheetColumnObs]) -> dict[str, list[str]]:
    samples: dict[str, list[str]] = defaultdict(list)
    for obs in observations:
        for v in obs.values:
            if not v:
                continue
            bucket = samples[obs.raw_column]
            if len(bucket) >= MAX_SAMPLE_STORE:
                break
            if v not in bucket:
                bucket.append(v)
    return samples


def refine_with_pairwise_similarity(observations: list[SheetColumnObs]) -> None:
    """Within each sheet, if two columns have highly overlapping values, align canonical."""
    by_sheet: dict[tuple, list[SheetColumnObs]] = defaultdict(list)
    for obs in observations:
        by_sheet[(obs.file_path, obs.output_sheet)].append(obs)

    for sheet_obs in by_sheet.values():
        for i, a in enumerate(sheet_obs):
            for b in sheet_obs[i + 1:]:
                overlap = value_overlap_ratio(a.values, b.values)
                if overlap < 0.75:
                    continue
                # Prefer higher-confidence; tie-break by header hint for person vs company
                if a.canonical != b.canonical:
                    if a.confidence >= b.confidence:
                        b.canonical = a.canonical
                        b.confidence = max(b.confidence, a.confidence * 0.95)
                        b.reason += f"; paired overlap {overlap:.0%} with '{a.raw_column}'"
                    else:
                        a.canonical = b.canonical
                        a.confidence = max(a.confidence, b.confidence * 0.95)
                        a.reason += f"; paired overlap {overlap:.0%} with '{b.raw_column}'"


def run(folder: str, output_path: str, overwrite: bool = False) -> None:
    folder = os.path.abspath(folder)
    output_path = os.path.abspath(output_path)

    if not os.path.isdir(folder):
        sys.exit(f"ERROR: folder not found: {folder}")
    if os.path.exists(output_path) and not overwrite:
        sys.exit(f"ERROR: output exists: {output_path}\nPass --overwrite to replace.")

    files = [f for f in find_source_files(folder) if os.path.abspath(f) != output_path]
    if not files:
        sys.exit(f"ERROR: no supported files in {folder}")

    print(f"Scanning {len(files)} file(s) in: {folder}")
    all_observations: list[SheetColumnObs] = []
    all_failures: list[dict] = []

    for i, fp in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {os.path.relpath(fp, folder)}")
        obs, failures = read_workbook_sheets(fp, folder)
        all_observations.extend(obs)
        all_failures.extend(failures)
        print(f"  -> {len(obs)} column observations")

    if not all_observations:
        sys.exit("ERROR: no column data extracted.")

    # Profile each sheet-column observation
    for obs in all_observations:
        obs.content_scores = profile_column_values(obs.values)
        obs.canonical, obs.confidence, obs.reason = choose_canonical(obs.content_scores, obs.raw_column)

    refine_with_pairwise_similarity(all_observations)

    global_samples = build_global_samples(all_observations)

    # ── Column_Samples ───────────────────────────────────────────────────────
    raw_columns = sorted({obs.raw_column for obs in all_observations}, key=str.lower)
    sample_rows = []
    for col in raw_columns:
        col_obs = [o for o in all_observations if o.raw_column == col]
        all_scores = Counter()
        for o in col_obs:
            for k, v in o.content_scores.items():
                all_scores[k] += v
        dominant = all_scores.most_common(1)[0][0] if all_scores else "unknown"
        canonical_votes = Counter(o.canonical for o in col_obs)
        suggested = canonical_votes.most_common(1)[0][0]
        samples = global_samples.get(col, [])[:SAMPLE_VALUES_OUT]
        row = {
            "raw_column_name": col,
            "suggested_canonical": suggested,
            "display_group": DISPLAY_NAMES.get(suggested, suggested),
            "dominant_content_type": dominant,
            "sheet_occurrences": len({(o.file_path, o.output_sheet) for o in col_obs}),
            "file_occurrences": len({o.file_path for o in col_obs}),
            "avg_confidence": round(sum(o.confidence for o in col_obs) / len(col_obs), 3),
        }
        for j in range(SAMPLE_VALUES_OUT):
            row[f"sample_{j + 1}"] = samples[j] if j < len(samples) else ""
        sample_rows.append(row)
    column_samples_df = pd.DataFrame(sample_rows)

    # ── Merge_Sheet_Suggestion ───────────────────────────────────────────────
    merge_groups: dict[str, list[str]] = defaultdict(list)
    merge_notes: dict[str, list[str]] = defaultdict(list)

    for col in raw_columns:
        col_obs = [o for o in all_observations if o.raw_column == col]
        canonical_votes = Counter(o.canonical for o in col_obs)
        suggested, vote_count = canonical_votes.most_common(1)[0]
        merge_groups[suggested].append(col)
        if len(canonical_votes) > 1:
            alt = "; ".join(f"{k}({v})" for k, v in canonical_votes.most_common()[1:3])
            merge_notes[suggested].append(f"{col}: also seen as {alt} in some sheets")

    merge_rows = []
    for canonical in CANONICAL_FIELDS:
        cols = merge_groups.get(canonical, [])
        if not cols:
            continue
        cols_sorted = sorted(cols, key=str.lower)
        basis_parts = []
        for c in cols_sorted:
            col_obs = [o for o in all_observations if o.raw_column == c]
            dom = Counter(o.canonical for o in col_obs).most_common(1)[0]
            samples = global_samples.get(c, [])[:2]
            ex = samples[0][:60] if samples else ""
            basis_parts.append(f"{c} [{dom[0]}, n={dom[1]}] e.g. {ex!r}")

        merge_rows.append({
            "merge_line": f"{DISPLAY_NAMES.get(canonical, canonical)}: {', '.join(cols_sorted)}",
            "canonical_field": canonical,
            "display_name": DISPLAY_NAMES.get(canonical, canonical),
            "raw_columns": ", ".join(cols_sorted),
            "raw_column_count": len(cols_sorted),
            "merge_basis": "data content profiling + header hints + within-sheet value overlap",
            "column_details": " | ".join(basis_parts),
            "conflicts_or_notes": " ; ".join(merge_notes.get(canonical, [])) or "",
        })
    merge_suggestion_df = pd.DataFrame(merge_rows)

    # ── Per_Sheet_Mapping ────────────────────────────────────────────────────
    mapping_rows = []
    for obs in sorted(all_observations, key=lambda o: (o.file_path, o.output_sheet, o.col_index)):
        top_data = sorted(obs.content_scores.items(), key=lambda x: -x[1])[:3]
        data_hint = ", ".join(f"{k}={v:.2f}" for k, v in top_data if v > 0)
        mapping_rows.append({
            "file_path": obs.file_path,
            "file_name": obs.file_name,
            "output_sheet": obs.output_sheet,
            "source_sheet_name": obs.source_sheet_name,
            "raw_column": obs.raw_column,
            "suggested_canonical": obs.canonical,
            "display_group": DISPLAY_NAMES.get(obs.canonical, obs.canonical),
            "confidence": round(obs.confidence, 3),
            "detection_reason": obs.reason,
            "content_type_scores": data_hint,
            "sample_value_1": obs.values[0] if obs.values else "",
            "sample_value_2": obs.values[1] if len(obs.values) > 1 else "",
        })
    per_sheet_df = pd.DataFrame(mapping_rows)

    # ── Column_Profiles (raw column aggregate) ───────────────────────────────
    profile_rows = []
    for col in raw_columns:
        col_obs = [o for o in all_observations if o.raw_column == col]
        agg_scores = Counter()
        for o in col_obs:
            for k, v in o.content_scores.items():
                agg_scores[k] += v
        n = len(col_obs) or 1
        avg_scores = {k: round(agg_scores[k] / n, 3) for k in CANONICAL_FIELDS}
        canonical_votes = Counter(o.canonical for o in col_obs)
        profile_rows.append({
            "raw_column_name": col,
            "suggested_canonical": canonical_votes.most_common(1)[0][0],
            "header_hint": header_canonical_hint(col)[0],
            "header_hint_confidence": header_canonical_hint(col)[1],
            **{f"avg_{k}": avg_scores[k] for k in CANONICAL_FIELDS},
            "canonical_vote_breakdown": ", ".join(f"{k}:{v}" for k, v in canonical_votes.most_common()),
        })
    profiles_df = pd.DataFrame(profile_rows)

    # ── Ambiguous columns report (Name-like conflicts) ───────────────────────
    ambiguous_rows = []
    by_sheet = defaultdict(list)
    for obs in all_observations:
        by_sheet[(obs.file_path, obs.output_sheet, obs.source_sheet_name)].append(obs)

    for key, sheet_obs in by_sheet.items():
        fp, out_sheet, src_sheet = key
        name_like = [o for o in sheet_obs if re.match(r"^(name|fullname|company)$", o.raw_column.strip(), re.I)]
        if len(name_like) < 2:
            continue
        desc = " | ".join(
            f"{o.raw_column}->{o.canonical} ({o.confidence:.2f}) sample={o.values[0][:40]!r}"
            for o in name_like
        )
        ambiguous_rows.append({
            "file_path": fp,
            "output_sheet": out_sheet,
            "source_sheet_name": src_sheet,
            "name_like_columns": ", ".join(o.raw_column for o in name_like),
            "sheet_level_mapping": desc,
            "note": "Same header text may map differently per sheet — use this row to decide merges",
        })
    ambiguous_df = pd.DataFrame(ambiguous_rows)

    # ── Skipped files / Run summary ──────────────────────────────────────────
    skipped_df = pd.DataFrame(all_failures, columns=["file_path", "scope", "reason"])
    summary_df = pd.DataFrame([{
        "input_folder": folder,
        "files_scanned": len(files),
        "files_or_sheets_failed": len(all_failures),
        "sheet_column_observations": len(all_observations),
        "unique_raw_columns": len(raw_columns),
        "suggested_canonical_groups": len(merge_rows),
        "sample_values_per_column": SAMPLE_VALUES_OUT,
        "output_file": output_path,
    }])

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sanitize_df(merge_suggestion_df).to_excel(writer, sheet_name="Merge_Sheet_Suggestion", index=False)
        sanitize_df(column_samples_df).to_excel(writer, sheet_name="Column_Samples", index=False)
        sanitize_df(per_sheet_df).to_excel(writer, sheet_name="Per_Sheet_Mapping", index=False)
        sanitize_df(profiles_df).to_excel(writer, sheet_name="Column_Profiles", index=False)
        if not ambiguous_df.empty:
            sanitize_df(ambiguous_df).to_excel(writer, sheet_name="Ambiguous_Name_Like", index=False)
        if not skipped_df.empty:
            sanitize_df(skipped_df).to_excel(writer, sheet_name="Skipped_Files", index=False)
        sanitize_df(summary_df).to_excel(writer, sheet_name="Run_Summary", index=False)

    print("\nDone.")
    print(f"  Unique raw columns : {len(raw_columns)}")
    print(f"  Canonical groups   : {len(merge_rows)}")
    print(f"  Sheet observations : {len(all_observations)}")
    if all_failures:
        print(f"  Failed files/sheets: {len(all_failures)} (see Skipped_Files sheet)")
    print(f"  Output             : {output_path}")
    print("\nMerge groups preview:")
    for _, row in merge_suggestion_df.iterrows():
        print(f"  {row['display_name']}: {row['raw_columns']}")


def main():
    parser = argparse.ArgumentParser(
        description="Suggest column merges using data content + per-sheet mapping."
    )
    parser.add_argument("folder", nargs="?", default=DEFAULT_FOLDER)
    parser.add_argument(
        "--output", "-o", default=DEFAULT_OUTPUT,
        help="output xlsx path",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    run(args.folder, args.output, overwrite=args.overwrite)


if __name__ == "__main__":
    main()
