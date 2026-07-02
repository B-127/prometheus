"""Strict validation + security checks for files being appended to the workbook.

Philosophy: the master workbook is precious and lives on the user's device, so an
input file is treated as UNTRUSTED until proven to exactly match the target tab's
structure. We read values only (never evaluate formulas or macros), enforce size/
row caps, restrict file types, sanitise text against spreadsheet-formula injection,
and reject on the first structural mismatch.
"""
from __future__ import annotations
import os
import zipfile
import pandas as pd

from schema import (SPEC, RANGES, GROUPS_TAB2, PERIOD_TYPES, MONTHS, TAB1, TAB2,
                    MAX_FILE_BYTES, MAX_INPUT_ROWS, ALLOWED_INPUT_EXT)


class ValidationError(Exception):
    pass


# --- security gate ---------------------------------------------------------
def security_check(path: str):
    """Reject anything dangerous before we even parse it."""
    if not os.path.isfile(path):
        raise ValidationError(f"file not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in ALLOWED_INPUT_EXT:
        raise ValidationError(
            f"extension '{ext}' not allowed. Use .xlsx or .csv "
            f"(macro formats like .xlsm/.xls are blocked).")
    size = os.path.getsize(path)
    if size == 0:
        raise ValidationError("file is empty.")
    if size > MAX_FILE_BYTES:
        raise ValidationError(f"file too large ({size} bytes > {MAX_FILE_BYTES}).")
    if ext == ".xlsx":
        if not zipfile.is_zipfile(path):
            raise ValidationError("'.xlsx' is not a valid Office Open XML file.")
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if any("vbaProject" in n for n in names):
                raise ValidationError("workbook contains a macro project (vbaProject) — blocked.")
            # zip-bomb guard: total uncompressed size sanity
            total = sum(i.file_size for i in z.infolist())
            if total > 200 * 1024 * 1024:
                raise ValidationError("uncompressed contents unreasonably large — blocked.")


def _read(path: str, sheet: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False, nrows=MAX_INPUT_ROWS + 1)
    else:
        # values only; openpyxl does not execute formulas/macros
        xls = pd.ExcelFile(path, engine="openpyxl")
        target = sheet if sheet in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=target, dtype=str, nrows=MAX_INPUT_ROWS + 1)
    if len(df) > MAX_INPUT_ROWS:
        raise ValidationError(f"too many rows (> {MAX_INPUT_ROWS}).")
    return df


_DANGEROUS_PREFIX = ("=", "+", "-", "@", "\t", "\r")


def _sanitise_text(v):
    """Neutralise spreadsheet-formula injection in free-text cells."""
    if v is None:
        return v
    s = str(v)
    if s and s[0] in _DANGEROUS_PREFIX:
        return "'" + s          # leading apostrophe forces text in Excel
    return s


# --- structural validation -------------------------------------------------
def validate(path: str, tab: str):
    """Return (clean_df, report). Raises ValidationError on structural failure."""
    if tab not in SPEC:
        raise ValidationError(f"unknown tab '{tab}'.")
    spec = SPEC[tab]
    security_check(path)
    raw = _read(path, tab)

    # exact column set (order-independent, but no missing / no extras)
    cols = list(raw.columns)
    missing = [c for c in spec["cols"] if c not in cols]
    extra = [c for c in cols if c not in spec["cols"]]
    if missing:
        raise ValidationError(f"missing required columns: {missing}")
    if extra:
        raise ValidationError(f"unexpected columns present: {extra}")
    df = raw[spec["cols"]].copy()       # reorder to canonical

    report = {"tab": tab, "rows_in": len(df), "rejected": [], "warnings": []}
    if len(df) == 0:
        raise ValidationError("no data rows.")

    # coerce numerics; collect per-row errors
    numeric_cols = [c for c in spec["cols"] if c in RANGES]
    good_rows, problems = [], []
    for i, row in df.iterrows():
        errs = []
        rec = {}
        for c in spec["cols"]:
            val = row[c]
            blank = (val is None) or (str(val).strip() == "") or (str(val).lower() == "nan")
            if c in spec["required"] and blank:
                errs.append(f"{c} is required")
            if c in numeric_cols and not blank:
                try:
                    num = float(str(val).replace(",", ""))
                except ValueError:
                    errs.append(f"{c}='{val}' is not numeric"); rec[c] = None; continue
                lo, hi = RANGES[c]
                if not (lo <= num <= hi):
                    errs.append(f"{c}={num} out of range [{lo},{hi}]")
                rec[c] = num
            else:
                rec[c] = None if blank else _sanitise_text(val)
        # controlled vocabularies
        if rec.get("Base Year Used") not in spec["bases"]:
            errs.append(f"Base Year Used '{rec.get('Base Year Used')}' not in {sorted(spec['bases'])}")
        if rec.get("Month") not in MONTHS and not (str(rec.get("Month")).strip()==""):
            errs.append(f"Month '{rec.get('Month')}' invalid")
        if tab == TAB2:
            if rec.get("Group") not in GROUPS_TAB2:
                errs.append(f"Group '{rec.get('Group')}' not recognised")
            if rec.get("Period Type") not in PERIOD_TYPES:
                errs.append(f"Period Type '{rec.get('Period Type')}' invalid")
        if errs:
            problems.append({"row": int(i) + 2, "errors": errs})   # +2: header + 1-index
        else:
            good_rows.append(rec)

    report["rejected"] = problems
    report["rows_valid"] = len(good_rows)
    if problems:
        # 100%-accuracy stance: any structural error fails the whole file.
        raise ValidationError(
            f"{len(problems)} invalid row(s). First issues: "
            + "; ".join(f"row {p['row']}: {p['errors'][0]}" for p in problems[:5]))

    clean = pd.DataFrame(good_rows, columns=spec["cols"])
    return clean, report


def find_duplicates(clean: pd.DataFrame, existing: pd.DataFrame, tab: str):
    """Return key tuples in clean that already exist (warning only)."""
    key = SPEC[tab]["key"]
    if existing.empty:
        return []
    e = existing.copy()
    for k in key:
        e[k] = e[k].astype(str)
    have = set(map(tuple, e[key].astype(str).values.tolist()))
    dups = [tuple(r) for r in clean[key].astype(str).values.tolist() if tuple(r) in have]
    return dups
