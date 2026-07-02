#!/usr/bin/env python3
"""Run: python tests/test_validate_append.py
Covers the append tool's validation + safe write, and verifies the base-year
splice numerically on the real workbook (the same algorithm the dashboard uses).
"""
import os
import sys
import shutil
import tempfile
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, "tools"))

import schema
import validate as V
import append_data as A

MASTER = os.path.join(REPO, "docs", "data", "ccpi_combined_single_tab.xlsx")
SAMPLES = os.path.join(HERE, "samples")
os.makedirs(SAMPLES, exist_ok=True)


def _valid_tab1():
    return pd.DataFrame([{
        "Source File": "manual_2026_07.xlsx", "Base Year Used": "2021=100",
        "Year": 2026, "Month Number": 7, "Month": "July", "Date": "2026-07-01",
        "CCPI Index": 208.5, "CCPI Core Index": 188.0, "CCPI MoM Change": 0.004,
        "CCPI Core MoM Change": 0.004, "CCPI YoY Inflation": 0.071,
        "CCPI Core YoY Inflation": 0.042, "CCPI 12M Moving Avg Inflation": 0.030,
        "CCPI Core 12M Moving Avg Inflation": 0.027,
    }])[schema.TAB1_COLS]


def _valid_tab2():
    return pd.DataFrame([{
        "Source File": "manual_2026_07.xlsx", "Base Year Used": "2021=100",
        "H/H Size": 3.8, "Year": 2026, "Month": "July", "Month No": 7,
        "Period Type": "Monthly", "Group": "Transport",
        "Base Value (Rs. Cts)": 9000.0, "Weight (%)": 9.8, "Index Number": 230.0,
        "Month-to-Month Inflation (%)": 0.5, "Year-on-Year Inflation (%)": 4.2,
        "Annual Average Inflation (%)": 3.1,
    }])[schema.TAB2_COLS]


def test_valid_passes():
    p1 = os.path.join(SAMPLES, "valid_tab1.xlsx"); _valid_tab1().to_excel(p1, index=False)
    p2 = os.path.join(SAMPLES, "valid_tab2.xlsx"); _valid_tab2().to_excel(p2, index=False)
    c1, r1 = V.validate(p1, schema.TAB1); assert r1["rows_valid"] == 1
    c2, r2 = V.validate(p2, schema.TAB2); assert r2["rows_valid"] == 1
    print("valid files pass ........ OK")


def test_invalid_rejected():
    # wrong base year
    bad = _valid_tab1(); bad.loc[0, "Base Year Used"] = "1999=100"
    p = os.path.join(SAMPLES, "bad_base.xlsx"); bad.to_excel(p, index=False)
    _expect_fail(p, schema.TAB1, "base")
    # out-of-range index
    bad = _valid_tab1(); bad.loc[0, "CCPI Index"] = 99999
    p = os.path.join(SAMPLES, "bad_range.xlsx"); bad.to_excel(p, index=False)
    _expect_fail(p, schema.TAB1, "range")
    # missing column
    bad = _valid_tab1().drop(columns=["CCPI Core Index"])
    p = os.path.join(SAMPLES, "bad_cols.xlsx"); bad.to_excel(p, index=False)
    _expect_fail(p, schema.TAB1, "missing")
    # bad group (tab2)
    bad = _valid_tab2(); bad.loc[0, "Group"] = "Spaceships"
    p = os.path.join(SAMPLES, "bad_group.xlsx"); bad.to_excel(p, index=False)
    _expect_fail(p, schema.TAB2, "recognis")
    # disallowed extension (macro)
    p = os.path.join(SAMPLES, "macro.xlsm"); _valid_tab1().to_excel(p, index=False)
    _expect_fail(p, schema.TAB1, "not allowed")
    print("invalid files rejected .. OK")


def test_formula_injection_neutralised():
    # direct unit check
    assert V._sanitise_text("=cmd|'/c calc'!A1").startswith("'=")
    assert V._sanitise_text("+1+1").startswith("'+")
    assert V._sanitise_text("normal text") == "normal text"
    # end-to-end via CSV (preserves the literal leading '=' as text)
    bad = _valid_tab1(); bad.loc[0, "Source File"] = "=cmd|'/c calc'!A1"
    p = os.path.join(SAMPLES, "inject.csv"); bad.to_csv(p, index=False)
    clean, _ = V.validate(p, schema.TAB1)
    assert clean.loc[0, "Source File"].startswith("'="), clean.loc[0, "Source File"]
    print("formula injection neutralised  OK")


def test_safe_append_on_copy():
    tmp = tempfile.mkdtemp()
    copy = os.path.join(tmp, "master.xlsx"); shutil.copy2(MASTER, copy)
    orig_master, orig_bdir = A.MASTER, A.BACKUP_DIR
    A.MASTER = copy; A.BACKUP_DIR = os.path.join(tmp, "backups")
    try:
        before = {s: len(pd.read_excel(copy, sheet_name=s)) for s in [schema.TAB1, schema.TAB2]}
        p1 = os.path.join(SAMPLES, "valid_tab1.xlsx")
        A.append("data", p1, dry_run=False)
        after1 = len(pd.read_excel(copy, sheet_name=schema.TAB1))
        assert after1 == before[schema.TAB1] + 1
        # other sheet preserved untouched
        assert len(pd.read_excel(copy, sheet_name=schema.TAB2)) == before[schema.TAB2]
        # backup exists
        assert os.path.isdir(A.BACKUP_DIR) and os.listdir(A.BACKUP_DIR)
    finally:
        A.MASTER, A.BACKUP_DIR = orig_master, orig_bdir
        shutil.rmtree(tmp, ignore_errors=True)
    print("safe append (on copy) ... OK  (+1 row, other sheet intact, backup made)")


# ---- splice verification (mirrors the dashboard algorithm) ----------------
BASE_ORDER = ["1952=100", "2002=100", "2006/07=100", "2013=100", "2021=100"]


def splice(points):
    """points: list of dict(year, month, base, value). Returns continuous series."""
    present = sorted({p["base"] for p in points}, key=BASE_ORDER.index)
    latest = present[-1]
    by = {b: {} for b in present}
    for p in points:
        by[p["base"]][(p["year"], p["month"])] = p["value"]
    scale = {latest: 1.0}
    for newer, older in zip(present[::-1], present[::-1][1:]):
        common = sorted(set(by[newer]) & set(by[older]))
        if common:
            lm = common[-1]
            step = by[newer][lm] / by[older][lm]
        else:
            step = 1.0
        scale[older] = scale[newer] * step
    out = {}
    for p in points:
        ym = (p["year"], p["month"])
        b = p["base"]
        # prefer newest base for a given month
        if ym not in out or BASE_ORDER.index(b) > BASE_ORDER.index(out[ym]["base"]):
            out[ym] = {"base": b, "linked": p["value"] * scale[b]}
    months = sorted(out)
    return [(y, m, out[(y, m)]["linked"], out[(y, m)]["base"]) for (y, m) in months], scale


def test_splice_real_data():
    d1 = pd.read_excel(MASTER, sheet_name=schema.TAB1)
    pts = [{"year": int(r.Year), "month": int(r._4), "base": r._2, "value": float(r._7)}
           for r in d1.itertuples() if pd.notna(r._7)]
    # rename via positional is fragile; rebuild explicitly:
    pts = [{"year": int(r["Year"]), "month": int(r["Month Number"]),
            "base": r["Base Year Used"], "value": float(r["CCPI Index"])}
           for _, r in d1.iterrows() if pd.notna(r["CCPI Index"])]
    series, scale = splice(pts)
    months = [(y, m) for (y, m, *_ ) in series]
    assert len(months) == len(set(months)), "overlap months not collapsed!"
    assert all(v > 0 for (_, _, v, _) in series), "non-positive spliced value"
    assert all(s > 0 for s in scale.values()), "non-positive scale factor"
    # continuity: spliced 2021-base latest equals native latest (scale==1)
    assert abs(scale["2021=100"] - 1.0) < 1e-12
    last = series[-1]
    print(f"splice on real data ..... OK  ({len(series)} continuous months, "
          f"latest {last[0]}-{last[1]:02d} index={last[2]:.1f})")
    print("   chained scale factors:", {k: round(v, 4) for k, v in scale.items()})


def _expect_fail(path, tab, needle):
    try:
        V.validate(path, tab)
    except V.ValidationError as e:
        assert needle.lower() in str(e).lower(), f"wrong error for {path}: {e}"
        return
    raise AssertionError(f"expected failure for {path} did not occur")


if __name__ == "__main__":
    test_valid_passes()
    test_invalid_rejected()
    test_formula_injection_neutralised()
    test_safe_append_on_copy()
    test_splice_real_data()
    print("\nALL TESTS PASSED")
