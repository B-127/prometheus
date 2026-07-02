# Sri Lanka CCPI Dashboard

A static dashboard that visualises the CCPI workbook **directly in the browser**, plus a
**local, validated tool** for appending new data to the workbook. The workbook in
`docs/data/` is the single source of truth.

## How it fits together

```
docs/data/ccpi_combined_single_tab.xlsx   <- source of truth (you update this)
        │
        ├── docs/index.html   reads it in the browser, splices base years, charts it
        │                     (GitHub Pages — static, no server, nothing uploaded)
        │
        └── tools/append_data.py   validates a new file & appends it (runs on YOUR machine)
```

- **Dashboard:** open the Pages site. It fetches the workbook, rescales every base
  year onto the latest (2021=100) into one continuous series, marks each base
  change, and **flags every month that exists under more than one base** (amber band
  + an "Overlap?" column) so cross-base comparisons aren't distorted.
- **Appending data:** you run one command locally; the file never leaves your computer.

## The two tabs (and a critical gotcha)

| | Tab 1 `CCPI Data` | Tab 2 `CCPI Subgroup Breakdown` |
|---|---|---|
| Grain | one row per month | one row per month × group |
| Inflation stored as | **decimals** (0.022 = 2.2%) | **percent** (10.3 = 10.3%) |
| Key | Base + Year + Month Number | Base + Year + Month No + Period Type + Group |

The dashboard handles both conventions and always displays inflation as %. The
append tool never normalises values — it appends them exactly, so keep each file in
its tab's native convention.

## Appending new data (the feature)

```bash
pip install -r requirements.txt

# dry run first — validates and shows a report, writes nothing
python tools/append_data.py --tab data     --input new_movements.xlsx --dry-run
python tools/append_data.py --tab subgroup  --input new_subgroup.xlsx --dry-run

# real run — backs up, then appends atomically
python tools/append_data.py --tab data     --input new_movements.xlsx
```

`--tab data` → Tab 1 (Movements of CCPI & CCPI core). `--tab subgroup` → Tab 2.
Your input file must have **exactly** the target tab's columns; anything that doesn't
match the structure, vocabulary, or sane numeric ranges is rejected and nothing is
written. After a successful append, commit the updated workbook to publish it.

## Security (what protects your device + the master file)

- **Static dashboard** — no upload endpoint, no server, no secrets; it only *reads*.
- **Input files are untrusted:** only `.xlsx`/`.csv` accepted; **macro formats blocked**
  (`.xlsm`/`.xls`), workbooks containing a VBA project rejected, size capped (25 MB),
  row count capped, zip-bomb guard.
- **Values only** — the tool never evaluates formulas or runs macros from input files.
- **Spreadsheet-formula-injection neutralised** — any text cell starting with `= + - @`
  is escaped before it can be written.
- **The master is protected on write** — timestamped backup before every append, an
  exclusive lock to prevent concurrent writes, and an atomic temp-then-replace so an
  interrupted run can never corrupt the workbook. Backups live in `backups/`.
- **Strict, fail-closed validation** — a single bad row rejects the whole file.

## Deploy the dashboard (GitHub Pages)

Settings → **Pages** → Source: *Deploy from a branch* → Branch **main**, folder **/docs** →
Save. The `.nojekyll` file is already included so the dashboard (not the README) is served.
Live at `https://<user>.github.io/<repo>/`. Updates whenever you commit a new workbook.

Preview locally: `cd docs && python -m http.server 8000` → open `http://localhost:8000`.

## Tests

```bash
python tests/test_validate_append.py
```
Covers valid/invalid validation, formula-injection neutralisation, a safe append on a
**copy** of the master, and a numeric check of the base-year splice on the real data.

---

## File-change map (vs the previous `sl-inflation-mvp` repo)

This is effectively a **new repo** — the old PDF-collection pipeline is removed per your
instruction to drop the data-collection work.

**Add (new):**
- `docs/index.html` — the new dashboard (reads the workbook, splices, flags overlaps)
- `docs/.nojekyll`
- `docs/data/ccpi_combined_single_tab.xlsx` — your workbook (source of truth)
- `tools/schema.py`, `tools/validate.py`, `tools/append_data.py` — the append feature
- `tests/test_validate_append.py`
- `requirements.txt`, `.gitignore`, `README.md`

**Remove (no longer used):**
- `src/` (extract.py, download.py, splice.py, store.py, pipeline.py, validate.py, schema.py, extract_llm.py)
- `scripts/run_monthly.py`, `scripts/backfill.py`, `scripts/build_web_json.py`
- `.github/workflows/refresh.yml` — there is no automated scraping anymore
- the old `data/` outputs (`inflation_ccpi.xlsx`, CSV mirrors) and old `docs/` dashboard

**Keep (conceptually carried over, rewritten here):** the base-year splice idea and the
validation-first mindset — both reimplemented around your actual workbook.
