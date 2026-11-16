# Global Data Validator

A Python toolkit for comparing two tabular datasets — different file formats,
different structures, same underlying records — and producing a clear,
field-level diff. It auto-detects join keys, flags structural and content
mismatches, and renders the results into a formatted Excel report.

Built to answer a recurring data-engineering question: *"are these two
exports of the same data actually the same?"* — without hand-writing a new
comparison script for every pair of files.

## What it does

- **Reads mixed formats** — XLSX, CSV, TSV, TXT, Parquet, JSON — into a
  common in-memory representation (Polars), with format sniffing for
  delimiter, encoding, quoting, and header detection on delimited files.
- **Input integrity checks** — row/column consistency, duplicate headers,
  banner/decorator row detection, before any comparison happens.
- **Normalization** — trims whitespace, unifies null tokens, canonicalizes
  numeric strings (e.g. `0042` → `42`, `12.0` → `12`), and optionally
  parses date-like strings — all conservative, logged transformations so
  results stay explainable.
- **Automatic key detection** — searches for a single column (or a
  composite pair) that is non-null and unique on both sides to use as the
  join key; falls back to an interactive key-builder when nothing
  qualifies outright.
- **Field-level comparison** — null-safe equality across matched records,
  with per-field mismatch counts, examples, and severity ordering.
- **Excel reporting** — fills a formatted workbook template with the
  summary and detail findings (overview block, mismatch tables, styled
  borders) so results are easy to review outside the terminal.

## Tech stack

- **Polars** — primary engine for reading, normalizing, and comparing data
- **Pandas** — Excel I/O interop and a few dataframe conveniences
- **NumPy** — numeric helpers
- **OpenPyXL** — Excel report template filling / styling
- **PyArrow** — Parquet support

Python 3.12+ is the reference environment.

## Project layout

```
├── global_validator.py           # Main validation pipeline (CLI)
├── report_cd.py                  # Excel report generator (reads validator's text logs)
├── run_validation_and_report.py  # Convenience wrapper: validator + report in one step
├── key_creator.py                # Standalone script to probe/build a key column on one file
├── utils.py                      # Small, reusable data-prep helpers (dedup, header inspection, date fixups)
├── validator_config.example.json # Config template (copy to validator_config.json)
├── requirements.txt
├── sample_data/                  # Tiny synthetic CSVs for a quick demo
├── templates/
│   └── report_base_template.xlsx # Blank formatted workbook used by report_cd.py
├── CODE_MAINTENANCE_GUIDE.md      # Per-function map of the codebase, for maintainers
└── LICENSE
```

`logs/` and `reports/` are created at run time and are git-ignored.

## How it works, in short

1. `global_validator.py` loads both datasets, runs integrity checks, applies
   normalization, detects (or lets you choose) the join key, aligns records,
   and compares every shared field. It writes
   `logs/summary_<report_name>.txt` and `logs/detail_<report_name>.txt`.
2. `report_cd.py` parses those two text logs and fills an Excel template
   (`templates/report_base_template.xlsx`, included in this repo — a plain
   two-sheet workbook; swap in your own formatted version, or adapt
   `fill_template()`, for different styling) to produce
   `reports/report_<report_name>.xlsx`.
3. `run_validation_and_report.py` just runs both steps back to back with the
   defaults from `validator_config.json`.

## Getting started

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

cp validator_config.example.json validator_config.json
```

Edit `validator_config.json`:

- `primary_file`, `secondary_file` — your two input datasets
- `report_name` — used to name the output log/report files
- optional: `mapping_file` (column rename table), sheet overrides, and the
  normalization flags described inline in the file

## Example usage — full walkthrough

A tiny synthetic dataset is included under `sample_data/` — two CSVs
describing the same set of fictional orders, with a few deliberate
mismatches (a quantity that differs, a price that differs, a blank field,
one row only in each file) so the diff output has something to show. This
walkthrough has been run end to end against that data; the steps and
output below are real, not illustrative.

1. Set up the environment as in [Getting started](#getting-started) above.

2. Point the config at the sample data (edit `validator_config.json`):

   ```json
   {
     "primary_file": "sample_data/primary.csv",
     "secondary_file": "sample_data/secondary.csv",
     "report_name": "demo"
   }
   ```

3. Run the validator on its own to produce the text logs:

   ```bash
   python global_validator.py
   cat logs/summary_demo.txt
   cat logs/detail_demo.txt
   ```

   For this sample data, auto key detection picks `order_id`, finds 7
   matching records (1 row only in primary, 1 only in secondary), and flags
   3 field-level mismatches: a `product` that's blank on one side, a
   `quantity` off by 2, and a `unit_price` off by 0.50.

4. Build the Excel report from those logs (template already included under
   `templates/`):

   ```bash
   python report_cd.py
   ```

   This produces `reports/report_demo.xlsx` with two sheets:
   `High_Level_Overview` (row/column counts, match rate, a green
   VALIDATED / red NOT VALIDATED badge) and `Dettaglio_Disallineamenti`
   (the non-common keys plus one bordered block per mismatched field,
   each showing the key, primary value, and secondary value).

5. Or run both steps in one command:

   ```bash
   python run_validation_and_report.py
   ```

Override the config path on any script with `--config /path/to/file.json`.
To use your own formatted workbook instead of the included plain template,
set `"excel_template_path"` in `validator_config.json` or pass
`--template-path` to `report_cd.py`.

## Notes

- `key_creator.py` and `utils.py` are standalone helper scripts for
  one-off, exploratory data prep outside the main pipeline — edit the
  placeholder paths at the top before running them directly.
- No real-world or employer data is included or referenced anywhere in
  this repository; `sample_data/` is entirely fictional.

## License

MIT — see [LICENSE](LICENSE).
