# Code Maintenance Guide

Reference for maintainers: high-level map of the codebase, then per-function stubs for **`global_validator.py`** and **`report_cd.py`**. Other scripts (`utils.py`, `key_creator.py`, `run_validation_and_report.py`) are thin helpers and are not enumerated here.

---

## 1. Macro areas (index)

| Area | Responsibility |
|------|------------------|
| **Configuration & logging** | JSON settings, stable log filenames, session banner in summary file. |
| **Input integrity** | Detect encoding/separator/quoting for CSV/TXT; row-count consistency; Excel sheet shape and duplicate headers; optional interactive stripping of banner/empty/decorator rows. |
| **File reading** | Load primary/secondary into Polars from XLSX, CSV, TSV, TXT, Parquet, JSON; Excel pre-clean prompts. |
| **Normalization** | Trim, null-token unification, numeric canonicalization (including leading zeros), conservative date-like string handling. |
| **Structural analysis** | Dimensions, common/only columns, types, duplicate column names. |
| **Key detection & matching** | Auto single-column key, auto composite pair, interactive key builder (null fill, duplicate suffix/remove on last key column), record match stats. |
| **Content comparison** | Polars null-safe inequality on aligned keys; summary metrics; optional comparison when no key. |
| **Deep analysis** | Per-field mismatch listing with examples (feeds `detail_<name>.txt`). |
| **Statistics** | Null/numeric/text summaries for exploratory logging. |
| **Column mapping** | Optional rename map from external table. |
| **Orchestration** | `run_validation` wires phases; `main` CLI. |
| **Excel report (separate module)** | Parse text reports, fill `templates/report_base_template.xlsx`, write `reports/report_<name>.xlsx`. |

---

## 2. `global_validator.py`

### Module-level

| Function | Brief description |
|----------|-------------------|
| `_safe_report_filename_segment` | Sanitizes `report_name` for use in log filenames. |
| `configure_logging` | Attaches file + console handlers; writes `logs/summary_<name>.txt` (overwrite per run). |
| `_coerce_sheet_value` | Normalizes JSON `null`/sheet index for Excel sheet selection. |
| `load_settings` | Loads and validates `validator_config.json` into `ValidatorSettings`. |
| `write_detail_report` | Serializes deep mismatch content to `logs/detail_<name>.txt`. |
| `finalize_summary_log` | Closes the file handler after validation (summary path already fixed). |
| `run_input_integrity_checks` | Runs integrity inspectors on both inputs; logs findings. |
| `run_validation` | End-to-end pipeline: read → normalize → map → analyze → compare → deep report dict. |
| `parse_args` | CLI: `--config` path. |
| `main` | Loads settings, configures logging, runs validation, writes detail + finalizes summary. |

### `FileConfig`

| Method | Brief description |
|--------|-------------------|
| `set_files` | Stores and checks existence of primary/secondary paths. |
| `get_file_info` | Basic metadata (size, modified time) for a path. |

### `FileIntegrityInspector`

| Method | Brief description |
|--------|-------------------|
| `detect_delimited_format` | Sniffs encoding, separator, quote, header, logical-row reconstruction flag for CSV/TXT. |
| `inspect_delimited_integrity` | Validates consistent column counts per row. |
| `inspect_excel_integrity` | Rows/columns, duplicate header names. |
| `_detect_line_ending` | CR/LF/CRLF detection. |
| `_detect_best_separator` | Scores candidate delimiters; ignores trivial single-column splits. |
| `_detect_header` | Heuristic: first line looks like header vs data. |
| `_should_reconstruct_logical_rows` | Whether to merge wrapped physical lines. |
| `_reconstruct_logical_rows` | Rejoins continuation lines for TXT/CSV. |
| `_is_decorator_line` | Detects dashed/placeholder lines for pre-cleaning. |
| `_cell_to_text` | Normalizes Excel cell values to text for inspection. |
| `_prompt_cleaning_confirmation` | Terminal yes/no per candidate row removal (Excel/TXT). |

### `FileReader`

| Method | Brief description |
|--------|-------------------|
| `read_file` | Dispatcher by extension to format-specific readers. |
| `_read_excel` | Pandas → Polars; optional `skiprows` after interactive banner removal. |
| `_read_csv` | Polars read with detected options. |
| `_read_tsv` | Tab-separated variant. |
| `_read_txt` | Uses integrity hints; optional cleaning + logical rows + `read_csv`. |
| `_read_parquet` | Polars native. |
| `_read_json` | Polars/pandas path for JSON tables. |
| `get_excel_sheets` | Lists sheet names for logging/selection. |

### `DataNormalizer`

| Method | Brief description |
|--------|-------------------|
| `normalize_dataframe` | Column rename trim, per-column string normalization, numeric/date helpers per settings. |
| `_normalize_column_names` | Strips header whitespace. |
| `_normalize_string_column` | Trim, null tokens, numeric canonicalization, optional date-like parsing with fallback. |
| `_looks_numeric` | Quick test if string looks numeric. |
| `_canonical_numeric_string` | Strips leading zeros on request; normalizes decimal text. |
| `_try_parse_date_to_iso` | Multi-strategy date parse; returns ISO or `None` to keep original. |

### `DataAnalyzer`

| Method | Brief description |
|--------|-------------------|
| `analyze_structure` | Row/column counts, dimension deltas. |
| `analyze_columns` | Common columns, primary-only / secondary-only columns. |
| `analyze_data_types` | Per-dataset type summaries. |
| `_find_duplicate_columns` | Duplicate header detection helper. |
| `run_full_analysis` | Bundles structure + columns + types. |

### `ContentComparator`

| Method | Brief description |
|--------|-------------------|
| `_null_safe_difference_mask` | Polars expr: true when values differ including NULL vs non-NULL. |
| `find_matching_records` | Detects keys (auto or interactive), aligns datasets, returns match stats and downstream payloads. |
| `_detect_key_columns` | Single-column then composite pair search under no-null uniqueness rules. |
| `_find_single_unique_key` | First column with zero nulls and full uniqueness on both sides. |
| `_interactive_key_creator` | Terminal flow: pick columns, fill nulls, resolve duplicates. |
| `_prompt_key_column_selection` | Parses user index list for key columns. |
| `_fill_key_nulls` | Fills key-column nulls with user-provided token. |
| `_get_column_null_counts` | Null counts for candidate columns. |
| `_get_columns_with_low_null_rate` | Filters columns under null-rate threshold for suggestions. |
| `_get_duplicate_key_groups` | Groups rows with duplicate composite keys. |
| `_print_duplicate_groups_side_by_side` | Prints duplicate counts primary vs secondary. |
| `_dedup_by_suffix_pandas` | Deduplicates by suffixing `_1`, `_2`… on **last** key column. |
| `_dedup_by_remove_pandas` | Keeps first row per duplicate key. |
| `_evaluate_key_columns` | Null/duplicate stats for chosen keys. |
| `_print_duplicate_key_details` | Terminal detail for duplicate keys in one dataset. |
| `_find_composite_key` | Brute-force pair of columns with full uniqueness both sides. |
| `_create_composite_key` | Builds UTF-8 composite key column with `-` separator. |
| `_analyze_value_differences` | Field-by-field comparison on matching keys; builds summary for reports. |
| `_generate_content_analysis_report` | Shapes high-level mismatch statistics for logging. |
| `_compare_without_keys` | Fallback structural comparison when no usable key exists. |

### `DeepContentReporter` *(extends `ContentComparator`)*

| Method | Brief description |
|--------|-------------------|
| `analyze_deep_content_differences` | Orders fields by severity; fills `field_details` / examples for detail report. |
| `_analyze_single_field_deep` | Per-field Polars join and null-safe mismatch extraction with examples. |

### `StatisticalAnalyzer`

| Method | Brief description |
|--------|-------------------|
| `analyze_data_quality` | Orchestrates null/numeric/text quality sections. |
| `_analyze_null_values` | Per-column null rates. |
| `_analyze_numeric_columns` | Basic numeric stats. |
| `_analyze_text_columns` | Text length / uniqueness style stats. |

### `ColumnMapper`

| Method | Brief description |
|--------|-------------------|
| `_load_mapping_file` | Reads mapping XLSX/CSV into rename dict (trimmed keys/values). |
| `apply_mapping` | Renames columns on a Polars frame when enabled. |

---

## 3. `report_cd.py`

| Function | Brief description |
|----------|-------------------|
| `_safe_report_segment` | Same filename sanitization as validator for `report_name`. |
| `parse_args` | CLI: `--config`, optional overrides for summary/detail/template/output paths. |
| `resolve_paths_from_config` | Reads JSON; defaults: `logs/summary_<name>.txt`, `logs/detail_<name>.txt`, `templates/report_base_template.xlsx`, `reports/report_<name>.xlsx`. |
| `_to_int` | Parses integers with comma thousands from log text. |
| `_dataset_file_label` | Chooses display file label (basename) for Excel headers. |
| `parse_summary` | Regex-based parser for metrics and keys from summary log text. |
| `_only_key_lists_for_excel` | Builds sorted unique key lists from parsed “only samples” lines. |
| `parse_detail` | Parses `FIELD:` blocks and `idx \| key \| primary \| secondary` lines. |
| `fill_template` | Loads template workbook, fills overview + detail sheets (Non Common Keys block, misalignment triples, caps, borders, styling). |
| `fill_template` → `apply_block_borders` *(nested)* | Thick outer / thin inner borders for one column block. |
| `main` | Resolves paths, validates files exist, runs `fill_template`, prints output path. |

---

## 4. Suggested maintenance workflow

1. **Behavior change in validation** → start from `run_validation` and the class handling that phase (`FileReader`, `DataNormalizer`, `ContentComparator`, etc.).
2. **Output text format change** → update `write_detail_report` / logging in `ContentComparator` / `DeepContentReporter`, then adjust `parse_summary` / `parse_detail` in `report_cd.py` if the Excel pipeline must follow.
3. **Excel layout change** → edit `templates/report_base_template.xlsx` and align cell addresses / merge logic inside `fill_template`.
