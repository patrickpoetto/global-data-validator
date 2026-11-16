"""
Utility snippets for one-off data preparation tasks.

Small, reusable helpers for common pre-processing steps around the main
validation pipeline (deduplication, quick header inspection). Each function
is self-contained and safe to import or copy into a one-off script.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


# ============================================================
# 1) REMOVE PERFECT DUPLICATES (ALL COLUMNS)
# ============================================================
def remove_perfect_duplicates(input_path: str, output_path: str) -> None:
    """
    Remove rows that are exact duplicates across all columns.
    """
    src = Path(input_path)
    dst = Path(output_path)

    df = pd.read_excel(src) if src.suffix.lower() in {".xlsx", ".xls"} else pd.read_csv(src)
    total_rows = len(df)
    duplicate_rows = int(df.duplicated().sum())
    dedup_df = df.drop_duplicates()

    if dst.suffix.lower() in {".xlsx", ".xls"}:
        dedup_df.to_excel(dst, index=False)
    else:
        dedup_df.to_csv(dst, index=False)

    print(f"Total rows: {total_rows}")
    print(f"Duplicate rows removed: {duplicate_rows}")
    print(f"Remaining rows: {len(dedup_df)}")
    print(f"Saved to: {dst}")


# ============================================================
# 2) REMOVE DUPLICATES BY CONCATENATED KEY COLUMNS
# ============================================================
def remove_duplicates_by_concat_key(
    input_path: str,
    output_path: str,
    key_columns: Iterable[str],
    separator: str = "-",
) -> None:
    """
    Remove duplicates based on a concatenated key built from selected columns.
    """
    src = Path(input_path)
    dst = Path(output_path)

    df = pd.read_excel(src) if src.suffix.lower() in {".xlsx", ".xls"} else pd.read_csv(src)

    key_columns = list(key_columns)
    missing = [c for c in key_columns if c not in df.columns]
    if missing:
        raise ValueError(f"Missing key columns: {missing}")

    total_rows = len(df)
    df["_key_concat"] = df[key_columns].astype(str).agg(separator.join, axis=1)
    dedup_df = df.drop_duplicates(subset="_key_concat").drop(columns=["_key_concat"])

    if dst.suffix.lower() in {".xlsx", ".xls"}:
        dedup_df.to_excel(dst, index=False)
    else:
        dedup_df.to_csv(dst, index=False)

    print(f"Initial rows: {total_rows}")
    print(f"Final rows: {len(dedup_df)}")
    print(f"Duplicates removed: {total_rows - len(dedup_df)}")
    print(f"Saved to: {dst}")


# ============================================================
# 3) PRINT FILE HEADERS
# ============================================================
def print_headers(input_path: str, csv_separator: str = "|") -> None:
    """
    Print one header per line from a CSV or Excel file.
    """
    src = Path(input_path)

    if src.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(src, nrows=0)
    else:
        df = pd.read_csv(src, sep=csv_separator, nrows=0)

    print("Headers:")
    for col in df.columns:
        print(col)


# ============================================================
# 4) CONVERT EXCEL SERIAL / TEXT DATES TO ISO STRINGS
# ============================================================
def convert_excel_dates(file_path: str, columns: Iterable[str]) -> None:
    """
    Normalize a set of date-like columns (Excel serial numbers or free-text
    dates) to ISO "YYYY-MM-DD HH:MM:SS" strings, printing a short before/after
    sample for each converted column.
    """
    print(f"[START] Reading file: {file_path}")

    df = pd.read_excel(file_path)
    print(f"[INFO] Rows: {len(df)}, Columns: {len(df.columns)}")

    def convert_value(val):
        if pd.isna(val):
            return val

        if isinstance(val, (int, float)):
            try:
                dt = pd.to_datetime(val, unit="D", origin="1899-12-30")
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                return val

        if isinstance(val, str):
            try:
                dt = pd.to_datetime(val, errors="coerce")
                if pd.notna(dt):
                    return dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                return val

        return val

    for col in columns:
        print(f"\n[COLUMN] {col}")

        if col not in df.columns:
            print("[WARNING] Column not found")
            continue

        original_values = df[col].copy()
        df[col] = df[col].apply(convert_value)

        converted_mask = original_values != df[col]
        num_converted = converted_mask.sum()
        print(f"[INFO] Cells converted: {num_converted}")

        sample = df.loc[converted_mask, col].head(4)
        original_sample = original_values.loc[sample.index]

        if len(sample) > 0:
            print("[SAMPLE] Conversion examples:")
            for before, after in zip(original_sample, sample):
                print(f"  {before}  -->  {after}")
        else:
            print("[SAMPLE] No conversion performed")

    print("\n[WRITE] Overwriting file...")
    df.to_excel(file_path, index=False)
    print("[DONE] File updated successfully.")


# ============================================================
# 5) SUFFIX DUPLICATE KEY VALUES (_1, _2, ...)
# ============================================================
def suffix_duplicate_keys(file_path: str, column_name: str) -> None:
    """
    For a given column, append `_1`, `_2`, ... to repeated values so each
    row ends up with a unique key, keeping the first occurrence unchanged.
    """
    print(f"[START] Reading file: {file_path}")

    df = pd.read_excel(file_path, dtype={column_name: str})
    print(f"[INFO] Rows: {df.shape[0]}, Columns: {df.shape[1]}")

    counts = df.groupby(column_name).size()
    duplicates = counts[counts > 1]

    if duplicates.empty:
        print(f"\n[No duplicates found in column {column_name}]")
    else:
        print(f"\n[DUPLICATES] Column: {column_name}")
        for val, count in duplicates.items():
            print(f"  {val} -> {count} occurrences")

    def add_suffix(group):
        if len(group) == 1:
            return group
        new_values = [group.iloc[0]] + [
            f"{val}_{i}" for i, val in enumerate(group.iloc[1:], start=1)
        ]
        return pd.Series(new_values, index=group.index)

    df[column_name] = df.groupby(column_name)[column_name].transform(add_suffix)

    print("\n[WRITE] Writing file...")
    df.to_excel(file_path, index=False)
    print("[DONE] File updated")


# ============================================================
# Example usage (replace with your own paths when running directly)
# ============================================================
if __name__ == "__main__":
    # remove_perfect_duplicates("path/to/input.xlsx", "path/to/output.xlsx")
    # remove_duplicates_by_concat_key("path/to/input.csv", "path/to/output.csv", ["key_col_1", "key_col_2"])
    # print_headers("path/to/input.csv", csv_separator="|")
    # convert_excel_dates("path/to/input.xlsx", ["date_column_1", "date_column_2"])
    # suffix_duplicate_keys("path/to/input.xlsx", "id_column")
    pass
