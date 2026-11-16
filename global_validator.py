#!/usr/bin/env python3
"""
Global Data Validator - Universal file comparison and validation tool
=====================================================================

Comprehensive tool for validating and comparing datasets across multiple file formats.
Supports Excel (.xlsx), CSV (.csv), Parquet (.parquet), JSON (.json), TSV (.tsv), and TXT (.txt). Default separator is tab. ,
Generates detailed reports with embedded analysis.

Author: TXT E-Tech - Patrick Poetto
Version: 1.0.0
"""

import polars as pl
import pandas as pd  # Required for Excel file reading compatibility
import numpy as np
import logging
import json
import argparse
import csv
import io
import re
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Union, Any, Tuple
from dataclasses import dataclass
from collections import Counter
import sys


# =============================================================================
# LOGGER CONFIGURATION
# =============================================================================

def _safe_report_filename_segment(report_name: str) -> str:
    """Filesystem-safe segment from config report_name (stable across runs)."""
    s = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in report_name.strip())
    return s or "default"


def configure_logging(report_name: str) -> None:
    """
    Configure file + console logging. Summary file is logs/summary_<report_name>.txt (no timestamp;
    overwritten on each run). Session time is still logged inside the file.
    """
    global logger
    logs_dir = Path("logs")
    logs_dir.mkdir(exist_ok=True)
    safe = _safe_report_filename_segment(report_name)
    log_file = logs_dir / f"summary_{safe}.txt"
    
    logger = logging.getLogger("GlobalValidator")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    
    file_handler = logging.FileHandler(log_file, mode="w", encoding="utf-8")
    file_formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(file_formatter)
    file_handler.setLevel(logging.INFO)
    
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter("%(message)s"))
    console_handler.setLevel(logging.INFO)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    logger.info("=" * 80)
    logger.info("GLOBAL DATA VALIDATOR - VALIDATION SESSION STARTED")
    logger.info("=" * 80)
    logger.info("")
    logger.info(f"Log file: {log_file}")
    logger.info("")
    logger.info(f"Session timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("")


# =============================================================================
# FILE CONFIGURATION SECTION
# =============================================================================

class FileConfig:
    """
    Configuration class for file paths and validation parameters.
    Handles the different formats and validation settings.
    """
    
    def __init__(self):
        """Initialize file configuration with default values."""
        self.primary_file: Optional[Path] = None
        self.secondary_file: Optional[Path] = None
        self.supported_formats = {'.xlsx', '.csv', '.parquet', '.json', '.tsv', '.txt'}
        
    def set_files(self, primary_path: str, secondary_path: str) -> bool:
        """
        Set and validate file paths for comparison.
        
        Args:
            primary_path (str): Path to primary file
            secondary_path (str): Path to secondary file
            
        Returns:
            bool: True if files are valid, False otherwise
        """
        try:
            self.primary_file = Path(primary_path)
            self.secondary_file = Path(secondary_path)
            
            # Validate file existence
            if not self.primary_file.exists():
                logger.error(f"Primary file not found: {self.primary_file}")
                return False
                
            if not self.secondary_file.exists():
                logger.error(f"Secondary file not found: {self.secondary_file}")
                return False
            
            # Validate file formats
            if self.primary_file.suffix.lower() not in self.supported_formats:
                logger.error(f"Unsupported primary file format: {self.primary_file.suffix}")
                return False
                
            if self.secondary_file.suffix.lower() not in self.supported_formats:
                logger.error(f"Unsupported secondary file format: {self.secondary_file.suffix}")
                return False
            
            logger.info("File Configuration:")
            logger.info(f"  Primary file:   {self.primary_file}")
            logger.info(f"  Secondary file: {self.secondary_file}")
            logger.info("")
            
            return True
            
        except Exception as e:
            logger.error(f"Error configuring files: {str(e)}")
            return False
    
    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        """
        Get information about a file - high level properties.
        
        Args:
            file_path (Path): Path to file
            
        Returns:
            Dict[str, Any]: File information dictionary
        """
        try:
            stat = file_path.stat()
            return {
                'name': file_path.name,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'format': file_path.suffix.lower(),
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception as e:
            logger.error(f"Error getting file info for {file_path}: {str(e)}")
            return {}


# =============================================================================
# APPLICATION CONFIGURATION
# =============================================================================

DEFAULT_CONFIG_PATH = Path("validator_config.json")


@dataclass
class ValidatorSettings:
    """Runtime settings loaded from a centralized JSON configuration file."""
    primary_file: str
    secondary_file: str
    primary_sheet: Optional[Union[str, int]] = None
    secondary_sheet: Optional[Union[str, int]] = None
    mapping_file: Optional[str] = None
    apply_mapping_to: str = "both"
    run_deep_analysis: bool = True
    normalize_before_compare: bool = True
    normalize_numeric_strings: bool = True
    strip_leading_zeros_numeric: bool = True
    convert_dates: bool = True
    report_name: str = "default"


def _coerce_sheet_value(sheet_value: Any) -> Optional[Union[str, int]]:
    """Convert JSON sheet values into runtime sheet parameters."""
    if sheet_value is None:
        return None

    if isinstance(sheet_value, int):
        return sheet_value

    if isinstance(sheet_value, str):
        cleaned = sheet_value.strip()
        if not cleaned:
            return None
        if cleaned.isdigit():
            return int(cleaned)
        return cleaned

    return None


def load_settings(config_path: Union[str, Path]) -> ValidatorSettings:
    """Load validator settings from JSON config file."""
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_file}. "
            f"Create it from the provided validator_config.json template."
        )

    with config_file.open("r", encoding="utf-8") as handle:
        raw_config = json.load(handle)

    required_fields = ["primary_file", "secondary_file"]
    missing_fields = [field for field in required_fields if not raw_config.get(field)]
    if missing_fields:
        raise ValueError(f"Missing required config fields: {', '.join(missing_fields)}")

    return ValidatorSettings(
        primary_file=str(raw_config["primary_file"]),
        secondary_file=str(raw_config["secondary_file"]),
        primary_sheet=_coerce_sheet_value(raw_config.get("primary_sheet")),
        secondary_sheet=_coerce_sheet_value(raw_config.get("secondary_sheet")),
        mapping_file=raw_config.get("mapping_file"),
        apply_mapping_to=str(raw_config.get("apply_mapping_to", "both")),
        run_deep_analysis=bool(raw_config.get("run_deep_analysis", True)),
        normalize_before_compare=bool(raw_config.get("normalize_before_compare", True)),
        normalize_numeric_strings=bool(raw_config.get("normalize_numeric_strings", True)),
        strip_leading_zeros_numeric=bool(raw_config.get("strip_leading_zeros_numeric", True)),
        convert_dates=bool(raw_config.get("convert_dates", True)),
        report_name=str(raw_config.get("report_name", "default")).strip() or "default"
    )


def write_detail_report(
    results: Dict[str, Any],
    report_name: str
) -> Path:
    """Write deep mismatch details into a dedicated text report."""
    safe_report_name = _safe_report_filename_segment(report_name)
    base_dir = Path("logs")
    deep_target = base_dir / f"detail_{safe_report_name}.txt"
    deep_target.parent.mkdir(parents=True, exist_ok=True)

    deep = results.get("deep_content_analysis", {})
    deep_lines = [
        "GLOBAL VALIDATOR - DEEP CONTENT REPORT",
        "=" * 80,
        "",
        f"Key columns: {deep.get('key_columns', [])}",
        f"Total matching records: {deep.get('total_matching_records', 0)}",
        f"Fields analyzed: {deep.get('fields_analyzed', 0)}",
        f"Fields with differences: {deep.get('fields_with_differences', 0)}",
        ""
    ]

    field_details = deep.get("field_details", {})
    sorted_field_items = sorted(
        field_details.items(),
        key=lambda item: (
            item[1].get("differences_found", 0),
            item[1].get("difference_percentage", 0.0)
        ),
        reverse=True
    )
    for field_name, field_data in sorted_field_items:
        deep_lines.append("-" * 80)
        deep_lines.append(f"FIELD: {field_name}")
        deep_lines.append(f"Comparisons: {field_data.get('total_comparisons', 0)}")
        deep_lines.append(f"Differences: {field_data.get('differences_found', 0)}")
        deep_lines.append(f"Difference %: {field_data.get('difference_percentage', 0.0):.4f}")
        deep_lines.append("")
        deep_lines.append("# | KEY | PRIMARY | SECONDARY")

        for idx, example in enumerate(field_data.get("examples", []), 1):
            deep_lines.append(
                f"{idx} | {example.get('key', '')} | {example.get('primary', '')} | {example.get('secondary', '')}"
            )
        deep_lines.append("")

    deep_target.write_text("\n".join(deep_lines), encoding="utf-8")
    return deep_target


def finalize_summary_log(report_name: str) -> Path:
    """
    Close the file log handler. Summary is already logs/summary_<name>.txt (no rename).
    """
    safe_report_name = _safe_report_filename_segment(report_name)
    summary_path = Path("logs") / f"summary_{safe_report_name}.txt"

    file_handler: Optional[logging.FileHandler] = None
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            file_handler = handler
            break

    if file_handler is None:
        raise RuntimeError("No file logger handler available")

    file_handler.flush()
    file_handler.close()
    logger.removeHandler(file_handler)

    return summary_path.resolve()


# =============================================================================
# FILE READERS SECTION
# =============================================================================

class FileIntegrityInspector:
    """
    Detect file format details and perform integrity checks.

    Focus:
    - CSV/TXT: encoding, separator, quote char, line ending, column consistency
    - XLSX: selected sheet shape and duplicate header detection
    """

    ENCODINGS_TO_TRY = ["utf-8", "utf-8-sig", "windows-1252", "iso-8859-1", "cp1252", "latin1", "cp850"]
    SEPARATORS_TO_TRY = [",", ";", "\t", "|"]

    @staticmethod
    def detect_delimited_format(file_path: Path, file_kind: str) -> Dict[str, Any]:
        """Autodetect encoding, separator and basic parsing hints."""
        best_candidate: Optional[Dict[str, Any]] = None

        for encoding in FileIntegrityInspector.ENCODINGS_TO_TRY:
            try:
                with open(file_path, "r", encoding=encoding, errors="strict") as handle:
                    sample_lines = [handle.readline().rstrip("\n\r") for _ in range(80)]
            except Exception:
                continue

            non_empty = [line for line in sample_lines if line.strip()]
            if not non_empty:
                continue

            line_ending = FileIntegrityInspector._detect_line_ending(file_path)
            separator, consistency, avg_cols = FileIntegrityInspector._detect_best_separator(non_empty)
            quote_char = '"' if any('"' in line for line in non_empty[:20]) else None
            has_header = FileIntegrityInspector._detect_header(non_empty[0], separator)
            logical_row_mode = (
                file_kind == "txt" and FileIntegrityInspector._should_reconstruct_logical_rows(non_empty, separator)
            )

            candidate = {
                "encoding": encoding,
                "separator": separator,
                "quote_char": quote_char,
                "line_ending": line_ending,
                "has_header": has_header,
                "consistency_score": consistency,
                "avg_columns": avg_cols,
                "logical_row_reconstruction": logical_row_mode
            }

            if best_candidate is None or candidate["consistency_score"] > best_candidate["consistency_score"]:
                best_candidate = candidate

        if best_candidate is None:
            # Conservative fallback
            best_candidate = {
                "encoding": "utf-8",
                "separator": "|" if file_kind == "txt" else ",",
                "quote_char": '"',
                "line_ending": "unknown",
                "has_header": True,
                "consistency_score": 0.0,
                "avg_columns": 0,
                "logical_row_reconstruction": False
            }

        return best_candidate

    @staticmethod
    def inspect_delimited_integrity(file_path: Path, format_info: Dict[str, Any]) -> Dict[str, Any]:
        """Check row consistency for CSV/TXT files."""
        encoding = format_info["encoding"]
        separator = format_info["separator"]
        quote_char = format_info.get("quote_char") or '"'
        reconstruct = bool(format_info.get("logical_row_reconstruction", False))

        row_lengths: List[int] = []
        sample_issues: List[Dict[str, Any]] = []
        total_rows = 0

        if reconstruct:
            with open(file_path, "r", encoding=encoding, errors="replace") as handle:
                raw_lines = [line.rstrip("\n\r") for line in handle]
            logical_rows = FileIntegrityInspector._reconstruct_logical_rows(raw_lines, separator)
            for idx, row in enumerate(logical_rows, 1):
                fields = row.split(separator)
                row_len = len(fields)
                row_lengths.append(row_len)
                total_rows += 1
        else:
            with open(file_path, "r", encoding=encoding, errors="replace", newline="") as handle:
                reader = csv.reader(
                    handle,
                    delimiter=separator,
                    quotechar=quote_char,
                    doublequote=True,
                    strict=False
                )
                for idx, row in enumerate(reader, 1):
                    row_len = len(row)
                    row_lengths.append(row_len)
                    total_rows += 1

        if not row_lengths:
            return {
                "total_rows": 0,
                "expected_columns": 0,
                "is_consistent": True,
                "column_distribution": {},
                "issue_samples": []
            }

        distribution = dict(sorted(Counter(row_lengths).items(), key=lambda x: x[0]))
        expected_columns = max(distribution.items(), key=lambda x: x[1])[0]

        # Build issue samples (max 10)
        if reconstruct:
            with open(file_path, "r", encoding=encoding, errors="replace") as handle:
                raw_lines = [line.rstrip("\n\r") for line in handle]
            logical_rows = FileIntegrityInspector._reconstruct_logical_rows(raw_lines, separator)
            for idx, row in enumerate(logical_rows, 1):
                cols = len(row.split(separator))
                if cols != expected_columns:
                    sample_issues.append({"row_number": idx, "columns": cols, "preview": row[:220]})
                    if len(sample_issues) >= 10:
                        break
        else:
            with open(file_path, "r", encoding=encoding, errors="replace", newline="") as handle:
                reader = csv.reader(
                    handle,
                    delimiter=separator,
                    quotechar=quote_char,
                    doublequote=True,
                    strict=False
                )
                for idx, row in enumerate(reader, 1):
                    cols = len(row)
                    if cols != expected_columns:
                        preview = separator.join(row)[:220]
                        sample_issues.append({"row_number": idx, "columns": cols, "preview": preview})
                        if len(sample_issues) >= 10:
                            break

        consistent_rows = distribution.get(expected_columns, 0)
        consistency_ratio = consistent_rows / total_rows if total_rows > 0 else 1.0
        is_fully_consistent = consistent_rows == total_rows

        return {
            "total_rows": total_rows,
            "expected_columns": expected_columns,
            "is_consistent": is_fully_consistent,
            "consistency_ratio": consistency_ratio,
            "column_distribution": distribution,
            "issue_samples": sample_issues
        }

    @staticmethod
    def inspect_excel_integrity(file_path: Path, sheet_name: Optional[Union[str, int]]) -> Dict[str, Any]:
        """Collect structural integrity info for Excel files."""
        if sheet_name is None:
            effective_sheet: Union[str, int] = 0
        else:
            effective_sheet = sheet_name

        pandas_df = pd.read_excel(file_path, sheet_name=effective_sheet, dtype=str)
        columns = [str(col) for col in pandas_df.columns]
        duplicate_headers = {
            col: count for col, count in Counter(columns).items() if count > 1
        }

        return {
            "sheet": effective_sheet,
            "rows": int(pandas_df.shape[0]),
            "columns": int(pandas_df.shape[1]),
            "duplicate_headers": duplicate_headers,
            "is_consistent": len(duplicate_headers) == 0
        }

    @staticmethod
    def _detect_line_ending(file_path: Path) -> str:
        """Detect dominant line ending style from raw bytes."""
        try:
            with open(file_path, "rb") as handle:
                chunk = handle.read(8192)
            crlf = chunk.count(b"\r\n")
            lf = chunk.count(b"\n")
            if crlf > 0 and crlf == lf:
                return "CRLF"
            if lf > 0:
                return "LF"
            return "unknown"
        except Exception:
            return "unknown"

    @staticmethod
    def _detect_best_separator(lines: List[str]) -> Tuple[str, float, int]:
        import csv
        from collections import Counter
        import statistics

        def safe_split(line: str, sep: str):
            try:
                return next(csv.reader([line], delimiter=sep, quotechar='"'))
            except Exception:
                return line.split(sep)

        best_sep = ","
        best_score = float("-inf")
        best_avg_cols = 0

        for sep in FileIntegrityInspector.SEPARATORS_TO_TRY:

            parsed = [safe_split(line, sep) for line in lines if line.strip()]
            if not parsed:
                continue

            col_counts = [len(row) for row in parsed]

            # drop non existent separators 
            if max(col_counts) <= 1:
                continue

            avg_cols = sum(col_counts) / len(col_counts)

            distribution = Counter(col_counts)
            mode_cols, mode_freq = max(distribution.items(), key=lambda x: x[1])
            consistency = mode_freq / len(col_counts)

            std_dev = statistics.pstdev(col_counts) if len(col_counts) > 1 else 0

            # balanced score
            score = avg_cols * 2 + consistency * 5 - std_dev * 3

            if score > best_score:
                best_score = score
                best_sep = sep
                best_avg_cols = int(avg_cols)

        # fallback
        if best_score == float("-inf"):
            try:
                sample = "\n".join(lines[:50])
                dialect = csv.Sniffer().sniff(sample)
                return dialect.delimiter, 1.0, 0
            except Exception:
                return ",", 0.0, 0

        return best_sep, best_score, best_avg_cols

    @staticmethod
    def _detect_header(first_line: str, separator: str) -> bool:
        """Simple heuristic: headers are usually mostly non-numeric."""
        fields = [field.strip() for field in first_line.split(separator)]
        if not fields:
            return False
        non_numeric = 0
        for field in fields:
            numeric_candidate = field.replace(".", "").replace("-", "").replace(",", "")
            if not numeric_candidate.isdigit():
                non_numeric += 1
        return non_numeric >= max(1, int(len(fields) * 0.5))

    @staticmethod
    def _should_reconstruct_logical_rows(lines: List[str], separator: str) -> bool:
        """
        Detect wrapped-row TXT exports where continuation lines do not start
        with the separator.
        """
        if separator != "|":
            return False
        starts_with_sep = sum(1 for line in lines if line.startswith(separator))
        non_empty = sum(1 for line in lines if line.strip())
        if non_empty == 0:
            return False
        ratio = starts_with_sep / non_empty
        # Heuristic: enough anchor rows, but not all rows are already anchored.
        return 0.4 <= ratio < 0.98

    @staticmethod
    def _reconstruct_logical_rows(raw_lines: List[str], separator: str) -> List[str]:
        """Rebuild logical rows by appending continuation lines."""
        if separator != "|":
            return [line for line in raw_lines if line.strip()]

        logical_rows: List[str] = []
        current = ""
        for line in raw_lines:
            if not line.strip():
                continue
            if line.startswith(separator):
                if current:
                    logical_rows.append(current)
                current = line
            else:
                current += line
        if current:
            logical_rows.append(current)
        return logical_rows

    @staticmethod
    def _is_decorator_line(value: str) -> bool:
        """Detect separator/decorator rows (-----, =====, ||||||, etc.)."""
        stripped = value.strip()
        if not stripped:
            return False
        if len(stripped) < 3:
            return False
        return bool(re.fullmatch(r"[-=|_+~\s]+", stripped))

    @staticmethod
    def _cell_to_text(value: Any) -> str:
        """Convert cell value to normalized text for heuristics."""
        if value is None:
            return ""
        text = str(value).strip()
        if text.lower() in {"nan", "none", "null"}:
            return ""
        return text

    @staticmethod
    def _prompt_cleaning_confirmation(file_label: str, candidates: List[Dict[str, Any]]) -> bool:
        """
        Show candidate rows in terminal and ask final confirmation.
        Returns True if user confirms removal.
        """
        if not candidates:
            return False
        if not sys.stdin.isatty():
            print(f"\n[CLEANING] {file_label}: interactive confirmation unavailable (non-interactive terminal).")
            return False

        print("\n" + "=" * 72)
        print(f"CLEANING CANDIDATES DETECTED - {file_label}")
        print("=" * 72)
        for item in candidates:
            print(f"[{item['row_number']}] ({item['reason']}) {item['preview']}")
        print("")
        answer = input("Remove ALL candidate rows above? (y/N): ").strip().lower()
        return answer == "y"


def run_input_integrity_checks(
    primary_file: Path,
    secondary_file: Path,
    primary_sheet: Optional[Union[str, int]],
    secondary_sheet: Optional[Union[str, int]]
) -> Dict[str, Any]:
    """
    Run input integrity checks before full dataset loading.

    This preflight is especially useful for CSV/TXT where encoding and separator
    may vary across vendor/source systems.
    """
    logger.info("")
    logger.info("=" * 60)
    logger.info("INPUT FORMAT AND INTEGRITY CHECK")
    logger.info("=" * 60)
    logger.info("")

    checks: Dict[str, Any] = {}
    files_to_check = [
        ("primary", primary_file, primary_sheet),
        ("secondary", secondary_file, secondary_sheet),
    ]

    for label, file_path, sheet_name in files_to_check:
        ext = file_path.suffix.lower()
        logger.info(f"[{label.upper()}] {file_path.name} ({ext})")

        if ext in {".csv", ".txt"}:
            fmt = FileIntegrityInspector.detect_delimited_format(
                file_path,
                file_kind="txt" if ext == ".txt" else "csv"
            )
            integrity = FileIntegrityInspector.inspect_delimited_integrity(file_path, fmt)
            logger.info(
                f"  Detected -> encoding={fmt['encoding']}, sep={repr(fmt['separator'])}, "
                f"line_ending={fmt['line_ending']}, quote={repr(fmt.get('quote_char'))}"
            )
            logger.info(
                f"  Integrity -> expected_columns={integrity['expected_columns']}, "
                f"rows={integrity['total_rows']:,}, consistency={integrity['consistency_ratio'] * 100:.2f}%"
            )
            if not integrity["is_consistent"]:
                logger.error("  ❌ Integrity check failed: inconsistent column count across rows.")
                logger.error("  ❌ Strict mode enabled: all rows must have identical column count.")
                logger.error("  Sample problematic rows:")
                for issue in integrity["issue_samples"][:5]:
                    logger.error(
                        f"    row={issue['row_number']} cols={issue['columns']} preview='{issue['preview']}'"
                    )
                raise ValueError(
                    f"{label} file integrity check failed: inconsistent column count detected "
                    f"(expected {integrity['expected_columns']} columns)."
                )
            logger.info("")
            checks[label] = {"format": fmt, "integrity": integrity}

        elif ext in {".xlsx", ".xls"}:
            excel_info = FileIntegrityInspector.inspect_excel_integrity(file_path, sheet_name)
            logger.info(
                f"  Excel -> sheet={excel_info['sheet']}, rows={excel_info['rows']:,}, "
                f"columns={excel_info['columns']}"
            )
            if excel_info["duplicate_headers"]:
                logger.warning(f"  ⚠ Duplicate headers detected: {excel_info['duplicate_headers']}")
            else:
                logger.info("  Header integrity -> no duplicate headers")
            logger.info("")
            checks[label] = {"format": {"type": "excel"}, "integrity": excel_info}

        else:
            logger.info("  Integrity preflight not implemented for this format; proceeding with standard load.")
            logger.info("")
            checks[label] = {"format": {"type": ext}, "integrity": {"status": "not_checked"}}

    return checks


class FileReader:
    """
    File reader that converts different formats to Polars DataFrames. 
    Polars is used instead of pandas for data manipulation and computations. Pandas is only needed for .xlsx compatibility.
    Supports Excel (.xlsx), CSV (.csv), Parquet (.parquet), JSON (.json), TSV (.tsv), TXT (.txt).
    """
    
    @staticmethod
    def read_file(file_path: Path, sheet_name: Optional[Union[str, int]] = None) -> pl.DataFrame:
        """
        Read file and convert to Polars DataFrame regardless of format.
        
        Args:
            file_path (Path): Path to the file
            sheet_name (Optional[Union[str, int]]): Sheet name/index for Excel files
            
        Returns:
            pl.DataFrame: Loaded data as Polars DataFrame
            
        Raises:
            ValueError: If file format is not supported
            FileNotFoundError: If file doesn't exist
        """
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        file_format = file_path.suffix.lower()
        logger.info(f"Reading {file_format} file: {file_path.name}")
        logger.info("")
        
        try:
            if file_format == '.xlsx':
                return FileReader._read_excel(file_path, sheet_name)
            elif file_format == '.csv':
                return FileReader._read_csv(file_path)
            elif file_format == '.tsv':
                return FileReader._read_tsv(file_path)
            elif file_format == '.txt':
                return FileReader._read_txt(file_path)
            elif file_format == '.parquet':
                return FileReader._read_parquet(file_path)
            elif file_format == '.json':
                return FileReader._read_json(file_path)
            else:
                raise ValueError(f"Unsupported file format: {file_format}")
                
        except Exception as e:
            logger.error(f"Error reading file {file_path}: {str(e)}")
            raise
    
    @staticmethod
    def _read_excel(file_path: Path, sheet_name: Optional[Union[str, int]] = None) -> pl.DataFrame:
        """Read Excel file using pandas then convert to Polars."""
        try:
            effective_sheet: Union[str, int] = sheet_name if sheet_name is not None else 0
            skiprows = 0
            excel_cleaning_report = {
                "candidates_detected": 0,
                "rows_removed": 0,
                "interactive_used": False
            }

            # Preview top rows to detect optional leading banner/decorator noise.
            preview_df = pd.read_excel(
                file_path,
                sheet_name=effective_sheet,
                header=None,
                dtype=str,
                nrows=30
            )
            excel_candidates: List[Dict[str, Any]] = []
            for ridx in range(preview_df.shape[0]):
                row_values = [FileIntegrityInspector._cell_to_text(v) for v in preview_df.iloc[ridx].tolist()]
                non_empty = [v for v in row_values if v != ""]
                if not non_empty:
                    excel_candidates.append({
                        "row_number": ridx + 1,
                        "reason": "empty",
                        "preview": "<empty>"
                    })
                    continue
                if all(FileIntegrityInspector._is_decorator_line(v) for v in non_empty):
                    excel_candidates.append({
                        "row_number": ridx + 1,
                        "reason": "decorator",
                        "preview": " | ".join(non_empty[:5])[:200]
                    })
                    continue
                if len(non_empty) == 1:
                    excel_candidates.append({
                        "row_number": ridx + 1,
                        "reason": "leading_banner",
                        "preview": non_empty[0][:200]
                    })
                    continue
                break

            if excel_candidates:
                excel_cleaning_report["candidates_detected"] = len(excel_candidates)
                if FileIntegrityInspector._prompt_cleaning_confirmation(file_path.name, excel_candidates):
                    skiprows = len(excel_candidates)
                    excel_cleaning_report["rows_removed"] = skiprows
                    excel_cleaning_report["interactive_used"] = True

            # Use pandas for Excel reading (Polars uses pandas under the hood anyway)
            if sheet_name is not None:
                pandas_df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, skiprows=skiprows)
                logger.info(f"  Sheet: {sheet_name}")
                logger.info("")
            else:
                pandas_df = pd.read_excel(file_path, dtype=str, skiprows=skiprows)
                logger.info(f"  Sheet: default (first sheet)")
                logger.info("")
            if excel_cleaning_report["candidates_detected"] > 0:
                logger.info(
                    "  Excel pre-cleaning:"
                    f" candidates={excel_cleaning_report['candidates_detected']},"
                    f" removed={excel_cleaning_report['rows_removed']}"
                )
                logger.info("")
            
            # Convert to Polars
            polars_df = pl.from_pandas(pandas_df)
            logger.info(f"  Loaded: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            raise
    
    @staticmethod
    def _read_csv(file_path: Path) -> pl.DataFrame:
        """Read CSV file with robust parsing for different formats."""
        logger.info(f"  CSV parser - analyzing file format...")
        csv_info = FileIntegrityInspector.detect_delimited_format(file_path, file_kind="csv")
        logger.info(
            f"  Detected CSV settings: encoding={csv_info['encoding']}, separator={repr(csv_info['separator'])}, "
            f"line_ending={csv_info['line_ending']}"
        )

        # Try Polars with detected settings
        try:
            polars_df = pl.read_csv(
                file_path, 
                separator=csv_info['separator'],
                encoding=csv_info['encoding'],
                quote_char=csv_info.get("quote_char") or '"',
                infer_schema_length=0  # Keep all columns as strings initially
            )
            logger.info(f"  Polars success: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df
            
        except Exception as polars_error:
            logger.debug(f"  Polars failed, trying pandas: {str(polars_error)[:100]}")
            
            # Step 3: Fallback to pandas with detected settings
            try:
                pandas_df = pd.read_csv(
                    file_path, 
                    sep=csv_info['separator'],
                    encoding=csv_info['encoding'],
                    quotechar=csv_info.get("quote_char") or '"',
                    dtype=str,
                    low_memory=False
                )
                polars_df = pl.from_pandas(pandas_df)
                logger.info(f"  Pandas success: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
                logger.info("")
                return polars_df
                
            except Exception as pandas_error:
                logger.error(f"Both Polars and pandas failed: {str(pandas_error)[:100]}")
                raise
    
    # @staticmethod
    # def _analyze_csv_format(file_path: Path) -> dict:
    #     """Analyze CSV file format to determine optimal parsing settings."""
    #     encodings_to_try = ['utf-8', 'utf-8-sig', 'windows-1252', 'iso-8859-1']
        
    #     for encoding in encodings_to_try:
    #         try:
    #             with open(file_path, 'r', encoding=encoding) as f:
    #                 # Read first few lines
    #                 lines = [f.readline().rstrip() for _ in range(5)]
                
    #             # Remove BOM if present (utf-8-sig handles this, but let's be explicit)
    #             if lines and lines[0].startswith('\ufeff'):
    #                 lines[0] = lines[0][1:]
                
    #             # Detect separator
    #             separators = [',', ';', '\t', '|']
    #             separator_counts = {}
                
    #             for line in lines[:3]:  # Check first 3 lines
    #                 if line.strip():  # Skip empty lines
    #                     for sep in separators:
    #                         count = line.count(sep)
    #                         separator_counts[sep] = separator_counts.get(sep, 0) + count
                
    #             # Choose the most common separator
    #             best_separator = max(separator_counts.items(), key=lambda x: x[1])[0] if separator_counts else ','
                
    #             # Check if we need to skip rows (detect headers vs data)
    #             skip_rows = 0
    #             if lines:
    #                 # Simple heuristic: if first line has many non-numeric fields, it's likely a header
    #                 first_line_parts = lines[0].split(best_separator)
    #                 non_numeric = sum(1 for part in first_line_parts if part.strip() and not part.strip().replace('.', '').replace('-', '').isdigit())
    #                 if non_numeric > len(first_line_parts) * 0.5:  # More than 50% non-numeric
    #                     skip_rows = 0  # It's a header, don't skip
                    
    #             logger.debug(f"CSV format detected:")
    #             logger.debug(f"    Encoding: {encoding}")
    #             logger.debug(f"    Separator: {repr(best_separator)}")
    #             logger.debug(f"    Skip rows: {skip_rows}")
                
    #             return {
    #                 'encoding': encoding,
    #                 'separator': best_separator,
    #                 'skip_rows': skip_rows
    #             }
                
    #         except Exception:
    #             continue
        
    #     # Fallback to defaults
    #     return {
    #         'encoding': 'utf-8',
    #         'separator': ',',
    #         'skip_rows': 0
    #     }
    
    @staticmethod
    def _read_tsv(file_path: Path) -> pl.DataFrame:
        """Read TSV file using Polars with tab separator."""
        try:
            polars_df = pl.read_csv(file_path, separator='\t')
            logger.info(f"  Loaded: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df
            
        except Exception as e:
            logger.error(f"Error reading TSV file: {str(e)}")
            raise
    
    @staticmethod
    def _read_txt(file_path: Path) -> pl.DataFrame:
        """Read TXT file with robust format detection and logical-row support."""
        txt_info = FileIntegrityInspector.detect_delimited_format(file_path, file_kind="txt")
        encoding = txt_info["encoding"]
        separator = txt_info["separator"]
        quote_char = txt_info.get("quote_char") or '"'
        reconstruct = bool(txt_info.get("logical_row_reconstruction", False))

        logger.info(
            f"  Detected TXT settings: encoding={encoding}, separator={repr(separator)}, "
            f"logical_row_reconstruction={reconstruct}, line_ending={txt_info['line_ending']}"
        )

        try:
            txt_cleaning_report = {
                "candidates_detected": 0,
                "rows_removed": 0,
                "interactive_used": False
            }

            if reconstruct:
                with open(file_path, "r", encoding=encoding, errors="replace") as handle:
                    raw_lines = [line.rstrip("\n\r") for line in handle]
                candidates: List[Dict[str, Any]] = []
                cleaned_lines: List[str] = []
                seen_structured_row = False
                for idx, line in enumerate(raw_lines, 1):
                    stripped = line.strip()
                    if not stripped:
                        candidates.append({"row_number": idx, "reason": "empty", "preview": "<empty>"})
                        continue
                    if FileIntegrityInspector._is_decorator_line(stripped):
                        candidates.append({"row_number": idx, "reason": "decorator", "preview": stripped[:220]})
                        continue
                    if not seen_structured_row and stripped.count(separator) == 0:
                        candidates.append({"row_number": idx, "reason": "leading_banner", "preview": stripped[:220]})
                        continue
                    if stripped.count(separator) > 0:
                        seen_structured_row = True
                    cleaned_lines.append(line)

                if candidates:
                    txt_cleaning_report["candidates_detected"] = len(candidates)
                    if FileIntegrityInspector._prompt_cleaning_confirmation(file_path.name, candidates):
                        raw_lines = cleaned_lines
                        txt_cleaning_report["rows_removed"] = len(candidates)
                        txt_cleaning_report["interactive_used"] = True

                logical_rows = FileIntegrityInspector._reconstruct_logical_rows(raw_lines, separator)
                payload = "\n".join(logical_rows)
                pandas_df = pd.read_csv(
                    io.StringIO(payload),
                    sep=separator,
                    quotechar=quote_char,
                    dtype=str,
                    low_memory=False
                )
            else:
                with open(file_path, "r", encoding=encoding, errors="replace") as handle:
                    raw_lines = [line.rstrip("\n\r") for line in handle]
                candidates = []
                cleaned_lines = []
                seen_structured_row = False
                for idx, line in enumerate(raw_lines, 1):
                    stripped = line.strip()
                    if not stripped:
                        candidates.append({"row_number": idx, "reason": "empty", "preview": "<empty>"})
                        continue
                    if FileIntegrityInspector._is_decorator_line(stripped):
                        candidates.append({"row_number": idx, "reason": "decorator", "preview": stripped[:220]})
                        continue
                    if not seen_structured_row and stripped.count(separator) == 0:
                        candidates.append({"row_number": idx, "reason": "leading_banner", "preview": stripped[:220]})
                        continue
                    if stripped.count(separator) > 0:
                        seen_structured_row = True
                    cleaned_lines.append(line)

                input_lines = raw_lines
                if candidates:
                    txt_cleaning_report["candidates_detected"] = len(candidates)
                    if FileIntegrityInspector._prompt_cleaning_confirmation(file_path.name, candidates):
                        input_lines = cleaned_lines
                        txt_cleaning_report["rows_removed"] = len(candidates)
                        txt_cleaning_report["interactive_used"] = True

                pandas_df = pd.read_csv(
                    io.StringIO("\n".join(input_lines)),
                    sep=separator,
                    quotechar=quote_char,
                    dtype=str,
                    low_memory=False,
                )

            if txt_cleaning_report["candidates_detected"] > 0:
                logger.info(
                    "  TXT pre-cleaning:"
                    f" candidates={txt_cleaning_report['candidates_detected']},"
                    f" removed={txt_cleaning_report['rows_removed']}"
                )
                logger.info("")

            polars_df = pl.from_pandas(pandas_df)
            logger.info(f"  Loaded: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df

        except Exception as e:
            logger.error(f"Error reading TXT file: {str(e)}")
            raise
    
    @staticmethod
    def _read_parquet(file_path: Path) -> pl.DataFrame:
        """Read Parquet file using Polars native reader."""
        try:
            polars_df = pl.read_parquet(file_path)
            logger.info(f"  Loaded: {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df
            
        except Exception as e:
            logger.error(f"Error reading Parquet file: {str(e)}")
            raise
    
    @staticmethod
    def _read_json(file_path: Path) -> pl.DataFrame:
        """Read JSON file using Polars native reader."""
        try:
            # Try reading as NDJSON first (newline-delimited JSON)
            polars_df = pl.read_ndjson(file_path)
            logger.info(f"  Loaded (NDJSON): {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
            logger.info("")
            return polars_df
            
        except Exception:
            # Fallback to regular JSON via pandas
            try:
                logger.info("  NDJSON failed, trying regular JSON format...")
                logger.info("")
                with open(file_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                
                # Handle different JSON structures
                if isinstance(json_data, list):
                    pandas_df = pd.DataFrame(json_data)
                elif isinstance(json_data, dict):
                    pandas_df = pd.DataFrame([json_data])
                else:
                    raise ValueError("JSON must be a list of objects or a single object")
                
                polars_df = pl.from_pandas(pandas_df)
                logger.info(f"  Loaded (JSON): {polars_df.shape[0]} rows, {polars_df.shape[1]} columns")
                logger.info("")
                return polars_df
                
            except Exception as e:
                logger.error(f"Error reading JSON file: {str(e)}")
                raise
    
    @staticmethod
    def get_excel_sheets(file_path: Path) -> List[str]:
        """
        Get list of sheet names in Excel file.
        
        Args:
            file_path (Path): Path to Excel file
            
        Returns:
            List[str]: List of sheet names
        """
        try:
            with pd.ExcelFile(file_path) as xls:
                sheets = xls.sheet_names
                logger.info(f"Excel sheets in {file_path.name}: {sheets}")
                logger.info("")
                return sheets
                
        except Exception as e:
            logger.error(f"Error getting Excel sheets: {str(e)}")
            return []


# =============================================================================
# DATA NORMALIZATION
# =============================================================================

class DataNormalizer:
    """
    Conservative in-memory normalizer for cross-source dataset comparisons.

    Safety principles:
    - Never writes back to source files
    - Applies only high-confidence transformations
    - Logs what was normalized
    """

    NULL_TOKENS = {"", "-", "--", "NULL", "NONE", "NAN", "NA", "N/A"}
    SENTINEL_DATE_PREFIXES = {'0001-01-01'}
    DATE_FORMATS = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y%m%d",
        "%d %b %Y",
        "%d %B %Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
    )

    @staticmethod
    def normalize_dataframe(
        df: pl.DataFrame,
        dataset_name: str,
        normalize_numeric_strings: bool = True,
        strip_leading_zeros_numeric: bool = True,
        convert_dates: bool = True
    ) -> Tuple[pl.DataFrame, Dict[str, Any]]:
        """Normalize values and return normalized dataframe + report."""
        report: Dict[str, Any] = {
            "dataset": dataset_name,
            "trimmed_columns": [],
            "numeric_normalized_columns": [],
            "null_standardized_columns": []
        }

        # 1) Trim column names safely.
        normalized_col_names = DataNormalizer._normalize_column_names(df.columns)
        if normalized_col_names != df.columns:
            renamed_df = df.select([
                pl.col(old).alias(new)
                for old, new in zip(df.columns, normalized_col_names)
            ])
            changed = [
                {"from": old, "to": new}
                for old, new in zip(df.columns, normalized_col_names)
                if old != new
            ]
            report["trimmed_columns"] = changed
            df = renamed_df

        # 2) Base string normalization for every column (trim + null tokens + optional numeric canonicalization).
        for col in df.columns:
            base_result = DataNormalizer._normalize_string_column(
                df,
                col,
                normalize_numeric_strings,
                strip_leading_zeros_numeric,
                convert_dates,
            )
            df = base_result["df"]
            if base_result["numeric_changed"]:
                report["numeric_normalized_columns"].append(col)
            if base_result["null_changed"]:
                report["null_standardized_columns"].append(col)

        return df, report

    @staticmethod
    def _normalize_column_names(columns: List[str]) -> List[str]:
        """Trim and deduplicate column names."""
        dedup_counter: Dict[str, int] = {}
        result: List[str] = []
        for original in columns:
            cleaned = str(original).strip()
            if cleaned in dedup_counter:
                dedup_counter[cleaned] += 1
                cleaned = f"{cleaned}_{dedup_counter[cleaned]}"
            else:
                dedup_counter[cleaned] = 0
            result.append(cleaned)
        return result

    @staticmethod
    def _normalize_string_column(
        df: pl.DataFrame,
        col: str,
        normalize_numeric_strings: bool,
        strip_leading_zeros_numeric: bool,
        convert_dates: bool
    ) -> Dict[str, Any]:
        """Apply base value normalization to one column."""
        before_values = df.select(pl.col(col).cast(pl.Utf8, strict=False)).to_series().to_list()

        def normalize_value(value: Any) -> Optional[str]:
            if value is None:
                return None
            raw = str(value)
            stripped = raw.strip()
            if stripped.upper() in DataNormalizer.NULL_TOKENS:
                return None
            # Sentinel date handling and parsing is part of the date conversion logic.
            if convert_dates and any(stripped.startswith(prefix) for prefix in DataNormalizer.SENTINEL_DATE_PREFIXES):
                return None
            canonical: Optional[str] = None
            if normalize_numeric_strings:
                canonical = DataNormalizer._canonical_numeric_string(
                    stripped,
                    strip_leading_zeros_numeric=strip_leading_zeros_numeric
                )
            if convert_dates:
                # Conservative date normalization with solid fallback:
                # if parse is uncertain, keep the original value unchanged.
                date_candidate = canonical if canonical is not None else stripped
                parsed_date = DataNormalizer._try_parse_date_to_iso(date_candidate)
                if parsed_date is not None:
                    return parsed_date
            return canonical if canonical is not None else stripped

        after_values = [normalize_value(value) for value in before_values]

        numeric_changed = False
        null_changed = False
        for before, after in zip(before_values, after_values):
            if before is None and after is None:
                continue
            before_str = None if before is None else str(before)
            if before_str != after:
                if after is None:
                    null_changed = True
                else:
                    # Count as numeric change only when source looked numeric.
                    compact_before = before_str.strip() if before_str is not None else ""
                    if DataNormalizer._looks_numeric(compact_before):
                        numeric_changed = True
            if numeric_changed and null_changed:
                break

        normalized_df = df.with_columns(pl.Series(col, after_values))
        return {"df": normalized_df, "numeric_changed": numeric_changed, "null_changed": null_changed}

    @staticmethod
    def _looks_numeric(value: str) -> bool:
        """Check if a string resembles a numeric token."""
        if not value:
            return False
        try:
            float(value.replace(",", "."))
            return True
        except Exception:
            return False

    @staticmethod
    def _canonical_numeric_string(value: str, strip_leading_zeros_numeric: bool) -> Optional[str]:
        """
        Canonicalize numeric-like strings conservatively.

        Example: "12.0" -> "12" to avoid false mismatches.
        """
        if not value:
            return None
        candidate = value.replace(",", ".")
        try:
            number = float(candidate)
        except Exception:
            return None

        # Normalize clear integer-like float tokens.
        if "." in candidate and number.is_integer():
            return str(int(number))
        # Optional: normalize pure-digit leading zeros (e.g. 0042 -> 42).
        if strip_leading_zeros_numeric and value.isdigit() and len(value) > 1:
            return value.lstrip("0") or "0"
        return None

    @staticmethod
    def _try_parse_date_to_iso(value: str) -> Optional[str]:
        """
        Parse date-like values conservatively and return ISO date (YYYY-MM-DD).
        Returns None when parsing is uncertain/unavailable.
        """
        if not value:
            return None

        # 1) Explicit textual formats only (no permissive fuzzy parser).
        for fmt in DataNormalizer.DATE_FORMATS:
            try:
                dt = datetime.strptime(value, fmt)
                return dt.strftime("%Y-%m-%d")
            except Exception:
                continue

        # 2) Excel serial dates in a realistic range.
        try:
            numeric = float(value)
            if numeric.is_integer():
                serial = int(numeric)
                if 20000 <= serial <= 70000:
                    dt = datetime(1899, 12, 30) + timedelta(days=serial)
                    return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

        return None

# =============================================================================
# DATA ANALYSIS SECTION
# =============================================================================

class DataAnalyzer:
    """
    Data analyzer for comparing datasets structure and content.
    Performs exploratory analysis without assumptions about data content.
    """
    
    def __init__(self, primary_df: pl.DataFrame, secondary_df: pl.DataFrame, 
                 primary_name: str = "Primary", secondary_name: str = "Secondary"):
        """
        Initialize analyzer with two datasets.
        
        Args:
            primary_df (pl.DataFrame): First dataset
            secondary_df (pl.DataFrame): Second dataset
            primary_name (str): Name for primary dataset in reports
            secondary_name (str): Name for secondary dataset in reports
        """
        self.primary_df = primary_df
        self.secondary_df = secondary_df
        self.primary_name = primary_name
        self.secondary_name = secondary_name
        
    def analyze_structure(self) -> Dict[str, Any]:
        """
        Analyze and compare basic structure of both datasets.
        
        Returns:
            Dict[str, Any]: Structure analysis results
        """
        logger.info("="*60)
        logger.info("STRUCTURAL ANALYSIS")
        logger.info("="*60)
        logger.info("")

        # Basic shape comparison
        primary_shape = self.primary_df.shape
        secondary_shape = self.secondary_df.shape
        
        logger.info("Dataset Dimensions:")
        logger.info("")
        logger.info(f"  {self.primary_name:<12}: {primary_shape[0]:>8,} rows × {primary_shape[1]:>3} columns")
        logger.info(f"  {self.secondary_name:<12}: {secondary_shape[0]:>8,} rows × {secondary_shape[1]:>3} columns")
        logger.info("")

        # Row comparison
        row_diff = primary_shape[0] - secondary_shape[0]
        col_diff = primary_shape[1] - secondary_shape[1]
        
        logger.info("")
        logger.info("Dimension Differences:")
        logger.info("")
        if row_diff == 0:
            logger.info(f"  Rows: ✓ Same number of rows ({primary_shape[0]:,})")
        else:
            logger.info(f"  Rows: ⚠ Difference of {abs(row_diff):,} rows ({'+' if row_diff > 0 else ''}{row_diff:,})")
            
        if col_diff == 0:
            logger.info(f"  Columns: ✓ Same number of columns ({primary_shape[1]})")
        else:
            logger.info(f"  Columns: ⚠ Difference of {abs(col_diff)} columns ({'+' if col_diff > 0 else ''}{col_diff})")
        
        return {
            'primary_shape': primary_shape,
            'secondary_shape': secondary_shape,
            'row_difference': row_diff,
            'column_difference': col_diff,
            'same_dimensions': row_diff == 0 and col_diff == 0
        }
    
    def analyze_columns(self) -> Dict[str, Any]:
        """
        Analyze and compare column structure between datasets.
        
        Returns:
            Dict[str, Any]: Column analysis results
        """
        logger.info("")
        logger.info("="*60)
        logger.info("COLUMN STRUCTURE ANALYSIS")
        logger.info("="*60)
        logger.info("")
        
        # Get column lists
        primary_cols = set(self.primary_df.columns)
        secondary_cols = set(self.secondary_df.columns)
        
        # Find common and different columns
        common_cols = primary_cols & secondary_cols
        primary_only = primary_cols - secondary_cols
        secondary_only = secondary_cols - primary_cols
        
        # Analyze column duplicates for each dataset
        primary_col_list = list(self.primary_df.columns)
        secondary_col_list = list(self.secondary_df.columns)
        
        # Find duplicates in primary
        primary_unique_cols = len(set(primary_col_list))
        primary_total_cols = len(primary_col_list)
        primary_duplicates = primary_total_cols - primary_unique_cols
        
        # Find duplicates in secondary
        secondary_unique_cols = len(set(secondary_col_list))
        secondary_total_cols = len(secondary_col_list)
        secondary_duplicates = secondary_total_cols - secondary_unique_cols
        
        logger.info("Column Overview:")
        logger.info("")
        logger.info(f"  {self.primary_name} columns:   {primary_total_cols:>3} total, {primary_unique_cols:>3} unique, {primary_duplicates:>3} duplicates")
        logger.info(f"  {self.secondary_name} columns:  {secondary_total_cols:>3} total, {secondary_unique_cols:>3} unique, {secondary_duplicates:>3} duplicates")
        logger.info("")
        
        # Show duplicate column details if any exist
        if primary_duplicates > 0:
            primary_duplicate_names = self._find_duplicate_columns(primary_col_list)
            logger.info(f"  {self.primary_name} duplicate columns:")
            for col_name, count in primary_duplicate_names.items():
                logger.info(f"    '{col_name}' appears {count} times")
            logger.info("")
        
        if secondary_duplicates > 0:
            secondary_duplicate_names = self._find_duplicate_columns(secondary_col_list)
            logger.info(f"  {self.secondary_name} duplicate columns:")
            for col_name, count in secondary_duplicate_names.items():
                logger.info(f"    '{col_name}' appears {count} times")
            logger.info("")
        
        logger.info(f"  Common columns:      {len(common_cols):>3}")
        logger.info(f"  {self.primary_name} only:       {len(primary_only):>3}")
        logger.info(f"  {self.secondary_name} only:      {len(secondary_only):>3}")
        
        # Display column differences (only if there are differences)
        if primary_only:
            logger.info("")
            logger.info(f"Columns only in {self.primary_name}: {', '.join(sorted(primary_only)[:10])}")
            if len(primary_only) > 10:
                logger.info(f"  ... and {len(primary_only) - 10} more")
        
        if secondary_only:
            logger.info("")
            logger.info(f"Columns only in {self.secondary_name}: {', '.join(sorted(secondary_only)[:10])}")
            if len(secondary_only) > 10:
                logger.info(f"  ... and {len(secondary_only) - 10} more")
        
        # Column alignment assessment
        alignment_score = len(common_cols) / max(len(primary_cols), len(secondary_cols)) * 100
        
        logger.info("")
        logger.info("Column Alignment Assessment:")
        logger.info("")
        if alignment_score == 100:
            logger.info(f"  ✓ Perfect alignment: {alignment_score:.1f}% (identical column structure)")
        elif alignment_score >= 80:
            logger.info(f"  ✓ Good alignment: {alignment_score:.1f}% (mostly similar structure)")
        elif alignment_score >= 50:
            logger.info(f"  ⚠ Moderate alignment: {alignment_score:.1f}% (some structural differences)")
        else:
            logger.info(f"  ✗ Poor alignment: {alignment_score:.1f}% (significant structural differences)")
        
        # Prepare duplicate analysis results
        primary_duplicate_info = {}
        secondary_duplicate_info = {}
        
        if primary_duplicates > 0:
            primary_duplicate_info = self._find_duplicate_columns(primary_col_list)
        
        if secondary_duplicates > 0:
            secondary_duplicate_info = self._find_duplicate_columns(secondary_col_list)
        
        return {
            'primary_columns': list(primary_cols),
            'secondary_columns': list(secondary_cols),
            'primary_total_columns': primary_total_cols,
            'primary_unique_columns': primary_unique_cols,
            'primary_duplicate_columns': primary_duplicates,
            'primary_duplicate_details': primary_duplicate_info,
            'secondary_total_columns': secondary_total_cols,
            'secondary_unique_columns': secondary_unique_cols,
            'secondary_duplicate_columns': secondary_duplicates,
            'secondary_duplicate_details': secondary_duplicate_info,
            'common_columns': list(common_cols),
            'primary_only_columns': list(primary_only),
            'secondary_only_columns': list(secondary_only),
            'alignment_score': alignment_score,
            'perfect_alignment': alignment_score == 100
        }
    
    def analyze_data_types(self) -> Dict[str, Any]:
        """
        Analyze and compare data types for common columns.
        
        Returns:
            Dict[str, Any]: Data type analysis results
        """
        logger.info("")
        logger.info("="*60)
        logger.info("DATA TYPE ANALYSIS")
        logger.info("="*60)
        logger.info("")
        
        # Get common columns
        common_cols = set(self.primary_df.columns) & set(self.secondary_df.columns)
        
        if not common_cols:
            logger.info("No common columns found for data type comparison.")
            return {'common_columns': [], 'type_matches': [], 'type_mismatches': []}
        
        # Compare data types for common columns
        type_matches = []
        type_mismatches = []
        
        logger.info(f"Data Type Comparison for Common Columns ({len(common_cols)}):")
        logger.info("")
        logger.info(f"{'Column':<30} {'Primary Type':<30} {'Secondary Type':<30} {'Match':<10}")
        logger.info("-" * 100)
        logger.info("")
        
        for col in sorted(common_cols):
            primary_type = str(self.primary_df[col].dtype)
            secondary_type = str(self.secondary_df[col].dtype)
            
            match = primary_type == secondary_type
            match_symbol = "✓" if match else "✗"
            
            logger.info(f"{col:<30} {primary_type:<30} {secondary_type:<30} {match_symbol:<10}")
            
            if match:
                type_matches.append({
                    'column': col,
                    'type': primary_type
                })
            else:
                type_mismatches.append({
                    'column': col,
                    'primary_type': primary_type,
                    'secondary_type': secondary_type
                })
        
        # Summary
        logger.info("")
        logger.info("Data Type Summary:")
        logger.info("")
        logger.info(f"  Type matches:    {len(type_matches):>3} columns")
        logger.info(f"  Type mismatches: {len(type_mismatches):>3} columns")
        
        if type_mismatches:
            type_compatibility = len(type_matches) / len(common_cols) * 100
            logger.info(f"  Type compatibility: {type_compatibility:.1f}%")
            
            if type_compatibility < 100:
                logger.info("")
                logger.info("⚠ Type mismatches may affect data comparison accuracy.")
        
        return {
            'common_columns': list(common_cols),
            'type_matches': type_matches,
            'type_mismatches': type_mismatches,
            'type_compatibility': len(type_matches) / len(common_cols) * 100 if common_cols else 0
        }
    
    def _find_duplicate_columns(self, column_list: List[str]) -> Dict[str, int]:
        """
        Find duplicate column names and their counts.
        
        Args:
            column_list (List[str]): List of column names (may contain duplicates)
            
        Returns:
            Dict[str, int]: Dictionary with column names that appear more than once and their counts
        """
        from collections import Counter
        
        column_counts = Counter(column_list)
        # Return only columns that appear more than once
        duplicates = {col: count for col, count in column_counts.items() if count > 1}
        return duplicates

    
    def run_full_analysis(self) -> Dict[str, Any]:
        """
        Run complete analysis comparing both datasets.
        
        Returns:
            Dict[str, Any]: Complete analysis results
        """
        logger.info("")
        
        results = {}
        results['structure'] = self.analyze_structure()
        results['columns'] = self.analyze_columns()
        results['data_types'] = self.analyze_data_types()
        
        logger.info("")
        logger.info("="*60)
        logger.info("STRUCTURAL ANALYSIS COMPLETE")
        logger.info("="*60)
        logger.info("")
        
        return results


# =============================================================================
# CONTENT COMPARISON ENGINE
# =============================================================================

class ContentComparator:
    """
    A content comparison engine for dataset analysis.
    Performs row-by-row matching and value difference detection.
    """
    
    def __init__(self, primary_df: pl.DataFrame, secondary_df: pl.DataFrame,
                 primary_name: str = "Primary", secondary_name: str = "Secondary"):
        """
        Initialize content comparator.
        
        Args:
            primary_df (pl.DataFrame): First dataset
            secondary_df (pl.DataFrame): Second dataset
            primary_name (str): Name for primary dataset
            secondary_name (str): Name for secondary dataset
        """
        self.primary_df = primary_df
        self.secondary_df = secondary_df
        self.primary_name = primary_name
        self.secondary_name = secondary_name
        self.common_columns = [col for col in primary_df.columns if col in secondary_df.columns]

    @staticmethod
    def _null_safe_difference_mask(left_expr: pl.Expr, right_expr: pl.Expr) -> pl.Expr:
        """
        Build a null-safe difference mask.

        This treats values as different when one side is null and the other is not,
        while keeping null/null pairs as equal.
        """
        return (
            (left_expr.is_null() ^ right_expr.is_null()) |
            (
                left_expr.is_not_null() &
                right_expr.is_not_null() &
                (left_expr != right_expr)
            )
        )
    
    def find_matching_records(self, key_columns: List[str] = None) -> Dict[str, Any]:
        """
        Find matching records between datasets based on key columns.
        
        Args:
            key_columns (List[str]): Columns to use as composite key
            
        Returns:
            Dict[str, Any]: Matching analysis results
        """
        logger.info("")
        logger.info("="*60)
        logger.info("CONTENT MATCHING ANALYSIS")
        logger.info("="*60)
        
        if not self.common_columns:
            logger.warning("No common columns found for key detection; fallback comparison will be used.")
            return self._compare_without_keys()
        
        # Auto-detect key columns if not provided
        if not key_columns:
            key_columns = self._detect_key_columns()
            
        if not key_columns:
            logger.warning("Auto key detection failed; switching to interactive key creator mode.")
            key_columns = self._interactive_key_creator()

        if not key_columns:
            logger.warning("Interactive key creator did not produce a valid key; fallback comparison will be used.")
            return self._compare_without_keys()
        
        logger.info(f"Using key columns: {', '.join(key_columns)}")
        logger.info("")
        
        try:
            # Create composite keys
            primary_with_key = self._create_composite_key(self.primary_df, key_columns, "primary_key")
            secondary_with_key = self._create_composite_key(self.secondary_df, key_columns, "secondary_key")
            
            # Find matches and differences using key values as-is.
            primary_raw_keys = primary_with_key.select("primary_key").to_series().to_list()
            secondary_raw_keys = secondary_with_key.select("secondary_key").to_series().to_list()
            
            primary_keys = {key for key in primary_raw_keys if key is not None}
            secondary_keys = {key for key in secondary_raw_keys if key is not None}

            matching_keys = primary_keys & secondary_keys
            primary_only_keys = primary_keys - secondary_keys
            secondary_only_keys = secondary_keys - primary_keys
            
            logger.info("Record Matching Results:")
            logger.info(f"  Total {self.primary_name} records:   {len(primary_keys):>6,}")
            logger.info(f"  Total {self.secondary_name} records:  {len(secondary_keys):>6,}")
            logger.info(f"  Matching records:        {len(matching_keys):>6,}")
            logger.info(f"  {self.primary_name} only:           {len(primary_only_keys):>6,}")
            logger.info(f"  {self.secondary_name} only:          {len(secondary_only_keys):>6,}")
            logger.info("")
            
            # Show sample keys from each dataset.
            logger.info("Sample Keys:")
            primary_raw_samples = primary_raw_keys[:5]
            secondary_raw_samples = secondary_raw_keys[:5]
            
            logger.info(f"  {self.primary_name} raw key samples: {primary_raw_samples}")
            logger.info(f"  {self.secondary_name} raw key samples: {secondary_raw_samples}")
            
            # Show samples of matching keys if any exist
            if matching_keys:
                matching_samples = list(matching_keys)[:5]
                logger.info(f"  Matching key samples: {matching_samples}")
            
            # Show samples of unique keys if any exist
            if primary_only_keys:
                primary_only_samples = list(primary_only_keys)#[:3]
                logger.info(f"  {self.primary_name} only samples: {primary_only_samples}")
            
            if secondary_only_keys:
                secondary_only_samples = list(secondary_only_keys)#[:3]
                logger.info(f"  {self.secondary_name} only samples: {secondary_only_samples}")
            
            logger.info("")
            
            # Calculate match percentage
            total_unique_records = len(primary_keys | secondary_keys)
            match_percentage = (len(matching_keys) / total_unique_records * 100) if total_unique_records > 0 else 0
            
            if match_percentage >= 90:
                logger.info(f"✓ Excellent match: {match_percentage:.1f}% records align")
            elif match_percentage >= 70:
                logger.info(f"✓ Good match: {match_percentage:.1f}% records align")
            elif match_percentage >= 50:
                logger.info(f"⚠ Moderate match: {match_percentage:.1f}% records align")
            else:
                logger.info(f"✗ Poor match: {match_percentage:.1f}% records align")
            
            # Reuse keyed DataFrames for content analysis.
            primary_keyed = primary_with_key
            secondary_keyed = secondary_with_key
            
            # Analyze value differences for matching records
            value_differences = self._analyze_value_differences(
                primary_keyed, secondary_keyed, matching_keys, key_columns
            )
            
            return {
                'key_columns': key_columns,
                'matching_records': len(matching_keys),
                'primary_only_records': len(primary_only_keys),
                'secondary_only_records': len(secondary_only_keys),
                'match_percentage': match_percentage,
                'value_differences': value_differences
            }
            
        except Exception as e:
            logger.error(f"Error in record matching: {str(e)}")
            return {}
    
    def _detect_key_columns(self) -> List[str]:
        """
        Auto-detect key columns with proper uniqueness validation.
        
        Strategy:
        1. Find single unique column (no nulls, all unique values)
        2. Try composite keys (2 columns max) from all common columns
        3. Fallback to no keys
        
        Returns:
            List[str]: Key column names or empty list if none found
        """
        if not self.common_columns:
            logger.info("  No common columns found for key detection")
            return []
        
        logger.info("  Detecting key columns...")
        
        # Step 1: Search for single unique column
        single_key = self._find_single_unique_key()
        if single_key:
            logger.info(f"  ✅ Found single unique key: '{single_key}'")
            return [single_key]
        
        # Step 2: Search for composite keys (2 columns) from all columns
        composite_key = self._find_composite_key()
        if composite_key:
            logger.info(f"  ✅ Found composite key: {' + '.join(composite_key)}")
            return composite_key
        
        # Step 3: Fallback - no suitable keys found
        logger.info("  ⚠ No unique key columns found - using fallback comparison")
        return []
    
    def _find_single_unique_key(self) -> Optional[str]:
        """Find a single column that uniquely identifies each row."""
        primary_rows = self.primary_df.height
        secondary_rows = self.secondary_df.height
        
        for col in self.common_columns:
            try:
                # # Check primary dataset
                primary_nulls = self.primary_df.select(pl.col(col).is_null().sum()).item()
                primary_unique = self.primary_df.select(pl.col(col).n_unique()).item()
                
                # Check secondary dataset  
                secondary_nulls = self.secondary_df.select(pl.col(col).is_null().sum()).item()
                secondary_unique = self.secondary_df.select(pl.col(col).n_unique()).item()
                
                # Strict mode: no nulls and all values unique.
                primary_is_key = (primary_nulls == 0) and (primary_unique == primary_rows)
                secondary_is_key = (secondary_nulls == 0) and (secondary_unique == secondary_rows)
                
                if primary_is_key and secondary_is_key:
                    logger.info(f"    Testing '{col}': ✅ Unique in both datasets")
                    return col
                    
            except Exception as e:
                logger.info(f"    Testing '{col}': ❌ Error - {str(e)}")
                continue
        
        return None
    
    def _interactive_key_creator(self) -> List[str]:
        """
        Terminal-only interactive fallback for key selection and key creation.

        Flow:
        1) Step 1: auto-check single key (strict: zero nulls + unique in both datasets)
        2) Step 2: auto-check pair keys (strict: zero nulls + unique in both datasets)
        3) Step 3: manual key selection from suggested columns (<10% nulls)
        4) Validate: nulls first, then duplicates
        5) If nulls exist -> ask to fill with conventional value 'na'
        6) If duplicates exist -> ask to either suffix (_1,_2,...) or remove duplicates
        """
        if not sys.stdin.isatty():
            print("\n[KEY CREATOR] Interactive mode is not available in non-interactive execution.")
            return []

        print("\n" + "=" * 72)
        print("KEY CREATOR (INTERACTIVE FALLBACK)")
        print("=" * 72)
        print("Auto key detection failed for strict rules.")

        # ------------------------------------------------------------
        # Step 1: auto-check single key
        # ------------------------------------------------------------
        single_key = self._find_single_unique_key()
        if single_key:
            print("\n[Step 1] Single key found automatically.")
            print(f"  Key column: {single_key}")
            logger.info(f"Interactive key creator accepted auto single key: {single_key}")
            return [single_key]
        print("\n[Step 1] Auto-check single key: FAILED (no strict single key).")

        # ------------------------------------------------------------
        # Step 2: auto-check pair keys
        # ------------------------------------------------------------
        pair_key = self._find_composite_key()
        if pair_key:
            print("\n[Step 2] Pair key found automatically.")
            print(f"  Key columns: {pair_key[0]} + {pair_key[1]}")
            logger.info(f"Interactive key creator accepted auto pair key: {pair_key}")
            return pair_key
        print("\n[Step 2] Auto-check pair keys: FAILED (no strict pair key).")

        # ------------------------------------------------------------
        # Step 3: propose manual selection
        # ------------------------------------------------------------
        null_threshold = 0.10
        suggested_columns = self._get_columns_with_low_null_rate(null_threshold)
        if not suggested_columns:
            print(f"\n[Step 3] No columns found with < {null_threshold*100:.0f}% nulls in BOTH datasets.")
            print("  As a fallback, you can still try ANY common column.")
            suggested_columns = self.common_columns

        print("\n[Step 3] Suggested key columns (lowest null risk):")
        for idx, col in enumerate(suggested_columns, 1):
            print(f"  {idx:>3}. {col}")
        print("")

        selected = self._prompt_key_column_selection(suggested_columns)
        if not selected:
            print("\n[KEY CREATOR] No valid columns selected. Stopping interactive key creation.")
            return []

        print(f"\n[KEY CREATOR] Selected key columns: {selected}")

        # ------------------------------------------------------------
        # Validate key columns: NULLS (must be zero unless user fills)
        # ------------------------------------------------------------
        primary_null_counts = self._get_column_null_counts(self.primary_df, selected)
        secondary_null_counts = self._get_column_null_counts(self.secondary_df, selected)

        has_nulls = any(v > 0 for v in primary_null_counts.values()) or any(v > 0 for v in secondary_null_counts.values())
        if has_nulls:
            print("\n" + "!" * 72)
            print("WARNING: Key columns MUST have ZERO null values.")
            print("If you fill nulls, matching can become incorrect and duplicates can be introduced.")
            print("Continue ONLY if you fully understand the risk.")
            print("!" * 72)

            print("\nPrimary null counts (selected key columns):")
            for col in selected:
                print(f"  - {col}: {primary_null_counts[col]}")
            print("\nSecondary null counts (selected key columns):")
            for col in selected:
                print(f"  - {col}: {secondary_null_counts[col]}")

            print("")
            print("Decision required:")
            answer = input("Do you want to fill null/blank key values with the conventional value 'na'? Type YES to continue: ").strip().upper()
            if answer != "YES":
                print("\n[KEY CREATOR] User rejected null filling. Stopping interactive key creation.")
                return []

            # Minimal indispensable logging: action taken.
            logger.warning("Interactive key creator: filling NULLs in selected key columns with 'na'.")
            self.primary_df = self._fill_key_nulls(self.primary_df, selected, "na")
            self.secondary_df = self._fill_key_nulls(self.secondary_df, selected, "na")

            primary_null_counts = self._get_column_null_counts(self.primary_df, selected)
            secondary_null_counts = self._get_column_null_counts(self.secondary_df, selected)
            if any(v > 0 for v in primary_null_counts.values()) or any(v > 0 for v in secondary_null_counts.values()):
                print("\n[KEY CREATOR] Nulls still exist AFTER fill. Cannot safely build a strict key.")
                return []

        # ------------------------------------------------------------
        # Validate key uniqueness: DUPLICATES (must be zero unless user resolves)
        # ------------------------------------------------------------
        dup_primary = self._get_duplicate_key_groups(self.primary_df, selected)
        dup_secondary = self._get_duplicate_key_groups(self.secondary_df, selected)

        has_dups = dup_primary.height > 0 or dup_secondary.height > 0
        if has_dups:
            print("\n" + "#" * 72)
            print("HIGH RISK: Duplicate key values detected.")
            print("This creates one-to-many matches and can massively distort the detail report.")
            print("You MUST resolve duplicates before proceeding.")
            print("#" * 72)

            self._print_duplicate_groups_side_by_side(dup_primary, dup_secondary)

            print("\nChoose a resolution strategy:")
            print("  1) Deduplicate by appending suffixes to the FIRST key column (_1, _2, ...).")
            print("  2) Remove duplicate records for duplicate keys (keep FIRST occurrence).")
            print("")
            answer = input("Type 1 or 2 (or CANCEL to abort): ").strip().upper()
            if answer == "CANCEL" or answer not in {"1", "2"}:
                print("\n[KEY CREATOR] User aborted duplicate resolution. Stopping interactive key creation.")
                return []

            logger.warning("Interactive key creator: resolving duplicate keys (user requested action).")
            if answer == "1":
                self.primary_df = self._dedup_by_suffix_pandas(self.primary_df, selected)
                self.secondary_df = self._dedup_by_suffix_pandas(self.secondary_df, selected)
            else:
                self.primary_df = self._dedup_by_remove_pandas(self.primary_df, selected)
                self.secondary_df = self._dedup_by_remove_pandas(self.secondary_df, selected)

            # Re-check strictness
            primary_null_counts = self._get_column_null_counts(self.primary_df, selected)
            secondary_null_counts = self._get_column_null_counts(self.secondary_df, selected)
            if any(v > 0 for v in primary_null_counts.values()) or any(v > 0 for v in secondary_null_counts.values()):
                print("\n[KEY CREATOR] Nulls exist AFTER duplicate resolution. Cannot safely proceed.")
                return []

            dup_primary = self._get_duplicate_key_groups(self.primary_df, selected)
            dup_secondary = self._get_duplicate_key_groups(self.secondary_df, selected)
            if dup_primary.height > 0 or dup_secondary.height > 0:
                print("\n[KEY CREATOR] Duplicates still exist AFTER resolution. Cannot safely proceed.")
                return []

        print("\n[KEY CREATOR] SUCCESS: strict key built (no nulls, unique in both datasets).")
        logger.info(f"Interactive key creator selected strict key columns: {selected}")
        return selected

    def _prompt_key_column_selection(self, columns: List[str]) -> List[str]:
        """Ask user to provide key columns as numbered indexes separated by '-'."""
        raw = input("Insert column numbers for key (e.g. 1-3 or 7): ").strip()
        if not raw:
            return []

        parts = [part.strip() for part in raw.split("-") if part.strip()]
        if not parts:
            return []

        indexes: List[int] = []
        for part in parts:
            if not part.isdigit():
                return []
            idx = int(part)
            if idx < 1 or idx > len(columns):
                return []
            indexes.append(idx - 1)

        # Keep insertion order, remove duplicates.
        seen = set()
        selected: List[str] = []
        for idx in indexes:
            if idx not in seen:
                seen.add(idx)
                selected.append(columns[idx])
        return selected

    def _fill_key_nulls(self, df: pl.DataFrame, key_columns: List[str], fill_value: str) -> pl.DataFrame:
        """Fill null/blank-like values only on selected key columns."""
        # Align with DataNormalizer tokens to prevent sentinel values from surviving.
        null_tokens = {"", "-", "--", "NULL", "NONE", "NAN", "NA", "N/A"}
        updated = df
        for col in key_columns:
            updated = updated.with_columns(
                pl.when(
                    pl.col(col).is_null() |
                    pl.col(col).cast(pl.Utf8, strict=False).str.strip_chars().is_in(list(null_tokens))
                )
                .then(pl.lit(fill_value))
                .otherwise(pl.col(col).cast(pl.Utf8, strict=False).str.strip_chars())
                .alias(col)
            )
        return updated

    def _get_column_null_counts(self, df: pl.DataFrame, columns: List[str]) -> Dict[str, int]:
        """Return null counts per column (exactly on Polars nulls)."""
        counts: Dict[str, int] = {}
        for col in columns:
            counts[col] = int(df.select(pl.col(col).is_null().sum()).item())
        return counts

    def _get_columns_with_low_null_rate(self, threshold: float) -> List[str]:
        """Pick columns where BOTH datasets have null rate <= threshold."""
        candidates: List[str] = []
        primary_rows = self.primary_df.height
        secondary_rows = self.secondary_df.height
        
        for col in self.common_columns:
            try:
                p_nulls = int(self.primary_df.select(pl.col(col).is_null().sum()).item())
                s_nulls = int(self.secondary_df.select(pl.col(col).is_null().sum()).item())
                p_rate = (p_nulls / primary_rows) if primary_rows > 0 else 0.0
                s_rate = (s_nulls / secondary_rows) if secondary_rows > 0 else 0.0
                if p_rate <= threshold and s_rate <= threshold:
                    candidates.append(col)
            except Exception:
                continue
        return candidates

    def _get_duplicate_key_groups(self, df: pl.DataFrame, key_columns: List[str]) -> pl.DataFrame:
        """Return duplicate groups for the strict composite key."""
        df_with_key = self._create_composite_key(df, key_columns, "Key_temp")
        dup = (
            df_with_key
            .filter(pl.col("Key_temp").is_not_null())
            .group_by("Key_temp")
            .agg(pl.len().alias("count"))
            .filter(pl.col("count") > 1)
            .sort("count", descending=True)
        )
        return dup

    def _print_duplicate_groups_side_by_side(self, dup_primary: pl.DataFrame, dup_secondary: pl.DataFrame) -> None:
        """Print duplicate groups in one combined table (primary and secondary counts side-by-side)."""
        if dup_primary.height == 0 and dup_secondary.height == 0:
            print("No duplicate key groups detected.")
            return

        p = dup_primary.rename({"count": "primary_count"}).to_pandas()
        s = dup_secondary.rename({"count": "secondary_count"}).to_pandas()

        # Outer merge to keep all keys.
        combined = p.merge(s, on="Key_temp", how="outer")
        if "primary_count" not in combined.columns:
            combined["primary_count"] = 0
        if "secondary_count" not in combined.columns:
            combined["secondary_count"] = 0
        combined = combined.fillna(0)
        combined = combined.sort_values(by="primary_count", ascending=False)

        print("")
        print("Duplicate key groups (side-by-side):")
        print(combined.to_string(index=False))

    def _dedup_by_suffix_pandas(self, df: pl.DataFrame, key_columns: List[str]) -> pl.DataFrame:
        """
        Deduplicate by appending _1, _2, ... to the LAST key column
        within each duplicate composite-key group.
        """
        pdf = df.to_pandas()
        # Suffix must be applied to the LAST column so the composite key
        # becomes like: aaaa-bb, aaaa-bb_1, ...
        first_col = key_columns[-1]

        key_temp = pdf[key_columns].astype(str).agg("-".join, axis=1)
        pdf["_key_temp"] = key_temp
        pdf["_dup_pos"] = pdf.groupby("_key_temp").cumcount()

        mask = pdf["_dup_pos"] > 0
        pdf.loc[mask, first_col] = pdf.loc[mask, first_col].astype(str) + "_" + pdf.loc[mask, "_dup_pos"].astype(str)

        pdf = pdf.drop(columns=["_key_temp", "_dup_pos"])
        return pl.from_pandas(pdf)

    def _dedup_by_remove_pandas(self, df: pl.DataFrame, key_columns: List[str]) -> pl.DataFrame:
        """Deduplicate by removing duplicate records (keep FIRST occurrence per key)."""
        pdf = df.to_pandas()
        key_temp = pdf[key_columns].astype(str).agg("-".join, axis=1)
        pdf["_key_temp"] = key_temp
        pdf = pdf.drop_duplicates(subset=["_key_temp"], keep="first")
        pdf = pdf.drop(columns=["_key_temp"])
        return pl.from_pandas(pdf)

    def _evaluate_key_columns(self, primary_df: pl.DataFrame, secondary_df: pl.DataFrame, key_columns: List[str]) -> Dict[str, int]:
        """Return null and duplicate statistics for selected key columns."""
        primary_with_key = self._create_composite_key(primary_df, key_columns, "_eval_key")
        secondary_with_key = self._create_composite_key(secondary_df, key_columns, "_eval_key")

        primary_rows = primary_with_key.height
        secondary_rows = secondary_with_key.height
        primary_nulls = primary_with_key.select(pl.col("_eval_key").is_null().sum()).item()
        secondary_nulls = secondary_with_key.select(pl.col("_eval_key").is_null().sum()).item()
        primary_unique = primary_with_key.select(pl.col("_eval_key").n_unique()).item()
        secondary_unique = secondary_with_key.select(pl.col("_eval_key").n_unique()).item()
                                
        return {
            "primary_nulls": int(primary_nulls),
            "secondary_nulls": int(secondary_nulls),
            "primary_duplicates": int(primary_rows - primary_unique),
            "secondary_duplicates": int(secondary_rows - secondary_unique),
        }

    def _print_duplicate_key_details(self, df: pl.DataFrame, key_columns: List[str], dataset_name: str) -> None:
        """Print duplicate key groups in terminal (interactive mode only)."""
        df_with_key = self._create_composite_key(df, key_columns, "Key_temp")
        dup = (
            df_with_key
            .filter(pl.col("Key_temp").is_not_null())
            .group_by("Key_temp")
            .agg(pl.len().alias("count"))
            .filter(pl.col("count") > 1)
            .sort("count", descending=True)
        )
        print("")
        print(f"--- DUPLICATE KEY DETAILS [{dataset_name}] ---")
        if dup.height == 0:
            print("No duplicate keys found.")
            return
        with pl.Config(tbl_rows=-1, tbl_cols=-1, fmt_str_lengths=100):
            print(dup)
    
    def _find_composite_key(self) -> Optional[List[str]]:
        """Find a combination of 2 columns that together uniquely identify rows."""
        from itertools import combinations

        primary_rows = self.primary_df.height
        secondary_rows = self.secondary_df.height

        for col_pair in combinations(self.common_columns, 2):
            try:
                col1, col2 = col_pair

                primary_composite = self.primary_df.select(
                    [
                        pl.concat_str(
                            [pl.col(col1).cast(pl.Utf8), pl.col(col2).cast(pl.Utf8)],
                            separator="-",
                        ).alias("composite")
                    ]
                )
                primary_nulls = primary_composite.select(pl.col("composite").is_null().sum()).item()
                primary_unique = primary_composite.select(pl.col("composite").n_unique()).item()

                secondary_composite = self.secondary_df.select(
                    [
                        pl.concat_str(
                            [pl.col(col1).cast(pl.Utf8), pl.col(col2).cast(pl.Utf8)],
                            separator="-",
                        ).alias("composite")
                    ]
                )
                secondary_nulls = secondary_composite.select(pl.col("composite").is_null().sum()).item()
                secondary_unique = secondary_composite.select(pl.col("composite").n_unique()).item()

                primary_is_key = (primary_nulls == 0) and (primary_unique == primary_rows)
                secondary_is_key = (secondary_nulls == 0) and (secondary_unique == secondary_rows)

                if primary_is_key and secondary_is_key:
                    logger.info(f"    Testing '{col1}' + '{col2}': ✅ Unique composite key")
                    return [col1, col2]

            except Exception as e:
                logger.info(
                    f"    Testing '{col_pair[0]}' + '{col_pair[1]}': ❌ Error - {str(e)}"
                )
                continue

        return None

    def _create_composite_key(self, df: pl.DataFrame, key_columns: List[str], key_name: str) -> pl.DataFrame:
        """Create composite key from multiple columns."""
        if len(key_columns) == 1:
            return df.with_columns(pl.col(key_columns[0]).cast(pl.Utf8).alias(key_name))
        else:
            # Combine multiple columns with separator
            key_expr = pl.concat_str([pl.col(col).cast(pl.Utf8) for col in key_columns], separator="-")
            return df.with_columns(key_expr.alias(key_name))
    
    def _analyze_value_differences(self, primary_keyed_df: pl.DataFrame, secondary_keyed_df: pl.DataFrame, 
                                 matching_keys: set, key_columns: List[str]) -> Dict[str, Any]:
        """Analyze value differences across ALL matching records."""
        if not matching_keys or len(matching_keys) == 0:
            return {}
        
        logger.info("")
        logger.info("="*60)
        logger.info("COMPREHENSIVE CONTENT ANALYSIS")
        logger.info("="*60)
        logger.info("")
        
        # Prepare for full dataset comparison
        key_name = "_temp_key"
        total_matching_records = len(matching_keys)
        differences_summary = {}
        total_field_comparisons = 0
        total_differences_found = 0
        
        # Create temporary key column for joining from keyed DataFrames.
        if len(key_columns) == 1:
            primary_keyed = primary_keyed_df.with_columns(pl.col("primary_key").alias(key_name))
            secondary_keyed = secondary_keyed_df.with_columns(pl.col("secondary_key").alias(key_name))
        else:
            primary_keyed = self._create_composite_key(primary_keyed_df, key_columns, key_name)
            secondary_keyed = self._create_composite_key(secondary_keyed_df, key_columns, key_name)
        
        
        logger.info(f"Comparing {total_matching_records:,} matching records across all common fields...")
        logger.info("")
        
        # Analyze each common column (except key columns)
        content_columns = [col for col in self.common_columns if col not in key_columns]
        
        
        if len(content_columns) == 0:
            logger.warning("⚠ No content columns to analyze (all common columns are keys)")
            return {
                'total_comparisons': 0,
                'total_differences': 0,
                'field_analysis': {},
                'data_integrity_score': 100.0
            }
        
        for col_idx, col in enumerate(content_columns, 1):
            try:
                logger.debug(f"Processing column {col_idx}/{len(content_columns)}: {col}")
                
                # Join datasets on key and compare this column
                comparison_df = primary_keyed.select([key_name, col]).join(
                    secondary_keyed.select([key_name, col]), 
                    on=key_name, 
                    how="inner",
                    suffix="_secondary"
                ).filter(pl.col(key_name).is_in(list(matching_keys)))
                
                if comparison_df.height == 0:
                    logger.debug(f"Skipping column '{col}' - no data after join")
                    continue
                
                # Count differences for this column
                col_secondary = f"{col}_secondary"
                differences_df = comparison_df.filter(
                    self._null_safe_difference_mask(pl.col(col), pl.col(col_secondary))
                )
                
                differences_count = differences_df.height
                total_comparisons_for_col = comparison_df.height
                total_field_comparisons += total_comparisons_for_col
                total_differences_found += differences_count
                
                logger.debug(f"Column '{col}': {total_comparisons_for_col} comparisons, {differences_count} differences")
                
                # Calculate percentage
                difference_percentage = (differences_count / total_comparisons_for_col * 100) if total_comparisons_for_col > 0 else 0
                
                # Collect examples of differences (up to 4 examples)
                examples = []
                if differences_count > 0 and differences_df.height > 0:
                    try:
                        # Get up to 4 examples of differences
                        sample_differences = differences_df.limit(4)
                        for row in sample_differences.iter_rows(named=True):
                            primary_val = row[col] if row[col] is not None else "NULL"
                            secondary_val = row[col_secondary] if row[col_secondary] is not None else "NULL"
                            examples.append({
                                'primary': str(primary_val),
                                'secondary': str(secondary_val)
                            })
                    except Exception as e:
                        logger.debug(f"Error collecting examples for column '{col}': {str(e)}")
                
                # Store summary with examples
                differences_summary[col] = {
                    'total_comparisons': total_comparisons_for_col,
                    'differences_found': differences_count,
                    'match_percentage': 100 - difference_percentage,
                    'difference_percentage': difference_percentage,
                    'examples': examples
                }
                
            except Exception as e:
                logger.debug(f"Error analyzing column '{col}': {str(e)}")
                continue
        
        # Calculate overall statistics
        overall_match_percentage = ((total_field_comparisons - total_differences_found) / total_field_comparisons * 100) if total_field_comparisons > 0 else 100
        
        # Generate professional summary report
        self._generate_content_analysis_report(
            differences_summary, 
            total_matching_records, 
            total_field_comparisons, 
            total_differences_found, 
            overall_match_percentage,
            len(content_columns)
        )
        
        return {
            'total_matching_records': total_matching_records,
            'total_field_comparisons': total_field_comparisons,
            'total_differences_found': total_differences_found,
            'overall_match_percentage': overall_match_percentage,
            'differences_by_column': differences_summary,
            'columns_analyzed': len(content_columns)
        }
    
    def _generate_content_analysis_report(self, differences_summary: Dict, total_matching_records: int,
                                        total_field_comparisons: int, total_differences_found: int,
                                        overall_match_percentage: float, columns_analyzed: int):
        """Generate a comprehensive, professional content analysis report."""
        
        logger.info("CONTENT ANALYSIS RESULTS:")
        logger.info("")
        
        # Overall Statistics
        logger.info("Overall Statistics:")
        logger.info(f"   Records Analyzed:      {total_matching_records:>10,}")
        logger.info(f"   Fields Analyzed:       {columns_analyzed:>10,}")
        logger.info(f"   Total Comparisons:     {total_field_comparisons:>10,}")
        logger.info(f"   Differences Found:     {total_differences_found:>10,}")
        logger.info("")
        
        # Overall Data Integrity Assessment
        if overall_match_percentage >= 99.5:
            integrity_status = "EXCELLENT"
            integrity_desc = "Data integrity is outstanding"
        elif overall_match_percentage >= 95.0:
            integrity_status = "GOOD"
            integrity_desc = "Data integrity is acceptable with minor discrepancies"
        elif overall_match_percentage >= 85.0:
            integrity_status = "MODERATE"
            integrity_desc = "Data integrity shows moderate discrepancies requiring review"
        else:
            integrity_status = "POOR"
            integrity_desc = "Data integrity shows significant discrepancies requiring immediate attention"
        
        logger.info(f"Data Integrity Score: {overall_match_percentage:.2f}% - {integrity_status}")
        logger.info(f"   {integrity_desc}")
        logger.info("")
        
        # Field-by-field analysis
        if differences_summary:
            logger.info("Field-by-Field Analysis:")
            logger.info("")
            
            # Sort by match percentage (worst first for attention)
            sorted_fields = sorted(differences_summary.items(), 
                                 key=lambda x: x[1]['match_percentage'])
            
            perfect_matches = []
            issues_found = []
            
            for field_name, stats in sorted_fields:
                match_pct = stats['match_percentage']
                diff_count = stats['differences_found']
                total_comp = stats['total_comparisons']
                
                if match_pct == 100.0:
                    perfect_matches.append(field_name)
                else:
                    issues_found.append({
                        'field': field_name,
                        'match_pct': match_pct,
                        'differences': diff_count,
                        'total': total_comp
                    })
            
            # Report issues first (most important)
            if issues_found:
                logger.info("   Fields with Discrepancies:")
                for issue in issues_found:
                    if issue['match_pct'] >= 99.0:
                        severity_dot = "🟢"
                    elif issue['match_pct'] >= 95.0:
                        severity_dot = "🟡"
                    else:
                        severity_dot = "🔴"
                    
                    logger.info(
                        f"   {issue['field']:<40} {issue['match_pct']:>6.2f}% match {severity_dot} ({issue['differences']:,}/{issue['total']:,} differ)"
                    )
                    
                    # Show examples if available
                    field_data = differences_summary.get(issue['field'], {})
                    examples = field_data.get('examples', [])
                    if examples:
                        logger.info(f"      Examples of differences:")
                        for i, example in enumerate(examples, 1):
                            primary_val = example['primary']
                            secondary_val = example['secondary']
                            logger.info(f"        {i}. Primary: '{primary_val}' → Secondary: '{secondary_val}'")
                        logger.info("")
                    else:
                        logger.info("")
            
            # Report perfect matches
            if perfect_matches:
                logger.info(f"   ✅ Perfect Matches ({len(perfect_matches)} fields):")
                # Group perfect matches for cleaner output
                if len(perfect_matches) <= 10:
                    for field in perfect_matches:
                        logger.info(f"      • {field}")
                else:
                    # Show first few and summarize
                    for field in perfect_matches[:5]:
                        logger.info(f"      • {field}")
                    logger.info(f"      ... and {len(perfect_matches) - 5} more fields")
                logger.info("")
        
        # Professional Summary
        logger.info("EXECUTIVE SUMMARY:")
        logger.info("")
        
        if total_differences_found == 0:
            logger.info("   PERFECT MATCH: All compared data fields are identical between datasets.")
            logger.info("   Data consistency is 100% - datasets are fully synchronized.")
        else:
            discrepancy_rate = (total_differences_found / total_field_comparisons * 100)
            logger.info(f"   Data Consistency: {overall_match_percentage:.2f}%")
            logger.info(f"   Discrepancy Rate: {discrepancy_rate:.2f}%")
            
            # Count issues from differences_summary
            issues_count = sum(1 for stats in differences_summary.values() if stats['match_percentage'] < 100.0)
            
            if issues_count == 1:
                logger.info(f"   Focus Area: 1 field requires attention")
            elif issues_count > 1:
                logger.info(f"   Focus Areas: {issues_count} fields require attention")
            
            if discrepancy_rate < 1.0:
                logger.info("   Recommendation: Minor discrepancies detected - review flagged fields")
            elif discrepancy_rate < 5.0:
                logger.info("   Recommendation: Moderate discrepancies - investigate data sources and processes")
            else:
                logger.info("   Recommendation: Significant discrepancies - comprehensive data audit recommended")
        
        logger.info("")
        logger.info("="*60)
    
    def _compare_without_keys(self) -> Dict[str, Any]:
        """
        Fallback comparison when no suitable keys are found.
        Provides statistical comparison without row-by-row matching.
        """
        logger.info("")
        logger.info("FALLBACK MODE: Statistical comparison without unique keys")
        logger.info("")
        logger.info("⚠ WARNING: Cannot perform precise row-by-row matching")
        logger.info("  Reason: No columns found with unique values in both datasets")
        logger.info("  Impact: Comparison limited to statistical analysis")
        logger.info("")
        
        # Compare dataset sizes
        primary_rows = self.primary_df.height
        secondary_rows = self.secondary_df.height
        row_difference = abs(primary_rows - secondary_rows)
        
        logger.info("Dataset Size Comparison:")
        logger.info(f"  {self.primary_name}: {primary_rows:,} rows")
        logger.info(f"  {self.secondary_name}: {secondary_rows:,} rows")
        logger.info(f"  Difference: {row_difference:,} rows")
        
        if row_difference == 0:
            logger.info("  ✅ Same number of rows")
        elif row_difference < primary_rows * 0.1:  # Less than 10% difference
            logger.info(f"  🟡 Small difference ({row_difference/max(primary_rows, secondary_rows)*100:.1f}%)")
        else:
            logger.info(f"  🔴 Significant difference ({row_difference/max(primary_rows, secondary_rows)*100:.1f}%)")
        
        logger.info("")
        
        # Compare unique value distributions
        stats_comparison = {}
        logger.info("Column Value Distribution Comparison:")
        
        for col in self.common_columns[:10]:  # Analyze up to 10 columns
            try:
                primary_unique = self.primary_df.select(pl.col(col).n_unique()).item()
                secondary_unique = self.secondary_df.select(pl.col(col).n_unique()).item()
                primary_nulls = self.primary_df.select(pl.col(col).is_null().sum()).item()
                secondary_nulls = self.secondary_df.select(pl.col(col).is_null().sum()).item()
                
                stats_comparison[col] = {
                    'primary_unique': primary_unique,
                    'secondary_unique': secondary_unique,
                    'primary_nulls': primary_nulls,
                    'secondary_nulls': secondary_nulls,
                    'unique_difference': abs(primary_unique - secondary_unique),
                    'null_difference': abs(primary_nulls - secondary_nulls)
                }
                
                # Assess similarity
                unique_similarity = 1 - (abs(primary_unique - secondary_unique) / max(primary_unique, secondary_unique))
                if unique_similarity > 0.9:
                    similarity_icon = "GOOD"
                elif unique_similarity > 0.7:
                    similarity_icon = "BAD"
                else:
                    similarity_icon = "VERY BAD"
                
                logger.info(f"  {similarity_icon} {col}:")
                logger.info(f"    Unique values: {primary_unique} vs {secondary_unique}")
                logger.info(f"    Null values: {primary_nulls} vs {secondary_nulls}")
                
            except Exception as e:
                logger.info(f"  ❌ {col}: Error analyzing - {str(e)}")
                continue
        
        logger.info("")
        logger.info("💡 RECOMMENDATION:")
        logger.info("  Consider adding a unique identifier column to enable precise matching")
        logger.info("  Or manually specify key columns if you know which should be unique")
        
        return {
            'fallback_mode': True,
            'row_comparison': {
                'primary_rows': primary_rows,
                'secondary_rows': secondary_rows,
                'row_difference': row_difference
            },
            'column_comparison': stats_comparison,
            'recommendation': "Add unique identifier or manually specify keys"
        }


# =============================================================================
# DEEP CONTENT REPORTING ENGINE
# =============================================================================

class DeepContentReporter(ContentComparator):
    """
    Deep content reporter integrated in the main validator workflow.
    Produces detailed per-field mismatch reports for matching keys.
    """

    def analyze_deep_content_differences(
        self,
        key_columns: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """Run deep per-field mismatch analysis for matching records."""
        logger.debug("Starting deep content analysis")

        if not self.common_columns:
            logger.debug("No common columns found for deep content comparison")
            return {}

        # If the caller provides key columns (e.g. from summary matching),
        # reuse them to avoid a second interactive prompt and to keep
        # detail aligned with summary matching.
        if key_columns is None:
            key_columns = self._detect_key_columns()

        if not key_columns:
            # Interactive fallback (terminal-only).
            key_columns = self._interactive_key_creator()

        if not key_columns:
            logger.debug("No suitable key columns found for deep content comparison")
            return {}

        primary_with_key = self._create_composite_key(self.primary_df, key_columns, "primary_key")
        secondary_with_key = self._create_composite_key(self.secondary_df, key_columns, "secondary_key")

        primary_raw_keys = primary_with_key.select("primary_key").to_series().to_list()
        secondary_raw_keys = secondary_with_key.select("secondary_key").to_series().to_list()

        primary_keys = {key for key in primary_raw_keys if key is not None}
        secondary_keys = {key for key in secondary_raw_keys if key is not None}
        matching_keys = primary_keys & secondary_keys

        if not matching_keys:
            logger.debug("No matching records available for deep analysis")
            return {
                "key_columns": key_columns,
                "total_matching_records": 0,
                "fields_analyzed": 0,
                "field_details": {}
            }

        primary_normalized = primary_with_key
        secondary_normalized = secondary_with_key

        content_columns = [col for col in self.common_columns if col not in key_columns]
        field_details: Dict[str, Any] = {}

        for idx, column_name in enumerate(content_columns, 1):
            logger.debug(f"[{idx}/{len(content_columns)}] Field deep analysis: '{column_name}'")
            field_result = self._analyze_single_field_deep(
                primary_normalized,
                secondary_normalized,
                column_name,
                matching_keys
            )
            if field_result is not None:
                field_details[column_name] = field_result

        fields_with_differences = sum(
            1 for value in field_details.values() if value["differences_found"] > 0
        )

        logger.debug(
            f"Deep content analysis completed: {len(content_columns)} fields, {fields_with_differences} with differences"
        )

        return {
            "key_columns": key_columns,
            "total_matching_records": len(matching_keys),
            "fields_analyzed": len(content_columns),
            "fields_with_differences": fields_with_differences,
            "field_details": field_details
        }

    def _analyze_single_field_deep(
        self,
        primary_df: pl.DataFrame,
        secondary_df: pl.DataFrame,
        field_name: str,
        matching_keys: set
    ) -> Optional[Dict[str, Any]]:
        """Analyze one field and return deep mismatch details."""
        primary_key_name = "primary_key"
        secondary_key_name = "secondary_key"
        secondary_field_name = f"{field_name}_secondary"

        if field_name not in primary_df.columns or field_name not in secondary_df.columns:
            logger.debug(f"Skipping '{field_name}': missing in one of the datasets")
            return None

        comparison_df = primary_df.select([primary_key_name, field_name]).join(
            secondary_df.select([secondary_key_name, field_name]),
            left_on=primary_key_name,
            right_on=secondary_key_name,
            how="inner",
            suffix="_secondary"
        ).filter(pl.col(primary_key_name).is_in(list(matching_keys)))

        if comparison_df.height == 0:
            logger.debug(f"Skipping '{field_name}': no comparable rows")
            return {
                "total_comparisons": 0,
                "differences_found": 0,
                "difference_percentage": 0.0,
                "examples": []
            }

        differences_df = comparison_df.filter(
            self._null_safe_difference_mask(pl.col(field_name), pl.col(secondary_field_name))
        )

        differences_count = differences_df.height
        total_comparisons = comparison_df.height
        difference_percentage = (
            differences_count / total_comparisons * 100 if total_comparisons > 0 else 0.0
        )

        logger.debug(
            f"Field '{field_name}' comparisons: {total_comparisons:,}, differences: {differences_count:,} ({difference_percentage:.2f}%)"
        )

        examples: List[Dict[str, str]] = []
        for row in differences_df.iter_rows(named=True):
            key_value = str(row[primary_key_name])
            primary_value = "NULL" if row[field_name] is None else str(row[field_name])
            secondary_value = "NULL" if row[secondary_field_name] is None else str(row[secondary_field_name])
            examples.append(
                {
                    "key": key_value,
                    "primary": primary_value,
                    "secondary": secondary_value
                }
            )

        return {
            "total_comparisons": total_comparisons,
            "differences_found": differences_count,
            "difference_percentage": difference_percentage,
            "examples": examples
        }


# =============================================================================
# STATISTICAL ANALYSIS ENGINE
# =============================================================================

class StatisticalAnalyzer:
    """
    Statistical analysis engine for data profiling.
    Provides statistics and data quality metrics.
    """
    
    def __init__(self, primary_df: pl.DataFrame, secondary_df: pl.DataFrame,
                 primary_name: str = "Primary", secondary_name: str = "Secondary"):
        """Initialize statistical analyzer."""
        self.primary_df = primary_df
        self.secondary_df = secondary_df
        self.primary_name = primary_name
        self.secondary_name = secondary_name
        self.common_columns = list(set(primary_df.columns) & set(secondary_df.columns))
    
    def analyze_data_quality(self) -> Dict[str, Any]:
        """Data quality analysis focused on missing values."""
        logger.info("")
        logger.info("="*60)
        logger.info("DATA QUALITY ANALYSIS")
        logger.info("="*60)
        
        results = {}
        
        # Null values analysis (main focus)
        results['null_analysis'] = self._analyze_null_values()
        
        # Numeric columns statistics (commented out)
        # results['numeric_stats'] = self._analyze_numeric_columns()
        
        # Text columns analysis (commented out)
        # results['text_stats'] = self._analyze_text_columns()
        
        return results
    
    def _analyze_null_values(self) -> Dict[str, Any]:
        """Enhanced analysis of null/missing values with detailed reporting."""
        logger.info("Missing Values Analysis:")
        logger.info("")
        
        null_analysis = {}
        overall_summary = {}
        
        for dataset_name, df in [(self.primary_name, self.primary_df), (self.secondary_name, self.secondary_df)]:
            logger.info(f"{dataset_name} Dataset:")
            
            null_counts = {}
            total_rows = df.height
            total_columns = len(df.columns)
            
            # Calculate null statistics for each column
            columns_with_nulls = 0
            total_null_cells = 0
            
            for col in df.columns:
                try:
                    null_count = df.select(pl.col(col).is_null().sum()).item()
                    null_percentage = (null_count / total_rows * 100) if total_rows > 0 else 0
                    null_counts[col] = {'count': null_count, 'percentage': null_percentage}
                    
                    if null_count > 0:
                        columns_with_nulls += 1
                        total_null_cells += null_count
                        
                except Exception:
                    null_counts[col] = {'count': 0, 'percentage': 0}
            
            # Overall dataset summary
            total_cells = total_rows * total_columns
            overall_null_percentage = (total_null_cells / total_cells * 100) if total_cells > 0 else 0
            clean_columns = total_columns - columns_with_nulls
            
            logger.info(f"  Dataset Overview:")
            logger.info(f"    Total rows: {total_rows:,}")
            logger.info(f"    Total columns: {total_columns}")
            logger.info(f"    Columns with missing values: {columns_with_nulls} ({columns_with_nulls/total_columns*100:.1f}%)")
            logger.info(f"    Clean columns (no nulls): {clean_columns} ({clean_columns/total_columns*100:.1f}%)")
            logger.info(f"    Overall missing data: {overall_null_percentage:.2f}% ({total_null_cells:,} null cells)")
            logger.info("")
            
            # Categorize columns by null percentage
            no_nulls = []
            low_nulls = []      # 0-5%
            medium_nulls = []   # 5-20%
            high_nulls = []     # 20-50%
            critical_nulls = [] # >50%
            
            for col, stats in null_counts.items():
                pct = stats['percentage']
                if pct == 0:
                    no_nulls.append(col)
                elif pct <= 5:
                    low_nulls.append((col, stats))
                elif pct <= 20:
                    medium_nulls.append((col, stats))
                elif pct <= 50:
                    high_nulls.append((col, stats))
                else:
                    critical_nulls.append((col, stats))
            
            # Report by severity
            if critical_nulls:
                logger.info(f"  🔴 Columns with >50% missing data ({len(critical_nulls)} columns):")
                for col, stats in sorted(critical_nulls, key=lambda x: x[1]['percentage'], reverse=True)[:10]:
                    logger.info(f"    {col}: {stats['percentage']:.1f}% ({stats['count']:,} nulls)")
                if len(critical_nulls) > 10:
                    logger.info(f"    ... and {len(critical_nulls) - 10} more columns")
                logger.info("")
            
            if high_nulls:
                logger.info(f"  🔴 Columns with 20-50% missing data ({len(high_nulls)} columns):")
                for col, stats in sorted(high_nulls, key=lambda x: x[1]['percentage'], reverse=True)[:10]:
                    logger.info(f"    {col}: {stats['percentage']:.1f}% ({stats['count']:,} nulls)")
                if len(high_nulls) > 10:
                    logger.info(f"    ... and {len(high_nulls) - 10} more columns")
                logger.info("")
            
            if medium_nulls:
                logger.info(f"  🟡 Columns with 5-20% missing data ({len(medium_nulls)} columns):")
                for col, stats in sorted(medium_nulls, key=lambda x: x[1]['percentage'], reverse=True)[:10]:
                    logger.info(f"    {col}: {stats['percentage']:.1f}% ({stats['count']:,} nulls)")
                if len(medium_nulls) > 10:
                    logger.info(f"    ... and {len(medium_nulls) - 10} more columns")
                logger.info("")
            
            if low_nulls:
                logger.info(f"  🟢 Columns with <5% missing data ({len(low_nulls)} columns):")
                for col, stats in sorted(low_nulls, key=lambda x: x[1]['percentage'], reverse=True)[:10]:
                    logger.info(f"    {col}: {stats['percentage']:.1f}% ({stats['count']:,} nulls)")
                if len(low_nulls) > 10:
                    logger.info(f"    ... and {len(low_nulls) - 10} more columns")
                logger.info("")
            
            if no_nulls:
                logger.info(f"  ✅ CLEAN: {len(no_nulls)} columns with no missing values")
                if len(no_nulls) <= 10:
                    logger.info(f"    {', '.join(no_nulls)}")
                else:
                    logger.info(f"    {', '.join(no_nulls[:10])} ... and {len(no_nulls) - 10} more")
                logger.info("")
            
            # Data quality assessment
            if overall_null_percentage == 0:
                quality_status = "✅ EXCELLENT - No missing data"
            elif overall_null_percentage < 1:
                quality_status = "🟢 VERY GOOD - Minimal missing data"
            elif overall_null_percentage < 5:
                quality_status = "🟢 GOOD - Low missing data"
            elif overall_null_percentage < 15:
                quality_status = "🟡 FAIR - Moderate missing data"
            elif overall_null_percentage < 30:
                quality_status = "🔴 POOR - High missing data"
            else:
                quality_status = "🔴 CRITICAL - Excessive missing data"
            
            logger.info(f"  Data Quality Assessment: {quality_status}")
            logger.info("")
            
            # Store results
            null_analysis[dataset_name] = null_counts
            overall_summary[dataset_name] = {
                'total_rows': total_rows,
                'total_columns': total_columns,
                'columns_with_nulls': columns_with_nulls,
                'total_null_cells': total_null_cells,
                'overall_null_percentage': overall_null_percentage,
                'quality_status': quality_status,
                'critical_columns': len(critical_nulls),
                'high_risk_columns': len(high_nulls),
                'medium_risk_columns': len(medium_nulls),
                'low_risk_columns': len(low_nulls),
                'clean_columns': len(no_nulls)
            }
        
        # Cross-dataset comparison
        if len(null_analysis) == 2:
            logger.info("Cross-Dataset Missing Values Comparison:")
            datasets = list(null_analysis.keys())
            primary_summary = overall_summary[datasets[0]]
            secondary_summary = overall_summary[datasets[1]]
            
            logger.info(f"  Overall Data Quality:")
            logger.info(f"    {datasets[0]}: {primary_summary['overall_null_percentage']:.2f}% missing")
            logger.info(f"    {datasets[1]}: {secondary_summary['overall_null_percentage']:.2f}% missing")
            
            diff = abs(primary_summary['overall_null_percentage'] - secondary_summary['overall_null_percentage'])
            if diff < 1:
                logger.info(f"    🟢 Similar data quality (±{diff:.2f}%)")
            elif diff < 5:
                logger.info(f"    🟡 Moderate difference (±{diff:.2f}%)")
            else:
                logger.info(f"    🔴 Significant difference (±{diff:.2f}%)")
            
            logger.info("")
        
        null_analysis['summary'] = overall_summary
        return null_analysis
    
    def _analyze_numeric_columns(self) -> Dict[str, Any]:
        """Analyze numeric columns statistics."""
        logger.info("Numeric Columns Statistics:")
        
        numeric_stats = {}
        
        for dataset_name, df in [(self.primary_name, self.primary_df), (self.secondary_name, self.secondary_df)]:
            dataset_stats = {}
            
            # Find numeric columns
            numeric_cols = []
            for col in df.columns:
                try:
                    if df.select(pl.col(col)).dtypes[0] in [pl.Int64, pl.Int32, pl.Float64, pl.Float32]:
                        numeric_cols.append(col)
                except Exception:
                    continue
            
            if numeric_cols:
                logger.info(f"  {dataset_name} - {len(numeric_cols)} numeric columns found")
                
                for col in numeric_cols[:5]:  # Limit to first 5
                    try:
                        stats = df.select([
                            pl.col(col).min().alias('min'),
                            pl.col(col).max().alias('max'),
                            pl.col(col).mean().alias('mean'),
                            pl.col(col).median().alias('median'),
                            pl.col(col).std().alias('std')
                        ]).to_dicts()[0]
                        
                        dataset_stats[col] = stats
                        
                        # Format display
                        mean_val = stats['mean'] if stats['mean'] is not None else 0
                        logger.info(f"    {col}: mean={mean_val:.2f}, range=[{stats['min']}, {stats['max']}]")
                        
                    except Exception:
                        continue
            else:
                logger.info(f"  {dataset_name}: No numeric columns found")
            
            numeric_stats[dataset_name] = dataset_stats
        
        return numeric_stats
    
    def _analyze_text_columns(self) -> Dict[str, Any]:
        """Analyze text columns characteristics."""
        logger.info("Text Columns Analysis:")
        
        text_stats = {}
        
        for dataset_name, df in [(self.primary_name, self.primary_df), (self.secondary_name, self.secondary_df)]:
            dataset_stats = {}
            
            # Find text columns
            text_cols = []
            for col in df.columns:
                try:
                    if df.select(pl.col(col)).dtypes[0] in [pl.Utf8, pl.String]:
                        text_cols.append(col)
                except Exception:
                    continue
            
            if text_cols:
                logger.info(f"  {dataset_name} - {len(text_cols)} text columns found")
                
                for col in text_cols[:5]:  # Limit to first 5
                    try:
                        unique_count = df.select(pl.col(col).n_unique()).item()
                        total_count = df.height
                        uniqueness = (unique_count / total_count * 100) if total_count > 0 else 0
                        
                        dataset_stats[col] = {
                            'unique_values': unique_count,
                            'uniqueness_percentage': uniqueness
                        }
                        
                        logger.info(f"    {col}: {unique_count:,} unique values ({uniqueness:.1f}% unique)")
                        
                    except Exception:
                        continue
            else:
                logger.info(f"  {dataset_name}: No text columns found")
            
            text_stats[dataset_name] = dataset_stats
        
        return text_stats


# Configured in main() via configure_logging(report_name)
logger = logging.getLogger("GlobalValidator")
logger.addHandler(logging.NullHandler())

# =============================================================================
# COLUMN MAPPING SYSTEM
# =============================================================================

class ColumnMapper:
    """
    Column mapping system for renaming columns based on external mapping files.
    Supports optional remapping of primary, secondary, or both datasets.
    """
    
    def __init__(self, mapping_file_path: Optional[str] = None):
        """
        Initialize column mapper.
        
        Args:
            mapping_file_path (Optional[str]): Path to mapping file (xlsx/csv) or None to skip mapping
        """
        self.mapping_file_path = mapping_file_path
        self.mapping_dict = {}
        self.is_enabled = mapping_file_path is not None and mapping_file_path.strip() != ""
        
        if self.is_enabled:
            self._load_mapping_file()
    
    def _load_mapping_file(self) -> None:
        """Load and validate mapping file."""
        try:
            logger.info("")
            logger.info("="*60)
            logger.info("COLUMN MAPPING SYSTEM")
            logger.info("="*60)
            logger.info(f"Loading mapping file: {Path(self.mapping_file_path).name}")
            
            # Use FileReader to load mapping file
            mapping_df = FileReader.read_file(Path(self.mapping_file_path))
            
            # Validate mapping file structure
            if mapping_df.shape[1] < 2:
                raise ValueError(f"Mapping file must have at least 2 columns, found {mapping_df.shape[1]}")
            
            # Get first two columns (original_name, new_name)
            original_col = mapping_df.columns[0]
            new_col = mapping_df.columns[1]
            
            # Create mapping dictionary with automatic trimming of spaces
            mapping_pairs = mapping_df.select([original_col, new_col]).to_pandas()
            raw_dict = dict(zip(mapping_pairs.iloc[:, 0], mapping_pairs.iloc[:, 1]))
            
            # Clean and trim all keys and values
            self.mapping_dict = {}
            for k, v in raw_dict.items():
                if pd.notna(k) and pd.notna(v):
                    clean_key = str(k).strip()
                    clean_value = str(v).strip()
                    if clean_key != "" and clean_value != "":
                        self.mapping_dict[clean_key] = clean_value
            
            logger.info(f"  Loaded {len(self.mapping_dict)} column mappings")
            logger.info(f"  Sample mappings:")
            for i, (orig, new) in enumerate(list(self.mapping_dict.items())):
                logger.info(f"    '{orig}' → '{new}'")
            logger.info("")
            logger.info("  ✅ All mapping keys and values have been trimmed of leading/trailing spaces")
            logger.info("")
            
        except Exception as e:
            logger.error(f"Failed to load mapping file: {str(e)}")
            self.is_enabled = False
            self.mapping_dict = {}
    
    def apply_mapping(self, df: pl.DataFrame, dataset_name: str) -> pl.DataFrame:
        """
        Apply column mapping to a dataframe.
        
        Args:
            df (pl.DataFrame): Input dataframe
            dataset_name (str): Name of dataset for logging
            
        Returns:
            pl.DataFrame: Dataframe with renamed columns
        """
        if not self.is_enabled or not self.mapping_dict:
            return df
        
        logger.info(f"Applying column mapping to {dataset_name} dataset:")
        
        original_columns = df.columns
        mapped_columns = []
        unmapped_columns = []
        columns_with_spaces = []
        
        # Check for columns with leading/trailing spaces
        for col in original_columns:
            if str(col) != str(col).strip():
                columns_with_spaces.append(f"'{col}' → '{str(col).strip()}'")
        
        if columns_with_spaces:
            logger.info(f"  ⚠️  Found {len(columns_with_spaces)} column(s) with leading/trailing spaces:")
            for space_col in columns_with_spaces[:5]:  # Show first 5
                logger.info(f"    {space_col}")
            if len(columns_with_spaces) > 5:
                logger.info(f"    ... and {len(columns_with_spaces) - 5} more")
            logger.info("  ✅ These will be automatically trimmed during mapping")
            logger.info("")
        
        # Create new column names with automatic trimming
        new_column_names = []
        name_counter = {}  # Track duplicate names
        
        for col in original_columns:
            # Try exact match first
            if col in self.mapping_dict:
                new_name = self.mapping_dict[col]
            else:
                # Try trimmed match (handle columns with trailing spaces)
                trimmed_col = str(col).strip()
                if trimmed_col in self.mapping_dict:
                    new_name = self.mapping_dict[trimmed_col]
                    mapped_columns.append(f"'{col}' (trimmed: '{trimmed_col}') → '{new_name}'")
                else:
                    # No mapping found - keep original but trim spaces
                    new_name = str(col).strip()
                    unmapped_columns.append(col)
            
            # Handle duplicate names by adding suffix
            if new_name in name_counter:
                name_counter[new_name] += 1
                unique_name = f"{new_name}_{name_counter[new_name]}"
                logger.warning(f"  ⚠ Duplicate column name '{new_name}' detected, renaming to '{unique_name}'")
                new_column_names.append(unique_name)
            else:
                name_counter[new_name] = 0
                new_column_names.append(new_name)
                
            # Add to mapped columns if it was actually mapped (not just trimmed)
            if col in self.mapping_dict:
                mapped_columns.append(f"'{col}' → '{new_name}'")
        
        # Apply renaming
        try:
            renamed_df = df.select([pl.col(old_name).alias(new_name) 
                                   for old_name, new_name in zip(original_columns, new_column_names)])
            
            # Log results
            logger.info(f"  Mapping Results:")
            logger.info(f"    Total columns: {len(original_columns)}")
            logger.info(f"    Mapped columns: {len(mapped_columns)}")
            logger.info(f"    Unmapped columns: {len(unmapped_columns)}")
            
            if mapped_columns:
                logger.info(f"  ✅ Successfully mapped columns:")
                for mapping in mapped_columns:  # Show first 5
                    logger.info(f"    {mapping}")
            
            if unmapped_columns:
                logger.info(f"  ⚠ Unmapped columns (kept original names):")
                unmapped_display = ', '.join(unmapped_columns[:10])
                if len(unmapped_columns) > 10:
                    unmapped_display += f" ... and {len(unmapped_columns) - 10} more"
                logger.info(f"    {unmapped_display}")
            
            mapping_success_rate = (len(mapped_columns) / len(original_columns)) * 100
            if mapping_success_rate == 100:
                logger.info(f" Perfect mapping: {mapping_success_rate:.0f}% columns mapped")
            elif mapping_success_rate >= 80:
                logger.info(f" Good mapping: {mapping_success_rate:.1f}% columns mapped")
            elif mapping_success_rate >= 50:
                logger.info(f" Partial mapping: {mapping_success_rate:.1f}% columns mapped")
            else:
                logger.info(f" Poor mapping: {mapping_success_rate:.1f}% columns mapped")
            
            logger.info("")
            return renamed_df
            
        except Exception as e:
            logger.error(f"Error applying column mapping to {dataset_name}: {str(e)}")
            logger.info(f"  Returning original dataset without mapping")
            logger.info("")
            return df


# Main execution function
def run_validation(
    primary_path: str,
    secondary_path: str,
    primary_sheet: Optional[Union[str, int]] = None,
    secondary_sheet: Optional[Union[str, int]] = None,
                  mapping_file: Optional[str] = None, 
    apply_mapping_to: str = "both",
    run_deep_analysis: bool = True,
    normalize_before_compare: bool = True,
    normalize_numeric_strings: bool = True,
    strip_leading_zeros_numeric: bool = True,
    convert_dates: bool = True
) -> Dict[str, Any]:
    """
    Run complete validation workflow between two files with optional column mapping.
    
    Args:
        primary_path (str): Path to primary file
        secondary_path (str): Path to secondary file  
        primary_sheet (Optional[Union[str, int]]): Sheet name/index for primary Excel file
        secondary_sheet (Optional[Union[str, int]]): Sheet name/index for secondary Excel file
        mapping_file (Optional[str]): Path to column mapping file (xlsx/csv) or None
        apply_mapping_to (str): Apply mapping to "primary", "secondary", or "both"
        run_deep_analysis (bool): Enable deep per-field mismatch report
        normalize_before_compare (bool): Enable in-memory value normalization before comparison
        normalize_numeric_strings (bool): Canonicalize values like 12.0 -> 12
        strip_leading_zeros_numeric (bool): Normalize pure numeric tokens like 0042 -> 42
        
    Returns:
        Dict[str, Any]: Complete validation results
    """
    try:
        # Initialize configuration
        config = FileConfig()
        
        # Set file paths
        if not config.set_files(primary_path, secondary_path):
            logger.error("File validation failed")
            return {}
        
        # Initialize column mapper
        mapper = ColumnMapper(mapping_file)

        # Preflight format/integrity check
        integrity_checks = run_input_integrity_checks(
            config.primary_file,
            config.secondary_file,
            primary_sheet,
            secondary_sheet
        )
        
        # Load datasets
        logger.info("Loading datasets...")
        primary_df = FileReader.read_file(config.primary_file, primary_sheet)
        secondary_df = FileReader.read_file(config.secondary_file, secondary_sheet)
        
        # Apply column mapping if enabled
        if mapper.is_enabled:
            if apply_mapping_to in ["primary", "both"]:
                primary_df = mapper.apply_mapping(primary_df, "Primary")
            
            if apply_mapping_to in ["secondary", "both"]:
                secondary_df = mapper.apply_mapping(secondary_df, "Secondary")
            
            if apply_mapping_to not in ["primary", "secondary", "both"]:
                logger.warning(f"Invalid apply_mapping_to value: '{apply_mapping_to}'. Using 'both' as default.")
                primary_df = mapper.apply_mapping(primary_df, "Primary")
                secondary_df = mapper.apply_mapping(secondary_df, "Secondary")

        normalization_reports = {}
        if normalize_before_compare:
            logger.info("")
            logger.info("=" * 60)
            logger.info("DATA NORMALIZATION (IN-MEMORY)")
            logger.info("=" * 60)
            logger.info("")
            primary_df, primary_norm_report = DataNormalizer.normalize_dataframe(
                primary_df,
                dataset_name="primary",
                normalize_numeric_strings=normalize_numeric_strings,
                strip_leading_zeros_numeric=strip_leading_zeros_numeric,
                convert_dates=convert_dates,
            )
            secondary_df, secondary_norm_report = DataNormalizer.normalize_dataframe(
                secondary_df,
                dataset_name="secondary",
                normalize_numeric_strings=normalize_numeric_strings,
                strip_leading_zeros_numeric=strip_leading_zeros_numeric,
                convert_dates=convert_dates,
            )
            normalization_reports = {
                "primary": primary_norm_report,
                "secondary": secondary_norm_report
            }
            logger.info(
                "  Normalized columns summary:"
                f" primary(numeric={len(primary_norm_report['numeric_normalized_columns'])},"
                f" null={len(primary_norm_report['null_standardized_columns'])}) |"
                f" secondary(numeric={len(secondary_norm_report['numeric_normalized_columns'])},"
                f" null={len(secondary_norm_report['null_standardized_columns'])})"
            )
            logger.info("")
        
        # Run analysis
        primary_name = config.primary_file.stem
        secondary_name = config.secondary_file.stem
        
        # Structural analysis
        analyzer = DataAnalyzer(primary_df, secondary_df, primary_name, secondary_name)
        results = analyzer.run_full_analysis()
        results["input_integrity"] = integrity_checks
        results["normalization"] = normalization_reports
        
        # Content comparison
        comparator = ContentComparator(primary_df, secondary_df, primary_name, secondary_name)
        results['content_comparison'] = comparator.find_matching_records()

        # Deep content analysis (integrated)
        if run_deep_analysis:
            # Reuse the potentially updated DataFrames from the content comparator
            # (e.g. after interactive key creation) so detail matching aligns with summary.
            deep_reporter = DeepContentReporter(comparator.primary_df, comparator.secondary_df, primary_name, secondary_name)
            # Keep deep verbose output out of summary log. Full details are written
            # into the dedicated detail report file.
            previous_level = logger.level
            logger.setLevel(logging.WARNING)
            try:
                content_comparison = results.get("content_comparison", {})
                key_columns = content_comparison.get("key_columns") if isinstance(content_comparison, dict) else None
                results['deep_content_analysis'] = deep_reporter.analyze_deep_content_differences(
                    key_columns=key_columns
                )
            finally:
                logger.setLevel(previous_level)
        
        # Statistical analysis
        stats_analyzer = StatisticalAnalyzer(primary_df, secondary_df, primary_name, secondary_name)
        results['statistical_analysis'] = stats_analyzer.analyze_data_quality()
        
        logger.info("")
        logger.info("=" * 60)
        logger.info("VALIDATION COMPLETED SUCCESSFULLY")
        logger.info("=" * 60)
        
        return results
        
    except Exception as e:
        logger.error(f"Validation failed: {str(e)}")
        return {}


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(
        description="Global Data Validator with integrated deep content reporting"
    )
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG_PATH),
        help="Path to the centralized JSON config file"
    )
    return parser.parse_args()


def main() -> int:
    """Application entrypoint."""
    try:
        args = parse_args()
        settings = load_settings(args.config)
        configure_logging(settings.report_name)

        logger.info("Global Data Validator Configuration:")
        logger.info(f"  Primary file: {Path(settings.primary_file).name}")
        logger.info(f"  Secondary file: {Path(settings.secondary_file).name}")
        logger.info(
            f"  Primary sheet: {settings.primary_sheet if settings.primary_sheet is not None else 'Default'}"
        )
        logger.info(
            f"  Secondary sheet: {settings.secondary_sheet if settings.secondary_sheet is not None else 'Default'}"
        )
        logger.info(f"  Column mapping target: {settings.apply_mapping_to}")
        logger.info(f"  Deep analysis: {'Enabled' if settings.run_deep_analysis else 'Disabled'}")
        logger.info(f"  Normalization: {'Enabled' if settings.normalize_before_compare else 'Disabled'}")
        logger.info(
            f"  Strip leading zero numeric tokens: {'Enabled' if settings.strip_leading_zeros_numeric else 'Disabled'}"
        )
        logger.info("  Key detection: Auto")
        logger.info("  Report directory: logs")
        logger.info(f"  Report name: {settings.report_name}")
        logger.info("")

        results = run_validation(
            primary_path=settings.primary_file,
            secondary_path=settings.secondary_file,
            primary_sheet=settings.primary_sheet,
            secondary_sheet=settings.secondary_sheet,
            mapping_file=settings.mapping_file,
            apply_mapping_to=settings.apply_mapping_to,
            run_deep_analysis=settings.run_deep_analysis,
            normalize_before_compare=settings.normalize_before_compare,
            normalize_numeric_strings=settings.normalize_numeric_strings,
            strip_leading_zeros_numeric=settings.strip_leading_zeros_numeric,
            convert_dates=settings.convert_dates,
        )

        deep_report_path = write_detail_report(results, settings.report_name)
        logger.info("")
        logger.info(f"Deep report written to: {deep_report_path}")
        logger.info("")

        summary_report_path = finalize_summary_log(settings.report_name)
        print(f"Summary report written to: {summary_report_path}")
        print(f"Deep report written to: {deep_report_path}")
        return 0

    except Exception as exc:
        logger.error(f"Application error: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())