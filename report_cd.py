#!/usr/bin/env python3
"""
Build final Excel report from Global Validator summary/detail text reports.

This script fills a fixed 2-sheet template:
- High_Level_Overview
- Dettaglio_Disallineamenti

Default blank workbook: templates/report_base_template.xlsx (layout, colors, merges).
Override with --template-path or optional excel_template_path in validator_config.json.
Default inputs: logs/summary_<report_name>.txt, logs/detail_<report_name>.txt;
default output: reports/report_<report_name>.xlsx.
"""

from __future__ import annotations

import argparse
import ast
import json
import re
from copy import copy
from pathlib import Path
from typing import Any, Dict, List, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# Max data rows written per field block in the Excel detail sheet (full count stays in header).
DETAIL_ROWS_EXPORT_CAP = 15_000
DEFAULT_CONFIG_PATH = Path("validator_config.json")
LOGS_DIR = Path("logs")
REPORTS_DIR = Path("reports")
TEMPLATES_DIR = Path("templates")
REPORT_BASE_TEMPLATE = TEMPLATES_DIR / "report_base_template.xlsx"


def _safe_report_segment(report_name: str) -> str:
    s = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in report_name.strip())
    return s or "default"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate template-based validation Excel report.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help="validator_config.json (used for report_name, template path, default log paths)",
    )
    parser.add_argument(
        "--summary-path",
        type=Path,
        default=None,
        help="Override: default is logs/summary_<report_name>.txt from config",
    )
    parser.add_argument(
        "--detail-path",
        type=Path,
        default=None,
        help="Override: default is logs/detail_<report_name>.txt from config",
    )
    parser.add_argument(
        "--template-path",
        type=Path,
        default=None,
        help="Override: default is templates/report_base_template.xlsx (or excel_template_path in config)",
    )
    parser.add_argument(
        "--output-path",
        type=Path,
        default=None,
        help="Override: default is reports/report_<report_name>.xlsx",
    )
    return parser.parse_args()


def resolve_paths_from_config(args: argparse.Namespace) -> Tuple[Path, Path, Path, Path]:
    """Return summary_path, detail_path, template_path, output_path."""
    raw = json.loads(args.config.read_text(encoding="utf-8"))
    report_name = str(raw.get("report_name", "default")).strip() or "default"
    safe = _safe_report_segment(report_name)

    summary_path = args.summary_path or (LOGS_DIR / f"summary_{safe}.txt")
    detail_path = args.detail_path or (LOGS_DIR / f"detail_{safe}.txt")

    if args.template_path is not None:
        template_path = args.template_path
    else:
        tpl = raw.get("excel_template_path")
        template_path = Path(tpl) if tpl else REPORT_BASE_TEMPLATE

    if args.output_path is not None:
        output_path = args.output_path
    else:
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        output_path = REPORTS_DIR / f"report_{safe}.xlsx"

    return summary_path, detail_path, template_path, output_path


def _to_int(value: str) -> int:
    cleaned = value.replace(",", "").strip()
    return int(cleaned) if cleaned else 0


def _dataset_file_label(summary: Dict[str, Any], role: str) -> str:
    """Prefer actual filename; fall back to dataset display name."""
    path_key = f"{role}_file"
    raw = (summary.get(path_key) or "").strip()
    if raw:
        return Path(raw).name
    return str(summary.get(f"{role}_name") or role).strip()


def parse_summary(summary_path: Path) -> Dict[str, Any]:
    text = summary_path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()

    summary: Dict[str, Any] = {
        "report_name": "",
        "primary_file": "",
        "secondary_file": "",
        "primary_name": "Primary",
        "secondary_name": "Secondary",
        "primary_rows": 0,
        "primary_cols": 0,
        "secondary_rows": 0,
        "secondary_cols": 0,
        "common_columns": 0,
        "primary_only_columns": 0,
        "secondary_only_columns": 0,
        "matching_records": 0,
        "primary_only_rows": 0,
        "secondary_only_rows": 0,
        "total_comparisons": 0,
        "differences_found": 0,
        "fields_analyzed": 0,
        "key_columns": "",
        "primary_only_key_samples": [],
        "secondary_only_key_samples": [],
    }

    # Header config block
    for line in lines:
        m = re.search(r"Report name:\s*(.+)$", line)
        if m:
            summary["report_name"] = m.group(1).strip()
        if "Primary file:" in line:
            summary["primary_file"] = line.split("Primary file:", 1)[1].strip()
        elif "Secondary file:" in line:
            summary["secondary_file"] = line.split("Secondary file:", 1)[1].strip()

    # Dataset dimensions lines
    dim_pattern = re.compile(r"\|\s*INFO\s*\|\s+([^:]+):\s*([\d,]+)\s*rows\s*×\s*([\d,]+)\s*columns")
    dim_hits: List[Tuple[str, int, int]] = []
    for line in lines:
        m = dim_pattern.search(line)
        if m:
            dataset_name = m.group(1).strip()
            dim_hits.append((dataset_name, _to_int(m.group(2)), _to_int(m.group(3))))
    if len(dim_hits) >= 2:
        summary["primary_name"], summary["primary_rows"], summary["primary_cols"] = dim_hits[0]
        summary["secondary_name"], summary["secondary_rows"], summary["secondary_cols"] = dim_hits[1]

    # "X only" appears in two sections:
    # - Column Overview -> only columns
    # - Record Matching Results -> only rows
    # Keep first occurrence for columns and last occurrence for rows.
    primary_only_hits: List[int] = []
    secondary_only_hits: List[int] = []

    # Column counts
    for line in lines:
        m = re.search(r"Common columns:\s*([\d,]+)", line)
        if m:
            summary["common_columns"] = _to_int(m.group(1))
        m = re.search(rf"{re.escape(summary['primary_name'])}\s+only:\s*([\d,]+)", line)
        if m:
            primary_only_hits.append(_to_int(m.group(1)))
        m = re.search(rf"{re.escape(summary['secondary_name'])}\s+only:\s*([\d,]+)", line)
        if m:
            secondary_only_hits.append(_to_int(m.group(1)))

    if primary_only_hits:
        summary["primary_only_columns"] = primary_only_hits[0]
        summary["primary_only_rows"] = primary_only_hits[-1]
    if secondary_only_hits:
        summary["secondary_only_columns"] = secondary_only_hits[0]
        summary["secondary_only_rows"] = secondary_only_hits[-1]

    # Matching results
    for line in lines:
        m = re.search(r"Using key columns:\s*(.+)", line)
        if m:
            summary["key_columns"] = m.group(1).strip()
        m = re.search(r"Matching records:\s*([\d,]+)", line)
        if m:
            summary["matching_records"] = _to_int(m.group(1))
        m = re.search(r"Fields Analyzed:\s*([\d,]+)", line)
        if m:
            summary["fields_analyzed"] = _to_int(m.group(1))
        m = re.search(r"Total Comparisons:\s*([\d,]+)", line)
        if m:
            summary["total_comparisons"] = _to_int(m.group(1))
        m = re.search(r"Differences Found:\s*([\d,]+)", line)
        if m:
            summary["differences_found"] = _to_int(m.group(1))

    # Optional sample keys
    for line in lines:
        if "only samples:" in line:
            try:
                payload = line.split("only samples:", 1)[1].strip()
                parsed = ast.literal_eval(payload) if payload.startswith("[") else []
            except Exception:
                parsed = []
            if summary["primary_name"] in line:
                summary["primary_only_key_samples"] = parsed
            elif summary["secondary_name"] in line:
                summary["secondary_only_key_samples"] = parsed

    return summary


def _only_key_lists_for_excel(summary: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    """String keys, sorted; capped elsewhere for row export."""
    pk = summary.get("primary_only_key_samples") or []
    sk = summary.get("secondary_only_key_samples") or []
    pk_s = sorted({str(x) for x in pk})
    sk_s = sorted({str(x) for x in sk})
    return pk_s, sk_s


def parse_detail(detail_path: Path) -> Dict[str, List[Dict[str, str]]]:
    lines = detail_path.read_text(encoding="utf-8", errors="replace").splitlines()
    field_rows: Dict[str, List[Dict[str, str]]] = {}
    current_field = ""

    for line in lines:
        if line.startswith("FIELD: "):
            current_field = line.replace("FIELD: ", "", 1).strip()
            if current_field not in field_rows:
                field_rows[current_field] = []
            continue
        if re.match(r"^\d+\s+\|\s+", line):
            parts = [part.strip() for part in line.split("|")]
            # Expected: idx | key | primary | secondary
            if len(parts) >= 4 and current_field:
                field_rows[current_field].append(
                    {
                        "key": parts[1],
                        "primary": parts[2],
                        "secondary": parts[3],
                    }
                )
    return field_rows


def fill_template(
    template_path: Path,
    output_path: Path,
    summary: Dict[str, Any],
    detail_rows_by_field: Dict[str, List[Dict[str, str]]],
) -> None:
    wb = load_workbook(template_path)
    ws_overview = wb["High_Level_Overview"]
    ws_detail = wb["Dettaglio_Disallineamenti"]

    # Same banner look as overview title row (theme-based fill from template A1).
    _overview_title_ref = ws_overview["A1"]
    overview_title_banner_fill = copy(_overview_title_ref.fill)
    overview_title_banner_font = copy(_overview_title_ref.font)

    # ---------------------------
    # Sheet 1: High_Level_Overview
    # ---------------------------
    # Title from report_name (fallback to summary filename stem).
    title = summary["report_name"].strip() if summary.get("report_name") else output_path.stem
    overview_title_cell = ws_overview["A1"]
    overview_title_cell.value = title.upper()
    overview_title_cell.fill = copy(overview_title_banner_fill)
    overview_title_cell.font = copy(overview_title_banner_font)

    ws_overview["A3"] = summary["primary_name"]
    ws_overview["B3"] = summary["primary_rows"]
    ws_overview["C3"] = summary["primary_cols"]

    ws_overview["A4"] = summary["secondary_name"]
    ws_overview["B4"] = summary["secondary_rows"]
    ws_overview["C4"] = summary["secondary_cols"]

    # Last DB Update -> run date (today), both rows.
    run_date = datetime.now().date()
    ws_overview["D3"] = run_date
    ws_overview["D4"] = run_date

    # Validation time range -> placeholder requested by user.
    ws_overview["E3"] = "TO BE COMPLETED"
    ws_overview["E4"] = "TO BE COMPLETED"

    ws_overview["C7"] = summary["matching_records"]
    ws_overview["D7"] = f'Key Field: "{summary["key_columns"]}"'
    ws_overview["A8"] = f"{summary['primary_name']} Only Rows"
    ws_overview["A9"] = f"{summary['secondary_name']} Only Rows"
    ws_overview["C8"] = summary["primary_only_rows"]
    ws_overview["C9"] = summary["secondary_only_rows"]

    ws_overview["C12"] = summary["common_columns"]
    ws_overview["A13"] = f"{summary['primary_name']} Only Columns"
    ws_overview["A14"] = f"{summary['secondary_name']} Only Columns"
    ws_overview["C13"] = summary["primary_only_columns"]
    ws_overview["C14"] = summary["secondary_only_columns"]

    ws_overview["C17"] = summary["matching_records"]
    ws_overview["C18"] = summary["fields_analyzed"] if summary["fields_analyzed"] else summary["common_columns"]
    ws_overview["C19"] = summary["total_comparisons"]
    ws_overview["C20"] = summary["differences_found"]
    ws_overview["C21"] = (
        summary["differences_found"] / summary["total_comparisons"]
        if summary["total_comparisons"] > 0
        else 0
    )

    # VALIDATED badge (column F, typically merged across row 3-4 in template)
    # Green + VALIDATED when no misalignments, otherwise red + NOT VALIDATED.
    validated_ok = summary["differences_found"] == 0
    badge_cell = ws_overview["F2"]
    badge_cell.value = "VALIDATED" if validated_ok else "NOT VALIDATED"
    badge_cell.fill = PatternFill(
        fill_type="solid",
        fgColor=("00B050" if validated_ok else "C00000")
    )
    badge_cell.font = Font(color="FFFFFF", bold=True)

    # ---------------------------
    # Sheet 2: Dettaglio_Disallineamenti
    # ---------------------------
    # Row 1 banner: same fill/font as High_Level_Overview A1 (template).
    field_fill = PatternFill(fill_type="solid", fgColor="FFC000")
    header_fill = PatternFill(fill_type="solid", fgColor="FFD966")
    # No solid white fill: that hides worksheet gridlines in Excel; default empty fill keeps gridlines.
    clear_fill = PatternFill()
    body_font = Font(color="000000")
    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    thin_side = Side(border_style="thin", color="000000")
    thick_side = Side(border_style="thick", color="000000")

    def apply_block_borders(start_col: int, end_col: int, start_row: int, end_row: int) -> None:
        """Apply thick outer borders and thin inner borders to one field block."""
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                left = thick_side if c == start_col else thin_side
                right = thick_side if c == end_col else thin_side
                top = thick_side if r == start_row else thin_side
                bottom = thick_side if r == end_row else thin_side
                ws_detail.cell(row=r, column=c).border = Border(
                    left=left, right=right, top=top, bottom=bottom
                )

    # Clear existing merged ranges and values, keep sheet and styles template-friendly.
    for merged_range in list(ws_detail.merged_cells.ranges):
        ws_detail.unmerge_cells(str(merged_range))
    for row in ws_detail.iter_rows(min_row=1, max_row=ws_detail.max_row, min_col=1, max_col=ws_detail.max_column):
        for cell in row:
            cell.value = None
            # Strip template fills/fonts (e.g. orange row) — we re-apply styles on written cells.
            cell.fill = clear_fill
            cell.font = body_font
            cell.border = Border()

    # Layout: cols A–B = Non Common Keys, col C = spacer, col D+ = misalignments (unchanged layout).
    NON_COMMON_END_COL = 2
    SPACER_COL = 3
    MISMATCH_START_COL = 4
    banner_text_mismatch = (
        f"Found Misalignments Detail — max {DETAIL_ROWS_EXPORT_CAP:,} rows exported per field"
    )

    pk_list, sk_list = _only_key_lists_for_excel(summary)
    n_pk = int(summary.get("primary_only_rows") or 0)
    n_sk = int(summary.get("secondary_only_rows") or 0)
    pk_export = pk_list[:DETAIL_ROWS_EXPORT_CAP]
    sk_export = sk_list[:DETAIL_ROWS_EXPORT_CAP]
    primary_lbl = _dataset_file_label(summary, "primary")
    secondary_lbl = _dataset_file_label(summary, "secondary")

    # Keep only fields with at least one mismatch row.
    non_empty_fields = [(field, rows) for field, rows in detail_rows_by_field.items() if rows]
    total_groups = len(non_empty_fields)
    mismatch_width = total_groups * 3 + max(0, total_groups - 1) if total_groups else 3
    last_sheet_col = max(NON_COMMON_END_COL, SPACER_COL + mismatch_width)

    # --- Row 1: Non Common Keys (A:B) ---
    ws_detail.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NON_COMMON_END_COL)
    nc_banner = ws_detail.cell(row=1, column=1, value="Non Common Keys")
    nc_banner.fill = copy(overview_title_banner_fill)
    nc_banner.font = copy(overview_title_banner_font)
    nc_banner.alignment = center

    # --- Row 2: two yellow headers (file only + counts from summary) ---
    p_head = ws_detail.cell(row=2, column=1, value=f"{primary_lbl} only ({n_pk})")
    s_head = ws_detail.cell(row=2, column=2, value=f"{secondary_lbl} only ({n_sk})")
    for h in (p_head, s_head):
        h.fill = field_fill
        h.font = bold_font
        h.alignment = center

    # --- Row 3+: only-key lists (side by side, padded) ---
    max_k_rows = max(len(pk_export), len(sk_export))
    for i in range(max_k_rows):
        r = 3 + i
        c1 = ws_detail.cell(row=r, column=1, value=pk_export[i] if i < len(pk_export) else None)
        c2 = ws_detail.cell(row=r, column=2, value=sk_export[i] if i < len(sk_export) else None)
        c1.font = body_font
        c2.font = body_font
        c1.alignment = center
        c2.alignment = center

    left_end_row = 2 if max_k_rows == 0 else 2 + max_k_rows
    apply_block_borders(1, NON_COMMON_END_COL, 1, left_end_row)
    ws_detail.column_dimensions["A"].width = 22
    ws_detail.column_dimensions["B"].width = 22
    ws_detail.column_dimensions["C"].width = 3

    max_block_rows = left_end_row

    # --- Misalignment block (from column D) ---
    if not non_empty_fields:
        ws_detail.merge_cells(
            start_row=1,
            start_column=MISMATCH_START_COL,
            end_row=1,
            end_column=MISMATCH_START_COL + 2,
        )
        top_m = ws_detail.cell(row=1, column=MISMATCH_START_COL, value=banner_text_mismatch)
        top_m.fill = copy(overview_title_banner_fill)
        top_m.font = copy(overview_title_banner_font)
        top_m.alignment = center
        ws_detail.merge_cells(
            start_row=2,
            start_column=MISMATCH_START_COL,
            end_row=2,
            end_column=MISMATCH_START_COL + 2,
        )
        nm = ws_detail.cell(row=2, column=MISMATCH_START_COL, value="No field-level mismatches found.")
        nm.fill = field_fill
        nm.font = bold_font
        nm.alignment = center
        max_block_rows = max(max_block_rows, 2)
        apply_block_borders(MISMATCH_START_COL, MISMATCH_START_COL + 2, 1, 2)
    else:
        ws_detail.merge_cells(
            start_row=1,
            start_column=MISMATCH_START_COL,
            end_row=1,
            end_column=MISMATCH_START_COL + mismatch_width - 1,
        )
        top_m = ws_detail.cell(row=1, column=MISMATCH_START_COL, value=banner_text_mismatch)
        top_m.fill = copy(overview_title_banner_fill)
        top_m.font = copy(overview_title_banner_font)
        top_m.alignment = center

        start_col = MISMATCH_START_COL
        for field_name, rows in non_empty_fields:
            full_count = len(rows)
            rows_export = rows[:DETAIL_ROWS_EXPORT_CAP]
            ws_detail.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=start_col + 2)
            field_title = ws_detail.cell(row=2, column=start_col, value=f"{field_name} ({full_count})")
            field_title.fill = field_fill
            field_title.font = bold_font
            field_title.alignment = center

            h1 = ws_detail.cell(row=3, column=start_col, value="Key")
            h2 = ws_detail.cell(row=3, column=start_col + 1, value=summary["primary_name"])
            h3 = ws_detail.cell(row=3, column=start_col + 2, value=summary["secondary_name"])
            for h in (h1, h2, h3):
                h.fill = header_fill
                h.font = bold_font
                h.alignment = center

            for idx, item in enumerate(rows_export, 4):
                c1 = ws_detail.cell(row=idx, column=start_col, value=item.get("key"))
                c2 = ws_detail.cell(row=idx, column=start_col + 1, value=item.get("primary"))
                c3 = ws_detail.cell(row=idx, column=start_col + 2, value=item.get("secondary"))
                for dc in (c1, c2, c3):
                    dc.font = body_font
                c1.alignment = center
                c2.alignment = center
                c3.alignment = center

            block_end_row = max(3, 3 + len(rows_export))
            max_block_rows = max(max_block_rows, block_end_row)
            apply_block_borders(start_col, start_col + 2, 2, block_end_row)

            ws_detail.column_dimensions[get_column_letter(start_col)].width = 20
            ws_detail.column_dimensions[get_column_letter(start_col + 1)].width = 24
            ws_detail.column_dimensions[get_column_letter(start_col + 2)].width = 24
            start_col += 4

    for r in range(1, max_block_rows + 1):
        for c in range(1, last_sheet_col + 1):
            ws_detail.cell(row=r, column=c).alignment = center

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    args = parse_args()
    summary_path, detail_path, template_path, output_path = resolve_paths_from_config(args)
    if not summary_path.is_file():
        raise SystemExit(f"Summary file not found: {summary_path}")
    if not detail_path.is_file():
        raise SystemExit(f"Detail file not found: {detail_path}")
    if not template_path.is_file():
        raise SystemExit(f"Excel template not found: {template_path}")

    summary = parse_summary(summary_path)
    detail_rows_by_field = parse_detail(detail_path)
    fill_template(template_path, output_path, summary, detail_rows_by_field)
    print(f"Excel report generated: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())