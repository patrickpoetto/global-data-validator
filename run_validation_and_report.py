#!/usr/bin/env python3
"""
Run Global Validator then build the Excel report using validator_config.json defaults.

Usage (from project root):
  python run_validation_and_report.py

Requires: logs/summary_<report_name>.txt and logs/detail_<report_name>.txt produced by
the validator; report goes to reports/report_<report_name>.xlsx.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parent
    py = sys.executable
    v = subprocess.run([py, str(root / "global_validator.py")], cwd=root)
    if v.returncode != 0:
        return v.returncode
    r = subprocess.run([py, str(root / "report_cd.py")], cwd=root)
    return r.returncode


if __name__ == "__main__":
    sys.exit(main())
