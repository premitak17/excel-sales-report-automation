"""Command-line interface for the sales report generator."""

from __future__ import annotations

import argparse
from pathlib import Path

from sales_report.report import create_sales_report
from sales_report.validation import clean_sales_data, load_sales_data

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate sales data and create a formatted Excel report."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=PROJECT_ROOT / "data" / "sample_sales.csv",
        help="Path to a .csv, .xlsx, or .xlsm sales file.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "output" / "sales_report.xlsx",
        help="Path for the generated Excel report.",
    )
    args = parser.parse_args()

    try:
        data = load_sales_data(args.input)
        result = clean_sales_data(data)
        metrics = create_sales_report(result, args.output)
    except (FileNotFoundError, PermissionError, ValueError, OSError) as exc:
        parser.exit(2, f"Error: {exc}\n")

    print(f"Report created: {args.output.resolve()}")
    print(f"Total revenue: {metrics['total_revenue']:,.2f}")
    print(
        "Rows included: "
        f"{metrics['valid_rows']} valid, "
        f"{metrics['duplicate_rows']} duplicate, "
        f"{metrics['rejected_rows']} rejected."
    )
    return 0