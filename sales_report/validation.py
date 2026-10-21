"""Input loading and sales-row validation."""

from __future__ import annotations

from dataclasses import dataclass
from math import isfinite
from pathlib import Path

import pandas as pd

REQUIRED_COLUMNS = ("order_id", "date", "product", "quantity", "unit_price")


@dataclass
class ValidationResult:
    """Validated sales records and rows set aside for review."""

    valid: pd.DataFrame
    duplicates: pd.DataFrame
    rejected: pd.DataFrame


def load_sales_data(path: str | Path) -> pd.DataFrame:
    """Load a CSV or Excel workbook, raising a clear error for unsupported input."""
    input_path = Path(path)
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    suffix = input_path.suffix.lower()
    try:
        if suffix == ".csv":
            return pd.read_csv(input_path, encoding="utf-8-sig")
        if suffix in {".xlsx", ".xlsm"}:
            return pd.read_excel(input_path, engine="openpyxl")
    except (pd.errors.ParserError, pd.errors.EmptyDataError, ValueError, OSError) as exc:
        raise ValueError(f"Could not read input file '{input_path}': {exc}") from exc

    raise ValueError(
        f"Unsupported input format '{suffix}'. Use a .csv, .xlsx, or .xlsm file."
    )


def _normalize_column_name(column: object) -> str:
    normalized = "_".join(str(column).split()).lower().replace("-", "_")
    if normalized in {"order_date", "sales_date"}:
        return "date"
    return normalized


def clean_sales_data(data: pd.DataFrame) -> ValidationResult:
    """Normalize columns, separate exact duplicates, and validate each remaining row."""
    if data.empty and len(data.columns) == 0:
        raise ValueError("The input file has no columns.")

    normalized_columns = [_normalize_column_name(column) for column in data.columns]
    if len(normalized_columns) != len(set(normalized_columns)):
        raise ValueError("Input contains duplicate column names after normalization.")

    frame = data.copy()
    frame.columns = normalized_columns
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in frame]
    if missing_columns:
        missing = ", ".join(missing_columns)
        raise ValueError(
            f"Missing required column(s): {missing}. Required columns are: "
            f"{', '.join(REQUIRED_COLUMNS)}."
        )

    frame = frame.loc[:, REQUIRED_COLUMNS].copy()
    frame.insert(0, "source_row", range(2, len(frame) + 2))

    duplicate_mask = frame.duplicated(subset=list(REQUIRED_COLUMNS), keep="first")
    duplicate_rows = frame.loc[duplicate_mask].copy()
    first_rows = frame.loc[~duplicate_mask].copy()

    if not duplicate_rows.empty:
        first_row_by_record = (
            first_rows.drop_duplicates(subset=list(REQUIRED_COLUMNS))
            .set_index(list(REQUIRED_COLUMNS))["source_row"]
        )
        duplicate_rows["duplicate_of_row"] = [
            int(first_row_by_record.loc[tuple(row[column] for column in REQUIRED_COLUMNS)])
            for _, row in duplicate_rows.iterrows()
        ]

    first_rows["order_id"] = first_rows["order_id"].astype("string").str.strip()
    first_rows["product"] = first_rows["product"].astype("string").str.strip()
    first_rows["date"] = pd.to_datetime(first_rows["date"], errors="coerce")
    first_rows["quantity"] = pd.to_numeric(first_rows["quantity"], errors="coerce")
    first_rows["unit_price"] = pd.to_numeric(first_rows["unit_price"], errors="coerce")

    issues_by_row: list[str] = []
    for _, row in first_rows.iterrows():
        issues: list[str] = []
        for column in ("order_id", "product"):
            if pd.isna(row[column]) or not row[column]:
                issues.append(f"missing {column}")
        if pd.isna(row["date"]):
            issues.append("missing or invalid date")
        if pd.isna(row["quantity"]):
            issues.append("missing or invalid quantity")
        elif not isfinite(float(row["quantity"])):
            issues.append("missing or invalid quantity")
        elif row["quantity"] <= 0:
            issues.append("quantity must be greater than zero")
        if pd.isna(row["unit_price"]):
            issues.append("missing or invalid unit_price")
        elif not isfinite(float(row["unit_price"])):
            issues.append("missing or invalid unit_price")
        elif row["unit_price"] < 0:
            issues.append("unit_price cannot be negative")
        issues_by_row.append("; ".join(issues))

    first_rows["issues"] = issues_by_row
    rejected = first_rows.loc[first_rows["issues"] != ""].copy()
    valid = first_rows.loc[first_rows["issues"] == ""].drop(columns="issues").copy()

    valid["quantity"] = valid["quantity"].astype(float)
    valid["unit_price"] = valid["unit_price"].astype(float)
    return ValidationResult(
        valid=valid.reset_index(drop=True),
        duplicates=duplicate_rows.reset_index(drop=True),
        rejected=rejected.reset_index(drop=True),
    )