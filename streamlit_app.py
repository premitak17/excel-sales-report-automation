"""Streamlit interface for the sales report automation."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

from sales_report.report import create_sales_report
from sales_report.validation import clean_sales_data, load_sales_data

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}


def main() -> None:
    st.set_page_config(page_title="Excel Sales Report Automation", page_icon="📊")
    st.title("Excel Sales Report Automation")
    st.write(
        "Upload a sales CSV or Excel workbook to validate the data, review sales "
        "summaries, and download a formatted Excel report."
    )

    uploaded_file = st.file_uploader(
        "Choose a sales data file",
        type=["csv", "xlsx", "xlsm"],
        help="Required columns: order_id, date, product, quantity, and unit_price.",
    )
    if uploaded_file is None:
        st.info("Upload a CSV, XLSX, or XLSM file to get started.")
        return

    file_suffix = Path(uploaded_file.name).suffix.lower()
    if file_suffix not in SUPPORTED_EXTENSIONS:
        st.error("Unsupported file type. Please upload a CSV, XLSX, or XLSM file.")
        return

    try:
        with TemporaryDirectory(prefix="sales_report_") as temporary_directory:
            temporary_path = Path(temporary_directory)
            input_path = temporary_path / f"uploaded_sales{file_suffix}"
            output_path = temporary_path / "sales_report.xlsx"
            input_path.write_bytes(uploaded_file.getvalue())

            data = load_sales_data(input_path)
            result = clean_sales_data(data)
            metrics = create_sales_report(result, output_path)
            report_bytes = output_path.read_bytes()

        monthly_sales = pd.read_excel(
            BytesIO(report_bytes), sheet_name="Monthly Sales", engine="openpyxl"
        )
        top_products = pd.read_excel(
            BytesIO(report_bytes), sheet_name="Top Products", engine="openpyxl"
        )
    except (
        BadZipFile,
        InvalidFileException,
        OSError,
        pd.errors.EmptyDataError,
        pd.errors.ParserError,
        ValueError,
    ) as exc:
        st.error(f"Could not process '{uploaded_file.name}': {exc}")
        return

    st.subheader("Validation summary")
    status_columns = st.columns(3)
    status_columns[0].metric("Valid rows", metrics["valid_rows"])
    status_columns[1].metric("Duplicate rows", metrics["duplicate_rows"])
    status_columns[2].metric("Rejected rows", metrics["rejected_rows"])

    st.subheader("Sales metrics")
    metric_columns = st.columns(3)
    metric_columns[0].metric("Total revenue", f"{metrics['total_revenue']:,.2f}")
    metric_columns[1].metric("Unique orders", metrics["unique_orders"])
    metric_columns[2].metric("Units sold", f"{metrics['units_sold']:,.2f}")

    valid_preview = result.valid.copy()
    valid_preview["revenue"] = valid_preview["quantity"] * valid_preview["unit_price"]

    preview_tabs = st.tabs(
        ["Clean data", "Monthly sales", "Top products", "Duplicates", "Rejected rows"]
    )
    with preview_tabs[0]:
        st.dataframe(valid_preview, use_container_width=True)
    with preview_tabs[1]:
        st.dataframe(monthly_sales, use_container_width=True)
    with preview_tabs[2]:
        st.dataframe(top_products, use_container_width=True)
    with preview_tabs[3]:
        st.dataframe(result.duplicates, use_container_width=True)
    with preview_tabs[4]:
        st.dataframe(result.rejected, use_container_width=True)

    st.download_button(
        "Download Excel report",
        data=report_bytes,
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
