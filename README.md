# Excel Sales Report Automation

A small, local Python project that validates sales data from CSV or Excel and
creates a formatted `.xlsx` workbook. It uses pandas for data processing and
openpyxl for Excel input/output and workbook styling. It does not use a paid
API or send data outside your computer.

## Project structure

```text
excel_sales_report_automation/
├── data/
│   └── sample_sales.csv       # Example data, including review cases
├── output/                    # Generated report destination
├── sales_report/
│   ├── __init__.py
│   ├── __main__.py            # Enables python -m sales_report
│   ├── cli.py                 # Command-line options
│   ├── report.py              # Aggregations and Excel formatting
│   └── validation.py          # File loading, cleaning, and validation
├── tests/
│   └── test_sales_report.py
├── requirements.txt
└── run_report.bat             # Windows convenience launcher
```

## Windows setup

Open PowerShell in this project folder, then create and activate a virtual
environment and install the dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If PowerShell blocks activation, run the commands below from Command Prompt
after creating the environment:

```bat
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

Generate a report from the included sample:

```powershell
python -m sales_report
```

The default input is `data\sample_sales.csv`; the report is written to
`output\sales_report.xlsx`. Specify another file and destination with:

```powershell
python -m sales_report --input "path\to\sales.xlsx" --output "path\to\report.xlsx"
```

Supported input files are `.csv`, `.xlsx`, and `.xlsm`. Column names are
case-insensitive and may use spaces or hyphens in place of underscores. The
required columns are `order_id`, `date`, `product`, `quantity`, and `unit_price`.
The date column may also be named `order_date`, `Order Date`, or `Sales Date`;
column names are trimmed and matched without regard to capitalization.
Quantity must be greater than zero; unit price must be zero or greater.

## Data handling and report sheets

Exact repeated records are identified as duplicates; only their first
occurrence is included in the calculations. Rows with missing or invalid
required values are excluded from the calculations. Neither type of row is
discarded: duplicates appear on the **Duplicates** sheet and invalid rows,
with reasons, appear on **Rejected Rows**. The workbook also contains:

- **Summary** — total revenue, unique orders, units, and row counts
- **Monthly Sales** — sales units and revenue by calendar month
- **Top Products** — products ranked by units sold
- **Clean Data** — validated rows, with calculated revenue

The sample deliberately includes one duplicate and three rejected records so
the audit sheets can be inspected immediately.
Revenue is displayed to two decimal places without a currency symbol, because
the input format does not specify a currency.

You can also run the default sample report by double-clicking `run_report.bat`
or calling it from PowerShell:

```powershell
.\run_report.bat
```

## Streamlit web interface

With the project environment and dependencies installed, open PowerShell in
the project folder and start the interface:

```powershell
streamlit run streamlit_app.py
```

The app accepts CSV, XLSX, and XLSM uploads, shows validation results and sales
previews, and provides the generated Excel report as a download. Uploaded data
and the intermediate report are handled in a temporary folder that is removed
after processing.

## Run tests

With the virtual environment active:

```powershell
python -m unittest discover -s tests -v
```