import json
import os
import sys
import datetime
import ast

from sync_stock_lists import sync_stock_lists, StockListError

# Resolve paths relative to this script's own location (not a hardcoded absolute
# path) so this works unchanged after copying the project to another machine/OS.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def count_generated_rows(source, constant):
    """Number of rows in a script's generated stock-list constant (for the markdown cells)."""
    marker = f"{constant} = "
    start = source.index(marker) + len(marker)
    end = source.index("\n]\n", start) + 2
    return len(ast.literal_eval(source[start:end]))


def update_notebook():
    nb_path = os.path.join(SCRIPT_DIR, "Stock_Market_Analysis_SGX_NSE.ipynb")

    # Step 0: regenerate both scripts' hardcoded stock lists from 'stock list.xlsx', so the
    # notebook below is built from the current spreadsheet. Aborts before touching the
    # notebook if the spreadsheet has any invalid rows.
    try:
        sync_stock_lists()
    except StockListError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        print("Notebook not updated.", file=sys.stderr)
        sys.exit(1)

    # Read SGX code
    with open(os.path.join(SCRIPT_DIR, "SGX_Stock_Analysis.py"), "r", encoding="utf-8") as f:
        sgx_lines = f.readlines()

    sgx_code_lines = []
    start_collecting = False
    for line in sgx_lines:
        # NOTE: trigger on FALLBACK_STI_30 (the *first* module-level constant after the
        # docstring/imports/logging setup), which comes before SGX_INDEX_STOCK_ROWS too -
        # get_sti30_tickers() and get_sgx_universe() both reference it on their fallback path.
        if "FALLBACK_STI_30" in line:
            start_collecting = True
        if "def main():" in line:
            break
        if start_collecting:
            sgx_code_lines.append(line)

    sgx_exec = """
# --- Run SGX Extraction ---
today_str = datetime.date.today().strftime('%Y-%m-%d')
sgx_filename = f'SGX_Stock_Analysis_{today_str}.xlsx'

print(f'🚀 Starting SGX Extraction (Default Index Mode: STI + S Cap + L & M Cap + Personal)...')
# Uses the built-in SGX_INDEX_STOCKS list above (generated from 'stock list.xlsx' when this
# notebook was last synced) - no file upload needed.
sgx_summary = extract_index_stocks(max_workers=10)
display(sgx_summary.head(10))

print(f'📊 Extracting 52-Week Daily Historical Data...')
sgx_history, sgx_active_companies = extract_52w_history_all(sgx_summary, max_workers=10)

print(f'📑 Extracting 5-Year Fundamentals History (Revenue, Profit, EPS, P/E, P/S, P/B, margins, D/E)...')
sgx_financials = extract_financial_history_all(sgx_summary, max_workers=6, max_years=4)

save_results(sgx_summary, sgx_history, sgx_active_companies, sgx_filename, sheet_name='SGX Index Summary', financial_histories=sgx_financials)
print(f'✅ SGX Analysis complete! Saved as: {sgx_filename}')
print(f'   Sheets: Summary -> Financials Detail (via each Name) -> Price Charts (via each Symbol)')

# Trigger download if in Google Colab
try:
    from google.colab import files
    files.download(sgx_filename)
except ImportError:
    pass
"""
    sgx_full_code = "# --- SGX Analysis Engine ---\n" + "".join(sgx_code_lines) + sgx_exec

    # Read NSE code
    with open(os.path.join(SCRIPT_DIR, "NSE_Stock_Analysis.py"), "r", encoding="utf-8") as f:
        nse_lines = f.readlines()

    nse_code_lines = []
    start_collecting = False
    for line in nse_lines:
        if "FALLBACK_NIFTY_50" in line:
            start_collecting = True
        if "def main():" in line:
            break
        if start_collecting:
            nse_code_lines.append(line)

    nse_exec = """
# --- Run NSE Extraction ---
today_str = datetime.date.today().strftime('%Y-%m-%d')
nse_filename = f'NSE_Stock_Analysis_{today_str}.xlsx'

print(f'🚀 Starting NSE Extraction (Top 80 + NIFTY 50)...')
nse_summary = extract_top80_with_nifty50(max_workers=10, top_n=80, mode='robust')
display(nse_summary.head(10))

print(f'📊 Extracting 52-Week Daily Historical Data...')
nse_history, nse_active_companies = extract_52w_history_all(nse_summary, max_workers=10)

print(f'📑 Extracting 5-Year Fundamentals History (Revenue, Profit, EPS, P/E, P/S, P/B, margins, D/E)...')
nse_financials = extract_financial_history_all(nse_summary, max_workers=6, max_years=4)

save_results(nse_summary, nse_history, nse_active_companies, nse_filename, financial_histories=nse_financials)
print(f'✅ NSE Analysis complete! Saved as: {nse_filename}')
print(f'   Sheets: Summary -> Financials Detail (via each Name) -> Price Charts (via each Symbol)')

# Trigger download if in Google Colab
try:
    from google.colab import files
    files.download(nse_filename)
except ImportError:
    pass
"""
    nse_full_code = "# --- NSE Analysis Engine ---\n" + "".join(nse_code_lines) + nse_exec

    sgx_stock_count = count_generated_rows("".join(sgx_lines), "SGX_INDEX_STOCK_ROWS")
    nse_personal_count = count_generated_rows("".join(nse_lines), "NSE_PERSONAL_STOCK_ROWS")

    # Shared "Step 1" cell: pip install + the imports/logging setup both scripts need,
    # consolidated once since both run in the same Colab kernel (no need to import
    # pandas/yfinance twice). This is hand-composed, not sliced from either .py file
    # (each script has its own copy of this same setup, since both must stay
    # independently runnable per CLAUDE.md) - kept here rather than left as a bare
    # hand-edited notebook cell, so it goes through the same sync path as everything
    # else and doesn't quietly drift from what the two scripts actually need.
    shared_setup_code = """!pip install -q yfinance openpyxl beautifulsoup4 pandas requests

import os
import re
import sys
import time
import datetime
import logging
import warnings
from typing import List, Dict, Any, Optional, Set
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
import yfinance as yf

# Suppress minor yfinance scraping warnings and logger noise
warnings.filterwarnings('ignore', category=FutureWarning)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# Suppress Colab/Jupyter's own internal datetime.utcnow() DeprecationWarning
# noise (from jupyter_client's message-timestamping code - not this project's
# code, and not fixable from here; see CLAUDE.md). Purely cosmetic, but it
# clutters cell output enough to make the actual pipeline logs hard to read.
warnings.filterwarnings('ignore', category=DeprecationWarning)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger('StockAnalysis')
print('✅ All libraries successfully imported!')
"""

    # Update notebook
    with open(nb_path, "r", encoding="utf-8") as f:
        nb = json.load(f)

    # Remove a previously-inserted "Step 2: Mount Google Drive" cell pair if present
    # (at index 3/4) - NSE's personal stock additions are hardcoded directly in
    # NSE_Stock_Analysis.py now (see PERSONAL_NSE_STOCKS), and SGX's embedded fallback
    # already covers its own personal additions, so neither script needs
    # 'stock list.xlsx' pulled in from Drive anymore. Keeps re-running this script
    # idempotent - only strips the cells once, doesn't error if already gone.
    drive_step_present = (
        len(nb["cells"]) > 3
        and "Mount Google Drive" in "".join(nb["cells"][3].get("source", []))
    )
    if drive_step_present:
        del nb["cells"][3:5]

    # JSON represents lines as a list of strings with newlines for code cells, or just a single string. Let's make it a list or a string.
    # Cell indices: intro=0, deps md/code=1/2, SGX md/code=3/4, NSE md/code=5/6.
    nb["cells"][2]["source"] = [line + ("\n" if not line.endswith("\n") else "") for line in shared_setup_code.split("\n")]
    nb["cells"][4]["source"] = [line + ("\n" if not line.endswith("\n") else "") for line in sgx_full_code.split("\n")]
    nb["cells"][6]["source"] = [line + ("\n" if not line.endswith("\n") else "") for line in nse_full_code.split("\n")]

    # Keep the intro / step markdown cells in sync with what the pipeline actually does now.
    intro_md = f"""# 📈 SGX & NSE Stock Market Capitalization & Historical Analysis
### Complete Pipeline for Singapore Exchange (SGX Index: STI + S Cap + L & M Cap) & National Stock Exchange of India (NSE Top 80 + NIFTY 50)

This Google Colab notebook extracts live fundamentals, 52-week pricing metrics, trailing P/E, sector, industry, a 52-week daily historical time-series, and a 5-year fundamentals history for all top listed companies across **SGX** and **NSE**.

#### Key Features:
- **SGX Analysis**: {sgx_stock_count} stocks - STI, S Cap, and L & M Cap index constituents plus personal additions, built into the code below (generated from `stock list.xlsx` when this notebook was last synced).
- **NSE Analysis**: Top 80 companies by market cap with automatic inclusion of all 50 NIFTY 50 constituents plus {nse_personal_count} built-in personal additions (with `fast`, `robust`, and `full` modes).
- **Multi-Worksheet Excel Output**: Exports a formatted `.xlsx` workbook containing:
  - **Sheet 1 (`Summary`)**: Snapshot metrics formatted to 2 decimal places with B/M/T abbreviations.
  - **Sheet 2 (`52W Daily History`)**: Horizontal time-series layout with 2-tier merged headers (`CODE - Name` over metric sub-columns `Close`, `% Change`, `Market Cap`, `Shares Out.`, `Volume`).
  - **Sheet 3 (`Financials Detail`)**: A stacked 5-year fundamentals block per company (Revenue, Net Income, EPS, P/E, P/S, P/B, Gross/Operating Margin, Debt/Equity) — click any company's **Name** on the Summary sheet to jump straight there.
  - **Sheet 4 (`Price Charts`)**: A compact 1-year closing-price line chart per company — click any company's **Symbol** on the Summary sheet (or its header on the 52W Daily History sheet) to jump straight to its chart.
- **Google Colab Ready**: Includes automatic direct file download triggers for Colab environments. No file uploads or Google Drive mount needed - both pipelines are fully self-contained.
"""
    sgx_step_md = f"""## 🇸🇬 Step 2: Singapore Exchange (SGX Index: STI + S Cap + L & M Cap)
Execute the cell below to run the SGX extraction pipeline. It will:
1. Load the built-in list of {sgx_stock_count} SGX stocks (index constituents plus personal additions).
2. Convert every market cap to SGD using live Yahoo Finance forex rates.
3. Extract snapshot metrics, 52-week historical daily prices, and 5-year fundamentals history.
4. Generate and download `SGX_Stock_Analysis_YYYY-MM-DD.xlsx` (Summary → Financials Detail → Price Charts, all cross-linked).
"""
    nse_step_md = """## 🇮🇳 Step 3: Indian Stock Exchange (NSE Top 80 + NIFTY 50)
Execute the cell below to run the NSE extraction pipeline.

**Mode Options:**
- `'robust'` (Default): Scans **NIFTY 200** (~200 stocks, ~15s, 100% Top 80 coverage).
- `'fast'`: Scans **NIFTY 100** (~100 stocks, ~5s).
- `'full'`: Scans **all 2,553+ NSE listed stocks** (~2m).

Also extracts 52-week historical daily prices and 5-year fundamentals history, producing the same Summary → Financials Detail → Price Charts cross-linked layout as the SGX workbook.
"""
    nb["cells"][0]["source"] = [line + "\n" for line in intro_md.split("\n")][:-1]
    nb["cells"][3]["source"] = [line + "\n" for line in sgx_step_md.split("\n")][:-1]
    nb["cells"][5]["source"] = [line + "\n" for line in nse_step_md.split("\n")][:-1]

    with open(nb_path, "w", encoding="utf-8") as f:
        json.dump(nb, f, indent=1)

if __name__ == "__main__":
    update_notebook()
