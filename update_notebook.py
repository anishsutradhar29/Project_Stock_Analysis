import json
import os
import sys
import datetime

# Resolve paths relative to this script's own location (not a hardcoded absolute
# path) so this works unchanged after copying the project to another machine/OS.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def update_notebook():
    nb_path = os.path.join(SCRIPT_DIR, "Stock_Market_Analysis_SGX_NSE.ipynb")

    # Read SGX code
    with open(os.path.join(SCRIPT_DIR, "SGX_Stock_Analysis.py"), "r", encoding="utf-8") as f:
        sgx_lines = f.readlines()

    sgx_code_lines = []
    start_collecting = False
    for line in sgx_lines:
        # NOTE: trigger on FALLBACK_STI_30 (the *first* module-level constant after the
        # docstring/imports/logging setup) rather than SGX_INDEX_STOCK_MAPPINGS - the old
        # trigger silently dropped FALLBACK_STI_30 itself, which get_sti30_tickers() and
        # get_sgx_universe() both reference on their fallback path (NameError if it's ever hit).
        if "FALLBACK_STI_30" in line or "SGX_INDEX_STOCK_MAPPINGS" in line or "load_sgx_index_stock_list" in line:
            start_collecting = True
        if "def main():" in line:
            break
        if start_collecting:
            sgx_code_lines.append(line)

    sgx_exec = """
# --- Run SGX Extraction ---
today_str = datetime.date.today().strftime('%Y-%m-%d')
sgx_filename = f'SGX_Stock_Analysis_{today_str}.xlsx'

print(f'🚀 Starting SGX Extraction (Default Index Mode: STI + S Cap + L & M Cap)...')
# Default mode extracts from 'SGX Index stock list.xlsx' (falls back to an embedded
# 89-stock list automatically if that file isn't uploaded to this Colab session).
sgx_summary = extract_index_stocks(excel_path='SGX Index stock list.xlsx', max_workers=10)
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

    # Update notebook
    with open(nb_path, "r", encoding="utf-8") as f:
        nb = json.load(f)

    # JSON represents lines as a list of strings with newlines for code cells, or just a single string. Let's make it a list or a string.
    nb["cells"][4]["source"] = [line + ("\n" if not line.endswith("\n") else "") for line in sgx_full_code.split("\n")]
    nb["cells"][6]["source"] = [line + ("\n" if not line.endswith("\n") else "") for line in nse_full_code.split("\n")]

    # Keep the intro / step markdown cells in sync with what the pipeline actually does now.
    intro_md = """# 📈 SGX & NSE Stock Market Capitalization & Historical Analysis
### Complete Pipeline for Singapore Exchange (SGX Index: STI + S Cap + L & M Cap) & National Stock Exchange of India (NSE Top 80 + NIFTY 50)

This Google Colab notebook extracts live fundamentals, 52-week pricing metrics, trailing P/E, sector, industry, a 52-week daily historical time-series, and a 5-year fundamentals history for all top listed companies across **SGX** and **NSE**.

#### Key Features:
- **SGX Analysis**: STI, S Cap, and L & M Cap index constituents (89 companies), read from `SGX Index stock list.xlsx`.
- **NSE Analysis**: Top 80 companies by market cap with automatic inclusion of all 50 NIFTY 50 constituents (with `fast`, `robust`, and `full` modes).
- **Multi-Worksheet Excel Output**: Exports a formatted `.xlsx` workbook containing:
  - **Sheet 1 (`Summary`)**: Snapshot metrics formatted to 2 decimal places with B/M/T abbreviations.
  - **Sheet 2 (`52W Daily History`)**: Horizontal time-series layout with 2-tier merged headers (`CODE - Name` over metric sub-columns `Close`, `% Change`, `Market Cap`, `Shares Out.`, `Volume`).
  - **Sheet 3 (`Financials Detail`)**: A stacked 5-year fundamentals block per company (Revenue, Net Income, EPS, P/E, P/S, P/B, Gross/Operating Margin, Debt/Equity) — click any company's **Name** on the Summary sheet to jump straight there.
  - **Sheet 4 (`Price Charts`)**: A compact 1-year closing-price line chart per company — click any company's **Symbol** on the Summary sheet (or its header on the 52W Daily History sheet) to jump straight to its chart.
- **Google Colab Ready**: Includes automatic direct file download triggers for Colab environments.
"""
    sgx_step_md = """## 🇸🇬 Step 2: Singapore Exchange (SGX Index: STI + S Cap + L & M Cap)
Execute the cell below to run the SGX extraction pipeline. It will:
1. Load the 89 SGX index constituents from `SGX Index stock list.xlsx` (falls back to an embedded list if the file isn't uploaded).
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
