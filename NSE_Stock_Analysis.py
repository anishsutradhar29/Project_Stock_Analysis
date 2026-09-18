"""
================================================================================
NSE & NIFTY 50 Market Cap Extractor (Top 80 + NIFTY 50 Benchmark Analysis)
================================================================================
Collects the latest snapshot of market capitalization, 52-week pricing metrics,
trailing and forward P/E, sector, industry, and custom 52-week range positioning 
for all top National Stock Exchange of India (NSE) listed stocks.

Key Features:
- Latest Snapshot Metrics:
  1. 52-Week High & Low
  2. 52-Week Range (52w High - 52w Low)
  3. Custom Range Metric (%): (52w range - (today's close - 52w low)) / 52w range (%)
  4. Trailing & Forward P/E
  5. Sector & Industry
  6. Today's Close Price & Outstanding Shares
  7. Formatted Market Cap (INR / Crores / Lakh Crores / Trillions)
- Top 80 Extraction & Ranking across broad NSE universe (Nifty 100/200/500).
- NIFTY 50 Catch-Up Inclusion: Automatically appends any NIFTY 50 constituent outside
  the top 80 to the end of the list.
- Single Merged Inclusion Tag: 'NIFTY 50', 'Top 80 Only', or 'NIFTY 50 (Rank #...)'.
- 52-Week Historical Daily Data Worksheet (Horizontal Side-by-Side company layout).
- Polished Multi-Worksheet Excel Output with 2-tier merged headers and 2dp / B / M abbreviation.
================================================================================
"""

import os
import re
import sys
import time
import logging
import argparse
import datetime
import warnings
from typing import List, Dict, Any, Optional, Set
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
import yfinance as yf

# Suppress minor yfinance scraping warnings and logger noise
warnings.filterwarnings("ignore", category=FutureWarning)
logging.getLogger("yfinance").setLevel(logging.CRITICAL)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("NSEMarketCap")

# Fallback top NIFTY 50 constituents in case scraping is blocked
FALLBACK_NIFTY_50 = [
    ("RELIANCE", "Reliance Industries Ltd", "Energy"),
    ("TCS", "Tata Consultancy Services Ltd", "Information Technology"),
    ("HDFCBANK", "HDFC Bank Ltd", "Financial Services"),
    ("BHARTIARTL", "Bharti Airtel Ltd", "Telecommunication"),
    ("ICICIBANK", "ICICI Bank Ltd", "Financial Services"),
    ("INFY", "Infosys Ltd", "Information Technology"),
    ("SBIN", "State Bank of India", "Financial Services"),
    ("ITC", "ITC Ltd", "Consumer Goods"),
    ("HINDUNILVR", "Hindustan Unilever Ltd", "Consumer Goods"),
    ("LT", "Larsen & Toubro Ltd", "Construction"),
    ("BAJFINANCE", "Bajaj Finance Ltd", "Financial Services"),
    ("HCLTECH", "HCL Technologies Ltd", "Information Technology"),
    ("MARUTI", "Maruti Suzuki India Ltd", "Automobile"),
    ("SUNPHARMA", "Sun Pharmaceutical Industries Ltd", "Healthcare"),
    ("TATAMOTORS", "Tata Motors Ltd", "Automobile"),
    ("KOTAKBANK", "Kotak Mahindra Bank Ltd", "Financial Services"),
    ("ONGC", "Oil & Natural Gas Corporation Ltd", "Energy"),
    ("NTPC", "NTPC Ltd", "Power"),
    ("AXISBANK", "Axis Bank Ltd", "Financial Services"),
    ("TITAN", "Titan Company Ltd", "Consumer Goods"),
    ("ADANIENT", "Adani Enterprises Ltd", "Metals & Mining"),
    ("ADANIPORTS", "Adani Ports & Special Economic Zone Ltd", "Services"),
    ("COALINDIA", "Coal India Ltd", "Metals & Mining"),
    ("POWERGRID", "Power Grid Corporation of India Ltd", "Power"),
    ("WIPRO", "Wipro Ltd", "Information Technology"),
    ("M&M", "Mahindra & Mahindra Ltd", "Automobile"),
    ("BAJAJFINSV", "Bajaj Finserv Ltd", "Financial Services"),
    ("ULTRACEMCO", "UltraTech Cement Ltd", "Construction Materials"),
    ("ASIANPAINT", "Asian Paints Ltd", "Consumer Durables"),
    ("JSWSTEEL", "JSW Steel Ltd", "Metals & Mining"),
    ("TATASTEEL", "Tata Steel Ltd", "Metals & Mining"),
    ("GRASIM", "Grasim Industries Ltd", "Construction Materials"),
    ("NESTLEIND", "Nestle India Ltd", "Consumer Goods"),
    ("TECHM", "Tech Mahindra Ltd", "Information Technology"),
    ("HINDALCO", "Hindalco Industries Ltd", "Metals & Mining"),
    ("SBILIFE", "SBI Life Insurance Company Ltd", "Financial Services"),
    ("HDFCLIFE", "HDFC Life Insurance Company Ltd", "Financial Services"),
    ("BRITANNIA", "Britannia Industries Ltd", "Consumer Goods"),
    ("CIPLA", "Cipla Ltd", "Healthcare"),
    ("DRREDDY", "Dr. Reddy's Laboratories Ltd", "Healthcare"),
    ("EICHERMOT", "Eicher Motors Ltd", "Automobile"),
    ("TATACONSUM", "Tata Consumer Products Ltd", "Consumer Goods"),
    ("DIVISLAB", "Divi's Laboratories Ltd", "Healthcare"),
    ("APOLLOHOSP", "Apollo Hospitals Enterprise Ltd", "Healthcare"),
    ("HEROMOTOCO", "Hero MotoCorp Ltd", "Automobile"),
    ("BAJAJ-AUTO", "Bajaj Auto Ltd", "Automobile"),
    ("BPCL", "Bharat Petroleum Corporation Ltd", "Energy"),
    ("SHRIRAMFIN", "Shriram Finance Ltd", "Financial Services"),
    ("INDUSINDBK", "IndusInd Bank Ltd", "Financial Services"),
    ("BEL", "Bharat Electronics Ltd", "Capital Goods"),
]

# Personal stock additions: company names (as they appear in the "Company" column
# of stock list.xlsx's "NSE" worksheet) -> (Code, Display Name). These are guaranteed
# inclusion in the final output (Inclusion_Reason "Personal"), same idea as the NIFTY 50
# guarantee below, but purely additive - unlike SGX's index list, the NSE pipeline
# already has a full stock universe without this file, so there's no embedded fallback
# list here; if the file/sheet is missing, personal additions are simply skipped.
NSE_PERSONAL_STOCK_MAPPINGS = {
    "ACC LTD": ("ACC", "ACC Ltd"),
    "AMBUJA CEMENTS LTD": ("AMBUJACEM", "Ambuja Cements Ltd"),
    "ASHOK LEYLAND LTD": ("ASHOKLEY", "Ashok Leyland Ltd"),
    "INDIAN HOTELS CO LTD": ("INDHOTEL", "Indian Hotels Co Ltd"),
    "ITC HOTELS LIMITED": ("ITCHOTELS", "ITC Hotels Ltd"),
    "KWALITY WALLS (INDIA) LIMITED": ("KWIL", "Kwality Wall's (India) Ltd"),
    "MARICO LTD (MARICO INDUSTRIES LTD)": ("MARICO", "Marico Ltd"),
    "SIEMENS ENERGY INDIA LIMITED": ("ENRIN", "Siemens Energy India Ltd"),
}


def clean_text(text: str) -> str:
    """Strip Wikipedia citation brackets [1], [zh], etc."""
    return re.sub(r"\[.*?\]", "", text).strip()


def format_currency(value: Optional[float], currency: str = "INR") -> str:
    """Format numeric currency value to readable shorthand (T, B, M, K)."""
    if value is None or pd.isna(value) or value <= 0:
        return "N/A"
    
    prefix = f"{currency} " if currency else "₹"
    if value >= 1_000_000_000_000:
        return f"{prefix}{value / 1_000_000_000_000:.2f}T"
    elif value >= 1_000_000_000:
        return f"{prefix}{value / 1_000_000_000:.2f}B"
    elif value >= 1_000_000:
        return f"{prefix}{value / 1_000_000:.2f}M"
    elif value >= 1_000:
        return f"{prefix}{value / 1_000:.2f}K"
    return f"{prefix}{value:.2f}"


def format_abbreviated(value: Optional[float]) -> str:
    """Format large numbers (shares, volume, market cap) with B/M suffix and 2 decimal points."""
    if value is None or pd.isna(value):
        return "N/A"
    try:
        val = float(value)
    except (ValueError, TypeError):
        return str(value)
        
    abs_val = abs(val)
    if abs_val >= 1_000_000_000_000:
        return f"{val / 1_000_000_000_000:.2f}T"
    elif abs_val >= 1_000_000_000:
        return f"{val / 1_000_000_000:.2f}B"
    elif abs_val >= 1_000_000:
        return f"{val / 1_000_000:.2f}M"
    elif abs_val >= 1_000:
        return f"{val / 1_000:.2f}K"
    else:
        return f"{val:.2f}"


def safe_float(val: Any) -> Optional[float]:
    """Safely convert value to float, handling strings like 'Infinity', None, NaN."""
    if val is None or pd.isna(val):
        return None
    try:
        f = float(val)
        return f if f == f else None  # Check NaN
    except (ValueError, TypeError):
        return None


def get_nifty50_tickers(with_suffix: bool = True, timeout: int = 10) -> List[str]:
    """
    Extract the 50 NIFTY 50 constituent symbols from Wikipedia or official CSV.
    """
    url = "https://en.wikipedia.org/wiki/NIFTY_50"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    try:
        response = requests.get(url, headers=headers, timeout=timeout)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            table = soup.find("table", {"id": "constituents"}) or soup.find("table", {"class": "wikitable"})

            if table:
                rows = table.find_all("tr")[1:]
                tickers = []
                for row in rows:
                    cols = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                    if len(cols) >= 2:
                        sym = clean_text(cols[1]).strip().upper()
                        tickers.append(sym)
                if len(tickers) >= 45:
                    logger.info(f"Discovered {len(tickers)} NIFTY 50 constituents from Wikipedia.")
                    return [f"{c}.NS" if with_suffix else c for c in tickers]
    except Exception as e:
        logger.warning(f"Failed to scrape NIFTY 50 from Wikipedia: {e}")

    # Fallback to official Nifty 50 CSV list from niftyindices
    try:
        csv_url = "https://niftyindices.com/IndexConstituent/ind_nifty50list.csv"
        resp = requests.get(csv_url, headers=headers, timeout=timeout)
        if resp.status_code == 200:
            lines = [line.strip().split(",") for line in resp.text.strip().split("\n")[1:] if line.strip()]
            tickers = [cols[2].strip().upper() for cols in lines if len(cols) >= 3 and cols[2].strip()]
            if len(tickers) >= 45:
                logger.info(f"Discovered {len(tickers)} NIFTY 50 constituents from NiftyIndices CSV.")
                return [f"{c}.NS" if with_suffix else c for c in tickers]
    except Exception:
        pass

    logger.info("Using hardcoded fallback list for NIFTY 50 constituents.")
    return [f"{code}.NS" if with_suffix else code for code, _, _ in FALLBACK_NIFTY_50]


def load_nse_personal_stocks(excel_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Extract personal stock additions from the shared 'stock list.xlsx' workbook's
    'NSE' worksheet (Inclusion reason, Company, and an optional Ticker column).
    Purely additive - if the file or sheet isn't present, returns an empty list
    rather than falling back to anything, since the rest of the pipeline already
    has a full stock universe without it.

    A non-blank Ticker cell (e.g. 'ACC.NS') is used directly - no name lookup
    needed. A blank one falls back to matching the Company name against
    NSE_PERSONAL_STOCK_MAPPINGS; unmatched rows are skipped with a warning.

    Returns a list of dicts: [{'code': '...', 'symbol': '...NS', 'name': '...'}, ...]
    """
    candidate_paths = []
    if excel_path:
        candidate_paths.append(excel_path)
    candidate_paths.extend([
        "stock list.xlsx",
        os.path.join(os.getcwd(), "stock list.xlsx"),
    ])
    try:
        # See the matching note in SGX_Stock_Analysis.py's load_sgx_index_stock_list:
        # __file__ is only defined when running as a real .py file, not inside a
        # notebook cell - this candidate is simply skipped there.
        candidate_paths.append(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "stock list.xlsx")
        )
    except NameError:
        pass

    resolved_path = None
    for p in candidate_paths:
        if p and os.path.exists(p):
            resolved_path = p
            break

    if not resolved_path:
        logger.info("'stock list.xlsx' not found - skipping personal stock additions.")
        return []

    try:
        import openpyxl
        wb = openpyxl.load_workbook(resolved_path)
        if "NSE" not in wb.sheetnames:
            logger.info("'stock list.xlsx' has no 'NSE' worksheet - skipping personal stock additions.")
            return []
        sheet = wb["NSE"]

        stocks = []
        seen_codes = set()
        for row in sheet.iter_rows(values_only=True):
            if not any(row):
                continue
            vals = [str(c).strip() for c in row if c is not None]
            if "Company" in vals and ("Inclusion reason" in vals or "Inclusion Reason" in vals):
                continue  # header row

            non_none = [c for c in row if c is not None]
            if len(non_none) < 2:
                continue
            comp = str(non_none[1]).strip()
            # Optional 3rd column: a real ticker (e.g. 'ACC.NS') straight from the sheet -
            # present since the "Ticker" column was added. Present and non-blank takes
            # priority over name-matching below; absent/blank falls back to it unchanged.
            ticker = str(non_none[2]).strip() if len(non_none) >= 3 else None

            if ticker:
                sym = ticker
                code = sym.split(".")[0]
                dict_entry = NSE_PERSONAL_STOCK_MAPPINGS.get(comp.upper())
                display_name = dict_entry[1] if dict_entry else comp
            else:
                mapping = NSE_PERSONAL_STOCK_MAPPINGS.get(comp.upper())
                if mapping is None:
                    logger.warning(f"No NSE ticker mapping for personal stock '{comp}' - skipping (add a Ticker column value, or add it to NSE_PERSONAL_STOCK_MAPPINGS).")
                    continue
                code, display_name = mapping
                sym = f"{code}.NS"

            if code not in seen_codes:
                stocks.append({"code": code, "symbol": sym, "name": display_name})
                seen_codes.add(code)

        if stocks:
            logger.info(f"Loaded {len(stocks)} personal stock(s) from '{resolved_path}' ('NSE' worksheet).")
        return stocks

    except Exception as e:
        logger.warning(f"Error reading personal stocks from '{resolved_path}': {e}. Skipping personal additions.")
        return []


def get_nse_universe(nifty50_codes: Set[str], mode: str = "robust") -> List[Dict[str, Any]]:
    """
    Fetch NSE listed equities according to the chosen analysis mode:
      - 'fast'   : NIFTY 100 constituents (~100 Large Cap stocks, fastest ~5-10s)
      - 'robust' : NIFTY 200 constituents (~200 Large + Mid Cap stocks, default ~15-20s, 100% Top 80 coverage)
      - 'full'   : All listed equities on the NSE master directory (~2,550+ stocks, ~1.5-2.5m)
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    }
    securities: List[Dict[str, Any]] = []
    seen_codes: Set[str] = set()

    mode = mode.lower().strip()
    logger.info(f"Loading NSE stock universe in [{mode.upper()}] mode...")

    # --- Mode: FULL (Entire NSE Directory of ~2,550+ stocks) ---
    if mode == "full":
        full_equity_url = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
        try:
            logger.info("Fetching complete master directory of all NSE listed equities (~2,550+ stocks)...")
            resp = requests.get(full_equity_url, headers=headers, timeout=15)
            if resp.status_code == 200:
                lines = [l.strip().split(",") for l in resp.text.strip().split("\n")[1:] if l.strip()]
                for cols in lines:
                    if len(cols) >= 2:
                        code = cols[0].strip().upper().replace('"', '')
                        name = cols[1].strip().strip('"')
                        series = cols[2].strip().upper() if len(cols) >= 3 else "EQ"
                        
                        if code and code not in seen_codes:
                            securities.append({
                                "code": code,
                                "symbol": f"{code}.NS",
                                "name": name,
                                "series": series,
                                "is_nifty50": code in nifty50_codes
                            })
                            seen_codes.add(code)
                            
                if len(securities) > 1000:
                    logger.info(f"Successfully loaded full NSE universe: {len(securities)} listed equities.")
                    return securities
        except Exception as e:
            logger.warning(f"Could not load full NSE equity list ({e}). Falling back to Nifty 500...")

    # --- Mode: FAST (NIFTY 100) or ROBUST (NIFTY 200, Default) ---
    if mode == "fast":
        target_urls = ["https://niftyindices.com/IndexConstituent/ind_nifty100list.csv"]
    else:  # 'robust' (default) or fallback
        target_urls = [
            "https://niftyindices.com/IndexConstituent/ind_nifty200list.csv",
            "https://niftyindices.com/IndexConstituent/ind_nifty100list.csv",
        ]

    for url in target_urls:
        try:
            resp = requests.get(url, headers=headers, timeout=12)
            if resp.status_code == 200:
                lines = [l.strip().split(",") for l in resp.text.strip().split("\n")[1:] if l.strip()]
                for cols in lines:
                    if len(cols) >= 3:
                        name = cols[0].strip().strip('"')
                        code = cols[2].strip().upper()
                        if code and code not in seen_codes:
                            securities.append({
                                "code": code,
                                "symbol": f"{code}.NS",
                                "name": name,
                                "is_nifty50": code in nifty50_codes
                            })
                            seen_codes.add(code)
                if len(securities) >= 90:
                    logger.info(f"Loaded {len(securities)} NSE listed companies from {url.split('/')[-1]}.")
                    break
        except Exception as e:
            logger.warning(f"Could not load {url}: {e}")

    # Ensure all NIFTY 50 constituents are guaranteed present
    for n50_code in nifty50_codes:
        if n50_code not in seen_codes:
            securities.append({
                "code": n50_code,
                "symbol": f"{n50_code}.NS",
                "name": n50_code,
                "is_nifty50": True
            })
            seen_codes.add(n50_code)

    if not securities:
        for code, name, _ in FALLBACK_NIFTY_50:
            securities.append({
                "code": code,
                "symbol": f"{code}.NS",
                "name": name,
                "is_nifty50": True
            })

    return securities


def fetch_latest_ticker_data(security: Dict[str, Any], max_retries: int = 2) -> Dict[str, Any]:
    """
    Fetch the latest market cap, 52-week pricing, trailing/forward P/E, sector, and industry for a single NSE ticker.
    Calculates the 52-week range and the percentage positioning metric.
    """
    symbol = security["symbol"]
    code = security["code"]
    initial_name = security.get("name", code)
    is_nifty50 = security.get("is_nifty50", False)

    record = {
        "Symbol": symbol,
        "Code": code,
        "Name": initial_name,
        "Is_NIFTY50": is_nifty50,
        "_MarketCapRaw": None,       # Internal only for sorting; not exported
        "MarketCap": "N/A",
        "TodaysClose": None,
        "SharesOutstanding": None,
        "Currency": "INR",
        "52WHigh": None,
        "52WLow": None,
        "52WRange": None,
        "Current_Proximity_to_52W_High": None,
        "Trailing_PE": None,
        "Sector": "N/A",
        "Industry": "N/A",
        "DividendYield": None,
    }

    for attempt in range(1, max_retries + 1):
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info or {}
            
            if info:
                name = info.get("shortName") or info.get("longName")
                if name and not ("," in name and "0P00" in name):
                    record["Name"] = name
                
                # Market Cap (raw kept internally for sorting)
                raw_mktcap = safe_float(info.get("marketCap"))
                
                # Today's Close / Current Price
                price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"))
                record["TodaysClose"] = round(price, 2) if price is not None else None
                record["Currency"] = info.get("currency") or "INR"
                
                # Fallback: calculate market cap from price * shares if missing
                if raw_mktcap is None and price is not None and info.get("sharesOutstanding"):
                    shares = safe_float(info.get("sharesOutstanding"))
                    if shares:
                        raw_mktcap = price * shares

                # Fallback: use fast_info for market cap
                if raw_mktcap is None:
                    try:
                        raw_mktcap = safe_float(getattr(ticker.fast_info, 'market_cap', None))
                    except Exception:
                        pass

                record["_MarketCapRaw"] = raw_mktcap

                # Shares Outstanding (needed to compute per-day Market Cap on the 52W History sheet)
                shares_out = safe_float(info.get("sharesOutstanding") or info.get("impliedSharesOutstanding"))
                record["SharesOutstanding"] = round(shares_out, 0) if shares_out is not None else None

                # 52-Week High & Low
                h52 = safe_float(info.get("fiftyTwoWeekHigh"))
                l52 = safe_float(info.get("fiftyTwoWeekLow"))
                record["52WHigh"] = round(h52, 2) if h52 is not None else None
                record["52WLow"]  = round(l52, 2) if l52 is not None else None

                # 52-Week Range & Proximity to 52W High metric
                # Formula: (price - 52wLow) / (52wHigh - 52wLow) * 100
                # 0% = price is AT the 52W low; 100% = price is AT the 52W high
                if h52 is not None and l52 is not None:
                    rng = h52 - l52
                    record["52WRange"] = round(rng, 2)
                    
                    if rng > 0 and price is not None:
                        proximity_pct = ((price - l52) / rng) * 100.0
                        record["Current_Proximity_to_52W_High"] = f"{proximity_pct:.2f}%"

                # P/E ratio, Sector, Industry
                trailing_pe = safe_float(info.get("trailingPE"))
                record["Trailing_PE"]  = round(trailing_pe, 2) if trailing_pe is not None else None
                record["Sector"]       = info.get("sector")   or "N/A"
                record["Industry"]     = info.get("industry") or "N/A"
                div_yield = safe_float(info.get("dividendYield"))
                record["DividendYield"] = round(div_yield, 4) if div_yield is not None else None

            # If ticker.info failed or returned nothing (e.g. rate limit), fallback to direct chart API
            if not info or record["TodaysClose"] is None:
                try:
                    c_url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=1d&interval=1d"
                    c_headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
                    }
                    c_resp = requests.get(c_url, headers=c_headers, timeout=8)
                    if c_resp.status_code == 200:
                        c_data = c_resp.json()
                        c_res = c_data.get("chart", {}).get("result", [{}])[0]
                        c_meta = c_res.get("meta", {})
                        
                        if c_meta:
                            name = c_meta.get("shortName") or c_meta.get("longName")
                            if name and not ("," in name and "0P00" in name):
                                record["Name"] = name
                            
                            price = safe_float(c_meta.get("regularMarketPrice") or c_meta.get("chartPreviousClose"))
                            if price is not None:
                                record["TodaysClose"] = round(price, 2)
                                record["Currency"] = c_meta.get("currency") or "INR"
                                
                            h52 = safe_float(c_meta.get("fiftyTwoWeekHigh"))
                            l52 = safe_float(c_meta.get("fiftyTwoWeekLow"))
                            if h52 is not None:
                                record["52WHigh"] = round(h52, 2)
                            if l52 is not None:
                                record["52WLow"] = round(l52, 2)
                                
                            if h52 is not None and l52 is not None:
                                rng = h52 - l52
                                record["52WRange"] = round(rng, 2)
                                if rng > 0 and price is not None:
                                    proximity_pct = ((price - l52) / rng) * 100.0
                                    record["Current_Proximity_to_52W_High"] = f"{proximity_pct:.2f}%"
                                    
                            # Fallback for Market Cap if still None (chart API lacks it)
                            if record["_MarketCapRaw"] is None:
                                try:
                                    record["_MarketCapRaw"] = safe_float(getattr(ticker.fast_info, 'market_cap', None))
                                except Exception:
                                    pass
                except Exception:
                    pass

            # Finalize formatted market cap string
            if record["_MarketCapRaw"] is not None and record["_MarketCapRaw"] > 0:
                record["MarketCap"] = format_currency(record["_MarketCapRaw"], record["Currency"])
                return record
            else:
                if attempt == max_retries:
                    return record

        except Exception as e:
            if attempt == max_retries:
                return record
            time.sleep(0.3 * attempt)

    return record


def extract_top80_with_nifty50(
    max_workers: int = 10,
    top_n: int = 80,
    mode: str = "robust",
    stock_list_file: Optional[str] = None
) -> pd.DataFrame:
    """
    Extract Top 80 NSE companies by market cap + guaranteed inclusion of all NIFTY 50
    constituents + guaranteed inclusion of any "Personal" stocks from stock list.xlsx's
    'NSE' worksheet. Supports mode: 'fast' (Nifty 100), 'robust' (Nifty 200, default),
    'full' (all 2,550+ stocks).
    """
    start_time = time.time()
    logger.info(f"Starting NSE Top {top_n} & NIFTY 50 market cap extraction in [{mode.upper()}] mode...")

    # 1. Get NIFTY 50 Constituents & Broad NSE Universe
    nifty50_symbols = get_nifty50_tickers(with_suffix=False)
    nifty50_codes = set(nifty50_symbols)
    logger.info(f"Found {len(nifty50_codes)} NIFTY 50 constituent codes.")

    securities = get_nse_universe(nifty50_codes, mode=mode)

    # Merge in any personal stocks not already covered by the broad universe above,
    # so they get fetched and ranked in the same pass (same guarantee pattern used
    # for NIFTY 50 further below).
    personal_stocks = load_nse_personal_stocks(excel_path=stock_list_file)
    personal_codes = {p["code"] for p in personal_stocks}
    existing_codes = {s["code"] for s in securities}
    for p in personal_stocks:
        if p["code"] not in existing_codes:
            securities.append({
                "code": p["code"],
                "symbol": p["symbol"],
                "name": p["name"],
                "is_nifty50": p["code"] in nifty50_codes,
            })
            existing_codes.add(p["code"])

    total_count = len(securities)
    logger.info(f"Querying {total_count} NSE instruments across {max_workers} concurrent threads...")

    # 2. Fetch snapshot data concurrently
    results = []
    fetch_start = time.time()
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_sec = {executor.submit(fetch_latest_ticker_data, sec): sec for sec in securities}
        
        for future in as_completed(future_to_sec):
            data = future.result()
            results.append(data)
            completed_count += 1

            if completed_count % 25 == 0 or completed_count == total_count:
                elapsed = time.time() - fetch_start
                pct = (completed_count / total_count) * 100
                rate = completed_count / elapsed if elapsed > 0 else 0
                logger.info(f"Progress: {completed_count}/{total_count} ({pct:.1f}%) | {rate:.1f} tickers/sec")

    # 3. Targeted Retry Pass for any missing NIFTY 50 constituents or Personal stocks -
    #    both are guaranteed-inclusion picks (see steps 6-7 below), so it's worth a
    #    dedicated retry rather than letting a transient fetch failure silently drop them.
    guaranteed_codes = nifty50_codes | personal_codes
    failed_nifty = [r for r in results if r["Code"] in guaranteed_codes and (r["_MarketCapRaw"] is None or r["_MarketCapRaw"] <= 0)]
    if failed_nifty:
        logger.info(f"Running dedicated retry pass for {len(failed_nifty)} guaranteed-inclusion constituent(s)...")
        time.sleep(1.0)
        for item in failed_nifty:
            refetched = fetch_latest_ticker_data(
                {"symbol": item["Symbol"], "code": item["Code"], "name": item["Name"], "is_nifty50": item["Code"] in nifty50_codes},
                max_retries=3
            )
            if refetched["_MarketCapRaw"] is not None:
                for idx, orig in enumerate(results):
                    if orig["Code"] == item["Code"]:
                        results[idx] = refetched
                        break
            time.sleep(0.2)

    full_df = pd.DataFrame(results)

    # 4. Global Ranking across entire NSE universe
    full_df["_SortCap"] = full_df["_MarketCapRaw"].fillna(-1)
    full_df = full_df.sort_values(by=["_SortCap"], ascending=False).reset_index(drop=True)
    full_df = full_df.drop(columns=["_SortCap"])

    # Assign NSE_Rank (1-indexed) only to rows with valid market cap
    valid_cap_mask = full_df["_MarketCapRaw"].notnull() & (full_df["_MarketCapRaw"] > 0)
    full_df["NSE_Rank"] = None
    full_df.loc[valid_cap_mask, "NSE_Rank"] = range(1, int(valid_cap_mask.sum()) + 1)

    # 5. Extract Top N Companies
    top_n_df = full_df.iloc[:top_n].copy()

    # Single Merged Tag: "NIFTY 50" if in index, else "Top 80 Only"
    top_n_df["Inclusion_Reason"] = top_n_df["Code"].apply(
        lambda c: "NIFTY 50" if c in nifty50_codes else f"Top {top_n} Only"
    )

    top_n_nifty_codes = set(top_n_df[top_n_df["Code"].isin(nifty50_codes)]["Code"])
    missing_nifty_codes = nifty50_codes - top_n_nifty_codes

    logger.info(f"NIFTY 50 constituents inside Top {top_n}: {len(top_n_nifty_codes)}/50")
    if missing_nifty_codes:
        logger.info(f"NIFTY 50 constituents outside Top {top_n} ({len(missing_nifty_codes)}): {sorted(list(missing_nifty_codes))}")
        
        # 6. Extract the missing NIFTY 50 constituents and append to end of list
        missing_nifty_df = full_df[full_df["Code"].isin(missing_nifty_codes)].copy()
        missing_nifty_df["Inclusion_Reason"] = missing_nifty_df["NSE_Rank"].apply(
            lambda r: f"NIFTY 50 (Rank #{r})" if pd.notnull(r) else "NIFTY 50"
        )
        combined_df = pd.concat([top_n_df, missing_nifty_df], ignore_index=True)
    else:
        logger.info(f"All 50 NIFTY constituents are already present in the Top {top_n}!")
        combined_df = top_n_df

    # 7. Guarantee inclusion of "Personal" stocks (from stock list.xlsx's 'NSE'
    #    worksheet), same pattern as the NIFTY 50 guarantee above. A stock already
    #    present for another reason (e.g. NIFTY 50 or Top N) keeps that reason and
    #    gets "+ Personal" appended, rather than losing the original context.
    if personal_codes:
        in_combined_mask = combined_df["Code"].isin(personal_codes)
        combined_df.loc[in_combined_mask, "Inclusion_Reason"] = combined_df.loc[in_combined_mask, "Inclusion_Reason"].apply(
            lambda r: r if "Personal" in str(r) else f"{r} + Personal"
        )

        already_covered = set(combined_df.loc[in_combined_mask, "Code"])
        missing_personal_codes = personal_codes - already_covered
        if missing_personal_codes:
            missing_personal_df = full_df[full_df["Code"].isin(missing_personal_codes)].copy()
            missing_personal_df["Inclusion_Reason"] = missing_personal_df["NSE_Rank"].apply(
                lambda r: f"Personal (Rank #{r})" if pd.notnull(r) else "Personal"
            )
            combined_df = pd.concat([combined_df, missing_personal_df], ignore_index=True)

    # Drop internal columns
    combined_df = combined_df.drop(columns=["Is_NIFTY50", "_MarketCapRaw"], errors="ignore")

    ordered_cols = [
        "NSE_Rank", "Symbol", "Code", "Name",
        "Inclusion_Reason", "MarketCap", "TodaysClose", "SharesOutstanding",
        "52WHigh", "52WLow", "52WRange", "Current_Proximity_to_52W_High",
        "Trailing_PE",
        "Sector", "Industry", "Currency", "DividendYield",
    ]
    existing_cols = [c for c in ordered_cols if c in combined_df.columns]
    combined_df = combined_df[existing_cols]

    elapsed_total = time.time() - start_time
    logger.info(f"Extraction completed in {elapsed_total:.2f}s. Total output rows: {len(combined_df)}.")
    return combined_df


def print_top80_nifty_table(df: pd.DataFrame):
    """
    Display a formatted terminal table with all metrics.
    """
    W = 142
    print("\n" + "=" * W)
    print(" TOP 80 NSE COMPANIES & NIFTY 50 BENCHMARK CONSTITUENTS")
    print("=" * W)

    header = (
        f"{'Rank':<5} | {'Symbol':<12} | {'Company Name':<22} | "
        f"{'Inclusion':<20} | {'Market Cap':<14} | {'Close':<8} | "
        f"{'52W H':<8} | {'52W L':<8} | {'52W Rng':<8} | {'Prox. 52W High':<15} | "
        f"{'Tr.PE':<6} | {'Sector'}"
    )
    print(header)
    print("-" * W)

    overflow_count = 0

    for _, row in df.iterrows():
        nse_rank   = f"#{int(row['NSE_Rank'])}" if pd.notnull(row.get('NSE_Rank')) else "N/A"
        symbol     = str(row["Symbol"])
        name       = (row["Name"][:20] + "..") if len(str(row["Name"])) > 22 else str(row["Name"])
        reason     = str(row["Inclusion_Reason"])
        mkt_cap    = str(row.get("MarketCap", "N/A"))

        p = safe_float(row.get('TodaysClose'))
        price = f"{p:.2f}" if p is not None else "N/A"

        h = safe_float(row.get('52WHigh'))
        h52 = f"{h:.2f}" if h is not None else "N/A"

        l = safe_float(row.get('52WLow'))
        l52 = f"{l:.2f}" if l is not None else "N/A"

        r = safe_float(row.get('52WRange'))
        rng = f"{r:.2f}" if r is not None else "N/A"

        prox = str(row.get("Current_Proximity_to_52W_High", "N/A"))
        if prox in ("nan", "", "None"):
            prox = "N/A"

        tpe = safe_float(row.get('Trailing_PE'))
        trailing_pe = f"{tpe:.2f}" if tpe is not None else "N/A"

        sec_str = str(row.get("Sector", "N/A"))
        sector = (sec_str[:18] + "..") if len(sec_str) > 20 else sec_str

        if "Rank #" in reason and overflow_count == 0:
            print("-" * W)
            print(" >>> NIFTY 50 CONSTITUENTS OUTSIDE TOP 80 (APPENDED BELOW):")
            print("-" * W)

        if "Rank #" in reason:
            overflow_count += 1

        print(
            f"{nse_rank:<5} | {symbol:<12} | {name:<22} | "
            f"{reason:<20} | {mkt_cap:<14} | {price:<8} | "
            f"{h52:<8} | {l52:<8} | {rng:<8} | {prox:<15} | "
            f"{trailing_pe:<6} | {sector}"
        )

    print("=" * W)
    nifty_in_top80 = df[df["Inclusion_Reason"] == "NIFTY 50"].shape[0]
    top80_only     = df[df["Inclusion_Reason"] == "Top 80 Only"].shape[0]
    print(f"NIFTY 50 Members in Top 80      : {nifty_in_top80} / 50")
    print(f"Non-NIFTY Top 80 Companies      : {top80_only}")
    print(f"NIFTY 50 Members outside Top 80 : {overflow_count} (Appended to end)")
    print(f"Total Companies in Output       : {len(df)}\n")


def fetch_ticker_52w_history(symbol: str, code: str, name: str, shares_outstanding: float = None, max_retries: int = 3) -> pd.DataFrame:
    """
    Fetch 52-week (1 year) historical daily trading data for a single NSE ticker.
    Uses direct Yahoo Chart API with session headers to bypass rate limits.
    Collects Date, Symbol, Code, Name, Close, DailyChangePct, MarketCap, SharesOutstanding, Volume.
    """
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=1y&interval=1d"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
    
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                results = data.get("chart", {}).get("result")
                if results and len(results) > 0:
                    chart = results[0]
                    timestamps = chart.get("timestamp", [])
                    indicators = chart.get("indicators", {}).get("quote", [{}])[0]
                    closes = indicators.get("close", [])
                    volumes = indicators.get("volume", [])
                    
                    if timestamps and closes:
                        dates = [datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d") for ts in timestamps]
                        
                        df_hist = pd.DataFrame({
                            "Date": dates,
                            "Close": closes,
                            "Volume": volumes
                        }).dropna(subset=["Close"])
                        
                        df_hist["Symbol"] = symbol
                        df_hist["Code"] = code
                        df_hist["Name"] = name
                        
                        # % Change in daily close price
                        df_hist["DailyChangePct"] = df_hist["Close"].pct_change() * 100.0
                        
                        # Outstanding shares and historical daily market cap calculation
                        df_hist["SharesOutstanding"] = shares_outstanding
                        if shares_outstanding is not None:
                            df_hist["MarketCap"] = (df_hist["Close"] * shares_outstanding).round(2)
                        else:
                            df_hist["MarketCap"] = None
                            
                        df_hist["Close"] = df_hist["Close"].round(2)
                        df_hist["DailyChangePct"] = df_hist["DailyChangePct"].round(2)
                        df_hist["Volume"] = df_hist["Volume"].fillna(0).astype(int)
                        
                        target_cols = [
                            "Date", "Symbol", "Code", "Name", "Close",
                            "DailyChangePct", "MarketCap", "SharesOutstanding", "Volume"
                        ]
                        existing_cols = [c for c in target_cols if c in df_hist.columns]
                        return df_hist[existing_cols]

            time.sleep(0.3 * attempt)
        except Exception as e:
            if attempt == max_retries:
                logger.warning(f"Failed to fetch 52W history for {symbol}: {e}")
            time.sleep(0.4 * attempt)
            
    return pd.DataFrame()


def extract_52w_history_all(summary_df: pd.DataFrame, max_workers: int = 10):
    """
    Extract 52-week historical daily data for all companies present in summary_df concurrently.
    Reformats the dataset horizontally such that each new company is added as a column group 
    to the right (aligned along Date rows) in the order of the Top 80 / NIFTY 50 rank.
    Returns (wide_hist_df, active_companies_metadata).
    """
    logger.info(f"Starting 52-week daily history extraction for {len(summary_df)} NSE companies...")
    start_time = time.time()
    
    tasks = []
    ordered_symbols = []
    meta_map = {}
    
    for _, row in summary_df.iterrows():
        sym = row["Symbol"]
        code = row["Code"]
        name = row["Name"]
        ordered_symbols.append(sym)
        meta_map[sym] = (code, name)
        tasks.append({
            "symbol": sym,
            "code": code,
            "name": name,
            "shares": safe_float(row.get("SharesOutstanding"))
        })
        
    hist_results = {}
    completed_count = 0
    total_count = len(tasks)
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {
            executor.submit(
                fetch_ticker_52w_history,
                t["symbol"],
                t["code"],
                t["name"],
                t["shares"]
            ): t for t in tasks
        }
        
        for future in concurrent.futures.as_completed(future_to_task):
            t = future_to_task[future]
            try:
                hist_df = future.result()
                if not hist_df.empty:
                    hist_results[t["symbol"]] = hist_df
            except Exception as e:
                logger.error(f"Error fetching history for {t['symbol']}: {e}")
                
            completed_count += 1
            if completed_count % 15 == 0 or completed_count == total_count:
                logger.info(f"History progress: {completed_count}/{total_count} ({completed_count / total_count * 100:.1f}%)")

    # Merge all historical data horizontally by Date
    wide_hist_df = None
    active_companies = []
    
    for sym in ordered_symbols:
        if sym not in hist_results or hist_results[sym].empty:
            continue
        sub = hist_results[sym][["Date", "Close", "DailyChangePct", "MarketCap", "SharesOutstanding", "Volume"]].copy()
        code, name = meta_map.get(sym, (sym, sym))
        active_companies.append((sym, code, name))
        
        label = f"{code} ({name})"
        sub = sub.rename(columns={
            "Close": f"{label} - Close",
            "DailyChangePct": f"{label} - % Change",
            "MarketCap": f"{label} - Market Cap",
            "SharesOutstanding": f"{label} - Shares Out.",
            "Volume": f"{label} - Volume",
        })
        
        if wide_hist_df is None:
            wide_hist_df = sub
        else:
            wide_hist_df = pd.merge(wide_hist_df, sub, on="Date", how="outer")

    if wide_hist_df is not None and not wide_hist_df.empty:
        wide_hist_df = wide_hist_df.sort_values(by="Date", ascending=True).reset_index(drop=True)
    else:
        wide_hist_df = pd.DataFrame()
        
    elapsed = time.time() - start_time
    logger.info(f"52-week history extraction completed in {elapsed:.2f}s for {len(active_companies)} companies ({len(wide_hist_df)} trading days).")
    return wide_hist_df, active_companies


# ---------------------------------------------------------------------------
# 5-Year Fundamentals History (Revenue, Profit, EPS, P/E, P/S, P/B, Margins, D/E)
# ---------------------------------------------------------------------------
# Candidate row labels for each fundamentals line item, since yfinance's
# statement row naming has shifted slightly across versions.
FUNDAMENTAL_ROW_ALIASES = {
    "revenue": ["Total Revenue", "TotalRevenue"],
    "gross_profit": ["Gross Profit", "GrossProfit"],
    "operating_income": ["Operating Income", "OperatingIncome"],
    "net_income": ["Net Income", "Net Income Common Stockholders", "NetIncome"],
    "diluted_eps": ["Diluted EPS", "DilutedEPS"],
    "basic_eps": ["Basic EPS", "BasicEPS"],
    "diluted_shares": ["Diluted Average Shares", "DilutedAverageShares"],
    "basic_shares": ["Basic Average Shares", "BasicAverageShares"],
    "equity": ["Stockholders Equity", "Total Equity Gross Minority Interest", "Common Stock Equity"],
    "total_debt": ["Total Debt", "TotalDebt"],
}


def find_statement_row(df: Optional[pd.DataFrame], aliases: List[str]) -> Optional[pd.Series]:
    """Look up a financial statement line item by trying several possible label spellings."""
    if df is None or df.empty:
        return None
    index_lower = {str(i).strip().lower(): i for i in df.index}
    for alias in aliases:
        key = alias.strip().lower()
        if key in index_lower:
            return df.loc[index_lower[key]]
    return None


def get_price_on_or_before(price_history: pd.DataFrame, target_date) -> Optional[float]:
    """From a Date/Close history DataFrame, return the close on the closest trading day at/before target_date."""
    if price_history is None or price_history.empty:
        return None
    eligible = price_history[price_history["Date"] <= target_date]
    if eligible.empty:
        return None
    return safe_float(eligible.iloc[-1]["Close"])


def fetch_long_price_history(symbol: str, range_str: str = "5y", max_retries: int = 2) -> pd.DataFrame:
    """Fetch multi-year daily closing prices via the Yahoo Chart API (ascending Date/Close columns)."""
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range={range_str}&interval=1d"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=12)
            if resp.status_code == 200:
                results = resp.json().get("chart", {}).get("result")
                if results:
                    chart = results[0]
                    timestamps = chart.get("timestamp", [])
                    closes = chart.get("indicators", {}).get("quote", [{}])[0].get("close", [])
                    if timestamps and closes:
                        dates = [datetime.datetime.fromtimestamp(ts).date() for ts in timestamps]
                        hist = pd.DataFrame({"Date": dates, "Close": closes}).dropna(subset=["Close"])
                        return hist.sort_values("Date").reset_index(drop=True)
            time.sleep(0.3 * attempt)
        except Exception:
            time.sleep(0.4 * attempt)
    return pd.DataFrame()


def fetch_financial_history(symbol: str, code: str, name: str, max_years: int = 5) -> pd.DataFrame:
    """
    Build a per-fiscal-year fundamentals table (chronological, oldest year first): Revenue,
    Net Income, EPS, Dividend Per Share, P/E, P/S, P/B, Gross Margin, Operating Margin,
    and Debt/Equity.

    Historical P/E, P/S, and P/B are approximations: each fiscal year's reported diluted
    EPS / average diluted shares / stockholders' equity is combined with the closing share
    price on (or immediately before) that fiscal year-end date. Yahoo Finance's free annual
    statements go back up to 5 years for most tickers, but occasionally fewer - whatever is
    actually available (capped at max_years) is returned.

    Dividend Per Share is also an approximation, same spirit as P/E/P/S/P/B above: yfinance
    doesn't expose a per-fiscal-year dividend figure directly (dividends aren't a line item
    on the income statement or balance sheet), so each fiscal year's total is reconstructed
    by summing `Ticker.dividends` (a real per-payment historical series, ex-dividend date ->
    per-share amount) over that fiscal year's ~12-month window.
    """
    columns = [
        "FiscalYearEnd", "Revenue", "NetIncome", "EPS", "DividendPerShare", "P/E", "P/S", "P/B",
        "GrossMargin", "OperatingMargin", "DebtToEquity", "Currency",
    ]
    empty = pd.DataFrame(columns=columns)

    try:
        ticker = yf.Ticker(symbol)
        income = ticker.income_stmt
        balance = ticker.balance_sheet
        if income is None or income.empty:
            return empty

        currency = getattr(ticker.fast_info, "currency", None) or "N/A"

        revenue_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["revenue"])
        gross_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["gross_profit"])
        op_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["operating_income"])
        ni_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["net_income"])
        eps_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["diluted_eps"])
        if eps_row is None:
            eps_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["basic_eps"])
        shares_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["diluted_shares"])
        if shares_row is None:
            shares_row = find_statement_row(income, FUNDAMENTAL_ROW_ALIASES["basic_shares"])
        equity_row = find_statement_row(balance, FUNDAMENTAL_ROW_ALIASES["equity"])
        debt_row = find_statement_row(balance, FUNDAMENTAL_ROW_ALIASES["total_debt"])

        if revenue_row is None and ni_row is None:
            return empty

        # Fiscal year-end period columns only (some yfinance versions add a "TTM" column).
        # Take the most recent `max_years` periods, then reverse to chronological order
        # (oldest first, latest last) for display.
        periods = sorted([c for c in income.columns if hasattr(c, "year")], reverse=True)[:max_years]
        if not periods:
            return empty
        periods = list(reversed(periods))

        price_hist = fetch_long_price_history(symbol, range_str="5y")

        try:
            dividends = ticker.dividends
            if dividends is not None and not dividends.empty:
                dividends = dividends.copy()
                dividends.index = pd.to_datetime(dividends.index).tz_localize(None)
        except Exception:
            dividends = None

        rows = []
        prev_fye = None
        for period in periods:
            fye = period.date() if hasattr(period, "date") else period
            revenue = safe_float(revenue_row.get(period)) if revenue_row is not None else None
            gross_profit = safe_float(gross_row.get(period)) if gross_row is not None else None
            op_income = safe_float(op_row.get(period)) if op_row is not None else None
            net_income = safe_float(ni_row.get(period)) if ni_row is not None else None
            eps = safe_float(eps_row.get(period)) if eps_row is not None else None
            shares = safe_float(shares_row.get(period)) if shares_row is not None else None
            equity = safe_float(equity_row.get(period)) if equity_row is not None else None
            debt = safe_float(debt_row.get(period)) if debt_row is not None else None

            price = get_price_on_or_before(price_hist, fye)

            pe = (price / eps) if (price is not None and eps not in (None, 0)) else None
            market_cap_fye = (price * shares) if (price is not None and shares) else None
            ps = (market_cap_fye / revenue) if (market_cap_fye is not None and revenue not in (None, 0)) else None
            pb = (market_cap_fye / equity) if (market_cap_fye is not None and equity not in (None, 0)) else None
            gross_margin = (gross_profit / revenue * 100.0) if (gross_profit is not None and revenue not in (None, 0)) else None
            op_margin = (op_income / revenue * 100.0) if (op_income is not None and revenue not in (None, 0)) else None
            debt_to_equity = (debt / equity) if (debt is not None and equity not in (None, 0)) else None

            # Dividend Per Share for this fiscal year: sum of per-share dividend payments
            # (ex-dividend date -> amount, from yfinance) falling within this fiscal year's
            # ~12-month window - since the previous period's fiscal year-end if known,
            # otherwise approximated as 1 year before this period's fiscal year-end.
            dps = None
            if dividends is not None and not dividends.empty:
                window_end = pd.Timestamp(fye)
                window_start = prev_fye if prev_fye is not None else (window_end - pd.DateOffset(years=1))
                in_window = dividends[(dividends.index > window_start) & (dividends.index <= window_end)]
                if not in_window.empty:
                    dps = float(in_window.sum())
            prev_fye = pd.Timestamp(fye)

            rows.append({
                "FiscalYearEnd": fye.strftime("%Y-%m-%d") if hasattr(fye, "strftime") else str(fye),
                "Revenue": revenue,
                "NetIncome": net_income,
                "EPS": round(eps, 2) if eps is not None else None,
                "DividendPerShare": round(dps, 2) if dps is not None else None,
                "P/E": round(pe, 2) if pe is not None else None,
                "P/S": round(ps, 2) if ps is not None else None,
                "P/B": round(pb, 2) if pb is not None else None,
                "GrossMargin": round(gross_margin, 2) if gross_margin is not None else None,
                "OperatingMargin": round(op_margin, 2) if op_margin is not None else None,
                "DebtToEquity": round(debt_to_equity, 2) if debt_to_equity is not None else None,
                "Currency": currency,
            })

        return pd.DataFrame(rows, columns=columns)
    except Exception as e:
        logger.warning(f"Could not build financial history for {symbol}: {e}")
        return empty


def extract_financial_history_all(summary_df: pd.DataFrame, max_workers: int = 6, max_years: int = 5) -> Dict[str, pd.DataFrame]:
    """
    Fetch up to 5 years of fundamentals history (Revenue, Profit, EPS, P/E, P/S, P/B, margins,
    D/E) for every company in summary_df. Uses a lower default thread count since income
    statement / balance sheet lookups are heavier calls than the basic snapshot fetch.
    Returns a dict mapping Symbol -> fundamentals DataFrame (only for symbols with data).
    """
    logger.info(f"Starting fundamentals history extraction for {len(summary_df)} companies...")
    start_time = time.time()

    tasks = [
        {"symbol": row["Symbol"], "code": row["Code"], "name": row["Name"]}
        for _, row in summary_df.iterrows()
    ]
    results: Dict[str, pd.DataFrame] = {}
    completed = 0
    total = len(tasks)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {
            executor.submit(fetch_financial_history, t["symbol"], t["code"], t["name"], max_years): t
            for t in tasks
        }
        for future in as_completed(future_to_task):
            t = future_to_task[future]
            try:
                fin_df = future.result()
                if fin_df is not None and not fin_df.empty:
                    results[t["symbol"]] = fin_df
            except Exception as e:
                logger.error(f"Error fetching fundamentals for {t['symbol']}: {e}")
            completed += 1
            if completed % 10 == 0 or completed == total:
                logger.info(f"Fundamentals progress: {completed}/{total} ({completed / total * 100:.1f}%)")

    elapsed = time.time() - start_time
    logger.info(f"Fundamentals history extraction completed in {elapsed:.2f}s for {len(results)}/{total} companies.")
    return results


def save_results(
    df_summary: pd.DataFrame,
    df_history: pd.DataFrame,
    active_companies: list,
    output_path: str,
    financial_histories: Optional[Dict[str, pd.DataFrame]] = None,
):
    """
    Save extracted dataset. If output is an Excel file (.xlsx), both summary and
    historical daily data are saved as separate worksheets in the same workbook.
    Formats numbers across both worksheets to 2 decimal places and abbreviates large
    numbers (Market Cap, Shares Outstanding, Volume) with B / M / T suffixes.

    If financial_histories is provided (Symbol -> fundamentals DataFrame), a "Financials
    Detail" worksheet is added with a stacked 5-year fundamentals block per company, and
    each company's Name cell on the Summary sheet becomes a hyperlink jumping straight to
    its block (kept as ONE extra visible sheet rather than one hidden sheet per company,
    since Excel refuses to navigate hyperlinks into a truly hidden worksheet).
    """
    ext = os.path.splitext(output_path)[1].lower()
    
    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
        
    if ext in [".xlsx", ".xls"]:
        try:
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            df_summary_fmt = df_summary.copy()

            # SharesOutstanding stays out of the Summary sheet specifically (still used
            # internally to compute the 52W History sheet's Market Cap column, via the
            # untouched df_summary passed into extract_52w_history_all before this point -
            # dropping it only from this formatted copy doesn't affect that).
            df_summary_fmt = df_summary_fmt.drop(columns=["SharesOutstanding"], errors="ignore")

            if "Volume" in df_summary_fmt.columns:
                df_summary_fmt["Volume"] = df_summary_fmt["Volume"].apply(
                    lambda x: format_abbreviated(x) if pd.notnull(x) else "N/A"
                )
            for col in ["TodaysClose", "52WHigh", "52WLow", "52WRange", "Trailing_PE", "Forward_PE"]:
                if col in df_summary_fmt.columns:
                    df_summary_fmt[col] = df_summary_fmt[col].apply(
                        lambda x: f"{float(x):.2f}" if pd.notnull(x) and safe_float(x) is not None else x
                    )

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                # 1. Write formatted Summary Sheet
                df_summary_fmt.to_excel(writer, sheet_name="Top 80 & NIFTY 50 Summary", index=False)
                
                wb = writer.book
                ws_summary = wb["Top 80 & NIFTY 50 Summary"]
                sum_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                sum_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                center_align = Alignment(horizontal="center", vertical="center")
                
                for col_num in range(1, len(df_summary_fmt.columns) + 1):
                    cell = ws_summary.cell(row=1, column=col_num)
                    cell.font = sum_header_font
                    cell.fill = sum_header_fill
                    cell.alignment = center_align
                    col_letter = get_column_letter(col_num)
                    ws_summary.column_dimensions[col_letter].width = max(12, len(str(cell.value or "")) + 3)

                # 2. Write formatted 52W Daily History Sheet
                history_ws = None
                history_col_anchor: Dict[str, int] = {}  # Symbol -> its start column on this sheet
                if df_history is not None and not df_history.empty:
                    ws = wb.create_sheet("52W Daily History")
                    history_ws = ws
                    
                    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    sub_header_font = Font(name="Calibri", size=10, bold=True, color="1F2937")
                    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                    sub_header_fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
                    
                    thin_border = Border(
                        left=Side(style="thin", color="D1D5DB"),
                        right=Side(style="thin", color="D1D5DB"),
                        top=Side(style="thin", color="D1D5DB"),
                        bottom=Side(style="thin", color="D1D5DB")
                    )
                    
                    metrics = ["Close", "% Change", "Market Cap", "Shares Out.", "Volume"]
                    
                    # Column 1: Date (Merged vertically across Row 1 & 2)
                    ws["A1"] = "Date"
                    ws.merge_cells("A1:A2")
                    ws["A1"].font = header_font
                    ws["A1"].fill = header_fill
                    ws["A1"].alignment = center_align
                    ws["A1"].border = thin_border
                    ws["A2"].border = thin_border
                    
                    col_idx = 2
                    for sym, code, name in active_companies:
                        start_col = col_idx
                        end_col = col_idx + len(metrics) - 1
                        history_col_anchor[sym] = start_col

                        start_letter = get_column_letter(start_col)
                        end_letter = get_column_letter(end_col)
                        
                        # Row 1: Stock Name and Code (Merged across 5 columns)
                        ws.cell(row=1, column=start_col, value=f"{code} - {name}")
                        ws.merge_cells(f"{start_letter}1:{end_letter}1")
                        
                        for c in range(start_col, end_col + 1):
                            cell = ws.cell(row=1, column=c)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = center_align
                            cell.border = thin_border
                            
                        # Row 2: Metric names
                        for i, m in enumerate(metrics):
                            sub_cell = ws.cell(row=2, column=start_col + i, value=m)
                            sub_cell.font = sub_header_font
                            sub_cell.fill = sub_header_fill
                            sub_cell.alignment = center_align
                            sub_cell.border = thin_border
                            
                        col_idx += len(metrics)
                        
                    # Write Historical Daily Data starting at Row 3 with formatting applied
                    date_vals = df_history["Date"].tolist()
                    for r_idx, d_val in enumerate(date_vals):
                        ws.cell(row=3 + r_idx, column=1, value=d_val)
                        
                    data_cols = [c for c in df_history.columns if c != "Date"]
                    for c_idx, col_name in enumerate(data_cols, start=2):
                        col_values = df_history[col_name].tolist()
                        
                        is_pct = col_name.endswith("% Change")
                        is_close = col_name.endswith("Close")
                        is_abbrev = any(col_name.endswith(suffix) for suffix in ["Market Cap", "Shares Out.", "Volume"])
                        
                        for r_idx, val in enumerate(col_values):
                            if pd.notnull(val):
                                f_val = safe_float(val)
                                if f_val is not None and is_close:
                                    # Keep Close as a REAL number (not a display string) so the
                                    # "Price Charts" sheet can reference it directly instead of
                                    # duplicating the data - only the number_format changes how
                                    # it looks (still "1234.56"), the underlying value is numeric.
                                    cell = ws.cell(row=3 + r_idx, column=c_idx, value=f_val)
                                    cell.number_format = "0.00"
                                    continue
                                if f_val is not None:
                                    if is_pct:
                                        formatted_cell = f"{f_val:+.2f}%"
                                    elif is_abbrev:
                                        formatted_cell = format_abbreviated(f_val)
                                    else:
                                        formatted_cell = f"{f_val:.2f}"
                                else:
                                    formatted_cell = str(val)
                                ws.cell(row=3 + r_idx, column=c_idx, value=formatted_cell)

                    ws.column_dimensions["A"].width = 13
                    for col in range(2, col_idx):
                        col_letter = get_column_letter(col)
                        ws.column_dimensions[col_letter].width = 13

                # 3. Write "Financials Detail" sheet: one stacked 5-year fundamentals block
                #    per company. The company Name cell on the Summary sheet hyperlinks
                #    straight to its block here, instead of scattering one hidden sheet
                #    per company (which Excel can't hyperlink-navigate to anyway).
                if financial_histories:
                    ws_fin = wb.create_sheet("Financials Detail")
                    fin_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    fin_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                    fin_sub_font = Font(name="Calibri", size=10, bold=True, color="1F2937")
                    fin_sub_fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
                    back_link_font = Font(name="Calibri", size=9, italic=True, color="6B7280")
                    hyperlink_font = Font(name="Calibri", size=11, color="0563C1", underline="single")

                    fin_metric_cols = [
                        "Fiscal Year End", "Revenue", "Net Income", "EPS", "Dividend/Unit", "P/E", "P/S", "P/B",
                        "Gross Margin", "Operating Margin", "Debt/Equity",
                    ]
                    n_fin_cols = len(fin_metric_cols)

                    anchor_rows: Dict[str, int] = {}  # Symbol -> row of its block header
                    row_cursor = 1
                    ordered_symbols = [s for s in df_summary["Symbol"].tolist() if s in financial_histories]

                    for sym in ordered_symbols:
                        fin_df = financial_histories[sym]
                        match = df_summary.loc[df_summary["Symbol"] == sym]
                        code = match["Code"].iloc[0] if not match.empty else sym
                        name = match["Name"].iloc[0] if not match.empty else sym
                        currency = fin_df["Currency"].iloc[0] if "Currency" in fin_df.columns and not fin_df.empty else "N/A"

                        anchor_rows[sym] = row_cursor

                        header_cell = ws_fin.cell(row=row_cursor, column=1, value=f"{code} - {name}  (Reported in {currency})")
                        ws_fin.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=n_fin_cols)
                        header_cell.font = fin_header_font
                        header_cell.fill = fin_header_fill
                        header_cell.alignment = Alignment(horizontal="left", vertical="center")
                        for c in range(1, n_fin_cols + 1):
                            ws_fin.cell(row=row_cursor, column=c).fill = fin_header_fill
                        row_cursor += 1

                        for c_idx, col_name in enumerate(fin_metric_cols, start=1):
                            sub_cell = ws_fin.cell(row=row_cursor, column=c_idx, value=col_name)
                            sub_cell.font = fin_sub_font
                            sub_cell.fill = fin_sub_fill
                            sub_cell.alignment = center_align
                        row_cursor += 1

                        for _, fin_row in fin_df.iterrows():
                            values = [
                                fin_row.get("FiscalYearEnd"),
                                format_currency(fin_row.get("Revenue"), currency) if pd.notnull(fin_row.get("Revenue")) else "N/A",
                                format_currency(fin_row.get("NetIncome"), currency) if pd.notnull(fin_row.get("NetIncome")) else "N/A",
                                f"{fin_row['EPS']:.2f}" if pd.notnull(fin_row.get("EPS")) else "N/A",
                                f"{fin_row['DividendPerShare']:.2f}" if pd.notnull(fin_row.get("DividendPerShare")) else "N/A",
                                f"{fin_row['P/E']:.2f}" if pd.notnull(fin_row.get("P/E")) else "N/A",
                                f"{fin_row['P/S']:.2f}" if pd.notnull(fin_row.get("P/S")) else "N/A",
                                f"{fin_row['P/B']:.2f}" if pd.notnull(fin_row.get("P/B")) else "N/A",
                                f"{fin_row['GrossMargin']:.2f}%" if pd.notnull(fin_row.get("GrossMargin")) else "N/A",
                                f"{fin_row['OperatingMargin']:.2f}%" if pd.notnull(fin_row.get("OperatingMargin")) else "N/A",
                                f"{fin_row['DebtToEquity']:.2f}" if pd.notnull(fin_row.get("DebtToEquity")) else "N/A",
                            ]
                            for c_idx, val in enumerate(values, start=1):
                                ws_fin.cell(row=row_cursor, column=c_idx, value=val)
                            row_cursor += 1

                        back_cell = ws_fin.cell(row=row_cursor, column=1, value="↑ Back to Summary")
                        back_cell.hyperlink = "#'Top 80 & NIFTY 50 Summary'!A1"
                        back_cell.font = back_link_font
                        row_cursor += 2  # spacer row before next block

                    ws_fin.column_dimensions["A"].width = 30
                    for col in range(2, n_fin_cols + 1):
                        ws_fin.column_dimensions[get_column_letter(col)].width = 14

                    # 4. Hyperlink each company's Name cell on the Summary sheet to its block above
                    if "Name" in df_summary.columns and "Symbol" in df_summary.columns:
                        name_col_idx = list(df_summary.columns).index("Name") + 1
                        for r_idx, sym in enumerate(df_summary["Symbol"].tolist(), start=2):  # row 1 = header
                            if sym in anchor_rows:
                                cell = ws_summary.cell(row=r_idx, column=name_col_idx)
                                cell.hyperlink = f"#'Financials Detail'!A{anchor_rows[sym]}"
                                cell.font = hyperlink_font

                # 5. Write "Price Charts" sheet: one compact 1-Year Close-price line chart
                #    per company, stacked vertically. Excel charts need real numeric cells
                #    (not the 2dp display-formatted strings on "52W Daily History"), so each
                #    chart is backed by a small Date/Close numeric table tucked into two
                #    HIDDEN helper columns on the same sheet (hidden COLUMNS are fine to
                #    hyperlink to/from - only hidden SHEETS break Excel's hyperlink navigation).
                if df_history is not None and not df_history.empty and active_companies and history_ws is not None:
                    from openpyxl.chart import LineChart, Reference
                    from openpyxl.chart.marker import Marker

                    ws_charts = wb.create_sheet("Price Charts")
                    chart_header_font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
                    chart_back_font = Font(name="Calibri", size=9, italic=True, color="6B7280")
                    chart_hyperlink_font = Font(name="Calibri", size=11, color="0563C1", underline="single")

                    BLOCK_SPACING = 18  # rows reserved per company block (header + chart + gap)
                    n_dates = len(df_history)
                    HIST_DATA_START_ROW = 3  # "52W Daily History": rows 1-2 are headers, data from row 3

                    chart_anchor_rows: Dict[str, int] = {}  # Symbol -> row of its block header
                    row_cursor = 1

                    for sym, code, name in active_companies:
                        if sym not in history_col_anchor:
                            continue
                        # "Close" is always the first of the 5 metric sub-columns per company
                        # on "52W Daily History" - reference that sheet's own numeric cells
                        # directly rather than duplicating the data onto this sheet, so there's
                        # nothing here that can go stale or collide between companies.
                        close_col_idx = history_col_anchor[sym]

                        chart_anchor_rows[sym] = row_cursor

                        header_cell = ws_charts.cell(row=row_cursor, column=1, value=f"{code} - {name}")
                        header_cell.font = chart_header_font
                        back_cell = ws_charts.cell(row=row_cursor + 1, column=1, value="↑ Back to Summary")
                        back_cell.hyperlink = "#'Top 80 & NIFTY 50 Summary'!A1"
                        back_cell.font = chart_back_font

                        currency_match = df_summary.loc[df_summary["Symbol"] == sym, "Currency"]
                        currency = currency_match.iloc[0] if not currency_match.empty else ""

                        chart = LineChart()
                        chart.title = f"{code} — 1Y Close Price"
                        chart.style = 2
                        chart.height = 7
                        chart.width = 15
                        chart.legend = None

                        # Visible, labeled axes: dates along X (thinned to ~monthly tick
                        # labels so 252 trading days don't cram into unreadable text),
                        # close price along Y.
                        chart.x_axis.delete = False
                        chart.x_axis.title = "Date"
                        chart.x_axis.tickLblSkip = 21
                        chart.x_axis.tickMarkSkip = 21
                        chart.y_axis.delete = False
                        chart.y_axis.title = f"Close Price ({currency})" if currency else "Close Price"

                        # Auto-scale the Y axis to this stock's actual 1Y trading range (with a
                        # small margin) instead of Excel's default of anchoring the axis at 0.
                        # Most stocks trade well above 0, so a 0-anchored axis crushes a year's
                        # worth of real fluctuation into a thin band near the top of the chart.
                        close_values = [
                            history_ws.cell(row=r, column=close_col_idx).value
                            for r in range(HIST_DATA_START_ROW, HIST_DATA_START_ROW + n_dates)
                        ]
                        close_values = [v for v in close_values if isinstance(v, (int, float))]
                        if close_values:
                            data_min, data_max = min(close_values), max(close_values)
                            value_range = data_max - data_min
                            padding = value_range * 0.05 if value_range > 0 else max(data_max * 0.05, 0.5)
                            chart.y_axis.scaling.min = max(0.0, round(data_min - padding, 2))
                            chart.y_axis.scaling.max = round(data_max + padding, 2)

                        cats = Reference(history_ws, min_col=1, min_row=HIST_DATA_START_ROW, max_row=HIST_DATA_START_ROW + n_dates - 1)
                        values_ref = Reference(history_ws, min_col=close_col_idx, min_row=HIST_DATA_START_ROW, max_row=HIST_DATA_START_ROW + n_dates - 1)
                        chart.add_data(values_ref, titles_from_data=False)
                        chart.set_categories(cats)
                        for s in chart.series:
                            s.marker = Marker(symbol="none")
                            s.smooth = False
                            s.graphicalProperties.line.width = 20000  # ~1.6pt
                            s.graphicalProperties.line.solidFill = "1E3A8A"

                        ws_charts.add_chart(chart, f"A{row_cursor + 2}")
                        row_cursor += BLOCK_SPACING

                    ws_charts.column_dimensions["A"].width = 40

                    # 6. Hyperlink Symbol cells on Summary, and company headers on "52W Daily
                    #    History", straight to each company's chart above.
                    if "Symbol" in df_summary.columns:
                        symbol_col_idx = list(df_summary.columns).index("Symbol") + 1
                        for r_idx, sym in enumerate(df_summary["Symbol"].tolist(), start=2):
                            if sym in chart_anchor_rows:
                                cell = ws_summary.cell(row=r_idx, column=symbol_col_idx)
                                cell.hyperlink = f"#'Price Charts'!A{chart_anchor_rows[sym]}"
                                cell.font = chart_hyperlink_font

                    if history_ws is not None:
                        for sym, start_col in history_col_anchor.items():
                            if sym in chart_anchor_rows:
                                history_ws.cell(row=1, column=start_col).hyperlink = f"#'Price Charts'!A{chart_anchor_rows[sym]}"

            logger.info(f"Multi-sheet Excel workbook with formatted cells saved to: {os.path.abspath(output_path)}")
        except Exception as e:
            logger.error(f"Failed to write Excel workbook: {e}")
    elif ext == ".json":
        base, _ = os.path.splitext(output_path)
        df_summary.to_json(output_path, orient="records", indent=2)
        if df_history is not None and not df_history.empty:
            hist_path = f"{base}_history.json"
            df_history.to_json(hist_path, orient="records", indent=2)
            logger.info(f"History dataset saved to: {os.path.abspath(hist_path)}")
        logger.info(f"Summary dataset successfully saved to: {os.path.abspath(output_path)}")
    else:
        base, _ = os.path.splitext(output_path)
        df_summary.to_csv(output_path, index=False)
        if df_history is not None and not df_history.empty:
            hist_path = f"{base}_history.csv"
            df_history.to_csv(hist_path, index=False)
            logger.info(f"History dataset saved to: {os.path.abspath(hist_path)}")
        logger.info(f"Summary dataset successfully saved to: {os.path.abspath(output_path)}")


def main():
    parser = argparse.ArgumentParser(
        description="Extract Top 80 NSE companies by latest market cap, 52W range metrics, NIFTY 50 constituents, and 52W historical daily data."
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=10,
        help="Number of concurrent worker threads (default: 10)."
    )
    parser.add_argument(
        "--top",
        type=int,
        default=80,
        help="Base top N companies by market cap to extract (default: 80)."
    )
    parser.add_argument(
        "--mode",
        type=str,
        choices=["fast", "robust", "full"],
        default="robust",
        help="Scanning mode: 'fast' (NIFTY 100 ~100 stocks), 'robust' (NIFTY 200 ~200 stocks, default), 'full' (Entire NSE ~2,550+ stocks)."
    )
    parser.add_argument(
        "--stock-list-file",
        type=str,
        default="stock list.xlsx",
        help="Path to the shared stock list Excel file, 'NSE' worksheet, for personal stock additions (default: 'stock list.xlsx')."
    )
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    default_filename = f"NSE_Stock_Analysis_{today_str}.xlsx"

    parser.add_argument(
        "--output",
        type=str,
        default=default_filename,
        help=f"Output filepath (.xlsx, .csv, or .json, default: {default_filename})."
    )
    parser.add_argument(
        "--no-financials",
        action="store_true",
        dest="skip_financials",
        help="Skip the 5-year fundamentals extraction (Revenue, Profit, EPS, P/E, P/S, P/B, margins, D/E)."
    )
    parser.add_argument(
        "--financials-years",
        type=int,
        default=5,
        help="Number of trailing fiscal years of fundamentals to extract (default: 5; fewer if Yahoo doesn't have that much history)."
    )
    parser.add_argument(
        "--financials-workers",
        type=int,
        default=6,
        help="Concurrent worker threads for fundamentals extraction (default: 6, kept low since these are heavier calls)."
    )

    args = parser.parse_args()

    # 1. Extract Top 80 & NIFTY 50 Summary
    df_summary = extract_top80_with_nifty50(
        max_workers=args.workers,
        top_n=args.top,
        mode=args.mode,
        stock_list_file=args.stock_list_file
    )

    print_top80_nifty_table(df_summary)

    # 2. Extract 52-Week Historical Daily Data for all extracted companies
    df_history, active_companies = extract_52w_history_all(
        summary_df=df_summary,
        max_workers=args.workers
    )

    # 3. Extract 5-Year Fundamentals History (Revenue, Profit, EPS, P/E, P/S, P/B, margins, D/E)
    financial_histories = None
    if not args.skip_financials:
        financial_histories = extract_financial_history_all(
            summary_df=df_summary,
            max_workers=args.financials_workers,
            max_years=args.financials_years
        )

    # 4. Save all tables to output (multi-worksheet workbook if Excel)
    save_results(df_summary, df_history, active_companies, args.output, financial_histories=financial_histories)


if __name__ == "__main__":
    main()
