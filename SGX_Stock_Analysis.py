"""
================================================================================
SGX Stock Analysis & Index Benchmark Extractor (STI, S Cap, L & M Cap)
================================================================================
Collects the latest snapshot of market capitalization, 52-week pricing metrics,
trailing P/E, sector, industry, and custom 52-week range positioning for:
1. SGX Index Constituents (Default Mode):
   - Straits Times Index (STI) [30 constituents]
   - Small Cap Index (S Cap) [40 constituents]
   - Large & Mid Cap Index (L & M Cap) [19 constituents]
   Hardcoded in SGX_INDEX_STOCK_ROWS, generated from 'stock list.xlsx' ('SGX' worksheet)
   by sync_stock_lists.py - the file itself isn't read at runtime.
2. Market Cap Scanning Mode (--scan-market-cap / --mode market-cap):
   - Full SGX universe extraction, Top N ranking, and STI catch-up inclusion.

Key Features:
- Latest Snapshot Metrics:
  1. 52-Week High
  2. 52-Week Low
  3. Trailing P/E
  4. Sector
  5. Industry
  6. Latest Close Price
  7. 52-Week Range (52w High - 52w Low)
  8. Custom Range Metric (%): (Latest Close - 52w Low) / 52w Range (%) [Nearness to 52W High]
- Inclusion Reason: Reflects the index the stock is from ('STI', 'S Cap', 'L & M Cap')
  or Top X / STI catch-up when market cap scanning is used.
- Multi-Threaded Concurrent Extraction with automatic retry pass.
- Multi-Format Export: Excel (.xlsx with 52W history sheet), CSV, and JSON.
================================================================================
"""

import os
import re
import sys
import time
import datetime
import logging
import argparse
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
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("SGX_Stock_Analysis")

# Official fallback list of current 30 STI constituents
FALLBACK_STI_30 = [
    ("D05", "DBS Group Holdings Ltd", "Financial Services"),
    ("O39", "Oversea-Chinese Banking Corp Ltd", "Financial Services"),
    ("U11", "United Overseas Bank Ltd", "Financial Services"),
    ("Z74", "Singapore Telecommunications Ltd", "Communication Services"),
    ("S63", "ST Engineering Ltd", "Industrials"),
    ("S68", "Singapore Exchange Ltd", "Financial Services"),
    ("C6L", "Singapore Airlines Ltd", "Industrials"),
    ("BN4", "Keppel Ltd", "Industrials"),
    ("C38U", "CapitaLand Integrated Commercial Trust", "Real Estate"),
    ("A17U", "CapitaLand Ascendas REIT", "Real Estate"),
    ("9CI", "CapitaLand Investment Ltd", "Real Estate"),
    ("U96", "Sembcorp Industries Ltd", "Utilities"),
    ("F34", "Wilmar International Ltd", "Consumer Defensive"),
    ("BS6", "Yangzijiang Shipbuilding Holdings Ltd", "Industrials"),
    ("Y92", "Thai Beverage PCL", "Consumer Defensive"),
    ("J36", "Jardine Matheson Holdings Ltd", "Industrials"),
    ("H78", "Hongkong Land Holdings Ltd", "Real Estate"),
    ("G13", "Genting Singapore Ltd", "Consumer Cyclical"),
    ("C09", "City Developments Ltd", "Real Estate"),
    ("5E2", "Seatrium Limited", "Energy"),
    ("N2IU", "Mapletree Pan Asia Commercial Trust", "Real Estate"),
    ("M44U", "Mapletree Logistics Trust", "Real Estate"),
    ("ME8U", "Mapletree Industrial Trust", "Real Estate"),
    ("AJBU", "Keppel DC REIT", "Real Estate"),
    ("BUOU", "Frasers Logistics & Commercial Trust", "Real Estate"),
    ("J69U", "Frasers Centrepoint Trust", "Real Estate"),
    ("D01", "DFI Retail Group Holdings Ltd", "Consumer Defensive"),
    ("S58", "SATS Ltd", "Industrials"),
    ("U14", "UOL Group Ltd", "Real Estate"),
    ("V03", "Venture Corporation Ltd", "Technology"),
]

# SGX default-mode stock universe: every STI / S Cap / L & M Cap constituent plus any
# "Personal" additions, as (Index, Company, Ticker) rows. The block between the markers
# below is GENERATED from the 'SGX' worksheet of 'stock list.xlsx' by sync_stock_lists.py
# (run automatically by update_notebook.py) - hardcoded rather than read from the file at
# runtime, so the script (and the Colab notebook built from it) never needs the file.
# To add/remove a stock: edit the spreadsheet, then re-run the sync - don't edit by hand.
# >>> BEGIN GENERATED STOCK LIST (sync_stock_lists.py) - edit stock list.xlsx, not this block
SGX_INDEX_STOCK_ROWS = [
    ("STI", "Mapletree PanAsia Co", "N2IU.SI"),
    ("STI", "CapLand Ascendas REI", "A17U.SI"),
    ("STI", "Mapletree Log Tr", "M44U.SI"),
    ("STI", "JMH USD", "J36.SI"),
    ("STI", "Frasers Cpt Tr", "J69U.SI"),
    ("STI", "Mapletree Ind Tr", "ME8U.SI"),
    ("STI", "Keppel DC Reit", "AJBU.SI"),
    ("STI", "Genting Sing", "G13.SI"),
    ("STI", "Frasers L&C Tr", "BUOU.SI"),
    ("STI", "CapitaLandInvest", "9CI.SI"),
    ("STI", "DFIRG USD", "D01.SI"),
    ("STI", "SIA", "C6L.SI"),
    ("STI", "CapLand IntCom T", "C38U.SI"),
    ("STI", "Singtel", "Z74.SI"),
    ("STI", "SATS", "S58.SI"),
    ("STI", "Seatrium Ltd", "5E2.SI"),
    ("STI", "UOL", "U14.SI"),
    ("STI", "ThaiBev", "Y92.SI"),
    ("STI", "Sembcorp Ind", "U96.SI"),
    ("STI", "CityDev", "C09.SI"),
    ("STI", "Venture Corp", "V03.SI"),
    ("STI", "Keppel", "BN4.SI"),
    ("STI", "ST Engineering", "S63.SI"),
    ("STI", "UOB", "U11.SI"),
    ("STI", "HongkongLand USD", "H78.SI"),
    ("STI", "Wilmar Intl", "F34.SI"),
    ("STI", "SGX", "S68.SI"),
    ("STI", "OCBC Bank", "O39.SI"),
    ("STI", "DBS Grp", "D05.SI"),
    ("STI", "YZJ Shipbldg SGD", "BS6.SI"),
    ("S Cap", "First Reit", "AW9U.SI"),
    ("S Cap", "YZJ Fin Hldg", "YF8.SI"),
    ("S Cap", "Aztech Gbl", "8AZ.SI"),
    ("S Cap", "Prime US REIT USD", "OXMU.SI"),
    ("S Cap", "Yanlord Land", "Z25.SI"),
    ("S Cap", "Delfi", "P34.SI"),
    ("S Cap", "CapLand India T", "CY6U.SI"),
    ("S Cap", "CapLand China T", "AU8U.SI"),
    ("S Cap", "UtdHampshReitUSD", "ODBU.SI"),
    ("S Cap", "Samudera Shipping", "S56.SI"),
    ("S Cap", "Manulife Reit USD", "BTOU.SI"),
    ("S Cap", "Far East HTrust", "Q5T.SI"),
    ("S Cap", "StarhillGbl Reit", "P40U.SI"),
    ("S Cap", "Daiwa Hse Log Tr", "DHLU.SI"),
    ("S Cap", "Wee Hur", "E3B.SI"),
    ("S Cap", "KORE Reit USD", "CMOU.SI"),
    ("S Cap", "Cosco Shp", "F83.SI"),
    ("S Cap", "Straits Trading", "S20.SI"),
    ("S Cap", "PropNex", "OYY.SI"),
    ("S Cap", "SingPost", "S08.SI"),
    ("S Cap", "China Aviation", "G92.SI"),
    ("S Cap", "Lendlease Reit", "JYEU.SI"),
    ("S Cap", "Yoma Strategic", "Z59.SI"),
    ("S Cap", "DigiCore Reit USD", "DCRU.SI"),
    ("S Cap", "AIMS APAC Reit", "O5RU.SI"),
    ("S Cap", "Nanofilm", "MZH.SI"),
    ("S Cap", "Stoneweg EU Trust EU", "SET.SI"),
    ("S Cap", "BRC Asia", "BEC.SI"),
    ("S Cap", "Food Empire", "F03.SI"),
    ("S Cap", "Riverstone", "AP4.SI"),
    ("S Cap", "Frencken", "E28.SI"),
    ("S Cap", "Hong Fok", "H30.SI"),
    ("S Cap", "Valuetronics", "BN2.SI"),
    ("S Cap", "Centurion", "OU8.SI"),
    ("S Cap", "Geo Energy Res", "RE4.SI"),
    ("S Cap", "Sasseur Reit", "CRPU.SI"),
    ("S Cap", "UMS", "558.SI"),
    ("S Cap", "AEM SGD", "AWX.SI"),
    ("S Cap", "Civmec", "P9D.SI"),
    ("S Cap", "Bumitama Agri", "P8Z.SI"),
    ("L & M Cap", "CapLand Ascott T", "HMN.SI"),
    ("L & M Cap", "RaffIes Medical", "BSL.SI"),
    ("L & M Cap", "ESR REIT", "9A4U.SI"),
    ("L & M Cap", "SIA Engineering", "S59.SI"),
    ("L & M Cap", "ComfortDelGro", "C52.SI"),
    ("L & M Cap", "CDL HTrust", "J85.SI"),
    ("L & M Cap", "Keppel Reit", "K71U.SI"),
    ("L & M Cap", "IFAST", "AIY.SI"),
    ("L & M Cap", "ParkwayLife Reit", "C2PU.SI"),
    ("L & M Cap", "OUEREIT", "TS0U.SI"),
    ("L & M Cap", "Olam Group", "VC2.SI"),
    ("L & M Cap", "NetLink NBN Tr", "CJLU.SI"),
    ("L & M Cap", "HPH Trust USD", "NS8U.SI"),
    ("L & M Cap", "Suntec Reit", "T82U.SI"),
    ("L & M Cap", "StarHub", "CC3.SI"),
    ("L & M Cap", "Kep Infra Tr", "A7RU.SI"),
    ("L & M Cap", "Golden Agri-Res", "E5H.SI"),
    ("L & M Cap", "Sheng Siong", "OV8.SI"),
    ("L & M Cap", "First Resources", "EB5.SI"),
    ("Personal", "Micro Mechanics", "5DD.SI"),
    ("Personal", "Jardine C&C", "C07.SI"),
    ("Personal", "Rex International", "5WH.SI"),
]
# <<< END GENERATED STOCK LIST

# Pre-compiled default-mode stock list, derived from the rows above. "type" is left as
# None here and worked out per stock after its data is fetched (see classify_sgx_security).
SGX_INDEX_STOCKS = [
    {
        "code": ticker.split(".")[0],
        "symbol": ticker,
        "type": None,
        "name": company,
        "index": index_name,
        "is_sti30": index_name == "STI",
    }
    for index_name, company, ticker in SGX_INDEX_STOCK_ROWS
]

# SGX business trusts in the default-mode list. Unlike REITs (identifiable from Yahoo's
# industry, which always starts with "REIT"), business trusts report ordinary industries
# (e.g. "Telecom Services", "Marine Shipping"), so they're listed explicitly. When adding a
# new business trust to stock list.xlsx, add its ticker here too - classify_sgx_security
# logs a warning for any non-REIT whose name looks like a trust but isn't in this set.
SGX_BUSINESS_TRUSTS = {
    "CY6U.SI",  # CapitaLand India Trust
    "CJLU.SI",  # NetLink NBN Trust
    "NS8U.SI",  # Hutchison Port Holdings Trust
    "A7RU.SI",  # Keppel Infrastructure Trust
}


def classify_sgx_security(symbol: str, industry: Optional[str], name: Optional[str]) -> str:
    """
    Work out a default-mode stock's security type ('reits', 'businesstrusts', or 'stocks')
    from its fetched Yahoo industry plus the SGX_BUSINESS_TRUSTS list. Returns 'N/A' if the
    industry couldn't be fetched and the stock isn't a known business trust.
    """
    if symbol in SGX_BUSINESS_TRUSTS:
        return "businesstrusts"
    if not industry or industry == "N/A":
        return "N/A"
    if industry.startswith("REIT"):
        return "reits"
    name_words = (name or "").lower().replace(".", " ").split()
    if "trust" in name_words or "tr" in name_words:
        logger.warning(
            f"{symbol} ({name}) looks like a trust but isn't a REIT or in SGX_BUSINESS_TRUSTS - "
            f"classified as 'stocks'. If it's a business trust, add it to SGX_BUSINESS_TRUSTS."
        )
    return "stocks"


def clean_text(text: str) -> str:
    """Strip Wikipedia citation brackets [1], [zh], etc."""
    return re.sub(r"\[.*?\]", "", text).strip()


def format_currency(value: Optional[float], currency: str = "SGD") -> str:
    """Format numeric currency value to readable shorthand ($T, $B, $M, $K)."""
    if value is None or pd.isna(value) or value <= 0:
        return "N/A"
    
    prefix = f"{currency} " if currency else "$"
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


def market_cap_number_format(currency: str) -> str:
    """
    Excel custom number format that *displays* a raw market cap like format_currency does
    ("SGD 1.23B") while the cell keeps the real number, so it sorts/filters numerically.
    Each trailing comma scales the shown value down by 1,000 (",,," = billions).
    """
    p = f'"{currency} "' if currency else ""
    return f'[>=1000000000000]{p}0.00,,,,"T";[>=1000000000]{p}0.00,,,"B";{p}0.00,,"M"'


def parse_percent_text(value: Any) -> Optional[float]:
    """'96.73%' -> 96.73 (None for 'N/A'/blank) - for writing percent columns as numbers."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    return safe_float(str(value).strip().rstrip("%"))


def safe_float(val: Any) -> Optional[float]:
    """Safely convert value to float, handling strings like 'Infinity', None, NaN."""
    if val is None or pd.isna(val):
        return None
    try:
        f = float(val)
        return f if f == f else None  # Check NaN
    except (ValueError, TypeError):
        return None


# Common non-SGD currencies reported by SGX-listed companies
_SGX_FOREIGN_CURRENCIES = ["USD", "CNY", "HKD", "AUD", "GBP", "EUR", "JPY", "THB", "MYR", "IDR", "INR"]


def get_sgd_fx_rates() -> Dict[str, float]:
    """
    Fetch live exchange rates from Yahoo Finance's native forex tickers
    (e.g. USDSGD=X, CNYSGD=X) to convert any foreign-currency market cap to SGD.
    No third-party currency API required — fully native via yfinance.
    Returns a dict mapping source currency -> rate to multiply by to get SGD.
    e.g. {"USD": 1.35, "CNY": 0.187, "HKD": 0.173, ...}  (SGD per 1 unit of currency)
    """
    rates: Dict[str, float] = {"SGD": 1.0}  # SGD -> SGD is always 1:1
    pairs_to_fetch = [f"{ccy}SGD=X" for ccy in _SGX_FOREIGN_CURRENCIES]

    try:
        logger.info("Fetching SGD FX rates from Yahoo Finance forex tickers...")
        # yf.download fetches multiple tickers in one request — most efficient
        raw = yf.download(
            pairs_to_fetch,
            period="1d",
            interval="1d",
            auto_adjust=True,
            progress=False,
        )

        # yf.download returns a MultiIndex DataFrame when multiple tickers given
        if not raw.empty:
            close_data = raw["Close"] if "Close" in raw.columns else raw
            for pair in pairs_to_fetch:
                ccy = pair.replace("SGD=X", "")
                try:
                    rate = None
                    if isinstance(close_data, pd.DataFrame) and pair in close_data.columns:
                        series = close_data[pair].dropna()
                        if not series.empty:
                            rate = float(series.iloc[-1])
                    elif isinstance(close_data, pd.Series):
                        series = close_data.dropna()
                        if not series.empty:
                            rate = float(series.iloc[-1])
                    if rate and rate > 0:
                        rates[ccy] = rate
                except Exception:
                    pass
    except Exception as e:
        logger.warning(f"yf.download FX batch failed ({e}). Falling back to individual ticker lookups...")

    # For any currency still missing, try individual Ticker().info lookups
    missing = [ccy for ccy in _SGX_FOREIGN_CURRENCIES if ccy not in rates]
    for ccy in missing:
        try:
            pair = f"{ccy}SGD=X"
            info = yf.Ticker(pair).info
            rate = safe_float(info.get("regularMarketPrice") or info.get("previousClose"))
            if rate and rate > 0:
                rates[ccy] = rate
        except Exception:
            pass

    fetched = {k: round(v, 6) for k, v in rates.items() if k != "SGD"}
    logger.info(f"SGD FX rates fetched: {fetched}")
    return rates


def get_sti30_tickers(with_suffix: bool = True, timeout: int = 10) -> List[str]:
    """
    Extract the 30 Straits Times Index (STI 30) constituent symbols.
    """
    url = "https://en.wikipedia.org/wiki/Straits_Times_Index"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    try:
        response = requests.get(url, headers=headers, timeout=timeout)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            tables = soup.find_all("table", {"class": "wikitable"})

            for table in tables:
                headers_text = [th.get_text(strip=True) for th in table.find_all("th")]
                if any("stock" in h.lower() for h in headers_text):
                    rows = table.find_all("tr")[1:]
                    tickers = []
                    for row in rows:
                        cols = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                        if len(cols) >= 2:
                            raw_code = clean_text(cols[0])
                            code = raw_code.replace("SGX:", "").replace("SGX", "").strip().upper()
                            tickers.append(code)
                    if len(tickers) == 30:
                        return [f"{c}.SI" if with_suffix else c for c in tickers]
    except Exception:
        pass

    return [f"{code}.SI" if with_suffix else code for code, _, _ in FALLBACK_STI_30]


def get_sgx_universe(
    include_types: Optional[List[str]] = None,
    sti_codes: Optional[Set[str]] = None
) -> List[Dict[str, Any]]:
    """
    Query the SGX securities directory to get all active listed stock codes.
    """
    if include_types is None:
        include_types = ["stocks", "reits", "businesstrusts"]

    if sti_codes is None:
        sti_codes = set(get_sti30_tickers(with_suffix=False))

    url = "https://api.sgx.com/securities/v1.1?params=nc,fn,p"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.sgx.com",
        "Referer": "https://www.sgx.com/",
    }

    securities: List[Dict[str, Any]] = []
    seen_codes: Set[str] = set()

    try:
        logger.info("Discovering listed companies from SGX directory feed...")
        response = requests.get(url, headers=headers, timeout=12)
        if response.status_code == 200:
            raw_data = response.json().get("data", {}).get("prices", [])
            for item in raw_data:
                code = item.get("nc")
                sec_type = item.get("type", "stocks")
                if not code or not isinstance(code, str):
                    continue

                code = code.strip().upper()
                if code in seen_codes:
                    continue

                if "all" in include_types or sec_type in include_types or code in sti_codes:
                    securities.append({
                        "code": code,
                        "symbol": f"{code}.SI",
                        "type": sec_type,
                        "name": item.get("fn") or code,
                        "is_sti30": code in sti_codes
                    })
                    seen_codes.add(code)

            # Ensure all 30 STI constituents are guaranteed to be present
            for sti_code in sti_codes:
                if sti_code not in seen_codes:
                    securities.append({
                        "code": sti_code,
                        "symbol": f"{sti_code}.SI",
                        "type": "reits" if sti_code.endswith("U") else "stocks",
                        "name": sti_code,
                        "is_sti30": True
                    })
                    seen_codes.add(sti_code)

            logger.info(f"Discovered {len(securities)} listed instruments.")
            if securities:
                return securities
    except Exception as e:
        logger.warning(f"Live directory query issue ({e}). Using fallback ticker list.")

    # Fallback
    for code, name, _ in FALLBACK_STI_30:
        securities.append({
            "code": code,
            "symbol": f"{code}.SI",
            "type": "reits" if code.endswith("U") else "stocks",
            "name": name,
            "is_sti30": True
        })
    return securities


def fetch_latest_ticker_data(
    security: Dict[str, Any],
    fx_rates: Optional[Dict[str, float]] = None,
    max_retries: int = 2
) -> Dict[str, Any]:
    """
    Fetch the latest market cap, 52-week pricing, trailing P/E, sector, and industry for a single ticker.
    Calculates the 52-week range and the percentage positioning metric.

    fx_rates: dict mapping currency code -> SGD rate (e.g. {"USD": 1.35, "CNY": 0.187}).
              If provided, _MarketCapSGD is computed for accurate SGD-normalised ranking.
              MarketCap column is displayed in SGD for all tickers.
    """
    if fx_rates is None:
        fx_rates = {"SGD": 1.0}

    symbol = security["symbol"]
    code = security["code"]
    sec_type = security["type"]
    initial_name = security.get("name", code)
    is_sti30 = security.get("is_sti30", False)
    index_name = security.get("index")

    record = {
        "Symbol": symbol,
        "Code": code,
        "Name": initial_name,
        "Type": sec_type,
        "Is_STI30": is_sti30,
        "Index": index_name,
        "_MarketCapRaw": None,   # In native currency, for reference
        "_MarketCapSGD": None,   # Converted to SGD, used for ranking
        "MarketCap": "N/A",      # Displayed as SGD
        "TodaysClose": None,
        "SharesOutstanding": None,
        "Currency": "SGD",       # Native currency of the ticker
        "52WHigh": None,
        "52WLow": None,
        "52WRange": None,
        "Current_Proximity_to_52W_High": "N/A",
        "Trailing_PE": None,
        "Sector": "N/A",
        "Industry": "N/A",
        "DividendYield": None,
    }

    for attempt in range(1, max_retries + 1):
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info

            if info:
                name = info.get("shortName") or info.get("longName")
                if name and not ("," in name and "0P00" in name):
                    record["Name"] = name

                # Market Cap (raw in native currency)
                raw_mktcap = safe_float(info.get("marketCap"))

                # Today's Close / Current Price
                price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"))
                record["TodaysClose"] = round(price, 2) if price is not None else None
                record["Currency"] = info.get("currency") or "SGD"

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

            # Fallback to direct Yahoo Chart API if yf.Ticker().info is empty / rate-limited
            if not info or record["TodaysClose"] is None:
                try:
                    c_url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=1d&interval=1d"
                    c_headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
                    }
                    c_resp = requests.get(c_url, headers=c_headers, timeout=8)
                    if c_resp.status_code == 200:
                        c_meta = c_resp.json().get("chart", {}).get("result", [{}])[0].get("meta", {})
                        if c_meta:
                            name = c_meta.get("shortName") or c_meta.get("longName")
                            if name and not ("," in name and "0P00" in name):
                                record["Name"] = name
                            price = safe_float(c_meta.get("regularMarketPrice") or c_meta.get("chartPreviousClose"))
                            if price is not None:
                                record["TodaysClose"] = round(price, 2)
                                record["Currency"] = c_meta.get("currency") or "SGD"
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

            # --- SGD Normalisation ---
            # Convert market cap to SGD using the pre-fetched Yahoo Finance FX rate.
            # _MarketCapSGD is what we sort and rank by; MarketCap is displayed in SGD.
            if record["_MarketCapRaw"] is not None and record["_MarketCapRaw"] > 0:
                native_ccy = record["Currency"]
                rate_to_sgd = fx_rates.get(native_ccy, 1.0)  # Default 1.0 if no rate (treat as SGD)
                if native_ccy not in fx_rates and native_ccy != "SGD":
                    logger.warning(f"{symbol}: No FX rate for {native_ccy} -> SGD. Treating as 1:1 (may affect ranking).")
                mktcap_sgd = record["_MarketCapRaw"] * rate_to_sgd
                record["_MarketCapSGD"] = mktcap_sgd
                # Display market cap in SGD regardless of native currency
                record["MarketCap"] = format_currency(mktcap_sgd, "SGD")
                return record
            else:
                if attempt == max_retries:
                    return record

        except Exception as e:
            if attempt == max_retries:
                return record
            time.sleep(0.3 * attempt)

    return record


def extract_index_stocks(
    max_workers: int = 10
) -> pd.DataFrame:
    """
    DEFAULT MODE:
    1. Uses the hardcoded SGX_INDEX_STOCKS list (STI, S Cap, L & M Cap, plus any Personal
       additions - generated from 'stock list.xlsx' by sync_stock_lists.py).
    2. Converts all market caps to SGD using Yahoo Finance native forex tickers (e.g. USDSGD=X).
    3. Fetches latest metrics (52W High, 52W Low, 52W Range, Proximity to 52W High,
       Trailing P/E, Sector, Industry, DividendYield).
    4. Ranks all index constituents by SGD-normalised Market Cap for accurate cross-currency ranking.
    5. Sets Inclusion_Reason to reflect the index the stock is from ('STI', 'S Cap', or 'L & M Cap').
    """
    # 1. Load the hardcoded index stock list (plus Personal additions)
    securities = list(SGX_INDEX_STOCKS)
    total_count = len(securities)
    logger.info(f"Loaded {total_count} stocks across STI, S Cap, L & M Cap, and Personal additions.")

    # 2. Fetch live SGD FX rates via Yahoo Finance native forex tickers
    fx_rates = get_sgd_fx_rates()

    # 3. Concurrent ticker data extraction
    logger.info(f"Starting latest data extraction for {total_count} index constituents ({max_workers} threads)...")
    results = []
    start_time = time.time()
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_sec = {executor.submit(fetch_latest_ticker_data, sec, fx_rates): sec for sec in securities}
        
        for future in as_completed(future_to_sec):
            data = future.result()
            results.append(data)
            completed_count += 1

            if completed_count % 20 == 0 or completed_count == total_count:
                elapsed = time.time() - start_time
                pct = (completed_count / total_count) * 100
                rate = completed_count / elapsed if elapsed > 0 else 0
                logger.info(f"Progress: {completed_count}/{total_count} ({pct:.1f}%) | {rate:.1f} tickers/sec")

    # Retry pass for any constituents where market cap failed
    failed = [r for r in results if r.get("_MarketCapSGD") is None or r.get("_MarketCapSGD") <= 0]
    if failed:
        logger.info(f"Running retry pass for {len(failed)} constituents...")
        time.sleep(1.0)
        for item in failed:
            sec_meta = next((s for s in securities if s["code"] == item["Code"]), {
                "symbol": item["Symbol"], "code": item["Code"], "type": item["Type"],
                "name": item["Name"], "index": item.get("Index"), "is_sti30": item.get("Is_STI30", False)
            })
            refetched = fetch_latest_ticker_data(sec_meta, fx_rates=fx_rates, max_retries=3)
            if refetched.get("_MarketCapSGD") is not None:
                for idx, orig in enumerate(results):
                    if orig["Code"] == item["Code"]:
                        results[idx] = refetched
                        break
            time.sleep(0.2)

    # Security type (reits / businesstrusts / stocks) from each stock's fetched industry -
    # done here, after the retry pass, so it uses the final fetched data for every row.
    for record in results:
        if record.get("Type") is None:
            record["Type"] = classify_sgx_security(record["Symbol"], record.get("Industry"), record.get("Name"))

    full_df = pd.DataFrame(results)

    # 4. Sort descending by SGD-normalised market cap across all index constituents
    full_df = full_df.sort_values(by="_MarketCapSGD", ascending=False, na_position="last").reset_index(drop=True)
    full_df["SGX_Rank"] = range(1, len(full_df) + 1)

    # 5. Inclusion Reason reflects the index the stock is from (STI, S Cap, L & M Cap)
    full_df["Inclusion_Reason"] = full_df["Index"].fillna("Index")

    # Drop internal/intermediate columns
    # Keep the numeric SGD market cap (as _MarketCapValue) for the Excel Summary sheet,
    # which writes it as a real number; drop the other internal columns.
    full_df = full_df.rename(columns={"_MarketCapSGD": "_MarketCapValue"})
    full_df = full_df.drop(columns=["Is_STI30", "Index", "_MarketCapRaw"], errors="ignore")

    # Final export column order
    ordered_cols = [
        "SGX_Rank", "Symbol", "Code", "Name", "Type",
        "Inclusion_Reason", "MarketCap", "TodaysClose", "SharesOutstanding",
        "52WHigh", "52WLow", "52WRange", "Current_Proximity_to_52W_High",
        "Trailing_PE",
        "Sector", "Industry", "Currency", "DividendYield",
        "_MarketCapValue",
    ]
    existing_cols = [c for c in ordered_cols if c in full_df.columns]
    combined_df = full_df[existing_cols]

    elapsed_total = time.time() - start_time
    logger.info(f"Index extraction completed in {elapsed_total:.2f}s. Total output rows: {len(combined_df)}.")
    return combined_df


def extract_top50_with_sti30(
    include_types: Optional[List[str]] = None,
    max_workers: int = 10,
    top_n: int = 50
) -> pd.DataFrame:
    """
    MARKET CAP SCANNING MODE:
    1. Extracts latest market cap, 52-week metrics, and fundamentals for all SGX stocks.
    2. Converts all market caps to SGD using Yahoo Finance native forex tickers (e.g. USDSGD=X).
    3. Ranks all instruments by SGD-normalised Market Cap for accurate cross-currency comparison.
    4. Selects Top N and tags Inclusion_Reason ('STI' or 'Top X').
    5. Automatically finds any STI 30 members that did not make the Top N,
       appends them onto the end, and tags them as 'STI (Rank #...)'.
    """
    # 1. Get STI 30 codes
    sti_codes: Set[str] = set(get_sti30_tickers(with_suffix=False))
    logger.info(f"Loaded {len(sti_codes)} STI 30 benchmark constituent codes.")

    # 2. Fetch live SGD FX rates via Yahoo Finance native forex tickers
    fx_rates = get_sgd_fx_rates()

    # 3. Get SGX universe
    securities = get_sgx_universe(include_types=include_types, sti_codes=sti_codes)
    total_count = len(securities)
    logger.info(f"Starting latest data extraction for {total_count} instruments ({max_workers} threads)...")

    results = []
    start_time = time.time()
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_sec = {executor.submit(fetch_latest_ticker_data, sec, fx_rates): sec for sec in securities}
        
        for future in as_completed(future_to_sec):
            data = future.result()
            results.append(data)
            completed_count += 1

            if completed_count % 50 == 0 or completed_count == total_count:
                elapsed = time.time() - start_time
                pct = (completed_count / total_count) * 100
                rate = completed_count / elapsed if elapsed > 0 else 0
                logger.info(f"Progress: {completed_count}/{total_count} ({pct:.1f}%) | {rate:.1f} tickers/sec")

    # Targeted Retry Pass for any missing STI 30 constituents
    failed_sti = [r for r in results if r["Code"] in sti_codes and (r["_MarketCapSGD"] is None or r["_MarketCapSGD"] <= 0)]
    if failed_sti:
        logger.info(f"Running dedicated retry pass for {len(failed_sti)} STI constituents...")
        time.sleep(1.0)
        for item in failed_sti:
            refetched = fetch_latest_ticker_data(
                {"symbol": item["Symbol"], "code": item["Code"], "type": item["Type"], "name": item["Name"], "is_sti30": True},
                fx_rates=fx_rates,
                max_retries=3
            )
            if refetched["_MarketCapSGD"] is not None:
                for idx, orig in enumerate(results):
                    if orig["Code"] == item["Code"]:
                        results[idx] = refetched
                        break
            time.sleep(0.2)

    full_df = pd.DataFrame(results)

    # 4. Sort descending by SGD-normalised market cap across entire SGX universe
    full_df = full_df.sort_values(by="_MarketCapSGD", ascending=False, na_position="last").reset_index(drop=True)
    full_df["SGX_Rank"] = range(1, len(full_df) + 1)

    # 5. Extract Top N and assign single merged 'Inclusion_Reason' column
    top50_df = full_df.head(top_n).copy()
    top50_df["Inclusion_Reason"] = top50_df["Is_STI30"].apply(
        lambda is_sti: "STI" if is_sti else f"Top {top_n}"
    )

    # Identify STI 30 members within the Top N
    top50_sti_codes = set(top50_df[top50_df["Is_STI30"]]["Code"])
    missing_sti_codes = sti_codes - top50_sti_codes

    logger.info(f"STI 30 constituents inside Top {top_n}: {len(top50_sti_codes)}/30")
    if missing_sti_codes:
        logger.info(f"STI 30 constituents outside Top {top_n} ({len(missing_sti_codes)}): {sorted(list(missing_sti_codes))}")
        
        # 6. Extract the missing STI 30 constituents and append to end of list
        missing_sti_df = full_df[full_df["Code"].isin(missing_sti_codes)].copy()
        missing_sti_df["Inclusion_Reason"] = missing_sti_df["SGX_Rank"].apply(
            lambda r: f"STI (Rank #{r})" if pd.notnull(r) else "STI"
        )
        
        # Combine Top N + remaining STI 30
        combined_df = pd.concat([top50_df, missing_sti_df], ignore_index=True)
    else:
        logger.info(f"All 30 STI constituents are already present in the Top {top_n}!")
        combined_df = top50_df

    # Drop internal/redundant columns
    # Keep the numeric SGD market cap (as _MarketCapValue) for the Excel Summary sheet,
    # which writes it as a real number; drop the other internal columns.
    combined_df = combined_df.rename(columns={"_MarketCapSGD": "_MarketCapValue"})
    combined_df = combined_df.drop(columns=["Is_STI30", "Index", "_MarketCapRaw"], errors="ignore")

    # Final export column order
    ordered_cols = [
        "SGX_Rank", "Symbol", "Code", "Name", "Type",
        "Inclusion_Reason", "MarketCap", "TodaysClose", "SharesOutstanding",
        "52WHigh", "52WLow", "52WRange", "Current_Proximity_to_52W_High",
        "Trailing_PE",
        "Sector", "Industry", "Currency", "DividendYield",
        "_MarketCapValue",
    ]
    existing_cols = [c for c in ordered_cols if c in combined_df.columns]
    combined_df = combined_df[existing_cols]

    elapsed_total = time.time() - start_time
    logger.info(f"Market cap scan completed in {elapsed_total:.2f}s. Total output rows: {len(combined_df)}.")
    return combined_df


def print_stock_table(df: pd.DataFrame):
    """
    Display a formatted terminal table with all metrics.
    Adapts header and summary stats based on analysis mode (Index vs Market Cap).
    """
    W = 138
    is_index_mode = any(r in df["Inclusion_Reason"].values for r in ["S Cap", "L & M Cap"])
    
    print("\n" + "=" * W)
    if is_index_mode:
        print(" SGX INDEX BENCHMARK CONSTITUENTS (STI, S Cap, L & M Cap)")
    else:
        print(" TOP SGX COMPANIES & STI BENCHMARK CONSTITUENTS")
    print("=" * W)

    header = (
        f"{'Rank':<5} | {'Symbol':<8} | {'Company Name':<22} | "
        f"{'Inclusion':<18} | {'Market Cap':<12} | {'Close':<7} | "
        f"{'52W H':<7} | {'52W L':<7} | {'52W Rng':<7} | {'Prox. 52W High':<15} | "
        f"{'Tr.PE':<6} | {'Sector'}"
    )
    print(header)
    print("-" * W)

    overflow_count = 0

    for _, row in df.iterrows():
        sgx_rank   = f"#{int(row['SGX_Rank'])}" if pd.notnull(row.get('SGX_Rank')) else "N/A"
        symbol     = row["Symbol"]
        name       = (row["Name"][:20] + "..") if len(str(row["Name"])) > 22 else row["Name"]
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
            print(" >>> STI CONSTITUENTS OUTSIDE TOP N (APPENDED BELOW):")
            print("-" * W)

        if "Rank #" in reason:
            overflow_count += 1

        print(
            f"{sgx_rank:<5} | {symbol:<8} | {name:<22} | "
            f"{reason:<18} | {mkt_cap:<12} | {price:<7} | "
            f"{h52:<7} | {l52:<7} | {rng:<7} | {prox:<15} | "
            f"{trailing_pe:<6} | {sector}"
        )

    print("=" * W)
    if is_index_mode:
        sti_count = df[df["Inclusion_Reason"] == "STI"].shape[0]
        scap_count = df[df["Inclusion_Reason"] == "S Cap"].shape[0]
        lmcap_count = df[df["Inclusion_Reason"] == "L & M Cap"].shape[0]
        print(f"STI Index Constituents        : {sti_count}")
        print(f"S Cap Index Constituents      : {scap_count}")
        print(f"L & M Cap Index Constituents  : {lmcap_count}")
    else:
        sti_in_top = df[df["Inclusion_Reason"] == "STI"].shape[0]
        non_sti = df[df["Inclusion_Reason"].str.startswith("Top")].shape[0]
        print(f"STI Members in Top Selection  : {sti_in_top} / 30")
        print(f"Non-STI Top Companies         : {non_sti}")
        print(f"STI Members outside Top Rank  : {overflow_count} (Appended to end)")
    print(f"Total Companies in Output     : {len(df)}\n")


# Backwards compatibility alias
print_top50_sti_table = print_stock_table


def fetch_ticker_52w_history(symbol: str, code: str, name: str, shares_outstanding: float = None, max_retries: int = 3) -> pd.DataFrame:
    """
    Fetch 52-week (1 year) historical daily trading data for a single ticker.
    Uses direct Yahoo Chart API (with custom browser headers and session pooling)
    which bypasses yfinance client-side rate limits.
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
                        import datetime
                        dates = [datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d") for ts in timestamps]
                        
                        df_hist = pd.DataFrame({
                            "Date": dates,
                            "Close": closes,
                            "Volume": volumes
                        }).dropna(subset=["Close"])
                        
                        df_hist["Symbol"] = symbol
                        df_hist["Code"] = code
                        df_hist["Name"] = name
                        
                        # Calculate % change in daily close price
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


def extract_52w_history_all(summary_df: pd.DataFrame, max_workers: int = 5):
    """
    Extract 52-week historical daily data for all companies present in summary_df concurrently.
    Reformats the dataset horizontally such that each new company is added as a column group 
    to the right (aligned along Date rows) in the order of the Top 50 / STI rank.
    Returns (wide_hist_df, active_companies_metadata).
    """
    logger.info(f"Starting 52-week daily history extraction for {len(summary_df)} companies...")
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
            if completed_count % 10 == 0 or completed_count == total_count:
                logger.info(f"History progress: {completed_count}/{total_count} ({completed_count / total_count * 100:.1f}%)")

    # Merge all historical data horizontally by Date (each company added to columns on the right)
    wide_hist_df = None
    active_companies = []
    
    for sym in ordered_symbols:
        if sym not in hist_results or hist_results[sym].empty:
            continue
        sub = hist_results[sym][["Date", "Close", "DailyChangePct", "MarketCap", "SharesOutstanding", "Volume"]].copy()
        code, name = meta_map.get(sym, (sym, sym))
        active_companies.append((sym, code, name))
        
        # Clean prefix used for flat CSV / DataFrame representation
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


def fetch_fx_history(from_ccy: str, to_ccy: str, range_str: str = "5y") -> pd.DataFrame:
    """
    Fetch multi-year daily FX history as Date/Close columns, where Close = units of `to_ccy`
    per 1 unit of `from_ccy` (e.g. SGD -> THB via Yahoo's "SGDTHB=X"). Falls back to the
    inverse pair (inverting each rate) if Yahoo doesn't list the direct one. Returns an
    empty DataFrame if neither is available.
    """
    hist = fetch_long_price_history(f"{from_ccy}{to_ccy}=X", range_str=range_str)
    if hist.empty:
        inverse = fetch_long_price_history(f"{to_ccy}{from_ccy}=X", range_str=range_str)
        inverse = inverse[inverse["Close"] > 0] if not inverse.empty else inverse
        if not inverse.empty:
            hist = inverse.assign(Close=1.0 / inverse["Close"])
    return hist


def fetch_financial_history(symbol: str, code: str, name: str, max_years: int = 5) -> pd.DataFrame:
    """
    Build a per-fiscal-year fundamentals table (chronological, oldest year first): Revenue,
    Net Income, EPS, Dividend Per Unit, P/E, P/S, P/B, Gross Margin, Operating Margin,
    and Debt/Equity.

    Historical P/E, P/S, and P/B are approximations: each fiscal year's reported diluted
    EPS / average diluted shares / stockholders' equity is combined with the closing share
    price on (or immediately before) that fiscal year-end date, in the company's native
    reporting currency (NOT SGD-normalised, unlike the Summary sheet's Market Cap column).
    Yahoo Finance's free annual statements go back up to 5 years for most tickers, but
    occasionally fewer - whatever is actually available (capped at max_years) is returned.

    Dividend Per Unit is also an approximation, same spirit as P/E/P/S/P/B above: yfinance
    doesn't expose a per-fiscal-year dividend figure directly (dividends aren't a line item
    on the income statement or balance sheet), so each fiscal year's total is reconstructed
    by summing `Ticker.dividends` (a real per-payment historical series, ex-dividend date ->
    per-unit amount - "per unit" since many SGX constituents here are REITs/trusts) over
    that fiscal year's ~12-month window.
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

        # Statements (Revenue, EPS, equity, ...) are in the company's *reporting* currency,
        # while prices and dividends are in its *trading* currency - these differ for many
        # SGX listings (e.g. Thai Beverage trades in SGD but reports in THB). Valuation ratios
        # need both sides in the same currency, so prices/dividends are converted below.
        trading_currency = getattr(ticker.fast_info, "currency", None) or "N/A"
        try:
            reporting_currency = (ticker.info or {}).get("financialCurrency") or trading_currency
        except Exception:
            logger.warning(f"{symbol}: could not read reporting currency; assuming {trading_currency}")
            reporting_currency = trading_currency
        currency = reporting_currency

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

        # Historical FX (trading -> reporting currency), applied at each fiscal year-end date
        # rather than today's spot rate. None = same currency, no conversion needed.
        fx_hist = None
        if reporting_currency != trading_currency and "N/A" not in (reporting_currency, trading_currency):
            fx_hist = fetch_fx_history(trading_currency, reporting_currency, range_str="5y")
            if fx_hist.empty:
                logger.warning(
                    f"{symbol}: no {trading_currency}->{reporting_currency} FX history; "
                    f"leaving P/E, P/S, P/B and dividends blank rather than mixing currencies"
                )

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

            fx_rate = 1.0 if fx_hist is None else get_price_on_or_before(fx_hist, fye)
            price = get_price_on_or_before(price_hist, fye)
            if price is not None:
                price = price * fx_rate if fx_rate is not None else None

            pe = (price / eps) if (price is not None and eps not in (None, 0)) else None
            market_cap_fye = (price * shares) if (price is not None and shares) else None
            ps = (market_cap_fye / revenue) if (market_cap_fye is not None and revenue not in (None, 0)) else None
            pb = (market_cap_fye / equity) if (market_cap_fye is not None and equity not in (None, 0)) else None
            gross_margin = (gross_profit / revenue * 100.0) if (gross_profit is not None and revenue not in (None, 0)) else None
            op_margin = (op_income / revenue * 100.0) if (op_income is not None and revenue not in (None, 0)) else None
            debt_to_equity = (debt / equity) if (debt is not None and equity not in (None, 0)) else None

            # Dividend/Distribution Per Unit for this fiscal year: sum of per-unit dividend
            # payments (ex-dividend date -> amount, from yfinance) falling within this fiscal
            # year's ~12-month window - since the previous period's fiscal year-end if known,
            # otherwise approximated as 1 year before this period's fiscal year-end.
            dps = None
            if dividends is not None and not dividends.empty:
                window_end = pd.Timestamp(fye)
                window_start = prev_fye if prev_fye is not None else (window_end - pd.DateOffset(years=1))
                in_window = dividends[(dividends.index > window_start) & (dividends.index <= window_end)]
                if not in_window.empty and fx_rate is not None:
                    dps = float(in_window.sum()) * fx_rate
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
    sheet_name: str = "SGX Stock Summary",
    financial_histories: Optional[Dict[str, pd.DataFrame]] = None,
):
    """
    Save extracted dataset. If output is an Excel file (.xlsx), both summary and
    historical daily data are saved as separate worksheets in the same workbook.
    Formats numbers across both worksheets to 2 decimal places and abbreviates large
    numbers (Market Cap, Shares Outstanding, Volume) with B / M suffixes.

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
            from openpyxl.worksheet.table import Table, TableStyleInfo

            # Prepare formatted Summary DataFrame for Excel Sheet 1
            df_summary_fmt = df_summary.copy()

            # SharesOutstanding stays out of the Summary sheet specifically (still used
            # internally to compute the 52W History sheet's Market Cap column, via the
            # untouched df_summary passed into extract_52w_history_all before this point -
            # dropping it only from this formatted copy doesn't affect that).
            df_summary_fmt = df_summary_fmt.drop(columns=["SharesOutstanding"], errors="ignore")

            # Numeric columns are written as real numbers (not pre-formatted text) so the
            # Summary table's filter/sort dropdowns order them numerically. Their display
            # (2 decimals, "SGD 1.23B", "96.73%") comes from Excel number formats applied
            # after writing - see summary_number_formats below.
            if "_MarketCapValue" in df_summary_fmt.columns:
                # Swap the display string ("SGD 1.23B") for the underlying number.
                cap_values = pd.to_numeric(df_summary_fmt["_MarketCapValue"], errors="coerce")
                df_summary_fmt["MarketCap"] = cap_values.where(cap_values > 0)
                df_summary_fmt = df_summary_fmt.drop(columns=["_MarketCapValue"])
            if "Current_Proximity_to_52W_High" in df_summary_fmt.columns:
                df_summary_fmt["Current_Proximity_to_52W_High"] = df_summary_fmt["Current_Proximity_to_52W_High"].apply(parse_percent_text)
            summary_number_formats = {
                "SGX_Rank": "0",
                "NSE_Rank": "0",
                "TodaysClose": "#,##0.00",
                "52WHigh": "#,##0.00",
                "52WLow": "#,##0.00",
                "52WRange": "#,##0.00",
                "Current_Proximity_to_52W_High": '0.00"%"',
                "Trailing_PE": "#,##0.00",
                "Forward_PE": "#,##0.00",
                "DividendYield": '0.00"%"',   # Yahoo reports this in percent points already (4.08 = 4.08%)
            }
            for col in list(summary_number_formats) + ["MarketCap"]:
                if col in df_summary_fmt.columns and (col != "MarketCap" or pd.api.types.is_numeric_dtype(df_summary_fmt[col])):
                    df_summary_fmt[col] = pd.to_numeric(df_summary_fmt[col], errors="coerce")

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                # 1. Write formatted Summary Sheet
                df_summary_fmt.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Polish Summary Sheet formatting
                wb = writer.book
                ws_summary = wb[sheet_name]
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

                # Format the Summary sheet's full range as a real Excel Table (not just
                # colored header cells) - gives it native header filter/sort dropdowns and
                # banded rows for free, no manual autofilter/formatting needed. Direct cell
                # formatting (the navy header fill/font above) still renders on top of the
                # table style, so the header keeps its look.
                if not df_summary_fmt.empty:
                    last_col_letter = get_column_letter(len(df_summary_fmt.columns))
                    last_row = len(df_summary_fmt) + 1
                    summary_table = Table(displayName="SummaryTable", ref=f"A1:{last_col_letter}{last_row}")
                    summary_table.tableStyleInfo = TableStyleInfo(
                        name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False
                    )
                    ws_summary.add_table(summary_table)
                    ws_summary.freeze_panes = "A2"

                # Display formats for the numeric columns (values themselves stay numbers).
                # Market Cap uses SGD (every SGX market cap is converted to SGD).
                summary_cols = list(df_summary_fmt.columns)
                currency_col = summary_cols.index("Currency") + 1 if "Currency" in summary_cols else None
                market_cap_numeric = "MarketCap" in summary_cols and pd.api.types.is_numeric_dtype(df_summary_fmt["MarketCap"])
                for c_idx, col_name in enumerate(summary_cols, start=1):
                    fmt = summary_number_formats.get(col_name)
                    if col_name == "MarketCap" and market_cap_numeric:
                        fmt = "market_cap"
                    if not fmt:
                        continue
                    for r_idx in range(2, len(df_summary_fmt) + 2):
                        cell = ws_summary.cell(row=r_idx, column=c_idx)
                        if fmt == "market_cap":
                            cell.number_format = market_cap_number_format("SGD")
                        else:
                            cell.number_format = fmt

                # 2. Write formatted 52W Daily History Sheet
                history_ws = None
                history_col_anchor: Dict[str, int] = {}  # Symbol -> its start column on this sheet
                if df_history is not None and not df_history.empty:
                    ws = wb.create_sheet("52W Daily History")
                    history_ws = ws

                    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    sub_header_font = Font(name="Calibri", size=10, bold=True, color="1F2937")
                    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")     # Dark Navy
                    sub_header_fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid") # Soft Indigo
                    
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
                    
                    # Populate Top Merged Company Headers and Sub-headers
                    col_idx = 2
                    for sym, code, name in active_companies:
                        start_col = col_idx
                        end_col = col_idx + len(metrics) - 1
                        history_col_anchor[sym] = start_col

                        start_letter = get_column_letter(start_col)
                        end_letter = get_column_letter(end_col)

                        # Row 1: Stock Name and Code (Merged across its 5 metric columns)
                        ws.cell(row=1, column=start_col, value=f"{code} - {name}")
                        ws.merge_cells(f"{start_letter}1:{end_letter}1")
                        
                        for c in range(start_col, end_col + 1):
                            cell = ws.cell(row=1, column=c)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = center_align
                            cell.border = thin_border
                            
                        # Row 2: Metric column names
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
                        
                        # Identify metric type by column suffix to apply appropriate 2dp / B / M abbreviation
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

                    # Auto-fit column widths
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
                        back_cell.hyperlink = f"#'{sheet_name}'!A1"
                        back_cell.font = back_link_font
                        row_cursor += 2  # spacer row before next block

                    ws_fin.column_dimensions["A"].width = 30
                    for col in range(2, n_fin_cols + 1):
                        ws_fin.column_dimensions[get_column_letter(col)].width = 14

                    # 4. Hyperlink each company's Name cell on the Summary sheet to its block above
                    if "Name" in df_summary.columns and "Symbol" in df_summary.columns:
                        name_col_idx = list(df_summary_fmt.columns).index("Name") + 1
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
                        back_cell.hyperlink = f"#'{sheet_name}'!A1"
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
                        symbol_col_idx = list(df_summary_fmt.columns).index("Symbol") + 1
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
        df_summary.drop(columns=["_MarketCapValue"], errors="ignore").to_json(output_path, orient="records", indent=2)
        if df_history is not None and not df_history.empty:
            hist_path = f"{base}_history.json"
            df_history.to_json(hist_path, orient="records", indent=2)
            logger.info(f"History dataset saved to: {os.path.abspath(hist_path)}")
        logger.info(f"Summary dataset successfully saved to: {os.path.abspath(output_path)}")
    else:
        base, _ = os.path.splitext(output_path)
        df_summary.drop(columns=["_MarketCapValue"], errors="ignore").to_csv(output_path, index=False)
        if df_history is not None and not df_history.empty:
            hist_path = f"{base}_history.csv"
            df_history.to_csv(hist_path, index=False)
            logger.info(f"History dataset saved to: {os.path.abspath(hist_path)}")
        logger.info(f"Summary dataset successfully saved to: {os.path.abspath(output_path)}")


def main():
    parser = argparse.ArgumentParser(
        description="SGX Stock Analysis: Analyzes SGX Index constituents (STI, S Cap, L & M Cap) by default, with optional full market-cap scanning."
    )
    parser.add_argument(
        "--scan-market-cap", "--market-cap",
        action="store_true",
        dest="scan_market_cap",
        help="Scan entire SGX market by market cap (Top N + STI catch-up) instead of the default index stock list."
    )
    parser.add_argument(
        "--mode",
        type=str,
        choices=["index", "market-cap"],
        default="index",
        help="Analysis mode: 'index' (default, the built-in index + Personal stock list) or 'market-cap' (scans entire SGX universe)."
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
        default=50,
        help="Base top N companies by market cap to extract when market-cap mode is used (default: 50)."
    )
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    default_filename = f"SGX_Stock_Analysis_{today_str}.xlsx"

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

    is_market_cap_mode = args.scan_market_cap or (args.mode == "market-cap")

    if is_market_cap_mode:
        logger.info(f"Running SGX Analysis in MARKET-CAP SCAN mode (Top {args.top} + STI)...")
        df_summary = extract_top50_with_sti30(
            max_workers=args.workers,
            top_n=args.top
        )
        sheet_title = f"Top {args.top} & STI Summary"
    else:
        logger.info("Running SGX Analysis in DEFAULT INDEX mode (built-in index + Personal stock list)...")
        df_summary = extract_index_stocks(
            max_workers=args.workers
        )
        sheet_title = "SGX Index Summary"

    print_stock_table(df_summary)

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
    save_results(df_summary, df_history, active_companies, args.output, sheet_name=sheet_title, financial_histories=financial_histories)


if __name__ == "__main__":
    main()

