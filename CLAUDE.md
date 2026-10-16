# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Stock Market Analysis (NSE + SGX)

Extracts live market-cap/valuation snapshots, 52-week daily price history, and
5-year fundamentals history for NSE (India) and SGX (Singapore) listed
companies, and writes each to a formatted, cross-linked Excel workbook. A
Jupyter notebook mirrors both scripts so the whole pipeline can run on Google
Colab without any local setup.

This file is meant to travel with the project (plain markdown, no
machine-specific content) so a Claude Code session opened in a copy of this
directory — on any OS — has the same context as the session that wrote it.

## Files

- `NSE_Stock_Analysis.py` / `SGX_Stock_Analysis.py` — the two standalone,
  independently runnable pipelines. **These are the source of truth for all
  logic.** Each contains one *generated* block (its stock list, between
  `BEGIN/END GENERATED STOCK LIST` markers) that is rewritten from
  `stock list.xlsx` by `sync_stock_lists.py` — never hand-edit inside those
  markers. See "Stock lists" below.
- `stock list.xlsx` — **the source of truth for which stocks are analysed.**
  One worksheet each (`SGX`, `NSE`), rows of Index/Inclusion-reason + Company
  + Ticker. Read only at sync time, never at pipeline runtime.
- `sync_stock_lists.py` — reads `stock list.xlsx`, validates every row, and
  regenerates the stock-list block in both scripts. Runs automatically as the
  first step of `update_notebook.py`; can also be run on its own.
- `Stock_Market_Analysis_SGX_NSE.ipynb` — Colab notebook. Generated/kept in
  sync from the two scripts above; never hand-edit its code cells directly,
  see "Notebook sync" below.
- `update_notebook.py` — **the real sync script.** Runs `sync_stock_lists.py`
  first, then slices the current, literal contents of both `.py` files into the
  notebook's two code cells (plus a short exec block calling `main()`'s
  pipeline) and refreshes the markdown cells to match. Run this after *any*
  change to either script or to `stock list.xlsx`.
- `create_notebook.py` — one-off script that originally built the notebook
  from a hand-condensed rewrite of the two scripts. Superseded by
  `update_notebook.py`; keep for reference but don't run it (it would
  regenerate the notebook from a stale, drifted copy of the logic).
- `requirements.txt` — `pandas`, `yfinance`, `requests`, `beautifulsoup4`,
  `openpyxl`. All pure-Python / cross-platform, nothing OS-specific.
  (`openpyxl` is needed both for writing the output workbooks and for
  `sync_stock_lists.py` reading `stock list.xlsx`.)
- `.gitignore` — excludes all `*.xlsx` except `stock list.xlsx` (so generated
  outputs under any `--output` name are ignored), `*.csv`, `*.json`,
  `archive/`, `venv/`, Python/Jupyter caches, and Windows/WSL
  `*:Zone.Identifier` metadata files. Generated outputs are reproducible by
  re-running the scripts, so they're deliberately not tracked.
- `archive/` — local-only (gitignored) folder of past output workbooks, kept
  in dated subfolders for comparing runs over time. Not used by any script.
- `update_and_share.sh` — one-command "save progress" wrapper: runs
  `update_notebook.py` (which syncs the stock lists first), compile-checks all
  four `.py` files, `git add`/`commit` (skipped if nothing changed) /`push`,
  then refreshes the shared Google Drive copies of both the notebook and
  `stock list.xlsx` via `rclone` (see "Sharing workflow" below). Uses the
  project's `venv/bin/python` if a `venv/` exists (falls back to `python3`) -
  the system Python may not have `openpyxl`. Run this instead of the
  individual steps whenever a script or the stock list changes.

## Running

```
python NSE_Stock_Analysis.py [--mode fast|robust|full] [--top 80] [--workers 10]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
python SGX_Stock_Analysis.py [--mode index|market-cap] [--top 50] [--workers 10]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
python sync_stock_lists.py [--xlsx "stock list.xlsx"]
```

Default output is `<NSE|SGX>_Stock_Analysis_<today>.xlsx`. Both scripts also
accept `.csv`/`.json` output, in which case the fundamentals/chart sheets
below don't apply (Excel-only features), and Summary values stay as the
pre-formatted display strings (e.g. `MarketCap` = `"SGD 1.23B"`).

## Excel workbook design (and *why* it's built this way)

Each workbook has (up to) four sheets, in this order, with hyperlinks
threading them together:

1. **Summary** — one row per company, ranked by market cap, formatted as a
   real Excel Table (`SummaryTable`: header filter/sort dropdowns, banded
   rows, frozen header row). Deliberately does *not* include
   `SharesOutstanding` (see below) even though the underlying data is
   fetched.
2. **52W Daily History** — horizontal layout: each company gets a 5-column
   block (`Close`, `% Change`, `Market Cap`, `Shares Out.`, `Volume`) with a
   2-tier merged header, rows are trading days.
3. **Financials Detail** — a stacked block per company: up to 5 fiscal years
   (oldest → newest, top → bottom) of Revenue, Net Income, EPS, Dividend/Unit,
   P/E, P/S, P/B, Gross/Operating Margin, Debt/Equity.
4. **Price Charts** — one compact 1-year Close-price line chart per company.

Navigation: **Name** cell on Summary → jumps to that company's block on
Financials Detail; **Symbol** cell on Summary (and the company header on 52W
Daily History) → jumps to that company's chart on Price Charts. Each
destination block has a "↑ Back to Summary" link back. (A company Yahoo has
no fundamentals for gets no Financials block, so its Name cell has no link.)

Several non-obvious decisions are baked into this, worth knowing before
changing it:

- **Summary numeric cells are real numbers, displayed via Excel number
  formats — not pre-formatted text.** Text cells sort alphabetically in the
  table's filter dropdowns ("9.5" after "10.2"), so `save_results` writes
  numbers and sets `cell.number_format` per column (`summary_number_formats`):
  - Prices / 52W High/Low/Range / P/E: `#,##0.00`; ranks: `0`.
  - `Current_Proximity_to_52W_High` and `DividendYield`: the number is in
    percent points (81.87 = 81.87%; Yahoo already reports dividend yield that
    way), shown with `0.00"%"` — a literal `%` suffix, *not* Excel's `0.00%`
    (which would multiply by 100). Proximity is computed as a `"81.87%"`
    string at fetch time and parsed back to a number in `save_results`
    (`parse_percent_text`).
  - `MarketCap`: the extraction functions keep the raw market cap (SGD for
    SGX, native currency for NSE) in an internal `_MarketCapValue` column, and
    `save_results` writes that number in place of the display string.
    `market_cap_number_format()` builds a custom format that still *displays*
    it like `"SGD 210.77B"` / `"INR 16.37T"` (trailing commas each scale by
    1,000). Number filters therefore work on the full figure (> 1 billion
    means typing `1000000000`). `_MarketCapValue` is dropped from CSV/JSON
    output and never shown.
  - Missing values are blank cells (sort to the bottom), not `"N/A"`.
  - The Name/Symbol hyperlink loops locate their columns in `df_summary_fmt`
    (the frame actually written), not `df_summary` — the two differ by the
    dropped `SharesOutstanding`/`_MarketCapValue` columns.

- **Financials Detail / Price Charts are separate *visible* sheets, not one
  hidden sheet per company.** Excel refuses to navigate a hyperlink into a
  truly hidden worksheet (you get "Reference isn't valid" unless the user
  manually unhides it first) — there's no way around this without a VBA
  macro (`.xlsm`), which isn't worth the complexity here. One extra visible
  sheet per concern, with hyperlinked anchors, gets the same "don't clutter
  the workbook with 90 tabs" result without fighting that limitation.

- **Price Charts references the "52W Daily History" sheet's cells directly
  — it does NOT duplicate the data.** This went through two failed
  approaches before landing here, worth knowing if chart data ever looks
  wrong again:
  1. First attempt copied each company's Close/Date series into two hidden
     helper columns on the Price Charts sheet, using the *same* row cursor
     as the visual header placement (~18 rows) — but a trading year is
     ~252 rows, so company 2's data overwrote most of company 1's, etc.
     Fixed by giving the hidden copy its own independent row cursor.
  2. That still produced blank charts: **Excel excludes hidden rows/columns
     from anything a chart plots by default.** Since the whole backing
     table was hidden, every chart resolved to nothing.
  3. Final fix: the 52W Daily History sheet's `Close` sub-column is written
     as a real numeric float with `cell.number_format = "0.00"` (previously
     it was a Python-formatted *string* like `"1234.56"`, which a chart
     can't plot at all). Price Charts' `LineChart` series now reference
     `'52W Daily History'!$<col>$3:$<col>$<n>` directly — no copy, no
     hidden columns, nothing to go stale or collide. The other 52W History
     columns (`% Change`, `Market Cap`, `Shares Out.`, `Volume`) are still
     pre-formatted display strings since nothing needs to chart them.
  - Charts have visible X (Date, thinned to ~monthly labels via
    `tickLblSkip=21`) and Y (`Close Price (<currency>)`) axes.
  - The Y axis is explicitly scaled to each company's actual 1Y Close range
    (min/max read back from the same "52W Daily History" cells, ±5% padding,
    floored at 0) rather than left on Excel's auto-scaling. Auto-scaling
    anchors at 0 by default, which flattens a year of real fluctuation into a
    thin band for any stock that trades well above 0.

- **Historical P/E, P/S, P/B are approximations, computed by the script —
  Yahoo has no historical valuation ratios** (its P/E/P/B are current-only;
  those are what the Summary sheet's `Trailing_PE` shows). In
  `fetch_financial_history`, each fiscal year's reported diluted EPS /
  average diluted shares / stockholders' equity (from `yfinance`'s
  `income_stmt`/`balance_sheet`) is combined with the closing share price on
  (or immediately before) that fiscal year-end date:
  P/E = price ÷ EPS; P/S and P/B = (price × average diluted shares) ÷
  revenue / equity. Why these differ from a paid vendor's figures: price is
  taken at fiscal year-end (vendors often use the results-announcement date
  or a yearly average), and the share count is the year's *average* diluted
  shares, not year-end shares outstanding. Margins and Debt/Equity are plain
  arithmetic on reported figures, not approximations.
  Years requested default to 5 (`--financials-years`), capped by however many
  Yahoo actually has. **In practice this is usually 4, not 5** - Yahoo's free
  API currently only serves ~4 years of annual `income_stmt`/`balance_sheet`
  data regardless of what's requested (confirmed via multiple reports on
  yfinance's own GitHub, including from its maintainer); the 5th year only
  shows up on Yahoo Finance's paid "Plus" tier, which `yfinance` has no way
  to authenticate as. Not fixable from this project's side - if 5 full years
  is a hard requirement, it would mean switching to a different (likely
  paid) financial data provider, not a parameter tweak.

- **Each fiscal year's EPS and share count are cross-checked against net
  income** (`reconcile_eps_and_shares`, called in `fetch_financial_history`)
  because Yahoo's data has two known glitches that silently break the ratios.
  A one-off check of all 191 stocks (36 with splits/bonus issues since 2020)
  found:
  - **Split adjustment:** historical prices were adjusted at every split, and
    EPS in every testable case, but share counts *not* always — HDFC Bank's
    FY2023-24 counts were left pre-split (its 2-for-1 was Aug 2025, while
    FY2025 was adjusted), halving those years' P/S and P/B. P/E (price ÷ EPS)
    was never affected.
  - **Units:** Prime US REIT (`OXMU.SI`) FY2022 EPS was reported in cents
    (−2.08 instead of −0.0208), making P/E 100× off.
  The fix: EPS × shares should be ≈ net income (±15%). If it's ≈ 1/R, where R
  is the product of `Ticker.splits` ratios dated after that fiscal year-end,
  shares are multiplied by R. If it's ≈ R, EPS is divided by R. If it's ≈ 100,
  EPS is divided by 100. Any other mismatch is left as reported and only
  logged at INFO level ("left as reported"). That case is normal for REITs and
  trusts (part of net income goes to perpetual-securities holders) and for
  companies with large minority interests or demergers. Across all 191 stocks
  this changed exactly the 3 years above; ~27 other years are only logged.

- **Financials Detail is in each company's *reporting* currency, with
  prices and dividends converted into it at each fiscal year-end's FX rate.**
  Statements are in the reporting currency (Yahoo `info["financialCurrency"]`)
  but prices and dividends are in the *trading* currency
  (`fast_info.currency`), and the two differ for many SGX listings (e.g.
  Thai Beverage trades in SGD, reports in THB; Yangzijiang CNY; Wilmar USD).
  Dividing an SGD price by THB EPS used to make Thai Beverage's P/E ~0.4
  instead of ~11. Now `fetch_financial_history` fetches daily FX history
  (`fetch_fx_history`, e.g. `SGDTHB=X`, falling back to the inverted reverse
  pair, via the same Yahoo chart API as prices) and converts each year's
  price and dividend total at *that* fiscal year-end's rate — a historical
  rate per year, not today's spot rate. If no FX history is available, that
  stock's P/E, P/S, P/B and Dividend/Unit are left blank (with a logged
  warning) rather than mixing currencies. The block header label ("Reported
  in THB") is the reporting currency. This is still *not* SGD-converted like
  the Summary sheet's Market Cap. Same-currency stocks (all NSE, SGD-reporting
  SGX) are unaffected. The same code is in both scripts.

- `SharesOutstanding` used to be silently `None` for every row (never
  actually captured from `yfinance`'s `.info`, despite `save_results`
  already having formatting logic ready for it) — this made the 52W History
  sheet's `Market Cap`/`Shares Out.` columns permanently blank. Fixed in
  both scripts' `fetch_latest_ticker_data`. It's still fetched onto
  `df_summary` today (52W History's per-day Market Cap calc reads it from
  there), but `save_results` drops it from the *Summary sheet's* formatted
  copy specifically (`df_summary_fmt`) before writing - the underlying
  `df_summary` is untouched, so don't "fix" this by removing the fetch.

- **Dividend/Unit (Financials Detail) is also an approximation.** yfinance
  doesn't expose a per-fiscal-year dividend figure directly (dividends
  aren't a line item on the income statement or balance sheet). Each fiscal
  year's total is reconstructed by summing `Ticker.dividends` (a real
  per-payment historical series, ex-dividend date → per-share/unit amount)
  over that fiscal year's ~12-month window, ending at that year's fiscal
  year-end date (same window logic already used for the P/E/P/S/P/B
  approximation above). Grouping is by ex-dividend date, so a final dividend
  for FY2024 that goes ex-dividend in early 2025 lands in FY2025.

## Stock lists (`stock list.xlsx` → `sync_stock_lists.py` → hardcoded in each script)

Both scripts use **hardcoded** stock lists at runtime — neither reads
`stock list.xlsx` when it runs. This is deliberate: on Colab (or any other
machine) the spreadsheet isn't there, and pulling it from Google Drive at
runtime was tried and dropped as too fiddly to share. Instead the
spreadsheet is the single place to edit the lists, and `sync_stock_lists.py`
turns it into code:

```
edit + SAVE stock list.xlsx  →  python update_notebook.py  (or ./update_and_share.sh)
                                 └─ sync_stock_lists.py rewrites the generated block in both .py files
                                 └─ notebook rebuilt from the updated .py files
```

- **`SGX` worksheet → `SGX_INDEX_STOCK_ROWS`** in `SGX_Stock_Analysis.py`:
  SGX's *entire* default-mode (`--mode index`) universe — all STI / S Cap /
  L & M Cap constituents plus rows tagged `Personal` (currently 92 = 89 + 3).
  `SGX_INDEX_STOCKS` (the dicts the pipeline uses) is derived from it just
  below the markers. `--mode market-cap` ignores this list entirely (it scans
  the whole SGX market via SGX's own API), so Personal picks only appear in
  that mode if they qualify on market cap.
- **`NSE` worksheet → `NSE_PERSONAL_STOCK_ROWS`** in `NSE_Stock_Analysis.py`:
  *additions* on top of NSE's downloaded universe (Top N by market cap + every
  NIFTY 50 constituent), currently 13. `PERSONAL_NSE_STOCKS` is derived from
  it. They're *guaranteed* to appear in the output, tagged
  `Inclusion_Reason = "Personal"`, or `"<existing reason> + Personal"` if the
  stock already qualified another way (e.g. also NIFTY 50).
- **Sheet format:** a header row (`Index` / `Inclusion reason`, `Company`,
  `Ticker`) then one row per stock. Data needn't start in column A (empty
  cells are skipped). Every cell must be filled in — there is no name-lookup
  fallback any more; the Ticker is the real yfinance symbol (`5DD.SI`,
  `ACC.NS`). The Company name is only a fallback label: the output's Name
  column comes from Yahoo's `shortName`.
- **Validation is all-or-nothing:** before writing anything, the sync checks
  every row (exactly 3 filled cells; ticker ends `.SI` / `.NS`; no duplicate
  tickers within a sheet; SGX Index ∈ `STI`, `S Cap`, `L & M Cap`,
  `Personal`; NSE reason = `Personal`). Any problem → every problem is
  printed with its Excel row number, exit code 1, and neither script is
  touched (and `update_notebook.py` aborts before touching the notebook). It
  also fails if a script's BEGIN/END markers are missing or duplicated.
- The sync only rewrites a script if its block actually changed, prints
  `+added / -removed` per script, and is deterministic (sheet order, strings
  via `json.dumps` so apostrophes etc. are safe and round-trip checked).
- **It reads the last *saved* version of the file** — unsaved edits still
  open in Excel are silently not picked up. Save (and ideally close) the
  workbook first.
- **SGX `Type` column (stocks / reits / businesstrusts) is not in the
  spreadsheet** — it's worked out per stock after fetching, in
  `extract_index_stocks` via `classify_sgx_security()`: Yahoo `industry`
  starting with `"REIT"` → `reits` (true for every SGX REIT, and no non-REIT,
  as of when this was written); ticker in the hand-maintained
  `SGX_BUSINESS_TRUSTS` set → `businesstrusts` (business trusts report
  ordinary industries like "Telecom Services", so they can't be detected);
  otherwise `stocks`; `N/A` if the industry couldn't be fetched. **When adding
  a business trust to the spreadsheet, also add its ticker to
  `SGX_BUSINESS_TRUSTS`** — as a reminder, a non-REIT whose Yahoo name
  contains the word "Trust"/"Tr" but isn't in the set logs a warning.
  Market-cap mode keeps the type SGX's own API supplies instead.

## Cross-platform notes

This project was authored on Windows and has been ported to other machines.
Scripts that previously had **hardcoded absolute `C:\Users\...` paths** now
resolve paths relative to the script's own location
(`os.path.dirname(os.path.abspath(__file__))`) instead —
`update_notebook.py` and `sync_stock_lists.py` (reading/writing the `.py`
files, the notebook and `stock list.xlsx`).

If you add new file paths anywhere in this project, use the same pattern
(relative to cwd, or relative to `__file__`) rather than an absolute path —
there is no other OS-specific code left in the project (no `subprocess`,
`winreg`, or platform-conditional logic).

A related gotcha: `__file__` is only defined when a `.py` file is actually
executed as a script — it doesn't exist inside a notebook cell (Jupyter or
Colab both run cells with no backing file), and `update_notebook.py` copies
both pipelines' source into the notebook verbatim. An earlier
`__file__`-based path inside `SGX_Stock_Analysis.py` raised `NameError` as
soon as the notebook ran that cell (that code is gone now, along with the
runtime Excel loader it belonged to). Keep this in mind for any future
`__file__` usage in either pipeline script — the sync-only scripts
(`update_notebook.py`, `sync_stock_lists.py`) never run in the notebook, so
they're fine.

## Notebook sync workflow

After changing either `.py` script or `stock list.xlsx`:

```
python update_notebook.py
```

This first runs `sync_stock_lists()` (see "Stock lists" above), then reads
both scripts' real source (from their first module-level constant —
`FALLBACK_STI_30` / `FALLBACK_NIFTY_50` — through `def main():`, exclusive;
the generated stock-list blocks come after those constants, so they're
included) into the notebook's two code cells, appends a short "run the
pipeline and download" block, and refreshes the markdown description cells to
match — including the stock counts, read from the freshly generated lists
rather than hardcoded. It does **not** re-run the pipeline itself — it only
regenerates the notebook's static cell contents.

The notebook has a fixed 7-cell layout the script addresses **by index**
(intro=0, deps/imports md/code=1/2, SGX md/code=3/4, NSE md/code=5/6). An
earlier revision of this project inserted a Drive-mount cell pair at index 3
(to pull `stock list.xlsx` from Google Drive at runtime) - that's been
removed since the stock lists are hardcoded, so neither script needs the file
at runtime. `update_notebook.py` still checks for and strips that old cell
pair if it finds one (idempotently, via a "Mount Google Drive" string match at
index 3), so an older copy of the notebook self-heals on the next sync.
Cell 2 (shared `pip install` + imports + logging/warnings setup, used by
both SGX and NSE) is hand-composed in
`update_notebook.py` rather than sliced from either script - each script has
its own copy of this same setup (both must stay independently runnable), so
the notebook consolidates it once rather than importing everything twice in
the same kernel. It's still templated in `update_notebook.py` and rewritten
on every sync, same as every other cell - **not** meant to be hand-edited in
the `.ipynb` either, even though it isn't literally sliced from a `.py` file.
If you ever restructure the notebook's
cells by hand or add a new step, update these indices in
`update_notebook.py` to match, or the wrong cell's source gets overwritten.

Sanity-check after running it:

```
python -m py_compile NSE_Stock_Analysis.py SGX_Stock_Analysis.py sync_stock_lists.py update_notebook.py
```

(and optionally `compile()` each notebook code cell's joined source, skipping
the `!pip` line, since a bad slice boundary would be a silent notebook bug,
not a `.py` one). To test a sync without touching the real notebook, copy the
project files into a scratch directory and run `update_notebook.py` there —
it resolves every path relative to its own location.

## Sharing workflow (git + Google Drive)

This project is a git repository with a private GitHub remote (`origin`).
Since the repo is private, viewers without GitHub collaborator access can't
use Colab's "Open from GitHub" flow (it authenticates as the viewer, and a
private repo will simply refuse anyone without access). Instead, a copy of
the notebook lives on Google Drive and is shared from there — kept in sync
with the actual project files via `rclone` (a Drive-authenticated CLI), not
Google Drive's desktop sync client.

Why not the Drive desktop client: this project's working directory lives in
WSL2's native Linux filesystem, which Windows (and the Drive desktop app)
can only reach through a virtual network path (`\\wsl.localhost\...`) — the
desktop client's folder-watching sync doesn't support that path reliably.
`rclone` sidesteps this entirely by talking to the Drive API directly from
Linux, so no Drive desktop install is needed.

`update_and_share.sh` re-uploads the notebook to the *same* Drive file path
each time (`rclone copyto`), overwriting its content in place rather than
creating a new file — this keeps the same Drive file ID, so a share link
already handed to someone keeps working without re-sharing after updates.
It does the same for `stock list.xlsx`, uploaded to the same Drive folder
(`Colab Notebooks/Project_Stock_Analysis/`) right after the notebook.

**Nothing reads `stock list.xlsx` from Drive** — the notebook carries the
stock lists baked into its code (as of the last sync), so a Colab session
started from the Drive-hosted notebook never needs the file. It's uploaded on
every `update_and_share.sh` run purely as a convenience backup copy alongside
the notebook. Close the workbook in Excel before running the script, so the
sync reads the saved version and `rclone` doesn't upload a half-written file
(Excel's `~$stock list.xlsx` lock file is a sign it's still open).

## Known open items

- The historical P/E/P/S/P/B approximation (see above) hasn't been
  spot-checked against a paid data vendor — flagged to the user as a "verify
  later" item, not yet resolved. (Split/bonus-issue adjustment *has* been
  checked across every stock in both lists — see the EPS/share-count
  cross-check above.)
