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
  independently runnable pipelines. **These are the source of truth.**
- `Stock_Market_Analysis_SGX_NSE.ipynb` — Colab notebook. Generated/kept in
  sync from the two scripts above; never hand-edit its code cells directly,
  see "Notebook sync" below.
- `create_notebook.py` — one-off script that originally built the notebook
  from a hand-condensed rewrite of the two scripts. Superseded by
  `update_notebook.py`; keep for reference but don't run it (it would
  regenerate the notebook from a stale, drifted copy of the logic).
- `update_notebook.py` — **the real sync script.** Slices the current,
  literal contents of both `.py` files into the notebook's two code cells
  (plus a short exec block calling `main()`'s pipeline) and refreshes the
  markdown cells to match. Run this after *any* change to either script.
- `stock list.xlsx` — shared by both scripts, one worksheet each (`SGX`, `NSE`).
  Each row is Index/Inclusion-reason + Company + an optional Ticker column. See
  "Shared stock list.xlsx" below for how the two scripts use it differently.
- `requirements.txt` — `pandas`, `yfinance`, `requests`, `beautifulsoup4`,
  `openpyxl`. All pure-Python / cross-platform, nothing OS-specific.
- `.gitignore` — excludes generated pipeline outputs (`*_Stock_Analysis_*.xlsx`,
  `*.csv`, `*.json`), Python/Jupyter caches, and Windows/WSL
  `*:Zone.Identifier` metadata files. Generated outputs are reproducible by
  re-running the scripts, so they're deliberately not tracked.
- `update_and_share.sh` — one-command "save progress" wrapper: runs
  `update_notebook.py`, compile-checks both scripts, `git add`/`commit`
  (skipped if nothing changed) /`push`, then refreshes the shared Google
  Drive copies of both the notebook and `stock list.xlsx` via `rclone` (see
  "Sharing workflow" below). Run this instead of the individual steps
  whenever a script or the stock list changes.

## Running

```
python NSE_Stock_Analysis.py [--mode fast|robust|full] [--top 80] [--workers 10]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
python SGX_Stock_Analysis.py [--mode index|market-cap] [--index-file "stock list.xlsx"]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
```

NSE also accepts `--stock-list-file "stock list.xlsx"` (same file, `NSE` worksheet,
purely additive - see "Shared stock list.xlsx" below).

Default output is `<NSE|SGX>_Stock_Analysis_<today>.xlsx`. Both scripts also
accept `.csv`/`.json` output, in which case the fundamentals/chart sheets
below don't apply (Excel-only features).

## Excel workbook design (and *why* it's built this way)

Each workbook has (up to) four sheets, in this order, with hyperlinks
threading them together:

1. **Summary** — one row per company, ranked by market cap. Deliberately does
   *not* include `SharesOutstanding` (see below) even though the underlying
   data is fetched.
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
destination block has a "↑ Back to Summary" link back.

Several non-obvious decisions are baked into this, worth knowing before
changing it:

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

- **Historical P/E, P/S, P/B are approximations.** Yahoo Finance doesn't
  provide historical valuation ratios directly. Each fiscal year's reported
  diluted EPS / average diluted shares / stockholders' equity (from
  `yfinance`'s `income_stmt`/`balance_sheet`) is combined with the closing
  share price on (or immediately before) that fiscal year-end date. Years
  requested default to 5 (`--financials-years`), capped by however many
  Yahoo actually has. **In practice this is usually 4, not 5** - Yahoo's free
  API currently only serves ~4 years of annual `income_stmt`/`balance_sheet`
  data regardless of what's requested (confirmed via multiple reports on
  yfinance's own GitHub, including from its maintainer); the 5th year only
  shows up on Yahoo Finance's paid "Plus" tier, which `yfinance` has no way
  to authenticate as. Not fixable from this project's side - if 5 full years
  is a hard requirement, it would mean switching to a different (likely
  paid) financial data provider, not a parameter tweak.

- **SGX fundamentals are in each company's native reporting currency**
  (labeled per block, e.g. "Reported in USD"), *not* SGD-converted like the
  Summary sheet's Market Cap — converting a multi-year series properly needs
  historical FX rates per year, not today's spot rate, so this was left
  unconverted rather than silently approximated.

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
  approximation above).

## Shared stock list.xlsx (SGX index + both scripts' "Personal" additions)

One workbook, two worksheets, read very differently by the two scripts:

- **`SGX` worksheet** is `load_sgx_index_stock_list()`'s **primary source of
  SGX's entire default-mode stock universe** - all 89 STI/S Cap/L & M Cap
  index constituents, plus any rows tagged `Personal` in the Index column
  (currently 3: Micro Mechanics, Jardine C&C, Rex International). If the file
  is missing entirely, falls back to an embedded copy of the same data
  (`DEFAULT_SGX_INDEX_STOCKS`).
- **`NSE` worksheet** is **purely additive** for `extract_top80_with_nifty50()`
  - NSE already has a full stock universe (NIFTY 50 + Top N market-cap
    ranking) without this file. Rows here (currently 8 personal picks) are
  merged into the fetch pass and *guaranteed* to appear in the final output,
  same mechanism already used to guarantee NIFTY 50 constituents that fall
  outside the Top N: tagged `Inclusion_Reason = "Personal"`, or `"<existing
  reason> + Personal"` if the stock already qualified another way (e.g. also
  Top 80). If the file or `NSE` worksheet is missing, personal additions are
  silently skipped - no embedded fallback here, since NSE doesn't need one.

Each row is `Index`/`Inclusion reason` + `Company` + an optional **`Ticker`**
column (added after the two loaders originally only matched by Company name):

- **A non-blank Ticker cell (e.g. `5DD.SI`, `ACC.NS`) is used directly** -
  the real yfinance symbol, no name lookup involved. This is the preferred
  way to add a new stock now: one spreadsheet edit, no code change.
- **A blank Ticker cell falls back to name-matching** against
  `SGX_INDEX_STOCK_MAPPINGS` / `NSE_PERSONAL_STOCK_MAPPINGS` (hardcoded
  name → code/symbol dicts in each script) - exact match, then (SGX only) a
  fuzzy match, then (SGX only) a crude last-resort guess from the company
  name's first 4 letters, which is often wrong for anything not already in
  the dict. All 100 rows currently in the file have their Ticker filled in,
  so this fallback path only matters for a *newly added* row that skips the
  Ticker column.
- **SGX worksheet name matters**: the loader checks `"SGX"` first (the
  current name), falls back to the legacy `"Sheet2"` name for older file
  copies, then `wb.active` as a last resort.
- Adding a stock to the `NSE` worksheet with a name not in
  `NSE_PERSONAL_STOCK_MAPPINGS` **and** no Ticker filled in gets skipped
  with a logged warning, not a crude guess (unlike SGX's fallback) - NSE
  ticker codes aren't reliably derivable from a company name, so this
  script doesn't try.

## Cross-platform notes

This project was authored on Windows and has been ported to other machines.
Two scripts previously had **hardcoded absolute `C:\Users\...` paths** that
would silently break outside that one machine — both are now resolved
relative to the script's own location (`os.path.dirname(os.path.abspath(__file__))`)
instead:
- `update_notebook.py` (reading/writing the `.py` files and the notebook)
- `SGX_Stock_Analysis.py`'s `load_sgx_index_stock_list()` fallback path

If you add new file paths anywhere in this project, use the same pattern
(relative to cwd, or relative to `__file__`) rather than an absolute path —
there is no other OS-specific code left in the project (no `subprocess`,
`winreg`, or platform-conditional logic).

A related gotcha: `__file__` is only defined when a `.py` file is actually
executed as a script — it doesn't exist inside a notebook cell (Jupyter or
Colab both run cells with no backing file), and `update_notebook.py` copies
this same source into the notebook verbatim. `load_sgx_index_stock_list()`'s
`__file__`-based candidate path used to raise `NameError` as soon as the
notebook ran that cell. Fixed by wrapping just that one candidate in
`try/except NameError` so it's silently skipped in a notebook context (the
`os.getcwd()`/bare-relative-path candidates already in the same list cover
Colab). Keep this in mind for any future `__file__` usage that might end up
in a notebook cell via the sync script.

## Notebook sync workflow

After changing either `.py` script:

```
python update_notebook.py
```

This reads both scripts' real source (from their first module-level
constant — `FALLBACK_STI_30` / `FALLBACK_NIFTY_50` — through `def main():`,
exclusive) into the notebook's two code cells, appends a short "run the
pipeline and download" block, and refreshes the markdown description cells
to match. It does **not** re-run the pipeline itself — it only regenerates
the notebook's static cell contents.

The notebook has a fixed 9-cell layout the script addresses **by index**
(intro=0, deps md/code=1/2, Drive-mount md/code=3/4, SGX md/code=5/6, NSE
md/code=7/8) - the Drive-mount cell pair is inserted once, idempotently
(checked via a "Mount Google Drive" string match at index 3), so re-running
this script never duplicates it. If you ever restructure the notebook's
cells by hand or add a new step, update these indices in
`update_notebook.py` to match, or the wrong cell's source gets overwritten.

Sanity-check after running it:

```
python -m py_compile NSE_Stock_Analysis.py SGX_Stock_Analysis.py
```

(and optionally `compile()` each notebook code cell's joined source, since a
bad slice boundary would be a silent notebook bug, not a `.py` one).

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

**Why `stock list.xlsx` also needs to be on Drive, not just GitHub:** a
Colab session started from the Drive-hosted notebook has no link back to
this GitHub repo or your local disk — it's a fresh, empty runtime. The
notebook's **Step 2** cell mounts Drive and copies `stock list.xlsx` from
that same Drive folder into the session's working directory before SGX/NSE
run, so personal stock additions don't need re-uploading by hand each
session. That cell fails soft (prints a message, doesn't error) if it's not
running in Colab or the file isn't in Drive yet — SGX falls back to its
embedded list either way; NSE just skips personal additions for that run.
The Drive path is hardcoded in that cell (`update_notebook.py`'s
`drive_step_code`) to match wherever `update_and_share.sh` uploads it -
keep both in sync if that folder path ever changes.

## Known open item

The historical P/E/P/S/P/B approximation (see above) hasn't been spot-checked
against a paid data vendor — flagged to the user as a "verify later" item, not
yet resolved.
