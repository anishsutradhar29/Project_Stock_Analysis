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
- `SGX Index stock list.xlsx` — source list of the 89 SGX index constituents
  (STI / S Cap / L&M Cap) that `SGX_Stock_Analysis.py`'s default mode reads.
  Falls back to an embedded 89-stock dict if this file is missing.
- `requirements.txt` — `pandas`, `yfinance`, `requests`, `beautifulsoup4`,
  `openpyxl`. All pure-Python / cross-platform, nothing OS-specific.
- `.gitignore` — excludes generated pipeline outputs (`*_Stock_Analysis_*.xlsx`,
  `*.csv`, `*.json`), Python/Jupyter caches, and Windows/WSL
  `*:Zone.Identifier` metadata files. Generated outputs are reproducible by
  re-running the scripts, so they're deliberately not tracked.
- `update_and_share.sh` — one-command "save progress" wrapper: runs
  `update_notebook.py`, compile-checks both scripts, `git add`/`commit`
  (skipped if nothing changed) /`push`, then refreshes a shared Google Drive
  copy of the notebook via `rclone` (see "Sharing workflow" below). Run this
  instead of the individual steps whenever a script changes.

## Running

```
python NSE_Stock_Analysis.py [--mode fast|robust|full] [--top 80] [--workers 10]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
python SGX_Stock_Analysis.py [--mode index|market-cap] [--index-file "SGX Index stock list.xlsx"]
                              [--no-financials] [--financials-years 5] [--output FILE.xlsx]
```

Default output is `<NSE|SGX>_Stock_Analysis_<today>.xlsx`. Both scripts also
accept `.csv`/`.json` output, in which case the fundamentals/chart sheets
below don't apply (Excel-only features).

## Excel workbook design (and *why* it's built this way)

Each workbook has (up to) four sheets, in this order, with hyperlinks
threading them together:

1. **Summary** — one row per company, ranked by market cap.
2. **52W Daily History** — horizontal layout: each company gets a 5-column
   block (`Close`, `% Change`, `Market Cap`, `Shares Out.`, `Volume`) with a
   2-tier merged header, rows are trading days.
3. **Financials Detail** — a stacked block per company: up to 5 fiscal years
   (oldest → newest, top → bottom) of Revenue, Net Income, EPS, P/E, P/S,
   P/B, Gross/Operating Margin, Debt/Equity.
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
  Yahoo actually has (usually 5, sometimes fewer for smaller/newer tickers).

- **SGX fundamentals are in each company's native reporting currency**
  (labeled per block, e.g. "Reported in USD"), *not* SGD-converted like the
  Summary sheet's Market Cap — converting a multi-year series properly needs
  historical FX rates per year, not today's spot rate, so this was left
  unconverted rather than silently approximated.

- `SharesOutstanding` used to be silently `None` for every row (never
  actually captured from `yfinance`'s `.info`, despite `save_results`
  already having formatting logic ready for it) — this made the 52W History
  sheet's `Market Cap`/`Shares Out.` columns permanently blank. Fixed in
  both scripts' `fetch_latest_ticker_data`.

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
the notebook's static cell contents. Sanity-check after running it:

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

## Known open item

The historical P/E/P/S/P/B approximation (see above) hasn't been spot-checked
against a paid data vendor — flagged to the user as a "verify later" item, not
yet resolved.
