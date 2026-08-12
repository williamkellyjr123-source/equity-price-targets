# Equity Price Targets — automated

A self-refreshing analyst price-target tracker. Every weekday after US market close,
GitHub Actions fetches fresh data (prices, analyst targets, dividends, earnings dates,
P/E) for every symbol in `tickers.txt` via Yahoo Finance and rebuilds
`Equity_Price_Targets.xlsx` — no computer needs to be on, and it costs nothing.

## One-time setup (about 5 minutes)

1. Create a free account at github.com if you don't have one.
2. Click **New repository** → name it `equity-price-targets` → set it to **Private**
   (or Public) → Create.
3. On the empty repo page choose **uploading an existing file** and drag in everything
   from this folder — including the `.github` folder (if your browser won't upload the
   folder, create the file manually: **Add file → Create new file**, type
   `.github/workflows/refresh.yml` as the name, and paste the contents). Commit.
4. Go to the **Actions** tab → enable workflows if prompted → open
   **Refresh Equity Price Targets** → **Run workflow**. In ~2 minutes the first
   refresh commits a fully populated `Equity_Price_Targets.xlsx`.
5. Tell Claude the repo URL so your scheduled delivery task can pick the file up.

That's it. From now on it refreshes itself every trading day at 22:35 UTC.

## Adding or removing stocks

Edit `tickers.txt` (one Yahoo-style symbol per line — `BRK-B`, not `BRK.B`) right in
the GitHub web editor, or tell Claude. The next refresh picks it up automatically.

## Running on your own computer instead (optional)

The same scripts work locally:

    pip install -r requirements.txt
    python fetch_data.py
    python build_workbook.py data.json Equity_Price_Targets.xlsx

Schedule those two commands with Task Scheduler (Windows) or cron/launchd (Mac) if
you'd rather not use GitHub.

## Notes

- Data comes from Yahoo Finance via the `yfinance` library — free but unofficial;
  the fetch retries each symbol up to 4 times and fails the run (keeping yesterday's
  data) if more than 20% of symbols error out.
- Ex-dividend dates older than 12 months are treated as "no current dividend" and
  blanked, so suspended dividends (e.g. INTC) never show stale dates.
- All derived columns in the spreadsheet (upside %, dividend yield, P/E) are live
  Excel formulas, so the sheet stays internally consistent.
