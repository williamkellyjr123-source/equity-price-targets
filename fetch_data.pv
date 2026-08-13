"""Fetch equity data via yfinance for every symbol in tickers.txt -> data.json."""
import json, time, datetime, pathlib
from concurrent.futures import ThreadPoolExecutor
import yfinance as yf

HERE = pathlib.Path(__file__).parent
TICKERS = [s.strip().upper() for s in (HERE / "tickers.txt").read_text().splitlines()
           if s.strip() and not s.startswith("#")]

def epoch_to_date(v):
    try:
        return datetime.datetime.fromtimestamp(int(v)).strftime("%Y-%m-%d")
    except Exception:
        return None

def fetch(sym, cutoff, as_of):
    for attempt in range(4):
        try:
            info = yf.Ticker(sym).info
            price = info.get("currentPrice") or info.get("regularMarketPrice")
            if not price:
                raise ValueError("no price returned")
            earnings = epoch_to_date(info.get("earningsTimestampStart") or info.get("earningsTimestamp"))
            exdiv = epoch_to_date(info.get("exDividendDate"))
            div_rate = info.get("dividendRate")
            # a stale ex-div date (>12 months old) means no current dividend
            if exdiv and datetime.datetime.strptime(exdiv, "%Y-%m-%d") < cutoff:
                exdiv, div_rate = None, None
            return {"ticker": sym.replace("-", "."), "name": info.get("shortName") or sym,
                    "industry": (info.get("industry") or "").title() or None,
                    "country": info.get("country"),
                    "price": round(float(price), 2),
                    "target_avg": info.get("targetMeanPrice"),
                    "target_min": info.get("targetLowPrice"),
                    "target_max": info.get("targetHighPrice"),
                    "analysts": info.get("numberOfAnalystOpinions"),
                    "market_cap": info.get("marketCap"),
                    "div_rate": div_rate, "ex_div": exdiv, "earnings": earnings,
                    "eps_ttm": info.get("trailingEps"), "as_of": as_of}
        except Exception as e:
            if attempt == 3:
                return {"ticker": sym, "error": str(e)[:120]}
            time.sleep(2 * (attempt + 1))

def main():
    now = datetime.datetime.now()
    cutoff = now - datetime.timedelta(days=365)
    as_of = now.strftime("%Y-%m-%d")
    # gentler concurrency for large universes: fewer workers, Yahoo tolerates this well
    with ThreadPoolExecutor(max_workers=4) as ex:
        results = list(ex.map(lambda s: fetch(s, cutoff, as_of), TICKERS))
    rows = [r for r in results if r and "error" not in r]
    failed = [[r["ticker"], r["error"]] for r in results if r and "error" in r]
    json.dump({"as_of": now.strftime("%Y-%m-%d %H:%M"), "rows": rows, "failed": failed},
              open(HERE / "data.json", "w"), indent=1)
    print(f"fetched {len(rows)}/{len(TICKERS)}; {len(failed)} failed")
    if failed:
        print("failed:", failed[:40], "..." if len(failed) > 40 else "")
    if len(rows) < len(TICKERS) * 0.5:
        raise SystemExit("over half the symbols failed (likely rate limiting) — "
                         "not committing a gutted dataset; re-run the workflow")

if __name__ == "__main__":
    main()
