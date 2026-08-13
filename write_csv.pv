"""Convert data.json to data.csv so Google Sheets can IMPORTDATA it."""
import csv, json, pathlib

HERE = pathlib.Path(__file__).parent
d = json.load(open(HERE / "data.json"))
cols = ["ticker", "name", "industry", "country", "price", "target_avg", "target_min",
        "target_max", "analysts", "market_cap", "div_rate", "ex_div", "earnings",
        "eps_ttm", "as_of"]
with open(HERE / "data.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(cols)
    for r in sorted(d["rows"], key=lambda x: -(x.get("market_cap") or 0)):
        w.writerow([r.get(c) if r.get(c) is not None else "" for c in cols])
print(f"wrote data.csv ({len(d['rows'])} rows)")
