#!/usr/bin/env python3
"""Build one TradingView chart page per watchlist tab from the published Google Sheet.

Reads every tab whose column E header is "Market Cap" (the equity tabs), and writes
<tab>.html plus an index.html into the output folder.

Usage:
    python build_pages.py --out ./watchlist-pages
    python build_pages.py --source watchlist.xlsx --out ./watchlist-pages   # local file instead
"""
import argparse
import html
import io
import json
import urllib.request
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

SHEET_XLSX = ("https://docs.google.com/spreadsheets/d/e/2PACX-1vQWZEUeoRauMtLXrP6HqdZOXiH55q-"
              "H84BnaUqmJcf0Si5JzIW6wdnRONeCyNyF9lWWgcuYCcTiz69T/pub?output=xlsx")
CAD_PREFIXES = {"TSE", "CVE", "NEO", "CNSX"}  # GoogleFinance prefixes quoted in CAD

PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__ – weekly, 2 years</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Condensed:wght@400;600&display=swap" rel="stylesheet">
<style>
  :root {
    color-scheme: dark;
    --bg: #000000; --card: #0d0d0d; --ink: #f2f2f2; --muted: #8c8c8c;
    --rule: #262626; --accent: #ffffff;
    --chart-min: 700px;   /* minimum chart width: change this to adjust columns */
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--ink);
         font-family: "IBM Plex Sans Condensed", "Arial Narrow", system-ui, sans-serif; }
  a { color: var(--muted); }
  header.page { padding: 28px 24px 8px; }
  header.page h1 { margin: 0; font-size: 1.6rem; font-weight: 600; }
  header.page p { margin: 4px 0 0; color: var(--muted); font-size: 0.95rem; }
  .grid { display: grid; gap: 20px; padding: 16px 24px 24px;
          grid-template-columns: repeat(auto-fill, minmax(min(var(--chart-min), 100%), 1fr)); }
  .card { background: var(--card); border: 1px solid var(--rule); border-top: 2px solid var(--accent);
          border-radius: 6px; overflow: hidden; }
  .card .head { display: flex; align-items: baseline; gap: 12px; padding: 12px 16px;
                border-bottom: 1px solid var(--rule); }
  .ticker { font-weight: 600; font-size: 1.25rem; color: var(--accent); }
  .name { font-size: 1rem; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .cap { font-size: 0.95rem; color: var(--muted); white-space: nowrap; }
  .chart { height: 420px; }
  .chart .tradingview-widget-container, .chart .tradingview-widget-container__widget { height: 100%; width: 100%; }
  nav.sectors { display: flex; flex-wrap: wrap; gap: 6px; padding: 16px 24px 0; }
  nav.sectors a { padding: 6px 12px; border: 1px solid var(--rule); border-radius: 6px;
                  background: var(--card); color: var(--ink); text-decoration: none; font-size: 0.9rem; }
  nav.sectors a:hover { border-color: var(--accent); }
  nav.sectors a[aria-current="page"] { background: var(--accent); color: var(--bg); border-color: var(--accent); }
  footer { padding: 0 24px 32px; color: var(--muted); font-size: 0.85rem; }
</style>
</head>
<body>
<nav class="sectors">
__NAV__
</nav>
<header class="page">
  <h1>__TITLE__</h1>
  <p>Weekly candles, trailing two years</p>
</header>
<main class="grid" id="grid"></main>
<footer>Charts are live from TradingView. Market caps from Google Finance as of __ASOF__, in each listing’s currency; sorted using USD/CAD __FX__.</footer>
<script>
  const stocks = __DATA__;
  const fmtCap = s => {
    if (s.cap == null) return "n/a";
    const v = s.cap / 1e9, cur = s.cur === "CAD" ? "C$" : "US$";
    return v >= 1 ? cur + v.toFixed(1) + "B" : cur + Math.round(v * 1000) + "M";
  };
  const esc = x => String(x).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
  const grid = document.getElementById("grid");
  stocks.forEach(s => {
    const card = document.createElement("section");
    card.className = "card";
    card.innerHTML = `
      <div class="head">
        <span class="ticker">${esc(s.t)}</span>
        <span class="name" title="${esc(s.name)}">${esc(s.name)}</span>
        <span class="cap">Mkt cap ≈ ${fmtCap(s)}</span>
      </div>
      <div class="chart"><div class="tradingview-widget-container">
        <div class="tradingview-widget-container__widget"></div></div></div>`;
    grid.appendChild(card);
    const script = document.createElement("script");
    script.src = "https://s3.tradingview.com/external-embedding/embed-widget-advanced-chart.js";
    script.async = true;
    script.innerHTML = JSON.stringify({
      autosize: true, symbol: s.tv, interval: "W", range: "24M", timezone: "America/Toronto",
      theme: "dark", backgroundColor: "#0d0d0d", gridColor: "rgba(255, 255, 255, 0.06)",
      style: "1", locale: "en", hide_top_toolbar: true, hide_legend: false, hide_side_toolbar: true,
      allow_symbol_change: false, save_image: false, withdateranges: false, details: false,
      calendar: false, support_host: "https://www.tradingview.com"
    });
    card.querySelector(".tradingview-widget-container").appendChild(script);
  });
</script>
</body>
</html>
"""

INDEX = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Watchlists</title>
<style>
  :root { color-scheme: dark; }
  body { margin: 0; padding: 28px 24px; background: #000; color: #f2f2f2;
         font-family: "IBM Plex Sans Condensed", "Arial Narrow", system-ui, sans-serif; }
  h1 { font-size: 1.6rem; font-weight: 600; margin: 0 0 16px; }
  ul { list-style: none; padding: 0; display: grid; gap: 8px;
       grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); }
  a { display: block; padding: 10px 14px; border: 1px solid #262626; border-radius: 6px;
      background: #0d0d0d; color: #fff; text-decoration: none; }
  a:hover { border-color: #fff; }
  span { color: #8c8c8c; }
</style></head><body>
<h1>Watchlists</h1>
<ul>
__LINKS__
</ul>
</body></html>
"""


def load(source):
    if source.startswith("http"):
        with urllib.request.urlopen(source) as r:
            data = io.BytesIO(r.read())
        return load_workbook(data, data_only=True)
    return load_workbook(source, data_only=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=SHEET_XLSX)
    ap.add_argument("--out", default="watchlist-pages")
    ap.add_argument("--usdcad", type=float, default=1.38, help="FX rate used only for sorting")
    args = ap.parse_args()

    wb = load(args.source)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    asof = date.today().strftime("%b %d, %Y")
    pages = []

    for ws in wb:
        if str(ws.cell(1, 5).value or "").strip().lower() != "market cap":
            continue
        rows = []
        for t, name, tv, gf, cap, *_ in ws.iter_rows(min_row=2, values_only=True):
            if not tv:
                continue
            cur = "CAD" if str(gf or "").split(":")[0] in CAD_PREFIXES else "USD"
            cap = cap if isinstance(cap, (int, float)) and cap > 0 else None
            rows.append({"t": t or str(tv).split(":")[-1], "name": name or "", "tv": tv,
                         "cap": cap, "cur": cur})
        if not rows:
            continue

        def usd(r):
            if r["cap"] is None:
                return -1
            return r["cap"] / args.usdcad if r["cur"] == "CAD" else r["cap"]

        rows.sort(key=usd, reverse=True)
        pages.append((f"{ws.title}.html", ws.title.replace("_", " ").title(), rows))

    def nav(current):
        items = ['<a href="index.html"' + (' aria-current="page"' if current == "index.html" else "") + '>Home</a>']
        for f, t, _ in pages:
            cur = ' aria-current="page"' if f == current else ""
            items.append(f'<a href="{f}"{cur}>{html.escape(t)}</a>')
        return "\n".join(items)

    for fname, title, rows in pages:
        page = (PAGE.replace("__TITLE__", html.escape(title))
                    .replace("__NAV__", nav(fname))
                    .replace("__ASOF__", asof)
                    .replace("__FX__", f"{args.usdcad:.2f}")
                    .replace("__DATA__", json.dumps(rows, ensure_ascii=False, indent=1)))
        (out / fname).write_text(page, encoding="utf-8")
        missing = sum(r["cap"] is None for r in rows)
        print(f"{fname:20} {len(rows):3} stocks  ({missing} without market cap)")

    links = "\n".join(f'  <li><a href="{f}">{html.escape(t)} <span>· {len(r)}</span></a></li>'
                      for f, t, r in pages)
    (out / "index.html").write_text(INDEX.replace("__LINKS__", links), encoding="utf-8")
    print(f"\nWrote {len(pages)} pages + index.html to {out.resolve()}")


if __name__ == "__main__":
    main()
