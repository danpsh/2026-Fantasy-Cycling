#!/usr/bin/env python3
"""
Scrape Tour de France results from ProCyclingStats -> tdf-results.xlsx.

Parsed directly with selectolax (procyclingstats 0.2.8 can't read current PCS).
Each classification has its own page; on each we take the LARGEST results table
(the full standings, ignoring small "today" widgets). Riders are identified by
their URL slug and mapped to canonical roster spelling from tdf-startlist.js.

    pip install procyclingstats openpyxl   # (brings selectolax)
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py
"""
import os
import re
import json
import time
import unicodedata
import urllib.request
import urllib.parse

from selectolax.parser import HTMLParser
from openpyxl import Workbook

# curl_cffi impersonates a real Chrome TLS/HTTP2 fingerprint, which gets past
# ProCyclingStats' Cloudflare block (plain urllib gets 403 from datacenter IPs).
try:
    from curl_cffi import requests as cffi_requests
except Exception:
    cffi_requests = None

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# ProCyclingStats sits behind Cloudflare, which 403s datacenter IPs (i.e. GitHub
# Actions runners). When SCRAPER_API_KEY is set we route each request through
# ScraperAPI's residential IP pool to get a clean IP; when it's absent we hit PCS
# directly (works from a home/residential IP for local runs). Free tier is ~1,000
# requests/month; a full 21-stage scrape is ~21 requests.
SCRAPER_API_KEY = os.environ.get("SCRAPER_API_KEY", "").strip()

SCHED_2026 = {1: "2026-07-04", 2: "2026-07-05", 3: "2026-07-06", 4: "2026-07-07",
              5: "2026-07-08", 6: "2026-07-09", 7: "2026-07-10", 8: "2026-07-11",
              9: "2026-07-12", 10: "2026-07-14", 11: "2026-07-15", 12: "2026-07-16",
              13: "2026-07-17", 14: "2026-07-18", 15: "2026-07-19", 16: "2026-07-21",
              17: "2026-07-22", 18: "2026-07-23", 19: "2026-07-24", 20: "2026-07-25",
              21: "2026-07-26"}
MONTHS = {m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1)}

print(f"=== ENV === YEAR={YEAR} OUT={OUT} RACE={RACE}")
print("=== FETCH === via ScraperAPI proxy" if SCRAPER_API_KEY else "=== FETCH === direct (no proxy key)")


def pkey(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def load_canon():
    # tdf-startlist.js is a JS object literal (unquoted keys), not JSON, so pull
    # rider names with a regex rather than json.loads.
    try:
        with open("tdf-startlist.js", "r", encoding="utf-8") as f:
            t = f.read()
        names = re.findall(r'(?:"name"|name)\s*:\s*"([^"]+)"', t)
        canon = {pkey(n): n for n in names}
        if not canon:
            raise ValueError("no rider names found")
        print(f"loaded {len(canon)} canonical names from tdf-startlist.js")
        return canon
    except Exception as e:
        print("tdf-startlist.js NOT loaded (names will be slug-cased):", e)
        return {}


CANON = load_canon()


def fetch(url, tries=3):
    full = "https://www.procyclingstats.com/" + url
    headers = {
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.procyclingstats.com/",
    }
    # When a proxy key is present, hand PCS's URL to ScraperAPI and fetch THAT
    # instead. ScraperAPI supplies a residential IP + solves Cloudflare, then
    # returns the target page's HTML. render=false (default) is enough for PCS's
    # server-rendered tables; country_code=us keeps the IP pool consistent.
    if SCRAPER_API_KEY:
        target = "https://api.scraperapi.com/?" + urllib.parse.urlencode(
            {"api_key": SCRAPER_API_KEY, "url": full, "country_code": "us"})
        req_headers = {"User-Agent": UA}
    else:
        target = full
        req_headers = headers

    last = None
    for i in range(tries):
        try:
            if cffi_requests is not None:
                r = cffi_requests.get(target, headers=req_headers, impersonate="chrome", timeout=60)
                if r.status_code == 200:
                    return r.text
                last = f"HTTP {r.status_code}"
            else:
                req = urllib.request.Request(target, headers=req_headers)
                with urllib.request.urlopen(req, timeout=60) as resp:
                    return resp.read().decode("utf-8", "replace")
        except Exception as e:
            last = str(e)
        time.sleep(2 * (i + 1))
    raise RuntimeError(f"fetch failed for {url}: {last}")


def name_from_href(href):
    m = re.search(r"rider/([^/?#\"]+)", href or "")
    if not m:
        return ""
    slug = m.group(1).replace("-", " ").strip()
    return CANON.get(pkey(slug)) or " ".join(w.capitalize() for w in slug.split())


def table_names(table):
    body = table.css_first("tbody") or table
    out, seen = [], set()
    for tr in body.css("tr"):
        a = tr.css_first('a[href*="rider/"]')
        if not a:
            continue
        nm = name_from_href(a.attributes.get("href", ""))
        if nm and nm not in seen:
            seen.add(nm)
            out.append(nm)
    return out


def classify(html, debug=False):
    """Return {stage,gc,points,kom,youth} name-lists.

    PCS puts all five standings on one stage page, in DOM order
    stage, GC, points, KOM, youth, interleaved with small daily sub-tables.
    The five cumulative standings are the first five results tables with
    enough riders (>=20); the daily sub-results (n<20) are skipped.
    """
    kept = []
    for i, t in enumerate(HTMLParser(html).css("table.results")):
        names = table_names(t)
        if debug:
            print(f"    [{i}] n={len(names)} top3={names[:3]}")
        if len(names) >= 20:
            kept.append(names)
    order = ["stage", "gc", "points", "kom", "youth"]
    return {k: (kept[i] if i < len(kept) else []) for i, k in enumerate(order)}


def parse_date(html, n):
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", html)
    if m and m.group(2).lower() in MONTHS:
        return f"{int(m.group(3)):04d}-{MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
    return SCHED_2026.get(n, "") if YEAR == "2026" else ""


def pad(names, n):
    return (names[:n] + [""] * n)[:n]


HEADER = (["Date", "Stage"] + ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
          + [f"GC #{i}" for i in range(1, 11)] + [f"Points #{i}" for i in range(1, 4)]
          + [f"Mountain #{i}" for i in range(1, 4)] + [f"Youth #{i}" for i in range(1, 4)])


def main():
    rows = []
    for n in range(1, MAX_STAGES + 1):
        b = f"race/{RACE}/{YEAR}/stage-{n}"
        try:
            html = fetch(b)
        except Exception as e:
            print(f"stage {n}: fetch error {e}")
            continue
        c = classify(html)
        if not c["stage"]:
            continue
        # Stage 1 is a TTT: the PCS "stage" table ranks by TEAM (whole squads tie),
        # which distorts a 2-manager fantasy league. By league rule, stage 1 placement
        # points are awarded off the individual GC times instead.
        stage_place = c["gc"] if n == 1 else c["stage"]
        if n == 1:
            print("stage 1 (TTT): using individual GC times for stage placements")
        row = ([parse_date(html, n), n]
               + pad(stage_place, 10)
               + pad(c["gc"], 10)
               + pad(c["points"], 3)
               + pad(c["kom"], 3)
               + pad(c["youth"], 3))
        rows.append(row)
        print(f"stage {n}: {row[0]} | win {row[2]} | GC {row[12]} | Pts {row[22]} | KOM {row[25]} | Yth {row[28]}")

    if not rows:
        print("\nNo completed stages found; nothing written.")
        return
    wb = Workbook(); ws = wb.active; ws.title = "Results"; ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(OUT)
    print(f"\nWrote {len(rows)} stage(s) to {OUT}")


if __name__ == "__main__":
    main()
