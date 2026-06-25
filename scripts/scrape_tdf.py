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
import unicodedata
import urllib.request

from selectolax.parser import HTMLParser
from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

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


def pkey(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def load_canon():
    try:
        with open("tdf-startlist.js", "r", encoding="utf-8") as f:
            t = f.read()
        t = t[t.index("{"):t.rindex("}") + 1]
        canon = {pkey(r["name"]): r["name"] for r in json.loads(t).get("riders", [])}
        print(f"loaded {len(canon)} canonical names from tdf-startlist.js")
        return canon
    except Exception as e:
        print("tdf-startlist.js NOT loaded (names will be slug-cased):", e)
        return {}


CANON = load_canon()


def fetch(url):
    full = "https://www.procyclingstats.com/" + url
    req = urllib.request.Request(full, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", "replace")


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
    print("\n=== STAGE 20 TABLE LABELS ===")
    classify(fetch(f"race/{RACE}/{YEAR}/stage-20"), debug=True)

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
        row = ([parse_date(html, n), n]
               + pad(c["stage"], 10)
               + pad(c["gc"], 10)
               + pad(c["points"], 3)
               + pad(c["kom"], 3)
               + pad(c["youth"], 3))
        rows.append(row)
        print(f"stage {n}: {row[0]} | win {row[2]} | GC {row[12]} | Pts {row[22]} | KOM {row[24]} | Yth {row[27]}")

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
