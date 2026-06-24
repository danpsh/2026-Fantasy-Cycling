#!/usr/bin/env python3
"""
Scrape Tour de France results from ProCyclingStats -> tdf-results.xlsx.

On PCS each classification is its OWN page:
    stage finish : race/<race>/<year>/stage-<n>
    GC           : race/<race>/<year>/stage-<n>-gc
    points       : race/<race>/<year>/stage-<n>-points
    KOM          : race/<race>/<year>/stage-<n>-kom
    youth        : race/<race>/<year>/stage-<n>-youth
Each page's .results() returns that ranked table.

Includes a raw-HTTP probe (browser User-Agent) so the log shows whether the
runner can reach PCS at all (HTTP 200 vs 403/blocked).

    pip install procyclingstats openpyxl
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py
"""
import os
import sys
import traceback
import urllib.request

from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

print("=== ENV ===")
print(f"YEAR={YEAR} OUT={OUT} RACE={RACE}")
try:
    from importlib.metadata import version
    print("procyclingstats version:", version("procyclingstats"))
except Exception as e:
    print("version lookup failed:", e)


def probe(url):
    full = "https://www.procyclingstats.com/" + url
    print(f"\n=== RAW PROBE {full} ===")
    try:
        req = urllib.request.Request(full, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=25) as r:
            body = r.read().decode("utf-8", "replace")
            low = body.lower()
            print(f"HTTP {r.status}, {len(body)} bytes")
            print("has results table:", ("resulttable" in low or 'class="results' in low or "<table" in low))
            print("looks blocked:", ("cloudflare" in low or "captcha" in low or "just a moment" in low))
            print("snippet:", " ".join(body[:250].split()))
    except Exception as e:
        print("RAW PROBE ERROR:", e)


probe(f"race/{RACE}/{YEAR}/stage-1")

Stage = None
try:
    from procyclingstats import Stage as _Stage
    Stage = _Stage
except Exception:
    print("import procyclingstats FAILED:")
    traceback.print_exc(file=sys.stdout)


def fmt_name(pcs_name):
    if not pcs_name:
        return ""
    toks = str(pcs_name).split()
    surname = [t for t in toks if t == t.upper()]
    given = [t for t in toks if t != t.upper()]
    return (" ".join(given) + " " + " ".join(w.capitalize() for w in surname)).strip()


def name_of(row):
    if isinstance(row, dict):
        for k in ("rider_name", "rider", "name"):
            if row.get(k):
                return row[k]
    return ""


def rank_of(row, fb):
    if isinstance(row, dict):
        for k in ("rank", "position", "place"):
            if row.get(k) not in (None, ""):
                try:
                    return int(row[k])
                except (TypeError, ValueError):
                    pass
    return fb


def page_results(url, debug=False):
    """Return the ranked .results() list for a PCS page URL ('' on failure)."""
    if not Stage:
        return []
    try:
        rows = Stage(url).results() or []
        if debug:
            print(f"  {url} -> {len(rows)} rows; first: {rows[0] if rows else None}")
        return rows
    except Exception as e:
        if debug:
            print(f"  {url} -> ERROR {e}")
            traceback.print_exc(file=sys.stdout)
        return []


def ranked(rows, n):
    out = [""] * n
    for i, row in enumerate(rows or []):
        r = rank_of(row, i + 1)
        if 1 <= r <= n:
            out[r - 1] = fmt_name(name_of(row))
    return out


def stage_date(n):
    if not Stage:
        return ""
    try:
        return Stage(f"race/{RACE}/{YEAR}/stage-{n}").date()
    except Exception:
        return ""


HEADER = (["Date", "Stage"] + ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
          + [f"GC #{i}" for i in range(1, 11)] + [f"Points #{i}" for i in range(1, 4)]
          + [f"Mountain #{i}" for i in range(1, 4)] + [f"Youth #{i}" for i in range(1, 4)])


def main():
    print("\n=== STAGE 1 DETAIL ===")
    base = f"race/{RACE}/{YEAR}/stage-1"
    page_results(base, debug=True)
    page_results(base + "-gc", debug=True)

    rows = []
    for n in range(1, MAX_STAGES + 1):
        b = f"race/{RACE}/{YEAR}/stage-{n}"
        results = page_results(b)
        if not results:
            continue
        row = ([stage_date(n), n]
               + ranked(results, 10)
               + ranked(page_results(b + "-gc"), 10)
               + ranked(page_results(b + "-points"), 3)
               + ranked(page_results(b + "-kom"), 3)
               + ranked(page_results(b + "-youth"), 3))
        rows.append(row)
        print(f"stage {n}: {row[0]} winner {row[2]}")

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
