#!/usr/bin/env python3
"""
Scrape Tour de France stage results from ProCyclingStats into the column layout
the Fantasy Cycling engine expects (tdf-results.xlsx).

Verbose/diagnostic build: prints exactly what the procyclingstats library
returns so we can confirm method names and field keys.

    pip install procyclingstats openpyxl
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py
"""
import os
import sys
import traceback

import procyclingstats
from procyclingstats import Stage
from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))

print(f"procyclingstats version: {getattr(procyclingstats, '__version__', '?')}")
print(f"YEAR={YEAR} OUT={OUT} RACE={RACE}")

HEADER = (
    ["Date", "Stage"]
    + ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
    + [f"GC #{i}" for i in range(1, 11)]
    + [f"Points #{i}" for i in range(1, 4)]
    + [f"Mountain #{i}" for i in range(1, 4)]
    + [f"Youth #{i}" for i in range(1, 4)]
)


def fmt_name(pcs_name):
    if not pcs_name:
        return ""
    toks = str(pcs_name).split()
    surname = [t for t in toks if t == t.upper()]
    given = [t for t in toks if t != t.upper()]
    return (" ".join(given) + " " + " ".join(w.capitalize() for w in surname)).strip()


def name_of(row):
    """Pull the rider's display name from a result row, whatever the key is called."""
    if isinstance(row, dict):
        for k in ("rider_name", "rider", "name"):
            if row.get(k):
                return row[k]
    return ""


def rank_of(row, fallback):
    if isinstance(row, dict):
        for k in ("rank", "position", "place"):
            if row.get(k) not in (None, ""):
                try:
                    return int(row[k])
                except (TypeError, ValueError):
                    pass
    return fallback


def get_classification(stage, method_name, parsed):
    """Return a list of result rows for a classification, trying the method then parse()."""
    try:
        v = getattr(stage, method_name)()
        if v:
            return v
    except Exception as e:
        print(f"    .{method_name}() raised: {e}", file=sys.stderr)
    if isinstance(parsed, dict) and parsed.get(method_name):
        return parsed[method_name]
    return []


def ranked(rows, n):
    out = [""] * n
    for i, row in enumerate(rows or []):
        r = rank_of(row, i + 1)
        if 1 <= r <= n:
            out[r - 1] = fmt_name(name_of(row))
    return out


def scrape_stage(n, verbose):
    url = f"race/{RACE}/{YEAR}/stage-{n}"
    stage = Stage(url)
    parsed = None
    try:
        parsed = stage.parse()
        if verbose:
            keys = list(parsed.keys()) if isinstance(parsed, dict) else type(parsed)
            print(f"  parse() keys: {keys}")
    except Exception as e:
        if verbose:
            print(f"  parse() raised: {e}", file=sys.stderr)

    results = get_classification(stage, "results", parsed)
    if verbose and results:
        print(f"  first result row: {results[0]}")
    if not results:
        return None

    try:
        date = stage.date()
    except Exception:
        date = (parsed or {}).get("date", "") if isinstance(parsed, dict) else ""

    row = [date, n]
    row += ranked(results, 10)
    row += ranked(get_classification(stage, "gc", parsed), 10)
    row += ranked(get_classification(stage, "points", parsed), 3)
    row += ranked(get_classification(stage, "kom", parsed), 3)
    row += ranked(get_classification(stage, "youth", parsed), 3)
    return row


def main():
    rows = []
    for n in range(1, MAX_STAGES + 1):
        try:
            r = scrape_stage(n, verbose=(n == 1))
        except Exception as e:
            print(f"stage {n}: ERROR {e}", file=sys.stderr)
            traceback.print_exc()
            r = None
        if r is None:
            print(f"stage {n}: no results")
            continue
        print(f"stage {n}: {r[0]} winner {r[2]}")
        rows.append(r)

    if not rows:
        print("No completed stages found; nothing written.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(OUT)
    print(f"Wrote {len(rows)} stage(s) to {OUT}")


if __name__ == "__main__":
    main()
