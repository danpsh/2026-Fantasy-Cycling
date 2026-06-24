#!/usr/bin/env python3
"""
Scrape Tour de France stage results from ProCyclingStats and write them in the
exact column layout the Fantasy Cycling scoring engine expects (tdf-results.xlsx).

Run locally:
    pip install procyclingstats openpyxl
    YEAR=2026 OUT=tdf-results.xlsx python scripts/scrape_tdf.py

Validate against last year before the Tour:
    YEAR=2025 OUT=tdf-results-2025-test.xlsx python scripts/scrape_tdf.py

The engine reads, per stage row:
    Date, Stage, 1st..10th, GC #1..#10, Points #1-3, Mountain #1-3, Youth #1-3
It rebuilds the whole file each run from every completed stage (idempotent).
"""
import os
import sys

from procyclingstats import Stage
from openpyxl import Workbook

YEAR = os.environ.get("YEAR", "2026")
OUT = os.environ.get("OUT", "tdf-results.xlsx")
RACE = os.environ.get("RACE_SLUG", "tour-de-france")
MAX_STAGES = int(os.environ.get("MAX_STAGES", "21"))

HEADER = (
    ["Date", "Stage"]
    + ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
    + [f"GC #{i}" for i in range(1, 11)]
    + [f"Points #{i}" for i in range(1, 4)]
    + [f"Mountain #{i}" for i in range(1, 4)]
    + [f"Youth #{i}" for i in range(1, 4)]
)


def fmt_name(pcs_name):
    """'VINGEGAARD Jonas' -> 'Jonas Vingegaard'. Surname tokens are ALL-CAPS."""
    if not pcs_name:
        return ""
    toks = pcs_name.split()
    surname = [t for t in toks if t == t.upper()]
    given = [t for t in toks if t != t.upper()]
    surname_title = " ".join(w.capitalize() for w in surname)
    return (" ".join(given) + " " + surname_title).strip()


def ranked(rows, n):
    """Return rider names ordered by rank 1..n (PCS rows have 'rank' + 'rider_name')."""
    out = [""] * n
    if not rows:
        return out
    for row in rows:
        rank = row.get("rank")
        try:
            r = int(rank)
        except (TypeError, ValueError):
            continue
        if 1 <= r <= n:
            out[r - 1] = fmt_name(row.get("rider_name", ""))
    return out


def safe(method):
    try:
        v = method()
        return v or []
    except Exception:
        return []


def scrape_stage(n):
    url = f"race/{RACE}/{YEAR}/stage-{n}"
    stage = Stage(url)
    results = safe(stage.results)
    if not results:
        return None  # not raced yet (or no finishers parsed)
    try:
        date = stage.date()
    except Exception:
        date = ""
    row = [date, n]
    row += ranked(results, 10)
    row += ranked(safe(stage.gc), 10)
    row += ranked(safe(stage.points), 3)
    row += ranked(safe(stage.kom), 3)
    row += ranked(safe(stage.youth), 3)
    return row


def main():
    rows = []
    for n in range(1, MAX_STAGES + 1):
        try:
            r = scrape_stage(n)
        except Exception as e:
            print(f"stage {n}: error {e}", file=sys.stderr)
            r = None
        if r is None:
            print(f"stage {n}: no results yet — stopping")
            break
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
