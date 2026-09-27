#!/usr/bin/env python3
"""csv_cleanup.py - clean a messy CSV and write a cleaned CSV plus an .xlsx sheet.

Usage: python3 csv_cleanup.py input.csv [-o OUTPUT_PREFIX]
Requires Python 3.8+ and openpyxl (pip install openpyxl). MIT licence.
"""
import argparse, csv, re, sys
from datetime import datetime
from pathlib import Path

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d", "%b %d, %Y", "%d %b %Y"]


def snake(name, i):
    s = re.sub(r"[^0-9a-zA-Z]+", "_", name.strip()).strip("_").lower()
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", s)
    return s or f"column_{i + 1}"


def to_number(v):
    t = re.sub(r"[,$€£\s]", "", v)
    if t.endswith("%"):
        t = t[:-1]
    if re.fullmatch(r"-?\d+", t):
        return int(t)
    if re.fullmatch(r"-?\d*\.\d+", t):
        return float(t)
    return None


def to_date(v):
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(v, f).date()
        except ValueError:
            pass
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("input")
    ap.add_argument("-o", "--output", help="output prefix (default: <input>_clean)")
    a = ap.parse_args()
    src = Path(a.input)
    prefix = a.output or str(src.with_suffix("")) + "_clean"
    with open(src, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        sys.exit("empty file")
    log = []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    rows = [[c.strip() for c in r] for r in rows]
    header, body = rows[0], rows[1:]
    n0 = len(body)
    body = [r for r in body if any(r)]
    log.append(f"dropped {n0 - len(body)} fully empty row(s)")
    keep = [j for j in range(width) if header[j] or any(r[j] for r in body)]
    log.append(f"dropped {width - len(keep)} fully empty column(s)")
    header = [header[j] for j in keep]
    body = [[r[j] for j in keep] for r in body]
    seen, out = set(), []
    for r in body:
        if tuple(r) not in seen:
            seen.add(tuple(r))
            out.append(r)
    log.append(f"removed {len(body) - len(out)} duplicate row(s)")
    body = out
    new_h, used = [], {}
    for i, h in enumerate(header):
        s = snake(h, i)
        used[s] = used.get(s, 0) + 1
        new_h.append(s if used[s] == 1 else f"{s}_{used[s]}")
    for o, n in zip(header, new_h):
        if o != n:
            log.append(f"renamed header '{o}' -> '{n}'")
    for j, name in enumerate(new_h):
        vals = [r[j] for r in body if r[j]]
        if not vals:
            continue
        nums = [to_number(v) for v in vals]
        if all(x is not None for x in nums):
            for r in body:
                if r[j]:
                    r[j] = to_number(r[j])
            log.append(f"column '{name}': converted to numbers")
            continue
        ds = [to_date(v) for v in vals]
        if all(x is not None for x in ds):
            for r in body:
                if r[j]:
                    r[j] = to_date(r[j])
            log.append(f"column '{name}': converted to dates (ISO)")
    with open(prefix + ".csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(new_h)
        w.writerows([[c.isoformat() if hasattr(c, "isoformat") else c for c in r] for r in body])
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "cleaned"
    ws.append(new_h)
    for r in body:
        ws.append(r)
    ws.freeze_panes = "A2"
    wb.save(prefix + ".xlsx")
    print(f"read {n0} data row(s) from {src}")
    for l in log:
        print(" -", l)
    print(f"wrote {prefix}.csv and {prefix}.xlsx ({len(body)} rows, {len(new_h)} columns)")


if __name__ == "__main__":
    main()
