"""Read the BS-HS8 file: which input-output industry (BS code) makes each HS8 tariff item.

Two columns, CSV or Excel (first sheet, or --concordance-sheet). The columns are found by their content: the column
with the most BS/NP/GS industry codes, and the column with the most 8-digit HS codes (written 01012100, 0101.21.00 or
as a number, where Excel has dropped the leading zero). A header row is skipped automatically.
An HS8 code listed with more than one industry is split equally between them (the file lists the split).
"""
import csv
import os
import re

from .iot import IND

HS8 = re.compile(r"^\d{8}$")


def _hs8(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"[.\s]", "", str(v).strip())
    if s.isdigit() and 7 <= len(s) <= 8:
        s = s.zfill(8)
    return s if HS8.match(s) else None


def _rows(path, sheet=None):
    """Rows of the file. Excel: the named sheet, else the first sheet that has industry codes. CSV: the delimiter
    (comma, semicolon, tab or |) is detected. A one-column file with "BS111A00 01012100" in each cell is split."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        from .iot import sheet_rows
        wb = load_workbook(path, read_only=True, data_only=True)
        sheets = [wb[sheet]] if sheet else wb.worksheets
        rows = []
        for ws in sheets:
            rows = [list(r) for r in sheet_rows(ws)]
            if any(isinstance(v, str) and IND.match(v.strip()) for r in rows for v in r):
                break
    else:
        text = open(path, encoding="utf-8-sig", errors="replace").read()
        try:
            dialect = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|")
            rows = list(csv.reader(text.splitlines(), dialect))
        except csv.Error:
            rows = list(csv.reader(text.splitlines()))
    if rows and max((len(r) for r in rows), default=0) == 1:
        rows = [re.split(r"[\s,;|]+", str(r[0]).strip()) if r and r[0] is not None else [] for r in rows]
    return rows


def read(path, sheet=None):
    """[(hs8, bs code, share)] with share = 1 / (number of industries listed for that HS8)."""
    rows = _rows(path, sheet)
    if not rows:
        raise ValueError(f"{path} is empty")
    ncol = max((len(r) for r in rows), default=0)
    col = lambda c: [r[c] if c < len(r) else None for r in rows]
    n_bs = {c: sum(1 for v in col(c) if isinstance(v, str) and IND.match(v.strip())) for c in range(ncol)}
    n_hs = {c: sum(1 for v in col(c) if _hs8(v)) for c in range(ncol)}
    head = [r[:4] for r in rows[:5]]
    if ncol < 2 or max(n_bs.values(), default=0) == 0:
        raise ValueError(f"{os.path.basename(path)}: no column of industry codes (BS..., NP..., GS...) found; the file was read as "
                         f"{len(rows)} rows x {ncol} columns. First rows: {head}")
    bs_col = max(n_bs, key=n_bs.get)
    others = {c: v for c, v in n_hs.items() if c != bs_col}
    if not others or max(others.values()) == 0:
        raise ValueError(f"{os.path.basename(path)}: industry codes found in column {bs_col + 1}, but no column of 8-digit HS codes. "
                         f"First rows: {head}")
    hs_col = max(others, key=others.get)
    pairs, seen = [], set()
    for r in rows:
        bs = str(r[bs_col]).strip() if bs_col < len(r) and r[bs_col] is not None else ""
        h = _hs8(r[hs_col] if hs_col < len(r) else None)
        if IND.match(bs) and h and (h, bs) not in seen:
            seen.add((h, bs)); pairs.append((h, bs))
    if not pairs:
        raise ValueError(f"No (BS code, HS8) pairs found in {path}")
    n = {}
    for h, _ in pairs:
        n[h] = n.get(h, 0) + 1
    info = dict(path=os.path.basename(path), bs_column=bs_col + 1, hs8_column=hs_col + 1, lines=len(pairs),
                hs8=len(n), split=sum(1 for v in n.values() if v > 1))
    return [(h, bs, 1.0 / n[h]) for h, bs in pairs], info
