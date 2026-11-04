"""Read Statistics Canada's input-output workbook with separate tables by origin ("IOT" workbook, on request).

Sheets (recognised by name; or give them with --iot-sheets total=...,domestic=...,us=...):
  Total            all products                        -> output (Total use), intermediate inputs, household spending,
                                                          imports by origin (INTIM* columns, for the market model)
  Domestic         domestically produced products       -> domestic input flows and household spending on domestic output
  Import from US   products imported from the U.S.     -> U.S. input flows and household purchases of U.S. imports
  Imports          all imported products (optional)    -> checks only

Layout of each sheet (found automatically):
  rows    = supplying industries (products): code column (BS/NP/GS codes) and the title in the next column
  columns = using industries and final demand: code row and the title in the next row
  values  = C$ thousands, from the column after the row titles and the row after the column titles.
  The output column may be called "Total use", "Total" or "TOTAL". Pivot-table lines ("Grand Total", "Row Labels")
  are ignored. A code given only in a title as "... [BS336400]" is also recognised.

Everything stays at the industry level: no aggregation to sectors or goods.
"""
import hashlib
import json
import os
import re

import numpy as np

BRACKET = re.compile(r"\[([^\]]+)\]\s*$")
IND = re.compile(r"^(BS|NP|GS)[A-Za-z0-9]{4,8}$")
CODELIKE = re.compile(r"^[A-Z]{2,5}[A-Za-z0-9]{2,8}$")
KINDS = ("total", "domestic", "us", "imports")
IMPORT_COLS = ("INTIMUS", "INTIMCN", "INTIMMX", "INTIMRW")
SKIP = ("grand total", "row labels", "column labels", "none", "sum of value", "")


def is_industry(code):
    return bool(IND.match(code or ""))


def _code_of(code_cell, title_cell):
    """Code from its own cell if it looks like a code, else from '[CODE]' at the end of the title, else the title."""
    v = str(code_cell).strip() if code_cell is not None else ""
    t = str(title_cell or "").strip()
    if v.lower() in ("total", "total use", "total uses") or t.lower() in ("total use",) and not v:
        return "Total use"
    if CODELIKE.match(v):
        return v
    m = BRACKET.search(t)
    if m:
        return m.group(1)
    if t.lower() in ("total", "total use", "total uses"):
        return "Total use"
    return v or t


def _kind(sheet_name):
    s = sheet_name.lower().replace("_", " ").replace("-", " ")
    if "domestic" in s:
        return "domestic"
    if "import" in s and ("us" in s.split() or "u.s" in s or "united states" in s):
        return "us"
    if "import" in s:
        return "imports"
    if "total" in s:
        return "total"
    return None


def sheet_rows(ws):
    """All rows of a sheet as tuples of values. In read-only mode openpyxl trusts the size recorded in the file, which some
    exporting systems write wrongly (e.g. "A1:A1"): reset it so every row and column is read."""
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()
    return [tuple(r) for r in ws.iter_rows(values_only=True)]


def find_layout(ws, look=30):
    """(code_row, code_col): the row with the most industry codes across the top and the column with the most
    industry codes down the left."""
    grid = ws if isinstance(ws, list) else list(ws.iter_rows(values_only=True))
    grid = grid[:400]
    count = lambda vals: sum(1 for v in vals if isinstance(v, str) and IND.match(v.strip()))
    ncol = max((len(r) for r in grid), default=0)
    rows_ok = [r for r in range(min(look, len(grid))) if count(grid[r]) > 0]
    cols_ok = [c for c in range(min(10, ncol)) if count(r[c] if c < len(r) else None for r in grid) > 0]
    if not rows_ok or not cols_ok:
        where = getattr(ws, "title", "sheet") if not isinstance(ws, list) else "sheet"
        raise ValueError(f"IOT {where}: no industry codes (BS..., NP..., GS...) found "
                         + ("across the top (first 30 rows)" if not rows_ok else "down the left (first 10 columns)")
                         + f". The sheet was read as {len(grid)} rows x {ncol} columns. First rows: {[r[:6] for r in grid[:8]]}")
    code_row = max(rows_ok, key=lambda r: count(grid[r])) + 1
    code_col = max(cols_ok, key=lambda c: count(r[c] if c < len(r) else None for r in grid)) + 1
    return code_row, code_col


def read_sheet(ws, layout=None, log=print):
    """({supply code: {use code: value}}, {code: title}, layout)."""
    rows = sheet_rows(ws)
    try:
        code_row, code_col = layout or find_layout(rows)
    except ValueError as e:
        raise ValueError(str(e).replace("IOT sheet", f"IOT sheet '{ws.title}'")) from None
    title_row, title_col = code_row + 1, code_col + 1
    log(f"  sheet '{ws.title}': column codes in row {code_row}, titles row {title_row}; "
        f"row codes in column {code_col}, titles column {title_col}; values from row {title_row + 1}, column {title_col + 1}")
    cols, titles = {}, {}
    hdr_c, hdr_t = rows[code_row - 1], rows[title_row - 1]
    for c in range(title_col, len(hdr_c)):
        code = _code_of(hdr_c[c], hdr_t[c] if c < len(hdr_t) else None)
        if code.lower() in SKIP:
            continue
        cols[c] = code; titles.setdefault(code, str((hdr_t[c] if c < len(hdr_t) else None) or code))
    cells = {}
    for r in rows[title_row:]:
        if len(r) < title_col:
            continue
        code = _code_of(r[code_col - 1], r[title_col - 1])
        if code.lower() in SKIP:
            continue
        titles.setdefault(code, str(r[title_col - 1] or code))
        row = cells.setdefault(code, {})
        for c, uc in cols.items():
            v = r[c] if c < len(r) else None
            if isinstance(v, (int, float)) and v != 0:
                row[uc] = row.get(uc, 0.0) + float(v)
    return cells, titles, (code_row, code_col)


def read_workbook(path, sheets=None, log=print):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    names = {}
    if sheets:
        names = dict(sheets)
    else:
        for n in wb.sheetnames:
            k = _kind(n)
            if k and k not in names:
                names[k] = n
    missing = [k for k in ("total", "domestic", "us") if k not in names]
    if missing:
        raise ValueError(f"{os.path.basename(path)}: no sheet found for {', '.join(missing)} (sheets: {wb.sheetnames}). "
                         "Use --iot-sheets total=<sheet>,domestic=<sheet>,us=<sheet>")
    out = {}
    for k, n in names.items():
        out[k] = read_sheet(wb[n], log=log)
    return out, names


def build(tables, names):
    """Industry-level arrays (C$ thousands). Industries = BS/NP/GS codes of the Total table, in table order."""
    tot, titles, _ = tables["total"]
    inds = [c for c in tot if is_industry(c)]
    for row in tot.values():
        for c in row:
            if is_industry(c) and c not in inds:
                inds.append(c)
    ix = {c: k for k, c in enumerate(inds)}
    n = len(inds)

    def flows(cells):
        Z = np.zeros((n, n)); H = np.zeros(n)
        for sup, row in cells.items():
            a = ix.get(sup)
            if a is None:
                continue
            for c, v in row.items():
                b = ix.get(c)
                if b is not None:
                    Z[a, b] += v
                elif c.startswith("PEC"):
                    H[a] += v
        return Z, H

    Z, H = flows(tot)
    Zd, Hd = flows(tables["domestic"][0])
    Zus, Hus = flows(tables["us"][0])
    X = np.array([tot.get(c, {}).get("Total use", 0.0) for c in inds])
    IMP = np.array([[-tot.get(c, {}).get(k, 0.0) for k in IMPORT_COLS] for c in inds])   # imports by origin, positive
    out = dict(inds=inds, names=[BRACKET.sub("", titles.get(c, c)).strip() for c in inds], X=X.tolist(), Z=Z.tolist(),
               H=H.tolist(), Zd=Zd.tolist(), Hd=Hd.tolist(), Zus=Zus.tolist(), Hus=Hus.tolist(), IMP=IMP.tolist(),
               Htot=float(H.sum()), sheets=names, has_imports="imports" in tables)
    if "imports" in tables:
        Zm, Hm = flows(tables["imports"][0]); out.update(Zm=Zm.tolist(), Hm=Hm.tolist())
    return out


def load(path, processed_dir, sheets=None, log=print):
    """IOT arrays for the workbook at `path`, cached in data/processed. The cache is tied to the file's size and
    modification time, so editing or replacing the workbook rebuilds it automatically."""
    st = os.stat(path)
    key = hashlib.md5(f"{os.path.abspath(path)}|{st.st_size}|{st.st_mtime}|{sheets}".encode()).hexdigest()[:10]
    out = os.path.join(processed_dir, f"iot_{os.path.splitext(os.path.basename(path))[0]}_{key}.json")
    if os.path.exists(out):
        return json.load(open(out))
    log(f"Reading the IOT workbook {os.path.basename(path)} ...")
    tables, names = read_workbook(path, sheets, log)
    d = build(tables, names)
    d["source"] = os.path.basename(path)
    os.makedirs(processed_dir, exist_ok=True)
    json.dump(d, open(out, "w"))
    return d
