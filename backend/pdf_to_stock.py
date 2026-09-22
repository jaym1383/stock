#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_to_stock.py
================
Reads "stock.pdf" (Oracle "Stock on Hand by Subinventory" report, JayMart
Mobile format) sitting next to this script, and produces "stock.xlsx" in
the same layout as the attached example workbook:

    Category | Model | Storage | Color | Qty | Item Code

Rules applied (per the instructions found in the "คำแนะนำ" sheet of the
example workbook, and cross-checked against the example data):

  1. All "DEMO ..." items are dropped completely.
  2. All SIM items are dropped completely (identified by UOM "ใบ" rather
     than "MOBILE").
  3. "Category" only separates the brand (OPPO / Samsung / Apple / Xiaomi
     / Infinix / vivo / iQOO / TECNO / Honor ...). Smartphones and Tablets
     are NOT split into separate categories/sheets - everything goes in
     one flat table (this also matches "Goods and Services For Upgrade"
     rows being merged in with the rest, exactly like the example file).
  4. RAM and ROM are combined into a single "Storage" column, e.g.
     "4GB RAM + 128GB ROM" -> "4/128GB". If a product only has one memory
     figure (e.g. iPhone) that lone figure is used as-is (e.g. "128GB").
  5. Rows are de-duplicated: the same Category+Model+Storage+Color
     combination appearing in several locators is summed into one row
     (this matches the way the example workbook already collapses
     multiple stockroom locations of the same variant into a single Qty).

Dependencies
------------
    pip install pdfplumber openpyxl

Usage
-----
    python pdf_to_stock.py
        (expects ./stock.pdf, writes ./stock.xlsx)

    python pdf_to_stock.py path/to/input.pdf path/to/output.xlsx
        (explicit paths)
"""

import sys
import re
from collections import defaultdict

try:
    import pdfplumber
except ImportError:
    sys.exit("Missing dependency 'pdfplumber'. Install with:\n"
              "    pip install pdfplumber")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependency 'openpyxl'. Install with:\n"
              "    pip install openpyxl")


# --------------------------------------------------------------------------
# 1. PDF -> raw records
# --------------------------------------------------------------------------

# A record "starts" on a line that begins with the (long) locator prefix
# number and ends with "MOBILE <qty>" or "ใบ <qty>" (UOM + quantity).
# Oracle wraps long descriptions onto one or two extra lines; those
# continuation lines are either "<short locator suffix> <text>" or just
# "<text>" with no leading number at all.
START_RE = re.compile(r'^(\d{6,})\s+(.*?)\s+(MOBILE|\u0e43\u0e1a)\s+(\d+)\s*$')
CONT_LEADING_NUM_RE = re.compile(r'^(\d{1,4})\s+(.*)$')
# Oracle always prints the locator's short suffix number on its own line
# under the record, whether or not the description itself wrapped. When
# the description did NOT wrap, that line is nothing but the bare number
# - it carries no text and must be discarded, not appended as content.
PURE_LOCATOR_SUFFIX_RE = re.compile(r'^\d{1,6}$')

# Lines that are page furniture / report chrome and never part of a
# description - matched by prefix.
BOILERPLATE_PREFIXES = (
    "Program ID", "User ID", "Group by Category", "From Category",
    "From Subinventory", "From Item", "Locator Item Code Description",
    "Subinventory :", "Group By :", "Total Item Code By Category",
    "Grand Total", "Total by",
)

# This text is printed once per page in the left margin (a "Condition"
# label) and can land in the middle of a wrapped line; strip it out
# wherever it occurs rather than trying to special-case its position.
CONDITION_NOISE = "1383-01 \u0e2a\u0e34\u0e19\u0e04\u0e49\u0e32\u0e14\u0e35"


def extract_records(pdf_path):
    """Return a list of dicts: {item_code, description, uom, qty}."""
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]
    full_text = "\n".join(pages_text)

    records = []
    current = None

    for raw_line in full_text.splitlines():
        line = raw_line.replace(CONDITION_NOISE, " ").strip()
        if not line:
            continue
        if any(line.startswith(p) for p in BOILERPLATE_PREFIXES):
            continue

        m = START_RE.match(line)
        if m:
            if current is not None:
                records.append(current)
            current = {
                "item_code": m.group(1),
                "description": [m.group(2).strip()],
                "uom": m.group(3),
                "qty": int(m.group(4)),
                "_code_suffix_seen": False,
            }
            continue

        # Continuation line for the record currently being built.
        if current is not None:
            if PURE_LOCATOR_SUFFIX_RE.match(line):
                if not current["_code_suffix_seen"]:
                    current["item_code"] += line
                    current["_code_suffix_seen"] = True
                continue
            m2 = CONT_LEADING_NUM_RE.match(line)
            if m2:
                if not current["_code_suffix_seen"]:
                    current["item_code"] += m2.group(1)
                    current["_code_suffix_seen"] = True
                text = m2.group(2).strip()
            else:
                text = line
            if text:
                current["description"].append(text)

    if current is not None:
        records.append(current)

    for r in records:
        r["description"] = re.sub(r"\s+", " ", " ".join(r["description"])).strip()
        r.pop("_code_suffix_seen", None)

    return records


# --------------------------------------------------------------------------
# 2. Filtering (drop DEMO + SIM)
# --------------------------------------------------------------------------

def keep_record(rec):
    if rec["uom"] != "MOBILE":          # drops SIM cards (UOM = "\u0e43\u0e1a")
        return False
    if rec["description"].upper().startswith("DEMO"):
        return False
    return True


# --------------------------------------------------------------------------
# 3. Description -> Category / Model / RAM-ROM / Color
# --------------------------------------------------------------------------

BRAND_PATTERNS = [
    (re.compile(r'^SAM(?:SUNG)?\s*', re.I), "Samsung"),
    (re.compile(r'^OPPO\s*', re.I), "OPPO"),
    (re.compile(r'^iPhone\s*', re.I), "Apple"),
    (re.compile(r'^Infinix\s*', re.I), "Infinix"),
    (re.compile(r'^Vivo\s*', re.I), "vivo"),
    (re.compile(r'^iQOO\s*', re.I), "iQOO"),
    (re.compile(r'^Xiaomi\s*', re.I), "Xiaomi"),
    (re.compile(r'^TECNO\s*', re.I), "TECNO"),
    (re.compile(r'^Honor\s*', re.I), "Honor"),
]

SUFFIX_RE = re.compile(
    r'(\s*-\s*|\s+)(New SP|Keyboard Cover|Upgrade|NEW|New)\s*$'
)

PAREN_RE = re.compile(r'\([^()]*\)')

STORAGE_RE = re.compile(
    r'\d+\s*/\s*\d+\s*(?:GB|TB)'   # 8/256GB, 16/1TB
    r'|\d+\s*(?:GB|TB)'            # 128GB
    r'|\d+\s*/\s*\d+'              # 12/256  (unit missing in source)
)

CONN_RE = re.compile(r'^(LTE|5G|4G|3G|Wi[- ]?Fi)\b\s*', re.I)


def split_brand(description):
    for pattern, category in BRAND_PATTERNS:
        if pattern.match(description):
            return category, pattern.sub("", description, count=1)
    # Unknown brand: use the first word as-is so nothing is silently lost.
    parts = description.split(" ", 1)
    return parts[0], (parts[1] if len(parts) > 1 else "")


def normalize_storage(raw):
    norm = re.sub(r"\s+", "", raw)
    if not re.search(r"(GB|TB)$", norm, re.I):
        norm += "GB"
    return norm


def parse_description(description):
    """Return (category, model, storage, color) for one item description."""
    text = description.strip()

    # 1. Pull off a trailing qualifier ("- New", "NEW", "Upgrade", ...)
    #    so it can be re-attached to the Model instead of the Color.
    suffix = ""
    m = SUFFIX_RE.search(text)
    if m:
        has_dash = "-" in m.group(1)
        suffix = (" - " if has_dash else " ") + m.group(2)
        text = text[: m.start()].strip()

    # 2. Brand -> Category, remainder of the string.
    category, rest = split_brand(text)

    # 3. Drop parenthetical item codes / carrier tags, e.g. "(CPH2801)",
    #    "(SM-A176BZBJTHL)", "(AWN)". (A rare descriptive parenthetical
    #    like "(Titanium)" placed after the color is dropped too - it
    #    only ever qualifies a color that is already stated in words.)
    rest = PAREN_RE.sub(" ", rest)
    rest = re.sub(r"\s+", " ", rest).strip()

    # 4. Find the storage token.
    sm = STORAGE_RE.search(rest)
    if sm:
        model_part1 = rest[: sm.start()].strip()
        storage = normalize_storage(sm.group(0))
        after = rest[sm.end():].strip()
    else:
        # No recognizable storage token - keep everything as the model
        # and leave storage/color blank rather than guessing.
        model_part1 = rest
        storage = ""
        after = ""

    # 5. A connectivity tag (LTE/5G/...) that ended up *after* the
    #    storage token belongs with the model, not the color.
    extra = []
    while True:
        cm = CONN_RE.match(after)
        if not cm:
            break
        extra.append(cm.group(1))
        after = after[cm.end():].strip()

    color = after.strip()
    model = model_part1
    if extra:
        model = (model + " " + " ".join(extra)).strip()
    if suffix:
        model = model + suffix

    model = re.sub(r"\s+", " ", model).strip()
    color = re.sub(r"\s+", " ", color).strip()

    return category, model, storage, color


# --------------------------------------------------------------------------
# 4. Aggregate
# --------------------------------------------------------------------------

def build_rows(records):
    totals = defaultdict(int)
    item_codes = defaultdict(set)
    for rec in records:
        if not keep_record(rec):
            continue
        category, model, storage, color = parse_description(rec["description"])
        key = (category, model, storage, color)
        totals[key] += rec["qty"]
        item_code = str(rec.get("item_code") or "").strip()
        if item_code:
            item_codes[key].add(item_code)

    rows = [list(key) + [qty, ", ".join(sorted(item_codes.get(key, [])))] for key, qty in totals.items()]
    rows.sort(key=lambda r: (r[0], r[1], r[2], r[3]))
    return rows


# --------------------------------------------------------------------------
# 5. Write the workbook: Category | Model | Storage | Color | Qty |
#    Item Code, bold header, banded rows, autofilter.
# --------------------------------------------------------------------------

def write_workbook(rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock"

    headers = ["Category", "Model", "Storage", "Color", "Qty", "Item Code"]
    ws.append(headers)

    for row in rows:
        ws.append(row)

    n_rows = len(rows) + 1
    n_cols = len(headers)

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(bold=True, color="FFFFFF")
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    widths = {"A": 14, "B": 38, "C": 12, "D": 22, "E": 8, "F": 20}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"

    table_ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
    table = Table(displayName="StockTable", ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True
    )
    ws.add_table(table)

    wb.save(out_path)


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main():
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "stock.pdf"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "stock.xlsx"

    print(f"Reading: {pdf_path}")
    records = extract_records(pdf_path)
    print(f"  -> {len(records)} raw line-items found in the PDF")

    kept = [r for r in records if keep_record(r)]
    print(f"  -> {len(kept)} kept after removing DEMO / SIM items "
          f"(sum of Qty = {sum(r['qty'] for r in kept)})")

    rows = build_rows(records)
    print(f"  -> {len(rows)} distinct Category/Model/RAM-ROM/Color rows "
          f"(sum of Qty = {sum(r[4] for r in rows)})")

    write_workbook(rows, out_path)
    print(f"Written: {out_path}")


if __name__ == "__main__":
    main()
