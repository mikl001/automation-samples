"""Turn a folder of messy monthly order exports into a formatted Excel report.

    python build_report.py --input ../excel-sales-report/data --output Sales_Report_2025.xlsx

Cleans what real webshop exports get wrong (inconsistent product spelling and
spacing, mixed country codes/names, duplicated lines, blank rows), then writes a
report with a summary dashboard, the clean data and a data-quality log.
Standard library + xlsxwriter only.
"""
import argparse
import csv
import re
import time
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import xlsxwriter

PRODUCTS = {  # canonical name -> (SKU, category)
    "Ethiopia Yirgacheffe 250g": ("CB-ETH-250", "Coffee Beans"),
    "Colombia Huila 250g": ("CB-COL-250", "Coffee Beans"),
    "Brazil Santos 1kg": ("CB-BRA-1KG", "Coffee Beans"),
    "House Espresso Blend 500g": ("CB-ESP-500", "Coffee Beans"),
    "Japanese Sencha 100g": ("TE-SEN-100", "Tea"),
    "Earl Grey Supreme 100g": ("TE-EGR-100", "Tea"),
    "Ceremonial Matcha 30g": ("TE-MAT-030", "Tea"),
    "V60 Dripper Ceramic": ("EQ-V60-02", "Equipment"),
    "Hand Grinder Steel Burr": ("EQ-GRD-HND", "Equipment"),
    "Gooseneck Kettle 1L": ("EQ-KTL-GOS", "Equipment"),
    "Paper Filters 100pcs": ("AC-FLT-100", "Accessories"),
    "Stoneware Mug 350ml": ("AC-MUG-CER", "Accessories"),
}
COUNTRIES = {
    "NL": "Netherlands", "NETHERLANDS": "Netherlands", "NEDERLAND": "Netherlands",
    "BE": "Belgium", "BELGIUM": "Belgium", "BELGIE": "Belgium", "BELGIË": "Belgium",
    "DE": "Germany", "GERMANY": "Germany", "DEUTSCHLAND": "Germany",
    "FR": "France", "FRANCE": "France",
}


def key(text: str) -> str:
    """Case- and whitespace-insensitive lookup key."""
    return re.sub(r"\s+", " ", text or "").strip().lower()


PRODUCT_BY_KEY = {key(name): name for name in PRODUCTS}


def load(folder: Path, log: Counter):
    rows, seen = [], set()
    for path in sorted(folder.glob("*.csv")):
        log["files"] += 1
        with path.open(encoding="utf-8", newline="") as f:
            for raw in csv.DictReader(f):
                log["raw_lines"] += 1
                if not (raw.get("OrderID") or "").strip():
                    log["blank_lines"] += 1
                    continue
                fingerprint = tuple(raw.values())
                if fingerprint in seen:
                    log["duplicates"] += 1
                    continue
                seen.add(fingerprint)
                product = PRODUCT_BY_KEY.get(key(raw["Product"]))
                if product is None:
                    log["unknown_product"] += 1
                    continue
                if product != raw["Product"]:
                    log["product_spelling_fixed"] += 1
                country_raw = raw["Country"].strip().upper()
                country = COUNTRIES.get(country_raw)
                if country is None:
                    log["unknown_country"] += 1
                    continue
                if country != raw["Country"]:
                    log["country_standardised"] += 1
                qty, price, disc = int(raw["Qty"]), float(raw["UnitPrice"]), float(raw["DiscountPct"])
                sku, category = PRODUCTS[product]
                rows.append({
                    "OrderID": raw["OrderID"].strip(), "OrderDate": date.fromisoformat(raw["OrderDate"]),
                    "Country": country, "Channel": raw["Channel"].strip().title(), "SKU": sku,
                    "Product": product, "Category": category, "Qty": qty, "UnitPrice": price,
                    "Discount": disc, "Revenue": round(qty * price * (1 - disc), 2),
                })
    log["clean_lines"] = len(rows)
    return rows


def write_report(rows, log: Counter, out: Path) -> None:
    wb = xlsxwriter.Workbook(str(out))
    ink, muted, slate, copper, sage, paper = "#1F2A30", "#6B7780", "#2F3E46", "#C8773A", "#6E8B74", "#F4F5F7"
    title = wb.add_format({"bold": True, "font_size": 18, "font_color": "#FFFFFF", "bg_color": slate, "valign": "vcenter"})
    band = wb.add_format({"bg_color": slate})
    sub = wb.add_format({"font_size": 9, "font_color": "#C9D1D6", "bg_color": slate})
    label = wb.add_format({"font_size": 8, "bold": True, "font_color": muted, "bg_color": "#FFFFFF", "top": 1, "left": 1, "right": 1, "border_color": "#DADFE3", "indent": 1})
    big = {"font_size": 20, "bold": True, "font_color": ink, "bg_color": "#FFFFFF", "left": 1, "right": 1, "bottom": 1, "border_color": "#DADFE3", "indent": 1, "valign": "vcenter"}
    big_eur, big_int = wb.add_format({**big, "num_format": "€#,##0"}), wb.add_format({**big, "num_format": "#,##0"})
    big_eur2 = wb.add_format({**big, "num_format": "€#,##0.00"})
    head = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": slate})
    eur, eur2, pct, dfmt = (wb.add_format({"num_format": f}) for f in ("€#,##0", "€#,##0.00", "0%", "yyyy-mm-dd"))
    mfmt = wb.add_format({"num_format": "mmm yyyy"})
    bg = wb.add_format({"bg_color": paper})

    # --- aggregates
    by_month, by_cat, by_country, by_product = defaultdict(float), defaultdict(float), defaultdict(float), defaultdict(float)
    for r in rows:
        by_month[r["OrderDate"].replace(day=1)] += r["Revenue"]
        by_cat[r["Category"]] += r["Revenue"]
        by_country[r["Country"]] += r["Revenue"]
        by_product[r["Product"]] += r["Revenue"]
    revenue = sum(r["Revenue"] for r in rows)
    orders = len({r["OrderID"] for r in rows})

    # --- Summary sheet
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2, bg)
    ws.set_column("B:M", 11, bg)
    ws.set_column("N:N", 2, bg)
    ws.set_row(0, 34)
    ws.merge_range("A1:N1", "  Bean & Leaf — sales report 2025", title)
    ws.merge_range("A2:N2", f"  Generated by build_report.py from {log['files']} monthly exports · {time.strftime('%d %b %Y %H:%M')}", sub)
    tiles = [("REVENUE", revenue, big_eur), ("ORDERS", orders, big_int), ("AVG ORDER", revenue / orders, big_eur2),
             ("UNITS", sum(r["Qty"] for r in rows), big_int)]
    for i, (lab, val, fmt) in enumerate(tiles):
        col = 1 + i * 3
        ws.merge_range(3, col, 3, col + 2, lab, label)
        ws.merge_range(4, col, 5, col + 2, val, fmt)

    calc = wb.add_worksheet("Calc")
    calc.write_row(0, 0, ["Month", "Revenue", "", "Category", "Revenue", "", "Country", "Revenue"])
    for i, m in enumerate(sorted(by_month)):
        calc.write_datetime(i + 1, 0, _dt(m), mfmt)
        calc.write_number(i + 1, 1, round(by_month[m], 2))
    for i, (k, v) in enumerate(sorted(by_cat.items(), key=lambda kv: -kv[1])):
        calc.write(i + 1, 3, k)
        calc.write_number(i + 1, 4, round(v, 2))
    for i, (k, v) in enumerate(sorted(by_country.items(), key=lambda kv: -kv[1])):
        calc.write(i + 1, 6, k)
        calc.write_number(i + 1, 7, round(v, 2))
    calc.hide()

    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Revenue", "categories": ["Calc", 1, 0, len(by_month), 0], "values": ["Calc", 1, 1, len(by_month), 1],
                   "fill": {"color": slate}, "gap": 60})
    ch.set_title({"name": "Revenue by month", "name_font": {"size": 11, "color": ink}})
    ch.set_legend({"none": True})
    ch.set_x_axis({"num_format": "mmm", "num_font": {"color": muted}})
    ch.set_y_axis({"num_format": "€#,##0", "num_font": {"color": muted}, "major_gridlines": {"visible": True, "line": {"color": "#DADFE3"}}, "line": {"none": True}})
    ch.set_chartarea({"border": {"none": True}})
    ch.set_size({"width": 560, "height": 300})  # columns B:H — must not slide under the category chart at I8
    ws.insert_chart("B8", ch)

    cc = wb.add_chart({"type": "bar"})
    cc.add_series({"name": "Revenue", "categories": ["Calc", 1, 3, len(by_cat), 3], "values": ["Calc", 1, 4, len(by_cat), 4],
                   "fill": {"color": copper}, "gap": 60, "data_labels": {"value": True, "num_format": "€#,##0"}})
    cc.set_title({"name": "Revenue by category", "name_font": {"size": 11, "color": ink}})
    cc.set_legend({"none": True})
    cc.set_y_axis({"reverse": True, "num_font": {"color": muted}})
    cc.set_x_axis({"visible": False, "major_gridlines": {"visible": False}})
    cc.set_chartarea({"border": {"none": True}})
    cc.set_size({"width": 420, "height": 300})
    ws.insert_chart("I8", cc)

    ws.write("B25", "Top 5 products", wb.add_format({"bold": True, "font_size": 11, "bg_color": paper}))
    for i, (p, v) in enumerate(sorted(by_product.items(), key=lambda kv: -kv[1])[:5]):
        ws.write(25 + i, 1, p, bg)
        ws.write_number(25 + i, 5, round(v), wb.add_format({"num_format": "€#,##0", "bg_color": paper}))
    ws.write("I25", "Data quality — fixed automatically", wb.add_format({"bold": True, "font_size": 11, "bg_color": paper}))
    dq = [("Raw lines read", log["raw_lines"]), ("Blank lines removed", log["blank_lines"]),
          ("Duplicate lines removed", log["duplicates"]), ("Product names corrected", log["product_spelling_fixed"]),
          ("Country values standardised", log["country_standardised"]), ("Clean order lines", log["clean_lines"])]
    for i, (k, v) in enumerate(dq):
        ws.write(25 + i, 8, k, bg)
        ws.write_number(25 + i, 12, v, wb.add_format({"num_format": "#,##0", "bg_color": paper}))
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.print_area("A1:N32")

    # --- Clean data sheet
    ds = wb.add_worksheet("Clean data")
    cols = ["OrderID", "OrderDate", "Country", "Channel", "SKU", "Product", "Category", "Qty", "UnitPrice", "Discount", "Revenue"]
    fmts = {"OrderDate": dfmt, "UnitPrice": eur2, "Discount": pct, "Revenue": eur2}
    data = [[r[c] if c != "OrderDate" else _dt(r[c]) for c in cols] for r in rows]
    ds.add_table(0, 0, len(rows), len(cols) - 1, {"data": data, "style": "Table Style Light 9",
                 "columns": [{"header": c, "format": fmts.get(c)} for c in cols]})
    ds.set_column("A:A", 16)
    ds.set_column("B:E", 12)
    ds.set_column("F:F", 26)
    ds.set_column("G:K", 12)
    ds.freeze_panes(1, 0)
    wb.close()


def _dt(d: date):
    from datetime import datetime
    return datetime(d.year, d.month, d.day)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--input", type=Path, default=Path(__file__).parent.parent / "excel-sales-report" / "data")
    ap.add_argument("--output", type=Path, default=Path(__file__).parent / "Sales_Report_2025.xlsx")
    args = ap.parse_args()
    started = time.perf_counter()
    log = Counter()
    rows = load(args.input, log)
    write_report(rows, log, args.output)
    print(f"Read {log['files']} files, {log['raw_lines']:,} lines")
    print(f"  removed {log['blank_lines']} blank and {log['duplicates']} duplicate lines")
    print(f"  corrected {log['product_spelling_fixed']} product names, standardised {log['country_standardised']} country values")
    print(f"  unknown products: {log['unknown_product']}, unknown countries: {log['unknown_country']}")
    print(f"Wrote {args.output.name}: {log['clean_lines']:,} clean lines in {time.perf_counter() - started:.2f} s")


if __name__ == "__main__":
    main()
