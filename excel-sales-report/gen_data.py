"""Generate 12 messy monthly order exports for a fictional online coffee & tea shop.

The mess is deliberate and typical of real webshop exports: inconsistent product
spelling and spacing, mixed country codes/names, duplicated lines, blank rows.
Seeded, so the output is reproducible.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 20250101
OUT = Path(__file__).parent / "data"

# SKU, canonical name, category, list price (EUR)
PRODUCTS = [
    ("CB-ETH-250", "Ethiopia Yirgacheffe 250g", "Coffee Beans", 12.50),
    ("CB-COL-250", "Colombia Huila 250g", "Coffee Beans", 11.00),
    ("CB-BRA-1KG", "Brazil Santos 1kg", "Coffee Beans", 29.00),
    ("CB-ESP-500", "House Espresso Blend 500g", "Coffee Beans", 18.50),
    ("TE-SEN-100", "Japanese Sencha 100g", "Tea", 9.50),
    ("TE-EGR-100", "Earl Grey Supreme 100g", "Tea", 7.90),
    ("TE-MAT-030", "Ceremonial Matcha 30g", "Tea", 21.00),
    ("EQ-V60-02", "V60 Dripper Ceramic", "Equipment", 27.00),
    ("EQ-GRD-HND", "Hand Grinder Steel Burr", "Equipment", 64.00),
    ("EQ-KTL-GOS", "Gooseneck Kettle 1L", "Equipment", 49.00),
    ("AC-FLT-100", "Paper Filters 100pcs", "Accessories", 5.50),
    ("AC-MUG-CER", "Stoneware Mug 350ml", "Accessories", 14.00),
]
# How the same product shows up in sloppy exports.
def messy_name(name: str, rng: random.Random) -> str:
    r = rng.random()
    if r < 0.70:
        return name
    if r < 0.80:
        return name.lower()
    if r < 0.88:
        return name.upper()
    if r < 0.95:
        return "  " + name.replace(" ", "  ", 1) + " "
    return name + "  "

COUNTRY_VARIANTS = {
    "Netherlands": ["NL", "Netherlands", "netherlands", "Nederland", " NL"],
    "Belgium": ["BE", "Belgium", "belgie", "België"],
    "Germany": ["DE", "Germany", "Deutschland", "germany "],
    "France": ["FR", "France", "france"],
}
COUNTRY_WEIGHTS = {"Netherlands": 0.46, "Belgium": 0.21, "Germany": 0.22, "France": 0.11}
CHANNELS = [("Webshop", 0.62), ("Marketplace", 0.28), ("Wholesale", 0.10)]
# Seasonality: coffee sells more in autumn/winter, gifts spike in Nov/Dec.
MONTH_ORDERS = [210, 190, 205, 180, 170, 160, 150, 165, 215, 240, 300, 360]


def pick(rng: random.Random, weighted):
    items, weights = zip(*weighted)
    return rng.choices(items, weights=weights, k=1)[0]


def main() -> None:
    rng = random.Random(SEED)
    OUT.mkdir(exist_ok=True)
    order_no = 1000
    for month, n_orders in enumerate(MONTH_ORDERS, start=1):
        rows = []
        start = date(2025, month, 1)
        days = (date(2025 + month // 12, month % 12 + 1, 1) - start).days
        for _ in range(n_orders):
            order_no += 1
            order_id = f"BL-2025-{order_no:06d}"
            d = start + timedelta(days=rng.randrange(days))
            country = pick(rng, COUNTRY_WEIGHTS.items())
            country_raw = rng.choice(COUNTRY_VARIANTS[country])
            channel = pick(rng, CHANNELS)
            n_lines = rng.choices([1, 2, 3, 4], weights=[0.45, 0.30, 0.17, 0.08])[0]
            for sku, name, category, price in rng.sample(PRODUCTS, n_lines):
                qty = rng.choices([1, 2, 3, 6], weights=[0.62, 0.24, 0.09, 0.05])[0]
                if channel == "Wholesale":
                    qty *= rng.choice([5, 10, 12])
                discount = 0.0
                if channel == "Wholesale":
                    discount = 0.15
                elif month in (11, 12) and rng.random() < 0.35:
                    discount = 0.10
                rows.append([order_id, d.isoformat(), country_raw, channel,
                             messy_name(name, rng), category, qty, f"{price:.2f}", f"{discount:.2f}"])
        # ~2% duplicated lines and a few blank rows, like a re-run export.
        for row in rng.sample(rows, max(1, len(rows) // 50)):
            rows.insert(rng.randrange(len(rows)), list(row))
        rows.sort(key=lambda r: (r[1], r[0]))
        for _ in range(rng.randint(1, 3)):
            rows.insert(rng.randrange(len(rows)), [""] * 9)
        path = OUT / f"orders_2025-{month:02d}.csv"
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["OrderID", "OrderDate", "Country", "Channel", "Product",
                        "Category", "Qty", "UnitPrice", "DiscountPct"])
            w.writerows(rows)
        print(f"{path.name}: {len(rows)} rows")


if __name__ == "__main__":
    main()
