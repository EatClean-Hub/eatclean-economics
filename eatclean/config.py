import os

# ─── DRIVE PATHS ──────────────────────────────────────────────────────────────
DRIVE_ROOT = os.path.join(
    os.path.expanduser("~"),
    "Library", "CloudStorage",
    "GoogleDrive-admin@eatcleanme.com",
    "My Drive", "02 Automation"
)
PATHS = {
    "daily_selections": os.path.join(DRIVE_ROOT, "01 Daily_selections"),
    "invoices":         os.path.join(DRIVE_ROOT, "02 Invoices"),
    "reference":        os.path.join(DRIVE_ROOT, "03 Reference"),
    "outputs":          os.path.join(DRIVE_ROOT, "04 Outputs"),
}

# ─── FIXED COSTS ──────────────────────────────────────────────────────────────
DELIVERY_PER_DAY  = 14.00   # AED — LogX
PACKAGING_PER_DAY =  2.50   # AED
VAT_DIVISOR       =  1.05
TARGET_MARGIN            = 40.0
CLASSPASS_PRICE_PER_ORDER = round(356.25 / 1.05, 4)  # 356.25 incl VAT → 339.2857 ex-VAT
CLASSPASS_PRICE   = 356.25   # Fixed ex-VAT revenue per ClassPass order (AED)    # % minimum

# ─── SCOPE ────────────────────────────────────────────────────────────────────
# Months with daily selection data available (reconciliation possible)
SELECTION_MONTHS = ["2026-03", "2026-04", "2026-05"]

# Invoice sheet names → year
INVOICE_SHEETS = {
    "Daily Order-MAR26": 2026,
    "Daily Order-APR26": 2026,
}

# ─── PLAN ROUTING ─────────────────────────────────────────────────────────────
# HP clients — identified from orders file, always separate sheet
# Routing uses orders file as source of truth (built at runtime in engine.py)

VALOREM_KEYWORDS = ["VALOREM"]

# ─── DISH PRICE BENCHMARKS (max per dish, ex-VAT) ─────────────────────────────
# Source: Menu Composition doc Jan 2026
# Structure: (plan_group, meal_type, tier) → max_price_ex_vat
BENCHMARKS = {
    # Fitness group: Balanced, New Moms, Gluten Free, Sugar Free, Vegetarian
    ("fitness", "Breakfast", "LOWCAL"): 11.50,
    ("fitness", "Breakfast", "LEAN"):   12.50,
    ("fitness", "Breakfast", "BUILD"):  16.00,
    ("fitness", "Mains",     "LOWCAL"): 22.50,
    ("fitness", "Mains",     "LEAN"):   24.00,
    ("fitness", "Mains",     "BUILD"):  28.50,
    ("fitness", "Snack",     "LOWCAL"):  6.50,
    ("fitness", "Snack",     "LEAN"):    7.50,
    ("fitness", "Snack",     "BUILD"):  11.00,
    # Vegan / Vegetarian
    ("vegan",   "Breakfast", "LOWCAL"): 12.00,
    ("vegan",   "Breakfast", "LEAN"):   13.00,
    ("vegan",   "Breakfast", "BUILD"):  13.50,
    ("vegan",   "Mains",     "LOWCAL"): 16.00,
    ("vegan",   "Mains",     "LEAN"):   18.50,
    ("vegan",   "Mains",     "BUILD"):  22.50,
    ("vegan",   "Snack",     "LOWCAL"):  6.50,
    ("vegan",   "Snack",     "LEAN"):    7.50,
    ("vegan",   "Snack",     "BUILD"):  10.00,
    # Pescatarian
    ("pesca",   "Breakfast", "LOWCAL"): 11.50,
    ("pesca",   "Breakfast", "LEAN"):   12.50,
    ("pesca",   "Breakfast", "BUILD"):  16.00,
    ("pesca",   "Mains",     "LOWCAL"): 22.50,
    ("pesca",   "Mains",     "LEAN"):   24.00,
    ("pesca",   "Mains",     "BUILD"):  28.50,
    ("pesca",   "Snack",     "LOWCAL"):  6.50,
    ("pesca",   "Snack",     "LEAN"):    7.50,
    ("pesca",   "Snack",     "BUILD"):  11.00,
}

# Pre-computed min/max per (meal_type, tier) for fast flagging
BENCH_MIN = {}
BENCH_MAX = {}
for (pg, mt, tier), price in BENCHMARKS.items():
    key = (mt, tier)
    BENCH_MIN[key] = min(BENCH_MIN.get(key, 9999), price)
    BENCH_MAX[key] = max(BENCH_MAX.get(key, 0),    price)

# ─── MANUAL PRICE OVERRIDES ───────────────────────────────────────────────────
# Dishes missing or mismatched in price list — confirmed manually
MANUAL_PRICES = {
    # ── HISTORICAL ALIASES — old names still appearing in selections ──────────
    "Keto Fajitas LOWCAL":   18.9333,   # → now Keto Beef Fajitas LOWCAL in Nutribot
    "Carrots Pancake LOWCAL":  9.3429,  # → now Carrot Pancakes LOWCAL in Nutribot

    # ── SPECIAL CASES — name in selections has no tier suffix ─────────────────
    "LEAN BEEF BURGER":  21.22,   # selections use no tier; price list has LEAN variant
    "Meatballs & Mash":  19.80,   # selections sometimes omit tier

    # ── GENUINELY MISSING FROM PRICE LIST ────────────────────────────────────
    # Dishes that appear in selections but have no matching entry in price list
    "Keto Beef Gyro With Tzatsiki LEAN":   20.5333,  # only LOWCAL+BUILD in price list
    "LEAN BEEF BURGER":                    21.2200,  # no-tier version, Valorem uses this
    "Meatballs & Mash":                    20.7900,  # no-tier version
    "Lemon Garlic Tuna Pasta LEAN":          21.95,
    "Vermicelli Shrimp Noodle Bowl LEAN":    17.2571,
    "Vermicelli Shrimp Noodle Bowl LOWCAL":  16.8667,
    "Vermicelli Shrimp Noodle Bowl BUILD":   17.9905,
    "Lime Ginger Energy Bites LEAN":          3.91,
    "Lime Ginger Energy Bites LOWCAL":        3.91,
    "Low Carb Egg & Hash LEAN":               8.51,
    "Low Carb Egg & Hash LOWCAL":             9.98,
    "Low Carb Egg Wrap LEAN":                15.96,
    "Low Carb Egg Wrap":                     15.20,
    "Moroccan Chicken Bowl LOWCAL":          13.16,
    "Moroccan Chicken Bowl":                 13.16,
    "Moroccan Meatballs & Couscous LOWCAL":  20.06,
    "Grilled Chicken Burger":                12.90,
    "Nori Wraps BUILD":                      13.86,
    "Fish With Spinach & Sweet Potato Mash LOWCAL": 15.2667,
    "Fish With Spinach & Sweet Potato Mash LEAN":   15.2667,
    "(v) Southwest Pasta Salad LEAN":        14.80,

    # ── VALOREM HISTORICAL — ALL CAPS names from past orders (Valorem formatting) ─
    # These appeared in selections before Valorem corrected their order format.
    # Going forward, ALL CAPS dish names will trigger an error flag in the report.
    "BUFFALO CHICKEN BOWL LOWCAL":           16.91,
    "BUFFALO CHICKEN BOWL LEAN":             17.91,
    "BUFFALO CHICKEN BOWL BUILD":            19.91,
    "ITALIAN HERB CHUCK ROAST LOWCAL":       18.40,
    "ITALIAN HERB CHUCK ROAST LEAN":         20.3048,
    "ITALIAN HERB CHUCK ROAST BUILD":        22.3048,
    "LEAN BEEF BURGER":                      21.22,
    "LEAN BEEF BURGER LOWCAL":               19.90,
    "LEAN BEEF BURGER LEAN":                 21.22,
    "LEAN BEEF BURGER BUILD":                23.50,
}


# ─── INVOICE COLUMN STRUCTURE ─────────────────────────────────────────────────
# Each week block: type_col, dish_col, qty_col, price_col (1-indexed)
INVOICE_WEEK_COLS = [
    {"week": 1, "type_col": 1,  "dish_col": 2,  "qty_col": 3,  "price_col": 4},
    {"week": 2, "type_col": 7,  "dish_col": 8,  "qty_col": 9,  "price_col": 10},
    {"week": 3, "type_col": 13, "dish_col": 14, "qty_col": 15, "price_col": 16},
    {"week": 4, "type_col": 19, "dish_col": 20, "qty_col": 21, "price_col": 22},
    {"week": 5, "type_col": 25, "dish_col": 26, "qty_col": 27, "price_col": 28},
]
