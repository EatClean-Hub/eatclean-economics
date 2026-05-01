import pandas as pd
from config import DELIVERY_PER_DAY, PACKAGING_PER_DAY, TARGET_MARGIN, BENCH_MAX


# ══════════════════════════════════════════════════════════════════════════════
# RECONCILIATION — Invoice vs Selections
# ══════════════════════════════════════════════════════════════════════════════

def reconcile(df_invoice: pd.DataFrame, df_selections: pd.DataFrame) -> pd.DataFrame:
    """
    Compare invoice quantities vs selection quantities per day per dish.
    Only for dates present in both datasets.

    Returns DataFrame with columns:
    date | dish_name | invoice_qty | selection_qty | diff | diff_pct | flag
    """
    if df_invoice.empty or df_selections.empty:
        return pd.DataFrame()

    # Aggregate invoice: date × dish → total qty
    inv = (
        df_invoice[df_invoice["bench_flag"] != "SKIP"]
        .groupby(["date", "dish_name"])["qty"]
        .sum()
        .reset_index()
        .rename(columns={"qty": "invoice_qty"})
    )

    # Aggregate selections: date × dish → total qty
    sel = (
        df_selections[df_selections["group"].isin(["regular", "valorem"])]
        .groupby(["date", "dish_name"])["qty"]
        .sum()
        .reset_index()
        .rename(columns={"qty": "selection_qty"})
    )

    # Only reconcile dates that exist in BOTH
    common_dates = set(inv["date"].dt.date) & set(sel["date"].dt.date)
    inv = inv[inv["date"].dt.date.isin(common_dates)]
    sel = sel[sel["date"].dt.date.isin(common_dates)]

    merged = pd.merge(inv, sel, on=["date", "dish_name"], how="outer").fillna(0)
    merged["diff"]     = merged["invoice_qty"] - merged["selection_qty"]
    merged["diff_pct"] = merged.apply(
        lambda r: round(r["diff"] / r["selection_qty"] * 100, 1)
        if r["selection_qty"] > 0 else None,
        axis=1
    )
    merged["flag"] = merged.apply(_reconcile_flag, axis=1)
    return merged.sort_values(["date", "flag", "dish_name"])


def _reconcile_flag(row) -> str:
    if row["diff"] == 0:
        return "OK"
    if row["selection_qty"] == 0:
        return "IN INVOICE ONLY"
    if row["invoice_qty"] == 0:
        return "IN SELECTIONS ONLY"
    if abs(row["diff"]) > 2:
        return "MISMATCH"
    return "MINOR DIFF"


# ══════════════════════════════════════════════════════════════════════════════
# MARGIN MONITOR — Per Order
# ══════════════════════════════════════════════════════════════════════════════

def calculate_margins(df_orders: pd.DataFrame,
                      df_invoice: pd.DataFrame) -> pd.DataFrame:
    """
    For each order, calculate actual margin using invoice cost for delivered days.
    Returns enriched orders DataFrame with cost + margin columns.
    """
    if df_invoice.empty:
        return df_orders.copy()

    # Daily kitchen cost per person from invoice (using selection data for attribution)
    # Since invoice is aggregate, use total daily invoice cost ÷ customer count as proxy
    # TODO: Replace with per-customer cost once customer-level invoice mapping is built
    daily_invoice_total = (
        df_invoice.groupby("date")["line_total_ex_vat"]
        .sum()
        .reset_index()
        .rename(columns={"line_total_ex_vat": "daily_invoice_total"})
    )

    results = []
    for _, order in df_orders.iterrows():
        start = order["Order date from"]
        end   = order["Order date to"]

        # Filter invoice days within this order's period
        order_days = daily_invoice_total[
            (daily_invoice_total["date"] >= start) &
            (daily_invoice_total["date"] <= end)
        ]

        days_with_data = len(order_days)
        net_rev        = order["net_revenue"]
        paid_days      = order["Delivery period"]

        row = order.to_dict()
        row["days_with_data"] = days_with_data

        # We need per-customer cost — using selections-based cost if available
        # For now flag orders where we have data
        row["data_coverage"] = (
            "✅ Full" if days_with_data >= paid_days
            else f"⏳ {days_with_data}/{paid_days} days"
        )
        row["margin_flag"] = "⬜ No data"
        results.append(row)

    return pd.DataFrame(results)


def calculate_order_margins_from_selections(
        df_orders: pd.DataFrame,
        df_selections: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate per-order margin using selections-based daily cost.
    This is the primary margin calculation.
    """
    if df_selections.empty:
        return df_orders.copy()

    # Daily cost per person
    daily_cost_per_person = (
        df_selections.groupby(["person", "date"])["line_cost"]
        .sum()
        .reset_index()
        .rename(columns={"line_cost": "day_kitchen_cost"})
    )

    from parsers import _normalize
    results = []

    for _, order in df_orders.iterrows():
        person_norm = order["name_norm"]
        start       = order["Order date from"]
        end         = order["Order date to"]
        net_rev     = float(order["net_revenue"])
        paid_days   = int(order["Delivery period"])

        # Find matching person in selections
        sel_persons = daily_cost_per_person.copy()
        sel_persons["person_norm"] = sel_persons["person"].apply(_normalize)
        person_data = sel_persons[
            (sel_persons["person_norm"] == person_norm) &
            (sel_persons["date"] >= start) &
            (sel_persons["date"] <= end)
        ]

        days_delivered  = len(person_data)
        total_kitchen   = person_data["day_kitchen_cost"].sum()
        delivery_cost   = days_delivered * DELIVERY_PER_DAY
        packaging_cost  = days_delivered * PACKAGING_PER_DAY
        total_cost      = total_kitchen + delivery_cost + packaging_cost
        margin_aed      = net_rev - total_cost
        margin_pct      = round(margin_aed / net_rev * 100, 1) if net_rev > 0 else None

        row = order.to_dict()
        row["days_delivered"]   = days_delivered
        row["kitchen_cost"]     = round(total_kitchen, 2)
        row["delivery_cost"]    = round(delivery_cost, 2)
        row["packaging_cost"]   = round(packaging_cost, 2)
        row["total_cost"]       = round(total_cost, 2)
        row["margin_aed"]       = round(margin_aed, 2)
        row["margin_pct"]       = margin_pct
        row["data_coverage"]    = (
            "✅ Full" if days_delivered >= paid_days
            else f"⏳ {days_delivered}/{paid_days} days"
        )
        row["margin_flag"] = (
            "⬜ No data"     if days_delivered == 0
            else f"🔴 {margin_pct}%" if (margin_pct is not None and margin_pct < TARGET_MARGIN)
            else f"✅ {margin_pct}%"
        )
        results.append(row)

    return pd.DataFrame(results)


# ══════════════════════════════════════════════════════════════════════════════
# DISH PRICE VIOLATIONS
# ══════════════════════════════════════════════════════════════════════════════

def get_violations(df_invoice: pd.DataFrame) -> pd.DataFrame:
    """Return all invoice lines where dish price exceeds benchmarks."""
    return df_invoice[
        df_invoice["bench_flag"].isin(["EXCEEDS ALL", "CHECK"])
    ].copy().sort_values(["bench_flag", "date", "meal_type"])


# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════

def dishes_ranking(df_selections: pd.DataFrame) -> pd.DataFrame:
    """Rank dishes by total quantity ordered."""
    if df_selections.empty:
        return pd.DataFrame()
    return (
        df_selections[df_selections["group"].isin(["regular", "valorem"])]
        .groupby("dish_name")
        .agg(
            total_qty   =("qty",       "sum"),
            days_ordered=("date",      "nunique"),
            total_cost  =("line_cost", "sum"),
        )
        .sort_values("total_qty", ascending=False)
        .reset_index()
    )


def weekly_summary(df_invoice: pd.DataFrame) -> pd.DataFrame:
    """Weekly invoice totals."""
    if df_invoice.empty:
        return pd.DataFrame()
    df = df_invoice.copy()
    df["week_start"] = df["date"] - pd.to_timedelta(
        df["date"].dt.dayofweek, unit="d"
    )
    return (
        df.groupby(["week_start", "month"])
        .agg(
            total_ex_vat     =("line_total_ex_vat", "sum"),
            lines            =("dish_name",         "count"),
            violations_exceed=("bench_flag", lambda x: (x == "EXCEEDS ALL").sum()),
            violations_check =("bench_flag", lambda x: (x == "CHECK").sum()),
        )
        .round(2)
        .reset_index()
        .sort_values("week_start")
    )


def monthly_summary(df_invoice: pd.DataFrame) -> pd.DataFrame:
    """Monthly invoice totals."""
    if df_invoice.empty:
        return pd.DataFrame()
    return (
        df_invoice.groupby("month")
        .agg(
            total_ex_vat     =("line_total_ex_vat", "sum"),
            lines            =("dish_name",         "count"),
            violations_exceed=("bench_flag", lambda x: (x == "EXCEEDS ALL").sum()),
            violations_check =("bench_flag", lambda x: (x == "CHECK").sum()),
        )
        .round(2)
        .reset_index()
    )
