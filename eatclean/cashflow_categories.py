"""
Vendor/description pattern -> 2026 P&L bucket mapping.

Buckets mirror the March 2026 income statement structure.
Each rule: (regex pattern on description, bucket_group, bucket_line, sign_hint)
  sign_hint: "in" (inflow expected), "out" (outflow expected), None (either)

Rules are evaluated in order. First match wins.
Unmatched rows are routed to "UNCATEGORIZED" for manual review.
"""

import re

# (pattern, group, line, sign_hint)
RULES = [
    # ----- REVENUE (inflows) -----
    # Network International is the Stripe settlement processor for UAE merchants.
    # Every "From NETWORK INTERNATIONAL LLC" inflow on this account is Stripe revenue.
    (r"NETWORK INTERNATIONAL", "Revenue", "Stripe", "in"),
    (r"\bSTRIPE\b", "Revenue", "Stripe", "in"),
    (r"MEALPLANET|Meal\s*Planet", "Revenue", "Meal Planet", "in"),
    (r"Tipalti|ClassPass|Class Pass", "Revenue", "Class Pass", "in"),
    (r"\bTabby\b",   "Revenue", "Tabby", "in"),
    (r"Pointspay",   "Revenue", "Pointspay", "in"),

    # ----- COGS (outflows) -----
    (r"Basiligo|BASILIGO|BSI\d+",            "COGS", "COGS Basiligo", "out"),
    (r"L O G X|LOGX|LGX\d+",                 "COGS", "COGS Delivery fee LogX", "out"),
    (r"Quick Pack|QuickPack",                "COGS", "COGS Packaging", "out"),
    (r"Fanan|Sticker",                       "COGS", "COGS Stickers", "out"),

    # ----- SALARIES -----
    # UAE employees (confirmed per name). Recruited / terminated per user notes.
    (r"MARYFEL SANTANDER|Maryfel Santander",          "Salaries", "Salaries and Wages UAE", "out"),
    (r"Jona Maureen Valerie",                          "Salaries", "Salaries and Wages UAE", "out"),
    (r"Nidheesh Ambika Krishnan",                      "Salaries", "Salaries and Wages UAE", "out"),
    (r"John Kenneth Jereza Ong",                       "Salaries", "Salaries and Wages UAE", "out"),
    # UpWork: remote contractors
    (r"Upwork ref|Upwork\b",                 "Salaries", "Salaries and Wages Remote", "out"),
    (r"Medical Insurance|\bInsurance\b",     "Salaries", "Insurance expenses", "out"),
    (r"Leave salary",                        "Salaries", "Leave Salary", "out"),
    (r"\bGratuity\b",                        "Salaries", "Gratuities", "out"),

    # ----- ADVERTISING & MARKETING -----
    (r"FACEBK|FACEBOOK|META ADS|Meta\*",     "Advertising", "Facebook/Instagram Expenses", "out"),
    (r"GOOGLE\*ADS|Google Ads|GOOGLE ADS",   "Advertising", "Google Ads expenses", "out"),
    (r"Adsquiz",                             "Advertising", "Other advertising", "out"),
    (r"Dubizzle",                            "Advertising", "Other advertising", "out"),

    # ----- RECRUITING (per user: LinkedIn = recruiting cost -> Other expenses) -----
    (r"Linkedin|LinkedIn",                   "Other expenses", "Other", "out"),

    # ----- MISC MARKETING -----
    (r"\bDesco\b",                           "Advertising", "Other advertising", "out"),

    # ----- CRM + PLATFORMS -----
    (r"Altconnect|Nutribot",                 "CRM + Platforms", "Nutribot", "out"),
    (r"Shopify",                             "CRM + Platforms", "Shopify + website", "out"),
    (r"Klaviyo",                             "CRM + Platforms", "Shopify + website", "out"),
    (r"Zoho",                                "CRM + Platforms", "Suscription fee", "out"),
    (r"GoDaddy|godaddy",                     "CRM + Platforms", "GoDaddy", "out"),
    (r"Google Workspace|Google Space",       "CRM + Platforms", "Google Space", "out"),
    (r"Predis",                              "CRM + Platforms", "Suscription fee", "out"),
    (r"Supliful",                            "CRM + Platforms", "Shopify + website", "out"),
    (r"Anthropic",                           "CRM + Platforms", "Suscription fee", "out"),

    # ----- SALES FEES -----
    # Wio doesn't typically show Stripe/MealPlanet fees as line items (they net before payout).
    # But any direct charge marked as a fee from these providers:
    (r"Stripe Fee",                          "Sales Fees", "Stripe Fees", "out"),
    (r"Meal Planet Fee",                     "Sales Fees", "Meal Planet Fees", "out"),
    (r"ClassPass Fee",                       "Sales Fees", "ClassPass Fees", "out"),
    (r"Pointspay Fee",                       "Sales Fees", "Pointspay Fees", "out"),
    (r"Tabby Fee",                           "Sales Fees", "Tabby Fees", "out"),

    # ----- ACCOUNTING -----
    (r"Ahsan Accountiing|Ahsan Accounting",  "Accounting", "Audit Fee", "out"),
    (r"Accountant|Accounting",               "Accounting", "Salary - Accountant", "out"),
    (r"Audit",                               "Accounting", "Audit Fee", "out"),

    # ----- FINANCE COSTS (bank fees, FX fees) -----
    (r"Foreign exchange transaction fee",    "Finance costs", "Bank Fees and Charges", "out"),
    (r"Subscription fee for \w+ 2026",       "Finance costs", "Bank Fees and Charges", "out"),  # Wio Grow plan
    (r"Wio|Bank Fee|Monthly fee|card fee",   "Finance costs", "Bank Fees and Charges", "out"),

    # ----- VAT / TAXES (cash flow only, not in P&L as an expense) -----
    (r"Federal Tax Authority|\bFTA\b",       "VAT", "VAT payments", "out"),

    # ----- LEGAL / ADMIN -----
    (r"Rent|Office Rent",                    "Legal costs", "Rent Expense", "out"),
    (r"Trade License|DSO.*License",          "Legal costs", "Trade license expenses", "out"),
    (r"Dubai integrated econo|DIEZ",         "Legal costs", "Trade license expenses", "out"),

    # ----- OTHER EXPENSES -----
    (r"Consultancy",                         "Other expenses", "Consultancy charges", "out"),
    (r"Partner.*Salar",                      "Other expenses", "Partner'sSalaries", "out"),
    (r"Travel|Flight|Hotel",                 "Other expenses", "Travel expenses", "out"),
    (r"Donation|One tree",                   "Other expenses", "Donation expense", "out"),
]


def categorize(description: str, amount: float) -> tuple[str, str]:
    """
    Return (bucket_group, bucket_line) for a transaction.
    If no pattern matches, returns ("UNCATEGORIZED", sign-hint string).
    """
    desc = description or ""
    for pat, group, line, _sign in RULES:
        if re.search(pat, desc, flags=re.IGNORECASE):
            return group, line
    tag = "UNCATEGORIZED (inflow)" if amount > 0 else "UNCATEGORIZED (outflow)"
    return "UNCATEGORIZED", tag


# The ordered P&L bucket structure for report rendering.
BUCKET_ORDER = [
    ("Revenue", [
        "Stripe", "Meal Planet", "Class Pass", "Tabby", "Pointspay", "Sales - local",
    ]),
    ("COGS", [
        "COGS Basiligo", "COGS Delivery fee LogX", "COGS Packaging", "COGS Stickers",
    ]),
    ("Salaries", [
        "Salaries and Wages UAE", "Salaries and Wages Remote",
        "Insurance expenses", "Visa expenses", "Leave Salary",
        "Employee benefits", "Gratuities",
    ]),
    ("Advertising", [
        "Agency Retainer Fee", "Google Ads expenses",
        "Facebook/Instagram Expenses", "Other advertising",
        "Commission for EC Refferal",
    ]),
    ("CRM + Platforms", [
        "Nutribot", "Shopify + website", "Suscription fee",
        "Google Space", "GoDaddy", "IT & Internet ",
    ]),
    ("Sales Fees", [
        "Stripe Fees", "Meal Planet Fees", "ClassPass Fees",
        "Pointspay Fees", "Tabby Fees",
    ]),
    ("Accounting", [
        "Salary - Accountant", "Audit Fee",
    ]),
    ("Finance costs", [
        "Bank Fees and Charges",
    ]),
    ("Legal costs", [
        "Rent Expense", "Trade license expenses",
    ]),
    ("VAT", [
        "VAT payments",
    ]),
    ("Other expenses", [
        "Consultancy charges", "Partner'sSalaries", "Travel expenses",
        "Donation expense", "Other",
    ]),
    ("UNCATEGORIZED", [
        "UNCATEGORIZED (inflow)", "UNCATEGORIZED (outflow)",
    ]),
]
