"""Sample tables of the use cases and what they must give: one source for make_workbooks.py, which
writes the workbooks, and run_lo.py, which checks the scripts in LibreOffice."""
from collections import defaultdict
import math

# Use case 1: rules users edit
RULES = [
    ["Name", "When", "Discount"],
    ["Volume", "o.qty >= 100", 0.08],
    ["Loyal EU customer", "o.region == 'EU' && o.years >= 3", 0.05],
    ["Clearance", "o.sku.startsWith('CL-')", 0.15],
    ["Big ticket", "o.qty * o.price > 5000", 0.10],
]
ORDERS = [
    ["order", "region", "years", "sku", "qty", "price"],
    ["A-101", "EU", 4, "ST-200", 20, 35],
    ["A-102", "US", 1, "CL-310", 5, 12],
    ["A-103", "EU", 1, "ST-200", 150, 35],
    ["A-104", "US", 6, "ST-900", 60, 120],
    ["A-105", "EU", 3, "CL-110", 200, 3],
]
# the best discount that applies to each order (checked by hand against RULES)
DISCOUNTS = [
    ["Order", "Discount", "Rule"],
    ["A-101", 0.05, "Loyal EU customer"],
    ["A-102", 0.15, "Clearance"],
    ["A-103", 0.10, "Big ticket"],
    ["A-104", 0.10, "Big ticket"],
    ["A-105", 0.15, "Clearance"],
]
# a mistake in a rule: what the user writes and the message they get
RULE_MISTAKES = [
    ("o.qty >= (100", "Rule 'Volume': Unbalanced opening or closing parentheses."),
    ("o.qty >=", "Rule 'Volume': Expected an expression after >="),
    ("o.skuu.startsWith('CL-')", "Rule 'Volume' uses o.skuu, but Orders has no column 'skuu'"),
    ("o.qtty >= 100", "Rule 'Volume' uses o.qtty, but Orders has no column 'qtty'"),
]

# Use case 3: regular expressions everywhere (results checked against JavaScript)
REGEX = [
    ('RX_TEST', ["ES-0042", r"^[A-Z]{2}-\d{4}$"], True),
    ('RX_TEST', ["es-0042", r"^[A-Z]{2}-\d{4}$"], False),
    ('RX_TEST', ["es-0042", r"^[A-Z]{2}-\d{4}$", True], True),
    ('RX_EXTRACT', ["Pallet: 12 boxes, 340 kg", r"(\d+)\s*kg", 1], "340"),
    ('RX_EXTRACT', ["Pallet: 12 boxes, 340 kg", r"\d+"], "12"),
    ('RX_EXTRACT', ["no numbers", r"\d+"], ""),
    ('RX_EXTRACT', ["Order #A-77 shipped", r"(?<=#)[A-Z]-\d+"], "A-77"),
    ('RX_EXTRACT', ["2026-11-15", r"(?<year>\d{4})-(?<month>\d{2})", 2], "11"),
    ('RX_REPLACE', ["  too    many   spaces ", r"\s+", " "], " too many spaces "),
    ('RX_REPLACE', ["Smith, John", r"^(\w+),\s*(\w+)$", "$2 $1"], "John Smith"),
    ('RX_COUNT', ["one two  three", r"\b\w+\b"], 3),
    ('RX_TEST', ["x", "(unclosed"], "#VALUE!"),
]

# Use case 5: a sales table, three regions, four products, three months
def sales():
    regions, products, months = ["North", "South", "East"], ["Bolts", "Nuts", "Washers", "Brackets"], ["2026-01", "2026-02", "2026-03"]
    rows = [["region", "product", "month", "units", "price"]]
    for i in range(2, 38):
        rows.append([regions[(i - 2) % 3], products[(i - 2) % 4], months[((i - 2) // 12) % 3],
                     10 + (i * 7) % 23, 1 + ((i * 3) % 5) * 0.5])
    return rows


def round2(x):
    return math.floor(x * 100 + 0.5) / 100


def sales_summary(rows):
    """What sales-summary.vas must give, computed here on its own."""
    total, by_region, by_month = 0.0, {}, {}
    for region, product, month, units, price in rows[1:]:
        revenue = units * price
        total += revenue
        g = by_region.setdefault(region, [0, 0.0, defaultdict(float)])
        g[0] += units
        g[1] += revenue
        g[2][product] += revenue
        by_month[month] = by_month.get(month, 0.0) + revenue
    regions = [[r, u, round2(rev), round2(100 * rev / total), max(ps, key=lambda p: (ps[p], -list(ps).index(p)))]
               for r, (u, rev, ps) in by_region.items()]
    regions.sort(key=lambda r: -r[3])
    months, previous = [], 0
    for m in sorted(by_month):
        months.append([m, round2(by_month[m]), "" if previous == 0 else round2(100 * (by_month[m] - previous) / previous)])
        previous = by_month[m]
    return ([["Region", "Units", "Revenue", "Share %", "Top product"]] + regions,
            [["Month", "Revenue", "Change %"]] + months)
