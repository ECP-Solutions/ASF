"""The sample data of PythonExamples.xlsx, the same on every run (a fixed seed)."""
import datetime
import random

REGIONS = ['North', 'South', 'East', 'West']
PRODUCTS = ['Widget', 'Gadget', 'Gizmo']
PRICES = {'Widget': 12.5, 'Gadget': 40.0, 'Gizmo': 7.25}


def sales(n=60):
    """Orders: Order, Date, Region, Product, Qty, Price, Amount. Two outsized orders on purpose."""
    rnd = random.Random(7)
    start = datetime.date(2026, 1, 5)
    rows = [['Order', 'Date', 'Region', 'Product', 'Qty', 'Price', 'Amount']]
    for i in range(n):
        product = rnd.choice(PRODUCTS)
        qty = rnd.randint(1, 20)
        if i in (17, 44):
            qty = 180 + i                  # the outliers
        price = PRICES[product]
        rows.append(['SO-%04d' % (1001 + i), start + datetime.timedelta(days=i * 3), rnd.choice(REGIONS),
                     product, qty, price, round(qty * price, 2)])
    return rows


def monthly():
    """24 months of a series with a trend and a yearly pattern: Month, Visitors."""
    rnd = random.Random(11)
    season = [0.85, 0.8, 0.95, 1.0, 1.05, 1.15, 1.25, 1.2, 1.05, 1.0, 0.95, 1.1]
    return [['Month', 'Visitors']] + [[m + 1, round((1000 + 25 * m) * season[m % 12] + rnd.uniform(-30, 30))]
                                      for m in range(24)]


def regions():
    """Region | Manager | Target: a table to join to the sales."""
    return [['Region', 'Manager', 'Target'], ['North', 'Ana Ruiz', 3000], ['South', 'Ben Okafor', 2500],
            ['East', 'Chen Wei', 2800], ['West', 'Dana Kowalski', 3200]]


def customers():
    """Names as typed (case, spaces, accents) and a clean list to match them against."""
    typed = ['  josé  müller', "o'brien SEAN", 'mcdonald, ronald', 'Ana de la ROSA', 'acme corp.', 'Globex Corporation',
             'initech llc', 'Umbrela Corp']
    clean = ['Acme Corporation', 'Globex Corporation', 'Initech LLC', 'Umbrella Corporation', 'Stark Industries']
    return typed, clean
