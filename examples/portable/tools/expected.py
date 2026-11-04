"""Independent reference of the five examples: writes expected/<sheet> result.csv from data/*.csv.

    python3 examples/portable/tools/expected.py

run_lo.py compares what the scripts write in LibreOffice with these files. The rules follow the
comments of each script; rounding is half away from zero, as scripts/lib/table.vas does it.
"""
import csv, datetime, math, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)


def cell(v):
    """A CSV field as the sheet holds it: number, date serial number or text."""
    if re.fullmatch(r"-?\d+(\.\d+)?", v):
        return float(v)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return float((datetime.date.fromisoformat(v) - datetime.date(1899, 12, 30)).days)
    return v


def load(name):
    with open(os.path.join(BASE, "data", name + ".csv"), encoding="utf-8", newline="") as f:
        rows = [[cell(v) for v in r] for r in csv.reader(f)]
    header = rows[0]
    recs = []
    for i, r in enumerate(rows[1:], start=2):
        rec = dict(zip(header, r))
        rec["_row"] = i
        recs.append(rec)
    return recs


def round_to(x, d):
    f = 10.0 ** d
    if x < 0:
        return -math.floor(-x * f + 0.5) / f
    return math.floor(x * f + 0.5) / f


def r2(x):
    return round_to(x, 2)


def budget():
    items = load("Budget")
    chapters, amounts, counts, total = [], {}, {}, 0.0
    for it in items:
        a = it["Quantity"] * it["UnitPrice"]
        if it["Chapter"] not in amounts:
            chapters.append(it["Chapter"])
            amounts[it["Chapter"]] = 0.0
            counts[it["Chapter"]] = 0
        amounts[it["Chapter"]] += a
        counts[it["Chapter"]] += 1
        total += a
    out = [["Chapter", "Items", "Amount", "Share %"]]
    for ch in chapters:
        out.append([ch, counts[ch], r2(amounts[ch]), r2(100 * amounts[ch] / total)])
    out.append(["TOTAL", len(items), r2(total), 100])
    return out


def workreports():
    units = ["m", "m2", "m3", "kg", "h", "ud"]
    budget_unit = {b["Code"]: b["Unit"] for b in load("Budget")}
    out = [["Row", "Field", "Value", "Problem"]]
    for r in load("WorkReports"):
        rep = lambda f, p: out.append([r["_row"], f, r[f], p])
        if not isinstance(r["Date"], float):
            rep("Date", "not a date")
        elif r["Date"] % 7 == 1:
            rep("Date", "Sunday: needs an overtime authorisation")
        if not re.fullmatch(r"[A-Z]{2}[0-9]{4}", str(r["Code"])):
            rep("Code", "expected two capital letters and four digits")
        elif r["Code"] not in budget_unit:
            rep("Code", "not in the budget")
        if str(r["Unit"]).strip() == "":
            rep("Unit", "missing")
        elif r["Unit"] not in units:
            rep("Unit", "unknown unit, expected one of " + ", ".join(units))
        elif r["Code"] in budget_unit and budget_unit[r["Code"]] != r["Unit"]:
            rep("Unit", "the budget measures this item in " + budget_unit[r["Code"]])
        if not isinstance(r["Quantity"], float):
            rep("Quantity", "not a number")
        elif r["Quantity"] <= 0:
            rep("Quantity", "must be greater than zero")
        if str(r["Crew"]).strip() == "":
            rep("Crew", "missing")
    return out


def progress():
    out = [["Code", "Description", "ToDateQty", "Progress %", "MonthAmount", "ToDateAmount", "Status"]]
    bt = mt = tt = 0.0
    for r in load("Progress"):
        q = r["PreviousQty"] + r["MonthQty"]
        status = ""
        if q > r["BudgetQty"]:
            status = "OVER BUDGET"
        elif q == r["BudgetQty"]:
            status = "COMPLETED"
        elif r["MonthQty"] == 0:
            status = "NO PROGRESS"
        bt += r["BudgetQty"] * r["UnitPrice"]
        mt += r["MonthQty"] * r["UnitPrice"]
        tt += q * r["UnitPrice"]
        out.append([r["Code"], r["Description"], q, r2(100 * q / r["BudgetQty"]),
                    r2(r["MonthQty"] * r["UnitPrice"]), r2(q * r["UnitPrice"]), status])
    out.append(["TOTAL", "", "", r2(100 * tt / bt), r2(mt), r2(tt), ""])
    return out


def price_revision():
    a, b, c, d = 0.15, 0.40, 0.30, 0.15
    rows = load("PriceIndices")
    base = rows[0]
    out = [["Period", "Certified", "Kt", "Revised", "Adjustment"]]
    cert = rev = 0.0
    for r in rows[1:]:
        kt = round_to(a + b * r["Labour"] / base["Labour"] + c * r["Materials"] / base["Materials"]
                      + d * r["Energy"] / base["Energy"], 3)
        amount = r2(r["Certified"] * kt)
        cert += r["Certified"]
        rev += amount
        out.append([r["Period"], r["Certified"], kt, amount, r2(amount - r["Certified"])])
    out.append(["TOTAL", r2(cert), "", r2(rev), r2(rev - cert)])
    return out


def cleanup():
    units = {"ml": "m", "mts": "m", "un": "ud", "u": "ud", "pcs": "ud", "kgs": "kg", "hr": "h", "hrs": "h"}
    out = [["Code", "Unit", "Description", "Changes"]]
    for r in load("Cleanup"):
        raw = re.sub(r"\s+", " ", str(r["Code"])).strip()
        m = re.fullmatch(r"([A-Za-z]{2})[-_ .]?0*([0-9]{1,4})", raw)
        code = m.group(1).upper() + m.group(2).rjust(4, "0") if m else str(r["Code"])
        u = str(r["Unit"]).strip().lower().replace("²", "2").replace("³", "3").replace(".", "")
        unit = units.get(u, u)
        t = re.sub(r"\s+", " ", str(r["Description"])).strip()
        text = t[:1].upper() + t[1:]
        changes = [n for n, new, old in (("code", code, r["Code"]), ("unit", unit, r["Unit"]),
                                         ("description", text, r["Description"])) if new != str(old)]
        out.append([code, unit, text, ", ".join(changes)])
    return out


EXAMPLES = {"Budget": budget, "WorkReports": workreports, "Progress": progress,
            "PriceIndices": price_revision, "Cleanup": cleanup}


def main():
    os.makedirs(os.path.join(BASE, "expected"), exist_ok=True)
    for sheet, fn in EXAMPLES.items():
        with open(os.path.join(BASE, "expected", sheet + " result.csv"), "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f, lineterminator="\n")
            for row in fn():
                w.writerow([("%.10g" % v) if isinstance(v, float) else v for v in row])
        print(sheet, "result.csv")


if __name__ == "__main__":
    main()
