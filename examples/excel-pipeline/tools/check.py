"""Checks the pipeline in LibreOffice, headless: every stage of the workbook through ASF_CALL and
ASF_RUN, each fed with what the stage before it gave, against ref.py fed with its own results.

    python3 tools/check.py [ASF repository]

The repository (default: three folders up, examples/excel-pipeline/tools in it) gives the engine
and the add-in's modules; the library folder is this pipeline's library/. Needs LibreOffice and
python3-uno. Exits with 0 when every stage matches."""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import data  # noqa: E402
import ref  # noqa: E402
from lo import Office  # noqa: E402


def norm(v):
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return round(float(v), 9)
    return v


def table(t):
    if isinstance(t, (list, tuple)) and t and isinstance(t[0], (list, tuple)):
        return [[norm(v) for v in row] for row in t]
    return norm(t)


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "..", "..")
    failures = 0

    def check(what, got, expected):
        nonlocal failures
        ok = table(got) == table(expected)
        failures += not ok
        print(f"{'ok' if ok else 'DIFFERS':8s} {what}")
        if not ok:
            g, e = table(got), table(expected)
            if isinstance(g, list) and isinstance(e, list):
                for i, (a, b) in enumerate(zip(g, e)):
                    if a != b:
                        print(f"         row {i + 1}: expected {b}\n                  got      {a}")
                if len(g) != len(e):
                    print(f"         {len(e)} rows expected, {len(g)} got")
            else:
                print(f"         expected {e!r}\n         got      {g!r}")

    with Office(root) as call:
        # the readers of numbers and dates
        num = "import { toNumber } from './tables.vas'; return toNumber($1);"
        for v in ["$1,250.00", "1.250,50", "(3.00)", "-12", "15%", " 7.25 ", "12,5", "1,250", "1.234.567",
                  "€ 9,99", "+4", "abc", "1.5.2", "--1", 42]:
            want = ref.to_number(v)
            check(f"number {v!r}", call("ASF_RUN", num, v), "" if want is None else want)
        date = "import { toDate } from './tables.vas'; return toDate($1, $2);"
        for v, order in [("2026-03-05", ""), ("2026/3/5", ""), ("05/03/2026", "dmy"), ("05/03/2026", "mdy"),
                         ("31/02/2026", "dmy"), ("29/02/2024", "dmy"), ("5-3-26", "dmy"), ("5.3.2026", "dmy"),
                         ("2026-03-05 10:30", ""), (46086, ""), ("March 5", "")]:
            want = ref.to_date(v, order or "dmy")
            check(f"date {v!r} {order}", call("ASF_RUN", date, v, order), "" if want is None else want)

        # the workbook's chain: each stage gets what the one before gave
        raw, customers, rules = data.RAW, data.CUSTOMERS, data.RULES
        clean = call("ASF_CALL", "clean", raw, data.CLEAN_SPEC)
        check("1 clean", clean, ref.clean(raw, data.CLEAN_SPEC))
        unique = call("ASF_CALL", "dedupe", clean, "Order", "last")
        check("2 dedupe", unique, ref.dedupe(clean, "Order", "last"))
        issues = call("ASF_CALL", "validate", unique, rules)
        check("3 validate: issues", issues, ref.validate(unique, rules))
        valid = call("ASF_CALL", "validate", unique, rules, "valid")
        check("4 validate: valid rows", valid, ref.validate(unique, rules, "valid"))
        flagged = call("ASF_CALL", "validate", unique, rules, "flagged")
        check("  validate: flagged", flagged, ref.validate(unique, rules, "flagged"))
        enriched = call("ASF_CALL", "lookup", valid, customers, "Customer=CustomerID", "Name, Region, Segment")
        check("5 lookup", enriched, ref.lookup(valid, customers, "Customer=CustomerID", "Name, Region, Segment"))
        sales = call("ASF_RUN", data.AMOUNT_SCRIPT, enriched)
        check("6 amount (ASF_RUN, script in a cell)", sales, ref.amount(enriched))
        measures = "sum:Amount, count:Order, avg:Amount, distinct:Customer, min:Date, max:Date"
        check("7 group by region", call("ASF_CALL", "group", sales, "Region", measures), ref.group(sales, "Region", measures))
        check("  group by channel and segment", call("ASF_CALL", "group", sales, "Channel, Segment", "sum:Amount, count:*"),
              ref.group(sales, "Channel, Segment", "sum:Amount, count:*"))
        check("  group: the whole table", call("ASF_CALL", "group", sales, "", "sum:Amount, count:*"), ref.group(sales, "", "sum:Amount, count:*"))
        check("8 pivot region x product", call("ASF_CALL", "pivot", sales, "Region", "Product", "sum:Amount"),
              ref.pivot(sales, "Region", "Product", "sum:Amount"))
        check("  pivot count", call("ASF_CALL", "pivot", sales, "Segment", "Channel"), ref.pivot(sales, "Segment", "Channel"))
        check("9 profile of the raw orders", call("ASF_CALL", "profile", raw), ref.profile(raw))
        check("  top.vas, the example of the README", call("ASF_CALL", "top", sales, "Amount", 5), ref.top(sales, "Amount", 5))

        # a stage reading the cells of an array formula: the table padded with empty rows and columns
        padded = [list(row) + ["", ""] for row in sales] + [[""] * (len(sales[0]) + 2)] * 20
        check("  padded like an array formula's cells", call("ASF_CALL", "group", padded, "Region", measures), ref.group(sales, "Region", measures))

        # mistakes in the arguments give #VALUE! with a message that names them
        for args, message in [(("clean", raw, "Qty:numbr"), "Unknown kind 'numbr'"),
                              (("clean", raw, "Quantity:number"), "No column 'Quantity'"),
                              (("group", sales, "Region", "total:Amount"), "Unknown measure 'total'"),
                              (("lookup", valid, customers, "Client"), "No column 'Client'"),
                              (("validate", unique, [["Column", "Rule", "Value"], ["Qty", "between", 1]]), "Unknown rule 'between'"),
                              (("dedupe", [["#VALUE!", "#VALUE!"], ["#VALUE!", "#VALUE!"]], "Order"), "the stage before it failed")]:
            got = call("ASF_CALL", *args)
            check(f"mistake: {message}", isinstance(got, str) and got.startswith("#ERR:#VALUE!") and message in got, True)
    print("all stages match" if not failures else f"{failures} check(s) differ")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
