"""Times every stage of the pipeline in LibreOffice on 1,000 orders made from the sample ones.

    python3 tools/timing.py [ASF repository]
"""
import os, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import data  # noqa: E402
from lo import Office  # noqa: E402

root = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "..", "..")
body = [r for r in data.RAW[1:] if r[0]]
raw = [data.RAW[0]] + [[f"SO-{2000 + i}"] + list(body[i % len(body)][1:]) for i in range(1000)]
with Office(root) as call:
    t = [time.time()]
    clean = call("ASF_CALL", "clean", raw, data.CLEAN_SPEC); t.append(time.time())
    unique = call("ASF_CALL", "dedupe", clean, "Order", "last"); t.append(time.time())
    valid = call("ASF_CALL", "validate", unique, data.RULES, "valid"); t.append(time.time())
    enriched = call("ASF_CALL", "lookup", valid, data.CUSTOMERS, "Customer=CustomerID", "Name, Region, Segment"); t.append(time.time())
    sales = call("ASF_RUN", data.AMOUNT_SCRIPT, enriched); t.append(time.time())
    call("ASF_CALL", "group", sales, "Region", "sum:Amount, count:Order, avg:Amount, distinct:Customer, min:Date, max:Date"); t.append(time.time())
    call("ASF_CALL", "pivot", sales, "Region", "Product", "sum:Amount"); t.append(time.time())
    call("ASF_CALL", "profile", raw); t.append(time.time())
    for i, name in enumerate(["clean", "dedupe", "validate", "lookup", "amount", "group", "pivot", "profile"]):
        print(f"{name:9s} {t[i + 1] - t[i]:6.2f} s")
    print(f"total     {t[-1] - t[0]:6.2f} s")
