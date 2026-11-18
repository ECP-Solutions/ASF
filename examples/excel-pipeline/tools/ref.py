"""The pipeline's stages written again in Python, from their descriptions: what check.py compares the
scripts of the library with. A table is a list of rows, the first the header; an empty cell is None
or ''."""
import datetime
import math
import re


def js_text(v):
    """A value as ASF writes it in a template."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def js_round(x, places):
    f = 10 ** places
    return math.floor(x * f + 0.5) / f


def trim_table(t):
    keep = [c for c, h in enumerate(t[0]) if not blank(h)]
    out = []
    for row in t:
        cells = [row[c] for c in keep]
        if not all(blank(v) for v in cells):
            out.append(cells)
    return out


def tidy(s):
    if "  " in s or "\t" in s or "\n" in s or "\r" in s:
        s = re.sub(r"\s+", " ", s)
    return s.strip(" ")


def index_of(t):
    return {js_text(h).strip().lower(): c for c, h in enumerate(t[0])}


def col(index, name):
    return index[name.strip().lower()]


def list_of(spec):
    return [p.strip() for p in js_text(spec).split(",") if p.strip()]


def to_number(v):
    if is_number(v):
        return float(v)
    if not isinstance(v, str):
        return None
    s = re.sub(r"[\s$€£¥]", "", v)
    if s == "":
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]
    if s.startswith("-"):
        negative, s = not negative, s[1:]
    elif s.startswith("+"):
        s = s[1:]
    percent = s.endswith("%")
    if percent:
        s = s[:-1]
    dot, comma = s.rfind("."), s.rfind(",")
    if dot >= 0 and comma >= 0:
        s = s.replace(".", "").replace(",", ".") if comma > dot else s.replace(",", "")
    elif comma >= 0:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3}){2,}", s):
        s = s.replace(".", "")
    if not re.fullmatch(r"(\d+(\.\d*)?|\.\d+)", s):
        return None
    n = float(s)
    if percent:
        n /= 100
    return -n if negative else n


def serial(y, m, d):
    try:
        return float((datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days)
    except ValueError:
        return None


def to_date(v, order):
    if is_number(v):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip()
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})([ T].*)?", s)
    if m:
        return serial(int(m[1]), int(m[2]), int(m[3]))
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})([ T].*)?", s)
    if not m:
        return None
    y = int(m[3]) + (2000 if len(m[3]) == 2 else 0)
    return serial(y, int(m[1]), int(m[2])) if order == "mdy" else serial(y, int(m[2]), int(m[1]))


def clean(t, spec):
    t = trim_table(t)
    index = index_of(t)
    kinds = {}
    for item in list_of(spec):
        name, kind = item.split(":", 1)
        kinds[col(index, name)] = kind.strip().lower()
    out = [[tidy(h) if isinstance(h, str) else h for h in t[0]]]
    for row in t[1:]:
        cells = []
        for c, v in enumerate(row):
            if isinstance(v, str):
                v = tidy(v)
            if c in kinds:
                v = convert(v, kinds[c])
            cells.append(v)
        out.append(cells)
    return out


def convert(v, kind):
    if blank(v):
        return ""
    if kind == "number":
        p = to_number(v)
    elif kind in ("date", "date-dmy"):
        p = to_date(v, "dmy")
    elif kind == "date-mdy":
        p = to_date(v, "mdy")
    elif kind == "lower":
        return js_text(v).lower()
    elif kind == "upper":
        return js_text(v).upper()
    elif kind == "title":
        return " ".join(w[:1].upper() + w[1:] for w in js_text(v).lower().split(" "))
    else:
        return js_text(v)
    return v if p is None else p


def dedupe(t, keys, keep="first"):
    t = trim_table(t)
    index = index_of(t)
    cols = [col(index, k) for k in list_of(keys)] or list(range(len(t[0])))
    chosen = {}
    for r, row in enumerate(t[1:], start=1):
        key = "|".join(js_text(row[c]).strip().lower() for c in cols)
        if keep == "last" or key not in chosen:
            chosen[key] = r
    rows = set(chosen.values())
    return [t[0]] + [row for r, row in enumerate(t[1:], start=1) if r in rows]


def validate(t, rules, mode="issues"):
    t = trim_table(t)
    rules = trim_table(rules)
    index = index_of(t)
    checks = []
    for name, rule, value in rules[1:]:
        name, rule = js_text(name).strip(), js_text(rule).strip().lower()
        if rule in ("min", "max"):
            value = to_number(value)
        elif rule == "oneof":
            value = list_of(value)
        checks.append((col(index, name), name, rule, value))
    seen = {}
    for c, name, rule, value in checks:
        if rule == "unique":
            rows = {}
            for r, row in enumerate(t[1:], start=2):
                v = js_text(row[c]).strip().lower()
                if v:
                    rows[v] = f"{rows[v]}, {r}" if v in rows else f"{r}"
            seen[c] = rows

    def problem(check, v, r):
        c, name, rule, value = check
        if blank(v):
            return "is required" if rule == "required" else ""
        if rule in ("number", "date") and not is_number(v):
            return f"is not a {rule}"
        if rule in ("min", "max"):
            if not is_number(v):
                return "is not a number"
            if rule == "min" and v < value:
                return f"is below {js_text(value)}"
            if rule == "max" and v > value:
                return f"is above {js_text(value)}"
        if rule == "pattern" and not re.search(value, js_text(v)):
            return f"does not match {value}"
        if rule == "oneof" and js_text(v).lower() not in [x.lower() for x in value]:
            return f"is not one of {', '.join(value)}"
        if rule == "unique":
            rows = seen[c][js_text(v).strip().lower()]
            if rows != f"{r}":
                return f"is repeated (rows {rows})"
        return ""

    issues = [["Row", t[0][0], "Column", "Value", "Problem"]]
    valid, flagged = [t[0]], [t[0] + ["Issues"]]
    for r, row in enumerate(t[1:], start=2):
        found = []
        for check in checks:
            p = problem(check, row[check[0]], r)
            if p and f"{check[1]} {p}" not in found:
                issues.append([r, row[0], check[1], row[check[0]], f"{check[1]} {p}"])
                found.append(f"{check[1]} {p}")
        if not found:
            valid.append(row)
        flagged.append(row + ["; ".join(found)])
    if mode == "valid":
        return valid
    if mode == "flagged":
        return flagged
    if len(issues) == 1:
        issues.append(["", "", "", "", "No issues"])
    return issues


def lookup(t, ref, key, bring=""):
    t, ref = trim_table(t), trim_table(ref)
    index, rindex = index_of(t), index_of(ref)
    left, _, right = key.partition("=")
    lk, rk = col(index, left), col(rindex, right or left)
    cols = [col(rindex, n) for n in list_of(bring)] or [c for c in range(len(ref[0])) if c != rk]
    row_of = {}
    for row in ref[1:]:
        k = js_text(row[rk]).strip().lower()
        if k and k not in row_of:
            row_of[k] = row
    out = [t[0] + [ref[0][c] for c in cols] + ["Matched"]]
    for row in t[1:]:
        match = row_of.get(js_text(row[lk]).strip().lower())
        out.append(row + [match[c] if match else "" for c in cols] + [match is not None])
    return out


def amount(t):
    t = trim_table(t)
    index = index_of(t)
    q, p = col(index, "Qty"), col(index, "Price")
    return [t[0] + ["Amount"]] + [row + [js_round(row[q] * row[p], 2)] for row in t[1:]]


def group(t, by, measures):
    t = trim_table(t)
    index = index_of(t)
    cols = [col(index, n) for n in list_of(by)]
    ms = []
    for item in list_of(measures) or ["count:*"]:
        kind, name = (x.strip() for x in item.split(":", 1))
        ms.append((kind.lower(), None if name == "*" else col(index, name), name))
    groups = {}
    for row in t[1:]:
        key = "|".join(js_text(row[c]).strip() for c in cols).lower()
        if key not in groups:
            groups[key] = ([("(blank)" if blank(row[c]) else row[c]) for c in cols], [[0.0, 0, None, None, set()] for _ in ms])
        stats = groups[key][1]
        for (kind, c, _), s in zip(ms, stats):
            if c is None:
                s[1] += 1
                continue
            v = row[c]
            if blank(v):
                continue
            s[1] += 1
            if is_number(v):
                s[0] += v
                s[2] = v if s[2] is None or v < s[2] else s[2]
                s[3] = v if s[3] is None or v > s[3] else s[3]
            s[4].add(js_text(v).strip().lower())
    labels = {"sum": "Sum of", "avg": "Average of", "min": "Min of", "max": "Max of", "distinct": "Distinct", "count": "Count of"}
    out = [[t[0][c] for c in cols] + ["Rows" if n == "*" else f"{labels[k]} {n}" for k, _, n in ms]]
    for key in sorted(groups):
        values, stats = groups[key]
        row = list(values)
        for (kind, _, _), s in zip(ms, stats):
            row.append({"sum": js_round(s[0], 6), "avg": js_round(s[0] / s[1], 6) if s[1] else "",
                        "min": "" if s[2] is None else s[2], "max": "" if s[3] is None else s[3],
                        "distinct": len(s[4]), "count": s[1]}[kind])
        out.append(row)
    return out


def pivot(t, row_name, col_name, measure="count"):
    t = trim_table(t)
    index = index_of(t)
    rc, cc = col(index, row_name), col(index, col_name)
    kind, vc = "count", None
    if measure and measure.lower() != "count":
        kind, name = (x.strip() for x in measure.split(":", 1))
        kind, vc = kind.lower(), col(index, name)
    key = lambda v: "(blank)" if blank(v) else js_text(v).strip()
    cells, rows, cols = {}, {}, {}
    for row in t[1:]:
        rk, ck = key(row[rc]), key(row[cc])
        rows[rk.lower()], cols[ck.lower()] = rk, ck
        for k in (f"{rk.lower()}|{ck.lower()}", f"{rk.lower()}|", f"|{ck.lower()}", "|"):
            acc = cells.setdefault(k, [0.0, 0])
            if kind == "count":
                acc[1] += 1
            elif is_number(row[vc]):
                acc[0] += row[vc]
                acc[1] += 1

    def value(k):
        if k not in cells:
            return ""
        s, n = cells[k]
        if kind == "count":
            return n
        if kind == "avg":
            return js_round(s / n, 6) if n else ""
        return js_round(s, 6)
    out = [[f"{t[0][rc]} / {t[0][cc]}"] + [cols[c] for c in sorted(cols)] + ["Total"]]
    for r in sorted(rows):
        out.append([rows[r]] + [value(f"{r}|{c}") for c in sorted(cols)] + [value(f"{r}|")])
    out.append(["Total"] + [value(f"|{c}") for c in sorted(cols)] + [value("|")])
    return out


def profile(t):
    t = trim_table(t)
    out = [["Column", "Filled", "Empty", "Distinct", "Numbers", "Numeric text", "Min", "Max", "Most frequent", "Times"]]
    for c, name in enumerate(t[0]):
        values = [row[c] for row in t[1:] if not blank(row[c])]
        numbers = [v for v in values if is_number(v)]
        counts, labels = {}, {}
        for v in values:
            k = js_text(v).strip().lower()
            counts[k] = counts.get(k, 0) + 1
            labels.setdefault(k, js_text(v).strip())
        top, times = "", 0
        for k, n in counts.items():
            if n > times:
                top, times = labels[k], n
        numeric_text = sum(1 for v in values if not is_number(v) and to_number(v) is not None)
        out.append([name, len(values), len(t) - 1 - len(values), len(counts), len(numbers), numeric_text,
                    min(numbers) if numbers else "", max(numbers) if numbers else "", top, times])
    return out


def top(t, name, n):
    t = trim_table(t)
    c = col(index_of(t), name)
    rows = sorted(t[1:], key=lambda row: -row[c])
    return [t[0]] + rows[:int(n)]
