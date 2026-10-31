"""Compare regex corpus output with the expected output, case by case.

    python3 compare.py CASES EXPECTED ACTUAL [--show=N]

EXPECTED and ACTUAL may name a member of a zip: expected/full.zip/c1.full.txt
Prints the number of differing cases and, for the first N (default 10), the decoded case and
the first field that differs. Exit status 0 when every case matches.
"""
import os, re, sys, zipfile


def read_lines(path):
    if not os.path.exists(path) and ".zip" + os.sep in path.replace("/", os.sep):
        archive, member = re.split(r"\.zip[\\/]", path, maxsplit=1)
        with zipfile.ZipFile(archive + ".zip") as z:
            text = z.read(member).decode("utf-8")
    else:
        with open(path, encoding="utf-8", newline="") as f:
            text = f.read()
    return [l.rstrip("\r") for l in text.split("\n") if l.rstrip("\r")]


def dec(h):
    return "".join(chr(int(h[i:i + 4], 16)) for i in range(0, len(h), 4))


def pretty(field):
    """Hex runs (UTF-16 code units, 4 digits each) shown as Python string literals."""
    return re.sub(r"(?<![0-9a-f])((?:[0-9a-f]{4})+)(?![0-9a-f])", lambda m: repr(dec(m.group(1))), field)


def first_diff(a, b):
    sep = "\t" if "\t" in a or "\t" in b else " "
    fa, fb = a.split(sep), b.split(sep)
    for k in range(max(len(fa), len(fb))):
        x = fa[k] if k < len(fa) else "<none>"
        y = fb[k] if k < len(fb) else "<none>"
        if x != y:
            return k, x, y
    return -1, "", ""


def compare(cases_path, expected_path, actual_path, show=10):
    cases = read_lines(cases_path)
    exp = read_lines(expected_path)
    act = read_lines(actual_path)
    bad = 0
    for i in range(max(len(exp), len(act))):
        e = exp[i] if i < len(exp) else "<missing>"
        a = act[i] if i < len(act) else "<missing>"
        if e == a:
            continue
        bad += 1
        if bad <= show:
            parts = cases[i].split("\t") if i < len(cases) else ["", ""]
            flags = parts[2] if len(parts) > 2 else ""
            print(f"case {i + 1}: pattern {dec(parts[0])!r} subject {dec(parts[1]) if len(parts) > 1 else ''!r} flags {flags!r}")
            k, x, y = first_diff(e, a)
            print(f"    field {k + 1} expected: {pretty(x)[:300]}")
            print(f"    field {k + 1} actual:   {pretty(y)[:300]}")
    print(f"{os.path.basename(actual_path)}: {max(len(exp), len(act))} cases, {bad} differ")
    return bad


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--show=")]
    show = int(next((a.split("=")[1] for a in sys.argv[1:] if a.startswith("--show=")), "10"))
    if len(args) != 3:
        sys.exit(__doc__)
    sys.exit(1 if compare(*args, show=show) else 0)
