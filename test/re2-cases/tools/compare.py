"""Compare Re2Cases.bas output with RE2's results.

    python3 test/re2-cases/tools/compare.py WORK OUT [SUITE ...] [--show=N]

Every case (pattern, subject) has two checks: the leftmost match and the match anchored at both
ends. A check that differs is put in the first category that explains it:

    utf16      the subject has a char above U+FFFF, which ASF sees as two UTF-16 code units
               while RE2 sees one code point
    bytes      the pattern has \\C, which matches one byte in RE2 and one char in ASF, and the
               subject is not ASCII
    ecmascript ASF gives what JavaScript's RegExp gives (Node.js, tools/js_check.js): RE2
               follows Perl where ECMAScript differs, e.g. on empty iterations of a loop and on
               captures kept from earlier iterations
    rejected   ASF rejected the pattern (E<number>)
    error      ASF raised while matching (X<number>)
    groups     same span, different group text
    span       different match, or match against no match

Checks where RE2 matched part of a UTF-8 sequence are skipped. Without Node.js the ecmascript
check is left out. Prints the counts per suite and the first N (default 5) differences per
category, decoded, and writes OUT/summary.json. Exit status 0 when nothing differs outside the
utf16, bytes and ecmascript categories.
"""
import json, os, shutil, subprocess, sys, tempfile
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
CATEGORIES = ["utf16", "bytes", "ecmascript", "rejected", "error", "groups", "span"]
EXPLAINED = {"utf16", "bytes", "ecmascript"}


def suites(work):
    return sorted(n[:-4] for n in os.listdir(os.path.join(work, "cases")) if n.endswith(".txt"))


def dec(h):
    return "".join(chr(int(h[i:i + 4], 16)) for i in range(0, len(h), 4)).encode("utf-16", "surrogatepass").decode("utf-16")


def read_cases(path):
    subjects, patterns = [], []
    with open(path, encoding="utf-8") as f:
        for ln in f:
            ln = ln.rstrip("\n")
            if ln.startswith("S "):
                subjects.append(dec(ln[2:]))
            elif ln.startswith("R "):
                translated, original = ln[2:].split("\t")
                patterns.append((dec(original), dec(translated)))
    return subjects, patterns


def category(original, subject, expected, actual, js):
    if any(ord(c) > 0xFFFF for c in subject):
        return "utf16"
    if "\\C" in original and any(ord(c) > 0x7F for c in subject):
        return "bytes"
    if js is not None and js == actual:
        return "ecmascript"
    if actual.startswith("E"):
        return "rejected"
    if actual.startswith("X"):
        return "error"
    if "|" in expected and "|" in actual and expected.split("|")[0] == actual.split("|")[0]:
        return "groups"
    return "span"


def pretty(field):
    if field in ("-", "~") or field[:1] in ("E", "X"):
        return field
    span, *groups = field.split("|")
    return span + "".join(" " + repr(dec(g)) for g in groups)


def enc(s):
    b = s.encode("utf-16-le", "surrogatepass")
    return "".join("%04x" % (b[k] | b[k + 1] << 8) for k in range(0, len(b), 2))


def run_js(checks):
    """JavaScript's result for each (translated pattern, subject, kind), None without Node.js."""
    node = shutil.which("node")
    if not node or not checks:
        return [None] * len(checks)
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = os.path.join(tmp, "in.txt"), os.path.join(tmp, "out.txt")
        with open(src, "w", encoding="utf-8", newline="\n") as f:
            for translated, subject, kind in checks:
                f.write(f"{enc(translated)}\t{enc(subject)}\t{kind[0].upper()}\n")
        subprocess.run([node, os.path.join(HERE, "js_check.js"), src, dst], check=True)
        with open(dst, encoding="utf-8") as f:
            return f.read().split("\n")[:len(checks)]


def compare_suite(work, out, num, counts, samples, show):
    subjects, patterns = read_cases(os.path.join(work, "cases", num + ".txt"))
    differing = []
    with open(os.path.join(work, "expected", num + ".txt"), encoding="utf-8") as fe, \
         open(os.path.join(out, num + ".txt"), encoding="utf-8", errors="replace") as fa:
        for (original, translated) in patterns:
            for subject in subjects:
                exp = fe.readline().rstrip("\n").split("\t")
                act = (fa.readline().rstrip("\r\n") or "<missing>\t<missing>").split("\t")
                for kind, e, a in zip(("leftmost", "anchored"), exp, act + ["<missing>"]):
                    counts["checks"] += 1
                    if e == "~":
                        counts["skipped"] += 1
                    elif e == a:
                        counts["pass"] += 1
                    else:
                        differing.append((original, translated, subject, kind, e, a))
    js = run_js([(t, s, k) for _, t, s, k, _, _ in differing])
    for (original, translated, subject, kind, e, a), j in zip(differing, js):
        cat = category(original, subject, e, a, j)
        counts[cat] += 1
        if len(samples[cat]) < show:
            samples[cat].append(f"{num} {kind}: RE2 {original!r} (ASF {translated!r}) on {subject!r}: "
                                f"RE2 {pretty(e)}, ASF {pretty(a)}" + ("" if j is None else f", JavaScript {pretty(j)}"))


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    args = [a for a in argv if not a.startswith("--")]
    work, out = os.path.abspath(args[0]), os.path.abspath(args[1])
    show = int(opts.get("--show", "5"))
    nums = args[2:] or suites(work)
    total, per_suite = Counter(), {}
    samples = defaultdict(list)
    print(f"{'suite':6} {'checks':>9} {'pass':>9} {'skipped':>8} " + " ".join(f"{c:>8}" for c in CATEGORIES))
    for num in nums:
        counts = Counter()
        compare_suite(work, out, num, counts, samples, show)
        per_suite[num] = dict(counts)
        total.update(counts)
        print(f"{num:6} {counts['checks']:9} {counts['pass']:9} {counts['skipped']:8} " + " ".join(f"{counts[c]:8}" for c in CATEGORIES))
    print(f"{'total':6} {total['checks']:9} {total['pass']:9} {total['skipped']:8} " + " ".join(f"{total[c]:8}" for c in CATEGORIES))
    for cat in CATEGORIES:
        for s in samples[cat]:
            print(f"  [{cat}] {s}")
    with open(os.path.join(out, "summary.json"), "w", encoding="utf-8") as f:
        json.dump({"total": dict(total), "suites": per_suite, "samples": dict(samples)}, f, indent=2, ensure_ascii=False)
    unexplained = sum(total[c] for c in CATEGORIES if c not in EXPLAINED)
    return 1 if unexplained else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
