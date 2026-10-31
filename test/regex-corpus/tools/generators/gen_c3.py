"""Generate random regex cases shared by ASF and JavaScript. Output: one case per line, hex-encoded fields.

cases/c3.txt: python3 gen_c3.py 21 6000 c3.txt
"""
import random, sys

seed, count, out = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
rnd = random.Random(seed)
named = set()
LITS = ["a", "b", "c", "x", "1", " ", "_", "-", "\\x61", "\\u0062", "\\t", "\\$", "\\^", "\\\\", "\u00e9", "\u4e2d"]
SUBJ_ALPHABET = "aabbcx1 _-\nA\\$^\t\u00e9\u4e2d"


def atom(depth, groups):
    global named
    r = rnd.random()
    if r < 0.30:
        return rnd.choice(LITS)
    if r < 0.38:
        return "."
    if r < 0.50:
        return rnd.choice(["[ab]", "[^a]", "[a-c]", "[^ ]", "\\d", "\\w", "\\s", "\\W", "\\D", "[x1]", "[\\d_]", "[\\D]", "[\\W_]", "[^\\S]", "[\\x61-c]", "[\\u4e00-\\u9fff]", "[^\\u00e0-\\u00ff]", "\\S", "[$^]"])
    if r < 0.56:
        return rnd.choice(["\\b", "\\B", "^", "$"])
    if r < 0.62 and groups[0] > 0:
        n = rnd.randint(1, groups[0])
        if n in named and rnd.random() < 0.5:
            return "(?:\\k<g" + str(n) + ">)"
        return "(?:\\" + str(n) + ")"
    if depth < 3 and r < 0.80:
        kind = rnd.random()
        if kind < 0.55:
            groups[0] += 1
            idx = groups[0]
            if rnd.random() < 0.3:
                named.add(idx)
                return "(?<g" + str(idx) + ">" + alt(depth + 1, groups) + ")"
            return "(" + alt(depth + 1, groups) + ")"
        if kind < 0.75:
            return "(?:" + alt(depth + 1, groups) + ")"
        if kind < 0.85:
            return "(?=" + alt(depth + 1, groups) + ")"
        if kind < 0.93:
            return "(?!" + alt(depth + 1, groups) + ")"
        fixed = "".join(rnd.choice(["a", "b", "x", "1", " ", "."]) for _ in range(rnd.randint(1, 2)))
        return rnd.choice(["(?<=", "(?<!"]) + fixed + ")"
    return rnd.choice(LITS)


def quant(a):
    if a in ("\\b", "\\B", "^", "$") or a.startswith("(?=") or a.startswith("(?!") or a.startswith("(?<"):
        return a
    r = rnd.random()
    if r < (0.35 if a.startswith("(") else 0.6):
        return a
    q = rnd.choice(["*", "+", "?", "{2}", "{1,2}", "{0,1}", "{2,}"])
    if rnd.random() < 0.3:
        q += "?"
    return a + q


def seq(depth, groups):
    return "".join(quant(atom(depth, groups)) for _ in range(rnd.randint(0 if depth else 1, 4)))


def alt(depth, groups):
    parts = [seq(depth, groups)]
    while rnd.random() < 0.4:
        parts.append(seq(depth, groups))
    return "|".join(parts)


def hexs(s):
    return "".join(f"{ord(c):04x}" for c in s)


with open(out, "w") as f:
    for _ in range(count):
        groups = [0]
        named.clear()
        p = alt(0, groups)
        a1, a2 = rnd.random() < 0.15, rnd.random() < 0.15
        if a1 or a2:
            p = ("^" if a1 else "") + "(?:" + p + ")" + ("$" if a2 else "")
        flags = "".join(fl for fl in "ims" if rnd.random() < 0.2)
        subj = "".join(rnd.choice(SUBJ_ALPHABET) for _ in range(rnd.randint(0, 8)))
        f.write(f"{hexs(p)}\t{hexs(subj)}\t{flags}\n")
