"""Prepare RE2's test cases (sihlfall/regex-test-cases) for Re2Cases.bas.

    python3 test/re2-cases/tools/prepare.py [--source=DIR] [--work=DIR]

Clones sihlfall/regex-test-cases at the pinned commit into WORK/regex-test-cases unless --source
names a checkout, then writes for every suite WORK/cases/NNN.txt (subjects and patterns in the
format Re2Cases.bas reads) and WORK/expected/NNN.txt (RE2's leftmost-first results in the format
Re2Cases.bas writes). WORK defaults to test/re2-cases/work.

RE2 syntax that ASF spells differently is translated (see README.md); RE2 reports byte offsets
in UTF-8, which become UTF-16 code unit offsets, and groups become their text, since ASF's
public API returns the text of each group.
"""
import json, os, re, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
REPO = "https://github.com/sihlfall/regex-test-cases"
COMMIT = "0ab3381ad388eaa44576fe33fa66aa3458c61c43"

# RE2's POSIX classes, as ASF class members
POSIX = {
    "alnum": "0-9A-Za-z", "alpha": "A-Za-z", "ascii": "\\x00-\\x7f", "blank": "\\t ",
    "cntrl": "\\x00-\\x1f\\x7f", "digit": "0-9", "graph": "!-~", "lower": "a-z",
    "print": " -~", "punct": "!-/:-@\\[-`{-~", "space": "\\t\\n\\v\\f\\r ", "upper": "A-Z",
    "word": "0-9A-Za-z_", "xdigit": "0-9A-Fa-f",
}
# RE2's \s is [\t\n\f\r ]; ASF's also has \v
SPACE = "\\t\\n\\f\\r "
NOT_SPACE = "\\x00-\\x08\\x0b\\x0e-\\x1f!-\\uffff"


def translate(p):
    """An ASF pattern that means what the RE2 pattern p means on RE2's default flags."""
    out, i, n = [], 0, len(p)
    while i < n:
        c = p[i]
        if c == "\\" and i + 1 < n:
            e = p[i + 1]
            out.append({"C": "[\\s\\S]", "A": "^", "z": "$", "a": "\\x07",
                        "s": "[" + SPACE + "]", "S": "[^" + SPACE + "]"}.get(e, c + e))
            i += 2
        elif c == ".":
            # RE2's dot excludes only \n
            out.append("[^\\n]")
            i += 1
        elif c == "[":
            j, cls = i + 1, ["["]
            if j < n and p[j] == "^":
                cls.append("^")
                j += 1
            if j < n and p[j] == "]":
                # a leading ] is a member in RE2
                cls.append("\\]")
                j += 1
            while j < n and p[j] != "]":
                m = re.match(r"\[:(\^?)(\w+):\]", p[j:])
                if m and m.group(2) in POSIX and not m.group(1):
                    cls.append(POSIX[m.group(2)])
                    j += m.end()
                elif p[j] == "\\" and j + 1 < n:
                    e = p[j + 1]
                    cls.append({"s": SPACE, "S": NOT_SPACE, "a": "\\x07"}.get(e, "\\" + e))
                    j += 2
                else:
                    cls.append(p[j])
                    j += 1
            cls.append("]")
            out.append("".join(cls))
            i = j + 1
        else:
            out.append(c)
            i += 1
    return "".join(out)


def enc(s):
    return "".join("%04x" % u for u in utf16_units(s))


def utf16_units(s):
    b = s.encode("utf-16-le", "surrogatepass")
    return [b[k] | b[k + 1] << 8 for k in range(0, len(b), 2)]


def offset_map(s):
    """UTF-8 byte offset -> UTF-16 code unit offset, at every char boundary of s."""
    m, b, u = {0: 0}, 0, 0
    for ch in s:
        b += len(ch.encode("utf-8", "surrogatepass"))
        u += 2 if ord(ch) > 0xFFFF else 1
        m[b] = u
    return m


def expected(res, s, omap):
    """One RE2 result (byte offsets, [] for no match) in Re2Cases.bas's output format, "~" when
    an offset falls inside a UTF-8 sequence."""
    if not res:
        return "-"
    if any(v >= 0 and v not in omap for v in res):
        # \C matched part of a UTF-8 sequence: no UTF-16 offset for it
        return "~"
    units = utf16_units(s)
    line = "%d,%d" % (omap[res[0]], omap[res[1]])
    for k in range(2, len(res), 2):
        a, z = res[k], res[k + 1]
        text = units[omap[a]:omap[z]] if a >= 0 else []
        line += "|" + "".join("%04x" % u for u in text)
    return line


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    work = os.path.abspath(opts.get("--work", os.path.join(BASE, "work")))
    src = opts.get("--source")
    if not src:
        src = os.path.join(work, "regex-test-cases")
        if not os.path.isdir(src):
            subprocess.run(["git", "clone", "-q", REPO, src], check=True)
        subprocess.run(["git", "-C", src, "checkout", "-q", COMMIT], check=True)
    for sub in ("cases", "expected"):
        os.makedirs(os.path.join(work, sub), exist_ok=True)
    total = 0
    for name in sorted(os.listdir(os.path.join(src, "test-cases"))):
        if not name.endswith(".json"):
            continue
        with open(os.path.join(src, "test-cases", name), encoding="utf-8") as f:
            suite = json.load(f)
        with open(os.path.join(src, "test-results", "re2", name), encoding="utf-8") as f:
            results = json.load(f)
        num = name[:3]
        maps = [offset_map(s) for s in suite["strs"]]
        with open(os.path.join(work, "cases", num + ".txt"), "w", encoding="utf-8", newline="\n") as f:
            for s in suite["strs"]:
                f.write("S " + enc(s) + "\n")
            for p in suite["regexs"]:
                f.write("R " + enc(translate(p)) + "\t" + enc(p) + "\n")
        with open(os.path.join(work, "expected", num + ".txt"), "w", encoding="utf-8", newline="\n") as f:
            for i in range(len(suite["regexs"])):
                for j, s in enumerate(suite["strs"]):
                    whole, leftmost = results[i][j][0], results[i][j][1]
                    f.write(expected(leftmost, s, maps[j]) + "\t" + expected(whole, s, maps[j]) + "\n")
        pairs = len(suite["regexs"]) * len(suite["strs"])
        total += pairs
        print(f"{name}: {len(suite['regexs'])} patterns x {len(suite['strs'])} subjects = {pairs} cases")
    print(f"{total} cases in {work}")


if __name__ == "__main__":
    main(sys.argv[1:])
