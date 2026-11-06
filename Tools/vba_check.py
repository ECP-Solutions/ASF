"""Static checks for compile errors of VBA that LibreOffice Basic lets through.

    python3 Tools/vba_check.py                 # src/*.cls and src/ASF bytecode/*.cls
    python3 Tools/vba_check.py src/ASF_VM.cls

The LibreOffice edition (libreoffice/tools/run_tests.py) runs the test suites, but LibreOffice
Basic compiles code that the VBA editor refuses. This script finds three of those errors:

- ByRef argument type mismatch: a ByRef parameter (the default) declared with a type other than
  Variant takes only a variable or array element of that same type. A Variant, or an element of
  a Variant that holds an array (pinfo(1)), does not compile; an expression (a call, (x),
  obj.Prop) does.
- Duplicate declaration in current scope: a name declared twice in a procedure, or declared again
  after being a parameter.
- Member already exists in an object module from which this object module derives: a member of a
  class module named as a method of IUnknown or IDispatch (Invoke, QueryInterface...).

Calls through an object variable are checked when the method name belongs to one module only.
The exit status is 1 when something is found.
"""
import glob, os, re, sys

PROC = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(Function|Sub|Property (?:Get|Let|Set))\s+(\w+)\s*\((.*)\)"
                  r"\s*(?:As\s+(?:New\s+)?([\w.]+)(\(\))?)?\s*$", re.I)
END = re.compile(r"^\s*End\s+(Function|Sub|Property)\b", re.I)
MODULE_DECL = re.compile(r"^\s*(?:Dim|Private|Public|Global)\s+(?!Sub\b|Function\b|Property\b|Const\b|Type\b|Enum\b|Declare\b|Event\b)(.*)$", re.I)
LOCAL_DECL = re.compile(r"^\s*(?:Dim|Static)\s+(.*)$", re.I)
COLLECTION_METHODS = {"add", "remove", "item", "count", "exists", "keys", "items"}
DISPATCH_MEMBERS = {"queryinterface", "addref", "release", "gettypeinfocount", "gettypeinfo", "getidsofnames", "invoke"}


def strip_comment(line):
    out, quoted = [], False
    for ch in line:
        if ch == '"':
            quoted = not quoted
        elif ch == "'" and not quoted:
            break
        out.append(ch)
    return "".join(out)


def logical_lines(text):
    """(first line number, code) with continuations joined and comments removed"""
    out, buf, start = [], "", 0
    for n, raw in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        code = strip_comment(raw)
        if not buf:
            start = n
        if code.rstrip().endswith(" _"):
            buf += code.rstrip()[:-1] + " "
            continue
        out.append((start, buf + code))
        buf = ""
    return out


def split_args(s):
    args, depth, quoted, cur = [], 0, False, ""
    for ch in s:
        if ch == '"':
            quoted = not quoted
        if not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                args.append(cur.strip())
                cur = ""
                continue
        cur += ch
    if cur.strip() or args:
        args.append(cur.strip())
    return args


def closing_paren(code, start):
    depth = 0
    for k in range(start, len(code)):
        if code[k] == "(":
            depth += 1
        elif code[k] == ")":
            depth -= 1
            if depth == 0:
                return k
    return len(code)


def parse_params(s):
    params = []
    for p in split_args(s):
        m = re.match(r"(?:Optional\s+)?(?:(ByVal|ByRef)\s+)?(?:(ParamArray)\s+)?(\w+)(\(\))?(?:\s+As\s+(?:New\s+)?([\w.]+))?", p, re.I)
        if m:
            params.append({"name": m.group(3), "byval": (m.group(1) or "").lower() == "byval", "paramarray": bool(m.group(2)),
                           "type": (m.group(5) or "Variant") + ("()" if m.group(4) else ""), "array": bool(m.group(4))})
    return params


def parse_decls(rest):
    """a, b As T, c() As U -> {name: (type, is array)}; a name without As is a Variant"""
    found = []
    for part in split_args(rest.split(":")[0]):
        m = re.match(r"(?:WithEvents\s+)?(\w+)(\s*\([^)]*\))?(?:\s+As\s+(?:New\s+)?([\w.]+))?", part.strip(), re.I)
        if m:
            found.append((m.group(1).lower(), m.group(3) or "Variant", bool(m.group(2))))
    return found


def read_module(path):
    lines = logical_lines(open(path, encoding="latin-1").read())
    procs, modvars, i = {}, {}, 0
    while i < len(lines):
        n, code = lines[i]
        m = PROC.match(code)
        if m:
            body, j = [], i + 1
            while j < len(lines) and not END.match(lines[j][1]):
                body.append(lines[j])
                j += 1
            procs[m.group(2).lower()] = {"name": m.group(2), "kind": m.group(1).lower(), "params": parse_params(m.group(3)),
                                         "ret": m.group(4) or "Variant", "retarray": bool(m.group(5)), "body": body}
            i = j + 1
            continue
        d = MODULE_DECL.match(code)
        if d:
            for name, vtype, is_array in parse_decls(d.group(1)):
                modvars[name] = (vtype, is_array)
        i += 1
    return procs, modvars


def local_decls(code):
    """the declarations of a line, also after a colon: Dim a As Long: Dim b"""
    for seg in re.split(r":(?=\s*(?:Dim|Static)\b)", code):
        d = LOCAL_DECL.match(seg)
        if d:
            yield from parse_decls(d.group(1))


def arg_type(arg, scope):
    """the declared type of a variable or array element argument; None for an expression"""
    m = re.fullmatch(r"(\w+)(\s*\([^()]*(?:\([^()]*\)[^()]*)*\))?", arg.strip())
    if not m or m.group(1).lower() not in scope:
        return None
    vtype, is_array = scope[m.group(1).lower()]
    index = m.group(2)
    if index is None or index.replace(" ", "") == "()":
        return vtype + "()" if is_array else vtype
    if is_array:
        return vtype
    if vtype.lower() == "variant":
        return "Variant"
    return None


def calls(code, path, modules):
    """(procedure, argument list) of the calls in a line"""
    own = modules[path][0]

    def other(name):
        if name.lower() in COLLECTION_METHODS:
            return None
        found = [p for procs, _ in modules.values() for n, p in procs.items() if n == name.lower()]
        return found[0] if len(found) == 1 else None

    # Name (x), y at the start of a statement is a call without parentheses: (x) is an expression
    statement = re.match(r"^\s*(Me\.)?(\w+)\s+\(", code, re.I)
    for m in re.finditer(r"(?<![\w.])(Me\.)?(\w+)\s*\(", code, re.I):
        proc = own.get(m.group(2).lower())
        if proc and not (statement and m.start() == statement.start(2) - len(m.group(1) or "")):
            yield proc, split_args(code[m.end():closing_paren(code, m.end() - 1)])
    for m in re.finditer(r"(?<!Me)\.(\w+)\s*\(", code, re.I):
        proc = other(m.group(1))
        if proc:
            yield proc, split_args(code[m.end():closing_paren(code, m.end() - 1)])
    st = re.match(r"^\s*(?:Call\s+)?(Me\.)?(\w+)\s+(?!=)(.+)$", code, re.I)
    if st and st.group(2).lower() in own:
        yield own[st.group(2).lower()], split_args(st.group(3))
    st = re.match(r"^\s*(?:Call\s+)?(?!Me\.)[\w.]+\.(\w+)\s+(?![=(])(.+)$", code, re.I)
    if st:
        proc = other(st.group(1))
        if proc:
            yield proc, split_args(st.group(2))


def check(paths):
    modules = {path: read_module(path) for path in paths}
    problems = []
    for path, (procs, modvars) in modules.items():
        for proc in procs.values():
            if path.lower().endswith(".cls") and proc["name"].lower() in DISPATCH_MEMBERS:
                problems.append(f"{path}: {proc['name']}: a class member cannot be named as a member of IDispatch")
            scope = dict(modvars)
            declared = {}
            for prm in proc["params"]:
                scope[prm["name"].lower()] = (prm["type"].replace("()", ""), prm["array"])
                declared[prm["name"].lower()] = None
            if proc["kind"] in ("function", "property get"):
                scope[proc["name"].lower()] = (proc["ret"], proc["retarray"])
            for n, code in proc["body"]:
                for name, vtype, is_array in local_decls(code):
                    if name in declared:
                        where = "a parameter" if declared[name] is None else f"declared at line {declared[name]}"
                        problems.append(f"{path}:{n}: {proc['name']}: '{name}' declared again ({where})")
                    else:
                        declared[name] = n
                    scope[name] = (vtype, is_array)
            for n, code in proc["body"]:
                for callee, args in calls(code, path, modules):
                    for prm, arg in zip(callee["params"], args):
                        if prm["byval"] or prm["paramarray"] or prm["type"].lower() == "variant" or ":=" in arg:
                            continue
                        vtype = arg_type(arg, scope)
                        if vtype is not None and vtype.lower() != prm["type"].lower():
                            problems.append(f"{path}:{n}: {proc['name']}: {arg} ({vtype}) passed to ByRef "
                                            f"{prm['name']} As {prm['type']} of {callee['name']}")
    return problems


if __name__ == "__main__":
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    paths = sys.argv[1:] or sorted(glob.glob(os.path.join(root, "src", "*.cls")) + glob.glob(os.path.join(root, "src", "ASF bytecode", "*.cls")))
    found = check(paths)
    for p in found:
        print(p)
    print(f"{len(paths)} module(s), {len(found)} problem(s)")
    sys.exit(1 if found else 0)
