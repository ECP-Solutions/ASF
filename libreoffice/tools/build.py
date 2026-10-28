"""Generate the LibreOffice Basic edition of ASF from the VBA sources.

    python3 libreoffice/tools/build.py            # writes libreoffice/dist/ASF/ and libreoffice/dist/ASF.oxt

src/*.cls stays the single source of truth. Each class is rewritten for LibreOffice Basic
(Option VBASupport 1) by the passes below; every pass works around one documented
difference between VBA and LibreOffice Basic. A few classes are replaced by hand-written
versions in libreoffice/overrides/, and libreoffice/compat/ASF_LOCompat.bas holds the
helpers the rewritten code calls and the factories other libraries use. The VBA
Expressions classes are left out: @(...) runs on VBAExpressionsLib, the library of the
VBA Expressions extension for LibreOffice.

The functions here are also used by run_tests.py to build the test library.
"""
import glob, io, os, re, sys, zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LO_DIR = os.path.join(ROOT, "libreoffice")
LIB_NAME = "ASF"
VERSION = "3.2.0"
EXTENSION_ID = "com.ecp-solutions.asf"

# classes replaced by libreoffice/overrides/<name>.bas
OVERRIDES = ("ASF_Map",)
# VBA Expressions classes: the LibreOffice edition uses the official port instead,
# the VBAExpressionsLib library of the VBA Expressions extension
VBA_EXPRESSIONS = ("VBAexpressions", "VBAexpressionsScope", "UDFunctions", "VBAcallBack")
VBA_EXPRESSIONS_LIB = "VBAExpressionsLib"
# passes that change engine-sensitive code are not applied to the regex engine,
# which only needs the generic ones
REGEX_ENGINE = "ASF_RegexEngine"


def load(path):
    with open(path, encoding="latin-1", newline="") as f:
        return f.read().replace("\r\n", "\n")


# ---------------------------------------------------------------- text helpers
def code_mask(line):
    """True for each character inside a string literal (quotes included)."""
    mask, q = [], False
    for ch in line:
        if ch == '"':
            q = not q
        mask.append(q or ch == '"')
    return mask


def find_code(line, kw, start=0):
    """Index of kw outside string literals, -1 when absent."""
    mask = code_mask(line)
    i = line.find(kw, start)
    while i != -1 and any(mask[i:i + len(kw)]):
        i = line.find(kw, i + 1)
    return i


def split_top(t):
    """Split an argument list on commas outside parentheses and strings."""
    parts, depth, q, cur = [], 0, False, ""
    for ch in t:
        if ch == '"':
            q = not q
        if not q and ch == "(":
            depth += 1
        if not q and ch == ")":
            depth -= 1
        if not q and ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    return parts


def sub_code(pattern, repl, s, flags=0):
    """re.sub applied to code only: string literals and comments are left alone."""
    out = []
    for line in s.split("\n"):
        mask = code_mask(line)
        comment = next((i for i, ch in enumerate(line) if ch == "'" and not mask[i]), len(line))
        segs, cur, inq = [], "", None
        for i, ch in enumerate(line[:comment]):
            if mask[i] != inq and cur:
                segs.append((inq, cur))
                cur = ""
            inq = mask[i]
            cur += ch
        if cur:
            segs.append((inq, cur))
        line = "".join(seg if q else re.sub(pattern, repl, seg, flags=flags) for q, seg in segs) + line[comment:]
        out.append(line)
    return "\n".join(out)


# ---------------------------------------------------------------- generic passes
PROC_START = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(?:Function|Sub|Property (?:Get|Let|Set)) \w+", re.I)
PROC_END = re.compile(r"^\s*End (?:Function|Sub|Property)\b", re.I)
DIM_LINE = re.compile(r"^(\s*)Dim ([^:\"']+?)(?:\s*:\s*(.*))?$")
FOR_EACH = re.compile(r"^(\s*)For Each (\w+) In (.+)$")


def strip_vba_header(s):
    """Drop the VERSION/BEGIN..END block and Attribute lines of exported VBA modules."""
    out, in_header = [], False
    for ln in s.split("\n"):
        if ln.startswith("VERSION 1.0 CLASS"):
            in_header = True
            continue
        if in_header:
            if ln.strip() == "END":
                in_header = False
            continue
        if re.match(r"^Attribute \w+(\.\w+)? *= ", ln.strip()):
            continue
        out.append(ln)
    return "\n".join(out)


def debug_print(s):
    """Debug.Print does not exist in LibreOffice: send it to ASF_LOCompat.LogLine."""
    out = []
    for ln in s.split("\n"):
        i = find_code(ln, "Debug.Print")
        if i == -1:
            out.append(ln)
            continue
        args = ln[i + len("Debug.Print"):]
        mask = code_mask(args)
        args = "".join(" & " if ch == ";" and not mask[k] else ch for k, ch in enumerate(args))
        out.append(ln[:i] + "LogLine" + args)
    return "\n".join(out)


def hoist_dims(code):
    """VBA reserves every Dim at procedure entry; LibreOffice runs Dim where it stands,
    which re-initialises variables declared inside loops. Move them to the top."""
    lines = code.split("\n")
    out, i = [], 0
    while i < len(lines):
        ln = lines[i]
        if PROC_START.match(ln) and not ln.strip().startswith("'"):
            header = [ln]
            while header[-1].rstrip().endswith(" _"):
                i += 1
                header.append(lines[i])
            i += 1
            body, dims = [], []
            while i < len(lines) and not PROC_END.match(lines[i]):
                m = DIM_LINE.match(lines[i])
                if m:
                    d = "    Dim " + m.group(2).rstrip()
                    if d not in dims:
                        dims.append(d)
                    body.append(m.group(1) + m.group(3) if m.group(3) else "")
                else:
                    body.append(lines[i])
                i += 1
            out.extend(header)
            out.extend(dims)
            out.extend(body)
            if i < len(lines):
                out.append(lines[i])
            i += 1
            continue
        out.append(ln)
        i += 1
    return "\n".join(out)


# new, empty storage for an array of each type (see fresh_redim)
FRESH_ARRAY = {"variant": "Array()", "string": "LOEmptyStrings()", "long": "LOEmptyLongs()", "double": "LOEmptyDoubles()"}


def declared_names(decl):
    """{lower-case name: lower-case type} of a Dim or parameter list."""
    names = {}
    for part in split_top(decl):
        m = re.match(r"^\s*(?:Optional\s+)?(?:ByRef\s+|ByVal\s+)?(?:ParamArray\s+)?(\w+)(?:\([^)]*\))?(?:\s+As\s+(\w+))?", part, re.I)
        if m:
            names[m.group(1).lower()] = (m.group(2) or "variant").lower()
    return names


def fresh_redim(code):
    """LibreOffice's ReDim resizes the array in place, so copies already stored elsewhere
    are wiped too: give the array new storage before a ReDim without Preserve.
    VBA's ReDim also declares an undeclared array, LibreOffice's does not."""
    lines = code.split("\n")
    module_vars, inside = {}, False
    for ln in lines:
        if PROC_START.match(ln):
            inside = True
        elif PROC_END.match(ln):
            inside = False
        elif not inside:
            m = re.match(r"^\s*(?:Private|Public|Dim)\s+(?!Const\b|Type\b|Enum\b|Declare\b|Sub\b|Function\b|Property\b)(.+)$", ln, re.I)
            if m:
                module_vars.update(declared_names(m.group(1).split("'")[0]))
    out, proc, local = [], None, None
    for ln in lines:
        if PROC_START.match(ln):
            m = re.search(r"\((.*)\)", ln)
            local = declared_names(m.group(1)) if m else {}
            proc, undeclared = [], []
            out.append(ln)
            continue
        if local is None:
            out.append(ln)
            continue
        if PROC_END.match(ln):
            out.extend(f"    Dim {n}() As Variant" for n in undeclared)
            out.extend(proc)
            out.append(ln)
            local = None
            continue
        m = DIM_LINE.match(ln)
        if m:
            local.update(declared_names(m.group(2)))
        m = re.match(r"^(\s*)ReDim\s+(?!Preserve\b)(.+)$", ln.split("'")[0], re.I)
        if m:
            for part in split_top(m.group(2)):
                name = re.match(r"\s*(\w+)", part).group(1)
                key = name.lower()
                if key not in local and key not in module_vars:
                    undeclared.append(name)
                    local[key] = "variant"
                fresh = FRESH_ARRAY.get(local.get(key, module_vars.get(key)))
                if fresh:
                    proc.append(f"{m.group(1)}{name} = {fresh}")
        proc.append(ln)
    return "\n".join(out)


def rewrite_for_each(code):
    """LibreOffice's For Each variable keeps stale array storage across items of mixed
    types; iterate a fresh variable and copy each item."""
    out, stack = [], []
    for ln in code.split("\n"):
        code_part = ln.split("'")[0]
        m = FOR_EACH.match(ln)
        if m:
            ind, var, src = m.groups()
            out.append(f"{ind}Dim {var}LoIt As Variant")
            out.append(f"{ind}For Each {var}LoIt In {src}")
            out.append(f"{ind}    {var} = Empty: If IsObject({var}LoIt) Then Set {var} = {var}LoIt Else {var} = {var}LoIt")
            stack.append(var)
            continue
        if re.match(r"^\s*For \w+\s*=", code_part):
            stack.append(None)
        mn = re.match(r"^(\s*)Next(?:\s+(\w+))?\s*$", code_part)
        if mn and stack:
            var = stack.pop()
            if var is not None:
                ln = f"{mn.group(1)}Next {var}LoIt"
        out.append(ln)
    return "\n".join(out)


def generic(s):
    s = strip_vba_header(s)
    s = s.replace("pNamedCaptures.Add (Array(gname, val)), gname", "pNamedCaptures.Add Array(gname, val), gname")
    s = debug_print(s)
    s = s.replace("VBA.Split(", "VBASplit(").replace("VBA.Replace(", "VBAReplace(")
    s = re.sub(r"\bVBA\.(IsNumeric|IsDate|IsEmpty|IsNull|IsArray|IsObject)\(", r"\1(", s)
    # byte positions are only compared with 0 in ASF
    s = re.sub(r"\bInStrB\(", "InStr(", s)
    s = s.replace(" As Collection = Nothing)", " As Collection)")
    # 'val' clashes with LibreOffice's Val()
    s = sub_code(r"\bval\b", "valStr", s)
    s = s.replace("    vAssignVar NodeGet, node(key)\n", "    Dim ngTmp As Variant: vAssignVar ngTmp, node(key)\n    If IsObject(ngTmp) Then Set NodeGet = ngTmp Else NodeGet = ngTmp\n")
    # functions declared As T() return a Variant array
    s = re.sub(r"^((?:Public |Private )?Function [^\n]*\)) As \w+\(\)[ \t]*$", r"\1 As Variant", s, flags=re.M)
    s = re.sub(r"\btypeName\(([^()]*(?:\([^()]*\)[^()]*)*)\) = \"Collection\"", r"LOIsColl(\1)", s, flags=re.I)
    s = re.sub(r"\btypeName\(([^()]*)\) <> \"Collection\"", r"Not LOIsColl(\1)", s, flags=re.I)
    s = re.sub(r"\bCLng\(", "VBACLng(", s)
    s = re.sub(r"\bvbObjectError\b", "LOvbObjectError()", s)
    s = re.sub(r" _\n\s*", " ", s)
    s = rewrite_for_each(s)
    s = hoist_dims(s)
    return s


# ---------------------------------------------------------------- ASF passes
def vba_expressions_bridge(s):
    """Calls into VBAExpressionsLib, whose API differs from the VBA class in a few places."""
    s = re.sub(r"\bAs VBAexpressions\b", "As Object", s)
    s = s.replace("If gExprEvaluator Is Nothing Then Set gExprEvaluator = New VBAexpressions",
                  "If gExprEvaluator Is Nothing Then Set gExprEvaluator = LOTryEvaluator()")
    s = s.replace("New VBAexpressions", "LONewEvaluator()")
    s = s.replace("With GLOBALS_.gExprEvaluator", "With LOEvaluator(GLOBALS_)")
    # only the ASF variables the expression can name: VBAExpressionsLib adds every
    # unknown name to its scope, which VBA Expressions 1.2.0 fails on LibreOffice 24.2
    s = s.replace("""                 If IsArray(tmpVal) Then
                    .VarValue(CStr(key)) = .ArrayToString(tmpVal)""", """                 If InStr(1, expr, CStr(key), vbTextCompare) = 0 Then
                 ElseIf IsArray(tmpVal) Then
                    .VarValue(CStr(key)) = .ArrayToString(tmpVal)""")
    # VarValue became GetVarValue/LetVarValue, Eval needs its argument
    s = re.sub(r"(?m)^(\s*)\.VarValue\((.+?)\) = (.+)$", r"\1.LetVarValue \2, \3", s)
    s = re.sub(r"(?m)^(\s*)\.Eval\s*$", r'\1.Eval ""', s)
    s = s.replace("Eval = .Eval(strValues)", 'If IsMissing(strValues) Then\n            Eval = .Eval("")\n        Else\n            Eval = .Eval(strValues)\n        End If')
    # no castTypeData argument
    s = re.sub(r"\.ArrayFromString2\((.+?), True\)", r"LOExprArray(LOEvaluator(GLOBALS_), \1)", s)
    return s


def rewrite_if_line(line):
    """LibreOffice mis-parses single-line If ... Then ... Else; make it a block."""
    m = re.match(r"^(\s*)If ", line)
    if not m or find_code(line, "'") != -1:
        return line
    t = find_code(line, " Then ")
    if t == -1:
        return line
    e = find_code(line, " Else ", t)
    if e == -1:
        return line
    th, el = line[t + 6:e].strip(), line[e + 6:].strip()
    if not th or find_code(th, "If ") != -1 or find_code(el, "If ") == 0 or find_code(" " + th, " If ") != -1:
        return line
    ind = m.group(1)
    return f"{ind}If {line[len(ind) + 3:t]} Then\n{ind}    {th}\n{ind}Else\n{ind}    {el}\n{ind}End If"


def optional_empty(m):
    """Optional x As Variant = Empty is rejected; default it with IsMissing instead."""
    names = re.findall(r"Optional (\w+) As Variant = Empty", m.group(0))
    head = re.sub(r" = Empty", "", m.group(0))
    return head + "".join(f"\n    If IsMissing({n}) Then {n} = Empty" for n in names)


def dim_as_new(m):
    """Dim x As New Cls does not auto-instantiate user classes: Set it explicitly."""
    ind, rest = m.group(1), m.group(2)
    if ":" in rest or '"' in rest:
        return m.group(0)
    decl, sets = [], []
    for part in rest.split(","):
        mm = re.match(r"\s*(\w+)\s+As\s+New\s+(\w+)\s*$", part)
        if mm:
            decl.append(f"{mm.group(1)} As {mm.group(2)}")
            sets.append(f"Set {mm.group(1)} = New {mm.group(2)}")
        else:
            decl.append(part.strip())
    return f"{ind}Dim " + ", ".join(decl) + "".join(f"\n{ind}{x}" for x in sets)


def self_assign(s):
    """vAssignment FuncName, x: a function's own result cannot be passed ByRef."""
    out, cur, n = [], None, 0
    for l in s.split("\n"):
        m = re.match(r"^(?:Public |Private |Friend )?(?:Function|Property Get) (\w+)\(", l)
        if m:
            cur = m.group(1)
        elif re.match(r"^End (Function|Property|Sub)", l):
            cur = None
        mm = re.match(r"^(\s*)vAssignment (\w+), (.+?)\s*$", l)
        if cur and mm and mm.group(2).lower() == cur.lower():
            n += 1
            ind, t = mm.group(1), f"lo_sa{n}"
            out += [f"{ind}Dim {t} As Variant", f"{ind}vAssignment {t}, {mm.group(3)}", f"{ind}If IsObject({t}) Then",
                    f"{ind}    Set {cur} = {t}", f"{ind}Else", f"{ind}    {cur} = {t}", f"{ind}End If"]
            continue
        out.append(l)
    return "\n".join(out)


def self_member(s):
    """FuncName.Method inside FuncName is a recursive call in LibreOffice: build the
    result in a local object and return it on every exit."""
    def fix(m):
        name, body = m.group(1), m.group(0)
        if not re.search(r"(?m)^(?!\s*').*(?<![\w.])" + name + r"\.\w", body):
            return body
        ret = name + "LoRet"
        lines = body.split("\n")
        out = [lines[0], f"    Dim {ret} As Object"]
        for l in lines[1:-1]:
            l = re.sub(r"\bSet " + name + r" = ", f"Set {ret} = ", l)
            l = re.sub(r"(?<![\w.])" + name + r"\.(?=\w)", ret + ".", l)
            l = re.sub(r"\bExit Function\b", f"Set {name} = {ret}: Exit Function", l)
            out.append(l)
        out += [f"    Set {name} = {ret}", lines[-1]]
        return "\n".join(out)
    return re.sub(r"(?ms)^(?:Public |Private |Friend )?Function (\w+)\(.*?^End Function", fix, s)


def map_default_member(s):
    """m("key") on an ASF_Map variable: LibreOffice classes have no default member."""
    maps = set(re.findall(r"(?i)\b(\w+)(?:\(\))? As (?:New )?ASF_Map\b", s))
    colls = set(re.findall(r"(?i)\b(\w+)(?:\(\))? As (?:New )?Collection\b", s))
    for v in maps - colls:
        s = re.sub(r"(?<![.\w])" + v + r"\((?=[^)])", v + ".GetValue(", s)
    return s


_ELEM = [0]


def element_assign(s):
    """vAssignment arr(i), x writes into a copy of the element: assign directly."""
    out = []
    for l in s.split("\n"):
        mm = re.match(r"^(\s*)vAssignment (.+?)\s*$", l)
        if mm and find_code(mm.group(2), "'") == -1:
            parts = split_top(mm.group(2))
            if len(parts) == 2:
                target, value = parts[0].strip(), parts[1].strip()
                if re.match(r"^[\w.]+\(.*\)$", target) and find_code(value, ":") == -1:
                    _ELEM[0] += 1
                    ind, t = mm.group(1), f"lo_ea{_ELEM[0]}"
                    out += [f"{ind}Dim {t} As Variant", f"{ind}vAssignment {t}, {value}", f"{ind}If IsObject({t}) Then",
                            f"{ind}    Set {target} = {t}", f"{ind}Else", f"{ind}    {target} = {t}", f"{ind}End If"]
                    continue
        out.append(l)
    return "\n".join(out)


def variant_param_methods(s):
    """Calling a method on a Variant parameter makes LibreOffice reject the call whenever
    that argument is not an object. Work on a local copy and write it back on exit."""
    lines = s.split("\n")
    out, i = [], 0
    hdr = re.compile(r"^((?:Public |Private |Friend )?(Function|Sub|Property (?:Get|Let|Set)) (\w+)\()(.*)$")
    while i < len(lines):
        m = hdr.match(lines[i])
        if not m:
            out.append(lines[i])
            i += 1
            continue
        j = next(k for k in range(i + 1, len(lines)) if re.match(r"^End (Function|Sub|Property)", lines[k]))
        body = lines[i + 1:j]
        rest = m.group(4)
        depth, close = 1, None
        for ci, ch in enumerate(rest):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    close = ci
                    break
        inner = rest[:close] if close is not None else rest
        params = []
        for part in split_top(inner):
            pm = re.match(r"^\s*(?:Optional\s+)?(?:ByRef\s+|ByVal\s+)?(?:ParamArray\s+)?(\w+)(\(\))?(?:\s+As\s+(\w+))?", part)
            if pm and not pm.group(2) and pm.group(3) in (None, "Variant"):
                params.append(pm.group(1))
        text = "\n".join(body)
        hit = [q for q in params if re.search(r"(?<![\w.])" + q + r"\.\w", text)]
        if not hit:
            out.extend(lines[i:j + 1])
            i = j + 1
            continue
        newinner = inner
        for q in hit:
            newinner = re.sub(r"(?<![\w.])" + q + r"\b", q + "_lo", newinner, count=1)
        out.append(m.group(1) + newinner + (rest[close:] if close is not None else ""))
        for q in hit:
            out.append(f"    Dim {q} As Variant")
            out.append(f"    If Not IsMissing({q}_lo) Then vAssignment {q}, {q}_lo")
        back = ": ".join(f"vAssignment {q}_lo, {q}" for q in hit)
        kind = m.group(2).split()[0]
        for l in body:
            out.append(re.sub(r"\bExit " + kind + r"\b", back + ": Exit " + kind, l))
        out.append("    " + back)
        out.append(lines[j])
        i = j + 1
    return "\n".join(out)


def variant_default_member(s):
    """v("key") on a Variant holding an object: route through LODefault."""
    names = set()
    for m in re.finditer(r"(?m)^\s*Dim\s+(.*)$", s):
        for part in split_top(m.group(1)):
            pm = re.match(r"^\s*(\w+)(\(\))?(?:\s+As\s+(\w+))?\s*$", part.split(":")[0])
            if pm and not pm.group(2) and pm.group(3) in (None, "Variant"):
                names.add(pm.group(1))
    for m in re.finditer(r"(?m)^(?:Public |Private |Friend )?(?:Function|Sub|Property (?:Get|Let|Set)) \w+\((.*)\)", s):
        for part in split_top(m.group(1)):
            pm = re.match(r"^\s*(?:Optional\s+)?(?:ByRef\s+|ByVal\s+)?(\w+)(\(\))?(?:\s+As\s+(\w+))?", part)
            if pm and not pm.group(2) and pm.group(3) in (None, "Variant"):
                names.add(pm.group(1))
    for v in names:
        s = re.sub(r"(?<![.\w])" + v + r"\((\"[^\"]*\")\)", "LODefault(" + v + r", \1)", s)
    return s


def module_collections(s):
    """coll(i) on a class's module-level Collection fails in LibreOffice: use .Item
    before a member access, LOItem elsewhere (.Item(i) of an array item misreads)."""
    for v in re.findall(r"(?m)^(?:Private|Public|Dim) (\w+) As (?:New )?Collection\b", s):
        s = re.sub(r"(?<![.\w])" + v + r"\(([^()]*)\)(?=\.\w)", v + r".Item(\1)", s)
        s = re.sub(r"(?<![.\w])" + v + r"\(", "LOItem(" + v + ", ", s)
    return s


def member_collections(code, names):
    """obj.coll(i) for Public Collection members of ASF classes, as above."""
    for v in names:
        code = re.sub(r"\." + v + r"\(([^()]*)\)(?=\.\w)", "." + v + r".Item(\1)", code)
        code = re.sub(r"((?<![\w.])(?:\w+)?(?:\.\w+)*\.)" + v + r"\(", lambda m: "LOItem(" + m.group(1) + v + ", ", code)
    return code


def asf_passes(name, s):
    s = vba_expressions_bridge(s)
    s = re.sub(r"((?<![\w.])\w+(?:\.\w+)*\.GetValue\((?:\"[^\"]*\"|\w+)\))\((\w+)\)", r"LOItem(\1, \2)", s)
    if name != REGEX_ENGINE:
        # arrays are shared on assignment in LibreOffice: erasing the local after
        # returning it would empty the caller's copy
        s = re.sub(r"(?m)^(\s*)Erase (tmpResult|arr)\s*$", r"\1' Erase \2", s)
    s = map_default_member(s)
    s = module_collections(s)
    s = re.sub(r"(?m)^(\s+)Dim (.*\bAs New\b.*)$", dim_as_new, s)
    s = re.sub(r"(?m)^(?:Public |Private )?(?:Sub|Function) .*Optional \w+ As Variant = Empty.*$", optional_empty, s)
    s = "\n".join(rewrite_if_line(l) for l in s.split("\n"))
    s = self_assign(s)
    s = self_member(s)
    s = element_assign(s)
    if name != REGEX_ENGINE:
        s = variant_default_member(s)
        s = variant_param_methods(s)
        s = fresh_redim(s)
    s = re.sub(r"(?m)^(\s*(?:\w+\s*:\s*)?)(?!Set\b|Dim\b|If\b|Else|Case\b)(\w+(?:\(\w+\))?) = ((?:\w+\.)+GetValue\((?:[^()]|\([^()]*\))*\))\s*$",
               r"\1\2 = LOLet(\3)", s)
    # a String read with Get in Binary mode starts after a 2-byte length LibreOffice expects
    s = re.sub(r"(?m)^(\s*)(\w+) = FreeFile\n\s*Open (\w+) For Binary As #\2\n\s*(\w+) = LOF\(\2\)\n\s*(\w+) = Space\$\(\4\)\n\s*Get #\2, , \5\n\s*Close #\2$",
               r"\1\5 = LOReadTextFile(\3)", s)
    s = re.sub(r"\bCallByName\b", "LOCallByName", s)
    # LibreOffice will not call a method named Clone
    s = re.sub(r"\.Clone\b(\(\))?", ".LOCloneMap", s)
    s = re.sub(r"Public Function Clone\(\)", "Public Function LOCloneMap()", s)
    s = re.sub(r"\bSet Clone = ", "Set LOCloneMap = ", s)
    s = re.sub(r"(?m)^(\s*)ReDim (Preserve )?(\w+)\.Storage\((.+?) To (.+?)\)\s*$",
               lambda m: f"{m.group(1)}{m.group(3)}.Storage = LORedim({m.group(3)}.Storage, {m.group(4)}, {m.group(5)}, {'True' if m.group(2) else 'False'})", s)
    return s


def public_collections(repo):
    names = set()
    for path in glob.glob(os.path.join(repo, "src", "*.cls")):
        names |= set(re.findall(r"(?m)^Public (\w+) As (?:New )?Collection\b", load(path)))
    return names


def class_module(name, body):
    body = re.sub(r"(?i)\btypeName\(", "LOTypeName(", body)
    body += f'\nPublic Function LOClassName() As String\n    LOClassName = "{name}"\nEnd Function\n'
    return "Option VBASupport 1\nOption ClassModule\n" + body


def call_by_name(max_args=30):
    """LOCallByName(obj, name, kind, args...): ASF_LOCompat's dispatch, then CallByName."""
    cases = []
    for n in range(max_args + 1):
        args = "".join(f", a({i})" for i in range(n))
        cases.append(f"        Case {n}\n            If (kind And (VbLet Or VbSet)) <> 0 Then\n"
                     f"                CallByName obj, name, kind{args}\n            Else\n"
                     f"                r = CallByName(obj, name, kind{args})\n            End If")
    return ("\nPublic Function LOCallByName(obj As Object, ByVal name As String, ByVal kind As Integer, ParamArray args() As Variant) As Variant\n"
            "    Dim a() As Variant, r As Variant, i As Long\n"
            "    a = Array()\n"
            "    If UBound(args) >= LBound(args) Then ReDim a(0 To UBound(args) - LBound(args))\n"
            "    For i = 0 To UBound(a)\n"
            "        If IsObject(args(LBound(args) + i)) Then Set a(i) = args(LBound(args) + i) Else a(i) = args(LBound(args) + i)\n"
            "    Next i\n"
            "    If Not LOCallSpecial(obj, name, kind, a, r) Then\n"
            "        Select Case UBound(a) + 1\n" + "\n".join("    " + l for c in cases for l in c.split("\n")) + "\n        End Select\n"
            "    End If\n"
            "    If IsObject(r) Then Set LOCallByName = r Else LOCallByName = r\n"
            "End Function\n")


def library_modules(repo=ROOT):
    """[(name, code, 'class'|'normal')] of the ASF library."""
    pub = public_collections(repo)
    mods = []
    for path in sorted(glob.glob(os.path.join(repo, "src", "*.cls"))):
        name = os.path.splitext(os.path.basename(path))[0]
        if name in VBA_EXPRESSIONS:
            continue
        if name in OVERRIDES:
            body = load(os.path.join(LO_DIR, "overrides", name + ".bas"))
        else:
            body = asf_passes(name, generic(load(path)))
            body = member_collections(body, pub)
        mods.append((name, class_module(name, body), "class"))
    for name in ("ASF_LOCompat", "ASF_VBAExpressions"):
        code = load(os.path.join(LO_DIR, "compat", name + ".bas"))
        if name == "ASF_LOCompat":
            code += call_by_name()
        mods.append((name, "Option VBASupport 1\n" + code, "normal"))
    return mods


def normal_module(src, repo=ROOT):
    """A VBA standard module (tests, user code) rewritten like the library's."""
    s = generic(src)
    s = re.sub(r"(?i)\btypeName\(", "LOTypeName(", s)
    s = member_collections(s, public_collections(repo))
    return "Option VBASupport 1\n" + s


# ---------------------------------------------------------------- packaging
def module_xml(name, code, kind):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            f'<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{name}" '
            f'script:language="StarBasic" script:moduleType="{kind}">' + escape(code) + "</script:module>")


def script_xlb(lib, names):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
            f'<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{lib}" '
            'library:readonly="false" library:passwordprotected="false">\n'
            + "".join(f' <library:element library:name="{n}"/>\n' for n in names) + "</library:library>\n")


def dialog_xlb(lib):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
            f'<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{lib}" '
            'library:readonly="false" library:passwordprotected="false"/>\n')


def library_files(lib, mods):
    files = {f"{name}.xba": module_xml(name, code, kind) for name, code, kind in mods}
    files["script.xlb"] = script_xlb(lib, [n for n, _, _ in mods])
    files["dialog.xlb"] = dialog_xlb(lib)
    return files


def write_library(folder, lib, mods):
    os.makedirs(folder, exist_ok=True)
    for f in os.listdir(folder):
        if f.endswith((".xba", ".xlb")):
            os.remove(os.path.join(folder, f))
    for fname, text in library_files(lib, mods).items():
        with open(os.path.join(folder, fname), "w", encoding="utf-8", newline="\n") as f:
            f.write(text)


def extension_bytes(mods):
    description = f'''<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006"
             xmlns:xlink="http://www.w3.org/1999/xlink">
  <identifier value="{EXTENSION_ID}"/>
  <version value="{VERSION}"/>
  <display-name><name lang="en">ASF - Advanced Scripting Framework</name></display-name>
  <publisher><name xlink:href="https://github.com/ECP-Solutions/ASF" lang="en">ECP Solutions</name></publisher>
  <dependencies>
    <OpenOffice.org-minimal-version xmlns:d="http://openoffice.org/extensions/description/2006" d:name="OpenOffice.org 4.1" value="4.1"/>
  </dependencies>
  <extension-description><src lang="en" xlink:href="description-en.txt"/></extension-description>
</description>
'''
    manifest = f'''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
  <manifest:file-entry manifest:media-type="application/vnd.sun.star.basic-library" manifest:full-path="{LIB_NAME}/"/>
</manifest:manifest>
'''
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("META-INF/manifest.xml", manifest)
        z.writestr("description.xml", description)
        z.writestr("description-en.txt", "ASF scripting engine (JavaScript-like language, regex engine, "
                   "module system) as a LibreOffice Basic library named ASF.\n")
        for fname, text in library_files(LIB_NAME, mods).items():
            z.writestr(f"{LIB_NAME}/{fname}", text)
    return buf.getvalue()


def main():
    mods = library_modules()
    dist = os.path.join(LO_DIR, "dist")
    write_library(os.path.join(dist, LIB_NAME), LIB_NAME, mods)
    with open(os.path.join(dist, f"{LIB_NAME}.oxt"), "wb") as f:
        f.write(extension_bytes(mods))
    print(f"{len(mods)} modules -> {os.path.relpath(dist, ROOT)}/{LIB_NAME}/ and {LIB_NAME}.oxt")


if __name__ == "__main__":
    main()
