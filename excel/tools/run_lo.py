"""Checks the modules of the Excel add-in in LibreOffice, headless: every case of cases.py.

    python3 excel/tools/run_lo.py

Builds the ASF library from src/ (libreoffice/tools/build.py) with the modules of excel/addin and
the regex functions of the use cases, and calls the worksheet functions from Basic: a Range
argument goes in as Range.Value2 would give it (one cell as its value, more as a 2D array from 1),
and a result over several cells goes through the fitting an array formula gets. The library folder
is a temporary folder with the scripts of cases.LIBRARY, and the step limit is 100000. What needs
Excel itself (Application.Caller, the settings in the registry, MacroOptions) is not run. Exits with
0 when every case matches. Needs LibreOffice and python3-uno.
"""
import os, shutil, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, os.path.join(ROOT, "Tools", "stdRegex"))
sys.path.insert(0, HERE)
import build  # noqa: E402
from BenchmarkLO import start_office  # noqa: E402
from cases import CASES, HEADERS, LIBRARY, RIBBON_NAMES, RIBBON_RUNS, Error, Range, shape  # noqa: E402
import re  # noqa: E402
import xml.etree.ElementTree as ET  # noqa: E402

CHECK = '''Option VBASupport 1
' an argument from Python: "#ERR:2042" is an error value, a tuple of tuples a range
Private Function CellOf(v As Variant) As Variant
    If VarType(v) = vbString Then
        If Left(v, 5) = "#ERR:" Then CellOf = CVErr(CLng(Mid(v, 6))): Exit Function
    End If
    CellOf = v
End Function
Private Function Arg(v As Variant) As Variant
    If Not IsArray(v) Then Arg = CellOf(v): Exit Function
    Dim n As Long, m As Long, r As Long, c As Long, g() As Variant
    n = UBound(v) - LBound(v) + 1
    m = UBound(v(LBound(v))) - LBound(v(LBound(v))) + 1
    ' Range.Value2 of one cell is its value
    If n = 1 And m = 1 Then Arg = CellOf(v(LBound(v))(LBound(v(LBound(v))))): Exit Function
    ReDim g(1 To n, 1 To m)
    For r = 1 To n
        For c = 1 To m
            g(r, c) = CellOf(v(LBound(v) + r - 1)(LBound(v(LBound(v))) + c - 1))
        Next c
    Next r
    Arg = g
End Function
' fn with args; rows x columns: the cells of the formula. An error comes back as "#ERRV:" and its text
Function CallFn(fn As String, args As Variant, rows As Long, columns As Long) As Variant
    Dim a(0 To 4) As Variant, n As Long, i As Long, r As Variant
    n = UBound(args) - LBound(args) + 1
    For i = 0 To n - 1
        a(i) = Arg(args(LBound(args) + i))
    Next i
    Select Case fn & n
        Case "ASF_RUN1": r = ASFFunctions.ASF_RUN(a(0))
        Case "ASF_RUN2": r = ASFFunctions.ASF_RUN(a(0), a(1))
        Case "ASF_RUN3": r = ASFFunctions.ASF_RUN(a(0), a(1), a(2))
        Case "ASF_CALL1": r = ASFFunctions.ASF_CALL(a(0))
        Case "ASF_CALL2": r = ASFFunctions.ASF_CALL(a(0), a(1))
        Case "ASF_CALL3": r = ASFFunctions.ASF_CALL(a(0), a(1), a(2))
        Case "RX_TEST2": r = RegexFunctions.RX_TEST(a(0), a(1))
        Case "RX_EXTRACT3": r = RegexFunctions.RX_EXTRACT(a(0), a(1), a(2))
        Case "RX_REPLACE3": r = RegexFunctions.RX_REPLACE(a(0), a(1), a(2))
        Case "RX_COUNT2": r = RegexFunctions.RX_COUNT(a(0), a(1))
        Case Else: CallFn = "no case " & fn & n: Exit Function
    End Select
    r = ASFCells.FitTo(r, rows, columns)
    If IsError(r) Then CallFn = "#ERRV:" & ASFCells.ErrorText(r): Exit Function
    CallFn = r
End Function
Function LastMessage() As String
    LastMessage = ASFEngine.LastErrorMessage()
End Function
' the ribbon's library: the names it lists, a header value, a run of Run script
Function ScriptNames() As String
    Dim s As Variant, i As Long, item As Variant
    s = ASFLibrary.LibraryScripts()
    If IsEmpty(s) Then Exit Function
    For i = LBound(s) To UBound(s)
        item = s(i)
        ScriptNames = ScriptNames & item(1) & "|"
    Next i
End Function
Function Header(source As String, key As String) As String
    Header = ASFLibrary.HeaderValue(source, key)
End Function
Function RibbonRun(file As String, sel As Variant) As String
    RibbonRun = ASFLibrary.RunForRibbon(ASFSettings.LibraryFolder() & "/" & file, Arg(sel))
End Function
Function MakeFolder(path As String) As Boolean
    MakeFolder = ASFLibrary.EnsureFolder(path)
End Function
Function Registered() As Variant
    Registered = ASFRegister.FunctionInfo()
End Function
'''


def addin_module(path, library):
    with open(path, encoding="utf-8") as f:
        src = f.read()
    src = src.replace("Option Private Module\n", "")
    # the library folder and the step limit of the check
    src = src.replace('Environ$("APPDATA") & "\\ASF\\scripts"', '"' + library + '"')
    src = src.replace("DEFAULT_STEP_LIMIT As Long = 10000000", "DEFAULT_STEP_LIMIT As Long = 100000")
    name = os.path.splitext(os.path.basename(path))[0]
    return (name, build.normal_module(build.resolve_conditionals(src)), "normal")


def to_basic(v):
    if isinstance(v, Range):
        return tuple(tuple(to_basic(c) for c in row) for row in v.rows)
    if isinstance(v, Error):
        return "#ERR:%d" % Error.CODES[v.text]
    return v


def norm(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return "%.10g" % v
    if isinstance(v, Error):
        return "#ERRV:" + v.text
    if v is None:
        return ""
    return str(v)


def table(v):
    if isinstance(v, (list, tuple)):
        return [[norm(c) for c in row] for row in v]
    return norm(v)


def ribbon_problems():
    """The ribbon XML against ASFRibbon.bas: every onAction a Sub there, ids unique."""
    ns = "{http://schemas.microsoft.com/office/2009/07/customui}"
    tree = ET.parse(os.path.join(ROOT, "excel", "customUI", "customUI14.xml"))
    with open(os.path.join(ROOT, "excel", "addin", "ASFRibbon.bas"), encoding="utf-8") as f:
        subs = set(re.findall(r"(?m)^Public Sub (\w+)\(control As Object\)", f.read()))
    problems, ids = [], set()
    if tree.getroot().tag != ns + "customUI":
        problems.append("the root is not customUI of the 2009/07 schema")
    for el in tree.iter():
        i = el.get("id")
        if i is not None:
            if i in ids:
                problems.append(f"id {i} twice")
            ids.add(i)
        action = el.get("onAction")
        if action and (not action.startswith("ASFRibbon.") or action.split(".", 1)[1] not in subs):
            problems.append(f"onAction {action} is not a Sub of ASFRibbon")
    return problems


def registration_problems(info):
    """ASFRegister's table against the worksheet functions: every one there, with its arguments in
    order, [optional] ones bracketed, and every text within Excel's 255 characters."""
    sigs = {}
    for path in (os.path.join(ROOT, "excel", "addin", "ASFFunctions.bas"),
                 os.path.join(ROOT, "examples", "use-cases", "regex", "RegexFunctions.bas")):
        with open(path, encoding="utf-8") as f:
            src = re.sub(r" _\r?\n\s*", " ", f.read())
        for name, params in re.findall(r"(?m)^Public Function (\w+)\(([^)]*)\)", src):
            if name == "NewEngine":
                continue
            args = []
            for p in filter(None, (x.strip() for x in params.split(","))):
                optional = p.startswith("Optional ") or p.startswith("ParamArray ")
                p = re.sub(r"^(Optional |ParamArray |ByVal |ByRef )+", "", p)
                args.append((p.split()[0].split("(")[0], optional))
            sigs[name] = args
    problems = []
    listed = [f[0] for f in info]
    if sorted(listed) != sorted(sigs):
        problems.append(f"functions {sorted(listed)} instead of {sorted(sigs)}")
    for f in info:
        name, description, pairs = f[0], f[1], list(zip(f[2::2], f[3::2]))
        if len(description) > 255:
            problems.append(f"{name}: description of {len(description)} characters")
        want = sigs.get(name, [])
        got = [(a.strip("[].").strip(), a.startswith("[")) for a, _ in pairs]
        if got != want:
            problems.append(f"{name}: arguments {got} instead of {want}")
        for a, d in pairs:
            if len(d) > 255:
                problems.append(f"{name}: description of {a} of {len(d)} characters")
    return problems


def main():
    failures = 0

    def check(what, got, expected):
        nonlocal failures
        ok = got == expected
        failures += not ok
        print(f"{'ok' if ok else 'DIFFERS':8s} {what}")
        if not ok:
            print(f"         expected {expected!r}\n         got      {got!r}")

    check("ribbon: callbacks and ids", ribbon_problems(), [])
    ctx, proc, work = start_office()
    library = tempfile.mkdtemp(prefix="asf-addin-library-")
    for name, text in LIBRARY.items():
        with open(os.path.join(library, name), "w", encoding="utf-8") as f:
            f.write(text)
    try:
        smgr = ctx.ServiceManager
        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        addin = os.path.join(ROOT, "excel", "addin")
        mods = build.library_modules(ROOT)
        mods += [addin_module(os.path.join(addin, f), library) for f in sorted(os.listdir(addin)) if f.endswith(".bas")]
        mods += [addin_module(os.path.join(ROOT, "examples", "use-cases", "regex", "RegexFunctions.bas"), library)]
        mods += [("AddinCheck", build.normal_module(CHECK), "normal")]
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        call = lambda fn, *args: provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.AddinCheck.{fn}?language=Basic&location=application").invoke(args, (), ())[0]
        for case in CASES:
            rows, columns = shape(case["expected"])
            got = call("CallFn", case["fn"], tuple(to_basic(a) for a in case["args"]), rows, columns)
            ok = table(got) == table(case["expected"])
            failures += not ok
            print(f"{'ok' if ok else 'DIFFERS':8s} {case['fn']}: {case['name']}")
            if not ok:
                print(f"         expected {table(case['expected'])}\n         got      {table(got)}   ({call('LastMessage')})")
        # the library as the ribbon uses it
        check("ribbon: the scripts listed", call("ScriptNames"), RIBBON_NAMES)
        for source, key, value in HEADERS:
            check(f"header {key} of {source!r}", call("Header", source, key), value)
        for file, sel, message in RIBBON_RUNS:
            got = call("RibbonRun", file, to_basic(sel) if sel is not None else None)
            if message is None:
                check(f"Run script {file}: not there", got.startswith("No script"), True)
            else:
                check(f"Run script {file} with {to_basic(sel)}", got, message)
        nested = os.path.join(library, "made", "by", "EnsureFolder")
        check("the library folder made with its parents", (call("MakeFolder", nested), os.path.isdir(nested)), (True, True))
        # a hidden folder on the way, as AppData is on Windows. Here it is a folder whose name starts with
        # a dot, which LibreOffice on Linux lists anyway: this checks the walk, not Windows' hidden flag
        os.makedirs(os.path.join(library, ".hidden"))
        under_hidden = os.path.join(library, ".hidden", "ASF", "scripts")
        check("the library folder made under a hidden folder", (call("MakeFolder", under_hidden), os.path.isdir(under_hidden)), (True, True))
        check("a hidden folder that is there already", call("MakeFolder", os.path.join(library, ".hidden")), True)
        check("registration: every function with its arguments", registration_problems(call("Registered")), [])
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
        shutil.rmtree(library, ignore_errors=True)
    print("all cases match" if not failures else f"{failures} case(s) differ")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
