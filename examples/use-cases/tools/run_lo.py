"""Checks the use cases in LibreOffice, headless: the scripts and modules on the sample workbooks.

    python3 examples/use-cases/tools/run_lo.py

Builds the ASF library from src/ (libreoffice/tools/build.py) and, with the modules of the use
cases, runs: the discount rules of rules/discounts.xlsx through RulesHost's BuildConditions and
CheckRules and rules/discounts.vas, and four mistakes in a rule; pipeline/sales-summary.vas on
pipeline/sales.xlsx; and the functions of regex/RegexFunctionsCalc.bas in the cells of a
spreadsheet. Compares with tools/data.py and exits with 0 when everything matches. The Excel-only
parts (ApplyDiscountRules, SummarizeSales, WriteRows) are not run; the functions of
regex/RegexFunctions.bas are called from Basic. Needs LibreOffice and python3-uno.
"""
import os, shutil, socket, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(BASE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, HERE)
import build  # noqa: E402
import data  # noqa: E402

CHECK = '''Option VBASupport 1
' rows of a sheet (an array of arrays) as Range.Value2 gives them: a 2D array from 1
Function Grid(rows As Variant) As Variant
    Dim r As Long, c As Long, g() As Variant
    ReDim g(1 To UBound(rows) - LBound(rows) + 1, 1 To UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1)
    For r = 1 To UBound(g, 1)
        For c = 1 To UBound(g, 2)
            g(r, c) = rows(LBound(rows) + r - 1)(LBound(rows(LBound(rows))) + c - 1)
        Next c
    Next r
    Grid = g
End Function
' the discounts, as ApplyDiscountRules computes them, or the message it shows
Function Discounts(rulesRows As Variant, ordersRows As Variant, script As String) As Variant
    Dim rules As Variant, orders As Variant, problem As String
    rules = Grid(rulesRows): orders = Grid(ordersRows)
    problem = RulesHost.CheckRules(rules)
    If Len(problem) > 0 Then Discounts = problem: Exit Function
    Dim eng As New ASF, result As Variant
    result = eng.Run(eng.Compile(RulesHost.BuildConditions(rules) & eng.ReadTextFile(script)), orders, rules)
    If IsEmpty(result) Then Discounts = eng.LastError Else Discounts = ASFSheets.ToGrid(result)
End Function
' a function of RegexFunctions.bas (the Excel module) as text, #VALUE! for an error value
Function Rx(fn As String, a As Variant) As String
    Dim r As Variant
    Select Case fn & UBound(a)
        Case "RX_TEST1": r = RegexFunctions.RX_TEST(a(0), a(1))
        Case "RX_TEST2": r = RegexFunctions.RX_TEST(a(0), a(1), a(2))
        Case "RX_EXTRACT1": r = RegexFunctions.RX_EXTRACT(a(0), a(1))
        Case "RX_EXTRACT2": r = RegexFunctions.RX_EXTRACT(a(0), a(1), a(2))
        Case "RX_REPLACE2": r = RegexFunctions.RX_REPLACE(a(0), a(1), a(2))
        Case "RX_COUNT1": r = RegexFunctions.RX_COUNT(a(0), a(1))
    End Select
    If IsError(r) Then Rx = "#VALUE!" Else Rx = CStr(r)
End Function
' the two tables of the sales summary, as SummarizeSales writes them
Function Summary(salesRows As Variant, script As String) As Variant
    Dim eng As New ASF, result As Variant
    result = eng.Run(eng.Compile(eng.ReadTextFile(script)), Grid(salesRows))
    If IsEmpty(result) Then Summary = eng.LastError: Exit Function
    Summary = Array(ASFSheets.ToGrid(result(LBound(result))), ASFSheets.ToGrid(result(LBound(result) + 1)))
End Function
'''


def start_office():
    import uno
    from com.sun.star.connection import NoConnectException
    work = tempfile.mkdtemp(prefix="asf-use-cases-")
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             "-env:UserInstallation=" + uno.systemPathToFileUrl(os.path.join(work, "profile")),   # also a Windows path
                             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                            env=build.soffice_env(uno.systemPathToFileUrl(os.path.join(work, "profile"))))
    resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            build.check_profile(ctx, work)
            return ctx, proc, work
        except NoConnectException:
            time.sleep(0.5)
    proc.kill()
    sys.exit("could not connect to soffice")


def prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def norm(v):
    """A cell or a value for comparing: numbers to 10 significant digits, the rest as text."""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return "%.10g" % v
    return str(v)


def table(t):
    return [[norm(v) for v in row] for row in t]


def module(name, path):
    with open(path, encoding="utf-8") as f:
        src = f.read()
    return (name, build.normal_module(build.resolve_conditionals(src)), "normal")


def main():
    ctx, proc, work = start_office()
    import uno
    failures = 0

    def check(what, got, expected):
        nonlocal failures
        ok = got == expected
        failures += not ok
        print(f"{'ok' if ok else 'DIFFERS':8s} {what}")
        if not ok:
            print(f"         expected {expected}\n         got      {got}")

    docs = []
    try:
        smgr = ctx.ServiceManager
        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        mods = build.library_modules(ROOT) + [module("ASFSheets", os.path.join(BASE, "ASFSheets.bas")),
                                              module("RulesHost", os.path.join(BASE, "rules", "RulesHost.bas")),
                                              module("RegexFunctions", os.path.join(BASE, "regex", "RegexFunctions.bas")),
                                              ("UseCaseCheck", build.normal_module(CHECK), "normal")]
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        call = lambda fn, *args: provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.UseCaseCheck.{fn}?language=Basic&location=application").invoke(args, (), ())[0]
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def book(path):
            d = desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, (prop("Hidden", True),))
            docs.append(d)
            return d

        def rows(d, name):
            sh = d.getSheets().getByName(name)
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            end = cur.getRangeAddress()
            return sh.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow).getDataArray()

        # 1. rules users edit
        rules_book = book(os.path.join(BASE, "rules", "discounts.xlsx"))
        rules, orders = rows(rules_book, "Rules"), rows(rules_book, "Orders")
        script = os.path.join(BASE, "rules", "discounts.vas")
        check("rules: discounts of the sample orders", table(call("Discounts", rules, orders, script)), table(data.DISCOUNTS))
        for when, message in data.RULE_MISTAKES:
            bad = (rules[0], (rules[1][0], when, rules[1][2]))
            check(f"rules: mistake {when!r}", call("Discounts", bad, orders, script), message)

        # 5. sales summary
        sales_book = book(os.path.join(BASE, "pipeline", "sales.xlsx"))
        regions, months = call("Summary", rows(sales_book, "Sales"), os.path.join(BASE, "pipeline", "sales-summary.vas"))
        want_regions, want_months = data.sales_summary(data.sales())
        check("pipeline: revenue by region", table(regions), table(want_regions))
        check("pipeline: revenue by month", table(months), table(want_months))

        # 3. regex functions: the Excel module called from Basic, then the Calc module in cells
        for fn, args, expected in data.REGEX:
            got = call("Rx", fn, tuple(args))
            want = ("True" if expected else "False") if isinstance(expected, bool) else norm(expected)
            check(f"regex (Excel module): {fn}({', '.join(repr(a) for a in args)})", got, want)
        calc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))
        docs.append(calc)
        # in My Macros (the application's Standard library), where Calc also finds functions
        with open(os.path.join(BASE, "regex", "RegexFunctionsCalc.bas"), encoding="utf-8") as f:
            std = libs.getByName("Standard") if libs.hasByName("Standard") else libs.createLibrary("Standard")
            std.insertByName("RegexFunctions", f.read())
            libs.loadLibrary("Standard")
        sh = calc.getSheets().getByIndex(0)
        for i, (fn, args, _) in enumerate(data.REGEX):
            def arg(a):
                if isinstance(a, bool):
                    return "TRUE()" if a else "FALSE()"
                if isinstance(a, (int, float)):
                    return str(a)
                return '"' + a.replace('"', '""') + '"'
            sh.getCellByPosition(0, i).setFormula("=%s(%s)" % (fn, ";".join(arg(a) for a in args)))
        calc.calculateAll()
        for i, (fn, args, expected) in enumerate(data.REGEX):
            c = sh.getCellByPosition(0, i)
            if c.getError() != 0:
                got = "#VALUE!" if c.getError() == 519 else f"error {c.getError()}"
            elif isinstance(expected, bool):
                got = c.getValue() != 0
            elif isinstance(expected, (int, float)):
                got = c.getValue()
            else:
                got = c.getString()
            check(f"regex (Calc cells): {fn}({', '.join(repr(a) for a in args)})", norm(got), norm(expected))
    finally:
        for d in docs:
            try:
                d.close(True)
            except Exception:
                pass
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    print("all use cases match" if not failures else f"{failures} check(s) differ")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
