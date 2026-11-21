"""Checks Python in the LibreOffice edition (phase 5 of excel/python/DESIGN.md): PY_CALL in Calc
cells and py in ASF scripts, through the Python script of ASF.oxt.

    python libreoffice/tools/run_python.py

Run it with a Python that imports uno (LibreOffice's own on Windows). It builds ASF.oxt from src/,
starts LibreOffice headless with a temporary profile, installs the extension there, points it at a
temporary Python folder with one module (lo_check.py, below) and runs the checks in Basic: py in
ASF scripts, and PY_CALL as Calc calls it. Calc cells themselves are not checked: in a headless
LibreOffice no Basic function runs from a cell (#VALUE!, even for a one-line function). Your own
profile and Python folder are not used. Exits with 0 when every check passes.
"""
import os
import shutil
import sys
import tempfile
import textwrap

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "Tools", "stdRegex"))
import build  # noqa: E402
from BenchmarkLO import start_office  # noqa: E402

LO_CHECK = '''
import collections
from asfpy import export, xl_func


@export
def add(a, b):
    return a + b


@export
def scale(values: list, factor: float):
    return [v * factor for v in values]


@export
def table(rows):
    """A range arrives as a list of rows; one column as a flat list, as in Excel."""
    return '%d x %d' % (len(rows), len(rows[0])) if isinstance(rows[0], list) else 'list of %d' % len(rows)


@export
def echo(x):
    return repr(x)


@export
def grid():
    return [[1, 'a'], [2, None]]


@export
def fails():
    raise ValueError('on purpose')


@export
def counter(text):
    return collections.Counter(text)


@xl_func
def double(x):
    """Exported too, as module.function: Calc has no @xl_func functions of its own."""
    return x * 2


def Which():
    return 'Which'


def which():
    return 'which'
'''

# Basic of the checks: each function returns text, which main() compares
CHECKS = '''Option VBASupport 1
Function Setup(folder As String) As String
    GlobalScope.BasicLibraries.LoadLibrary "ASF"
    Setup = LOpython.PySetFolder(folder)
End Function

Function LastError() As String
    LastError = LOpython.PyLastError()
End Function

Function StartInfo() As String
    StartInfo = LOpython.PyStartInfo()
End Function

' Basic calls py itself
Function Direct() As String
    Dim py As Object
    Set py = LOpython.PyRoot(1).Proxy
    Direct = CStr(py.call("lo_check.add", 2, 40)) & " " & py.lastError
End Function

' a script with py of a level
Function ViaAsf(ByVal level As Long, ByVal code As String) As String
    Dim e As Object, r As Variant
    Set e = LOhelper.NewASF()
    If level > 0 Then e.InjectHostObject "py", LOpython.PyRoot(level)
    On Error GoTo Failed
    r = e.Run(e.Compile(code))
    ViaAsf = CStr(r)
    If ViaAsf = "" Then ViaAsf = "error: " & CStr(e.LastError)
    Exit Function
Failed:
    ViaAsf = "raised: " & Err.Description & " | " & LOpython.PyLastError()
End Function

' PY_CALL (the cell function) or PyCall (for Basic) from Basic, as Calc calls it: a range as a 2-D
' array from 1 (rows "range" for A1:A3 = 1, 2, 3)
Function PyCallText(ByVal fn As String, ByVal name As String, Optional a As Variant, Optional b As Variant) As String
    Dim r As Variant, g(1 To 3, 1 To 1) As Variant, q(1 To 2, 1 To 2) As Variant
    g(1, 1) = 1: g(2, 1) = 2: g(3, 1) = 3
    q(1, 1) = 1: q(1, 2) = 2: q(2, 1) = 3: q(2, 2) = 4
    If Not IsMissing(a) Then
        If VarType(a) = vbString Then
            If a = "range" Then
                a = g
            ElseIf a = "square" Then
                a = q
            End If
        End If
    End If
    If fn = "PyCall" Then
        If IsMissing(a) Then
            r = LOpython.PyCall(name)
        ElseIf IsMissing(b) Then
            r = LOpython.PyCall(name, a)
        Else
            r = LOpython.PyCall(name, a, b)
        End If
    Else
        If IsMissing(a) Then
            r = LOpython.PY_CALL(name)
        ElseIf IsMissing(b) Then
            r = LOpython.PY_CALL(name, a)
        Else
            r = LOpython.PY_CALL(name, a, b)
        End If
    End If
    If IsError(r) Then
        PyCallText = "error " & CLng(r)
    ElseIf IsArray(r) Then
        Dim i As Long, j As Long, s As String
        For i = LBound(r, 1) To UBound(r, 1)
            If i > LBound(r, 1) Then s = s & ";"
            For j = LBound(r, 2) To UBound(r, 2)
                If j > LBound(r, 2) Then s = s & ","
                s = s & CStr(r(i, j))
            Next j
        Next i
        PyCallText = s
    Else
        PyCallText = CStr(r)
    End If
End Function
'''


def CommandEnvironment():
    """An XCommandEnvironment that approves what installing the extension asks (no dialogs)."""
    import uno
    import unohelper
    from com.sun.star.task import XInteractionHandler
    from com.sun.star.ucb import XCommandEnvironment, XProgressHandler

    class Progress(unohelper.Base, XProgressHandler):
        def push(self, status):
            pass

        def update(self, status):
            pass

        def pop(self):
            pass

    class Handler(unohelper.Base, XInteractionHandler):
        def handle(self, request):
            approve = uno.getTypeByName("com.sun.star.task.XInteractionApprove")
            for c in request.getContinuations():
                if c.queryInterface(approve):
                    c.select()
                    return

    class Env(unohelper.Base, XCommandEnvironment):
        def getInteractionHandler(self):
            return Handler()

        def getProgressHandler(self):
            return Progress()

    return Env()


def main():
    import uno
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ctx, proc, work = start_office()
    failures = []
    checks = 0

    def check(what, got, want, contains=False):
        nonlocal checks
        checks += 1
        ok = want in str(got) if contains else str(got) == want
        if not ok:
            failures.append(what)
        print("%s %s: %s" % ("ok  " if ok else "FAIL", what, got))

    try:
        smgr = ctx.ServiceManager
        # the extension, built from src/, in the temporary profile
        oxt = os.path.join(work, "ASF.oxt")
        with open(oxt, "wb") as f:
            f.write(build.extension_bytes(build.library_modules()))
        manager = ctx.getValueByName("/singletons/com.sun.star.deployment.ExtensionManager")
        try:
            manager.addExtension(uno.systemPathToFileUrl(oxt), (), "user", manager.createAbortChannel(), CommandEnvironment())
        except Exception as exc:
            log = os.path.join(work, "profile", "user", "uno_packages", "cache", "log.txt")
            if os.path.isfile(log):
                print(open(log, encoding="utf-8", errors="replace").read()[-3000:])
            sys.exit("ASF.oxt didn't install: %r" % (exc,))
        folder = os.path.join(work, "python")
        os.makedirs(folder)
        with open(os.path.join(folder, "lo_check.py"), "w", encoding="utf-8") as f:
            f.write(textwrap.dedent(LO_CHECK))

        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary("ASFPyChecks")
        lib.insertByName("Checks", CHECKS)
        libs.loadLibrary("ASFPyChecks")
        provider = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")

        def basic(fn, *args):
            script = provider.getScript(f"vnd.sun.star.script:ASFPyChecks.Checks.{fn}?language=Basic&location=application")
            return script.invoke(args, (), ())[0]

        check("the Python folder", basic("Setup", folder), folder)
        info = basic("StartInfo")
        check("Python of LibreOffice, lo_check.py imported", "imported lo_check.py" in info and info.startswith("Python 3"), "True")
        print("     " + info.split("\n")[0])

        # py from Basic, and in ASF scripts
        check("a script without py", basic("ViaAsf", 0, "return 40 + 2;"), "42")
        check("py.call from Basic (the proxy), and py.lastError", basic("Direct"), "42 ")
        check("py.call in a script", basic("ViaAsf", 1, "return py.call('lo_check.add', 2, 40);"), "42")
        check("py.call with an ASF array", basic("ViaAsf", 1, "let r = py.call('lo_check.scale', [1, 2, 3], 10); return r.length;"), "3")
        check("an object: typeof", basic("ViaAsf", 1, "return typeof py.call('lo_check.counter', 'abcab');"), "object: <Counter>")
        check("an object's method", basic("ViaAsf", 1, "let c = py.call('lo_check.counter', 'abcab'); return c.total();"), "5")
        check("an object back to Python", basic("ViaAsf", 1, "let c = py.call('lo_check.counter', 'abcab'); return py.call('lo_check.add', c, c).total();"), "10")
        check("py.lastError, a property", basic("ViaAsf", 1, "py.clearError(); return 'x' + py.lastError + 'x';"), "xx")
        check("an exception: the script fails with Python's message",
              basic("ViaAsf", 1, "return py.call('lo_check.fails');"), "ValueError: on purpose", contains=True)
        check("  py.lastError", basic("ViaAsf", 1, "return py.lastError;"), "ValueError: on purpose")
        check("py.call refused: os.system", basic("ViaAsf", 1, "py.call('os.system', 'calc'); return 'ran';"),
              "NameError: no exported function 'os.system' in the Python folder", contains=True)
        check("py.module at level 1", basic("ViaAsf", 1, "py.module('math'); return 'ran';"),
              "AttributeError: '_Root' has no member 'module'", contains=True)
        check("py.module at level 2", basic("ViaAsf", 2, "return py.module('math').sqrt(16);"), "4")
        check("a method's name as written", basic("ViaAsf", 2, "return py.module('lo_check').Which();"), "Which")
        check("_ names refused", basic("ViaAsf", 2, "py.module('_thread'); return 'ran';"), "can't be used", contains=True)

        # PY_CALL as Calc calls it: Calc gives a range as a 2-D array from 1, and takes a 2-D array
        # back into the cells of an array formula. Calc cells themselves aren't checked: in a
        # headless LibreOffice no Basic function runs from a cell, not even a one-line one (#VALUE!)
        check("PY_CALL", basic("PyCallText", "PY_CALL", "lo_check.add", 2, 40), "42")
        check("PY_CALL of an @xl_func function", basic("PyCallText", "PY_CALL", "lo_check.double", 4), "8")
        check("PY_CALL with a one-column range: a flat list", basic("PyCallText", "PY_CALL", "lo_check.table", "range"), "list of 3")
        check("PY_CALL with a 2 x 2 range: a list of rows", basic("PyCallText", "PY_CALL", "lo_check.table", "square"), "2 x 2")
        check("PY_CALL with a range, a column back", basic("PyCallText", "PY_CALL", "lo_check.scale", "range", 10), "10;20;30")
        check("PY_CALL: a grid back", basic("PyCallText", "PY_CALL", "lo_check.grid"), "1,a;2,")
        check("PY_CALL: unknown name, the text #NAME?", basic("PyCallText", "PY_CALL", "os.system", "calc"), "#NAME?")
        check("PY_CALL: an exception, the text #VALUE!", basic("PyCallText", "PY_CALL", "lo_check.fails"), "#VALUE!")
        check("PyCall: unknown name, CVErr(2029)", basic("PyCallText", "PyCall", "os.system", "calc"), "error 2029")
        check("PyCall: an exception, CVErr(2015)", basic("PyCallText", "PyCall", "lo_check.fails"), "error 2015")
        check("  PyLastError", basic("LastError"), "ValueError: on purpose")
    finally:
        try:
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=30)
        except Exception:
            proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    print("%d checks, %d failed" % (checks, len(failures)))
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
