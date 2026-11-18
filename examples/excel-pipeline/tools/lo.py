"""Runs ASF_RUN and ASF_CALL in LibreOffice, headless, with the add-in's modules and this pipeline's
library folder: what check.py uses.

    with Office(root) as call:
        call("ASF_CALL", "clean", table, "Qty:number")

root is the ASF repository (its src/ and excel/ folders). A table is a list of rows; a result over
several cells comes back as a list of rows, an error value as "#ERR:" and its text.
"""
import os, shutil, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
LIBRARY = os.path.join(os.path.dirname(HERE), "library")

CHECK = '''Option VBASupport 1
Private Function CellOf(v As Variant) As Variant
    If VarType(v) = vbString Then
        If Left(v, 5) = "#ERR:" Then CellOf = CVErr(CLng(Mid(v, 6))): Exit Function
    End If
    CellOf = v
End Function
' a range as Range.Value2 gives it: one cell its value, more a 2D array from 1
Private Function Arg(v As Variant) As Variant
    If Not IsArray(v) Then Arg = CellOf(v): Exit Function
    Dim n As Long, m As Long, r As Long, c As Long, g() As Variant, row As Variant
    n = UBound(v) - LBound(v) + 1
    row = v(LBound(v))
    m = UBound(row) - LBound(row) + 1
    If n = 1 And m = 1 Then Arg = CellOf(row(LBound(row))): Exit Function
    ReDim g(1 To n, 1 To m)
    For r = 1 To n
        row = v(LBound(v) + r - 1)
        For c = 1 To m
            g(r, c) = CellOf(row(LBound(row) + c - 1))
        Next c
    Next r
    Arg = g
End Function
Function CallAsf(fn As String, args As Variant) As Variant
    Dim a(0 To 6) As Variant, n As Long, i As Long, r As Variant
    n = UBound(args) - LBound(args) + 1
    For i = 0 To n - 1
        a(i) = Arg(args(LBound(args) + i))
    Next i
    Select Case n
        Case 1: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0)) Else r = ASFFunctions.ASF_CALL(a(0))
        Case 2: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0), a(1)) Else r = ASFFunctions.ASF_CALL(a(0), a(1))
        Case 3: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0), a(1), a(2)) Else r = ASFFunctions.ASF_CALL(a(0), a(1), a(2))
        Case 4: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0), a(1), a(2), a(3)) Else r = ASFFunctions.ASF_CALL(a(0), a(1), a(2), a(3))
        Case 5: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0), a(1), a(2), a(3), a(4)) Else r = ASFFunctions.ASF_CALL(a(0), a(1), a(2), a(3), a(4))
        Case 6: If fn = "ASF_RUN" Then r = ASFFunctions.ASF_RUN(a(0), a(1), a(2), a(3), a(4), a(5)) Else r = ASFFunctions.ASF_CALL(a(0), a(1), a(2), a(3), a(4), a(5))
    End Select
    If IsError(r) Then CallAsf = "#ERR:" & ASFCells.ErrorText(r) & " " & ASFEngine.LastErrorMessage(): Exit Function
    CallAsf = r
End Function
'''


class Office:
    def __init__(self, root):
        self.root = os.path.abspath(root)

    def __enter__(self):
        sys.path.insert(0, os.path.join(self.root, "libreoffice", "tools"))
        sys.path.insert(0, os.path.join(self.root, "excel", "tools"))
        sys.path.insert(0, os.path.join(self.root, "Tools", "stdRegex"))
        import build
        import run_lo as addin
        from BenchmarkLO import start_office
        self.ctx, self.proc, self.work = start_office()
        smgr = self.ctx.ServiceManager
        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", self.ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        folder = os.path.join(self.root, "excel", "addin")
        mods = build.library_modules(self.root)
        mods += [addin.addin_module(os.path.join(folder, f), LIBRARY) for f in sorted(os.listdir(folder)) if f.endswith(".bas")]
        mods += [("PipelineCheck", build.normal_module(CHECK), "normal")]
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx).createScriptProvider("")
        script = provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.PipelineCheck.CallAsf?language=Basic&location=application")

        def call(fn, *args):
            return script.invoke((fn, tuple(to_basic(a) for a in args)), (), ())[0]
        return call

    def __exit__(self, *exc):
        try:
            self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        self.proc.kill()
        shutil.rmtree(self.work, ignore_errors=True)


def to_basic(v):
    if isinstance(v, (list, tuple)):
        rows = [r if isinstance(r, (list, tuple)) else [r] for r in v]
        return tuple(tuple(r) for r in rows)
    return v
