Attribute VB_Name = "AsfPySpike"
' The checks of phases 1a to 3 in Excel (excel/python/spike/README.md). Import this module into
' any workbook (not into ASF.xlam), then run AsfPySetup once and AsfPySpikeTest with
' asfpy64.xll or asfpy32.xll loaded. The results go to a new workbook.
Option Explicit

#If Win64 Then
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy64.xll" (ByVal level As Long, ByRef out As Variant) As Long
#Else
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy32.xll" (ByVal level As Long, ByRef out As Variant) As Long
#End If

Private ws As Object, row As Long, failures As Long, lastFailure As String, lastRaised As String

' Sets the venv (or Python install) and the Python folder that the add-in reads when it loads
Public Sub AsfPySetup()
    Dim venv As String, folder As String
    venv = InputBox("Folder of the venv (or of a Python install) whose bitness matches Excel's:", _
                    "ASF Python", GetSetting("ASF", "AddIn", "PythonVenv", ""))
    If LenB(venv) = 0 Then Exit Sub
    folder = InputBox("Python folder, with spike.py in it:", "ASF Python", _
                      GetSetting("ASF", "AddIn", "PythonFolder", Environ$("APPDATA") & "\ASF\python"))
    If LenB(folder) = 0 Then Exit Sub
    SaveSetting "ASF", "AddIn", "PythonVenv", venv
    SaveSetting "ASF", "AddIn", "PythonFolder", folder
    MsgBox "Saved. Now load the add-in (File > Options > Add-ins > Excel Add-ins > Browse), or restart Excel if it was loaded.", vbInformation, "ASF Python"
End Sub

Public Sub AsfPySpikeTest()
    Dim py As Variant, py2 As Variant, m As Variant, v As Variant, a(1 To 2, 1 To 2) As Variant
    Dim i As Long, t As Double, code As Long, numpy As Boolean, x As Variant, s2 As String, folder As String
    Set ws = Workbooks.Add.Worksheets(1)
    ws.Name = "Spike"
    row = 1
    failures = 0
    Put3 "ASF Python, phases 1a to 3", Application.Name & " " & Application.Version & ", " & IIf(Is64(), "64-bit", "32-bit"), ""

    ' --- py from VBA
    On Error Resume Next
    code = AsfPy_Root(1, py)
    If Err.Number <> 0 Then
        Put3 "AsfPy_Root", "", "Error " & Err.Number & ": " & Err.Description & ". Is the add-in loaded?"
        Exit Sub
    End If
    On Error GoTo 0
    If code <> 0 Then
        Put3 "AsfPy_Root", "", "Python didn't start: " & py & ". See " & Environ$("TEMP") & "\asfpy.log"
        Exit Sub
    End If
    Put3 "Start-up of the Python folder", CallPy(py, "startInfo", VbGet), ""
    ' each call goes through CallPy or PyObj: an error shows in the sheet, with its number and text
    Check "py.call add", CallPy(py, "call", VbMethod, "spike.add", 2, 40), 42
    v = CallPy(py, "call", VbMethod, "spike.echo", ChrW(937) & "mega", 7)
    If IsArray(v) Then
        Check "py.call echo: an array back", LBound(v) & ":" & UBound(v) & ":" & v(0) & ":" & v(1), "0:1:" & ChrW(937) & "mega:7"
    Else
        Check "py.call echo: an array back", v, "an array"
    End If
    a(1, 1) = 11: a(1, 2) = 12: a(2, 1) = 21: a(2, 2) = 22
    Check "py.call mean of a 2-D array", CallPy(py, "call", VbMethod, "spike.mean", a), 16.5
    Check "py.call refused: os.system", ErrorOf(py, "call", "os.system", "calc"), "NameError: no exported function 'os.system' in the Python folder"
    Put3 "  what CallByName raised for it", lastRaised, ""
    Check "py.lastError", Left$(CallPy(py, "lastError", VbGet), 9), "NameError"
    Check "py.module at level 1", ErrorOf(py, "module", "math", Empty), "AttributeError: '_Root' has no member 'module'"
    Check "py.__class__", ErrorOf(py, "__class__", Empty, Empty), "AttributeError: '__class__' can't be used"
    code = AsfPy_Root(2, py2)
    ' what a plain late-bound call raises, as against CallByName above (information only)
    On Error Resume Next
    v = py2.module("_hidden")
    Put3 "  what a late-bound call raised for py.module('_hidden')", Err.Number & ": " & Err.Description, ""
    On Error GoTo 0
    Set m = PyObj(py2, "module", "math")
    If m Is Nothing Then
        Check "py.module('math')", lastFailure, "an object"
    Else
        Check "py.module('math').sqrt(2)", RoundIf(CallPy(m, "sqrt", VbMethod, 2)), 1.414214
        Check "math.PI, by case", RoundIf(CallPy(m, "PI", VbGet)), 3.141593
    End If

    ' --- phase 1b: For Each and the default value, through VBA itself
    Set m = PyObj(py, "call", "spike.squares", 4)
    If m Is Nothing Then
        Check "For Each over a generator", lastFailure, "0,1,4,9,"
    Else
        s2 = ""
        On Error Resume Next
        For Each x In m
            s2 = s2 & x & ","
        Next x
        If Err.Number <> 0 Then s2 = "error " & Err.Number & ": " & Err.Description
        On Error GoTo 0
        Check "For Each over a generator", s2, "0,1,4,9,"
    End If
    Set m = PyObj(py2, "module", "collections")
    If Not m Is Nothing Then Set m = PyObj(m, "Counter", "abcab")
    If m Is Nothing Then
        Check "x = obj: a Counter's text", lastFailure, "Counter({'a': 2, 'b': 2, 'c': 1})"
    Else
        On Error Resume Next
        v = m
        If Err.Number <> 0 Then v = "error " & Err.Number & ": " & Err.Description
        On Error GoTo 0
        Check "x = obj: a Counter's text", v, "Counter({'a': 2, 'b': 2, 'c': 1})"
    End If
    t = Timer
    For i = 1 To 10000
        v = CallPy(py, "call", VbMethod, "spike.add", 1, 2)
    Next i
    Put3 "py.call from VBA, microseconds per call", Format((Timer - t) * 100, "0.0"), ""
    Set m = PyObj(py2, "module", "numpy")
    numpy = Not m Is Nothing
    If numpy Then
        Check "numpy through py: percentile", CallPy(m, "percentile", VbMethod, CallPy(py, "call", VbMethod, "spike.echo", 1, 2), 50), 1.5
    Else
        Put3 "numpy", "not available: " & lastFailure, ""
    End If

    ' --- worksheet functions
    Dim s As Object: Set s = ws.Parent.Worksheets.Add(After:=ws)
    s.Name = "Formulas"
    s.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    s.Range("B1:C3").Value = s.Range("A1:A3").Value
    s.Range("A10:A19").Formula = "=ROW()-9"
    Formula s, "E1", "=spike_version()", Empty
    Check "spike_version", Left$(s.Range("E1").Text, 9), "Python 3."
    Formula s, "E2", "=spike_add(2,3.5)", 5.5
    Formula s, "E3", "=spike_args3(1,""x"",)", "1.0 | 'x' | None"
    Formula s, "E4", "=spike_sum7(1,2,3,,5,6,7)", 24
    Formula s, "E5", "=spike_shape(A1:C3)", "3 rows x 3 columns"
    Formula s, "E6", "=spike_shape(A1:A3)", "list of 3"
    Formula s, "E7", "=spike_fail()", CVErr(xlErrValue)
    Formula s, "E8", "=spike_sqrt(-1)", CVErr(xlErrValue)
    Formula s, "E9", "=spike_sqrt(1/0)", CVErr(xlErrDiv0)
    Formula s, "E10", "=spike_sqrt(16)", 4
    Formula s, "E11", "=spike_text(""" & ChrW(937) & "mega"")", ChrW(937) & "mega " & ChrW(183) & " " & ChrW(252) & " " & ChrW(183) & " " & ChrW(&H4E2D) & ChrW(&H6587) & " " & ChrW(183) & " 5 chars"
    s.Range("G1:I2").FormulaArray = "=spike_grid(2,3)"
    Check "spike_grid(2,3), an array formula", Join2(s.Range("G1:I2").Value), "11,12,13;21,22,23"
    s.Range("K1:K3").FormulaArray = "=spike_column(A1:A3)"
    Check "spike_column(A1:A3)", Join2(s.Range("K1:K3").Value), "10;20;30"
    If numpy Then Formula s, "E12", "=spike_numpy(A10:A19)", 9.1

    ' --- recalculation: 2,000 cells
    s.Range("M1:M2000").Formula = "=spike_add(ROW(),1)"
    ' the clock first: with automatic calculation, Dirty calculates the cells at once
    t = Timer
    Recalc s.Range("M1:M2000")
    Put3 "2,000 cells of spike_add, ms to calculate", Format((Timer - t) * 1000, "0"), ""
    Check "M2000", s.Range("M2000").Value, 2001

    ' --- phase 1b: the type suffixes in formulas
    Formula s, "E13", "=spike_thread()", "MainThread"
    Formula s, "E14", "=spike_refused()", CVErr(xlErrName)
    s.Range("E15").Formula = "=spike_volatile()"
    v = s.Range("E15").Value
    Application.Calculate
    Check "=spike_volatile() counts each recalculation", IIf(s.Range("E15").Value > v, "more", "the same: " & s.Range("E15").Value & " after " & v), "more"
    ' a thread-safe function on Excel's calculation threads: 20,000 cells, multi-threaded
    Put3 "Multi-threaded calculation", IIf(Application.MultiThreadedCalculation.Enabled, Application.MultiThreadedCalculation.ThreadCount & " threads", "off: File > Options > Advanced > Formulas"), ""
    s.Range("O1:O20000").Formula = "=spike_threadsafe(ROW())"
    t = Timer
    Recalc s.Range("O1:O20000")
    Put3 "20,000 cells of spike_threadsafe, ms to calculate", Format((Timer - t) * 1000, "0"), ""
    s.Range("P1").Formula = "=SUMPRODUCT(--(O1:O20000=2*ROW(O1:O20000)))"
    Check "spike_threadsafe: right results of 20,000", s.Range("P1").Value, 20000

    ' --- phase 2: PY_CALL, type hints, Why this error?
    Formula s, "E17", "=PY_CALL(""spike.add"",2,40)", 42
    s.Range("R1:R3").FormulaArray = "=PY_CALL(""spike.scale"",A1:A3,10)"
    Check "=PY_CALL(""spike.scale"",A1:A3,10), an array formula", Join2(s.Range("R1:R3").Value), "10;20;30"
    Formula s, "E18", "=PY_CALL(""os.system"",""calc"")", CVErr(xlErrName)
    Formula s, "E19", "=spike_hints(2,DATE(2023,3,15),12,7)", "2 | datetime.date(2023, 3, 15) | '12' | [7.0]"
    s.Range("E20").Formula = "=spike_nextday(DATE(2023,3,15))"
    Check "=spike_nextday(DATE(2023,3,15)), a serial number", s.Range("E20").Value2, 45001
    ' what Why this error? will do: forget the last error, calculate the cell again, read it
    CallPy py, "clearError", VbMethod
    Recalc s.Range("E8")
    Check "Why this error? for =spike_sqrt(-1)", CallPy(py, "lastError", VbGet), "ValueError: expected a nonnegative input, got -1.0"
    Check "  its traceback names the function", IIf(InStr(CallPy(py, "lastTraceback", VbGet), "spike_sqrt") > 0, "yes", "no"), "yes"

    ' --- phase 2: AsfPyReload, with a module added, changed and removed in the Python folder
    folder = GetSetting("ASF", "AddIn", "PythonFolder", Environ$("APPDATA") & "\ASF\python")
    WriteReloadModule folder & "\aa_reload.py", 1
    Check "AsfPyReload, aa_reload.py added", Reload(), "done"
    Formula s, "E21", "=reload_check()", 1
    Formula s, "E22", "=PY_CALL(""aa_reload.version"")", 1
    WriteReloadModule folder & "\aa_reload.py", 22
    Check "AsfPyReload, aa_reload.py changed", Reload(), "done"
    Put3 "  =reload_check() before the cell is calculated again", s.Range("E21").Value, ""
    Recalc s.Range("E21:E22")
    Check "=reload_check() after the change", s.Range("E21").Value, 22
    Check "=PY_CALL(""aa_reload.version"") after the change", s.Range("E22").Value, 22
    Kill folder & "\aa_reload.py"
    Check "AsfPyReload, aa_reload.py removed", Reload(), "done"
    Recalc s.Range("E22")
    Check "=PY_CALL(""aa_reload.version"") after the removal", s.Range("E22").Value, CVErr(xlErrName)
    Formula s, "E23", "=spike_add(2,3.5)", 5.5

    ' --- phase 3: library scripts with py, through ASF.xlam
    Phase3Checks s

    ' --- phase 1b: the add-in unticked and ticked again (xlAutoClose, xlAutoOpen)
    Dim ai As Object
    For Each ai In Application.AddIns
        If LCase$(ai.Name) Like "asfpy*.xll" And ai.Installed Then Exit For
    Next ai
    If ai Is Nothing Then
        Put3 "Unticked and ticked again", "the add-in isn't in Application.AddIns: check by hand (README)", ""
    Else
        ' unticked, Excel doesn't call the function: an error (#VALUE!, as Excel keeps the name), not 5.5
        ai.Installed = False
        s.Range("E16").Formula = "=spike_add(2,3.5)"
        v = s.Range("E16").Value
        Check "Unticked: =spike_add(2,3.5) isn't calculated", IIf(IsError(v), "an error", v), "an error"
        Put3 "  the error", CStr(v), ""
        ai.Installed = True
        s.Range("E16").Formula = "=spike_add(2,3.5)"
        Check "Ticked again: =spike_add(2,3.5)", s.Range("E16").Value, 5.5
        Check "Ticked again: py.call still works", CallPy(py, "call", VbMethod, "spike.add", 2, 40), 42
    End If

    Put3 IIf(failures = 0, "All checks pass", failures & " checks fail"), "", ""
    Put3 "The add-in's log", Environ$("TEMP") & "\asfpy.log", ""
    ws.Columns("A:C").AutoFit
    ws.Activate
End Sub

' Phase 3: ASF_CALL of library scripts that ask for Python, with PythonAccess off, call and modules,
' and ASF_RUN, which never gets py. Needs ASF.xlam of phase 3 loaded next to the Python add-in. The
' scripts go to the library folder and are deleted afterwards; PythonAccess is put back as it was.
Private Sub Phase3Checks(s As Object)
    Dim lib As String, saved As String, wb As Object
    On Error Resume Next
    Set wb = Workbooks("ASF.xlam")
    On Error GoTo 0
    If wb Is Nothing Then
        Put3 "Phase 3", "ASF.xlam isn't loaded: its checks are skipped", ""
        Exit Sub
    End If
    lib = GetSetting("ASF", "AddIn", "LibraryFolder", Environ$("APPDATA") & "\ASF\scripts")
    saved = GetSetting("ASF", "AddIn", "PythonAccess", "")
    WriteText lib & "\py_check_call.vas", "// @asf python: call" & vbLf & "return py.call('spike.add', $1, 1);"
    WriteText lib & "\py_check_modules.vas", "// @asf python: modules" & vbLf & "return py.module('spike').Which();"
    WriteText lib & "\py_check_fails.vas", "// @asf python: call" & vbLf & "return py.call('spike.add', 'x', 1);"
    On Error GoTo PutBack
    SetPythonAccess "off"
    Formula s, "G20", "=ASF_CALL(""py_check_call"",41)", CVErr(xlErrValue)
    Check "  why, with PythonAccess off", WhyText(s.Range("G20")), _
          "the script asks for Python (python: call), which the setting PythonAccess (off) doesn't allow: it needs call or modules"
    SetPythonAccess "call"
    Formula s, "G21", "=ASF_CALL(""py_check_call"",41)", 42
    Formula s, "G22", "=ASF_CALL(""py_check_modules"")", CVErr(xlErrValue)
    Check "  why, modules with PythonAccess call", WhyText(s.Range("G22")), _
          "the script asks for Python (python: modules), which the setting PythonAccess (call) doesn't allow: it needs modules"
    SetPythonAccess "modules"
    ' the name as written: a lower-cased one would call which() and give "which"
    Formula s, "G23", "=ASF_CALL(""py_check_modules"")", "Which"
    Formula s, "G24", "=ASF_CALL(""py_check_fails"")", CVErr(xlErrValue)
    Check "  why: the script's message and Python's", InStr(WhyText(s.Range("G24")), "(Python: TypeError") > 0, True
    Formula s, "G25", "=ASF_RUN(""// @asf python: call""&CHAR(10)&""return 1;"")", CVErr(xlErrValue)
    Check "  why: ASF_RUN never gets py", WhyText(s.Range("G25")), _
          "Python is for the scripts of the library (ASF_CALL): a script in a formula can't ask for it"
Done:
    On Error Resume Next
    If LenB(saved) = 0 Then
        DeleteSetting "ASF", "AddIn", "PythonAccess"
        Application.Run "ASF.xlam!ASFEngine.ResetEngine"
    Else
        SetPythonAccess saved
    End If
    Kill lib & "\py_check_call.vas"
    Kill lib & "\py_check_modules.vas"
    Kill lib & "\py_check_fails.vas"
    Exit Sub
PutBack:
    Put3 "Phase 3 stopped", "error " & Err.Number & ": " & Err.Description, "FAIL"
    failures = failures + 1
    Resume Done
End Sub

Private Sub SetPythonAccess(ByVal value As String)
    Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "PythonAccess", value
End Sub

' what Why this error? shows for a cell with ASF_RUN or ASF_CALL: the cell calculated again, then
' the add-in's message
Private Function WhyText(cell As Object) As String
    Recalc cell
    WhyText = CStr(Application.Run("ASF.xlam!ASFEngine.LastErrorMessage"))
End Function

Private Sub WriteText(path As String, text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub

' a range calculated again, as Ctrl+Alt+F9 does: Range.Calculate alone skips cells that aren't
' dirty, such as a cell whose function the add-in just reloaded
Private Sub Recalc(r As Object)
    r.Dirty
    r.Calculate
End Sub

' a module for the reload checks: reload_check() and aa_reload.version() both return n
Private Sub WriteReloadModule(path As String, n As Long)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "from asfpy import export, xl_func"
    Print #f, ""
    Print #f, ""
    Print #f, "@xl_func"
    Print #f, "def reload_check():"
    Print #f, "    return " & n
    Print #f, ""
    Print #f, ""
    Print #f, "@export"
    Print #f, "def version():"
    Print #f, "    return " & n
    Close #f
End Sub

' the add-in's command AsfPyReload: "done", or the error
Private Function Reload() As String
    On Error GoTo Failed
    Application.Run "AsfPyReload"
    Reload = "done"
    Exit Function
Failed:
    Reload = "error " & Err.Number & ": " & Err.Description
End Function

Private Function Is64() As Boolean
#If Win64 Then
    Is64 = True
#End If
End Function

Private Sub Put3(a As String, b As Variant, c As Variant)
    ws.Cells(row, 1).Value = Label(a)
    ws.Cells(row, 2).Value = b
    ws.Cells(row, 3).Value = c
    row = row + 1
End Sub

Private Sub Check(what As String, got As Variant, want As Variant)
    Dim ok As Boolean
    If IsError(got) Or IsError(want) Then
        ok = IsError(got) And IsError(want)
        If ok Then ok = (CStr(got) = CStr(want))
    Else
        ok = (CStr(got) = CStr(want))
    End If
    If Not ok Then failures = failures + 1
    ws.Cells(row, 1).Value = Label(what)
    ws.Cells(row, 2).Value = IIf(IsError(got), CStr(got), got)
    ws.Cells(row, 3).Value = IIf(ok, "ok", "FAIL, expected " & CStr(want))
    row = row + 1
End Sub

Private Sub Formula(s As Object, cell As String, f As String, want As Variant)
    s.Range(cell).Formula = f
    If Not IsEmpty(want) Then Check f, s.Range(cell).Value, want
End Sub

' obj.name(args) through CallByName, for a value; on an error, "error <number>: <text>"
Private Function CallPy(obj As Variant, name As String, kind As VbCallType, ParamArray a() As Variant) As Variant
    On Error GoTo Failed
    Select Case UBound(a) - LBound(a) + 1
        Case 0: CallPy = CallByName(obj, name, kind)
        Case 1: CallPy = CallByName(obj, name, kind, a(0))
        Case 2: CallPy = CallByName(obj, name, kind, a(0), a(1))
        Case Else: CallPy = CallByName(obj, name, kind, a(0), a(1), a(2))
    End Select
    Exit Function
Failed:
    CallPy = "error " & Err.Number & ": " & Err.Description
    lastFailure = CallPy
End Function

' obj.name(args) through CallByName, for an object; Nothing on an error, kept in lastFailure
Private Function PyObj(obj As Variant, name As String, ParamArray a() As Variant) As Object
    On Error GoTo Failed
    Select Case UBound(a) - LBound(a) + 1
        Case 0: Set PyObj = CallByName(obj, name, VbMethod)
        Case 1: Set PyObj = CallByName(obj, name, VbMethod, a(0))
        Case Else: Set PyObj = CallByName(obj, name, VbMethod, a(0), a(1))
    End Select
    Exit Function
Failed:
    lastFailure = "error " & Err.Number & ": " & Err.Description
    Set PyObj = Nothing
End Function

Private Function RoundIf(v As Variant) As Variant
    If IsNumeric(v) Then RoundIf = Round(v, 6) Else RoundIf = v
End Function

' A call of py that Python refuses, through CallByName as ASF makes it. VBA raises an error (in
' Excel, error 440 "Automation error", without Python's text), and py.lastError names the
' exception: the result is that text, or "no error". What VBA raised goes to lastRaised.
Private Function ErrorOf(py As Variant, name As String, a1 As Variant, a2 As Variant) As String
    Dim v As Variant, raised As Boolean
    On Error Resume Next
    If IsEmpty(a1) Then
        v = CallByName(py, name, VbGet)
    ElseIf IsEmpty(a2) Then
        v = CallByName(py, name, VbMethod, a1)
    Else
        v = CallByName(py, name, VbMethod, a1, a2)
    End If
    raised = Err.Number <> 0
    lastRaised = IIf(raised, Err.Number & ": " & Err.Description, "no error")
    On Error GoTo 0
    If raised Then ErrorOf = CallPy(py, "lastError", VbGet) Else ErrorOf = "no error"
End Function

' a text that starts with = would become a formula in the cell
Private Function Label(s As String) As String
    If Left$(s, 1) = "=" Then Label = "'" & s Else Label = s
End Function

Private Function Join2(v As Variant) As String
    Dim r As Long, c As Long, out As String
    For r = LBound(v, 1) To UBound(v, 1)
        If r > LBound(v, 1) Then out = out & ";"
        For c = LBound(v, 2) To UBound(v, 2)
            If c > LBound(v, 2) Then out = out & ","
            If IsError(v(r, c)) Then out = out & CStr(v(r, c)) Else out = out & v(r, c)
        Next c
    Next r
    Join2 = out
End Function
