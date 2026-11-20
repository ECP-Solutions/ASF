Attribute VB_Name = "AsfPySpike"
' The checks of the phase 1a spike in Excel (excel/python/spike/README.md). Import this module into
' any workbook (not into ASF.xlam), then run AsfPySetup once and AsfPySpikeTest with
' asfpy64.xll or asfpy32.xll loaded. The results go to a new workbook.
Option Explicit

#If Win64 Then
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy64.xll" (ByVal level As Long, ByRef out As Variant) As Long
#Else
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy32.xll" (ByVal level As Long, ByRef out As Variant) As Long
#End If

Private ws As Object, row As Long, failures As Long, lastFailure As String

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
    Dim i As Long, t As Double, code As Long, numpy As Boolean
    Set ws = Workbooks.Add.Worksheets(1)
    ws.Name = "Spike"
    row = 1
    failures = 0
    Put3 "ASF Python, phase 1a", Application.Name & " " & Application.Version & ", " & IIf(Is64(), "64-bit", "32-bit"), ""

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
    Check "py.call refused: os.system", ErrorOf(py, "call", "os.system", "calc"), "NameError"
    Check "py.lastError", Left$(CallPy(py, "lastError", VbGet), 9), "NameError"
    Check "py.module at level 1", ErrorOf(py, "module", "math", Empty), "AttributeError"
    Check "py.__class__", ErrorOf(py, "__class__", Empty, Empty), "AttributeError"
    code = AsfPy_Root(2, py2)
    Set m = PyObj(py2, "module", "math")
    If m Is Nothing Then
        Check "py.module('math')", lastFailure, "an object"
    Else
        Check "py.module('math').sqrt(2)", RoundIf(CallPy(m, "sqrt", VbMethod, 2)), 1.414214
        Check "math.PI, by case", RoundIf(CallPy(m, "PI", VbGet)), 3.141593
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
    t = Timer
    s.Range("M1:M2000").Calculate
    Put3 "2,000 cells of spike_add, ms to calculate", Format((Timer - t) * 1000, "0"), ""
    Check "M2000", s.Range("M2000").Value, 2001

    Put3 IIf(failures = 0, "All checks pass", failures & " checks fail"), "", ""
    Put3 "The add-in's log", Environ$("TEMP") & "\asfpy.log", ""
    ws.Columns("A:C").AutoFit
    ws.Activate
End Sub

Private Function Is64() As Boolean
#If Win64 Then
    Is64 = True
#End If
End Function

Private Sub Put3(a As String, b As Variant, c As Variant)
    ws.Cells(row, 1).Value = a
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
    ws.Cells(row, 1).Value = what
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
        Case Else: Set PyObj = CallByName(obj, name, VbMethod, a(0))
    End Select
    Exit Function
Failed:
    lastFailure = "error " & Err.Number & ": " & Err.Description
    Set PyObj = Nothing
End Function

Private Function RoundIf(v As Variant) As Variant
    If IsNumeric(v) Then RoundIf = Round(v, 6) Else RoundIf = v
End Function

' the text of a Python error raised through CallByName, or "" when there was none
Private Function ErrorOf(obj As Variant, name As String, a1 As Variant, a2 As Variant) As String
    Dim v As Variant
    On Error Resume Next
    If IsEmpty(a1) Then
        v = CallByName(obj, name, VbGet)
    ElseIf IsEmpty(a2) Then
        v = CallByName(obj, name, VbMethod, a1)
    Else
        v = CallByName(obj, name, VbMethod, a1, a2)
    End If
    If Err.Number <> 0 Then ErrorOf = Split(Err.Description & ":", ":")(0)
    If Err.Number <> 0 And InStr(Err.Description, ":") = 0 Then ErrorOf = "error " & Err.Number & ": " & Err.Description
End Function

Private Function Join2(v As Variant) As String
    Dim r As Long, c As Long, out As String
    For r = LBound(v, 1) To UBound(v, 1)
        If r > LBound(v, 1) Then out = out & ";"
        For c = LBound(v, 2) To UBound(v, 2)
            If c > LBound(v, 2) Then out = out & ","
            out = out & v(r, c)
        Next c
    Next r
    Join2 = out
End Function
