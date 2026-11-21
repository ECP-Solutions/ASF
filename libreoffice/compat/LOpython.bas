' Python for ASF in LibreOffice: PY_CALL in cells and py for scripts, through asfpy_lo.py, the
' Python script of the ASF.oxt extension (excel/python/DESIGN.md, phase 5). The rules are those of
' the Excel add-in: PY_CALL and py.call reach functions marked @export in the Python folder only,
' and py.module, any module, only at level 2. The host decides what a script gets:
'
'   engine.InjectHostObject "py", LOpython.PyRoot(1)     ' py.call
'   engine.InjectHostObject "py", LOpython.PyRoot(2)     ' py.call and py.module
'   LOpython.PyRoot(1).Proxy.call("module.function", 2)  ' from Basic itself
'
' The script runs from the extension (location=user:uno_packages or share:uno_packages), never
' from a document: LibreOffice can embed Python in a document, and that comes with the document.

' The function fn of the extension's script; an error when the extension isn't installed
Private Function PyScript(ByVal fn As String) As Object
    Dim pip As Object, url As String, location As String, uri As String
    pip = GetDefaultContext().getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    url = pip.getPackageLocation("com.ecp-solutions.asf")
    If url = "" Then
        Err.Raise 5, "LOpython", "the ASF extension (ASF.oxt) isn't installed: Python comes with it"
    End If
    If InStr(url, "UNO_SHARED_PACKAGES") > 0 Then location = "share:uno_packages" Else location = "user:uno_packages"
    uri = "vnd.sun.star.script:" & Mid(url, InStrRev(url, "/") + 1) & "|python|asfpy_lo.py$" & fn & _
          "?language=Python&location=" & location
    Set PyScript = CreateUnoService("com.sun.star.script.provider.MasterScriptProviderFactory") _
        .createScriptProvider("").getScript(uri)
End Function

' fn of asfpy_lo.py with the arguments of the array args
Private Function PyRun(ByVal fn As String, args As Variant) As Variant
    Dim outIndex() As Variant, outArgs() As Variant, r As Variant
    r = PyScript(fn).invoke(args, outIndex, outArgs)
    If IsObject(r) Then Set PyRun = r Else PyRun = r
End Function

' py for an ASF script: level 1 py.call of exported functions, level 2 py.module too. An ASF_PyObject:
' Basic code that calls py itself takes its Proxy (LOpython.PyRoot(1).Proxy.call("m.f", 2))
Public Function PyRoot(ByVal level As Long) As Object
    Dim r As Variant
    LOhelper.PyValue PyRun("root", Array(level)), r
    Set PyRoot = r
End Function

' The Python folder ("" for the default: %APPDATA%\ASF\python on Windows, ~/.config/ASF/python
' elsewhere): its modules are imported, and PY_CALL and py find their functions there
Public Function PySetFolder(ByVal folder As String) As String
    PySetFolder = PyRun("set_folder", Array(folder))
End Function

' The Python folder's modules imported again, after they changed; what the load found
Public Function PyReload() As String
    PyReload = PyRun("reload", Array())
End Function

' What the last load of the Python folder found, and the version of Python
Public Function PyStartInfo() As String
    PyStartInfo = PyRun("start_info", Array())
End Function

' Python's last error, one line: from PY_CALL or py
Public Function PyLastError() As String
    PyLastError = PyRun("last_error", Array())
End Function

' =PY_CALL("module.function"; arguments...) in a cell: an exported function of the Python folder.
' A range arrives as a list of rows; a list of rows back fills the cells of an array formula. Calc
' takes no error value from a Basic function, so an unknown name gives the text #NAME?, an exception
' #VALUE! (PyLastError says why). Basic code calls PyCall, which gives error values (CVErr).
Public Function PY_CALL(name As Variant, Optional a1 As Variant, Optional a2 As Variant, Optional a3 As Variant, _
        Optional a4 As Variant, Optional a5 As Variant, Optional a6 As Variant, Optional a7 As Variant, _
        Optional a8 As Variant, Optional a9 As Variant, Optional a10 As Variant, Optional a11 As Variant, _
        Optional a12 As Variant, Optional a13 As Variant, Optional a14 As Variant, Optional a15 As Variant) As Variant
    Dim r As Variant
    r = PyCallArgs(name, Array(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15))
    If IsError(r) Then
        Select Case CLng(r)
            Case 2029: PY_CALL = "#NAME?"
            Case 2036: PY_CALL = "#NUM!"
            Case Else: PY_CALL = "#VALUE!"
        End Select
    Else
        PY_CALL = r
    End If
End Function

' PY_CALL for Basic code: the same call, with Excel's error values (CVErr(2029) #NAME?, CVErr(2015)
' #VALUE!, ...) for an unknown name or an exception
Public Function PyCall(name As Variant, Optional a1 As Variant, Optional a2 As Variant, Optional a3 As Variant, _
        Optional a4 As Variant, Optional a5 As Variant, Optional a6 As Variant, Optional a7 As Variant, _
        Optional a8 As Variant, Optional a9 As Variant, Optional a10 As Variant, Optional a11 As Variant, _
        Optional a12 As Variant, Optional a13 As Variant, Optional a14 As Variant, Optional a15 As Variant) As Variant
    PyCall = PyCallArgs(name, Array(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15))
End Function

' name with the arguments of given (Optional parameters, missing or not): a value, rows as a 2-D
' array, or an error value
Private Function PyCallArgs(name As Variant, given As Variant) As Variant
    Dim missing(1 To 15) As Boolean, n As Long, i As Long, args() As Variant
    For i = 1 To 15
        missing(i) = IsMissing(given(LBound(given) + i - 1))
    Next i
    ' the arguments left out at the end are dropped; one left out between two is empty
    n = 15
    Do While n > 0
        If Not missing(n) Then Exit Do
        n = n - 1
    Loop
    If n = 0 Then
        args = Array()
    Else
        ReDim args(0 To n - 1)
        For i = 1 To n
            If Not missing(i) Then args(i - 1) = RowsOf(given(LBound(given) + i - 1))
        Next i
    End If
    Dim r As Variant
    r = PyRun("py_call", Array(name, args))
    Select Case r(0)
        Case 0: PyCallArgs = r(1)
        Case 1: PyCallArgs = Grid(r(1))
        Case Else: PyCallArgs = CVErr(2000 + r(1))    ' 2029 #NAME?, 2015 #VALUE!, as Excel's xlErrName...
    End Select
End Function

' A range as rows for Python: Calc passes a 2-D array (from 1), which goes as an array of rows
Private Function RowsOf(v As Variant) As Variant
    Dim out() As Variant, r As Long, d As Long
    If IsArray(v) Then
        On Error Resume Next
        d = UBound(v, 2)
        If Err.Number <> 0 Then d = -1
        On Error GoTo 0
    End If
    If Not IsArray(v) Or d = -1 Then     ' a value, or an array of one dimension: as it is
        RowsOf = v
        Exit Function
    End If
    ReDim out(0 To UBound(v, 1) - LBound(v, 1))
    For r = LBound(v, 1) To UBound(v, 1)
        out(r - LBound(v, 1)) = RowOf(v, r)
    Next r
    RowsOf = out
End Function

' Row r of a 2-D array, as an array from 0 of its own: a row array ReDim'd again for each row reaches
' Python as None, though Basic still sees it
Private Function RowOf(v As Variant, ByVal r As Long) As Variant
    Dim row() As Variant, c As Long
    ReDim row(0 To UBound(v, 2) - LBound(v, 2))
    For c = LBound(v, 2) To UBound(v, 2)
        row(c - LBound(v, 2)) = v(r, c)
    Next c
    RowOf = row
End Function

' Python's rows as the 2-D array an array formula takes
Private Function Grid(rows As Variant) As Variant
    Dim g() As Variant, r As Long, c As Long, nr As Long, nc As Long
    nr = UBound(rows) - LBound(rows) + 1
    nc = UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1
    ReDim g(1 To nr, 1 To nc)
    For r = 1 To nr
        For c = 1 To nc
            g(r, c) = rows(LBound(rows) + r - 1)(LBound(rows(LBound(rows))) + c - 1)
        Next c
    Next r
    Grid = g
End Function
