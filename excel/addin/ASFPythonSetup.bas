Attribute VB_Name = "ASFPythonSetup"
' The Python buttons of the ASF tab, after the ribbon (ASFRibbon): Python studio, Run Python script,
' Python macros, Python console, Python settings and About Python. The Python add-in (asfpy64.xll,
' asfpy32.xll) does the work through its commands AsfPyRunScript, AsfPyMacros and AsfPyConsole; the settings are those it reads when
' Excel loads it (ASFSettings). Python runs on Excel for Windows only.
Option Explicit
Option Private Module

Private Const TITLE As String = "ASF Python"
Private Const MAX_SHOWN As Long = 900    ' a message box holds about 1024 characters
Private Const MAX_LISTED As Long = 30    ' and an InputBox prompt as much
Private Const MAX_ARGS As Long = 5       ' arguments typed for a macro of the list

' Run Python script: a .py file, picked in the Python folder, runs in Excel's Python as __main__;
' its output comes back. It may use Excel's object model (xl_app), as a command does.
Public Sub RunScript()
    Dim py As Object, path As String, out As Variant
    If Not Ready(py) Then Exit Sub
    path = PickFile(ASFSettings.PythonFolder())
    If LenB(path) = 0 Then Exit Sub
    On Error GoTo Failed
    out = Application.Run("AsfPyRunScript", path)
    On Error GoTo 0
    If IsError(out) Then
        MsgBox "The script didn't run: the Python add-in has no Python. About Python says why.", vbExclamation, TITLE
        Exit Sub
    End If
    ShowOutput "Run Python script: " & Mid$(path, InStrRev(path, "\") + 1), CStr(out), "The script ran and printed nothing."
    Exit Sub
Failed:
    MsgBox "Run Python script failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python macros: the commands (@xl_macro) of the Python folder; runs the one picked, asking for
' its arguments when it needs some
Public Sub Macros()
    Dim py As Object, list As Variant, lines() As String, parts() As String, prompt As String
    Dim i As Long, n As Long, choice As String, needed As Long, args As Variant, result As Variant
    If Not Ready(py) Then Exit Sub
    On Error GoTo Failed
    list = Application.Run("AsfPyMacros")
    On Error GoTo 0
    If IsError(list) Then Exit Sub
    If LenB(list) = 0 Then
        MsgBox "The Python folder has no macros. A function of a module there becomes one with @xl_macro:" & vbLf & vbLf & _
               "from asfpy import xl_macro" & vbLf & vbLf & "@xl_macro" & vbLf & "def tidy():" & vbLf & "    ...", vbInformation, TITLE
        Exit Sub
    End If
    lines = Split(list, vbLf)
    n = UBound(lines) + 1
    If n > MAX_LISTED Then n = MAX_LISTED
    For i = 0 To n - 1
        parts = Split(lines(i), vbTab)
        prompt = prompt & (i + 1) & ". " & parts(0)
        If parts(1) <> "0" Then prompt = prompt & " (" & parts(1) & " arguments)"
        If UBound(parts) >= 2 Then If LenB(parts(2)) > 0 Then prompt = prompt & ": " & parts(2)
        prompt = prompt & vbLf
    Next i
    choice = InputBox(prompt & vbLf & "Number of the macro to run:", TITLE & ": Python macros")
    If Not IsNumeric(choice) Then Exit Sub
    If CLng(choice) < 1 Or CLng(choice) > n Then Exit Sub
    parts = Split(lines(CLng(choice) - 1), vbTab)
    needed = CLng(parts(1))
    If needed > MAX_ARGS Then
        MsgBox parts(0) & " needs " & needed & " arguments: run it from VBA with Application.Run """ & parts(0) & """, ...", vbInformation, TITLE
        Exit Sub
    End If
    If needed > 0 Then
        args = MacroArguments(parts(0), needed)
        If IsEmpty(args) Then Exit Sub
    End If
    ASFPython.ClearError py
    On Error GoTo Failed
    Select Case needed
        Case 0: result = Application.Run(parts(0))
        Case 1: result = Application.Run(parts(0), args(0))
        Case 2: result = Application.Run(parts(0), args(0), args(1))
        Case 3: result = Application.Run(parts(0), args(0), args(1), args(2))
        Case 4: result = Application.Run(parts(0), args(0), args(1), args(2), args(3))
        Case Else: result = Application.Run(parts(0), args(0), args(1), args(2), args(3), args(4))
    End Select
    On Error GoTo 0
    If IsError(result) Then
        MsgBox parts(0) & " failed:" & vbLf & vbLf & Left$(ASFPython.ErrorReport(py), MAX_SHOWN), vbExclamation, TITLE
    ElseIf IsEmpty(result) Or IsArray(result) Then
        MsgBox parts(0) & " ran.", vbInformation, TITLE
    Else
        MsgBox parts(0) & " ran and gave:" & vbLf & vbLf & Left$(CStr(result), MAX_SHOWN), vbInformation, TITLE
    End If
    Exit Sub
Failed:
    MsgBox "Python macros failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python studio: the add-in's command AsfPyStudio opens it, or its window when it runs already
Public Sub Studio()
    Dim py As Object
    If Not Ready(py) Then Exit Sub
    On Error GoTo Failed
    Application.Run "AsfPyStudio"
    Exit Sub
Failed:
    MsgBox "Python studio failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python console: the add-in's Python at its prompt, in a console window of its own (its command
' AsfPyConsole), with pip("...") and the Python folder importable
Public Sub Console()
    Dim py As Object
    If Not Ready(py) Then Exit Sub
    On Error GoTo Failed
    Application.Run "AsfPyConsole"
    Exit Sub
Failed:
    MsgBox "Python console failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python settings: which Python, the Python folder and PythonAccess, each shown with its value;
' the add-in reads the first two when Excel loads it
Public Sub Settings()
    Dim choice As String, restart As Boolean
#If Mac Then
    MsgBox "Python needs the Python add-in, which runs on Excel for Windows.", vbInformation, TITLE
    Exit Sub
#End If
    Do
        choice = InputBox("1. Python: " & PythonInUse() & vbLf & _
                          "2. Python folder: " & ASFSettings.PythonFolder() & vbLf & _
                          "3. Python for library scripts (PythonAccess): " & ASFSettings.PythonAccess() & vbLf & vbLf & _
                          "Number of the setting to change (Cancel when done):", TITLE & ": settings")
        Select Case Trim$(choice)
            Case "": Exit Do
            Case "1": If ChoosePython() Then restart = True
            Case "2": If ChooseFolder() Then restart = True
            Case "3": ChooseAccess
        End Select
    Loop
    If restart Then MsgBox "Restart Excel to use the new settings: the Python add-in reads the Python and the Python folder when Excel loads it.", vbInformation, TITLE
End Sub

' About Python: the add-in, its Python, the settings and what the Python folder gave; offers the log
Public Sub About()
    Dim py As Object, message As String, text As String, info As String
    text = "Add-in: " & AddinPath() & vbLf & _
           "Python: " & PythonInUse() & vbLf & _
           "Python folder: " & ASFSettings.PythonFolder() & vbLf & _
           "PythonAccess: " & ASFSettings.PythonAccess() & vbLf & _
           "Log: " & LogPath() & vbLf & vbLf
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        text = text & message
    Else
        On Error Resume Next
        info = CStr(CallByName(py, "startInfo", VbGet))
        On Error GoTo 0
        If Len(info) > 600 Then info = Left$(info, 600) & "..."
        text = text & info
    End If
    If MsgBox(text & vbLf & vbLf & "Open the log?", vbYesNo + vbInformation + vbDefaultButton2, TITLE & ": about") = vbYes Then
        OpenText LogPath()
    End If
End Sub

' --- helpers ---------------------------------------------------------------------------------------

' the add-in loaded and Python running; else a message and False
Private Function Ready(ByRef py As Object) As Boolean
    Dim message As String
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        MsgBox message & vbLf & vbLf & "Python settings and About Python on the ASF tab help set it up.", vbExclamation, TITLE
    Else
        Ready = True
    End If
End Function

Private Function PythonInUse() As String
    Dim venv As String: venv = ASFSettings.PythonVenv()
    If LenB(venv) = 0 Then PythonInUse = "the add-in's own" Else PythonInUse = venv
End Function

Private Function LogPath() As String
    LogPath = Environ$("TEMP") & "\asfpy.log"
End Function

' the loaded asfpy*.xll, or where to load it from
Private Function AddinPath() As String
    Dim ai As Object
    On Error Resume Next
    For Each ai In Application.AddIns
        If LCase$(ai.Name) Like "asfpy*.xll" And ai.Installed Then
            AddinPath = ai.FullName
            Exit Function
        End If
    Next ai
    AddinPath = "not loaded (File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse)"
End Function

' Which Python: the add-in's own, or another (a venv or an install), picked here
Private Function ChoosePython() As Boolean
    Dim answer As Long, folder As String
    answer = MsgBox("Which Python should the Python add-in run?" & vbLf & vbLf & _
                    "Yes: the Python the add-in carries (recommended). Its packages come from Python packages." & vbLf & vbLf & _
                    "No: another Python of Excel's bitness, 3.11 or later: a venv or a python.org install, picked next.", _
                    vbYesNoCancel + vbQuestion, TITLE & ": Python")
    If answer = vbYes Then
        If LenB(ASFSettings.PythonVenv()) > 0 Then
            ASFSettings.ClearSetting "PythonVenv"
            ChoosePython = True
        End If
    ElseIf answer = vbNo Then
        folder = PickFolder("Folder of the venv or of the Python install", ASFSettings.PythonVenv())
        If LenB(folder) = 0 Then Exit Function
        If LenB(Dir(folder & "\pyvenv.cfg")) = 0 And LenB(Dir(folder & "\python3*.dll")) = 0 Then
            MsgBox folder & vbLf & "has neither pyvenv.cfg (a venv) nor python3XX.dll (an install): it isn't a Python.", vbExclamation, TITLE
            Exit Function
        End If
        ASFSettings.ChangeSetting "PythonVenv", folder
        ChoosePython = True
    End If
End Function

Private Function ChooseFolder() As Boolean
    Dim folder As String
    folder = PickFolder("Python folder: the modules whose functions Excel gets", ASFSettings.PythonFolder())
    If LenB(folder) = 0 Then Exit Function
    ASFSettings.ChangeSetting "PythonFolder", folder
    ChooseFolder = True
End Function

' PythonAccess applies at once: library scripts read it each time they run
Private Sub ChooseAccess()
    Dim choice As String, current As Long
    current = InStr("offcallmodules", ASFSettings.PythonAccess())
    current = IIf(current = 1, 1, IIf(current = 4, 2, 3))
    choice = InputBox("What library scripts whose header asks for Python may do:" & vbLf & vbLf & _
                      "1. off: no Python (the default)" & vbLf & _
                      "2. call: py.call of the functions exported with @export" & vbLf & _
                      "3. modules: py.module too, which imports any module: everything Python can do" & vbLf & vbLf & _
                      "Formulas (ASF_RUN) never get Python, whatever this says.", TITLE & ": PythonAccess", CStr(current))
    Select Case Trim$(choice)
        Case "1": ASFSettings.ChangeSetting "PythonAccess", "off"
        Case "2": ASFSettings.ChangeSetting "PythonAccess", "call"
        Case "3": ASFSettings.ChangeSetting "PythonAccess", "modules"
    End Select
End Sub

' the arguments of a macro, typed separated by commas: numbers as numbers, TRUE and FALSE as
' booleans, the rest as text; Empty when cancelled or too few
Private Function MacroArguments(ByVal name As String, ByVal needed As Long) As Variant
    Dim typed As String, parts() As String, args() As Variant, i As Long, v As String
    typed = InputBox(name & " needs " & needed & " arguments. Type them separated by commas:", TITLE & ": " & name)
    If LenB(typed) = 0 Then Exit Function
    parts = Split(typed, ",")
    If UBound(parts) + 1 < needed Then
        MsgBox name & " needs " & needed & " arguments.", vbExclamation, TITLE
        Exit Function
    End If
    ReDim args(0 To UBound(parts))
    For i = 0 To UBound(parts)
        v = Trim$(parts(i))
        If IsNumeric(v) Then
            args(i) = Val(v)
        ElseIf UCase$(v) = "TRUE" Or UCase$(v) = "FALSE" Then
            args(i) = (UCase$(v) = "TRUE")
        Else
            args(i) = v
        End If
    Next i
    MacroArguments = args
End Function

Private Function PickFile(ByVal folder As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)          ' msoFileDialogFilePicker
    fd.Title = TITLE & ": Python script to run"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Python scripts", "*.py"
    If LenB(folder) > 0 Then fd.InitialFileName = folder & "\"
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Private Function PickFolder(ByVal caption As String, ByVal start As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)          ' msoFileDialogFolderPicker
    fd.Title = TITLE & ": " & caption
    If LenB(start) > 0 Then fd.InitialFileName = start & "\"
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
End Function

' what a script printed: in a message box, or in Notepad when it's longer than one holds
Private Sub ShowOutput(ByVal caption As String, ByVal text As String, ByVal whenEmpty As String)
    Dim path As String, f As Integer
    If LenB(text) = 0 Then
        MsgBox whenEmpty, vbInformation, caption
    ElseIf Len(text) <= MAX_SHOWN Then
        MsgBox text, vbInformation, caption
    Else
        path = Environ$("TEMP") & "\asfpy-output.txt"
        f = FreeFile
        Open path For Output As #f
        Print #f, text;
        Close #f
        OpenText path
        MsgBox Left$(text, MAX_SHOWN) & "..." & vbLf & vbLf & "All of it is in " & path & ", opened in Notepad.", vbInformation, caption
    End If
End Sub

Private Sub OpenText(ByVal path As String)
    On Error Resume Next
    Shell "notepad.exe """ & path & """", vbNormalFocus
End Sub
