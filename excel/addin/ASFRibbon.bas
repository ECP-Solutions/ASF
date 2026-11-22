Attribute VB_Name = "ASFRibbon"
' The callbacks of the ASF tab (excel/customUI/customUI14.xml). They are Subs, so they don't show up
' as worksheet functions; control is the ribbon's IRibbonControl, declared As Object so that the
' module compiles without the Office library.
Option Explicit

Private Const TITLE As String = "ASF"
Private Const HELP_URL As String = "https://github.com/ECP-Solutions/ASF/blob/main/excel/README.md"
Private Const MAX_LISTED As Long = 30   ' an InputBox prompt holds about 1024 characters

' Run script: picks a script of the library and runs it, with the selection when its header asks
Public Sub OnRunScript(control As Object)
    Dim scripts As Variant, prompt As String, i As Long, n As Long, choice As String, item As Variant, sel As Variant
    scripts = ASFLibrary.LibraryScripts()
    If IsEmpty(scripts) Then
        MsgBox "There are no scripts (.vas files) in the library folder:" & vbLf & ASFSettings.LibraryFolder(), vbInformation, TITLE
        Exit Sub
    End If
    n = UBound(scripts) - LBound(scripts) + 1
    If n > MAX_LISTED Then n = MAX_LISTED
    For i = 1 To n
        item = scripts(LBound(scripts) + i - 1)
        prompt = prompt & i & ". " & item(1) & vbLf
    Next i
    choice = InputBox(prompt & vbLf & "Number of the script to run:", TITLE & ": Run script")
    If Not IsNumeric(choice) Then Exit Sub
    If CLng(choice) < 1 Or CLng(choice) > n Then Exit Sub
    item = scripts(LBound(scripts) + CLng(choice) - 1)
    If TypeName(Selection) = "Range" Then Set sel = Selection Else sel = Empty
    MsgBox ASFLibrary.RunForRibbon(item(0), sel), vbInformation, TITLE & ": " & item(1)
End Sub

' Library folder: opens it, made first when it is missing
Public Sub OnOpenLibrary(control As Object)
    Dim folder As String: folder = ASFSettings.LibraryFolder()
    If Not ASFLibrary.EnsureFolder(folder) Then
        MsgBox "The library folder can't be made:" & vbLf & folder, vbExclamation, TITLE
        Exit Sub
    End If
    On Error GoTo ShowPath
    ThisWorkbook.FollowHyperlink folder
    Exit Sub
ShowPath:
    MsgBox "The library folder is" & vbLf & folder, vbInformation, TITLE
End Sub

' Why this error?: calculates the active cell again and shows why its ASF_RUN or ASF_CALL gave
' #VALUE!, or, when the Python add-in is loaded, the Python error of its PY_CALL or Python function
Public Sub OnWhyError(control As Object)
    Dim cell As Object, formula As String, message As String, isAsf As Boolean, py As Object, notLoaded As String
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell with ASF_RUN, ASF_CALL or a Python function.", vbInformation, TITLE
        Exit Sub
    End If
    Set cell = ActiveCell
    formula = UCase$(cell.Formula)
    isAsf = InStr(formula, "ASF_RUN(") > 0 Or InStr(formula, "ASF_CALL(") > 0
    Set py = ASFPython.Root(1, notLoaded)
    If Not isAsf And py Is Nothing Then
        MsgBox "The formula of " & cell.Address(False, False) & " doesn't call ASF_RUN or ASF_CALL.", vbInformation, TITLE
        Exit Sub
    End If
    If Not py Is Nothing Then ASFPython.ClearError py
    ' the cell's own call, so that the arguments and the fitting are its own. Dirty first: Calculate
    ' alone skips a cell that isn't dirty, and then nothing runs
    cell.Dirty
    cell.Calculate
    If isAsf Then message = ASFEngine.LastErrorMessage()
    If LenB(message) = 0 And Not py Is Nothing Then message = ASFPython.ErrorReport(py)
    If LenB(message) = 0 And IsError(cell.Value) Then
        MsgBox cell.Address(False, False) & ": the error comes from an argument of the formula, not from " & _
               IIf(py Is Nothing, "the script.", "ASF or Python."), vbInformation, TITLE
    ElseIf LenB(message) = 0 Then
        MsgBox cell.Address(False, False) & ": no error from ASF" & IIf(py Is Nothing, ".", " or Python."), vbInformation, TITLE
    Else
        MsgBox cell.Address(False, False) & ": " & message, vbExclamation, TITLE
    End If
End Sub

' Python folder: opens the Python add-in's folder of modules, made first when it is missing
Public Sub OnPythonFolder(control As Object)
    Dim folder As String: folder = ASFSettings.PythonFolder()
    If LenB(folder) = 0 Then
        MsgBox "Python needs the Python add-in, which runs on Excel for Windows.", vbInformation, TITLE
        Exit Sub
    End If
    If Not ASFLibrary.EnsureFolder(folder) Then
        MsgBox "The Python folder can't be made:" & vbLf & folder, vbExclamation, TITLE
        Exit Sub
    End If
    On Error GoTo ShowPath
    ThisWorkbook.FollowHyperlink folder
    Exit Sub
ShowPath:
    MsgBox "The Python folder is" & vbLf & folder, vbInformation, TITLE
End Sub

' Reload Python: the Python add-in imports the modules of its folder again (its command
' AsfPyReload), then the open workbooks are calculated in full: a cell keeps its old value otherwise
Public Sub OnReloadPython(control As Object)
    Dim py As Object, message As String, info As String
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        MsgBox message, vbExclamation, TITLE
        Exit Sub
    End If
    On Error GoTo Failed
    Application.Run "AsfPyReload"
    Application.CalculateFull
    info = CStr(CallByName(py, "startInfo", VbGet))
    If Len(info) > 900 Then info = Left$(info, 900) & "..."
    MsgBox "The Python folder was imported again, and the workbooks calculated." & vbLf & vbLf & info, vbInformation, TITLE
    Exit Sub
Failed:
    MsgBox "Reload Python failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Run Python script: a .py file run in Excel's Python, its output shown (ASFPythonSetup)
Public Sub OnRunPythonScript(control As Object)
    ASFPythonSetup.RunScript
End Sub

' Python macros: the commands (@xl_macro) of the Python folder; runs the one picked
Public Sub OnPythonMacros(control As Object)
    ASFPythonSetup.Macros
End Sub

' Python console: the Python add-in's Python at its prompt, in a console window (ASFPythonSetup)
Public Sub OnPythonConsole(control As Object)
    ASFPythonSetup.Console
End Sub

' Python settings: which Python, the Python folder, PythonAccess
Public Sub OnPythonSettings(control As Object)
    ASFPythonSetup.Settings
End Sub

' About Python: the add-in, its Python, the settings, the start-up and the log
Public Sub OnAboutPython(control As Object)
    ASFPythonSetup.About
End Sub

' Python packages: pip for the Python add-in's Python, in a console window (its command AsfPyPackages)
Public Sub OnPythonPackages(control As Object)
    Dim py As Object, message As String
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        MsgBox message, vbExclamation, TITLE
        Exit Sub
    End If
    On Error GoTo Failed
    Application.Run "AsfPyPackages"
    Exit Sub
Failed:
    MsgBox "Python packages failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Reset engine: drops the engine of the worksheet functions and their compiled scripts
Public Sub OnResetEngine(control As Object)
    ASFEngine.ResetEngine
    MsgBox "The engine was reset: the next formula compiles its script again and reads the settings again.", vbInformation, TITLE
End Sub

Public Sub OnHelp(control As Object)
    On Error GoTo ShowUrl
    ThisWorkbook.FollowHyperlink HELP_URL
    Exit Sub
ShowUrl:
    MsgBox HELP_URL, vbInformation, TITLE
End Sub
