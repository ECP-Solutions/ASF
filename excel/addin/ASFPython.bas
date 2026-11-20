Attribute VB_Name = "ASFPython"
' Python in the scripts of the library, through the Python add-in (asfpy64.xll or asfpy32.xll,
' excel/python). Two keys: the script asks for it with a header, and the setting PythonAccess
' (ASFSettings) decides.
'   // @asf python: call       py.call('module.function', args...): functions exported with @export
'   // @asf python: modules    py.module('name') too: any module of the venv
' A script that asks for more than PythonAccess allows doesn't run, with a message naming the
' setting. Scripts of formulas (ASF_RUN) never get py: anyone can send a workbook.
Option Explicit
Option Private Module

#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy64.xll" (ByVal level As Long, ByRef out As Variant) As Long
#ElseIf VBA7 Then
Private Declare PtrSafe Function AsfPy_Root Lib "asfpy32.xll" (ByVal level As Long, ByRef out As Variant) As Long
#End If

' The level of Python a script asks for, against the level allowed (0 off, 1 call, 2 modules):
' 0 when it asks for none, 1 or 2, or -1 when it can't run, with message saying why
Public Function ScriptLevel(ByVal source As String, ByVal allowed As Long, ByRef message As String) As Long
    Dim asked As String
    asked = LCase$(ASFLibrary.HeaderValue(source, "python"))
    Select Case asked
        Case vbNullString
            ScriptLevel = 0
            Exit Function
        Case "call"
            ScriptLevel = 1
        Case "modules"
            ScriptLevel = 2
        Case Else
            message = "the header '// @asf python: " & asked & "' takes call or modules"
            ScriptLevel = -1
            Exit Function
    End Select
    If ScriptLevel > allowed Then
        message = "the script asks for Python (python: " & asked & "), which the setting PythonAccess (" & _
                  LevelName(allowed) & ") doesn't allow: it needs " & IIf(ScriptLevel = 1, "call or modules", "modules")
        ScriptLevel = -1
    End If
End Function

' The level PythonAccess allows: 0 off, 1 call, 2 modules
Public Function AllowedLevel() As Long
    Select Case ASFSettings.PythonAccess()
        Case "call": AllowedLevel = 1
        Case "modules": AllowedLevel = 2
    End Select
End Function

Private Function LevelName(ByVal level As Long) As String
    Select Case level
        Case 1: LevelName = "call"
        Case 2: LevelName = "modules"
        Case Else: LevelName = "off"
    End Select
End Function

' True when the script asks for Python at all: ASF_RUN refuses it
Public Function AsksForPython(ByVal source As String) As Boolean
    AsksForPython = LenB(ASFLibrary.HeaderValue(source, "python")) > 0
End Function

' py of a level (1: py.call; 2: py.module too) from the Python add-in; Nothing, with message, when
' the add-in isn't loaded or Python didn't start
Public Function Root(ByVal level As Long, ByRef message As String) As Object
#If Mac Then
    message = "Python needs the Python add-in (asfpy64.xll or asfpy32.xll), which runs on Excel for Windows"
#ElseIf VBA7 Then
    Dim out As Variant, code As Long
    On Error GoTo NotLoaded
    code = AsfPy_Root(level, out)
    On Error GoTo 0
    If code = 0 And IsObject(out) Then
        Set Root = out
    Else
        message = "Python didn't start in the Python add-in: " & CStr(out)
    End If
    Exit Function
NotLoaded:
    message = "the Python add-in (" & AddinName() & ") is not loaded: File > Options > Add-ins > Manage: Excel Add-ins"
#Else
    message = "Python needs the Python add-in (asfpy64.xll or asfpy32.xll) of Excel for Windows"
#End If
End Function

Private Function AddinName() As String
#If Win64 Then
    AddinName = "asfpy64.xll"
#Else
    AddinName = "asfpy32.xll"
#End If
End Function

' Forgets Python's last error: an error after it is the script's own
Public Sub ClearError(py As Object)
    On Error Resume Next
    CallByName py, "clearError", VbMethod
End Sub

' Python's last error, one line; "" when there was none. VBA's CallByName reports a failed call of
' py as error 440 "Automation error" without Python's text: this is where the text is
Public Function LastError(py As Object) As String
    On Error Resume Next
    LastError = CStr(CallByName(py, "lastError", VbGet))
End Function

' Python's last error and the end of its traceback, for Why this error?
Public Function ErrorReport(py As Object) As String
    Dim trace As String
    ErrorReport = LastError(py)
    If LenB(ErrorReport) = 0 Then Exit Function
    On Error Resume Next
    trace = CStr(CallByName(py, "lastTraceback", VbGet))
    On Error GoTo 0
    ' a message box holds about 1024 characters: the end of the traceback is the useful part
    If Len(trace) > 700 Then trace = "..." & Right$(trace, 700)
    If LenB(trace) > 0 Then ErrorReport = ErrorReport & vbLf & vbLf & trace
End Function

' The message of a failed run of a script with py: the engine's, and Python's error when it has one
Public Function WithPythonError(ByVal message As String, py As Object) As String
    Dim pyError As String
    WithPythonError = message
    If py Is Nothing Then Exit Function
    pyError = LastError(py)
    If LenB(pyError) > 0 Then WithPythonError = message & " (Python: " & pyError & ")"
End Function
