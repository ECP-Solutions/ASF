Attribute VB_Name = "ASFEngine"
' The engine the worksheet functions share: made on first use with the settings (ASFSettings), it
' keeps the scripts it compiled; ResetEngine drops it.
Option Explicit
Option Private Module

Private engine_ As ASF
Private runEnabled_ As Boolean   ' the settings, read when the engine is made
Private folder_ As String
Private sources_ As Collection   ' library scripts read: Array(file time, source), by path
Private textOf_ As Long          ' the program that writes a value as a template does
Private lastError_ As String     ' why the last call gave #VALUE!

' ASF_RUN: a script written in a formula or a cell
Public Function RunScript(ByVal script As String, ByRef args As Variant) As Variant
    Ready
    If Not runEnabled_ Then
        lastError_ = "ASF_RUN is turned off in the settings of the add-in: only the scripts of the library run (ASF_CALL)"
        RunScript = CVErr(XL_ERR_VALUE)
        Exit Function
    End If
    RunScript = RunSource(script, args)
End Function

' ASF_CALL: name.vas from the library folder
Public Function RunLibraryScript(ByVal name As String, ByRef args As Variant) As Variant
    Ready
    Dim path As String: path = LibraryScript(name)
    If LenB(path) = 0 Then
        lastError_ = "'" & name & "' is not the name of a script of the library"
        RunLibraryScript = CVErr(XL_ERR_VALUE)
        Exit Function
    End If
    Dim source As String
    If Not LibrarySource(path, source) Then
        lastError_ = "No script " & path
        RunLibraryScript = CVErr(XL_ERR_VALUE)
        Exit Function
    End If
    RunLibraryScript = RunSource(source, args)
End Function

' Why the last call of ASF_RUN or ASF_CALL gave #VALUE! (empty when it did not)
Public Function LastErrorMessage() As String
    LastErrorMessage = lastError_
End Function

' Drops the engine of the worksheet functions, its compiled scripts and the settings it read
Public Sub ResetEngine()
    Set engine_ = Nothing
    Set sources_ = Nothing
    textOf_ = 0
End Sub

' The text of a value as a template writes it: objects and nested arrays in cells
Public Function ValueText(ByRef v As Variant) As String
    On Error GoTo Failed
    If textOf_ = 0 Then textOf_ = engine_.Compile("return `${$1}`;")
    ValueText = CStr(engine_.RunArgs(textOf_, Array(v)))
    Exit Function
Failed:
    ValueText = TypeName(v)
End Function

Private Sub Ready()
    If Not engine_ Is Nothing Then Exit Sub
    folder_ = ASFSettings.LibraryFolder()
    runEnabled_ = ASFSettings.RunEnabled()
    Set engine_ = New ASF
    engine_.AppAccess = False
    engine_.WorkingDir = folder_
    engine_.ImportRoot = folder_
    engine_.MaxSteps = ASFSettings.StepLimit()
    Set sources_ = New Collection
    textOf_ = 0
End Sub

Private Function RunSource(ByVal source As String, ByRef args As Variant) As Variant
    Dim values() As Variant, n As Long, i As Long, idx As Long, compiled As Boolean, result As Variant
    On Error GoTo Failed
    n = UBound(args) - LBound(args) + 1
    If n > 0 Then
        ReDim values(1 To n)
        For i = 1 To n
            values(i) = ASFCells.ArgumentValue(args(LBound(args) + i - 1))
            ' an error value as an argument is the result, as in Excel's own functions
            If IsError(values(i)) Then
                lastError_ = vbNullString
                RunSource = values(i)
                Exit Function
            End If
        Next i
    End If
    idx = engine_.Compile(source)
    compiled = True
    If n > 0 Then
        SetValue result, engine_.RunArgs(idx, values)
    Else
        SetValue result, engine_.RunArgs(idx, Empty)
    End If
    If Not IsEmpty(engine_.LastError) Then
        If IsObject(engine_.LastError) Then lastError_ = "a throw of " & ValueText(engine_.LastError) Else lastError_ = CStr(engine_.LastError)
        RunSource = CVErr(XL_ERR_VALUE)
        Exit Function
    End If
    lastError_ = vbNullString
    RunSource = FitToCaller(ASFCells.CellResult(result))
    Exit Function
Failed:
    lastError_ = Err.Description
    ' a compile error leaves the engine as it was; any other may leave it half way through a run
    If compiled Then ResetEngine
    RunSource = CVErr(XL_ERR_VALUE)
End Function

' The result for the cells of the formula: an array formula over several cells (Excel 2016 and
' 2019, or entered with Ctrl+Shift+Enter) gets the result fitted to them
Private Function FitToCaller(ByRef result As Variant) As Variant
    Dim rows As Long, columns As Long, caller As Object
    On Error Resume Next
    Set caller = Application.Caller
    If Not caller Is Nothing Then
        rows = caller.Rows.Count
        columns = caller.Columns.Count
    End If
    On Error GoTo 0
    FitToCaller = ASFCells.FitTo(result, rows, columns)
End Function

' The path of name.vas in the library, or "" when name is not the bare name of a file
Private Function LibraryScript(ByVal name As String) As String
    name = Trim$(name)
    If LenB(name) = 0 Then Exit Function
    If InStr(name, "/") > 0 Or InStr(name, "\") > 0 Or InStr(name, ":") > 0 Or InStr(name, "..") > 0 Then Exit Function
    If LCase$(Right$(name, 4)) <> ".vas" Then name = name & ".vas"
    LibraryScript = engine_.JoinPath(folder_, name)
End Function

' The source of a library script, read again when its file changed; False when there is no file
Private Function LibrarySource(ByVal path As String, ByRef source As String) As Boolean
    Dim stamp As Date, entry As Variant
    On Error GoTo Missing
    stamp = FileDateTime(path)
    On Error Resume Next
    entry = sources_(path)
    On Error GoTo 0
    If IsArray(entry) Then
        If entry(0) = stamp Then
            source = entry(1)
            LibrarySource = True
            Exit Function
        End If
        sources_.Remove path
    End If
    source = engine_.ReadTextFile(path)
    sources_.Add Array(stamp, source), path
    LibrarySource = True
    Exit Function
Missing:
End Function

Private Sub SetValue(ByRef dst As Variant, ByRef v As Variant)
    If IsObject(v) Then Set dst = v Else dst = v
End Sub
