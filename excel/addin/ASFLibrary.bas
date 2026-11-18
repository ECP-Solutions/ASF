Attribute VB_Name = "ASFLibrary"
' The scripts of the library folder, as the ribbon runs them. A script may start with a header of
' comment lines:
'   // @asf name: Apply discount rules     the name the ribbon lists (else the file name)
'   // @asf appAccess: true                 the script may use Office objects (default: false)
'   // @asf args: selection                 $1 is the selected range (default: no arguments)
' With appAccess, $1 is the Range itself; without, its values (ASFCells). The library holds the
' scripts the user installed: they may have Office access, unlike the scripts of formulas.
Option Explicit
Option Private Module

' The .vas files of the library, sorted by name: an array of Array(path, name); Empty when none
Public Function LibraryScripts() As Variant
    Dim folder As String, sep As String, f As String, n As Long, i As Long, files() As String, found() As Variant
    folder = ASFSettings.LibraryFolder()
    sep = PathSeparatorOf(folder)
    If Right$(folder, 1) <> sep Then folder = folder & sep
    On Error GoTo NoFolder
    ' Dir with a pattern does not work on Mac: every file, filtered here
    f = Dir(folder)
    On Error GoTo 0
    ' the names first: reading a script checks its file with Dir, which would end this listing
    Do While LenB(f) > 0
        If LCase$(Right$(f, 4)) = ".vas" Then
            n = n + 1
            ReDim Preserve files(1 To n)
            files(n) = f
        End If
        f = Dir()
    Loop
    If n = 0 Then Exit Function
    ReDim found(1 To n)
    For i = 1 To n
        found(i) = Array(folder & files(i), ScriptName(folder & files(i), files(i)))
    Next i
    SortByName found
    LibraryScripts = found
    Exit Function
NoFolder:
End Function

' A header value of source: "" when the script has no such line
Public Function HeaderValue(ByVal source As String, ByVal key As String) As String
    Dim lines As Variant, i As Long, t As String, p As Long
    lines = Split(Replace(source, vbCr, vbNullString), vbLf)
    For i = LBound(lines) To UBound(lines)
        t = Trim$(lines(i))
        If LenB(t) = 0 Then
        ElseIf Left$(t, 2) <> "//" Then
            Exit For
        Else
            t = Trim$(Mid$(t, 3))
            If LCase$(Left$(t, 5)) = "@asf " Then
                t = Trim$(Mid$(t, 6))
                p = InStr(t, ":")
                If p > 0 Then
                    If LCase$(Trim$(Left$(t, p - 1))) = LCase$(key) Then
                        HeaderValue = Trim$(Mid$(t, p + 1))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' Runs the library script at path, as the ribbon does; selection: the selected Range, or Empty.
' The message for the user: the result, or why the script stopped
Public Function RunForRibbon(ByVal path As String, ByRef selection As Variant) As String
    Dim eng As ASF, source As String, idx As Long, result As Variant, appAccess As Boolean, arg As Variant
    On Error GoTo Failed
    Set eng = New ASF
    source = eng.ReadTextFile(path)
    If LenB(source) = 0 Then
        RunForRibbon = "No script " & path
        Exit Function
    End If
    appAccess = (LCase$(HeaderValue(source, "appAccess")) = "true")
    eng.AppAccess = appAccess
    eng.WorkingDir = ASFSettings.LibraryFolder()
    eng.ImportRoot = eng.WorkingDir
    eng.MaxSteps = ASFSettings.StepLimit()
    idx = eng.Compile(source)
    If LCase$(HeaderValue(source, "args")) = "selection" Then
        If IsObject(selection) Then
            If appAccess Then Set arg = selection Else arg = ASFCells.ArgumentValue(selection)
        Else
            arg = ASFCells.ArgumentValue(selection)
        End If
        SetValue result, eng.RunArgs(idx, Array(arg))
    Else
        SetValue result, eng.RunArgs(idx, Empty)
    End If
    If Not IsEmpty(eng.LastError) Then
        RunForRibbon = "The script stopped: " & ASFEngine.ValueText(eng.LastError)
    ElseIf IsEmpty(result) Then
        RunForRibbon = "Done."
    Else
        RunForRibbon = "Result: " & ASFEngine.ValueText(result)
    End If
    Exit Function
Failed:
    RunForRibbon = "The script could not run: " & Err.Description
End Function

' The folder, made when it is missing (with its parents); False when it can't be made. Folders are
' looked up with GetAttr, which sees hidden ones too: Dir(path, vbDirectory) doesn't, and AppData,
' on the way to the default library folder on Windows, is hidden
Public Function EnsureFolder(ByVal folder As String) As Boolean
    Dim sep As String, parts As Variant, path As String, i As Long
    If FolderExists(folder) Then
        EnsureFolder = True
        Exit Function
    End If
    sep = PathSeparatorOf(folder)
    parts = Split(folder, sep)
    For i = LBound(parts) To UBound(parts)
        If i = LBound(parts) Then path = parts(i) Else path = path & sep & parts(i)
        If LenB(parts(i)) > 0 And InStr(parts(i), ":") = 0 Then
            If Not FolderExists(path) Then
                ' a level that can't be made (a share's root, no rights) shows in the check at the end
                On Error Resume Next
                MkDir path
                On Error GoTo 0
            End If
        End If
    Next i
    EnsureFolder = FolderExists(folder)
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    FolderExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
Missing:
End Function

Private Function ScriptName(ByVal path As String, ByVal fileName As String) As String
    Dim eng As New ASF
    ScriptName = HeaderValue(eng.ReadTextFile(path), "name")
    If LenB(ScriptName) = 0 Then ScriptName = Left$(fileName, Len(fileName) - 4)
End Function

Private Function PathSeparatorOf(ByVal folder As String) As String
    If InStr(folder, "/") > 0 Then PathSeparatorOf = "/" Else PathSeparatorOf = "\"
End Function

Private Sub SortByName(ByRef items As Variant)
    Dim i As Long, j As Long, t As Variant, other As Variant
    For i = LBound(items) + 1 To UBound(items)
        t = items(i)
        j = i - 1
        Do While j >= LBound(items)
            other = items(j)
            If LCase$(other(1)) <= LCase$(t(1)) Then Exit Do
            items(j + 1) = items(j)
            j = j - 1
        Loop
        items(j + 1) = t
    Next i
End Sub

Private Sub SetValue(ByRef dst As Variant, ByRef v As Variant)
    If IsObject(v) Then Set dst = v Else dst = v
End Sub
