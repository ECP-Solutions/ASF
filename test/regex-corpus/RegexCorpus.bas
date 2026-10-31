Attribute VB_Name = "RegexCorpus"
' RegexCorpus
' Runs the regex corpora of test/regex-corpus against ASF_RegexEngine and compares the output
' with the expected files, in Excel (VBA) or the LibreOffice edition. See README.md there.
'
'   RunCorpusSuite "C:\ASF\test\regex-corpus"          ' c1..c3 against JavaScript's output
'   RunCorpusJs  caseFile, outFile                     ' one corpus, JavaScript-comparable output
'   RunCorpusFull caseFile, outFile                    ' one corpus, every public method
'   CompareCorpusOutput actualFile, expectedFile       ' differing lines, printed and counted
'
' Case files hold one case per line: pattern, subject and flags (i, m, s) separated by tabs,
' pattern and subject as UTF-16 code units in hex (4 digits each), so any char survives the
' file round trip. Output lines use the same encoding.
Option Explicit

Private Const MAX_STEPS As Long = 200000

' Runs c1, c2 and c3 in JavaScript mode and compares them with expected/js; True when all match
Public Function RunCorpusSuite(ByVal folder As String) As Boolean
    Dim sep As String: sep = PathSep(folder)
    Dim outDir As String: outDir = folder & sep & "out"
    If Len(Dir(outDir, vbDirectory)) = 0 Then MkDir outDir
    Dim c As Variant, bad As Long, total As Long
    Debug.Print "=== ASF_RegexEngine corpus suite (JavaScript mode) ==="
    For Each c In Array("c1", "c2", "c3")
        RunCorpusJs folder & sep & "cases" & sep & c & ".txt", outDir & sep & c & ".js.txt"
        bad = CompareCorpusOutput(outDir & sep & c & ".js.txt", folder & sep & "expected" & sep & "js" & sep & c & ".js.txt")
        total = total + bad
    Next c
    Debug.Print "=== Corpus suite done: " & total & " differing cases ==="
    RunCorpusSuite = (total = 0)
End Function

' Exec captures, ExecAll full matches and two global replacements per case, in the format of
' tools/js_run.js
Public Sub RunCorpusJs(ByVal caseFile As String, ByVal outFile As String)
    Dim fi As Integer, fo As Integer, ln As String
    Dim p As String, s As String, fl As String, res As String
    Dim r As Object, c As Object, all As Object, k As Long, allStr As String
    fi = FreeFile
    Open caseFile For Input As #fi
    fo = FreeFile
    Open outFile For Output As #fo
    Do While Not EOF(fi)
        Line Input #fi, ln
        If ReadCase(ln, p, s, fl) Then
            Set r = New ASF_RegexEngine
            On Error GoTo InitErr
            r.Init p, InStr(fl, "i") > 0, MAX_STEPS, InStr(fl, "m") > 0, InStr(fl, "s") > 0
            On Error GoTo ExecErr
            Set c = r.Exec(s)
            If c Is Nothing Then
                res = "NOMATCH"
            Else
                res = "M "
                For k = 1 To c.Count
                    If k > 1 Then res = res & ","
                    res = res & Enc(CStr(c(k)))
                Next k
            End If
            Set all = r.ExecAll(s)
            If all Is Nothing Then
                allStr = "NONE"
            Else
                allStr = vbNullString
                For k = 1 To all.Count
                    If k > 1 Then allStr = allStr & ","
                    allStr = allStr & Enc(CStr(all(k)(1)))
                Next k
            End If
            res = res & vbTab & allStr & vbTab & Enc(r.Replace(s, "<$&>")) & vbTab & Enc(r.Replace(s, "<$&|$1>"))
            On Error GoTo 0
            Print #fo, res
        End If
        GoTo NextCase
InitErr:
        Print #fo, "ERR"
        Resume NextCase
ExecErr:
        Print #fo, "XERR " & Err.Number & " " & Err.Description
        Resume NextCase
NextCase:
    Loop
    Close #fi
    Close #fo
End Sub

' Every public method per case, three times: as compiled, with ignoreCase and multiline flipped
' after Init, and after re-Init from the pattern cache
Public Sub RunCorpusFull(ByVal caseFile As String, ByVal outFile As String)
    Dim fi As Integer, fo As Integer, ln As String
    Dim p As String, s As String, fl As String, res As String
    Dim r As Object
    fi = FreeFile
    Open caseFile For Input As #fi
    fo = FreeFile
    Open outFile For Output As #fo
    Do While Not EOF(fi)
        Line Input #fi, ln
        If ReadCase(ln, p, s, fl) Then
            Set r = New ASF_RegexEngine
            On Error GoTo InitErr
            r.Init p, InStr(fl, "i") > 0, MAX_STEPS, InStr(fl, "m") > 0, InStr(fl, "s") > 0
            On Error GoTo 0
            res = AllOps(r, s)
            r.ignoreCase = Not r.ignoreCase
            r.multiline = Not r.multiline
            res = res & " |F " & AllOps(r, s)
            On Error GoTo InitErr
            r.Init "x(y)", False, MAX_STEPS
            r.Init p, InStr(fl, "i") > 0, MAX_STEPS, InStr(fl, "m") > 0, InStr(fl, "s") > 0
            On Error GoTo 0
            res = res & " |C " & AllOps(r, s)
            Print #fo, res
        End If
        GoTo NextCase
InitErr:
        Print #fo, res & " ERR " & Err.Number & " " & Err.Description
        Resume NextCase
NextCase:
        res = vbNullString
    Loop
    Close #fi
    Close #fo
End Sub

' Prints the first maxShown differing lines and returns how many lines differ
Public Function CompareCorpusOutput(ByVal actualFile As String, ByVal expectedFile As String, Optional ByVal maxShown As Long = 10) As Long
    Dim fa As Integer, fe As Integer, la As String, le As String, n As Long, bad As Long
    fa = FreeFile
    Open actualFile For Input As #fa
    fe = FreeFile
    Open expectedFile For Input As #fe
    Do While Not EOF(fe)
        Line Input #fe, le
        n = n + 1
        If EOF(fa) Then
            la = "<missing>"
        Else
            Line Input #fa, la
        End If
        If la <> le Then
            bad = bad + 1
            If bad <= maxShown Then
                Debug.Print "  line " & n & " expected: " & Left$(le, 160)
                Debug.Print "  line " & n & " actual:   " & Left$(la, 160)
            End If
        End If
    Loop
    Do While Not EOF(fa)
        Line Input #fa, la
        n = n + 1
        bad = bad + 1
    Loop
    Close #fa
    Close #fe
    Debug.Print Mid$(actualFile, InStrRev(Replace(actualFile, "\", "/"), "/") + 1) & ": " & n & " cases, " & bad & " differ"
    CompareCorpusOutput = bad
End Function

Private Function AllOps(ByVal r As Object, ByVal s As String) As String
    Dim o As String, t As Variant
    On Error GoTo OpErr
    o = "E" & CapsLine(r.Exec(s)) & NamedLine(r)
    o = o & " T" & CStr(r.test(s))
    o = o & " A" & AllLine(r.ExecAll(s)) & NamedLine(r)
    For Each t In Array("<$&>", "<$&|$1>", "[$`|$']", "$2$1$$", "$<n>$01$10x$9$<", "plain", "$0$&$")
        o = o & " R" & Enc(r.Replace(s, CStr(t)))
    Next t
    o = o & " R2" & Enc(r.Replace(s, "<$&|$1>")) & " R3" & Enc(r.Replace(s, "<$&|$1>"))
    o = o & " S" & CapsLine(r.Split(s))
    For Each t In Array(1, 2, Len(s), Len(s) + 1)
        o = o & " @" & CapsLine(r.ExecAt(s, CLng(t)))
    Next t
    o = o & NamedLine(r)
    AllOps = o
    Exit Function
OpErr:
    AllOps = o & " XERR " & Err.Number & " " & Err.Description
End Function

Private Function ReadCase(ByVal ln As String, ByRef p As String, ByRef s As String, ByRef fl As String) As Boolean
    If Len(ln) = 0 Then Exit Function
    Dim parts As Variant: parts = Split(ln, vbTab)
    p = Dec(CStr(parts(0)))
    s = vbNullString
    fl = vbNullString
    If UBound(parts) >= 1 Then s = Dec(CStr(parts(1)))
    If UBound(parts) >= 2 Then fl = parts(2)
    ReadCase = True
End Function

Private Function CapsLine(ByVal c As Object) As String
    Dim k As Long, o As String
    If c Is Nothing Then CapsLine = "-": Exit Function
    For k = 1 To c.Count
        If k > 1 Then o = o & ","
        o = o & Enc(CStr(c(k)))
    Next k
    CapsLine = "[" & o & "]"
End Function

Private Function AllLine(ByVal all As Object) As String
    Dim k As Long, o As String
    If all Is Nothing Then AllLine = "-": Exit Function
    For k = 1 To all.Count
        o = o & CapsLine(all(k))
    Next k
    AllLine = o
End Function

Private Function NamedLine(ByVal r As Object) As String
    Dim nc As Object, k As Long, o As String
    Set nc = r.LastNamedCaptures
    If nc Is Nothing Then NamedLine = "-": Exit Function
    For k = 1 To nc.Count
        o = o & Enc(CStr(nc(k)(0))) & "=" & Enc(CStr(nc(k)(1))) & ";"
    Next k
    NamedLine = "{" & o & "}"
End Function

Private Function Dec(ByVal h As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(h) Step 4
        o = o & ChrW(Val("&H" & Mid$(h, i, 4)))
    Next i
    Dec = o
End Function

Private Function Enc(ByVal s As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(s)
        o = o & Right$("000" & LCase$(Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&)), 4)
    Next i
    Enc = o
End Function

Private Function PathSep(ByVal folder As String) As String
    If InStr(folder, "\") > 0 Then PathSep = "\" Else PathSep = "/"
End Function
