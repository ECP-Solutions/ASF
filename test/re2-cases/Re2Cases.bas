Attribute VB_Name = "Re2Cases"
' Re2Cases
' Runs RE2's test cases, as prepared by tools/prepare.py from sihlfall/regex-test-cases, against
' ASF_RegexEngine in Excel (VBA) or the LibreOffice edition. See README.md there.
'
'   RunRe2Suite "C:\ASF\test\re2-cases"      ' every suite in work\cases, output to out\
'   RunRe2Cases caseFile, outFile              ' one suite
'
' A case file lists the subjects first, one "S <hex>" line each, then one "R <hex>" line per
' pattern; text is in UTF-16 code units, 4 hex digits each, and anything after a tab is ignored.
' For every pattern and subject the output has one line with two fields separated by a tab: the
' leftmost match, and the match of ^(?:pattern)$ (RE2's anchor at both ends). A match is written
' "start,end" in code units from 0, followed by "|" and the text of each group in hex; "-" is no
' match, "E<number>" a rejected pattern and "X<number>" an error while matching.
Option Explicit

Private Const MAX_STEPS As Long = 1000000

Public Sub RunRe2Suite(ByVal folder As String)
    Dim sep As String: If InStr(folder, "\") > 0 Then sep = "\" Else sep = "/"
    Dim outDir As String: outDir = folder & sep & "out"
    If Len(Dir(outDir, vbDirectory)) = 0 Then MkDir outDir
    Dim names As Collection: Set names = New Collection
    Dim f As String: f = Dir(folder & sep & "work" & sep & "cases" & sep & "*.txt")
    Do While Len(f) > 0
        names.Add f
        f = Dir()
    Loop
    Dim k As Long, t0 As Double
    For k = 1 To names.Count
        t0 = Timer
        RunRe2Cases folder & sep & "work" & sep & "cases" & sep & names(k), outDir & sep & names(k)
        Debug.Print names(k) & ": " & Format$(Timer - t0, "0") & " s"
    Next k
End Sub

Public Sub RunRe2Cases(ByVal caseFile As String, ByVal outFile As String)
    Dim fi As Integer, fo As Integer, ln As String, fld As String
    Dim subjects As Collection: Set subjects = New Collection
    ' two engines for the whole file: in LibreOffice an engine that compiled is never freed
    Dim rLeft As Object, rWhole As Object
    Set rLeft = New ASF_RegexEngine
    Set rWhole = New ASF_RegexEngine
    Dim leftErr As String, wholeErr As String
    Dim p As String, k As Long
    fi = FreeFile
    Open caseFile For Input As #fi
    fo = FreeFile
    Open outFile For Output As #fo
    Do While Not EOF(fi)
        Line Input #fi, ln
        fld = ln
        If InStr(fld, vbTab) > 0 Then fld = Left$(fld, InStr(fld, vbTab) - 1)
        If Left$(fld, 2) = "S " Then
            subjects.Add Dec(Mid$(fld, 3))
        ElseIf Left$(fld, 2) = "R " Then
            p = Dec(Mid$(fld, 3))
            leftErr = InitEngine(rLeft, p)
            wholeErr = InitEngine(rWhole, "^(?:" & p & ")$")
            For k = 1 To subjects.Count
                Print #fo, Leftmost(rLeft, leftErr, subjects(k)) & vbTab & WholeMatch(rWhole, wholeErr, subjects(k))
            Next k
        End If
    Loop
    Close #fi
    Close #fo
End Sub

' "" when the pattern compiles, else "E<number>"
Private Function InitEngine(ByVal r As Object, ByVal pattern As String) As String
    InitEngine = vbNullString
    On Error GoTo InitErr
    r.Init pattern, False, MAX_STEPS
    Exit Function
InitErr:
    InitEngine = "E" & Err.Number
End Function

' The first start position with a match, as Exec finds it, with the match's span
Private Function Leftmost(ByVal r As Object, ByVal errText As String, ByVal s As String) As String
    If Len(errText) > 0 Then Leftmost = errText: Exit Function
    Dim c As Object, p As Long
    On Error GoTo MatchErr
    For p = 1 To Len(s) + 1
        Set c = r.ExecAt(s, p)
        If Not c Is Nothing Then
            Leftmost = MatchLine(c, p)
            Exit Function
        End If
    Next p
    Leftmost = "-"
    Exit Function
MatchErr:
    Leftmost = "X" & Err.Number
End Function

Private Function WholeMatch(ByVal r As Object, ByVal errText As String, ByVal s As String) As String
    If Len(errText) > 0 Then WholeMatch = errText: Exit Function
    Dim c As Object
    On Error GoTo MatchErr
    Set c = r.ExecAt(s, 1)
    If c Is Nothing Then WholeMatch = "-" Else WholeMatch = MatchLine(c, 1)
    Exit Function
MatchErr:
    WholeMatch = "X" & Err.Number
End Function

Private Function MatchLine(ByVal c As Object, ByVal startPos As Long) As String
    Dim o As String, k As Long
    o = (startPos - 1) & "," & (startPos - 1 + Len(c(1)))
    For k = 2 To c.Count
        o = o & "|" & Enc(CStr(c(k)))
    Next k
    MatchLine = o
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
