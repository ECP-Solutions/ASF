Attribute VB_Name = "RegexBench"
' RegexBench
' Performance benchmarks for ASF_RegexEngine, for Excel (VBA) and the LibreOffice edition.
' Each workload times test, ExecAll and Replace on a fixed subject: every operation is
' repeated until it has run for at least MIN_SECONDS, and the time per call is printed in
' milliseconds to the Immediate window. Compare runs on the same machine only.
Option Explicit

Private Const MIN_SECONDS As Double = 0.5

Public Sub RunRegexBenchmarks()
    Debug.Print "=== ASF_RegexEngine benchmarks (ms per call) ==="
    Debug.Print "workload | test | ExecAll | Replace | matches"
    Bench "Literal scan", "needle", "", String(1024, "a") & "needle" & String(1024, "b"), "X"
    Bench "Short validation", "^[A-Z]{2}\d{6}$", "", "GB123456", "X"
    Bench "Capture extraction", "([A-Z]{3})-(\d{4})", "", Rep("ABC-1234, ", 100), "$2-$1"
    Bench "Alternation and ignore case", "(ERROR|WARN|INFO): ([A-Za-z]+)", "i", LogText(), "[$1] $2"
    Bench "Multiline anchors", "^([A-Z]+):(\d+)$", "m", Rep("ITEM:123" & vbLf, 100), "$2=$1"
    Bench "No-match scan", "Z\d{4}", "", String(1024, "a"), "X"
    Bench "Word boundaries", "\bcat\b", "", Rep("a cat scatter cat! ", 20), "dog"
    Bench "Backreference", "(\w)\1", "", Rep("bookkeeper committee balloon ", 20), "<$1>"
    Bench "E-mail addresses", "([\w.]+)@(\w+)\.(\w+)", "", Rep("contact: john.doe@example.com, ", 50), "$1 at $2"
    Bench "Alternation in a quantifier", "(?:foo|bar)+baz", "", Rep("foobarfoo baz foobarbaz ", 30), "#"
    Bench "Lazy body with backreference", "<(\w+)>(.*?)</\1>", "", Rep("<b>bold</b> <i>italic</i> ", 30), "[$2]"
    Bench "Lookaheads", "^(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{8,}$", "m", Rep("abcDEF12" & vbLf & "weakpass" & vbLf & "NoDigitsHere" & vbLf, 30), "ok"
    Bench "Quantified group", "(\d+,)*\d+", "", Rep("1,22,333,4444 ", 30), "#"
    Bench "Backtracking into a group", "((\w+)\s?)+;", "", Rep("one two three four five six; ", 10), "#"
    Debug.Print "=== Benchmarks done ==="
End Sub

Private Sub Bench(ByVal label As String, ByVal pattern As String, ByVal flags As String, ByVal subj As String, ByVal repl As String)
    Dim re As ASF_RegexEngine
    Set re = New ASF_RegexEngine
    re.Init pattern, InStr(flags, "i") > 0, 2000000000, InStr(flags, "m") > 0
    Dim all As Collection
    Set all = re.ExecAll(subj)
    Dim matches As Long
    If Not all Is Nothing Then matches = all.Count
    Debug.Print label & " | " & Format$(TimeOp(re, 1, subj, repl), "0.000") & " | " & _
                Format$(TimeOp(re, 2, subj, repl), "0.000") & " | " & _
                Format$(TimeOp(re, 3, subj, repl), "0.000") & " | " & matches
End Sub

' Milliseconds per call of op (1 test, 2 ExecAll, 3 Replace)
Private Function TimeOp(ByVal re As ASF_RegexEngine, ByVal op As Long, ByVal subj As String, ByVal repl As String) As Double
    Dim n As Long, i As Long
    Dim t0 As Double, el As Double
    Dim r As Variant
    n = 1
    Do
        t0 = Timer
        For i = 1 To n
            Select Case op
                Case 1: r = re.test(subj)
                Case 2: Set r = re.ExecAll(subj)
                Case 3: r = re.Replace(subj, repl)
            End Select
        Next i
        el = Timer - t0
        If el < 0 Then el = el + 86400
        If el >= MIN_SECONDS Or n >= 65536 Then Exit Do
        n = n * 4
    Loop
    TimeOp = el * 1000 / n
End Function

Private Function Rep(ByVal s As String, ByVal n As Long) As String
    Dim k As Long, out As String
    For k = 1 To n
        out = out & s
    Next k
    Rep = out
End Function

Private Function LogText() As String
    Dim words As Variant
    words = Array("ERROR: Disk", "warn: Fanspeed", "Info: Started", "error: Timeout", "WARN: Lowmem")
    Dim k As Long, out As String
    For k = 0 To 99
        out = out & words(k Mod 5) & vbLf
    Next k
    LogText = out
End Function
