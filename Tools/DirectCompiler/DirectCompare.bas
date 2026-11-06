Attribute VB_Name = "DirectCompare"
' DirectCompare: the stage 3 prototype (ASF_Direct, tokens -> bytecode) against the AST pipeline
' (ASF_Compiler -> ASF_Bytecode), in VBA. Import ASF_Direct.cls and this module next to the ASF
' classes (with ASF_BYTECODE = 1), put corpus.txt, bench1.vas and bench2.vas in the workbook's
' folder, and run RunDirectComparison. The results go to the Immediate window.
'
'   Parity: every program of corpus.txt must give the same disassembly on both compilers.
'   Timings: milliseconds to compile each benchmark script, tokenizer included.
Option Explicit

Private Const MIN_SECONDS As Double = 0.5

Public Sub RunDirectComparison()
    Dim eng As New ASF
    Dim folder As String: folder = ThisWorkbook.Path
    Dim progs As Variant
    progs = Split(Replace(eng.ReadTextFile(eng.JoinPath(folder, "corpus.txt")), vbCrLf, vbLf), vbLf & "----" & vbLf)
    Dim i As Long, same As Long, taken As Long, r As String
    For i = LBound(progs) To UBound(progs)
        If Len(Trim$(progs(i))) > 0 Then
            r = Parity(Trim$(Replace(progs(i), vbLf, " ")))
            If Left$(r, 9) = "NOT TAKEN" Then
                Debug.Print r & " | " & Left$(progs(i), 100)
            Else
                taken = taken + 1
                If r = "same" Then
                    same = same + 1
                Else
                    Debug.Print "DIFF: " & progs(i)
                    Debug.Print r
                End If
            End If
        End If
    Next i
    Debug.Print "parity: " & same & "/" & taken & " identical"
    Debug.Print "script | tokens -> AST -> bytecode | tokens -> bytecode (direct) | speed-up"
    Dim name As Variant, src As String, tAst As Double, tDirect As Double
    For Each name In Array("bench1.vas", "bench2.vas")
        src = eng.ReadTextFile(eng.JoinPath(folder, CStr(name)))
        tAst = TimeCompile(src, False)
        tDirect = TimeCompile(src, True)
        Debug.Print name & " | " & Format(tAst, "0.0") & " | " & Format(tDirect, "0.0") & " | " & Format(tAst / tDirect, "0.0") & "x"
    Next name
End Sub

Private Function Parity(ByVal src As String) As String
    On Error GoTo Failed
    Dim e As New ASF
    Dim idx As Long: idx = e.Compile(src)
    Dim bc As New ASF_Bytecode: bc.SetGlobals e.GetGlobals
    Dim tree As String: tree = bc.Disassemble(idx)
    Dim d As New ASF_Direct
    If Not d.Compile(src, e.GetGlobals) Then
        Parity = "NOT TAKEN: " & d.Reason
        Exit Function
    End If
    If d.Disassembly() = tree Then
        Parity = "same"
    Else
        Parity = "AST:" & vbLf & tree & vbLf & "DIRECT:" & vbLf & d.Disassembly()
    End If
    Exit Function
Failed:
    Parity = "ERROR " & Err.Number & " " & Err.Description
End Function

' Milliseconds per compilation of src from its text: a fresh engine each time
Private Function TimeCompile(ByVal src As String, ByVal direct As Boolean) As Double
    Dim n As Long, i As Long, t0 As Double, el As Double
    Dim e As ASF, bc As ASF_Bytecode, d As ASF_Direct, idx As Long
    n = 1
    Do
        el = 0
        For i = 1 To n
            Set e = New ASF
            t0 = Timer
            If direct Then
                Set d = New ASF_Direct
                d.Compile src, e.GetGlobals
            Else
                idx = e.Compile(src)
                Set bc = New ASF_Bytecode
                bc.SetGlobals e.GetGlobals
                bc.Prepare idx
            End If
            el = el + Timer - t0
        Next i
        If el >= MIN_SECONDS Or n >= 256 Then Exit Do
        n = n * 2
    Loop
    TimeCompile = el * 1000 / n
End Function
