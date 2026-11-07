Attribute VB_Name = "DirectCompare"
' DirectCompare: the direct compiler (ASF_Bytecode.CompileSource, tokens -> bytecode) against the AST
' path (ASF_Compiler -> ASF_Bytecode.Prepare), in VBA. Import this module next to the ASF classes
' (with ASF_BYTECODE = 1), put corpus.txt, bench1.vas and bench2.vas in the workbook's folder, and
' run RunDirectComparison. The results go to the Immediate window.
'
'   Parity: every program of corpus.txt must give the same disassembly on both paths.
'   Timings: milliseconds to compile each benchmark script, tokenizer included.
Option Explicit

Private Const MIN_SECONDS As Double = 0.5

Public Sub RunDirectComparison()
    Dim eng As New ASF
    Dim folder As String: folder = ThisWorkbook.Path
    Dim progs As Variant
    progs = Split(Replace(eng.ReadTextFile(eng.JoinPath(folder, "corpus.txt")), vbCrLf, vbLf), vbLf & "----" & vbLf)
    Dim i As Long, same As Long, taken As Long, errors As Long, r As String
    For i = LBound(progs) To UBound(progs)
        If Len(Trim$(progs(i))) > 0 Then
            r = Parity(Trim$(Replace(progs(i), vbLf, " ")))
            If Left$(r, 9) = "NOT TAKEN" Then
                Debug.Print r & " | " & Left$(progs(i), 100)
            ElseIf Left$(r, 9) = "ERROR AST" Then
                errors = errors + 1
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
    Debug.Print "parity: " & same & "/" & taken & " identical, " & errors & " rejected by ASF_Compiler"
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
    Dim phase As String: phase = "AST"
    Dim g As New ASF_Globals: g.ASF_InitGlobals
    Dim c As New ASF_Compiler: c.SetGlobals g
    Dim idx As Long: idx = c.CompileProgram(src)
    Dim bc As New ASF_Bytecode: bc.SetGlobals g
    Dim tree As String: tree = bc.Disassemble(idx)
    phase = "direct"
    Dim g2 As New ASF_Globals: g2.ASF_InitGlobals
    Dim d As New ASF_Bytecode: d.SetGlobals g2
    Dim idx2 As Long: idx2 = d.CompileSource(src, "@anon")
    If idx2 = 0 Then
        Parity = "NOT TAKEN: " & d.Reason
        Exit Function
    End If
    Dim direct As String: direct = d.Disassemble(idx2)
    If direct = tree Then
        Parity = "same"
    Else
        Parity = "AST:" & vbLf & tree & vbLf & "DIRECT:" & vbLf & direct
    End If
    Exit Function
Failed:
    Parity = "ERROR " & phase & " " & Err.Number & " " & Err.Description
End Function

' Milliseconds per compilation of src from its text: fresh globals each time
Private Function TimeCompile(ByVal src As String, ByVal direct As Boolean) As Double
    Dim n As Long, i As Long, t0 As Double, el As Double
    Dim g As ASF_Globals, c As ASF_Compiler, bc As ASF_Bytecode, idx As Long
    n = 1
    Do
        el = 0
        For i = 1 To n
            Set g = New ASF_Globals
            g.ASF_InitGlobals
            Set bc = New ASF_Bytecode
            bc.SetGlobals g
            t0 = Timer
            If direct Then
                bc.CompileSource src, "@anon"
            Else
                Set c = New ASF_Compiler
                c.SetGlobals g
                idx = c.CompileProgram(src)
                bc.Prepare idx
            End If
            el = el + Timer - t0
        Next i
        If el >= MIN_SECONDS Or n >= 256 Then Exit Do
        n = n * 2
    Loop
    TimeCompile = el * 1000 / n
End Function
