Attribute VB_Name = "AsfBench"
' AsfBench
' Performance benchmarks for ASF's compiler and VM, for Excel (VBA) and the LibreOffice edition.
' Each workload is an ASF script: its compile time is a fresh engine plus Compile, its run time
' is Run on the compiled program. Both are repeated until they have taken at least MIN_SECONDS,
' and the time per call is printed in milliseconds to the Immediate window together with the
' script's result, which must match the expected value. Compare runs on the same machine only.
'
'   RunAsfBenchmarks
Option Explicit

Private Const MIN_SECONDS As Double = 0.5

Public Sub RunAsfBenchmarks()
    Debug.Print "=== ASF benchmarks (ms per call) ==="
    Debug.Print "workload | compile | run | result | expected"
    Bench "Arithmetic loop", _
          "let s = 0; for (let i = 1, i <= 2000, i += 1) { s = s + i * 2 - 1; }; return(s);", "4000000"
    Bench "Recursive calls", _
          "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return(fib(14));", "377"
    Bench "String building", _
          "let s = ''; for (let i = 1, i <= 500, i += 1) { s = s + 'x' + i; }; return(s.length);", "1892"
    Bench "Array methods", _
          "let a = []; for (let i = 1, i <= 300, i += 1) { a.push(i); }; " & _
          "let t = a.map(fun(x) { return x * 2; }).filter(fun(x) { return x % 3 == 0; }).reduce(fun(acc, v) { return acc + v; }, 0); return(t);", "30300"
    Bench "Object properties", _
          "let o = { a: 0, b: 1, c: { d: 2 } }; for (let i = 1, i <= 1000, i += 1) { o.a = o.a + o.b + o.c.d; }; return(o.a);", "3000"
    Bench "Closures", _
          "fun counter() { let c = 0; return fun() { c = c + 1; return c; }; }; let f = counter(); let r = 0; " & _
          "for (let i = 1, i <= 500, i += 1) { r = f(); }; return(r);", "500"
    Bench "While and branches", _
          "let i = 0; let n = 0; while (i < 1000) { if (i % 3 == 0) { n = n + 2; } else { n = n - 1; }; i += 1; }; return(n);", "2"
    Bench "Template literals", _
          "let out = 0; for (let i = 1, i <= 300, i += 1) { let s = `item ${i}`; out = out + s.toUpperCase().length; }; return(out);", "2292"
    Bench "Array indexing", _
          "let a = [0, 0, 0, 0, 0, 0]; for (let i = 1, i <= 300, i += 1) { a[1 + i % 6] = a[1 + i % 6] + i; }; " & _
          "return(a[1] + a[2] + a[3] + a[4] + a[5] + a[6]);", "45150"
    Bench "Nested arrays", _
          "let m = [[0, 0, 0], [0, 0, 0], [0, 0, 0]]; for (let i = 1, i <= 300, i += 1) { m[1 + i % 3][1 + i % 2] = m[1 + i % 3][1 + i % 2] + i; }; " & _
          "return(m[1][1] + m[2][2] + m[3][1]);", "22500"
    Bench "Classes", _
          "class P { constructor(x) { this.x = x; } add(d) { this.x = this.x + d; return this.x; } }; " & _
          "let p = new P(0); for (let i = 1, i <= 300, i += 1) { p.add(i); }; return(p.x);", "45150"
    Debug.Print "=== ASF benchmarks done ==="
End Sub

Private Sub Bench(ByVal label As String, ByVal script As String, ByVal expected As String)
    Dim eng As ASF, idx As Long
    Set eng = New ASF
    idx = eng.Compile(script)
    Dim result As String: result = CStr(eng.Run(idx))
    Debug.Print label & " | " & Format$(TimeCompile(script), "0.000") & " | " & _
                Format$(TimeRun(eng, idx), "0.000") & " | " & result & " | " & IIf(result = expected, "ok", "expected " & expected)
End Sub

' Milliseconds per fresh engine and Compile
Private Function TimeCompile(ByVal script As String) As Double
    Dim n As Long, i As Long, t0 As Double, el As Double
    Dim eng As ASF, idx As Long
    n = 1
    Do
        t0 = Timer
        For i = 1 To n
            Set eng = New ASF
            idx = eng.Compile(script)
        Next i
        el = Timer - t0
        If el < 0 Then el = el + 86400
        If el >= MIN_SECONDS Or n >= 4096 Then Exit Do
        n = n * 2
    Loop
    TimeCompile = el * 1000 / n
End Function

' Milliseconds per Run of a compiled program
Private Function TimeRun(ByVal eng As ASF, ByVal idx As Long) As Double
    Dim n As Long, i As Long, t0 As Double, el As Double
    Dim r As Variant
    n = 1
    Do
        t0 = Timer
        For i = 1 To n
            r = eng.Run(idx)
        Next i
        el = Timer - t0
        If el < 0 Then el = el + 86400
        If el >= MIN_SECONDS Or n >= 4096 Then Exit Do
        n = n * 2
    Loop
    TimeRun = el * 1000 / n
End Function
