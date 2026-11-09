Attribute VB_Name = "VersionBench"
' VersionBench
' ASF 3.2.1 (src/old, the classes with the suffix _OLD, which run the AST) against ASF 4 (bytecode),
' one workload per aspect of the language. For each workload: the run time of a compiled program
' on each version, the cold time (a fresh engine, Compile and the first Run), and the result of each
' against the expected one. Times are milliseconds per call; each is repeated until it has taken
' MIN_SECONDS. Import the classes of src/ and of src/old/ to run it.
'
'   RunVersionBenchmarks
Option Explicit

Private Const MIN_SECONDS As Double = 0.3

Public Sub RunVersionBenchmarks()
    Debug.Print "=== ASF 3.2.1 against ASF 4 (ms per call) ==="
    Debug.Print "workload | 3.2.1 run | 4.0 run | speed-up | 3.2.1 cold | 4.0 cold | result | check"
    ' expressions and control flow
    Bench "Arithmetic, top level", "let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + i * 2 - 1; }; return s;", "90000"
    Bench "Arithmetic, in a function", "fun f() { let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + i * 2 - 1; }; return s; }; return f();", "90000"
    Bench "Comparisons and branches", "let n = 0; for (let i = 1, i <= 300, i += 1) { if (i % 3 == 0) { n += 2; } elseif (i % 3 == 1) { n -= 1; } else { n += 0; }; }; return n;", "100"
    Bench "While loop", "let i = 0; let n = 0; while (i < 300) { i += 1; n = n + i; }; return n;", "45150"
    Bench "Logical operators", "let n = 0; for (let i = 1, i <= 300, i += 1) { if ((i > 10 && i < 200) || i == 250) { n += 1; }; }; return n;", "190"
    Bench "Ternary", "let n = 0; for (let i = 1, i <= 300, i += 1) { n = n + (i % 2 == 0 ? 1 : -1); }; return n;", "0"
    Bench "Constant expressions", "let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + (2 * 3 + 4) / 5; }; return s;", "600"
    Bench "String concatenation", "let s = ''; for (let i = 1, i <= 300, i += 1) { s = s + 'x' + i; }; return s.length;", "1092"
    Bench "Ampersand operator", "let s = ''; for (let i = 1, i <= 300, i += 1) { s = s & i; }; return s.length;", "792"
    Bench "typeof", "let n = 0; for (let i = 1, i <= 300, i += 1) { if (typeof i == 'number') { n += 1; }; }; return n;", "300"
    Bench "Null coalescing", "let u; let n = 0; for (let i = 1, i <= 300, i += 1) { n = n + (u ?? 1); }; return n;", "300"
    ' functions
    Bench "Function calls", "fun add(a, b) { return a + b; }; let s = 0; for (let i = 1, i <= 300, i += 1) { s = add(s, i); }; return s;", "45150"
    Bench "Recursion (fib 14)", "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return fib(14);", "377"
    Bench "Deep recursion (300)", "fun down(n) { if (n == 0) { return 0; }; return 1 + down(n - 1); }; return down(300);", "300"
    Bench "Globals from a function", "let total = 0; fun bump(k) { total = total + k; }; for (let i = 1, i <= 300, i += 1) { bump(i); }; return total;", "45150"
    Bench "Error in a function", "fun bad() { return 1 / 0; }; let n = 0; for (let i = 1, i <= 100, i += 1) { if (bad() == undefined) { n += 1; }; }; return n;", "100"
    ' closures, arrays and objects
    Bench "Closure calls", "fun counter() { let c = 0; return fun() { c = c + 1; return c; }; }; let f = counter(); let r = 0; for (let i = 1, i <= 300, i += 1) { r = f(); }; return r;", "300"
    Bench "Closure creation", "let n = 0; for (let i = 1, i <= 300, i += 1) { let g = fun(x) { return x + i; }; n = g(n) - n; }; return n;", "300"
    Bench "Array literals", "let n = 0; for (let i = 1, i <= 300, i += 1) { let a = [i, i + 1, i + 2]; n = n + a.length; }; return n;", "900"
    Bench "Array index read", "let a = [1, 2, 3, 4, 5, 6]; let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + a[1 + i % 6]; }; return s;", "1050"
    Bench "Array index write", "let a = [0, 0, 0, 0, 0, 0]; for (let i = 1, i <= 300, i += 1) { a[1 + i % 6] = a[1 + i % 6] + i; }; return a[1] + a[2] + a[3] + a[4] + a[5] + a[6];", "45150"
    Bench "Nested arrays", "let m = [[0, 0, 0], [0, 0, 0], [0, 0, 0]]; for (let i = 1, i <= 300, i += 1) { m[1 + i % 3][1 + i % 2] = m[1 + i % 3][1 + i % 2] + i; }; return m[1][1] + m[2][2] + m[3][1];", "22500"
    Bench "Object literals", "let n = 0; for (let i = 1, i <= 300, i += 1) { let o = { a: i, b: 2 }; n = n + o.b; }; return n;", "600"
    Bench "Property reads", "let o = { a: 1, b: 2, c: { d: 3 } }; let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + o.a + o.b + o.c.d; }; return s;", "1800"
    Bench "Property writes", "let o = { a: 0, b: 1 }; for (let i = 1, i <= 300, i += 1) { o.a = o.a + o.b; }; return o.a;", "300"
    ' built-in methods and the rest of the language
    Bench "Array push", "let a = []; for (let i = 1, i <= 300, i += 1) { a.push(i); }; return a.length;", "300"
    Bench "map / filter / reduce", "let a = []; for (let i = 1, i <= 100, i += 1) { a.push(i); }; return a.map(fun(x) { return x * 2; }).filter(fun(x) { return x % 3 == 0; }).reduce(fun(acc, v) { return acc + v; }, 0);", "3366"
    Bench "Sort with a comparator", "let a = []; for (let i = 1, i <= 60, i += 1) { a.push((i * 37) % 61); }; let b = a.sort(fun(x, y) { return x - y; }); return b[1] + b[60];", "61"
    Bench "String methods", "let n = 0; for (let i = 1, i <= 300, i += 1) { n = n + ('Item ' + i).toUpperCase().trim().length; }; return n;", "2292"
    Bench "Template literals", "let n = 0; for (let i = 1, i <= 300, i += 1) { let s = `item ${i} of ${300}`; n = n + s.length; }; return n;", "4392"
    Bench "Math methods", "let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + Math.max(i, 150) - Math.floor(i / 2); }; return s;", "33825"
    Bench "switch", "let n = 0; for (let i = 1, i <= 300, i += 1) { switch (i % 3) { case 0 { n += 3; } case 1 { n += 1; } default { n += 2; } }; }; return n;", "600"
    Bench "for-of", "let a = []; for (let i = 1, i <= 300, i += 1) { a.push(i); }; let s = 0; for (let x of a) { s = s + x; }; return s;", "45150"
    Bench "for-in over an object", "let o = { a: 1, b: 2, c: 3, d: 4 }; let n = 0; for (let i = 1, i <= 75, i += 1) { for (let k in o) { n = n + k.length; }; }; return n;", "300"
    Bench "Spread and rest", "fun sum(...xs) { return xs.reduce(fun(a, b) { return a + b; }, 0); }; let n = 0; for (let i = 1, i <= 100, i += 1) { n = n + sum(...[i, 1, 2]); }; return n;", "5350"
    Bench "Destructuring", "let n = 0; for (let i = 1, i <= 300, i += 1) { let [a, b] = [i, 2]; n = n + a * b; }; return n;", "90300"
    Bench "Regex", "let re = regex(`[0-9]+`); let n = 0; for (let i = 1, i <= 100, i += 1) { if (re.test('abc' + i)) { n += 1; }; }; return n;", "100"
    Bench "Classes", "class P { constructor(x) { this.x = x; } add(d) { this.x = this.x + d; return this.x; } }; let p = new P(0); for (let i = 1, i <= 300, i += 1) { p.add(i); }; return p.x;", "45150"
    Bench "Inheritance and super", "class A { constructor(x) { this.x = x; } v() { return this.x; } }; class B extends A { constructor(x) { super(x); } w() { return this.v() + 1; } }; let n = 0; for (let i = 1, i <= 100, i += 1) { let b = new B(i); n = n + b.w(); }; return n;", "5150"
    Bench "Static methods", "class M { static sq(x) { return x * x; } }; let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + M.sq(i % 10); }; return s;", "8550"
    Bench "Bound method as a value", "class P { field k = 2; m(x) { return x * this.k; } }; let f = new P().m; let s = 0; for (let i = 1, i <= 300, i += 1) { s = s + f(i); }; return s;", "90300"
    ' modules
    Bench "Module function calls", "scwd('" & ModuleFolder() & "'); import { add } from './bench_module.vas'; let s = 0; for (let i = 1, i <= 300, i += 1) { s = add(s, i); }; return s;", "45150"
    RemoveModule
    Bench "try without throw", "let n = 0; for (let i = 1, i <= 300, i += 1) { try { n += 1; } catch (e) { n -= 1; }; }; return n;", "300"
    Bench "throw and catch", "let n = 0; for (let i = 1, i <= 100, i += 1) { try { throw i; } catch (e) { n = n + e; }; }; return n;", "5050"
    ' start-up and compilation
    Bench "Start-up (return 1)", "return 1;", "1"
    Bench "Long script", LongScript(), "300"
    Debug.Print "=== done ==="
End Sub

' A folder with the module the module workload imports
Private Function ModuleFolder() As String
    Dim sep As String: sep = "/"
    If InStr(CurDir$, "\") > 0 Then sep = "\"
    Dim f As Integer: f = FreeFile
    Open CurDir$ & sep & "bench_module.vas" For Output As #f
    Print #f, "fun add(a, b) { return a + b; }; export { add };"
    Close #f
    ModuleFolder = CurDir$
End Function

Private Sub RemoveModule()
    On Error Resume Next
    Kill CurDir$ & IIf(InStr(CurDir$, "\") > 0, "\", "/") & "bench_module.vas"
End Sub

' 300 statements, to weigh the compilers
Private Function LongScript() As String
    Dim s As String, i As Long
    s = "let n = 0; "
    For i = 1 To 300
        s = s & "n = n + 1; "
    Next i
    LongScript = s & "return n;"
End Function

Private Sub Bench(ByVal label As String, ByVal script As String, ByVal expected As String)
    Dim oldEng As ASF_OLD, newEng As ASF, oldIdx As Long, newIdx As Long
    Set oldEng = New ASF_OLD
    oldIdx = oldEng.Compile(script)
    Set newEng = New ASF
    newIdx = newEng.Compile(script)
    Dim oldResult As String, newResult As String
    oldResult = Shown(oldEng.Run(oldIdx))
    newResult = Shown(newEng.Run(newIdx))
    Dim oldRun As Double, newRun As Double
    oldRun = TimeRun(oldEng, oldIdx)
    newRun = TimeRun(newEng, newIdx)
    Dim check As String
    If newResult <> expected Then
        check = "4.0 gives " & newResult & ", expected " & expected
    ElseIf oldResult <> newResult Then
        check = "3.2.1 gives " & oldResult
    Else
        check = "ok"
    End If
    Debug.Print label & " | " & Format$(oldRun, "0.000") & " | " & Format$(newRun, "0.000") & " | " & _
                Format$(oldRun / IIf(newRun > 0, newRun, 0.001), "0.0") & "x | " & _
                Format$(TimeCold(script, False), "0.000") & " | " & Format$(TimeCold(script, True), "0.000") & " | " & _
                newResult & " | " & check
End Sub

Private Function Shown(ByVal v As Variant) As String
    If IsObject(v) Then
        Shown = "<object>"
    ElseIf IsArray(v) Then
        Shown = "<array>"
    Else
        Shown = CStr(v)
    End If
End Function

' Milliseconds per Run of a compiled program
Private Function TimeRun(ByVal eng As Object, ByVal idx As Long) As Double
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

' Milliseconds per fresh engine, Compile and first Run; isNew: ASF 4, else ASF 3.2.1
Private Function TimeCold(ByVal script As String, ByVal isNew As Boolean) As Double
    Dim n As Long, i As Long, t0 As Double, el As Double
    Dim eng As Object, r As Variant
    n = 1
    Do
        t0 = Timer
        For i = 1 To n
            If isNew Then Set eng = New ASF Else Set eng = New ASF_OLD
            r = eng.Run(eng.Compile(script))
        Next i
        el = Timer - t0
        If el < 0 Then el = el + 86400
        If el >= MIN_SECONDS Or n >= 4096 Then Exit Do
        n = n * 2
    Loop
    TimeCold = el * 1000 / n
End Function
