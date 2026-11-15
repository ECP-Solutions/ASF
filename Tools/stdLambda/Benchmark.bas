Attribute VB_Name = "LambdaBenchmark"
' stdLambda (stdVBA) against ASF on the same workloads, one per part of an engine: arithmetic,
' built-in functions, branches, loops, calls, recursion, callbacks over an array, properties of an
' object of the script and of a host object. Each workload is written once in each language with
' the same result; "Compile" is a fresh object plus compiling the source (stdLambda's protInit,
' bypassing the text cache of stdLambda.Create), "Run" is one call of the compiled object with the
' workload's arguments, as a callback is called. "Array transformations" is each library used as it
' is meant to be: stdLambda called from VBA once per item (as stdArray.Map calls it), ASF given the
' array and calling its own function once per item (map).
' VBA only (Excel): run it with Benchmark.ps1.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private source As String, args As Variant, expected As String, perItem As Boolean
Private engineName As String
Private lambda As Object, asfEngine As Object, asfIndex As Long
Private checksum As Long

Private Function Seconds() As Double
    Static frequency As Currency
    Dim counter As Currency
    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    Seconds = counter / frequency
End Function

Private Sub Configure(ByVal workload As String)
    Dim items As Collection
    Dim isAsf As Boolean: isAsf = (engineName = "ASF")
    Dim numbers() As Variant, i As Long
    perItem = False
    Select Case workload
        Case "Arithmetic"
            ' instruction throughput: an expression of the loop counter, 1000 times
            If isAsf Then
                source = "let s = 0; for (let i = 1, i <= 1000, i += 1) { s = s + (i * 3 - i / 2) * 2; }; return s;"
            Else
                source = "let s = 0: for i = 1 to 1000 let s = s + (i * 3 - i / 2) * 2 next: s"
            End If
            args = Array()
            expected = "2502500"
        Case "Math operations"
            ' built-in function dispatch: three functions per pass, 500 passes
            If isAsf Then
                source = "let s = 0; for (let i = 1, i <= 500, i += 1) { s = s + Math.abs(-i) + Math.floor(i / 3) + Math.sqrt(i * i); }; return s;"
            Else
                source = "let s = 0: for i = 1 to 500 let s = s + abs(-i) + int(i / 3) + sqr(i * i) next: s"
            End If
            args = Array()
            expected = "292083"
        Case "Conditional branches"
            ' jumps: two if/else per pass, 1000 passes
            If isAsf Then
                source = "let n = 0; for (let i = 1, i <= 1000, i += 1) { if (i % 3 == 0) { n = n + 2; } else { n = n - 1; }; if (i % 5 == 0) { n = n + 3; } else { n = n + 0; }; }; return n;"
            Else
                source = "let n = 0: for i = 1 to 1000 if i mod 3 = 0 then let n = n + 2 else let n = n - 1 end if i mod 5 = 0 then let n = n + 3 else let n = n + 0 end next: n"
            End If
            args = Array()
            expected = "599"
        Case "For/While loops"
            ' dispatch of the loops themselves: a while inside a for, 1000 passes in all
            If isAsf Then
                source = "let n = 0; for (let i = 1, i <= 40, i += 1) { let j = 0; while (j < 25) { j = j + 1; n = n + 1; }; }; return n;"
            Else
                source = "let n = 0: for i = 1 to 40 let j = 0 do while j < 25 let j = j + 1 let n = n + 1 loop next: n"
            End If
            args = Array()
            expected = "1000"
        Case "Function calls"
            ' frames: 500 calls of a two-argument function
            If isAsf Then
                source = "fun add(a, b) { return a + b; }; let s = 0; for (let i = 1, i <= 500, i += 1) { s = add(s, i); }; return s;"
            Else
                source = "fun add(a, b) a + b end: let s = 0: for i = 1 to 500 let s = add(s, i) next: s"
            End If
            args = Array()
            expected = "125250"
        Case "Fibonacci recursion"
            ' recursive calls: fib(15), 1973 calls
            If isAsf Then
                source = "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return fib($1);"
            Else
                source = "fun fib(n) if n <= 1 then n else fib(n - 1) + fib(n - 2) end end: fib($1)"
            End If
            args = Array(15)
            expected = "610"
        Case "Array transformations"
            ' callbacks: map over 300 items; stdLambda is called from VBA once per item, as
            ' stdArray.Map calls it, ASF maps the array it is given
            If isAsf Then
                source = "return $1.map(fun(x) { return x * 2 + 1; });"
            Else
                source = "$1 * 2 + 1"
                perItem = True
            End If
            ReDim numbers(1 To 300)
            For i = 1 To 300: numbers(i) = i: Next i
            args = Array(numbers)
            expected = "90600"
        Case "Object property access"
            ' lookups in an object the script makes: two properties per pass, 500 passes
            If isAsf Then
                source = "let o = {x: 3, y: 4}; let s = 0; for (let i = 1, i <= 500, i += 1) { s = s + o.x * o.y; }; return s;"
            Else
                source = "let o = dict(""x"", 3, ""y"", 4): let s = 0: for i = 1 to 500 let s = s + o.x * o.y next: s"
            End If
            args = Array()
            expected = "6000"
        Case "COM property access"
            ' a host object: Count once, then Item of each of 100 members of a Collection
            If isAsf Then
                source = "let n = $1.count; let s = 0; for (let i = 1, i <= n, i += 1) { s = s + $1.item(i); }; return s;"
            Else
                source = "let s = 0: for i = 1 to $1.Count let s = s + $1.Item(i) next: s"
            End If
            Set items = New Collection
            For i = 1 To 100: items.Add i: Next i
            args = Array(items)
            expected = "5050"
        Case Else
            Err.Raise 5, "LambdaBenchmark", "Unknown workload: " & workload
    End Select
End Sub

Private Sub CompileSource()
    If engineName = "ASF" Then
        Set asfEngine = New ASF
        asfEngine.AppAccess = True
        asfIndex = asfEngine.Compile(source)
    Else
        Set lambda = New stdLambda
        lambda.protInit 1, source, False, False
    End If
End Sub

Private Function RunOnce() As Variant
    Dim r As Variant, i As Long, mapped() As Variant
    If perItem Then
        ' the mapped array, as stdArray.Map builds it, then its sum (as for ASF's below)
        ReDim mapped(LBound(args(0)) To UBound(args(0)))
        For i = LBound(args(0)) To UBound(args(0))
            mapped(i) = lambda.Run(args(0)(i))
        Next i
        RunOnce = SumOf(mapped)
        Exit Function
    ElseIf engineName = "ASF" Then
        Select Case UBound(args)
            Case -1: r = asfEngine.Run(asfIndex)
            Case 0: r = asfEngine.Run(asfIndex, args(0))
            Case 1: r = asfEngine.Run(asfIndex, args(0), args(1))
        End Select
    Else
        Select Case UBound(args)
            Case -1: r = lambda.Run()
            Case 0: r = lambda.Run(args(0))
            Case 1: r = lambda.Run(args(0), args(1))
        End Select
    End If
    If IsArray(r) Then r = SumOf(r)
    RunOnce = r
End Function

Private Function SumOf(ByRef values As Variant) As Double
    Dim v As Variant
    For Each v In values
        SumOf = SumOf + v
    Next v
End Function

Private Function Validate() As String
    Dim r As Variant, same As Boolean
    r = RunOnce()
    If VarType(r) = vbBoolean Then
        same = (expected = IIf(r, "True", "False"))
    ElseIf IsNumeric(r) And VarType(r) <> vbString Then
        same = (Abs(CDbl(r) - Val(expected)) < 0.000000001)
    Else
        same = (CStr(r) = expected)
    End If
    If Not same Then Validate = "Returned " & CStr(r) & "; expected " & expected
End Function

Private Function Batch(ByVal operation As String, ByVal iterations As Long) As Double
    Dim started As Double, i As Long
    started = Seconds()
    If operation = "Compile" Then
        For i = 1 To iterations
            CompileSource
        Next
    Else
        For i = 1 To iterations
            checksum = checksum Xor Len(CStr(RunOnce()))
        Next
    End If
    Batch = Seconds() - started
End Function

Private Function Quote(ByVal value As String) As String
    Quote = """" & Replace(value, """", """""") & """"
End Function

'Returns one CSV row. Raw samples are microseconds per call. Five batches follow calibration
'to >= 30 ms, with a maximum of 16384 calls/batch.
Public Function Measure(ByVal whichEngine As String, ByVal workload As String, ByVal operation As String) As String
    On Error GoTo Failed
    Dim issue As String, iterations As Long, elapsed As Double, sample As Long, row As String, baseRow As String
    engineName = whichEngine
    Configure workload
    row = Quote(engineName) & "," & Quote(workload) & "," & Quote(operation)
    baseRow = row
    CompileSource
    issue = Validate()
    If Len(issue) > 0 Then
        Measure = row & ",mismatch," & Quote(issue) & ",0,,,,,"
        Exit Function
    End If
    iterations = 1
    Do
        elapsed = Batch(operation, iterations)
        If elapsed >= 0.03 Or iterations >= 16384 Then Exit Do
        iterations = iterations * 2
    Loop
    row = row & ",ok,," & iterations
    For sample = 1 To 5
        elapsed = Batch(operation, iterations)
        row = row & "," & Trim$(Str$(elapsed * 1000000# / iterations))
    Next
    Measure = row
    Exit Function
Failed:
    Measure = baseRow & ",error," & Quote(CStr(Err.Number) & " " & Err.Description) & ",0,,,,,"
End Function
