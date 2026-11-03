Attribute VB_Name = "LambdaBenchmark"
' stdLambda (stdVBA) against ASF on the same workloads. Each workload is written once in each
' language with the same result; "Compile" is a fresh object plus compiling the source (stdLambda's
' protInit, bypassing the text cache of stdLambda.Create), "Run" is one call of the compiled object
' with the workload's arguments, as a callback is called. "Callback over 300 items" is each library
' used as it is meant to be: stdLambda called from VBA once per item (as stdArray.Map calls it),
' ASF given the items and calling its own function once per item (forEach).
' Run it with Benchmark.ps1 (Excel) or BenchmarkLO.py (LibreOffice).
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
        Case "Arithmetic expression"
            If isAsf Then source = "return ($1 + $2) * 3 - $1 / 2;" Else source = "($1 + $2) * 3 - $1 / 2"
            args = Array(7, 5)
            expected = "32.5"
        Case "Predicate"
            If isAsf Then source = "return $1 % 3 == 0 && $1 > 10;" Else source = "$1 mod 3 = 0 and $1 > 10"
            args = Array(12)
            expected = "True"
        Case "String functions"
            If isAsf Then source = "return $1.trim().toUpperCase() + $1.length;" Else source = "ucase(trim($1)) & len($1)"
            args = Array("  abc  ")
            expected = "ABC7"
        Case "Conditional"
            If isAsf Then source = "return $1 > 5 ? 'big' : 'small';" Else source = "if $1 > 5 then ""big"" else ""small"""
            args = Array(9)
            expected = "big"
        Case "Local variables"
            If isAsf Then source = "let a = $1 * 2; let b = a + 3; return a * b;" Else source = "let a = $1 * 2: let b = a + 3: a * b"
            args = Array(4)
            expected = "88"
        Case "Object members"
            If isAsf Then source = "return $1.count + $1.item(2);" Else source = "$1.Count + $1.Item(2)"
            Set items = New Collection
            items.Add 10: items.Add 20: items.Add 30
            args = Array(items)
            expected = "23"
        Case "Loop"
            If isAsf Then
                source = "let s = 0; for (let i = 1, i <= 2000, i += 1) { s = s + i * 2 - 1; }; return s;"
            Else
                source = "let s = 0: for i = 1 to 2000 let s = s + i * 2 - 1 next: s"
            End If
            args = Array()
            expected = "4000000"
        Case "While and branches"
            If isAsf Then
                source = "let i = 0; let n = 0; while (i < 1000) { if (i % 3 == 0) { n = n + 2; } else { n = n - 1; }; i += 1; }; return n;"
            Else
                source = "let i = 0: let n = 0: do while i < 1000 if i mod 3 = 0 then let n = n + 2 else let n = n - 1 end let i = i + 1 loop: n"
            End If
            args = Array()
            expected = "2"
        Case "Recursion"
            If isAsf Then
                source = "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return fib($1);"
            Else
                source = "fun fib(n) if n <= 1 then n else fib(n - 1) + fib(n - 2) end end: fib($1)"
            End If
            args = Array(14)
            expected = "377"
        Case "Callback over 300 items"
            If isAsf Then
                source = "let s = 0; $1.forEach(fun(x) { s = s + x * 2 + 1; }); return s;"
            Else
                source = "$1 * 2 + 1"
                perItem = True
            End If
            ReDim numbers(1 To 300)
            For i = 1 To 300: numbers(i) = i: Next i
            args = Array(numbers)
            expected = "90600"
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
    Dim r As Variant, i As Long
    If perItem Then
        r = 0
        For i = LBound(args(0)) To UBound(args(0))
            r = r + lambda.Run(args(0)(i))
        Next i
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
    RunOnce = r
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
