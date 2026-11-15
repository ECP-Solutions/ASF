Attribute VB_Name = "LambdaBenchmark"
' stdLambda (stdVBA) against ASF on direct expression evaluation, where stdLambda is meant to be
' used: a short expression compiled once and called with arguments, as a callback is called. One
' workload per part of an engine; each is written once in each language with the same result.
' "Compile" is a fresh object plus compiling the source (stdLambda's protInit, bypassing the text
' cache of stdLambda.Create), "Run" is one call of the compiled object with the workload's
' arguments. "For/While loops" is a while loop of 5 passes in both engines, called 100 times from
' VBA; "Array transformations" calls stdLambda once per item (as stdArray.Map does) while ASF maps
' the array it is given.
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
Private hostCalls As Long              ' > 0: VBA calls the expression this many times, with 1, 2, 3...
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
    Dim items As Collection, numbers() As Variant, i As Long, obj As Object
    Dim isAsf As Boolean: isAsf = (engineName = "ASF")
    perItem = False: hostCalls = 0
    Select Case workload
        Case "Arithmetic"
            If isAsf Then source = "return ($1 + $2) * 3 - $1 / 2;" Else source = "($1 + $2) * 3 - $1 / 2"
            args = Array(7, 5)
            expected = "32.5"
        Case "Math operations"
            If isAsf Then
                source = "return Math.abs($1) + Math.sqrt($2) + Math.floor($1 / 3);"
            Else
                source = "abs($1) + sqr($2) + int($1 / 3)"
            End If
            args = Array(-7, 16)
            expected = "8"
        Case "Conditional branches"
            If isAsf Then source = "return $1 > 5 ? 'big' : 'small';" Else source = "if $1 > 5 then ""big"" else ""small"""
            args = Array(9)
            expected = "big"
        Case "For/While loops"
            ' a while loop of 5 passes in each engine, called 100 times from VBA with 1 to 100
            If isAsf Then
                source = "let i = 0; let s = 0; while (i < 5) { i = i + 1; s = s + i * $1; }; return s;"
            Else
                source = "let i = 0: let s = 0: do while i < 5 let i = i + 1 let s = s + i * $1 loop: s"
            End If
            hostCalls = 100
            args = Array()
            expected = "75750"
        Case "Function calls"
            If isAsf Then
                source = "fun sq(x) { return x * x; }; return sq($1) + sq($2);"
            Else
                source = "fun sq(x) x * x end: sq($1) + sq($2)"
            End If
            args = Array(3, 4)
            expected = "25"
        Case "Fibonacci recursion"
            ' fib(10): 177 calls
            If isAsf Then
                source = "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return fib($1);"
            Else
                source = "fun fib(n) if n <= 1 then n else fib(n - 1) + fib(n - 2) end end: fib($1)"
            End If
            args = Array(10)
            expected = "55"
        Case "Array transformations"
            ' x * 2 + 1 over 50 items: stdLambda called from VBA once per item, as stdArray.Map calls
            ' it; ASF maps the array it is given
            If isAsf Then
                source = "return $1.map(fun(x) { return x * 2 + 1; });"
            Else
                source = "$1 * 2 + 1"
                perItem = True
            End If
            ReDim numbers(1 To 50)
            For i = 1 To 50: numbers(i) = i: Next i
            args = Array(numbers)
            expected = "2600"
        Case "Object property access"
            ' each engine's own object: a Scripting.Dictionary for stdLambda (o.x reads its item), an
            ' ASF object for ASF
            If isAsf Then
                source = "return $1.x * $1.y;"
                Set obj = New ASF_Map
                obj.SetValue "x", 3
                obj.SetValue "y", 4
            Else
                source = "$1.x * $1.y"
                Set obj = CreateObject("Scripting.Dictionary")
                obj.Add "x", 3
                obj.Add "y", 4
            End If
            args = Array(obj)
            expected = "12"
        Case "COM property access"
            If isAsf Then source = "return $1.count + $1.item(2);" Else source = "$1.Count + $1.Item(2)"
            Set items = New Collection
            items.Add 10: items.Add 20: items.Add 30
            args = Array(items)
            expected = "23"
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
    If hostCalls > 0 Then
        r = 0
        For i = 1 To hostCalls
            If engineName = "ASF" Then r = r + asfEngine.Run(asfIndex, i) Else r = r + lambda.Run(i)
        Next i
        RunOnce = r
        Exit Function
    End If
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
