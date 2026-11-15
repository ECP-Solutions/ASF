Attribute VB_Name = "LambdaBenchmark"
' stdLambda (stdVBA) against ASF in seven categories: simple expressions, math functions, functions
' of the script, methods of COM objects that are not collections (Excel ranges, FileSystemObject),
' methods of COM collections (Collection, Dictionary), short programs and long programs. Each
' workload is written once in each language with the same result, checked before timing.
' "Compile" is a fresh object plus compiling the source (stdLambda's protInit, bypassing the text
' cache of stdLambda.Create), "Run" is one call of the compiled object with the workload's
' arguments; "Loop of 5, called 100 times" is called from VBA 100 times with 1 to 100.
' The stdLambda sources keep to what its parser handles: a for loop's variable is declared before
' the loop (a new one breaks a body that sets two variables), every if has its own end (an else if
' inside a function takes the function's end), true is lower case, and the members of a
' Dictionary are written in lower case (count, item, exists), else they read keys.
' VBA only (Excel): run it with Benchmark.ps1.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private source As String, args As Variant, expected As String
Private hostCalls As Long              ' > 0: VBA calls the program this many times, with 1, 2, 3...
Private engineName As String, workloadCategory As String
Private lambda As Object, asfEngine As Object, asfIndex As Long
Private checksum As Long
Private program As String

Private Function Seconds() As Double
    Static frequency As Currency
    Dim counter As Currency
    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    Seconds = counter / frequency
End Function

' category|workload, one per line, in the order Benchmark.ps1 runs them
Public Function Workloads() As String
    Dim s As String
    s = s & "Simple expressions|Sum and product" & vbLf
    s = s & "Simple expressions|Power and modulo" & vbLf
    s = s & "Simple expressions|Comparison and logic" & vbLf
    s = s & "Simple expressions|String concatenation" & vbLf
    s = s & "Math functions|Rounding and roots" & vbLf
    s = s & "Math functions|Trigonometry and logarithms" & vbLf
    s = s & "Math functions|String functions" & vbLf
    s = s & "Script functions|One function" & vbLf
    s = s & "Script functions|Nested calls" & vbLf
    s = s & "Script functions|Recursion" & vbLf
    s = s & "COM objects|Range properties" & vbLf
    s = s & "COM objects|Range methods" & vbLf
    s = s & "COM objects|FileSystemObject methods" & vbLf
    s = s & "COM collections|Collection items" & vbLf
    s = s & "COM collections|Dictionary lookups" & vbLf
    s = s & "COM collections|Collection loop" & vbLf
    s = s & "Short programs|Variables and branch" & vbLf
    s = s & "Short programs|Loop of 5, called 100 times" & vbLf
    s = s & "Short programs|String building" & vbLf
    s = s & "Long programs|Number theory" & vbLf
    s = s & "Long programs|Text report" & vbLf
    Workloads = s
End Function

Private Function CategoryOf(ByVal workload As String) As String
    Dim line As Variant
    For Each line In Split(Workloads(), vbLf)
        If Mid$(line, InStr(line, "|") + 1) = workload And Len(line) > 0 Then
            CategoryOf = Left$(line, InStr(line, "|") - 1)
            Exit Function
        End If
    Next line
    Err.Raise 5, "LambdaBenchmark", "Unknown workload: " & workload
End Function

' a cell of the first sheet, with the values the range workloads read
Private Function Cell(ByVal address As String) As Object
    With ThisWorkbook.Worksheets(1)
        .Range("B3").Value = 21
        .Range("C4").Value = 30
        .Range("D4").Value = 40
        Set Cell = .Range(address)
    End With
End Function

' a Collection of 10, 20, 30...
Private Function Items(ByVal count As Long) As Collection
    Dim c As New Collection, i As Long
    For i = 1 To count
        c.Add i * 10
    Next i
    Set Items = c
End Function

Private Function Lookup() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    d.Add "b", 2
    d.Add "c", 3
    Set Lookup = d
End Function

' one line of a long program: ASF reads the lines, stdLambda gets them joined by spaces, as
' stdLambda.CreateMultiline joins them
Private Sub L(ByVal text As String)
    If engineName = "ASF" Then
        program = program & text & vbLf
    Else
        program = program & text & " "
    End If
End Sub

Private Sub Configure(ByVal workload As String)
    Dim isAsf As Boolean: isAsf = (engineName = "ASF")
    hostCalls = 0
    workloadCategory = CategoryOf(workload)
    Select Case workload
        Case "Sum and product"
            If isAsf Then
                source = "return ($1 + $2) * 3 - $1 / 2;"
            Else
                source = "($1 + $2) * 3 - $1 / 2"
            End If
            args = Array(7, 5)
            expected = "32.5"
        Case "Power and modulo"
            If isAsf Then
                source = "return $1 ^ 2 + $2 % 4 - ($1 - $2) / 2;"
            Else
                source = "$1 ^ 2 + $2 mod 4 - ($1 - $2) / 2"
            End If
            args = Array(7, 5)
            expected = "49"
        Case "Comparison and logic"
            If isAsf Then
                source = "return ($1 > 5 && $2 <= 5) || $1 == $2 ? 'yes' : 'no';"
            Else
                source = "if $1 > 5 and $2 <= 5 or $1 = $2 then ""yes"" else ""no"""
            End If
            args = Array(7, 5)
            expected = "yes"
        Case "String concatenation"
            If isAsf Then
                source = "return $1 + '-' + $2 + '!';"
            Else
                source = "$1 & ""-"" & $2 & ""!"""
            End If
            args = Array("abc", 7)
            expected = "abc-7!"
        Case "Rounding and roots"
            If isAsf Then
                source = "return Math.abs($1) + Math.sqrt($2) + Math.floor($1 / 3);"
            Else
                source = "abs($1) + sqr($2) + int($1 / 3)"
            End If
            args = Array(-7, 16)
            expected = "8"
        Case "Trigonometry and logarithms"
            If isAsf Then
                source = "return Math.sin($1) ^ 2 + Math.cos($1) ^ 2 + Math.exp(Math.log($2)) + Math.atan(1) * 4;"
            Else
                source = "sin($1) ^ 2 + cos($1) ^ 2 + exp(log($2)) + atn(1) * 4"
            End If
            args = Array(7, 5)
            expected = "9.14159265358979"
        Case "String functions"
            If isAsf Then
                source = "return $1.slice(0, 3).toUpperCase() + $1.length + $1.substring(4, 7);"
            Else
                source = "ucase(left($1, 3)) & len($1) & mid($1, 5, 3)"
            End If
            args = Array("benchmark")
            expected = "BEN9hma"
        Case "One function"
            If isAsf Then
                source = "fun sq(x) { return x * x; }; return sq($1) + sq($2);"
            Else
                source = "fun sq(x) x * x end: sq($1) + sq($2)"
            End If
            args = Array(3, 4)
            expected = "25"
        Case "Nested calls"
            If isAsf Then
                source = "fun add(a, b) { return a + b; }; fun mul(a, b) { return a * b; }; fun poly(x) { return add(mul(x, x), mul(3, x)); }; return poly($1) + poly($2);"
            Else
                source = "fun add(a, b) a + b end: fun mul(a, b) a * b end: fun poly(x) add(mul(x, x), mul(3, x)) end: poly($1) + poly($2)"
            End If
            args = Array(3, 4)
            expected = "46"
        Case "Recursion"
            If isAsf Then
                source = "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return fib($1);"
            Else
                source = "fun fib(n) if n <= 1 then n else fib(n - 1) + fib(n - 2) end end: fib($1)"
            End If
            args = Array(10)
            expected = "55"
        Case "Range properties"
            If isAsf Then
                source = "return $1.value * 2 + $1.row + $1.column;"
            Else
                source = "$1.Value * 2 + $1.Row + $1.Column"
            End If
            args = Array(Cell("B3"))
            expected = "47"
        Case "Range methods"
            If isAsf Then
                source = "return $1.offset(1, 2).value + $1.worksheet.range('C4').value;"
            Else
                source = "$1.Offset(1, 2).Value + $1.Worksheet.Range(""C4"").Value"
            End If
            args = Array(Cell("B3"))
            expected = "70"
        Case "FileSystemObject methods"
            If isAsf Then
                source = "return $1.getbasename('report.txt').length + $1.getextensionname('report.txt').length + $1.buildpath('a', 'b').length;"
            Else
                source = "len($1.GetBaseName(""report.txt"")) + len($1.GetExtensionName(""report.txt"")) + len($1.BuildPath(""a"", ""b""))"
            End If
            args = Array(CreateObject("Scripting.FileSystemObject"))
            expected = "12"
        Case "Collection items"
            If isAsf Then
                source = "return $1.count + $1.item(2);"
            Else
                source = "$1.Count + $1.Item(2)"
            End If
            args = Array(Items(3))
            expected = "23"
        Case "Dictionary lookups"
            If isAsf Then
                source = "return $1.count + $1.item('b') + ($1.exists('z') ? 100 : 0);"
            Else
                source = "$1.count + $1.item(""b"") + (if $1.exists(""z"") then 100 else 0)"
            End If
            args = Array(Lookup())
            expected = "5"
        Case "Collection loop"
            If isAsf Then
                source = "let s = 0; for (let i = 1, i <= $1.count, i += 1) { s = s + $1.item(i); }; return s;"
            Else
                source = "let s = 0: for i = 1 to $1.count let s = s + $1.item(i) next: s"
            End If
            args = Array(Items(10))
            expected = "550"
        Case "Variables and branch"
            If isAsf Then
                source = "let a = $1 * 2; let b = a + $2; return b > 10 ? b - 10 : b;"
            Else
                source = "let a = $1 * 2: let b = a + $2: if b > 10 then b - 10 else b end"
            End If
            args = Array(7, 5)
            expected = "9"
        Case "Loop of 5, called 100 times"
            If isAsf Then
                source = "let s = 0; for (let i = 1, i <= 5, i += 1) { s = s + i * $1; }; return s;"
            Else
                source = "let s = 0: for i = 1 to 5 let s = s + i * $1 next: s"
            End If
            hostCalls = 100
            args = Array()
            expected = "75750"
        Case "String building"
            If isAsf Then
                source = "let s = ''; for (let i = 1, i <= 5, i += 1) { s = s + i; }; return s + $1;"
            Else
                source = "let s = """": for i = 1 to 5 let s = s & i next: s & $1"
            End If
            args = Array("!")
            expected = "12345!"
        Case "Number theory"
            If isAsf Then NumberTheoryAsf Else NumberTheoryLambda
            source = program
            args = Array(30)
            expected = "10097441"
        Case "Text report"
            If isAsf Then TextReportAsf Else TextReportLambda
            source = program
            args = Array(20, 37)
            expected = "120|939|94|..1:C;..2:A;..3:D;..4:C;"
    End Select
End Sub

Private Sub NumberTheoryAsf()
    program = vbNullString
    L "fun isprime(n) {"
    L "    let r = 1;"
    L "    if (n < 2) {"
    L "        r = 0;"
    L "    };"
    L "    let d = 2;"
    L "    while (d * d <= n && r == 1) {"
    L "        if (n % d == 0) {"
    L "            r = 0;"
    L "        };"
    L "        d = d + 1;"
    L "    };"
    L "    return r;"
    L "};"
    L "fun gcd(a, b) {"
    L "    let t = 0;"
    L "    while (b != 0) {"
    L "        t = b;"
    L "        b = a % b;"
    L "        a = t;"
    L "    };"
    L "    return a;"
    L "};"
    L "fun collatz(n) {"
    L "    let c = 0;"
    L "    while (n != 1) {"
    L "        if (n % 2 == 0) {"
    L "            n = n / 2;"
    L "        } else {"
    L "            n = 3 * n + 1;"
    L "        };"
    L "        c = c + 1;"
    L "    };"
    L "    return c;"
    L "};"
    L "let primes = 0;"
    L "let g = 0;"
    L "let steps = 0;"
    L "for (let k = 1, k <= $1, k += 1) {"
    L "    primes = primes + isprime(k);"
    L "    g = g + gcd(k, 12);"
    L "    steps = steps + collatz(k);"
    L "};"
    L "return primes * 1000000 + g * 1000 + steps;"
End Sub

Private Sub NumberTheoryLambda()
    program = vbNullString
    L "fun isprime(n)"
    L "    let r = 1"
    L "    if n < 2 then"
    L "        let r = 0"
    L "    end"
    L "    let d = 2"
    L "    do while d * d <= n and r = 1"
    L "        if n mod d = 0 then"
    L "            let r = 0"
    L "        end"
    L "        let d = d + 1"
    L "    loop"
    L "    r"
    L "end"
    L "fun gcd(a, b)"
    L "    let t = 0"
    L "    do while b <> 0"
    L "        let t = b"
    L "        let b = a mod b"
    L "        let a = t"
    L "    loop"
    L "    a"
    L "end"
    L "fun collatz(n)"
    L "    let c = 0"
    L "    do while n <> 1"
    L "        if n mod 2 = 0 then"
    L "            let n = n / 2"
    L "        else"
    L "            let n = 3 * n + 1"
    L "        end"
    L "        let c = c + 1"
    L "    loop"
    L "    c"
    L "end"
    L "let primes = 0"
    L "let g = 0"
    L "let steps = 0"
    L "let k = 0"
    L "for k = 1 to $1"
    L "    let primes = primes + isprime(k)"
    L "    let g = g + gcd(k, 12)"
    L "    let steps = steps + collatz(k)"
    L "next"
    L "primes * 1000000 + g * 1000 + steps"
End Sub

Private Sub TextReportAsf()
    program = vbNullString
    L "fun grade(score) {"
    L "    if (score >= 90) {"
    L "        return 'A';"
    L "    } elseif (score >= 75) {"
    L "        return 'B';"
    L "    } elseif (score >= 50) {"
    L "        return 'C';"
    L "    } else {"
    L "        return 'D';"
    L "    };"
    L "};"
    L "fun pad(s, width) {"
    L "    s = '' + s;"
    L "    while (s.length < width) {"
    L "        s = '.' + s;"
    L "    };"
    L "    return s;"
    L "};"
    L "fun score(k, seed) {"
    L "    return (k * seed + 17) % 101;"
    L "};"
    L "let report = '';"
    L "let total = 0;"
    L "let best = 0;"
    L "let v = 0;"
    L "for (let k = 1, k <= $1, k += 1) {"
    L "    v = score(k, $2);"
    L "    total = total + v;"
    L "    if (v > best) {"
    L "        best = v;"
    L "    };"
    L "    report = report + pad(k, 3) + ':' + grade(v) + ';';"
    L "};"
    L "return report.length + '|' + total + '|' + best + '|' + report.substring(0, 24);"
End Sub

Private Sub TextReportLambda()
    program = vbNullString
    L "fun grade(score)"
    L "    if score >= 90 then"
    L "        ""A"""
    L "    else if score >= 75 then"
    L "        ""B"""
    L "    else if score >= 50 then"
    L "        ""C"""
    L "    else"
    L "        ""D"""
    L "    end"
    L "    end"
    L "    end"
    L "end"
    L "fun pad(s, width)"
    L "    do while len(s) < width"
    L "        let s = ""."" & s"
    L "    loop"
    L "    s"
    L "end"
    L "fun score(k, seed)"
    L "    (k * seed + 17) mod 101"
    L "end"
    L "let report = """""
    L "let total = 0"
    L "let best = 0"
    L "let v = 0"
    L "let k = 0"
    L "for k = 1 to $1"
    L "    let v = score(k, $2)"
    L "    let total = total + v"
    L "    if v > best then"
    L "        let best = v"
    L "    end"
    L "    let report = report & pad(k, 3) & "":"" & grade(v) & "";"""
    L "next"
    L "len(report) & ""|"" & total & ""|"" & best & ""|"" & mid(report, 1, 24)"
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
    If hostCalls > 0 Then
        r = 0
        For i = 1 To hostCalls
            If engineName = "ASF" Then r = r + asfEngine.Run(asfIndex, i) Else r = r + lambda.Run(i)
        Next i
        RunOnce = r
        Exit Function
    End If
    If engineName = "ASF" Then
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
    workloadCategory = vbNullString
    Configure workload
    row = Quote(engineName) & "," & Quote(workloadCategory) & "," & Quote(workload) & "," & Quote(operation)
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
    If Len(baseRow) = 0 Then baseRow = Quote(engineName) & "," & Quote(workloadCategory) & "," & Quote(workload) & "," & Quote(operation)
    Measure = baseRow & ",error," & Quote(CStr(Err.Number) & " " & Err.Description) & ",0,,,,,"
End Function
