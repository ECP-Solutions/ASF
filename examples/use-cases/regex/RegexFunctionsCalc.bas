' RegexFunctions.bas for LibreOffice Calc: the same worksheet functions on the regex engine of the
' ASF extension (ASF.oxt). Add this module to the Standard library of the spreadsheet (Tools >
' Macros > Edit Macros) and use the functions in cells, with ; between arguments:
'   =RX_TEST(A2; "^[A-Z]{2}-\d{4}$")    =RX_EXTRACT(A2; "(\d+)\s*kg"; 1)    =RX_REPLACE(A2; "\s+"; " ")
' A bad pattern gives the text #VALUE! (a Basic function cannot return Calc's error values).
Option VBASupport 1
Option Explicit

Private Const BAD_PATTERN As String = "#VALUE!"
Private engine As Object

Private Function Compiled(ByVal pattern As String, ByVal ignoreCase As Boolean) As Object
    If engine Is Nothing Then
        GlobalScope.BasicLibraries.LoadLibrary("ASF")
        Set engine = ASF.LOhelper.NewRegex()
    End If
    engine.Init pattern, ignoreCase
    Set Compiled = engine
End Function

Function RX_TEST(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean) As Variant
    On Error GoTo Bad
    If IsMissing(ignoreCase) Then ignoreCase = False
    RX_TEST = Compiled(pattern, ignoreCase).test(text)
    Exit Function
Bad:
    RX_TEST = BAD_PATTERN
End Function

Function RX_EXTRACT(ByVal text As String, ByVal pattern As String, Optional ByVal group As Long, _
                    Optional ByVal ignoreCase As Boolean) As Variant
    On Error GoTo Bad
    If IsMissing(group) Then group = 0
    If IsMissing(ignoreCase) Then ignoreCase = False
    Dim m As Object
    Set m = Compiled(pattern, ignoreCase).Exec(text)
    If m Is Nothing Then
        RX_EXTRACT = ""
    ElseIf group + 1 > m.Count Then
        RX_EXTRACT = BAD_PATTERN
    Else
        RX_EXTRACT = m.Item(group + 1)
    End If
    Exit Function
Bad:
    RX_EXTRACT = BAD_PATTERN
End Function

Function RX_REPLACE(ByVal text As String, ByVal pattern As String, ByVal replacement As String, _
                    Optional ByVal ignoreCase As Boolean) As Variant
    On Error GoTo Bad
    If IsMissing(ignoreCase) Then ignoreCase = False
    RX_REPLACE = Compiled(pattern, ignoreCase).Replace(text, replacement)
    Exit Function
Bad:
    RX_REPLACE = BAD_PATTERN
End Function

Function RX_COUNT(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean) As Variant
    On Error GoTo Bad
    If IsMissing(ignoreCase) Then ignoreCase = False
    RX_COUNT = Compiled(pattern, ignoreCase).ExecAll(text).Count
    Exit Function
Bad:
    RX_COUNT = BAD_PATTERN
End Function
