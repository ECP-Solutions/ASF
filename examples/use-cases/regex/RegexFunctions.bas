Attribute VB_Name = "RegexFunctions"
' Worksheet functions on ASF's regex engine: pure VBA, so the same results in Excel for Windows
' and for Mac (which has no VBScript.RegExp), with JavaScript's regex syntax: lookbehind, named
' groups, lazy quantifiers, \b, backreferences. Import it with ASF_RegexEngine.cls.
'   =RX_TEST(A2, "^[A-Z]{2}-\d{4}$")          TRUE when the text matches
'   =RX_EXTRACT(A2, "(\d+)\s*kg", 1)          the first match, or one of its groups
'   =RX_REPLACE(A2, "\s+", " ")               every match replaced ($1, $2... for groups)
'   =RX_COUNT(A2, "\b\w+\b")                  the number of matches
' A bad pattern gives #VALUE!.
Option Explicit

Private engine As ASF_RegexEngine

' The engine with pattern loaded; it keeps the patterns it has compiled, so a column of formulas
' with the same pattern compiles it once
Private Function Compiled(ByVal pattern As String, ByVal ignoreCase As Boolean) As ASF_RegexEngine
    If engine Is Nothing Then Set engine = New ASF_RegexEngine
    engine.Init pattern, ignoreCase
    Set Compiled = engine
End Function

Public Function RX_TEST(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As Variant
    On Error GoTo Bad
    RX_TEST = Compiled(pattern, ignoreCase).test(text)
    Exit Function
Bad:
    RX_TEST = CVErr(2015) ' #VALUE!
End Function

Public Function RX_EXTRACT(ByVal text As String, ByVal pattern As String, Optional ByVal group As Long = 0, _
                           Optional ByVal ignoreCase As Boolean = False) As Variant
    On Error GoTo Bad
    Dim m As Collection
    Set m = Compiled(pattern, ignoreCase).Exec(text)
    If m Is Nothing Then
        RX_EXTRACT = vbNullString
    ElseIf group + 1 > m.Count Then
        RX_EXTRACT = CVErr(2015)
    Else
        RX_EXTRACT = m(group + 1)
    End If
    Exit Function
Bad:
    RX_EXTRACT = CVErr(2015)
End Function

Public Function RX_REPLACE(ByVal text As String, ByVal pattern As String, ByVal replacement As String, _
                           Optional ByVal ignoreCase As Boolean = False) As Variant
    On Error GoTo Bad
    RX_REPLACE = Compiled(pattern, ignoreCase).Replace(text, replacement)
    Exit Function
Bad:
    RX_REPLACE = CVErr(2015)
End Function

Public Function RX_COUNT(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As Variant
    On Error GoTo Bad
    RX_COUNT = Compiled(pattern, ignoreCase).ExecAll(text).Count
    Exit Function
Bad:
    RX_COUNT = CVErr(2015)
End Function
