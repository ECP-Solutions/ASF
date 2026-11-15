Attribute VB_Name = "RulesHost"
' Discounts from rules that users edit in the Rules sheet, without touching VBA.
' Sheets: Rules (Name | When | Discount), Orders (order | region | years | sku | qty | price),
' Discounts (written by ApplyDiscountRules). discounts.vas sits next to the workbook; import this
' module with the ASF class modules and ASFSheets.bas.
Option Explicit

Public Sub ApplyDiscountRules()
    Dim rules As Variant, orders As Variant, problem As String
    rules = ThisWorkbook.Worksheets("Rules").Range("A1").CurrentRegion.Value2
    orders = ThisWorkbook.Worksheets("Orders").Range("A1").CurrentRegion.Value2
    problem = CheckRules(rules)
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation, "Discount rules"
        Exit Sub
    End If

    Dim eng As New ASF, idx As Long, result As Variant
    idx = eng.Compile(BuildConditions(rules) & eng.ReadTextFile(eng.JoinPath(ThisWorkbook.Path, "discounts.vas")))
    result = eng.Run(idx, orders, rules)
    If IsEmpty(result) Then
        MsgBox "Stopped: " & eng.LastError, vbExclamation, "Discount rules"
        Exit Sub
    End If
    With ThisWorkbook.Worksheets("Discounts")
        .Cells.ClearContents
        WriteRows .Range("A1"), result
    End With
End Sub

' The When column as ASF functions: let conditions = [fun(o) { return (...); }, ...];
Public Function BuildConditions(ByVal rules As Variant) As String
    Dim r As Long, s As String
    For r = LBound(rules, 1) + 1 To UBound(rules, 1)
        If Len(s) > 0 Then s = s & "," & vbLf
        s = s & "    fun(o) { return (" & rules(r, LBound(rules, 2) + 1) & "); }"
    Next r
    BuildConditions = "let conditions = [" & vbLf & s & vbLf & "];" & vbLf
End Function

' Compiles each rule alone, so that a mistake names its rule; empty when all compile
Public Function CheckRules(ByVal rules As Variant) As String
    Dim r As Long, eng As ASF
    For r = LBound(rules, 1) + 1 To UBound(rules, 1)
        Set eng = New ASF
        On Error Resume Next
        eng.Compile "let o = {}; return (" & rules(r, LBound(rules, 2) + 1) & ");"
        If Err.Number <> 0 Then
            CheckRules = "Rule '" & rules(r, LBound(rules, 2)) & "': " & Err.Description
            Exit Function
        End If
        On Error GoTo 0
    Next r
End Function
