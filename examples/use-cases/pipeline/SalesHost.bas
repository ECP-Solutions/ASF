Attribute VB_Name = "SalesHost"
' Sales summary with ASF: Sales (region | product | month | units | price) is read once, summarised
' by sales-summary.vas, next to the workbook, and written to the sheet Summary once.
' Import it with the ASF class modules and ASFSheets.bas.
Option Explicit

Public Sub SummarizeSales()
    Dim eng As New ASF, data As Variant, result As Variant
    data = ThisWorkbook.Worksheets("Sales").Range("A1").CurrentRegion.Value2
    result = eng.Run(eng.Compile(eng.ReadTextFile(eng.JoinPath(ThisWorkbook.Path, "sales-summary.vas"))), data)
    If IsEmpty(result) Then
        MsgBox "Stopped: " & eng.LastError, vbExclamation, "Sales summary"
        Exit Sub
    End If
    With ThisWorkbook.Worksheets("Summary")
        .Cells.ClearContents
        WriteRows .Range("A1"), result(LBound(result))
        WriteRows .Range("H1"), result(LBound(result) + 1)
    End With
End Sub
