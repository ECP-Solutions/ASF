Attribute VB_Name = "ExcelHost"
' ExcelHost: runs the portable ASF examples in Excel.
' Needs the ASF class modules (src/*.cls) in the workbook, saved as .xlsm in examples/portable,
' next to the scripts folder, with one sheet per example holding its table from cell A1.
' CalcHost.bas is the same host for LibreOffice Calc: the scripts do not change.
Option Explicit

Public Sub RunBudget()
    RunExample ThisWorkbook, "budget.vas", "Budget", "", "Budget result"
End Sub

Public Sub RunWorkReports()
    RunExample ThisWorkbook, "workreports.vas", "WorkReports", "Budget", "WorkReports result"
End Sub

Public Sub RunProgress()
    RunExample ThisWorkbook, "progress.vas", "Progress", "", "Progress result"
End Sub

Public Sub RunPriceRevision()
    RunExample ThisWorkbook, "price-revision.vas", "PriceIndices", "", "PriceIndices result"
End Sub

Public Sub RunCleanup()
    RunExample ThisWorkbook, "cleanup.vas", "Cleanup", "", "Cleanup result"
End Sub

Public Sub RunAll()
    RunBudget
    RunWorkReports
    RunProgress
    RunPriceRevision
    RunCleanup
End Sub

' Runs script on the table of sheetName (and of extraSheet, as $2, when given); its rows go to outName
Public Sub RunExample(ByVal wb As Workbook, ByVal script As String, ByVal sheetName As String, _
                      ByVal extraSheet As String, ByVal outName As String)
    On Error GoTo Failed
    Dim eng As ASF: Set eng = New ASF
    Dim folder As String: folder = eng.JoinPath(wb.Path, "scripts")
    eng.WorkingDir = folder
    Dim result As Variant
    If Len(extraSheet) = 0 Then
        result = eng.Execute(eng.JoinPath(folder, script), ReadTable(wb, sheetName))
    Else
        result = eng.Execute(eng.JoinPath(folder, script), ReadTable(wb, sheetName), ReadTable(wb, extraSheet))
    End If
    WriteTable wb, outName, result
    Exit Sub
Failed:
    WriteTable wb, outName, "error " & Err.Number & ": " & Err.Description
End Sub

' The table around A1; blank cells as "", as LibreOffice gives them (ASF reads Empty as 0)
Private Function ReadTable(ByVal wb As Workbook, ByVal sheetName As String) As Variant
    Dim t As Variant, i As Long, j As Long
    t = wb.Worksheets(sheetName).Range("A1").CurrentRegion.Value2
    For i = 1 To UBound(t, 1)
        For j = 1 To UBound(t, 2)
            If IsEmpty(t(i, j)) Then t(i, j) = vbNullString
        Next j
    Next i
    ReadTable = t
End Function

' Writes the script's rows (an array of arrays) from A1 of sheetName, created if needed
Private Sub WriteTable(ByVal wb As Workbook, ByVal sheetName As String, ByVal rows As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.ClearContents
    If Not IsArray(rows) Then
        ws.Range("A1").Value2 = "Script error: " & CStr(rows)
        Exit Sub
    End If
    Dim n As Long, w As Long, i As Long, j As Long
    n = UBound(rows) - LBound(rows) + 1
    For i = LBound(rows) To UBound(rows)
        If UBound(rows(i)) - LBound(rows(i)) + 1 > w Then w = UBound(rows(i)) - LBound(rows(i)) + 1
    Next i
    Dim m() As Variant
    ReDim m(1 To n, 1 To w)
    For i = 1 To n
        For j = 1 To w
            m(i, j) = vbNullString
            If LBound(rows(LBound(rows) + i - 1)) + j - 1 <= UBound(rows(LBound(rows) + i - 1)) Then
                m(i, j) = rows(LBound(rows) + i - 1)(LBound(rows(LBound(rows) + i - 1)) + j - 1)
            End If
        Next j
    Next i
    ws.Range("A1").Resize(n, w).Value2 = m
    ws.Columns.AutoFit
End Sub
