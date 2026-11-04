' CalcHost: runs the portable ASF examples in LibreOffice Calc.
' Needs the ASF extension (ASF.oxt). The document is saved in examples/portable, next to the
' scripts folder, with one sheet per example holding its table from cell A1.
' ExcelHost.bas is the same host for Excel: the scripts do not change.
Option Explicit

Sub RunBudget()
    RunExample ThisComponent, "budget.vas", "Budget", "", "Budget result"
End Sub

Sub RunWorkReports()
    RunExample ThisComponent, "workreports.vas", "WorkReports", "Budget", "WorkReports result"
End Sub

Sub RunProgress()
    RunExample ThisComponent, "progress.vas", "Progress", "", "Progress result"
End Sub

Sub RunPriceRevision()
    RunExample ThisComponent, "price-revision.vas", "PriceIndices", "", "PriceIndices result"
End Sub

Sub RunCleanup()
    RunExample ThisComponent, "cleanup.vas", "Cleanup", "", "Cleanup result"
End Sub

Sub RunAll()
    RunBudget
    RunWorkReports
    RunProgress
    RunPriceRevision
    RunCleanup
End Sub

' Runs script on the table of sheetName (and of extraSheet, as $2, when given); its rows go to outName
Sub RunExample(doc As Object, script As String, sheetName As String, extraSheet As String, outName As String)
    On Error GoTo Failed
    GlobalScope.BasicLibraries.LoadLibrary("ASF")
    Dim sep As String: sep = GetPathSeparator()
    Dim folder As String: folder = DocumentFolder(doc) & sep & "scripts"
    Dim eng As Object: eng = ASF.LOhelper.NewASF()
    eng.WorkingDir = folder
    Dim result As Variant
    If extraSheet = "" Then
        result = eng.Execute(folder & sep & script, ReadTable(doc, sheetName))
    Else
        result = eng.Execute(folder & sep & script, ReadTable(doc, sheetName), ReadTable(doc, extraSheet))
    End If
    WriteTable doc, outName, result
    Exit Sub
Failed:
    WriteTable doc, outName, "error " & Err & " in line " & Erl & ": " & Error$
End Sub

Function DocumentFolder(doc As Object) As String
    Dim path As String: path = ConvertFromURL(doc.getURL())
    Dim i As Long: i = Len(path)
    Do While i > 0 And Mid(path, i, 1) <> GetPathSeparator()
        i = i - 1
    Loop
    DocumentFolder = Left(path, i - 1)
End Function

' The table around A1 as a 2D array from (1, 1), as Excel's Range.Value2 gives it
Function ReadTable(doc As Object, sheetName As String) As Variant
    Dim sheet As Object: sheet = doc.Sheets.getByName(sheetName)
    Dim cursor As Object: cursor = sheet.createCursorByRange(sheet.getCellRangeByName("A1"))
    cursor.collapseToCurrentRegion()
    Dim rows As Variant: rows = cursor.getDataArray()
    Dim t() As Variant, i As Long, j As Long, nCols As Long
    nCols = UBound(rows(0)) + 1
    ReDim t(1 To UBound(rows) + 1, 1 To nCols)
    For i = 0 To UBound(rows)
        For j = 0 To nCols - 1
            ' blank cells come as "", which is what the scripts expect
            t(i + 1, j + 1) = rows(i)(j)
        Next j
    Next i
    ReadTable = t
End Function

' Writes the script's rows (an array of arrays) from A1 of sheetName, created if needed
Sub WriteTable(doc As Object, sheetName As String, rows As Variant)
    Dim sheets As Object: sheets = doc.Sheets
    If Not sheets.hasByName(sheetName) Then sheets.insertNewByName(sheetName, sheets.Count)
    Dim sheet As Object: sheet = sheets.getByName(sheetName)
    ' values, dates, strings and formulas
    sheet.clearContents(23)
    If Not IsArray(rows) Then
        sheet.getCellByPosition(0, 0).setString("Script error: " & rows)
        Exit Sub
    End If
    Dim n As Long, w As Long, i As Long
    n = UBound(rows) - LBound(rows) + 1
    For i = LBound(rows) To UBound(rows)
        If UBound(rows(i)) - LBound(rows(i)) + 1 > w Then w = UBound(rows(i)) - LBound(rows(i)) + 1
    Next i
    Dim data() As Variant
    ReDim data(0 To n - 1)
    For i = 0 To n - 1
        data(i) = DataRow(rows(LBound(rows) + i), w)
    Next i
    sheet.getCellRangeByPosition(0, 0, w - 1, n - 1).setDataArray(data)
    sheet.getColumns().OptimalWidth = True
End Sub

' One row for setDataArray: w cells from 0, blanks as ""
Function DataRow(r As Variant, w As Long) As Variant
    Dim line() As Variant, j As Long
    ReDim line(0 To w - 1)
    For j = 0 To w - 1
        line(j) = ""
        If LBound(r) + j <= UBound(r) Then
            If Not IsEmpty(r(LBound(r) + j)) Then line(j) = r(LBound(r) + j)
        End If
    Next j
    DataRow = line
End Function
