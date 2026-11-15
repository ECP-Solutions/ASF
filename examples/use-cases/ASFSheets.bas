Attribute VB_Name = "ASFSheets"
' What ASF scripts and sheets exchange: Range.Value2 goes to a script as it is (ASF reads a 2D
' array as an array of rows); a script's array of rows comes back through ToGrid.
Option Explicit

' An ASF array of rows as a 2D array (1 To rows, 1 To columns) for Range.Value2; short rows are
' padded with empty cells
Public Function ToGrid(ByVal rows As Variant) As Variant
    Dim r As Long, c As Long, n As Long, m As Long, row As Variant, grid() As Variant
    n = UBound(rows) - LBound(rows) + 1
    For r = LBound(rows) To UBound(rows)
        If UBound(rows(r)) - LBound(rows(r)) + 1 > m Then m = UBound(rows(r)) - LBound(rows(r)) + 1
    Next r
    ReDim grid(1 To n, 1 To m)
    For r = 1 To n
        row = rows(LBound(rows) + r - 1)
        For c = 1 To UBound(row) - LBound(row) + 1
            grid(r, c) = row(LBound(row) + c - 1)
        Next c
    Next r
    ToGrid = grid
End Function

' Writes a script's array of rows from target down and right, in one assignment
Public Sub WriteRows(ByVal target As Object, ByVal rows As Variant)
    Dim grid As Variant: grid = ToGrid(rows)
    target.Resize(UBound(grid, 1), UBound(grid, 2)).Value2 = grid
End Sub
