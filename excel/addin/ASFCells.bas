Attribute VB_Name = "ASFCells"
' What cells and ASF scripts exchange, for the worksheet functions of the add-in.
'   Into a script: a cell gives its value; one row or one column of cells a flat array; a range of
'   several rows and columns an array of rows; errors inside a range become their text (#N/A...),
'   empty cells undefined.
'   Out of a script: an array of rows becomes a grid, a flat array a column, objects and nested
'   arrays their text; undefined and null an empty string (a function returning Empty shows 0).
Option Explicit
Option Private Module

' Excel's error values: CVErr(XL_ERR_VALUE) is #VALUE!
Public Const XL_ERR_NULL As Long = 2000
Public Const XL_ERR_DIV0 As Long = 2007
Public Const XL_ERR_VALUE As Long = 2015
Public Const XL_ERR_REF As Long = 2023
Public Const XL_ERR_NAME As Long = 2029
Public Const XL_ERR_NUM As Long = 2036
Public Const XL_ERR_NA As Long = 2042

' An argument of a worksheet function as a script takes it: a Range as its values, an omitted one
' as Empty. An error value is returned as is: the caller gives it back without running the script
Public Function ArgumentValue(ByRef arg As Variant) As Variant
    If IsError(arg) Then
        ' an omitted argument is an error value too (Missing): only Excel's errors are kept
        If ErrorText(arg) = "#ERROR" Then ArgumentValue = Empty Else ArgumentValue = arg
    ElseIf IsMissing(arg) Then
        ArgumentValue = Empty
    ElseIf IsObject(arg) Then
        Dim v As Variant: v = RangeValue(arg)
        If IsArray(v) Then ArgumentValue = CellValues(v) Else ArgumentValue = v
    ElseIf IsArray(arg) Then
        ArgumentValue = CellValues(arg)
    Else
        ArgumentValue = arg
    End If
End Function

' The values of a Range, the only object a cell formula passes (a procedure of its own: LibreOffice
' raises on arg.Value2 even in a branch that does not run)
Private Function RangeValue(ByVal cells As Object) As Variant
    RangeValue = cells.Value2
End Function

' The values of an array of cells, error values replaced by their text: one row or one column as a
' flat array from 1 (B2:B6 is a list, as in Excel), more as a 2D array, which the engine reads as an
' array of rows
Private Function CellValues(ByVal cells As Variant) As Variant
    Dim r As Long, c As Long
    If TwoDimensions(cells) Then
        For r = LBound(cells, 1) To UBound(cells, 1)
            For c = LBound(cells, 2) To UBound(cells, 2)
                If IsError(cells(r, c)) Then cells(r, c) = ErrorText(cells(r, c))
            Next c
        Next r
        If LBound(cells, 1) = UBound(cells, 1) Or LBound(cells, 2) = UBound(cells, 2) Then
            CellValues = Flat(cells)
            Exit Function
        End If
    Else
        ' an array constant as {1,2,3} has one dimension
        For r = LBound(cells) To UBound(cells)
            If IsError(cells(r)) Then cells(r) = ErrorText(cells(r))
        Next r
    End If
    CellValues = cells
End Function

' The cells of one row or one column, as an array from 1
Private Function Flat(ByRef cells As Variant) As Variant
    Dim items() As Variant, r As Long, c As Long, n As Long
    ReDim items(1 To (UBound(cells, 1) - LBound(cells, 1) + 1) * (UBound(cells, 2) - LBound(cells, 2) + 1))
    For r = LBound(cells, 1) To UBound(cells, 1)
        For c = LBound(cells, 2) To UBound(cells, 2)
            n = n + 1
            items(n) = cells(r, c)
        Next c
    Next r
    Flat = items
End Function

Private Function TwoDimensions(ByRef arr As Variant) As Boolean
    Dim lb As Long
    On Error GoTo OneOnly
    lb = LBound(arr, 2)
    TwoDimensions = True
OneOnly:
End Function

Public Function ErrorText(ByVal e As Variant) As String
    Select Case True
        Case e = CVErr(XL_ERR_NA): ErrorText = "#N/A"
        Case e = CVErr(XL_ERR_VALUE): ErrorText = "#VALUE!"
        Case e = CVErr(XL_ERR_DIV0): ErrorText = "#DIV/0!"
        Case e = CVErr(XL_ERR_REF): ErrorText = "#REF!"
        Case e = CVErr(XL_ERR_NAME): ErrorText = "#NAME?"
        Case e = CVErr(XL_ERR_NUM): ErrorText = "#NUM!"
        Case e = CVErr(XL_ERR_NULL): ErrorText = "#NULL!"
        Case Else: ErrorText = "#ERROR"
    End Select
End Function

' A script's result for cells: a value, or a 2D array (1 To rows, 1 To columns)
Public Function CellResult(ByRef v As Variant) As Variant
    If IsObject(v) Then
        CellResult = ASFEngine.ValueText(v)
    ElseIf IsArray(v) Then
        CellResult = Grid(v)
    Else
        CellResult = CellValue(v)
    End If
End Function

' A value for one cell
Private Function CellValue(ByRef v As Variant) As Variant
    If IsObject(v) Then
        CellValue = ASFEngine.ValueText(v)
    ElseIf IsArray(v) Then
        CellValue = ASFEngine.ValueText(v)
    ElseIf IsEmpty(v) Or IsNull(v) Then
        CellValue = vbNullString
    Else
        CellValue = v
    End If
End Function

' An ASF array as a 2D array: an array of rows as a grid (short rows padded with ""), any other
' array as one column; an empty array as ""
Private Function Grid(ByRef arr As Variant) As Variant
    Dim n As Long: n = ItemCount(arr)
    If n = 0 Then
        Grid = vbNullString
        Exit Function
    End If
    Dim lb As Long: lb = LBound(arr)
    Dim r As Long, c As Long, m As Long, rows As Boolean, item As Variant
    ' an array of rows: every item an array
    rows = True
    For r = 0 To n - 1
        If IsObject(arr(lb + r)) Then
            rows = False
        ElseIf Not IsArray(arr(lb + r)) Then
            rows = False
        ElseIf ItemCount(arr(lb + r)) > m Then
            m = ItemCount(arr(lb + r))
        End If
    Next r
    If m = 0 Then rows = False
    Dim g() As Variant
    If rows Then
        ReDim g(1 To n, 1 To m)
        For r = 1 To n
            item = arr(lb + r - 1)
            For c = 1 To m
                If c <= ItemCount(item) Then g(r, c) = CellValue(item(LBound(item) + c - 1)) Else g(r, c) = vbNullString
            Next c
        Next r
    Else
        ReDim g(1 To n, 1 To 1)
        For r = 1 To n
            g(r, 1) = CellValue(arr(lb + r - 1))
        Next r
    End If
    Grid = g
End Function

' The items of an ASF array: an empty one is Array() or a (-1 To -1) array
Private Function ItemCount(ByRef arr As Variant) As Long
    If UBound(arr) < LBound(arr) Or UBound(arr) = -1 Then Exit Function
    ItemCount = UBound(arr) - LBound(arr) + 1
End Function

' A result fitted to the cells of an array formula (rows x columns): cut, or padded with ""; a 1x1
' result as its value. rows = 0: the formula is in one cell, the result stays as it is (it spills)
Public Function FitTo(ByRef result As Variant, ByVal rows As Long, ByVal columns As Long) As Variant
    Dim r As Long, c As Long, n As Long, m As Long
    If IsArray(result) Then
        n = UBound(result, 1): m = UBound(result, 2)
    Else
        n = 1: m = 1
    End If
    If rows <= 1 And columns <= 1 Then
        If n = 1 And m = 1 And IsArray(result) Then FitTo = result(1, 1) Else FitTo = result
        Exit Function
    End If
    Dim g() As Variant
    ReDim g(1 To rows, 1 To columns)
    For r = 1 To rows
        For c = 1 To columns
            If r > n Or c > m Then
                g(r, c) = vbNullString
            ElseIf IsArray(result) Then
                g(r, c) = result(r, c)
            Else
                g(r, c) = result
            End If
        Next c
    Next r
    FitTo = g
End Function
