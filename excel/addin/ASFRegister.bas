Attribute VB_Name = "ASFRegister"
' The worksheet functions as Excel shows them: one table (FunctionInfo) of names, descriptions and
' arguments, used for
'   - Insert Function and the Function Arguments dialog (Ctrl+A after the name), with the category
'     ASF: RegisterFunctions, run by the build and again on every opening (ThisWorkbook), as Excel
'     does not keep the descriptions of the arguments in the file;
'   - the tips while typing a formula, as Excel shows them for its own functions, which needs the
'     free Excel-DNA IntelliSense add-in: it reads the sheet _IntelliSense_ that WriteIntelliSenseSheet
'     makes at the build.
' An optional argument is written [name], the arguments that repeat [args]...
Option Explicit
Option Private Module

Private Const CATEGORY As String = "ASF"
Private Const HELP_URL As String = "https://github.com/ECP-Solutions/ASF/blob/main/excel/README.md#worksheet-functions"

' Array(name, description, argument, its description, argument, its description...) per function
Public Function FunctionInfo() As Variant
    Const ARGS_TIP As String = "Values or ranges, read by the script as $1, $2...: a range is an array of rows, one cell its value"
    Const PATTERN_TIP As String = "The regular expression, in JavaScript syntax: \d, \b, (?<name>...), lookbehind, lazy quantifiers"
    Const CASE_TIP As String = "TRUE to ignore case; FALSE when omitted"
    FunctionInfo = Array( _
        Array("ASF_RUN", "Runs an ASF script with the arguments as $1, $2... An array of rows comes back as a grid, a flat array as a column; an error gives #VALUE!.", _
              "script", "The script, or the cell that holds it (text in a formula holds at most 255 characters)", _
              "[args]...", ARGS_TIP), _
        Array("ASF_CALL", "Runs a script of the ASF library folder with the arguments as $1, $2... An error gives #VALUE!.", _
              "name", "The name of the script: a .vas file of the library folder, with or without .vas", _
              "[args]...", ARGS_TIP), _
        Array("RX_TEST", "TRUE when the text matches the regular expression.", _
              "text", "The text to test", _
              "pattern", PATTERN_TIP, _
              "[ignoreCase]", CASE_TIP), _
        Array("RX_EXTRACT", "The first match of the regular expression in the text, or one of its groups; empty when there is none.", _
              "text", "The text to search", _
              "pattern", PATTERN_TIP, _
              "[group]", "0 or omitted for the whole match; 1, 2... for a group", _
              "[ignoreCase]", CASE_TIP), _
        Array("RX_REPLACE", "The text with every match of the regular expression replaced.", _
              "text", "The text to change", _
              "pattern", PATTERN_TIP, _
              "replacement", "What replaces each match: $1, $2... insert the groups", _
              "[ignoreCase]", CASE_TIP), _
        Array("RX_COUNT", "The number of matches of the regular expression in the text.", _
              "text", "The text to search", _
              "pattern", PATTERN_TIP, _
              "[ignoreCase]", CASE_TIP))
End Function

' The descriptions in Insert Function and the Function Arguments dialog
Public Sub RegisterFunctions()
    Dim f As Variant
    For Each f In FunctionInfo()
        Describe f
    Next f
End Sub

' The sheet Excel-DNA IntelliSense reads from a VBA add-in: A1 FunctionInfo and B1 the version 1.0,
' then one row per function: name, description, help topic, and the name and description of each
' argument. The first sheet of the add-in becomes it; the build runs this before saving
Public Sub WriteIntelliSenseSheet()
    Dim sh As Object, f As Variant, r As Long, i As Long, c As Long
    Set sh = ThisWorkbook.Worksheets(1)
    sh.Name = "_IntelliSense_"
    sh.Cells.Clear
    sh.Cells(1, 1).Value = "FunctionInfo"
    sh.Cells(1, 2).Value = "'1.0"
    r = 1
    For Each f In FunctionInfo()
        r = r + 1
        sh.Cells(r, 1).Value = f(0)
        sh.Cells(r, 2).Value = f(1)
        sh.Cells(r, 3).Value = HELP_URL
        c = 4
        For i = 2 To UBound(f)
            sh.Cells(r, c).Value = f(i)
            c = c + 1
        Next i
    Next f
End Sub

Private Sub Describe(ByVal f As Variant)
    Dim descriptions() As String, n As Long, i As Long
    n = (UBound(f) - 1) \ 2
    ReDim descriptions(0 To n - 1)
    For i = 0 To n - 1
        descriptions(i) = f(3 + 2 * i)
    Next i
    On Error Resume Next
    Application.MacroOptions Macro:=f(0), Description:=f(1), Category:=CATEGORY, ArgumentDescriptions:=descriptions
    If Err.Number <> 0 Then
        ' a version without ArgumentDescriptions (older Excel for Mac): the description alone
        Err.Clear
        Application.MacroOptions Macro:=f(0), Description:=f(1), Category:=CATEGORY
    End If
End Sub
