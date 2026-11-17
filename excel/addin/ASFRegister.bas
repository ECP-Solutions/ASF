Attribute VB_Name = "ASFRegister"
' The descriptions of the worksheet functions in Excel's Insert Function dialog, in the category
' ASF. The build (excel/tools/build.ps1) runs RegisterFunctions before saving the add-in, which keeps
' them.
Option Explicit
Option Private Module

Public Sub RegisterFunctions()
    Describe "ASF_RUN", "Runs an ASF script. Its arguments are $1, $2...: a range arrives as an array of rows. " & _
             "An array of rows comes back as a grid, a flat array as a column.", _
             Array("The script, or the cell that holds it", "A value or a range: $1, $2...")
    Describe "ASF_CALL", "Runs a script of the ASF library folder (name.vas). Its arguments are $1, $2...", _
             Array("The name of the script, without folder", "A value or a range: $1, $2...")
    Describe "RX_TEST", "TRUE when the text matches the regular expression (JavaScript syntax).", _
             Array("The text", "The pattern", "TRUE to ignore case")
    Describe "RX_EXTRACT", "The first match of the regular expression, or one of its groups.", _
             Array("The text", "The pattern", "The group: 0 for the whole match", "TRUE to ignore case")
    Describe "RX_REPLACE", "The text with every match replaced ($1, $2... for the groups).", _
             Array("The text", "The pattern", "The replacement", "TRUE to ignore case")
    Describe "RX_COUNT", "The number of matches of the regular expression.", _
             Array("The text", "The pattern", "TRUE to ignore case")
End Sub

Private Sub Describe(ByVal fn As String, ByVal description As String, ByVal arguments As Variant)
    Application.MacroOptions Macro:=fn, Description:=description, Category:="ASF", ArgumentDescriptions:=arguments
End Sub
