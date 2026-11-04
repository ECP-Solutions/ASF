' LibreOfficeHost: runs the portable ASF examples in Calc, Writer or Impress. The same module works
' in the three: add it to a spreadsheet, text document or presentation saved in examples/portable,
' next to the scripts folder, with the ASF extension (ASF.oxt) installed, and run one of the macros.
' OfficeHost.bas does the same in Excel, Word and PowerPoint: the scripts do not change.
Option Explicit

Sub RunHello : RunScript "hello.vas" : End Sub
Sub RunBudget : RunScript "budget.vas" : End Sub
Sub RunProgress : RunScript "progress.vas" : End Sub
Sub RunPriceRevision : RunScript "price-revision.vas" : End Sub
Sub RunWorkReports : RunScript "workreports.vas" : End Sub
Sub RunCleanup : RunScript "cleanup.vas" : End Sub

Sub RunAll
    Dim s As Variant, msg As String
    For Each s In Array("hello.vas", "budget.vas", "progress.vas", "price-revision.vas", "workreports.vas", "cleanup.vas")
        msg = msg & RunScriptOn(ThisComponent, ScriptsFolder(ThisComponent), s) & Chr(10)
    Next s
    MsgBox msg, 64, "ASF"
End Sub

Sub RunScript(script As String)
    MsgBox RunScriptOn(ThisComponent, ScriptsFolder(ThisComponent), script), 64, "ASF"
End Sub

' Runs folder/script on doc and returns the script's result, or its error
Function RunScriptOn(doc As Object, folder As String, script As String) As String
    GlobalScope.BasicLibraries.LoadLibrary("ASF")
    Dim eng As Object: eng = ASF.LOhelper.NewASF()
    eng.AppAccess = True
    eng.WorkingDir = folder
    Dim result As Variant
    result = eng.Execute(folder & GetPathSeparator() & script, doc)
    If IsEmpty(result) Then
        RunScriptOn = script & " stopped: " & eng.LastError
    Else
        RunScriptOn = result
    End If
End Function

' The scripts folder next to the document
Function ScriptsFolder(doc As Object) As String
    Dim path As String: path = ConvertFromURL(doc.getURL())
    Dim i As Long: i = Len(path)
    Do While i > 0 And Mid(path, i, 1) <> GetPathSeparator()
        i = i - 1
    Loop
    ScriptsFolder = Left(path, i) & "scripts"
End Function
