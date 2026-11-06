Attribute VB_Name = "OfficeHost"
' OfficeHost: runs the portable ASF examples in Excel, Word or PowerPoint. The same module works in
' the three: import it, with the ASF class modules (src/*.cls), into a workbook, document or
' presentation saved in examples/portable, next to the scripts folder, and run one of the macros.
' LibreOfficeHost.bas does the same in Calc, Writer and Impress: the scripts do not change.
Option Explicit

Public Sub RunHello():         RunScript "hello.vas":          End Sub
Public Sub RunBudget():        RunScript "budget.vas":         End Sub
Public Sub RunProgress():      RunScript "progress.vas":       End Sub
Public Sub RunPriceRevision(): RunScript "price-revision.vas": End Sub
Public Sub RunWorkReports():   RunScript "workreports.vas":    End Sub
Public Sub RunCleanup():       RunScript "cleanup.vas":        End Sub

Public Sub RunAll()
    Dim s As Variant, msg As String
    For Each s In Array("hello.vas", "budget.vas", "progress.vas", "price-revision.vas", "workreports.vas", "cleanup.vas")
        msg = msg & RunScriptOn(ActiveDoc(), CStr(s)) & vbLf
    Next s
    MsgBox msg, vbInformation, "ASF"
End Sub

Public Sub RunScript(ByVal script As String)
    MsgBox RunScriptOn(ActiveDoc(), script), vbInformation, "ASF"
End Sub

' Runs scripts\<script> on doc and returns the script's result, or its error
Public Function RunScriptOn(ByVal doc As Object, ByVal script As String) As String
    Dim eng As ASF: Set eng = New ASF
    Dim folder As String: folder = eng.JoinPath(doc.Path, "scripts")
    eng.AppAccess = True
    eng.WorkingDir = folder
    Dim result As Variant
    result = eng.Execute(eng.JoinPath(folder, script), doc)
    If IsEmpty(result) Then
        RunScriptOn = script & " stopped: " & CStr(eng.LastError)
    Else
        RunScriptOn = CStr(result)
    End If
End Function

' The open workbook, document or presentation, whichever application this is
Private Function ActiveDoc() As Object
    Select Case Application.Name
        Case "Microsoft Excel": Set ActiveDoc = CallByName(Application, "ActiveWorkbook", VbGet)
        Case "Microsoft Word": Set ActiveDoc = CallByName(Application, "ActiveDocument", VbGet)
        Case "Microsoft PowerPoint": Set ActiveDoc = CallByName(Application, "ActivePresentation", VbGet)
    End Select
End Function
