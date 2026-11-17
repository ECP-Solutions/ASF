Attribute VB_Name = "ASFFunctions"
' The worksheet functions of the ASF add-in, and the engine for the macros of other workbooks.
'   =ASF_RUN("return $1 * 1.16;", B2)          a script in the formula, or in a cell: ASF_RUN(A1, B2)
'   =ASF_CALL("discounts", A1:F200, H1:J20)    discounts.vas, from the library folder (ASFSettings)
' The arguments are $1, $2...; ASFCells says how values go in and come out. A compile error, an
' error or throw that no catch takes, or the step limit gives #VALUE!.
' Scripts run without access to Office objects (AppAccess off), import only from the library
' folder (ImportRoot), and stop after StepLimit loop iterations and calls (MaxSteps).
Option Explicit

' Runs the script, with the arguments as $1, $2...
Public Function ASF_RUN(ByVal script As String, ParamArray args() As Variant) As Variant
    Dim a As Variant: a = args
    ASF_RUN = ASFEngine.RunScript(script, a)
End Function

' Runs name.vas from the library folder, with the arguments as $1, $2...
Public Function ASF_CALL(ByVal name As String, ParamArray args() As Variant) As Variant
    Dim a As Variant: a = args
    ASF_CALL = ASFEngine.RunLibraryScript(name, a)
End Function

' A new engine for the macros of other workbooks, which reach it through a reference to the add-in
' (Tools > References > ASFAddIn):  Dim e As ASF: Set e = ASFAddIn.NewEngine()
Public Function NewEngine() As ASF
    Set NewEngine = New ASF
End Function
