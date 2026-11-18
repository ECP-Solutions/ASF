' The code of the add-in's ThisWorkbook module, which build.ps1 adds. The descriptions of the
' worksheet functions are set again on every opening: Excel does not keep those of their arguments
' in the file. (No Option Explicit: the module of a new workbook may have it already.)

Private Sub Workbook_Open()
    ASFRegister.RegisterFunctions
End Sub
