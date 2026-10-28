' VBAEhelper - creates VBAExpressionsLib evaluators for the ASF library.
'
' LibreOffice resolves New against the libraries loaded when a module is compiled, and
' compiles a module on its first call: this one is first called once LOhelper has
' loaded VBAExpressionsLib.
Option Explicit

Public Function NewVBAexpressions() As Object
    Set NewVBAexpressions = New VBAexpressions
End Function
