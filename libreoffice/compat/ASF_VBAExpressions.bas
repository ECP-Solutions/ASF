' LibreOffice resolves New against the libraries loaded when a module is compiled, and
' compiles a module on its first call: this one is first called once ASF_LOCompat has
' loaded VBAExpressionsLib.
Option Explicit

Public Function ASFNewVBAexpressions() As Object
    Set ASFNewVBAexpressions = New VBAexpressions
End Function
