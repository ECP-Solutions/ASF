VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AsfFakeHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A host object for TestRunner: what an object a host trusts looks like to a script, as the Python
' add-in's py does (InjectHostObject). Twice is a method, Child an object it hands out, Label a
' property. Import it into the test workbook next to TestRunner.bas.
Option Explicit

Public Function Twice(ByVal x As Variant) As Variant
    Twice = x * 2
End Function

Public Function Child() As AsfFakeHost
    Set Child = New AsfFakeHost
End Function

Public Property Get Label() As String
    Label = "fake host"
End Property
