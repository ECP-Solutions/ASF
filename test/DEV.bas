Attribute VB_Name = "DEV"
Option Explicit
Private expected As Variant
Private actual As Variant
Private scriptEngine As ASF
Private Assert As Object
Private Fakes As Object
Private Function GetResult(script As String, Optional verbose As Boolean = False) As Variant
    On Error Resume Next
    Dim idx As Long
    Set scriptEngine = New ASF
    
    With scriptEngine
        .verbose = verbose
        .EnableCallTrace = True
        idx = .Compile(script)
        .Run idx
        GetResult = .OUTPUT_
    End With
End Function

Private Sub test()
    Dim scriptEngine As New ASF
    Dim progIdx As Long
    Dim tmpResult As Variant
    With scriptEngine
        .AppAccess = True
        .OverrideCollMethods = True
        .verbose = True
        .EnableCallTrace = True
        progIdx = .Compile("  prototype.COM.ListRow asDictionary() {" & _
                            "    let headers = this.parent.listcolumns;" & _
                            "    let values = this.range.value2;" & _
                            "    let result = {};" & _
                            "    for (let i = 1, i <= headers.count, i+=1) {" & _
                            "        result.set(headers.item(i).name, values[1][i]);" & _
                            "    };" & _
                            "    return result;" & _
                            " };" & _
                            " let myData = $1.Sheets(1).ListObjects('Demo_Table').ListRows.map(" & _
                            "     fun(col) {" & _
                            "         return col.asDictionary().Get('ip_address');" & _
                            "      }" & _
                            " ); print(myData[5])")
        .Run progIdx, ThisWorkbook
'        ThisWorkbook.Sheets(1).Range("H14").value2 = .GetCallStackTrace
    End With
End Sub
'@TestMethod("calling_native")
Private Sub calling_native_function()
    On Error GoTo TestFail
    Dim asfGlobals As New ASF_Globals
    Dim progIdx  As Long
    
    With asfGlobals
        .ASF_InitGlobals
        .gExprEvaluator.DeclareUDF "ThisWBname", "UserDefFunctions"
    End With
    Set scriptEngine = New ASF
    With scriptEngine
        .SetGlobals asfGlobals
        .EnableCallTrace = True
        .verbose = True
        progIdx = .Compile("/*Get Thisworkbook name*/ return(@(ThisWBname()))")
        .Run progIdx
        actual = CStr(.OUTPUT_)
    End With
    expected = ThisWorkbook.name
    Assert.AreEqual expected, actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("host_VBA_Expressions")
Private Sub host_sees_array_mutation()
    On Error GoTo TestFail
    Dim asfGlobals As New ASF_Globals
    Dim progIdx  As Long
    
    With asfGlobals
        .ASF_InitGlobals
        .gExprEvaluator.DeclareUDF "HostCheckSecondNestedElement", "UserDefFunctions"
    End With
    Set scriptEngine = New ASF
    With scriptEngine
        .EnableCallTrace = True
        .verbose = True
        .SetGlobals asfGlobals
        progIdx = .Compile("a=@({{7;8;9}}); a[1][2]=42; return(@(HostCheckSecondNestedElement(a)));")
        .Run progIdx
        actual = CStr(.OUTPUT_)
    End With
    expected = "42"
    Assert.AreEqual expected, actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("nested_member")
Private Sub nested_member_index_LValue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult(" o = { a: [ {v:1}, {v:2} ] } ;" & _
                "o.a[2].v = o.a[2].v + 5 ; return(o.a[2].v + 2)"))
    expected = "9"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("module_system")
Private Sub module_system_prototype_imports()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .AppAccess = True
        .InjectVariable "wd", wd
        .EnableCallTrace = True
        .verbose = True
        On Error GoTo TestFail
        actual = .Execute(wd & "\main_prototype.vas", ThisWorkbook)
        Debug.Print .GetCallStackTrace
    End With
    expected = 255
    Assert.AreEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

