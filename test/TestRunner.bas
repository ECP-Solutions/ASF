Attribute VB_Name = "TestRunner"
Option Explicit
Private expected As Variant
Private actual As Variant
Private scriptEngine As ASF

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Function GetResult(script As String, Optional verbose As Boolean = False) As Variant
    On Error Resume Next
    Dim idx As Long
    Set scriptEngine = New ASF
    
    With scriptEngine
        scriptEngine.verbose = verbose
        idx = .Compile(script)
        .Run idx
        GetResult = .OUTPUT_
    End With
End Function

'@TestMethod("arith_simple")
Private Sub arith_simple()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(1 + 2 * 3);"))
    expected = "7"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("arith_precedence")
Private Sub arith_precedence()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(1 + 2 * 3 / 4^2);"))
    expected = "1.375"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("paren_grouping")
Private Sub paren_grouping()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return((1 + 2) * 3);"))
    expected = "9"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("negation_unary")
Private Sub negation_unary()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "print(-5 + 3, !false, !true);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:-2, True, False"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("power_right_assoc")
Private Sub power_right_assoc()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(2 ^ 3 ^ 2);"))
    expected = "512"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("shortc_and")
Private Sub shortc_and()
    On Error GoTo TestFail
    
    actual = CBool(GetResult("x = false; return(x && (1/0));"))
    expected = CBool("False")
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("shortc_or")
Private Sub shortc_or()
    On Error GoTo TestFail
    
    actual = CBool(GetResult("x = true; return(x || (1/0));"))
    expected = CBool("True")
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("ternary_operator")
Private Sub ternary_operator()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return( 1 < 2 ? 'yes' : 'no' )"))
    expected = "yes"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("left_shift")
Private Sub left_shift()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(5<<1)"))
    expected = "10"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("right_shift")
Private Sub right_shift()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("x=-100; x>>=5; return(x)"))
    expected = "-4"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("compound_assignment_plus_equals")
Private Sub compound_assignment_plus_equals()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=2; a += 3; return(a);"))
    expected = "5"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("if_chain_same_line")
Private Sub if_chain_same_line()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=2; if (a==1) { print('one') } elseif (a==2) { print('two') } else { print('other') }; print('done');", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:'two', PRINT:'done'"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("if_multiline")
Private Sub if_multiline()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=3;" & _
                    "if (a==1) {" & _
                    "  print('one')" & _
                    "} elseif (a==2) {" & _
                    "  print('two')" & _
                    "} elseif (a==3) {" & _
                    "  print('three')" & _
                    "} else {" & _
                    "  print('other')" & _
                    "};" & _
                    "print('end');", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:'three', PRINT:'end'"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("for_simple")
Private Sub for_simple()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("s=0; for(i=1,i<=3,i=i+1) { s = s + i }; return(s);"))
    expected = "6"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("for_break_continue")
Private Sub for_break_continue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("s=0; for(i=1,i<=5,i=i+1) {" & _
                                    "if (i==3) { continue }" & _
                                    "if (i==5) { break } s = s + i };" & _
                                    "return(s);"))
    expected = "7"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("while_break_continue")
Private Sub while_break_continue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("i=1; s=0; while (i <= 5) {" & _
                                        "if (i==2) { i = i + 1 ; continue }" & _
                                        "if (i==5) { break }" & _
                                        "s = s + i ; i = i + 1 };" & _
                                        "return(s);"))
    expected = "8"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("switch_case")
Private Sub switch_case()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("c='blue'; switch(c) {" & _
                                        "case 'red' { return('warm') }" & _
                                        "case 'blue' { return('cool') }" & _
                                        "default { return('other') } }"))
    expected = "cool"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("try_catch")
Private Sub try_catch()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("try { x = 1/0 }" & _
                            "catch { return('caught') }"))
    expected = "caught"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("function_basic")
Private Sub function_basic()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun add(a,b) { return a + b }; return(add(2,3));"))
    expected = "5"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("function_scope_isolation")
Private Sub function_scope_isolation()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=5; fun f(a) { a = a + 1 ; print(a) } ; f(a); print(a);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:6, PRINT:5"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("recursion_fib_arrays")
Private Sub recursion_fib_arrays()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "fun fib(n) {" & _
                            " if (n <= 2) { return 1 }; return fib(n-1) + fib(n-2)" & _
                            "} ;" & _
                            "a = [];" & _
                            "for(i=1,i<=6,i=i+1) {" & _
                                "a[i] = fib(i)" & _
                            "};" & _
                            "print(a[1]); print(a[6]);" & _
                            "print(a);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:1, PRINT:8, PRINT:[ 1, 1, 2, 3, 5, 8 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("recursion_fib_single")
Private Sub recursion_fib_single()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun fib(n) {" & _
                                " if (n <= 2) { return 1 }; return fib(n-1) + fib(n-2)" & _
                            "} ;" & _
                            "return(fib(15));"))
    expected = "610"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("closure_shared_write")
Private Sub closure_shared_write()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = 1; f = fun() { a = a + 1; return a }; print(f()); print(a);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:2, PRINT:2"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("closure_multiple_instances")
Private Sub closure_multiple_instances()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = 0; fun make() { return fun() { a = a + 1 ; return a } };" & _
                "f1 = make(); f2 = make();" & _
                "print(f1()); print(f2()); print(a);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:1, PRINT:2, PRINT:2"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("array_literal_and_length")
Private Sub array_literal_and_length()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[10,20,30]; print(a[2]); print(a.length);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:20, PRINT:3"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("array_of_arrays_length")
Private Sub array_of_arrays_length()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [] ; a[1] = [7,8] ; a[3] = [9,10,11] ;" & _
                "print(a[1]); print(a[3]); print(a[3].length)", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 7, 8 ], PRINT:[ 9, 10, 11 ], PRINT:3"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("object_literal_and_member")
Private Sub object_literal_and_member()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "o = { x: 10, y: 'hi' } ; print(o.x) ; o.x = o.x + 5 ; print(o.x)", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:10, PRINT:15"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("nested_member_index_LValue")
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

'@TestMethod("method_call_on_member")
Private Sub method_call_on_member()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("o = { v: 10, incr: fun(x) { return x + 1 } } ; return(o.incr(o.v))"))
    expected = "11"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("anon_func_as_arg")
Private Sub anon_func_as_arg()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun apply(f,x) { return f(x) } ; return(apply(fun(y) { return y * 2 }, 5))"))
    expected = "10"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("anon_func_closure_arg")
Private Sub anon_func_closure_arg()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a = 5; fun apply(f) { return f() }; return(apply(fun() { return a + 1 }))"))
    expected = "6"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("vbexpr_embedded")
Private Sub vbexpr_embedded()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a = @({1;0;4});" & _
                            " b = @({1;1;6});" & _
                            " c = @({-3;0;-10});" & _
                            " d = @({2;3;4});" & _
                            " return(@(MROUND(LUDECOMP(ARRAY(a;b;c));4)))"))
    expected = "{{-3;0;-10};{-0.3333;1;2.6667};{-0.3333;0;0.6667}}"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("calling_native_function")
Sub calling_native_function()
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

'@TestMethod("map_nested_array")
Private Sub map_nested_array()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [1,[2,[3,[4]]]]; b = a.map(fun(x) { return x * 10 }); print(b);)", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 10, [ 20, [ 30, [ 40 ] ] ] ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("map_array_of_objects")
Private Sub map_array_of_objects()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [{ k: 1, arr: [10,20] }, { k: 2, arr: [30,[40,50]] }];" & _
                        "b = a.map(fun(o){return { k: o.k * 2, arr: o.arr.map(fun(x){ return x + 1 })}; });" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ { k: 2, arr: [ 11, 21 ] }, { k: 4, arr: [ 31, [ 41, 51 ] ] } ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("map_closure_capture")
Private Sub map_closure_capture()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "mul = fun(factor){return fun(x){ return x * factor };};" & _
                        "a = [1,2,3]; b = a.map(mul(5));" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 5, 10, 15 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("map_returning_nested_array")
Private Sub map_returning_nested_array()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "print( [1,2].map(fun(x){ return [x,x] }) );", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ [ 1, 1 ], [ 2, 2 ] ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("map_returning_objects_and_arrays")
Private Sub map_returning_objects_and_arrays()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [1,2];" & _
                        "b = a.map(fun(n){return {orig: n,pair: [n, n*n],nested: [ [n, n+1], { v: n*n } ]};});" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ { orig: 1, pair: [ 1, 1 ], nested: [ [ 1, 2 ], { v: 1 } ] }, { orig: 2, pair: [ 2, 4 ], nested: [ [ 2, 3 ], { v: 4 } ] } ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("mapping_mixed_types")
Private Sub mapping_mixed_types()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [1,'x',[2,'y',[3]]];" & _
                        "b = a.map(fun(x){if (IsArray(x)) {return x} elseif (IsNumeric(x)) {return x*3} else {return x}};);" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 3, 'x', [ 6, 'y', [ 9 ] ] ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("filter_simple")
Private Sub filter_simple()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a = [1,2,3,4];" & _
                        "b = a.filter(fun(x){ return x % 2 == 0 });" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 2, 4 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("filter_nested_arrays")
Private Sub filter_nested_arrays()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[1,[2,3],4,[5]];" & _
                        "b=a.filter(fun(x){ return IsArray(x) });" & _
                        "print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ [ 2, 3 ], [ 5 ] ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("reduce_with_initial")
Private Sub reduce_with_initial()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3,4]; return(a.reduce(fun(acc,x){ return acc + x }, 0));"))
    expected = "10"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("reduce_with_NO_initial")
Private Sub reduce_with_NO_initial()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3]; return(a.reduce(fun(acc,x){ return acc + x }));"))
    expected = "6"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("slice_tail_only")
Private Sub slice_tail_only()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[10,20,30,40]; b=a.slice(2); print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 20, 30, 40 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("slice_start_end")
Private Sub slice_start_end()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=['ant', 'bison', 'camel', 'duck', 'elephant']; b=a.slice(3,5); print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 'camel', 'duck' ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("pop_push")
Private Sub pop_push()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[1,2]; a.push(3); a.push(4); x = a.pop(); print(a); print(x);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], 4"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("range_default")
Private Sub range_default()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "print(range(3));", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 0, 1, 2 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("range_custom")
Private Sub range_custom()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "print(range(1,3));", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 2 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("range_with_step")
Private Sub range_with_step()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "print(range(1,10,2));", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 3, 5, 7, 9 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("flatten_full")
Private Sub flatten_full()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[1,[2,3],[4,[5]]]; b = flatten(a); print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 2, 3, 4, 5 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("flatten_depth_one")
Private Sub flatten_depth_one()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=[1,[2,[3]]]; b = flatten(a,1); print(b);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 2, [ 3 ] ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("clone_array")
Private Sub clone_array()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "o = { x: 1, a: [1,2] }; c = clone(o); c.a.push(3); print(o.a); print(c.a);", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:[ 1, 2 ], PRINT:[ 1, 2, 3 ]"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

'@TestMethod("filter_reduce_chain")
Private Sub filter_reduce_chain()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3,4,5]; return(a.filter(fun(x){ return x > 2 }).reduce(fun(acc,x){ return acc + x }, 0));"))
    expected = "12"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
