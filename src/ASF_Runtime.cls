VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Runtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The library of ASF's values, which ASF_VM calls: methods of arrays, strings, objects, collections,
' Math and regex, built-in functions, printing, host objects and their prototypes, @(...) and the
' files of modules
Option Explicit
Private Const VTP_MAX_DEPTH As Long = 8    ' safety depth limit
Private Const VTP_MAX_ITEMS_INLINE As Long = 8 ' prefer inline for small containers
Private VERBOSE_ As Boolean
Private GLOBALS_ As ASF_Globals
Private INJECTED_VARIABLES_ As Collection
Private VM_ As Object                   ' the VM while it runs: it runs the closures it made
Private METHOD_NAME_ As String          ' the name of the method being dispatched, as written
Private METHOD_WROTE_ As Boolean        ' it changed its array, METHOD_WRITTEN_: the new array
Private METHOD_WRITTEN_ As Variant
Private ENABLE_CALL_TRACE_ As Boolean
Private CAST_INJECTED_VARIABLES_ As Boolean
Private APPLICATION_ACCESS_ As Boolean
Private tmpVar As Variant
Private COM_PROTOTYPES_ As Collection ' COM Prototype storage for monkey patching
Private COLLECTIONS_METHOD_OVERRIDE_ As Boolean
Private MathObj As Collection
Private RANDOMIZED_ As Boolean

Public Property Set Globals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Property
Public Property Get Globals() As ASF_Globals
    Set Globals = GLOBALS_
End Property
Public Sub InjectVariable(varName As String, VarValue As Variant)
    If INJECTED_VARIABLES_ Is Nothing Then
        Set INJECTED_VARIABLES_ = New Collection
    End If
    On Error Resume Next
    INJECTED_VARIABLES_.Remove (varName)
    If CAST_INJECTED_VARIABLES_ And Not APPLICATION_ACCESS_ Then
        INJECTED_VARIABLES_.Add Array(varName, CastInjectedVar(VarValue)), varName
    Else
        vAssignment tmpVar, VarValue
        If IsArray(tmpVar) Then tmpVar = ArrayToASF2(tmpVar)
        INJECTED_VARIABLES_.Add Array(varName, tmpVar), varName
    End If
End Sub
Private Sub InitializeCOMPrototypes()
    If COM_PROTOTYPES_ Is Nothing Then
        Set COM_PROTOTYPES_ = New Collection
    End If
End Sub
' Register a prototype method for a COM class type
Public Sub RegisterCOMPrototype(classType As String, methodName As String, funcIndex As Long, params As Collection, Optional ByVal closure As Variant)
    InitializeCOMPrototypes
    
    Dim classKey As String: classKey = UCase$(classType)
    Dim prototypeMethods As Collection
    
    ' Get or create methods collection for this class type
    On Error Resume Next
    Set prototypeMethods = COM_PROTOTYPES_(classKey)
    On Error GoTo 0
    
    If prototypeMethods Is Nothing Then
        Set prototypeMethods = New Collection
        COM_PROTOTYPES_.Add prototypeMethods, classKey
    End If
    
    ' Store method info
    Dim methodInfo As ASF_Map: Set methodInfo = New ASF_Map
    methodInfo.SetValue "funcIndex", funcIndex
    methodInfo.SetValue "params", params
    ' the function ASF_VM compiled, which it runs
    If Not IsMissing(closure) Then
        If IsObject(closure) Then methodInfo.SetValue "closure", closure
    End If
    
    ' Remove existing method if present
    On Error Resume Next
    prototypeMethods.Remove UCase$(methodName)
    On Error GoTo 0
    
    prototypeMethods.Add methodInfo, UCase$(methodName)
End Sub
' Check if a prototype method exists for a COM object type and method name
Private Function HasCOMPrototype(obj As Variant, methodName As String) As Boolean
    InitializeCOMPrototypes
    
    On Error GoTo NoPrototype
    
    Dim objectType As String: objectType = GetCOMObjectType(obj)
    If objectType = vbNullString Then GoTo NoPrototype
    
    Dim prototypeMethods As Collection
    Set prototypeMethods = COM_PROTOTYPES_(UCase$(objectType))
    
    Dim methodInfo As ASF_Map
    Set methodInfo = prototypeMethods(UCase$(methodName))
    
    HasCOMPrototype = True
    Exit Function
    
NoPrototype:
    HasCOMPrototype = False
End Function
' Get prototype method info for a COM object type and method name
Private Function GetCOMPrototype(obj As Variant, methodName As String) As ASF_Map
    InitializeCOMPrototypes
    
    On Error GoTo NoPrototype
    
    Dim objectType As String: objectType = GetCOMObjectType(obj)
    If objectType = vbNullString Then GoTo NoPrototype
    
    Dim prototypeMethods As Collection
    Set prototypeMethods = COM_PROTOTYPES_(UCase$(objectType))
    
    Set GetCOMPrototype = prototypeMethods(UCase$(methodName))
    Exit Function
    
NoPrototype:
    Set GetCOMPrototype = Nothing
End Function
' Determine COM object type using TypeName
Private Function GetCOMObjectType(obj As Variant) As String
    If IsObject(obj) Then
        GetCOMObjectType = typeName(obj)
    Else
        GetCOMObjectType = vbNullString
    End If
End Function
Public Property Get EnableCallTrace() As Boolean
    EnableCallTrace = ENABLE_CALL_TRACE_
End Property
Public Property Let EnableCallTrace(aValue As Boolean)
    ENABLE_CALL_TRACE_ = aValue
End Property
Public Property Get GetCallStack() As Collection
    Set GetCallStack = GLOBALS_.gCallStack
End Property
Public Property Get CastInjectedVars() As Boolean
    CastInjectedVars = CAST_INJECTED_VARIABLES_
End Property
Public Property Let CastInjectedVars(aValue As Boolean)
    CAST_INJECTED_VARIABLES_ = aValue
End Property
Public Property Get AppAccess() As Boolean
    AppAccess = APPLICATION_ACCESS_
End Property
Public Property Let AppAccess(aValue As Boolean)
    APPLICATION_ACCESS_ = aValue
End Property
Public Property Get OverrideCollMethods() As Boolean
    OverrideCollMethods = COLLECTIONS_METHOD_OVERRIDE_
End Property
Public Property Let OverrideCollMethods(aValue As Boolean)
    COLLECTIONS_METHOD_OVERRIDE_ = aValue
End Property
Public Function GetCallStackTrace() As String
    Dim sb As String
    Dim i As Long
    
    For i = 1 To GLOBALS_.gCallStack.Count
        Dim trace As ASF_Map
        Set trace = GLOBALS_.gCallStack(i)
        
        Select Case trace.GetValue("type")
            Case "CALL"
                sb = sb & "CALL: " & trace.GetValue("name")
                
                ' Show arguments
                Dim args As Collection
                Set args = trace.GetValue("args")
                If args.Count > 0 Then
                    sb = sb & "("
                    Dim j As Long
                    For j = 1 To args.Count
                        If j > 1 Then sb = sb & ", "
                        sb = sb & CStr(args(j))
                    Next j
                    sb = sb & ")"
                Else
                    sb = sb & "()"
                End If
                
                ' Show return value if available
                If trace.Exists("returnValue") Then
                    sb = sb & " -> " & CStr(trace.GetValue("returnValue"))
                End If
                
                sb = sb & vbCrLf
                
            Case "ERROR"
                sb = sb & "ERROR in " & trace.GetValue("name") & ": " & trace.GetValue("error") & vbCrLf
        End Select
    Next i
    
    GetCallStackTrace = sb
End Function

' Clear the call stack
Public Sub ClearCallStack()
    Set GLOBALS_.gCallStack = New Collection
End Sub
Public Function VarToASF(VarValue As Variant) As Variant
    VarToASF = CastInjectedVar(VarValue)
End Function
Private Function CastInjectedVar(VarValue As Variant) As Variant
    If IsArray(VarValue) Then
        CastInjectedVar = ArrayToASF(VarValue)
    ElseIf IsVBAcollection(VarValue) Then
        CastInjectedVar = CollToASF(VarValue)
    ElseIf typeName(VarValue) = "ASF_Map" Then
        Set CastInjectedVar = VarValue
    ElseIf typeName(VarValue) = "ASF_RegexEngine" Then
        Set CastInjectedVar = VarValue
    ElseIf IsObject(VarValue) Then
        CastInjectedVar = typeName(VarValue)
    Else
        CastInjectedVar = CastType(VarValue)
    End If
End Function
Private Function CastType(aValue As Variant) As Variant
    If typeName(aValue) = "String" Then
        CastType = CStr(aValue)
    ElseIf typeName(aValue) = "Boolean" Then
        CastType = CBool(aValue)
    ElseIf IsNumeric(aValue) Then
        CastType = CDbl(aValue)
    Else
        CastType = aValue
    End If
End Function
' Convert native VBA array -> ASF one-based internal array representation
Private Function ArrayToASF(arr As Variant) As Variant
    On Error GoTo ErrHandler
    Dim lb As Long, ub As Long
    Dim tmpResult() As Variant
    lb = LBound(arr): ub = UBound(arr)
    Dim n As Long: n = ub - lb + 1
    If n <= 0 Then ArrayToASF = Array(): Exit Function
    Dim i As Long
    ReDim tmpResult(1 To n)
    If Is2Darray(arr) Then
        Dim ii As Long
        Dim lb2 As Long, ub2 As Long
        lb2 = LBound(arr, 2): ub2 = UBound(arr, 2)
        Dim tmpElm() As Variant
        For i = lb To ub
            ReDim tmpElm(1 To ub2 - lb2 + 1)
            For ii = lb2 To ub2
                If IsArray(arr(i, ii)) Then
                    tmpElm(1 + ii - lb2) = ArrayToASF(arr(i, ii))
                ElseIf IsVBAcollection(arr(i, ii)) Then
                    tmpElm(1 + ii - lb2) = CollToASF(arr(i, ii))(1)
                ElseIf IsObject(arr(i, ii)) Then
                    tmpElm(1 + ii - lb2) = typeName(arr(i, ii))
                Else
                    tmpElm(1 + ii - lb2) = CastType(arr(i, ii))
                End If
            Next ii
            tmpResult(1 + i - lb) = tmpElm
        Next i
    Else
        For i = lb To ub
            If IsArray(arr(i)) Then
                tmpResult(1 + i - lb) = ArrayToASF(arr(i))
            ElseIf IsVBAcollection(arr(i)) Then
                tmpResult(1 + i - lb) = CollToASF(arr(i))(1)
            ElseIf IsObject(arr(i)) Then
                tmpResult(1 + i - lb) = typeName(arr(i))
            Else
                tmpResult(1 + i - lb) = CastType(arr(i))
            End If
        Next i
    End If
    ArrayToASF = tmpResult
    Erase tmpResult
    Exit Function

ErrHandler:
    ArrayToASF = Array("[<error>]")
    err.Clear
End Function
' Convert native VBA array -> ASF no casting
Private Function ArrayToASF2(arr As Variant) As Variant
    On Error GoTo ErrHandler
    Dim lb As Long, ub As Long
    Dim tmpResult() As Variant
    lb = LBound(arr): ub = UBound(arr)
    Dim n As Long: n = ub - lb + 1
    If n <= 0 Then ArrayToASF2 = Array(): Exit Function
    Dim i As Long
    ReDim tmpResult(1 To n)
    If Is2Darray(arr) Then
        Dim ii As Long
        Dim lb2 As Long, ub2 As Long
        lb2 = LBound(arr, 2): ub2 = UBound(arr, 2)
        Dim tmpElm() As Variant
        For i = lb To ub
            ReDim tmpElm(1 To ub2 - lb2 + 1)
            For ii = lb2 To ub2
                vAssignment tmpElm(1 + ii - lb2), arr(i, ii)
            Next ii
            tmpResult(1 + i - lb) = tmpElm
        Next i
    Else
        For i = lb To ub
            vAssignment tmpResult(1 + i - lb), arr(i)
        Next i
    End If
    ArrayToASF2 = tmpResult
    Erase tmpResult
    Exit Function

ErrHandler:
    err.Clear
    ArrayToASF2 = Array()
End Function
' Convert native ASF to 2 dim VBA array, no casting
Private Function ASFto2Darray(arr As Variant) As Variant
    On Error GoTo ErrHandler
    Dim lb As Long, ub As Long
    Dim lb2 As Long, ub2 As Long
    Dim tmpResult() As Variant
    lb = LBound(arr): ub = UBound(arr)
    Dim n As Long: n = ub - lb + 1
    If n <= 0 Then ASFto2Darray = Array(): Exit Function
    Dim i As Long
    Dim j As Long
    Dim tmpLng As Long
    Dim m As Long
    For i = lb To ub
        tmpLng = UBound(arr(i)) - LBound(arr(i)) + 1
        If tmpLng > m Then m = tmpLng
    Next i
    ReDim tmpResult(1 To n, 1 To m)
    For i = lb To ub
        lb2 = LBound(arr(i))
        ub2 = UBound(arr(i))
        For j = lb2 To ub2
            vAssignment tmpResult(i - lb + 1, j - lb2 + 1), arr(i)(j)
        Next j
    Next i
    ASFto2Darray = tmpResult
    Erase tmpResult
    Exit Function

ErrHandler:
    err.Clear
    ASFto2Darray = Array()
End Function
' Convert native VBA Collection -> ASF one-based internal array representation
Private Function CollToASF(coll As Variant, Optional CastElem As Boolean = True) As Variant
    On Error GoTo ErrHandler
    Dim tmpResult() As Variant
    Dim tmpElm As Variant
    Dim i As Long: i = 1
    If coll.Count = 0 Then CollToASF = Array(): Exit Function
    ReDim tmpResult(1 To coll.Count)
    For Each tmpElm In coll
        If IsArray(tmpElm) Then
            If CastElem Then
                tmpResult(i) = ArrayToASF(tmpElm)
            Else
                tmpResult(i) = ArrayToASF2(tmpElm)
            End If
        ElseIf IsVBAcollection(tmpElm) Then
            tmpResult(i) = CollToASF(tmpElm, CastElem)
        ElseIf typeName(tmpElm) = "String" Then
            tmpResult(i) = CStr(tmpElm)
        ElseIf typeName(tmpElm) = "Boolean" Then
            tmpResult(i) = CBool(tmpElm)
        ElseIf IsNumeric(tmpElm) Then
            tmpResult(i) = CDbl(tmpElm)
        Else
            If CastElem Then
                If Not IsObject(tmpElm) Then
                    tmpResult(i) = tmpElm
                Else
                    tmpResult(i) = typeName(tmpElm)
                End If
            Else
                vAssignment tmpResult(i), tmpElm
            End If
        End If
        i = i + 1
    Next tmpElm
    CollToASF = tmpResult
    Erase tmpResult
    Exit Function

ErrHandler:
    CollToASF = Array("[<error>]")
    err.Clear
End Function

Private Function Is2Darray(arr As Variant) As Boolean
    Dim d As Long
    
    On Error GoTo err_Handler
    d = UBound(arr, 2)
    Is2Darray = True
    Exit Function
err_Handler:
    Is2Darray = False
End Function
Public Property Get verbose() As Boolean
    verbose = VERBOSE_
End Property
Public Property Let verbose(aValue As Boolean)
    VERBOSE_ = aValue
End Property

' ---------------------------------------------------------------- services for ASF_VM
Public Function FormatValue(v As Variant) As String
    FormatValue = ValueToStringForPrint(v)
End Function
Public Function TypeOfValue(v As Variant) As String
    TypeOfValue = builtin_typeof(v)
End Function
' What a host parameter becomes before it is bound to $1, $2...
Public Function PrepareParam(v As Variant) As Variant
    If CAST_INJECTED_VARIABLES_ And Not APPLICATION_ACCESS_ Then
        vAssignment PrepareParam, CastInjectedVar(v)
    ElseIf IsArray(v) And Not IsObject(v) Then
        PrepareParam = ArrayToASF2(v)
    Else
        vAssignment PrepareParam, v
    End If
End Function
' base.prop as this VM reads it: members of strings, arrays, collections and host objects
Public Function MemberOf(ByRef target As Variant, ByVal prop As String) As Variant
    vAssignment MemberOf, MemberValue(target, prop)
End Function
' base.prop = value on a value that is not an ASF object: a host object's property
Public Sub SetMemberOf(ByRef target As Variant, ByVal prop As String, ByRef value As Variant)
    If IsMapValue(target) Then
        Dim m As ASF_Map: Set m = target
        m.SetValue prop, value
    Else
        SetObjectProperty target, prop, value
    End If
End Sub
' receiver.name(args...) as this VM calls it (methods of arrays, strings, objects, Math, regex and host
' objects); wrote and written give the array a mutating method (push, sort...) writes back
Public Function CallMemberOn(ByRef target As Variant, ByVal name As String, ByRef args As Variant, _
                             ByRef wrote As Boolean, ByRef written As Variant) As Variant
    Dim isMath As Boolean
    If IsObject(target) Then isMath = (target Is MathObj)
    vAssignment CallMemberOn, MethodCallValue(target, name, ArrayToCollection(args), isMath, wrote, written)
End Function
' A built-in function (range, flatten, clone, isArray, isNumeric, forEach, regex, cwd, scwd)
Public Function CallBuiltin(ByVal name As String, ByRef args As Variant) As Variant
    Dim result As Variant
    If EvalNamedBuiltin(name, ArrayToCollection(args), result) Then
        vAssignment CallBuiltin, result
    End If
End Function
Private Function ArrayToCollection(ByRef args As Variant) As Collection
    Dim out As Collection: Set out = New Collection
    Dim i As Long
    If IsArray(args) Then
        For i = LBound(args) To UBound(args)
            out.Add args(i)
        Next i
    End If
    Set ArrayToCollection = out
End Function
' The Math object, shared with ASF_VM so that members of Math resolve alike
Public Property Get MathObject() As Collection
    If MathObj Is Nothing Then Set MathObj = New Collection
    Set MathObject = MathObj
End Property
' ASF_VM sets itself here while it runs, and Nothing after: the closures it made run there
Public Sub SetHost(host As Object)
    Set VM_ = host
End Sub
' The path of a module as import resolves it
Public Function ResolveModule(ByVal modulePath As String) As String
    ResolveModule = ResolveModulePath(modulePath)
End Function
' @(expr), with the variables the code can see (later names win)
Public Function VBAExpression(ByVal expr As String, ByVal names As Variant, ByVal values As Variant) As Variant
    Dim vars As ASF_Map: Set vars = New ASF_Map
    Dim elm As Variant, i As Long
    If Not INJECTED_VARIABLES_ Is Nothing Then
        For Each elm In INJECTED_VARIABLES_
            vars.SetValue CStr(elm(LBound(elm))), elm(UBound(elm))
        Next elm
    End If
    For i = LBound(names) To UBound(names)
        vars.SetValue CStr(names(i)), values(i)
    Next i
    vAssignment VBAExpression, EvalVBExpressionWithScope(expr, vars)
End Function
' Variables given with InjectVariable, as Array(name, value) items
Public Property Get InjectedVariables() As Collection
    Set InjectedVariables = INJECTED_VARIABLES_
End Property

' base.prop on a value: own properties and static methods of objects and classes, the built-in
' members of arrays and strings, the constants of Math, and the properties of host objects
Private Function MemberValue(ByRef baseVal As Variant, ByVal prop As String) As Variant
    On Error GoTo ErrHandler
    If IsMapValue(baseVal) Then
        Dim baseMap As ASF_Map: Set baseMap = baseVal
        Dim classMethods As Variant
        ' a class keeps its methods aside
        If baseMap.TryGetValue("classMethods", classMethods) Then
            Dim method_ As Variant
            For Each method_ In classMethods
                If method_.GetValue("name") = prop Then
                    vAssignment MemberValue, method_
                    Exit For
                End If
            Next method_
        Else
            ' a property the object does not have is undefined, not the Nothing of GetValue
            Dim propVal As Variant
            If baseMap.TryGetValue(prop, propVal) Then
                vAssignment MemberValue, propVal
            End If
        End If
        Exit Function
    End If
    If typeName(baseVal) = "String" Or IsArray(baseVal) Then
        If IsObject(baseVal) Then
            vAssignment MemberValue, CallObjectMethod(baseVal, prop)
        Else
            ' the member can be a method called without parentheses
            Dim wrote As Boolean, written As Variant
            vAssignment MemberValue, MethodCallValue(baseVal, prop, New Collection, False, wrote, written)
        End If
        Exit Function
    End If
    ' Math.PI and the other constants are read as properties
    If IsObject(baseVal) Then
        If baseVal Is MathObj Then
            Dim constVal As Variant
            If MathConstant(prop, constVal) Then
                MemberValue = constVal
                Exit Function
            End If
        End If
    End If
    vAssignment MemberValue, CallObjectMethod(baseVal, prop)
    Exit Function
ErrHandler:
    If err.Number = (vbObjectError + 5007) Then
        err.Raise err.Number, err.Source, err.Description
    Else
        err.Clear
        vAssignment MemberValue, Empty
    End If
End Function
' target.name(args) on values: a built-in method, a prototype of a host
' object, or a function the object holds, called with this; wrote and written: the array a method
' changed, for the caller to write back
Private Function MethodCallValue(ByRef target As Variant, ByVal name As String, evaluated As Collection, ByVal isMath As Boolean, _
                                 ByRef wrote As Boolean, ByRef written As Variant) As Variant
    Dim propN As String, baseValLocal As Variant, origBaseVal As Variant, tmpResult As Variant
    vAssignment baseValLocal, target
    METHOD_NAME_ = name
    propN = LCase$(name)
    METHOD_WROTE_ = False
    Dim handled As Boolean
    handled = DispatchMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult, isMath)
    wrote = METHOD_WROTE_
    If wrote Then vAssignment written, METHOD_WRITTEN_
    METHOD_WROTE_ = False
    METHOD_WRITTEN_ = Empty
    If handled Then
        vAssignment MethodCallValue, tmpResult
        Exit Function
    End If
    If HasCOMPrototype(origBaseVal, propN) Then
        vAssignment MethodCallValue, CallObjectMethod(origBaseVal, propN, evaluated)
        Exit Function
    End If
    If Not IsObject(origBaseVal) And (IsArray(origBaseVal) Or VarType(origBaseVal) = vbString) Then
        err.Raise vbObjectError + 5012, "VM.EvalCall", "'" & propN & "' is not a method of " & IIf(IsArray(origBaseVal), "an array", "a string")
    End If
    Dim callee As Variant
    vAssignment callee, MemberValue(target, name)
    If Not IsMapValue(callee) Then Exit Function
    vAssignment MethodCallValue, CallClosure(callee, evaluated, baseValLocal)
End Function
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub

Private Function EvalNamedBuiltin(ByVal fname As String, evaluated As Collection, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim nameLcase As String: nameLcase = LCase$(fname)
    If nameLcase = "range" Then
        ' range([start,] end[, step]) -> returns an array (python-style range semantics)
        Dim rcount As Long: rcount = evaluated.Count
        Dim rs As Double, re As Double, rstep As Double
        Dim optionBase_local As Long: optionBase_local = 1
        If rcount = 1 Then
            rs = 0
            re = CDbl(evaluated(1))
            rstep = 1
        ElseIf rcount = 2 Then
            rs = CDbl(evaluated(1))
            re = CDbl(evaluated(2))
            rstep = 1
        ElseIf rcount >= 3 Then
            rs = CDbl(evaluated(1))
            re = CDbl(evaluated(2))
            rstep = CDbl(evaluated(3))
        Else
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If
        Dim tmpArr() As Variant
        Dim vals As New Collection
        Dim vv As Double: vv = rs
        If rstep = 0 Then rstep = 1
        If rstep > 0 Then
            Do While vv < re
                vals.Add vv
                vv = vv + rstep
            Loop
        Else
            Do While vv > re
                vals.Add vv
                vv = vv + rstep
            Loop
        End If
        If vals.Count = 0 Then
            ReDim tmpArr(optionBase_local To optionBase_local)
        Else
            ReDim tmpArr(optionBase_local To optionBase_local + vals.Count - 1)
            Dim ri As Long
            For ri = 1 To vals.Count
                vAssignment tmpArr(optionBase_local + ri - 1), vals(ri)
            Next ri
        End If
        vAssignment tmpResult, tmpArr
        GoTo exitfun
    End If
    If nameLcase = "flatten" Then
        If evaluated.Count = 0 Then
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If
        If evaluated.Count = 1 Then
            vAssignment tmpResult, builtin_flatten(evaluated(1))
        Else
            vAssignment tmpResult, builtin_flatten(evaluated(1), evaluated(2))
        End If
        GoTo exitfun
    End If
    If nameLcase = "clone" Then
        If evaluated.Count = 0 Then
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If
        vAssignment tmpResult, builtin_clone(evaluated(1))
        GoTo exitfun
    End If
    If nameLcase = "isarray" Then
        If evaluated.Count = 0 Then
            vAssignment tmpResult, False
        Else
            Dim v0 As Variant: v0 = evaluated(1)
            On Error Resume Next
            Dim isArr As Boolean: isArr = IsArray(v0)
            If err.Number <> 0 Then
                err.Clear
                isArr = False
            End If
            vAssignment tmpResult, CBool(isArr)
        End If
        GoTo exitfun
    End If
    If nameLcase = "isnumeric" Then
        If evaluated.Count = 0 Then
            vAssignment tmpResult, False
        Else
            Dim v1 As Variant: v1 = evaluated(1)
            Dim numericRes As Boolean
            On Error Resume Next
            ' reject arrays and objects quickly
            If IsArray(v1) Then
                numericRes = False
            ElseIf IsObject(v1) Then
                numericRes = False
            Else
                numericRes = VBA.IsNumeric(v1) ' use VBA's IsNumeric for broad compatibility
            End If
            If err.Number <> 0 Then
                err.Clear
                numericRes = False
            End If
            vAssignment tmpResult, CBool(numericRes)
        End If
        GoTo exitfun
    End If
    If nameLcase = "foreach" Then
        ' foreach(arr, fn) | foreach(obj, fn) -> same behavior as method-style forEach
        If evaluated.Count < 2 Then
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If

        Dim feTmpVal As Variant: vAssignment feTmpVal, evaluated(1)
        Dim cbFn As Variant: vAssignment cbFn, evaluated(2)

        If typeName(cbFn) <> "ASF_Map" Or cbFn.GetValue("type") <> "Closure" Then
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If

        If Not IsArray(feTmpVal) And typeName(feTmpVal) <> "ASF_Map" Then
            vAssignment tmpResult, Empty
            GoTo exitfun
        End If

        If typeName(feTmpVal) = "ASF_Map" Then
            Dim foreachKeys2 As Collection: Set foreachKeys2 = feTmpVal.keys()
            Dim fk2 As Variant
            Dim args_obj2 As Collection
            Dim tmpRes_obj2 As Variant

            For Each fk2 In foreachKeys2
                Set args_obj2 = New Collection
                args_obj2.Add feTmpVal.GetValue(CStr(fk2))
                args_obj2.Add CStr(fk2)
                args_obj2.Add feTmpVal
                vAssignment tmpMapObj, cbFn
                vAssignment tmpRes_obj2, CallClosure(tmpMapObj, args_obj2, feTmpVal)
            Next fk2
        Else
            Dim lb2 As Long: lb2 = LBound(feTmpVal)
            Dim ub2 As Long: ub2 = UBound(feTmpVal)
            Dim optBase2 As Long: optBase2 = 1
            Dim idx2 As Long: idx2 = 0
            Dim j As Long
            Dim callArgs2 As Collection
            For j = lb2 To ub2
                Set callArgs2 = New Collection
                callArgs2.Add feTmpVal(j)
                callArgs2.Add (optBase2 + idx2)
                callArgs2.Add feTmpVal
                Dim tmpRc As Variant
                vAssignment tmpMapObj, cbFn
                vAssignment tmpRc, CallClosure(tmpMapObj, callArgs2, feTmpVal)
                idx2 = idx2 + 1
            Next j
        End If
        ' return original array/object for convenience
        vAssignment tmpResult, feTmpVal
        GoTo exitfun
    End If
    If nameLcase = "regex" Then
        Set tmpResult = New ASF_RegexEngine
        Set tmpResult.SharedCache = GLOBALS_.regexCache
        Select Case evaluated.Count
            Case 1
                ' pattern
                tmpResult.Init CStr(evaluated(1))
            Case Is <= 2
                ' pattern, ignoreCase
                tmpResult.Init CStr(evaluated(1)), CBool(evaluated(2))
            Case Is <= 3
                ' pattern, ignoreCase flag, multiline flag
                tmpResult.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3))
            Case Is <= 4
                ' pattern, ignoreCase flag, multiline flag, dotAll flag
                tmpResult.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3)), CBool(evaluated(4))
            Case Is <= 5
                ' pattern, ignoreCase flag, multiline flag, dotAll flag, global flag
                tmpResult.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3)), CBool(evaluated(4)), CBool(evaluated(5))
        End Select
        GoTo exitfun
    End If
    ' Retrieve the current working directory
    If nameLcase = "cwd" Then
        tmpResult = GLOBALS_.CURRENT_MODULE_PATH
        GoTo exitfun
    End If
    ' Set the current working directory
    If nameLcase = "scwd" Then
        GLOBALS_.CURRENT_MODULE_PATH = CStr(evaluated(1))
        GoTo exitfun
    End If
    Exit Function
exitfun:
    EvalNamedBuiltin = True
End Function

' The built-in methods of a value (arrays, strings, objects, collections, Math, regex, host objects):
' False when it has none of that name. A method that changes its array calls WriteBack with it
Private Function DispatchMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant, ByVal isMath As Boolean) As Boolean
    vAssignment origBaseVal, baseValLocal
    If IsVBAcollection(baseValLocal) Then
        If COLLECTIONS_METHOD_OVERRIDE_ Then
            If BinarySearch(GLOBALS_.builtinArrayMethods, propN) = -1 Then
                DispatchMethod = CollectionMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult, isMath)
                Exit Function
            End If
        End If
        If Not isMath And Not APPLICATION_ACCESS_ Then
            baseValLocal = CollToASF(baseValLocal, False)
        End If
    End If
    Dim baseType As String: baseType = typeName(baseValLocal)
    If IsArray(baseValLocal) Then
        DispatchMethod = ArrayMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
    ElseIf baseType = "ASF_RegexEngine" Then
        DispatchMethod = RegexMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
    ElseIf baseType = "String" Then
        DispatchMethod = StringMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
    ElseIf baseType = "ASF_Map" Then
        DispatchMethod = MapMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
    ElseIf baseType = "Collection" Then
        DispatchMethod = CollectionMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult, isMath)
    Else
        vAssignment tmpResult, CallObjectMethod(baseValLocal, propN, evaluated)
        DispatchMethod = True
    End If
End Function
' A method that changes its receiver gives the new array to the caller, which writes it back
Private Sub WriteBack(ByRef value As Variant)
    METHOD_WROTE_ = True
    vAssignment METHOD_WRITTEN_, value
End Sub

' Array methods
Private Function ArrayMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim optBase_sp As Long, len_sp As Long
    Dim optBase_u2 As Long
    Dim len_base As Long
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim startRaw_sp As Long, startN_sp As Long, deleteCount_sp As Long
    Dim tempIn() As Variant, newTemp() As Variant
    Dim insertCount As Long, newSize As Long
    Dim i_sp As Long, j_sp As Long, cur As Long
    Dim insertedItem As Variant
    ' Array methods (map/filter/reduce/slice/push/pop)
    Select Case propN
        Case "length"
            If UBound(baseValLocal) <> -1 Then
                tmpResult = UBound(baseValLocal) - LBound(baseValLocal) + 1
            Else
                tmpResult = 0
            End If
            GoTo exitfun
        Case "map"
            ArrayMethod = ArrMap(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "filter"
            ArrayMethod = ArrFilter(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "reduce"
            ArrayMethod = ArrReduce(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "slice"
            ArrayMethod = ArrSlice(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "push"
            ArrayMethod = ArrPush(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "pop"
            ArrayMethod = ArrPop(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "foreach"
            ArrayMethod = ArrForeach(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "unique"
            ArrayMethod = ArrUnique(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "concat"
            ArrayMethod = ArrConcat(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "tostring", "join"
            ArrayMethod = ArrTostring(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "shift"
            ArrayMethod = ArrShift(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "unshift"
            ' unshift: add elements at start, returns new length
            If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
            optBase_u2 = 1
            lb_b = LBound(baseValLocal)
            ub_b = UBound(baseValLocal)
            len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
            insertCount = evaluated.Count
            newSize = len_base + insertCount
            If newSize <= 0 Then
                ReDim outArr(optBase_u2 To optBase_u2)
            Else
                ReDim newTemp(0 To newSize - 1)
                cur = 0
                ' copy inserted args (in order)
                For j_sp = 1 To evaluated.Count
                    vAssignment insertedItem, evaluated(j_sp)
                    newTemp(cur) = insertedItem
                    cur = cur + 1
                Next j_sp
                ' copy old items
                For i_sp = 0 To len_base - 1
                    newTemp(cur) = baseValLocal(lb_b + i_sp)
                    cur = cur + 1
                Next i_sp
                ' convert to option-base array
                ReDim outArr(optBase_u2 To optBase_u2 + newSize - 1)
                For i_sp = 0 To newSize - 1
                    outArr(optBase_u2 + i_sp) = newTemp(i_sp)
                Next i_sp
            End If
            vAssignment tmpResult, IIf(newSize <= 0, 0, (UBound(outArr) - LBound(outArr) + 1))
            baseValLocal = outArr
            WriteBack baseValLocal
            GoTo exitfun
        Case "delete"
            ' delete(index): remove element at index (user-space) and return true/false
            If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
            optBase_u2 = 1
            lb_b = LBound(baseValLocal)
            ub_b = UBound(baseValLocal)
            len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
            startRaw_sp = CLng(evaluated(1))
            startN_sp = NormalizeIndex(startRaw_sp, len_base, optBase_u2) ' zero-based offset
            If startN_sp >= len_base Then
                vAssignment tmpResult, False: GoTo exitfun
            End If
            ' build tempIn
            If len_base > 0 Then
                ReDim tempIn(0 To len_base - 1)
                For i_sp = 0 To len_base - 1
                    vAssignment tempIn(i_sp), baseValLocal(lb_b + i_sp)
                Next i_sp
            Else
                ReDim tempIn(0 To -1)
            End If
            ' build new array without the specified element
            newSize = len_base - 1
            If newSize <= 0 Then
                ReDim outArr(optBase_u2 To optBase_u2)
            Else
                ReDim newTemp(0 To newSize - 1)
                cur = 0
                For i_sp = 0 To len_base - 1
                    If i_sp <> startN_sp Then
                        vAssignment newTemp(cur), tempIn(i_sp)
                        cur = cur + 1
                    End If
                Next i_sp
                ReDim outArr(optBase_u2 To optBase_u2 + newSize - 1)
                For i_sp = 0 To newSize - 1
                    vAssignment outArr(optBase_u2 + i_sp), newTemp(i_sp)
                Next i_sp
            End If
            baseValLocal = outArr
            vAssignment tmpResult, True
            ' write back to original container via ResolvePath
            WriteBack baseValLocal
            GoTo exitfun
        Case "splice"
            ArrayMethod = ArrSplice(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "tospliced"
            ' toSpliced(start, deleteCount?, ...items) -> non-mutating version of splice
            If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
            optBase_sp = 1
            lb_b = LBound(baseValLocal)
            ub_b = UBound(baseValLocal)
            len_sp = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)

            startRaw_sp = CLng(evaluated(1))
            startN_sp = NormalizeIndex(startRaw_sp, len_sp, optBase_sp) ' zero-based offset within logical array
            If evaluated.Count >= 2 Then
                deleteCount_sp = CLng(evaluated(2))
                If deleteCount_sp < 0 Then deleteCount_sp = 0
            Else
                deleteCount_sp = IIf(len_sp - startN_sp > 0, len_sp - startN_sp, 0)
            End If
            ' prepare tempIn from baseValLocal (zero-based)
            If len_sp > 0 Then
                ReDim tempIn(0 To len_sp - 1)
                For i_sp = 0 To len_sp - 1
                    vAssignment tempIn(i_sp), baseValLocal(lb_b + i_sp)
                Next i_sp
            Else
                ReDim tempIn(0 To -1)
            End If
            ' build newTemp: prefix, inserted items, suffix
            insertCount = IIf(evaluated.Count > 2, evaluated.Count - 2, 0)
            newSize = len_sp - deleteCount_sp + insertCount
            If newSize > 0 Then ReDim newTemp(0 To newSize - 1) Else ReDim newTemp(0 To -1)
            cur = 0
            ' copy prefix
            For i_sp = 0 To startN_sp - 1
                If i_sp <= UBound(tempIn) Then
                    vAssignment newTemp(cur), tempIn(i_sp)
                    cur = cur + 1
                End If
            Next i_sp
            ' insert new items (in order)
            For i_sp = 3 To evaluated.Count
                vAssignment insertedItem, evaluated(i_sp)
                vAssignment newTemp(cur), insertedItem
                cur = cur + 1
            Next i_sp
            ' copy suffix
            For i_sp = startN_sp + deleteCount_sp To len_sp - 1
                If i_sp <= UBound(tempIn) Then
                    vAssignment newTemp(cur), tempIn(i_sp)
                    cur = cur + 1
                End If
            Next i_sp
            ' convert newTemp -> option-base array
            If newSize <= 0 Then
                ReDim outArr(optBase_sp To optBase_sp)
            Else
                ReDim outArr(optBase_sp To optBase_sp + newSize - 1)
                For i_sp = 0 To newSize - 1
                    vAssignment outArr(optBase_sp + i_sp), newTemp(i_sp)
                Next i_sp
            End If
            vAssignment tmpResult, outArr
            GoTo exitfun
        Case "at"
            ArrayMethod = ArrAt(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "copywithin"
            ArrayMethod = ArrCopywithin(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "entries"
            ArrayMethod = ArrEntries(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "every"
            ArrayMethod = ArrEvery(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "find", "findindex", "findlast", "findlastindex"
            ArrayMethod = ArrFind(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "from"
            ArrayMethod = ArrFrom(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "includes", "indexof", "lastindexof"
            ArrayMethod = ArrIncludes(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "of"
            ArrayMethod = ArrOf(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "reverse"
            ArrayMethod = ArrReverse(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "toreversed"
            ArrayMethod = ArrToreversed(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "some"
            ArrayMethod = ArrSome(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "with"
            ArrayMethod = ArrWith(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "sort", "tosorted"
            ArrayMethod = ArrSort(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case Else
            Exit Function
    End Select
    Exit Function
exitfun:
    ArrayMethod = True
End Function

Private Function ArrMap(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim ii As Long
    Dim outArr() As Variant
    ' map(fn) -> return new array with fn applied to each element (recursively tolerates nested arrays)
    If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim mapper As Variant: vAssignment mapper, evaluated(1)
    If typeName(mapper) <> "ASF_Map" Or mapper.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty: GoTo exitfun
    End If

    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim baseLB As Long, baseUB As Long
    On Error Resume Next
    baseLB = LBound(baseValLocal)
    baseUB = UBound(baseValLocal)
    If err.Number <> 0 Then
        err.Clear
        vAssignment tmpResult, Empty: GoTo exitfun
    End If
    Dim countEls As Long: countEls = baseUB - baseLB + 1
    If countEls < 0 Then countEls = 0

    Dim optBaseLocal As Long: optBaseLocal = 1
    If countEls = 0 Then
        ReDim outArr(optBaseLocal To optBaseLocal)
    Else
        ReDim outArr(optBaseLocal To optBaseLocal + countEls - 1)
    End If

    vAssignment tmpMapObj, mapper
    For ii = 0 To countEls - 1
        Dim srcIndex As Long: srcIndex = baseLB + ii
        Dim resVal As Variant
        ' Recursively apply mapper to nested arrays (preserves nested structure)
        vAssignment resVal, map_apply_recursive(baseValLocal(srcIndex), _
                                tmpMapObj, baseValLocal, optBaseLocal)
        vAssignment outArr(optBaseLocal + ii), resVal
    Next ii

    vAssignment tmpResult, outArr
    GoTo exitfun

    Exit Function
exitfun:
    ArrMap = True
End Function

Private Function ArrFilter(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' filter(fn) -> return new array with elements where fn(elem) truthy
    If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim pred As Variant: vAssignment pred, evaluated(1)
    If typeName(pred) <> "ASF_Map" Or pred.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty: GoTo exitfun
    End If
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim tmpCol As New Collection
    Dim jj As Long
    Dim keep As Variant
    Dim ca As Collection
    For jj = LBound(baseValLocal) To UBound(baseValLocal)
        Set ca = New Collection
        ca.Add baseValLocal(jj)
        vAssignment tmpMapObj, pred
        vAssignment keep, CallClosure(tmpMapObj, ca, baseValLocal)
        If IsTruthy(keep) Then tmpCol.Add baseValLocal(jj)
    Next jj
    ' build array from tmpCol
    Dim outA() As Variant
    Dim ob As Long: ob = 1
    If tmpCol.Count = 0 Then
        ReDim outA(ob To ob)
    Else
        ReDim outA(ob To ob + tmpCol.Count - 1)
        For jj = 1 To tmpCol.Count
            vAssignment outA(ob + jj - 1), tmpCol(jj)
        Next jj
    End If
    vAssignment tmpResult, outA
    GoTo exitfun
    Exit Function
exitfun:
    ArrFilter = True
End Function

Private Function ArrReduce(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' reduce(fn, initial?) -> apply reduction
    If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim rfn As Variant: vAssignment rfn, evaluated(1)
    If typeName(rfn) <> "ASF_Map" Or rfn.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty: GoTo exitfun
    End If
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim acc As Variant
    Dim startIdx As Long
    If evaluated.Count >= 2 Then
        vAssignment acc, evaluated(2)
        startIdx = LBound(baseValLocal)
    Else
        ' use first element as initial
        vAssignment acc, baseValLocal(LBound(baseValLocal))
        startIdx = LBound(baseValLocal) + 1
    End If
    Dim kk As Long
    Dim carr As Collection
    For kk = startIdx To UBound(baseValLocal)
        Set carr = New Collection
        carr.Add acc
        carr.Add baseValLocal(kk)
        vAssignment tmpMapObj, rfn
        vAssignment acc, CallClosure(tmpMapObj, carr, baseValLocal)
    Next kk
    vAssignment tmpResult, acc
    GoTo exitfun
    Exit Function
exitfun:
    ArrReduce = True
End Function

Private Function ArrSlice(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim optBase As Long
    Dim si As Long
    ' slice(start, end) -> non-mutating, respects __option_base
    Dim s As Long, e As Long
    Dim lenA As Long
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun

    ' number of elements in the array (independent of its LBound)
    lenA = UBound(baseValLocal) - LBound(baseValLocal) + 1

    ' runtime option base for arrays (0 or 1 typically)
    optBase = 1

    Dim sRaw As Long, eRaw As Long

    ' Interpret arguments in the user's indexing space (option base)
    If evaluated.Count = 0 Then
        ' default: full slice
        s = 0: e = lenA
    ElseIf evaluated.Count = 1 Then
        sRaw = CLng(evaluated(1))
        ' convert start to zero-based offset (and support negative indices)
        If sRaw < 0 Then
            s = lenA + sRaw
        Else
            s = sRaw - optBase
        End If
        e = lenA
    Else
        sRaw = CLng(evaluated(1))
        eRaw = CLng(evaluated(2))
        If sRaw < 0 Then
            s = lenA + sRaw
        Else
            s = sRaw - optBase
        End If
        If eRaw < 0 Then
            e = lenA + eRaw
        Else
            e = eRaw - optBase
        End If
    End If

    ' Normalize & clamp to valid ranges (zero-based offsets)
    If s < 0 Then s = 0
    If e < 0 Then e = 0
    If s > lenA Then s = lenA
    If e > lenA Then e = lenA
    If e < s Then e = s

    Dim cnt As Long: cnt = e - s

    Dim outS() As Variant
    Dim ob2 As Long: ob2 = optBase

    If cnt <= 0 Then
        ' Produce minimal array using same option base  keep conservative behavior:
        ' if you prefer an empty array representation, you can return Empty instead.
        ReDim outS(ob2 To ob2)
    Else
        ReDim outS(ob2 To ob2 + cnt - 1)
        Dim pidx As Long: pidx = 0
        Dim baseL As Long: baseL = LBound(baseValLocal)
        For si = s To e - 1
            vAssignment outS(ob2 + pidx), baseValLocal(baseL + si)
            pidx = pidx + 1
        Next si
    End If

    vAssignment tmpResult, outS
    GoTo exitfun
    Exit Function
exitfun:
    ArrSlice = True
End Function

Private Function ArrPush(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim insertCount As Long
    ' push(...values) -> mutates array, returns new length
    Dim offs As Long
    If Not IsArray(baseValLocal) Or UBound(baseValLocal) = -1 Then
        ' create new array and set it back
        Dim na() As Variant
        ReDim na(1 To 1)
        baseValLocal = na
        offs = 1 'Empty array
    Else
        offs = 0
    End If
    insertCount = evaluated.Count
    Dim oldLB As Long: oldLB = LBound(baseValLocal)
    Dim oldUB As Long: oldUB = UBound(baseValLocal)
    Dim oldLen As Long: oldLen = oldUB - oldLB + 1
    Dim newLen As Long: newLen = oldLen + insertCount
    Dim newArr As Variant
    ReDim newArr(oldLB To oldLB + newLen - 1 - offs)
    Dim zz As Long
    For zz = oldLB To oldUB - offs
        vAssignment newArr(zz), baseValLocal(zz)
    Next zz
    For zz = 1 To insertCount
        vAssignment newArr(oldLB + oldLen + zz - 1 - offs), evaluated(zz)
    Next zz
    ' write back to original container via ResolvePath
    WriteBack newArr
    vAssignment tmpResult, newLen
    GoTo exitfun
    Exit Function
exitfun:
    ArrPush = True
End Function

Private Function ArrPop(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' pop() -> remove last element, return it
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim lbp As Long: lbp = LBound(baseValLocal)
    Dim ubp As Long: ubp = UBound(baseValLocal)
    If ubp < lbp Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim popped As Variant: vAssignment popped, baseValLocal(ubp)
    If ubp = lbp Then
        Dim zeroA() As Variant
        ReDim zeroA(lbp To lbp)
        vAssignment baseValLocal, zeroA
    Else
        ReDim Preserve baseValLocal(lbp To ubp - 1)
    End If
    ' write back:
    WriteBack baseValLocal
    vAssignment tmpResult, popped
    GoTo exitfun
    Exit Function
exitfun:
    ArrPop = True
End Function

Private Function ArrForeach(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim optBase As Long
    Dim ii As Long
    ' forEach(fn) -> invoke fn(value, index, array) for each element; returns the original array
    If evaluated.Count = 0 Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim cb As Variant: vAssignment cb, evaluated(1)
    If typeName(cb) <> "ASF_Map" Or cb.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    ' must be an array to iterate
    If Not IsArray(baseValLocal) Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim lb As Long: lb = LBound(baseValLocal)
    Dim ub As Long: ub = UBound(baseValLocal)
    optBase = 1
    Dim idxLocal As Long: idxLocal = 0
    Dim callArgs As Collection
    For ii = lb To ub
        Set callArgs = New Collection
        ' value
        callArgs.Add baseValLocal(ii)
        ' index in user-space (respect option base)
        callArgs.Add (optBase + idxLocal)
        ' array reference as third arg
        callArgs.Add baseValLocal
        ' call closure with this = base array (so callback can use this)
        Dim tmpResCall As Variant
        vAssignment tmpMapObj, cb
        vAssignment tmpResCall, CallClosure(tmpMapObj, callArgs, baseValLocal)
        ' discard tmpResCall (forEach ignores return)
        idxLocal = idxLocal + 1
    Next ii

    ' return original array to allow chaining
    vAssignment tmpResult, baseValLocal
    GoTo exitfun
    Exit Function
exitfun:
    ArrForeach = True
End Function

Private Function ArrUnique(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim k As Long
    Dim outArr() As Variant
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_u As Long: optBase_u = 1
    Dim lb_u As Long: lb_u = LBound(baseValLocal)
    Dim ub_u As Long: ub_u = UBound(baseValLocal)
    Dim len_u As Long: len_u = IIf(ub_u >= lb_u, ub_u - lb_u + 1, 0)

    ' Fast temp buffer (worst-case size = len_u)
    Dim tmpBuf() As Variant
    If len_u <= 0 Then
        ReDim tmpBuf(0 To 0) ' empty
    Else
        ReDim tmpBuf(0 To len_u - 1)
    End If
    Dim outCount As Long: outCount = 0
    Dim i_u As Long
    Dim v_i As Variant
    For i_u = lb_u To ub_u
        vAssignment v_i, baseValLocal(i_u)
        Dim found As Boolean: found = False
        Dim j_u As Long
        For j_u = 0 To outCount - 1
            If DeepEqual(tmpBuf(j_u), v_i) Then
                found = True
                Exit For
            End If
        Next j_u
        If Not found Then
            ' append to tmpBuf
            vAssignment tmpBuf(outCount), v_i
            outCount = outCount + 1
        End If
    Next i_u
    ' Build result array honoring option base
    If outCount = 0 Then
        ReDim outArr(optBase_u To optBase_u)
    Else
        ReDim outArr(optBase_u To optBase_u + outCount - 1)
        For k = 0 To outCount - 1
            vAssignment outArr(optBase_u + k), tmpBuf(k)
        Next k
    End If
    vAssignment tmpResult, outArr
    GoTo exitfun
    Exit Function
exitfun:
    ArrUnique = True
End Function

Private Function ArrConcat(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim alb As Long, aub As Long
    Dim len_base As Long, totalCount As Long
    Dim outArr() As Variant, writeIdx As Long
    Dim avTmp As Variant
    Dim lb_b As Long, ub_b As Long
     ' concat: concat arrays/values into the callee array
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_c As Long: optBase_c = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    If ub_b >= lb_b Then len_base = ub_b - lb_b + 1 Else len_base = 0

    Dim argc As Long: argc = evaluated.Count
    Dim p As Long
    totalCount = len_base

    ' first pass: compute total size
    For p = 1 To argc
        vAssignment avTmp, evaluated(p)
        If IsArray(avTmp) Then
            alb = LBound(avTmp)
            aub = UBound(avTmp)
            If aub >= alb Then totalCount = totalCount + (aub - alb + 1)
        Else
            totalCount = totalCount + 1
        End If
    Next p

    ' allocate output array honoring option base
    If totalCount = 0 Then
        ReDim outArr(optBase_c To optBase_c)
    Else
        ReDim outArr(optBase_c To optBase_c + totalCount - 1)
    End If

    ' second pass: copy base and args into outArr
    writeIdx = 0
    Dim ci As Long, q As Long
    If len_base > 0 Then
        For ci = lb_b To ub_b
            vAssignment outArr(optBase_c + writeIdx), baseValLocal(ci)
            writeIdx = writeIdx + 1
        Next ci
    End If
    For p = 1 To argc
        vAssignment avTmp, evaluated(p)
        If IsArray(avTmp) Then
            alb = LBound(avTmp)
            aub = UBound(avTmp)
            If aub >= alb Then
                For q = alb To aub
                    vAssignment outArr(optBase_c + writeIdx), avTmp(q)
                    writeIdx = writeIdx + 1
                Next q
            End If
        Else
            vAssignment outArr(optBase_c + writeIdx), avTmp
            writeIdx = writeIdx + 1
        End If
    Next p
    vAssignment tmpResult, outArr
    ' write back to original container
    WriteBack tmpResult
    GoTo exitfun
    Exit Function
exitfun:
    ArrConcat = True
End Function

Private Function ArrTostring(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim sep As String
    Dim lb_val As Long, ub_val As Long, len_val As Long
    Dim parts() As String
    Dim ii_j As Long
    Dim valj As Variant
    ' toString / join: join(customSep)
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    sep = ", "
    ' join() passes one empty argument: the default separator, as join(undefined) in JavaScript
    If evaluated.Count >= 1 Then
        If Not IsEmpty(evaluated(1)) Then sep = CStr(evaluated(1))
    End If
    lb_val = LBound(baseValLocal)
    ub_val = UBound(baseValLocal)
    len_val = IIf(ub_val >= lb_val, ub_val - lb_val + 1, 0)
    If len_val <= 0 Then
        ' empty -> empty string
        vAssignment tmpResult, vbNullString
        GoTo exitfun
    End If
    ReDim parts(1 To len_val)
    For ii_j = 0 To len_val - 1
        vAssignment valj, baseValLocal(lb_val + ii_j)
        If IsArray(valj) Or (IsObject(valj) And typeName(valj) = "ASF_Map") Then
            parts(ii_j + 1) = ValueToStringForPrint(valj)
        Else
            parts(ii_j + 1) = CStr(valj)
        End If
    Next ii_j
    Dim sb As String: sb = parts(1)
    For ii_j = 2 To len_val
        sb = sb & sep & parts(ii_j)
    Next ii_j
    vAssignment tmpResult, sb
    GoTo exitfun
    Exit Function
exitfun:
    ArrTostring = True
End Function

Private Function ArrShift(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim popped2 As Variant
    Dim newA_s() As Variant
    Dim idx_s As Long, k_s As Long
    ' shift: remove first element, mutate array, return popped value
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim ob_s As Long: ob_s = 1
    Dim lb_s As Long: lb_s = LBound(baseValLocal)
    Dim ub_s As Long: ub_s = UBound(baseValLocal)
    If lb_s > ub_s Then
        vAssignment tmpResult, Empty: GoTo exitfun
    End If
    vAssignment popped2, baseValLocal(lb_s)
    ' create new array without first element
    If lb_s = ub_s Then
        ReDim baseValLocal(ob_s To ob_s)
        ' empty arr = single slot by convention
    Else
        ReDim newA_s(ob_s To ob_s + (ub_s - lb_s - 1))
        idx_s = 0
        For k_s = lb_s + 1 To ub_s
            vAssignment newA_s(ob_s + idx_s), baseValLocal(k_s)
            idx_s = idx_s + 1
        Next k_s
        baseValLocal = newA_s
    End If
    vAssignment tmpResult, popped2
    ' write back to original container via ResolvePath
    WriteBack baseValLocal
    GoTo exitfun
    Exit Function
exitfun:
    ArrShift = True
End Function

Private Function ArrSplice(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim startRaw_sp As Long, startN_sp As Long, deleteCount_sp As Long
    Dim removedCount As Long
    Dim tempIn() As Variant, removedArr() As Variant, newTemp() As Variant
    Dim insertCount As Long, newSize As Long
    Dim i_sp As Long, j_sp As Long, cur As Long
    Dim insertedItem As Variant
    Dim removedOut() As Variant
    ' splice(start, deleteCount?, ...items): mutating splice - returns array of deleted elements
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_sp As Long: optBase_sp = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    Dim len_sp As Long: len_sp = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    startRaw_sp = CLng(evaluated(1))
    startN_sp = NormalizeIndex(startRaw_sp, len_sp, optBase_sp) ' zero-based offset
    If evaluated.Count >= 2 Then
        deleteCount_sp = CLng(evaluated(2))
        If deleteCount_sp < 0 Then deleteCount_sp = 0
    Else
        deleteCount_sp = IIf(len_sp - startN_sp > 0, len_sp - startN_sp, 0)
    End If
    removedCount = IIf(len_sp - startN_sp > 0, _
                    IIf(deleteCount_sp < (len_sp - startN_sp), deleteCount_sp, len_sp - startN_sp), 0)
    ' prepare tempIn
    If len_sp > 0 Then
        ReDim tempIn(0 To len_sp - 1)
        For i_sp = 0 To len_sp - 1
            vAssignment tempIn(i_sp), baseValLocal(lb_b + i_sp)
        Next i_sp
    Else
        ReDim tempIn(0 To -1)
    End If
    ' removed elements
    If removedCount > 0 Then
        ReDim removedArr(0 To removedCount - 1)
        For i_sp = 0 To removedCount - 1
            vAssignment removedArr(i_sp), tempIn(startN_sp + i_sp)
        Next i_sp
    Else
        ReDim removedArr(0 To -1)
    End If
    ' insertion count
    insertCount = IIf(evaluated.Count > 2, evaluated.Count - 2, 0)
    newSize = len_sp - removedCount + insertCount
    If newSize > 0 Then
        ReDim newTemp(0 To newSize - 1)
    Else
        ReDim newTemp(0 To -1)
    End If
    cur = 0
    ' copy prefix
    For i_sp = 0 To startN_sp - 1
        If i_sp <= UBound(tempIn) Then
            vAssignment newTemp(cur), tempIn(i_sp)
            cur = cur + 1
        End If
    Next i_sp
    ' inserted items
    For j_sp = 3 To evaluated.Count
        vAssignment insertedItem, evaluated(j_sp)
        vAssignment newTemp(cur), insertedItem
        cur = cur + 1
    Next j_sp
    ' copy suffix
    For i_sp = startN_sp + removedCount To len_sp - 1
        If i_sp <= UBound(tempIn) Then
            vAssignment newTemp(cur), tempIn(i_sp)
            cur = cur + 1
        End If
    Next i_sp
    ' convert newTemp to option-base array
    If newSize <= 0 Then
        ReDim outArr(optBase_sp To optBase_sp)
    Else
        ReDim outArr(optBase_sp To optBase_sp + newSize - 1)
        For i_sp = 0 To newSize - 1
            vAssignment outArr(optBase_sp + i_sp), newTemp(i_sp)
        Next i_sp
    End If
    baseValLocal = outArr
    ' build removedOut
    If removedCount <= 0 Then
        ReDim removedOut(optBase_sp To optBase_sp)
    Else
        ReDim removedOut(optBase_sp To optBase_sp + removedCount - 1)
        For i_sp = 0 To removedCount - 1
            vAssignment removedOut(optBase_sp + i_sp), removedArr(i_sp)
        Next i_sp
    End If
    vAssignment tmpResult, removedOut
    ' write back to original container via ResolvePath
    WriteBack baseValLocal
    GoTo exitfun
    Exit Function
exitfun:
    ArrSplice = True
End Function

Private Function ArrAt(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim optBase_u2 As Long
    Dim len_base As Long
    Dim lb_b As Long, ub_b As Long
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    optBase_u2 = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    Dim idxRaw_sp As Long: idxRaw_sp = CLng(evaluated(1))
    Dim idxN_sp As Long: idxN_sp = NormalizeIndex(idxRaw_sp, len_base, optBase_u2)
    If idxN_sp >= len_base Then vAssignment tmpResult, Empty: GoTo exitfun
    vAssignment tmpResult, baseValLocal(lb_b + idxN_sp)
    GoTo exitfun
    Exit Function
exitfun:
    ArrAt = True
End Function

Private Function ArrCopywithin(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim tempIn() As Variant
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_cw As Long: optBase_cw = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    Dim len_cw As Long: len_cw = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)

    Dim targetRaw_cw As Long: targetRaw_cw = CLng(IIf(evaluated.Count >= 1, evaluated(1), optBase_cw))
    Dim startRaw_cw As Long: startRaw_cw = CLng(IIf(evaluated.Count >= 2, evaluated(2), optBase_cw))
    Dim endRaw_cw As Long: endRaw_cw = CLng(IIf(evaluated.Count >= 3, evaluated(3), optBase_cw + len_cw))

    Dim targN As Long: targN = NormalizeIndex(targetRaw_cw, len_cw, optBase_cw)
    Dim startN_cw As Long: startN_cw = NormalizeIndex(startRaw_cw, len_cw, optBase_cw)
    Dim endN_cw As Long: endN_cw = NormalizeIndex(endRaw_cw, len_cw, optBase_cw)
    If endN_cw < startN_cw Then endN_cw = startN_cw
    Dim srcLen_cw As Long: srcLen_cw = endN_cw - startN_cw
    Dim i_cw As Long
    ' build tempIn from baseValLocal
    If len_cw > 0 Then
        ReDim tempIn(0 To len_cw - 1)
        For i_cw = 0 To len_cw - 1
            vAssignment tempIn(i_cw), baseValLocal(lb_b + i_cw)
        Next i_cw
    Else
        ReDim tempIn(0 To -1)
    End If
    ' build chunk to copy (length = srcLen_cw)
    If srcLen_cw <= 0 Then
        ReDim tempArr(0 To -1)
    Else
        ReDim tempArr(0 To srcLen_cw - 1)
        For i_cw = 0 To srcLen_cw - 1
            If (startN_cw + i_cw) <= UBound(tempIn) Then
                vAssignment tempArr(i_cw), tempIn(startN_cw + i_cw)
            Else
                tempArr(i_cw) = Empty
            End If
        Next i_cw
    End If
    Dim j_cw As Long
    Dim destPos As Long
    ' copy into destination (mutating)
    For j_cw = 0 To UBound(tempArr)
        destPos = targN + j_cw
        If destPos >= 0 And destPos <= UBound(tempIn) Then
            vAssignment tempIn(destPos), tempArr(j_cw)
        End If
    Next j_cw
    ' convert tempIn back to option-base array
    If len_cw <= 0 Then
        ReDim outArr(optBase_cw To optBase_cw)
    Else
        ReDim outArr(optBase_cw To optBase_cw + len_cw - 1)
        For i_cw = 0 To len_cw - 1
            vAssignment outArr(optBase_cw + i_cw), tempIn(i_cw)
        Next i_cw
    End If
    baseValLocal = outArr
    vAssignment tmpResult, baseValLocal
    ' write back to original container via ResolvePath
    WriteBack baseValLocal
    GoTo exitfun
    Exit Function
exitfun:
    ArrCopywithin = True
End Function

Private Function ArrEntries(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim len_base As Long
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim newTemp() As Variant
    Dim i_sp As Long
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim ob_entries As Long: ob_entries = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    If len_base <= 0 Then
        ReDim outArr(ob_entries To ob_entries)
        vAssignment tmpResult, outArr
        GoTo exitfun
    End If
    ReDim newTemp(0 To len_base - 1)
    For i_sp = 0 To len_base - 1
        Dim pairArr_local() As Variant
        ReDim pairArr_local(0 To 1)
        vAssignment pairArr_local(0), ob_entries + i_sp
        vAssignment pairArr_local(1), baseValLocal(lb_b + i_sp)
        vAssignment newTemp(i_sp), pairArr_local
    Next i_sp
    ReDim outArr(ob_entries To ob_entries + len_base - 1)
    For i_sp = 0 To len_base - 1
        vAssignment outArr(ob_entries + i_sp), newTemp(i_sp)
    Next i_sp
    vAssignment tmpResult, outArr
    GoTo exitfun
    Exit Function
exitfun:
    ArrEntries = True
End Function

Private Function ArrEvery(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim len_base As Long
    Dim lb_b As Long, ub_b As Long
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim cb_every As Variant: vAssignment cb_every, evaluated(1)
    If typeName(cb_every) <> "ASF_Map" Or cb_every.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    Dim ob_ev As Long: ob_ev = 1
    Dim idx_ev As Long
    Dim i_ev As Long
    Dim args_ev As Collection
    Dim res_ev As Variant
    For i_ev = 0 To len_base - 1
        Set args_ev = New Collection
        args_ev.Add baseValLocal(lb_b + i_ev)
        args_ev.Add (ob_ev + idx_ev)
        args_ev.Add baseValLocal
        vAssignment tmpMapObj, cb_every
        vAssignment res_ev, CallClosure(tmpMapObj, args_ev, baseValLocal)
        If Not IsTruthy(res_ev) Then
            vAssignment tmpResult, False: GoTo exitfun
        End If
        idx_ev = idx_ev + 1
    Next i_ev
    vAssignment tmpResult, True
    GoTo exitfun
    Exit Function
exitfun:
    ArrEvery = True
End Function

Private Function ArrFind(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim len_base As Long
    Dim lb_b As Long, ub_b As Long
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim cb_find As Variant: vAssignment cb_find, evaluated(1)
    If typeName(cb_find) <> "ASF_Map" Or cb_find.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim ob_f As Long: ob_f = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)

    If propN = "findlast" Or propN = "findlastindex" Then
        Dim i_f As Long
        Dim args_f As Collection
        Dim res_f As Variant
        For i_f = len_base - 1 To 0 Step -1
            Set args_f = New Collection
            args_f.Add baseValLocal(lb_b + i_f)
            args_f.Add (ob_f + i_f)
            args_f.Add baseValLocal
            vAssignment tmpMapObj, cb_find
            vAssignment res_f, CallClosure(tmpMapObj, args_f, baseValLocal)
            If IsTruthy(res_f) Then
                If propN = "findlast" Then vAssignment tmpResult, baseValLocal(lb_b + i_f) Else vAssignment tmpResult, (ob_f + i_f)
                GoTo exitfun
            End If
        Next i_f
        If propN = "findlast" Then vAssignment tmpResult, Empty Else vAssignment tmpResult, -1
        GoTo exitfun
    Else
        Dim args_f2 As Collection
        Dim res_f2 As Variant
        For i_f = 0 To len_base - 1
            Set args_f2 = New Collection
            args_f2.Add baseValLocal(lb_b + i_f)
            args_f2.Add (ob_f + i_f)
            args_f2.Add baseValLocal
            vAssignment tmpMapObj, cb_find
            vAssignment res_f2, CallClosure(tmpMapObj, args_f2, baseValLocal)
            If IsTruthy(res_f2) Then
                If propN = "find" Then vAssignment tmpResult, baseValLocal(lb_b + i_f) Else vAssignment tmpResult, (ob_f + i_f)
                GoTo exitfun
            End If
        Next i_f
        If propN = "find" Then vAssignment tmpResult, Empty Else vAssignment tmpResult, -1
        GoTo exitfun
    End If
    Exit Function
exitfun:
    ArrFind = True
End Function

Private Function ArrFrom(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim argColl As Collection
    Dim tmpMapObj As ASF_Map
    Dim thisVal As Variant
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim tempIn() As Variant
    Dim newSize As Long
    Dim i_sp As Long
    If evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim val_from As Variant: vAssignment val_from, evaluated(1)
    Dim opt_from As Long: opt_from = 1

    ' Build a zero-based tempIn[] with the source elements and also
    ' build srcArrOption (option-base array)
    If IsArray(val_from) Then

        ' copy source into option-base array and tempIn
        Dim srcArrOption As Variant
        vAssignment srcArrOption, ShallowCopyArray(val_from, opt_from)
        lb_b = LBound(srcArrOption)
        ub_b = UBound(srcArrOption)
        newSize = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
        If newSize > 0 Then
            ReDim tempIn(0 To newSize - 1)
            For i_sp = 0 To newSize - 1
                tempIn(i_sp) = srcArrOption(lb_b + i_sp)
            Next i_sp
        Else
            ReDim tempIn(0 To -1)
        End If

    ElseIf typeName(val_from) = "String" Then
        ' string -> array of characters (option base)
        If Len(val_from) = 0 Then
            ReDim outArr(opt_from To opt_from)
            vAssignment tmpResult, outArr
            GoTo exitfun
        End If
        newSize = Len(val_from)
        ReDim tempIn(0 To newSize - 1)
        ReDim srcArrOption(opt_from To opt_from + newSize - 1)
        For i_sp = 1 To newSize
            tempIn(i_sp - 1) = mid$(val_from, i_sp, 1)
            srcArrOption(opt_from + i_sp - 1) = tempIn(i_sp - 1)
        Next i_sp

    Else
        ' single value -> one-element source
        newSize = 1
        ReDim tempIn(0 To 0)
        tempIn(0) = val_from
        ReDim srcArrOption(opt_from To opt_from)
        srcArrOption(opt_from) = val_from
    End If

    ' optional mapping closure (mapFn) and optional thisArg
    Dim comp As Variant
    If evaluated.Count >= 2 Then vAssignment comp, evaluated(2)
    If evaluated.Count >= 3 Then vAssignment thisVal, evaluated(3)

    ' If mapping function is a closure, apply it; otherwise just copy
    If newSize <= 0 Then
        ReDim outArr(opt_from To opt_from)
        vAssignment tmpResult, outArr
        GoTo exitfun
    End If

    ReDim outArr(opt_from To opt_from + newSize - 1)
    Dim mapResult As Variant

    If Not IsEmpty(comp) And Not IsNull(comp) Then
        ' only apply mapping if comp is a Closure map object
        If typeName(comp) = "ASF_Map" Then
            If comp.Exists("type") And CStr(comp.GetValue("type")) = "Closure" Then
                For i_sp = 0 To newSize - 1
                    Set argColl = New Collection
                    argColl.Add tempIn(i_sp)
                    argColl.Add (opt_from + i_sp)
                    argColl.Add srcArrOption
                    vAssignment tmpMapObj, comp
                    vAssignment mapResult, CallClosure(tmpMapObj, argColl, thisVal)
                    outArr(opt_from + i_sp) = mapResult
                Next i_sp
                vAssignment tmpResult, outArr
                GoTo exitfun
            End If
        End If
    End If

    ' fallback -> direct copy
    For i_sp = 0 To newSize - 1
        outArr(opt_from + i_sp) = tempIn(i_sp)
    Next i_sp
    vAssignment tmpResult, outArr
    GoTo exitfun
    Exit Function
exitfun:
    ArrFrom = True
End Function

Private Function ArrIncludes(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim len_base As Long
    Dim lb_b As Long, ub_b As Long
    Dim foundIdx As Long
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim needle As Variant: vAssignment needle, evaluated(1)
    Dim optBase_in As Long: optBase_in = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    foundIdx = -1
    Dim i_in As Long
    If propN = "lastindexof" Then
        For i_in = len_base - 1 To 0 Step -1
            If DeepEqual(baseValLocal(lb_b + i_in), needle) Then
                foundIdx = optBase_in + i_in
                Exit For
            End If
        Next i_in
    Else
        For i_in = 0 To len_base - 1
            If DeepEqual(baseValLocal(lb_b + i_in), needle) Then
                foundIdx = optBase_in + i_in
                Exit For
            End If
        Next i_in
    End If
    If propN = "includes" Then
        vAssignment tmpResult, (foundIdx <> -1)
    Else
        If foundIdx = -1 Then vAssignment tmpResult, -1 Else vAssignment tmpResult, foundIdx
    End If
    GoTo exitfun
    Exit Function
exitfun:
    ArrIncludes = True
End Function

Private Function ArrOf(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim outArr() As Variant
    Dim newSize As Long
    Dim i_sp As Long
    Dim optBase_of As Long: optBase_of = 1
    newSize = evaluated.Count
    If newSize <= 0 Then
        ReDim outArr(optBase_of To optBase_of)
    Else
        ReDim outArr(optBase_of To optBase_of + newSize - 1)
        For i_sp = 1 To evaluated.Count
            vAssignment outArr(optBase_of + i_sp - 1), evaluated(i_sp)
        Next i_sp
    End If
    vAssignment tmpResult, outArr
    GoTo exitfun
    Exit Function
exitfun:
    ArrOf = True
End Function

Private Function ArrReverse(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim len_base As Long
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim newTemp() As Variant
    Dim i_sp As Long
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_rev As Long: optBase_rev = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    If len_base <= 0 Then
        ReDim outArr(optBase_rev To optBase_rev)
        vAssignment tmpResult, outArr
        GoTo exitfun
    End If
    ReDim newTemp(0 To len_base - 1)
    For i_sp = 0 To len_base - 1
        newTemp(i_sp) = baseValLocal(lb_b + (len_base - 1 - i_sp))
    Next i_sp
    ReDim outArr(optBase_rev To optBase_rev + len_base - 1)
    For i_sp = 0 To len_base - 1
        outArr(optBase_rev + i_sp) = newTemp(i_sp)
    Next i_sp
    baseValLocal = outArr
    vAssignment tmpResult, baseValLocal
    ' write back to original container via ResolvePath
    WriteBack baseValLocal
    GoTo exitfun
    Exit Function
exitfun:
    ArrReverse = True
End Function

Private Function ArrToreversed(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim len_base As Long
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim newTemp() As Variant
    Dim i_sp As Long
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_tr As Long: optBase_tr = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    If len_base <= 0 Then
        ReDim outArr(optBase_tr To optBase_tr)
        vAssignment tmpResult, outArr
        GoTo exitfun
    End If
    ReDim newTemp(0 To len_base - 1)
    For i_sp = 0 To len_base - 1
        newTemp(i_sp) = baseValLocal(lb_b + (len_base - 1 - i_sp))
    Next i_sp
    ReDim outArr(optBase_tr To optBase_tr + len_base - 1)
    For i_sp = 0 To len_base - 1
        outArr(optBase_tr + i_sp) = newTemp(i_sp)
    Next i_sp
    vAssignment tmpResult, outArr
    GoTo exitfun
    Exit Function
exitfun:
    ArrToreversed = True
End Function

Private Function ArrSome(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    Dim len_base As Long
    Dim lb_b As Long, ub_b As Long
    If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim cb_some As Variant: vAssignment cb_some, evaluated(1)
    If typeName(cb_some) <> "ASF_Map" Or cb_some.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    Dim ob_som As Long: ob_som = 1
    Dim i_som As Long
    Dim args_s As Collection
    Dim res_s As Variant
    For i_som = 0 To len_base - 1
        Set args_s = New Collection
        args_s.Add baseValLocal(lb_b + i_som)
        args_s.Add (ob_som + i_som)
        args_s.Add baseValLocal
        vAssignment tmpMapObj, cb_some
        vAssignment res_s, CallClosure(tmpMapObj, args_s, baseValLocal)
        If IsTruthy(res_s) Then
            vAssignment tmpResult, True: GoTo exitfun
        End If
    Next i_som
    vAssignment tmpResult, False
    GoTo exitfun
    Exit Function
exitfun:
    ArrSome = True
End Function

Private Function ArrWith(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim lb_b As Long, ub_b As Long
    If Not IsArray(baseValLocal) Or evaluated.Count < 2 Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_w As Long: optBase_w = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    Dim len_w As Long: len_w = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
    Dim idxRaw_w As Long: idxRaw_w = CLng(evaluated(1))
    Dim idxN_w As Long: idxN_w = NormalizeIndex(idxRaw_w, len_w, optBase_w)
    Dim copy_w As Variant: vAssignment copy_w, ShallowCopyArray(baseValLocal, optBase_w)
    Dim baseL_w As Long: baseL_w = LBound(copy_w)
    vAssignment copy_w(baseL_w + idxN_w), evaluated(2)
    vAssignment tmpResult, copy_w
    GoTo exitfun
    Exit Function
exitfun:
    ArrWith = True
End Function

Private Function ArrSort(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim comp As Variant
    Dim thisVal As Variant
    Dim outArr() As Variant
    Dim lb_b As Long, ub_b As Long
    Dim tempIn() As Variant
    Dim i_sp As Long
    If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
    Dim optBase_srt As Long: optBase_srt = 1
    lb_b = LBound(baseValLocal)
    ub_b = UBound(baseValLocal)
    Dim n_s As Long: n_s = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)

    ' get comparator if any
    If evaluated.Count >= 1 Then vAssignment comp, evaluated(1) Else comp = Empty

    If n_s <= 1 Then
        If propN = "sort" Then
            ' ensure mutation writes back even if single element (keeps behavior consistent)
            WriteBack baseValLocal
            vAssignment tmpResult, baseValLocal
        Else
            vAssignment tmpResult, baseValLocal
        End If
        GoTo exitfun
    End If

    ' build zero-based tempIn array
    ReDim tempIn(0 To n_s - 1)
    For i_sp = 0 To n_s - 1
        vAssignment tempIn(i_sp), baseValLocal(lb_b + i_sp)
    Next i_sp

    ' determine if comp is a closure
    Dim hasComp As Boolean: hasComp = False
    If Not IsEmpty(comp) And Not IsNull(comp) Then
        If typeName(comp) = "ASF_Map" Then
            If comp.Exists("type") And CStr(comp.GetValue("type")) = "Closure" Then hasComp = True
        End If
    End If

    ' quicksort in-place on tempIn. Pass baseValLocal as thisVal (so comparator gets same this as other calls)
    Call QS_QuickSort(tempIn, 0, n_s - 1, comp, hasComp, baseValLocal)

    ' convert sorted tempIn back to option-base array
    ReDim outArr(optBase_srt To optBase_srt + n_s - 1)
    For i_sp = 0 To n_s - 1
        vAssignment outArr(optBase_srt + i_sp), tempIn(i_sp)
    Next i_sp

    If propN = "sort" Then
        baseValLocal = outArr
        ' persist mutation to original container (LValue)
        WriteBack baseValLocal
        vAssignment tmpResult, baseValLocal
    Else
        vAssignment tmpResult, outArr
    End If

    GoTo exitfun
    Exit Function
exitfun:
    ArrSort = True
End Function


' Regex methods
Private Function RegexMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Expose regex module methods with the Regex object
    Select Case propN
        Case "escape"
            tmpResult = baseValLocal.Escape(CStr(evaluated(1)))
            GoTo exitfun
        Case "exec"
            tmpResult = RegexCapsToVariantArray(baseValLocal.Exec(CStr(evaluated(1))))
            GoTo exitfun
        Case "execall"
            tmpResult = RegexMatchesToVariantArray(baseValLocal.ExecAll(CStr(evaluated(1))))
            GoTo exitfun
        Case "execat"
            tmpResult = RegexCapsToVariantArray(baseValLocal.ExecAt(CStr(evaluated(1)), CLng(evaluated(2))))
            GoTo exitfun
        Case "init"
            On Error GoTo init_Error_Handler
            Select Case evaluated.Count
                Case 1
                    ' pattern
                    baseValLocal.Init CStr(evaluated(1))
                Case Is <= 2
                    ' pattern, ignoreCase
                    baseValLocal.Init CStr(evaluated(1)), CBool(evaluated(2))
                Case Is <= 3
                    ' pattern, ignoreCase flag, multiline flag
                    baseValLocal.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3))
                Case Is <= 4
                    ' pattern, ignoreCase flag, multiline flag, dotAll flag
                    baseValLocal.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3)), CBool(evaluated(4))
                Case Is <= 5
                    ' pattern, ignoreCase flag, multiline flag, dotAll flag, global flag
                    baseValLocal.Init CStr(evaluated(1)), CBool(evaluated(2)), , CBool(evaluated(3)), CBool(evaluated(4)), CBool(evaluated(5))
            End Select
            tmpResult = True
init_Return:
            GoTo exitfun
init_Error_Handler:
            tmpResult = False
            Resume init_Return
        Case "getpattern"
            tmpResult = baseValLocal.pattern
            GoTo exitfun
        Case "setpattern"
            baseValLocal.pattern = CStr(evaluated(1))
            GoTo exitfun
        Case "getdotall"
            tmpResult = baseValLocal.dotAll
            GoTo exitfun
        Case "getglobal"
            tmpResult = baseValLocal.globalMatch
            GoTo exitfun
        Case "setglobal"
            baseValLocal.globalMatch = CBool(evaluated(1))
            GoTo exitfun
        Case "setdotall"
            baseValLocal.dotAll = CBool(evaluated(1))
            GoTo exitfun
        Case "getignorecase"
            tmpResult = baseValLocal.ignoreCase
            GoTo exitfun
        Case "setignorecase"
            baseValLocal.ignoreCase = CBool(evaluated(1))
            GoTo exitfun
        Case "getmaxmatchsteps"
            tmpResult = baseValLocal.MaxMatchSteps
            GoTo exitfun
        Case "setmaxmatchsteps"
            baseValLocal.MaxMatchSteps = CLng(evaluated(1))
            GoTo exitfun
        Case "getmultiline"
            tmpResult = baseValLocal.multiline
            GoTo exitfun
        Case "setmultiline"
            baseValLocal.multiline = CBool(evaluated(1))
            GoTo exitfun
        Case "replace"
            tmpResult = baseValLocal.Replace(CStr(evaluated(1)), CStr(evaluated(2)))
            GoTo exitfun
        Case "split"
            tmpResult = RegexCapsToVariantArray(baseValLocal.Split(CStr(evaluated(1))))
            GoTo exitfun
        Case "test"
            tmpResult = baseValLocal.test(CStr(evaluated(1)))
            GoTo exitfun
    End Select
    Exit Function
exitfun:
    RegexMethod = True
End Function

' String methods
Private Function StringMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim nidx As Long
    Dim eidx As Long
    Dim strOB As Long
    'Builtins string methods
    Select Case propN
        Case "length"
            tmpResult = Len(baseValLocal)
            GoTo exitfun
        Case "at"
            If evaluated.Count Then
                nidx = NormalizeIndex(CLng(evaluated(1)), Len(baseValLocal), strOB) + 1
            Else
                nidx = 1
            End If
            tmpResult = mid(CStr(baseValLocal), nidx, 1)
            GoTo exitfun
        Case "charat"
            If evaluated.Count Then
                nidx = CLng(evaluated(1)) + 1
            Else
                nidx = 1
            End If
            tmpResult = mid(CStr(baseValLocal), nidx, 1)
            GoTo exitfun
        Case "charcodeat"
            If evaluated.Count Then
                nidx = CLng(evaluated(1)) + 1
            Else
                nidx = 1
            End If
            tmpResult = AscW(mid(CStr(baseValLocal), nidx, 1))
            GoTo exitfun
        Case "concat"
            StringMethod = StrConcat(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "endswith"
            StringMethod = StrEndswith(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "fromcharcode"
            tmpResult = ChrW$(CLng(evaluated(1)))
            GoTo exitfun
        Case "includes"
            StringMethod = StrIncludes(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "indexof"
            StringMethod = StrIndexof(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "lastindexof"
            StringMethod = StrLastindexof(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "localecompare"
            ' A tiny VBA based implementation, does not behave like javascript
            tmpResult = StrComp(baseValLocal, evaluated(1), vbBinaryCompare)
            GoTo exitfun
        Case "match", "matchall"
            StringMethod = StrMatch(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "padend"
            StringMethod = StrPadend(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "padstart"
            StringMethod = StrPadstart(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "repeat"
            StringMethod = StrRepeat(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "replace", "replaceall"
            StringMethod = StrReplace(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "slice"
            nidx = NormalizeIndex(CLng(evaluated(1)), Len(baseValLocal), strOB) + 1
            Select Case evaluated.Count
                Case 1
                    tmpResult = mid(CStr(baseValLocal), nidx)
                Case 2
                    eidx = NormalizeIndex(CLng(evaluated(2)), Len(baseValLocal), strOB) + 1
                    tmpResult = mid(CStr(baseValLocal), nidx, eidx - nidx)
            End Select
            GoTo exitfun
        Case "split"
            StringMethod = StrSplit(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "startswith"
            StringMethod = StrStartswith(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "substring"
            StringMethod = StrSubstring(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "tolowercase"
            tmpResult = LCase$(baseValLocal)
            GoTo exitfun
        Case "tostring", "valueof"
            tmpResult = CStr(baseValLocal)
            GoTo exitfun
        Case "touppercase"
            tmpResult = UCase$(baseValLocal)
            GoTo exitfun
        Case "trim"
            tmpResult = Trim$(baseValLocal)
            GoTo exitfun
        Case "trimstart"
            tmpResult = LTrim$(baseValLocal)
            GoTo exitfun
        Case "trimend"
            tmpResult = RTrim$(baseValLocal)
            GoTo exitfun
    End Select
    Exit Function
exitfun:
    StringMethod = True
End Function

Private Function StrConcat(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim concChr As String
    Dim ccElm As Variant
    If evaluated.Count > 1 Then
        concChr = evaluated(1)
    Else
        concChr = vbNullString
    End If
    tmpResult = CStr(baseValLocal)
    For Each ccElm In evaluated
        If ccElm <> concChr Then
            tmpResult = tmpResult & concChr & ccElm
        End If
    Next ccElm
    GoTo exitfun
    Exit Function
exitfun:
    StrConcat = True
End Function

Private Function StrEndswith(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim endStr As String
    Dim ewComStr As String
    ewComStr = CStr(evaluated(1))
    endStr = right$(baseValLocal, Len(ewComStr))
    tmpResult = (endStr = ewComStr)
    GoTo exitfun
    Exit Function
exitfun:
    StrEndswith = True
End Function

Private Function StrIncludes(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim icStartPos As Long
    If evaluated.Count > 1 Then
        icStartPos = CLng(evaluated(2)) + 1
    Else
        icStartPos = 1
    End If
    tmpResult = (InStrB(icStartPos, baseValLocal, evaluated(1), vbBinaryCompare) > 0)
    GoTo exitfun
    Exit Function
exitfun:
    StrIncludes = True
End Function

Private Function StrIndexof(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim ioStartPos As Long
    If evaluated.Count > 1 Then
        ioStartPos = CLng(evaluated(2)) + 1
    Else
        ioStartPos = 1
    End If
    tmpResult = InStr(ioStartPos, baseValLocal, evaluated(1), vbBinaryCompare) - 1
    GoTo exitfun
    Exit Function
exitfun:
    StrIndexof = True
End Function

Private Function StrLastindexof(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim lioStartPos As Long
    Dim tmpLIOres As Long
    Dim LIOprev As Long
    If evaluated.Count > 1 Then
        lioStartPos = CLng(evaluated(2)) + 1
    Else
        lioStartPos = 1
    End If
    tmpLIOres = InStr(lioStartPos, baseValLocal, evaluated(1), vbBinaryCompare)
    Do While tmpLIOres
        LIOprev = tmpLIOres
        tmpLIOres = InStr(tmpLIOres + 1, baseValLocal, evaluated(1), vbBinaryCompare)
    Loop
    tmpResult = LIOprev - 1
    GoTo exitfun
    Exit Function
exitfun:
    StrLastindexof = True
End Function

Private Function StrMatch(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim foundIdx As Long
    Dim sourceStr As String: sourceStr = CStr(baseValLocal)
    Dim searchArg As Variant: searchArg = evaluated(1)
    Dim isSlashRegex As Boolean: isSlashRegex = VM_IsSlashRegexString(searchArg)
    If Not isSlashRegex Then
        ' treat as plain string -> match: if exact match return array [full] else Nothing / empty
        If propN = "matchall" Then
            Dim matchesList() As Variant: ReDim matchesList(1 To 1)
            ' naive: find all occurrences of the literal substring
            Dim startAt2 As Long: startAt2 = 1
            foundIdx = 0
            Do
                Dim pFind As Long: pFind = InStr(startAt2, sourceStr, CStr(searchArg), vbBinaryCompare)
                If pFind = 0 Then Exit Do
                foundIdx = foundIdx + 1
                ReDim Preserve matchesList(1 To foundIdx)
                matchesList(foundIdx) = mid$(sourceStr, pFind, Len(CStr(searchArg)))
                startAt2 = pFind + Len(CStr(searchArg))
                If Len(CStr(searchArg)) = 0 Then startAt2 = startAt2 + 1
            Loop
            tmpResult = matchesList
            GoTo exitfun
        Else
            Dim pOne As Long: pOne = InStr(1, sourceStr, CStr(searchArg), vbBinaryCompare)
            If pOne = 0 Then
                tmpResult = Empty
            Else
                tmpResult = Array(mid$(sourceStr, pOne, Len(CStr(searchArg))))
            End If
            GoTo exitfun
        End If
    Else
        Dim pat2 As String, f2 As String
        VM_ParseSlashRegex CStr(searchArg), pat2, f2
        Dim fi_i As Boolean: fi_i = (InStr(f2, "i") > 0)
        Dim fi_m As Boolean: fi_m = (InStr(f2, "m") > 0)
        Dim fi_s As Boolean: fi_s = (InStr(f2, "s") > 0)
        Dim fi_g As Boolean: fi_g = (InStr(f2, "g") > 0)

        Dim r2 As New ASF_RegexEngine
        Set r2.SharedCache = GLOBALS_.regexCache
        r2.Init pat2, fi_i, GLOBALS_.re.MaxMatchSteps, fi_m, fi_s

        ' If caller asked for matchAll method
        If LCase$(propN) = "matchall" Then
            ' matchAll requires global flag
            If Not fi_g Then
                err.Raise vbObjectError + 102, "VM", "matchAll requires the 'g' flag on the regular expression"
            End If

            ' Collect full match + groups arrays for every non-overlapping match
            Dim resultsColl As New Collection
            Dim sPosA As Long: sPosA = 1
            Dim Lsrc As Long: Lsrc = Len(sourceStr)
            Do While sPosA <= Lsrc + 1
                Dim capsA As Collection
                Dim tryStartA As Long
                Dim foundAtA As Long: foundAtA = 0
                For tryStartA = sPosA To Lsrc + 1
                    Set capsA = r2.ExecAt(sourceStr, tryStartA)
                    If Not (capsA Is Nothing) Then
                        foundAtA = tryStartA
                        Exit For
                    End If
                Next tryStartA

                If foundAtA = 0 Then Exit Do

                ' Append full match + groups as an array
                resultsColl.Add VM_CapturesToArray(capsA)

                Dim fullMatchA As String: fullMatchA = CStr(capsA.Item(1))
                Dim fLenA As Long: fLenA = Len(fullMatchA)
                If fLenA = 0 Then
                    sPosA = foundAtA + 1
                Else
                    sPosA = foundAtA + fLenA
                End If
            Loop

            If resultsColl.Count = 0 Then
                tmpResult = Empty
            Else
                ' convert Collection-of-arrays to a Variant() array for caller
                Dim outArrA() As Variant
                Dim iA As Long: iA = 0
                Dim itA As Variant
                For Each itA In resultsColl
                    iA = iA + 1
                    ReDim Preserve outArrA(1 To iA)
                    outArrA(iA) = itA
                Next itA
                tmpResult = outArrA
            End If
            GoTo exitfun
        End If

        ' Caller did not ask matchAll; determine global (g) behavior
        If fi_g Then
            ' Global match: return array of full-match strings, groups ignored
            Dim outMatches() As Variant
            Dim sPos As Long: sPos = 1
            Dim L2 As Long: L2 = Len(sourceStr)
            Dim idxOut As Long: idxOut = 0

            Do While sPos <= L2 + 1
                Dim caps2 As Collection
                Dim sTry As Long
                Dim foundAt2 As Long: foundAt2 = 0
                For sTry = sPos To L2 + 1
                    Set caps2 = r2.ExecAt(sourceStr, sTry)
                    If Not (caps2 Is Nothing) Then
                        foundAt2 = sTry
                        Exit For
                    End If
                Next sTry

                If foundAt2 = 0 Then Exit Do

                idxOut = idxOut + 1
                ReDim Preserve outMatches(1 To idxOut)
                outMatches(idxOut) = CStr(caps2.Item(1))

                Dim fLen As Long: fLen = Len(outMatches(idxOut))
                If fLen = 0 Then
                    sPos = foundAt2 + 1
                Else
                    sPos = foundAt2 + fLen
                End If
            Loop

            If idxOut = 0 Then
                tmpResult = Empty
            Else
                tmpResult = outMatches
            End If
            GoTo exitfun
        Else
            ' Non-global match: return first full match + groups (array), like JS String.prototype.match without g
            Dim caps3 As Collection
            Set caps3 = r2.Exec(sourceStr)
            If caps3 Is Nothing Then
                tmpResult = Empty
            Else
                tmpResult = VM_CapturesToArray(caps3)
            End If
            GoTo exitfun
        End If
    End If
    Exit Function
exitfun:
    StrMatch = True
End Function

Private Function StrPadend(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim pePadStr As String
    Dim tpeLen As Long
    Dim cpeLen As Long
    If evaluated.Count > 1 Then
        pePadStr = CStr(evaluated(2))
    Else
        pePadStr = " "
    End If
    tpeLen = CLng(evaluated(1))
    cpeLen = Len(baseValLocal)
    If tpeLen > cpeLen Then
        tmpResult = baseValLocal & makeStringForPadding(pePadStr, tpeLen - cpeLen)
    Else
        tmpResult = baseValLocal
    End If
    GoTo exitfun
    Exit Function
exitfun:
    StrPadend = True
End Function

Private Function StrPadstart(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim psPadStr As String
    Dim tpsLen As Long
    Dim cpsLen As Long
    If evaluated.Count > 1 Then
        psPadStr = CStr(evaluated(2))
    Else
        psPadStr = " "
    End If
    tpsLen = CLng(evaluated(1))
    cpsLen = Len(baseValLocal)
    If tpsLen > cpsLen Then
        tmpResult = makeStringForPadding(psPadStr, tpsLen - cpsLen) & baseValLocal
    Else
        tmpResult = baseValLocal
    End If
    GoTo exitfun
    Exit Function
exitfun:
    StrPadstart = True
End Function

Private Function StrRepeat(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim repCount As Long
    Dim repStr As String
    Dim rpI As Long
    repCount = CLng(evaluated(1))
    For rpI = 1 To repCount
        repStr = repStr & baseValLocal
    Next rpI
    tmpResult = repStr
    GoTo exitfun
    Exit Function
exitfun:
    StrRepeat = True
End Function

Private Function StrReplace(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim argColl As Collection
    Dim tmpMapObj As ASF_Map
    ' replace(searchValue, replaceValue)  -> first occurrence
    ' replaceAll(searchValue, replaceValue) -> all occurrences
    If evaluated.Count = 0 Then
        tmpResult = CStr(baseValLocal)
        GoTo exitfun
    End If

    Dim rawSearch As Variant: vAssignment rawSearch, evaluated(1)
    Dim isRegex As Boolean: isRegex = VM_IsSlashRegexString(rawSearch)

    Dim src As String: src = CStr(baseValLocal)
    Dim replRaw As Variant
    If evaluated.Count >= 2 Then vAssignment replRaw, evaluated(2) Else replRaw = Empty

    Dim isGlobal As Boolean: isGlobal = (propN = "replaceall")

    If isRegex Then
        Dim pattern As String, flags As String
        VM_ParseSlashRegex CStr(rawSearch), pattern, flags
        Dim flag_i As Boolean: flag_i = (InStr(flags, "i") > 0)
        Dim flag_m As Boolean: flag_m = (InStr(flags, "m") > 0)
        Dim flag_s As Boolean: flag_s = (InStr(flags, "s") > 0)
        Dim flag_g As Boolean: flag_g = (InStr(flags, "g") > 0)
        If isGlobal And Not flag_g Then
            err.Raise vbObjectError + 103, "VM", "replaceAll must be called with a global regular expression (missing 'g' flag)"
        End If

        Dim doGlobal As Boolean: doGlobal = flag_g Or isGlobal

        With GLOBALS_.re
            ' Init with flags (ignoreCase, maxSteps, multiline, dotAll)
            .Init pattern, flag_i, .MaxMatchSteps, flag_m, flag_s

            Dim outRepStr As String: outRepStr = vbNullString
            Dim L As Long: L = Len(src)
            Dim searchPos As Long: searchPos = 1

            Do While searchPos <= L + 1
                Dim caps As Collection
                Dim startTry As Long
                Dim foundAt As Long: foundAt = 0

                ' locate the next match by trying ExecAt at increasing start positions
                For startTry = searchPos To L + 1
                    Set caps = .ExecAt(src, startTry)
                    If Not (caps Is Nothing) Then
                        foundAt = startTry
                        Exit For
                    End If
                Next startTry

                If foundAt = 0 Then
                    ' append remainder and break
                    If searchPos <= L Then outRepStr = outRepStr & mid$(src, searchPos)
                    Exit Do
                Else
                    Dim fullMatch As String: fullMatch = CStr(caps.Item(1))
                    Dim matchLen As Long: matchLen = Len(fullMatch)
                    Dim matchStart As Long: matchStart = foundAt
                    Dim matchEnd As Long: matchEnd = matchStart + matchLen - 1

                    ' append chunk before match
                    If matchStart > searchPos Then outRepStr = outRepStr & mid$(src, searchPos, matchStart - searchPos)

                    ' compute replacement text
                    Dim replacedText As String: replacedText = vbNullString

                    ' --- replacement is a function/closure? ---
                    If Not IsEmpty(replRaw) And (IsObject(replRaw) Or typeName(replRaw) = "ASF_Map") Then
                        Set argColl = New Collection
                        argColl.Add fullMatch
                        Dim gi As Long
                        For gi = 2 To caps.Count
                            argColl.Add CStr(caps.Item(gi))
                        Next gi
                        ' offset - JS uses zero-based
                        argColl.Add (matchStart - 1)
                        argColl.Add src
                        vAssignment tmpMapObj, replRaw
                        Dim cbRes As Variant
                        vAssignment cbRes, CallClosure(tmpMapObj, argColl)
                        If IsObject(cbRes) Then
                            replacedText = ValueToStringForPrint(cbRes)
                        ElseIf IsNull(cbRes) Then
                            replacedText = "null"
                        ElseIf IsEmpty(cbRes) Then
                            replacedText = vbNullString
                        Else
                            replacedText = CStr(cbRes)
                        End If

                    Else
                        ' replacement is a string: expand $ sequences (JS-like)
                        Dim replStr As String: replStr = vbNullString
                        If Not IsEmpty(replRaw) Then replStr = CStr(replRaw)
                        replacedText = VM_ExpandReplacementFromCapturesJS(caps, replStr, src, matchStart, matchEnd)
                    End If

                    outRepStr = outRepStr & replacedText

                    ' advance search position; a zero-length match keeps the char at its position
                    If matchLen = 0 Then
                        If matchStart <= L Then outRepStr = outRepStr & mid$(src, matchStart, 1)
                        searchPos = matchStart + 1
                    Else
                        searchPos = matchEnd + 1
                    End If
                    If Not doGlobal Then
                        ' append remainder and exit
                        If searchPos <= L Then outRepStr = outRepStr & mid$(src, searchPos)
                        Exit Do
                    End If
                End If
            Loop

            vAssignment tmpResult, outRepStr
            GoTo exitfun
        End With
    Else
        ' empty search -> no-op (avoid infinite insertion)
        Dim searchStr As String: searchStr = CStr(rawSearch)
        If Len(searchStr) = 0 Then
            tmpResult = CStr(baseValLocal)
            GoTo exitfun
        End If
        Dim pos As Long: pos = 1
        Dim foundPos As Long
        Dim matchCount As Long: matchCount = 0

        ' For replace (first match) behave like JS: only first occurrence
        If Not isGlobal Then
            foundPos = InStr(pos, src, searchStr, vbBinaryCompare)
            If foundPos = 0 Then
                tmpResult = src: GoTo exitfun
            End If
            Dim beforePart As String: beforePart = left$(src, foundPos - 1)
            Dim afterPart As String: afterPart = mid$(src, foundPos + Len(searchStr))
            Dim replacementText As String
            replacementText = VM_StringReplacementText(replRaw, mid$(src, foundPos, Len(searchStr)), foundPos - 1, src)
            tmpResult = beforePart & replacementText & afterPart
            GoTo exitfun
        End If

        ' replaceAll loop
        Do
            foundPos = InStr(pos, src, searchStr, vbBinaryCompare)
            If foundPos = 0 Then
                ' append tail and exit
                outRepStr = outRepStr & mid$(src, pos)
                Exit Do
            End If
            ' append portion before match
            If foundPos > pos Then outRepStr = outRepStr & mid$(src, pos, foundPos - pos)
            ' compute replacement
            Dim repTxt As String
            repTxt = VM_StringReplacementText(replRaw, mid$(src, foundPos, Len(searchStr)), foundPos - 1, src)
            outRepStr = outRepStr & repTxt
            matchCount = matchCount + 1
            ' advance pos after this match (prevent overlapping infinite loops)
            pos = foundPos + Len(searchStr)
            ' continue searching
        Loop
        tmpResult = outRepStr
        GoTo exitfun
    End If
    Exit Function
exitfun:
    StrReplace = True
End Function

Private Function StrSplit(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim spSepStr As String
    If evaluated.Count = 2 Then
        Dim spLimit As Long
        spLimit = CLng(evaluated(2))
    Else
        spLimit = -1
    End If
    spSepStr = evaluated(1)
    Dim spIt As Long
    Dim tmpSParr() As String
    Dim tmpSplitted() As String
    If spSepStr = vbNullString Then
        If spLimit = -1 Then spLimit = Len(baseValLocal)
        If spLimit > 0 Then
            ReDim tmpSParr(1 To spLimit)
        End If
        For spIt = 1 To spLimit
            tmpSParr(spIt) = mid$(baseValLocal, spIt, 1)
        Next spIt
    Else
        Dim spItc As Long
        tmpSplitted = Split(baseValLocal, evaluated(1), -1, vbBinaryCompare)
        If spLimit = -1 Then spLimit = UBound(tmpSplitted) - LBound(tmpSplitted) + 1
        spItc = LBound(tmpSplitted)
        If spLimit > 0 Then
            ReDim tmpSParr(1 To spLimit)
        End If
        For spIt = 1 To spLimit
            tmpSParr(spIt) = tmpSplitted(spItc)
            spItc = spItc + 1
        Next spIt
    End If
    tmpResult = tmpSParr
    GoTo exitfun
    Exit Function
exitfun:
    StrSplit = True
End Function

Private Function StrStartswith(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim startStr As String
    Dim swComStr As String
    swComStr = CStr(evaluated(1))
    startStr = mid$(baseValLocal, 1, Len(swComStr))
    tmpResult = (startStr = swComStr)
    GoTo exitfun
    Exit Function
exitfun:
    StrStartswith = True
End Function

Private Function StrSubstring(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim nidx As Long
    Dim eidx As Long
    nidx = CLng(evaluated(1)) + 1
    If evaluated.Count = 2 Then
        eidx = CLng(evaluated(2))
        If eidx < nidx Then
            Dim tmpSv As Long
            tmpSv = eidx
            eidx = nidx - 1
            nidx = tmpSv + 1
            If nidx <= 0 Then nidx = 1
        End If
    Else
        eidx = Len(baseValLocal)
    End If
    If nidx < 0 Or eidx < 0 Then
        tmpResult = vbNullString
    Else
        tmpResult = mid(CStr(baseValLocal), nidx, eidx - nidx + 1)
    End If
    GoTo exitfun
    Exit Function
exitfun:
    StrSubstring = True
End Function


' Map methods
Private Function MapMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' an object's own function member wins over the builtin method of the same name, as in
    ' JavaScript: class methods such as set(), get() or has() are called, not shadowed
    If HasOwnFunction(baseValLocal, METHOD_NAME_) Then Exit Function
    ' ASF_Map (object) methods
    Select Case propN
        Case "keys"
            MapMethod = MapKeysMethod(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "values"
            MapMethod = MapValues(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "entries"
            MapMethod = MapEntries(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "size", "length"
            ' Return count of keys in object
            tmpResult = baseValLocal.Count
            GoTo exitfun

        Case "haskey", "has"
            ' Check if object has a specific key
            If evaluated.Count = 0 Then
                tmpResult = False
            Else
                tmpResult = baseValLocal.Exists(CStr(evaluated(1)))
            End If
            GoTo exitfun

        Case "get"
            MapMethod = MapGet(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "set"
            ' Set key-value pair and return the object (for chaining)
            If evaluated.Count >= 2 Then
                baseValLocal.SetValue CStr(evaluated(1)), evaluated(2)
            End If
            vAssignment tmpResult, baseValLocal
            GoTo exitfun

        Case "delete", "remove"
            MapMethod = MapDelete(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "clear"
            ' Remove all keys
            baseValLocal.Clear
            vAssignment tmpResult, baseValLocal
            GoTo exitfun

        Case "isempty"
            ' Check if object has no keys
            tmpResult = (baseValLocal.Count = 0)
            GoTo exitfun

        Case "clone"
            ' Create shallow copy of object
            vAssignment tmpResult, baseValLocal.Clone()
            GoTo exitfun

        Case "merge"
            MapMethod = MapMerge(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "foreach"
            MapMethod = MapForeach(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "map"
            MapMethod = MapMap(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "filter"
            MapMethod = MapFilter(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "some"
            MapMethod = MapSome(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case "every"
            MapMethod = MapEvery(evaluated, propN, baseValLocal, origBaseVal, tmpResult)
        Case Else
            Exit Function
    End Select
    Exit Function
exitfun:
    MapMethod = True
End Function

Private Function MapKeysMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Return array of object keys
    Dim keysCol As Collection: Set keysCol = baseValLocal.keys()
    Dim keysArr() As Variant
    Dim optBase_keys As Long: optBase_keys = 1
    If keysCol.Count = 0 Then
        ReDim keysArr(optBase_keys To optBase_keys)
    Else
        ReDim keysArr(optBase_keys To optBase_keys + keysCol.Count - 1)
        Dim ki As Long
        For ki = 1 To keysCol.Count
            keysArr(optBase_keys + ki - 1) = keysCol(ki)
        Next ki
    End If
    tmpResult = keysArr
    GoTo exitfun

    Exit Function
exitfun:
    MapKeysMethod = True
End Function

Private Function MapValues(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Return array of object values
    Dim itemsCol As Collection: Set itemsCol = baseValLocal.items()
    Dim valsArr() As Variant
    Dim optBase_vals As Long: optBase_vals = 1
    If itemsCol.Count = 0 Then
        ReDim valsArr(optBase_vals To optBase_vals)
    Else
        ReDim valsArr(optBase_vals To optBase_vals + itemsCol.Count - 1)
        Dim vi As Long
        For vi = 1 To itemsCol.Count
            vAssignment valsArr(optBase_vals + vi - 1), itemsCol(vi)
        Next vi
    End If
    tmpResult = valsArr
    GoTo exitfun

    Exit Function
exitfun:
    MapValues = True
End Function

Private Function MapEntries(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Return array of [key, value] pairs
    Dim entriesKeys As Collection: Set entriesKeys = baseValLocal.keys()
    Dim entriesItems As Collection: Set entriesItems = baseValLocal.items()
    Dim entriesArr() As Variant
    Dim optBase_ent As Long: optBase_ent = 1
    If entriesKeys.Count = 0 Then
        ReDim entriesArr(optBase_ent To optBase_ent)
    Else
        ReDim entriesArr(optBase_ent To optBase_ent + entriesKeys.Count - 1)
        Dim ei As Long
        For ei = 1 To entriesKeys.Count
            Dim pairArr() As Variant
            ReDim pairArr(0 To 1)
            pairArr(0) = entriesKeys(ei)
            vAssignment pairArr(1), entriesItems(ei)
            entriesArr(optBase_ent + ei - 1) = pairArr
        Next ei
    End If
    tmpResult = entriesArr
    GoTo exitfun

    Exit Function
exitfun:
    MapEntries = True
End Function

Private Function MapGet(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Get value for key with optional default
    If evaluated.Count = 0 Then
        tmpResult = Empty
    Else
        Dim getKey As String: getKey = CStr(evaluated(1))
        If baseValLocal.Exists(getKey) Then
            vAssignment tmpResult, baseValLocal.GetValue(getKey)
        Else
            ' Return default value if provided, otherwise Empty
            If evaluated.Count >= 2 Then
                vAssignment tmpResult, evaluated(2)
            Else
                tmpResult = Empty
            End If
        End If
    End If
    GoTo exitfun

    Exit Function
exitfun:
    MapGet = True
End Function

Private Function MapDelete(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Remove a key and return true/false
    If evaluated.Count = 0 Then
        tmpResult = False
    Else
        Dim delKey As String: delKey = CStr(evaluated(1))
        If baseValLocal.Exists(delKey) Then
            baseValLocal.Remove delKey
            tmpResult = True
        Else
            tmpResult = False
        End If
    End If
    GoTo exitfun

    Exit Function
exitfun:
    MapDelete = True
End Function

Private Function MapMerge(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    ' Merge another object into this one
    If evaluated.Count = 0 Then
        vAssignment tmpResult, baseValLocal
    Else
        Dim mergeObj As Variant: vAssignment mergeObj, evaluated(1)
        If typeName(mergeObj) = "ASF_Map" Then
            Dim mergeKeys As Collection: Set mergeKeys = mergeObj.keys()
            Dim mk As Variant
            For Each mk In mergeKeys
                baseValLocal.SetValue CStr(mk), mergeObj.GetValue(CStr(mk))
            Next mk
        End If
        vAssignment tmpResult, baseValLocal
    End If
    GoTo exitfun

    Exit Function
exitfun:
    MapMerge = True
End Function

Private Function MapForeach(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' Iterate over entries with callback(value, key, object)
    If evaluated.Count = 0 Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim cb_obj As Variant: vAssignment cb_obj, evaluated(1)
    If typeName(cb_obj) <> "ASF_Map" Or cb_obj.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim foreachKeys As Collection: Set foreachKeys = baseValLocal.keys()
    Dim fk As Variant
    Dim args_obj As Collection
    Dim tmpRes_obj As Variant
    For Each fk In foreachKeys
        Set args_obj = New Collection
        args_obj.Add baseValLocal.GetValue(CStr(fk))
        args_obj.Add CStr(fk)
        args_obj.Add baseValLocal
        vAssignment tmpMapObj, cb_obj
        vAssignment tmpRes_obj, CallClosure(tmpMapObj, args_obj, baseValLocal)
    Next fk

    vAssignment tmpResult, baseValLocal
    GoTo exitfun

    Exit Function
exitfun:
    MapForeach = True
End Function

Private Function MapMap(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' Map over entries, return new object with transformed values
    If evaluated.Count = 0 Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim cb_map As Variant: vAssignment cb_map, evaluated(1)
    If typeName(cb_map) <> "ASF_Map" Or cb_map.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim newMap As New ASF_Map
    Dim mapKeys As Collection: Set mapKeys = baseValLocal.keys()
    Dim mapk As Variant
    Dim args_map As Collection
    For Each mapk In mapKeys
        Set args_map = New Collection
        args_map.Add baseValLocal.GetValue(CStr(mapk))
        args_map.Add CStr(mapk)
        args_map.Add baseValLocal
        vAssignment tmpMapObj, cb_map
        Dim mapRes As Variant
        vAssignment mapRes, CallClosure(tmpMapObj, args_map, baseValLocal)
        newMap.SetValue CStr(mapk), mapRes
    Next mapk

    vAssignment tmpResult, newMap
    GoTo exitfun

    Exit Function
exitfun:
    MapMap = True
End Function

Private Function MapFilter(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' Filter entries, return new object with entries that pass test
    If evaluated.Count = 0 Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim cb_filter As Variant: vAssignment cb_filter, evaluated(1)
    If typeName(cb_filter) <> "ASF_Map" Or cb_filter.GetValue("type") <> "Closure" Then
        vAssignment tmpResult, Empty
        GoTo exitfun
    End If

    Dim filteredMap As New ASF_Map
    Dim filterKeys As Collection: Set filterKeys = baseValLocal.keys()
    Dim filterk As Variant
    Dim args_filter As Collection
    Dim filterRes As Variant
    For Each filterk In filterKeys
        Set args_filter = New Collection
        args_filter.Add baseValLocal.GetValue(CStr(filterk))
        args_filter.Add CStr(filterk)
        args_filter.Add baseValLocal
        vAssignment tmpMapObj, cb_filter
        vAssignment filterRes, CallClosure(tmpMapObj, args_filter, baseValLocal)
        If IsTruthy(filterRes) Then
            filteredMap.SetValue CStr(filterk), baseValLocal.GetValue(CStr(filterk))
        End If
    Next filterk

    vAssignment tmpResult, filteredMap
    GoTo exitfun

    Exit Function
exitfun:
    MapFilter = True
End Function

Private Function MapSome(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' Test if at least one entry passes test
    If evaluated.Count = 0 Then
        tmpResult = False
        GoTo exitfun
    End If

    Dim cb_some_obj As Variant: vAssignment cb_some_obj, evaluated(1)
    If typeName(cb_some_obj) <> "ASF_Map" Or cb_some_obj.GetValue("type") <> "Closure" Then
        tmpResult = False
        GoTo exitfun
    End If

    Dim someKeys As Collection: Set someKeys = baseValLocal.keys()
    Dim somek As Variant
    Dim args_some As Collection
    For Each somek In someKeys
        Set args_some = New Collection
        args_some.Add baseValLocal.GetValue(CStr(somek))
        args_some.Add CStr(somek)
        args_some.Add baseValLocal
        vAssignment tmpMapObj, cb_some_obj
        Dim someRes As Variant
        vAssignment someRes, CallClosure(tmpMapObj, args_some, baseValLocal)
        If IsTruthy(someRes) Then
            tmpResult = True
            GoTo exitfun
        End If
    Next somek

    tmpResult = False
    GoTo exitfun

    Exit Function
exitfun:
    MapSome = True
End Function

Private Function MapEvery(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant) As Boolean
    Dim tmpMapObj As ASF_Map
    ' Test if all entries pass test
    If evaluated.Count = 0 Then
        tmpResult = True
        GoTo exitfun
    End If

    Dim cb_every_obj As Variant: vAssignment cb_every_obj, evaluated(1)
    If typeName(cb_every_obj) <> "ASF_Map" Or cb_every_obj.GetValue("type") <> "Closure" Then
        tmpResult = False
        GoTo exitfun
    End If

    Dim everyKeys As Collection: Set everyKeys = baseValLocal.keys()
    Dim everyk As Variant
    Dim args_every As Collection
    For Each everyk In everyKeys
        Set args_every = New Collection
        args_every.Add baseValLocal.GetValue(CStr(everyk))
        args_every.Add CStr(everyk)
        args_every.Add baseValLocal
        vAssignment tmpMapObj, cb_every_obj
        Dim everyRes As Variant
        vAssignment everyRes, CallClosure(tmpMapObj, args_every, baseValLocal)
        If Not IsTruthy(everyRes) Then
            tmpResult = False
            GoTo exitfun
        End If
    Next everyk

    tmpResult = True
    GoTo exitfun

    Exit Function
exitfun:
    MapEvery = True
End Function


' Collection methods
Private Function CollectionMethod(evaluated As Collection, ByRef propN As String, ByRef baseValLocal As Variant, ByRef origBaseVal As Variant, ByRef tmpResult As Variant, ByVal isMath As Boolean) As Boolean
    If isMath Then
        If isMath Then
            If MathConstant(propN, tmpResult) Then GoTo exitfun
            Select Case propN
                Case "random"
                    ' a different sequence on every run, as in JavaScript
                    If Not RANDOMIZED_ Then
                        Randomize
                        RANDOMIZED_ = True
                    End If
                    tmpResult = Rnd()
                    GoTo exitfun
                Case "abs"
                    tmpResult = Abs(evaluated(1))
                    GoTo exitfun
                Case "acos"
                    If evaluated(1) = -1 Then
                        tmpResult = 4 * Atn(1)
                    ElseIf evaluated(1) = 1 Then
                        tmpResult = 0
                    ElseIf evaluated(1) > 1 Or evaluated(1) < -1 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Atn(-1 * evaluated(1) / _
                                        Sqr(-1 * evaluated(1) * evaluated(1) + 1)) + 2 * Atn(1)
                    End If
                    GoTo exitfun
                Case "acosh"
                    If evaluated(1) < 1 Then
                        tmpResult = "NaN"
                    ElseIf evaluated(1) = 1 Then
                        tmpResult = 0
                    Else
                        tmpResult = Log(evaluated(1) + Sqr(evaluated(1) * evaluated(1) - 1))
                    End If
                    GoTo exitfun
                Case "asin"
                    If evaluated(1) = -1 Then
                        tmpResult = -2 * Atn(1)
                    ElseIf evaluated(1) = 1 Then
                        tmpResult = 2 * Atn(1)
                    ElseIf evaluated(1) = 0 Then
                        tmpResult = 0
                    ElseIf evaluated(1) > 1 Or evaluated(1) < -1 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Atn(evaluated(1) / Sqr(-1 * evaluated(1) * evaluated(1) + 1))
                    End If
                    GoTo exitfun
                Case "asinh"
                    tmpResult = Log(evaluated(1) + Sqr(evaluated(1) * evaluated(1) + 1))
                    GoTo exitfun
                Case "atan"
                    tmpResult = Atn(evaluated(1))
                    GoTo exitfun
                Case "atan2" 'Arguments => (y, x)
                    If evaluated(1) = 0 And evaluated(2) = 0 Then
                        tmpResult = 0
                    ElseIf evaluated(1) = 0 And evaluated(2) < 0 Then
                        tmpResult = 4 * Atn(1)
                    Else
                        tmpResult = Atn(evaluated(1) / evaluated(2))
                    End If
                    GoTo exitfun
                Case "atanh"
                    If evaluated(1) >= 1 Or evaluated(1) <= -1 Then
                        tmpResult = "NaN"
                    ElseIf evaluated(1) = 0 Then
                        tmpResult = 0
                    Else
                        tmpResult = Log((1 + evaluated(1)) / (1 - evaluated(1))) / 2
                    End If
                    GoTo exitfun
                Case "cbrt"
                    tmpResult = Sgn(evaluated(1)) * Abs(evaluated(1)) ^ (1 / 3)
                    GoTo exitfun
                Case "ceil"
                    tmpResult = Fix(evaluated(1))
                    tmpResult = tmpResult - ((evaluated(1) <> tmpResult) And (evaluated(1) > 0))
                    GoTo exitfun
                Case "cos"
                    tmpResult = Cos(evaluated(1))
                    GoTo exitfun
                Case "cosh"
                    tmpResult = (Exp(evaluated(1)) + Exp(-1 * evaluated(1))) / 2
                    GoTo exitfun
                Case "exp"
                    tmpResult = Exp(evaluated(1))
                    GoTo exitfun
                Case "expm1"
                    tmpResult = Exp(evaluated(1)) - 1
                    GoTo exitfun
                Case "floor"
                    tmpResult = Fix(evaluated(1))
                    tmpResult = tmpResult + ((evaluated(1) <> tmpResult) And (evaluated(1) < 0))
                    GoTo exitfun
                Case "hypot"
                    Dim cat As Variant
                    tmpResult = 0
                    For Each cat In evaluated
                        tmpResult = tmpResult + cat * cat
                    Next cat
                    tmpResult = tmpResult ^ 0.5
                    GoTo exitfun
                Case "log"
                    If evaluated(1) <= 0 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Log(evaluated(1))
                    End If
                    GoTo exitfun
                Case "log10"
                    If evaluated(1) <= 0 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Log(evaluated(1)) / Log(10)
                    End If
                    GoTo exitfun
                Case "log1p"
                    If evaluated(1) <= 0 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Log(evaluated(1) + 1)
                    End If
                    GoTo exitfun
                Case "log2"
                    If evaluated(1) <= 0 Then
                        tmpResult = "NaN"
                    Else
                        tmpResult = Log(evaluated(1)) / Log(2)
                    End If
                    GoTo exitfun
                Case "max"
                    ' no arguments give Empty, as other built-ins called without their arguments do
                    Dim maxEl As Variant
                    tmpResult = Empty
                    If evaluated.Count > 0 Then
                        tmpResult = evaluated(1)
                        For Each maxEl In evaluated
                            If maxEl > tmpResult Then
                                tmpResult = maxEl
                            End If
                        Next maxEl
                    End If
                    GoTo exitfun
                Case "min"
                    Dim minEl As Variant
                    tmpResult = Empty
                    If evaluated.Count > 0 Then
                        tmpResult = evaluated(1)
                        For Each minEl In evaluated
                            If minEl < tmpResult Then
                                tmpResult = minEl
                            End If
                        Next minEl
                    End If
                    GoTo exitfun
                Case "pow"
                    tmpResult = evaluated(1) ^ evaluated(2)
                    GoTo exitfun
                Case "round"
                    tmpResult = Round(evaluated(1))
                    GoTo exitfun
                Case "sign"
                    tmpResult = Sgn(evaluated(1))
                    GoTo exitfun
                Case "sin"
                    tmpResult = Sin(evaluated(1))
                    GoTo exitfun
                Case "sinh"
                    tmpResult = (Exp(evaluated(1)) - Exp(-1 * evaluated(1))) / 2
                    GoTo exitfun
                Case "sqrt"
                    tmpResult = Sqr(evaluated(1))
                    GoTo exitfun
                Case "tan"
                    tmpResult = Tan(evaluated(1))
                    GoTo exitfun
                Case "tanh"
                    tmpResult = (Exp(evaluated(1)) - Exp(-1 * evaluated(1))) / _
                                (Exp(evaluated(1)) + Exp(-1 * evaluated(1)))
                    GoTo exitfun
            End Select
        Else
            ' Delegate the evaluation to the Aplication
            vAssignment tmpResult, CallObjectMethod(baseValLocal, propN, evaluated)
            GoTo exitfun
        End If
    Else
        vAssignment tmpResult, CallObjectMethod(baseValLocal, propN, evaluated)
        GoTo exitfun
    End If
    Exit Function
exitfun:
    CollectionMethod = True
End Function


Private Function makeStringForPadding(padStr As String, tLen As Long) As String
    Dim curLen As Long
    Dim padStrLen As Long: padStrLen = Len(padStr)
    Dim tmpResult As String
    Do While curLen < tLen
        curLen = curLen + padStrLen
        If curLen > tLen Then
            tmpResult = tmpResult & mid(padStr, 1, tLen - curLen + padStrLen)
        Else
            tmpResult = tmpResult & padStr
        End If
    Loop
    makeStringForPadding = tmpResult
End Function
' ---------------------
' Utilities
' ---------------------
Private Function IsTruthy(v As Variant) As Boolean
    If IsObject(v) Then IsTruthy = Not v Is Nothing: Exit Function
    If IsNull(v) Then IsTruthy = False: Exit Function
    If IsEmpty(v) Then IsTruthy = False: Exit Function
    If VarType(v) = vbBoolean Then IsTruthy = CBool(v): Exit Function
    If IsNumeric(v) Then IsTruthy = (CDbl(v) <> 0): Exit Function
    If VarType(v) = vbString Then
        If IsBoolean(CStr(v)) Then IsTruthy = CBool(v) Else IsTruthy = (CStr(v) <> vbNullString)
        Exit Function
    End If
    IsTruthy = True
End Function

Private Function IsBoolean(ByRef expression As String) As Boolean
    IsBoolean = (LCase$(expression) = "true")
    If Not IsBoolean Then IsBoolean = (LCase$(expression) = "false")
End Function

' True when the map has a member named prop holding a closure or a method
Private Function HasOwnFunction(ByRef mapVal As Variant, ByVal prop As String) As Boolean
    Dim m As ASF_Map: Set m = mapVal
    Dim member As Variant
    If Not m.TryGetValue(prop, member) Then Exit Function
    If Not IsMapValue(member) Then Exit Function
    Dim f As ASF_Map: Set f = member
    Select Case f.GetValue("type")
        Case "Closure", "BoundMethod", "Method"
            HasOwnFunction = True
    End Select
End Function

' True when v holds an ASF_Map; TypeOf builds no string, unlike TypeName, but needs an object
' The constants of Math, as in JavaScript; the name in any case, as Math's methods
Private Function MathConstant(ByVal name As String, ByRef value As Variant) As Boolean
    MathConstant = True
    Select Case UCase$(name)
        Case "PI": value = 4 * Atn(1)
        Case "E": value = Exp(1)
        Case "LN2": value = Log(2)
        Case "LN10": value = Log(10)
        Case "LOG2E": value = 1 / Log(2)
        Case "LOG10E": value = 1 / Log(10)
        Case "SQRT2": value = Sqr(2)
        Case "SQRT1_2": value = Sqr(0.5)
        Case Else: MathConstant = False
    End Select
End Function

Private Function IsMapValue(ByRef v As Variant) As Boolean
    If IsObject(v) Then IsMapValue = TypeOf v Is ASF_Map
End Function

Private Function ValueToStringForPrint(v As Variant) As String
    Dim visited As New Collection
    ValueToStringForPrint = ValueToStringWithCtx(v, visited, 0)
End Function
' core recursive printer with context
Private Function ValueToStringWithCtx(v As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler

    ' Depth guard
    If depth > VTP_MAX_DEPTH Then
        ValueToStringWithCtx = "..."
        Exit Function
    End If
    ' Null / Empty
    If IsNull(v) Then
        ValueToStringWithCtx = "NULL": Exit Function
    End If
    If IsEmpty(v) Then
        ValueToStringWithCtx = vbNullString: Exit Function
    End If

    ' Scalars
    If VarType(v) = vbString Then
        ValueToStringWithCtx = "'" & CStr(v) & "'": Exit Function
    End If
    If VarType(v) = vbBoolean Then
        If CBool(v) Then ValueToStringWithCtx = "True" Else ValueToStringWithCtx = "False"
        Exit Function
    End If
    If IsNumeric(v) Then
        ValueToStringWithCtx = CStr(v): Exit Function
    End If

    ' Arrays (native VBA arrays)
    If IsArray(v) Then
        If Not IsObject(v) Then
            ValueToStringWithCtx = ArrayToString(v, visited, depth)
        Else
            ValueToStringWithCtx = "<" & typeName(v) & ">"
        End If
        Exit Function
    End If

    ' Objects
    If IsObject(v) Then
        Dim tn As String: tn = typeName(v)
        Select Case tn
            Case "ASF_Map"
                ' cycle detection by object identity
                Dim i As Long
                For i = 1 To visited.Count
                    If visited(i) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next i
                visited.Add v
                ValueToStringWithCtx = MapToString(v, visited, depth + 1)
                visited.Remove visited.Count
                Exit Function

            Case "Collection"
                ' cycle detection
                Dim j As Long
                For j = 1 To visited.Count
                    If visited(j) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next j
                visited.Add v
                ValueToStringWithCtx = CollectionToString(v, visited, depth + 1)
                visited.Remove visited.Count
                Exit Function

            Case Else
                ' generic objects: try to call a ToString-like property if present, else fallback
                ValueToStringWithCtx = ObjectToString(v, visited, depth + 1)
                Exit Function
        End Select
    End If

    ' Fallback
    ValueToStringWithCtx = CStr(v)
    Exit Function

ErrHandler:
    ' On any unexpected error, return a safe placeholder and continue
    On Error Resume Next
    ValueToStringWithCtx = "[error: " & err.Number & "]"
    err.Clear
End Function

' Convert Map -> string
Private Function MapToString(m As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim keys As Collection: Set keys = m.keys
    Dim kcnt As Long: kcnt = keys.Count

    If kcnt = 0 Then
        MapToString = "{}": Exit Function
    End If

    ' For small maps and shallow depth prefer inline representation
    If kcnt <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim parts() As String
        ReDim parts(1 To kcnt)
        Dim i As Long
        For i = 1 To kcnt
            Dim key As String: key = CStr(keys(i))
            Dim val As Variant: vAssignment val, m.GetValue(key)
            parts(i) = CStr(key) & ": " & ValueToStringWithCtx(val, visited, depth)
        Next i
        MapToString = "{ " & Join(parts, ", ") & " }"
        Exit Function
    End If

    ' Multi-line pretty print
    Dim sb As String
    Dim indent As String: indent = String(depth * 2, " ")
    Dim innerIndent As String: innerIndent = String((depth + 1) * 2, " ")
    sb = "{"
    Dim first As Boolean: first = True
    Dim kk As Variant
    For Each kk In keys
        If Not first Then sb = sb & vbCrLf
        sb = sb & innerIndent & CStr(kk) & ": " & ValueToStringWithCtx(m.GetValue(CStr(kk)), visited, depth + 1)
        first = False
    Next kk
    sb = sb & vbCrLf & indent & "}"
    MapToString = sb
    Exit Function

ErrHandler:
    MapToString = "{<error>}"
    err.Clear
End Function

' Convert Collection -> string (treat as list)
Private Function CollectionToString(col As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim n As Long: n = col.Count
    If n = 0 Then CollectionToString = "[]": Exit Function
    If n <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim tmp() As String: ReDim tmp(1 To n)
        Dim ii As Long
        For ii = 1 To n
            tmp(ii) = ValueToStringWithCtx(col(ii), visited, depth)
        Next ii
        CollectionToString = "[ " & Join(tmp, ", ") & " ]"
        Exit Function
    End If

    Dim sb As String: sb = "["
    Dim i As Long
    Dim indent As String: indent = String((depth + 1) * 2, " ")
    For i = 1 To n
        If i > 1 Then sb = sb & vbCrLf
        sb = sb & indent & ValueToStringWithCtx(col(i), visited, depth + 1)
    Next i
    sb = sb & vbCrLf & String(depth * 2, " ") & "]"
    CollectionToString = sb
    Exit Function

ErrHandler:
    CollectionToString = "[<error>]"
    err.Clear
End Function

' Convert native VBA array -> string
Private Function ArrayToString(arr As Variant, visited As Collection, depth As Long) As String
    On Error GoTo ErrHandler
    Dim lb As Long, ub As Long
    lb = LBound(arr): ub = UBound(arr)
    Dim n As Long: n = ub - lb + 1
    If n <= 0 Then ArrayToString = "[]": Exit Function
    If n <= VTP_MAX_ITEMS_INLINE And depth <= 2 Then
        Dim tmp() As String: ReDim tmp(1 To n)
        Dim i As Long
        For i = lb To ub
            tmp(i - lb + 1) = ValueToStringWithCtx(arr(i), visited, depth)
        Next i
        ArrayToString = "[ " & Join(tmp, ", ") & " ]"
        Exit Function
    End If

    Dim sb As String: sb = "["
    Dim indent As String: indent = String((depth + 1) * 2, " ")
    Dim ii As Long
    For ii = lb To ub
        If ii > lb Then sb = sb & vbCrLf
        sb = sb & indent & ValueToStringWithCtx(arr(ii), visited, depth + 1)
    Next ii
    sb = sb & vbCrLf & String(depth * 2, " ") & "]"
    ArrayToString = sb
    Exit Function

ErrHandler:
    ArrayToString = "[<error>]"
    err.Clear
End Function

' Generic object to string fallback:
' - If the object is a Map-like (has Keys and GetValue), will attempt to treat it As ASF_Map.
' - Else TypeName + simple to-string
Private Function ObjectToString(obj As Variant, visited As Collection, depth As Long) As String
    On Error GoTo fallback
    Dim tn As String: tn = typeName(obj)

    ' Attempt Map-like duck typing: presence of Keys and GetValue
    ' (use On Error to bail out if methods missing)
    On Error GoTo Fallback2
    ' if successful, treat As ASF_Map
    Dim i As Long
    For i = 1 To visited.Count
        If visited(i) Is obj Then
            ObjectToString = "[Circular]": Exit Function
        End If
    Next i
    visited.Add obj
    ObjectToString = MapToString(obj, visited, depth)
    visited.Remove visited.Count
    If ObjectToString = "{<error>}" Then GoTo Fallback2
    Exit Function

Fallback2:
    ' Not a Map-like object: try default string
    On Error GoTo fallback
    ObjectToString = "<" & tn & ">"
    Exit Function

fallback:
    ObjectToString = "<object>"
    err.Clear
End Function

 ' VBAexpressions integration helpers
 ' ---------------------
 ' Evaluate a raw VBAexpressions expression string using a VBAexpressions instance,
 ' seeding it with the current ASF scope variables so VB expressions can reference ASF variables.
Private Function EvalVBExpressionWithScope(expr As String, vars As ASF_Map) As Variant
     On Error GoTo ErrHandler
     With GLOBALS_.gExprEvaluator
         ' Create expression in evaluator
         .Create expr
    
         ' the variables it can see
         Dim tmpVal As Variant
         Dim keyCol As Collection
         Dim key As Variant
         Set keyCol = vars.keys
         For Each key In keyCol
             vAssignment tmpVal, vars.GetValue(CStr(key))
             If IsArray(tmpVal) Then
                .VarValue(CStr(key)) = .ArrayToString(tmpVal)
             Else
                If Not IsObject(tmpVal) Then
                    .VarValue(CStr(key)) = CStr(tmpVal)
                End If
            End If
         Next key
    
         ' Evaluate
         .Eval
         If .ErrorType = 0 Then
             vAssignment EvalVBExpressionWithScope, cTypeExprRes(.result)
         Else
             ' On error, raise to caller; the try/catch at Exec layer can handle it
             err.Raise vbObjectError + 7001, "VM.EvalVBExpressionWithScope", "VBAexpressions eval error"
         End If
     End With
     Exit Function
ErrHandler:
     ' convert to runtime log and return Empty
     If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VBAexpr error: " & err.Description
     err.Clear
     EvalVBExpressionWithScope = Empty
End Function

Private Function cTypeExprRes(ByRef aResult As Variant) As Variant
    Dim tmpArr() As Variant
    On Error GoTo err_Handler
    If mid(aResult, 1, 2) = "{{" Then
        'VBAexpressions array was returned
        With GLOBALS_.gExprEvaluator
            vAssignment tmpArr, .ArrayFromString2(CStr(aResult), True)
            vAssignment cTypeExprRes, tmpArr
        End With
    Else
        vAssignment cTypeExprRes, aResult
    End If
    Exit Function
err_Handler:
    vAssignment cTypeExprRes, vbEmpty
End Function

' ---------------------
' Closures: a callback (map, sort...) calls a closure of ASF_VM, {type: 'Closure', fn, env}
' ---------------------
Private Function CallClosure(ByVal closureMap As ASF_Map, ByVal evaluatedArgs As Collection, Optional thisVal As Variant) As Variant
    ' a closure of ASF_VM: a function number and an environment
    If VM_ Is Nothing Then Exit Function
    If Not closureMap.Exists("fn") Then Exit Function
    If IsMissing(thisVal) Then
        vAssignment CallClosure, VM_.CallFromRuntime(closureMap, evaluatedArgs, False, Empty)
    Else
        vAssignment CallClosure, VM_.CallFromRuntime(closureMap, evaluatedArgs, True, thisVal)
    End If
End Function

' ---------------------
' Builtin helper functions (range/flatten/clone) and helpers for array methods
' ---------------------

' Flatten builtin with optional depth
' Usage:
'   builtin_flatten(value)         -> full flatten (old behavior)
'   builtin_flatten(value, depth) -> flatten up to 'depth' levels
Private Function builtin_flatten(v As Variant, Optional depthParam As Variant) As Variant
    Dim flattenAll As Boolean
    Dim maxDepth As Long

    ' Determine depth semantics:
    ' - If no depthParam or it's Empty/Non-numeric -> full flatten (use -1 sentinel)
    ' - If numeric and < 0 -> full flatten (-1)
    ' - Else use provided non-negative integer depth
    If IsMissing(depthParam) Then
        flattenAll = True
        maxDepth = -1
    ElseIf IsEmpty(depthParam) Then
        flattenAll = True
        maxDepth = -1
    Else
        On Error Resume Next
        Dim dVal As Double
        dVal = CDbl(depthParam)
        If err.Number <> 0 Then
            err.Clear
            flattenAll = True
            maxDepth = -1
        Else
            On Error GoTo 0
            If dVal < 0 Then
                flattenAll = True
                maxDepth = -1
            Else
                flattenAll = False
                maxDepth = CLng(dVal)
            End If
        End If
    End If

    ' Collect flattened elements into a Collection
    Dim col As New Collection
    If flattenAll Then
        Call flatten_recursive_full(v, col)
    Else
        Call flatten_recursive_depth(v, col, maxDepth)
    End If

    ' Build output array using __option_base
    Dim ob As Long: ob = 1
    Dim outArr() As Variant
    If col.Count = 0 Then
        ReDim outArr(ob To ob)
    Else
        ReDim outArr(ob To ob + col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            outArr(ob + i - 1) = col(i)
        Next i
    End If

    vAssignment builtin_flatten, outArr
End Function

' Full flatten
Private Sub flatten_recursive_full(v As Variant, ByRef col As Collection)
    If IsArray(v) Then
        Dim i As Long
        For i = LBound(v) To UBound(v)
            Call flatten_recursive_full(v(i), col)
        Next i
    Else
        col.Add v
    End If
End Sub

' Depth-limited flatten: maxDepth indicates how many levels to flatten.
' - If maxDepth = 0 and v is an array -> add the array itself as a single element (no flatten)
' - If maxDepth > 0 and v is an array -> recurse with maxDepth-1
Private Sub flatten_recursive_depth(v As Variant, ByRef col As Collection, ByVal maxDepth As Long)
    If IsArray(v) Then
        If maxDepth < 0 Then
            ' Do not flatten this array level  push the array object as an element.
            col.Add v
        Else
            Dim i As Long
            For i = LBound(v) To UBound(v)
                Call flatten_recursive_depth(v(i), col, maxDepth - 1)
            Next i
        End If
    Else
        col.Add v
    End If
End Sub
Private Function builtin_clone(v As Variant) As Variant
    If IsArray(v) Then
        Dim lb As Long: lb = LBound(v)
        Dim ub As Long: ub = UBound(v)
        Dim outArr() As Variant
        ReDim outArr(lb To ub)
        Dim i As Long
        For i = lb To ub
            outArr(i) = builtin_clone(v(i))
        Next i
        vAssignment builtin_clone, outArr
        Exit Function
    End If
    If IsObject(v) Then
        If typeName(v) = "ASF_Map" Then
            Dim m As ASF_Map: Set m = v
            Dim mc As ASF_Map: Set mc = m.Clone()
            vAssignment builtin_clone, mc
            Exit Function
        End If
    End If
    vAssignment builtin_clone, v
End Function

 ' Recursively applies mapper closure to value. If value is an array, map each element recursively.
' mapper : closure map (type="Closure")
' progScope : current scope stack (used for __option_base and to pass as runtime context)
' parentThis : optional 'this' value passed to CallClosure (typically the containing array)
' optBase: the indexing base
Private Function map_apply_recursive(value As Variant, mapper As ASF_Map, _
                                    Optional parentThis As Variant, _
                                    Optional optBase As Long = 1) As Variant
    On Error GoTo ErrHandler

    ' If the value is an array, create an output array with same length and map each element recursively
    If IsArray(value) Then
        Dim lb As Long, ub As Long
        On Error Resume Next
        lb = LBound(value)
        ub = UBound(value)
        If err.Number <> 0 Then
            err.Clear
            vAssignment map_apply_recursive, Empty
            Exit Function
        End If
        On Error GoTo ErrHandler

        Dim cnt As Long: cnt = ub - lb + 1
        If cnt < 0 Then cnt = 0
        Dim outArr() As Variant
        If cnt = 0 Then
            ReDim outArr(optBase To optBase)
        Else
            ReDim outArr(optBase To optBase + cnt - 1)
        End If

        Dim i As Long
        For i = 0 To cnt - 1
            Dim srcIdx As Long: srcIdx = lb + i
            Dim childRes As Variant
            ' Recurse: pass the current array as parentThis so mapper can use it if needed
            vAssignment childRes, map_apply_recursive(value(srcIdx), mapper, value, optBase)
            vAssignment outArr(optBase + i), childRes
        Next i

        vAssignment map_apply_recursive, outArr
        Exit Function
    End If

    ' If not an array, call the mapper closure with the value as single argument.
    Dim args As New Collection
    args.Add value
    Dim result As Variant
    vAssignment result, CallClosure(mapper, args, parentThis)
    args.Remove args.Count
    vAssignment map_apply_recursive, result
    Exit Function

ErrHandler:
    ' On mapper errors propagate as runtime error so calling code can handle (consistent with earlier semantics)
    err.Raise err.Number, err.Source, err.Description
End Function

Private Function builtin_typeof(v As Variant) As String
    On Error Resume Next
    ' Null / VBA Null
    If IsNull(v) Then
        builtin_typeof = "null"
        Exit Function
    End If

    ' Empty / uninitialized
    If IsEmpty(v) Then
        builtin_typeof = "undefined"
        Exit Function
    End If

    ' Array
    Dim isArr As Boolean
    isArr = False
    If Not IsError(v) Then
        On Error Resume Next
        If err.Number = 0 Then isArr = IsArray(v) Else err.Clear
        On Error GoTo 0
    End If
    If isArr Then
        If Not IsObject(v) Then
            builtin_typeof = "array"
        Else
            builtin_typeof = "object" & ": <" & typeName(v) & ">"
        End If
        Exit Function
    End If

    ' ASF_Map (objects and closures)  detect function closures specially
    If IsObject(v) Then
        ' if it's an ASF_Map, look inside for "type" = "Closure"
        Dim typeN As String: typeN = typeName(v)
        If typeN = "ASF_Map" Then
            If v.Exists("type") Then
                Dim tP As Variant: tP = v.GetValue("type")
                If CStr(tP) = "Closure" Then
                    builtin_typeof = "function"
                    Exit Function
                End If
            End If
        End If
        ' generic object or COM/host objects
        builtin_typeof = "object" & ": <" & typeN & ">"
        Exit Function
    End If

    ' Boolean (True/False)
    If VarType(v) = vbBoolean Then
        builtin_typeof = "boolean"
        Exit Function
    End If

    ' String
    If VarType(v) = vbString Then
        builtin_typeof = "string"
        Exit Function
    End If
    
    ' Numeric
    If IsNumeric(v) Then
        builtin_typeof = "number"
        Exit Function
    End If

    ' Fallback: unknown
    builtin_typeof = "unknown"
End Function

' ---------------------------
' Array method helpers
' ---------------------------

' Deep equality for nested arrays and ASF_Map (objects)
Private Function DeepEqual(a As Variant, b As Variant) As Boolean
    On Error GoTo notEqual
    ' same reference?
    If IsObject(a) And IsObject(b) Then
        If typeName(a) = "ASF_Map" And typeName(b) = "ASF_Map" Then
            ' compare keys count and values
            Dim keysA As Collection: Set keysA = a.keys()
            Dim keysB As Collection: Set keysB = b.keys()
            If keysA.Count <> keysB.Count Then GoTo notEqual
            Dim k As Variant
            For Each k In keysA
                If Not b.Exists(CStr(k)) Then GoTo notEqual
                If Not DeepEqual(a.GetValue(CStr(k)), b.GetValue(CStr(k))) Then GoTo notEqual
            Next k
            DeepEqual = True
            Exit Function
        Else
            ' fallback to object reference equality
            DeepEqual = (a Is b)
            Exit Function
        End If
    End If

    If IsArray(a) And IsArray(b) Then
        Dim la As Long: la = UBound(a) - LBound(a) + 1
        Dim lb As Long: lb = UBound(b) - LBound(b) + 1
        If la <> lb Then GoTo notEqual
        Dim i As Long
        For i = LBound(a) To UBound(a)
            If Not DeepEqual(a(i), b(i)) Then GoTo notEqual
        Next i
        DeepEqual = True
        Exit Function
    End If

    ' primitive comparison (numbers, strings, booleans)
    If IsNull(a) And IsNull(b) Then
        DeepEqual = True
        Exit Function
    End If
    If IsEmpty(a) And IsEmpty(b) Then
        DeepEqual = True
        Exit Function
    End If
    If VarType(a) = VarType(b) Then
        If VarType(a) = vbString Then
            DeepEqual = (CStr(a) = CStr(b))
        ElseIf (VarType(a) = vbDouble Or VarType(a) = vbSingle Or VarType(a) = vbInteger Or VarType(a) = vbLong Or VarType(a) = vbCurrency) Then
            DeepEqual = (CDbl(a) = CDbl(b))
        Else
            DeepEqual = (a = b)
        End If
        Exit Function
    Else
        ' try string compare fallback
        DeepEqual = (CStr(a) = CStr(b))
        Exit Function
    End If

notEqual:
    DeepEqual = False
    Exit Function
End Function

' Helper: normalize an index (user-space index to zero-based offset). Accept negative indexing.
Private Function NormalizeIndex(idxRaw As Long, lenA As Long, optBase As Long) As Long
    Dim idx As Long
    If idxRaw < 0 Then
        idx = lenA + idxRaw  ' negative => from end (zero-based)
    Else
        idx = idxRaw - optBase ' convert to zero-based offset
    End If
    If idx < 0 Then idx = 0
    If idx > lenA Then idx = lenA
    NormalizeIndex = idx
End Function

' Helper: create shallow copy of array preserving option base
Private Function ShallowCopyArray(arr As Variant, optBase As Long) As Variant
    If Not IsArray(arr) Then ShallowCopyArray = Empty: Exit Function
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim lenA As Long: lenA = ub - lb + 1
    Dim out() As Variant
    ReDim out(optBase To optBase + lenA - 1)
    Dim i As Long
    For i = 0 To lenA - 1
        vAssignment out(optBase + i), arr(lb + i)
    Next i
    ShallowCopyArray = out
End Function

' QS_Compare: returns -1/0/1 using optional closure comparator; fallback to numeric/string compare
Private Function QS_Compare(ByVal a As Variant, ByVal b As Variant, ByVal comp As Variant, ByVal hasComp As Boolean, ByVal thisVal As Variant) As Long
    On Error GoTo fallback_cmp
    Dim cres As Variant
    Dim tmpMapObj As ASF_Map
    
    If hasComp Then
        Dim args_cmp As New Collection
        args_cmp.Add a
        args_cmp.Add b
        vAssignment tmpMapObj, comp
        vAssignment cres, CallClosure(tmpMapObj, args_cmp, thisVal)
        If IsNumeric(cres) Then
            Dim d As Double: d = CDbl(cres)
            If d < 0 Then
                QS_Compare = -1
            ElseIf d > 0 Then
                QS_Compare = 1
            Else
                QS_Compare = 0
            End If
            Exit Function
        Else
            If IsTruthy(cres) Then
                QS_Compare = 1
                Exit Function
            Else
                QS_Compare = 0
                Exit Function
            End If
        End If
    End If

fallback_cmp:
    err.Clear
    ' default comparator: numeric when possible, otherwise string compare
    If IsNumeric(a) And IsNumeric(b) Then
        Dim da As Double: da = CDbl(a)
        Dim db As Double: db = CDbl(b)
        If da < db Then
            QS_Compare = -1
        ElseIf da > db Then
            QS_Compare = 1
        Else
            QS_Compare = 0
        End If
    Else
        Dim sa As String: sa = CStr(a)
        Dim sb As String: sb = CStr(b)
        If sa < sb Then
            QS_Compare = -1
        ElseIf sa > sb Then
            QS_Compare = 1
        Else
            QS_Compare = 0
        End If
    End If
End Function

' QS_Swap: swap two elements in a Variant array
Private Sub QS_Swap(ByRef arr() As Variant, ByVal i As Long, ByVal j As Long)
    Dim tmp As Variant
    vAssignment tmp, arr(i)
    vAssignment arr(i), arr(j)
    vAssignment arr(j), tmp
End Sub

' QS_QuickSort: in-place quicksort (Hoare-style partition)
Private Sub QS_QuickSort(ByRef arr() As Variant, ByVal low As Long, ByVal high As Long, ByVal comp As Variant, ByVal hasComp As Boolean, ByVal thisVal As Variant)
    Dim i As Long, j As Long
    If low >= high Then Exit Sub

    i = low
    j = high
    Dim mid As Long: vAssignment mid, (low + high) \ 2
    Dim pivot As Variant: vAssignment pivot, arr(mid)

    Do While i <= j
        Do While QS_Compare(arr(i), pivot, comp, hasComp, thisVal) < 0
            i = i + 1
        Loop
        Do While QS_Compare(arr(j), pivot, comp, hasComp, thisVal) > 0
            j = j - 1
        Loop
        If i <= j Then
            QS_Swap arr, i, j
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QS_QuickSort arr, low, j, comp, hasComp, thisVal
    If i < high Then QS_QuickSort arr, i, high, comp, hasComp, thisVal
End Sub

' Helper: compute replacement text (supports string or Closure)
Private Function VM_StringReplacementText(ByVal replRaw As Variant, ByVal matchStr As String, ByVal matchIndexZeroBased As Long, ByVal originalString As String) As String
    On Error GoTo fallback_str

    If Not IsEmpty(replRaw) And typeName(replRaw) = "ASF_Map" Then
        If replRaw.Exists("type") And CStr(replRaw.GetValue("type")) = "Closure" Then
            Dim callArgsR As New Collection
            callArgsR.Add matchStr
            callArgsR.Add matchIndexZeroBased
            callArgsR.Add originalString
            Dim callRes As Variant
            Dim tmpMapObj As ASF_Map
            vAssignment tmpMapObj, replRaw
            vAssignment callRes, CallClosure(tmpMapObj, callArgsR, originalString)
            If IsArray(callRes) Or (IsObject(callRes) And typeName(callRes) = "ASF_Map") Then
                VM_StringReplacementText = ValueToStringForPrint(callRes)
                Exit Function
            ElseIf IsNull(callRes) Then
                VM_StringReplacementText = "null": Exit Function
            ElseIf IsEmpty(callRes) Then
                VM_StringReplacementText = vbNullString: Exit Function
            Else
                VM_StringReplacementText = CStr(callRes): Exit Function
            End If
        Else
            VM_StringReplacementText = CStr(replRaw)
            Exit Function
        End If
    End If

fallback_str:
    err.Clear
    If IsNull(replRaw) Then
        VM_StringReplacementText = "null"
    ElseIf IsEmpty(replRaw) Then
        VM_StringReplacementText = vbNullString
    Else
        VM_StringReplacementText = CStr(replRaw)
    End If
End Function

' Expand replacement string using JS-like $ sequences:
' $$ -> $
' $& -> full match
' $` -> prefix (text before match)
' $' -> suffix (text after match)
' $n or $nn -> nth capture (1-based). If not present -> empty string.
Private Function VM_ExpandReplacementFromCapturesJS( _
        ByVal caps As Collection, _
        ByVal repl As String, _
        ByVal original As String, _
        ByVal matchStart As Long, _
        ByVal matchEnd As Long) As String

    Dim out As String: out = vbNullString
    Dim i As Long: i = 1
    Dim Lr As Long: Lr = Len(repl)
    Dim nCaps As Long: nCaps = 0
    If Not (caps Is Nothing) Then nCaps = caps.Count

    Dim prefix As String: prefix = vbNullString
    Dim suffix As String: suffix = vbNullString
    If matchStart > 1 Then prefix = left$(original, matchStart - 1)
    If matchEnd < Len(original) Then suffix = mid$(original, matchEnd + 1)

    Do While i <= Lr
        Dim ch As String: ch = mid$(repl, i, 1)
        If ch = "$" And i < Lr Then
            Dim nextCh As String: nextCh = mid$(repl, i + 1, 1)
            Select Case nextCh
                Case "$"
                    out = out & "$"
                    i = i + 2
                Case "&"
                    If nCaps >= 1 Then
                        out = out & CStr(caps.Item(1))
                    End If
                    i = i + 2
                Case "`"
                    out = out & prefix
                    i = i + 2
                Case "'"
                    out = out & suffix
                    i = i + 2
                Case "0" To "9"
                    ' parse up to two digits (JS supports $1..$99)
                    Dim j As Long: j = i + 1
                    Dim numStr As String: numStr = vbNullString
                    Do While j <= Lr And Len(numStr) < 2 And mid$(repl, j, 1) >= "0" And mid$(repl, j, 1) <= "9"
                        numStr = numStr & mid$(repl, j, 1)
                        j = j + 1
                    Loop
                    If numStr <> vbNullString Then
                        Dim idx As Long: idx = CLng(numStr)
                        ' JS: $1 maps to capture group 1 (caps.Item(2) in our caps layout)
                        If idx >= 1 And idx <= (nCaps - 1) Then
                            out = out & CStr(caps.Item(idx + 1))
                        Else
                            ' unknown group -> empty string
                            out = out & vbNullString
                        End If
                        i = i + 1 + Len(numStr)
                    Else
                        ' single $ followed by non-digit (shouldn't reach here)
                        out = out & "$": i = i + 1
                    End If
                Case Else
                    ' Not a recognized sequence -> keep literal $
                    out = out & "$"
                    i = i + 1
            End Select
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    VM_ExpandReplacementFromCapturesJS = out
End Function


' Convert a Collection of captures -> 1-based VBA array (full match then groups)
Private Function VM_CapturesToArray(ByVal caps As Collection) As Variant
    If caps Is Nothing Then
        VM_CapturesToArray = Array()
        Exit Function
    End If
    Dim n As Long: n = caps.Count
    Dim arr() As Variant: ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = caps.Item(i)
    Next i
    VM_CapturesToArray = arr
End Function

' Parse a JS-like /pattern/flags string into pattern and flags
Private Sub VM_ParseSlashRegex(ByVal reSrc As String, ByRef outPattern As String, ByRef outFlags As String)
    outPattern = vbNullString
    outFlags = vbNullString
    If Len(reSrc) < 2 Then Exit Sub
    If left$(reSrc, 1) <> "/" Then Exit Sub
    Dim lastSlash As Long: lastSlash = InStrRev(reSrc, "/")
    If lastSlash <= 1 Then Exit Sub
    outPattern = mid$(reSrc, 2, lastSlash - 2)
    If lastSlash < Len(reSrc) Then outFlags = mid$(reSrc, lastSlash + 1)
End Sub

' Helper: test whether an evaluated search value is a slash-regex string
Private Function VM_IsSlashRegexString(ByVal v As Variant) As Boolean
    If Not (VarType(v) = vbString) Then VM_IsSlashRegexString = False: Exit Function
    If Len(CStr(v)) >= 2 And left$(CStr(v), 1) = "/" And InStrRev(CStr(v), "/") > 1 Then
        VM_IsSlashRegexString = True
    Else
        VM_IsSlashRegexString = False
    End If
End Function
' Convert a Collection of captures (caps) to a Variant containing a 1-D VBA array.
' - If caps Is Nothing -> returns Empty
' - If caps.Count = 0 -> returns a Variant containing an empty array (Array())
' - If any element is a Collection, conversion recurses producing nested arrays.
Private Function RegexCapsToVariantArray(ByVal caps As Collection) As Variant
    If caps Is Nothing Then
        RegexCapsToVariantArray = Empty
        Exit Function
    End If

    Dim n As Long: n = caps.Count
    If n = 0 Then
        RegexCapsToVariantArray = Array() ' empty array inside Variant
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        Dim v As Variant: v = caps.Item(i)
        If IsObject(v) Then
            On Error Resume Next
            Dim tn As String: tn = typeName(v)
            On Error GoTo 0
            If tn = "Collection" Then
                out(i - 1) = RegexCapsToVariantArray(v) ' recursive
            Else
                out(i - 1) = v
            End If
        Else
            out(i - 1) = v
        End If
    Next i

    RegexCapsToVariantArray = out
End Function

' Convert a collection-of-captures (e.g. matchAll result: Collection of Collections)
' into a Variant containing a 1-D array of 1-D arrays (or nested arrays).
' Returns Empty when matches Is Nothing.
Private Function RegexMatchesToVariantArray(ByVal matches As Collection) As Variant
    If matches Is Nothing Then
        RegexMatchesToVariantArray = Empty
        Exit Function
    End If

    Dim m As Long: m = matches.Count
    If m = 0 Then
        RegexMatchesToVariantArray = Array()
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(0 To m - 1)

    Dim i As Long
    For i = 1 To m
        Dim caps As Collection
        Set caps = matches.Item(i)
        out(i - 1) = RegexCapsToVariantArray(caps)
    Next i

    RegexMatchesToVariantArray = out
End Function

' Push a call onto the call stack trace
Public Sub PushCallTrace(funcName As String, args As Collection)
    If Not ENABLE_CALL_TRACE_ Then Exit Sub
    
    Dim trace As New ASF_Map
    trace.Add "type", "CALL"
    trace.Add "name", funcName
    trace.Add "timestamp", Now
    
    ' Store arguments
    Dim argsList As New Collection
    If Not args Is Nothing Then
        Dim i As Long
        For i = 1 To args.Count
            argsList.Add ValueToStringForPrint(args(i))
        Next i
    End If
    trace.SetValue "args", argsList
    
    GLOBALS_.gCallStack.Add trace
End Sub

' Pop a call from the call stack trace and record return value
Public Sub PopCallTrace(funcName As String, returnValue As Variant)
    If Not ENABLE_CALL_TRACE_ Then Exit Sub
    
    ' Find the most recent matching call
    Dim i As Long
    For i = GLOBALS_.gCallStack.Count To 1 Step -1
        Dim trace As ASF_Map
        Set trace = GLOBALS_.gCallStack(i)
        If trace.GetValue("type") = "CALL" And trace.GetValue("name") = funcName Then
            ' Update with return info
            trace.SetValue "returned", Now
            trace.SetValue "returnValue", ValueToStringForPrint(returnValue)
            Exit Sub
        End If
    Next i
End Sub

' Record an error in the call trace
Public Sub TraceError(funcName As String, errDesc As String)
    If Not ENABLE_CALL_TRACE_ Then Exit Sub
    
    Dim trace As New ASF_Map
    trace.Add "type", "ERROR"
    trace.Add "name", funcName
    trace.Add "timestamp", Now
    trace.Add "error", errDesc
    
    GLOBALS_.gCallStack.Add trace
End Sub
' ---------------------
' Module System: Resolve Module Path
' ---------------------
Private Function ResolveModulePath(modulePath As String) As String
    ' Simple resolution: if path starts with ./ or ../ treat as relative
    ' Otherwise treat as absolute or package name

    Dim cleanPath As String: cleanPath = modulePath
    If left$(cleanPath, 1) = "'" And right$(cleanPath, 1) = "'" Then
        cleanPath = mid$(cleanPath, 2, Len(cleanPath) - 2)
    End If
    
    ' For relative paths, resolve against current module directory
    Dim relPath As String: relPath = Replace(cleanPath, "\", "/")
    If left$(relPath, 2) = "./" Or left$(relPath, 3) = "../" Then
        Dim currentDir As String
        If GLOBALS_.CURRENT_MODULE_PATH <> vbNullString Then
            Dim basePath As String: basePath = Replace(GLOBALS_.CURRENT_MODULE_PATH, "\", "/")
            Dim lastSlash As Long
            lastSlash = InStrRev(basePath, "/")
            If lastSlash > 0 Then
                If GLOBALS_.gModuleExports.Exists(GLOBALS_.CURRENT_MODULE_PATH) Then
                    ' Inside a running module the path is the module file: use its directory
                    currentDir = left$(basePath, lastSlash)
                ElseIf lastSlash < Len(basePath) Then
                    currentDir = basePath & "/"
                Else
                    currentDir = basePath
                End If
            Else
                currentDir = vbNullString
            End If
        Else
            currentDir = vbNullString
        End If
        ' Consume the leading ./ and ../ segments: each ../ goes up one directory
        Dim parentSlash As Long
        Do
            If left$(relPath, 2) = "./" Then
                relPath = mid$(relPath, 3)
            ElseIf left$(relPath, 3) = "../" Then
                relPath = mid$(relPath, 4)
                If currentDir = vbNullString Or right$(currentDir, 3) = "../" Then
                    ' No directory left to leave: keep the step
                    currentDir = currentDir & "../"
                ElseIf Len(currentDir) > 1 Then
                    parentSlash = InStrRev(currentDir, "/", Len(currentDir) - 1)
                    ' At a root ("/", "C:/") there is no parent: stay there
                    If parentSlash > 0 Then currentDir = left$(currentDir, parentSlash)
                End If
            Else
                Exit Do
            End If
        Loop
        ResolveModulePath = currentDir & relPath
    Else
        ' Absolute path or module name
        ResolveModulePath = cleanPath
    End If
End Function

Public Function ReadModuleSource(modulePath As String) As String
    Dim filePath As String: filePath = ResolveModulePath(modulePath)
    
    ' Add .vas extension if not present
    If Not FileExists(filePath) Then
        If right$(filePath, 4) <> ".vas" Then
            filePath = filePath & ".vas"
        End If
    End If
    
    If Not FileExists(filePath) Then
        err.Raise vbObjectError + 9012, "VM.ReadModuleSource", "Module file not found: " & modulePath
    End If
    ReadModuleSource = ReadTextFile(filePath)
End Function
Public Function ReadTextFile(filePath As String) As String
    If FileExists(filePath) Then
        Dim EOFile As Long
        Dim FileHandled As Integer
        Dim OutputStr As String
        FileHandled = FreeFile
        Open filePath For Binary As #FileHandled
        EOFile = LOF(FileHandled)
        OutputStr = Space$(EOFile)
        Get #FileHandled, , OutputStr
        Close #FileHandled
        ReadTextFile = OutputStr
        OutputStr = vbNullString
    End If
End Function
Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = CBool(LenB(Dir(filePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function
Private Sub SetObjectProperty(obj As Variant, propName As Variant, propValue As Variant)
    Dim vbaParam As Variant: vAssignment vbaParam, CastVBAparam(propValue)
    If Not IsObject(propValue) Then
        CallByName obj, propName, VbLet, vbaParam
    Else
        CallByName obj, propName, VbSet, vbaParam
    End If
    ' Trace: Pop call with return value
    PopCallTrace CStr(propName), vbaParam
End Sub
Private Function CallObjectMethod(obj As Variant, methodName As Variant, Optional args As Collection) As Variant
    Dim tmpResult As Variant
    If HasCOMPrototype(obj, CStr(methodName)) Then
        Dim prototypeInfo As ASF_Map: Set prototypeInfo = GetCOMPrototype(obj, CStr(methodName))
        If Not prototypeInfo Is Nothing Then
            Dim funcIdx As Long: funcIdx = prototypeInfo.GetValue("funcIndex")
            Dim prototypeArgs As Collection
            If Not IsMissing(args) And Not args Is Nothing Then
                Set prototypeArgs = args
            Else
                Set prototypeArgs = New Collection
            End If
            ' Trace: Push call
            PushCallTrace CStr(methodName), prototypeArgs
            Dim protoClosure As Variant
            If prototypeInfo.TryGetValue("closure", protoClosure) Then vAssignment tmpResult, CallClosure(protoClosure, prototypeArgs, obj)
        End If
    Else
        PushCallTrace CStr(methodName), args
        If args Is Nothing Then
            vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod)
        Else
            Dim nArgs As Long: nArgs = args.Count
            If nArgs = 1 Then
                If IsEmpty(args(1)) Then nArgs = nArgs - 1
            End If
            With args
                'A lengthy and visually overwhelming implementation.
                Select Case nArgs
                    Case 0
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod)
                    Case 1
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)))
                    Case 2
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)))
                    Case 3
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)))
                    Case 4
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)))
                    Case 5
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)))
                    Case 6
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)))
                    Case 7
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)))
                    Case 8
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)))
                    Case 9
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)))
                    Case 10
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)))
                    Case 11
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)))
                    Case 12
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)))
                    Case 13
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)))
                    Case 14
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)))
                    Case 15
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)))
                    Case 16
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)))
                    Case 17
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)))
                    Case 18
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)))
                    Case 19
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)))
                    Case 20
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)))
                    Case 21
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)))
                    Case 22
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)))
                    Case 23
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)))
                    Case 24
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)))
                    Case 25
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)))
                    Case 26
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)), CastVBAparam(.Item(26)))
                    Case 27
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)), CastVBAparam(.Item(26)), CastVBAparam(.Item(27)))
                    Case 28
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)), CastVBAparam(.Item(26)), CastVBAparam(.Item(27)), CastVBAparam(.Item(28)))
                    Case 29
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)), CastVBAparam(.Item(26)), CastVBAparam(.Item(27)), CastVBAparam(.Item(28)), CastVBAparam(.Item(29)))
                    Case 30
                        vAssignment tmpResult, CallByName(obj, methodName, VbGet Or VbMethod, CastVBAparam(.Item(1)), CastVBAparam(.Item(2)), CastVBAparam(.Item(3)), CastVBAparam(.Item(4)), CastVBAparam(.Item(5)), CastVBAparam(.Item(6)), CastVBAparam(.Item(7)), CastVBAparam(.Item(8)), CastVBAparam(.Item(9)), CastVBAparam(.Item(10)), CastVBAparam(.Item(11)), CastVBAparam(.Item(12)), CastVBAparam(.Item(13)), CastVBAparam(.Item(14)), CastVBAparam(.Item(15)), CastVBAparam(.Item(16)), CastVBAparam(.Item(17)), CastVBAparam(.Item(18)), CastVBAparam(.Item(19)), CastVBAparam(.Item(20)), CastVBAparam(.Item(21)), CastVBAparam(.Item(22)), CastVBAparam(.Item(23)), CastVBAparam(.Item(24)), CastVBAparam(.Item(25)), CastVBAparam(.Item(26)), CastVBAparam(.Item(27)), CastVBAparam(.Item(28)), CastVBAparam(.Item(29)), CastVBAparam(.Item(30)))
                    Case Else
                        err.Raise vbObjectError + 513, "CallObjectMethod", "Too many arguments: maximum supported is 30."
                End Select
            End With
        End If
    End If
    If IsArray(tmpResult) And Not IsObject(tmpResult) Then
        CallObjectMethod = ArrayToASF2(tmpResult)
    Else
        vAssignment CallObjectMethod, tmpResult
    End If
    ' Trace: Pop call with return value
    PopCallTrace CStr(methodName), CallObjectMethod
End Function

Private Function CastVBAparam(aValue As Variant) As Variant
    If IsArray(aValue) And Not IsObject(aValue) Then
        CastVBAparam = ASFto2Darray(aValue)
    Else
        vAssignment CastVBAparam, aValue
    End If
End Function
Private Function IsVBAcollection(obj As Variant) As Boolean
    Dim typeN As String: typeN = typeName(obj)
    IsVBAcollection = (BinarySearch(GLOBALS_.knownMSOcollections, typeN) <> -1)
End Function

' Searches for an item in an 1D array. Returns
' the index of the item when found and -1 when missing values.
' The data must be supplied sorted
Private Function BinarySearch(ByRef TargetArr As Variant, ByRef Target As Variant) As Long

    Dim tmpR As Boolean
    Dim BottomIdx As Double
    Dim TopIdx As Double
    Dim MiddleIdx As Double
    
    BottomIdx = LBound(TargetArr)
    TopIdx = UBound(TargetArr)
    Do
        MiddleIdx = BottomIdx + Ceiling((TopIdx - BottomIdx) / 2)
        tmpR = (Target = TargetArr(MiddleIdx))
        If Target > TargetArr(MiddleIdx) Then
            BottomIdx = MiddleIdx + 1
        Else
            TopIdx = MiddleIdx - 1
        End If
    Loop While Not tmpR And TopIdx >= BottomIdx
    If tmpR Then
        BinarySearch = MiddleIdx
    Else
        BinarySearch = -1
    End If
End Function
Private Function Ceiling(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Ceiling = tmpResult - ((value <> tmpResult) And (value > 0))
End Function
' ---------------------
' End VM class
' ---------------------

