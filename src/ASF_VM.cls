VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_VM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------
' VM class
' ---------------------
' Class Module: VM (AST executor)
Option Explicit
Private Const VTP_MAX_DEPTH As Long = 8    ' safety depth limit
Private Const VTP_MAX_ITEMS_INLINE As Long = 8 ' prefer inline for small containers
Private VERBOSE_ As Boolean
Private GLOBALS_ As ASF_Globals
Private OUTPUT__ As Variant
Private INJECTED_VARIABLES_ As Collection
Private ENABLE_CALL_TRACE_ As Boolean
Private CAST_INJECTED_VARIABLES_ As Boolean
Private APPLICATION_ACCESS_ As Boolean
Private Const PLACEHOLDER_CHR_ As String = "$"
Private tmpVar As Variant
Private emptyColl As New Collection

' Executes AST nodes produced by Compiler (AST).
' Uses Map node types and ScopeStack. Logs to gRuntimeLog.
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
    ElseIf TypeName(VarValue) = "Collection" Then
        CastInjectedVar = CollToASF(VarValue)
    ElseIf TypeName(VarValue) = "ASF_Map" Then
        Set CastInjectedVar = VarValue
    ElseIf TypeName(VarValue) = "ASF_RegexEngine" Then
        Set CastInjectedVar = VarValue
    ElseIf IsObject(VarValue) Then
        CastInjectedVar = TypeName(VarValue)
    Else
        CastInjectedVar = CastType(VarValue)
    End If
End Function
Private Function CastType(aValue As Variant) As Variant
    If TypeName(aValue) = "String" Then
        CastType = CStr(aValue)
    ElseIf TypeName(aValue) = "Boolean" Then
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
                ElseIf TypeName(arr(i, ii)) = "Collection" Then
                    tmpElm(1 + ii - lb2) = CollToASF(arr(i, ii))(1)
                ElseIf IsObject(arr(i, ii)) Then
                    tmpElm(1 + ii - lb2) = TypeName(arr(i, ii))
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
            ElseIf TypeName(arr(i)) = "Collection" Then
                tmpResult(1 + i - lb) = CollToASF(arr(i))(1)
            ElseIf IsObject(arr(i)) Then
                tmpResult(1 + i - lb) = TypeName(arr(i))
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
    If coll.Count = 0 Then CollToASF = Array(): Exit Function
    Dim i As Long
    ReDim tmpResult(1 To coll.Count)
    For i = 1 To coll.Count
        If IsArray(coll(i)) Then
            If CastElem Then
                tmpResult(i) = ArrayToASF(coll(i))
            Else
                tmpResult(i) = ArrayToASF2(coll(i))
            End If
        ElseIf TypeName(coll(i)) = "Collection" Then
            tmpResult(i) = CollToASF(coll(i), CastElem)
        ElseIf TypeName(coll(i)) = "String" Then
            tmpResult(i) = CStr(coll(i))
        ElseIf TypeName(coll(i)) = "Boolean" Then
            tmpResult(i) = CBool(coll(i))
        ElseIf IsNumeric(coll(i)) Then
            tmpResult(i) = CDbl(coll(i))
        Else
            If CastElem Then
                If Not IsObject(coll(i)) Then
                    tmpResult(i) = coll(i)
                Else
                    tmpResult(i) = TypeName(coll(i))
                End If
            Else
                vAssignment tmpResult(i), coll(i)
            End If
        End If
    Next i
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
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Property Get verbose() As Boolean
    verbose = VERBOSE_
End Property
Public Property Let verbose(aValue As Boolean)
    VERBOSE_ = aValue
End Property

' Node helpers (Map-based)
Private Function MakeNode(nodeType As String) As ASF_Map
    Dim m As New ASF_Map
    m.Add "type", nodeType
    Set MakeNode = m
End Function

Public Sub RunProgramByIndex(idx As Long, Optional params As Variant = Empty)
    GLOBALS_.ASF_InitGlobals
    If idx < 1 Or idx > GLOBALS_.gPrograms.Count Then Exit Sub
    Dim p As Variant: p = GLOBALS_.gPrograms(idx)
    Dim progName As String: progName = p(0)
    Dim stmts As Collection: Set stmts = p(1)
    Dim rawScope As Collection: Set rawScope = p(2)
'    Dim optionBase As Long: optionBase = IIf(UBound(p) >= 3, p(3), 1)
    Dim progScope As ASF_ScopeStack: Set progScope = New ASF_ScopeStack
    progScope.LoadRaw rawScope
'    progScope.SetValue "__option_base", optionBase
    progScope.SetValue "__option_base", 1 'Enforce option base=1
    Dim MathObj As New Collection
    progScope.SetValue "Math", MathObj
    ' Variables injection
    If Not INJECTED_VARIABLES_ Is Nothing Then
        Dim elm As Variant
        For Each elm In INJECTED_VARIABLES_
            progScope.SetValue CStr(elm(LBound(elm))), elm(UBound(elm))
        Next elm
    End If
    ' Placeholders
    If Not IsEmpty(params) Then
        If IsArray(params) Then
            Dim pk As Long
            Dim phid As Long
            For pk = LBound(params) To UBound(params)
                phid = phid + 1
                If CAST_INJECTED_VARIABLES_ And Not APPLICATION_ACCESS_ Then
                    progScope.SetValue "$" & CStr(phid), CastInjectedVar(params(pk))
                Else
                    If IsArray(params(pk)) And Not IsObject(params(pk)) Then
                        progScope.SetValue "$" & CStr(phid), ArrayToASF2(params(pk))
                    Else
                        progScope.SetValue "$" & CStr(phid), params(pk)
                    End If
                End If
            Next pk
        Else
            If CAST_INJECTED_VARIABLES_ And Not APPLICATION_ACCESS_ Then
                progScope.SetValue "$1", CastInjectedVar(params)
            Else
                progScope.SetValue "$1", params
            End If
        End If
    End If
    progScope.Push
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "RUN Program: " & progName
    Dim i As Long
    For i = 1 To stmts.Count
        Dim ctrl As String
        ctrl = ExecuteStmtNode(stmts(i), progScope)
        If ctrl = "RETURN" Then
            ' Trace: Push call
            PushCallTrace progName, emptyColl
            vAssignment OUTPUT__, progScope.GetValue("__return")
            ' Trace: Pop call with return value
            PopCallTrace progName, OUTPUT__
            Exit For
        End If
        If ctrl = "ERR" Then Exit For
    Next i
    progScope.Pop
End Sub

' Execute a statement node (Map). Return control signals: vbnullstring, "BREAK", "CONTINUE", "RETURN", "ERR"
Private Function ExecuteStmtNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    On Error GoTo ErrHandler
    Dim tP As String: tP = node.GetValue("type")
    Dim rval As Variant
    Dim i As Long
    Select Case tP
        Case "Print"
            Dim args As Collection: Set args = node.GetValue("args")
            Dim outParts As New Collection
            For i = 1 To args.Count
                Dim v As Variant: vAssignment v, EvalExprNode(args(i), progScope)
                outParts.Add ValueToStringForPrint(v)
            Next i
            Dim sb As String: sb = vbNullString
            For i = 1 To outParts.Count
                If i > 1 Then sb = sb & ", "
                sb = sb & outParts(i)
            Next i
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "PRINT:" & sb
            Debug.Print sb
            ExecuteStmtNode = vbNullString
            Exit Function

        Case "Assign"
            Dim left As ASF_Map: Set left = node.GetValue("left")
            Dim right As ASF_Map: Set right = node.GetValue("right")
            vAssignment rval, EvalExprNode(right, progScope)
            HandleAssignment left, rval, progScope
            ExecuteStmtNode = vbNullString
            Exit Function

        Case "ExprStmt"
            Dim res As Variant: vAssignment res, EvalExprNode(node.GetValue("expr"), progScope)
            ExecuteStmtNode = vbNullString
            Exit Function

        Case "If"
            ExecuteStmtNode = ExecIfNode(node, progScope)
            Exit Function

        Case "For"
            ExecuteStmtNode = ExecForNode(node, progScope)
            Exit Function

        Case "While"
            ExecuteStmtNode = ExecWhileNode(node, progScope)
            Exit Function

        Case "Break"
            ExecuteStmtNode = "BREAK": Exit Function
        Case "Continue"
            ExecuteStmtNode = "CONTINUE": Exit Function

        Case "Return"
            Dim rex As ASF_Map: Set rex = node.GetValue("expr")
            If Not rex Is Nothing Then vAssignment rval, EvalExprNode(rex, progScope) Else rval = Empty
            progScope.SetValue "__return", rval
            ExecuteStmtNode = "RETURN": Exit Function

        Case "TryCatch"
            ExecuteStmtNode = ExecTryCatchNode(node, progScope)
            Exit Function

        Case "Switch"
            ExecuteStmtNode = ExecSwitchNode(node, progScope)
            Exit Function
        Case "ClassDecl"
            ExecuteStmtNode = ExecClassDecl(node, progScope)
            Exit Function
        Case "ArrayDestructuring"
            ExecuteStmtNode = ExecArrayDestruct(node, progScope)
            Exit Function
        Case "Import"
            ExecuteStmtNode = ExecModImport(node, progScope)
            Exit Function
        Case "Export"
            ExecuteStmtNode = ExecModExport(node, progScope)
            Exit Function
        Case Else
            ' unknown node type
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "Unknown statement node: " & tP
            ExecuteStmtNode = vbNullString
            Exit Function
    End Select

ErrHandler:
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM statement error: " & err.Description
    err.Clear
    ExecuteStmtNode = "ERR"
End Function

Private Sub AssignToArray(arrName As String, idxV As Variant, val As Variant, progScope As ASF_ScopeStack)
    Dim arr As Variant
    Dim pos As Long
    Dim ub As Long
    Dim lb As Long
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    
    arr = progScope.GetValue(arrName)
    pos = CLng(idxV)
    If pos < optionBase Then
        err.Raise vbObjectError + 5001, "VM.AssignToArray", "Invalid array index (must be >=1)"
    End If
    If Not IsArray(arr) Then
        If IsEmpty(arr) Then
            arr = Array()
        Else
            err.Raise vbObjectError + 5002, "VM.AssignToArray", "Not an array"
        End If
    End If
    ub = -1
    If IsArray(arr) Then
        ub = UBound(arr)
    End If
    If ub < 0 Then
        ReDim arr(1 To pos)
    Else
        lb = LBound(arr)
        If pos > (ub - lb + 1) Then
            ReDim Preserve arr(lb To lb + (pos - optionBase))
        End If
    End If
    arr(LBound(arr) + (pos - optionBase)) = val
    progScope.SetValue arrName, arr
End Sub

Private Function EvalMemberNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    If Not node.Exists("prop") And Not node.Exists("property") Then _
        err.Raise vbObjectError + 5010, "VM.EvalMemberNode", "Missing prop/property in Member node"
    
    On Error GoTo ErrHandler
    Dim propKey As String: propKey = IIf(node.Exists("prop"), "prop", "property")
    Dim prop As String: prop = CStr(node.GetValue(propKey))

    Dim baseVal As Variant
    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)

    If TypeName(baseVal) = "ASF_Map" Then
        'Eval static method on a class
        If baseVal.Exists("classMethods") Then
            Dim classMethods As Collection
            Dim method_ As Variant
            Set classMethods = baseVal.GetValue("classMethods")
            For Each method_ In classMethods
                If method_.GetValue("name") = prop Then
                    vAssignment EvalMemberNode, method_
                    Exit For
                End If
            Next method_
            Set classMethods = Nothing
        Else
            vAssignment EvalMemberNode, baseVal.GetValue(prop)
        End If
        Exit Function
    End If

    ' If base is an ASF_ScopeStack, try to read property if present
    If TypeName(baseVal) = "ASF_ScopeStack" Then
        If baseVal.Has(CStr(prop)) Then
            vAssignment EvalMemberNode, baseVal.GetValue(CStr(prop))
        Else
            vAssignment EvalMemberNode, Empty
        End If
        Exit Function
    End If

    ' If base is an array/string, the member can be a call with missing parents
    If TypeName(baseVal) = "String" Or IsArray(baseVal) Then
        If IsObject(baseVal) Then 'An object identified as an array
            vAssignment EvalMemberNode, CallObjectMethod(baseVal, prop)
        Else
            Dim tmpNode As ASF_Map: Set tmpNode = MakeNode("Call")
            Dim tmpArgs As New Collection
            tmpNode.SetValue "callee", node
            tmpNode.SetValue "args", tmpArgs
            vAssignment EvalMemberNode, EvalExprNode(tmpNode, progScope)
            Set tmpArgs = Nothing
        End If
        Exit Function
    End If
    
    ' Delegate to Application.
    vAssignment EvalMemberNode, CallObjectMethod(baseVal, prop)
    Exit Function

ErrHandler:
    ' Preserve existing explicit map-access errors; convert other unexpected errors to Empty.
    If err.Number = (vbObjectError + 5007) Then
        err.Raise err.Number, err.Source, err.Description
    Else
        err.Clear
        vAssignment EvalMemberNode, Empty
    End If
End Function
Private Function EvalIndexNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    If Not node.Exists("index") Then err.Raise vbObjectError + 5011, "VM.EvalIndexNode", "Missing index in Index node"
    Dim baseVal As Variant
    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)
    Dim idxV As Variant: vAssignment idxV, EvalExprNode(node.GetValue("index"), progScope)
    Dim pos As Long: pos = CLng(idxV)
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.EvalIndexNode", "Invalid index (must >= base)"
    If IsArray(baseVal) Then
        Dim lb As Long: lb = LBound(baseVal)
        Dim ub As Long: ub = UBound(baseVal)
        Dim adjustedPos As Long: adjustedPos = lb + (pos - optionBase)
        If adjustedPos > ub Then err.Raise vbObjectError + 5005, "VM.EvalIndexNode", "Index out of bounds"
        vAssignment EvalIndexNode, baseVal(adjustedPos)
    ElseIf TypeName(baseVal) = "ASF_Map" Then
        vAssignment EvalIndexNode, baseVal.GetValue(pos)
    Else
        err.Raise vbObjectError + 5006, "VM.EvalIndexNode", "Cannot index non-array/non-map"
    End If
End Function

Private Function EvalArrayNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    Dim items As Collection: Set items = node.GetValue("items")
    Dim arr As Variant
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    
    If items Is Nothing Or items.Count = 0 Then
        ReDim arr(-1 To -1)
    Else
        ' First pass: count total elements (accounting for spreads)
        Dim totalCount As Long: totalCount = 0
        Dim ii As Long
        Dim itemNode As ASF_Map
        Dim spreadArr As Variant
        
        For ii = 1 To items.Count
            Set itemNode = items(ii)
            If TypeName(itemNode) = "ASF_Map" And itemNode.GetValue("type") = "SpreadElement" Then
                ' Evaluate the spread argument
                vAssignment spreadArr, EvalExprNode(itemNode.GetValue("argument"), progScope)
                
                ' JavaScript behavior: spread strings into characters
                If TypeName(spreadArr) = "String" Then
                    totalCount = totalCount + Len(spreadArr)
                ElseIf IsArray(spreadArr) Then
                    Dim spreadLen As Long
                    If UBound(spreadArr) <> -1 Then
                        spreadLen = UBound(spreadArr) - LBound(spreadArr) + 1
                        If spreadLen > 0 Then totalCount = totalCount + spreadLen
                    End If
                Else
                    ' Non-iterable: treat as single element if not empty
                    If Not IsEmpty(spreadArr) Then
                        totalCount = totalCount + 1
                    End If
                End If
            Else
                totalCount = totalCount + 1
            End If
        Next ii
        
        ' Allocate array
        If totalCount = 0 Then
            ReDim arr(optionBase To optionBase)
        Else
            ReDim arr(optionBase To optionBase + totalCount - 1)
        End If
        
        ' Second pass: fill array
        Dim writeIdx As Long: writeIdx = 0
        Dim jj As Long
        Dim ci As Long
        Dim StrLen As Long
        
        For ii = 1 To items.Count
            Set itemNode = items(ii)
            If TypeName(itemNode) = "ASF_Map" And itemNode.GetValue("type") = "SpreadElement" Then
                ' Spread the elements
                vAssignment spreadArr, EvalExprNode(itemNode.GetValue("argument"), progScope)
                
                If TypeName(spreadArr) = "String" Then
                    ' JavaScript behavior: spread string into individual characters
                    StrLen = Len(spreadArr)
                    For ci = 1 To StrLen
                        vAssignment arr(optionBase + writeIdx), mid$(spreadArr, ci, 1)
                        writeIdx = writeIdx + 1
                    Next ci
                ElseIf IsArray(spreadArr) Then
                    ' Spread array elements
                    If UBound(spreadArr) <> -1 Then
                        For jj = LBound(spreadArr) To UBound(spreadArr)
                            vAssignment arr(optionBase + writeIdx), spreadArr(jj)
                            writeIdx = writeIdx + 1
                        Next jj
                    End If
                Else
                    ' Non-iterable: add as single element (lenient behavior)
                    vAssignment arr(optionBase + writeIdx), spreadArr
                    writeIdx = writeIdx + 1
                End If
            Else
                ' Regular element
                vAssignment arr(optionBase + writeIdx), EvalExprNode(itemNode, progScope)
                writeIdx = writeIdx + 1
            End If
        Next ii
    End If
    
    vAssignment EvalArrayNode, arr
    Erase arr
End Function

Private Function ParsePath(ByVal path As String) As Collection
    Dim col As New Collection
    Dim pos As Long: pos = 1
    Dim lenP As Long: lenP = Len(path)
    Dim currentPart As String
    While pos <= lenP
        Dim ch As String: ch = mid(path, pos, 1)
        If ch = "." Then
            If currentPart <> vbNullString Then col.Add currentPart
            currentPart = vbNullString
            pos = pos + 1
        ElseIf ch = "[" Then
            If currentPart <> vbNullString Then col.Add currentPart
            currentPart = vbNullString
            Dim startB As Long: startB = pos
            pos = pos + 1
            While pos <= lenP And mid(path, pos, 1) <> "]"
                pos = pos + 1
            Wend
            If pos > lenP Then err.Raise vbObjectError + 5003, "VM.ParsePath", "Unmatched bracket in path"
            currentPart = mid(path, startB, pos - startB + 1)
            col.Add currentPart
            currentPart = vbNullString
            pos = pos + 1
        Else
            currentPart = currentPart & ch
            pos = pos + 1
        End If
    Wend
    If currentPart <> vbNullString Then col.Add currentPart
    Set ParsePath = col
End Function

Private Function ResolvePath(node As ASF_Map, progScope As ASF_ScopeStack, Optional forAssign As Boolean = False) As ASF_Map
    Dim res As New ASF_Map
    Dim tP As String: tP = node.GetValue("type")
    If tP = "Variable" Then
        If (InStr(node.GetValue("name"), ".") > 0 Or InStr(node.GetValue("name"), "[") > 0) Then
            ' Fallback for flat names (if Compiler didn't nest)
            Dim path As String: path = node.GetValue("name")
            Dim parts As Collection: Set parts = ParsePath(path) ' Add ParsePath from previous patch
            Dim current As Variant
            vAssignment current, progScope.GetValue(CStr(parts(1)))
            Dim ii As Long
            For ii = 2 To parts.Count - IIf(forAssign, 1, 0)
                Dim part As String: part = parts(ii)
                If left(part, 1) = "[" Then
                    Dim indexStr As String: indexStr = mid(part, 2, Len(part) - 2)
                    Dim indexV As Variant: indexV = EvalVBExpressionWithScope(indexStr, progScope) ' Twist: use VBEXPR for fallback computed
                    current = GetIndexedValue(current, indexV, progScope)
                Else
                    If TypeName(current) = "ASF_Map" Then
                        vAssignment current, current.GetValue(part)
                    Else
                        err.Raise vbObjectError + 5007, "VM.ResolvePath", "Cannot access property on non-map"
                    End If
                End If
            Next ii
            If forAssign Then
                res.SetValue "container", current
                res.SetValue "key", parts(parts.Count)
                res.SetValue "kind", IIf(left(parts(parts.Count), 1) = "[", "index", "prop")
            Else
                res.SetValue "value", current
            End If
        Else
            If forAssign Then
                res.SetValue "container", progScope
                res.SetValue "key", node.GetValue("name")
                res.SetValue "kind", "scopeVar"
            Else
                res.SetValue "value", progScope.GetValue(node.GetValue("name"))
            End If
        End If
    Else
        ' Nested nodes (preferred)
        Select Case tP
            Case "Member"
                Dim baseRes As ASF_Map: Set baseRes = ResolvePath(node.GetValue("base"), progScope, forAssign)
                If forAssign Then
                    ' Check if base resolution returned a direct value (e.g., from Call)
                    ' instead of a container reference
                    If baseRes.Exists("value") And Not baseRes.Exists("container") Then
                        ' Base is a Call or similar that returned its result as "value"
                        ' Use that value as the container
                        res.SetValue "container", baseRes.GetValue("value")
                    Else
                        ' Normal case - use container from base resolution
                        res.SetValue "container", baseRes("container")
                    End If
                    res.SetValue "key", node.GetValue("prop")
                    res.SetValue "kind", "prop"
                    Dim pKey As String
                    Dim pKind As String
                    ' carry forward parent property/key info if present on the base resolution
                    If baseRes.Exists("key") Then
                        pKey = baseRes("key")
                        res.SetValue "parentKey", pKey
                    End If
                    If baseRes.Exists("kind") Then
                        pKind = baseRes("kind")
                        res.SetValue "parentKind", pKind
                    End If
                    res.SetValue "parentObj", baseRes
                Else
                    res.SetValue "value", EvalMemberNode(node, progScope)
                End If
            Case "Index"
                Dim baseRes2 As ASF_Map: Set baseRes2 = ResolvePath(node.GetValue("base"), progScope, forAssign)
                If forAssign Then
                    ' We need to preserve the parent container _and_ the parent property name (if any)
                    ' so that assignments to an indexed property (e.g. o.a[2]) can write back correctly.
                    If baseRes2.Exists("value") And Not baseRes2.Exists("container") Then
                        res.SetValue "container", baseRes2.GetValue("value")
                    Else
                        res.SetValue "container", baseRes2("container")
                    End If
                    res.SetValue "key", EvalExprNode(node.GetValue("index"), progScope)
                    res.SetValue "kind", "index"
                    ' carry forward parent property/key info if present on the base resolution
                    If baseRes2.Exists("key") Then res.SetValue "parentKey", baseRes2("key")
                    If baseRes2.Exists("kind") Then res.SetValue "parentKind", baseRes2("kind")
                    res.SetValue "parentObj", baseRes2
                Else
                    res.SetValue "value", EvalIndexNode(node, progScope)
                End If
            Case "Call"
                ' A Call node in an assignment path (e.g., obj.method().prop = value)
                ' Execute the call and return its result as "value"
                res.SetValue "value", EvalExprNode(node, progScope)
        End Select
    End If
    Set ResolvePath = res
End Function

Private Function GetIndexedValue(baseVal As Variant, idxV As Variant, progScope As ASF_ScopeStack) As Variant
    Dim pos As Long: pos = CLng(idxV)
    Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
    If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.GetIndexedValue", "Invalid index"
    If IsArray(baseVal) Then
        Dim lb As Long: lb = LBound(baseVal)
        Dim adjustedPos As Long: adjustedPos = lb + (pos - optionBase)
        If adjustedPos > UBound(baseVal) Then err.Raise vbObjectError + 5005, "VM.GetIndexedValue", "Index out of bounds"
        vAssignment GetIndexedValue, baseVal(adjustedPos)
    ElseIf TypeName(baseVal) = "ASF_Map" Then
        vAssignment GetIndexedValue, baseVal.GetValue(pos)
    Else
        err.Raise vbObjectError + 5006, "VM.GetIndexedValue", "Cannot index non-array/non-map"
    End If
End Function

Private Function ResolveLValue(node As ASF_Map, progScope As ASF_ScopeStack) As ASF_Map
    Set ResolveLValue = ResolvePath(node, progScope, True)
    If node.GetValue("type") = "Index" Then
        If node.GetValue("index").GetValue("type") <> "Literal" Then
            ResolveLValue.Add "computed", True ' Twist: flag for lazy
        End If
    End If
End Function

' ---------------------
' ExecIfNode: Evaluate condition, then execute proper block
' ---------------------
Private Function ExecIfNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim condNode As ASF_Map: Set condNode = node.GetValue("cond")
    If IsTruthy(EvalExprNode(condNode, progScope)) Then
        Dim thenStmts As Collection: Set thenStmts = node.GetValue("then")
        Dim si As Long
        For si = 1 To thenStmts.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(thenStmts(si), progScope)
            If ctrl <> vbNullString Then ExecIfNode = ctrl: Exit Function
        Next si
        ExecIfNode = vbNullString
        Exit Function
    End If
    Dim elseifConds As Collection: Set elseifConds = node.GetValue("elseif_conds")
    Dim elseifBlocks As Collection: Set elseifBlocks = node.GetValue("elseif_blocks")
    Dim idx As Long
    For idx = 1 To elseifConds.Count
        If IsTruthy(EvalExprNode(elseifConds(idx), progScope)) Then
            Dim bl As Collection: Set bl = elseifBlocks(idx)
            Dim si2 As Long
            For si2 = 1 To bl.Count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(bl(si2), progScope)
                If ctrl2 <> vbNullString Then ExecIfNode = ctrl2: Exit Function
            Next si2
            ExecIfNode = vbNullString
            Exit Function
        End If
    Next idx
    If node.GetValue("hasElse") Then
        Dim els As Collection: Set els = node.GetValue("else")
        Dim ee As Long
        For ee = 1 To els.Count
            Dim ctrl3 As String: ctrl3 = ExecuteStmtNode(els(ee), progScope)
            If ctrl3 <> vbNullString Then ExecIfNode = ctrl3: Exit Function
        Next ee
    End If
    ExecIfNode = vbNullString
End Function

' ---------------------
' ExecForNode
' ---------------------
Private Function ExecForNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    On Error GoTo ExecForErr

    Dim kind As String
    If node.Exists("kind") Then
        kind = LCase$(CStr(node.GetValue("kind")))
    Else
        kind = "classic"
    End If

    ' Classic for(...) branch variables (may be Nothing for in/of)
    Dim initNode As ASF_Map: Set initNode = IIf(node.Exists("init"), node.GetValue("init"), Nothing)
    Dim condNode As ASF_Map: Set condNode = IIf(node.Exists("cond"), node.GetValue("cond"), Nothing)
    Dim stepNode As ASF_Map: Set stepNode = IIf(node.Exists("step"), node.GetValue("step"), Nothing)
    Dim body As Collection: Set body = node.GetValue("body")

    If kind = "in" Or kind = "of" Then
        ' for-in / for-of semantics
        Dim eachNode As ASF_Map: Set eachNode = node.GetValue("each")
        Dim iterNode As ASF_Map: Set iterNode = node.GetValue("iter")

        ' Evaluate iterable once (like JS for-of), per semantics
        Dim iterVal As Variant
        vAssignment iterVal, EvalExprNode(iterNode, progScope)

        ' Option base (for index translation)
        Dim optBase As Long: optBase = progScope.GetValue("__option_base")

        ' Prepare iterator lists
        Dim keysCol As Collection: Set keysCol = New Collection

        ' Arrays
        If IsArray(iterVal) Then
            Dim lb As Long: lb = LBound(iterVal)
            Dim ub As Long: ub = UBound(iterVal)
            Dim idxA As Long
            For idxA = lb To ub
                ' for-in yields indices (adjusted); for-of yields values
                If kind = "in" Then
                    keysCol.Add (optBase + (idxA - lb))
                Else
                    keysCol.Add iterVal(idxA)
                End If
            Next idxA

        ' Strings
        ElseIf VarType(iterVal) = vbString Then
            Dim str As String: str = CStr(iterVal)
            Dim nch As Long: nch = Len(str)
            Dim iChar As Long
            For iChar = 1 To nch
                If kind = "in" Then
                    keysCol.Add (optBase + (iChar - 1))   ' index-like
                Else
                    keysCol.Add mid$(str, iChar, 1)
                End If
            Next iChar

        ' ASF_Map objects (attempt to obtain keys)
        ElseIf Not IsEmpty(iterVal) And Not IsNull(iterVal) And IsObject(iterVal) And TypeName(iterVal) = "ASF_Map" Then
            Set keysCol = GetMapKeys(iterVal) ' returns Collection of keys (strings or numeric)
            ' for-of: replace keysCol with values collection
            If kind = "of" Then
                Dim valsCol As New Collection
                Dim kk As Long
                For kk = 1 To keysCol.Count
                    Dim kkey As Variant: kkey = keysCol(kk)
                    Dim valv As Variant
                    On Error Resume Next
                    vAssignment valv, iterVal.GetValue(CStr(kkey))
                    On Error GoTo ExecForErr
                    valsCol.Add valv
                Next kk
                Set keysCol = valsCol
            End If

        Else
            ' Unknown iterable -> nothing to iterate
            Set keysCol = New Collection
        End If

        ' Iterate over keysCol
        Dim itCount As Long: itCount = keysCol.Count
        Dim iterIndex As Long
        For iterIndex = 1 To itCount
            Dim curVal As Variant: vAssignment curVal, keysCol(iterIndex)

            ' Bind the 'each' target (if provided) -> we need to assign LValue
            If Not eachNode Is Nothing Then
                ' For "in": assign key (for arrays this is numeric index adjusted to option base)
                ' For "of": assign value
                Dim toAssign As Variant: vAssignment toAssign, curVal

                ' If 'each' is a simple IDENT variable or any lvalue AST node,
                ' reuse HandleAssignment which performs ResolveLValue and writes back to container.
                ' But HandleAssignment expects an AST node (Variable/Member/Index)
                On Error Resume Next
                Call HandleAssignment(eachNode, toAssign, progScope)
                If err.Number <> 0 Then
                    ' If assignment failed (e.g., eachNode not an LValue) clear error and skip binding
                    err.Clear
                End If
                On Error GoTo ExecForErr
            End If

            ' Execute body statements
            Dim s As Long
            For s = 1 To body.Count
                Dim ctrl As String: ctrl = ExecuteStmtNode(body(s), progScope)
                If ctrl = "BREAK" Then
                    GoTo ForInEnd
                ElseIf ctrl = "CONTINUE" Then
                    Exit For ' continue outer loop: move to next iterIndex
                ElseIf ctrl = "RETURN" Or ctrl = "ERR" Then
                    ExecForNode = ctrl
                    Exit Function
                End If
            Next s
        Next iterIndex

ForInEnd:
        ExecForNode = vbNullString
        Exit Function
    End If

    ' ---------- classic for(init; cond; step) fallback ----------
    If Not initNode Is Nothing Then
        Dim initType As String
        initType = initNode.GetValue("type")
        Select Case initType
            Case "Assign", "Print", "If", "For", "While", "TryCatch", "Switch", "Return", "Break", "Continue", "ExprStmt"
                ' Already a statement node — execute directly so side-effects happen.
                ExecuteStmtNode initNode, progScope
            Case Else
                ' Expression node — wrap into ExprStmt for side-effect evaluation.
                ExecuteStmtNode MakeNodeExprStmt(initNode), progScope
        End Select
    End If

ForStart:
    Dim condOk As Boolean: condOk = True
    If Not condNode Is Nothing Then
        condOk = IsTruthy(EvalExprNode(condNode, progScope))
    End If
    If Not condOk Then GoTo ForEnd

    Dim s2 As Long
    For s2 = 1 To body.Count
        Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(body(s2), progScope)
        If ctrl2 = "BREAK" Then GoTo ForEnd
        If ctrl2 = "CONTINUE" Then Exit For
        If ctrl2 = "RETURN" Or ctrl2 = "ERR" Then ExecForNode = ctrl2: Exit Function
    Next s2

    If Not stepNode Is Nothing Then
        Dim stepType As String
        stepType = stepNode.GetValue("type")
        Select Case stepType
            Case "Assign", "Print", "If", "For", "While", "TryCatch", "Switch", "Return", "Break", "Continue", "ExprStmt"
                ExecuteStmtNode stepNode, progScope
            Case Else
                ExecuteStmtNode MakeNodeExprStmt(stepNode), progScope
        End Select
    End If
    GoTo ForStart

ForEnd:
    ExecForNode = vbNullString
    Exit Function

ExecForErr:
    ' propagate error as VM runtime error marker
    ExecForNode = "ERR"
    If err.Number <> 0 Then
        ' You may prefer to re-raise or log error — currently we return "ERR" for caller to handle
    End If
End Function

' Helper to wrap an expression node into an ExprStmt node
Private Function MakeNodeExprStmt(expr As ASF_Map) As ASF_Map
    Dim m As ASF_Map: Set m = MakeNode("ExprStmt")
    m.SetValue "expr", expr
    Set MakeNodeExprStmt = m
End Function

' ---------------------
' ExecWhileNode
' ---------------------
Private Function ExecWhileNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim condNode As ASF_Map: Set condNode = node.GetValue("cond")
    Dim body As Collection: Set body = node.GetValue("body")

WhileStart:
    If Not IsTruthy(EvalExprNode(condNode, progScope)) Then GoTo WhileEnd
    Dim i As Long
    For i = 1 To body.Count
        Dim ctrl As String: ctrl = ExecuteStmtNode(body(i), progScope)
        If ctrl = "BREAK" Then GoTo WhileEnd
        If ctrl = "CONTINUE" Then Exit For
        If ctrl = "RETURN" Or ctrl = "ERR" Then ExecWhileNode = ctrl: Exit Function
    Next i
    GoTo WhileStart

WhileEnd:
    ExecWhileNode = vbNullString
End Function

' ---------------------
' ExecTryCatchNode
' ---------------------
Private Function ExecTryCatchNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    On Error GoTo TryErr
    Dim tryStmts As Collection: Set tryStmts = node.GetValue("try")
    Dim i As Long
    For i = 1 To tryStmts.Count
        Dim ctrl As String: ctrl = ExecuteStmtNode(tryStmts(i), progScope)
        If ctrl = "RETURN" Or ctrl = "ERR" Then
            ExecTryCatchNode = ctrl
            If ctrl = "ERR" Then GoTo TryErr
            Exit Function
        End If
    Next i
    ExecTryCatchNode = vbNullString
    Exit Function

TryErr:
    err.Clear
    Dim catchStmts As Collection: Set catchStmts = node.GetValue("catch")
    If Not catchStmts Is Nothing Then
        Dim j As Long
        For j = 1 To catchStmts.Count
            Dim cctrl As String: cctrl = ExecuteStmtNode(catchStmts(j), progScope)
            If cctrl = "RETURN" Or cctrl = "ERR" Then ExecTryCatchNode = cctrl: Exit Function
        Next j
    End If
    ExecTryCatchNode = vbNullString
End Function

' ---------------------
' ExecSwitchNode
' ---------------------
Private Function ExecSwitchNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim switchVal As Variant: switchVal = EvalExprNode(node.GetValue("expr"), progScope)
    Dim cases As Collection: Set cases = node.GetValue("cases")
    Dim i As Long
    Dim Match As Boolean
    For i = 1 To cases.Count
        Dim pair As Collection: Set pair = cases(i)
        Dim caseExpr As ASF_Map: Set caseExpr = pair(1)
        Dim blockStmts As Collection: Set blockStmts = pair(2)
        Match = (EvalExprNode(caseExpr, progScope) = switchVal)
        If Match Then
            Dim s As Long
            For s = 1 To blockStmts.Count
                Dim ctrl As String: ctrl = ExecuteStmtNode(blockStmts(s), progScope)
                If ctrl = "BREAK" Then ExecSwitchNode = vbNullString: Exit Function
                If ctrl = "RETURN" Or ctrl = "ERR" Then ExecSwitchNode = ctrl: Exit Function
            Next s
            Exit For
        End If
    Next i
    If Not Match Then
        Dim defBlk As Collection: Set defBlk = node.GetValue("default")
        If Not defBlk Is Nothing Then
            Dim d As Long
            For d = 1 To defBlk.Count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(defBlk(d), progScope)
                If ctrl2 = "BREAK" Then ExecSwitchNode = vbNullString: Exit Function
                If ctrl2 = "RETURN" Or ctrl2 = "ERR" Then ExecSwitchNode = ctrl2: Exit Function
            Next d
        End If
    End If
    ExecSwitchNode = vbNullString
End Function

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
Private Function EvalObjectNode(node As ASF_Map, progScope As ASF_ScopeStack) As ASF_Map
    Dim items As Collection: Set items = node.GetValue("items")
    Dim resultMap As New ASF_Map
    
    If items Is Nothing Then
        Set EvalObjectNode = resultMap
        Exit Function
    End If
    
    Dim ii As Long
    Dim pair As Collection
    Dim elm As Variant
    Dim keyStr As String
    Dim valNode As ASF_Map
    Dim spreadSrc As Variant
    Dim spreadKeys As Collection
    Dim jj As Long
    Dim si As Long
    
    For ii = 1 To items.Count
        Set pair = items(ii)
        keyStr = CStr(pair(1))
        
        If keyStr = "__SPREAD__" Then
            ' This is a spread property
            Set valNode = pair(2)
            If valNode.GetValue("type") = "SpreadProperty" Then
                ' Evaluate the spread source
                vAssignment spreadSrc, EvalExprNode(valNode.GetValue("argument"), progScope)
                
                If TypeName(spreadSrc) = "ASF_Map" Then
                    ' JavaScript behavior: merge all enumerable properties
                    Set spreadKeys = spreadSrc.keys()
                    If spreadKeys.Count > 0 Then
                        For Each elm In spreadKeys
                            resultMap.SetValue CStr(elm), spreadSrc.GetValue(CStr(elm))
                        Next elm
                    End If
                ElseIf IsArray(spreadSrc) Then
                    ' JavaScript behavior: array spreads as numeric keys (0-based)
                    For jj = LBound(spreadSrc) To UBound(spreadSrc)
                        Dim keyIndex As String
                        keyIndex = CStr(jj - LBound(spreadSrc) + 1) ' Convert to 1-based  DIFFERS FROM JS
                        resultMap.SetValue keyIndex, spreadSrc(jj)
                    Next jj
                ElseIf TypeName(spreadSrc) = "String" Then
                    ' JavaScript behavior: string spreads as character index (0-based)
                    For si = 1 To Len(spreadSrc)
                        resultMap.SetValue CStr(si), mid$(spreadSrc, si, 1) ' Convert to 1-based  DIFFERS FROM JS
                    Next si
                Else
                    ' JavaScript behavior: primitives (null, undefined, numbers, booleans) are ignored
                    ' Do nothing - matches JavaScript {...null}, {...undefined}, {...42}, etc.
                End If
            End If
        Else
            ' Regular property
            vAssignment valNode, pair(2)
            resultMap.SetValue keyStr, EvalExprNode(valNode, progScope)
        End If
    Next ii
    
    Set EvalObjectNode = resultMap
End Function
' ---------------------
' Expression evaluator: Evaluate Expr AST nodes to runtime values
' ---------------------
Private Function EvalExprNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    If node Is Nothing Then EvalExprNode = Empty: Exit Function
    Dim tP As String: tP = node.GetValue("type")
    Dim items As Collection
    Dim tmpResult As Variant
    Dim pi As Long
    Dim funcIdx As Long
    Dim argColl As Collection
    Dim tmpMapObj As ASF_Map
    Dim calleeExpr As ASF_Map
    Dim thisVal As Variant
    Dim fname As String
    Dim optBase As Long
    Dim ii As Long
    Dim k As Long
    
    Dim sep As String
    Dim lb_val As Long, ub_val As Long, len_val As Long
    Dim parts() As String
    Dim ii_j As Long
    Dim valj As Variant
    Dim popped2 As Variant
    Dim newA_s() As Variant
    Dim idx_s As Long, k_s As Long
    
    Dim optBase_u2 As Long
    Dim alb As Long, aub As Long
    Dim len_base As Long, totalCount As Long
    Dim outArr() As Variant, writeIdx As Long
    Dim avTmp As Variant
    Dim lb_b As Long, ub_b As Long
    
    Dim startRaw_sp As Long, startN_sp As Long, deleteCount_sp As Long
    Dim removedCount As Long
    Dim tempIn() As Variant, removedArr() As Variant, newTemp() As Variant
    Dim insertCount As Long, newSize As Long
    Dim i_sp As Long, j_sp As Long, cur As Long
    Dim insertedItem As Variant
    Dim removedOut() As Variant
    Dim foundIdx As Long
    Dim si As Long
    
    Select Case tP
        Case "FuncLiteral", "AnonFunc"
            ' Create closure (capture env by reference for shared-write closures)
            Dim cparams As Collection: vAssignment cparams, node.GetValue("params")
            Dim cbody As Collection: vAssignment cbody, node.GetValue("body")
            Dim cmap As ASF_Map
            Set cmap = New ASF_Map
            Dim envCopy As ASF_ScopeStack
            cmap.Add "type", "Closure"
            cmap.SetValue "params", cparams
            cmap.SetValue "body", cbody
            ' capture current scope reference (shared-write)
            Set envCopy = New ASF_ScopeStack
            envCopy.LoadRawByRef progScope.RawByRef
            cmap.SetValue "env", envCopy
            ' Add rest parameter if present
            If node.Exists("restParam") Then
                cmap.SetValue "restParam", node.GetValue("restParam")
            End If
            vAssignment tmpResult, cmap
            GoTo exitfun
        Case "VBAexpr"
            ' Forced VBAexpression node (from @(...)). The node stores the raw VB expression string in "expr".
            Dim rawVB As String: rawVB = node.GetValue("expr")
            vAssignment tmpResult, EvalVBExpressionWithScope(rawVB, progScope)
            GoTo exitfun
        Case "Object"
            ' Build and return a Map containing the evaluated properties
            vAssignment tmpResult, EvalObjectNode(node, progScope)
            GoTo exitfun
        Case "Member"
            vAssignment tmpResult, EvalMemberNode(node, progScope)
            GoTo exitfun
        Case "Index"
            vAssignment tmpResult, EvalIndexNode(node, progScope)
            GoTo exitfun
        Case "Array"
            vAssignment tmpResult, EvalArrayNode(node, progScope)
            GoTo exitfun
        Case "Literal"
            vAssignment tmpResult, node.GetValue("value")
            GoTo exitfun
        Case "Template"
            Dim partsC As Collection: Set partsC = node.GetValue("parts")
            Dim outStr As String
            For pi = 1 To partsC.Count
                Dim pNode As ASF_Map: Set pNode = partsC(pi)
                If pNode Is Nothing Then
                    ' skip
                ElseIf pNode.GetValue("type") = "Literal" Then
                    ' literal part - use its value (could be number/string)
                    Dim litV As Variant: vAssignment litV, pNode.GetValue("value")
                    outStr = outStr & CStr(litV)
                Else
                    ' evaluate expression and append stringified result
                    Dim evs As Variant
                    vAssignment evs, EvalExprNode(pNode, progScope)
                    ' pretty-print arrays/maps/collections; fallback to CStr for primitives
                    If IsArray(evs) Or (IsObject(evs) And TypeName(evs) = "ASF_Map") Then
                        outStr = outStr & ValueToStringForPrint(evs)
                    ElseIf IsNull(evs) Then
                        outStr = outStr & "null"
                    ElseIf IsEmpty(evs) Then
                        outStr = outStr & vbNullString
                    Else
                        outStr = outStr & CStr(evs)
                    End If
                End If
            Next pi
            vAssignment tmpResult, outStr
            GoTo exitfun
        Case "Variable"
            ' Robust variable resolution:
            '  - return actual value from scope (covers closures stored As ASF_Map AST nodes)
            '  - if scope has no value, fallback to named functions table (gFuncTable)
            '  - optional fallback to gFuncObjects registry
            Dim vname As String: vname = CStr(node.GetValue("name"))
            Dim vVal As Variant
            vAssignment vVal, progScope.GetValue(vname)

            ' If scope contains a non-empty value, return it — this covers closures
            ' stored As ASF_Map/AST nodes (FuncLiteral/Closure/etc.).
            If Not IsEmpty(vVal) Then
                ' If it's a Map and looks like a function node, return as callable
                On Error Resume Next
                If TypeName(vVal) = "ASF_Map" Then
                    Dim maybeType As Variant
                    maybeType = Empty
                    On Error Resume Next
                    maybeType = vVal.GetValue("type")
                    On Error GoTo 0
                    If Not IsEmpty(maybeType) Then
                        Dim tt As String: tt = CStr(maybeType)
                        ' Common names for function-like AST nodes / closures:
                        If tt = "FuncLiteral" Or tt = "Function" Or tt = "Closure" Or tt = "AnonFunction" Then
                            vAssignment tmpResult, vVal
                            GoTo exitfun
                        End If
                    End If
                End If
                ' Not necessarily a function-like Map — still return stored value.
                vAssignment tmpResult, vVal
                GoTo exitfun
            End If

            ' Not in scope (or Empty). Fallback: named compiled function table.
            On Error Resume Next
            If GLOBALS_.gFuncTable.Exists(vname) Then
                Dim fIdxVal As Variant: fIdxVal = GLOBALS_.gFuncTable.GetValue(vname)
                On Error GoTo 0
                If Not IsEmpty(fIdxVal) Then
                    funcIdx = CLng(fIdxVal)

                    ' Build a Closure map with the shape CallClosure expects:
                    ' "type" = "Closure", "params" = Collection, "body" = Collection, "env" = ScopeStack
                    Dim closureMap As New ASF_Map
                    closureMap.Add "type", "Closure"

                    ' params: convert stored gFuncParams (likely an array) into a Collection
                    Dim paramsCol As Collection
                    Dim restParamCol As Collection
                    Set paramsCol = New Collection
                    Dim paVar As Variant
                    Dim it As Variant
                    If GLOBALS_.gFuncParams.Exists(vname) Then
                        paVar = GLOBALS_.gFuncParams.GetValue(vname)
                        If IsArray(paVar) Then
                            If ArrayIsInit(paVar) Then
                                For pi = LBound(paVar) To UBound(paVar)
                                    paramsCol.Add CStr(paVar(pi))
                                Next pi
                            Else
                                paramsCol.Add vbNullString
                            End If
                        ElseIf TypeName(paVar) = "Collection" Then
                            ' already a collection — copy it
                            For Each it In paVar
                                paramsCol.Add it
                            Next it
                        End If
                    End If
                    closureMap.SetValue "params", paramsCol
                    
                    ' body and env come from gPrograms(funcIdx)
                    Dim pinfo As Variant
                    On Error Resume Next
                    pinfo = GLOBALS_.gPrograms(funcIdx)
                    If err.Number <> 0 Then
                        err.Clear
                        ' fallback: no program info — return Empty
                        EvalExprNode = Empty
                        Exit Function
                    End If
                    On Error GoTo 0

                    ' pinfo layout: Array(name, stmtsCollection, rawScopeCollection)
                    Dim bodyStmts As Collection
                    Set bodyStmts = pinfo(1)
                    closureMap.SetValue "body", bodyStmts

                    ' env: create ScopeStack and LoadRaw with stored raw scope (so closure has env object)
                    Dim envScope As ASF_ScopeStack
                    Set envScope = New ASF_ScopeStack
                    envScope.LoadRawByRef progScope.RawByRef
                    Set pinfo(2) = envScope.RawByRef
                    closureMap.SetValue "env", envScope

                    ' optionally store a name/funcIdx for debugging
                    closureMap.SetValue "name", vname
                    closureMap.SetValue "funcIdx", funcIdx

                    ' Add rest parameter if stored (index 3 in pinfo array)
                    If IsArray(pinfo) Then
                        If UBound(pinfo) >= 3 Then
                            If Not IsEmpty(pinfo(3)) Then
                                closureMap.SetValue "restParam", pinfo(3)
                            End If
                        End If
                    End If
                    vAssignment tmpResult, closureMap
                    GoTo exitfun
                End If
            End If
            On Error GoTo 0

            ' Not found — log and return Empty
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM: Variable '" & vname & "' is undefined or not callable (returned Empty)."
            vAssignment tmpResult, Empty
            GoTo exitfun
        Case "SuperCall"
            ' super() call - get parent class constructor and execute it
            Dim superArgs As Collection: Set superArgs = node.GetValue("args")
            Dim evaluatedSuper As Collection: Set evaluatedSuper = New Collection
            Dim sa As Long
            For sa = 1 To superArgs.Count
                evaluatedSuper.Add EvalExprNode(superArgs(sa), progScope)
            Next sa
            
            ' Get 'this' from current scope
            Dim thisInstance As Variant
            vAssignment thisInstance, progScope.GetValue("this")
            
            If TypeName(thisInstance) = "ASF_Map" And thisInstance.GetValue("type") = "Instance" Then
                Dim instanceClassName As String: instanceClassName = thisInstance.GetValue("className")
                
                ' Get the class definition
                Dim classDefSuper As Variant
                vAssignment classDefSuper, progScope.GetValue(instanceClassName)
                
                If TypeName(classDefSuper) = "ASF_Map" And classDefSuper.GetValue("type") = "Class" Then
                    ' Get parent class
                    If classDefSuper.Exists("superClass") Then
                        Dim parentClassName As String: parentClassName = classDefSuper.GetValue("superClass")
                        Dim parentClassDef As Variant
                        vAssignment parentClassDef, progScope.GetValue(parentClassName)
                        
                        If TypeName(parentClassDef) = "ASF_Map" And parentClassDef.Exists("constructor") Then
                            Dim parentCtor As ASF_Map: Set parentCtor = parentClassDef.GetValue("constructor")
                            
                            ' Execute parent constructor with this instance and provided args
                            vAssignment tmpMapObj, thisInstance
                            Call ExecuteSuperConstructor(parentCtor, tmpMapObj, evaluatedSuper, progScope)
                        End If
                    End If
                End If
            End If
            
            vAssignment tmpResult, Empty
            GoTo exitfun
        Case "Call"
            ' Call can be either:
            '  - Call with "name" (string) and args (legacy)
            '  - Call with "callee" (an expression node) and args
            '       (new: allows variable holding closures)

            Dim hasCalleeExpr As Boolean
            If node.Exists("name") Then
                fname = CStr(node.GetValue("name"))
            End If
            If fname = vbNullString Then
                ' try to derive name from callee expression when callee is a simple Variable node
                If node.Exists("callee") Then
                    hasCalleeExpr = True
                    Set calleeExpr = node.GetValue("callee")
                    If Not calleeExpr Is Nothing Then
                        If calleeExpr.GetValue("type") = "Variable" Then
                            fname = CStr(calleeExpr.GetValue("name"))
                        End If
                    End If
                End If
            End If

            Dim argsColl As Collection: Set argsColl = node.GetValue("args")
            Dim evaluated As Collection
            Set evaluated = New Collection

            ' Evaluate arguments and expand spreads (JavaScript-compatible)
            Dim argNode As ASF_Map
            Dim spreadVal As Variant
            Dim sci As Long

            For k = 1 To argsColl.Count
                Set argNode = argsColl(k)
                
                If TypeName(argNode) = "ASF_Map" And argNode.GetValue("type") = "SpreadElement" Then
                    ' Evaluate spread and expand
                    vAssignment spreadVal, EvalExprNode(argNode.GetValue("argument"), progScope)
                    
                    If TypeName(spreadVal) = "String" Then
                        ' JavaScript behavior: spread string into individual characters
                        For sci = 1 To Len(spreadVal)
                            evaluated.Add mid$(spreadVal, sci, 1)
                        Next sci
                    ElseIf IsArray(spreadVal) Then
                        ' Spread array elements
                        For si = LBound(spreadVal) To UBound(spreadVal)
                            evaluated.Add spreadVal(si)
                        Next si
                    Else
                        ' Non-iterable: add as single argument (lenient behavior)
                        evaluated.Add spreadVal
                    End If
                Else
                    ' Regular argument
                    evaluated.Add EvalExprNode(argNode, progScope)
                End If
            Next k
            
            If Not hasCalleeExpr Then
                If node.Exists("callee") Then
                    Set calleeExpr = node.GetValue("callee")
                Else
                    Set calleeExpr = Nothing
                End If
            End If
            ' If callee is a Member node -> method-builtins (map/filter/reduce/...)
            If Not calleeExpr Is Nothing Then
                If calleeExpr.Exists("type") And CStr(calleeExpr.GetValue("type")) = "Member" Then
                    ' --- method builtins on Member callee ---
                    Dim propN As String: propN = LCase$(CStr(calleeExpr.GetValue("prop")))
                    Dim baseValLocal As Variant
                    
                    vAssignment baseValLocal, EvalExprNode(calleeExpr.GetValue("base"), progScope)
                    If TypeName(baseValLocal) = "Collection" Then
                        Dim i As Long
                        Dim knownMethod As Boolean
                        For i = LBound(GLOBALS_.builtinArrayMethods) To UBound(GLOBALS_.builtinArrayMethods)
                            knownMethod = (GLOBALS_.builtinArrayMethods(i) = propN)
                            If knownMethod Then Exit For
                        Next i
                        If Not knownMethod Then
                            Dim isSpecialColl As Boolean
                            Dim tmpBase As Variant
                            Dim tmpBaseLocal As Variant
                            vAssignment tmpBaseLocal, calleeExpr.GetValue("base")
                            If tmpBaseLocal.Exists("name") Then
                                isSpecialColl = tmpBaseLocal.GetValue("name") = "Math"
                            End If
                            GoTo CallCollMethod
                        Else
                            baseValLocal = CollToASF(baseValLocal, False)
                        End If
                    End If
                    If IsArray(baseValLocal) Then
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
                                ' map(fn) -> return new array with fn applied to each element (recursively tolerates nested arrays)
                                If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim mapper As Variant: vAssignment mapper, evaluated(1)
                                If TypeName(mapper) <> "ASF_Map" Or mapper.GetValue("type") <> "Closure" Then
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
                            
                                Dim optBaseLocal As Long: optBaseLocal = progScope.GetValue("__option_base")
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
                                                            tmpMapObj, progScope, baseValLocal, optBaseLocal)
                                    vAssignment outArr(optBaseLocal + ii), resVal
                                Next ii
                            
                                vAssignment tmpResult, outArr
                                GoTo exitfun
    
                            Case "filter"
                                ' filter(fn) -> return new array with elements where fn(elem) truthy
                                If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim pred As Variant: vAssignment pred, evaluated(1)
                                If TypeName(pred) <> "ASF_Map" Or pred.GetValue("type") <> "Closure" Then
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
                                Dim ob As Long: ob = progScope.GetValue("__option_base")
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
                            Case "reduce"
                                ' reduce(fn, initial?) -> apply reduction
                                If evaluated.Count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim rfn As Variant: vAssignment rfn, evaluated(1)
                                If TypeName(rfn) <> "ASF_Map" Or rfn.GetValue("type") <> "Closure" Then
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
                            Case "slice"
                                ' slice(start, end) -> non-mutating, respects __option_base
                                Dim s As Long, e As Long
                                Dim lenA As Long
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                            
                                ' number of elements in the array (independent of its LBound)
                                lenA = UBound(baseValLocal) - LBound(baseValLocal) + 1
                            
                                ' runtime option base for arrays (0 or 1 typically)
                                optBase = progScope.GetValue("__option_base")
                            
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
                                    ' Produce minimal array using same option base — keep conservative behavior:
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
                            Case "push"
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
                                Dim newArr() As Variant
                                ReDim newArr(oldLB To oldLB + newLen - 1 - offs)
                                Dim zz As Long
                                For zz = oldLB To oldUB - offs
                                    vAssignment newArr(zz), baseValLocal(zz)
                                Next zz
                                For zz = 1 To insertCount
                                    vAssignment newArr(oldLB + oldLen + zz - 1 - offs), evaluated(zz)
                                Next zz
                                ' write back to original container via ResolvePath
                                WriteToContainer calleeExpr, progScope, newArr
                                vAssignment tmpResult, newLen
                                GoTo exitfun
                            Case "pop"
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                vAssignment tmpResult, popped
                                GoTo exitfun
                            Case "foreach"
                                ' forEach(fn) -> invoke fn(value, index, array) for each element; returns the original array
                                If evaluated.Count = 0 Then
                                    vAssignment tmpResult, Empty
                                    GoTo exitfun
                                End If
                            
                                Dim cb As Variant: vAssignment cb, evaluated(1)
                                If TypeName(cb) <> "ASF_Map" Or cb.GetValue("type") <> "Closure" Then
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
                                optBase = progScope.GetValue("__option_base")
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
                            Case "unique"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_u As Long: optBase_u = progScope.GetValue("__option_base")
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
                            Case "concat"
                                 ' concat: concat arrays/values into the callee array
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_c As Long: optBase_c = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, tmpResult
                                GoTo exitfun
                            Case "tostring", "join"
                                ' toString / join: join(customSep)
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                sep = ", "
                                If evaluated.Count >= 1 Then sep = CStr(evaluated(1))
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
                                    If IsArray(valj) Or (IsObject(valj) And TypeName(valj) = "ASF_Map") Then
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
                            Case "shift"
                                ' shift: remove first element, mutate array, return popped value
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim ob_s As Long: ob_s = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "unshift"
                                ' unshift: add elements at start, returns new length
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                optBase_u2 = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "delete"
                                ' delete(index): remove element at index (user-space) and return true/false
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                optBase_u2 = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "splice"
                                ' splice(start, deleteCount?, ...items): mutating splice - returns array of deleted elements
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_sp As Long: optBase_sp = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "tospliced"
                                ' toSpliced(start, deleteCount?, ...items) -> non-mutating version of splice
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                optBase_sp = progScope.GetValue("__option_base")
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
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                optBase_u2 = progScope.GetValue("__option_base")
                                lb_b = LBound(baseValLocal)
                                ub_b = UBound(baseValLocal)
                                len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
                                Dim idxRaw_sp As Long: idxRaw_sp = CLng(evaluated(1))
                                Dim idxN_sp As Long: idxN_sp = NormalizeIndex(idxRaw_sp, len_base, optBase_u2)
                                If idxN_sp >= len_base Then vAssignment tmpResult, Empty: GoTo exitfun
                                vAssignment tmpResult, baseValLocal(lb_b + idxN_sp)
                                GoTo exitfun
                            Case "copywithin"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_cw As Long: optBase_cw = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "entries"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim ob_entries As Long: ob_entries = progScope.GetValue("__option_base")
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
                            Case "every"
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim cb_every As Variant: vAssignment cb_every, evaluated(1)
                                If TypeName(cb_every) <> "ASF_Map" Or cb_every.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
                                lb_b = LBound(baseValLocal)
                                ub_b = UBound(baseValLocal)
                                len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
                                Dim ob_ev As Long: ob_ev = progScope.GetValue("__option_base")
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
                            Case "find", "findindex", "findlast", "findlastindex"
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim cb_find As Variant: vAssignment cb_find, evaluated(1)
                                If TypeName(cb_find) <> "ASF_Map" Or cb_find.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim ob_f As Long: ob_f = progScope.GetValue("__option_base")
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
                            Case "from"
                                If evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim val_from As Variant: vAssignment val_from, evaluated(1)
                                Dim opt_from As Long: opt_from = progScope.GetValue("__option_base")
                            
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
                            
                                ElseIf TypeName(val_from) = "String" Then
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
                                    If TypeName(comp) = "ASF_Map" Then
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
                            Case "includes", "indexof", "lastindexof"
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim needle As Variant: vAssignment needle, evaluated(1)
                                Dim optBase_in As Long: optBase_in = progScope.GetValue("__option_base")
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
                            Case "of"
                                Dim optBase_of As Long: optBase_of = progScope.GetValue("__option_base")
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
                            Case "reverse"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_rev As Long: optBase_rev = progScope.GetValue("__option_base")
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
                                WriteToContainer calleeExpr, progScope, baseValLocal
                                GoTo exitfun
                            Case "toreversed"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_tr As Long: optBase_tr = progScope.GetValue("__option_base")
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
                            Case "some"
                                If Not IsArray(baseValLocal) Or evaluated.Count = 0 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim cb_some As Variant: vAssignment cb_some, evaluated(1)
                                If TypeName(cb_some) <> "ASF_Map" Or cb_some.GetValue("type") <> "Closure" Then vAssignment tmpResult, Empty: GoTo exitfun
                                lb_b = LBound(baseValLocal)
                                ub_b = UBound(baseValLocal)
                                len_base = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
                                Dim ob_som As Long: ob_som = progScope.GetValue("__option_base")
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
                            Case "with"
                                If Not IsArray(baseValLocal) Or evaluated.Count < 2 Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_w As Long: optBase_w = progScope.GetValue("__option_base")
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
                            Case "sort", "tosorted"
                                If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                                Dim optBase_srt As Long: optBase_srt = progScope.GetValue("__option_base")
                                lb_b = LBound(baseValLocal)
                                ub_b = UBound(baseValLocal)
                                Dim n_s As Long: n_s = IIf(ub_b >= lb_b, ub_b - lb_b + 1, 0)
                            
                                ' get comparator if any
                                If evaluated.Count >= 1 Then vAssignment comp, evaluated(1) Else comp = Empty
                            
                                If n_s <= 1 Then
                                    If propN = "sort" Then
                                        ' ensure mutation writes back even if single element (keeps behavior consistent)
                                        WriteToContainer calleeExpr, progScope, baseValLocal
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
                                    If TypeName(comp) = "ASF_Map" Then
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
                                    WriteToContainer calleeExpr, progScope, baseValLocal
                                    vAssignment tmpResult, baseValLocal
                                Else
                                    vAssignment tmpResult, outArr
                                End If
                            
                                GoTo exitfun
                            Case Else
                                GoTo namedFunctions
                        End Select
                    ElseIf TypeName(baseValLocal) = "ASF_RegexEngine" Then
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
                    ElseIf TypeName(baseValLocal) = "String" Then
                        'Builtins string methods
                        Dim nidx As Long
                        Dim eidx As Long
                        Dim strOB As Long: strOB = 0
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
                            Case "endswith"
                                Dim endStr As String
                                Dim ewComStr As String
                                ewComStr = CStr(evaluated(1))
                                endStr = right$(baseValLocal, Len(ewComStr))
                                tmpResult = (endStr = ewComStr)
                                GoTo exitfun
                            Case "fromcharcode"
                                tmpResult = ChrW$(CLng(evaluated(1)))
                                GoTo exitfun
                            Case "includes"
                                Dim icStartPos As Long
                                If evaluated.Count > 1 Then
                                    icStartPos = CLng(evaluated(2)) + 1
                                Else
                                    icStartPos = 1
                                End If
                                tmpResult = (InStrB(icStartPos, baseValLocal, evaluated(1), vbBinaryCompare) > 0)
                                GoTo exitfun
                            Case "indexof"
                                Dim ioStartPos As Long
                                If evaluated.Count > 1 Then
                                    ioStartPos = CLng(evaluated(2)) + 1
                                Else
                                    ioStartPos = 1
                                End If
                                tmpResult = InStr(ioStartPos, baseValLocal, evaluated(1), vbBinaryCompare) - 1
                                GoTo exitfun
                            Case "lastindexof"
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
                            Case "localecompare"
                                ' A tiny VBA based implementation, does not behave like javascript
                                tmpResult = StrComp(baseValLocal, evaluated(1), vbBinaryCompare)
                                GoTo exitfun
                            Case "match", "matchall"
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
                                        Do While sPosA <= Lsrc
                                            Dim capsA As Collection
                                            Dim tryStartA As Long
                                            Dim foundAtA As Long: foundAtA = 0
                                            For tryStartA = sPosA To Lsrc
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
                                                sPosA = IIf(foundAtA + 1 > Lsrc, Lsrc + 1, foundAtA + 1)
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
                                    
                                        Do While sPos <= L2
                                            Dim caps2 As Collection
                                            Dim sTry As Long
                                            Dim foundAt2 As Long: foundAt2 = 0
                                            For sTry = sPos To L2
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
                                                sPos = IIf(foundAt2 + 1 > L2, L2 + 1, foundAt2 + 1)
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
                            Case "padend"
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
                            Case "padstart"
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
                            Case "repeat"
                                Dim repCount As Long
                                Dim repStr As String
                                Dim rpI As Long
                                repCount = CLng(evaluated(1))
                                For rpI = 1 To repCount
                                    repStr = repStr & baseValLocal
                                Next rpI
                                tmpResult = repStr
                                GoTo exitfun
                            Case "replace", "replaceall"
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
    
                                    Dim doGlobal As Boolean: doGlobal = flag_g Or isGlobal
    
                                    With GLOBALS_.re
                                        ' Init with flags (ignoreCase, maxSteps, multiline, dotAll)
                                        .Init pattern, flag_i, .MaxMatchSteps, flag_m, flag_s
    
                                        Dim outRepStr As String: outRepStr = vbNullString
                                        Dim L As Long: L = Len(src)
                                        Dim searchPos As Long: searchPos = 1
    
                                        Do While searchPos <= L
                                            Dim caps As Collection
                                            Dim startTry As Long
                                            Dim foundAt As Long: foundAt = 0
    
                                            ' locate the next match by trying ExecAt at increasing start positions
                                            For startTry = searchPos To L
                                                Set caps = .ExecAt(src, startTry)
                                                If Not (caps Is Nothing) Then
                                                    foundAt = startTry
                                                    Exit For
                                                End If
                                            Next startTry
    
                                            If foundAt = 0 Then
                                                ' append remainder and break
                                                outRepStr = outRepStr & mid$(src, searchPos)
                                                Exit Do
                                            Else
                                                Dim fullMatch As String: fullMatch = CStr(caps.Item(1))
                                                Dim matchLen As Long: matchLen = Len(fullMatch)
                                                Dim matchStart As Long: matchStart = foundAt
                                                Dim matchEnd As Long: matchEnd = matchStart + matchLen - 1
    
                                                ' Safety: avoid infinite loop for zero-length matches
                                                If matchLen = 0 Then
                                                    ' append the next char (if any) and advance
                                                    If searchPos <= L Then outRepStr = outRepStr & mid$(src, searchPos, 1)
                                                    searchPos = searchPos + 1
                                                    If searchPos > L Then Exit Do
                                                    If Not doGlobal Then
                                                        ' single replacement only -> append remainder and exit
                                                        outRepStr = outRepStr & mid$(src, searchPos)
                                                        Exit Do
                                                    Else
                                                        ' continue global
                                                        GoTo ContinueDo
                                                    End If
                                                End If
    
                                                ' append chunk before match
                                                If matchStart > searchPos Then outRepStr = outRepStr & mid$(src, searchPos, matchStart - searchPos)
    
                                                ' compute replacement text
                                                Dim replacedText As String: replacedText = vbNullString
    
                                                ' --- replacement is a function/closure? ---
                                                If Not IsEmpty(replRaw) And (IsObject(replRaw) Or TypeName(replRaw) = "ASF_Map") Then
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
    
                                                ' advance search position according to global / single replacement
                                                searchPos = matchEnd + 1
                                                If Not doGlobal Then
                                                    ' append remainder and exit
                                                    If searchPos <= L Then outRepStr = outRepStr & mid$(src, searchPos)
                                                    Exit Do
                                                End If
                                            End If
ContinueDo:
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
                            Case "startswith"
                                Dim startStr As String
                                Dim swComStr As String
                                swComStr = CStr(evaluated(1))
                                startStr = mid$(baseValLocal, 1, Len(swComStr))
                                tmpResult = (startStr = swComStr)
                                GoTo exitfun
                            Case "substring"
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
                    ElseIf TypeName(baseValLocal) = "ASF_Map" Then
                        ' ASF_Map (object) methods
                        Select Case propN
                            Case "keys"
                                ' Return array of object keys
                                Dim keysCol As Collection: Set keysCol = baseValLocal.keys()
                                Dim keysArr() As Variant
                                Dim optBase_keys As Long: optBase_keys = progScope.GetValue("__option_base")
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
                                
                            Case "values"
                                ' Return array of object values
                                Dim itemsCol As Collection: Set itemsCol = baseValLocal.items()
                                Dim valsArr() As Variant
                                Dim optBase_vals As Long: optBase_vals = progScope.GetValue("__option_base")
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
                                
                            Case "entries"
                                ' Return array of [key, value] pairs
                                Dim entriesKeys As Collection: Set entriesKeys = baseValLocal.keys()
                                Dim entriesItems As Collection: Set entriesItems = baseValLocal.items()
                                Dim entriesArr() As Variant
                                Dim optBase_ent As Long: optBase_ent = progScope.GetValue("__option_base")
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
                                
                            Case "set"
                                ' Set key-value pair and return the object (for chaining)
                                If evaluated.Count >= 2 Then
                                    baseValLocal.SetValue CStr(evaluated(1)), evaluated(2)
                                End If
                                vAssignment tmpResult, baseValLocal
                                GoTo exitfun
                                
                            Case "delete", "remove"
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
                                ' Merge another object into this one
                                If evaluated.Count = 0 Then
                                    vAssignment tmpResult, baseValLocal
                                Else
                                    Dim mergeObj As Variant: vAssignment mergeObj, evaluated(1)
                                    If TypeName(mergeObj) = "ASF_Map" Then
                                        Dim mergeKeys As Collection: Set mergeKeys = mergeObj.keys()
                                        Dim mk As Variant
                                        For Each mk In mergeKeys
                                            baseValLocal.SetValue CStr(mk), mergeObj.GetValue(CStr(mk))
                                        Next mk
                                    End If
                                    vAssignment tmpResult, baseValLocal
                                End If
                                GoTo exitfun
                                
                            Case "foreach"
                                ' Iterate over entries with callback(value, key, object)
                                If evaluated.Count = 0 Then
                                    vAssignment tmpResult, Empty
                                    GoTo exitfun
                                End If
                                
                                Dim cb_obj As Variant: vAssignment cb_obj, evaluated(1)
                                If TypeName(cb_obj) <> "ASF_Map" Or cb_obj.GetValue("type") <> "Closure" Then
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
                                
                            Case "map"
                                ' Map over entries, return new object with transformed values
                                If evaluated.Count = 0 Then
                                    vAssignment tmpResult, Empty
                                    GoTo exitfun
                                End If
                                
                                Dim cb_map As Variant: vAssignment cb_map, evaluated(1)
                                If TypeName(cb_map) <> "ASF_Map" Or cb_map.GetValue("type") <> "Closure" Then
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
                                
                            Case "filter"
                                ' Filter entries, return new object with entries that pass test
                                If evaluated.Count = 0 Then
                                    vAssignment tmpResult, Empty
                                    GoTo exitfun
                                End If
                                
                                Dim cb_filter As Variant: vAssignment cb_filter, evaluated(1)
                                If TypeName(cb_filter) <> "ASF_Map" Or cb_filter.GetValue("type") <> "Closure" Then
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
                                
                            Case "some"
                                ' Test if at least one entry passes test
                                If evaluated.Count = 0 Then
                                    tmpResult = False
                                    GoTo exitfun
                                End If
                                
                                Dim cb_some_obj As Variant: vAssignment cb_some_obj, evaluated(1)
                                If TypeName(cb_some_obj) <> "ASF_Map" Or cb_some_obj.GetValue("type") <> "Closure" Then
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
                                
                            Case "every"
                                ' Test if all entries pass test
                                If evaluated.Count = 0 Then
                                    tmpResult = True
                                    GoTo exitfun
                                End If
                                
                                Dim cb_every_obj As Variant: vAssignment cb_every_obj, evaluated(1)
                                If TypeName(cb_every_obj) <> "ASF_Map" Or cb_every_obj.GetValue("type") <> "Closure" Then
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
                                
                            Case Else
                                GoTo namedFunctions
                        End Select
                    ElseIf TypeName(baseValLocal) = "Collection" Then
CallCollMethod:
                        If isSpecialColl Then
                            ' Check for Math object
                            Dim baseLocal As Variant
                            vAssignment baseLocal, tmpBaseLocal
                            If baseLocal.GetValue("name") = "Math" Then
                                Select Case propN
                                    Case "E"
                                        tmpResult = Exp(1)
                                        GoTo exitfun
                                    Case "PI"
                                        tmpResult = 4 * Atn(1)
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
                                        Dim maxEl As Variant
                                        tmpResult = 0
                                        For Each maxEl In evaluated
                                            If maxEl > tmpResult Then
                                                tmpResult = maxEl
                                            End If
                                        Next maxEl
                                        GoTo exitfun
                                    Case "min"
                                        Dim minEl As Variant
                                        tmpResult = 0
                                        For Each minEl In evaluated
                                            If minEl < tmpResult Then
                                                tmpResult = minEl
                                            End If
                                        Next minEl
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
                    Else
                        vAssignment tmpResult, CallObjectMethod(baseValLocal, propN, evaluated)
                        GoTo exitfun
                    End If
                    ' --- end method builtins ---
                End If
            End If
namedFunctions:
            ' If callee is a Variable AST -> treat as named call (builtins like range, IsArray, IsNumeric, flatten, clone)
            If fname = vbNullString Then
                If Not calleeExpr Is Nothing Then
                    If calleeExpr.Exists("type") And CStr(calleeExpr.GetValue("type")) = "Variable" Then
                        fname = CStr(calleeExpr.GetValue("name"))
                    End If
                Else
                    ' check "name" property instead of a callee AST
                    If node.Exists("name") Then fname = CStr(node.GetValue("name")) Else fname = vbNullString
                End If
            End If
            If fname <> vbNullString Then
                ' --- builtin named functions ---
                Dim nameLcase As String: nameLcase = LCase$(fname)
                If nameLcase = "range" Then
                    ' range([start,] end[, step]) -> returns an array (python-style range semantics)
                    Dim rcount As Long: rcount = evaluated.Count
                    Dim rs As Double, re As Double, rstep As Double
                    Dim optionBase_local As Long: optionBase_local = progScope.GetValue("__option_base")
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
                    vAssignment tmpMapObj, node.GetValue("args")(1)
                    If evaluated.Count = 1 Then
                        vAssignment tmpResult, builtin_flatten(EvalExprNode(tmpMapObj, progScope), progScope)
                    Else
                        vAssignment tmpResult, builtin_flatten(EvalExprNode(tmpMapObj, progScope), progScope, evaluated(2))
                    End If
                    GoTo exitfun
                End If
                If nameLcase = "clone" Then
                    If evaluated.Count = 0 Then
                        vAssignment tmpResult, Empty
                        GoTo exitfun
                    End If
                    vAssignment tmpMapObj, node.GetValue("args")(1)
                    vAssignment tmpResult, builtin_clone(EvalExprNode(tmpMapObj, progScope), progScope)
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
                
                    If TypeName(cbFn) <> "ASF_Map" Or cbFn.GetValue("type") <> "Closure" Then
                        vAssignment tmpResult, Empty
                        GoTo exitfun
                    End If
                
                    If Not IsArray(feTmpVal) And TypeName(feTmpVal) <> "ASF_Map" Then
                        vAssignment tmpResult, Empty
                        GoTo exitfun
                    End If
                
                    If TypeName(feTmpVal) = "ASF_Map" Then
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
                        Dim optBase2 As Long: optBase2 = progScope.GetValue("__option_base")
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
                ' --- end builtins ---
            End If
            If Not hasCalleeExpr Then
                ' named call
                If GLOBALS_.gFuncTable.Exists(fname) Then
                    funcIdx = CLng(GLOBALS_.gFuncTable.GetValue(fname))
                    vAssignment tmpResult, CallFuncByIndex_AST(funcIdx, evaluated)
                End If
            Else
                ' dynamic callee: evaluate callee expression -> it must yield a closure map or call via VB
                Dim calleeVal As Variant
                Dim hasThisVal As Boolean: hasThisVal = False
                ' If callee is a member or index expression, compute base as `this` for method call binding.
                If Not calleeExpr Is Nothing Then
                    Dim ct As String: ct = IIf(calleeExpr.Exists("type"), CStr(calleeExpr.GetValue("type")), vbNullString)
                    If ct = "Member" Or ct = "Index" Then
                        hasThisVal = True
                        vAssignment thisVal, baseValLocal
                    End If
                End If
                vAssignment calleeVal, EvalExprNode(calleeExpr, progScope)
                If TypeName(calleeVal) = "ASF_Map" Then
                    Dim cM As ASF_Map: Set cM = calleeVal
                    If calleeVal.Exists("type") Then
                        Dim calleeType As String: calleeType = calleeVal.GetValue("type")
                        If calleeType = "Class" Then
                            ' This is a class instantiation with 'new'
                            vAssignment tmpResult, InstantiateClass(cM, evaluated, progScope)
                            GoTo exitfun
                        ElseIf calleeType = "BoundMethod" Then
                            ' Execute bound method
                            Dim methodNode2 As ASF_Map: Set methodNode2 = calleeVal.GetValue("method")
                            Dim boundInstance As ASF_Map: Set boundInstance = calleeVal.GetValue("instance")
                            
                            vAssignment tmpResult, ExecuteBoundMethod(methodNode2, boundInstance, evaluated, progScope)
                            GoTo exitfun
                        ' Execute static method
                        ElseIf calleeType = "Method" Then
                            vAssignment tmpMapObj, calleeVal
                            vAssignment tmpResult, ExecuteClassMethod(tmpMapObj, evaluated, progScope)
                            GoTo exitfun
                        End If
                    End If
                    If cM.GetValue("type") = "Closure" Then
                        If hasThisVal Then
                            vAssignment tmpResult, CallClosure(cM, evaluated, thisVal)
                        Else
                            vAssignment tmpResult, CallClosure(cM, evaluated)
                        End If
                    Else
                        ' not a closure -> fail gracefully
                        vAssignment tmpResult, Empty
                    End If
                Else
                    vAssignment tmpResult, Empty
                End If
            End If
            GoTo exitfun

        Case "Unary"
            Dim op As String: op = node.GetValue("op")
            Dim ev As Variant: vAssignment ev, EvalExprNode(node.GetValue("expr"), progScope)
            If op = "!" Then
                vAssignment tmpResult, Not IsTruthy(ev)
            ElseIf op = "-" Then
                vAssignment tmpResult, -CDbl(ev)
            ElseIf op = "typeof" Then
                vAssignment tmpResult, builtin_typeof(ev)
            Else
                vAssignment tmpResult, ev
            End If
            GoTo exitfun
        Case "Ternary"
            Dim cnd As Variant: cnd = EvalExprNode(node.GetValue("cond"), progScope)
            If IsTruthy(cnd) Then
                vAssignment tmpResult, EvalExprNode(node.GetValue("trueExpr"), progScope)
            Else
                vAssignment tmpResult, EvalExprNode(node.GetValue("falseExpr"), progScope)
            End If
            GoTo exitfun
        Case "Binary"
            Dim lop As Variant: lop = EvalExprNode(node.GetValue("left"), progScope)
            Dim rop As Variant
            Dim op2 As String: op2 = node.GetValue("op")
            If op2 = "&&" Then
                If Not IsTruthy(lop) Then vAssignment tmpResult, False: GoTo exitfun
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) And IsTruthy(rop))
                GoTo exitfun
            ElseIf op2 = "||" Then
                If IsTruthy(lop) Then vAssignment tmpResult, True: GoTo exitfun
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                vAssignment tmpResult, (IsTruthy(lop) Or IsTruthy(rop))
                GoTo exitfun
            Else
                vAssignment rop, EvalExprNode(node.GetValue("right"), progScope)
                Select Case op2
                    Case "+"
                        vAssignment tmpResult, SafeAdd(lop, rop)
                    Case "-": vAssignment tmpResult, lop - rop
                    Case "*": vAssignment tmpResult, lop * rop
                    Case "/"
                        If rop = 0 Then err.Raise vbObjectError + 2001, "VM.EvalExprNode", "Division by zero" Else vAssignment tmpResult, lop / rop
                    Case "%": vAssignment tmpResult, lop Mod rop
                    Case "^": vAssignment tmpResult, lop ^ rop
                    Case "&": vAssignment tmpResult, CStr(lop) & CStr(rop)
                    Case "|": vAssignment tmpResult, lop Or rop
                    Case "==", "=": vAssignment tmpResult, (IIf(IsNull(lop), Empty, lop) = IIf(IsNull(rop), Empty, rop))
                    Case "!=": vAssignment tmpResult, (lop <> rop)
                    Case "<": vAssignment tmpResult, (lop < rop)
                    Case ">": vAssignment tmpResult, (lop > rop)
                    Case "<=": vAssignment tmpResult, (lop <= rop)
                    Case ">=": vAssignment tmpResult, (lop >= rop)
                    Case "<<": vAssignment tmpResult, lshift(lop, rop)
                    Case ">>": vAssignment tmpResult, rshift(lop, rop)
                    Case Else
                        vAssignment tmpResult, Empty
                End Select
                GoTo exitfun
            End If
    End Select

    vAssignment tmpResult, Empty
exitfun:
    vAssignment EvalExprNode, tmpResult
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
Private Function ReturnCollection(ByRef aColl As Variant) As Collection
    Set ReturnCollection = aColl
End Function
Private Function ArrayIsInit(aArray As Variant) As Boolean
    Dim ub As Long
    ArrayIsInit = True
    On Error GoTo err_Handler
    ub = UBound(aArray)
    Exit Function
err_Handler:
    err.Clear
    ArrayIsInit = False
End Function
' Call function program by index (AST). "args" is Collection of evaluated arg values.
Private Function CallFuncByIndex_AST(funcIdx As Long, args As Collection) As Variant
    Dim p As Variant: p = GLOBALS_.gPrograms(funcIdx)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim callScope As New ASF_ScopeStack
    Dim i As Long, k As Long
    callScope.LoadRaw rawScope
    callScope.Push
    Dim fname As String: fname = p(0)
    
    ' Trace: Push call
    PushCallTrace fname, args
    
    Dim pa As Variant
    ' set params from gFuncParams
    If GLOBALS_.gFuncParams.Exists(fname) Then
        pa = GLOBALS_.gFuncParams.GetValue(fname)
        For i = LBound(pa) To UBound(pa)
            If i - LBound(pa) + 1 <= args.Count Then
                callScope.SetLocalValue CStr(pa(i)), args(i - LBound(pa) + 1)
            Else
                callScope.SetLocalValue CStr(pa(i)), Empty
            End If
        Next i
    End If
    Dim stmts As Collection: Set stmts = p(1)
    Dim retVal As Variant
    retVal = Empty
    
    On Error GoTo FuncError
    
    For i = 1 To stmts.Count
        Dim ctrl As String: ctrl = ExecuteStmtNode(stmts(i), callScope)
        If ctrl = "RETURN" Then
            vAssignment retVal, callScope.GetValue("__return")
            ' Trace: Pop call with return value
            PopCallTrace fname, retVal
            vAssignment CallFuncByIndex_AST, retVal
            callScope.Pop
            Exit Function
        ElseIf ctrl = "ERR" Then
            TraceError fname, "Error during function execution"
            err.Raise vbObjectError + 3000, "VM.CallFunc", "Error during function execution"
        End If
    Next i
    
    ' Trace: Pop call with empty return
    PopCallTrace fname, Empty
    vAssignment CallFuncByIndex_AST, Empty
    callScope.Pop
    Exit Function
    
FuncError:
    TraceError fname, err.Description
    callScope.Pop
    err.Raise err.Number, err.Source, err.Description
End Function

' ---------------------
' Utilities adapted from previous VM
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

Private Function SafeAdd(a As Variant, b As Variant) As Variant
    If TypeName(a) <> "String" And TypeName(b) <> "String" Then
        SafeAdd = a + b
    Else
        SafeAdd = IIf(IsNull(a), a, CStr(a)) & IIf(IsNull(b), b, CStr(b))
    End If
End Function

Private Function lshift(a As Variant, b As Variant) As Variant 'a<<b
    If IsNumeric(a) And IsNumeric(b) Then
        lshift = Floor(Floor(a) * (2 ^ Abs(Floor(b))))
    Else
        Exit Function
    End If
End Function

Private Function rshift(a As Variant, b As Variant) As Variant 'a>>b
    If IsNumeric(a) And IsNumeric(b) Then
        rshift = Floor(Floor(a) / (2 ^ Abs(Floor(b))))
    Else
        Exit Function
    End If
End Function

Private Function Floor(ByRef value As Variant) As Long
    Dim tmpResult As Double
    tmpResult = CDbl(Fix(value))
    Floor = CLng(tmpResult + ((CDbl(value) <> tmpResult) And (CDbl(value) < 0)))
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
            ValueToStringWithCtx = "<" & TypeName(v) & ">"
        End If
        Exit Function
    End If

    ' Objects
    If IsObject(v) Then
        Dim tn As String: tn = TypeName(v)
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
    Dim tn As String: tn = TypeName(obj)

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

' Utility to escape short strings for printing if you want (optional)
Private Function EscapeStringForPrint(s As String) As String
    ' currently returns s raw; adapt if you want quoted output
    EscapeStringForPrint = s
End Function
 ' VBAexpressions integration helpers
 ' ---------------------
 ' Evaluate a raw VBAexpressions expression string using a VBAexpressions instance,
 ' seeding it with the current ASF scope variables so VB expressions can reference ASF variables.
Private Function EvalVBExpressionWithScope(expr As String, progScope As ASF_ScopeStack) As Variant
     On Error GoTo ErrHandler
     With GLOBALS_.gExprEvaluator
         ' Create expression in evaluator
         .Create expr
    
         ' Inject variables from progScope (shadowing: global frames first, then locals)
         Dim frame As Variant
         Dim m As ASF_Map
         Dim keyCol As Collection
         Dim key As Variant
         For Each frame In progScope.Raw
             Set m = frame
             Dim tmpVal As Variant
             Set keyCol = m.keys
             For Each key In keyCol
                 vAssignment tmpVal, m.GetValue(CStr(key))
                 If IsArray(tmpVal) Then
                    .VarValue(CStr(key)) = .ArrayToString(tmpVal)
                 Else
                    If Not IsObject(tmpVal) Then
                        .VarValue(CStr(key)) = CStr(tmpVal)
                    End If
                End If
             Next key
         Next frame
    
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
 ' Closure / function-value support (runtime helpers)
 ' ---------------------

 ' Create and call closures (closureMap is a Map with keys:
 '  "type" = "Closure", "params" = Collection, "body" = Collection (stmts), "env" = ScopeStack)
Private Function CallClosure(closureMap As ASF_Map, evaluatedArgs As Collection, Optional thisVal As Variant) As Variant
    On Error GoTo ErrHandler
    
    ' Get function name for tracing (if available)
    Dim traceName As String
    If closureMap.Exists("name") Then
        traceName = CStr(closureMap.GetValue("name"))
    Else
        traceName = "<anonymous>"
    End If
    
    ' Trace: Push call
    PushCallTrace traceName, evaluatedArgs
    
    Dim env As ASF_ScopeStack
    Set env = closureMap.GetValue("env") ' shared reference (shared-write semantics)

    ' push a new frame for this call
    env.Push
    Dim i As Long
    Dim pa As Variant
    Dim elm As Variant
    Dim fc As Boolean
    Dim emptyArgs As Long: emptyArgs = 0
    
    vAssignment pa, closureMap.GetValue("params")
    If IsObject(pa) Then
       fc = (Not pa Is Nothing)
    Else
       fc = Not IsEmpty(pa)
    End If
    If fc Then
        For Each elm In pa
            i = i + 1
            If CStr(elm) = vbNullString Then emptyArgs = emptyArgs + 1
            If i <= evaluatedArgs.Count Then
                env.SetLocalValue CStr(elm), evaluatedArgs(i)
            Else
                env.SetLocalValue CStr(elm), Empty
            End If
        Next elm
    End If
    
    ' Handle rest parameters
    If closureMap.Exists("restParam") Then
        Dim restParamName As String: restParamName = closureMap.GetValue("restParam")
        Dim restArr() As Variant
        Dim regularParamCount As Long
        
        ' Count regular parameters
        If i > 1 Then
            regularParamCount = i
        Else
            regularParamCount = i - emptyArgs
        End If
        
        Dim totalArgCount As Long: totalArgCount = evaluatedArgs.Count
        Dim optBase As Long: optBase = env.GetValue("__option_base")
        
        If totalArgCount > regularParamCount Then
            ' Check for empty params
            If totalArgCount - regularParamCount = 1 Then
                If IsEmpty(evaluatedArgs(regularParamCount + 1)) Then
                    ' Return the empty array
                    ReDim restArr(-1 To -1)
                Else
                    ReDim restArr(optBase To optBase)
                    vAssignment restArr(optBase), evaluatedArgs(regularParamCount + 1)
                End If
            Else
                ' Collect remaining arguments into rest array
                Dim restCount As Long: restCount = totalArgCount - regularParamCount
                ReDim restArr(optBase To optBase + restCount - 1)
                
                Dim ri As Long
                For ri = 0 To restCount - 1
                    vAssignment restArr(optBase + ri), evaluatedArgs(regularParamCount + 1 + ri)
                Next ri
            End If
        Else
            ' No extra arguments, create empty array
            ReDim restArr(-1 To -1)
        End If
        
        ' Bind rest parameter to the rest array
        env.SetLocalValue restParamName, restArr
    End If
    
    ' set 'this' if provided
    If Not IsMissing(thisVal) Then
        env.SetLocalValue "this", thisVal
    End If

    ' execute body
    Dim stmts As Collection: Set stmts = closureMap.GetValue("body")
    Dim ctrl As String
    Dim stmtIdx As Long
    Dim retVal As Variant
    retVal = Empty
    
    For stmtIdx = 1 To stmts.Count
        ctrl = ExecuteStmtNode(stmts(stmtIdx), env)
        If ctrl = "RETURN" Then
            vAssignment retVal, env.GetValue("__return")
            ' Trace: Pop call with return value
            PopCallTrace traceName, retVal
            vAssignment CallClosure, retVal
            env.Pop
            Exit Function
        ElseIf ctrl = "ERR" Then
            TraceError traceName, "Error during closure execution"
            err.Raise vbObjectError + 8001, "VM.CallClosure", "Error during closure execution"
        End If
    Next stmtIdx

    ' normal return -> Empty
    ' Trace: Pop call with empty return
    PopCallTrace traceName, Empty
    env.Pop
    CallClosure = Empty
    Exit Function
    
ErrHandler:
    TraceError traceName, err.Description
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "CallClosure error: " & err.Description
    err.Clear
    PopCallTrace traceName, Empty
    CallClosure = Empty
End Function
 ' ----------------- LValue resolution and helpers -----------------

Private Sub HandleAssignment(left As ASF_Map, ByVal rval As Variant, progScope As ASF_ScopeStack)
    Dim lval As ASF_Map: Set lval = ResolveLValue(left, progScope)
    Dim key As Variant: vAssignment key, lval("key")
    If lval.Exists("computed") Then vAssignment key, EvalExprNode(left.GetValue("index"), progScope) ' Lazy compute
    Dim container As Variant: vAssignment container, lval("container")
    Select Case lval("kind")
        Case "scopeVar"
            If mid$(key, 1, 1) <> PLACEHOLDER_CHR_ Then
                progScope.SetValue CStr(key), rval
            End If
        Case "prop"
            If TypeName(container) <> "ASF_Map" Then
                Dim tValue As Variant
                vAssignment tValue, ResolveIndexProp(lval)
                Select Case TypeName(tValue)
                    Case "ASF_Map", "ASF_ScopeStack"
                        tValue.SetValue CStr(key), rval
                    Case Else
                        ' Perform a call by name
                        SetObjectProperty tValue, key, rval
                End Select
            Else
                container.SetValue CStr(key), rval
            End If
        Case "index"
            Dim pos As Long: pos = CLng(key)
            Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
            If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.HandleAssignment", "Invalid index"
        
            ' If ResolvePath provided parentKey this means the array is stored under a property/variable:
            Dim parentKey As Variant
            Dim keyChain As New Collection
            Dim baseKey As Variant: baseKey = lval("key")
            If lval.Exists("parentKey") Then vAssignment parentKey, lval("parentKey") Else parentKey = Empty
            Do While lval.Exists("parentKind") And lval("parentKind") = "index"
                keyChain.Add parentKey
                vAssignment lval, lval("parentObj")
                vAssignment parentKey, lval("parentKey")
            Loop
            keyChain.Add baseKey
            If Not IsEmpty(parentKey) Then
                ' parent container holds the array under parentKey (could be progScope or Map)
                Dim arrVal As Variant
                If TypeName(container) = "ASF_Map" Then
                    vAssignment arrVal, container.GetValue(CStr(parentKey))
                ElseIf TypeName(container) = "ASF_ScopeStack" Then
                    vAssignment arrVal, container.GetValue(CStr(parentKey))
                Else
                    ' fallback: maybe container itself is the array
                    vAssignment arrVal, container
                End If
                ' Initialize the array if needed
                If UBound(arrVal) = -1 Then ReDim Preserve arrVal(1 To 1)
                ' Ensure arrVal is an array we can index into (AssignToArrayValueInPlace will create/resize).
                If keyChain.Count > 1 Then 'nested indexes
                    Call AssignToArrayValueInPlace(arrVal, pos, rval, optionBase, keyChain)
                Else
                    Call AssignToArrayValueInPlace(arrVal, pos, rval)
                End If
                ' Write the updated array back into the parent container
                If TypeName(container) = "ASF_Map" Then
                    container.SetValue CStr(parentKey), arrVal
                ElseIf TypeName(container) = "ASF_ScopeStack" Then
                    container.SetValue CStr(parentKey), arrVal
                Else
                    ' if container itself was an array variable, try to overwrite it:
                    container = arrVal
                End If
            Else
                ' container is the array itself
                If IsArray(container) Then
                    Call AssignToArrayValueInPlace(container, pos, rval)
                Else
                    err.Raise vbObjectError + 5002, "VM.HandleAssignment", "Cannot index non-array/non-map for set"
                End If
            End If
        Case "Member"
            Dim baseVal As Variant: vAssignment baseVal, EvalExprNode(left.GetValue("base"), progScope)
            If TypeName(baseVal) <> "ASF_Map" Then err.Raise vbObjectError + 5007, "VM.HandleAssignment", "Cannot set property on non-map"
            Dim propKey As String: propKey = IIf(left.Exists("prop"), "prop", "property")
            baseVal.SetValue left.GetValue(propKey), rval
    End Select
End Sub
' UpdateNestedArrayByPath
' Mutates a jagged Variant array "root" at the nested path specified by idxPath (Array or Collection of indexes).
' root is passed ByRef and will be updated (the top-level will be overwritten if needed).
' idxPath: either an Array(...) or a Collection of numeric indexes (user-space, using optBase)
' newVal: value to write at target slot
' optBase: 0 or 1 (default 1)
' createMissing: if True, create missing inner arrays as needed (default True)
' Returns True on success, False on failure.
Public Function UpdateNestedArrayByPath(ByRef root As Variant, ByVal idxPath As Variant, ByVal newVal As Variant, _
                                       Optional ByVal optBase As Long = 1, Optional ByVal createMissing As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim idxs() As Long, depth As Long
    depth = 0

    ' --- Normalize idxPath to 1..depth long array idxs() ---
    If IsArray(idxPath) Then
        Dim lbp As Long, ubp As Long
        lbp = LBound(idxPath): ubp = UBound(idxPath)
        depth = ubp - lbp + 1
        If depth <= 0 Then UpdateNestedArrayByPath = False: Exit Function
        ReDim idxs(1 To depth)
        Dim ii As Long, posi As Long: posi = 1
        For ii = lbp To ubp
            If LBound(root) <> optBase Then
                idxs(posi) = NormalizeIndex(CLng(idxPath(ii)), UBound(root) - LBound(root) + 1, optBase)
            Else
                idxs(posi) = CLng(idxPath(ii))
            End If
            posi = posi + 1
        Next ii
    Else
        ' single scalar index
        depth = 1
        ReDim idxs(1 To 1)
        idxs(1) = CLng(idxPath)
    End If

    ' Ensure root is an array (create if allowed)
    If Not IsArray(root) Then
        If createMissing Then
            ReDim root(optBase To optBase) ' single-slot per ASF convention
        Else
            UpdateNestedArrayByPath = False: Exit Function
        End If
    End If

    ' containers(k) will hold the array value at level k (1-based levels)
    ' posAbs(k) will hold the absolute position (array index) at that level (LBound..UBound inclusive)
    Dim containers() As Variant, posAbs() As Long
    ReDim containers(1 To depth)
    ReDim posAbs(1 To depth)

    ' initialize level 1 container to root
    vAssignment containers(1), root

    Dim level As Long
    For level = 1 To depth
        Dim arr As Variant
        vAssignment arr, containers(level)           ' may be a copy - we'll write back after changes
        Dim arrLB As Long, arrUB As Long, arrLen As Long
        If IsArray(arr) Then
            arrLB = LBound(arr): arrUB = UBound(arr)
            If arrUB >= arrLB Then arrLen = arrUB - arrLB + 1 Else arrLen = 0
        Else
            arrLB = 0: arrUB = -1: arrLen = 0
        End If

        ' compute normalized relative index (0-based offset) from user-provided index (supports negatives)
        Dim userIdx As Long: userIdx = idxs(level)
        Dim relOff As Long
        relOff = NormalizeIndexOffset(userIdx, arrLen, optBase)

        ' absolute position inside arr expected
        Dim absPos As Long
        absPos = arrLB + relOff

        ' If array is empty or position outside bounds, resize to include absPos
        If Not IsArray(arr) Or absPos < arrLB Or absPos > arrUB Then
            If Not createMissing Then
                UpdateNestedArrayByPath = False: Exit Function
            End If
            ' compute new bounds
            Dim newLB As Long, newUB As Long
            If Not IsArray(arr) Then
                newLB = absPos: newUB = absPos
            Else
                newLB = arrLB: newUB = arrUB
                If absPos < newLB Then newLB = absPos
                If absPos > newUB Then newUB = absPos
            End If
            Dim tmp() As Variant
            ReDim tmp(newLB To newUB)
            ' copy existing values if any
            If IsArray(arr) Then
                Dim k As Long
                For k = arrLB To arrUB
                    vAssignment tmp(k), arr(k)
                Next k
            End If
            vAssignment arr, tmp
            arrLB = newLB: arrUB = newUB: arrLen = arrUB - arrLB + 1
            ' store back into containers(level) - we'll propagate later
            vAssignment containers(level), arr
            ' propagate upward so parents hold updated child arrays
            Call PropagateUp(containers, posAbs, level, root)
        End If

        ' store absolute position and ensure containers(level) holds updated arr
        posAbs(level) = absPos
        vAssignment containers(level), arr

        If level < depth Then
            ' need to step into child at arr(absPos); ensure it's an array (create if necessary)
            Dim child As Variant
            vAssignment child, arr(absPos)

            If Not IsArray(child) Then
                If Not createMissing Then
                    UpdateNestedArrayByPath = False: Exit Function
                End If
                ' create single-slot inner array
                Dim newInner() As Variant
                ReDim newInner(optBase To optBase)
                vAssignment child, newInner
                ' assign into arr and propagate
                vAssignment arr(absPos), child
                vAssignment containers(level), arr
                Call PropagateUp(containers, posAbs, level, root)
            End If

            ' set next level container to child (fresh copy)
            vAssignment containers(level + 1), arr(absPos)
            ' continue loop
        Else
            ' final level: assign newVal into arr(absPos)
            vAssignment arr(absPos), newVal
            vAssignment containers(level), arr
            ' propagate upward to root
            Call PropagateUp(containers, posAbs, level, root)
            UpdateNestedArrayByPath = True
            Exit Function
        End If
    Next level

    UpdateNestedArrayByPath = False
    Exit Function

ErrHandler:
    UpdateNestedArrayByPath = False
End Function

' Helper: compute zero-based relative offset for user index, supports negatives
' userIdx: may be negative to access from end, or absolute user base (optBase)
' arrLen: length of array (may be 0)
' optBase: user base (0 or 1)
Private Function NormalizeIndexOffset(ByVal userIdx As Long, ByVal arrLen As Long, ByVal optBase As Long) As Long
    Dim idx As Long: idx = userIdx
    If idx < 0 Then
        ' negative: -1 => last element
        If arrLen > 0 Then
            idx = arrLen + idx + optBase
        Else
            ' no elements; treat negative as optBase (will become index 0 offset)
            idx = optBase
        End If
    End If
    ' relative zero-based offset
    NormalizeIndexOffset = idx - optBase
    If NormalizeIndexOffset < 0 Then NormalizeIndexOffset = 0
End Function

' PropagateUp(containers, posAbs, changedLevel, root)
' - containers(1..N) holds arrays at each level (may be copies)
' - posAbs(1..N) holds absolute positions chosen at each level (valid for levels <= changedLevel)
' - changedLevel: the highest level whose containers(changedLevel) has just been modified
' This writes containers(changedLevel) into containers(changedLevel-1)(posAbs(changedLevel-1)),
' then that container into its parent, etc, and finally updates the top-level "root" variable.
Private Sub PropagateUp(ByRef containers As Variant, ByRef posAbs As Variant, ByVal changedLevel As Long, ByRef root As Variant)
    Dim lvl As Long
    For lvl = changedLevel To 2 Step -1
        Dim parentArr As Variant
        vAssignment parentArr, containers(lvl - 1)
        vAssignment parentArr(posAbs(lvl - 1)), containers(lvl)
        vAssignment containers(lvl - 1), parentArr
    Next lvl
    ' finally update root to containers(1)
    vAssignment root, containers(1)
End Sub

Private Function ResolveIndexProp(ByRef lval) As Variant
    Dim baseVariable As String
    Dim pObject As Variant
    Dim keys As New Collection
    Dim i As Long
    Dim container As Variant: vAssignment container, lval("container")
    
    If Not lval.Exists("parentObj") Then
        vAssignment ResolveIndexProp, container
        Exit Function
    End If
    
    vAssignment pObject, lval("parentObj")
    Do While pObject.Exists("parentObj")
        If pObject.Exists("key") Then keys.Add pObject("key")
        vAssignment pObject, pObject("parentObj")
        If Not pObject.Exists("parentKey") Then
            Exit Do
        End If
    Loop
    
    If Not pObject.Exists("key") Then
        ' Return container directly
        vAssignment ResolveIndexProp, container
        Exit Function
    End If
    
    baseVariable = pObject("key")
    Select Case TypeName(container)
        Case "ASF_Map", "ASF_ScopeStack"
            ' Container is an ASF structure
            vAssignment pObject, container.GetValue(baseVariable)
            If keys.Count > 0 Then
                For i = keys.Count To 2 Step -1
                    vAssignment pObject, pObject.GetValue(keys(i))
                Next
                vAssignment ResolveIndexProp, pObject(keys(1))
            Else
                vAssignment ResolveIndexProp, pObject
            End If
        Case Else
            ' Container is a native VBA object or array
            vAssignment ResolveIndexProp, container
    End Select
End Function
Private Function indexDecompose(indexStr As String) As Variant
    
End Function

Private Function ResolveProp(ByRef lval) As Variant
    Dim baseVariable As String
    Dim pObject As Variant

    vAssignment pObject, lval("parentObj")
    Do While pObject.Exists("parentObj")
        vAssignment pObject, pObject("parentObj")
        If Not pObject.Exists("parentKey") Then
            Exit Do
        End If
    Loop
    baseVariable = pObject("key")
    Dim container As Variant: vAssignment container, lval("container")
    vAssignment pObject, container.GetValue(baseVariable)
    vAssignment ResolveProp, pObject
End Function
Private Sub AssignToArrayByName(arrVarName As String, idx As Long, newValue As Variant, progScope As ASF_ScopeStack)
     Dim arrVal As Variant: arrVal = progScope.GetValue(arrVarName)
     If IsEmpty(arrVal) Then
         Dim tmpA() As Variant
         ReDim tmpA(1 To 0)
         arrVal = tmpA
     End If
     Call AssignToArrayValueInPlace(arrVal, idx, newValue)
     progScope.SetValue arrVarName, arrVal
End Sub

Private Sub AssignToArrayValueInPlace(ByRef arrVal As Variant, idx As Long, _
                                newValue As Variant, Optional oBase As Long, _
                                Optional keypath As Collection = Nothing)
     If Not IsArray(arrVal) Then
         Dim na() As Variant
         ReDim na(1 To idx)
         na(idx) = newValue
         arrVal = na
         Exit Sub
     End If
     Dim lb As Long, ub As Long
     Dim pos As Long
     If Not keypath Is Nothing Then
        Dim tmpsltCopy As Variant
        Dim tmpSlot As Variant
        Dim i As Long
        For i = keypath.Count - 1 To 1 Step -1
            If i = keypath.Count - 1 Then
                If LBound(arrVal) <> oBase Then
                    vAssignment tmpSlot, arrVal(NormalizeIndex(keypath(i), UBound(arrVal) - LBound(arrVal) + 1, oBase))
                Else
                    vAssignment tmpSlot, arrVal(keypath(i))
                End If
            Else
                If LBound(tmpSlot) <> oBase Then
                    vAssignment tmpsltCopy, tmpSlot
                    vAssignment tmpSlot, tmpsltCopy(NormalizeIndex(keypath(i), UBound(tmpSlot) - LBound(tmpSlot) + 1, oBase))
                Else
                    vAssignment tmpsltCopy, tmpSlot
                    vAssignment tmpSlot, tmpsltCopy(keypath(i))
                End If
            End If
        Next i
        If IsEmpty(tmpSlot) Then
            Dim emptySlot() As Variant
            ReDim emptySlot(oBase To oBase)
            vAssignment tmpSlot, emptySlot
        End If
        If IsArray(tmpSlot) Then
            Dim pathToSlot() As Variant
            lb = LBound(tmpSlot)
            ub = UBound(tmpSlot)
            pos = lb + idx - 1
            If pos < lb Then
                err.Raise vbObjectError + 9120, "VM.AssignToArrayValueInPlace", "Index out of bounds (below LBound)"
            End If
            If pos > ub Then
                ReDim Preserve tmpSlot(lb To pos)
            End If
            vAssignment tmpSlot(pos), newValue
            ReDim pathToSlot(1 To keypath.Count - 1)
            vAssignment pathToSlot(1), keypath(keypath.Count)
            For i = 1 To keypath.Count - 2
                vAssignment pathToSlot(1 + i), keypath(i)
            Next i
            UpdateNestedArrayByPath arrVal, pathToSlot, tmpSlot, oBase, True
        Else
            If LBound(arrVal) <> oBase Then
                vAssignment arrVal(NormalizeIndex(keypath(keypath.Count - 1), _
                                    UBound(arrVal) - LBound(arrVal) + 1, oBase)), newValue
            Else
                vAssignment arrVal(keypath(keypath.Count - 1)), newValue
            End If
        End If
     Else
        lb = LBound(arrVal)
        ub = UBound(arrVal)
        pos = lb + idx - 1
        If pos < lb Then
            err.Raise vbObjectError + 9120, "VM.AssignToArrayValueInPlace", "Index out of bounds (below LBound)"
        End If
        If pos > ub Then
            ReDim Preserve arrVal(lb To pos)
        End If
        vAssignment arrVal(pos), newValue
    End If
End Sub
 ' RebaseArray: recursively rebase a Variant (possibly jagged) array so each nested
' Variant array uses the given optionBase as its LBound.
' - arr: Variant expected to be an array (if not, the function returns arr unchanged)
' - optionBase: desired lower bound for every array that is encountered (0 or 1 typically)
' Returns: a Variant array (newly allocated) where every nested Variant array is re-based.
Public Function RebaseArray(arr As Variant, optionBase As Long) As Variant
    On Error GoTo ErrHandler

    ' If not an array, return as-is
    If Not IsArray(arr) Then
        RebaseArray = arr
        Exit Function
    End If

    Dim lb As Long, ub As Long
    lb = LBound(arr)
    ub = UBound(arr)

    ' handle empty-like arrays gracefully: if ub < lb produce single-slot array per ASF convention
    Dim Count As Long
    If ub >= lb Then
        Count = ub - lb + 1
    Else
        Count = 0
    End If

    Dim tmpArr() As Variant
    If Count <= 0 Then
        ' empty -> keep single-slot convention
        ReDim tmpArr(optionBase To optionBase)
        RebaseArray = tmpArr
        Exit Function
    Else
        ReDim tmpArr(optionBase To optionBase + Count - 1)
    End If

    Dim i As Long, destIdx As Long
    For i = lb To ub
        destIdx = optionBase + (i - lb)
        Dim v As Variant
        v = arr(i)

        If IsArray(v) Then
            ' recursive rebase for nested arrays
            tmpArr(destIdx) = RebaseArray(v, optionBase)
        Else
            ' non-array: copy as-is
            tmpArr(destIdx) = v
        End If
    Next i

    RebaseArray = tmpArr
    Exit Function

ErrHandler:
    ' On error return the original array unchanged (fail-safe)
    RebaseArray = arr
End Function

'-----------------------------------------------------------------------
' GetElementMapFromIndexNode
' Given an Index AST node, resolve the array container and return the actual
' element Map stored at the index (creating a Map element if empty).
'-----------------------------------------------------------------------
Private Function GetElementMapFromIndexNode(indexNode As ASF_Map, progScope As ASF_ScopeStack) As ASF_Map
    Dim resolved As Variant
    resolved = ResolveLValue(indexNode, progScope)
    If IsEmpty(resolved) Then
        Set GetElementMapFromIndexNode = Nothing
        Exit Function
    End If
    Dim kind As String: kind = CStr(resolved(0))
    Select Case kind
    Case "arrayInScope"
        Dim arrName As String: arrName = CStr(resolved(1))
        Dim idx As Long: idx = CLng(resolved(2))
        Dim arrVal As Variant: arrVal = progScope.GetValue(arrName)
        If IsEmpty(arrVal) Then
            Dim tmpA() As Variant
            ReDim tmpA(1 To 0)
            arrVal = tmpA
        End If
        Dim lb As Long: lb = LBound(arrVal)
        Dim pos As Long: pos = lb + idx - 1
        If pos > UBound(arrVal) Then
            ' extend array so the slot exists
            ReDim Preserve arrVal(lb To pos)
        End If
        Dim elem As Variant: elem = arrVal(pos)
        If IsEmpty(elem) Then
            Dim nm As ASF_Map: Set nm = MakeNode("ASF_Map")
            arrVal(pos) = nm
            progScope.SetValue arrName, arrVal
            Set GetElementMapFromIndexNode = nm
            Exit Function
        ElseIf TypeName(elem) = "ASF_Map" Then
            Set GetElementMapFromIndexNode = elem
            Exit Function
        Else
            ' not a Map stored in the array slot
            err.Raise vbObjectError + 9021, "VM.GetElementMapFromIndexNode", "Array element is not an object"
        End If
    Case "arrayInMap"
        ' resolved shape: Array("arrayInMap", mapObj, index, propName)
        Dim mapObj As ASF_Map: Set mapObj = resolved(1)
        Dim theIdx As Long: theIdx = CLng(resolved(2))
        Dim propName As String: propName = CStr(resolved(3))
        Dim arrVal2 As Variant: arrVal2 = mapObj.GetValue(propName)
        If IsEmpty(arrVal2) Then
            Dim tmpB() As Variant
            ReDim tmpB(1 To 0)
            arrVal2 = tmpB
        End If
        Dim lb2 As Long: lb2 = LBound(arrVal2)
        Dim pos2 As Long: pos2 = lb2 + theIdx - 1
        If pos2 > UBound(arrVal2) Then
            ReDim Preserve arrVal2(lb2 To pos2)
        End If
        Dim elem2 As Variant: elem2 = arrVal2(pos2)
        If IsEmpty(elem2) Then
            Dim nm2 As ASF_Map: Set nm2 = MakeNode("ASF_Map")
            arrVal2(pos2) = nm2
            mapObj.SetValue propName, arrVal2
            Set GetElementMapFromIndexNode = nm2
            Exit Function
        ElseIf TypeName(elem2) = "ASF_Map" Then
            Set GetElementMapFromIndexNode = elem2
            Exit Function
        Else
            err.Raise vbObjectError + 9022, "VM.GetElementMapFromIndexNode", "Array element is not an object (map)"
        End If
    Case Else
        err.Raise vbObjectError + 9023, "VM.GetElementMapFromIndexNode", "Unsupported resolved kind for index -> " & kind
    End Select
End Function
' ---------------------
' Builtin helper functions (range/flatten/clone) and helpers for array methods
' ---------------------

' Flatten builtin with optional depth
' Usage:
'   builtin_flatten(value, progScope)         -> full flatten (old behavior)
'   builtin_flatten(value, progScope, depth) -> flatten up to 'depth' levels
Private Function builtin_flatten(v As Variant, progScope As ASF_ScopeStack, Optional depthParam As Variant) As Variant
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
    Dim ob As Long: ob = progScope.GetValue("__option_base")
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
            ' Do not flatten this array level — push the array object as an element.
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
Private Function builtin_clone(v As Variant, progScope As ASF_ScopeStack) As Variant
    If IsArray(v) Then
        Dim lb As Long: lb = LBound(v)
        Dim ub As Long: ub = UBound(v)
        Dim outArr() As Variant
        ReDim outArr(lb To ub)
        Dim i As Long
        For i = lb To ub
            outArr(i) = builtin_clone(v(i), progScope)
        Next i
        vAssignment builtin_clone, outArr
        Exit Function
    End If
    If IsObject(v) Then
        If TypeName(v) = "ASF_Map" Then
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
                                    progScope As ASF_ScopeStack, Optional parentThis As Variant, _
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
            vAssignment childRes, map_apply_recursive(value(srcIdx), mapper, progScope, value, optBase)
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
            builtin_typeof = "object" & ": <" & TypeName(v) & ">"
        End If
        Exit Function
    End If

    ' ASF_Map (objects and closures) — detect function closures specially
    If IsObject(v) Then
        ' if it's an ASF_Map, look inside for "type" = "Closure"
        Select Case TypeName(v)
            Case "ASF_Map"
                If v.Exists("type") Then
                    Dim tP As Variant: tP = v.GetValue("type")
                    If CStr(tP) = "Closure" Then
                        builtin_typeof = "function"
                        Exit Function
                    End If
                End If
                ' generic ASF_Map object
                builtin_typeof = "object"
                Exit Function
            Case "ASF", "ASF_Compiler", "ASF_Globals", "ASF_Parser", _
                "ASF_RegexEngine", "ASF_ScopeStack", "ASF_VM"
                builtin_typeof = "object"
                Exit Function
        End Select
        ' other COM/host objects
        builtin_typeof = "object" & ": <" & TypeName(v) & ">"
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

' Convert VBA fixed/dynamic array to a 0-based logical list (Collection)
Private Function ArrayToList(arr As Variant) As Collection
    Dim col As New Collection
    If Not IsArray(arr) Then
        Set ArrayToList = col
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        col.Add arr(i)
    Next i
    Set ArrayToList = col
End Function

' Convert 0-based Collection back to array using runtime option base
Private Function ListToArray(col As Collection, optBase As Long) As Variant
    Dim outArr() As Variant
    If col.Count = 0 Then
        ReDim outArr(optBase To optBase)
    Else
        ReDim outArr(optBase To optBase + col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            vAssignment outArr(optBase + i - 1), col(i)
        Next i
    End If
    ListToArray = outArr
End Function

' Deep equality for nested arrays and ASF_Map (objects)
Private Function DeepEqual(a As Variant, b As Variant) As Boolean
    On Error GoTo notEqual
    ' same reference?
    If IsObject(a) And IsObject(b) Then
        If TypeName(a) = "ASF_Map" And TypeName(b) = "ASF_Map" Then
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

' Helper: get element at logical index (option base aware)
Private Function GetElementAt(arr As Variant, userIndex As Long, optBase As Long) As Variant
    If Not IsArray(arr) Then GetElementAt = Empty: Exit Function
    Dim baseL As Long: baseL = LBound(arr)
    Dim off As Long: off = userIndex - optBase
    Dim pos As Long: pos = baseL + off
    If pos < LBound(arr) Or pos > UBound(arr) Then GetElementAt = Empty Else vAssignment GetElementAt, arr(pos)
End Function

' Helper: set element at logical index (resizing if necessary)
Private Function SetElementAt(ByRef arr As Variant, userIndex As Long, optBase As Long, value As Variant) As Variant
    ' arr passed by reference - ensure bounds
    Dim ub As Long, lb As Long
    If Not IsArray(arr) Then
        ' create minimal array
        ReDim arr(optBase To optBase)
        vAssignment arr(optBase), value
        vAssignment SetElementAt, arr
        Exit Function
    End If
    lb = LBound(arr)
    ub = UBound(arr)
    Dim off As Long: off = userIndex - optBase
    Dim pos As Long: pos = lb + off
    If pos < lb Then
        ' grow lower bound (rare). Simplest: create new array with same elements shifted.
        Dim tmp() As Variant
        ReDim tmp(pos To pos + (ub - lb))
        Dim i As Long, j As Long: j = lb
        For i = pos To pos + (ub - lb)
            If i = pos Then
                vAssignment tmp(i), value
            Else
                vAssignment tmp(i), arr(j)
                j = j + 1
            End If
        Next i
        arr = tmp
    ElseIf pos > ub Then
        ' need to extend upper bound
        Dim newUB As Long: newUB = pos
        Dim tmp2() As Variant
        ReDim tmp2(lb To newUB)
        Dim k As Long
        For k = lb To ub
            vAssignment tmp2(k), arr(k)
        Next k
        vAssignment tmp2(pos), value
        arr = tmp2
    Else
        vAssignment arr(pos), value
    End If
    vAssignment SetElementAt, arr
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

' Helper Write value back to a property
Private Sub WriteToContainer(ByRef calleeExpr As ASF_Map, ByRef progScope As ASF_ScopeStack, tValue As Variant)
    Dim rp As ASF_Map
    Dim cont As Variant
    Dim propKey As String
    
    Set rp = ResolvePath(calleeExpr.GetValue("base"), progScope, True)
    If rp.Exists("kind") Then
        Select Case rp.GetValue("kind")
            Case "scopeVar"
                progScope.SetValue CStr(rp.GetValue("key")), tValue
            Case "prop"
                vAssignment cont, ResolveProp(rp)
                propKey = CStr(rp.GetValue("key"))
                If IsObject(cont) Then
                    Select Case TypeName(cont)
                        Case "ASF_Map"
                            ' set property on map
                            cont.SetValue propKey, tValue
                        Case "ASF_ScopeStack"
                            ' set named variable in scope stack
                            cont.SetValue propKey, tValue
                        Case Else
                            ' try a late-bound SetValue, otherwise fallback to current progScope assignment
                            On Error Resume Next
                            cont.SetValue propKey, tValue
                            If err.Number <> 0 Then
                                err.Clear
                                progScope.SetValue propKey, tValue
                            End If
                    End Select
                Else
                    ' couldn't resolve a container object — fall back to storing in current scope
                    progScope.SetValue propKey, tValue
                End If
        End Select
    Else
        'On the fly value without variable
    End If
End Sub
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

' Returns a Collection of keys (1-based).
Private Function GetMapKeys(mapObj As Variant) As Collection
    Dim out As New Collection
    On Error GoTo GetKeysFallback

    If TypeName(mapObj) <> "ASF_Map" Then
        Set GetMapKeys = out
        Exit Function
    End If

    ' Try: mapObj.Keys -> returns Collection
    On Error Resume Next
    Dim tmpKeys As Variant
    tmpKeys = Empty
    vAssignment tmpKeys, mapObj.keys
    If err.Number = 0 Then
        Dim k As Variant
        For Each k In tmpKeys
            out.Add k
        Next k
        Set GetMapKeys = out
        Exit Function
    End If
    err.Clear

GetKeysFallback:
    ' If we reach here, return empty collection
    Set GetMapKeys = out
End Function

' Helper: compute replacement text (supports string or Closure)
Private Function VM_StringReplacementText(ByVal replRaw As Variant, ByVal matchStr As String, ByVal matchIndexZeroBased As Long, ByVal originalString As String) As String
    On Error GoTo fallback_str

    If Not IsEmpty(replRaw) And TypeName(replRaw) = "ASF_Map" Then
        If replRaw.Exists("type") And CStr(replRaw.GetValue("type")) = "Closure" Then
            Dim callArgsR As New Collection
            callArgsR.Add matchStr
            callArgsR.Add matchIndexZeroBased
            callArgsR.Add originalString
            Dim callRes As Variant
            Dim tmpMapObj As ASF_Map
            vAssignment tmpMapObj, replRaw
            vAssignment callRes, CallClosure(tmpMapObj, callArgsR, originalString)
            If IsArray(callRes) Or (IsObject(callRes) And TypeName(callRes) = "ASF_Map") Then
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
            Dim tn As String: tn = TypeName(v)
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

' Convert captures to a string array (forces CStr for each capture; nested collections become nested string arrays).
Private Function RegexCapsToStringArray(ByVal caps As Collection) As Variant
    If caps Is Nothing Then
        RegexCapsToStringArray = Empty
        Exit Function
    End If

    Dim n As Long: n = caps.Count
    If n = 0 Then
        RegexCapsToStringArray = Array()
        Exit Function
    End If

    Dim out() As String
    ReDim out(0 To n - 1)

    Dim i As Long
    For i = 1 To n
        Dim v As Variant: v = caps.Item(i)
        If IsObject(v) Then
            If TypeName(v) = "Collection" Then
                ' we can't store an array in a String() direct; return a Variant array for caller
                ' so fall back to RegexCapsToVariantArray if nested structure detected
                RegexCapsToStringArray = RegexCapsToVariantArray(caps)
                Exit Function
            Else
                out(i - 1) = CStr(v)
            End If
        ElseIf IsEmpty(v) Then
            out(i - 1) = vbNullString
        ElseIf IsNull(v) Then
            out(i - 1) = "Null"
        Else
            out(i - 1) = CStr(v)
        End If
    Next i

    RegexCapsToStringArray = out
End Function
' Execute class declaration
Private Function ExecClassDecl(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim className As String: className = node.GetValue("name")
    
    ' Create class constructor function
    Dim classMap As New ASF_Map
    classMap.Add "type", "Class"
    classMap.SetValue "name", className
    classMap.SetValue "fields", node.GetValue("fields")
    classMap.SetValue "classMethods", node.GetValue("classMethods")
    
    If node.Exists("constructor") Then
        classMap.SetValue "constructor", node.GetValue("constructor")
    End If
    
    If node.Exists("superClass") Then
        classMap.SetValue "superClass", node.GetValue("superClass")
    End If
    
    ' Store class in scope
    progScope.SetValue className, classMap
    
    ExecClassDecl = vbNullString
End Function
' Execute array destructuring
Private Function ExecArrayDestruct(node As ASF_Map, progScope As ASF_ScopeStack) As String
    ' Handle: [a, b, c] = sourceArray or [a, b, ...rest] = sourceArray
    Dim srcVal As Variant
    vAssignment srcVal, EvalExprNode(node.GetValue("source"), progScope)
    
    If Not IsArray(srcVal) Then
        err.Raise vbObjectError + 6001, "VM.ExecuteStmtNode", "Cannot destructure non-array value"
    End If
    
    Dim destTargets As Collection: Set destTargets = node.GetValue("targets")
    Dim srcLB As Long: srcLB = LBound(srcVal)
    Dim srcUB As Long: srcUB = UBound(srcVal)
    Dim srcLen As Long: srcLen = srcUB - srcLB + 1
    Dim optBase As Long: optBase = progScope.GetValue("__option_base")
    
    ' Assign to regular targets
    Dim ti As Long
    Dim targetNode As ASF_Map
    Dim targetName As String
    Dim srcIdx As Long
    For ti = 1 To destTargets.Count
        Set targetNode = destTargets(ti)
        targetName = targetNode.GetValue("name")
        srcIdx = srcLB + (ti - 1)
        ' Assign value or Empty if out of bounds
        If srcIdx >= srcLB And srcIdx <= srcUB Then
            progScope.SetValue targetName, srcVal(srcIdx)
        Else
            progScope.SetValue targetName, Empty
        End If
    Next ti
    
    ' Handle rest element if present
    If node.Exists("restTarget") Then
        Dim restName As String: restName = node.GetValue("restTarget")
        Dim restStartIdx As Long: restStartIdx = destTargets.Count
        
        ' Calculate how many elements remain for rest
        Dim remainingCount As Long
        remainingCount = srcLen - restStartIdx
        If remainingCount > 0 Then
            ' Create rest array
            Dim restArr() As Variant
            ReDim restArr(optBase To optBase + remainingCount - 1)
            
            Dim ri As Long, srcPos As Long
            srcPos = srcLB + restStartIdx
            For ri = optBase To optBase + remainingCount - 1
                If srcPos <= srcUB Then
                    vAssignment restArr(ri), srcVal(srcPos)
                    srcPos = srcPos + 1
                Else
                    restArr(ri) = Empty
                End If
            Next ri
            
            progScope.SetValue restName, restArr
        Else
            ' No remaining elements, set to empty array
            Dim emptyRestArr() As Variant
            ReDim emptyRestArr(-1 To -1)
            progScope.SetValue restName, emptyRestArr
        End If
    End If
    ExecArrayDestruct = vbNullString
End Function
' Execute module import
Private Function ExecModImport(node As ASF_Map, progScope As ASF_ScopeStack) As String
    ' Handle import statement
    Dim modulePath As String: modulePath = node.GetValue("source")
    Dim resolvedPath As String: resolvedPath = ResolveModulePath(modulePath)
    
    ' Load module if not already loaded
    Dim moduleObj As ASF_Map
    Set moduleObj = LoadModule(resolvedPath)
    
    ' Import default
    If node.Exists("defaultImport") Then
        Dim defaultName As String: defaultName = node.GetValue("defaultImport")
        If moduleObj.Exists("defaultExport") Then
            progScope.SetValue defaultName, moduleObj.GetValue("defaultExport")
        Else
            progScope.SetValue defaultName, Empty
        End If
    End If
    
    ' Import namespace
    If node.Exists("namespaceImport") Then
        Dim nsName As String: nsName = node.GetValue("namespaceImport")
        ' Create namespace object with all exports
        Dim nsObj As New ASF_Map
        Dim exportsMap As ASF_Map: Set exportsMap = moduleObj.GetValue("exports")
        Dim expKey As Variant
        For Each expKey In exportsMap.keys
            nsObj.SetValue CStr(expKey), exportsMap.GetValue(CStr(expKey))
        Next expKey
        progScope.SetValue nsName, nsObj
    End If
    
    ' Import named
    If node.Exists("namedImports") Then
        Dim namedImports As Collection: Set namedImports = node.GetValue("namedImports")
        Dim importSpec As ASF_Map
        Dim ii As Long
        For ii = 1 To namedImports.Count
            Set importSpec = namedImports(ii)
            Dim importedName As String: importedName = importSpec.GetValue("imported")
            Dim localName As String: localName = importSpec.GetValue("local")
            
            Dim exportsMap2 As ASF_Map: Set exportsMap2 = moduleObj.GetValue("exports")
            If exportsMap2.Exists(importedName) Then
                progScope.SetValue localName, exportsMap2.GetValue(importedName)
            Else
                err.Raise vbObjectError + 9001, "VM.ExecuteStmtNode", "Export '" & importedName & "' not found in module '" & resolvedPath & "'"
            End If
        Next ii
    End If
    ExecModImport = vbNullString
End Function
' Execute module export
Private Function ExecModExport(node As ASF_Map, progScope As ASF_ScopeStack) As String
    ' Handle export statement
    ' Exports are collected during module execution and stored in module object
    ' Handles the execution-time behavior
    
    If node.GetValue("isDefault") Then
        ' Export default: evaluate expression and store
        Dim defaultValue As Variant
        vAssignment defaultValue, EvalExprNode(node.GetValue("expression"), progScope)
        
        ' Store in module's export registry (requires module context)
        If GLOBALS_.gModuleExports.Exists(GLOBALS_.CURRENT_MODULE_PATH) Then
            Dim modExports As ASF_Map
            Set modExports = GLOBALS_.gModuleExports.GetValue(GLOBALS_.CURRENT_MODULE_PATH)
            modExports.SetValue "__default", defaultValue
        End If
    Else
        ' Named exports
        If node.Exists("namedExports") Then
            Dim namedExportsList As Collection: Set namedExportsList = node.GetValue("namedExports")
            Dim exportSpec2 As ASF_Map
            Dim jj As Long
            
            If GLOBALS_.gModuleExports.Exists(GLOBALS_.CURRENT_MODULE_PATH) Then
                Set modExports = GLOBALS_.gModuleExports.GetValue(GLOBALS_.CURRENT_MODULE_PATH)
                
                For jj = 1 To namedExportsList.Count
                    Set exportSpec2 = namedExportsList(jj)
                    Dim localVarName As String: localVarName = exportSpec2.GetValue("local")
                    Dim exportedVarName As String: exportedVarName = exportSpec2.GetValue("exported")

                    ' Get value from current scope or gFuncTable
                    Dim exportVarNode As ASF_Map
                    Set exportVarNode = MakeNode("Variable")
                    exportVarNode.SetValue "name", localVarName
                    Dim exportValue As Variant
                    vAssignment exportValue, EvalExprNode(exportVarNode, progScope)
                    modExports.SetValue exportedVarName, exportValue
                Next jj
            End If
        ElseIf node.Exists("declarationType") Then
            ' Function/class export - the declaration will execute first
            ' Then we mark it for export
            Dim declName As String: declName = node.GetValue("declarationName")

            ' Store the declaration name for post-processing
            If GLOBALS_.gModuleExports.Exists(GLOBALS_.CURRENT_MODULE_PATH) Then
                Set modExports = GLOBALS_.gModuleExports.GetValue(GLOBALS_.CURRENT_MODULE_PATH)
                modExports.SetValue "__pending_" & declName, True
            End If
        End If
    End If
    ExecModExport = vbNullString
End Function
' Instantiate a class
Private Function InstantiateClass(classMap As ASF_Map, args As Collection, progScope As ASF_ScopeStack) As Variant
    On Error GoTo ErrHandler
    
    ' Create instance object
    Dim instance As ASF_Map
    Set instance = New ASF_Map
    instance.Add "type", "Instance"
    instance.SetValue "className", classMap.GetValue("name")
    
    ' Handle inheritance - copy parent fields and methods
    If classMap.Exists("superClass") Then
        Dim superClassName As String: superClassName = classMap.GetValue("superClass")
        Dim superClassMap As Variant
        vAssignment superClassMap, progScope.GetValue(superClassName)
        
        If TypeName(superClassMap) = "ASF_Map" And superClassMap.GetValue("type") = "Class" Then
            ' Initialize parent fields
            If superClassMap.Exists("fields") Then
                Call InitializeFields(instance, superClassMap.GetValue("fields"), progScope)
            End If
            
            ' Copy parent methods
            If superClassMap.Exists("classMethods") Then
                Dim parentMethods As Collection: Set parentMethods = superClassMap.GetValue("classMethods")
                Dim pm As Variant
                Dim pmNode As ASF_Map
                For Each pm In parentMethods
                    Set pmNode = pm
                    Dim methodName As String: methodName = pmNode.GetValue("name")
                    instance.SetValue methodName, CreateBoundMethod(pmNode, instance)
                Next pm
            End If
        End If
    End If
    
    ' Initialize fields with defaults
    If classMap.Exists("fields") Then
        Call InitializeFields(instance, classMap.GetValue("fields"), progScope)
    End If
    
    ' Bind methods to instance
    If classMap.Exists("classMethods") Then
        Dim methods As Collection: Set methods = classMap.GetValue("classMethods")
        Dim m As Variant
        For Each m In methods
            Dim methodNode As ASF_Map: Set methodNode = m
            Dim mName As String: mName = methodNode.GetValue("name")
            Dim isStaticMethod As Boolean: isStaticMethod = False
            If methodNode.Exists("isStatic") Then
                isStaticMethod = methodNode.GetValue("isStatic")
            End If
            
            ' Only bind non-static methods to instance
            If Not isStaticMethod Then
                instance.SetValue mName, CreateBoundMethod(methodNode, instance)
            End If
        Next m
    End If
    
    ' Call constructor if exists
    If classMap.Exists("constructor") Then
        Dim ctorNode As ASF_Map: Set ctorNode = classMap.GetValue("constructor")
        Call ExecuteConstructor(ctorNode, instance, args, progScope)
    ElseIf classMap.Exists("superClass") Then
        ' Call parent constructor if no constructor defined
        Dim superClassName2 As String: superClassName2 = classMap.GetValue("superClass")
        Dim superClassMap2 As Variant
        vAssignment superClassMap2, progScope.GetValue(superClassName2)
        
        If TypeName(superClassMap2) = "ASF_Map" And superClassMap2.Exists("constructor") Then
            Dim parentCtor As ASF_Map: Set parentCtor = superClassMap2.GetValue("constructor")
            Call ExecuteConstructor(parentCtor, instance, args, progScope)
        End If
    End If
    
    vAssignment InstantiateClass, instance
    Exit Function
    
ErrHandler:
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "InstantiateClass error: " & err.Description
    err.Clear
    vAssignment InstantiateClass, Empty
End Function

' Initialize instance fields
Private Sub InitializeFields(instance As ASF_Map, fields As Collection, progScope As ASF_ScopeStack)
    Dim f As Variant
    For Each f In fields
        Dim fieldNode As ASF_Map: Set fieldNode = f
        Dim fieldName As String: fieldName = fieldNode.GetValue("name")
        
        Dim defaultValue As Variant
        If fieldNode.Exists("default") Then
            Dim defaultExpr As ASF_Map: Set defaultExpr = fieldNode.GetValue("default")
            vAssignment defaultValue, EvalExprNode(defaultExpr, progScope)
        Else
            defaultValue = Empty
        End If
        
        instance.SetValue fieldName, defaultValue
    Next f
End Sub

' Create a bound method (closure with 'this' bound to instance)
Private Function CreateBoundMethod(methodNode As ASF_Map, instance As ASF_Map) As ASF_Map
    Dim boundMethod As New ASF_Map
    boundMethod.Add "type", "BoundMethod"
    boundMethod.SetValue "method", methodNode
    boundMethod.SetValue "instance", instance
    Set CreateBoundMethod = boundMethod
End Function

' Execute constructor
Private Sub ExecuteConstructor(ctorNode As ASF_Map, instance As ASF_Map, args As Collection, progScope As ASF_ScopeStack)
    ' Create execution scope
    Dim ctorScope As New ASF_ScopeStack
    ctorScope.LoadRawByRef progScope.RawByRef
    ctorScope.Push
    
    ' Bind 'this' to instance
    ctorScope.SetLocalValue "this", instance
    
    ' Bind parameters
    If ctorNode.Exists("params") Then
        Dim params As Collection: Set params = ctorNode.GetValue("params")
        Dim i As Long
        For i = 1 To params.Count
            If i <= args.Count Then
                ctorScope.SetLocalValue CStr(params(i)), args(i)
            Else
                ctorScope.SetLocalValue CStr(params(i)), Empty
            End If
        Next i
    End If
    
    ' Execute body
    If ctorNode.Exists("body") Then
        Dim body As Collection: Set body = ctorNode.GetValue("body")
        Dim j As Long
        For j = 1 To body.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(body(j), ctorScope)
            If ctrl = "RETURN" Or ctrl = "ERR" Then Exit For
        Next j
    End If
    
    ctorScope.Pop
End Sub
' Execute bound method
Private Function ExecuteBoundMethod(methodNode As ASF_Map, instance As ASF_Map, args As Collection, progScope As ASF_ScopeStack) As Variant
    On Error GoTo ErrHandler
    
    Dim methodName As String
    If methodNode.Exists("name") Then
        methodName = methodNode.GetValue("name")
    Else
        methodName = "<method>"
    End If
    
    ' Trace: Push call
    PushCallTrace methodName, args
    
    ' Create method scope
    Dim methodScope As New ASF_ScopeStack
    methodScope.LoadRawByRef progScope.RawByRef
    methodScope.Push
    
    ' Bind 'this' to instance
    methodScope.SetLocalValue "this", instance
    
    ' Bind parameters
    If methodNode.Exists("params") Then
        Dim params As Collection: Set params = methodNode.GetValue("params")
        Dim i As Long
        For i = 1 To params.Count
            If i <= args.Count Then
                methodScope.SetLocalValue CStr(params(i)), args(i)
            Else
                methodScope.SetLocalValue CStr(params(i)), Empty
            End If
        Next i
    End If
    
    ' Execute body
    Dim retVal As Variant
    retVal = Empty
    
    If methodNode.Exists("body") Then
        Dim body As Collection: Set body = methodNode.GetValue("body")
        Dim j As Long
        For j = 1 To body.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(body(j), methodScope)
            If ctrl = "RETURN" Then
                vAssignment retVal, methodScope.GetValue("__return")
                ' Trace: Pop call with return value
                PopCallTrace methodName, retVal
                vAssignment ExecuteBoundMethod, retVal
                methodScope.Pop
                Exit Function
            ElseIf ctrl = "ERR" Then
                TraceError methodName, "Error during method execution"
                Exit For
            End If
        Next j
    End If
    
    methodScope.Pop
    ' Trace: Pop call with empty return
    PopCallTrace methodName, retVal
    vAssignment ExecuteBoundMethod, retVal
    Exit Function
    
ErrHandler:
    TraceError methodName, err.Description
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "ExecuteBoundMethod error: " & err.Description
    err.Clear
    PopCallTrace methodName, Empty
    vAssignment ExecuteBoundMethod, Empty
End Function

Private Function ExecuteClassMethod(methodNode As ASF_Map, args As Collection, progScope As ASF_ScopeStack) As Variant
    On Error GoTo ErrHandler
    
    ' Create method scope
    Dim methodScope As New ASF_ScopeStack
    methodScope.LoadRawByRef progScope.RawByRef
    methodScope.Push
    
    ' Bind parameters
    If methodNode.Exists("params") Then
        Dim params As Collection: Set params = methodNode.GetValue("params")
        Dim i As Long
        For i = 1 To params.Count
            If i <= args.Count Then
                methodScope.SetLocalValue CStr(params(i)), args(i)
            Else
                methodScope.SetLocalValue CStr(params(i)), Empty
            End If
        Next i
    End If
    
    ' Execute body
    If methodNode.Exists("body") Then
        Dim body As Collection: Set body = methodNode.GetValue("body")
        Dim j As Long
        For j = 1 To body.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(body(j), methodScope)
            If ctrl = "RETURN" Then
                vAssignment ExecuteClassMethod, methodScope.GetValue("__return")
                methodScope.Pop
                Exit Function
            ElseIf ctrl = "ERR" Then
                Exit For
            End If
        Next j
    End If
    
    methodScope.Pop
    vAssignment ExecuteClassMethod, Empty
    Exit Function
    
ErrHandler:
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "ExecuteClassMethod error: " & err.Description
    err.Clear
    vAssignment ExecuteClassMethod, Empty
End Function
' Execute parent constructor (called via super())
Private Sub ExecuteSuperConstructor(ctorNode As ASF_Map, instance As ASF_Map, args As Collection, progScope As ASF_ScopeStack)
    ' Create execution scope
    Dim superScope As New ASF_ScopeStack
    superScope.LoadRawByRef progScope.RawByRef
    superScope.Push
    
    ' Bind 'this' to the SAME instance (important!)
    superScope.SetLocalValue "this", instance
    
    ' Bind parameters
    If ctorNode.Exists("params") Then
        Dim params As Collection: Set params = ctorNode.GetValue("params")
        Dim i As Long
        For i = 1 To params.Count
            If i <= args.Count Then
                superScope.SetLocalValue CStr(params(i)), args(i)
            Else
                superScope.SetLocalValue CStr(params(i)), Empty
            End If
        Next i
    End If
    
    ' Execute body
    If ctorNode.Exists("body") Then
        Dim body As Collection: Set body = ctorNode.GetValue("body")
        Dim j As Long
        For j = 1 To body.Count
            Dim ctrl As String: ctrl = ExecuteStmtNode(body(j), superScope)
            If ctrl = "RETURN" Or ctrl = "ERR" Then Exit For
        Next j
    End If
    
    superScope.Pop
End Sub
' Push a call onto the call stack trace
Private Sub PushCallTrace(funcName As String, args As Collection)
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
Private Sub PopCallTrace(funcName As String, returnValue As Variant)
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
Private Sub TraceError(funcName As String, errDesc As String)
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
    If left$(cleanPath, 2) = "./" Or left$(cleanPath, 3) = "../" Then
        Dim currentDir As String
        If GLOBALS_.CURRENT_MODULE_PATH <> vbNullString Then
            Dim lastSlash As Long
            lastSlash = InStrRev(Replace(GLOBALS_.CURRENT_MODULE_PATH, "\", "/"), "/")
            If lastSlash > 0 Then
                If lastSlash < Len(GLOBALS_.CURRENT_MODULE_PATH) Then
                    currentDir = Replace(GLOBALS_.CURRENT_MODULE_PATH, "\", "/") & "/"
                Else
                    currentDir = Replace(GLOBALS_.CURRENT_MODULE_PATH, "\", "/")
                End If
            Else
                currentDir = vbNullString
            End If
        Else
            currentDir = vbNullString
        End If
        ResolveModulePath = currentDir & mid$(cleanPath, InStr(1, cleanPath, "/") + 1)
    Else
        ' Absolute path or module name
        ResolveModulePath = cleanPath
    End If
End Function

' ---------------------
' Module System: Load Module
' ---------------------
Public Function LoadModule(modulePath As String) As ASF_Map
    ' Check if module is already loaded
    If GLOBALS_.gModuleRegistry.Exists(modulePath) Then
        Dim existingModule As ASF_Map
        Set existingModule = GLOBALS_.gModuleRegistry.GetValue(modulePath)
        If existingModule.GetValue("loaded") Then
            Set LoadModule = existingModule
            Exit Function
        End If
    End If
    
    ' Check for circular dependency
    Dim k As Long
    For k = 1 To GLOBALS_.gLoadingModules.Count
        If CStr(GLOBALS_.gLoadingModules(k)) = modulePath Then
            err.Raise vbObjectError + 9010, "VM.LoadModule", "Circular dependency detected: " & modulePath
        End If
    Next k
    
    ' Mark module as currently loading
    GLOBALS_.gLoadingModules.Add modulePath
    
    Dim moduleSource As String
    On Error Resume Next
    moduleSource = ReadModuleSource(modulePath)
    If err.Number <> 0 Then
        On Error GoTo 0
        err.Raise vbObjectError + 9011, "VM.LoadModule", "Failed to load module: " & modulePath
    End If
    On Error GoTo 0
    
    ' Compile module
    Dim compiler As New ASF_Compiler
    compiler.SetGlobals GLOBALS_
    compiler.CurrentModulePath = modulePath
    compiler.IsModuleMode = True
    
    Dim moduleIdx As Long
    moduleIdx = compiler.CompileProgram(moduleSource, "@module:" & modulePath)
    
    ' Create module object
    Dim moduleObj As New ASF_Map
    moduleObj.SetValue "path", modulePath
    moduleObj.SetValue "loaded", False
    
    ' Create module exports map
    Dim moduleExports As New ASF_Map
    moduleObj.SetValue "exports", moduleExports
    GLOBALS_.gModuleExports.SetValue modulePath, moduleExports
    
    ' Store current module path for execution
    Dim previousModulePath As String: previousModulePath = GLOBALS_.CURRENT_MODULE_PATH
    GLOBALS_.CURRENT_MODULE_PATH = modulePath
    
    ' Execute module
    Dim prevProgIdx As Long: prevProgIdx = 0
    If moduleIdx > 0 And moduleIdx <= GLOBALS_.gPrograms.Count Then
        Dim p As Variant: p = GLOBALS_.gPrograms(moduleIdx)
        Dim stmts As Collection: Set stmts = p(1)
        Dim rawScope As Collection: Set rawScope = p(2)
        
        Dim moduleScope As New ASF_ScopeStack
        moduleScope.LoadRaw rawScope
        moduleScope.SetValue "__option_base", 1
        moduleScope.Push
        
        ' Execute module statements
        Dim mi As Long
        For mi = 1 To stmts.Count
            Dim ctrl As String
            ctrl = ExecuteStmtNode(stmts(mi), moduleScope)
            If ctrl = "RETURN" Or ctrl = "ERR" Then Exit For
        Next mi
        
        moduleScope.Pop
        
        ' Process pending exports (functions/classes declared with export)
        Dim exportKeys As Variant
        vAssignment exportKeys, moduleExports.keys
        Dim ek As Variant
        For Each ek In exportKeys
            Dim ekStr As String: ekStr = CStr(ek)
            If left$(ekStr, 10) = "__pending_" Then
                Dim pendingName As String: pendingName = mid$(ekStr, 11)
                ' Get the value from module scope and export it
                If moduleScope.Has(pendingName) Then
                    moduleExports.SetValue pendingName, moduleScope.GetValue(pendingName)
                End If
                ' Remove pending marker
                moduleExports.Remove ekStr
            End If
        Next ek
        
        ' Handle default export
        If moduleExports.Exists("__default") Then
            moduleObj.SetValue "defaultExport", moduleExports.GetValue("__default")
            moduleExports.Remove "__default"
        End If
    End If
    
    ' Restore module path
    GLOBALS_.CURRENT_MODULE_PATH = previousModulePath
    
    ' Mark module as loaded
    moduleObj.SetValue "loaded", True
    
    ' Store in registry
    GLOBALS_.gModuleRegistry.SetValue modulePath, moduleObj
    
    ' Remove from loading list
    For k = 1 To GLOBALS_.gLoadingModules.Count
        If CStr(GLOBALS_.gLoadingModules(k)) = modulePath Then
            GLOBALS_.gLoadingModules.Remove k
            Exit For
        End If
    Next k
    
    Set LoadModule = moduleObj
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
Private Sub CreateSubFolders(FullFilePath As String)
    Dim cPath As String
    Dim SubFolders() As String
    Dim j As Long
    
    SubFolders() = Split(FullFilePath, "\")
    cPath = SubFolders(LBound(SubFolders))
    For j = LBound(SubFolders) + 1 To UBound(SubFolders) - 1
        cPath = cPath & "\" & SubFolders(j)
        If Not FolderExists(cPath) Then
            MkDir cPath           'Create the subdirectory
        End If
    Next j
End Sub
Private Function FolderExists(ByVal filePath As String) As Boolean
    FolderExists = CBool(LenB(Dir(filePath, vbDirectory + vbHidden)))
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
    ' Trace: Push call
    PushCallTrace CStr(methodName), args
    Dim tmpResult As Variant
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
' ---------------------
' End VM class
' ---------------------

