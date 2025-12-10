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

' Executes AST nodes produced by Compiler (AST).
' Uses Map node types and ScopeStack. Logs to gRuntimeLog.

Public Sub SetGlobals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Sub
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

Public Sub RunProgramByIndex(idx As Long)
    GLOBALS_.ASF_InitGlobals
    If idx < 1 Or idx > GLOBALS_.gPrograms.count Then Exit Sub
    Dim p As Variant: p = GLOBALS_.gPrograms(idx)
    Dim progName As String: progName = p(0)
    Dim stmts As Collection: Set stmts = p(1)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim optionBase As Long: optionBase = IIf(UBound(p) >= 3, p(3), 1)
    Dim progScope As New ASF_ScopeStack
    progScope.LoadRaw rawScope
    progScope.SetValue "__option_base", optionBase
    progScope.Push
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "RUN Program: " & progName
    Dim i As Long
    For i = 1 To stmts.count
        Dim ctrl As String
        ctrl = ExecuteStmtNode(stmts(i), progScope)
        If ctrl = "RETURN" Then
            vAssignment OUTPUT__, progScope.GetValue("__return")
            Exit For
        End If
        If ctrl = "ERR" Then Exit For
    Next i
    progScope.Pop
End Sub

' Execute a statement node (Map). Return control signals: "", "BREAK", "CONTINUE", "RETURN", "ERR"
Private Function ExecuteStmtNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    On Error GoTo ErrHandler
    Dim tp As String: tp = node.GetValue("type")
    Dim rval As Variant
    Dim i As Long
    Select Case tp
        Case "Print"
            Dim args As Collection: Set args = node.GetValue("args")
            Dim outParts As New Collection
            For i = 1 To args.count
                Dim v As Variant: vAssignment v, EvalExprNode(args(i), progScope)
                outParts.Add ValueToStringForPrint(v)
            Next i
            Dim sb As String: sb = ""
            For i = 1 To outParts.count
                If i > 1 Then sb = sb & ", "
                sb = sb & outParts(i)
            Next i
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "PRINT:" & sb
            Debug.Print sb
            ExecuteStmtNode = ""
            Exit Function

        Case "Assign"
            Dim left As ASF_Map: Set left = node.GetValue("left")
            Dim right As ASF_Map: Set right = node.GetValue("right")
            vAssignment rval, EvalExprNode(right, progScope)
            HandleAssignment left, rval, progScope
            If left.GetValue("type") = "Variable" Then
                Dim lName As String: lName = left.GetValue("name")
            End If
            ExecuteStmtNode = ""
            Exit Function

        Case "ExprStmt"
            Dim res As Variant: res = EvalExprNode(node.GetValue("expr"), progScope)
            ExecuteStmtNode = ""
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

        Case Else
            ' unknown node type
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "Unknown statement node: " & tp
            ExecuteStmtNode = ""
            Exit Function
    End Select

ErrHandler:
    If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM statement error: " & err.Description
    err.Clear
    ExecuteStmtNode = "ERR"
End Function

' Reuse your AssignToArray logic
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

'Private Function EvalMemberNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
'    If Not node.Exists("prop") And Not node.Exists("property") Then err.Raise vbObjectError + 5010, "VM.EvalMemberNode", "Missing prop/property in Member node"
'    Dim propKey As String: propKey = IIf(node.Exists("prop"), "prop", "property")  ' Twist: auto-detect
'    Dim prop As String: prop = node.GetValue(propKey)
'    Dim baseVal As Variant
'    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)
'    If TypeName(baseVal) <> "ASF_Map" Then err.Raise vbObjectError + 5007, "VM.EvalMemberNode", "Cannot access property on non-map"
'    vAssignment EvalMemberNode, baseVal.GetValue(prop)
'End Function
Private Function EvalMemberNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    If Not node.Exists("prop") And Not node.Exists("property") Then _
        err.Raise vbObjectError + 5010, "VM.EvalMemberNode", "Missing prop/property in Member node"
    
    On Error GoTo ErrHandler
    Dim propKey As String: propKey = IIf(node.Exists("prop"), "prop", "property")
    Dim prop As String: prop = CStr(node.GetValue(propKey))

    Dim baseVal As Variant
    vAssignment baseVal, EvalExprNode(node.GetValue("base"), progScope)

    ' If base is a Map (object) -> return property value as before
    If TypeName(baseVal) = "ASF_Map" Then
        vAssignment EvalMemberNode, baseVal.GetValue(prop)
        Exit Function
    End If

    ' If base is an array -> avoid hard errors in other cases.
    If IsArray(baseVal) Then
        ' Not a recognized array property -> return Empty (do not raise).
        vAssignment EvalMemberNode, Empty
        Exit Function
    End If

    ' If base is an ASF_ScopeStack, try to read property if present
    If TypeName(baseVal) = "ASF_ScopeStack" Then
        If baseVal.Exists(CStr(prop)) Then
            vAssignment EvalMemberNode, baseVal.GetValue(CStr(prop))
        Else
            vAssignment EvalMemberNode, Empty
        End If
        Exit Function
    End If

    ' For any other base type, return Empty (safe fallback).
    vAssignment EvalMemberNode, Empty
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
    If items Is Nothing Or items.count = 0 Then
        ReDim arr(optionBase To optionBase)
    Else
        ReDim arr(optionBase To optionBase + items.count - 1)
        Dim ii As Long
        For ii = 1 To items.count
            vAssignment arr(optionBase + ii - 1), EvalExprNode(items(ii), progScope)
        Next ii
    End If
    vAssignment EvalArrayNode, arr
End Function

Private Function ParsePath(ByVal path As String) As Collection
    Dim col As New Collection
    Dim pos As Long: pos = 1
    Dim lenP As Long: lenP = Len(path)
    Dim currentPart As String
    While pos <= lenP
        Dim ch As String: ch = Mid(path, pos, 1)
        If ch = "." Then
            If currentPart <> "" Then col.Add currentPart
            currentPart = ""
            pos = pos + 1
        ElseIf ch = "[" Then
            If currentPart <> "" Then col.Add currentPart
            currentPart = ""
            Dim startB As Long: startB = pos
            pos = pos + 1
            While pos <= lenP And Mid(path, pos, 1) <> "]"
                pos = pos + 1
            Wend
            If pos > lenP Then err.Raise vbObjectError + 5003, "VM.ParsePath", "Unmatched bracket in path"
            currentPart = Mid(path, startB, pos - startB + 1)
            col.Add currentPart
            currentPart = ""
            pos = pos + 1
        Else
            currentPart = currentPart & ch
            pos = pos + 1
        End If
    Wend
    If currentPart <> "" Then col.Add currentPart
    Set ParsePath = col
End Function

Private Function ResolvePath(node As ASF_Map, progScope As ASF_ScopeStack, Optional forAssign As Boolean = False) As ASF_Map
    Dim res As New ASF_Map
    Dim tp As String: tp = node.GetValue("type")
    If tp = "Variable" Then
        If (InStr(node.GetValue("name"), ".") > 0 Or InStr(node.GetValue("name"), "[") > 0) Then
            ' Fallback for flat names (if Compiler didn't nest)
            Dim path As String: path = node.GetValue("name")
            Dim parts As Collection: Set parts = ParsePath(path) ' Add ParsePath from previous patch
            Dim current As Variant
            vAssignment current, progScope.GetValue(CStr(parts(1)))
            Dim ii As Long
            For ii = 2 To parts.count - IIf(forAssign, 1, 0)
                Dim part As String: part = parts(ii)
                If left(part, 1) = "[" Then
                    Dim indexStr As String: indexStr = Mid(part, 2, Len(part) - 2)
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
                res.SetValue "key", parts(parts.count)
                res.SetValue "kind", IIf(left(parts(parts.count), 1) = "[", "index", "prop")
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
        Select Case tp
            Case "Member"
                Dim baseRes As ASF_Map: Set baseRes = ResolvePath(node.GetValue("base"), progScope, forAssign)
                If forAssign Then
                    res.SetValue "container", baseRes("container")
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
                    res.SetValue "container", baseRes2("container")
                    res.SetValue "key", EvalExprNode(node.GetValue("index"), progScope)
                    res.SetValue "kind", "index"
                    ' carry forward parent property/key info if present on the base resolution
                    If baseRes2.Exists("key") Then res.SetValue "parentKey", baseRes2("key")
                    If baseRes2.Exists("kind") Then res.SetValue "parentKind", baseRes2("kind")
                    res.SetValue "parentObj", baseRes2
                Else
                    res.SetValue "value", EvalIndexNode(node, progScope)
                End If
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
        For si = 1 To thenStmts.count
            Dim ctrl As String: ctrl = ExecuteStmtNode(thenStmts(si), progScope)
            If ctrl <> "" Then ExecIfNode = ctrl: Exit Function
        Next si
        ExecIfNode = ""
        Exit Function
    End If
    Dim elseifConds As Collection: Set elseifConds = node.GetValue("elseif_conds")
    Dim elseifBlocks As Collection: Set elseifBlocks = node.GetValue("elseif_blocks")
    Dim idx As Long
    For idx = 1 To elseifConds.count
        If IsTruthy(EvalExprNode(elseifConds(idx), progScope)) Then
            Dim bl As Collection: Set bl = elseifBlocks(idx)
            Dim si2 As Long
            For si2 = 1 To bl.count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(bl(si2), progScope)
                If ctrl2 <> "" Then ExecIfNode = ctrl2: Exit Function
            Next si2
            ExecIfNode = ""
            Exit Function
        End If
    Next idx
    If node.GetValue("hasElse") Then
        Dim els As Collection: Set els = node.GetValue("else")
        Dim ee As Long
        For ee = 1 To els.count
            Dim ctrl3 As String: ctrl3 = ExecuteStmtNode(els(ee), progScope)
            If ctrl3 <> "" Then ExecIfNode = ctrl3: Exit Function
        Next ee
    End If
    ExecIfNode = ""
End Function

' ---------------------
' ExecForNode
' ---------------------
Private Function ExecForNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim initNode As ASF_Map: Set initNode = node.GetValue("init")
    Dim condNode As ASF_Map: Set condNode = node.GetValue("cond")
    Dim stepNode As ASF_Map: Set stepNode = node.GetValue("step")
    Dim body As Collection: Set body = node.GetValue("body")
    Dim condOk As Boolean
    
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
    condOk = True
    If Not condNode Is Nothing Then
        condOk = IsTruthy(EvalExprNode(condNode, progScope))
    End If
    If Not condOk Then GoTo ForEnd

    Dim s As Long
    For s = 1 To body.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(body(s), progScope)
        If ctrl = "BREAK" Then GoTo ForEnd
        If ctrl = "CONTINUE" Then Exit For
        If ctrl = "RETURN" Or ctrl = "ERR" Then ExecForNode = ctrl: Exit Function
    Next s

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
    ExecForNode = ""
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
    For i = 1 To body.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(body(i), progScope)
        If ctrl = "BREAK" Then GoTo WhileEnd
        If ctrl = "CONTINUE" Then Exit For
        If ctrl = "RETURN" Or ctrl = "ERR" Then ExecWhileNode = ctrl: Exit Function
    Next i
    GoTo WhileStart

WhileEnd:
    ExecWhileNode = ""
End Function

' ---------------------
' ExecTryCatchNode
' ---------------------
Private Function ExecTryCatchNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    On Error GoTo TryErr
    Dim tryStmts As Collection: Set tryStmts = node.GetValue("try")
    Dim i As Long
    For i = 1 To tryStmts.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(tryStmts(i), progScope)
        If ctrl = "RETURN" Or ctrl = "ERR" Then
            ExecTryCatchNode = ctrl
            If ctrl = "ERR" Then GoTo TryErr
            Exit Function
        End If
    Next i
    ExecTryCatchNode = ""
    Exit Function

TryErr:
    err.Clear
    Dim catchStmts As Collection: Set catchStmts = node.GetValue("catch")
    If Not catchStmts Is Nothing Then
        Dim j As Long
        For j = 1 To catchStmts.count
            Dim cctrl As String: cctrl = ExecuteStmtNode(catchStmts(j), progScope)
            If cctrl = "RETURN" Or cctrl = "ERR" Then ExecTryCatchNode = cctrl: Exit Function
        Next j
    End If
    ExecTryCatchNode = ""
End Function

' ---------------------
' ExecSwitchNode
' ---------------------
Private Function ExecSwitchNode(node As ASF_Map, progScope As ASF_ScopeStack) As String
    Dim switchVal As Variant: switchVal = EvalExprNode(node.GetValue("expr"), progScope)
    Dim cases As Collection: Set cases = node.GetValue("cases")
    Dim i As Long
    Dim match As Boolean
    For i = 1 To cases.count
        Dim pair As Collection: Set pair = cases(i)
        Dim caseExpr As ASF_Map: Set caseExpr = pair(1)
        Dim blockStmts As Collection: Set blockStmts = pair(2)
        match = (EvalExprNode(caseExpr, progScope) = switchVal)
        If match Then
            Dim s As Long
            For s = 1 To blockStmts.count
                Dim ctrl As String: ctrl = ExecuteStmtNode(blockStmts(s), progScope)
                If ctrl = "BREAK" Then ExecSwitchNode = "": Exit Function
                If ctrl = "RETURN" Or ctrl = "ERR" Then ExecSwitchNode = ctrl: Exit Function
            Next s
        End If
    Next i
    If Not match Then
        Dim defBlk As Collection: Set defBlk = node.GetValue("default")
        If Not defBlk Is Nothing Then
            Dim d As Long
            For d = 1 To defBlk.count
                Dim ctrl2 As String: ctrl2 = ExecuteStmtNode(defBlk(d), progScope)
                If ctrl2 = "BREAK" Then ExecSwitchNode = "": Exit Function
                If ctrl2 = "RETURN" Or ctrl2 = "ERR" Then ExecSwitchNode = ctrl2: Exit Function
            Next d
        End If
    End If
    ExecSwitchNode = ""
End Function

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
' ---------------------
' Expression evaluator: Evaluate Expr AST nodes to runtime values
' ---------------------
Private Function EvalExprNode(node As ASF_Map, progScope As ASF_ScopeStack) As Variant
    If node Is Nothing Then EvalExprNode = Empty: Exit Function
    Dim tp As String: tp = node.GetValue("type")
    Dim items As Collection
    Dim tmpResult As Variant
    Dim pi As Long
    Dim funcIdx As Long
    Dim baseExpr As ASF_Map
    Dim tmpMapObj As ASF_Map
    Dim calleeExpr As ASF_Map
    Dim fname As String
    
    Select Case tp
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
            vAssignment tmpResult, cmap
            GoTo exitfun
        Case "VBAexpr"
            ' Forced VBAexpression node (from @(...)). The node stores the raw VB expression string in "expr".
            Dim rawVB As String: rawVB = node.GetValue("expr")
            vAssignment tmpResult, EvalVBExpressionWithScope(rawVB, progScope)
            GoTo exitfun
        Case "Object"
            ' Build and return a Map containing the evaluated properties
            Set items = node.GetValue("items")
            Dim om As New ASF_Map
            Dim valNode As ASF_Map
            Dim pair As Collection
            Dim key As String
            For pi = 1 To items.count
                Set pair = items(pi)
                key = CStr(pair(1))
                Set valNode = pair(2)
                om.SetValue key, EvalExprNode(valNode, progScope)
            Next pi
            ' Return Map object (as Variant holding the object)
            vAssignment tmpResult, om
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
                    Set paramsCol = New Collection
                    If GLOBALS_.gFuncParams.Exists(vname) Then
                        Dim paVar As Variant: paVar = GLOBALS_.gFuncParams.GetValue(vname)
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
                            Dim it As Variant
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

                    vAssignment tmpResult, closureMap
                    GoTo exitfun
                End If
            End If
            On Error GoTo 0

            ' Not found — log and return Empty
            If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "VM: Variable '" & vname & "' is undefined or not callable (returned Empty)."
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
            If fname = "" Then
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
            ' builtin length function
            If fname = ".__len__" Then
                Dim argn As Collection: Set argn = node.GetValue("args")
                Dim av As Variant: av = EvalExprNode(argn(1), progScope)
                If Not IsArray(av) Then
                    vAssignment tmpResult, 0
                Else
                    vAssignment tmpResult, UBound(av) - LBound(av) + 1
                End If
                GoTo exitfun
            End If

            Dim argsColl As Collection: Set argsColl = node.GetValue("args")
            Dim evaluated As Collection
            Set evaluated = New Collection
            Dim k As Long
            For k = 1 To argsColl.count
                evaluated.Add EvalExprNode(argsColl(k), progScope)
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
                    ' --- method builtins on Member callee (map/filter/reduce/slice/push/pop) ---
                    Dim propN As String: propN = LCase$(CStr(calleeExpr.GetValue("prop")))
                    Dim baseValLocal As Variant
                    vAssignment baseValLocal, EvalExprNode(calleeExpr.GetValue("base"), progScope)
                    Select Case propN
                        Case "map"
                            ' map(fn) -> return new array with fn applied to each element (recursively tolerates nested arrays)
                            If evaluated.count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
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
                            Dim outArr() As Variant
                            If countEls = 0 Then
                                ReDim outArr(optBaseLocal To optBaseLocal)
                            Else
                                ReDim outArr(optBaseLocal To optBaseLocal + countEls - 1)
                            End If
                        
                            Dim ii As Long
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
                            If evaluated.count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
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
                            If tmpCol.count = 0 Then
                                ReDim outA(ob To ob)
                            Else
                                ReDim outA(ob To ob + tmpCol.count - 1)
                                For jj = 1 To tmpCol.count
                                    vAssignment outA(ob + jj - 1), tmpCol(jj)
                                Next jj
                            End If
                            vAssignment tmpResult, outA
                            GoTo exitfun
                        Case "reduce"
                            ' reduce(fn, initial?) -> apply reduction
                            If evaluated.count < 1 Then vAssignment tmpResult, Empty: GoTo exitfun
                            Dim rfn As Variant: vAssignment rfn, evaluated(1)
                            If TypeName(rfn) <> "ASF_Map" Or rfn.GetValue("type") <> "Closure" Then
                                vAssignment tmpResult, Empty: GoTo exitfun
                            End If
                            If Not IsArray(baseValLocal) Then vAssignment tmpResult, Empty: GoTo exitfun
                            Dim acc As Variant
                            Dim startIdx As Long
                            If evaluated.count >= 2 Then
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
                            Dim optBase As Long: optBase = progScope.GetValue("__option_base")
                        
                            Dim sRaw As Long, eRaw As Long
                        
                            ' Interpret arguments in the user's indexing space (option base)
                            If evaluated.count = 0 Then
                                ' default: full slice
                                s = 0: e = lenA
                            ElseIf evaluated.count = 1 Then
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
                                Dim si As Long
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
                            If Not IsArray(baseValLocal) Then
                                ' create new array and set it back
                                Dim na() As Variant
                                ReDim na(1 To 1)
                                baseValLocal = na
                            End If
                            Dim insertCount As Long: insertCount = evaluated.count
                            Dim oldLB As Long: oldLB = LBound(baseValLocal)
                            Dim oldUB As Long: oldUB = UBound(baseValLocal)
                            Dim oldLen As Long: oldLen = oldUB - oldLB + 1
                            Dim newLen As Long: newLen = oldLen + insertCount
                            Dim newArr() As Variant
                            ReDim newArr(oldLB To oldLB + newLen - 1)
                            Dim zz As Long
                            For zz = oldLB To oldUB
                                vAssignment newArr(zz), baseValLocal(zz)
                            Next zz
                            For zz = 1 To insertCount
                                vAssignment newArr(oldLB + oldLen + zz - 1), evaluated(zz)
                            Next zz
                            ' write back to original container via ResolvePath
                            Dim rp As ASF_Map: Set rp = ResolvePath(calleeExpr.GetValue("base"), progScope, True)
                            If rp.GetValue("kind") = "scopeVar" Then
                                progScope.SetValue CStr(rp.GetValue("key")), newArr
                            ElseIf rp.GetValue("kind") = "prop" Then
                                Dim cont As Variant: cont = rp.GetValue("container")
                                If TypeName(cont) = "ASF_Map" Then
                                    cont.SetValue CStr(rp.GetValue("key")), newArr
                                ElseIf TypeName(cont) = "ASF_ScopeStack" Then
                                    cont.SetValue CStr(rp.GetValue("key")), newArr
                                End If
                            End If
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
                            Dim rp2 As ASF_Map: Set rp2 = ResolvePath(calleeExpr.GetValue("base"), progScope, True)
                            If rp2.GetValue("kind") = "scopeVar" Then
                                progScope.SetValue CStr(rp2.GetValue("key")), baseValLocal
                            ElseIf rp2.GetValue("kind") = "prop" Then
                                Dim ccont As Variant: vAssignment ccont, rp2.GetValue("container")
                                If TypeName(ccont) = "ASF_Map" Then
                                    ccont.SetValue CStr(rp2.GetValue("key")), baseValLocal
                                ElseIf TypeName(ccont) = "ASF_ScopeStack" Then
                                    ccont.SetValue CStr(rp2.GetValue("key")), baseValLocal
                                End If
                            End If
                            vAssignment tmpResult, popped
                        Case Else
                            GoTo namedFunctions
                    End Select
                    ' --- end method builtins ---
                    GoTo exitfun
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
                    If node.Exists("name") Then fname = CStr(node.GetValue("name")) Else fname = ""
                End If
            End If
            If fname <> "" Then
                ' --- builtin named functions ---
                If LCase$(fname) = "range" Then
                    ' range([start,] end[, step]) -> returns an array (python-style range semantics)
                    Dim rcount As Long: rcount = evaluated.count
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
                    If vals.count = 0 Then
                        ReDim tmpArr(optionBase_local To optionBase_local)
                    Else
                        ReDim tmpArr(optionBase_local To optionBase_local + vals.count - 1)
                        Dim ri As Long
                        For ri = 1 To vals.count
                            vAssignment tmpArr(optionBase_local + ri - 1), vals(ri)
                        Next ri
                    End If
                    vAssignment tmpResult, tmpArr
                    GoTo exitfun
                End If
                If LCase$(fname) = "flatten" Then
                    If evaluated.count = 0 Then
                        vAssignment tmpResult, Empty
                        GoTo exitfun
                    End If
                    vAssignment tmpMapObj, node.GetValue("args")(1)
                    If evaluated.count = 1 Then
                        vAssignment tmpResult, builtin_flatten(EvalExprNode(tmpMapObj, progScope), progScope)
                    Else
                        vAssignment tmpResult, builtin_flatten(EvalExprNode(tmpMapObj, progScope), progScope, evaluated(2))
                    End If
                    GoTo exitfun
                End If
                If LCase$(fname) = "clone" Then
                    If evaluated.count = 0 Then
                        vAssignment tmpResult, Empty
                        GoTo exitfun
                    End If
                    vAssignment tmpMapObj, node.GetValue("args")(1)
                    vAssignment tmpResult, builtin_clone(EvalExprNode(tmpMapObj, progScope), progScope)
                    GoTo exitfun
                End If
                If LCase$(fname) = "isarray" Then
                    If evaluated.count = 0 Then
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
                If LCase$(fname) = "isnumeric" Then
                    If evaluated.count = 0 Then
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
                Dim thisVal As Variant
                ' If callee is a member or index expression, compute base as `this` for method call binding.
                If Not calleeExpr Is Nothing Then
                    Dim ct As String: ct = IIf(calleeExpr.Exists("type"), CStr(calleeExpr.GetValue("type")), "")
                    If ct = "Member" Or ct = "Index" Then
                        hasThisVal = True
                        Set baseExpr = calleeExpr.GetValue("base")
                        vAssignment thisVal, EvalExprNode(baseExpr, progScope)
                    End If
                End If
                vAssignment calleeVal, EvalExprNode(calleeExpr, progScope)
                If TypeName(calleeVal) = "ASF_Map" Then
                    Dim cM As ASF_Map: Set cM = calleeVal
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
            Dim ev As Variant: ev = EvalExprNode(node.GetValue("expr"), progScope)
            If op = "!" Then
                vAssignment tmpResult, Not IsTruthy(ev)
            ElseIf op = "-" Then
                vAssignment tmpResult, -CDbl(ev)
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
                    Case "==", "=": vAssignment tmpResult, (lop = rop)
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

Private Function ReturnCollection(ByRef aColl As Variant) As Collection
    Set ReturnCollection = aColl
End Function
Private Function ArrayIsInit(aArray As Variant) As Boolean
    Dim ub As Long
    ArrayIsInit = True
    On Error GoTo err_handler
    ub = UBound(aArray)
    Exit Function
err_handler:
    err.Clear
    ArrayIsInit = False
End Function
' Call function program by index (AST). "args" is Collection of evaluated arg values.
Private Function CallFuncByIndex_AST(funcIdx As Long, args As Collection) As Variant
    Dim p As Variant: p = GLOBALS_.gPrograms(funcIdx)
    Dim rawScope As Collection: Set rawScope = p(2)
    Dim callScope As New ASF_ScopeStack
    Dim i As Long
    callScope.LoadRaw rawScope
    callScope.Push
    Dim fname As String: fname = p(0)
    ' set params from gFuncParams
    If GLOBALS_.gFuncParams.Exists(fname) Then
        Dim pa As Variant: pa = GLOBALS_.gFuncParams.GetValue(fname)
        For i = LBound(pa) To UBound(pa)
            If i - LBound(pa) + 1 <= args.count Then
                callScope.SetLocalValue CStr(pa(i)), args(i - LBound(pa) + 1)
            Else
                callScope.SetLocalValue CStr(pa(i)), Empty
            End If
        Next i
    End If

    Dim stmts As Collection: Set stmts = p(1)
    For i = 1 To stmts.count
        Dim ctrl As String: ctrl = ExecuteStmtNode(stmts(i), callScope)
        If ctrl = "RETURN" Then
            CallFuncByIndex_AST = callScope.GetValue("__return")
            callScope.Pop
            Exit Function
        ElseIf ctrl = "ERR" Then
            err.Raise vbObjectError + 3000, "VM.CallFunc", "Error during function execution"
        End If
    Next i
    CallFuncByIndex_AST = Empty
    callScope.Pop
End Function

' ---------------------
' Utilities adapted from your previous VM
' ---------------------
Private Function IsTruthy(v As Variant) As Boolean
    If IsObject(v) Then IsTruthy = Not v Is Nothing: Exit Function
    If IsNull(v) Then IsTruthy = False: Exit Function
    If IsEmpty(v) Then IsTruthy = False: Exit Function
    If VarType(v) = vbBoolean Then IsTruthy = CBool(v): Exit Function
    If IsNumeric(v) Then IsTruthy = (CDbl(v) <> 0): Exit Function
    If VarType(v) = vbString Then
        If IsBoolean(CStr(v)) Then IsTruthy = CBool(v) Else IsTruthy = (CStr(v) <> "")
        Exit Function
    End If
    IsTruthy = True
End Function

Private Function IsBoolean(ByRef expression As String) As Boolean
    IsBoolean = (LCase(expression) = "true")
    If Not IsBoolean Then IsBoolean = (LCase(expression) = "false")
End Function

Private Function SafeAdd(a As Variant, b As Variant) As Variant
    If IsNumeric(a) And IsNumeric(b) Then SafeAdd = a + b Else SafeAdd = CStr(a) & CStr(b)
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
        ValueToStringWithCtx = "": Exit Function
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
        ValueToStringWithCtx = ArrayToString(v, visited, depth)
        Exit Function
    End If

    ' Objects
    If IsObject(v) Then
        Dim tn As String: tn = TypeName(v)
        Select Case tn
            Case "ASF_Map"
                ' cycle detection by object identity
                Dim i As Long
                For i = 1 To visited.count
                    If visited(i) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next i
                visited.Add v
                ValueToStringWithCtx = MapToString(v, visited, depth + 1)
                visited.Remove visited.count
                Exit Function

            Case "Collection"
                ' cycle detection
                Dim j As Long
                For j = 1 To visited.count
                    If visited(j) Is v Then
                        ValueToStringWithCtx = "[Circular]": Exit Function
                    End If
                Next j
                visited.Add v
                ValueToStringWithCtx = CollectionToString(v, visited, depth + 1)
                visited.Remove visited.count
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
    Dim kcnt As Long: kcnt = keys.count

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
            Dim val As Variant: val = m.GetValue(key)
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
    Dim n As Long: n = col.count
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
    On Error GoTo Fallback
    Dim tn As String: tn = TypeName(obj)

    ' Attempt Map-like duck typing: presence of Keys and GetValue
    ' (use On Error to bail out if methods missing)
    Dim dummy As Collection
    Dim tryKeys As Collection
    On Error GoTo Fallback2
    Set tryKeys = obj.keys
    ' if successful, treat As ASF_Map
    Dim i As Long
    For i = 1 To visited.count
        If visited(i) Is obj Then
            ObjectToString = "[Circular]": Exit Function
        End If
    Next i
    visited.Add obj
    ObjectToString = MapToString(obj, visited, depth)
    visited.Remove visited.count
    Exit Function

Fallback2:
    ' Not a Map-like object: try default string
    On Error GoTo Fallback
    ObjectToString = "<" & tn & ">"
    Exit Function

Fallback:
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
             Set keyCol = m.keys
             For Each key In keyCol
                 .VarValue(CStr(key)) = m.GetValue(CStr(key))
             Next key
         Next frame
    
         ' Evaluate
         .Eval
         If .ErrorType = 0 Then
             EvalVBExpressionWithScope = .result
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

 ' ---------------------
 ' Closure / function-value support (runtime helpers)
 ' ---------------------

 ' Create and call closures (closureMap is a Map with keys:
 '  "type" = "Closure", "params" = Collection, "body" = Collection (stmts), "env" = ScopeStack)
 Private Function CallClosure(closureMap As ASF_Map, evaluatedArgs As Collection, Optional thisVal As Variant) As Variant
     On Error GoTo ErrHandler
     Dim env As ASF_ScopeStack
     Set env = closureMap.GetValue("env") ' shared reference (shared-write semantics)

     ' push a new frame for this call
     env.Push
     Dim i As Long
     Dim pa As Variant
     Dim elm As Variant
     Dim fc As Boolean
     
     vAssignment pa, closureMap.GetValue("params")
     If IsObject(pa) Then
        fc = (Not pa Is Nothing)
     Else
        fc = Not IsEmpty(pa)
     End If
     If fc Then
         For Each elm In pa
             i = i + 1
             If i <= evaluatedArgs.count Then
                 env.SetLocalValue CStr(elm), evaluatedArgs(i)
             Else
                 env.SetLocalValue CStr(elm), Empty
             End If
         Next elm
     End If

     ' set 'this' if provided
     If Not IsMissing(thisVal) Then
         env.SetLocalValue "this", thisVal
     End If

     ' execute body
     Dim stmts As Collection: Set stmts = closureMap.GetValue("body")
     Dim ctrl As String
     Dim stmtIdx As Long
     For stmtIdx = 1 To stmts.count
         ctrl = ExecuteStmtNode(stmts(stmtIdx), env)
         If ctrl = "RETURN" Then
             vAssignment CallClosure, env.GetValue("__return")
             env.Pop
             Exit Function
         ElseIf ctrl = "ERR" Then
             err.Raise vbObjectError + 8001, "VM.CallClosure", "Error during closure execution"
         End If
     Next stmtIdx

     ' normal return -> Empty
     env.Pop
     CallClosure = Empty
     Exit Function
ErrHandler:
     If VERBOSE_ Then GLOBALS_.gRuntimeLog.Add "CallClosure error: " & err.Description
     err.Clear
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
            progScope.SetValue CStr(key), rval
        Case "prop"
            If TypeName(container) <> "ASF_Map" Then
                Dim tValue As Variant
                vAssignment tValue, ResolveIndexProp(lval)
                tValue.SetValue CStr(key), rval
            Else
                container.SetValue CStr(key), rval
            End If
        Case "index"
            Dim pos As Long: pos = CLng(key)
            Dim optionBase As Long: optionBase = progScope.GetValue("__option_base")
            If pos < optionBase Then err.Raise vbObjectError + 5001, "VM.HandleAssignment", "Invalid index"
        
            ' If ResolvePath provided parentKey this means the array is stored under a property/variable:
            Dim parentKey As Variant
            If lval.Exists("parentKey") Then vAssignment parentKey, lval("parentKey") Else parentKey = Empty
        
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
        
                ' Ensure arrVal is an array we can index into (AssignToArrayValueInPlace will create/resize).
                Call AssignToArrayValueInPlace(arrVal, pos, rval)
        
                ' Write the updated array back into the parent container
                If TypeName(container) = "ASF_Map" Then
                    container.SetValue CStr(parentKey), arrVal
                ElseIf TypeName(container) = "ASF_ScopeStack" Then
                    container.SetValue CStr(parentKey), arrVal
                Else
                    ' if container itself was an array variable (rare), try to overwrite it:
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
Private Function ResolveIndexProp(ByRef lval) As Variant
    Dim baseVariable As String
    Dim pObject As Variant
    Dim keys As New Collection
    Dim i As Long
    vAssignment pObject, lval("parentObj")
    Do While pObject.Exists("parentObj")
        keys.Add pObject("key")
        vAssignment pObject, pObject("parentObj")
        If Not pObject.Exists("parentKey") Then
            Exit Do
        End If
    Loop
    baseVariable = pObject("key")
    Dim container As Variant: vAssignment container, lval("container")
    vAssignment pObject, container.GetValue(baseVariable)
    For i = keys.count To 2 Step -1
        vAssignment pObject, pObject.GetValue(keys(i))
    Next
    vAssignment ResolveIndexProp, pObject(keys(1))
End Function
 Sub AssignToArrayByName(arrVarName As String, idx As Long, newValue As Variant, progScope As ASF_ScopeStack)
     Dim arrVal As Variant: arrVal = progScope.GetValue(arrVarName)
     If IsEmpty(arrVal) Then
         Dim tmpA() As Variant
         ReDim tmpA(1 To 0)
         arrVal = tmpA
     End If
     Call AssignToArrayValueInPlace(arrVal, idx, newValue)
     progScope.SetValue arrVarName, arrVal
 End Sub

 Sub AssignToArrayValueInPlace(ByRef arrVal As Variant, idx As Long, newValue As Variant)
     If Not IsArray(arrVal) Then
         Dim na() As Variant
         ReDim na(1 To idx)
         na(idx) = newValue
         arrVal = na
         Exit Sub
     End If
     Dim lb As Long: lb = LBound(arrVal)
     Dim ub As Long: ub = UBound(arrVal)
     Dim pos As Long: pos = lb + idx - 1
     If pos < lb Then
         err.Raise vbObjectError + 9120, "VM.AssignToArrayValueInPlace", "Index out of bounds (below LBound)"
     End If
     If pos > ub Then
         ReDim Preserve arrVal(lb To pos)
     End If
     vAssignment arrVal(pos), newValue
 End Sub
'-----------------------------------------------------------------------
' GetElementMapFromIndexNode
' Given an Index AST node, resolve the array container and return the actual
' element Map stored at the index (creating a Map element if empty).
'-----------------------------------------------------------------------
Function GetElementMapFromIndexNode(indexNode As ASF_Map, progScope As ASF_ScopeStack) As ASF_Map
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
    If col.count = 0 Then
        ReDim outArr(ob To ob)
    Else
        ReDim outArr(ob To ob + col.count - 1)
        Dim i As Long
        For i = 1 To col.count
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

 ' ----------------- end LValue helpers -----------------
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
    args.Remove args.count
    vAssignment map_apply_recursive, result
    Exit Function

ErrHandler:
    ' On mapper errors propagate as runtime error so calling code can handle (consistent with earlier semantics)
    err.Raise err.Number, err.Source, err.Description
End Function

' ---------------------
' End VM class
' ---------------------

