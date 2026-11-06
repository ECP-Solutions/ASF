VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Direct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ASF_Direct: prototype of stage 3 (see README.md). Compiles a program from its tokens to the
' bytecode of ASF_Bytecode in one pass, without an AST: a subset of the language (top level code;
' expressions, assignments, if, while, for, return, print). Operands are emitted at once; two
' constants are folded afterwards, and x = x + k, x + k and comparisons in conditions become the
' same superinstructions as ASF_Bytecode makes. Parts the order of emission needs later (the
' target of an assignment, the step of a for, a condition) are token ranges parsed again.
Option Explicit

Private Const OP_PUSHK As Long = 2
Private Const OP_LOADG As Long = 6
Private Const OP_STOREG As Long = 7
Private Const OP_ADD As Long = 8
Private Const OP_SUB As Long = 9
Private Const OP_MUL As Long = 10
Private Const OP_DIV As Long = 11
Private Const OP_MOD As Long = 12
Private Const OP_POW As Long = 13
Private Const OP_EQ As Long = 14
Private Const OP_NE As Long = 15
Private Const OP_LT As Long = 16
Private Const OP_LE As Long = 17
Private Const OP_GT As Long = 18
Private Const OP_GE As Long = 19
Private Const OP_CAT As Long = 20
Private Const OP_SHL As Long = 21
Private Const OP_SHR As Long = 22
Private Const OP_NOT As Long = 23
Private Const OP_NEG As Long = 24
Private Const OP_TYPEOF As Long = 25
Private Const OP_JMP As Long = 26
Private Const OP_JMPF As Long = 27
Private Const OP_JMPT As Long = 28
Private Const OP_ANDJ As Long = 29
Private Const OP_ORJ As Long = 30
Private Const OP_TOBOOL As Long = 31
Private Const OP_NULLJ As Long = 32
Private Const OP_POP As Long = 33
Private Const OP_RET As Long = 35
Private Const OP_RET0 As Long = 36
Private Const OP_PRINT As Long = 37
Private Const OP_INCG As Long = 39
Private Const OP_JNLT As Long = 40
Private Const OP_ADDK As Long = 46
Private Const OP_NEWARR As Long = 55
Private Const OP_NEWOBJ As Long = 56
Private Const OP_GETIDX As Long = 57
Private Const OP_GETIDXW As Long = 58
Private Const OP_SETIDXR As Long = 59
Private Const OP_GETPROP As Long = 60
Private Const OP_SETPROPV As Long = 61
Private Const OP_HALT As Long = 1
Private Const NOT_TAKEN As Long = -2147212480

' tokens, without comments
Private tt_() As String, tv_() As Variant, nT_ As Long
Private p_ As Long, lim_ As Long        ' the parser's position and the end (exclusive) of its range

' code
Private op_() As Long, a_() As Long, b_() As Long, dl_() As Long, nOp_ As Long  ' dl: stack effect
Private k_() As Variant, nK_ As Long, kMap_ As ASF_Map
Private gSlot_ As ASF_Map, gNames_ As Collection
Private parser_ As ASF_Parser
Private reason_ As String

' ---------------------------------------------------------------- compiling
' Compiles src: True when the prototype takes it; Disassembly then gives the code
Public Function Compile(ByVal src As String, aGlobals As ASF_Globals) As Boolean
    On Error GoTo NotTaken
    Set parser_ = New ASF_Parser
    parser_.SetGlobals aGlobals
    LoadTokens parser_.Tokenize(src)
    ReDim op_(1 To 256): ReDim a_(1 To 256): ReDim b_(1 To 256): ReDim dl_(1 To 256): nOp_ = 0
    ReDim k_(1 To 64): nK_ = 0
    Set kMap_ = New ASF_Map
    Set gSlot_ = New ASF_Map
    Set gNames_ = New Collection
    ScanGlobals
    StatementsIn 1, nT_ + 1
    Emit OP_HALT, 0, 0, 0
    Compile = True
    Exit Function
NotTaken:
    reason_ = Err.Description
    Compile = False
End Function

Public Property Get Reason() As String
    Reason = reason_
End Property

Private Sub LoadTokens(toks As Collection)
    Dim t As Variant, n As Long
    ReDim tt_(1 To toks.Count + 1): ReDim tv_(1 To toks.Count + 1)
    For Each t In toks
        If t(0) <> "COMMENT" Then
            n = n + 1
            tt_(n) = t(0)
            tv_(n) = t(1)
        End If
    Next t
    nT_ = n
End Sub

Private Sub NotSupported(ByVal what As String)
    Err.Raise NOT_TAKEN, "ASF_Direct", what
End Sub

' The names the program assigns, in order: its first global slots, as ASF_Bytecode numbers them
Private Sub ScanGlobals()
    Dim i As Long, j As Long, nm As String
    For i = 2 To nT_
        If tt_(i) = "OP" Then
            Select Case tv_(i)
                Case "=", "+=", "-=", "*=", "/="
                    j = TargetStart(i)
                    If j > 0 Then
                        If StatementStart(j) Then
                            nm = RootName(CStr(tv_(j)))
                            If Len(nm) > 0 Then GlobalSlot nm
                        End If
                    End If
            End Select
        End If
    Next i
End Sub

' The first token of the target of the assignment operator at i: an identifier followed by
' indexes only (a property assignment assigns no variable); 0 when none
Private Function TargetStart(ByVal i As Long) As Long
    Dim j As Long: j = i - 1
    Do While j >= 1
        If tt_(j) = "PAREN" And tv_(j) = "]" Then
            Dim d As Long: d = 0
            Do While j >= 1
                If tt_(j) = "PAREN" Then
                    If tv_(j) = "]" Then d = d + 1
                    If tv_(j) = "[" Then d = d - 1
                End If
                If d = 0 Then Exit Do
                j = j - 1
            Loop
            j = j - 1
        ElseIf tt_(j) = "IDENT" Then
            TargetStart = j
            Exit Function
        Else
            Exit Function
        End If
    Loop
End Function

' True when token j begins a statement (or a part of a for header)
Private Function StatementStart(ByVal j As Long) As Boolean
    If j = 1 Then StatementStart = True: Exit Function
    Select Case tt_(j - 1)
        Case "SEP": StatementStart = (tv_(j - 1) = ";" Or tv_(j - 1) = ",")
        Case "PAREN"
            If tv_(j - 1) = "{" Or tv_(j - 1) = "}" Then
                StatementStart = True
            ElseIf tv_(j - 1) = "(" And j > 2 Then
                If tt_(j - 2) = "IDENT" Then StatementStart = (LCase$(tv_(j - 2)) = "for")
            End If
    End Select
End Function

Private Function RootName(ByVal s As String) As String
    Dim k As Long: k = InStr(s, "[")
    If k > 0 Then RootName = Left$(s, k - 1) Else RootName = s
    If InStr(RootName, ".") > 0 Then RootName = vbNullString
End Function

Private Function GlobalSlot(ByVal name As String) As Long
    If gSlot_.Exists(name) Then
        GlobalSlot = gSlot_.GetValue(name)
    Else
        gNames_.Add name
        GlobalSlot = gNames_.Count
        gSlot_.SetValue name, GlobalSlot
    End If
End Function

' ---------------------------------------------------------------- emitting
Private Sub Emit(ByVal op As Long, ByVal a As Long, ByVal b As Long, ByVal stackDelta As Long)
    nOp_ = nOp_ + 1
    If nOp_ > UBound(op_) Then
        ReDim Preserve op_(1 To 2 * nOp_): ReDim Preserve a_(1 To 2 * nOp_): ReDim Preserve b_(1 To 2 * nOp_)
        ReDim Preserve dl_(1 To 2 * nOp_)
    End If
    op_(nOp_) = op: a_(nOp_) = a: b_(nOp_) = b: dl_(nOp_) = stackDelta
End Sub

Private Function ConstKey(ByRef v As Variant) As String
    If IsEmpty(v) Then
        ConstKey = "E"
    ElseIf IsNull(v) Then
        ConstKey = "N"
    Else
        ConstKey = VarType(v) & ":" & CStr(v)
    End If
End Function

Private Function Constant(ByRef v As Variant) As Long
    Dim key As String: key = ConstKey(v)
    If kMap_.Exists(key) Then
        Constant = kMap_.GetValue(key)
        Exit Function
    End If
    nK_ = nK_ + 1
    If nK_ > UBound(k_) Then ReDim Preserve k_(1 To 2 * nK_)
    k_(nK_) = v
    kMap_.SetValue key, nK_
    Constant = nK_
End Function

' Drops the last instruction, a PUSHK: and its constant, when it was the last one added
Private Sub DropPush()
    Dim k As Long: k = a_(nOp_)
    If k = nK_ And Not KUsedBefore(k, nOp_) Then
        kMap_.Remove ConstKey(k_(nK_))
        k_(nK_) = Empty
        nK_ = nK_ - 1
    End If
    nOp_ = nOp_ - 1
End Sub

' True when an instruction before at refers to constant k
Private Function KUsedBefore(ByVal k As Long, ByVal at As Long) As Boolean
    Dim i As Long
    For i = at - 1 To 1 Step -1
        Select Case op_(i)
            Case OP_PUSHK, OP_ADDK, OP_GETPROP, OP_SETPROPV
                If a_(i) = k Then KUsedBefore = True: Exit Function
            Case OP_INCG
                If b_(i) = k Then KUsedBefore = True: Exit Function
        End Select
    Next i
End Function

Private Sub Patch(ByVal at As Long, ByVal target As Long)
    a_(at) = target
End Sub

Private Sub PatchAll(jumps As Collection, ByVal target As Long)
    Dim j As Variant
    For Each j In jumps
        a_(j) = target
    Next j
End Sub

' ---------------------------------------------------------------- statements
' The statements between token from and token upTo (exclusive)
Private Sub StatementsIn(ByVal from As Long, ByVal upTo As Long)
    Dim i As Long: i = from
    Do While i < upTo
        i = Statement(i, upTo)
    Loop
End Sub

' One statement from token i; the token after it
Private Function Statement(ByVal i As Long, ByVal upTo As Long) As Long
    If tt_(i) = "SEP" And tv_(i) = ";" Then
        Statement = i + 1
        Exit Function
    End If
    Dim e As Long
    If tt_(i) = "IDENT" Then
        Select Case LCase$(tv_(i))
            Case "if": Statement = IfStatement(i, upTo): Exit Function
            Case "while": Statement = WhileStatement(i): Exit Function
            Case "for": Statement = ForStatement(i): Exit Function
            Case "return"
                e = StatementEnd(i + 1, upTo)
                If e = i + 1 Then
                    Emit OP_RET0, 0, 0, 0
                Else
                    Expr i + 1, e
                    Emit OP_RET, 0, 0, -1
                End If
                Statement = e
                Exit Function
            Case "print"
                Statement = PrintStatement(i, upTo)
                Exit Function
            Case "break", "continue", "fun", "class", "switch", "try", "throw", "import", "export", "prototype"
                NotSupported "statement " & tv_(i)
        End Select
    End If
    e = StatementEnd(i, upTo)
    SimpleStatement i, e
    Statement = e
End Function

' The end of a simple statement: the ; at depth 0, or upTo
Private Function StatementEnd(ByVal i As Long, ByVal upTo As Long) As Long
    Dim d As Long
    Do While i < upTo
        If tt_(i) = "PAREN" Then
            Select Case tv_(i)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf tt_(i) = "SEP" And tv_(i) = ";" And d = 0 Then
            Exit Do
        End If
        i = i + 1
    Loop
    StatementEnd = i
End Function

' An assignment (x = v, x op= v, o.p = v, a[i] = v) or an expression
Private Sub SimpleStatement(ByVal i As Long, ByVal e As Long)
    Dim j As Long, d As Long
    For j = i To e - 1
        If tt_(j) = "PAREN" Then
            Select Case tv_(j)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf tt_(j) = "OP" And d = 0 Then
            Select Case tv_(j)
                Case "=", "+=", "-=", "*=", "/="
                    Assignment i, j, e
                    Exit Sub
            End Select
        End If
    Next j
    Expr i, e
    Emit OP_POP, 0, 0, -1
End Sub

' target (tokens i..eq-1) op value (eq+1..e-1)
Private Sub Assignment(ByVal i As Long, ByVal eq As Long, ByVal e As Long)
    Dim op As String: op = tv_(eq)
    Dim start As Long: start = nOp_
    If op = "=" Then
        Expr eq + 1, e
    Else
        ' x op= v is x = x op v
        Expr i, eq
        Dim lEnd As Long: lEnd = nOp_
        Expr eq + 1, e
        EmitBinaryOp Left$(op, 1), start + 1, lEnd
    End If
    ' x = x + constant: one instruction
    If nOp_ - start = 2 And eq = i + 1 And tt_(i) = "IDENT" Then
        If op_(start + 1) = OP_LOADG And op_(start + 2) = OP_ADDK Then
            If InStr(CStr(tv_(i)), "[") = 0 Then
                If a_(start + 1) = GlobalSlot(CStr(tv_(i))) Then
                    op_(start + 1) = OP_INCG
                    b_(start + 1) = a_(start + 2)
                    dl_(start + 1) = 0
                    nOp_ = start + 1
                    Exit Sub
                End If
            End If
        End If
    End If
    StoreTarget i, eq
End Sub

' Stores the value on top in the target written in tokens i..e-1: x, x.p, x[i], and chains of them
Private Sub StoreTarget(ByVal i As Long, ByVal e As Long)
    Dim root As String, segs As Collection
    TargetSegments i, e, root, segs
    StoreSegments root, segs, segs.Count
End Sub

' root and its segments: Array("P", name) or Array("I", from, upTo)
Private Sub TargetSegments(ByVal i As Long, ByVal e As Long, ByRef root As String, ByRef segs As Collection)
    Set segs = New Collection
    If tt_(i) <> "IDENT" Then NotSupported "assignment target"
    Dim s As String: s = tv_(i)
    Dim k As Long, j As Long, from As Long, upTo As Long, c As Long
    k = InStr(s, "[")
    If k = 0 Then
        root = s
    Else
        Dim idx As Collection
        SplitCollapsed s, root, idx
        For j = 1 To idx.Count
            AppendTokens CStr(idx(j)), from, upTo
            segs.Add Array("I", from, upTo)
        Next j
    End If
    j = i + 1
    Do While j < e
        If tt_(j) = "SYM" And tv_(j) = "." And j + 1 < e Then
            segs.Add Array("P", CStr(tv_(j + 1)))
            j = j + 2
        ElseIf tt_(j) = "PAREN" And tv_(j) = "[" Then
            c = Matching(j)
            segs.Add Array("I", j + 1, c)
            j = c + 1
        Else
            NotSupported "assignment target"
        End If
    Loop
End Sub

' As ASF_Bytecode.EmitStoreTop: obj.p = v sets the property; c[i] = v writes into the container and
' stores the container back where it came from
Private Sub StoreSegments(ByVal root As String, segs As Collection, ByVal n As Long)
    If n = 0 Then
        Emit OP_STOREG, GlobalSlot(root), 0, -1
        Exit Sub
    End If
    Dim seg As Variant: seg = segs(n)
    If seg(0) = "P" Then
        LoadSegments root, segs, n - 1, False
        Emit OP_SETPROPV, Constant(CStr(seg(1))), 0, -2
    Else
        LoadSegments root, segs, n - 1, True
        Expr seg(1), seg(2)
        Emit OP_SETIDXR, 0, 0, -2
        StoreSegments root, segs, n - 1
    End If
End Sub

' The value of root and its first n segments; forWrite: as a container about to be written, whose
' indexed parts read with GETIDXW (EmitContainerForWrite)
Private Sub LoadSegments(ByVal root As String, segs As Collection, ByVal n As Long, ByVal forWrite As Boolean)
    If n = 0 Then
        Emit OP_LOADG, GlobalSlot(root), 0, 1
        Exit Sub
    End If
    Dim seg As Variant: seg = segs(n)
    If forWrite And seg(0) = "I" Then
        LoadSegments root, segs, n - 1, True
        Expr seg(1), seg(2)
        Emit OP_GETIDXW, 0, 0, -1
    Else
        LoadSegments root, segs, n - 1, False
        If seg(0) = "P" Then
            Emit OP_GETPROP, Constant(CStr(seg(1))), 0, 0
        Else
            Expr seg(1), seg(2)
            Emit OP_GETIDX, 0, 0, -1
        End If
    End If
End Sub

Private Sub SplitCollapsed(ByVal s As String, ByRef name As String, ByRef idx As Collection)
    Set idx = New Collection
    Dim k As Long: k = InStr(s, "[")
    name = Left$(s, k - 1)
    Dim d As Long, st As Long, i As Long, ch As String
    For i = k To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "[" Then
            If d = 0 Then st = i + 1
            d = d + 1
        ElseIf ch = "]" Then
            d = d - 1
            If d = 0 Then idx.Add Mid$(s, st, i - st)
        ElseIf d = 0 Then
            NotSupported "collapsed member"
        End If
    Next i
End Sub

' The tokens of text (an index of a collapsed identifier, a[1]) go after the program's: from and
' upTo give their range
Private Sub AppendTokens(ByVal text As String, ByRef from As Long, ByRef upTo As Long)
    Dim toks As Collection: Set toks = parser_.Tokenize(text)
    If nT_ + toks.Count + 1 > UBound(tt_) Then
        ReDim Preserve tt_(1 To 2 * (nT_ + toks.Count + 1)): ReDim Preserve tv_(1 To 2 * (nT_ + toks.Count + 1))
    End If
    from = nT_ + 1
    Dim t As Variant
    For Each t In toks
        If t(0) <> "COMMENT" Then
            nT_ = nT_ + 1
            tt_(nT_) = t(0)
            tv_(nT_) = t(1)
        End If
    Next t
    upTo = nT_ + 1
    Do While upTo > from
        If Not (tt_(upTo - 1) = "SEP" And tv_(upTo - 1) = ";") Then Exit Do
        upTo = upTo - 1
    Loop
End Sub

Private Function PrintStatement(ByVal i As Long, ByVal upTo As Long) As Long
    Dim e As Long: e = StatementEnd(i, upTo)
    If Not (tt_(i + 1) = "PAREN" And tv_(i + 1) = "(") Then NotSupported "print"
    Dim close_ As Long: close_ = Matching(i + 1)
    Dim n As Long, a As Long, j As Long, d As Long
    a = i + 2
    For j = i + 2 To close_
        If j = close_ Or (tt_(j) = "SEP" And tv_(j) = "," And d = 0) Then
            If j > a Then
                Expr a, j
                n = n + 1
            End If
            a = j + 1
        ElseIf tt_(j) = "PAREN" Then
            Select Case tv_(j)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        End If
    Next j
    Emit OP_PRINT, n, 0, -n
    PrintStatement = e
End Function

' The token that closes the bracket at i
Private Function Matching(ByVal i As Long) As Long
    Dim d As Long
    Do
        If tt_(i) = "PAREN" Then
            Select Case tv_(i)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
            If d = 0 Then Matching = i: Exit Function
        End If
        i = i + 1
        If i > nT_ Then NotSupported "unbalanced"
    Loop
End Function

' if (c) { } elseif (c) { } else { }
Private Function IfStatement(ByVal i As Long, ByVal upTo As Long) As Long
    Dim ends As Collection: Set ends = New Collection
    Dim falses As Collection, c As Long, b As Long
    c = Matching(i + 1)
    Set falses = New Collection
    JumpIfFalse i + 2, c, falses
    b = Matching(c + 1)
    StatementsIn c + 2, b
    Emit OP_JMP, 0, 0, 0: ends.Add nOp_
    PatchAll falses, nOp_ + 1
    i = b + 1
    Do While i < upTo
        If tt_(i) <> "IDENT" Then Exit Do
        If LCase$(tv_(i)) <> "elseif" Then Exit Do
        c = Matching(i + 1)
        Set falses = New Collection
        JumpIfFalse i + 2, c, falses
        b = Matching(c + 1)
        StatementsIn c + 2, b
        Emit OP_JMP, 0, 0, 0: ends.Add nOp_
        PatchAll falses, nOp_ + 1
        i = b + 1
    Loop
    If i < upTo Then
        If tt_(i) = "IDENT" Then
            If LCase$(tv_(i)) = "else" Then
                b = Matching(i + 1)
                StatementsIn i + 2, b
                i = b + 1
            End If
        End If
    End If
    PatchAll ends, nOp_ + 1
    IfStatement = i
End Function

Private Function WhileStatement(ByVal i As Long) As Long
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    Dim c As Long: c = Matching(i + 1)
    JumpIfFalse i + 2, c, exits
    Dim b As Long: b = Matching(c + 1)
    StatementsIn c + 2, b
    Emit OP_JMP, loopTop, 0, 0
    PatchAll exits, nOp_ + 1
    WhileStatement = b + 1
End Function

' for (init, cond, step) { }: the step is emitted after the body
Private Function ForStatement(ByVal i As Long) As Long
    Dim c As Long: c = Matching(i + 1)
    Dim parts(1 To 4) As Long, n As Long, j As Long, d As Long
    parts(1) = i + 2: n = 1
    For j = i + 2 To c - 1
        If tt_(j) = "PAREN" Then
            Select Case tv_(j)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf tt_(j) = "SEP" And d = 0 And (tv_(j) = "," Or tv_(j) = ";") Then
            n = n + 1
            If n > 3 Then NotSupported "for"
            parts(n) = j + 1
        End If
    Next j
    If n <> 3 Then NotSupported "for-of/in"
    parts(4) = c + 1
    If parts(2) - 1 > parts(1) Then SimpleStatement parts(1), parts(2) - 1
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    If parts(3) - 1 > parts(2) Then JumpIfFalse parts(2), parts(3) - 1, exits
    Dim b As Long: b = Matching(c + 1)
    StatementsIn c + 2, b
    If parts(4) - 1 > parts(3) Then SimpleStatement parts(3), parts(4) - 1
    Emit OP_JMP, loopTop, 0, 0
    PatchAll exits, nOp_ + 1
    ForStatement = b + 1
End Function

' ---------------------------------------------------------------- conditions
' The expression of tokens i..e-1 as a condition: jumps get the jumps taken when it is false
Private Sub JumpIfFalse(ByVal i As Long, ByVal e As Long, jumps As Collection)
    Dim at As Long, op As String
    Unwrap i, e
    op = TopOperator(i, e, at)
    Select Case op
        Case "&&"
            JumpIfFalse i, at, jumps
            JumpIfFalse at + 1, e, jumps
        Case "||"
            Dim trues As Collection: Set trues = New Collection
            JumpIfTrue i, at, trues
            JumpIfFalse at + 1, e, jumps
            PatchAll trues, nOp_ + 1
        Case "!"
            JumpIfTrue i + 1, e, jumps
        Case "<", "<=", ">", ">=", "==", "=", "!="
            Expr i, e
            op_(nOp_) = CompareJump(op)
            dl_(nOp_) = -2
            jumps.Add nOp_
        Case Else
            Expr i, e
            Emit OP_JMPF, 0, 0, -1
            jumps.Add nOp_
    End Select
End Sub

Private Sub JumpIfTrue(ByVal i As Long, ByVal e As Long, jumps As Collection)
    Dim at As Long, op As String
    Unwrap i, e
    op = TopOperator(i, e, at)
    Select Case op
        Case "||"
            JumpIfTrue i, at, jumps
            JumpIfTrue at + 1, e, jumps
        Case "&&"
            Dim falses As Collection: Set falses = New Collection
            JumpIfFalse i, at, falses
            JumpIfTrue at + 1, e, jumps
            PatchAll falses, nOp_ + 1
        Case "!"
            JumpIfFalse i + 1, e, jumps
        Case Else
            Expr i, e
            Emit OP_JMPT, 0, 0, -1
            jumps.Add nOp_
    End Select
End Sub

' (x): the AST has no node for parentheses
Private Sub Unwrap(ByRef i As Long, ByRef e As Long)
    Do While e - i >= 2
        If Not (tt_(i) = "PAREN" And tv_(i) = "(") Then Exit Do
        If Matching(i) <> e - 1 Then Exit Do
        i = i + 1: e = e - 1
    Loop
End Sub

Private Function CompareJump(ByVal op As String) As Long
    Select Case op
        Case "<": CompareJump = OP_JNLT
        Case "<=": CompareJump = OP_JNLT + 1
        Case ">": CompareJump = OP_JNLT + 2
        Case ">=": CompareJump = OP_JNLT + 3
        Case "==", "=": CompareJump = OP_JNLT + 4
        Case "!=": CompareJump = OP_JNLT + 5
    End Select
End Function

' The root of the expression in tokens i..e-1, as the parser builds it: the binary operator of
' lowest precedence at depth 0 (the last of them, as they group to the left; the first for ??),
' "!" for a ! over the whole expression, "?" for a conditional, "" otherwise
Private Function TopOperator(ByVal i As Long, ByVal e As Long, ByRef at As Long) As String
    Dim j As Long, d As Long, best As Long, bestPrec As Long, prec As Long, prevOperand As Boolean
    bestPrec = 99
    For j = i To e - 1
        If tt_(j) = "PAREN" Then
            Select Case tv_(j)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
            prevOperand = True
        ElseIf tt_(j) = "OP" And d = 0 Then
            If tv_(j) = "?" Then
                TopOperator = "?"
                Exit Function
            End If
            If prevOperand Then prec = BinaryPrec(CStr(tv_(j))) Else prec = 0
            If prec > 0 Then
                If prec < bestPrec Or (prec = bestPrec And prec <> 1) Then
                    bestPrec = prec
                    best = j
                End If
            End If
            prevOperand = False
        ElseIf tt_(j) = "SYM" Then
            prevOperand = False
        Else
            prevOperand = True
        End If
    Next j
    If best > 0 Then
        at = best
        TopOperator = tv_(best)
    ElseIf tt_(i) = "OP" And tv_(i) = "!" Then
        TopOperator = "!"
    End If
End Function

' ---------------------------------------------------------------- expressions
' Emits the expression of tokens i..e-1
Private Sub Expr(ByVal i As Long, ByVal e As Long)
    Dim saveP As Long, saveL As Long
    saveP = p_: saveL = lim_
    p_ = i: lim_ = e
    TernaryExpr
    If p_ < lim_ Then NotSupported "tokens after an expression"
    p_ = saveP: lim_ = saveL
End Sub

Private Function Peek(ByVal kind As String, ByVal value As String) As Boolean
    If p_ < lim_ Then Peek = (tt_(p_) = kind And tv_(p_) = value)
End Function

' cond ? a : b: the condition is parsed again as a jump, as ASF_Bytecode emits it
Private Sub TernaryExpr()
    Dim start As Long: start = p_
    Dim mark As Long: mark = nOp_
    Dim saveK As Long, saveG As Long
    saveK = nK_: saveG = gNames_.Count
    BinaryExpr 1
    If Not Peek("OP", "?") Then Exit Sub
    Dim q As Long: q = p_
    Undo mark, saveK, saveG
    Dim falses As Collection: Set falses = New Collection
    Dim saveL As Long: saveL = lim_
    JumpIfFalse start, q, falses
    p_ = q + 1: lim_ = saveL
    TernaryExpr
    If Not Peek("OP", ":") Then NotSupported "ternary"
    p_ = p_ + 1
    Emit OP_JMP, 0, 0, -1
    Dim endJump As Long: endJump = nOp_
    PatchAll falses, nOp_ + 1
    TernaryExpr
    Patch endJump, nOp_ + 1
End Sub

' Drops the code emitted since mark, its new constants and its new globals
Private Sub Undo(ByVal mark As Long, ByVal saveK As Long, ByVal saveG As Long)
    Do While nK_ > saveK
        kMap_.Remove ConstKey(k_(nK_))
        k_(nK_) = Empty
        nK_ = nK_ - 1
    Loop
    Do While gNames_.Count > saveG
        gSlot_.Remove CStr(gNames_(gNames_.Count))
        gNames_.Remove gNames_.Count
    Loop
    nOp_ = mark
End Sub

Private Function BinaryPrec(ByVal op As String) As Long
    Select Case op
        Case "??": BinaryPrec = 1
        Case "||": BinaryPrec = 2
        Case "&&": BinaryPrec = 3
        Case "==", "=", "!=": BinaryPrec = 4
        Case "<", ">", "<=", ">=": BinaryPrec = 5
        Case "<<", ">>": BinaryPrec = 6
        Case "+", "-", "&": BinaryPrec = 7
        Case "*", "/", "%": BinaryPrec = 8
        Case "^": BinaryPrec = 9
    End Select
End Function

' Precedence climbing as ASF_Compiler.ParseBinaryNode; each operand is emitted when parsed
Private Sub BinaryExpr(ByVal minPrec As Long)
    Dim lStart As Long, rStart As Long, op As String, prec As Long, endJump As Long
    lStart = nOp_ + 1
    UnaryExpr
    Do While p_ < lim_
        If tt_(p_) <> "OP" Then Exit Do
        op = tv_(p_)
        prec = BinaryPrec(op)
        If prec = 0 Or prec < minPrec Then Exit Do
        p_ = p_ + 1
        Select Case op
            Case "&&", "||"
                If op = "&&" Then Emit OP_ANDJ, 0, 0, -1 Else Emit OP_ORJ, 0, 0, -1
                endJump = nOp_
                BinaryExpr prec + 1
                Emit OP_TOBOOL, 0, 0, 0
                Patch endJump, nOp_ + 1
            Case "??"
                Emit OP_NULLJ, 0, 0, -1
                endJump = nOp_
                BinaryExpr prec
                Patch endJump, nOp_ + 1
            Case Else
                rStart = nOp_ + 1
                If prec = 9 Then BinaryExpr prec Else BinaryExpr prec + 1
                EmitBinaryOp op, lStart, rStart - 1
        End Select
    Loop
End Sub

' The operator over the left operand (lStart..lEnd) and the right one (lEnd+1..nOp_): two constants
' fold into one, x + constant is ADDK
Private Sub EmitBinaryOp(ByVal op As String, ByVal lStart As Long, ByVal lEnd As Long)
    Dim l As Variant, r As Variant, k As Variant, folded As Boolean
    If lEnd = lStart And nOp_ = lEnd + 1 Then
        If op_(lStart) = OP_PUSHK And op_(nOp_) = OP_PUSHK Then
            l = k_(a_(lStart)): r = k_(a_(nOp_))
            If VarType(l) = vbDouble And VarType(r) = vbDouble Then
                Select Case op
                    Case "+": k = l + r: folded = True
                    Case "-": k = l - r: folded = True
                    Case "*": k = l * r: folded = True
                    Case "/"
                        If r <> 0 Then k = l / r: folded = True
                End Select
            ElseIf VarType(l) = vbString And VarType(r) = vbString Then
                If op = "+" Then k = l & r: folded = True
            End If
            If folded Then
                DropPush
                DropPush
                Emit OP_PUSHK, Constant(k), 0, 1
                Exit Sub
            End If
        End If
    End If
    Dim code As Long: code = BinaryCode(op)
    If code = OP_ADD And nOp_ = lEnd + 1 Then
        If op_(nOp_) = OP_PUSHK Then
            If VarType(k_(a_(nOp_))) = vbDouble Then
                op_(nOp_) = OP_ADDK
                dl_(nOp_) = 0
                Exit Sub
            End If
        End If
    End If
    Emit code, 0, 0, -1
End Sub

Private Function BinaryCode(ByVal op As String) As Long
    Select Case op
        Case "+": BinaryCode = OP_ADD
        Case "-": BinaryCode = OP_SUB
        Case "*": BinaryCode = OP_MUL
        Case "/": BinaryCode = OP_DIV
        Case "%": BinaryCode = OP_MOD
        Case "^": BinaryCode = OP_POW
        Case "==", "=": BinaryCode = OP_EQ
        Case "!=": BinaryCode = OP_NE
        Case "<": BinaryCode = OP_LT
        Case "<=": BinaryCode = OP_LE
        Case ">": BinaryCode = OP_GT
        Case ">=": BinaryCode = OP_GE
        Case "&": BinaryCode = OP_CAT
        Case "<<": BinaryCode = OP_SHL
        Case ">>": BinaryCode = OP_SHR
        Case Else: NotSupported "operator " & op
    End Select
End Function

Private Sub UnaryExpr()
    If p_ < lim_ Then
        If tt_(p_) = "OP" And (tv_(p_) = "!" Or tv_(p_) = "-") Then
            Dim op As String: op = tv_(p_)
            p_ = p_ + 1
            Dim start As Long: start = nOp_ + 1
            UnaryExpr
            If op = "-" And nOp_ = start And op_(nOp_) = OP_PUSHK Then
                If VarType(k_(a_(nOp_))) = vbDouble Then
                    Dim v As Variant: v = -k_(a_(nOp_))
                    DropPush
                    Emit OP_PUSHK, Constant(v), 0, 1
                    Exit Sub
                End If
            End If
            If op = "!" Then Emit OP_NOT, 0, 0, 0 Else Emit OP_NEG, 0, 0, 0
            Exit Sub
        End If
        If tt_(p_) = "IDENT" And LCase$(tv_(p_)) = "typeof" Then
            p_ = p_ + 1
            UnaryExpr
            Emit OP_TYPEOF, 0, 0, 0
            Exit Sub
        End If
    End If
    PrimaryExpr
End Sub

Private Sub PrimaryExpr()
    If p_ >= lim_ Then
        Emit OP_PUSHK, Constant(Empty), 0, 1
        Exit Sub
    End If
    Dim t As String: t = tt_(p_)
    Dim v As Variant: v = tv_(p_)
    Select Case t
        Case "NUMBER"
            Emit OP_PUSHK, Constant(CDbl(v)), 0, 1
            p_ = p_ + 1
        Case "STRING"
            If Left$(v, 1) = "`" Then NotSupported "template"
            Emit OP_PUSHK, Constant(CStr(v)), 0, 1
            p_ = p_ + 1
        Case "IDENT"
            p_ = p_ + 1
            Select Case LCase$(v)
                Case "true": Emit OP_PUSHK, Constant(True), 0, 1
                Case "false": Emit OP_PUSHK, Constant(False), 0, 1
                Case "null": Emit OP_PUSHK, Constant(Null), 0, 1
                Case "undefined": Emit OP_PUSHK, Constant(Empty), 0, 1
                Case Else
                    If InStr(v, "[") > 0 Then
                        Dim name As String, idx As Collection, k As Long, from As Long, upTo As Long
                        SplitCollapsed CStr(v), name, idx
                        Emit OP_LOADG, GlobalSlot(name), 0, 1
                        For k = 1 To idx.Count
                            AppendTokens CStr(idx(k)), from, upTo
                            Expr from, upTo
                            Emit OP_GETIDX, 0, 0, -1
                        Next k
                    Else
                        If InStr(v, ".") > 0 Then NotSupported "collapsed member"
                        Emit OP_LOADG, GlobalSlot(CStr(v)), 0, 1
                    End If
            End Select
        Case "PAREN"
            Select Case v
                Case "("
                    Dim c As Long: c = Matching(p_)
                    Dim saveL As Long: saveL = lim_
                    p_ = p_ + 1: lim_ = c
                    TernaryExpr
                    lim_ = saveL: p_ = c + 1
                Case "["
                    ArrayLiteral
                Case "{"
                    ObjectLiteral
                Case Else
                    NotSupported "bracket"
            End Select
        Case Else
            NotSupported "token " & t
    End Select
    Postfix
End Sub

Private Sub Postfix()
    Do While p_ < lim_
        If tt_(p_) = "SYM" And tv_(p_) = "." Then
            If p_ + 1 >= lim_ Then NotSupported "member"
            If p_ + 2 < lim_ Then
                If tt_(p_ + 2) = "PAREN" And tv_(p_ + 2) = "(" Then NotSupported "method call"
            End If
            Emit OP_GETPROP, Constant(CStr(tv_(p_ + 1))), 0, 0
            p_ = p_ + 2
        ElseIf tt_(p_) = "PAREN" And tv_(p_) = "[" Then
            Dim c As Long: c = Matching(p_)
            Expr p_ + 1, c
            Emit OP_GETIDX, 0, 0, -1
            p_ = c + 1
        ElseIf tt_(p_) = "PAREN" And tv_(p_) = "(" Then
            NotSupported "call"
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ArrayLiteral()
    Dim c As Long: c = Matching(p_)
    Dim n As Long, a As Long, j As Long, d As Long
    a = p_ + 1
    For j = p_ + 1 To c
        If j = c Or (tt_(j) = "SEP" And tv_(j) = "," And d = 0) Then
            If j > a Then
                Expr a, j
                n = n + 1
            End If
            a = j + 1
        ElseIf tt_(j) = "PAREN" Then
            Select Case tv_(j)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf tt_(j) = "SPREAD" Then
            NotSupported "spread"
        End If
    Next j
    Emit OP_NEWARR, n, 0, 1 - n
    p_ = c + 1
End Sub

Private Sub ObjectLiteral()
    Dim c As Long: c = Matching(p_)
    Dim n As Long, j As Long
    j = p_ + 1
    Do While j < c
        If Not (tt_(j) = "IDENT" Or tt_(j) = "STRING") Then NotSupported "object key"
        If Not (tt_(j + 1) = "OP" And tv_(j + 1) = ":") Then NotSupported "object entry"
        Emit OP_PUSHK, Constant(CStr(tv_(j))), 0, 1
        Dim e As Long, d As Long: e = j + 2: d = 0
        Do While e < c
            If tt_(e) = "PAREN" Then
                Select Case tv_(e)
                    Case "(", "[", "{": d = d + 1
                    Case ")", "]", "}": d = d - 1
                End Select
            ElseIf tt_(e) = "SEP" And tv_(e) = "," And d = 0 Then
                Exit Do
            End If
            e = e + 1
        Loop
        Expr j + 2, e
        n = n + 1
        j = e + 1
    Loop
    Emit OP_NEWOBJ, n, 0, 1 - 2 * n
    p_ = c + 1
End Sub

' ---------------------------------------------------------------- output
' The code as ASF_Bytecode.Disassemble writes it, for a program of one function
Public Function Disassembly() As String
    Dim names As Variant
    names = Split("HALT PUSHK PUSHE LOADL STOREL LOADG STOREG ADD SUB MUL DIV MOD POW EQ NE LT LE GT GE CAT SHL SHR " & _
                  "NOT NEG TYPEOF JMP JMPF JMPT ANDJ ORJ TOBOOL NULLJ POP CALLF RET RET0 PRINT INCL INCG " & _
                  "JNLT JNLE JNGT JNGE JNEQ JNNE ADDK DUP LOADC STOREC LOADE STOREE MKCLOS MKFUNC CALL NEWARR NEWOBJ " & _
                  "GETIDX GETIDXW SETIDXR GETPROP SETPROPV", " ")
    Dim out As String, i As Long, depth As Long, maxDepth As Long
    For i = 1 To nOp_
        depth = depth + dl_(i)
        If depth > maxDepth Then maxDepth = depth
    Next i
    out = "fn 1 @main entry 1 params 0 locals 0 stack " & (maxDepth + 2) & " cells 0 captures 0" & vbLf
    For i = 1 To nK_
        out = out & "k" & i & " = " & TypeName(k_(i))
        If Not IsNull(k_(i)) Then out = out & " " & CStr(k_(i))
        out = out & vbLf
    Next i
    For i = 1 To nOp_
        out = out & i & vbTab & names(op_(i) - 1) & " " & a_(i) & " " & b_(i)
        Select Case op_(i)
            Case OP_PUSHK, OP_ADDK, OP_GETPROP, OP_SETPROPV
                out = out & "  ; " & TypeName(k_(a_(i)))
                If Not IsNull(k_(a_(i))) Then out = out & " " & CStr(k_(a_(i)))
        End Select
        out = out & vbLf
    Next i
    Disassembly = out
End Function

Public Property Get InstructionCount() As Long
    InstructionCount = nOp_
End Property
