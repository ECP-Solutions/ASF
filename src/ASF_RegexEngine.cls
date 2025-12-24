VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_RegexEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ASF_RegexEngine.cls - Full-featured regex engine for ASF (single class module)
' Features implemented:
'  - Precompile pattern into an AST (no reparsing on Exec)
'  - Literals, . (any char)
'  - Character classes: [abc], [a-z], negation [^...]
'  - Escapes: \d \w \s and common escapes (\n, \r, \t)
'  - Quantifiers: *, +, ?, {m,n} with {m,} and exact {n}
'  - Greedy and lazy quantifiers (use trailing ? after quantifier to mark lazy)
'  - Grouping: ( ... ) with precise capture start/end tracking
'  - Alternation: | via `alt` AST node
'  - Anchors: ^ and $
'  - IgnoreCase (i) flag
'  - DotAll (s) controlling whether '.' matches newlines
'  - Multiline (m) flag
'  - Public API: Init, Test, Exec, Replace, Split
'
' Design notes:
'  - AST nodes are represented as VBA Collections with string keys; helper
'    accessors (NodeAdd/NodeGet/NodeExists) are provided.
'  - Matching is backtracking-based; defensive limits and optimizations are
'    implemented to reduce pathological backtracking where practical.
'
' Notes & Limitations:
'     + Lookbehind is fixed-width only.
'     + Atomic groups protection is capture-slot based. When an atomic group matches it locks
'       the capture slots that changed during the atomic group's internal match. Those locked
'       slots will not be restored by RestoreCaps. This is correct semantics: atomic group
'       prevents backtracking into the group.
'     + Possessive quantifiers are supported. In the quantifier attempt loop your possessive
'       flag prevents trying smaller k values.
'     + Performance: atomic groups and lookarounds use snapshots. Possessive quantifiers will
'       often be faster (no backtracking), but atomic groups may make some matches irreversible.
Option Explicit

' ---------------- Public config -----------------------------------------
Private Const DEFAULT_MAX_STEPS As Long = 200000
' --- debugging toggle: set True to print detailed trace during matching
Private Const DEBUG_TRACE As Boolean = False

' ---------------- Public state ------------------------------------------
Private pPattern As String
Private pIgnoreCase As Boolean
Private pMaxMatchSteps As Long
Private pAST As Collection
Private pGroupCount As Long
Private pMultiline As Boolean
Private pDotAll As Boolean

' Current subject & captures
Private CurSubject As String
Private capsStart() As Long
Private capsEnd() As Long
Private atomicLocks() As Boolean
Private MatchSteps As Long

' ---------------- Node helpers ------------------------------------------
Private Function NewNode(ByVal typ As String) As Collection
    Dim c As New Collection
    c.Add typ, "type"
    Set NewNode = c
End Function

Private Sub NodeAdd(ByRef node As Collection, ByVal key As String, ByVal value As Variant)
    If NodeExists(node, key) Then
        On Error Resume Next
        node.Remove key
        On Error GoTo 0
    End If
    node.Add value, key
End Sub

Private Sub vAssignVar(ByRef dest As Variant, ByVal src As Variant)
    If IsObject(src) Then Set dest = src Else dest = src
End Sub

Private Function NodeGet(ByVal node As Collection, ByVal key As String) As Variant
    On Error GoTo notfound
    vAssignVar NodeGet, node(key)
    Exit Function
notfound:
    err.Clear
    NodeGet = Empty
End Function

Private Function NodeExists(ByVal node As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant: vAssignVar tmp, node(key)
    If err.Number = 0 Then NodeExists = True Else NodeExists = False
    err.Clear
    On Error GoTo 0
End Function

' ---------------- Public API -------------------------------------------
Public Sub Init(ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False, Optional ByVal maxSteps As Long = DEFAULT_MAX_STEPS, Optional ByVal multiline As Boolean = False, Optional ByVal dotAll As Boolean = False)
    pPattern = pattern
    pIgnoreCase = ignoreCase
    pMultiline = multiline
    pDotAll = dotAll
    If maxSteps <= 0 Then pMaxMatchSteps = DEFAULT_MAX_STEPS Else pMaxMatchSteps = maxSteps

    ' Build AST
    Set pAST = CompilePattern(pattern)

    ' Run annotate/validation pass (will raise vbObjectError + 10 on invalid lookbehinds)
    AnnotateLookbehinds pAST

    ' compute group count after validation
    pGroupCount = GetGroupCount(pAST)
End Sub
Public Property Get pattern() As String: pattern = pPattern: End Property
Public Property Let pattern(ByVal v As String): Init v, pIgnoreCase, pMaxMatchSteps, pMultiline, pDotAll: End Property
Public Property Get ignoreCase() As Boolean: ignoreCase = pIgnoreCase: End Property
Public Property Let ignoreCase(ByVal v As Boolean): pIgnoreCase = v: End Property
Public Property Get MaxMatchSteps() As Long: MaxMatchSteps = pMaxMatchSteps: End Property
Public Property Let MaxMatchSteps(ByVal v As Long): pMaxMatchSteps = v: End Property
Public Property Get multiline() As Boolean: multiline = pMultiline: End Property
Public Property Let multiline(ByVal v As Boolean): pMultiline = v: End Property
Public Property Get dotAll() As Boolean: dotAll = pDotAll: End Property
Public Property Let dotAll(ByVal v As Boolean): pDotAll = v: End Property

Public Function Test(ByVal subj As String) As Boolean
    Dim res As Collection
    Set res = Exec(subj)
    Test = Not (res Is Nothing)
End Function

' ---------------- Parser & AST precomputation ----------------------------
Private Function CompilePattern(ByVal pat As String) As Collection
    Dim ast As New Collection
    NodeAdd ast, "raw", pat

    Dim anchoredStart As Boolean: anchoredStart = False
    Dim anchoredEnd As Boolean: anchoredEnd = False
    If Len(pat) > 0 And SafeMid(pat, 1, 1) = "^" Then anchoredStart = True: pat = SafeMid(pat, 2)
    If Len(pat) > 0 And right$(pat, 1) = "$" Then anchoredEnd = True: pat = left$(pat, Len(pat) - 1)
    NodeAdd ast, "anchoredStart", anchoredStart
    NodeAdd ast, "anchoredEnd", anchoredEnd

    Dim pos As Long: pos = 1
    Dim groupCounter As Long: groupCounter = 0
    Dim body As Collection: Set body = ParseAlternation(pat, pos, Len(pat), groupCounter)
    NodeAdd ast, "body", body

    ' Precompute minLen
    Dim mlen As Long: mlen = ComputeMinLen(body)
    NodeAdd ast, "minLen", mlen
    
    ' Validate/lookbehind annotate
    AnnotateLookbehinds ast
    
    Set CompilePattern = ast
End Function

' ParseAlternation: sequences separated by '|'
Private Function ParseAlternation(ByVal pat As String, ByRef p As Long, ByVal pEnd As Long, ByRef groupCounter As Long) As Collection
    Dim firstSeq As Collection
    Set firstSeq = ParseSequence(pat, p, pEnd, groupCounter)

    If Not (p <= pEnd And SafeMid(pat, p, 1) = "|") Then
        Set ParseAlternation = firstSeq
        Exit Function
    End If

    Dim altNode As Collection: Set altNode = NewNode("alt")
    Dim alts As New Collection
    alts.Add firstSeq

    Do While p <= pEnd And SafeMid(pat, p, 1) = "|"
        p = p + 1 ' consume '|'
        Dim nextSeq As Collection
        Set nextSeq = ParseSequence(pat, p, pEnd, groupCounter)
        alts.Add nextSeq
    Loop

    NodeAdd altNode, "children", alts
    Set ParseAlternation = altNode
End Function

Private Function ParseSequence(ByVal pat As String, ByRef p As Long, ByVal pEnd As Long, ByRef groupCounter As Long) As Collection
    Dim seq As Collection: Set seq = NewNode("seq")
    Dim children As New Collection
    Do While p <= pEnd
        Dim ch As String: ch = SafeMid(pat, p, 1)
        If ch = ")" Or ch = "|" Then Exit Do
        Dim node As Collection: Set node = Nothing
        If ch = "\" Then
            p = p + 1
            If p > pEnd Then Exit Do
            Set node = ParseEscapeNode(SafeMid(pat, p, 1))
            p = p + 1
        ElseIf ch = "." Then
            Set node = NewNode("dot")
            p = p + 1
        ElseIf ch = "[" Then
            Dim clsEnd As Long: clsEnd = FindClassEnd(pat, p, pEnd)
            If clsEnd = -1 Then
                Set node = NewNode("lit"): NodeAdd node, "ch", "[": p = p + 1
            Else
                Set node = ParseClass(pat, p + 1, clsEnd - 1)
                p = clsEnd + 1
            End If
        ElseIf ch = "(" Then
            p = p + 1
            ' Special group or plain group?
            If p <= pEnd And SafeMid(pat, p, 1) = "?" Then
                p = p + 1
                If p <= pEnd Then
                    Dim nextCh As String: nextCh = SafeMid(pat, p, 1)
                    Select Case nextCh
                        Case "=" ' (?= ) positive lookahead
                            p = p + 1
                            Dim innerLA As Collection
                            Set innerLA = ParseSequence(pat, p, pEnd, groupCounter)
                            If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                            Dim laNode As Collection: Set laNode = NewNode("lookahead_pos")
                            NodeAdd laNode, "child", innerLA
                            Set node = laNode
                        Case "!" ' (?! ) negative lookahead
                            p = p + 1
                            Dim innerNL As Collection
                            Set innerNL = ParseSequence(pat, p, pEnd, groupCounter)
                            If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                            Dim nlaNode As Collection: Set nlaNode = NewNode("lookahead_neg")
                            NodeAdd nlaNode, "child", innerNL
                            Set node = nlaNode
                        Case "<" ' (?<= or (?<! ) lookbehinds
                            p = p + 1
                            If p <= pEnd Then
                                Dim sbch As String: sbch = SafeMid(pat, p, 1)
                                If sbch = "=" Then
                                    p = p + 1
                                    Dim innerLB As Collection
                                    Set innerLB = ParseSequence(pat, p, pEnd, groupCounter)
                                    If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                    Dim lbNode As Collection: Set lbNode = NewNode("lookbehind_pos")
                                    NodeAdd lbNode, "child", innerLB
                                    Set node = lbNode
                                ElseIf sbch = "!" Then
                                    p = p + 1
                                    Dim innerLBN As Collection
                                    Set innerLBN = ParseSequence(pat, p, pEnd, groupCounter)
                                    If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                    Dim lbnNode As Collection: Set lbnNode = NewNode("lookbehind_neg")
                                    NodeAdd lbnNode, "child", innerLBN
                                    Set node = lbnNode
                                Else
                                    ' Unknown (?<... ) construct -> treat as literal fallback
                                    Set node = NewNode("lit"): NodeAdd node, "ch", "("
                                End If
                            Else
                                Set node = NewNode("lit"): NodeAdd node, "ch", "("
                            End If
                        Case ">" ' (?> ) atomic group
                            p = p + 1
                            Dim innerAtomic As Collection
                            Set innerAtomic = ParseSequence(pat, p, pEnd, groupCounter)
                            If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                            Dim atomicNode As Collection: Set atomicNode = NewNode("atomic")
                            NodeAdd atomicNode, "child", innerAtomic
                            Set node = atomicNode
                        Case Else
                            ' Unknown ?something: fallback to literal '('
                            Set node = NewNode("lit"): NodeAdd node, "ch", "("
                    End Select
                Else
                    Set node = NewNode("lit"): NodeAdd node, "ch", "("
                End If
            Else
                ' normal capturing group
                groupCounter = groupCounter + 1
                Dim thisIdx As Long: thisIdx = groupCounter
                Dim inner As Collection
                Set inner = ParseSequence(pat, p, pEnd, groupCounter)
                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                Dim gnode As Collection: Set gnode = NewNode("group")
                NodeAdd gnode, "idx", thisIdx
                NodeAdd gnode, "child", inner
                Set node = gnode
            End If
        Else
            Set node = NewNode("lit"): NodeAdd node, "ch", ch
            p = p + 1
        End If

        ' quantifier (supports possessive + after quantifier)
        If p <= pEnd Then
            Dim qch As String: qch = SafeMid(pat, p, 1)
            If qch = "*" Or qch = "+" Or qch = "?" Or qch = "{" Then
                Dim qnode As Collection: Set qnode = NewNode("quant")
                NodeAdd qnode, "child", node
                NodeAdd qnode, "possessive", False
                If qch = "*" Then
                    NodeAdd qnode, "min", 0: NodeAdd qnode, "max", -1
                    p = p + 1
                ElseIf qch = "+" Then
                    NodeAdd qnode, "min", 1: NodeAdd qnode, "max", -1
                    p = p + 1
                ElseIf qch = "?" Then
                    NodeAdd qnode, "min", 0: NodeAdd qnode, "max", 1
                    p = p + 1
                ElseIf qch = "{" Then
                    Dim closePos As Long: closePos = InStr(p, pat, "}")
                    If closePos > 0 Then
                        Dim content As String: content = SafeMid(pat, p + 1, closePos - p - 1)
                        Dim parts() As String: parts = VBA.Split(content, ",")
                        Dim mi As Long, ma As Long
                        If UBound(parts) = 0 Then
                            If Len(Trim$(parts(0))) = 0 Then mi = 0: ma = -1 Else mi = CLng(Trim$(parts(0))): ma = mi
                        Else
                            mi = CLng(Trim$(parts(0)))
                            If Len(Trim$(parts(1))) = 0 Then ma = -1 Else ma = CLng(Trim$(parts(1)))
                        End If
                        NodeAdd qnode, "min", mi: NodeAdd qnode, "max", ma
                        p = closePos + 1
                    Else
                        NodeAdd qnode, "min", 0: NodeAdd qnode, "max", 1
                        p = p + 1
                    End If
                End If
                ' lazy or possessive flag
                If p <= pEnd Then
                    Dim nextQ As String: nextQ = SafeMid(pat, p, 1)
                    If nextQ = "?" Then
                        NodeAdd qnode, "lazy", True
                        p = p + 1
                    ElseIf nextQ = "+" Then
                        NodeAdd qnode, "possessive", True
                        NodeAdd qnode, "lazy", False
                        p = p + 1
                    Else
                        NodeAdd qnode, "lazy", False
                    End If
                Else
                    NodeAdd qnode, "lazy", False
                End If
                Set node = qnode
            End If
        End If

        children.Add node
    Loop

    ' Build seq node for the parsed children
    NodeAdd seq, "children", children

    ' If the current char is '|' then construct an alternation containing this seq and subsequent seqs
    If p <= pEnd And SafeMid(pat, p, 1) = "|" Then
        Dim altNode As Collection: Set altNode = NewNode("alt")
        Dim altChildren As New Collection
        ' add the seq we just parsed as first alternative
        altChildren.Add seq
        ' iterate over remaining alternatives
        Do While p <= pEnd And SafeMid(pat, p, 1) = "|"
            p = p + 1 ' consume '|'
            Dim nextSeq As Collection
            Set nextSeq = ParseSequence(pat, p, pEnd, groupCounter)
            ' if the nextSeq is empty sequence, ensure it's still a seq node (could be empty)
            altChildren.Add nextSeq
        Loop
        NodeAdd altNode, "children", altChildren
        Set ParseSequence = altNode
        Exit Function
    End If

    Set ParseSequence = seq
End Function


Private Function ParseEscapeNode(ByVal ch As String) As Collection
    Dim node As Collection
    Select Case ch
        Case "d": Set node = ParseClassFromSimple("[0-9]")
        Case "w": Set node = ParseClassFromSimple("[A-Za-z0-9_]")
        Case "s": Set node = ParseClassFromSimple("[ " & vbTab & vbLf & vbCr & "]")
        Case "n": Set node = NewNode("lit"): NodeAdd node, "ch", vbLf
        Case "r": Set node = NewNode("lit"): NodeAdd node, "ch", vbCr
        Case "t": Set node = NewNode("lit"): NodeAdd node, "ch", vbTab
        Case Else: Set node = NewNode("lit"): NodeAdd node, "ch", ch
    End Select
    Set ParseEscapeNode = node
End Function

' Parse class and build a 256-byte boolean map for quick matching (ASCII fast-path)
Private Function ParseClass(ByVal pat As String, ByVal pStart As Long, ByVal pEnd As Long) As Collection
    Dim node As Collection: Set node = NewNode("class")
    Dim chars As New Collection
    Dim neg As Boolean: neg = False
    Dim i As Long: i = pStart
    If i <= pEnd And SafeMid(pat, i, 1) = "^" Then neg = True: i = i + 1
    Do While i <= pEnd
        Dim ch As String: ch = SafeMid(pat, i, 1)
        If ch = "\" Then
            i = i + 1
            If i <= pEnd Then chars.Add ParseEscapeChar(SafeMid(pat, i, 1))
            i = i + 1
        ElseIf i + 2 <= pEnd And SafeMid(pat, i + 1, 1) = "-" Then
            chars.Add Array("range", SafeMid(pat, i, 1), SafeMid(pat, i + 2, 1))
            i = i + 3
        Else
            chars.Add ch
            i = i + 1
        End If
    Loop
    NodeAdd node, "chars", chars
    NodeAdd node, "neg", neg
    ' Precompute ASCII map
    Dim mapArr() As Boolean: ReDim mapArr(0 To 255) As Boolean
    Dim it As Variant
    For Each it In chars
        If IsArray(it) Then
            If it(0) = "range" Then
                Dim aCode As Long: aCode = AscW(it(1))
                Dim bCode As Long: bCode = AscW(it(2))
                Dim cc As Long
                For cc = aCode To bCode
                    If cc >= 0 And cc <= 255 Then mapArr(cc) = True
                Next cc
            End If
        Else
            Dim code As Long: code = AscW(it)
            If code >= 0 And code <= 255 Then mapArr(code) = True
        End If
    Next it
    NodeAdd node, "map", mapArr
    Set ParseClass = node
End Function

Private Function ParseClassFromSimple(ByVal spec As String) As Collection
    Dim node As Collection: Set node = NewNode("class")
    Dim chars As New Collection
    Dim i As Long: i = 2
    Do While i <= Len(spec) - 1
        Dim ch As String: ch = SafeMid(spec, i, 1)
        If i + 2 <= Len(spec) - 1 And SafeMid(spec, i + 1, 1) = "-" Then
            chars.Add Array("range", ch, SafeMid(spec, i + 2, 1))
            i = i + 3
        Else
            chars.Add ch
            i = i + 1
        End If
    Loop
    NodeAdd node, "chars", chars
    NodeAdd node, "neg", False
    ' build simple map
    Dim mapArr() As Boolean: ReDim mapArr(0 To 255) As Boolean
    Dim it As Variant
    For Each it In chars
        If IsArray(it) Then
            Dim aCode As Long: aCode = AscW(it(1))
            Dim bCode As Long: bCode = AscW(it(2))
            Dim cc As Long
            For cc = aCode To bCode
                If cc >= 0 And cc <= 255 Then mapArr(cc) = True
            Next cc
        Else
            Dim ccode As Long: ccode = AscW(it)
            If ccode >= 0 And ccode <= 255 Then mapArr(ccode) = True
        End If
    Next it
    NodeAdd node, "map", mapArr
    Set ParseClassFromSimple = node
End Function

Private Function ParseEscapeChar(ByVal ch As String) As String
    Select Case ch
        Case "n": ParseEscapeChar = vbLf
        Case "r": ParseEscapeChar = vbCr
        Case "t": ParseEscapeChar = vbTab
        Case Else: ParseEscapeChar = ch
    End Select
End Function

' ------------------ Matching engine (returns explicit end positions) --------
Private Function GetGroupCount(ByVal ast As Collection) As Long
    Dim maxIdx As Long: maxIdx = 0
    TraverseForGroups NodeGet(ast, "body"), maxIdx
    GetGroupCount = maxIdx
End Function

Private Sub TraverseForGroups(ByVal node As Collection, ByRef maxIdx As Long)
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "seq"
            Dim ch As Variant
            For Each ch In NodeGet(node, "children")
                TraverseForGroups ch, maxIdx
            Next
        Case "group"
            If NodeExists(node, "idx") Then If NodeGet(node, "idx") > maxIdx Then maxIdx = NodeGet(node, "idx")
            TraverseForGroups NodeGet(node, "child"), maxIdx
        Case "quant"
            TraverseForGroups NodeGet(node, "child"), maxIdx
        Case "alt"
            Dim altc As Variant
            For Each altc In NodeGet(node, "children")
                TraverseForGroups altc, maxIdx
            Next
    End Select
End Sub

Private Function MatchNode(ByVal node As Collection, ByVal subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    If DEBUG_TRACE Then Debug.Print "TRACE MatchNode ENTER type=" & CStr(NodeGet(node, "type")) & " pos=" & pos & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
    MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"

    Dim t As String: t = CStr(NodeGet(node, "type"))
    Dim res As Boolean: res = False
    Dim localEnd As Variant: localEnd = Empty

    Select Case t
        Case "seq"
            res = MatchSequence(NodeGet(node, "children"), subj, pos, localEnd)

        Case "lit"
            If pos > Len(subj) Then
                res = False
            Else
                Dim ch As String: ch = NodeGet(node, "ch")
                Dim sc As String: sc = SafeMid(subj, pos, 1)
                If pIgnoreCase Then
                    If LCase$(sc) = LCase$(ch) Then localEnd = pos: res = True
                Else
                    If sc = ch Then localEnd = pos: res = True
                End If
            End If

        Case "dot"
            If pos > Len(subj) Then
                res = False
            Else
                Dim dch As String: dch = SafeMid(subj, pos, 1)
                If (dch = vbLf Or dch = vbCr) And Not pDotAll Then
                    res = False
                Else
                    localEnd = pos: res = True
                End If
            End If

        Case "class"
            If pos > Len(subj) Then
                res = False
            Else
                If MatchCharClassNode(node, SafeMid(subj, pos, 1)) Then localEnd = pos: res = True
            End If

        Case "group"
            Dim gIdx As Long: gIdx = NodeGet(node, "idx")
            Dim prevStart As Long: prevStart = capsStart(gIdx)
            Dim prevEnd As Long: prevEnd = capsEnd(gIdx)
            capsStart(gIdx) = pos
            Dim childEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, childEnd) Then
                capsEnd(gIdx) = childEnd
                localEnd = childEnd
                res = True
            Else
                capsStart(gIdx) = prevStart
                capsEnd(gIdx) = prevEnd
                res = False
            End If

        Case "alt"
            Dim altChild As Variant
            For Each altChild In NodeGet(node, "children")
                Dim snapAlt() As Long: snapAlt = CopyCaps()
                Dim altEnd As Variant
                If MatchNode(altChild, subj, pos, altEnd) Then
                    localEnd = altEnd
                    res = True
                    Exit For
                Else
                    RestoreCaps snapAlt
                End If
            Next altChild

        Case "quant"
            ' This path is rarely used now since we handle quantifiers inline in MatchSequence.
            Dim mn As Long: mn = NodeGet(node, "min")
            Dim mx As Long: mx = NodeGet(node, "max")
            Dim lazy As Boolean: lazy = NodeGet(node, "lazy")
            Dim poss As Boolean: poss = NodeGet(node, "possessive")
            If mx = -1 Then mx = Len(subj) - pos + 1
            ' Fallback simple behaviour: attempt greedy expansion and accept largest if poss else backtracking handled by MatchSequence
            Dim occPos() As Long: ReDim occPos(0 To 0)
            Dim occCount As Long: occCount = 0
            Dim curPos As Long: curPos = pos
            Dim nodeEnd As Variant
            Do While occCount < mx
                Dim tempSnap() As Long: tempSnap = CopyCaps()
                If MatchNode(NodeGet(node, "child"), subj, curPos, nodeEnd) Then
                    occCount = occCount + 1
                    ReDim Preserve occPos(0 To occCount)
                    occPos(occCount) = nodeEnd
                    curPos = nodeEnd + 1
                Else
                    RestoreCaps tempSnap: Exit Do
                End If
            Loop
            If occCount >= mn Then
                localEnd = occPos(occCount)
                res = True
            Else
                res = False
            End If

        Case "lookahead_pos"
            ' Snapshot captures, test child at current pos, restore captures
            Dim snapLA() As Long: snapLA = CopyCaps()
            Dim laEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, laEnd) Then
                res = True
            Else
                res = False
            End If
            RestoreCaps snapLA

        Case "lookahead_neg"
            Dim snapLN() As Long: snapLN = CopyCaps()
            Dim lnEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, lnEnd) Then
                res = False
            Else
                res = True
            End If
            RestoreCaps snapLN

        Case "lookbehind_pos"
            ' Fixed-width lookbehind. Use annotated width if present; otherwise compute/validate.
            Dim widthLB As Long
            If NodeExists(node, "width") Then
                widthLB = CLng(NodeGet(node, "width"))
            Else
                Dim lbMin As Long: lbMin = ComputeMinLen(NodeGet(node, "child"))
                Dim lbMax As Long: lbMax = ComputeMaxLen(NodeGet(node, "child"))
                If lbMin = -1 Or lbMax = -1 Or lbMin <> lbMax Then
                    err.Raise vbObjectError + 10, "RegexEngine", "Unsupported variable-length lookbehind (must be fixed-width)."
                End If
                widthLB = lbMin
                NodeAdd node, "width", widthLB
            End If
        
            If widthLB < 1 Or (pos - widthLB) < 1 Then
                res = False
            Else
                Dim startLB As Long: startLB = pos - widthLB
                Dim snapLB() As Long: snapLB = CopyCaps()
                Dim lbEnd As Variant
                If MatchNode(NodeGet(node, "child"), subj, startLB, lbEnd) Then
                    ' require the inner match to end exactly at pos-1
                    If CLng(lbEnd) = pos - 1 Then
                        ' indicate we consumed nothing by returning pos-1 as localEnd
                        localEnd = pos - 1
                        res = True
                    Else
                        res = False
                    End If
                Else
                    res = False
                End If
                RestoreCaps snapLB
            End If
        
        Case "lookbehind_neg"
            Dim widthLBN As Long
            If NodeExists(node, "width") Then
                widthLBN = CLng(NodeGet(node, "width"))
            Else
                Dim lbnMin As Long: lbnMin = ComputeMinLen(NodeGet(node, "child"))
                Dim lbnMax As Long: lbnMax = ComputeMaxLen(NodeGet(node, "child"))
                If lbnMin = -1 Or lbnMax = -1 Or lbnMin <> lbnMax Then
                    err.Raise vbObjectError + 10, "RegexEngine", "Unsupported variable-length lookbehind (must be fixed-width)."
                End If
                widthLBN = lbnMin
                NodeAdd node, "width", widthLBN
            End If
        
            If widthLBN < 1 Or (pos - widthLBN) < 1 Then
                ' Not enough chars before pos -> negative lookbehind trivially true; consume nothing
                localEnd = pos - 1
                res = True
            Else
                Dim startLBN As Long: startLBN = pos - widthLBN
                Dim snapLBN() As Long: snapLBN = CopyCaps()
                Dim lbnEnd As Variant
                If MatchNode(NodeGet(node, "child"), subj, startLBN, lbnEnd) Then
                    If CLng(lbnEnd) = pos - 1 Then
                        ' inner matched immediately before pos -> negative fails
                        res = False
                    Else
                        localEnd = pos - 1
                        res = True
                    End If
                Else
                    ' inner didn't match -> negative succeeds; consume nothing
                    localEnd = pos - 1
                    res = True
                End If
                RestoreCaps snapLBN
            End If

        Case "atomic"
            ' Attempt to match child; if it succeeds protect changed capture slots from later restore
            Dim beforeSnap() As Long: beforeSnap = CopyCaps()
            Dim atomEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, atomEnd) Then
                res = True
                localEnd = atomEnd
                ' mark which capture slots changed and lock them
                Dim tot As Long: tot = pGroupCount
                Dim idx As Long
                For idx = 0 To tot
                    Dim beforeS As Long: beforeS = beforeSnap(idx)
                    Dim beforeE As Long: beforeE = beforeSnap(idx + tot + 1)
                    If capsStart(idx) <> beforeS Or capsEnd(idx) <> beforeE Then
                        atomicLocks(idx) = True
                    End If
                Next idx
            Else
                ' restore beforeSnap (child already handled restore in branches), keep atomicLocks unchanged
                RestoreCaps beforeSnap
                res = False
            End If

        Case Else
            res = False
    End Select

    ' synchronize localEnd -> capsEnd(0) and outEnd
    If res Then
        If Not IsEmpty(localEnd) Then
            capsEnd(0) = CLng(localEnd)
            If Not IsMissing(outEnd) Then outEnd = CLng(localEnd)
        Else
            If Not IsMissing(outEnd) Then outEnd = capsEnd(0)
        End If
    End If

    If DEBUG_TRACE Then Debug.Print "TRACE MatchNode EXIT type=" & t & " pos=" & pos & " res=" & CStr(res) & " localEnd=" & CStr(localEnd) & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
    MatchNode = res
End Function

Private Function MatchSequence(ByVal children As Collection, ByVal subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    If DEBUG_TRACE Then Debug.Print "TRACE MatchSequence start pos=" & pos & " children=" & children.count
    Dim seqSnap() As Long: seqSnap = CopyCaps()
    Dim curPos As Long: curPos = pos
    Dim i As Long

    For i = 1 To children.count
        Dim node As Collection: Set node = children(i)
        Dim nodeSnap() As Long: nodeSnap = CopyCaps()
        Dim nodeEnd As Variant
        Dim nodeType As String: nodeType = NodeGet(node, "type")

        If DEBUG_TRACE Then Debug.Print "TRACE  seq child#" & i & " type=" & nodeType & " curPos=" & curPos

        If nodeType = "quant" Then
            ' --- handle quantifier with full backtracking (inline remainder matching) ---
            Dim childNode As Collection: Set childNode = NodeGet(node, "child")
            Dim mn As Long: mn = NodeGet(node, "min")
            Dim mx As Long: mx = NodeGet(node, "max")
            Dim lazy As Boolean: lazy = NodeGet(node, "lazy")
            Dim possFlag As Boolean
            If NodeExists(node, "possessive") Then possFlag = NodeGet(node, "possessive") Else possFlag = False
            If mx = -1 Then mx = Len(subj) - curPos + 1
        
            ' save sequence baseline snapshot
            seqSnap = CopyCaps()
        
            ' collect all possible end positions by repeatedly matching child
            Dim occEnds() As Long
            Dim occCount As Long: occCount = 0
            Dim scanPos As Long: scanPos = curPos
            Do While occCount < mx
                Dim preSnap() As Long: preSnap = CopyCaps()
                Dim ndEnd As Variant
                If MatchNode(childNode, subj, scanPos, ndEnd) Then
                    occCount = occCount + 1
                    ReDim Preserve occEnds(1 To occCount)
                    occEnds(occCount) = CLng(ndEnd)
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant matched occ " & occCount & " nodeEnd=" & ndEnd
                    scanPos = CLng(ndEnd) + 1
                Else
                    ' restore captures from failed attempt and stop expanding
                    RestoreCaps preSnap
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant stop at occCount=" & occCount
                    Exit Do
                End If
            Loop
        
            ' If possessive, only try the maximal occCount (no backtracking to smaller k)
            If possFlag Then
                If occCount < mn Then
                    ' not enough occurrences even at max -> fail immediately
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant possessive failed: occCount < min"
                    RestoreCaps seqSnap
                    MatchSequence = False: Exit Function
                End If
        
                ' attempt remainder only with k = occCount
                If DEBUG_TRACE Then Debug.Print "TRACE  quant possessive trying k=" & occCount & " (mn=" & mn & " occCount=" & occCount & ")"
                RestoreCaps seqSnap
                Dim posTryP As Long: posTryP = curPos
                Dim tP As Long
                Dim okAllP As Boolean: okAllP = True
                For tP = 1 To occCount
                    Dim snapTmpP() As Long: snapTmpP = CopyCaps()
                    Dim ndEndP As Variant
                    If MatchNode(childNode, subj, posTryP, ndEndP) Then
                        posTryP = CLng(ndEndP) + 1
                    Else
                        RestoreCaps snapTmpP
                        okAllP = False
                        Exit For
                    End If
                Next tP
        
                If Not okAllP Then
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant possessive failed while building occurrences"
                    RestoreCaps seqSnap
                    MatchSequence = False: Exit Function
                End If
        
                ' Now attempt remainder once
                Dim remSuccessP As Boolean: remSuccessP = True
                Dim jP As Long
                Dim nxtEndP As Variant
                For jP = i + 1 To children.count
                    Dim nxtNodeP As Collection: Set nxtNodeP = children(jP)
                    If DEBUG_TRACE Then Debug.Print "TRACE   quant possessive trying remainder child j=" & jP & " at posTry=" & posTryP & " type=" & NodeGet(nxtNodeP, "type")
                    If Not MatchNode(nxtNodeP, subj, posTryP, nxtEndP) Then
                        remSuccessP = False
                        If DEBUG_TRACE Then Debug.Print "TRACE   possessive remainder child j=" & jP & " failed at pos=" & posTryP
                        Exit For
                    Else
                        posTryP = CLng(nxtEndP) + 1
                    End If
                Next jP
        
                If remSuccessP Then
                    outEnd = posTryP - 1
                    capsEnd(0) = CLng(outEnd)
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant possessive accepted k=" & occCount & " outEnd=" & outEnd
                    MatchSequence = True: Exit Function
                Else
                    ' possessive prevents trying smaller k -> fail the quantifier
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant possessive remainder failed -> overall fail"
                    RestoreCaps seqSnap
                    MatchSequence = False: Exit Function
                End If
            End If
        
            ' Decide iteration order (lazy: small->big, greedy: big->small)
            Dim kStart As Long, kEnd As Long, kStep As Long
            If lazy Then
                kStart = mn: kEnd = occCount: kStep = 1
            Else
                kStart = occCount: kEnd = mn: kStep = -1
            End If
        
            Dim k As Long
            For k = kStart To kEnd Step kStep
                If DEBUG_TRACE Then Debug.Print "TRACE  quant trying k=" & k & " (mn=" & mn & " occCount=" & occCount & ")"
                ' restore sequence baseline before attempting this k
                RestoreCaps seqSnap
        
                ' apply k repetitions of child to set captures accordingly
                Dim posTry As Long: posTry = curPos
                Dim t As Long
                Dim okAll As Boolean: okAll = True
                For t = 1 To k
                    Dim snapTmp() As Long: snapTmp = CopyCaps()
                    Dim ndEnd2 As Variant
                    If MatchNode(childNode, subj, posTry, ndEnd2) Then
                        posTry = CLng(ndEnd2) + 1
                    Else
                        RestoreCaps snapTmp
                        okAll = False
                        Exit For
                    End If
                Next t
        
                If Not okAll Then
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant k=" & k & " failed while building occurrences"
                    ' try next k
                    GoTo ContinueFor
                End If
        
                ' Now match the remainder *inline* (avoid building a separate 'rest' and recursing)
                Dim remSuccess As Boolean: remSuccess = True
                Dim j As Long
                Dim nxtEnd As Variant
                For j = i + 1 To children.count
                    Dim nxtNode As Collection: Set nxtNode = children(j)
                    If DEBUG_TRACE Then Debug.Print "TRACE   quant trying remainder child j=" & j & " at posTry=" & posTry & " type=" & NodeGet(nxtNode, "type")
                    If Not MatchNode(nxtNode, subj, posTry, nxtEnd) Then
                        remSuccess = False
                        If DEBUG_TRACE Then Debug.Print "TRACE   remainder child j=" & j & " failed at pos=" & posTry
                        Exit For
                    Else
                        posTry = CLng(nxtEnd) + 1
                    End If
                Next j
        
                If remSuccess Then
                    outEnd = posTry - 1
                    capsEnd(0) = CLng(outEnd)
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant accepted k=" & k & " outEnd=" & outEnd
                    MatchSequence = True: Exit Function
                Else
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant k=" & k & " remainder failed; will try next k"
                    ' continue to next k (seqSnap will be restored at loop top)
                End If
ContinueFor:
            Next k
        
            ' none of the k choices worked; restore seq baseline and fail
            If DEBUG_TRACE Then Debug.Print "TRACE  quant exhausted all k options -> fail"
            RestoreCaps seqSnap
            MatchSequence = False: Exit Function

        ElseIf nodeType = "group" Then
            Dim gIdx As Long: gIdx = NodeGet(node, "idx")
            If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group idx=" & gIdx & " set start=" & curPos
            capsStart(gIdx) = curPos
            Dim childEnd As Variant
            If Not MatchNode(NodeGet(node, "child"), subj, curPos, childEnd) Then
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group FAILED, restoring seqSnap"
                RestoreCaps seqSnap
                MatchSequence = False: Exit Function
            Else
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group OK nodeEnd=" & childEnd
                capsEnd(gIdx) = childEnd
                curPos = CLng(childEnd) + 1
            End If

        Else
            If Not MatchNode(node, subj, curPos, nodeEnd) Then
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " node FAILED, restoring seqSnap"
                RestoreCaps seqSnap
                MatchSequence = False: Exit Function
            Else
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " node OK nodeEnd=" & nodeEnd
                curPos = CLng(nodeEnd) + 1
            End If
        End If
    Next i

    outEnd = curPos - 1
    capsEnd(0) = CLng(outEnd)
    If DEBUG_TRACE Then Debug.Print "TRACE MatchSequence success outEnd=" & outEnd & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
    MatchSequence = True
End Function

' MatchQuantifier: tries to match between min..max occurrences.
' For greedy: try larger k first; for lazy: try smaller first.
' If possessive = True -> accept the maximal match and DO NOT try smaller k's.
Private Function MatchQuantifier(ByVal child As Collection, ByVal minc As Long, ByVal maxc As Long, ByVal lazy As Boolean, ByVal subj As String, ByVal posStart As Long, ByVal possessive As Boolean, ByRef outEnd As Long) As Boolean
    Dim saveSnapshot() As Long: saveSnapshot = CopyCaps()
    Dim occPos() As Long
    ReDim occPos(0 To 0)
    Dim occCount As Long: occCount = 0
    Dim curPos As Long: curPos = posStart
    Dim nodeEnd As Long

    ' Greedily consume up to maxc occurrences (or until child fails).
    Do While occCount < maxc
        Dim tempSnap() As Long: tempSnap = CopyCaps()
        If MatchNode(child, subj, curPos, nodeEnd) Then
            occCount = occCount + 1
            ReDim Preserve occPos(0 To occCount)
            occPos(occCount) = nodeEnd
            curPos = nodeEnd + 1
        Else
            ' restore partial child state (the last attempted child failed)
            RestoreCaps tempSnap
            Exit Do
        End If
    Loop

    ' If possessive: accept the maximal consumption (if >= min), do not attempt smaller k
    If possessive Then
        If occCount < minc Then
            MatchQuantifier = False
            Exit Function
        Else
            ' leave caps as they are (they reflect the maximal consumption)
            outEnd = occPos(occCount)
            MatchQuantifier = True
            Exit Function
        End If
    End If

    ' Non-possessive behavior proceeds as before:
    Dim k As Long
    If lazy Then
        ' Lazy: try k = min..occCount
        For k = minc To occCount
            RestoreCaps saveSnapshot
            Dim posTry As Long: posTry = posStart
            Dim t As Long
            Dim okAll As Boolean: okAll = True
            For t = 1 To k
                Dim tmp() As Long: tmp = CopyCaps()
                If MatchNode(child, subj, posTry, nodeEnd) Then
                    posTry = nodeEnd + 1
                Else
                    RestoreCaps tmp
                    okAll = False
                    Exit For
                End If
            Next t
            If okAll Then outEnd = posTry - 1: MatchQuantifier = True: Exit Function
        Next k
        MatchQuantifier = False
    Else
        ' Greedy: try k = occCount down to minc
        For k = occCount To minc Step -1
            RestoreCaps saveSnapshot
            Dim posTry2 As Long: posTry2 = posStart
            Dim u As Long
            Dim okAll2 As Boolean: okAll2 = True
            For u = 1 To k
                Dim tmp2() As Long: tmp2 = CopyCaps()
                If MatchNode(child, subj, posTry2, nodeEnd) Then
                    posTry2 = nodeEnd + 1
                Else
                    RestoreCaps tmp2
                    okAll2 = False
                    Exit For
                End If
            Next u
            If okAll2 Then outEnd = posTry2 - 1: MatchQuantifier = True: Exit Function
        Next k
        MatchQuantifier = False
    End If
End Function


' ---------------- Captures snapshot utilities -----------------------------
Private Function CopyCaps() As Long()
    Dim tot As Long: tot = pGroupCount
    Dim slots As Long: slots = tot + 1
    Dim res() As Long
    ReDim res(0 To slots * 2 - 1)
    Dim i As Long
    For i = 0 To tot
        res(i) = capsStart(i)
        res(i + slots) = capsEnd(i)
    Next i
    If DEBUG_TRACE Then
        Dim s As String: s = "TRACE CopyCaps -> ["
        For i = 0 To tot
            s = s & " s" & i & "=" & CStr(res(i)) & " e" & i & "=" & CStr(res(i + slots))
            If i < tot Then s = s & " |"
        Next i
        s = s & " ]"
        Debug.Print s
    End If
    CopyCaps = res
End Function

Private Sub RestoreCaps(ByRef snapshot() As Long)
    Dim slots As Long: slots = (UBound(snapshot) + 1) \ 2
    Dim tot As Long: tot = slots - 1
    Dim i As Long
    If DEBUG_TRACE Then
        Dim before As String: before = "TRACE RestoreCaps (before) -> ["
        For i = 0 To tot
            before = before & " s" & i & "=" & CStr(capsStart(i)) & " e" & i & "=" & CStr(capsEnd(i))
            If i < tot Then before = before & " |"
        Next i
        before = before & " ]"
        Debug.Print before
    End If

    For i = 0 To tot
        ' If atomicLocks(i) true, skip restoring this capture index to preserve atomic group side-effects
        If Not atomicLocks(i) Then
            capsStart(i) = snapshot(i)
            capsEnd(i) = snapshot(i + slots)
        End If
    Next i

    If DEBUG_TRACE Then
        Dim after As String: after = "TRACE RestoreCaps (after)  -> ["
        For i = 0 To tot
            after = after & " s" & i & "=" & CStr(capsStart(i)) & " e" & i & "=" & CStr(capsEnd(i))
            If i < tot Then after = after & " |"
        Next i
        after = after & " ]"
        Debug.Print after
    End If
End Sub

' ---------------- Char class matching with precomputed map -----------------
Private Function MatchCharClassNode(ByVal node As Collection, ByVal ch As String) As Boolean
    Dim code As Long: code = AscW(ch)
    If code >= 0 And code <= 255 Then
        Dim mapArr() As Boolean: mapArr = NodeGet(node, "map")
        Dim matched As Boolean: matched = mapArr(code)
        If NodeGet(node, "neg") Then MatchCharClassNode = Not matched Else MatchCharClassNode = matched
        Exit Function
    End If
    ' fallback for non-ASCII: brute-force check
    Dim items As Collection: Set items = NodeGet(node, "chars")
    Dim matched2 As Boolean: matched2 = False
    Dim i As Long
    For i = 1 To items.count
        Dim it As Variant: it = items(i)
        If IsArray(it) Then
            If it(0) = "range" Then
                If AscW(ch) >= AscW(it(1)) And AscW(ch) <= AscW(it(2)) Then matched2 = True
            End If
        Else
            If ch = it Then matched2 = True
        End If
        If matched2 Then Exit For
    Next i
    If NodeGet(node, "neg") Then MatchCharClassNode = Not matched2 Else MatchCharClassNode = matched2
End Function

' ---------------- Replacement helpers -----------------------------------
Private Function ExpandReplacement(ByVal caps As Collection, ByVal repl As String) As String
    Dim out As String: out = ""
    Dim i As Long: i = 1
    Do While i <= Len(repl)
        Dim ch As String: ch = SafeMid(repl, i, 1)
        If ch = "$" Then
            Dim j As Long: j = i + 1
            Dim num As String: num = ""
            Do While j <= Len(repl) And SafeMid(repl, j, 1) >= "0" And SafeMid(repl, j, 1) <= "9"
                num = num & SafeMid(repl, j, 1)
                j = j + 1
            Loop
            If num <> "" Then
                Dim idx As Long: idx = CLng(num)
                If idx >= 0 And idx <= caps.count - 1 Then out = out & caps.item(idx + 1)
                i = j
            Else
                out = out & "$": i = i + 1
            End If
        Else
            out = out & ch: i = i + 1
        End If
    Loop
    ExpandReplacement = out
End Function

Private Function BuildCaptureCollection() As Collection
    Dim col As New Collection
    If capsStart(0) >= 1 And capsEnd(0) >= capsStart(0) Then
        col.Add SafeMid(CurSubject, capsStart(0), capsEnd(0) - capsStart(0) + 1)
    Else
        col.Add ""
    End If
    Dim i As Long
    For i = 1 To pGroupCount
        If capsStart(i) >= 1 And capsEnd(i) >= capsStart(i) Then
            col.Add SafeMid(CurSubject, capsStart(i), capsEnd(i) - capsStart(i) + 1)
        Else
            col.Add ""
        End If
    Next i
    Set BuildCaptureCollection = col
End Function

' ---------------- Utilities ------------------------------------------------
Private Function FindClassEnd(ByVal pat As String, ByVal p As Long, ByVal pEnd As Long) As Long
    Dim i As Long
    Dim esc As Boolean: esc = False
    For i = p To pEnd
        Dim ch As String: ch = SafeMid(pat, i, 1)
        If esc Then
            esc = False
        ElseIf ch = "\" Then
            esc = True
        ElseIf ch = "]" Then
            FindClassEnd = i: Exit Function
        End If
    Next i
    FindClassEnd = -1
End Function

Private Function ComputeMinLen(ByVal node As Collection) As Long
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "lit", "dot", "class": ComputeMinLen = 1
        Case "seq"
            Dim sum As Long: sum = 0
            Dim ch As Variant
            For Each ch In NodeGet(node, "children")
                sum = sum + ComputeMinLen(ch)
            Next
            ComputeMinLen = sum
        Case "group"
            ComputeMinLen = ComputeMinLen(NodeGet(node, "child"))
        Case "quant"
            Dim mn As Long: mn = NodeGet(node, "min")
            If mn < 0 Then mn = 0
            ComputeMinLen = mn * ComputeMinLen(NodeGet(node, "child"))
        Case "alt"
            Dim minv As Long: minv = -1
            Dim chAlt As Variant
            For Each chAlt In NodeGet(node, "children")
                Dim vAlt As Long: vAlt = ComputeMinLen(chAlt)
                If minv = -1 Or vAlt < minv Then minv = vAlt
            Next
            If minv = -1 Then minv = 0
            ComputeMinLen = minv
        Case Else: ComputeMinLen = 0
    End Select
End Function

' ---------------- Public Exec / Replace / Split ---------------------------
' Exec: find first match anywhere (respects ^ anchor and multiline semantics)
Public Function Exec(ByVal subj As String) As Collection
    CurSubject = subj
    MatchSteps = 0
    Dim anchoredStart As Boolean: anchoredStart = NodeGet(pAST, "anchoredStart")
    Dim L As Long: L = Len(subj)
    Dim minLen As Long
    minLen = NodeGet(pAST, "minLen")

    ' Anchored start: if multiline, try positions that are start-of-string or immediately after a newline (CR or LF)
    If anchoredStart Then
        If pMultiline Then
            Dim tryPos As Long
            For tryPos = 1 To L
                If tryPos = 1 Or SafeMid(subj, tryPos - 1, 1) = vbLf Or SafeMid(subj, tryPos - 1, 1) = vbCr Then
                    If Not (ExecAt(subj, tryPos) Is Nothing) Then
                        Set Exec = BuildCaptureCollection()
                        Exit Function
                    End If
                End If
                If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
            Next tryPos
            Set Exec = Nothing
            Exit Function
        Else
            If ExecAt(subj, 1) Is Nothing Then
                Set Exec = Nothing
            Else
                Set Exec = BuildCaptureCollection()
            End If
            Exit Function
        End If
    End If

    Dim startLimit As Long
    If minLen > 0 Then startLimit = L - minLen + 1 Else startLimit = L
    If startLimit < 1 Then startLimit = 1
    Dim i As Long
    For i = 1 To startLimit
        If ExecAt(subj, i) Is Nothing Then
            ' no match at this start
        Else
            Set Exec = BuildCaptureCollection()
            Exit Function
        End If
        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
    Next i
    Set Exec = Nothing
End Function

' ExecAt: try to match body at absolute position startPos. Returns Collection (captures filled) or Nothing
Public Function ExecAt(ByVal subj As String, ByVal startPos As Long) As Collection
    If DEBUG_TRACE Then Debug.Print "TRACE ExecAt startPos=" & startPos & " subjLen=" & Len(subj) & " multiline=" & CStr(pMultiline) & " dotAll=" & CStr(pDotAll)
    CurSubject = subj
    MatchSteps = MatchSteps + 1

    Dim totalGroups As Long: totalGroups = pGroupCount
    ReDim capsStart(0 To totalGroups) As Long
    ReDim capsEnd(0 To totalGroups) As Long
    ' Initialize atomicLocks for this ExecAt attempt (no atomic locks initially)
    ReDim atomicLocks(0 To totalGroups) As Boolean
    Dim i As Long
    For i = 0 To totalGroups
        capsStart(i) = 0: capsEnd(i) = -1
        atomicLocks(i) = False
    Next i
    capsStart(0) = startPos
    If DEBUG_TRACE Then Debug.Print "TRACE ExecAt caps initialized caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"

    Dim outEnd As Variant
    Dim ok As Boolean
    ok = MatchNode(NodeGet(pAST, "body"), subj, startPos, outEnd)

    If DEBUG_TRACE Then Debug.Print "TRACE ExecAt after MatchNode ok=" & CStr(ok) & " outEnd=" & CStr(outEnd) & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"

    If ok Then
        ' if outEnd available, ensure capsEnd(0) is synchronized
        If Not IsEmpty(outEnd) Then capsEnd(0) = CLng(outEnd)
        If DEBUG_TRACE Then Debug.Print "TRACE ExecAt sync caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"

        ' anchored-end check
        If NodeGet(pAST, "anchoredEnd") Then
            Dim matchEndPos As Long: matchEndPos = capsEnd(0)
            Dim subjLen As Long: subjLen = Len(subj)
            If DEBUG_TRACE Then Debug.Print "TRACE anchoredEnd: matchEndPos=" & matchEndPos & " subjLen=" & subjLen
            Dim okEnd As Boolean: okEnd = False
            If matchEndPos < 0 Then
                If DEBUG_TRACE Then Debug.Print "TRACE anchoredEnd: bad matchEndPos, reject"
                Set ExecAt = Nothing: Exit Function
            End If
            If matchEndPos = subjLen Then
                okEnd = True
            ElseIf pMultiline Then
                If matchEndPos < subjLen Then
                    Dim nextCh As String: nextCh = SafeMid(subj, matchEndPos + 1, 1)
                    If nextCh = vbLf Or nextCh = vbCr Then okEnd = True
                    If matchEndPos + 1 < subjLen Then
                        If SafeMid(subj, matchEndPos + 1, 2) = vbCr & vbLf Then okEnd = True
                    End If
                End If
            End If
            If Not okEnd Then
                If DEBUG_TRACE Then Debug.Print "TRACE anchoredEnd check FAILED -> returning Nothing"
                Set ExecAt = Nothing: Exit Function
            Else
                If DEBUG_TRACE Then Debug.Print "TRACE anchoredEnd check OK"
            End If
        End If

        If DEBUG_TRACE Then Debug.Print "TRACE ExecAt returning captures caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
        Set ExecAt = BuildCaptureCollection()
    Else
        If DEBUG_TRACE Then Debug.Print "TRACE ExecAt MatchNode returned False -> no match"
        Set ExecAt = Nothing
    End If
End Function

' Replace and Split now use ExecAt to avoid SafeMid/InStr double-work
Public Function Replace(ByVal subj As String, ByVal replacement As String) As String
    Dim out As String: out = ""
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    Do
        Dim cs As Collection
        Set cs = ExecAt(subj, pos)
        If cs Is Nothing Then
            out = out & SafeMid(subj, pos)
            Exit Do
        Else
            Dim mstart As Long: mstart = capsStart(0)
            Dim mend As Long: mend = capsEnd(0)
            If mstart < pos Then mstart = pos
            out = out & SafeMid(subj, pos, mstart - pos)
            out = out & ExpandReplacement(cs, replacement)
            pos = mend + 1
            If pos > L Then Exit Do
        End If
    Loop
    Replace = out
End Function

Public Function Split(ByVal subj As String) As Collection
    Dim res As New Collection
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    Do
        Dim cs As Collection
        Set cs = ExecAt(subj, pos)
        If cs Is Nothing Then
            res.Add SafeMid(subj, pos)
            Exit Do
        Else
            Dim mstart As Long: mstart = capsStart(0)
            Dim mend As Long: mend = capsEnd(0)
            If mstart < pos Then mstart = pos
            res.Add SafeMid(subj, pos, mstart - pos)
            pos = mend + 1
            If pos > L Then
                res.Add ""
                Exit Do
            End If
        End If
    Loop
    Set Split = res
End Function

Private Function SafeMid(aStr As String, aStart As Long, Optional aLength As Long = -1) As String
    If aLength = -1 Then aLength = Len(aStr) - aStart + 1
    If aStart >= 1 Then
        SafeMid = mid$(aStr, aStart, aLength)
    Else
        SafeMid = vbNullString
    End If
End Function

Private Function ComputeMaxLen(ByVal node As Collection) As Long
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "lit", "dot", "class": ComputeMaxLen = 1
        Case "seq"
            Dim sum As Long: sum = 0
            Dim ch As Variant
            For Each ch In NodeGet(node, "children")
                Dim m As Long: m = ComputeMaxLen(ch)
                If m = -1 Then
                    ComputeMaxLen = -1: Exit Function
                End If
                sum = sum + m
            Next
            ComputeMaxLen = sum
        Case "group"
            ComputeMaxLen = ComputeMaxLen(NodeGet(node, "child"))
        Case "alt"
            Dim mx As Long: mx = 0
            Dim itm As Variant
            For Each itm In NodeGet(node, "children")
                Dim vm As Long: vm = ComputeMaxLen(itm)
                If vm = -1 Then
                    mx = -1: Exit For
                ElseIf vm > mx Then mx = vm
                End If
            Next
            ComputeMaxLen = mx
        Case "quant"
            Dim mn As Long: mn = NodeGet(node, "min")
            Dim mxChild As Long: mxChild = ComputeMaxLen(NodeGet(node, "child"))
            If mxChild = -1 Then
                If NodeGet(node, "max") = -1 Then
                    ComputeMaxLen = -1: Exit Function
                Else
                    ' child unbounded but quant bounded by explicit max
                End If
            End If
            Dim qmax As Long: qmax = NodeGet(node, "max")
            If qmax = -1 Then
                If mxChild = -1 Then
                    ComputeMaxLen = -1
                Else
                    ComputeMaxLen = -1
                End If
            Else
                If mxChild = -1 Then
                    ComputeMaxLen = -1
                Else
                    ComputeMaxLen = qmax * mxChild
                End If
            End If
        Case "lookahead_pos", "lookahead_neg"
            ComputeMaxLen = ComputeMaxLen(NodeGet(node, "child"))
        Case "lookbehind_pos", "lookbehind_neg"
            ComputeMaxLen = ComputeMaxLen(NodeGet(node, "child"))
        Case "atomic"
            ComputeMaxLen = ComputeMaxLen(NodeGet(node, "child"))
        Case Else
            ComputeMaxLen = -1
    End Select
End Function

Private Sub AnnotateLookbehinds(ByVal node As Collection)
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "seq"
            Dim ch As Variant
            For Each ch In NodeGet(node, "children")
                AnnotateLookbehinds ch
            Next
        Case "group", "atomic", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
            If NodeExists(node, "child") Then
                AnnotateLookbehinds NodeGet(node, "child")
            End If
            If t = "lookbehind_pos" Or t = "lookbehind_neg" Then
                Dim mn As Long: mn = ComputeMinLen(NodeGet(node, "child"))
                Dim mx As Long: mx = ComputeMaxLen(NodeGet(node, "child"))
                If mn = -1 Or mx = -1 Or mn <> mx Then
                    err.Raise vbObjectError + 10, "RegexEngine", "Unsupported variable-length lookbehind (must be fixed-width)."
                Else
                    NodeAdd node, "width", mn
                End If
            End If
        Case "quant"
            AnnotateLookbehinds NodeGet(node, "child")
        Case "alt"
            Dim it As Variant
            For Each it In NodeGet(node, "children")
                AnnotateLookbehinds it
            Next
    End Select
End Sub

