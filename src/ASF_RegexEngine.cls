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
'  - Non-capturing groups: (?:pattern)
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
Option Base 0

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
Private pNamedCaptures As Collection

Public Property Get LastNamedCaptures() As Collection
    Set LastNamedCaptures = pNamedCaptures
End Property
' ---------------- Node helpers ------------------------------------------
Private Function newNode(ByVal typ As String) As Collection
    Dim c As New Collection
    c.Add typ, "type"
    Set newNode = c
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
Public Sub Init(ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False, Optional ByVal MaxSteps As Long = DEFAULT_MAX_STEPS, Optional ByVal multiline As Boolean = False, Optional ByVal dotAll As Boolean = False)
    pPattern = pattern
    pIgnoreCase = ignoreCase
    pMultiline = multiline
    pDotAll = dotAll
    If MaxSteps <= 0 Then pMaxMatchSteps = DEFAULT_MAX_STEPS Else pMaxMatchSteps = MaxSteps

    ' Build AST
    Set pAST = CompilePattern(pattern)
    If Not pAST Is Nothing Then
        If Not NodeExists(pAST, "groupNodes") Then BuildGroupNodesFromAST pAST
    End If

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
    ' Reject unsupported inline-comment syntax early
    DetectUnsupportedCommentSyntax pat
    
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
    
    ' enforce lookbehind fixed-width constraint (raise at Init/Compile time)
    EnforceFixedWidthLookbehinds ast

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

    Dim altNode As Collection: Set altNode = newNode("alt")
    Dim alts As New Collection
    alts.Add firstSeq

    Do While p <= pEnd And SafeMid(pat, p, 1) = "|"
        p = p + 1 ' consume '|'
        Dim nextSeq As Collection
        Set nextSeq = ParseSequence(pat, p, pEnd, groupCounter)
        alts.Add nextSeq
    Loop

    NodeAdd altNode, "child", alts
    Set ParseAlternation = altNode
End Function

Private Function ParseSequence(pat As String, ByRef p As Long, pEnd As Long, _
                            ByRef groupCounter As Long, Optional stopAtPipeOrParen As Boolean = False) As Collection
    Dim seq As Collection: Set seq = newNode("seq")
    Dim children As New Collection
    Dim thisIdx As Long
    Dim ch As String
    Dim node As Collection

    Do While p <= pEnd
        ch = SafeMid(pat, p, 1)
        If ch = ")" Or ch = "|" Then Exit Do
        Set node = Nothing

        If ch = "\" Then
            p = p + 1
            If p > pEnd Then Exit Do
            Set node = ParseEscapeNode(SafeMid(pat, p, 1))
            p = p + 1

        ElseIf ch = "." Then
            Set node = newNode("dot")
            p = p + 1

        ElseIf ch = "[" Then
            Dim clsEnd As Long: clsEnd = FindClassEnd(pat, p, pEnd)
            If clsEnd = -1 Then
                Set node = newNode("lit"): NodeAdd node, "ch", "[": p = p + 1
            Else
                Set node = ParseClass(pat, p + 1, clsEnd - 1)
                p = clsEnd + 1
            End If

        ElseIf ch = "(" Then
            ' locate matching ')' deterministically
            Dim foundClose As Long
            foundClose = FindGroupEnd(pat, p, pEnd)
            If foundClose = -1 Then
                ' malformed -> treat '(' as literal
                Set node = newNode("lit"): NodeAdd node, "ch", "("
                p = p + 1
            Else
                ' consume '('
                p = p + 1
                If p <= pEnd And SafeMid(pat, p, 1) = "?" Then
                    p = p + 1
                    If p > pEnd Then
                        Set node = newNode("lit"): NodeAdd node, "ch", "("
                    Else
                        Dim nextCh As String: nextCh = SafeMid(pat, p, 1)
                        Select Case nextCh
                            Case "="
                                p = p + 1
                                Dim innerLA As Collection
                                Set innerLA = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                Dim laNode As Collection: Set laNode = newNode("lookahead_pos")
                                NodeAdd laNode, "child", innerLA
                                Set node = laNode

                            Case "!"
                                p = p + 1
                                Dim innerNL As Collection
                                Set innerNL = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                Dim nlaNode As Collection: Set nlaNode = newNode("lookahead_neg")
                                NodeAdd nlaNode, "child", innerNL
                                Set node = nlaNode

                            Case "<"
                                p = p + 1
                                If p <= pEnd Then
                                    Dim sbch As String: sbch = SafeMid(pat, p, 1)
                                    If sbch = "=" Then
                                        p = p + 1
                                        Dim innerLB As Collection
                                        Set innerLB = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                        If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                        Dim lbNode As Collection: Set lbNode = newNode("lookbehind_pos")
                                        NodeAdd lbNode, "child", innerLB
                                        Set node = lbNode
                                    ElseIf sbch = "!" Then
                                        p = p + 1
                                        Dim innerLBN As Collection
                                        Set innerLBN = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                        If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                        Dim lbnNode As Collection: Set lbnNode = newNode("lookbehind_neg")
                                        NodeAdd lbnNode, "child", innerLBN
                                        Set node = lbnNode
                                    Else
                                        ' named capture form (?<name>...)
                                        Dim nameBuf As String: nameBuf = vbNullString
                                        Do While p <= pEnd
                                            Dim nc As String: nc = SafeMid(pat, p, 1)
                                            If nc = ">" Then Exit Do
                                            If (nc >= "a" And nc <= "z") Or (nc >= "A" And nc <= "Z") Or (nc >= "0" And nc <= "9") Or nc = "_" Or nc = "-" Then
                                                nameBuf = nameBuf & nc
                                                p = p + 1
                                            Else
                                                Exit Do
                                            End If
                                        Loop
                                        If p <= pEnd And SafeMid(pat, p, 1) = ">" And Len(nameBuf) > 0 Then
                                            p = p + 1 ' consume '>'
                                            groupCounter = groupCounter + 1
                                            thisIdx = groupCounter
                                            Dim innerNamed As Collection
                                            Set innerNamed = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                            If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                            Dim gnodeNamed As Collection: Set gnodeNamed = newNode("group")
                                            NodeAdd gnodeNamed, "idx", thisIdx
                                            NodeAdd gnodeNamed, "child", innerNamed
                                            NodeAdd gnodeNamed, "name", nameBuf
                                            Set node = gnodeNamed
                                            If DEBUG_TRACE Then Debug.Print "TRACE PARSE named-group -> name=" & nameBuf & " idx=" & thisIdx
                                        Else
                                            Set node = newNode("lit"): NodeAdd node, "ch", "("
                                        End If
                                    End If
                                Else
                                    Set node = newNode("lit"): NodeAdd node, "ch", "("
                                End If

                            Case ">"
                                p = p + 1
                                Dim innerAtomic As Collection
                                Set innerAtomic = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                Dim atomicNode As Collection: Set atomicNode = newNode("atomic")
                                NodeAdd atomicNode, "child", innerAtomic
                                Set node = atomicNode

                            Case ":"
                                p = p + 1
                                Dim innerNC As Collection
                                Set innerNC = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                Dim ncNode As Collection: Set ncNode = newNode("group")
                                NodeAdd ncNode, "child", innerNC
                                NodeAdd ncNode, "noncap_syntax", True
                                Set node = ncNode

                            Case "R"
                                p = p + 1
                                If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                Dim rnode As Collection: Set rnode = newNode("recurse")
                                NodeAdd rnode, "target", 0
                                Set node = rnode

                            Case "("
                                ' conditional (?(cond)yes|no)
                                p = p + 1
                                Dim condStart As Long: condStart = p
                                Dim condEnd As Long: condEnd = -1
                                Do While p <= pEnd
                                    If SafeMid(pat, p, 1) = ")" Then
                                        condEnd = p - 1
                                        Exit Do
                                    End If
                                    p = p + 1
                                Loop
                                If condEnd = -1 Then
                                    Set node = newNode("lit"): NodeAdd node, "ch", "("
                                Else
                                    Dim condText As String: condText = mid$(pat, condStart, condEnd - condStart + 1)
                                    p = p + 1 ' consume ')'
                                    Dim yesNode As Collection: Set yesNode = ParseSequence(pat, p, foundClose - 1, groupCounter, True)
                                    Dim noNode As Collection: Set noNode = Nothing
                                    If p <= pEnd And SafeMid(pat, p, 1) = "|" Then
                                        p = p + 1
                                        Set noNode = ParseSequence(pat, p, foundClose - 1, groupCounter, True)
                                    End If
                                    If foundClose <= pEnd Then p = foundClose + 1
                                    Dim condNode As Collection: Set condNode = newNode("cond")
                                    NodeAdd condNode, "cond", condText
                                    NodeAdd condNode, "yes", yesNode
                                    If Not (noNode Is Nothing) Then NodeAdd condNode, "no", noNode
                                    Set node = condNode
                                End If

                            Case Else
                                ' recursion by number or inline flags
                                If SafeMid(pat, p, 1) >= "0" And SafeMid(pat, p, 1) <= "9" Then
                                    Dim ds As String: ds = vbNullString
                                    Do While p <= pEnd
                                        Dim dc As String: dc = SafeMid(pat, p, 1)
                                        If dc >= "0" And dc <= "9" Then
                                            ds = ds & dc
                                            p = p + 1
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                    If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                    Dim rn As Collection: Set rn = newNode("recurse")
                                    NodeAdd rn, "target", CLng(ds)
                                    Set node = rn
                                Else
                                    Dim flagsStr As String: flagsStr = vbNullString
                                    Do While p <= pEnd
                                        Dim fch As String: fch = SafeMid(pat, p, 1)
                                        If (fch >= "a" And fch <= "z") Or (fch >= "A" And fch <= "Z") Or fch = "-" Then
                                            flagsStr = flagsStr & fch
                                            p = p + 1
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                    If p <= pEnd And SafeMid(pat, p, 1) = ":" Then
                                        p = p + 1
                                        Dim innerFlagged As Collection
                                        Set innerFlagged = ParseSequence(pat, p, foundClose - 1, groupCounter)
                                        If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                                        Dim gnodeF As Collection: Set gnodeF = newNode("group")
                                        NodeAdd gnodeF, "noncap_syntax", True
                                        Dim normFlags As String: normFlags = vbNullString
                                        Dim kch As Long
                                        For kch = 1 To Len(flagsStr)
                                            Dim fc As String: fc = LCase$(mid$(flagsStr, kch, 1))
                                            If fc = "i" Or fc = "m" Or fc = "s" Or fc = "-" Then normFlags = normFlags & fc
                                        Next kch
                                        If Len(normFlags) > 0 Then
                                            NodeAdd gnodeF, "flags_on", normFlags
                                            PropagateFlags innerFlagged, normFlags
                                        End If
                                        NodeAdd gnodeF, "child", innerFlagged
                                        Set node = gnodeF
                                    Else
                                        err.Raise vbObjectError + 101, "RegexEngine", "Inline flags without ':' are not supported at compile time (use (?i:...) for group-scoped flags)."
                                    End If
                                End If
                        End Select
                    End If
                Else
                    ' normal capturing group
                    groupCounter = groupCounter + 1
                    thisIdx = groupCounter
                    Dim inner As Collection
                    Set inner = ParseSequence(pat, p, foundClose - 1, groupCounter)
                    If p <= pEnd And SafeMid(pat, p, 1) = ")" Then p = p + 1
                    Dim gnode As Collection: Set gnode = newNode("group")
                    NodeAdd gnode, "idx", thisIdx
                    NodeAdd gnode, "child", inner
                    Set node = gnode
                End If

                ' defensive: if inner parsing didn't advance to foundClose, jump past it
                If foundClose <> -1 Then
                    If p <= foundClose And SafeMid(pat, p, 1) <> ")" Then p = foundClose + 1
                End If
            End If

        Else
            Set node = newNode("lit"): NodeAdd node, "ch", ch
            p = p + 1
        End If

        ' quantifier (same semantics as original)
        If Not node Is Nothing And p <= pEnd Then
            Dim qch As String: qch = SafeMid(pat, p, 1)
            If qch = "*" Or qch = "+" Or qch = "?" Or qch = "{" Then
                Dim qnode As Collection: Set qnode = newNode("quant")
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

        If Not node Is Nothing Then children.Add node
    Loop

    NodeAdd seq, "child", children

    If Not stopAtPipeOrParen Then
        If p <= pEnd And SafeMid(pat, p, 1) = "|" Then
            Dim altNode As Collection: Set altNode = newNode("alt")
            Dim altChildren As New Collection
            altChildren.Add seq
            Do While p <= pEnd And SafeMid(pat, p, 1) = "|"
                p = p + 1
                Dim nextSeq As Collection
                Set nextSeq = ParseSequence(pat, p, pEnd, groupCounter)
                altChildren.Add nextSeq
            Loop
            NodeAdd altNode, "child", altChildren
            Set ParseSequence = altNode
            Exit Function
        End If
    End If

    Set ParseSequence = seq
End Function

Private Sub PropagateFlags(ByVal node As Variant, ByRef flags As String)
    If node Is Nothing Then Exit Sub
    On Error GoTo err_
    
    If TypeName(node) = "Collection" Then
        Dim t As String
        Dim elm As Variant
        For Each elm In node
            If TypeName(elm) = "Collection" Then
                If NodeExists(elm, "type") Then t = NodeGet(elm, "type")
                Select Case t
                    Case "lit", "dot"
                        Dim tmpNode As Collection
                        Set tmpNode = elm
                        NodeAdd tmpNode, "flags_on", flags
                End Select
                If NodeExists(elm, "child") Then
                    PropagateFlags NodeGet(elm, "child"), flags
                Else
                    PropagateFlags elm, flags
                End If
            End If
        Next elm
    End If
exit_:
    Exit Sub
err_:
    ' Be defensive in production: clear error and continue (or rethrow in debug).
    Debug.Print "PropagateFlags error: "; err.Number; err.Description
    err.Clear
    Resume exit_
End Sub

' ---------- Helper: Find matching ] for a class starting at startPos (where pat[startPos] = "[") ----------
Private Function FindClassEnd(ByVal pat As String, ByVal startPos As Long, ByVal pEnd As Long) As Long
    Dim pos As Long: pos = startPos
    Dim escaped As Boolean: escaped = False
    Dim ch As String
    ' We expect pat(startPos) = "["
    pos = pos + 1
    Do While pos <= pEnd
        ch = SafeMid(pat, pos, 1)
        If escaped Then
            escaped = False
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = "\" Then
            escaped = True
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = "]" Then
            FindClassEnd = pos
            Exit Function
        End If
        pos = pos + 1
ContinueDo:
    Loop
    FindClassEnd = -1
End Function

' ---------- Helper: Find matching ) for a group starting at startPos (where pat[startPos] = "(") ----------
Private Function FindGroupEnd(ByVal pat As String, ByVal startPos As Long, ByVal pEnd As Long) As Long
    Dim pos As Long: pos = startPos
    Dim depth As Long: depth = 0
    Dim inClass As Boolean: inClass = False
    Dim escaped As Boolean: escaped = False
    Dim ch As String

    ' start scanning after the '('
    pos = pos + 1
    Do While pos <= pEnd
        ch = SafeMid(pat, pos, 1)
        If escaped Then
            escaped = False
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = "\" Then
            escaped = True
            pos = pos + 1
            GoTo ContinueDo
        End If
        If inClass Then
            If ch = "]" Then inClass = False
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = "[" Then
            inClass = True
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = "(" Then
            depth = depth + 1
            pos = pos + 1
            GoTo ContinueDo
        End If
        If ch = ")" Then
            If depth = 0 Then
                FindGroupEnd = pos
                Exit Function
            Else
                depth = depth - 1
                pos = pos + 1
                GoTo ContinueDo
            End If
        End If
        pos = pos + 1
ContinueDo:
    Loop

    FindGroupEnd = -1
End Function

Private Function ParseEscapeNode(ByVal ch As String) As Collection
    Dim node As Collection
    Select Case ch
        Case "d": Set node = ParseClassFromSimple("[0-9]")
        Case "D"
            ' Non-digit escape: reuse the simple class and mark it negated
            Set node = ParseClassFromSimple("[0-9]")
            NodeAdd node, "neg", True
        Case "w": Set node = ParseClassFromSimple("[A-Za-z0-9_]")
        Case "W"
            ' Non-word
            Set node = ParseClassFromSimple("[A-Za-z0-9_]")
            NodeAdd node, "neg", True
        Case "s": Set node = ParseClassFromSimple("[ " & vbTab & vbLf & vbCr & "]")
        Case "S"
            ' Non-space
            Set node = ParseClassFromSimple("[ " & vbTab & vbLf & vbCr & "]")
            NodeAdd node, "neg", True
        Case "n": Set node = newNode("lit"): NodeAdd node, "ch", vbLf
        Case "r": Set node = newNode("lit"): NodeAdd node, "ch", vbCr
        Case "t": Set node = newNode("lit"): NodeAdd node, "ch", vbTab
        Case Else: Set node = newNode("lit"): NodeAdd node, "ch", ch
    End Select
    Set ParseEscapeNode = node
End Function


' Parse class and build a 256-byte boolean map for quick matching (ASCII fast-path)
Private Function ParseClass(ByVal pat As String, ByVal pStart As Long, ByVal pEnd As Long) As Collection
    Dim node As Collection: Set node = newNode("class")
    Dim Chars As New Collection
    Dim neg As Boolean: neg = False
    Dim i As Long: i = pStart
    If i <= pEnd And SafeMid(pat, i, 1) = "^" Then neg = True: i = i + 1
    Do While i <= pEnd
        Dim ch As String: ch = SafeMid(pat, i, 1)
        If ch = "\" Then
            i = i + 1
            If i <= pEnd Then
                Dim esc As String: esc = SafeMid(pat, i, 1)
                Select Case esc
                    Case "d"
                        ' digits
                        Chars.Add Array("range", "0", "9")
                    Case "w"
                        ' word chars: A-Z, a-z, 0-9, '_'
                        Chars.Add Array("range", "A", "Z")
                        Chars.Add Array("range", "a", "z")
                        Chars.Add Array("range", "0", "9")
                        Chars.Add "_"
                    Case "s"
                        ' whitespace
                        Chars.Add " "
                        Chars.Add vbTab
                        Chars.Add vbLf
                        Chars.Add vbCr
                    Case Else
                        ' fallback (n, r, t, or literal)
                        Chars.Add ParseEscapeChar(esc)
                End Select
            End If
            i = i + 1
        ElseIf i + 2 <= pEnd And SafeMid(pat, i + 1, 1) = "-" Then
            Chars.Add Array("range", SafeMid(pat, i, 1), SafeMid(pat, i + 2, 1))
            i = i + 3
        Else
            Chars.Add ch
            i = i + 1
        End If
    Loop
    NodeAdd node, "chars", Chars
    NodeAdd node, "neg", neg
    ' Precompute ASCII map
    Dim mapArr() As Boolean: ReDim mapArr(0 To 255) As Boolean
    Dim it As Variant
    For Each it In Chars
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
    Dim node As Collection: Set node = newNode("class")
    Dim Chars As New Collection
    Dim i As Long: i = 2
    Do While i <= Len(spec) - 1
        Dim ch As String: ch = SafeMid(spec, i, 1)
        If i + 2 <= Len(spec) - 1 And SafeMid(spec, i + 1, 1) = "-" Then
            Chars.Add Array("range", ch, SafeMid(spec, i + 2, 1))
            i = i + 3
        Else
            Chars.Add ch
            i = i + 1
        End If
    Loop
    NodeAdd node, "chars", Chars
    NodeAdd node, "neg", False
    ' build simple map
    Dim mapArr() As Boolean: ReDim mapArr(0 To 255) As Boolean
    Dim it As Variant
    For Each it In Chars
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

' Deterministic traversal for capturing maximum group index in an AST node tree.
Private Sub TraverseForGroups(ByVal node As Variant, ByRef maxIdx As Long)
    If node Is Nothing Then Exit Sub
    On Error GoTo err_
    
    If TypeName(node) = "Collection" Then
        Dim t As String
        Dim elm As Variant
        Dim tmpIdx As Long
        
        If NodeExists(node, "type") Then t = NodeGet(node, "type")
        If t = "group" Then
            If NodeExists(node, "idx") Then
                tmpIdx = NodeGet(node, "idx")
                If tmpIdx > maxIdx Then maxIdx = tmpIdx
            End If
        End If
        For Each elm In node
            If TypeName(elm) = "Collection" Then
                If NodeExists(elm, "type") Then t = NodeGet(elm, "type")
                If t = "group" Then
                    If NodeExists(elm, "idx") Then
                        tmpIdx = NodeGet(elm, "idx")
                        If tmpIdx > maxIdx Then maxIdx = tmpIdx
                    End If
                End If
                If NodeExists(elm, "child") Then
                    TraverseForGroups NodeGet(elm, "child"), maxIdx
                Else
                    TraverseForGroups elm, maxIdx
                End If
            End If
        Next elm
    End If
exit_:
    Exit Sub
err_:
    ' Be defensive in production: clear error and continue (or rethrow in debug).
    If DEBUG_TRACE Then Debug.Print "TraverseForGroups error: "; err.Number; err.Description
    err.Clear
    Resume exit_
End Sub

Private Function MatchNode(ByVal node As Collection, ByVal subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    If DEBUG_TRACE Then Debug.Print "TRACE MatchNode ENTER type=" & CStr(NodeGet(node, "type")) & " pos=" & pos & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
    MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"

    Dim t As String: t = CStr(NodeGet(node, "type"))
    Dim res As Boolean: res = False
    Dim localEnd As Variant: localEnd = Empty
    Dim inlineFlags As String
    
    Select Case t
        Case "seq"
            res = MatchSequence(NodeGet(node, "child"), subj, pos, localEnd)

        Case "lit"
            If pos > Len(subj) Then
                res = False
            Else
                Dim ch As String: ch = NodeGet(node, "ch")
                Dim sc As String: sc = SafeMid(subj, pos, 1)
                Dim iCase As Boolean
                Dim tOffCase As Boolean
                If pIgnoreCase Then
                    'Inline flags
                    If NodeExists(node, "flags_on") Then
                        inlineFlags = NodeGet(node, "flags_on")
                        tOffCase = InStrB(1, inlineFlags, "-i")
                    End If
                    If Not tOffCase Then
                        If LCase$(sc) = LCase$(ch) Then localEnd = pos: res = True
                    Else
                        ' Case sentitive
                        If sc = ch Then localEnd = pos: res = True
                    End If
                Else
                    If NodeExists(node, "flags_on") Then
                        inlineFlags = NodeGet(node, "flags_on")
                        iCase = InStrB(1, inlineFlags, "i") And (InStrB(1, inlineFlags, "-i") = 0)
                        If iCase Then
                            ' Case insensitive
                            If LCase$(sc) = LCase$(ch) Then localEnd = pos: res = True
                        Else
                            If sc = ch Then localEnd = pos: res = True
                        End If
                    Else
                        If sc = ch Then localEnd = pos: res = True
                    End If
                End If
            End If

        Case "dot"
            If pos > Len(subj) Then
                res = False
            Else
                Dim dAll As Boolean
                If NodeExists(node, "flags_on") Then
                    inlineFlags = NodeGet(node, "flags_on")
                    tOffCase = (InStrB(1, inlineFlags, "-s") > 0)
                    dAll = InStrB(1, inlineFlags, "s") And (InStrB(1, inlineFlags, "-s") = 0)
                End If
                If pDotAll And tOffCase Then
                    dAll = False
                Else
                    If Not dAll Then dAll = pDotAll
                End If
                Dim dch As String: dch = SafeMid(subj, pos, 1)
                If (dch = vbLf Or dch = vbCr) And Not dAll Then
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
            ' Handle both capturing and non-capturing groups.
            Dim hasIdx As Boolean: hasIdx = NodeExists(node, "idx")
            Dim gidx As Long
            If hasIdx Then gidx = CLng(NodeGet(node, "idx")) Else gidx = 0
        
            ' Support group-scoped inline flags (flags_on) during the child match.
            Dim flagsApplied As Boolean: flagsApplied = False
            Dim savedIgnore As Boolean, savedMultiline As Boolean, savedDotAll As Boolean
            If NodeExists(node, "flags_on") Then
                flagsApplied = True
                savedIgnore = pIgnoreCase
                savedMultiline = IIf(NodeExists(Me, "pMultiline"), pMultiline, False)
                savedDotAll = IIf(NodeExists(Me, "pDotAll"), pDotAll, False)
                Dim fstr As String: fstr = CStr(NodeGet(node, "flags_on"))
                If InStr(fstr, "i") > 0 Then pIgnoreCase = True
                If InStr(fstr, "m") > 0 Then pMultiline = True
                If InStr(fstr, "s") > 0 Then pDotAll = True
            End If
        
            ' --- FULL snapshot before touching ANY capture slots ---
            Dim snapGroup() As Long: snapGroup = CopyCaps()
        
            ' set group's start (only after we took full snapshot)
            If gidx >= 1 Then
                capsStart(gidx) = pos
            End If
        
            Dim childEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, childEnd) Then
                If gidx >= 1 Then capsEnd(gidx) = CLng(childEnd)
                localEnd = childEnd
                res = True
            Else
                ' restore entire capture state (safer than just restoring this group's two slots)
                RestoreCaps snapGroup
                res = False
            End If
        
            ' restore flags if applied (always do this after child path)
            If flagsApplied Then
                pIgnoreCase = savedIgnore
                pMultiline = savedMultiline
                pDotAll = savedDotAll
            End If
        Case "alt"
            Dim altChild As Variant
            For Each altChild In NodeGet(node, "child")
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
            If mx = -1 Then mx = Len(subj) - pos + 1
            ' Attempt greedy expansion and accept largest
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

        ' ---- Positive lookahead (?=...) ----
        Case "lookahead_pos"
            If Not NodeExists(node, "child") Then
                MatchNode = False: Exit Function
            End If

            Dim laChild As Collection: Set laChild = NodeGet(node, "child")
            Dim snapLA() As Long: snapLA = CopyCaps()
            Dim laEnd As Long
            Dim laMatched As Boolean

            ' try the inner pattern at current pos (non-consuming)
            laMatched = MatchNode(laChild, subj, pos, laEnd)

            ' always restore capture snapshot (non-consuming)
            RestoreCaps snapLA

            If laMatched Then
                outEnd = pos - 1          ' zero-width: leave curPos unchanged
                MatchNode = True
            Else
                MatchNode = False
            End If
            Exit Function

        ' ---- Negative lookahead (?!...) ----
        Case "lookahead_neg"
            If Not NodeExists(node, "child") Then
                MatchNode = False: Exit Function
            End If

            Dim lnChild As Collection: Set lnChild = NodeGet(node, "child")
            Dim snapLN() As Long: snapLN = CopyCaps()
            Dim lnEnd As Long
            Dim lnMatched As Boolean

            lnMatched = MatchNode(lnChild, subj, pos, lnEnd)

            ' restore (assertions do not consume)
            RestoreCaps snapLN

            If lnMatched Then
                MatchNode = False
            Else
                outEnd = pos - 1
                MatchNode = True
            End If
            Exit Function

        ' ---- Positive lookbehind (?<=...)  -- fixed-width only ----
        Case "lookbehind_pos"
            If Not NodeExists(node, "child") Or Not NodeExists(node, "width") Then
                MatchNode = False: Exit Function
            End If

            Dim lbChild As Collection: Set lbChild = NodeGet(node, "child")
            Dim lbWidth As Long: lbWidth = CLng(NodeGet(node, "width"))
            Dim snapLB() As Long
            Dim lbStart As Long, lbEnd As Long
            Dim lbMatched As Boolean

            ' not enough characters before pos -> fail
            If pos - 1 < lbWidth Then
                MatchNode = False: Exit Function
            End If

            lbStart = pos - lbWidth
            snapLB = CopyCaps()
            lbMatched = MatchNode(lbChild, subj, lbStart, lbEnd)
            RestoreCaps snapLB

            If lbMatched Then
                outEnd = pos - 1
                MatchNode = True
            Else
                MatchNode = False
            End If
            Exit Function

        ' ---- Negative lookbehind (?<!...)  -- fixed-width only ----
        Case "lookbehind_neg"
            If Not NodeExists(node, "child") Or Not NodeExists(node, "width") Then
                MatchNode = False: Exit Function
            End If

            Dim lbnChild As Collection: Set lbnChild = NodeGet(node, "child")
            Dim lbnWidth As Long: lbnWidth = CLng(NodeGet(node, "width"))
            Dim snapLBN() As Long
            Dim lbnStart As Long, lbnEnd As Long
            Dim lbnMatched As Boolean

            If pos - 1 < lbnWidth Then
                ' no room for inner to match -> negative succeeds
                outEnd = pos - 1
                MatchNode = True: Exit Function
            End If

            lbnStart = pos - lbnWidth
            snapLBN = CopyCaps()
            lbnMatched = MatchNode(lbnChild, subj, lbnStart, lbnEnd)
            RestoreCaps snapLBN

            If lbnMatched Then
                MatchNode = False
            Else
                outEnd = pos - 1
                MatchNode = True
            End If
            Exit Function

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
        Case "cond"
            ' Conditional node: fields "cond" (string like "1" or "name"), "yes" (node), optional "no" (node)
            If Not NodeExists(node, "cond") Then
                MatchNode = False: Exit Function
            End If
        
            Dim condText As String: condText = CStr(NodeGet(node, "cond"))
            Dim condSatisfied As Boolean: condSatisfied = False
        
            ' numeric?
            If IsNumeric(condText) Then
                Dim cIdx As Long: cIdx = CLng(condText)
                If cIdx >= 0 And cIdx <= UBound(capsEnd) Then
                    If capsEnd(cIdx) <> -1 Then condSatisfied = True
                End If
            Else
                ' named: find index by name
                Dim resolvedIdx As Long: resolvedIdx = FindGroupIndexByName(pAST, condText)
                If resolvedIdx <> -1 Then
                    If resolvedIdx >= 0 And resolvedIdx <= UBound(capsEnd) Then
                        If capsEnd(resolvedIdx) <> -1 Then condSatisfied = True
                    End If
                Else
                    condSatisfied = False
                End If
            End If
        
            ' Branch selection
            If condSatisfied Then
                If NodeExists(node, "yes") Then
                    Dim yesNode As Variant: vAssignVar yesNode, NodeGet(node, "yes")
                    Dim yesEnd As Variant
                    Dim snapY() As Long: snapY = CopyCaps()
                    If MatchNode(yesNode, subj, pos, yesEnd) Then
                        localEnd = yesEnd: res = True
                    Else
                        RestoreCaps snapY: res = False
                    End If
                Else
                    res = False
                End If
            Else
                If NodeExists(node, "no") Then
                    Dim noNode As Variant: vAssignVar noNode, NodeGet(node, "no")
                    Dim noEnd As Variant
                    Dim snapN() As Long: snapN = CopyCaps()
                    If MatchNode(noNode, subj, pos, noEnd) Then
                        localEnd = noEnd: res = True
                    Else
                        RestoreCaps snapN: res = False
                    End If
                Else
                    ' No 'no' branch -> treat as failure (consistent with your earlier choice)
                    res = False
                End If
            End If
        
        Case "recurse"
            If Not NodeExists(node, "target") Then
                MatchNode = False: Exit Function
            End If
        
            Dim targ As Long: targ = CLng(NodeGet(node, "target"))
            Dim recOk As Boolean: recOk = False
            Dim recEnd As Variant
        
            If targ = 0 Then
                ' (?R) -> recurse into whole root AST
                Dim snapR() As Long: snapR = CopyCaps()
                If MatchNode(NodeGet(pAST, "body"), subj, pos, recEnd) Then
                    recOk = True: localEnd = recEnd
                Else
                    RestoreCaps snapR: recOk = False
                End If
            Else
                ' recurse into group node child (if available)
                If NodeExists(pAST, "groupNodes") Then
                    Dim gnodes As Collection: Set gnodes = NodeGet(pAST, "groupNodes")
                    Dim rIdx As Long: rIdx = FindGroupIndexByOrdinal(pAST, targ)
                    On Error Resume Next
                    Dim targetNode As Variant
                    targetNode = gnodes(rIdx) ' index collection access
                    If err.Number <> 0 Then
                        err.Clear
                        targetNode = Empty
                    End If
                    On Error GoTo 0
        
                    If Not IsEmpty(targetNode) Then
                        If TypeName(targetNode(UBound(targetNode))) = "Collection" Then
                            Dim childNode As Collection: Set childNode = targetNode(UBound(targetNode))
                            Dim snapG() As Long: snapG = CopyCaps()
                            If MatchNode(childNode, subj, pos, recEnd) Then
                                recOk = True: localEnd = recEnd
                            Else
                                RestoreCaps snapG: recOk = False
                            End If
                        End If
                    End If
                End If
            End If
        
            If recOk Then
                res = True
            Else
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
    Dim seqSnap() As Long
    Dim rootSnap() As Long: rootSnap = CopyCaps()
    Dim grpSnap() As Long
    Dim tmpSnaps() As Long
    Dim curPos As Long: curPos = pos
    Dim i As Long

    For i = 1 To children.count
        Dim node As Collection: Set node = children(i)
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
            Dim occSnaps() As Variant   ' store snapshots (each an array of Long) after each successful occurrence
            Dim occCount As Long: occCount = 0
            Dim scanPos As Long: scanPos = curPos
            Do While occCount < mx
                Dim preSnap() As Long: preSnap = CopyCaps()
                Dim ndEnd As Variant
                If MatchNode(childNode, subj, scanPos, ndEnd) Then
                    occCount = occCount + 1
                    ReDim Preserve occEnds(1 To occCount)
                    occEnds(occCount) = CLng(ndEnd)
                    ' capture snapshot AFTER the successful match so we can deterministically restore it later
                    Dim afterSnap() As Long: afterSnap = CopyCaps()
                    ReDim Preserve occSnaps(1 To occCount)
                    occSnaps(occCount) = afterSnap
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
                ' restore capture state corresponding to occCount (if >0)
                tmpSnaps = occSnaps(occCount)
                If occCount >= 1 Then RestoreCaps tmpSnaps
                Dim posTryP As Long: posTryP = IIf(occCount >= 1, occEnds(occCount) + 1, curPos)
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
            Dim remChildren As Collection
            For k = kStart To kEnd Step kStep
                If DEBUG_TRACE Then Debug.Print "TRACE  quant trying k=" & k & " (mn=" & mn & " occCount=" & occCount & ")"
                ' restore sequence baseline before attempting this k
                RestoreCaps seqSnap
        
                ' If k>0 restore the saved snapshot for exactly k occurrences (deterministic)
                Dim posTry As Long
                If k >= 1 Then
                    ' restore the capture state after k occurrences
                    tmpSnaps = occSnaps(k)
                    RestoreCaps tmpSnaps
                    posTry = occEnds(k) + 1
                Else
                    posTry = curPos
                End If
        
                ' Now match the remainder *inline*
                Dim remSuccess As Boolean: remSuccess = True
                Dim j As Long
                Dim nxtEnd As Variant
                If i + 1 <= children.count Then
                    ' Create a sub-collection for remaining children
                    Set remChildren = New Collection
                    For j = i + 1 To children.count
                        remChildren.Add children(j)
                    Next j
                    
                    ' Match the remainder as a sequence (handles nested quantifiers correctly)
                    If Not MatchSequence(remChildren, subj, posTry, nxtEnd) Then
                        remSuccess = False
                    Else
                        ' Remainder matched successfully
                        posTry = CLng(nxtEnd) + 1
                    End If
                Else
                    ' No remainder children - success
                    remSuccess = True
                End If
        
                If remSuccess Then
                    outEnd = posTry - 1
                    capsEnd(0) = CLng(outEnd)
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant accepted k=" & k & " outEnd=" & outEnd
                    MatchSequence = True: Exit Function
                Else
                    If DEBUG_TRACE Then Debug.Print "TRACE  quant k=" & k & " remainder failed; will try next k"
                    ' continue to next k (seqSnap will be restored at loop top)
                End If
            Next k
        
            ' none of the k choices worked; restore seq baseline and fail
            If DEBUG_TRACE Then Debug.Print "TRACE  quant exhausted all k options -> fail"
            RestoreCaps seqSnap
            MatchSequence = False: Exit Function
            
        ElseIf nodeType = "group" Then
            Dim childEnd As Variant
            Dim gidx As Long
            ' guard: some group nodes may have idx=0 (non-capturing wrapper) — only touch capture arrays for idx>=1
            If NodeExists(node, "idx") Then
                gidx = CLng(NodeGet(node, "idx"))
            Else
                gidx = 0
            End If
            
            ' save sequence baseline snapshot
            grpSnap = CopyCaps()
            
            If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group idx=" & gidx & " set start=" & curPos

            If gidx >= 1 Then capsStart(gidx) = curPos

            If Not MatchNode(NodeGet(node, "child"), subj, curPos, childEnd) Then
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group FAILED, restoring seqSnap"
                ' restore everything (seqSnap contains the capture arrays)
                RestoreCaps grpSnap
                MatchSequence = False: Exit Function
            Else
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group OK nodeEnd=" & childEnd
                If gidx >= 1 Then capsEnd(gidx) = childEnd
                curPos = CLng(childEnd) + 1
            End If
        Else
            If Not MatchNode(node, subj, curPos, nodeEnd) Then
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " node FAILED, restoring seqSnap"
                RestoreCaps rootSnap
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
' Return a fresh, self-contained snapshot of the capture arrays.
' Snapshot layout (0-based):
'  indices 0 .. tot        => capsStart(0..tot)
'  indices tot+1 .. 2*(tot+1)-1 => capsEnd(0..tot)
Private Function CopyCaps() As Long()
    Dim tot As Long: tot = pGroupCount
    Dim snap() As Long
    ReDim snap(0 To 2 * tot + 1)    ' 2*(tot+1) elements
    Dim i As Long
    For i = 0 To tot
        snap(i) = capsStart(i)
        snap(i + tot + 1) = capsEnd(i)
    Next i
    CopyCaps = snap
End Function

' Restore capture arrays from a snapshot produced by CopyCaps.
' This routine is defensive: if the current pGroupCount differs from the
' snapshot's pGroupCount, it restores as many slots as possible and
' initializes any new slots to sensible defaults.
Private Sub RestoreCaps(ByRef snapshot() As Long)
    On Error GoTo bad_snapshot
    If (Not (IsArray(snapshot))) Then Exit Sub

    Dim ub As Long: ub = UBound(snapshot)
    If ub < 0 Then Exit Sub

    ' compute snapshot tot: (UB - 1) \ 2
    Dim snapTot As Long
    snapTot = (ub - 1) \ 2
    Dim snapOff As Long: snapOff = snapTot + 1

    Dim curTot As Long: curTot = pGroupCount

    Dim i As Long
    For i = 0 To curTot
        If i <= snapTot Then
            capsStart(i) = snapshot(i)
            capsEnd(i) = snapshot(i + snapOff)
        Else
            ' slot not present in snapshot -> reset to default
            capsStart(i) = 0
            capsEnd(i) = -1
        End If
    Next i

    ' harmless cleanup: attempt to free snapshot memory
    On Error Resume Next
    Erase snapshot
    Exit Sub

bad_snapshot:
    ' Defensive fallback: if snapshot was bogus, do nothing but avoid crashing.
    On Error Resume Next
    err.Clear
End Sub

' ---------------- Char class matching with precomputed map -----------------
Private Function MatchCharClassNode(ByVal node As Collection, ByVal ch As String) As Boolean
    Dim code As Long: code = AscW(ch)
    If code >= 0 And code <= 255 Then
        Dim mapArr() As Boolean: mapArr = NodeGet(node, "map")
        Dim matched As Boolean: matched = False

        ' direct ASCII fast-path
        If code >= 0 And code <= 255 Then
            matched = mapArr(code)
            ' if ignore-case, check the swapped-case ASCII code too (A<->a)
            If pIgnoreCase And Not matched Then
                If code >= AscW("A") And code <= AscW("Z") Then
                    Dim alt As Long: alt = code + 32 ' to lower
                    If alt >= 0 And alt <= 255 Then matched = mapArr(alt)
                ElseIf code >= AscW("a") And code <= AscW("z") Then
                    Dim alt2 As Long: alt2 = code - 32 ' to upper
                    If alt2 >= 0 And alt2 <= 255 Then matched = mapArr(alt2)
                End If
            End If
        End If

        If NodeGet(node, "neg") Then
            MatchCharClassNode = Not matched
        Else
            MatchCharClassNode = matched
        End If
        Exit Function
    End If

    ' Non-ASCII fallback: brute-force check, but respect ignore-case
    Dim items As Collection: Set items = NodeGet(node, "chars")
    Dim matched2 As Boolean: matched2 = False
    Dim i As Long
    For i = 1 To items.count
        Dim it As Variant: it = items(i)
        If IsArray(it) Then
            If it(0) = "range" Then
                If Not pIgnoreCase Then
                    If AscW(ch) >= AscW(it(1)) And AscW(ch) <= AscW(it(2)) Then matched2 = True
                Else
                    ' case-insensitive range test: compare lowercased codepoints where possible
                    Dim lowerCh As String: lowerCh = LCase$(ch)
                    Dim lowerA As String: lowerA = LCase$(it(1))
                    Dim lowerB As String: lowerB = LCase$(it(2))
                    If AscW(lowerCh) >= AscW(lowerA) And AscW(lowerCh) <= AscW(lowerB) Then matched2 = True
                End If
            End If
        Else
            If Not pIgnoreCase Then
                If ch = it Then matched2 = True
            Else
                If LCase$(ch) = LCase$(CStr(it)) Then matched2 = True
            End If
        End If
        If matched2 Then Exit For
    Next i
    If NodeGet(node, "neg") Then MatchCharClassNode = Not matched2 Else MatchCharClassNode = matched2
End Function
' ---------------- Replacement helpers -----------------------------------
Private Function ExpandReplacement(ByVal caps As Collection, ByVal repl As String, Optional ByVal namedColl As Collection = Nothing) As String
    Dim out As String: out = vbNullString
    Dim i As Long: i = 1
    Dim L As Long: L = Len(repl)

    If caps Is Nothing Then
        ExpandReplacement = repl: Exit Function
    End If

    Dim mstart As Long: mstart = -1
    Dim mend As Long: mend = -1
    Dim ch As String
    Dim nextCh As String
    Dim j As Long
    Dim nm As String
    Dim c2 As String
    Dim numStr As String
    Dim k As Long
    Dim cc As String
    Dim useIdx As Long
    Dim try2 As Long
    Dim tryOne As Long
    Dim capIndex As Long

    On Error Resume Next
    mstart = CLng(capsStart(0))
    mend = CLng(capsEnd(0))
    On Error GoTo 0

    Do While i <= L
        ch = SafeMid(repl, i, 1)
        If ch = "$" Then
            If i = L Then
                out = out & "$": i = i + 1: GoTo ContinueLoop
            End If
            nextCh = SafeMid(repl, i + 1, 1)
            Select Case nextCh
                Case "$"
                    out = out & "$": i = i + 2: GoTo ContinueLoop
                Case "&"
                    If caps.count >= 1 Then out = out & CStr(caps.Item(1))
                    i = i + 2: GoTo ContinueLoop
                Case "`"
                    If mstart >= 1 Then
                        If mstart > 1 Then out = out & SafeMid(CurSubject, 1, mstart - 1) Else out = out & vbNullString
                    Else
                        out = out & vbNullString
                    End If
                    i = i + 2: GoTo ContinueLoop
                Case "'"
                    If mend >= 1 Then
                        If mend < Len(CurSubject) Then out = out & SafeMid(CurSubject, mend + 1) Else out = out & vbNullString
                    Else
                        out = out & vbNullString
                    End If
                    i = i + 2: GoTo ContinueLoop
                Case "<"
                    ' Named capture syntax $<name>
                    j = i + 2
                    nm = vbNullString
                    Do While j <= L
                        c2 = SafeMid(repl, j, 1)
                        If c2 = ">" Then Exit Do
                        nm = nm & c2
                        j = j + 1
                    Loop
                    If j <= L And SafeMid(repl, j, 1) = ">" Then
                        ' found close '>'
                        If Not IsMissing(namedColl) And Not namedColl Is Nothing Then
                            On Error Resume Next
                            If NodeExists(namedColl, nm) Then
                                out = out & CStr(namedColl(nm)(1))
                            Else
                                ' missing named capture yields empty string
                                out = out & vbNullString
                            End If
                            On Error GoTo 0
                        Else
                            ' no namedColl provided -> try to find in caps by numeric fallback (not possible) -> empty
                            out = out & vbNullString
                        End If
                        i = j + 1
                        GoTo ContinueLoop
                    Else
                        ' no closing '>' - treat literally as "$<..."
                        out = out & "$<": i = i + 2: GoTo ContinueLoop
                    End If
                Case "0" To "9"
                    ' numeric capture: greedy 2-digit then fallback
                    numStr = vbNullString
                    k = i + 1
                    Do While k <= L And Len(numStr) < 2
                        cc = SafeMid(repl, k, 1)
                        If cc >= "0" And cc <= "9" Then
                            numStr = numStr & cc
                            k = k + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    useIdx = -1
                    If Len(numStr) = 2 Then
                        try2 = CLng(numStr)
                        If try2 >= 0 And try2 <= caps.count - 1 Then
                            useIdx = try2
                            i = i + 3
                        Else
                            Dim try1 As Long: try1 = CLng(left$(numStr, 1))
                            If try1 >= 0 And try1 <= caps.count - 1 Then useIdx = try1
                            i = i + 2
                        End If
                    ElseIf Len(numStr) = 1 Then
                        tryOne = CLng(numStr)
                        If tryOne >= 0 And tryOne <= caps.count - 1 Then useIdx = tryOne
                        i = i + 2
                    Else
                        out = out & "$": i = i + 1
                    End If
                    If useIdx >= 0 Then
                        capIndex = useIdx + 1
                        If capIndex >= 1 And capIndex <= caps.count Then out = out & CStr(caps.Item(capIndex)) Else out = out & vbNullString
                    End If
                    GoTo ContinueLoop
                Case Else
                    out = out & "$": i = i + 1: GoTo ContinueLoop
            End Select
        Else
            out = out & ch: i = i + 1
        End If
ContinueLoop:
    Loop
    ExpandReplacement = out
End Function
' This implementation only returns empty matches for non-multiple groups or subjects
Private Function BuildCaptureCollection() As Collection
    Dim col As New Collection
    If capsStart(0) >= 1 And capsEnd(0) >= capsStart(0) Then
        col.Add SafeMid(CurSubject, capsStart(0), capsEnd(0) - capsStart(0) + 1)
    Else
        If col.count = 0 Then
            col.Add vbNullString
        End If
    End If
    Dim i As Long
    For i = 1 To pGroupCount
        If capsStart(i) >= 1 And capsEnd(i) >= capsStart(i) Then
            col.Add SafeMid(CurSubject, capsStart(i), capsEnd(i) - capsStart(i) + 1)
        Else
            If pGroupCount = 1 Then
                col.Add vbNullString
            End If
        End If
    Next i
    Set BuildCaptureCollection = col
End Function

' ---------------- Utilities ------------------------------------------------
Private Function ComputeMinLen(ByVal node As Collection) As Long
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "lit", "dot", "class": ComputeMinLen = 1
        Case "seq"
            Dim SUM As Long: SUM = 0
            Dim ch As Variant
            For Each ch In NodeGet(node, "child")
                SUM = SUM + ComputeMinLen(ch)
            Next
            ComputeMinLen = SUM
        Case "group"
            ComputeMinLen = ComputeMinLen(NodeGet(node, "child"))
        Case "quant"
            Dim mn As Long: mn = NodeGet(node, "min")
            If mn < 0 Then mn = 0
            ComputeMinLen = mn * ComputeMinLen(NodeGet(node, "child"))
        Case "alt"
            Dim minv As Long: minv = -1
            Dim chAlt As Variant
            For Each chAlt In NodeGet(node, "child")
                Dim vAlt As Long: vAlt = ComputeMinLen(chAlt)
                If minv = -1 Or vAlt < minv Then minv = vAlt
            Next
            If minv = -1 Then minv = 0
            ComputeMinLen = minv
        Case Else: ComputeMinLen = 0
    End Select
End Function
Private Function ReplaceNewLines(subj As String) As String
    ReplaceNewLines = VBA.Replace(VBA.Replace(subj, "\n", vbLf), "\r", vbCr)
End Function

' ---------------- Public Exec / ExecAll / Replace / Split ---------------------------
' Exec: find first match anywhere (respects ^ anchor and multiline semantics)
Public Function Exec(ByVal subj As String) As Collection
    CurSubject = subj
    MatchSteps = 0
    Dim anchoredStart As Boolean: anchoredStart = NodeGet(pAST, "anchoredStart")
    Dim L As Long: L = Len(subj)
    Dim minLen As Long
    minLen = NodeGet(pAST, "minLen")
    
    subj = ReplaceNewLines(subj)
    ' Anchored start: if multiline, try positions that are start-of-string or immediately after a newline (CR or LF)
    If anchoredStart Then
        If pMultiline Then
            Dim tryPos As Long
            For tryPos = 1 To L
                If tryPos = 1 Or SafeMid(subj, tryPos - 1, 1) = vbLf Or SafeMid(subj, tryPos - 1, 1) = vbCr Then
                    If Not (ExecAt(subj, tryPos) Is Nothing) Then
                        Set Exec = BuildCaptureCollection()
                        Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
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
                Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
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
            Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
            Exit Function
        End If
        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
    Next i
    Set Exec = Nothing
End Function

' ExecAll: find all non-overlapping matches anywhere in subject
' Returns a Collection where each item is a Collection of captures (same structure as BuildCaptureCollection)
Public Function ExecAll(ByVal subj As String) As Collection
    CurSubject = subj
    MatchSteps = 0

    If DEBUG_TRACE Then Debug.Print "TRACE ExecAll start subjLen=" & Len(subj) & " anchoredStart=" & CStr(NodeGet(pAST, "anchoredStart")) & " minLen=" & CStr(NodeGet(pAST, "minLen"))

    Dim anchoredStart As Boolean: anchoredStart = NodeGet(pAST, "anchoredStart")
    Dim L As Long: L = Len(subj)
    Dim minLen As Long: minLen = NodeGet(pAST, "minLen")

    Dim results As New Collection

    ' --- anchored-start behaviour ---
    If anchoredStart Then
        If pMultiline Then
            ' Try only positions that are start-of-string or immediately after CR/LF
            Dim tryPos As Long
            For tryPos = 1 To L
                If tryPos = 1 Or SafeMid(subj, tryPos - 1, 1) = vbLf Or SafeMid(subj, tryPos - 1, 1) = vbCr Then
                    If DEBUG_TRACE Then Debug.Print "TRACE ExecAll trying anchored pos=" & tryPos
                    If Not (ExecAt(subj, tryPos) Is Nothing) Then
                        ' execAt returns captures in engine globals; BuildCaptureCollection wraps them
                        Dim matchColl As Collection
                        Set matchColl = BuildCaptureCollection()
                        results.Add matchColl
                        If DEBUG_TRACE Then Debug.Print "TRACE ExecAll anchored matched at " & tryPos & " matchLen=" & (CLng(capsEnd(0)) - CLng(capsStart(0)) + 1)
                        ' For anchored-start we continue scanning other anchored positions (do not skip to match end)
                    End If
                End If
                If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
            Next tryPos

            If results.count = 0 Then
                Set ExecAll = Nothing
            Else
                Set ExecAll = results
                Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
            End If
            Exit Function

        Else
            ' anchored at beginning only
            If DEBUG_TRACE Then Debug.Print "TRACE ExecAll trying anchored pos=1"
            If ExecAt(subj, 1) Is Nothing Then
                Set ExecAll = Nothing
            Else
                Dim firstColl As Collection
                Set firstColl = BuildCaptureCollection()
                results.Add firstColl
                Set ExecAll = results
                Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
            End If
            Exit Function
        End If
    End If

    ' --- non-anchored behaviour: scan start positions up to startLimit ---
    Dim startLimit As Long
    If minLen > 0 Then startLimit = L - minLen + 1 Else startLimit = L
    If startLimit < 1 Then startLimit = 1

    Dim pos As Long: pos = 1
    Do While pos <= startLimit
        If DEBUG_TRACE Then Debug.Print "TRACE ExecAll trying pos=" & pos & " (startLimit=" & startLimit & ")"
        If Not (ExecAt(subj, pos) Is Nothing) Then
            ' match found at pos; Build collection of captures and add
            Dim col As Collection
            Set col = BuildCaptureCollection()
            results.Add col

            ' compute match start/end and length from caps
            Dim mstart As Long: mstart = CLng(capsStart(0))
            Dim mend As Long: mend = CLng(capsEnd(0))
            Dim matchLen As Long: matchLen = mend - mstart + 1

            If DEBUG_TRACE Then Debug.Print "TRACE ExecAll matched at pos=" & pos & " mstart=" & mstart & " mend=" & mend & " len=" & matchLen

            ' guard mstart not before pos
            If mstart < pos Then mstart = pos

            ' Advance pos after the matched text (non-overlapping)
            If matchLen <= 0 Then
                ' zero-length match, advance at least one to avoid infinite loop
                pos = pos + 1
            Else
                pos = mend + 1
            End If

            ' recompute startLimit if minLen depends on subject length (not necessary here but safe)
            If minLen > 0 Then startLimit = L - minLen + 1 Else startLimit = L
            If startLimit < 1 Then startLimit = 1
        Else
            ' no match at this start -> advance to next start position
            pos = pos + 1
        End If

        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
    Loop

    If results.count = 0 Then
        Set ExecAll = Nothing
    Else
        Set ExecAll = results
        Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
    End If
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
        Call UpdateNamedCapturesFromAST    ' <- keep named captures in sync
    Else
        If DEBUG_TRACE Then Debug.Print "TRACE ExecAt MatchNode returned False -> no match"
        Set ExecAt = Nothing
    End If
End Function

Public Function Replace(ByVal subj As String, ByVal Replacement As String) As String
    Dim out As String: out = vbNullString
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    Dim foundAt As Long
    Dim cs As Collection
    Dim startTry As Long
    Dim mstart As Long
    Dim mend As Long
    Dim matchLen As Long

    ' Fast path: empty subject
    If L = 0 Then
        Replace = vbNullString
        Exit Function
    End If

    Do
        foundAt = 0

        ' Scan forward to locate the next match by trying ExecAt at increasing start positions
        For startTry = pos To L
            Set cs = ExecAt(subj, startTry)
            If Not (cs Is Nothing) Then
                foundAt = startTry
                Exit For
            End If
        Next startTry

        If foundAt = 0 Then
            ' no more matches -> append remainder and finish
            out = out & SafeMid(subj, pos)
            Exit Do
        Else
            mstart = CLng(capsStart(0))
            mend = CLng(capsEnd(0))

            ' ensure mstart is not before current pos (some engines set capture positions differently)
            If mstart < pos Then mstart = pos

            ' append portion before the match (may be empty)
            If mstart > pos Then
                out = out & SafeMid(subj, pos, mstart - pos)
            End If

            ' expand replacement using your helper (ExpandReplacement)
            out = out & ExpandReplacement(cs, Replacement, pNamedCaptures)

            ' advance position after the matched substring
            matchLen = mend - mstart + 1
            If matchLen <= 0 Then
                ' zero-length match: advance by one to avoid infinite loop
                If foundAt + 1 > L Then
                    pos = L + 1
                Else
                    pos = foundAt + 1
                End If
            Else
                pos = mend + 1
            End If

            ' if we've consumed the whole subject, append empty tail and exit
            If pos > L Then
                out = out & vbNullString ' explicit - nothing to append
                Exit Do
            End If
        End If
    Loop

    Replace = out
End Function


Public Function Split(ByVal subj As String) As Collection
    Dim res As New Collection
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    Dim foundAt As Long
    Dim cs As Collection
    Dim startTry As Long
    Dim mstart As Long
    Dim mend As Long
    Dim matchLen As Long
    
    Do
        foundAt = 0

        ' Scan forward to locate next match (try ExecAt at increasing start positions)
        For startTry = pos To L
            Set cs = ExecAt(subj, startTry)
            If Not (cs Is Nothing) Then
                foundAt = startTry
                Exit For
            End If
        Next startTry

        If foundAt = 0 Then
            ' no more matches -> append remainder and finish
            res.Add SafeMid(subj, pos)
            Exit Do
        Else
            mstart = CLng(capsStart(0))
            mend = CLng(capsEnd(0))

            ' ensure mstart is not before current pos
            If mstart < pos Then mstart = pos
            res.Add SafeMid(subj, pos, mstart - pos)
            ' advance position after the matched separator
            matchLen = mend - mstart + 1
            If matchLen <= 0 Then
                ' zero-length match: advance by one to avoid infinite loop
                pos = IIf(foundAt + 1 > L, L + 1, foundAt + 1)
            Else
                pos = mend + 1
            End If

            ' if we've consumed the whole subject, append empty tail
            If pos > L Then
                res.Add vbNullString
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
            Dim SUM As Long: SUM = 0
            Dim ch As Variant
            For Each ch In NodeGet(node, "child")
                Dim m As Long: m = ComputeMaxLen(ch)
                If m = -1 Then
                    ComputeMaxLen = -1: Exit Function
                End If
                SUM = SUM + m
            Next
            ComputeMaxLen = SUM
        Case "group"
            ComputeMaxLen = ComputeMaxLen(NodeGet(node, "child"))
        Case "alt"
            Dim mx As Long: mx = 0
            Dim itm As Variant
            For Each itm In NodeGet(node, "child")
                Dim vm As Long: vm = ComputeMaxLen(itm)
                If vm = -1 Then
                    mx = -1: Exit For
                ElseIf vm > mx Then mx = vm
                End If
            Next
            ComputeMaxLen = mx
        Case "quant"
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
            For Each ch In NodeGet(node, "child")
                AnnotateLookbehinds ch
            Next
        Case "group", "atomic", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg", "recurse"
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
            For Each it In NodeGet(node, "child")
                AnnotateLookbehinds it
            Next
    End Select
End Sub
' Return fixed width in characters for node, or -1 if variable/unbounded
Private Function ComputeFixedWidth(ByVal node As Collection) As Long
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "lit", "dot", "class"
            ComputeFixedWidth = 1
            Exit Function
        Case "seq"
            Dim SUM As Long: SUM = 0
            Dim ch As Variant
            For Each ch In NodeGet(node, "child")
                Dim w As Long: w = ComputeFixedWidth(ch)
                If w = -1 Then
                    ComputeFixedWidth = -1: Exit Function
                End If
                SUM = SUM + w
            Next
            ComputeFixedWidth = SUM
            Exit Function
        Case "group"
            ComputeFixedWidth = ComputeFixedWidth(NodeGet(node, "child"))
            Exit Function
        Case "quant"
            Dim child As Collection: Set child = NodeGet(node, "child")
            Dim cw As Long: cw = ComputeFixedWidth(child)
            If cw = -1 Then
                ComputeFixedWidth = -1: Exit Function
            End If
            Dim mn As Long: mn = NodeGet(node, "min")
            Dim mx As Long: mx = NodeGet(node, "max")
            ' if max is -1 (unbounded) or range not fixed -> variable
            If mx = -1 Then
                ComputeFixedWidth = -1: Exit Function
            End If
            If mn <> mx Then
                ComputeFixedWidth = -1: Exit Function
            End If
            ' fixed repeat count
            ComputeFixedWidth = mn * cw
            Exit Function
        Case "alt"
            ' alternation: all alternatives must have same fixed width for the alt to be fixed-width
            Dim firstW As Variant: firstW = Empty
            Dim childAlt As Variant
            For Each childAlt In NodeGet(node, "child")
                Dim aw As Long: aw = ComputeFixedWidth(childAlt)
                If aw = -1 Then
                    ComputeFixedWidth = -1: Exit Function
                End If
                If IsEmpty(firstW) Then
                    firstW = aw
                Else
                    If aw <> firstW Then
                        ComputeFixedWidth = -1: Exit Function
                    End If
                End If
            Next childAlt
            If IsEmpty(firstW) Then
                ComputeFixedWidth = 0
            Else
                ComputeFixedWidth = CLng(firstW)
            End If
            Exit Function
        Case Else
            ' unknown/complex node types: assume variable
            ComputeFixedWidth = -1
            Exit Function
    End Select
End Function
' Walk AST; for lookbehind nodes compute width and enforce fixed width.
Private Sub EnforceFixedWidthLookbehinds(ByVal ast As Collection)
    TraverseEnforce NodeGet(ast, "body")
End Sub

Private Sub TraverseEnforce(ByVal node As Collection)
    Dim t As String: t = NodeGet(node, "type")
    Select Case t
        Case "seq"
            Dim ch As Variant
            For Each ch In NodeGet(node, "child")
                TraverseEnforce ch
            Next
        Case "group"
            TraverseEnforce NodeGet(node, "child")
        Case "quant"
            TraverseEnforce NodeGet(node, "child")
        Case "alt"
            Dim altCh As Variant
            For Each altCh In NodeGet(node, "child")
                TraverseEnforce altCh
            Next
        Case "lookbehind_pos", "lookbehind_neg"
            ' compute fixed width of the inner child
            Dim child As Collection: Set child = NodeGet(node, "child")
            Dim w As Long: w = ComputeFixedWidth(child)
            ' store width for runtime use (lookbehind code expects node width)
            NodeAdd node, "width", w
            If w = -1 Then
                ' raise the specific error your tests expect
                err.Raise vbObjectError + 10, "RegexEngine", "Variable-length lookbehind not supported"
            End If
        Case Else
            ' other nodes need no special traversal
    End Select
End Sub
' Raise error if pattern contains unsupported (?#comment) syntax outside character classes.
Private Sub DetectUnsupportedCommentSyntax(ByVal pat As String)
    Dim i As Long
    Dim L As Long: L = Len(pat)
    Dim esc As Boolean: esc = False
    Dim inClass As Boolean: inClass = False

    For i = 1 To L
        Dim c As String: c = mid$(pat, i, 1)

        If esc Then
            esc = False
            GoTo ContinueLoop
        End If

        If c = "\" Then
            esc = True
            GoTo ContinueLoop
        End If

        If inClass Then
            If c = "]" Then inClass = False
            GoTo ContinueLoop
        Else
            If c = "[" Then
                inClass = True
                GoTo ContinueLoop
            End If
        End If

        ' look for literal "(?#" outside classes and not escaped
        If c = "(" Then
            If i + 2 <= L Then
                If mid$(pat, i + 1, 2) = "?#" Then
                    ' unsupported comment syntax
                    err.Raise vbObjectError + 20, "RegexEngine", "(?# comments are not supported"
                End If
            End If
        End If

ContinueLoop:
    Next i
End Sub

' Escape a string so it can be used as a literal regex pattern.
' Escapes the common regex metacharacters: \ . ^ $ * + ? ( ) [ ] { } | /
Public Function Escape(ByVal aInput As String) As String
    Dim specials As String
    specials = "\.^$*+?()[]{}|/"

    Dim outStr As String: outStr = vbNullString
    Dim i As Long, ch As String

    For i = 1 To Len(aInput)
        ch = mid$(aInput, i, 1)
        If InStr(1, specials, ch, vbBinaryCompare) > 0 Then
            outStr = outStr & "\" & ch
        Else
            If i = 1 Then
                outStr = outStr & "\" & ch
            Else
                outStr = outStr & ch
            End If
        End If
    Next i
    Escape = outStr
End Function

' ---- Regex runtime helper: returns group index (Long) given a name, or -1 if not found ----
Private Function FindGroupIndexByName(ByVal ast As Collection, ByVal name As String) As Long
    FindGroupIndexByName = -1
    On Error GoTo ExitHere
    If Not NodeExists(ast, "groupNodes") Then Exit Function
    Dim gnodes As Collection: Set gnodes = NodeGet(ast, "groupNodes")
    Dim lb As Long
    
    Dim it As Variant
    For Each it In gnodes
        lb = LBound(it)
        If it(lb) = name Then
            FindGroupIndexByName = CLng(it(lb + 1))
            Exit Function
        End If
    Next it

ExitHere:
    Exit Function
End Function
Private Function FindGroupIndexByOrdinal(ByVal ast As Collection, ByVal ordinal As Long) As Long
    FindGroupIndexByOrdinal = -1
    On Error GoTo ExitHere
    If Not NodeExists(ast, "groupNodes") Then Exit Function
    Dim gnodes As Collection: Set gnodes = NodeGet(ast, "groupNodes")
    Dim lb As Long
    Dim cIdx As Long
    
    Dim it As Variant
    For Each it In gnodes
        lb = LBound(it)
        cIdx = cIdx + 1
        If it(lb + 1) = ordinal Then
            FindGroupIndexByOrdinal = cIdx
            Exit Function
        End If
    Next it

ExitHere:
    Exit Function
End Function
' Build a groupNodes index for the provided AST (collection).
' Result: a Collection where each element is an Array(name As String, idx As Long, node As Variant)
' This function will attach the collection to the AST by calling: NodeAdd ast, "groupNodes", groupNodes
Private Function BuildGroupNodesFromAST(ByVal ast As Collection) As Collection
    On Error GoTo ErrHandler
    Dim groupNodes As New Collection

    If ast Is Nothing Then
        Set BuildGroupNodesFromAST = groupNodes
        Exit Function
    End If

    ' Walk AST and collect group nodes that have a "name"
    Call CollectGroupNodesRecursive(ast, groupNodes)

    ' Attach to AST if not already present (do not overwrite existing)
    If Not NodeExists(ast, "groupNodes") Then
        NodeAdd ast, "groupNodes", groupNodes
    End If

    Set BuildGroupNodesFromAST = groupNodes
    Exit Function

ErrHandler:
    ' defensive: return whatever we collected so far
    Set BuildGroupNodesFromAST = groupNodes
End Function


' Recursive helper that inspects node-like collections and adds named group info to outColl
Private Sub CollectGroupNodesRecursive(ByVal node As Variant, ByRef outColl As Collection)
    On Error Resume Next
    Dim t As String

    If node Is Nothing Then Exit Sub

    If NodeExists(node, "type") Then
        t = CStr(NodeGet(node, "type"))
        If t = "group" Then
            If NodeExists(node, "name") Then
                Dim nm As String: nm = CStr(NodeGet(node, "name"))
                Dim idx As Long: idx = 0
                If NodeExists(node, "idx") Then idx = CLng(NodeGet(node, "idx"))
                ' store pair (name, idx, nodeRef) — nodeRef kept for debugging/inspection
                outColl.Add Array(nm, idx, node)
            End If
        End If
    End If

    ' Try to recurse into likely child slots used by AST nodes.
    ' This covers "child" (Collection of seqs), "child", "base", "index", "left", "right", "body", "cases", etc.
    If NodeExists(node, "child") Then
        Dim ch As Variant
        For Each ch In NodeGet(node, "child")
            CollectGroupNodesRecursive ch, outColl
        Next ch
    End If

    If NodeExists(node, "child") Then CollectGroupNodesRecursive NodeGet(node, "child"), outColl
    If NodeExists(node, "base") Then CollectGroupNodesRecursive NodeGet(node, "base"), outColl
    If NodeExists(node, "index") Then CollectGroupNodesRecursive NodeGet(node, "index"), outColl
    If NodeExists(node, "left") Then CollectGroupNodesRecursive NodeGet(node, "left"), outColl
    If NodeExists(node, "right") Then CollectGroupNodesRecursive NodeGet(node, "right"), outColl
    If NodeExists(node, "body") Then CollectGroupNodesRecursive NodeGet(node, "body"), outColl
    If NodeExists(node, "cases") Then
        Dim c As Variant
        For Each c In NodeGet(node, "cases")
            CollectGroupNodesRecursive c, outColl
        Next c
    End If
End Sub
' Ensure AST.top has groupNodes built. Returns the collection (newly built or existing).
Private Function EnsureGroupNodes(ByVal ast As Collection) As Collection
    If ast Is Nothing Then
        Dim emptyColl As New Collection
        Set EnsureGroupNodes = emptyColl
        Exit Function
    End If

    If NodeExists(ast, "groupNodes") Then
        Set EnsureGroupNodes = NodeGet(ast, "groupNodes")
        Exit Function
    End If

    Set EnsureGroupNodes = BuildGroupNodesFromAST(ast)
End Function
' Rebuild the pNamedCaptures collection using the AST's groupNodes index and current caps arrays.
Private Sub UpdateNamedCapturesFromAST()
    On Error Resume Next
    Set pNamedCaptures = New Collection

    If pAST Is Nothing Then Exit Sub
    If Not IsArray(capsStart) Then Exit Sub

    ' Ensure index exists
    Dim gnodes As Collection
    Set gnodes = EnsureGroupNodes(pAST)
    If gnodes Is Nothing Then Exit Sub

    Dim i As Long
    Dim val As String
    Dim pair As Variant
    Dim gname As String
    Dim gidx As Long
    For i = 1 To gnodes.count
        pair = gnodes(i) ' Array(name, idx, nodeRef)
        gname = CStr(pair(0))
        gidx = CLng(pair(1))
        Dim sPos As Long
        Dim ePos As Long
        val = vbNullString
        If gidx > 0 Then
            If gidx <= UBound(capsStart) Then
                sPos = CLng(capsStart(gidx))
                ePos = CLng(capsEnd(gidx))
                If sPos >= 1 And ePos >= sPos Then
                    val = mid$(CurSubject, sPos, ePos - sPos + 1)
                Else
                    val = vbNullString
                End If
            End If
        End If
        pNamedCaptures.Add (Array(gname, val)), gname
    Next i
End Sub



