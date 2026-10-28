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
'  - Anchors: ^ and $ anywhere in the pattern (line anchors with multiline)
'  - Backreferences: \1..\n and named \k<name>
'  - Word boundaries: \b and \B
'  - Code point escapes: \xhh, \uhhhh, \f, \v, \0
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
'     + A quantified atom with internal choices, e.g. (?:a|ab)+, backtracks into each iteration for
'       the first QUANT_REC_LIMIT iterations; beyond that the remaining iterations take the first
'       alternative that matches, which only matters when the alternatives overlap.
Option Explicit
Option Base 0

' ---------------- Public config -----------------------------------------
Private Const DEFAULT_MAX_STEPS As Long = 200000
' --- debugging toggle: set True to print detailed trace during matching
Private Const DEBUG_TRACE As Boolean = False
' iterations of a choice-bearing quantifier matched recursively before falling back to the iterative scan
Private Const QUANT_REC_LIMIT As Long = 100
' (?R)/(?n) recursion depth and group nesting accepted before raising, to keep VBA stack use bounded
Private Const RECURSE_LIMIT As Long = 200
Private Const MAX_GROUP_NESTING As Long = 100

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
Private pQuantDepth As Long
Private pRecDepth As Long
' class bitmaps (char code 0..255, class index): pBuildMaps while compiling, pClassMaps once committed
Private pClassMaps() As Boolean
Private pBuildMaps() As Boolean
Private pBuildMapCount As Long
' chars that can start a match; pHasFirst is False when the pattern can match empty
Private pFirstMap() As Boolean
Private pFirstHigh As Boolean
Private pHasFirst As Boolean
Private pRecActive As String
Private pNamedCaptures As Collection
Private pNamedCount As Long
' match-start filters from the AST: leading literal and ^ at the start of the pattern
Private pPrefix As String
Private pPrefixNode As Collection
Private pAnchorNode As Collection
' subject as UTF-16 bytes for the start-position scan
Private pSubjBytes() As Byte
Private pCapsSize As Long
' compiled patterns keyed by pattern text; engines can share one through SharedCache
Private pCache As Collection
Private Const CACHE_LIMIT As Long = 64

Public Property Get LastNamedCaptures() As Collection
    Set LastNamedCaptures = pNamedCaptures
End Property
Public Property Get SharedCache() As Collection
    If pCache Is Nothing Then Set pCache = New Collection
    Set SharedCache = pCache
End Property
Public Property Set SharedCache(ByVal c As Collection)
    Set pCache = c
End Property
' ---------------- Node helpers ------------------------------------------
Private Function newNode(ByVal typ As String) As Collection
    Dim c As New Collection
    c.Add typ, "type"
    ' always present so FlagState never pays for a failed key lookup
    c.Add vbNullString, "flags_on"
    Set newNode = c
End Function

Private Sub NodeAdd(ByRef node As Collection, ByVal key As String, ByVal value As Variant)
    ' new keys are the common case: add first, replace only an existing key
    On Error GoTo replace_
    node.Add value, key
    Exit Sub
replace_:
    node.Remove key
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
    ' Compile (or reuse) first, so a rejected pattern leaves the previous pattern and flags in place
    If pCache Is Nothing Then Set pCache = New Collection
    Dim key As String: key = CacheKey(pattern)
    Dim st As Variant
    If Not CacheGet(key, st) Then
        st = CompileState(pattern)
        If pCache.Count >= CACHE_LIMIT Then
            Do While pCache.Count > 0
                pCache.Remove 1
            Loop
        End If
        pCache.Add st, key
    End If
    LoadState st

    pPattern = pattern
    pIgnoreCase = ignoreCase
    pMultiline = multiline
    pDotAll = dotAll
    If MaxSteps <= 0 Then pMaxMatchSteps = DEFAULT_MAX_STEPS Else pMaxMatchSteps = MaxSteps
End Sub

Private Function CacheGet(ByVal key As String, ByRef st As Variant) As Boolean
    On Error GoTo miss
    Dim cache As Collection: Set cache = pCache
    Dim v As Variant: v = cache(key)
    If Not IsArray(v) Then Exit Function
    st = v
    CacheGet = True
miss:
End Function

' Everything matching needs from a pattern; cached, so it must not depend on the flags
Private Function CompileState(ByVal pattern As String) As Variant
    Dim ast As Collection: Set ast = CompilePattern(pattern)
    Dim named As Collection: Set named = BuildGroupNodesFromAST(ast)
    ' fresh arrays, never resized in place: a cached state may hold the previous ones
    Dim firstMap() As Boolean: ReDim firstMap(0 To 255)
    pFirstMap = firstMap
    pFirstHigh = False
    pHasFirst = Not AddFirstChars(ast("body"))
    Dim maps As Variant
    If pBuildMapCount > 0 Then maps = pBuildMaps
    Dim prefix As String, prefixNode As Collection, anchorNode As Collection
    AnalyzeStart ast("body"), prefix, prefixNode, anchorNode
    CompileState = Array(ast, CLng(ast("groupCount")), pFirstMap, pFirstHigh, pHasFirst, maps, prefix, prefixNode, anchorNode, named.Count)
End Function

Private Sub LoadState(ByRef st As Variant)
    Set pAST = st(0)
    pGroupCount = st(1)
    pFirstMap = st(2)
    pFirstHigh = st(3)
    pHasFirst = st(4)
    If Not IsEmpty(st(5)) Then pClassMaps = st(5)
    pPrefix = st(6)
    Set pPrefixNode = st(7)
    Set pAnchorNode = st(8)
    pNamedCount = st(9)
End Sub

' Collection keys ignore case, so the pattern is keyed by its UTF-16 bytes in hex
Private Function CacheKey(ByRef pat As String) As String
    Dim b() As Byte: b = pat
    Dim s As String: s = Space$(2 * (UBound(b) + 1))
    Dim i As Long
    For i = 0 To UBound(b)
        Mid$(s, 2 * i + 1, 2) = Right$("0" & Hex$(b(i)), 2)
    Next i
    CacheKey = "k" & s
End Function

' Leading literal and leading ^ of a single-alternative pattern; zero-width assertions before the literal are skipped
Private Sub AnalyzeStart(ByVal body As Collection, ByRef prefix As String, ByRef prefixNode As Collection, ByRef anchorNode As Collection)
    Dim seq As Collection: Set seq = body
    If CStr(seq(1)) = "alt" Then
        If seq("child").Count <> 1 Then Exit Sub
        Set seq = seq("child")(1)
    End If
    If CStr(seq(1)) <> "seq" Then Exit Sub
    Dim nd As Variant, k As Long
    For Each nd In seq("child")
        k = k + 1
        Select Case CStr(nd(1))
            Case "bol", "wordb", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
                If Len(prefix) > 0 Then Exit For
                If k = 1 And CStr(nd(1)) = "bol" Then Set anchorNode = nd
            Case "lit"
                If prefixNode Is Nothing Then Set prefixNode = nd
                If CStr(nd("flags_on")) <> CStr(prefixNode("flags_on")) Then Exit For
                prefix = prefix & CStr(nd("ch"))
            Case Else
                Exit For
        End Select
    Next nd
End Sub

Private Sub EnsureCompiled()
    If pAST Is Nothing Then err.Raise vbObjectError + 4, "RegexEngine", "No pattern: call Init before matching"
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

Public Function test(ByVal subj As String) As Boolean
    EnsureCompiled
    test = FindFirst(subj)
    If test Then UpdateNamedCapturesFromAST
End Function

' ---------------- Parser & AST precomputation ----------------------------
Private Function CompilePattern(ByVal pat As String) As Collection
    ' Validate pattern structure before any parsing (raises on bad syntax).
    ' This subsumes DetectUnsupportedCommentSyntax and adds paren-balance,
    ' unclosed class, dangling backslash, and quantifier-operand checks.
    ValidatePattern pat

    Dim ast As New Collection
    NodeAdd ast, "raw", pat
    pBuildMapCount = 0
    Dim noMaps() As Boolean: pBuildMaps = noMaps

    Dim pos As Long: pos = 1
    Dim groupCounter As Long: groupCounter = 0
    Dim body As Collection: Set body = ParseAlternation(pat, pos, Len(pat), groupCounter)
    NodeAdd ast, "body", body
    NodeAdd ast, "groupCount", groupCounter
    
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
            ch = SafeMid(pat, p, 1)
            Dim codeCh As String
            If ParseCodeEscape(pat, p, pEnd, codeCh) Then
                Set node = newNode("lit"): NodeAdd node, "ch", codeCh
            ElseIf ch >= "1" And ch <= "9" Then
                Dim refNum As String: refNum = vbNullString
                Do While p <= pEnd
                    If SafeMid(pat, p, 1) < "0" Or SafeMid(pat, p, 1) > "9" Then Exit Do
                    refNum = refNum & SafeMid(pat, p, 1)
                    p = p + 1
                Loop
                Set node = newNode("backref"): NodeAdd node, "idx", CLng(refNum)
            ElseIf ch = "k" And SafeMid(pat, p + 1, 1) = "<" And InStr(p + 2, pat, ">") > 0 Then
                Dim refEnd As Long: refEnd = InStr(p + 2, pat, ">")
                Set node = newNode("backref"): NodeAdd node, "name", SafeMid(pat, p + 2, refEnd - p - 2)
                p = refEnd + 1
            Else
                Set node = ParseEscapeNode(ch)
                p = p + 1
            End If

        ElseIf ch = "." Then
            Set node = newNode("dot")
            p = p + 1

        ElseIf ch = "^" Then
            Set node = newNode("bol")
            p = p + 1

        ElseIf ch = "$" Then
            Set node = newNode("eol")
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
                                NodeAdd ncNode, "idx", 0
                                NodeAdd ncNode, "name", vbNullString
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
                                        NodeAdd gnodeF, "idx", 0
                                        NodeAdd gnodeF, "name", vbNullString
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
                    NodeAdd gnode, "name", vbNullString
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
            If qch = "{" Then
                If Not IsBraceQuantifier(pat, p, pEnd) Then qch = vbNullString
            End If
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
    
    If typeName(node) = "Collection" Then
        Dim t As String
        Dim elm As Variant
        For Each elm In node
            If typeName(elm) = "Collection" Then
                t = vbNullString
                If NodeExists(elm, "type") Then t = NodeGet(elm, "type")
                Select Case t
                    Case "lit", "dot", "class", "backref", "bol", "eol"
                        Dim tmpNode As Collection
                        Set tmpNode = elm
                        ' inner groups were parsed first, so their flags win
                        NodeAdd tmpNode, "flags_on", MergeFlags(flags, CStr(NodeGet(tmpNode, "flags_on")))
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

' 1 when flag f is switched on in a flags string such as "i" or "i-ms", -1 when switched off, 0 otherwise
Private Function ExplicitFlag(ByVal fl As String, ByVal f As String) As Long
    Dim dash As Long: dash = InStr(fl, "-")
    If dash > 0 Then
        If InStr(dash, fl, f) > 0 Then ExplicitFlag = -1: Exit Function
        fl = left$(fl, dash - 1)
    End If
    If InStr(fl, f) > 0 Then ExplicitFlag = 1
End Function

Private Function MergeFlags(ByVal outer As String, ByVal inner As String) As String
    Dim onPart As String, offPart As String
    Dim f As Variant
    Dim st As Long
    For Each f In Array("i", "m", "s")
        st = ExplicitFlag(inner, CStr(f))
        If st = 0 Then st = ExplicitFlag(outer, CStr(f))
        If st = 1 Then onPart = onPart & f
        If st = -1 Then offPart = offPart & f
    Next f
    If Len(offPart) > 0 Then MergeFlags = onPart & "-" & offPart Else MergeFlags = onPart
End Function

' Effective state of flag f for a node: its inline flags, else the engine default
Private Function FlagState(ByVal node As Collection, ByVal f As String, ByVal dflt As Boolean) As Boolean
    FlagState = dflt
    Dim fl As String: fl = node("flags_on")
    If Len(fl) = 0 Then Exit Function
    Select Case ExplicitFlag(fl, f)
        Case 1: FlagState = True
        Case -1: FlagState = False
    End Select
End Function

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
        Case "s": Set node = ParseClassFromSimple("[ " & vbTab & vbLf & vbCr & Chr$(11) & Chr$(12) & "]")
        Case "S"
            ' Non-space
            Set node = ParseClassFromSimple("[ " & vbTab & vbLf & vbCr & Chr$(11) & Chr$(12) & "]")
            NodeAdd node, "neg", True
        Case "b": Set node = newNode("wordb"): NodeAdd node, "neg", False
        Case "B": Set node = newNode("wordb"): NodeAdd node, "neg", True
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
    Dim lowCh As String, highCh As String
    Dim j As Long
    Do While i <= pEnd
        If ParseClassAtom(pat, i, pEnd, Chars, lowCh) Then
            ' a char may start a range; "-" before a set or the end stays literal
            If i < pEnd And SafeMid(pat, i, 1) = "-" Then
                j = i + 1
                If PeekClassChar(pat, j, pEnd, highCh) Then
                    If CodeOf(highCh) < CodeOf(lowCh) Then
                        err.Raise vbObjectError + 41, "RegexEngine", _
                            "Invalid pattern: character class range out of order '" & lowCh & "-" & highCh & "'"
                    End If
                    Chars.Add Array("range", lowCh, highCh)
                    i = j
                Else
                    Chars.Add lowCh
                End If
            Else
                Chars.Add lowCh
            End If
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
                Dim aCode As Long: aCode = CodeOf(it(1))
                Dim bCode As Long: bCode = CodeOf(it(2))
                If bCode > 255 Then bCode = 255
                Dim cc As Long
                For cc = aCode To bCode
                    mapArr(cc) = True
                Next cc
            End If
        Else
            Dim code As Long: code = CodeOf(it)
            If code <= 255 Then mapArr(code) = True
        End If
    Next it
    NodeAdd node, "mapIdx", RegisterClassMap(mapArr)
    Set ParseClass = node
End Function

' One class member at i: a char (returned in outCh, True) or a shorthand set (added to Chars, False); advances i
Private Function ParseClassAtom(ByVal pat As String, ByRef i As Long, ByVal pEnd As Long, ByRef Chars As Collection, ByRef outCh As String) As Boolean
    Dim ch As String: ch = SafeMid(pat, i, 1)
    If ch <> "\" Or i = pEnd Then
        outCh = ch: i = i + 1
        ParseClassAtom = True: Exit Function
    End If
    i = i + 1
    Dim esc As String: esc = SafeMid(pat, i, 1)
    Select Case esc
        Case "d"
            Chars.Add Array("range", "0", "9")
        Case "D"
            AddClassRanges Chars, "0-47,58-65535"
        Case "w"
            Chars.Add Array("range", "A", "Z")
            Chars.Add Array("range", "a", "z")
            Chars.Add Array("range", "0", "9")
            Chars.Add "_"
        Case "W"
            AddClassRanges Chars, "0-47,58-64,91-94,96-96,123-65535"
        Case "s"
            AddClassRanges Chars, "9-13,32-32"
        Case "S"
            AddClassRanges Chars, "0-8,14-31,33-65535"
        Case Else
            If (esc = "p" Or esc = "P") And SafeMid(pat, i + 1, 1) = "{" Then
                err.Raise vbObjectError + 46, "RegexEngine", "Invalid pattern: Unicode property escapes (\" & esc & "{...}) are not supported"
            End If
            If Not ParseCodeEscape(pat, i, pEnd, outCh) Then
                ' n, r, t, b, or the literal char
                outCh = ParseEscapeChar(esc)
                i = i + 1
            End If
            ParseClassAtom = True: Exit Function
    End Select
    i = i + 1
End Function

' True with the char when the class member at j is a single char (range end); advances j
Private Function PeekClassChar(ByVal pat As String, ByRef j As Long, ByVal pEnd As Long, ByRef outCh As String) As Boolean
    Dim scratch As New Collection
    PeekClassChar = ParseClassAtom(pat, j, pEnd, scratch, outCh)
End Function

' Add the chars that can start a match of node to pFirstMap/pFirstHigh; True when node can match empty
Private Function AddFirstChars(ByVal node As Collection) As Boolean
    Dim cc As Long
    Dim it As Variant
    Dim mapIdx As Long
    Select Case CStr(node(1))
        Case "lit"
            AddFirstChar CStr(node("ch"))
        Case "dot"
            AddAnyFirst
        Case "class"
            mapIdx = node("mapIdx")
            For cc = 0 To 255
                If pBuildMaps(cc, mapIdx) <> CBool(node("neg")) Then AddFirstChar ChrW$(cc)
            Next cc
            pFirstHigh = True
        Case "seq"
            AddFirstChars = True
            For Each it In node("child")
                If Not AddFirstChars(it) Then AddFirstChars = False: Exit For
            Next it
        Case "group", "atomic"
            AddFirstChars = AddFirstChars(node("child"))
        Case "alt"
            For Each it In node("child")
                If AddFirstChars(it) Then AddFirstChars = True
            Next it
        Case "quant"
            AddFirstChars = AddFirstChars(node("child")) Or (node("min") = 0)
        Case "wordb", "bol", "eol", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
            AddFirstChars = True
        Case Else
            ' backreferences, conditionals and recursion may start with anything
            AddAnyFirst
            AddFirstChars = True
    End Select
End Function

' Both cases count, since ignoreCase can change after Init; letters may also match chars above 255
Private Sub AddFirstChar(ByVal ch As String)
    Dim v As Variant
    For Each v In Array(ch, LCase$(ch), UCase$(ch))
        If CodeOf(CStr(v)) <= 255 Then pFirstMap(CodeOf(CStr(v))) = True Else pFirstHigh = True
    Next v
    If LCase$(ch) <> UCase$(ch) Then pFirstHigh = True
End Sub

Private Sub AddAnyFirst()
    Dim cc As Long
    For cc = 0 To 255
        pFirstMap(cc) = True
    Next cc
    pFirstHigh = True
End Sub

' Store a class bitmap in pBuildMaps and return its index
Private Function RegisterClassMap(ByRef mapArr() As Boolean) As Long
    If pBuildMapCount = 0 Then
        ReDim pBuildMaps(0 To 255, 0 To 0)
    Else
        ReDim Preserve pBuildMaps(0 To 255, 0 To pBuildMapCount)
    End If
    Dim cc As Long
    For cc = 0 To 255
        pBuildMaps(cc, pBuildMapCount) = mapArr(cc)
    Next cc
    RegisterClassMap = pBuildMapCount
    pBuildMapCount = pBuildMapCount + 1
End Function

' Add "a-b,c-d" code point ranges to a class
Private Sub AddClassRanges(ByRef Chars As Collection, ByVal spec As String)
    Dim rng As Variant
    Dim bounds() As String
    For Each rng In VBA.Split(spec, ",")
        bounds = VBA.Split(rng, "-")
        Chars.Add Array("range", ChrW$(CLng(bounds(0))), ChrW$(CLng(bounds(1))))
    Next rng
End Sub

' Code point of a char as 0..65535 (AscW is negative above &H7FFF)
Private Function CodeOf(ByVal s As String) As Long
    CodeOf = AscW(s) And &HFFFF&
End Function

' \xhh, \uhhhh, \f, \v and \0 at p (the char after the backslash); advances p past the escape
Private Function ParseCodeEscape(ByVal pat As String, ByRef p As Long, ByVal pEnd As Long, ByRef outCh As String) As Boolean
    Dim hexLen As Long
    Select Case SafeMid(pat, p, 1)
        Case "x": hexLen = 2
        Case "u": hexLen = 4
        Case "f": outCh = Chr$(12): p = p + 1: ParseCodeEscape = True: Exit Function
        Case "v": outCh = Chr$(11): p = p + 1: ParseCodeEscape = True: Exit Function
        Case "0"
            If SafeMid(pat, p + 1, 1) Like "[0-9]" Then Exit Function
            outCh = Chr$(0): p = p + 1: ParseCodeEscape = True: Exit Function
        Case Else: Exit Function
    End Select
    If p + hexLen > pEnd Then Exit Function
    Dim hx As String: hx = LCase$(SafeMid(pat, p + 1, hexLen))
    If hx Like "*[!0-9a-f]*" Then Exit Function
    Dim code As Long, k As Long
    For k = 1 To hexLen
        code = code * 16 + InStr("0123456789abcdef", mid$(hx, k, 1)) - 1
    Next k
    outCh = ChrW$(code)
    p = p + 1 + hexLen
    ParseCodeEscape = True
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
    NodeAdd node, "mapIdx", RegisterClassMap(mapArr)
    Set ParseClassFromSimple = node
End Function

Private Function ParseEscapeChar(ByVal ch As String) As String
    Select Case ch
        Case "b": ParseEscapeChar = Chr$(8)
        Case "n": ParseEscapeChar = vbLf
        Case "r": ParseEscapeChar = vbCr
        Case "t": ParseEscapeChar = vbTab
        Case Else: ParseEscapeChar = ch
    End Select
End Function

' ------------------ Matching engine (returns explicit end positions) --------

Private Function MatchNode(ByVal node As Collection, ByRef subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    If DEBUG_TRACE Then Debug.Print "TRACE MatchNode ENTER type=" & CStr(NodeGet(node, "type")) & " pos=" & pos & " caps0=[" & capsStart(0) & "," & capsEnd(0) & "]"
    MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"

    Dim t As String: t = CStr(node(1))
    Dim res As Boolean: res = False
    Dim localEnd As Variant: localEnd = Empty
    
    Select Case t
        Case "seq"
            res = MatchSequence(node("child"), subj, pos, localEnd)

        Case "lit", "dot", "class"
            If LeafMatches(node, t, subj, pos) Then localEnd = pos: res = True
            
        Case "group"
            ' Handle both capturing and non-capturing groups.
            Dim gidx As Long
            gidx = CLng(node("idx"))

            ' --- FULL snapshot before touching ANY capture slots ---
            Dim snapGroup() As Long: snapGroup = CopyCaps()
            If gidx >= 1 Then capsStart(gidx) = pos

            Dim childEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, childEnd) Then
                If gidx >= 1 Then capsEnd(gidx) = CLng(childEnd)
                localEnd = childEnd
                res = True
            Else
                RestoreCaps snapGroup
                res = False
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
            Dim qEnd As Variant
            Dim noTail As New Collection
            If MatchQuant(node, 0, noTail, subj, pos, qEnd) Then localEnd = qEnd: res = True

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

            ' non-consuming; a successful lookahead keeps its captures (ECMAScript)
            If laMatched Then
                capsStart(0) = snapLA(0)
                capsEnd(0) = snapLA(pGroupCount + 1)
            Else
                RestoreCaps snapLA
            End If

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
            If lbMatched Then
                capsStart(0) = snapLB(0)
                capsEnd(0) = snapLB(pGroupCount + 1)
            Else
                RestoreCaps snapLB
            End If

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

            ' re-entering the same target at the same position can never consume input (left recursion)
            Dim recKey As String: recKey = "|" & targ & "@" & pos & "|"
            If InStr(pRecActive, recKey) > 0 Then
                MatchNode = False: Exit Function
            End If
            pRecDepth = pRecDepth + 1
            If pRecDepth > RECURSE_LIMIT Then err.Raise vbObjectError + 3, "RegexEngine", "Match aborted: recursion too deep"
            Dim prevActive As String: prevActive = pRecActive
            pRecActive = pRecActive & recKey
        
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
                        If typeName(targetNode(UBound(targetNode))) = "Collection" Then
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
            pRecActive = prevActive
            pRecDepth = pRecDepth - 1
        
            If recOk Then
                res = True
            Else
                res = False
            End If

        Case "backref"
            Dim brIdx As Long
            If NodeExists(node, "name") Then
                brIdx = FindGroupIndexByName(pAST, CStr(NodeGet(node, "name")))
            Else
                brIdx = CLng(NodeGet(node, "idx"))
            End If
            ' A group that did not participate matches the empty string (ECMAScript/VBScript)
            Dim brLen As Long: brLen = 0
            If brIdx >= 1 And brIdx <= pGroupCount Then
                If capsStart(brIdx) >= 1 And capsEnd(brIdx) >= capsStart(brIdx) Then brLen = capsEnd(brIdx) - capsStart(brIdx) + 1
            End If
            If brLen = 0 Then
                localEnd = pos - 1: res = True
            ElseIf pos + brLen - 1 <= Len(subj) Then
                Dim brText As String: brText = mid$(subj, pos, brLen)
                Dim brCap As String: brCap = mid$(subj, capsStart(brIdx), brLen)
                If FlagState(node, "i", pIgnoreCase) Then brText = LCase$(brText): brCap = LCase$(brCap)
                If brText = brCap Then localEnd = pos + brLen - 1: res = True
            End If

        Case "wordb"
            If (IsWordCharAt(subj, pos - 1) = IsWordCharAt(subj, pos)) = CBool(NodeGet(node, "neg")) Then
                outEnd = pos - 1
                MatchNode = True
            Else
                MatchNode = False
            End If
            Exit Function

        Case "bol"
            If pos = 1 Then
                res = True
            ElseIf FlagState(node, "m", pMultiline) Then
                res = (SafeMid(subj, pos - 1, 1) = vbLf Or SafeMid(subj, pos - 1, 1) = vbCr)
            End If
            If res Then outEnd = pos - 1
            MatchNode = res
            Exit Function

        Case "eol"
            If pos = Len(subj) + 1 Then
                res = True
            ElseIf FlagState(node, "m", pMultiline) Then
                res = (SafeMid(subj, pos, 1) = vbLf Or SafeMid(subj, pos, 1) = vbCr)
            End If
            If res Then outEnd = pos - 1
            MatchNode = res
            Exit Function

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


' Flatten nested sequences.
Private Sub AppendNodeToMatchSequence(ByRef target As Collection, ByVal sourceNode As Collection)
    If CStr(sourceNode(1)) = "seq" Then
        Dim sourceChild As Variant
        For Each sourceChild In NodeGet(sourceNode, "child")
            target.Add sourceChild
        Next sourceChild
    Else
        target.Add sourceNode
    End If
End Sub

Private Function MatchSequence(ByVal children As Collection, ByRef subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    If DEBUG_TRACE Then Debug.Print "TRACE MatchSequence start pos=" & pos & " children=" & children.Count
    Dim rootSnap() As Long: rootSnap = CopyCaps()
    Dim grpSnap() As Long
    Dim curPos As Long: curPos = pos
    Dim i As Long

    For i = 1 To children.Count
        Dim node As Collection: Set node = children(i)
        Dim nodeEnd As Variant
        Dim nodeType As String: nodeType = node(1)

        If DEBUG_TRACE Then Debug.Print "TRACE  seq child#" & i & " type=" & nodeType & " curPos=" & curPos

        If nodeType = "quant" Then
            ' Quantifier followed by the rest of the sequence, with backtracking.
            If MatchQuant(node, 0, SliceChildren(children, i + 1), subj, curPos, outEnd) Then
                MatchSequence = True
            Else
                RestoreCaps rootSnap
                MatchSequence = False
            End If
            Exit Function

        ElseIf nodeType = "quant_cont" Then
            ' Next iteration of a quantifier whose child has internal choices.
            If MatchQuant(node("quant"), CLng(node("count")), SliceChildren(children, i + 1), subj, curPos, outEnd) Then
                MatchSequence = True
            Else
                RestoreCaps rootSnap
                MatchSequence = False
            End If
            Exit Function

        ElseIf nodeType = "quant_guard" Then
            ' An iteration beyond the minimum must consume input (ECMAScript RepeatMatcher).
            If curPos = CLng(node("start")) Then
                RestoreCaps rootSnap
                MatchSequence = False: Exit Function
            End If

        ElseIf nodeType = "alt" Then
            ' Try each alternative followed by the rest of the sequence.
            Dim altSnap() As Long: altSnap = CopyCaps()
            Dim altItem As Variant
            For Each altItem In NodeGet(node, "child")
                Dim altSeq As Collection: Set altSeq = New Collection
                AppendNodeToMatchSequence altSeq, altItem
                Dim altTail As Long
                For altTail = i + 1 To children.Count
                    altSeq.Add children(altTail)
                Next altTail
                Dim altEnd As Variant
                If MatchSequence(altSeq, subj, curPos, altEnd) Then
                    outEnd = altEnd
                    capsEnd(0) = CLng(outEnd)
                    MatchSequence = True: Exit Function
                End If
                RestoreCaps altSnap
            Next altItem
            RestoreCaps rootSnap
            MatchSequence = False: Exit Function

        ElseIf nodeType = "group" Then
            Dim gidx As Long
            Dim captureEndNode As Collection
            ' idx=0 is non-capturing; the capture_end marker is built once per group
            If NodeExists(node, "capEnd") Then
                Set captureEndNode = node("capEnd")
            Else
                gidx = CLng(node("idx"))
                Set captureEndNode = newNode("capture_end")
                captureEndNode.Add gidx, "idx"
                node.Add captureEndNode, "capEnd"
            End If
            gidx = captureEndNode("idx")

            ' Match group and tail as one backtracking sequence.
            grpSnap = CopyCaps()
            If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group idx=" & gidx & " set start=" & curPos
            If gidx >= 1 Then capsStart(gidx) = curPos

            Dim groupCombinedSeq As New Collection
            Dim groupCombinedEnd As Variant
            Dim groupTailIdx As Long

            AppendNodeToMatchSequence groupCombinedSeq, node("child")
            groupCombinedSeq.Add captureEndNode
            For groupTailIdx = i + 1 To children.Count
                groupCombinedSeq.Add children(groupTailIdx)
            Next groupTailIdx

            If MatchSequence(groupCombinedSeq, subj, curPos, groupCombinedEnd) Then
                outEnd = groupCombinedEnd
                capsEnd(0) = CLng(outEnd)
                If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group idx=" & gidx & " OK outEnd=" & outEnd
                MatchSequence = True: Exit Function
            End If

            If DEBUG_TRACE Then Debug.Print "TRACE   seq child#" & i & " group idx=" & gidx & " FAILED, restoring grpSnap"
            RestoreCaps grpSnap
            MatchSequence = False: Exit Function

        ElseIf nodeType = "capture_end" Then
            ' Record the capture boundary.
            Dim captureEndIdx As Long: captureEndIdx = node("idx")
            If captureEndIdx >= 1 Then capsEnd(captureEndIdx) = curPos - 1
            If DEBUG_TRACE Then Debug.Print "TRACE   capture_end idx=" & captureEndIdx & " end=" & (curPos - 1)
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

Private Function SliceChildren(ByVal children As Collection, ByVal fromIdx As Long) As Collection
    Dim res As New Collection
    Dim j As Long
    For j = fromIdx To children.Count
        res.Add children(j)
    Next j
    Set SliceChildren = res
End Function

' True when a node can match in at most one way at a given position (memoized on the node).
Private Function IsDeterministic(ByVal node As Collection) As Boolean
    If NodeExists(node, "det") Then IsDeterministic = node("det"): Exit Function
    Dim det As Boolean: det = True
    Dim subNode As Variant
    Select Case CStr(node(1))
        Case "seq"
            For Each subNode In node("child")
                If Not IsDeterministic(subNode) Then det = False: Exit For
            Next subNode
        Case "group"
            det = IsDeterministic(NodeGet(node, "child"))
        Case "quant"
            If NodeGet(node, "possessive") Or NodeGet(node, "min") = NodeGet(node, "max") Then
                det = IsDeterministic(NodeGet(node, "child"))
            Else
                det = False
            End If
        Case "alt"
            Dim alts As Collection: Set alts = NodeGet(node, "child")
            If alts.Count = 1 Then det = IsDeterministic(alts(1)) Else det = False
    End Select
    NodeAdd node, "det", det
    IsDeterministic = det
End Function

Private Sub CollectGroupIdx(ByVal node As Variant, ByRef idxList As String)
    If typeName(node) <> "Collection" Then Exit Sub
    If NodeExists(node, "type") Then
        If NodeGet(node, "type") = "group" Then
            If CLng(node("idx")) > 0 Then idxList = idxList & node("idx") & ","
        End If
    End If
    Dim elm As Variant
    For Each elm In node
        If typeName(elm) = "Collection" Then CollectGroupIdx elm, idxList
    Next elm
End Sub

' Captures inside a quantified atom start unset on every iteration (ECMAScript).
Private Sub ResetGroupCaps(ByVal qnode As Collection)
    If Not NodeExists(qnode, "capIdx") Then
        Dim lst As String
        CollectGroupIdx NodeGet(qnode, "child"), lst
        NodeAdd qnode, "capIdx", lst
    End If
    Dim parts() As String: parts = VBA.Split(CStr(NodeGet(qnode, "capIdx")), ",")
    Dim k As Long
    For k = 0 To UBound(parts)
        If Len(parts(k)) > 0 Then
            capsStart(CLng(parts(k))) = 0
            capsEnd(CLng(parts(k))) = -1
        End If
    Next k
End Sub

' Match quantifier qnode, count iterations already done, followed by tail.
' Children with internal choices are matched one iteration at a time so the tail can backtrack into them.
Private Function MatchQuant(ByVal qnode As Collection, ByVal count As Long, ByVal tail As Collection, ByRef subj As String, ByVal pos As Long, ByRef outEnd As Variant) As Boolean
    Dim childNode As Collection: Set childNode = qnode("child")
    Dim mn As Long: mn = qnode("min")
    Dim mx As Long: mx = qnode("max")
    Dim lazy As Boolean: lazy = qnode("lazy")
    Dim possFlag As Boolean: possFlag = qnode("possessive")
    Dim remMin As Long: remMin = mn - count
    If remMin < 0 Then remMin = 0
    Dim remMax As Long
    If mx = -1 Then remMax = -1 Else remMax = mx - count

    If possFlag Or pQuantDepth >= QUANT_REC_LIMIT Or IsDeterministic(childNode) Then
        MatchQuant = MatchQuantIter(childNode, remMin, remMax, lazy, possFlag, tail, subj, pos, outEnd)
        Exit Function
    End If

    Dim snap() As Long: snap = CopyCaps()
    If lazy And remMin = 0 Then
        If MatchSequence(tail, subj, pos, outEnd) Then MatchQuant = True: Exit Function
        RestoreCaps snap
    End If
    If remMax <> 0 Then
        Dim iter As Collection: Set iter = New Collection
        AppendNodeToMatchSequence iter, childNode
        If remMin = 0 Then
            Dim guardNode As Collection: Set guardNode = newNode("quant_guard")
            guardNode.Add pos, "start"
            iter.Add guardNode
        End If
        Dim contNode As Collection: Set contNode = newNode("quant_cont")
        contNode.Add qnode, "quant"
        contNode.Add count + 1, "count"
        iter.Add contNode
        Dim t As Variant
        For Each t In tail
            iter.Add t
        Next t
        ResetGroupCaps qnode
        pQuantDepth = pQuantDepth + 1
        Dim ok As Boolean: ok = MatchSequence(iter, subj, pos, outEnd)
        pQuantDepth = pQuantDepth - 1
        If ok Then MatchQuant = True: Exit Function
        RestoreCaps snap
    End If
    If Not lazy And remMin = 0 Then
        If MatchSequence(tail, subj, pos, outEnd) Then MatchQuant = True: Exit Function
        RestoreCaps snap
    End If
    MatchQuant = False
End Function

' Quantifier over a child with no internal choices: collect the occurrences, then try the tail for each count.
Private Function MatchQuantIter(ByVal childNode As Collection, ByVal mn As Long, ByVal mx As Long, ByVal lazy As Boolean, ByVal possFlag As Boolean, ByVal tail As Collection, ByRef subj As String, ByVal curPos As Long, ByRef outEnd As Variant) As Boolean
    Dim leafType As String: leafType = childNode(1)
    If leafType = "lit" Or leafType = "dot" Or leafType = "class" Then
        MatchQuantIter = MatchLeafRun(childNode, leafType, mn, mx, lazy, possFlag, tail, subj, curPos, outEnd)
        Exit Function
    End If

    Dim seqSnap() As Long: seqSnap = CopyCaps()
    Dim occEnds() As Long
    Dim occSnaps() As Variant
    Dim occCount As Long, occCap As Long
    Dim scanPos As Long: scanPos = curPos
    Dim preSnap() As Long
    Dim ndEnd As Variant
    Dim zeroLen As Boolean
    Do While mx = -1 Or occCount < mx
        preSnap = CopyCaps()
        If Not MatchNode(childNode, subj, scanPos, ndEnd) Then
            RestoreCaps preSnap
            Exit Do
        End If
        ' an empty iteration repeats identically; it only counts towards the minimum
        zeroLen = (CLng(ndEnd) < scanPos)
        If zeroLen And occCount >= mn Then
            RestoreCaps preSnap
            Exit Do
        End If
        Do
            occCount = occCount + 1
            If occCount > occCap Then
                occCap = occCap * 2 + 4
                ReDim Preserve occEnds(1 To occCap)
                ReDim Preserve occSnaps(1 To occCap)
            End If
            occEnds(occCount) = CLng(ndEnd)
            occSnaps(occCount) = CopyCaps()
        Loop While zeroLen And occCount < mn
        If zeroLen Then Exit Do
        scanPos = CLng(ndEnd) + 1
    Loop
    If occCount < mn Then
        RestoreCaps seqSnap
        MatchQuantIter = False: Exit Function
    End If

    ' possessive: only the maximal count; lazy: small to big; greedy: big to small
    Dim kStart As Long, kEnd As Long, kStep As Long
    If possFlag Then
        kStart = occCount: kEnd = occCount: kStep = 1
    ElseIf lazy Then
        kStart = mn: kEnd = occCount: kStep = 1
    Else
        kStart = occCount: kEnd = mn: kStep = -1
    End If
    Dim k As Long, posTry As Long
    Dim kSnap() As Long
    For k = kStart To kEnd Step kStep
        RestoreCaps seqSnap
        If k >= 1 Then
            kSnap = occSnaps(k)
            RestoreCaps kSnap
            posTry = occEnds(k) + 1
        Else
            posTry = curPos
        End If
        If MatchSequence(tail, subj, posTry, outEnd) Then
            MatchQuantIter = True: Exit Function
        End If
    Next k
    RestoreCaps seqSnap
    MatchQuantIter = False
End Function

' Quantified single-char node: captures cannot change, so occurrence k ends at curPos + k - 1
Private Function MatchLeafRun(ByVal childNode As Collection, ByVal leafType As String, ByVal mn As Long, ByVal mx As Long, ByVal lazy As Boolean, ByVal possFlag As Boolean, ByVal tail As Collection, ByRef subj As String, ByVal curPos As Long, ByRef outEnd As Variant) As Boolean
    Dim seqSnap() As Long: seqSnap = CopyCaps()
    Dim n As Long
    Do While n < mn
        If Not LeafStep(childNode, leafType, subj, curPos + n) Then RestoreCaps seqSnap: Exit Function
        n = n + 1
    Loop
    If lazy And Not possFlag Then
        ' extend one char at a time, trying the tail first
        Do
            If MatchSequence(tail, subj, curPos + n, outEnd) Then MatchLeafRun = True: Exit Function
            RestoreCaps seqSnap
            If n = mx Then Exit Do
            If Not LeafStep(childNode, leafType, subj, curPos + n) Then Exit Do
            n = n + 1
        Loop
    Else
        Do While mx = -1 Or n < mx
            If Not LeafStep(childNode, leafType, subj, curPos + n) Then Exit Do
            n = n + 1
        Loop
        ' possessive: only the maximal count; greedy: big to small
        Dim k As Long, kEnd As Long
        If possFlag Then kEnd = n Else kEnd = mn
        For k = n To kEnd Step -1
            If MatchSequence(tail, subj, curPos + k, outEnd) Then MatchLeafRun = True: Exit Function
            RestoreCaps seqSnap
        Next k
    End If
    RestoreCaps seqSnap
End Function

' One char of a leaf run, counted like a MatchNode call
Private Function LeafStep(ByVal node As Collection, ByVal t As String, ByRef subj As String, ByVal pos As Long) As Boolean
    MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"
    LeafStep = LeafMatches(node, t, subj, pos)
End Function

' lit, dot and class nodes: does the char at pos match?
Private Function LeafMatches(ByVal node As Collection, ByVal t As String, ByRef subj As String, ByVal pos As Long) As Boolean
    If pos > Len(subj) Then Exit Function
    Dim sc As String: sc = mid$(subj, pos, 1)
    Select Case t
        Case "lit"
            If FlagState(node, "i", pIgnoreCase) Then LeafMatches = (LCase$(sc) = LCase$(CStr(node("ch")))) Else LeafMatches = (sc = node("ch"))
        Case "dot"
            LeafMatches = (sc <> vbLf And sc <> vbCr) Or FlagState(node, "s", pDotAll)
        Case "class"
            LeafMatches = MatchCharClassNode(node, sc)
    End Select
End Function

' MatchQuantifier: tries to match between min..max occurrences.
' For greedy: try larger k first; for lazy: try smaller first.
' If possessive = True -> accept the maximal match and DO NOT try smaller k's.
Private Function MatchQuantifier(ByVal child As Collection, ByVal minc As Long, ByVal maxc As Long, ByVal lazy As Boolean, ByRef subj As String, ByVal posStart As Long, ByVal possessive As Boolean, ByRef outEnd As Long) As Boolean
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
    Exit Sub

bad_snapshot:
    ' Defensive fallback: if snapshot was bogus, do nothing but avoid crashing.
    On Error Resume Next
    err.Clear
End Sub

' ---------------- Char class matching with precomputed map -----------------
Private Function MatchCharClassNode(ByVal node As Collection, ByVal ch As String) As Boolean
    Dim iCase As Boolean: iCase = FlagState(node, "i", pIgnoreCase)
    Dim code As Long: code = CodeOf(ch)
    If code >= 0 And code <= 255 Then
        Dim mapIdx As Long: mapIdx = node("mapIdx")
        Dim matched As Boolean: matched = False

        ' direct ASCII fast-path
        If code >= 0 And code <= 255 Then
            matched = pClassMaps(code, mapIdx)
            ' if ignore-case, check the swapped-case ASCII code too (A<->a)
            If iCase And Not matched Then
                If code >= 65 And code <= 90 Then
                    matched = pClassMaps(code + 32, mapIdx)
                ElseIf code >= 97 And code <= 122 Then
                    matched = pClassMaps(code - 32, mapIdx)
                End If
            End If
        End If

        If node("neg") Then
            MatchCharClassNode = Not matched
        Else
            MatchCharClassNode = matched
        End If
        Exit Function
    End If

    ' Non-ASCII fallback: brute-force check, but respect ignore-case
    Dim items As Collection: Set items = node("chars")
    Dim matched2 As Boolean: matched2 = False
    Dim i As Long
    For i = 1 To items.Count
        Dim it As Variant: it = items(i)
        If IsArray(it) Then
            If it(0) = "range" Then
                If Not iCase Then
                    If CodeOf(ch) >= CodeOf(it(1)) And CodeOf(ch) <= CodeOf(it(2)) Then matched2 = True
                Else
                    ' case-insensitive range test: compare lowercased codepoints where possible
                    Dim lowerCh As String: lowerCh = LCase$(ch)
                    Dim lowerA As String: lowerA = LCase$(it(1))
                    Dim lowerB As String: lowerB = LCase$(it(2))
                    If CodeOf(lowerCh) >= CodeOf(lowerA) And CodeOf(lowerCh) <= CodeOf(lowerB) Then matched2 = True
                End If
            End If
        Else
            If Not iCase Then
                If ch = it Then matched2 = True
            Else
                If LCase$(ch) = LCase$(CStr(it)) Then matched2 = True
            End If
        End If
        If matched2 Then Exit For
    Next i
    If node("neg") Then MatchCharClassNode = Not matched2 Else MatchCharClassNode = matched2
End Function

' Word char test for \b and \B, same ASCII set as \w
Private Function IsWordCharAt(ByRef subj As String, ByVal pos As Long) As Boolean
    If pos < 1 Or pos > Len(subj) Then Exit Function
    Dim code As Long: code = AscW(mid$(subj, pos, 1))
    IsWordCharAt = (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 95
End Function
' ---------------- Replacement helpers -----------------------------------
' Unset groups yield vbNullString so group positions stay stable
Private Function BuildCaptureCollection() As Collection
    Dim col As New Collection
    If capsStart(0) >= 1 And capsEnd(0) >= capsStart(0) Then
        col.Add SafeMid(CurSubject, capsStart(0), capsEnd(0) - capsStart(0) + 1)
    Else
        If col.Count = 0 Then
            col.Add vbNullString
        End If
    End If
    Dim i As Long
    For i = 1 To pGroupCount
        If capsStart(i) >= 1 And capsEnd(i) >= capsStart(i) Then
            col.Add SafeMid(CurSubject, capsStart(i), capsEnd(i) - capsStart(i) + 1)
        Else
            col.Add vbNullString
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
' ---------------- Public Exec / ExecAll / Replace / Split ---------------------------
' Exec: find first match anywhere (anchors are matched by bol/eol nodes)
Public Function Exec(ByVal subj As String) As Collection
    EnsureCompiled
    If FindFirst(subj) Then
        Set Exec = BuildCaptureCollection()
        UpdateNamedCapturesFromAST
    Else
        Set Exec = Nothing
    End If
End Function

' First match at or after position 1; True leaves it in capsStart/capsEnd
Private Function FindFirst(ByRef subj As String) As Boolean
    BeginScan subj
    Dim lastStart As Long: lastStart = StartLimit(subj)
    Dim i As Long: i = NextCandidate(subj, 1, lastStart)
    Do While i > 0
        If TryAt(subj, i) Then FindFirst = True: Exit Function
        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
        i = NextCandidate(subj, i + 1, lastStart)
    Loop
End Function

' Last start position that leaves room for the pattern's minimum length
Private Function StartLimit(ByRef subj As String) As Long
    Dim ast As Collection: Set ast = pAST
    Dim minLen As Long: minLen = ast("minLen")
    If minLen > 0 Then StartLimit = Len(subj) - minLen + 1 Else StartLimit = Len(subj) + 1
    If StartLimit < 1 Then StartLimit = 1
End Function

Private Sub BeginScan(ByRef subj As String)
    CurSubject = subj
    pSubjBytes = subj
    MatchSteps = 0
End Sub

' ExecAll: find all non-overlapping matches anywhere in subject
' Returns a Collection where each item is a Collection of captures (same structure as BuildCaptureCollection)
Public Function ExecAll(ByVal subj As String) As Collection
    EnsureCompiled
    BeginScan subj
    Dim results As New Collection
    Dim lastStart As Long: lastStart = StartLimit(subj)
    Dim pos As Long: pos = NextCandidate(subj, 1, lastStart)
    Dim mstart As Long, mend As Long
    Do While pos > 0
        If TryAt(subj, pos) Then
            results.Add BuildCaptureCollection()
            UpdateNamedCapturesFromAST
            mstart = capsStart(0): mend = capsEnd(0)
            If mstart < pos Then mstart = pos
            ' non-overlapping; an empty match advances one position
            If mend - mstart + 1 <= 0 Then pos = pos + 1 Else pos = mend + 1
        Else
            pos = pos + 1
        End If
        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
        pos = NextCandidate(subj, pos, lastStart)
    Loop
    If results.Count = 0 Then Set ExecAll = Nothing Else Set ExecAll = results
End Function

' ExecAt: try to match body at absolute position startPos. Returns Collection (captures filled) or Nothing
Public Function ExecAt(ByVal subj As String, ByVal startPos As Long) As Collection
    EnsureCompiled
    If startPos < 1 Or startPos > Len(subj) + 1 Then Set ExecAt = Nothing: Exit Function
    If Not CanStartAt(subj, startPos) Then Set ExecAt = Nothing: Exit Function
    CurSubject = subj
    MatchSteps = 0
    If TryAt(subj, startPos) Then
        Set ExecAt = BuildCaptureCollection()
        UpdateNamedCapturesFromAST
    Else
        Set ExecAt = Nothing
    End If
End Function

' One match attempt at startPos; on success the match is in capsStart/capsEnd
Private Function TryAt(ByRef subj As String, ByVal startPos As Long) As Boolean
    MatchSteps = MatchSteps + 1
    pQuantDepth = 0
    pRecDepth = 0
    pRecActive = vbNullString
    Dim totalGroups As Long: totalGroups = pGroupCount
    If pCapsSize <> totalGroups + 1 Then
        ReDim capsStart(0 To totalGroups) As Long
        ReDim capsEnd(0 To totalGroups) As Long
        ReDim atomicLocks(0 To totalGroups) As Boolean
        pCapsSize = totalGroups + 1
    End If
    Dim i As Long
    For i = 0 To totalGroups
        capsStart(i) = 0: capsEnd(i) = -1
        atomicLocks(i) = False
    Next i
    capsStart(0) = startPos
    Dim outEnd As Variant
    Dim ast As Collection: Set ast = pAST
    If MatchNode(ast("body"), subj, startPos, outEnd) Then
        If Not IsEmpty(outEnd) Then capsEnd(0) = CLng(outEnd)
        TryAt = True
    End If
End Function

' Can a match start at p? Checks the leading ^, the leading literal and the first-char map
Private Function CanStartAt(ByRef subj As String, ByVal p As Long) As Boolean
    If Not pAnchorNode Is Nothing And p > 1 Then
        If Not FlagState(pAnchorNode, "m", pMultiline) Then Exit Function
        Dim prev As String: prev = mid$(subj, p - 1, 1)
        If prev <> vbLf And prev <> vbCr Then Exit Function
    End If
    If Len(pPrefix) > 0 Then
        If Not FlagState(pPrefixNode, "i", pIgnoreCase) Then
            CanStartAt = (mid$(subj, p, Len(pPrefix)) = pPrefix)
            Exit Function
        End If
    End If
    If pHasFirst Then
        If p > Len(subj) Then Exit Function
        Dim c0 As Long: c0 = CodeOf(mid$(subj, p, 1))
        If c0 > 255 Then
            If Not pFirstHigh Then Exit Function
        ElseIf Not pFirstMap(c0) Then
            Exit Function
        End If
    End If
    CanStartAt = True
End Function

' First position in p..lastPos where CanStartAt holds, 0 when none; reads pSubjBytes (see BeginScan)
Private Function NextCandidate(ByRef subj As String, ByVal p As Long, ByVal lastPos As Long) As Long
    Dim lineStart As Boolean, usePrefix As Boolean, c0 As Long
    If Not pAnchorNode Is Nothing Then
        If FlagState(pAnchorNode, "m", pMultiline) Then
            lineStart = True
        ElseIf p > 1 Then
            Exit Function
        ElseIf lastPos > 1 Then
            lastPos = 1
        End If
    End If
    If Len(pPrefix) > 0 Then usePrefix = Not FlagState(pPrefixNode, "i", pIgnoreCase)
    Dim L As Long: L = Len(subj)
    Do While p <= lastPos
        If usePrefix Then
            p = InStr(p, subj, pPrefix, vbBinaryCompare)
            If p = 0 Or p > lastPos Then Exit Function
        ElseIf pHasFirst Then
            If p > L Then Exit Function
            c0 = pSubjBytes(2 * p - 2) + 256& * pSubjBytes(2 * p - 1)
            If c0 > 255 Then
                If Not pFirstHigh Then GoTo nextPos
            ElseIf Not pFirstMap(c0) Then
                GoTo nextPos
            End If
        End If
        If lineStart And p > 1 Then
            c0 = pSubjBytes(2 * p - 4) + 256& * pSubjBytes(2 * p - 3)
            If c0 <> 10 And c0 <> 13 Then GoTo nextPos
        End If
        NextCandidate = p
        Exit Function
nextPos:
        p = p + 1
    Loop
End Function

Public Function Replace(ByVal subj As String, ByVal Replacement As String) As String
    EnsureCompiled
    BeginScan subj
    Dim toks As Collection: Set toks = CompileReplacement(Replacement)
    Dim buf As String, bufLen As Long
    Dim L As Long: L = Len(subj)
    Dim pos As Long: pos = 1
    Dim foundAt As Long, mstart As Long, mend As Long
    Do
        foundAt = NextCandidate(subj, pos, L + 1)
        Do While foundAt > 0
            If TryAt(subj, foundAt) Then Exit Do
            foundAt = NextCandidate(subj, foundAt + 1, L + 1)
        Loop
        If foundAt = 0 Then
            If pos <= L Then BufAppend buf, bufLen, mid$(subj, pos)
            Exit Do
        End If
        UpdateNamedCapturesFromAST
        mstart = capsStart(0): mend = capsEnd(0)
        If mstart < pos Then mstart = pos
        If mstart > pos Then BufAppend buf, bufLen, mid$(subj, pos, mstart - pos)
        AppendReplacement buf, bufLen, toks
        If mend - mstart + 1 <= 0 Then
            ' zero-length match: keep the char at the match position and advance past it
            If foundAt <= L Then BufAppend buf, bufLen, mid$(subj, foundAt, 1)
            pos = foundAt + 1
        Else
            pos = mend + 1
        End If
    Loop
    Replace = left$(buf, bufLen)
End Function

' Replacement as tokens, parsed once per call: Array(0, text), Array(1, group or -1), Array(2) for $`, Array(3) for $'
Private Function CompileReplacement(ByVal repl As String) As Collection
    Dim toks As New Collection
    Dim lit As String, nextCh As String, numStr As String
    Dim i As Long: i = 1
    Dim L As Long: L = Len(repl)
    Dim j As Long, n As Long, useIdx As Long
    Do While i <= L
        If mid$(repl, i, 1) <> "$" Or i = L Then
            lit = lit & mid$(repl, i, 1): i = i + 1
        Else
            nextCh = mid$(repl, i + 1, 1)
            Select Case nextCh
                Case "$"
                    lit = lit & "$": i = i + 2
                Case "&", "`", "'"
                    FlushLiteral toks, lit
                    toks.Add Array(InStr("&`'", nextCh), 0)
                    i = i + 2
                Case "<"
                    ' $<name> needs named groups and a closing '>'; a missing name gives an empty string
                    j = InStr(i + 2, repl, ">")
                    If pNamedCount = 0 Or j = 0 Then
                        lit = lit & "$<": i = i + 2
                    Else
                        FlushLiteral toks, lit
                        toks.Add Array(1, NamedGroupIndex(mid$(repl, i + 2, j - i - 2)))
                        i = j + 1
                    End If
                Case "0" To "9"
                    ' greedy two digits, then one; "$n" with no such group stays literal (ECMAScript)
                    numStr = nextCh
                    If i + 2 <= L Then
                        If mid$(repl, i + 2, 1) >= "0" And mid$(repl, i + 2, 1) <= "9" Then numStr = numStr & mid$(repl, i + 2, 1)
                    End If
                    useIdx = -1
                    n = CLng(numStr)
                    If Len(numStr) = 2 And n <= pGroupCount Then
                        useIdx = n: i = i + 3
                    Else
                        n = CLng(left$(numStr, 1))
                        If n <= pGroupCount Then useIdx = n
                        i = i + 2
                    End If
                    If useIdx >= 0 Then
                        FlushLiteral toks, lit
                        toks.Add Array(1, useIdx)
                    Else
                        lit = lit & "$" & left$(numStr, 1)
                    End If
                Case Else
                    lit = lit & "$": i = i + 1
            End Select
        End If
    Loop
    FlushLiteral toks, lit
    Set CompileReplacement = toks
End Function

Private Sub FlushLiteral(ByVal toks As Collection, ByRef lit As String)
    If Len(lit) > 0 Then toks.Add Array(0, lit): lit = vbNullString
End Sub

' First named group whose name matches, compared like Collection keys; -1 when none
Private Function NamedGroupIndex(ByVal nm As String) As Long
    NamedGroupIndex = -1
    Dim it As Variant, ast As Collection: Set ast = pAST
    For Each it In ast("groupNodes")
        If StrComp(CStr(it(0)), nm, vbTextCompare) = 0 Then NamedGroupIndex = CLng(it(1)): Exit Function
    Next it
End Function

Private Sub AppendReplacement(ByRef buf As String, ByRef bufLen As Long, ByVal toks As Collection)
    Dim t As Variant
    For Each t In toks
        Select Case t(0)
            Case 0
                BufAppend buf, bufLen, CStr(t(1))
            Case 1
                If t(1) >= 0 Then BufAppend buf, bufLen, CapText(CLng(t(1)))
            Case 2
                If capsStart(0) > 1 Then BufAppend buf, bufLen, left$(CurSubject, capsStart(0) - 1)
            Case 3
                If capsEnd(0) >= 1 And capsEnd(0) < Len(CurSubject) Then BufAppend buf, bufLen, mid$(CurSubject, capsEnd(0) + 1)
        End Select
    Next t
End Sub

Private Function CapText(ByVal i As Long) As String
    If capsStart(i) >= 1 And capsEnd(i) >= capsStart(i) Then CapText = mid$(CurSubject, capsStart(i), capsEnd(i) - capsStart(i) + 1)
End Function

' Append to a buffer that grows geometrically; the text is left$(buf, bufLen)
Private Sub BufAppend(ByRef buf As String, ByRef bufLen As Long, ByVal s As String)
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Sub
    If bufLen + n > Len(buf) Then buf = buf & Space$(bufLen + n + 64)
    Mid$(buf, bufLen + 1, n) = s
    bufLen = bufLen + n
End Sub


Public Function Split(ByVal subj As String) As Collection
    Dim res As New Collection
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    If L > 0 Then EnsureCompiled
    BeginScan subj
    Dim foundAt As Long
    Dim mstart As Long
    Dim mend As Long
    Do
        foundAt = NextCandidate(subj, pos, L)
        Do While foundAt > 0
            If TryAt(subj, foundAt) Then Exit Do
            foundAt = NextCandidate(subj, foundAt + 1, L)
        Loop

        If foundAt = 0 Then
            ' no more matches -> append remainder and finish
            res.Add SafeMid(subj, pos)
            Exit Do
        Else
            UpdateNamedCapturesFromAST
            mstart = CLng(capsStart(0))
            mend = CLng(capsEnd(0))

            ' ensure mstart is not before current pos
            If mstart < pos Then mstart = pos
            res.Add SafeMid(subj, pos, mstart - pos)
            ' advance position after the matched separator
            If mend - mstart + 1 <= 0 Then
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
            Dim qMax As Long: qMax = NodeGet(node, "max")
            If qMax = -1 Then
                If mxChild = -1 Then
                    ComputeMaxLen = -1
                Else
                    ComputeMaxLen = -1
                End If
            Else
                If mxChild = -1 Then
                    ComputeMaxLen = -1
                Else
                    ComputeMaxLen = qMax * mxChild
                End If
            End If
        Case "wordb", "bol", "eol"
            ComputeMaxLen = 0
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
        Case "wordb", "bol", "eol", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
            ComputeFixedWidth = 0
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
        Case "group", "quant", "atomic", "lookahead_pos", "lookahead_neg"
            TraverseEnforce NodeGet(node, "child")
        Case "cond"
            TraverseEnforce NodeGet(node, "yes")
            If NodeExists(node, "no") Then TraverseEnforce NodeGet(node, "no")
        Case "alt"
            Dim altCh As Variant
            For Each altCh In NodeGet(node, "child")
                TraverseEnforce altCh
            Next
        Case "lookbehind_pos", "lookbehind_neg"
            ' compute fixed width of the inner child
            Dim child As Collection: Set child = NodeGet(node, "child")
            TraverseEnforce child
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

' Error codes (vbObjectError + N):
'   20  (?# comment syntax (unsupported)
'   31  Dangling backslash at end of pattern
'   32  Unclosed character class [
'   33  Incomplete (?...) specifier or conditional
'   34  Named group (?<name missing closing >
'   35  Named group (?<> has empty name
'   36  Unknown (?< variant (not =, !, or valid name start)
'   37  Unknown (?X specifier
'   38  Unmatched closing )
'   39  Quantifier (*, +, ?, {m,n}) with no preceding quantifiable token
'   40  {m,n} quantifier where min > max
'   41  Character class range out of order, e.g. [z-a]
'   43  Unclosed group(s): unmatched ( at end of pattern
'   44  Backreference to a non-existent group (\n or \k<name>)
'   45  Named backreference \k< missing closing >
'   46  Unicode property escapes \p{..} / \P{..} (unsupported)
'   47  Groups nested deeper than MAX_GROUP_NESTING
'   48  {m,n} quantifier bound longer than 9 digits
' Match-time aborts: 1/2 too many steps, 3 recursion too deep, 4 no pattern (Init not called)
Private Sub ValidatePattern(ByVal pat As String)
    Dim L As Long: L = Len(pat)
    If L = 0 Then Exit Sub

    Dim i As Long: i = 1
    Dim c As String
    Dim depth As Long: depth = 0
    Dim prevQ As Boolean: prevQ = False
    Dim esc As Boolean: esc = False
    Dim advance As Long
    Dim spec As String
    Dim clEnd As Long
    Dim lbCh As String
    Dim gtPos As Long
    Dim condClose As Long
    Dim ri As Long
    Dim fi As Long
    Dim flagCh As String
    Dim nextMod As String
    Dim closeB As Long
    Dim qcontent As String
    Dim qparts() As String
    Dim qMin As Long
    Dim qMax As Long
    Dim capCount As Long
    Dim maxRef As Long
    Dim refStr As String
    Dim refNames As String
    Dim groupNames As String
    Dim refList() As String

    Do While i <= L
        c = mid$(pat, i, 1)
        advance = 1
        If esc Then
            esc = False
            ' \b and \B are assertions, like ^ and $ they cannot be quantified
            prevQ = (c <> "b" And c <> "B")
            If c >= "1" And c <= "9" Then
                ri = i
                Do While ri < L
                    If mid$(pat, ri + 1, 1) < "0" Or mid$(pat, ri + 1, 1) > "9" Then Exit Do
                    ri = ri + 1
                Loop
                refStr = mid$(pat, i, ri - i + 1)
                If Len(refStr) > 9 Then
                    maxRef = 1000000000
                ElseIf CLng(refStr) > maxRef Then
                    maxRef = CLng(refStr)
                End If
                advance = ri - i + 1
            ElseIf (c = "p" Or c = "P") And mid$(pat, i + 1, 1) = "{" Then
                err.Raise vbObjectError + 46, "RegexEngine", _
                    "Invalid pattern: Unicode property escapes (\" & c & "{...}) are not supported (position " & (i - 1) & ")"
            ElseIf c = "k" And mid$(pat, i + 1, 1) = "<" Then
                gtPos = InStr(i + 2, pat, ">")
                If gtPos = 0 Then
                    err.Raise vbObjectError + 45, "RegexEngine", _
                        "Invalid pattern: unclosed named backreference '\k<' at position " & (i - 1)
                End If
                refNames = refNames & "|" & mid$(pat, i + 2, gtPos - i - 2)
                advance = gtPos - i + 1
            End If
        Else
            Select Case c
                Case "\"
                    If i = L Then
                        err.Raise vbObjectError + 31, "RegexEngine", _
                            "Invalid pattern: dangling backslash at end of pattern"
                    End If
                    esc = True
                Case "["
                    clEnd = FindClassEnd(pat, i, L)
                    If clEnd = -1 Then
                        err.Raise vbObjectError + 32, "RegexEngine", _
                            "Invalid pattern: unclosed character class '[' at position " & i
                    End If
                    advance = clEnd - i + 1
                    prevQ = True
                Case "("
                    depth = depth + 1
                    If depth > MAX_GROUP_NESTING Then
                        err.Raise vbObjectError + 47, "RegexEngine", _
                            "Invalid pattern: groups nested deeper than " & MAX_GROUP_NESTING & " levels at position " & i
                    End If
                    If i + 1 <= L And mid$(pat, i + 1, 1) = "?" Then
                        If i + 2 > L Then
                            err.Raise vbObjectError + 33, "RegexEngine", _
                                "Invalid pattern: '(?' at end of pattern (position " & i & ")"
                        End If
                        spec = mid$(pat, i + 2, 1)
                        Select Case spec
                            Case ":", "=", "!", ">", "R"
                                advance = 3
                            Case "<"
                                ' lookbehind (?<= (?<! or named group (?<name>
                                If i + 3 > L Then
                                    err.Raise vbObjectError + 33, "RegexEngine", _
                                        "Invalid pattern: '(?<' at end of pattern (position " & i & ")"
                                End If
                                lbCh = mid$(pat, i + 3, 1)
                                If lbCh = "=" Or lbCh = "!" Then
                                    advance = 4
                                ElseIf lbCh = ">" Then
                                    err.Raise vbObjectError + 35, "RegexEngine", _
                                        "Invalid pattern: empty named group name '(?<>' at position " & i
                                ElseIf (lbCh >= "A" And lbCh <= "Z") Or _
                                       (lbCh >= "a" And lbCh <= "z") Or lbCh = "_" Then
                                    gtPos = InStr(i + 3, pat, ">")
                                    If gtPos = 0 Then
                                        err.Raise vbObjectError + 34, "RegexEngine", _
                                            "Invalid pattern: unclosed named group '(?<name' - missing '>' at position " & i
                                    End If
                                    advance = gtPos - i + 1
                                    capCount = capCount + 1
                                    groupNames = groupNames & "|" & mid$(pat, i + 3, gtPos - i - 3) & "|"
                                Else
                                    err.Raise vbObjectError + 36, "RegexEngine", _
                                        "Invalid pattern: unknown '(?<" & lbCh & _
                                        "' at position " & i & " (expected '=', '!' or a name start char)"
                                End If
                            Case "#"
                                ' inline comment - unsupported
                                err.Raise vbObjectError + 20, "RegexEngine", _
                                    "Invalid pattern: (?# comments are not supported (position " & i & ")"
                            Case "("
                                condClose = InStr(i + 3, pat, ")")
                                If condClose = 0 Then
                                    err.Raise vbObjectError + 33, "RegexEngine", _
                                        "Invalid pattern: unclosed conditional '(?(' at position " & i
                                End If
                                advance = condClose - i + 1
                            Case "0" To "9"
                                ri = i + 2
                                Do While ri <= L
                                    If mid$(pat, ri, 1) >= "0" And mid$(pat, ri, 1) <= "9" Then
                                        ri = ri + 1
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                advance = ri - i
                            Case Else
                                fi = i + 2
                                Do While fi <= L
                                    flagCh = mid$(pat, fi, 1)
                                    If flagCh = ":" Or flagCh = ")" Then Exit Do
                                    If (flagCh >= "a" And flagCh <= "z") Or _
                                       (flagCh >= "A" And flagCh <= "Z") Or flagCh = "-" Then
                                        fi = fi + 1
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                If fi <= L And _
                                   (mid$(pat, fi, 1) = ":" Or mid$(pat, fi, 1) = ")") Then
                                    advance = fi - i + 1
                                    If mid$(pat, fi, 1) = ")" Then
                                        depth = depth - 1
                                        If depth < 0 Then
                                            err.Raise vbObjectError + 38, "RegexEngine", _
                                                "Invalid pattern: unmatched ')' at position " & fi
                                        End If
                                    End If
                                Else
                                    err.Raise vbObjectError + 37, "RegexEngine", _
                                        "Invalid pattern: unknown group specifier '(?" & _
                                        spec & "' at position " & i
                                End If

                        End Select
                    Else
                        capCount = capCount + 1
                    End If
                    prevQ = False
                Case ")"
                    depth = depth - 1
                    If depth < 0 Then
                        err.Raise vbObjectError + 38, "RegexEngine", _
                            "Invalid pattern: unmatched ')' at position " & i
                    End If
                    prevQ = True
                Case "*", "+", "?"
                    If Not prevQ Then
                        err.Raise vbObjectError + 39, "RegexEngine", _
                            "Invalid pattern: quantifier '" & c & "' at position " & i & _
                            " has nothing to quantify"
                    End If
                    If i + 1 <= L Then
                        nextMod = mid$(pat, i + 1, 1)
                        If nextMod = "?" Or nextMod = "+" Then advance = 2
                    End If
                    prevQ = False
                Case "{"
                    If IsBraceQuantifier(pat, i, L) Then
                        closeB = InStr(i, pat, "}")
                        qcontent = mid$(pat, i + 1, closeB - i - 1)
                        qparts = VBA.Split(qcontent, ",")
                        If Not prevQ Then
                            err.Raise vbObjectError + 39, "RegexEngine", _
                                "Invalid pattern: quantifier '{" & qcontent & _
                                "}' at position " & i & " has nothing to quantify"
                        End If
                        If Len(qparts(0)) > 9 Or Len(qparts(UBound(qparts))) > 9 Then
                            err.Raise vbObjectError + 48, "RegexEngine", _
                                "Invalid pattern: quantifier bound {" & qcontent & "} is too large at position " & i
                        End If
                        qMin = CLng(qparts(0))
                        If UBound(qparts) = 1 Then
                            If Len(qparts(1)) > 0 Then
                                qMax = CLng(qparts(1))
                                If qMax < qMin Then
                                    err.Raise vbObjectError + 40, "RegexEngine", _
                                        "Invalid pattern: quantifier {" & qcontent & _
                                        "} has min > max (min=" & qMin & _
                                        ", max=" & qMax & ") at position " & i
                                End If
                            End If
                        End If
                        advance = closeB - i + 1
                        If i + advance <= L Then
                            nextMod = mid$(pat, i + advance, 1)
                            If nextMod = "?" Or nextMod = "+" Then advance = advance + 1
                        End If
                        prevQ = False
                    Else
                        prevQ = True
                    End If
                Case "|"
                    prevQ = False
                Case "^", "$"
                    prevQ = False
                Case Else
                    prevQ = True
            End Select
        End If
        i = i + advance
    Loop
    If depth > 0 Then
        err.Raise vbObjectError + 43, "RegexEngine", _
            "Invalid pattern: " & depth & " unclosed group(s) - " & _
            depth & " unmatched '(' in pattern"
    End If
    If maxRef > capCount Then
        err.Raise vbObjectError + 44, "RegexEngine", _
            "Invalid pattern: backreference \" & maxRef & " refers to a non-existent group (" & capCount & " defined)"
    End If
    If Len(refNames) > 0 Then
        refList = VBA.Split(refNames, "|")
        For ri = 1 To UBound(refList)
            If Len(refList(ri)) = 0 Or InStr(groupNames, "|" & refList(ri) & "|") = 0 Then
                err.Raise vbObjectError + 44, "RegexEngine", _
                    "Invalid pattern: backreference \k<" & refList(ri) & "> refers to a non-existent group"
            End If
        Next ri
    End If
End Sub

' True when pat has a {n}, {n,} or {n,m} quantifier at p, closing before pEnd
Private Function IsBraceQuantifier(ByVal pat As String, ByVal p As Long, ByVal pEnd As Long) As Boolean
    Dim closePos As Long: closePos = InStr(p, pat, "}")
    If closePos = 0 Or closePos > pEnd Then Exit Function
    Dim parts() As String: parts = VBA.Split(mid$(pat, p + 1, closePos - p - 1), ",")
    If UBound(parts) < 0 Or UBound(parts) > 1 Then Exit Function
    If Not IsDigits(parts(0)) Then Exit Function
    If UBound(parts) = 1 Then
        If Len(parts(1)) > 0 And Not IsDigits(parts(1)) Then Exit Function
    End If
    IsBraceQuantifier = True
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = Len(s) > 0 And Not (s Like "*[!0-9]*")
End Function

' True when pat compiles (structure, classes, references and lookbehind widths)
Public Function Validate(ByVal pat As String) As Boolean
    On Error GoTo fail_
    Dim ast As Collection: Set ast = CompilePattern(pat)
    Validate = True
    Exit Function
fail_:
    Validate = False
    err.Clear
End Function

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
            outStr = outStr & ch
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
    Dim groupNodes As New Collection
    CollectNamedGroups ast("body"), groupNodes
    NodeAdd ast, "groupNodes", groupNodes
    Set BuildGroupNodesFromAST = groupNodes
End Function

' Named groups in pre-order along "child" links; each item is Array(name, idx, node)
Private Sub CollectNamedGroups(ByVal node As Collection, ByRef outColl As Collection)
    Dim it As Variant
    Select Case CStr(node(1))
        Case "seq", "alt"
            For Each it In node("child")
                CollectNamedGroups it, outColl
            Next it
        Case "group"
            If Len(node("name")) > 0 Then outColl.Add Array(CStr(node("name")), CLng(node("idx")), node)
            CollectNamedGroups node("child"), outColl
        Case "quant", "atomic", "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
            CollectNamedGroups node("child"), outColl
    End Select
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
    If pNamedCount = 0 Then
        If pNamedCaptures Is Nothing Then
            Set pNamedCaptures = New Collection
        ElseIf pNamedCaptures.Count > 0 Then
            Set pNamedCaptures = New Collection
        End If
        Exit Sub
    End If
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
    For i = 1 To gnodes.Count
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
