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
'  - The parser builds AST nodes as VBA Collections with string keys (NodeAdd/NodeGet/
'    NodeExists); LowerProgram compiles them to Long arrays, which is all the matcher reads.
'  - Matching is backtracking-based; defensive limits and optimizations are
'    implemented to reduce pathological backtracking where practical.
'
' Notes & Limitations:
'     + Lookbehind is fixed-width only.
'     + Atomic groups, lookarounds, conditionals and recursion match on their own: the rest of
'       the pattern never backtracks into them.
'     + Possessive quantifiers are supported: only the maximal count is tried.
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
Private pPrefixId As Long
Private pAnchorId As Long
Private pMinLen As Long
' subject as UTF-16 bytes for the start-position scan
Private pSubjBytes() As Byte
Private pCapsSize As Long
Private pSubjLen As Long
' continuation frames of the matcher (see MatchHere) and the capture undo trail (see SetCap)
Private Const FR_SEQ As Long = 1
Private Const FR_CAPEND As Long = 2
Private Const FR_QCONT As Long = 3
Private Const FR_QGUARD As Long = 4
Private kKind() As Long
Private kNode() As Long
Private kNum() As Long
Private kTop As Long
Private kBase As Long
Private pStackCap As Long
Private trSlot() As Long
Private trOld() As Long
Private trTop As Long
Private pTrailCap As Long
' the compiled program (see LowerProgram); b* and pBuild* are its build arrays
Private Const T_LIT As Long = 1
Private Const T_DOT As Long = 2
Private Const T_CLASS As Long = 3
Private Const T_BOL As Long = 4
Private Const T_EOL As Long = 5
Private Const T_WORDB As Long = 6
Private Const T_BACKREF As Long = 7
Private Const T_LA_POS As Long = 8
Private Const T_LA_NEG As Long = 9
Private Const T_LB_POS As Long = 10
Private Const T_LB_NEG As Long = 11
Private Const T_ATOMIC As Long = 12
Private Const T_COND As Long = 13
Private Const T_RECURSE As Long = 14
Private Const T_SEQ As Long = 15
Private Const T_GROUP As Long = 16
Private Const T_ALT As Long = 17
Private Const T_QUANT As Long = 18
Private Const Q_LAZY As Long = 1
Private Const Q_POSS As Long = 2
Private Const Q_DET As Long = 4
Private Const E_I As Long = 1
Private Const E_M As Long = 2
Private Const E_S As Long = 4
Private pNodeCount As Long
Private nType() As Long
Private nA() As Long
Private nB() As Long
Private nKid() As Long
Private nOpt() As Long
Private nList() As Long
Private nCnt() As Long
Private pKids() As Long
Private pLo() As Long
Private pHi() As Long
Private nEff() As Long
Private pEffFor As Long
Private bType() As Long
Private bA() As Long
Private bB() As Long
Private bKid() As Long
Private bOpt() As Long
Private bList() As Long
Private bCnt() As Long
Private bKids() As Long
Private bLo() As Long
Private bHi() As Long
Private bGroup() As Long
Private pBuildCount As Long
Private pBuildKidCount As Long
Private pBuildRangeCount As Long
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
    ' always present, so it is read without a key check
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
    Dim prefix As String, prefixId As Long, anchorId As Long
    Dim prog As Variant: prog = LowerProgram(ast, prefix, prefixId, anchorId)
    CompileState = Array(ast, CLng(ast("groupCount")), pFirstMap, pFirstHigh, pHasFirst, maps, prefix, prefixId, anchorId, named.Count, CLng(ast("minLen")), prog)
End Function

Private Sub LoadState(ByRef st As Variant)
    Set pAST = st(0)
    pGroupCount = st(1)
    pFirstMap = st(2)
    pFirstHigh = st(3)
    pHasFirst = st(4)
    If Not IsEmpty(st(5)) Then pClassMaps = st(5)
    pPrefix = st(6)
    pPrefixId = st(7)
    pAnchorId = st(8)
    pNamedCount = st(9)
    pMinLen = st(10)
    Dim prog As Variant: prog = st(11)
    pNodeCount = prog(0)
    nType = prog(1): nA = prog(2): nB = prog(3): nKid = prog(4)
    nOpt = prog(5): nList = prog(6): nCnt = prog(7)
    pKids = prog(8): pLo = prog(9): pHi = prog(10)
    pEffFor = -1
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

' ------------------ Compiled program ----------------------------------------
' The AST is lowered once per pattern to parallel Long arrays indexed by node id (1 is the body),
' so matching reads array slots instead of keyed Collection items:
'   nType   T_* code
'   nA      lit: char code; class: map index; group: capture index (0: none); quant: min;
'           backref and cond: group (-1: none); lookbehind: width (-1: none); recurse: target number;
'           wordb: 1 for \B
'   nB      lit: lower-case code; class: 1 when negated; quant: max (-1: no limit);
'           cond: "no" branch (0: none); recurse: target node (0: none)
'   nKid    child of group, quant, lookaround and atomic; "yes" branch of cond
'   nList, nCnt  seq and alt: children in pKids; quant: capture indexes in its child, in pKids;
'                class: ranges in pLo/pHi
'   nOpt    quant: Q_* bits; lit, dot, class, backref, bol, eol: inline flags (see FlagBits)
' nEff holds each node's effective flags (E_* bits) for the engine's current defaults.

Private Function LowerProgram(ByVal ast As Collection, ByRef prefix As String, ByRef prefixId As Long, ByRef anchorId As Long) As Variant
    pBuildCount = 0: pBuildKidCount = 0: pBuildRangeCount = 0
    ' fresh arrays, never resized in place: a cached program may hold the previous ones
    FreshLongs bType: FreshLongs bA: FreshLongs bB: FreshLongs bKid: FreshLongs bOpt
    FreshLongs bList: FreshLongs bCnt: FreshLongs bKids: FreshLongs bLo: FreshLongs bHi
    ReDim bGroup(0 To CLng(ast("groupCount")))
    LowerNode ast("body"), ast
    ' recursion targets and quantifier details need the whole program
    Dim id As Long, k As Long
    For id = 1 To pBuildCount
        Select Case bType(id)
            Case T_RECURSE
                If bA(id) = 0 Then
                    bB(id) = 1
                ElseIf FindGroupIndexByOrdinal(ast, bA(id)) > 0 Then
                    bB(id) = bGroup(bA(id))
                End If
            Case T_QUANT
                If IsDeterministic(bKid(id)) Then bOpt(id) = bOpt(id) Or Q_DET
                k = pBuildKidCount
                bList(id) = k + 1
                AddGroupIdx bKid(id)
                bCnt(id) = pBuildKidCount - k
        End Select
    Next id
    AnalyzeStart prefix, prefixId, anchorId
    LowerProgram = Array(pBuildCount, bType, bA, bB, bKid, bOpt, bList, bCnt, bKids, bLo, bHi)
End Function

' True when node nd can match in at most one way at a given position
Private Function IsDeterministic(ByVal nd As Long) As Boolean
    Dim k As Long
    IsDeterministic = True
    Select Case bType(nd)
        Case T_SEQ
            For k = bList(nd) To bList(nd) + bCnt(nd) - 1
                If Not IsDeterministic(bKids(k)) Then IsDeterministic = False: Exit Function
            Next k
        Case T_GROUP
            IsDeterministic = IsDeterministic(bKid(nd))
        Case T_QUANT
            If (bOpt(nd) And Q_POSS) <> 0 Or bA(nd) = bB(nd) Then IsDeterministic = IsDeterministic(bKid(nd)) Else IsDeterministic = False
        Case T_ALT
            If bCnt(nd) = 1 Then IsDeterministic = IsDeterministic(bKids(bList(nd))) Else IsDeterministic = False
    End Select
End Function

' Append the capture indexes of the groups in nd's subtree to bKids
Private Sub AddGroupIdx(ByVal nd As Long)
    Dim k As Long
    Select Case bType(nd)
        Case T_SEQ, T_ALT
            For k = bList(nd) To bList(nd) + bCnt(nd) - 1
                AddGroupIdx bKids(k)
            Next k
        Case T_GROUP
            If bA(nd) > 0 Then AddKid bA(nd)
            AddGroupIdx bKid(nd)
        Case T_QUANT, T_LA_POS, T_LA_NEG, T_LB_POS, T_LB_NEG, T_ATOMIC
            AddGroupIdx bKid(nd)
        Case T_COND
            AddGroupIdx bKid(nd)
            If bB(nd) > 0 Then AddGroupIdx bB(nd)
    End Select
End Sub

' Leading literal and leading ^ of a single-alternative pattern; zero-width assertions before the literal are skipped
Private Sub AnalyzeStart(ByRef prefix As String, ByRef prefixId As Long, ByRef anchorId As Long)
    Dim sq As Long: sq = 1
    If bType(sq) = T_ALT Then
        If bCnt(sq) <> 1 Then Exit Sub
        sq = bKids(bList(sq))
    End If
    If bType(sq) <> T_SEQ Then Exit Sub
    Dim k As Long, nd As Long
    For k = 1 To bCnt(sq)
        nd = bKids(bList(sq) + k - 1)
        Select Case bType(nd)
            Case T_BOL, T_WORDB, T_LA_POS, T_LA_NEG, T_LB_POS, T_LB_NEG
                If Len(prefix) > 0 Then Exit For
                If k = 1 And bType(nd) = T_BOL Then anchorId = nd
            Case T_LIT
                If prefixId = 0 Then prefixId = nd
                If bOpt(nd) <> bOpt(prefixId) Then Exit For
                prefix = prefix & ChrW$(bA(nd))
            Case Else
                Exit For
        End Select
    Next k
End Sub

Private Sub FreshLongs(ByRef a() As Long)
    Dim r() As Long: ReDim r(0 To 15)
    a = r
End Sub

Private Function LowerNode(ByVal node As Collection, ByVal ast As Collection) As Long
    pBuildCount = pBuildCount + 1
    Dim id As Long: id = pBuildCount
    If id > UBound(bType) Then
        Dim cap As Long: cap = 2 * id
        ReDim Preserve bType(0 To cap): ReDim Preserve bA(0 To cap): ReDim Preserve bB(0 To cap)
        ReDim Preserve bKid(0 To cap): ReDim Preserve bOpt(0 To cap): ReDim Preserve bList(0 To cap)
        ReDim Preserve bCnt(0 To cap)
    End If
    Dim it As Variant, n As Long, k As Long, ch As String
    Select Case CStr(node(1))
        Case "lit"
            bType(id) = T_LIT
            ch = node("ch")
            bA(id) = AscW(ch) And &HFFFF&
            bB(id) = AscW(LCase$(ch)) And &HFFFF&
        Case "dot"
            bType(id) = T_DOT
        Case "class"
            bType(id) = T_CLASS
            bA(id) = node("mapIdx")
            If node("neg") Then bB(id) = 1
            bList(id) = pBuildRangeCount + 1
            For Each it In node("chars")
                If IsArray(it) Then AddRange CodeOf(it(1)), CodeOf(it(2)) Else AddRange CodeOf(it), CodeOf(it)
                n = n + 1
            Next it
            bCnt(id) = n
        Case "seq", "alt"
            If CStr(node(1)) = "seq" Then bType(id) = T_SEQ Else bType(id) = T_ALT
            Dim kids As Collection: Set kids = node("child")
            ' the child slots are taken before the children add their own
            bList(id) = pBuildKidCount + 1
            bCnt(id) = kids.Count
            For Each it In kids
                AddKid 0
            Next it
            For Each it In kids
                k = LowerNode(it, ast)
                bKids(bList(id) + n) = k
                n = n + 1
            Next it
        Case "group"
            bType(id) = T_GROUP
            bA(id) = node("idx")
            If bA(id) > 0 Then bGroup(bA(id)) = id
            k = LowerNode(node("child"), ast)
            bKid(id) = k
        Case "quant"
            bType(id) = T_QUANT
            bA(id) = node("min")
            bB(id) = node("max")
            If node("lazy") Then bOpt(id) = Q_LAZY
            If node("possessive") Then bOpt(id) = bOpt(id) Or Q_POSS
            k = LowerNode(node("child"), ast)
            bKid(id) = k
        Case "bol"
            bType(id) = T_BOL
        Case "eol"
            bType(id) = T_EOL
        Case "wordb"
            bType(id) = T_WORDB
            If node("neg") Then bA(id) = 1
        Case "backref"
            bType(id) = T_BACKREF
            If NodeExists(node, "name") Then bA(id) = FindGroupIndexByName(ast, CStr(node("name"))) Else bA(id) = node("idx")
        Case "lookahead_pos", "lookahead_neg", "lookbehind_pos", "lookbehind_neg"
            Select Case CStr(node(1))
                Case "lookahead_pos": bType(id) = T_LA_POS
                Case "lookahead_neg": bType(id) = T_LA_NEG
                Case "lookbehind_pos": bType(id) = T_LB_POS
                Case Else: bType(id) = T_LB_NEG
            End Select
            If NodeExists(node, "width") Then bA(id) = node("width") Else bA(id) = -1
            k = LowerNode(node("child"), ast)
            bKid(id) = k
        Case "atomic"
            bType(id) = T_ATOMIC
            k = LowerNode(node("child"), ast)
            bKid(id) = k
        Case "cond"
            bType(id) = T_COND
            Dim condText As String: condText = node("cond")
            bA(id) = -1
            On Error Resume Next
            If IsNumeric(condText) Then bA(id) = CLng(condText) Else bA(id) = FindGroupIndexByName(ast, condText)
            On Error GoTo 0
            k = LowerNode(node("yes"), ast)
            bKid(id) = k
            If NodeExists(node, "no") Then
                k = LowerNode(node("no"), ast)
                bB(id) = k
            End If
        Case "recurse"
            bType(id) = T_RECURSE
            bA(id) = node("target")
    End Select
    ch = node("flags_on")
    If Len(ch) > 0 Then bOpt(id) = bOpt(id) Or FlagBits(ch)
    LowerNode = id
End Function

Private Sub AddKid(ByVal v As Long)
    pBuildKidCount = pBuildKidCount + 1
    If pBuildKidCount > UBound(bKids) Then
        ReDim Preserve bKids(0 To 2 * pBuildKidCount)
    End If
    bKids(pBuildKidCount) = v
End Sub

Private Sub AddRange(ByVal lo As Long, ByVal hi As Long)
    pBuildRangeCount = pBuildRangeCount + 1
    If pBuildRangeCount > UBound(bLo) Then
        ReDim Preserve bLo(0 To 2 * pBuildRangeCount)
        ReDim Preserve bHi(0 To 2 * pBuildRangeCount)
    End If
    bLo(pBuildRangeCount) = lo
    bHi(pBuildRangeCount) = hi
End Sub

' Inline flags of a node as bits: 1, 2 and 4 switch i, m and s on; 8, 16 and 32 switch them off
Private Function FlagBits(ByVal fl As String) As Long
    Dim k As Long, bit As Long: bit = 1
    For k = 1 To 3
        Select Case ExplicitFlag(fl, mid$("ims", k, 1))
            Case 1: FlagBits = FlagBits Or bit
            Case -1: FlagBits = FlagBits Or 8 * bit
        End Select
        bit = bit * 2
    Next k
End Function

' Effective flags of every node for the current IgnoreCase, Multiline and DotAll
Private Sub ResolveFlags()
    Dim dflt As Long
    If pIgnoreCase Then dflt = E_I
    If pMultiline Then dflt = dflt Or E_M
    If pDotAll Then dflt = dflt Or E_S
    If dflt = pEffFor Then Exit Sub
    ReDim nEff(0 To pNodeCount)
    Dim i As Long
    For i = 1 To pNodeCount
        nEff(i) = (dflt And Not (nOpt(i) \ 8)) Or (nOpt(i) And 7)
    Next i
    pEffFor = dflt
End Sub

' ------------------ Matching engine -----------------------------------------
' Backtracking in continuation-passing style: MatchHere matches a node followed by the
' continuation, the frames on the k* stack above kBase (rest of a sequence, end of a capture,
' next iteration of a quantifier, empty-iteration guard), so no sequence is ever copied.
' Capture writes go through SetCap, which logs the old value on the tr* trail; a failed
' attempt undoes the trail back to its mark. Every matcher leaves the trail as it found it
' when it returns False.

' Match node nd at pos, then the continuation; outEnd is the last position of the whole match
Private Function MatchHere(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim t As Long: t = nType(nd)
    Dim nxt As Long
    Select Case t
        Case T_SEQ
            MatchHere = MatchSeq(nd, 1, pos, outEnd)
        Case T_GROUP
            MatchHere = MatchGroup(nd, pos, outEnd)
        Case T_ALT
            MatchHere = MatchAlt(nd, pos, outEnd)
        Case T_QUANT
            MatchHere = MatchQuant(nd, 0, pos, outEnd)
        Case Else
            MatchSteps = MatchSteps + 1
            If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"
            Select Case t
                Case T_LIT, T_DOT, T_CLASS
                    If LeafMatches(nd, pos) Then MatchHere = MatchRest(pos + 1, outEnd)
                Case T_LA_POS, T_LA_NEG, T_LB_POS, T_LB_NEG
                    MatchHere = MatchLook(nd, t, pos, outEnd)
                Case T_ATOMIC, T_COND, T_RECURSE
                    MatchHere = MatchIsolatedThenRest(nd, t, pos, outEnd)
                Case Else
                    If MatchSimple(nd, t, pos, nxt) Then MatchHere = MatchRest(nxt, outEnd)
            End Select
    End Select
End Function

' Children i.. of sequence sq, then the continuation: leaves and assertions are matched here in a
' loop; any other node is matched with the rest of the sequence pushed as a frame
Private Function MatchSeq(ByVal sq As Long, ByVal i As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim n As Long: n = nCnt(sq)
    Dim first As Long: first = nList(sq) - 1
    Dim nd As Long, t As Long, nxt As Long
    Do While i <= n
        nd = pKids(first + i)
        t = nType(nd)
        If t <= T_CLASS Then
            MatchSteps = MatchSteps + 1
            If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"
            If Not LeafMatches(nd, pos) Then Exit Function
            pos = pos + 1
        ElseIf t <= T_BACKREF Then
            MatchSteps = MatchSteps + 1
            If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"
            If Not MatchSimple(nd, t, pos, nxt) Then Exit Function
            pos = nxt
        Else
            If i < n Then
                PushFrame FR_SEQ, sq, i + 1
                MatchSeq = MatchHere(nd, pos, outEnd)
                kTop = kTop - 1
            Else
                MatchSeq = MatchHere(nd, pos, outEnd)
            End If
            Exit Function
        End If
        i = i + 1
    Loop
    MatchSeq = MatchRest(pos, outEnd)
End Function

' The continuation from pos: pop the top frame, run it, put it back
Private Function MatchRest(ByVal pos As Long, ByRef outEnd As Long) As Boolean
    If kTop = kBase Then
        outEnd = pos - 1
        MatchRest = True
        Exit Function
    End If
    Dim kind As Long: kind = kKind(kTop)
    Dim nd As Long: nd = kNode(kTop)
    Dim num As Long: num = kNum(kTop)
    Dim mark As Long
    kTop = kTop - 1
    Select Case kind
        Case FR_SEQ
            MatchRest = MatchSeq(nd, num, pos, outEnd)
        Case FR_CAPEND
            mark = trTop
            SetCap -num, pos - 1
            MatchRest = MatchRest(pos, outEnd)
            If Not MatchRest Then Undo mark
        Case FR_QCONT
            MatchRest = MatchQuant(nd, num, pos, outEnd)
        Case FR_QGUARD
            ' an iteration beyond the minimum must consume input (ECMAScript RepeatMatcher)
            If pos <> num Then MatchRest = MatchRest(pos, outEnd)
    End Select
    kTop = kTop + 1
    kKind(kTop) = kind
    kNode(kTop) = nd
    kNum(kTop) = num
End Function

' A node on its own, with no continuation: lookarounds, atomic groups, conditionals, recursion
Private Function MatchIsolated(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim saveBase As Long: saveBase = kBase
    kBase = kTop
    MatchIsolated = MatchHere(nd, pos, outEnd)
    kBase = saveBase
End Function

Private Sub PushFrame(ByVal kind As Long, ByVal nd As Long, ByVal num As Long)
    kTop = kTop + 1
    If kTop > pStackCap Then
        pStackCap = 2 * pStackCap + 32
        ReDim Preserve kKind(1 To pStackCap)
        ReDim Preserve kNode(1 To pStackCap)
        ReDim Preserve kNum(1 To pStackCap)
    End If
    kKind(kTop) = kind
    kNode(kTop) = nd
    kNum(kTop) = num
End Sub

' Set a capture slot, logging the old value: g >= 1 is the start of group g, g <= -1 the end of group -g
Private Sub SetCap(ByVal g As Long, ByVal v As Long)
    trTop = trTop + 1
    If trTop > pTrailCap Then
        pTrailCap = 2 * pTrailCap + 64
        ReDim Preserve trSlot(1 To pTrailCap)
        ReDim Preserve trOld(1 To pTrailCap)
    End If
    trSlot(trTop) = g
    If g > 0 Then
        trOld(trTop) = capsStart(g): capsStart(g) = v
    Else
        trOld(trTop) = capsEnd(-g): capsEnd(-g) = v
    End If
End Sub

' Undo capture writes back to a trail mark
Private Sub Undo(ByVal mark As Long)
    Dim g As Long
    Do While trTop > mark
        g = trSlot(trTop)
        If g > 0 Then capsStart(g) = trOld(trTop) Else capsEnd(-g) = trOld(trTop)
        trTop = trTop - 1
    Loop
End Sub

Private Function MatchGroup(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim gidx As Long: gidx = nA(nd)
    If gidx < 1 Then
        MatchGroup = MatchHere(nKid(nd), pos, outEnd)
        Exit Function
    End If
    Dim mark As Long: mark = trTop
    SetCap gidx, pos
    PushFrame FR_CAPEND, 0, gidx
    MatchGroup = MatchHere(nKid(nd), pos, outEnd)
    kTop = kTop - 1
    If Not MatchGroup Then Undo mark
End Function

Private Function MatchAlt(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim k As Long
    For k = nList(nd) To nList(nd) + nCnt(nd) - 1
        If MatchHere(pKids(k), pos, outEnd) Then MatchAlt = True: Exit Function
    Next k
End Function

' Lookarounds: the inner pattern is matched on its own; a positive one keeps its captures (ECMAScript)
Private Function MatchLook(ByVal nd As Long, ByVal t As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim at As Long: at = pos
    Dim negative As Boolean: negative = (t = T_LA_NEG Or t = T_LB_NEG)
    If t = T_LB_POS Or t = T_LB_NEG Then
        If nA(nd) < 0 Then Exit Function
        at = pos - nA(nd)
        If at < 1 Then
            ' no room before pos: the inner pattern cannot match
            If negative Then MatchLook = MatchRest(pos, outEnd)
            Exit Function
        End If
    End If
    Dim mark As Long: mark = trTop
    Dim e As Long
    Dim ok As Boolean: ok = MatchIsolated(nKid(nd), at, e)
    If negative Then
        Undo mark
        If ok Then Exit Function
    ElseIf Not ok Then
        Exit Function
    End If
    MatchLook = MatchRest(pos, outEnd)
    If Not MatchLook Then Undo mark
End Function

' atomic, cond and recurse: the node matches on its own, with no backtracking into it, then the continuation
Private Function MatchIsolatedThenRest(ByVal nd As Long, ByVal t As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim mark As Long: mark = trTop
    Dim e As Long, ok As Boolean
    Select Case t
        Case T_ATOMIC: ok = MatchIsolated(nKid(nd), pos, e)
        Case T_COND: ok = MatchCond(nd, pos, e)
        Case T_RECURSE: ok = MatchRecurse(nd, pos, e)
    End Select
    If Not ok Then
        Undo mark
        Exit Function
    End If
    MatchIsolatedThenRest = MatchRest(e + 1, outEnd)
    If Not MatchIsolatedThenRest Then Undo mark
End Function

' (?(cond)yes|no): cond is a group number or name; a missing branch fails
Private Function MatchCond(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim cIdx As Long: cIdx = nA(nd)
    Dim branch As Long
    If cIdx >= 0 And cIdx <= UBound(capsEnd) Then
        If capsEnd(cIdx) <> -1 Then branch = nKid(nd) Else branch = nB(nd)
    Else
        branch = nB(nd)
    End If
    If branch > 0 Then MatchCond = MatchIsolated(branch, pos, outEnd)
End Function

' (?R) and (?n): the whole pattern or group n, guarded against left recursion and depth
Private Function MatchRecurse(ByVal nd As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim targ As Long: targ = nA(nd)
    ' re-entering the same target at the same position can never consume input (left recursion)
    Dim recKey As String: recKey = "|" & targ & "@" & pos & "|"
    If InStr(pRecActive, recKey) > 0 Then Exit Function
    pRecDepth = pRecDepth + 1
    If pRecDepth > RECURSE_LIMIT Then err.Raise vbObjectError + 3, "RegexEngine", "Match aborted: recursion too deep"
    Dim prevActive As String: prevActive = pRecActive
    pRecActive = pRecActive & recKey
    If nB(nd) > 0 Then MatchRecurse = MatchIsolated(nB(nd), pos, outEnd)
    pRecActive = prevActive
    pRecDepth = pRecDepth - 1
End Function

' Zero-width assertions and backreferences: True when node nd matches at pos; nxt is where matching resumes
Private Function MatchSimple(ByVal nd As Long, ByVal t As Long, ByVal pos As Long, ByRef nxt As Long) As Boolean
    nxt = pos
    Select Case t
        Case T_BOL
            If pos = 1 Then
                MatchSimple = True
            ElseIf nEff(nd) And E_M Then
                MatchSimple = IsLineBreakAt(pos - 1)
            End If
        Case T_EOL
            If pos = pSubjLen + 1 Then
                MatchSimple = True
            ElseIf nEff(nd) And E_M Then
                MatchSimple = IsLineBreakAt(pos)
            End If
        Case T_WORDB
            MatchSimple = ((IsWordCharAt(pos - 1) = IsWordCharAt(pos)) = (nA(nd) <> 0))
        Case T_BACKREF
            MatchSimple = MatchBackref(nd, pos, nxt)
    End Select
End Function

' A group that did not participate matches the empty string (ECMAScript/VBScript)
Private Function MatchBackref(ByVal nd As Long, ByVal pos As Long, ByRef nxt As Long) As Boolean
    Dim brIdx As Long: brIdx = nA(nd)
    Dim brLen As Long
    If brIdx >= 1 And brIdx <= pGroupCount Then
        If capsStart(brIdx) >= 1 And capsEnd(brIdx) >= capsStart(brIdx) Then brLen = capsEnd(brIdx) - capsStart(brIdx) + 1
    End If
    nxt = pos + brLen
    If brLen = 0 Then MatchBackref = True: Exit Function
    If nxt - 1 > pSubjLen Then Exit Function
    Dim brText As String: brText = mid$(CurSubject, pos, brLen)
    Dim brCap As String: brCap = mid$(CurSubject, capsStart(brIdx), brLen)
    If nEff(nd) And E_I Then brText = LCase$(brText): brCap = LCase$(brCap)
    MatchBackref = (brText = brCap)
End Function

Private Function IsLineBreakAt(ByVal pos As Long) As Boolean
    If pos < 1 Or pos > pSubjLen Then Exit Function
    Dim code As Long: code = pSubjBytes(2 * pos - 2) + 256& * pSubjBytes(2 * pos - 1)
    IsLineBreakAt = (code = 10 Or code = 13)
End Function

' Captures inside a quantified atom start unset on every iteration (ECMAScript).
Private Sub ResetGroupCaps(ByVal qd As Long)
    Dim k As Long, g As Long
    For k = nList(qd) To nList(qd) + nCnt(qd) - 1
        g = pKids(k)
        If capsStart(g) <> 0 Then SetCap g, 0
        If capsEnd(g) <> -1 Then SetCap -g, -1
    Next k
End Sub

' Quantifier qd with count iterations done, then the continuation. A child with internal choices is
' matched one iteration at a time, with the next iteration as a frame, so the continuation can
' backtrack into each iteration.
Private Function MatchQuant(ByVal qd As Long, ByVal count As Long, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim child As Long: child = nKid(qd)
    Dim mn As Long: mn = nA(qd)
    Dim mx As Long: mx = nB(qd)
    Dim opt As Long: opt = nOpt(qd)
    Dim lazy As Boolean: lazy = (opt And Q_LAZY) <> 0
    Dim possFlag As Boolean: possFlag = (opt And Q_POSS) <> 0
    Dim remMin As Long: remMin = mn - count
    If remMin < 0 Then remMin = 0
    Dim remMax As Long
    If mx = -1 Then remMax = -1 Else remMax = mx - count

    If possFlag Or pQuantDepth >= QUANT_REC_LIMIT Or (opt And Q_DET) <> 0 Then
        MatchQuant = MatchQuantIter(child, remMin, remMax, lazy, possFlag, pos, outEnd)
        Exit Function
    End If

    If lazy And remMin = 0 Then
        If MatchRest(pos, outEnd) Then MatchQuant = True: Exit Function
    End If
    If remMax <> 0 Then
        Dim mark As Long: mark = trTop
        ResetGroupCaps qd
        PushFrame FR_QCONT, qd, count + 1
        If remMin = 0 Then PushFrame FR_QGUARD, 0, pos
        pQuantDepth = pQuantDepth + 1
        Dim ok As Boolean: ok = MatchHere(child, pos, outEnd)
        pQuantDepth = pQuantDepth - 1
        If remMin = 0 Then kTop = kTop - 2 Else kTop = kTop - 1
        If ok Then MatchQuant = True: Exit Function
        Undo mark
    End If
    If Not lazy And remMin = 0 Then MatchQuant = MatchRest(pos, outEnd)
End Function

' Quantifier over a child with no internal choices: each occurrence is matched on its own, then the
' continuation is tried for each count (greedy: most first; lazy: fewest first; possessive: most only)
Private Function MatchQuantIter(ByVal child As Long, ByVal mn As Long, ByVal mx As Long, ByVal lazy As Boolean, ByVal possFlag As Boolean, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    If nType(child) <= T_CLASS Then
        MatchQuantIter = MatchLeafRun(child, mn, mx, lazy, possFlag, pos, outEnd)
        Exit Function
    End If

    Dim mark0 As Long: mark0 = trTop
    Dim lazyRun As Boolean: lazyRun = lazy And Not possFlag
    Dim occEnds() As Long, occMarks() As Long
    Dim n As Long, occCap As Long
    Dim scanPos As Long: scanPos = pos
    Dim e As Long, preMark As Long
    Dim zeroLen As Boolean
    Do
        If lazyRun And n >= mn Then
            If MatchRest(scanPos, outEnd) Then MatchQuantIter = True: Exit Function
        End If
        If n = mx Then Exit Do
        preMark = trTop
        If Not MatchIsolated(child, scanPos, e) Then Exit Do
        ' an empty iteration repeats identically; it only counts towards the minimum
        zeroLen = (e < scanPos)
        If zeroLen And n >= mn Then
            Undo preMark
            Exit Do
        End If
        If zeroLen Then n = mn Else n = n + 1
        scanPos = e + 1
        If Not lazyRun Then
            If n > occCap Then
                occCap = occCap * 2 + 4
                ReDim Preserve occEnds(0 To occCap)
                ReDim Preserve occMarks(0 To occCap)
            End If
            occEnds(n) = e
            occMarks(n) = trTop
        End If
        If zeroLen Then
            If lazyRun Then
                If MatchRest(scanPos, outEnd) Then MatchQuantIter = True: Exit Function
            End If
            Exit Do
        End If
    Loop
    If Not lazyRun And n >= mn Then
        Dim k As Long, kEnd As Long, posTry As Long
        If possFlag Then kEnd = n Else kEnd = mn
        For k = n To kEnd Step -1
            If k = 0 Then
                Undo mark0
                posTry = pos
            Else
                Undo occMarks(k)
                posTry = occEnds(k) + 1
            End If
            If MatchRest(posTry, outEnd) Then MatchQuantIter = True: Exit Function
        Next k
    End If
    Undo mark0
End Function

' Quantified single-char node: captures cannot change, so occurrence k ends at pos + k - 1
Private Function MatchLeafRun(ByVal child As Long, ByVal mn As Long, ByVal mx As Long, ByVal lazy As Boolean, ByVal possFlag As Boolean, ByVal pos As Long, ByRef outEnd As Long) As Boolean
    Dim n As Long: n = LeafCount(child, pos, mn)
    If n < mn Then Exit Function
    If lazy And Not possFlag Then
        ' extend one char at a time, trying the continuation first
        Do
            If MatchRest(pos + n, outEnd) Then MatchLeafRun = True: Exit Function
            If n = mx Then Exit Do
            If LeafCount(child, pos + n, 1) = 0 Then Exit Do
            n = n + 1
        Loop
    Else
        If mx = -1 Then
            n = n + LeafCount(child, pos + n, -1)
        ElseIf n < mx Then
            n = n + LeafCount(child, pos + n, mx - n)
        End If
        Dim k As Long, kEnd As Long
        If possFlag Then kEnd = n Else kEnd = mn
        For k = n To kEnd Step -1
            If MatchRest(pos + k, outEnd) Then MatchLeafRun = True: Exit Function
        Next k
    End If
End Function

' Chars from pos on that match a lit, dot or class node, at most limit of them (-1: no limit).
' Each char tested counts as a step, the one that fails too.
Private Function LeafCount(ByVal nd As Long, ByVal pos As Long, ByVal limit As Long) As Long
    If limit = 0 Then Exit Function
    Dim last As Long: last = pSubjLen
    If limit > 0 Then
        If pos + limit - 1 < last Then last = pos + limit - 1
    End If
    Dim p As Long: p = pos
    Dim code As Long, want As Long
    Select Case nType(nd)
        Case T_LIT
            If nEff(nd) And E_I Then
                want = nB(nd)
                Do While p <= last
                    If LowerCode(pSubjBytes(2 * p - 2) + 256& * pSubjBytes(2 * p - 1)) <> want Then Exit Do
                    p = p + 1
                Loop
            Else
                want = nA(nd)
                Do While p <= last
                    If pSubjBytes(2 * p - 2) + 256& * pSubjBytes(2 * p - 1) <> want Then Exit Do
                    p = p + 1
                Loop
            End If
        Case T_DOT
            If nEff(nd) And E_S Then
                p = last + 1
            Else
                Do While p <= last
                    code = pSubjBytes(2 * p - 2) + 256& * pSubjBytes(2 * p - 1)
                    If code = 10 Or code = 13 Then Exit Do
                    p = p + 1
                Loop
            End If
        Case T_CLASS
            Dim iCase As Boolean: iCase = (nEff(nd) And E_I) <> 0
            Dim mapIdx As Long: mapIdx = nA(nd)
            Dim neg As Boolean: neg = (nB(nd) <> 0)
            Dim hit As Boolean
            Do While p <= last
                code = pSubjBytes(2 * p - 2) + 256& * pSubjBytes(2 * p - 1)
                If code <= 255 Then
                    hit = pClassMaps(code, mapIdx)
                    If iCase And Not hit Then
                        If code >= 65 And code <= 90 Then
                            hit = pClassMaps(code + 32, mapIdx)
                        ElseIf code >= 97 And code <= 122 Then
                            hit = pClassMaps(code - 32, mapIdx)
                        End If
                    End If
                    If hit = neg Then Exit Do
                ElseIf Not ClassHit(nd, code) Then
                    Exit Do
                End If
                p = p + 1
            Loop
    End Select
    LeafCount = p - pos
    MatchSteps = MatchSteps + LeafCount
    If limit < 0 Or LeafCount < limit Then MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"
End Function

' lit, dot and class nodes: does the char at pos match?
Private Function LeafMatches(ByVal nd As Long, ByVal pos As Long) As Boolean
    If pos > pSubjLen Then Exit Function
    Dim code As Long: code = pSubjBytes(2 * pos - 2) + 256& * pSubjBytes(2 * pos - 1)
    Select Case nType(nd)
        Case T_LIT
            If code = nA(nd) Then
                LeafMatches = True
            ElseIf nEff(nd) And E_I Then
                LeafMatches = (LowerCode(code) = nB(nd))
            End If
        Case T_DOT
            LeafMatches = (code <> 10 And code <> 13) Or (nEff(nd) And E_S) <> 0
        Case T_CLASS
            LeafMatches = ClassHit(nd, code)
    End Select
End Function

' ---------------- Char class matching with precomputed map -----------------
Private Function ClassHit(ByVal nd As Long, ByVal code As Long) As Boolean
    Dim iCase As Boolean: iCase = (nEff(nd) And E_I) <> 0
    Dim hit As Boolean
    If code <= 255 Then
        Dim mapIdx As Long: mapIdx = nA(nd)
        hit = pClassMaps(code, mapIdx)
        ' if ignore-case, check the swapped-case ASCII code too (A<->a)
        If iCase And Not hit Then
            If code >= 65 And code <= 90 Then
                hit = pClassMaps(code + 32, mapIdx)
            ElseIf code >= 97 And code <= 122 Then
                hit = pClassMaps(code - 32, mapIdx)
            End If
        End If
    Else
        ' above 255: test the ranges; ignoring case compares lower-cased code points
        Dim k As Long, lc As Long
        If iCase Then lc = LowerCode(code)
        For k = nList(nd) To nList(nd) + nCnt(nd) - 1
            If iCase Then
                hit = (lc >= LowerCode(pLo(k)) And lc <= LowerCode(pHi(k)))
            Else
                hit = (code >= pLo(k) And code <= pHi(k))
            End If
            If hit Then Exit For
        Next k
    End If
    ClassHit = (hit <> (nB(nd) <> 0))
End Function

' Code point of LCase$ of a char
Private Function LowerCode(ByVal code As Long) As Long
    If code < 128 Then
        If code >= 65 And code <= 90 Then LowerCode = code + 32 Else LowerCode = code
    Else
        LowerCode = CodeOf(LCase$(ChrW$(code)))
    End If
End Function

' Word char test for \b and \B, same ASCII set as \w
Private Function IsWordCharAt(ByVal pos As Long) As Boolean
    If pos < 1 Or pos > pSubjLen Then Exit Function
    Dim code As Long: code = pSubjBytes(2 * pos - 2) + 256& * pSubjBytes(2 * pos - 1)
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
    If pMinLen > 0 Then StartLimit = Len(subj) - pMinLen + 1 Else StartLimit = Len(subj) + 1
    If StartLimit < 1 Then StartLimit = 1
End Function

Private Sub BeginScan(ByRef subj As String)
    CurSubject = subj
    pSubjBytes = subj
    MatchSteps = 0
    ResolveFlags
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
    BeginScan subj
    If Not CanStartAt(subj, startPos) Then Set ExecAt = Nothing: Exit Function
    If TryAt(subj, startPos) Then
        Set ExecAt = BuildCaptureCollection()
        UpdateNamedCapturesFromAST
    Else
        Set ExecAt = Nothing
    End If
End Function

' One match attempt at startPos (subj is CurSubject); on success the match is in capsStart/capsEnd
Private Function TryAt(ByRef subj As String, ByVal startPos As Long) As Boolean
    MatchSteps = MatchSteps + 1
    pQuantDepth = 0
    pRecDepth = 0
    pRecActive = vbNullString
    kTop = 0: kBase = 0: trTop = 0
    pSubjLen = Len(subj)
    Dim totalGroups As Long: totalGroups = pGroupCount
    If pCapsSize <> totalGroups + 1 Then
        ReDim capsStart(0 To totalGroups) As Long
        ReDim capsEnd(0 To totalGroups) As Long
        pCapsSize = totalGroups + 1
    End If
    Dim i As Long
    For i = 0 To totalGroups
        capsStart(i) = 0: capsEnd(i) = -1
    Next i
    capsStart(0) = startPos
    Dim outEnd As Long
    If MatchHere(1, startPos, outEnd) Then
        capsEnd(0) = outEnd
        TryAt = True
    End If
End Function

' Can a match start at p? Checks the leading ^, the leading literal and the first-char map
Private Function CanStartAt(ByRef subj As String, ByVal p As Long) As Boolean
    If pAnchorId > 0 And p > 1 Then
        If (nEff(pAnchorId) And E_M) = 0 Then Exit Function
        Dim prev As String: prev = mid$(subj, p - 1, 1)
        If prev <> vbLf And prev <> vbCr Then Exit Function
    End If
    If Len(pPrefix) > 0 Then
        If (nEff(pPrefixId) And E_I) = 0 Then
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
    If pAnchorId > 0 Then
        If nEff(pAnchorId) And E_M Then
            lineStart = True
        ElseIf p > 1 Then
            Exit Function
        ElseIf lastPos > 1 Then
            lastPos = 1
        End If
    End If
    If Len(pPrefix) > 0 Then usePrefix = (nEff(pPrefixId) And E_I) = 0
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
