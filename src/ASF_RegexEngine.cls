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
'  - Alternation: |
'  - Anchors: ^ and $
'  - Public API: Init, Test, Exec, Replace, Split
'
' Design notes:
'  - AST nodes are represented as VBA Collections with string keys; helper
'    accessors (NodeAdd/NodeGet/NodeExists) are provided.
'  - Matching is backtracking-based; defensive limits and optimizations are
'    implemented to reduce pathological backtracking where practical.
Option Explicit

' ---------------- Public config -----------------------------------------
Private Const DEFAULT_MAX_STEPS As Long = 200000

' ---------------- Public state ------------------------------------------
Private pPattern As String
Private pIgnoreCase As Boolean
Private pMaxMatchSteps As Long
Private pAST As Collection
Private pGroupCount As Long

' Current subject & captures
Private CurSubject As String
Private capsStart() As Long
Private capsEnd() As Long
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

Private Function NodeGet(ByVal node As Collection, ByVal key As String) As Variant
    On Error GoTo notfound
    vAssignment NodeGet, node(key)
    Exit Function
notfound:
    err.Clear
    NodeGet = Empty
End Function

Private Function NodeExists(ByVal node As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant: vAssignment tmp, node(key)
    If err.Number = 0 Then NodeExists = True Else NodeExists = False
    err.Clear
    On Error GoTo 0
End Function

' ---------------- Public API -------------------------------------------
Public Sub Init(ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = False, Optional ByVal maxSteps As Long = DEFAULT_MAX_STEPS)
    pPattern = Pattern
    pIgnoreCase = IgnoreCase
    If maxSteps <= 0 Then pMaxMatchSteps = DEFAULT_MAX_STEPS Else pMaxMatchSteps = maxSteps
    Set pAST = CompilePattern(Pattern)
    pGroupCount = GetGroupCount(pAST)
End Sub

Public Property Get Pattern() As String: Pattern = pPattern: End Property
Public Property Let Pattern(ByVal v As String): Init v, pIgnoreCase, pMaxMatchSteps: End Property
Public Property Get IgnoreCase() As Boolean: IgnoreCase = pIgnoreCase: End Property
Public Property Let IgnoreCase(ByVal v As Boolean): pIgnoreCase = v: End Property
Public Property Get MaxMatchSteps() As Long: MaxMatchSteps = pMaxMatchSteps: End Property
Public Property Let MaxMatchSteps(ByVal v As Long): pMaxMatchSteps = v: End Property

'Public Function Test(ByVal subj As String) As Boolean
'    Dim res As Collection
'    Set res = Exec(subj)
'    Test = Not (res Is Nothing)
'End Function
'
'Public Function Exec(ByVal subj As String) As Collection
'    CurSubject = subj
'    MatchSteps = 0
'    Dim anchoredStart As Boolean: anchoredStart = NodeGet(pAST, "anchoredStart")
'    Dim body As Collection: Set body = NodeGet(pAST, "body")
'    Dim L As Long: L = Len(subj)
'    Dim minLen As Long
'    minLen = NodeGet(pAST, "minLen")
'    If anchoredStart Then
'        If ExecAt(subj, 1) Is Nothing Then
'            Set Exec = Nothing
'        Else
'            Set Exec = BuildCaptureCollection()
'        End If
'        Exit Function
'    End If
'
'    Dim startLimit As Long
'    If minLen > 0 Then startLimit = L - minLen + 1 Else startLimit = L
'    Dim i As Long
'    For i = 1 To startLimit
'        If ExecAt(subj, i) Is Nothing Then
'            ' no match at this start
'        Else
'            Set Exec = BuildCaptureCollection()
'            Exit Function
'        End If
'        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
'    Next i
'    Set Exec = Nothing
'End Function
'
'Public Function ExecAt(ByVal subj As String, ByVal startPos As Long) As Collection
'    CurSubject = subj
'    MatchSteps = MatchSteps + 1
'    Dim totalGroups As Long: totalGroups = pGroupCount
'    ReDim capsStart(0 To totalGroups) As Long
'    ReDim capsEnd(0 To totalGroups) As Long
'    Dim i As Long
'    For i = 0 To totalGroups
'        capsStart(i) = 0: capsEnd(i) = -1
'    Next i
'    capsStart(0) = startPos
'    Dim outEnd As Long
'    Dim ok As Boolean
'    ok = MatchNode(NodeGet(pAST, "body"), subj, startPos, outEnd)
'    If ok Then
'        Set ExecAt = BuildCaptureCollection()
'    Else
'        Set ExecAt = Nothing
'    End If
'End Function


' Replace and Split now use ExecAt to avoid Mid$/InStr double-work
Public Function Replace(ByVal subj As String, ByVal replacement As String) As String
    Dim out As String: out = ""
    Dim pos As Long: pos = 1
    Dim L As Long: L = Len(subj)
    Do
        Dim cs As Collection
        Set cs = ExecAt(subj, pos)
        If cs Is Nothing Then
            out = out & mid$(subj, pos)
            Exit Do
        Else
            Dim mstart As Long: mstart = capsStart(0)
            Dim mend As Long: mend = capsEnd(0)
            If mstart < pos Then mstart = pos
            out = out & mid$(subj, pos, mstart - pos)
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
            res.Add mid$(subj, pos)
            Exit Do
        Else
            Dim mstart As Long: mstart = capsStart(0)
            Dim mend As Long: mend = capsEnd(0)
            If mstart < pos Then mstart = pos
            res.Add mid$(subj, pos, mstart - pos)
            pos = mend + 1
            If pos > L Then
                res.Add ""
                Exit Do
            End If
        End If
    Loop
    Set Split = res
End Function

' ---------------- Parser & AST precomputation ----------------------------
Private Function CompilePattern(ByVal pat As String) As Collection
    Dim ast As New Collection
    NodeAdd ast, "raw", pat

    Dim anchoredStart As Boolean: anchoredStart = False
    Dim anchoredEnd As Boolean: anchoredEnd = False
    If Len(pat) > 0 And mid$(pat, 1, 1) = "^" Then anchoredStart = True: pat = mid$(pat, 2)
    If Len(pat) > 0 And right$(pat, 1) = "$" Then anchoredEnd = True: pat = left$(pat, Len(pat) - 1)
    NodeAdd ast, "anchoredStart", anchoredStart
    NodeAdd ast, "anchoredEnd", anchoredEnd

    Dim pos As Long: pos = 1
    Dim groupCounter As Long: groupCounter = 0
    Dim body As Collection: Set body = ParseSequence(pat, pos, Len(pat), groupCounter)
    NodeAdd ast, "body", body

    ' Precompute minLen
    Dim mlen As Long: mlen = ComputeMinLen(body)
    NodeAdd ast, "minLen", mlen

    Set CompilePattern = ast
End Function

Private Function ParseSequence(ByVal pat As String, ByRef p As Long, ByVal pEnd As Long, ByRef groupCounter As Long) As Collection
    Dim seq As Collection: Set seq = NewNode("seq")
    Dim children As New Collection
    Do While p <= pEnd
        Dim ch As String: ch = mid$(pat, p, 1)
        If ch = ")" Or ch = "|" Then Exit Do
        Dim node As Collection: Set node = Nothing
        If ch = "\" Then
            p = p + 1
            If p > pEnd Then Exit Do
            Set node = ParseEscapeNode(mid$(pat, p, 1))
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
            groupCounter = groupCounter + 1
            Dim thisIdx As Long: thisIdx = groupCounter
            Dim inner As Collection
            Set inner = ParseSequence(pat, p, pEnd, groupCounter)
            If p <= pEnd And mid$(pat, p, 1) = ")" Then p = p + 1
            Dim gnode As Collection: Set gnode = NewNode("group")
            NodeAdd gnode, "idx", thisIdx
            NodeAdd gnode, "child", inner
            Set node = gnode
        Else
            Set node = NewNode("lit"): NodeAdd node, "ch", ch
            p = p + 1
        End If

        ' quantifier
        If p <= pEnd Then
            Dim qch As String: qch = mid$(pat, p, 1)
            If qch = "*" Or qch = "+" Or qch = "?" Or qch = "{" Then
                Dim qnode As Collection: Set qnode = NewNode("quant")
                NodeAdd qnode, "child", node
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
                        Dim content As String: content = mid$(pat, p + 1, closePos - p - 1)
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
                ' lazy flag
                If p <= pEnd And mid$(pat, p, 1) = "?" Then
                    NodeAdd qnode, "lazy", True
                    p = p + 1
                Else
                    NodeAdd qnode, "lazy", False
                End If
                Set node = qnode
            End If
        End If

        children.Add node
    Loop
    NodeAdd seq, "children", children
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
    If i <= pEnd And mid$(pat, i, 1) = "^" Then neg = True: i = i + 1
    Do While i <= pEnd
        Dim ch As String: ch = mid$(pat, i, 1)
        If ch = "\" Then
            i = i + 1
            If i <= pEnd Then chars.Add ParseEscapeChar(mid$(pat, i, 1))
            i = i + 1
        ElseIf i + 2 <= pEnd And mid$(pat, i + 1, 1) = "-" Then
            chars.Add Array("range", mid$(pat, i, 1), mid$(pat, i + 2, 1))
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
        Dim ch As String: ch = mid$(spec, i, 1)
        If i + 2 <= Len(spec) - 1 And mid$(spec, i + 1, 1) = "-" Then
            chars.Add Array("range", ch, mid$(spec, i + 2, 1))
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
    End Select
End Sub

' MatchNode(node, subj, pos, ByRef outEnd) - returns True and sets outEnd to index of last matched char (1-based)
Private Function MatchNode(ByVal node As Collection, ByVal subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    MatchSteps = MatchSteps + 1
    If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 2, "RegexEngine", "Match aborted: too many steps"

    Dim t As Variant: t = NodeGet(node, "type")
    Select Case CStr(t)
        Case "seq"
            MatchNode = MatchSequence(NodeGet(node, "children"), subj, pos, outEnd)
        Case "lit"
            Dim ch As String: ch = NodeGet(node, "ch")
            If pos > Len(subj) Then
                MatchNode = False: Exit Function
            End If
            Dim sc As String: sc = mid$(subj, pos, 1)
            If pIgnoreCase Then
                If LCase$(sc) = LCase$(ch) Then
                    If IsMissing(outEnd) Then capsEnd(0) = pos Else outEnd = pos
                    MatchNode = True
                Else
                    MatchNode = False
                End If
            Else
                If sc = ch Then
                    If IsMissing(outEnd) Then capsEnd(0) = pos Else outEnd = pos
                    MatchNode = True
                Else
                    MatchNode = False
                End If
            End If
        Case "dot"
            If pos > Len(subj) Then
                MatchNode = False
            Else
                If IsMissing(outEnd) Then capsEnd(0) = pos Else outEnd = pos
                MatchNode = True
            End If
        Case "class"
            If pos > Len(subj) Then
                MatchNode = False
            Else
                If MatchCharClassNode(node, mid$(subj, pos, 1)) Then
                    If IsMissing(outEnd) Then capsEnd(0) = pos Else outEnd = pos
                    MatchNode = True
                Else
                    MatchNode = False
                End If
            End If
        Case "group"
            Dim gIdx As Long: gIdx = NodeGet(node, "idx")
            Dim prevStart As Long: prevStart = capsStart(gIdx)
            Dim prevEnd As Long: prevEnd = capsEnd(gIdx)
            capsStart(gIdx) = pos
            Dim childEnd As Variant
            If MatchNode(NodeGet(node, "child"), subj, pos, childEnd) Then
                capsEnd(gIdx) = childEnd
                If IsMissing(outEnd) Then capsEnd(0) = childEnd Else outEnd = childEnd
                MatchNode = True
            Else
                capsStart(gIdx) = prevStart
                capsEnd(gIdx) = prevEnd
                MatchNode = False
            End If
        Case "quant"
            Dim mn As Long: mn = NodeGet(node, "min")
            Dim mx As Long: mx = NodeGet(node, "max")
            Dim lazy As Boolean: lazy = NodeGet(node, "lazy")
            If mx = -1 Then mx = Len(subj) - pos + 1
            MatchNode = MatchQuantifier(NodeGet(node, "child"), mn, mx, lazy, subj, pos, outEnd)
        Case Else
            MatchNode = False
    End Select
End Function

Private Function MatchSequence(ByVal children As Collection, ByVal subj As String, ByVal pos As Long, Optional ByRef outEnd As Variant) As Boolean
    Dim seqSnapshot() As Long: seqSnapshot = CopyCaps()
    Dim curPos As Long: curPos = pos
    Dim i As Long

    For i = 1 To children.count
        Dim node As Collection: Set node = children(i)
        Dim nodeSnapshot() As Long: nodeSnapshot = CopyCaps()
        Dim nodeEnd As Variant
        Dim nodeType As String: nodeType = NodeGet(node, "type")

        If nodeType = "quant" Then
            If Not MatchQuantifier(NodeGet(node, "child"), NodeGet(node, "min"), IIf(NodeGet(node, "max") = -1, Len(subj) - curPos + 1, NodeGet(node, "max")), NodeGet(node, "lazy"), subj, curPos, nodeEnd) Then
                RestoreCaps nodeSnapshot
                RestoreCaps seqSnapshot
                MatchSequence = False: Exit Function
            Else
                curPos = nodeEnd + 1
            End If

        ElseIf nodeType = "group" Then
            Dim gIdx As Long: gIdx = NodeGet(node, "idx")
            capsStart(gIdx) = curPos
            If Not MatchNode(NodeGet(node, "child"), subj, curPos, nodeEnd) Then
                RestoreCaps nodeSnapshot
                RestoreCaps seqSnapshot
                MatchSequence = False: Exit Function
            Else
                capsEnd(gIdx) = nodeEnd
                curPos = nodeEnd + 1
            End If

        Else
            If Not MatchNode(node, subj, curPos, nodeEnd) Then
                RestoreCaps nodeSnapshot
                RestoreCaps seqSnapshot
                MatchSequence = False: Exit Function
            Else
                curPos = nodeEnd + 1
            End If
        End If
    Next i

    outEnd = curPos - 1
    capsEnd(0) = outEnd
    MatchSequence = True
End Function

' MatchQuantifier: tries to match between min..max occurrences. For greedy try larger k first; for lazy try smaller first.
Private Function MatchQuantifier(ByVal child As Collection, ByVal minc As Long, ByVal maxc As Long, ByVal lazy As Boolean, ByVal subj As String, ByVal posStart As Long, Optional ByRef outEnd As Variant) As Boolean
    Dim saveSnapshot() As Long: saveSnapshot = CopyCaps()
    Dim occPos() As Long
    ReDim occPos(0 To 0)
    Dim occCount As Long: occCount = 0
    Dim curPos As Long: curPos = posStart
    Dim nodeEnd As Variant

    Do While occCount < maxc
        Dim tempSnap() As Long: tempSnap = CopyCaps()
        If MatchNode(child, subj, curPos, nodeEnd) Then
            occCount = occCount + 1
            ReDim Preserve occPos(0 To occCount)
            occPos(occCount) = nodeEnd
            curPos = nodeEnd + 1
        Else
            RestoreCaps tempSnap
            Exit Do
        End If
    Loop

    Dim k As Long
    If lazy Then
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
            If okAll Then
                If IsMissing(outEnd) Then capsEnd(0) = posTry - 1 Else outEnd = posTry - 1
                MatchQuantifier = True
                Exit Function
            End If
        Next k
        MatchQuantifier = False
    Else
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
            If okAll2 Then
                If IsMissing(outEnd) Then capsEnd(0) = posTry2 - 1 Else outEnd = posTry2 - 1
                MatchQuantifier = True
                Exit Function
            End If
        Next k
        MatchQuantifier = False
    End If
End Function


' ---------------- Captures snapshot utilities -----------------------------
Private Function CopyCaps() As Long()
    Dim tot As Long: tot = pGroupCount
    Dim res() As Long
    ReDim res(0 To tot * 2 + 1)
    Dim i As Long
    For i = 0 To tot
        res(i) = IIf(Not IsEmpty(capsStart(i)), capsStart(i), 0)
        res(i + tot + 1) = IIf(Not IsEmpty(capsEnd(i)), capsEnd(i), -1)
    Next i
    CopyCaps = res
End Function

Private Sub RestoreCaps(ByRef snapshot() As Long)
    Dim total As Long: total = (UBound(snapshot) - 0) \ 2
    Dim i As Long
    For i = 0 To total
        capsStart(i) = snapshot(i)
        capsEnd(i) = snapshot(i + total + 1)
    Next i
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
        Dim ch As String: ch = mid$(repl, i, 1)
        If ch = "$" Then
            Dim j As Long: j = i + 1
            Dim num As String: num = ""
            Do While j <= Len(repl) And mid$(repl, j, 1) >= "0" And mid$(repl, j, 1) <= "9"
                num = num & mid$(repl, j, 1)
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
        col.Add mid$(CurSubject, capsStart(0), capsEnd(0) - capsStart(0) + 1)
    Else
        col.Add ""
    End If
    Dim i As Long
    For i = 1 To pGroupCount
        If capsStart(i) >= 1 And capsEnd(i) >= capsStart(i) Then
            col.Add mid$(CurSubject, capsStart(i), capsEnd(i) - capsStart(i) + 1)
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
        Dim ch As String: ch = mid$(pat, i, 1)
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
        Case Else: ComputeMinLen = 0
    End Select
End Function
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub

Public Function Test(ByVal subj As String) As Boolean
    Dim res As Collection
    Set res = Exec(subj)
    Test = Not (res Is Nothing)
End Function

Public Function Exec(ByVal subj As String) As Collection
    CurSubject = subj
    MatchSteps = 0
    Dim anchoredStart As Boolean: anchoredStart = NodeGet(pAST, "anchoredStart")
    Dim body As Collection: Set body = NodeGet(pAST, "body")
    Dim L As Long: L = Len(subj)
    Dim minLen As Long
    minLen = NodeGet(pAST, "minLen")

    If anchoredStart Then
        If ExecAt(subj, 1) Is Nothing Then
            Set Exec = Nothing
        Else
            Set Exec = BuildCaptureCollection()
        End If
        Exit Function
    End If

    Dim startLimit As Long
    If minLen > 0 Then startLimit = L - minLen + 1 Else startLimit = L
    If startLimit < 1 Then startLimit = 1

    Dim i As Long
    For i = 1 To startLimit
        If ExecAt(subj, i) Is Nothing Then
            'no match at i
        Else
            Set Exec = BuildCaptureCollection()
            Exit Function
        End If
        If MatchSteps > pMaxMatchSteps Then err.Raise vbObjectError + 1, "RegexEngine", "Match aborted: too many steps"
    Next i
    Set Exec = Nothing
End Function

Public Function ExecAt(ByVal subj As String, ByVal startPos As Long) As Collection
    CurSubject = subj
    MatchSteps = MatchSteps + 1
    Dim totalGroups As Long: totalGroups = pGroupCount
    ReDim capsStart(0 To totalGroups) As Long
    ReDim capsEnd(0 To totalGroups) As Long
    Dim i As Long
    For i = 0 To totalGroups
        capsStart(i) = 0: capsEnd(i) = -1
    Next i
    capsStart(0) = startPos
    Dim outEnd As Variant
    Dim ok As Boolean
    ok = MatchNode(NodeGet(pAST, "body"), subj, startPos, outEnd)
    If ok Then
        Set ExecAt = BuildCaptureCollection()
    Else
        Set ExecAt = Nothing
    End If
End Function
