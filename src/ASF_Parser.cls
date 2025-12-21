VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Parser
Option Explicit
Private Const d_Apostrophe As String = "'"

Private GLOBALS_ As ASF_Globals

Public Sub SetGlobals(ByRef aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Sub
Public Function Tokenize(src As String) As Collection
    Dim toks As New Collection
    Dim i As Long, n As Long
    Dim ch As String
    Dim nxt As String
    Dim j As Long, jC As Long, jB As Long
    n = Len(src)
    i = 1

Tokenize_MainLoop:
    Do While i <= n
        If i > n Then Exit Do ' FIX: Bounds
        ch = mid$(src, i, 1)
        ' comments and slash handling
        If ch = "/" And i < n Then
            nxt = mid$(src, i + 1, 1)
            If nxt = "/" Then
                jC = i + 2
                Do While jC <= n And mid$(src, jC, 1) <> vbCr And mid$(src, jC, 1) <> vbLf
                    jC = jC + 1
                Loop
                toks.Add Array("COMMENT", mid$(src, i + 2, jC - i - 2))
                i = jC
                GoTo Tokenize_SkipToLoopEnd
            ElseIf nxt = "*" Then
                jB = i + 2
                Dim foundClose As Boolean: foundClose = False
                Do While jB <= n - 1
                    If mid$(src, jB, 2) = "*/" Then
                        foundClose = True
                        Exit Do
                    End If
                    jB = jB + 1
                Loop
                If foundClose Then
                    toks.Add Array("COMMENT", mid$(src, i + 2, jB - i - 2))
                    i = jB + 2
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    toks.Add Array("COMMENT", mid$(src, i + 2))
                    i = n + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
            End If
        End If

        If ch = "#" Then
            Dim jH As Long: jH = i + 1
            Do While jH <= n And mid$(src, jH, 1) <> vbCr And mid$(src, jH, 1) <> vbLf
                jH = jH + 1
            Loop
            toks.Add Array("COMMENT", mid$(src, i + 1, jH - i - 1))
            i = jH
            GoTo Tokenize_SkipToLoopEnd
        End If

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                i = i + 1
                GoTo Tokenize_SkipToLoopEnd

            ' TEMPLATE literal handling: `...${ expr }...`
            Case "`"
                ' consume opening backtick
                i = i + 1
                Dim tplParts As New Collection
                Dim curLit As String: curLit = ""
                Dim Ls As Long: Ls = n
            
                Do While i <= Ls
                    Dim c As String: c = mid$(src, i, 1)
            
                    ' escape sequences - keep them as raw characters in the literal
                    If c = "\" Then
                        If i < Ls Then
                            curLit = curLit & mid$(src, i + 1, 1)
                            i = i + 2
                        Else
                            curLit = curLit & "\"
                            i = i + 1
                        End If
                        GoTo ContinueTemplateLoop
                    End If
            
                    ' placeholder start ${ ... }
                    If c = "$" And i < Ls And mid$(src, i + 1, 1) = "{" Then
                        ' flush current literal
                        If Len(curLit) > 0 Then
                            tplParts.Add Array("LITERAL", curLit)
                            curLit = ""
                        End If
                        ' consume "${"
                        i = i + 2
                        Dim depth As Long: depth = 1
                        Dim exprStart As Long: exprStart = i
                        Do While i <= Ls And depth > 0
                            Dim cc As String: cc = mid$(src, i, 1)
                            If cc = "\" Then
                                ' preserve escaped char in expression substring (skip next char)
                                If i < Ls Then i = i + 2 Else i = i + 1
                            Else
                                If cc = "{" Then
                                    depth = depth + 1
                                ElseIf cc = "}" Then
                                    depth = depth - 1
                                End If
                                i = i + 1
                            End If
                        Loop
                        ' exprStart .. i-2 is the inner expression (we stopped just after the closing '}')
                        Dim exprLen As Long: exprLen = i - exprStart - 1
                        Dim exprSrc As String
                        If exprLen >= 0 Then
                            exprSrc = mid$(src, exprStart, exprLen)
                        Else
                            exprSrc = ""
                        End If
                        tplParts.Add Array("EXPR", exprSrc)
                        GoTo ContinueTemplateLoop
                    End If
            
                    ' closing backtick ends template
                    If c = "`" Then
                        If Len(curLit) > 0 Then
                            tplParts.Add Array("LITERAL", curLit)
                            curLit = ""
                        End If
                        i = i + 1 ' consume closing backtick
            
                        ' Reconstruct a plain string token value containing backticks and ${...}
                        Dim reconstructed As String: reconstructed = ""
                        Dim pi As Long
                        For pi = 1 To tplParts.count
                            Dim p As Variant: p = tplParts(pi)
                            If p(0) = "LITERAL" Then
                                ' We keep the literal text as-is. No special escaping required here.
                                reconstructed = reconstructed & p(1)
                            ElseIf p(0) = "EXPR" Then
                                reconstructed = reconstructed & "${" & p(1) & "}"
                            End If
                        Next pi
                        ' Wrap in backticks so compiler can detect it easily
                        toks.Add Array("STRING", "`" & reconstructed & "`")
                        GoTo Tokenize_SkipToLoopEnd
                    End If
            
                    ' normal char, accumulate literal
                    curLit = curLit & c
                    i = i + 1
            
ContinueTemplateLoop:
                Loop
            
                ' Unterminated template: treat remainder as literal and emit as a STRING token
                If Len(curLit) > 0 Then tplParts.Add Array("LITERAL", curLit)
                Dim reconstructed2 As String: reconstructed2 = ""
                Dim pj As Long
                For pj = 1 To tplParts.count
                    Dim pp As Variant: pp = tplParts(pj)
                    If pp(0) = "LITERAL" Then
                        reconstructed2 = reconstructed2 & pp(1)
                    Else
                        reconstructed2 = reconstructed2 & "${" & pp(1) & "}"
                    End If
                Next pj
                toks.Add Array("STRING", "`" & reconstructed2 & "`")
                GoTo Tokenize_SkipToLoopEnd

            Case ","
                toks.Add Array("SEP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case ";"
                toks.Add Array("SEP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "(", ")", "{", "}", "[", "]"
                toks.Add Array("PAREN", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "+", "-", "*", "/", "^", "%", "&", "|"
                ' support +=, -=, *=, /=, ^=, %=, &=, |=
                If i < n And mid$(src, i + 1, 1) = "=" Then
                    toks.Add Array("OP", ch & "="): i = i + 2
                Else
                    toks.Add Array("OP", ch): i = i + 1
                End If
                GoTo Tokenize_SkipToLoopEnd
            Case "&"
                If i < n And mid$(src, i + 1, 1) = "&" Then toks.Add Array("OP", "&&"): i = i + 2 Else toks.Add Array("OP", "&"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "|"
                If i < n And mid$(src, i + 1, 1) = "|" Then toks.Add Array("OP", "||"): i = i + 2 Else toks.Add Array("SYM", "|"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "!"
                If i < n And mid$(src, i + 1, 1) = "=" Then toks.Add Array("OP", "!="): i = i + 2 Else toks.Add Array("OP", "!"): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "<", ">"
                Dim two As String
                Dim three As String
                If i < n Then two = mid$(src, i, 2) Else two = ""
                If two = ">=" Or two = "<=" Then
                    toks.Add Array("OP", two): i = i + 2
                Else
                    If i < n Then three = mid$(src, i, 3) Else three = ""
                    If three = "<<=" Or three = ">>=" Then 'Support for compound shift assignment
                        toks.Add Array("OP", three): i = i + 3
                    Else
                        If two = "<<" Or two = ">>" Then 'Support for eleft and right shifts
                            toks.Add Array("OP", two): i = i + 2
                        Else
                            toks.Add Array("OP", ch): i = i + 1
                        End If
                    End If
                End If
                GoTo Tokenize_SkipToLoopEnd
            Case "="
                If i < n And mid$(src, i + 1, 1) = "=" Then toks.Add Array("OP", "=="): i = i + 2 Else toks.Add Array("OP", "="): i = i + 1
                GoTo Tokenize_SkipToLoopEnd
            Case "?"
                toks.Add Array("OP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case ":"
                toks.Add Array("OP", ch): i = i + 1: GoTo Tokenize_SkipToLoopEnd
            Case "'"
                j = i + 1
                Dim sb As String: sb = ""
                Do While j <= n
                    Dim cj As String: cj = mid$(src, j, 1)
                    If cj = "\" Then
                        If j + 1 <= n Then
                            Dim nx As String: nx = mid$(src, j + 1, 1)
                            Select Case nx
                                Case "'": sb = sb & "'": j = j + 2: GoTo ContinueDo
                                Case "\": sb = sb & "\": j = j + 2: GoTo ContinueDo
                                Case "n": sb = sb & vbLf: j = j + 2: GoTo ContinueDo
                                Case "r": sb = sb & vbCr: j = j + 2: GoTo ContinueDo
                                Case "t": sb = sb & vbTab: j = j + 2: GoTo ContinueDo
                                Case Else: sb = sb & nx: j = j + 2: GoTo ContinueDo
                            End Select
                        Else
                            sb = sb & "\"
                            j = j + 1
                            GoTo ContinueDo
                        End If
                    ElseIf cj = "'" Then
                        Exit Do
                    Else
                        sb = sb & cj
                        j = j + 1
                    End If
ContinueDo:
                Loop
                If j <= n And mid$(src, j, 1) = "'" Then
                    toks.Add Array("STRING", sb)
                    i = j + 1
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    If GLOBALS_.STRICT_STRINGS Then err.Raise vbObjectError + 1001, "Parser.Tokenize", "Unterminated string literal"
                    toks.Add Array("STRING", sb)
                    i = n + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
            Case Else
                If ch Like "[0-9]" Then
                    Dim j2 As Long: j2 = i
                    Do While j2 <= n And mid$(src, j2, 1) Like "[0-9.]"
                        j2 = j2 + 1
                    Loop
                    toks.Add Array("NUMBER", mid$(src, i, j2 - i))
                    i = j2
                    GoTo Tokenize_SkipToLoopEnd
                ElseIf ch Like "[A-Za-z_]" Then
                    Dim j3 As Long: j3 = i
                    Do While j3 <= n And mid$(src, j3, 1) Like "[A-Za-z0-9_.]"
                        j3 = j3 + 1
                    Loop
                    toks.Add Array("IDENT", mid$(src, i, j3 - i))
                    i = j3
                    GoTo Tokenize_SkipToLoopEnd
                ' forced VBAexpression: @(...). Capture whole balanced parentheses block as a single VBEXPR token.
                ElseIf ch = "@" And i < n And mid$(src, i + 1, 1) = "(" Then
                    Dim jvb As Long: jvb = i + 2
                    Dim depthVB As Long: depthVB = 1
                    Dim vbContent As String: vbContent = ""
                    Do While jvb <= n And depthVB > 0
                        Dim cjvb As String: cjvb = mid$(src, jvb, 1)
                        If cjvb = "(" Then
                            depthVB = depthVB + 1
                            vbContent = vbContent & cjvb
                            jvb = jvb + 1
                        ElseIf cjvb = ")" Then
                            depthVB = depthVB - 1
                            If depthVB = 0 Then
                                ' matched closing ')', stop without adding this ')'
                                Exit Do
                            Else
                                vbContent = vbContent & cjvb
                                jvb = jvb + 1
                            End If
                        Else
                            vbContent = vbContent & cjvb
                            jvb = jvb + 1
                        End If
                    Loop
                    ' vbContent now contains inner characters between parentheses
                    toks.Add Array("VBEXPR", vbContent)
                    i = jvb + 1  ' position after ')'
                    GoTo Tokenize_SkipToLoopEnd
                Else
                    toks.Add Array("SYM", ch)
                    i = i + 1
                    GoTo Tokenize_SkipToLoopEnd
                End If
        End Select

Tokenize_SkipToLoopEnd:
    Loop

    ' collapse array index tokens: IDENT [ NUMBER ] -> IDENT "name[index]"
    Dim out As New Collection
    Dim idx As Long
    idx = 1
Tokenize_CollapseLoop:
    Do While idx <= toks.count
        If idx > toks.count Then Exit Do ' FIX: Bounds
        Dim tk As Variant: tk = toks(idx)
        If tk(0) = "IDENT" And idx + 3 <= toks.count Then
            If toks(idx + 1)(0) = "PAREN" And toks(idx + 1)(1) = "[" And toks(idx + 2)(0) = "NUMBER" And toks(idx + 3)(0) = "PAREN" And toks(idx + 3)(1) = "]" Then
                out.Add Array("IDENT", tk(1) & "[" & toks(idx + 2)(1) & "]")
                idx = idx + 4
                GoTo Tokenize_CollapseLoop
            End If
        End If
        out.Add tk
        idx = idx + 1
    Loop

    Set Tokenize = out
End Function


Public Function JoinTokens(tokens As Collection) As String
    Dim s As String: s = ""
    Dim i As Long
    For i = 1 To tokens.count
        If tokens(i)(0) = "STRING" Then
            s = s & ToLiteralString(CStr(tokens(i)(1))) & " "
        Else
            s = s & tokens(i)(1) & " "
        End If
    Next i
    JoinTokens = Trim$(s)
End Function
Private Function ToLiteralString(ByRef aString As String) As String
    ToLiteralString = d_Apostrophe & aString & d_Apostrophe
End Function
