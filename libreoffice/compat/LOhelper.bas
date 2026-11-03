' LOhelper - LibreOffice Basic support module for the ASF library.
'
' The ASF sources are written for VBA. tools/build.py rewrites them for LibreOffice Basic
' and routes the constructs LibreOffice handles differently through the helpers below.
' This module also holds the entry points for code outside the ASF library: New only
' finds a class of another library if that library was loaded before the calling module
' was compiled, so use the factories instead.
'
'   GlobalScope.BasicLibraries.LoadLibrary "ASF"
'   Dim engine As Object: Set engine = ASF.LOhelper.NewASF()
'   engine.Run engine.Compile("return 1 + 2;")
'   MsgBox engine.OUTPUT_
Option Explicit

Private pOutput As String
Private pOutputFile As String

' ---------------- Factories -------------------------------------------------
Public Function NewASF() As Object
    Set NewASF = New ASF
End Function

Public Function NewRegex() As Object
    Set NewRegex = New ASF_RegexEngine
End Function

Public Function NewASFMap() As Object
    Set NewASFMap = New ASF_Map
End Function

' ---------------- VBA Expressions ---------------------------------------------
' The @(...) expressions run on VBAExpressionsLib, the LibreOffice port shipped as the
' VBA Expressions extension. ASF works without it until a script evaluates one.
Public Function TryEvaluator() As Object
    On Error Resume Next
    With GlobalScope.BasicLibraries
        If .hasByName("VBAExpressionsLib") Then
            If Not .isLibraryLoaded("VBAExpressionsLib") Then .loadLibrary "VBAExpressionsLib"
        End If
    End With
    Set TryEvaluator = VBAEhelper.NewVBAexpressions()
    Err.Clear
End Function

Public Function NewEvaluator() As Object
    Dim ev As Object
    Set ev = TryEvaluator()
    If ev Is Nothing Then
        Err.Raise ObjectError() + 7100, "ASF", "VBA expressions need the VBA Expressions extension (VBAExpressionsLib library)"
    End If
    Set NewEvaluator = ev
End Function

Public Function Evaluator(g As Object) As Object
    If g.gExprEvaluator Is Nothing Then Set g.gExprEvaluator = NewEvaluator()
    Set Evaluator = g.gExprEvaluator
End Function

' ArrayFromString2 with VBA's castTypeData: numbers and booleans typed, strings unquoted
Public Function ExprArray(ev As Object, ByVal s As String) As Variant
    ExprArray = CastExprItems(ev.ArrayFromString2(s))
End Function

Private Function CastExprItems(ByVal a As Variant) As Variant
    If Not IsArray(a) Then
        CastExprItems = CastExprItem(a)
        Exit Function
    End If
    Dim t() As Variant
    Dim i As Long
    ReDim t(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        t(i) = CastExprItems(a(i))
    Next i
    CastExprItems = t
End Function

Private Function CastExprItem(ByVal v As Variant) As Variant
    Dim s As String
    s = CStr(v)
    If Len(s) >= 2 And Left$(s, 1) = "'" And Right$(s, 1) = "'" Then
        CastExprItem = Mid$(s, 2, Len(s) - 2)
    ElseIf IsNumeric(s) Then
        CastExprItem = CDbl(s)
    ElseIf LCase$(s) = "true" Or LCase$(s) = "false" Then
        CastExprItem = (LCase$(s) = "true")
    Else
        CastExprItem = s
    End If
End Function

' ---------------- Output (VBA's Debug.Print) ----------------------------------
' Text that VBA would send to the Immediate window: ASF print() and the runtime log.
Public Sub LogLine(Optional ByVal s As Variant)
    Dim txt As String
    If Not IsMissing(s) Then txt = CStr(s)
    pOutput = pOutput & txt & Chr$(10)
    If pOutputFile <> "" Then
        Dim fh As Integer: fh = FreeFile
        Open pOutputFile For Append As #fh
        Print #fh, txt
        Close #fh
    End If
End Sub

Public Function GetOutput() As String
    GetOutput = pOutput
End Function

Public Sub ClearOutput()
    pOutput = ""
End Sub

' Also append every output line to a text file; an empty path stops it
Public Sub SetOutputFile(ByVal path As String)
    pOutputFile = path
End Sub

' ---------------- Type names --------------------------------------------------
' TypeName reports every object as "Object" in LibreOffice; each generated class
' answers LOClassName, UNO objects their implementation name, Collections are
' recognised by Count.
Public Function TypeNameOf(v As Variant) As String
    If Not IsObject(v) Then
        TypeNameOf = TypeName(v)
        Exit Function
    End If
    TypeNameOf = ObjTypeName(v)
End Function

' Separate procedure: LibreOffice rejects a call into a procedure that invokes
' methods on a Variant parameter holding a non-object
Public Function ObjTypeName(o As Object) As String
    If o Is Nothing Then
        ObjTypeName = "Nothing"
        Exit Function
    End If
    On Error Resume Next
    Dim n As String
    If HasUnoInterfaces(o, "com.sun.star.uno.XInterface") Then
        n = o.getImplementationName()
        If n = "" Then n = "Object"
        ObjTypeName = n
        Err.Clear
        Exit Function
    End If
    n = o.LOClassName()
    If Err.Number = 0 And n <> "" Then
        ObjTypeName = n
        Exit Function
    End If
    Err.Clear
    Dim c As Long
    c = o.Count
    If Err.Number = 0 Then
        ObjTypeName = "Collection"
    Else
        ObjTypeName = "Object"
    End If
    Err.Clear
End Function

Public Function IsColl(ByVal v As Variant) As Boolean
    IsColl = (TypeNameOf(v) = "Collection")
End Function

' ---------------- Collections and default members -----------------------------
' coll(i) on a Collection that is a member of another object, or on a class's own
' module-level Collection, goes through a local Collection variable
Public Function Item(ByVal c As Object, ByVal k As Variant) As Variant
    Dim cc As Collection: Set cc = c
    If IsObject(cc(k)) Then
        Set Item = cc(k)
    Else
        Item = cc(k)
    End If
End Function

' v("key") on a Variant: LibreOffice classes have no default member, so call GetValue
Public Function DefaultValue(v As Variant, ByVal k As Variant) As Variant
    If IsObject(v) Then
        Dim r As Variant
        DefaultValueObj v, k, r
        If IsObject(r) Then
            Set DefaultValue = r
        Else
            DefaultValue = r
        End If
    Else
        DefaultValue = v(k)
    End If
End Function

Public Sub DefaultValueObj(o As Object, ByVal k As Variant, r As Variant)
    If ObjTypeName(o) = "Collection" Then
        Dim c As Collection: Set c = o
        If IsObject(c(k)) Then
            Set r = c(k)
        Else
            r = c(k)
        End If
    ElseIf IsObject(o.GetValue(k)) Then
        Set r = o.GetValue(k)
    Else
        r = o.GetValue(k)
    End If
End Sub

' x = obj.GetValue(k) without Set: VBA raises when the value is an object, LibreOffice
' silently stores it; ASF relies on the VBA behaviour under On Error Resume Next
Public Function LetValue(ByVal v As Variant) As Variant
    If IsObject(v) Then Err.Raise 91
    LetValue = v
End Function

' ---------------- VBA semantics LibreOffice lacks ------------------------------
Public Function ObjectError() As Long
    ObjectError = -2147221504
End Function

' VBA's CLng raises on a non-numeric string; LibreOffice returns 0
Public Function VBACLng(ByVal v As Variant) As Long
    If VarType(v) = vbString Then
        If Not IsNumeric(v) Then Err.Raise 13, , "Type mismatch"
    End If
    VBACLng = CLng(v)
End Function

Public Function VBASplit(ByVal s As String, ByVal d As String) As Variant
    VBASplit = Split(s, d)
End Function

Public Function VBAReplace(ByVal s As String, ByVal f As String, ByVal r As String) As String
    VBAReplace = Replace(s, f, r)
End Function

' VBA's Timer with millisecond resolution, for measuring elapsed time
Public Function TimerValue() As Double
    TimerValue = GetSystemTicks() / 1000
End Function

' ReDim of an array field of a user-defined type
Public Function RedimField(ByVal old As Variant, ByVal lo As Long, ByVal hi As Long, ByVal keep As Boolean) As Variant
    Dim t() As Variant
    ReDim t(lo To hi)
    If keep And IsArray(old) Then
        Dim i As Long
        For i = lo To hi
            If i >= LBound(old) And i <= UBound(old) Then
                If IsObject(old(i)) Then
                    Set t(i) = old(i)
                Else
                    t(i) = old(i)
                End If
            End If
        Next i
    End If
    RedimField = t
End Function

' New storage for a typed array about to be resized: LibreOffice's ReDim would
' otherwise resize the copies of it already stored elsewhere as well
Public Function EmptyStrings() As Variant
    Dim t() As String
    EmptyStrings = t
End Function

Public Function EmptyLongs() As Variant
    Dim t() As Long
    EmptyLongs = t
End Function

Public Function EmptyDoubles() As Variant
    Dim t() As Double
    EmptyDoubles = t
End Function

' ---------------- Files -------------------------------------------------------
' Script and module files are read as UTF-8
Public Function ReadTextFile(ByVal path As String) As String
    Dim sfa As Object, inp As Object, txt As Object
    Dim s As String
    Set sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    Set inp = sfa.openFileRead(ConvertToURL(path))
    Set txt = CreateUnoService("com.sun.star.io.TextInputStream")
    txt.setInputStream inp
    txt.setEncoding "UTF-8"
    Do While Not txt.isEOF()
        s = s & txt.readString(Array(), False)
    Loop
    txt.closeInput
    If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    ReadTextFile = s
End Function

' ---------------- CallByName ---------------------------------------------------
' LibreOffice's CallByName only reaches Basic class instances; Collections and UNO
' objects are dispatched here. CallMember itself is generated by tools/build.py.
Private Function CallSpecial(obj As Object, ByVal name As String, ByVal kind As Integer, a As Variant, r As Variant) As Boolean
    Dim n As Long
    n = UBound(a) - LBound(a) + 1
    If ObjTypeName(obj) = "Collection" Then
        Dim c As Collection: Set c = obj
        Select Case LCase$(name)
            Case "item"
                If IsObject(c(a(LBound(a)))) Then Set r = c(a(LBound(a))) Else r = c(a(LBound(a)))
            Case "count"
                r = c.Count
            Case "add"
                Select Case n
                    Case 1: c.Add a(LBound(a))
                    Case 2: c.Add a(LBound(a)), a(LBound(a) + 1)
                    Case 3: c.Add a(LBound(a)), a(LBound(a) + 1), a(LBound(a) + 2)
                    Case Else: c.Add a(LBound(a)), a(LBound(a) + 1), a(LBound(a) + 2), a(LBound(a) + 3)
                End Select
            Case "remove"
                c.Remove a(LBound(a))
            Case Else
                Err.Raise 438
        End Select
        CallSpecial = True
    ElseIf HasUnoInterfaces(obj, "com.sun.star.uno.XInterface") Then
        Dim inv As Object
        Set inv = CreateUnoService("com.sun.star.script.Invocation").createInstanceWithArguments(Array(obj))
        ' UNO names are case-sensitive, ASF passes them lower-cased
        Dim infos As Variant, info As Variant, member As String, i As Long
        infos = inv.getInfo()
        For i = LBound(infos) To UBound(infos)
            If LCase$(infos(i).aName) = LCase$(name) Then
                info = infos(i)
                member = info.aName
                Exit For
            End If
        Next i
        If member = "" Then Err.Raise 438, , "Object doesn't support this property or method: " & name
        If (kind And (VbLet Or VbSet)) <> 0 Then
            inv.setValue member, a(LBound(a))
        ElseIf info.eMemberType = com.sun.star.script.MemberType.METHOD Then
            Dim outIdx() As Variant, outArgs() As Variant
            Dim res As Variant
            res = inv.invoke(member, a, outIdx, outArgs)
            If IsObject(res) Then Set r = res Else r = res
        Else
            If IsObject(inv.getValue(member)) Then Set r = inv.getValue(member) Else r = inv.getValue(member)
        End If
        CallSpecial = True
    End If
End Function
