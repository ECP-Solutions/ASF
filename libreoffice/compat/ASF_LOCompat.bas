' ASF_LOCompat - LibreOffice Basic support module for the ASF library.
'
' The ASF sources are written for VBA. tools/build.py rewrites them for LibreOffice Basic
' and routes the constructs LibreOffice handles differently through the helpers below.
' This module also holds the entry points for code outside the ASF library: LibreOffice
' cannot create a class of another library with New, so use the factories instead.
'
'   BasicLibraries.LoadLibrary "ASF"
'   Dim engine As Object: Set engine = ASF.ASF_LOCompat.NewASF()
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
Public Function LOTryEvaluator() As Object
    On Error Resume Next
    With GlobalScope.BasicLibraries
        If .hasByName("VBAExpressionsLib") Then
            If Not .isLibraryLoaded("VBAExpressionsLib") Then .loadLibrary "VBAExpressionsLib"
        End If
    End With
    Set LOTryEvaluator = ASFNewVBAexpressions()
    Err.Clear
End Function

Public Function LONewEvaluator() As Object
    Dim ev As Object
    Set ev = LOTryEvaluator()
    If ev Is Nothing Then
        Err.Raise LOvbObjectError() + 7100, "ASF", "VBA expressions need the VBA Expressions extension (VBAExpressionsLib library)"
    End If
    Set LONewEvaluator = ev
End Function

Public Function LOEvaluator(g As Object) As Object
    If g.gExprEvaluator Is Nothing Then Set g.gExprEvaluator = LONewEvaluator()
    Set LOEvaluator = g.gExprEvaluator
End Function

' ArrayFromString2 with VBA's castTypeData: numbers and booleans typed, strings unquoted
Public Function LOExprArray(ev As Object, ByVal s As String) As Variant
    LOExprArray = LOCastExprItems(ev.ArrayFromString2(s))
End Function

Private Function LOCastExprItems(ByVal a As Variant) As Variant
    If Not IsArray(a) Then
        LOCastExprItems = LOCastExprItem(a)
        Exit Function
    End If
    Dim t() As Variant
    Dim i As Long
    ReDim t(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        t(i) = LOCastExprItems(a(i))
    Next i
    LOCastExprItems = t
End Function

Private Function LOCastExprItem(ByVal v As Variant) As Variant
    Dim s As String
    s = CStr(v)
    If Len(s) >= 2 And Left$(s, 1) = "'" And Right$(s, 1) = "'" Then
        LOCastExprItem = Mid$(s, 2, Len(s) - 2)
    ElseIf IsNumeric(s) Then
        LOCastExprItem = CDbl(s)
    ElseIf LCase$(s) = "true" Or LCase$(s) = "false" Then
        LOCastExprItem = (LCase$(s) = "true")
    Else
        LOCastExprItem = s
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
Public Function LOTypeName(v As Variant) As String
    If Not IsObject(v) Then
        LOTypeName = TypeName(v)
        Exit Function
    End If
    LOTypeName = LOObjTypeName(v)
End Function

' Separate procedure: LibreOffice rejects a call into a procedure that invokes
' methods on a Variant parameter holding a non-object
Public Function LOObjTypeName(o As Object) As String
    If o Is Nothing Then
        LOObjTypeName = "Nothing"
        Exit Function
    End If
    On Error Resume Next
    Dim n As String
    If HasUnoInterfaces(o, "com.sun.star.uno.XInterface") Then
        n = o.getImplementationName()
        If n = "" Then n = "Object"
        LOObjTypeName = n
        Err.Clear
        Exit Function
    End If
    n = o.LOClassName()
    If Err.Number = 0 And n <> "" Then
        LOObjTypeName = n
        Exit Function
    End If
    Err.Clear
    Dim c As Long
    c = o.Count
    If Err.Number = 0 Then
        LOObjTypeName = "Collection"
    Else
        LOObjTypeName = "Object"
    End If
    Err.Clear
End Function

Public Function LOIsColl(ByVal v As Variant) As Boolean
    LOIsColl = (LOTypeName(v) = "Collection")
End Function

' ---------------- Collections and default members -----------------------------
' coll(i) on a Collection that is a member of another object, or on a class's own
' module-level Collection, goes through a local Collection variable
Public Function LOItem(ByVal c As Object, ByVal k As Variant) As Variant
    Dim cc As Collection: Set cc = c
    If IsObject(cc(k)) Then
        Set LOItem = cc(k)
    Else
        LOItem = cc(k)
    End If
End Function

' v("key") on a Variant: LibreOffice classes have no default member, so call GetValue
Public Function LODefault(v As Variant, ByVal k As Variant) As Variant
    If IsObject(v) Then
        Dim r As Variant
        LODefaultObj v, k, r
        If IsObject(r) Then
            Set LODefault = r
        Else
            LODefault = r
        End If
    Else
        LODefault = v(k)
    End If
End Function

Public Sub LODefaultObj(o As Object, ByVal k As Variant, r As Variant)
    If LOObjTypeName(o) = "Collection" Then
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
Public Function LOLet(ByVal v As Variant) As Variant
    If IsObject(v) Then Err.Raise 91
    LOLet = v
End Function

' ---------------- VBA semantics LibreOffice lacks ------------------------------
Public Function LOvbObjectError() As Long
    LOvbObjectError = -2147221504
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

' ReDim of an array field of a user-defined type
Public Function LORedim(ByVal old As Variant, ByVal lo As Long, ByVal hi As Long, ByVal keep As Boolean) As Variant
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
    LORedim = t
End Function

' New storage for a typed array about to be resized: LibreOffice's ReDim would
' otherwise resize the copies of it already stored elsewhere as well
Public Function LOEmptyStrings() As Variant
    Dim t() As String
    LOEmptyStrings = t
End Function

Public Function LOEmptyLongs() As Variant
    Dim t() As Long
    LOEmptyLongs = t
End Function

Public Function LOEmptyDoubles() As Variant
    Dim t() As Double
    LOEmptyDoubles = t
End Function

' ---------------- Files -------------------------------------------------------
' Script and module files are read as UTF-8
Public Function LOReadTextFile(ByVal path As String) As String
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
    LOReadTextFile = s
End Function

' ---------------- CallByName ---------------------------------------------------
' LibreOffice's CallByName only reaches Basic class instances; Collections and UNO
' objects are dispatched here. LOCallByName itself is generated by tools/build.py.
Private Function LOCallSpecial(obj As Object, ByVal name As String, ByVal kind As Integer, a As Variant, r As Variant) As Boolean
    Dim n As Long
    n = UBound(a) - LBound(a) + 1
    If LOObjTypeName(obj) = "Collection" Then
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
        LOCallSpecial = True
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
        LOCallSpecial = True
    End If
End Function
