"""stdLambda for the LibreOffice edition, for BenchmarkLO.py.

stdLambda needs Windows: Scripting.Dictionary, VBScript.RegExp, rtcCallByName in VBE7.dll,
Implements and a predeclared instance. LibreOffice has none of them, so the class is patched:
dictionaries become LambdaDict, the tokeniser's regexes run on ASF_RegexEngine (LambdaRegExp),
members are called through CallByName (stdLambda's own Mac path), and VBA.Math./VBA.Conversion.
qualifiers are dropped. The parser and the evaluation loop are not changed.
"""
import re
import urllib.request

STDVBA_COMMIT = "230aea3aef43cd4bb78a728e7a18f70e11b9a77b"
RAW = "https://raw.githubusercontent.com/sancarn/stdVBA/{}/src/{}"

LAMBDA_DICT = '''Private pKeys() As String
Private pValues() As Variant
Private pCount As Long

Private Sub Class_Initialize()
    ReDim pKeys(0 To 7)
    ReDim pValues(0 To 7)
End Sub

Private Function IndexOf(ByVal key As String) As Long
    Dim i As Long
    For i = 0 To pCount - 1
        If pKeys(i) = key Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

Public Function Exists(ByVal key As String) As Boolean
    Exists = IndexOf(key) >= 0
End Function

Public Sub Add(ByVal key As String, ByVal value As Variant)
    If IndexOf(key) >= 0 Then Err.Raise 457, "LambdaDict", "Key already exists: " & key
    If pCount > UBound(pKeys) Then
        ReDim Preserve pKeys(0 To 2 * pCount - 1)
        ReDim Preserve pValues(0 To 2 * pCount - 1)
    End If
    pKeys(pCount) = key
    If IsObject(value) Then Set pValues(pCount) = value Else pValues(pCount) = value
    pCount = pCount + 1
End Sub

Public Function Item(ByVal key As String) As Variant
    Dim i As Long: i = IndexOf(key)
    If i < 0 Then Item = Empty: Exit Function
    If IsObject(pValues(i)) Then Set Item = pValues(i) Else Item = pValues(i)
End Function

Public Sub SetItem(ByVal key As String, ByVal value As Variant)
    Dim i As Long: i = IndexOf(key)
    If i < 0 Then
        Add key, value
    ElseIf IsObject(value) Then
        Set pValues(i) = value
    Else
        pValues(i) = value
    End If
End Sub

Public Function Count() As Long
    Count = pCount
End Function
'''

LAMBDA_REGEXP = '''Public pattern As String
Public ignoreCase As Boolean
Private pEngine As Object
Private pLast As String

Private Function Engine() As Object
    If pEngine Is Nothing Then
        Set pEngine = New ASF_RegexEngine
        pEngine.Init pattern, ignoreCase
    End If
    Set Engine = pEngine
End Function

Public Function test(ByVal s As String) As Boolean
    test = Engine().Test(s)
End Function

Public Function execute(ByVal s As String) As Object
    Dim c As Object: Set c = Engine().Exec(s)
    If c Is Nothing Then pLast = vbNullString Else pLast = c(1)
    Set execute = Me
End Function

Public Function Item0() As String
    Item0 = pLast
End Function
'''


def fetch(name):
    with urllib.request.urlopen(RAW.format(STDVBA_COMMIT, name)) as r:
        return r.read().decode("cp1252").replace("\r\n", "\n")


def preprocess(s, consts):
    """#If/#ElseIf/#Else/#End If on the given constants; #Const lines are dropped."""
    def value(cond):
        cond = cond.strip()
        if cond.lower().startswith("not "):
            return not value(cond[4:])
        return consts[cond]
    out, stack = [], []
    for ln in s.split("\n"):
        t = ln.strip()
        m = re.match(r"#If (.+) Then$", t, re.I)
        if m:
            taken = value(m.group(1))
            stack.append([taken, taken])
            continue
        m = re.match(r"#ElseIf (.+) Then$", t, re.I)
        if m:
            top = stack[-1]
            top[0] = not top[1] and value(m.group(1))
            top[1] = top[1] or top[0]
            continue
        if re.match(r"#Else$", t, re.I):
            top = stack[-1]
            top[0] = not top[1]
            top[1] = True
            continue
        if re.match(r"#End If$", t, re.I):
            stack.pop()
            continue
        if re.match(r"#Const ", t, re.I):
            continue
        if all(f[0] for f in stack):
            out.append(ln)
    return "\n".join(out)


def replace_once(s, old, new, count=1):
    if s.count(old) != count:
        raise SystemExit(f"stdLambda patch: expected {count} of {old!r}, found {s.count(old)}")
    return s.replace(old, new)


def patch(src):
    s = preprocess(src, {"Mac": True, "VBA7": True, "devMode": False})
    s = replace_once(s, "Implements stdICallable\n", "")
    s = re.sub(r'CreateObject\("Scripting\.Dictionary"\)', "New LambdaDict", s)
    s = replace_once(s, 'CreateObject("VBScript.Regexp")', "New LambdaRegExp")
    s = replace_once(s, "oMatch(0)", "oMatch.Item0()", 2)
    # no predeclared instance: every lambda gets its own (empty) extension table
    s = replace_once(s, "Set This.FunctionExtensions = stdLambda.oFunctExt", "Set This.FunctionExtensions = Nothing", 2)
    # a Type variable of a class starts unset: give it a fresh value
    s = replace_once(s, "    If Me Is stdLambda Then Set This.FunctionExtensions = New LambdaDict\n", "    This = NewTThis()\n")
    s = replace_once(s, 'TypeName(This.FunctionExtensions) = "Dictionary"', 'TypeName(This.FunctionExtensions) = "LambdaDict"')
    # default members of Scripting.Dictionary
    dicts = r"This\.(?:Singleton\.Cache|PerformanceCache|FunctionExtensions)"
    s = re.sub(r"Set (" + dicts + r")\(([^()]+)\) = (.+)", r"Call \1.SetItem(\2, \3)", s)
    s = re.sub(r"Let (" + dicts + r")\(([^()]+)\) = (.+)", r"Call \1.SetItem(\2, \3)", s)
    s = re.sub(r"(This\.(?:Singleton\.Cache|PerformanceCache|FunctionExtensions))\(", r"\1.Item(", s)
    # ReDim of an array inside a Type: the arrays become fields of the class
    for field in ("scopesArgCount", "scopes", "tokens", "operations"):
        s = re.sub(r"\bThis\." + field + r"\b", "p" + field[0].upper() + field[1:], s)
    # members through the edition's CallByName, which knows Collections and UNO objects
    s = re.sub(r"\bCallByName\(obj, funcName, callerType", "LOhelper.CallMember(obj, funcName, callerType", s)
    # enum members qualified by their enum
    s = re.sub(r"\b(?:LambdaType|IInstruction)\.(?=i\w)", "", s)
    # constant expressions
    s = replace_once(s, "vbGetOrMethod As Long = VbGet Or VbMethod", "vbGetOrMethod As Long = 3")
    # a field named Type
    s = replace_once(s, "    Type As TokenDefinition\n", "    Kind As TokenDefinition\n")
    s = re.sub(r"\.Type\b", ".Kind", s)
    # and one named name
    s = replace_once(s, "    name As String\n    Regex As String\n", "    TokName As String\n    Regex As String\n")
    s = re.sub(r"(\.Kind|getTokenDefinition)\.name\b", r"\1.TokName", s)
    s = re.sub(r"\bVBA\.(?:Math|Conversion|DateTime)\.", "", s)
    s = re.sub(r"\bVBA\.(?=\w+\()", "", s)
    return s


def variant_arrays(s):
    """LibreOffice cannot ReDim an array of a Type that holds an object: arrays of Types become
    Variant arrays whose elements are created by New<Type>() (stdLambdaTypes)."""
    r = lambda old, new, n=1: replace_once(s, old, new, n)
    s = r("Private pOperations() As Operation\n", "Private pOperations() As Variant\n")
    s = r("Private pTokens() As token\n", "Private pTokens As Variant\n")
    s = r("    Dim arr() As TokenDefinition\n", "    Dim arr() As Variant\n")
    s = r("    Dim defs() As TokenDefinition\n", "    Dim defs() As Variant\n")
    s = r("    Dim tokens() As token, iTokenDef As Long\n", "    Dim tokens() As Variant, iTokenDef As Long\n")
    s = r("                ReDim Preserve tokens(1 To iNumTokens)\n",
          "                ReDim Preserve tokens(1 To iNumTokens)\n                tokens(iNumTokens) = NewToken()\n")
    s = r("    ReDim Preserve tokens(1 To iNumTokens + 1)\n",
          "    ReDim Preserve tokens(1 To iNumTokens + 1)\n    tokens(iNumTokens + 1) = NewToken()\n")
    s = r("ByRef tokens() As token,", "ByRef tokens() As Variant,")
    s = r("    ReDim Preserve tokens(LBound(tokens) To (UBound(tokens) - iCountRemoved))\n    removeTokens = tokens\n",
          "    Dim kept() As Variant, k As Long\n    ReDim kept(LBound(tokens) To UBound(tokens) - iCountRemoved)\n"
          "    For k = LBound(kept) To UBound(kept)\n        kept(k) = tokens(k)\n    Next k\n    removeTokens = kept\n")
    s = r("Private Function ShiftTokens() As token", "Private Function ShiftTokens() As Variant")
    s = r("    Dim firstToken As token\n", "    Dim firstToken As Variant\n")
    s = r("Private Function evaluate(ByRef ops() As Operation,", "Private Function evaluate(ByRef ops() As Variant,")
    s = r("    Dim op As Operation\n", "    Dim op As Variant\n")
    i = s.index("Private Function getTokenDefinition(")
    j = s.index("End Function", i)
    s = s[:i] + ("Private Function getTokenDefinition(ByVal sName As String, ByVal sRegex As String, Optional ByVal ignoreCase As Boolean = True, Optional ByVal isKeyword As Boolean = False) As Variant\n"
                 "    Dim d As Variant: d = NewTokenDefinition()\n"
                 "    d.TokName = sName\n"
                 "    d.Regex = sRegex & IIf(isKeyword, \"\\b\", \"\")\n"
                 "    Set d.RegexObj = New LambdaRegExp\n"
                 "    d.RegexObj.pattern = \"^(?:\" & sRegex & IIf(isKeyword, \"\\b\", \"\") & \")\"\n"
                 "    d.RegexObj.ignoreCase = ignoreCase\n"
                 "    getTokenDefinition = d\n") + s[j:]
    s = r("""    With pOperations(This.iOperationIndex)
        .Instruction = kInstruction
        Call CopyVariant(.value, value)
    End With
""", """    FillOperations pOperations
    pOperations(This.iOperationIndex).Instruction = kInstruction
    If Not IsMissing(value) Then
        If IsObject(value) Then Set pOperations(This.iOperationIndex).value = value Else pOperations(This.iOperationIndex).value = value
    End If
""")
    s = r("    ReDim Preserve pOperations(0 To This.iOperationIndex)\nEnd Sub",
          "    ReDim Preserve pOperations(0 To This.iOperationIndex)\n    FillOperations pOperations\nEnd Sub")
    return s


TYPE_HELPERS = """
Public Function NewTThis() As TThis
    Dim t As TThis
    NewTThis = t
End Function

Public Function NewToken() As token
    Dim t As token
    NewToken = t
End Function

Public Function NewTokenDefinition() As TokenDefinition
    Dim t As TokenDefinition
    NewTokenDefinition = t
End Function

Public Function NewOperation() As Operation
    Dim t As Operation
    NewOperation = t
End Function

Public Sub FillOperations(ByRef ops As Variant)
    Dim i As Long
    For i = LBound(ops) To UBound(ops)
        If IsEmpty(ops(i)) Then ops(i) = NewOperation()
    Next i
End Sub
"""


def gosub_lines(s):
    """After Return, LibreOffice goes on at the next line: a statement after GoSub X: is skipped."""
    return re.sub(r"^(\s*)GoSub (\w+): ", r"\1GoSub \2\n\1", s, flags=re.M)


def result_copies(s):
    """LibreOffice cannot pass a function's own name ByRef: CopyVariant(F, x) inside F goes
    through a local."""
    out, fname = [], None
    for ln in s.split("\n"):
        m = re.match(r"^(?:Public |Private |Friend )?(?:Function|Property Get) (\w+)\(", ln)
        if m:
            fname = m.group(1)
            out.append(ln)
            out.append("    Dim cvResult As Variant")
            continue
        if fname and re.match(r"^End (Function|Property)", ln):
            fname = None
        if fname:
            m = re.match(r"^(.*?)Call CopyVariant\(" + fname + r", (.*)\)(\s*'[^\"]*)?$", ln)
            if m:
                ln = (f"{m.group(1)}Call CopyVariant(cvResult, {m.group(2)}): "
                      f"If IsObject(cvResult) Then Set {fname} = cvResult Else {fname} = cvResult")
        out.append(ln)
    return "\n".join(out)


def split_types(s):
    """LibreOffice has no Type in class modules: (class body, standard module with the Types)."""
    types = re.findall(r"^Private Type \w+\n.*?^End Type\n", s, flags=re.M | re.S)
    for t in types:
        s = s.replace(t, "")
    mod = "".join(types).replace("Private Type ", "Public Type ")
    for field in ("tokens() As token", "operations() As Operation", "scopes() As Variant", "scopesArgCount() As Variant"):
        mod = mod.replace("    " + field + "\n", "")
        s = s.replace("Private This As TThis\n", "Private This As TThis\nPrivate p" + field[0].upper() + field[1:] + "\n")
    mod += TYPE_HELPERS
    mod = mod.replace("Instruction As IInstruction", "Instruction As Long").replace("Lambda As stdLambda", "Lambda As Object")
    s = variant_arrays(s)
    s = result_copies(gosub_lines(s))
    # functions declared As T() return a Variant array
    s = re.sub(r"^((?:Public |Private )?Function [^\n]*\)) As \w+\(\)[ \t]*$", r"\1 As Variant", s, flags=re.M)
    return s, mod


def lo_modules(build, src=None):
    """[(name, code, kind)]: the patched stdLambda, its Types and the two shims."""
    src = src if src is not None else fetch("stdLambda.cls")
    body, types = split_types(build.hoist_dims(build.strip_vba_header(patch(src))))
    return [("stdLambda", build.class_module("stdLambda", body), "class"),
            ("stdLambdaTypes", "Option VBASupport 1\n" + types, "normal"),
            ("LambdaDict", build.class_module("LambdaDict", LAMBDA_DICT), "class"),
            ("LambdaRegExp", build.class_module("LambdaRegExp", LAMBDA_REGEXP), "class")]
