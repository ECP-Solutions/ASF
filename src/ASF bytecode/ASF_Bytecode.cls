VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Bytecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ASF_Bytecode: compiles the AST of a program, and the named functions it calls, to the bytecode
' that ASF_FastVM runs (see README.md). A program that uses something this compiler does not take
' is left to the tree VM: Prepare returns False and Reason says why.
Option Explicit

' opcodes: the numbers are the jump table of ASF_FastVM.Execute, keep both lists in step
Private Const OP_HALT As Long = 1
Private Const OP_PUSHK As Long = 2      ' a: constant
Private Const OP_PUSHE As Long = 3      ' undefined
Private Const OP_LOADL As Long = 4      ' a: local slot
Private Const OP_STOREL As Long = 5
Private Const OP_LOADG As Long = 6      ' a: global slot
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
Private Const OP_JMP As Long = 26       ' a: target
Private Const OP_JMPF As Long = 27      ' pops, jumps when falsy
Private Const OP_JMPT As Long = 28      ' pops, jumps when truthy
Private Const OP_ANDJ As Long = 29      ' falsy: False and jump; truthy: pops
Private Const OP_ORJ As Long = 30       ' truthy: True and jump; falsy: pops
Private Const OP_TOBOOL As Long = 31
Private Const OP_NULLJ As Long = 32     ' ??: defined: jumps keeping it; else pops
Private Const OP_POP As Long = 33
Private Const OP_CALLF As Long = 34     ' a: function, b: argument count
Private Const OP_RET As Long = 35
Private Const OP_RET0 As Long = 36
Private Const OP_PRINT As Long = 37     ' a: argument count
Private Const OP_INCL As Long = 38      ' local a = local a + constant b
Private Const OP_INCG As Long = 39      ' global a = global a + constant b
Private Const OP_JNLT As Long = 40      ' compares the two values on loopTop, jumps to a when false
Private Const OP_JNLE As Long = 41
Private Const OP_JNGT As Long = 42
Private Const OP_JNGE As Long = 43
Private Const OP_JNEQ As Long = 44
Private Const OP_JNNE As Long = 45
Private Const OP_ADDK As Long = 46      ' loopTop + constant a
Private Const OP_DUP As Long = 47
Private Const OP_LOADC As Long = 48     ' a: local slot holding a cell
Private Const OP_STOREC As Long = 49
Private Const OP_LOADE As Long = 50     ' a: entry of the closure's environment
Private Const OP_STOREE As Long = 51
Private Const OP_MKCLOS As Long = 52    ' a: function literal
Private Const OP_MKFUNC As Long = 53    ' a: named function, as a value
Private Const OP_CALL As Long = 54      ' a: argument count; the callee is under the arguments
Private Const OP_NEWARR As Long = 55    ' a: items
Private Const OP_NEWOBJ As Long = 56    ' a: key and value pairs
Private Const OP_GETIDX As Long = 57
Private Const OP_GETIDXW As Long = 58   ' GETIDX for a container about to be written: no bounds error
Private Const OP_SETIDXR As Long = 59   ' value, container, index -> the updated container
Private Const OP_GETPROP As Long = 60   ' a: key constant
Private Const OP_SETPROPV As Long = 61  ' value, object -> object.key = value; a: key constant
Private Const OP_CALLM As Long = 62     ' a: method name, b: arguments; the receiver is under them
Private Const OP_CALLMW As Long = 63    ' CALLM, then the receiver to write back (push, sort...)
Private Const OP_EXPAND As Long = 64    ' a: arguments, b: constant flagging the spread ones -> items, count
Private Const OP_CALLFV As Long = 65    ' CALLF with the argument count on the stack
Private Const OP_CALLV As Long = 66
Private Const OP_CALLMV As Long = 67
Private Const OP_CALLMWV As Long = 68
Private Const OP_CALLB As Long = 69     ' a: built-in name, b: arguments
Private Const OP_CALLBV As Long = 70
Private Const OP_TPL As Long = 71       ' a: parts of a template
Private Const OP_NEWARRS As Long = 72   ' a: items, b: constant flagging the spread ones
Private Const OP_JNSW As Long = 73      ' switch: compares with =, jumps to a when different
Private Const OP_ITERLIST As Long = 74  ' a: 0 for-of, 1 for-in: the values or keys to visit
Private Const OP_FORNEXT As Long = 75   ' list, counter -> the next item, or a jump to a at the end
Private Const OP_UNPACK As Long = 76    ' a: targets, b: 1 with a rest target
Private Const OP_CLASS As Long = 77     ' a: the class value to copy (constant), b: class number
Private Const OP_NEWINST As Long = 78   ' a: class name -> a new instance
Private Const OP_BINDM As Long = 79     ' a: method name, b: its function; pops the instance and binds the method to it
Private Const OP_CALLFT As Long = 80    ' a: function, b: arguments, under them this
Private Const OP_THROW As Long = 81     ' pops the value thrown
Private Const OP_IMPORT As Long = 82    ' a: the import (source, defaultImport, namespaceImport, named) -> the values it binds
Private Const OP_TREESTMT As Long = 83  ' a: a statement the tree VM runs (prototype, export)
Private Const OP_VBAEXPR As Long = 84   ' a: the VBA expression, b: the names of the values under it, one per line
Private Const OP_PROTO As Long = 85     ' a: the prototype (classType, methodName, funcIndex, params); pops its function
Private Const OP_EXPORT As Long = 86    ' a: exported name; pops the value (a module being loaded keeps it)
Private Const OP_EXPORTPROTO As Long = 87 ' a: the prototype, with exportName; pops its function
Private Const METHOD_SHIFT As Long = 1024 ' b of CALLM...: arguments + METHOD_SHIFT * method number

Private Const NOT_TAKEN As Long = -2147212480    ' vbObjectError + 9024 + 100: raised when the program is left to the tree VM

Private GLOBALS_ As ASF_Globals

' the code of every program prepared so far
Private op_() As Long, a_() As Long, b_() As Long, nOp_ As Long
Private k_() As Variant, nK_ As Long, kMap_ As ASF_Map
Private fnEntry_() As Long, fnNParams_() As Long, fnNLocals_() As Long, fnMaxStack_() As Long
Private fnName_() As String, nFn_ As Long
Private fnCellStart_() As Long, fnCellCount_() As Long  ' locals boxed in cells at the entry
Private cellSlot_() As Long, nCellSlot_ As Long
Private fnCapStart_() As Long, fnCapCount_() As Long    ' what MKCLOS copies into the environment
Private fnRest_() As Long, fnThis_() As Long, fnThisEnv_() As Long   ' slots of the rest parameter and this, -1 when none
Private capKind_() As Long, capIdx_() As Long, nCap_ As Long   ' kind 0: a local of the creator, 1: its environment
Private fnKind_() As Long               ' 1: a method (keeps a throw to itself), 2: a constructor (also gives its instance)
Private tryStart_() As Long, tryEnd_() As Long, tryHandler_() As Long, tryDepth_() As Long, nTry_ As Long
Private clsNew_() As Long, nCls_ As Long                ' class number -> its constructor function
Private progFn_ As ASF_Map              ' program index -> its main function, -1 when not taken
Private progReason_ As ASF_Map
Private progGlobals_ As ASF_Map         ' program index -> names of its global slots
Private progGBase_ As ASF_Map           ' program index -> the slot before its first global
Private nGTotal_ As Long                ' global slots of all programs: each program has its own range
Private reason_ As String
Private version_ As Long

' the program being compiled
Private gSlot_ As ASF_Map               ' global name -> slot
Private assignedGlobals_ As ASF_Map     ' the names the program assigns at its top level
Private gNames_() As String, nG_ As Long, gBase_ As Long
Private isModule_ As Boolean            ' the program is a module: export keeps values
Private pendingExports_ As Collection   ' export fun/class: the names exported when the module ends
Private fnOf_ As ASF_Map                ' named function -> its function number in this program
Private pending_ As Collection          ' named functions still to compile
Private classes_ As ASF_Map             ' class name -> its declaration, for the classes of the program
Private ctorFn_ As ASF_Map              ' class name -> the function of its constructor alone (for super)
Private pendingCls_ As Collection       ' class functions still to compile: Array(kind, declaration, node, function)
Private methodFns_ As Collection        ' Array(method declaration, its function) for the classes of the program

' functions: one context (an ASF_Map) each, with
'   kind      "main", "named" or "literal"
'   locals    name -> slot           celled  name -> True (a local a nested function uses)
'   free      name -> environment entry, with freeKind and freeIdx (Collections) for MKCLOS
'   parent    the enclosing context of a literal; node, the literal; fn, its function number
Private cx_ As ASF_Map                  ' the context being compiled
Private literals_ As Collection         ' contexts of the function literals of the program
Private pendingLits_ As Collection      ' literals still to compile
Private depth_ As Long, maxDepth_ As Long
Private breaks_ As Collection           ' one Collection of jump positions per enclosing loop
Private conts_ As Collection
Private tempCount_ As Long

Private Sub Class_Initialize()
    ReDim op_(1 To 256): ReDim a_(1 To 256): ReDim b_(1 To 256)
    ReDim k_(1 To 64)
    ReDim fnEntry_(1 To 16): ReDim fnNParams_(1 To 16): ReDim fnNLocals_(1 To 16)
    ReDim fnMaxStack_(1 To 16): ReDim fnName_(1 To 16)
    ReDim fnCellStart_(1 To 16): ReDim fnCellCount_(1 To 16)
    ReDim fnCapStart_(1 To 16): ReDim fnCapCount_(1 To 16)
    ReDim fnRest_(1 To 16): ReDim fnThis_(1 To 16): ReDim fnThisEnv_(1 To 16)
    ReDim cellSlot_(1 To 16): ReDim capKind_(1 To 16): ReDim capIdx_(1 To 16)
    ReDim fnKind_(1 To 16): ReDim tryStart_(1 To 4): ReDim tryEnd_(1 To 4): ReDim tryHandler_(1 To 4): ReDim tryDepth_(1 To 4): ReDim clsNew_(1 To 4)
    Set kMap_ = New ASF_Map
    Set progFn_ = New ASF_Map
    Set progReason_ = New ASF_Map
    Set progGlobals_ = New ASF_Map
    Set progGBase_ = New ASF_Map
End Sub

Public Sub SetGlobals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Sub

' ---------------------------------------------------------------- what ASF_FastVM reads
Public Property Get Reason() As String
    Reason = reason_
End Property
Public Property Get Version() As Long
    Version = version_
End Property
Public Function Ops() As Variant
    Ops = op_
End Function
Public Function ArgsA() As Variant
    ArgsA = a_
End Function
Public Function ArgsB() As Variant
    ArgsB = b_
End Function
Public Function Constants() As Variant
    Constants = k_
End Function
Public Function FunctionEntries() As Variant
    FunctionEntries = fnEntry_
End Function
Public Function FunctionParams() As Variant
    FunctionParams = fnNParams_
End Function
Public Function FunctionLocals() As Variant
    FunctionLocals = fnNLocals_
End Function
Public Function FunctionStacks() As Variant
    FunctionStacks = fnMaxStack_
End Function
Public Function FunctionCellStarts() As Variant
    FunctionCellStarts = fnCellStart_
End Function
Public Function FunctionCellCounts() As Variant
    FunctionCellCounts = fnCellCount_
End Function
Public Function CellSlots() As Variant
    CellSlots = cellSlot_
End Function
Public Function FunctionCaptureStarts() As Variant
    FunctionCaptureStarts = fnCapStart_
End Function
Public Function FunctionCaptureCounts() As Variant
    FunctionCaptureCounts = fnCapCount_
End Function
Public Function FunctionRests() As Variant
    FunctionRests = fnRest_
End Function
Public Function FunctionThis() As Variant
    FunctionThis = fnThis_
End Function
Public Function FunctionThisEnv() As Variant
    FunctionThisEnv = fnThisEnv_
End Function
Public Function FunctionKinds() As Variant
    FunctionKinds = fnKind_
End Function
Public Function TryStarts() As Variant
    TryStarts = tryStart_
End Function
Public Function TryEnds() As Variant
    TryEnds = tryEnd_
End Function
Public Function TryHandlers() As Variant
    TryHandlers = tryHandler_
End Function
Public Function TryDepths() As Variant
    TryDepths = tryDepth_
End Function
Public Function TryCount() As Long
    TryCount = nTry_
End Function
Public Function ClassConstructors() As Variant
    ClassConstructors = clsNew_
End Function
Public Function ClassCount() As Long
    ClassCount = nCls_
End Function
Public Function CaptureKinds() As Variant
    CaptureKinds = capKind_
End Function
Public Function CaptureIndexes() As Variant
    CaptureIndexes = capIdx_
End Function
Public Function FunctionName(ByVal fn As Long) As String
    FunctionName = fnName_(fn)
End Function
Public Function ProgramFunction(ByVal idx As Long) As Long
    ProgramFunction = progFn_.GetValue(CStr(idx))
End Function
' names of the global slots of program idx, from 1
Public Function ProgramGlobalBase(ByVal idx As Long) As Long
    ProgramGlobalBase = progGBase_.GetValue(CStr(idx))
End Function
Public Property Get GlobalCount() As Long
    GlobalCount = nGTotal_
End Property
Public Function ProgramGlobals(ByVal idx As Long) As Variant
    ProgramGlobals = progGlobals_.GetValue(CStr(idx))
End Function

' The code of program idx as text, one instruction per line
Public Function Disassemble(ByVal idx As Long) As String
    Dim names As Variant
    names = Split("HALT PUSHK PUSHE LOADL STOREL LOADG STOREG ADD SUB MUL DIV MOD POW EQ NE LT LE GT GE CAT SHL SHR " & _
                  "NOT NEG TYPEOF JMP JMPF JMPT ANDJ ORJ TOBOOL NULLJ POP CALLF RET RET0 PRINT INCL INCG " & _
                  "JNLT JNLE JNGT JNGE JNEQ JNNE ADDK DUP LOADC STOREC LOADE STOREE MKCLOS MKFUNC CALL NEWARR NEWOBJ " & _
                  "GETIDX GETIDXW SETIDXR GETPROP SETPROPV CALLM CALLMW EXPAND CALLFV CALLV CALLMV CALLMWV CALLB CALLBV TPL " & _
                  "NEWARRS JNSW ITERLIST FORNEXT UNPACK CLASS NEWINST BINDM CALLFT THROW IMPORT TREESTMT VBAEXPR PROTO EXPORT EXPORTPROTO", " ")
    Dim out As String, i As Long, fn As Long, k As Variant
    If Not Prepare(idx) Then
        Disassemble = "not taken: " & reason_
        Exit Function
    End If
    For fn = 1 To nFn_
        out = out & "fn " & fn & " " & fnName_(fn) & " entry " & fnEntry_(fn) & " params " & fnNParams_(fn) & _
              " locals " & fnNLocals_(fn) & " stack " & fnMaxStack_(fn) & " cells " & fnCellCount_(fn) & _
              " captures " & fnCapCount_(fn) & vbLf
    Next fn
    For i = 1 To nK_
        out = out & "k" & i & " = " & TypeName(k_(i)) & ConstantText(k_(i)) & vbLf
    Next i
    For i = 1 To nOp_
        out = out & i & vbTab & names(op_(i) - 1) & " " & a_(i) & " " & b_(i)
        Select Case op_(i)
            Case OP_PUSHK, OP_ADDK, OP_GETPROP, OP_SETPROPV, OP_CALLM, OP_CALLMW, OP_CALLMV, OP_CALLMWV, OP_CALLB, OP_CALLBV
                out = out & "  ; " & TypeName(k_(a_(i))) & ConstantText(k_(a_(i)))
        End Select
        out = out & vbLf
    Next i
    Disassemble = out
End Function

Private Function ConstantText(ByRef k As Variant) As String
    If Not IsObject(k) And Not IsNull(k) Then ConstantText = " " & CStr(k)
End Function

' ---------------------------------------------------------------- compiling a program
' Compiles program idx (once); False when it is left to the tree VM
Public Function Prepare(ByVal idx As Long) As Boolean
    If idx < 1 Or idx > GLOBALS_.gPrograms.Count Then
        reason_ = "no program " & idx
        Exit Function
    End If
    Dim key As String: key = CStr(idx)
    If progFn_.Exists(key) Then
        Prepare = (progFn_.GetValue(key) >= 0)
        If Not Prepare Then reason_ = progReason_.GetValue(key)
        Exit Function
    End If
    Dim saveOp As Long, saveK As Long, saveFn As Long, mainFn As Long, saveCell As Long, saveCap As Long
    Dim saveCls As Long, saveTry As Long, saveG As Long
    saveOp = nOp_: saveK = nK_: saveFn = nFn_: saveCell = nCellSlot_: saveCap = nCap_
    saveCls = nCls_: saveTry = nTry_: saveG = nGTotal_
    On Error GoTo NotTaken
    mainFn = CompileProgram(idx)
    progFn_.SetValue key, mainFn
    reason_ = vbNullString
    version_ = version_ + 1
    Prepare = True
    Exit Function
NotTaken:
    If Err.Number = NOT_TAKEN Then reason_ = Err.Description Else reason_ = "compiler error: " & Err.Description
    nOp_ = saveOp: nFn_ = saveFn: nCellSlot_ = saveCell: nCap_ = saveCap
    nCls_ = saveCls: nTry_ = saveTry: nGTotal_ = saveG
    If nK_ <> saveK Then
        Dim i As Long
        For i = saveK + 1 To nK_
            k_(i) = Empty
        Next i
        nK_ = saveK
        Set kMap_ = New ASF_Map
        For i = 1 To nK_
            If IsEmpty(k_(i)) Then
                kMap_.SetValue "E", i
            ElseIf IsNull(k_(i)) Then
                kMap_.SetValue "N", i
            ElseIf Not IsObject(k_(i)) Then
                kMap_.SetValue ConstKey(k_(i)), i
            End If
        Next i
    End If
    progFn_.SetValue key, -1
    progReason_.SetValue key, reason_
End Function

Private Sub NotSupported(ByVal what As String)
    Err.Raise NOT_TAKEN, "ASF_Bytecode", what
End Sub

Private Function CompileProgram(ByVal idx As Long) As Long
    Dim p As Variant: p = GLOBALS_.gPrograms(idx)
    Dim stmts As Collection: Set stmts = p(1)
    Set gSlot_ = New ASF_Map
    ReDim gNames_(1 To 16): nG_ = 0: gBase_ = nGTotal_
    isModule_ = (Left$(CStr(p(0)), 8) = "@module:")
    Set pendingExports_ = New Collection
    Set fnOf_ = New ASF_Map
    Set pending_ = New Collection
    Set classes_ = New ASF_Map
    Set ctorFn_ = New ASF_Map
    Set pendingCls_ = New Collection
    Set methodFns_ = New Collection
    Dim i As Long
    For i = 1 To stmts.Count
        If stmts(i).GetValue("type") = "ClassDecl" Then
            If classes_.Exists(CStr(stmts(i).GetValue("name"))) Then NotSupported "a class declared twice"
            classes_.SetValue CStr(stmts(i).GetValue("name")), stmts(i)
        End If
    Next i

    ' the names a program assigns at its loopTop level are its globals
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    CollectAssigned stmts, assigned
    For i = 1 To assigned.Count
        GlobalSlot assigned.key(i)
    Next i
    Set assignedGlobals_ = assigned

    Set literals_ = New Collection
    Set pendingLits_ = New Collection
    Dim mainFn As Long: mainFn = NewFunction("@main")
    Dim ctx As ASF_Map: Set ctx = NewContext("main", Nothing)
    AnalyzeLiteralsIn stmts, ctx
    BeginFunction ctx
    fnEntry_(mainFn) = nOp_ + 1
    EmitBlock stmts
    Dim nm As Variant
    For Each nm In pendingExports_
        EmitLoadName CStr(nm)
        Emit OP_EXPORT, Constant(CStr(nm)), 0, -1
    Next nm
    Emit OP_HALT, 0, 0, 0
    EndFunction mainFn, 0

    Dim name As String, work As Variant
    Do While pending_.Count > 0 Or pendingLits_.Count > 0 Or pendingCls_.Count > 0
        If pending_.Count > 0 Then
            name = pending_(1)
            pending_.Remove 1
            CompileNamedFunction name
        ElseIf pendingCls_.Count > 0 Then
            work = pendingCls_(1)
            pendingCls_.Remove 1
            CompileClassFunction work
        Else
            Set ctx = pendingLits_(1)
            pendingLits_.Remove 1
            CompileLiteral ctx
        End If
    Loop

    Dim names() As String
    ReDim names(0 To nG_)
    For i = 1 To nG_
        names(i) = gNames_(i)
    Next i
    progGlobals_.SetValue CStr(idx), names
    progGBase_.SetValue CStr(idx), gBase_
    nGTotal_ = gBase_ + nG_
    CompileProgram = mainFn
End Function

Private Sub CompileNamedFunction(ByVal name As String)
    Dim fIdx As Long: fIdx = GLOBALS_.gFuncTable.GetValue(name)
    Dim p As Variant: p = GLOBALS_.gPrograms(fIdx)
    Dim restName As String
    If UBound(p) >= 3 Then
        If Not IsEmpty(p(3)) Then restName = CStr(p(3))
    End If
    Dim body As Collection: Set body = p(1)
    Dim fn As Long: fn = fnOf_.GetValue(name)
    Dim ctx As ASF_Map: Set ctx = NewContext("named", Nothing)

    Dim params As Variant, i As Long
    If GLOBALS_.gFuncParams.TryGetValue(name, params) Then
        If IsArray(params) Then
            If ArrayHasItems(params) Then
                For i = LBound(params) To UBound(params)
                    If Len(CStr(params(i))) > 0 Then AddLocal ctx, CStr(params(i))
                Next i
            End If
        End If
    End If
    Dim nParams As Long: nParams = ctx.GetValue("nL")
    If Len(restName) > 0 Then fnRest_(fn) = AddLocal(ctx, restName)

    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    CollectAssigned body, assigned
    For i = 1 To assigned.Count
        If Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
    Next i
    Dim used As ASF_Map: Set used = New ASF_Map
    CollectUsed body, used
    If used.Exists("this") Then fnThis_(fn) = AddLocal(ctx, "this")

    AnalyzeLiteralsIn body, ctx
    BeginFunction ctx
    fnEntry_(fn) = nOp_ + 1
    EmitBlock body
    Emit OP_RET0, 0, 0, 0
    EndFunction fn, nParams
End Sub

Private Sub CompileLiteral(ctx As ASF_Map)
    Dim lit As ASF_Map: Set lit = ctx.GetValue("node")
    Dim fn As Long: fn = ctx.GetValue("fn")
    Dim locals As ASF_Map: Set locals = ctx.GetValue("locals")
    If ctx.Exists("restSlot") Then fnRest_(fn) = ctx.GetValue("restSlot")
    If locals.Exists("this") Then fnThis_(fn) = locals.GetValue("this")
    Dim free As ASF_Map: Set free = ctx.GetValue("free")
    If free.Exists("%outerthis") Then fnThisEnv_(fn) = free.GetValue("%outerthis")
    BeginFunction ctx
    fnEntry_(fn) = nOp_ + 1
    EmitBlock lit.GetValue("body")
    Emit OP_RET0, 0, 0, 0
    EndFunction fn, ctx.GetValue("nParams")
End Sub

' ---------------------------------------------------------------- classes
' A class declaration makes the class value of the tree VM, and functions for its constructor
' (which also sets the fields and binds the methods, as InstantiateClass does) and its methods
Private Sub EmitClassDecl(node As ASF_Map)
    If cx_.GetValue("kind") <> "main" Then NotSupported "a class declared inside a function"
    Dim name As String: name = node.GetValue("name")
    If Not classes_.Exists(name) Then NotSupported "a class declared inside a block"
    If Not classes_.GetValue(name) Is node Then NotSupported "a class declared inside a block"
    If node.Exists("superClass") Then
        If Not classes_.Exists(CStr(node.GetValue("superClass"))) Then NotSupported "a class extending one the program does not declare"
    End If
    nCls_ = nCls_ + 1
    If nCls_ > UBound(clsNew_) Then ReDim Preserve clsNew_(1 To 2 * nCls_)
    clsNew_(nCls_) = NewFunction("@new " & name)
    pendingCls_.Add Array("new", node, node, clsNew_(nCls_))
    ' the class value holds descriptors of its fields and methods, not their code
    Dim cls As ASF_Map: Set cls = New ASF_Map
    cls.Add "type", "Class"
    cls.SetValue "name", name
    Dim fields As Collection: Set fields = New Collection
    Dim item As Variant
    If Not node.GetValue("fields") Is Nothing Then
        For Each item In node.GetValue("fields")
            fields.Add CStr(item.GetValue("name"))
        Next item
    End If
    cls.SetValue "fields", fields
    Dim methods As Collection: Set methods = node.GetValue("classMethods")
    Dim descs As Collection: Set descs = New Collection
    Dim i As Long, fn As Long
    For i = 1 To methods.Count
        fn = NewFunction(name & "." & methods(i).GetValue("name"))
        methodFns_.Add Array(methods(i), fn)
        pendingCls_.Add Array("method", node, methods(i), fn)
        descs.Add MethodDescriptor(methods(i), fn)
    Next i
    cls.SetValue "classMethods", descs
    If node.Exists("constructor") Then cls.SetValue "constructor", MethodDescriptor(node.GetValue("constructor"), 0)
    If node.Exists("superClass") Then cls.SetValue "superClass", node.GetValue("superClass")
    Emit OP_CLASS, ObjectConstant(cls), nCls_, 1
    EmitStoreName name
End Sub

' {type: 'Method', name, params, isStatic, fn}: fn is a Long, unlike a closure's, and 0 for a constructor
Private Function MethodDescriptor(node As ASF_Map, ByVal fn As Long) As ASF_Map
    Dim d As ASF_Map: Set d = New ASF_Map
    d.Add "type", "Method"
    d.SetValue "name", CStr(node.GetValue("name"))
    Dim params As Collection: Set params = New Collection
    Dim item As Variant
    If node.Exists("params") Then
        For Each item In node.GetValue("params")
            params.Add CStr(item)
        Next item
    End If
    d.SetValue "params", params
    Dim isStatic As Boolean
    If node.Exists("isStatic") Then isStatic = node.GetValue("isStatic")
    d.SetValue "isStatic", isStatic
    If fn > 0 Then d.SetValue "fn", fn
    Set MethodDescriptor = d
End Function

' The function of the constructor of a class, alone: what super(...) calls
Private Function ConstructorFunction(cls As ASF_Map) As Long
    Dim name As String: name = cls.GetValue("name")
    If ctorFn_.Exists(name) Then
        ConstructorFunction = ctorFn_.GetValue(name)
    Else
        ConstructorFunction = NewFunction("@ctor " & name)
        ctorFn_.SetValue name, ConstructorFunction
        pendingCls_.Add Array("ctor", cls, cls.GetValue("constructor"), ConstructorFunction)
    End If
End Function

' work: Array(kind, class declaration, method or constructor node, function)
Private Sub CompileClassFunction(ByVal work As Variant)
    Dim cls As ASF_Map: Set cls = work(1)
    Dim node As ASF_Map: Set node = work(2)
    Dim fn As Long: fn = work(3)
    Dim ctx As ASF_Map: Set ctx = NewContext("named", Nothing)
    ctx.SetValue "class", cls
    Dim parent As ASF_Map, ctor As ASF_Map, body As Collection
    If cls.Exists("superClass") Then Set parent = classes_.GetValue(CStr(cls.GetValue("superClass")))
    If work(0) = "new" Then
        ' the constructor of the class, or else the one of its parent, runs in this function
        If cls.Exists("constructor") Then
            Set ctor = cls.GetValue("constructor")
        ElseIf Not parent Is Nothing Then
            If parent.Exists("constructor") Then
                Set ctor = parent.GetValue("constructor")
                ctx.SetValue "class", parent
            End If
        End If
        ctx.SetValue "ctor", True
        fnKind_(fn) = 2
    Else
        Set ctor = node
        fnKind_(fn) = 1
    End If
    If Not ctor Is Nothing Then
        If ctor.Exists("body") Then Set body = ctor.GetValue("body")
    End If

    Dim i As Long, params As Collection
    If Not ctor Is Nothing Then
        If ctor.Exists("params") Then Set params = ctor.GetValue("params")
    End If
    If Not params Is Nothing Then
        For i = 1 To params.Count
            AddLocal ctx, CStr(params(i))
        Next i
    End If
    Dim nParams As Long: nParams = ctx.GetValue("nL")
    fnThis_(fn) = AddLocal(ctx, "this")
    If Not body Is Nothing Then
        Dim assigned As ASF_Map: Set assigned = New ASF_Map
        CollectAssigned body, assigned
        For i = 1 To assigned.Count
            If Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
        Next i
    End If

    If work(0) = "new" Then
        If Not parent Is Nothing Then AnalyzeLiteralsIn parent.GetValue("fields"), ctx
        AnalyzeLiteralsIn cls.GetValue("fields"), ctx
    End If
    AnalyzeLiteralsIn body, ctx
    BeginFunction ctx
    fnEntry_(fn) = nOp_ + 1
    If work(0) = "new" Then
        Emit OP_NEWINST, Constant(CStr(cls.GetValue("name"))), 0, 1
        Emit OP_STOREL, fnThis_(fn), 0, -1
        ' as InstantiateClass: the parent's fields and all its methods, then the class's own
        If Not parent Is Nothing Then
            EmitFields parent.GetValue("fields"), fnThis_(fn)
            EmitBindMethods parent.GetValue("classMethods"), fnThis_(fn), True
        End If
        EmitFields cls.GetValue("fields"), fnThis_(fn)
        EmitBindMethods cls.GetValue("classMethods"), fnThis_(fn), False
    End If
    EmitBlock body
    If work(0) = "new" Then
        Emit OP_LOADL, fnThis_(fn), 0, 1
        Emit OP_RET, 0, 0, -1
    Else
        Emit OP_RET0, 0, 0, 0
    End If
    EndFunction fn, nParams
End Sub

Private Sub EmitFields(fields As Collection, ByVal thisSlot As Long)
    Dim i As Long
    If fields Is Nothing Then Exit Sub
    For i = 1 To fields.Count
        If fields(i).Exists("default") Then
            EmitExpr fields(i).GetValue("default")
        Else
            Emit OP_PUSHE, 0, 0, 1
        End If
        Emit OP_LOADL, thisSlot, 0, 1
        Emit OP_SETPROPV, Constant(CStr(fields(i).GetValue("name"))), 0, -2
    Next i
End Sub

Private Sub EmitBindMethods(methods As Collection, ByVal thisSlot As Long, ByVal withStatic As Boolean)
    Dim i As Long, isStatic As Boolean
    If methods Is Nothing Then Exit Sub
    For i = 1 To methods.Count
        isStatic = False
        If methods(i).Exists("isStatic") Then isStatic = methods(i).GetValue("isStatic")
        If withStatic Or Not isStatic Then
            Emit OP_LOADL, thisSlot, 0, 1
            Emit OP_BINDM, Constant(CStr(methods(i).GetValue("name"))), MethodFunctionOf(methods(i)), -1
        End If
    Next i
End Sub

' The function of a method declaration, as EmitClassDecl made it
Private Function MethodFunctionOf(node As ASF_Map) As Long
    Dim pair As Variant
    For Each pair In methodFns_
        If pair(0) Is node Then
            MethodFunctionOf = pair(1)
            Exit Function
        End If
    Next pair
    NotSupported "a method of a class not declared yet"
End Function

' super(args): the constructor of the parent of the class being compiled, with the same this
Private Sub EmitSuperCall(node As ASF_Map)
    Dim args As Collection: Set args = node.GetValue("args")
    Dim n As Long, i As Long, cls As ASF_Map, parent As ASF_Map
    If Not args Is Nothing Then n = args.Count
    If cx_.Exists("class") Then Set cls = cx_.GetValue("class")
    If Not cls Is Nothing Then
        If cls.Exists("superClass") Then Set parent = classes_.GetValue(CStr(cls.GetValue("superClass")))
    End If
    If Not parent Is Nothing Then
        If Not parent.Exists("constructor") Then Set parent = Nothing
    End If
    If parent Is Nothing Then
        For i = 1 To n
            EmitExpr args(i)
            Emit OP_POP, 0, 0, -1
        Next i
    Else
        EmitLoadName "this"
        For i = 1 To n
            EmitExpr args(i)
        Next i
        Emit OP_CALLFT, ConstructorFunction(parent), n, -n
        Emit OP_POP, 0, 0, -1
    End If
    Emit OP_PUSHE, 0, 0, 1
End Sub

Private Function ArrayHasItems(ByRef v As Variant) As Boolean
    On Error Resume Next
    ArrayHasItems = (UBound(v) >= LBound(v))
End Function

' ---------------------------------------------------------------- functions and slots
Private Function NewFunction(ByVal name As String) As Long
    nFn_ = nFn_ + 1
    If nFn_ > UBound(fnEntry_) Then
        ReDim Preserve fnEntry_(1 To 2 * nFn_): ReDim Preserve fnNParams_(1 To 2 * nFn_)
        ReDim Preserve fnNLocals_(1 To 2 * nFn_): ReDim Preserve fnMaxStack_(1 To 2 * nFn_)
        ReDim Preserve fnName_(1 To 2 * nFn_)
        ReDim Preserve fnCellStart_(1 To 2 * nFn_): ReDim Preserve fnCellCount_(1 To 2 * nFn_)
        ReDim Preserve fnCapStart_(1 To 2 * nFn_): ReDim Preserve fnCapCount_(1 To 2 * nFn_)
        ReDim Preserve fnRest_(1 To 2 * nFn_): ReDim Preserve fnThis_(1 To 2 * nFn_)
        ReDim Preserve fnThisEnv_(1 To 2 * nFn_): ReDim Preserve fnKind_(1 To 2 * nFn_)
    End If
    fnName_(nFn_) = name
    fnRest_(nFn_) = -1: fnThis_(nFn_) = -1: fnThisEnv_(nFn_) = -1: fnKind_(nFn_) = 0
    NewFunction = nFn_
End Function

' The function number of a named function in this program; queued to compile on first use
Private Function NamedFunction(ByVal name As String) As Long
    If fnOf_.Exists(name) Then
        NamedFunction = fnOf_.GetValue(name)
    Else
        NamedFunction = NewFunction(name)
        fnOf_.SetValue name, NamedFunction
        pending_.Add name
    End If
End Function

Private Sub BeginFunction(ctx As ASF_Map)
    Set cx_ = ctx
    depth_ = 0: maxDepth_ = 0
    Set breaks_ = New Collection
    Set conts_ = New Collection
End Sub

' The frame of fn: its locals, the ones boxed in cells, and what its closures capture
Private Sub EndFunction(ByVal fn As Long, ByVal nParams As Long)
    fnNParams_(fn) = nParams
    fnNLocals_(fn) = cx_.GetValue("nL")
    fnMaxStack_(fn) = maxDepth_ + 2
    Dim celled As ASF_Map: Set celled = cx_.GetValue("celled")
    Dim locals As ASF_Map: Set locals = cx_.GetValue("locals")
    Dim i As Long
    fnCellStart_(fn) = nCellSlot_ + 1
    fnCellCount_(fn) = celled.Count
    For i = 1 To celled.Count
        nCellSlot_ = nCellSlot_ + 1
        If nCellSlot_ > UBound(cellSlot_) Then
            ReDim Preserve cellSlot_(1 To 2 * nCellSlot_)
        End If
        cellSlot_(nCellSlot_) = locals.GetValue(celled.key(i))
    Next i
    Dim kinds As Collection: Set kinds = cx_.GetValue("freeKind")
    Dim idxs As Collection: Set idxs = cx_.GetValue("freeIdx")
    fnCapStart_(fn) = nCap_ + 1
    fnCapCount_(fn) = kinds.Count
    For i = 1 To kinds.Count
        nCap_ = nCap_ + 1
        If nCap_ > UBound(capKind_) Then
            ReDim Preserve capKind_(1 To 2 * nCap_)
            ReDim Preserve capIdx_(1 To 2 * nCap_)
        End If
        capKind_(nCap_) = kinds(i)
        capIdx_(nCap_) = idxs(i)
    Next i
End Sub

Private Function NewContext(ByVal kind As String, parent As ASF_Map) As ASF_Map
    Dim ctx As ASF_Map: Set ctx = New ASF_Map
    ctx.SetValue "kind", kind
    ctx.SetValue "locals", New ASF_Map
    ctx.SetValue "nL", 0
    ctx.SetValue "celled", New ASF_Map
    ctx.SetValue "free", New ASF_Map
    ctx.SetValue "freeKind", New Collection
    ctx.SetValue "freeIdx", New Collection
    ctx.SetValue "parent", parent
    Set NewContext = ctx
End Function

Private Function AddLocal(ctx As ASF_Map, ByVal name As String) As Long
    Dim locals As ASF_Map: Set locals = ctx.GetValue("locals")
    If locals.Exists(name) Then
        AddLocal = locals.GetValue(name)
    Else
        AddLocal = ctx.GetValue("nL")
        locals.SetValue name, AddLocal
        ctx.SetValue "nL", AddLocal + 1
    End If
End Function

Private Function GlobalSlot(ByVal name As String) As Long
    If gSlot_.Exists(name) Then
        GlobalSlot = gSlot_.GetValue(name)
    Else
        nG_ = nG_ + 1
        If nG_ > UBound(gNames_) Then ReDim Preserve gNames_(1 To 2 * nG_)
        gNames_(nG_) = name
        gSlot_.SetValue name, gBase_ + nG_
        GlobalSlot = gBase_ + nG_
    End If
End Function

' The names a block assigns, outside the function literals it holds
Private Sub CollectAssigned(stmts As Collection, found As ASF_Map)
    Dim i As Long
    For i = 1 To stmts.Count
        CollectAssignedIn stmts(i), found
    Next i
End Sub

Private Sub CollectAssignedIn(node As ASF_Map, found As ASF_Map)
    If node Is Nothing Then Exit Sub
    Dim t As String: t = node.GetValue("type")
    Dim target As ASF_Map
    Select Case t
        Case "Assign"
            Set target = node.GetValue("left")
            ' a[i] = v writes the array back to the variable a
            Do While target.GetValue("type") = "Index"
                Set target = target.GetValue("base")
                If target Is Nothing Then Exit Sub
            Loop
            If target.GetValue("type") = "Variable" Then found.SetValue CStr(target.GetValue("name")), True
        Case "If"
            CollectAssigned node.GetValue("then"), found
            Dim blocks As Collection: Set blocks = node.GetValue("elseif_blocks")
            Dim i As Long
            If Not blocks Is Nothing Then
                For i = 1 To blocks.Count
                    CollectAssigned blocks(i), found
                Next i
            End If
            If node.GetValue("hasElse") Then CollectAssigned node.GetValue("else"), found
        Case "For"
            If node.Exists("init") Then CollectAssignedIn node.GetValue("init"), found
            If node.Exists("step") Then CollectAssignedIn node.GetValue("step"), found
            If node.Exists("each") Then
                Set target = node.GetValue("each")
                If Not target Is Nothing Then
                    If target.GetValue("type") = "Variable" Then found.SetValue CStr(target.GetValue("name")), True
                End If
            End If
            CollectAssigned node.GetValue("body"), found
        Case "While"
            CollectAssigned node.GetValue("body"), found
        Case "Switch"
            Dim cases As Collection: Set cases = node.GetValue("cases")
            Dim pair As Variant, block As Collection
            If Not cases Is Nothing Then
                For Each pair In cases
                    Set block = pair(2)
                    CollectAssigned block, found
                Next pair
            End If
            If Not node.GetValue("default") Is Nothing Then CollectAssigned node.GetValue("default"), found
        Case "ArrayDestructuring"
            Dim targetItem As Variant
            For Each targetItem In node.GetValue("targets")
                found.SetValue CStr(targetItem.GetValue("name")), True
            Next targetItem
            If node.Exists("restTarget") Then found.SetValue CStr(node.GetValue("restTarget")), True
        Case "ClassDecl"
            found.SetValue CStr(node.GetValue("name")), True
        Case "Import"
            Dim names As Collection: Set names = ImportNames(node)
            Dim nm As Variant
            For Each nm In names
                found.SetValue CStr(nm), True
            Next nm
        Case "TryCatch"
            CollectAssigned node.GetValue("try"), found
            If node.Exists("catchVar") Then found.SetValue CStr(node.GetValue("catchVar")), True
            If Not node.GetValue("catch") Is Nothing Then CollectAssigned node.GetValue("catch"), found
    End Select
End Sub

' ---------------------------------------------------------------- closures
' The function literals of a body, outside other literals, each with its context: locals, and the
' variables it takes from enclosing functions (which become cells there)
Private Sub AnalyzeLiteralsIn(node As Variant, parent As ASF_Map)
    Dim i As Long, nd As ASF_Map, item As Variant
    If Not IsObject(node) Then Exit Sub
    If node Is Nothing Then Exit Sub
    If TypeOf node Is Collection Then
        For Each item In node
            AnalyzeLiteralsIn item, parent
        Next item
        Exit Sub
    End If
    If Not TypeOf node Is ASF_Map Then Exit Sub
    Set nd = node
    If nd.Exists("type") Then
        Select Case nd.GetValue("type")
            Case "FuncLiteral"
                AnalyzeLiteral nd, parent
                Exit Sub
            Case "ClassDecl"
                Exit Sub
        End Select
    End If
    For i = 1 To nd.Count
        If IsObject(nd.GetValue(nd.key(i))) Then AnalyzeLiteralsIn nd.GetValue(nd.key(i)), parent
    Next i
End Sub

Private Sub AnalyzeLiteral(lit As ASF_Map, parent As ASF_Map)
    Dim ctx As ASF_Map: Set ctx = NewContext("literal", parent)
    ctx.SetValue "node", lit
    Dim params As Collection: Set params = lit.GetValue("params")
    Dim i As Long
    If Not params Is Nothing Then
        For i = 1 To params.Count
            If Len(CStr(params(i))) > 0 Then AddLocal ctx, CStr(params(i))
        Next i
    End If
    ctx.SetValue "nParams", ctx.GetValue("nL")
    If lit.Exists("restParam") Then
        If Len(CStr(lit.GetValue("restParam"))) > 0 Then ctx.SetValue "restSlot", AddLocal(ctx, CStr(lit.GetValue("restParam")))
    End If
    Dim body As Collection: Set body = lit.GetValue("body")
    ' a name it assigns is its own unless an enclosing function or the program has it
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    CollectAssigned body, assigned
    For i = 1 To assigned.Count
        If Not EnclosingHas(parent, assigned.key(i)) And Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
    Next i
    literals_.Add ctx
    AnalyzeLiteralsIn body, ctx
    Dim used As ASF_Map: Set used = New ASF_Map
    CollectUsed body, used
    Dim locals As ASF_Map: Set locals = ctx.GetValue("locals")
    ' this: the receiver of a method call, or else the this of the enclosing function
    If used.Exists("this") Then
        AddLocal ctx, "this"
        ResolveFreeAs ctx, "this", "%outerthis"
    End If
    For i = 1 To used.Count
        If Not locals.Exists(used.key(i)) Then ResolveFree ctx, used.key(i)
    Next i
End Sub

' True when a function enclosing ctx (not the program) has name as a local
Private Function EnclosingHas(ByVal ctx As ASF_Map, ByVal name As String) As Boolean
    Dim locals As ASF_Map
    Do While Not ctx Is Nothing
        If ctx.GetValue("kind") = "main" Then Exit Function
        Set locals = ctx.GetValue("locals")
        If locals.Exists(name) Then
            EnclosingHas = True
            Exit Function
        End If
        Set ctx = ctx.GetValue("parent")
    Loop
End Function

' Makes name an entry of ctx's environment when an enclosing function has it; False otherwise
Private Function ResolveFree(ctx As ASF_Map, ByVal name As String) As Boolean
    ResolveFree = ResolveFreeAs(ctx, name, name)
End Function

' ResolveFree for name, kept in the environment as key
Private Function ResolveFreeAs(ctx As ASF_Map, ByVal name As String, ByVal key As String) As Boolean
    Dim free As ASF_Map: Set free = ctx.GetValue("free")
    If free.Exists(key) Then
        ResolveFreeAs = True
        Exit Function
    End If
    Dim parent As ASF_Map: Set parent = ctx.GetValue("parent")
    If parent Is Nothing Then Exit Function
    If parent.GetValue("kind") = "main" Then Exit Function
    Dim locals As ASF_Map: Set locals = parent.GetValue("locals")
    Dim kinds As Collection: Set kinds = ctx.GetValue("freeKind")
    Dim idxs As Collection: Set idxs = ctx.GetValue("freeIdx")
    Dim celled As ASF_Map
    If locals.Exists(name) Then
        Set celled = parent.GetValue("celled")
        celled.SetValue name, True
        kinds.Add 0
        idxs.Add locals.GetValue(name)
    ElseIf ResolveFree(parent, name) Then
        Dim pfree As ASF_Map: Set pfree = parent.GetValue("free")
        kinds.Add 1
        idxs.Add pfree.GetValue(name)
    Else
        Exit Function
    End If
    free.SetValue key, kinds.Count - 1
    ResolveFreeAs = True
End Function

' The names a body reads or writes, outside the function literals it holds
Private Sub CollectUsed(node As Variant, found As ASF_Map)
    Dim i As Long, nd As ASF_Map, item As Variant
    If Not IsObject(node) Then Exit Sub
    If node Is Nothing Then Exit Sub
    If TypeOf node Is Collection Then
        For Each item In node
            CollectUsed item, found
        Next item
        Exit Sub
    End If
    If Not TypeOf node Is ASF_Map Then Exit Sub
    Set nd = node
    If nd.Exists("type") Then
        Select Case nd.GetValue("type")
            Case "FuncLiteral": Exit Sub
            Case "Variable": found.SetValue CStr(nd.GetValue("name")), True
        End Select
    End If
    For i = 1 To nd.Count
        If IsObject(nd.GetValue(nd.key(i))) Then CollectUsed nd.GetValue(nd.key(i)), found
    Next i
End Sub

Private Function LiteralContext(lit As ASF_Map) As ASF_Map
    Dim ctx As Variant
    For Each ctx In literals_
        If ctx.GetValue("node") Is lit And ctx.GetValue("parent") Is cx_ Then
            Set LiteralContext = ctx
            Exit Function
        End If
    Next ctx
    NotSupported "a function literal the analysis did not see"
End Function

' ---------------------------------------------------------------- emitting
Private Sub Emit(ByVal op As Long, ByVal a As Long, ByVal b As Long, ByVal stackDelta As Long)
    nOp_ = nOp_ + 1
    If nOp_ > UBound(op_) Then
        ReDim Preserve op_(1 To 2 * nOp_): ReDim Preserve a_(1 To 2 * nOp_): ReDim Preserve b_(1 To 2 * nOp_)
    End If
    op_(nOp_) = op: a_(nOp_) = a: b_(nOp_) = b
    depth_ = depth_ + stackDelta
    If depth_ > maxDepth_ Then maxDepth_ = depth_
End Sub

Private Sub Patch(ByVal at As Long, ByVal target As Long)
    a_(at) = target
End Sub

Private Sub PatchAll(jumps As Collection, ByVal target As Long)
    Dim j As Variant
    For Each j In jumps
        a_(CLng(j)) = target
    Next j
End Sub

Private Function ConstKey(ByRef v As Variant) As String
    ConstKey = VarType(v) & ":" & CStr(v)
End Function

' An object in the constant table (a declaration the VM hands to the tree VM's values)
Private Function ObjectConstant(ByVal obj As Object) As Long
    nK_ = nK_ + 1
    If nK_ > UBound(k_) Then ReDim Preserve k_(1 To 2 * nK_)
    Set k_(nK_) = obj
    ObjectConstant = nK_
End Function

Private Function Constant(ByRef v As Variant) As Long
    Dim key As String
    If IsEmpty(v) Then
        key = "E"
    ElseIf IsNull(v) Then
        key = "N"
    Else
        key = ConstKey(v)
    End If
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

' ---------------------------------------------------------------- statements
Private Sub EmitBlock(stmts As Collection)
    If stmts Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To stmts.Count
        EmitStatement stmts(i)
    Next i
End Sub

Private Sub EmitStatement(node As ASF_Map)
    If node Is Nothing Then Exit Sub
    Select Case node.GetValue("type")
        Case "Assign": EmitAssign node
        Case "ExprStmt"
            EmitExpr node.GetValue("expr")
            Emit OP_POP, 0, 0, -1
        Case "If": EmitIf node
        Case "For": EmitFor node
        Case "While": EmitWhile node
        Case "Return"
            Dim rex As ASF_Map: Set rex = node.GetValue("expr")
            If cx_.Exists("ctor") Then
                ' a constructor gives its instance whatever it returns
                If Not rex Is Nothing Then
                    EmitExpr rex
                    Emit OP_POP, 0, 0, -1
                End If
                EmitLoadName "this"
                Emit OP_RET, 0, 0, -1
            ElseIf rex Is Nothing Then
                Emit OP_RET0, 0, 0, 0
            Else
                EmitExpr rex
                Emit OP_RET, 0, 0, -1
            End If
        Case "Print"
            Dim args As Collection: Set args = node.GetValue("args")
            Dim i As Long
            For i = 1 To args.Count
                EmitExpr args(i)
            Next i
            Emit OP_PRINT, args.Count, 0, -args.Count
        Case "Break"
            If breaks_.Count = 0 Then NotSupported "break outside a loop"
            ' in a try or a catch, the tree VM goes on with the next statement
            If Not breaks_(breaks_.Count) Is Nothing Then
                Emit OP_JMP, 0, 0, 0
                breaks_(breaks_.Count).Add nOp_
            End If
        Case "Continue"
            If conts_.Count = 0 Then NotSupported "continue outside a loop"
            ' in a switch, the tree VM goes on with the case's next statement
            If Not conts_(conts_.Count) Is Nothing Then
                Emit OP_JMP, 0, 0, 0
                conts_(conts_.Count).Add nOp_
            End If
        Case "Switch": EmitSwitch node
        Case "ArrayDestructuring": EmitDestructuring node
        Case "ClassDecl": EmitClassDecl node
        Case "TryCatch": EmitTry node
        Case "Import": EmitImport node
        Case "Prototype": EmitPrototype node
        Case "Export": EmitExport node
        Case "Throw"
            EmitExpr node.GetValue("expr")
            Emit OP_THROW, 0, 0, -1
        Case Else
            NotSupported "statement " & node.GetValue("type")
    End Select
End Sub

' A hidden variable of the function being compiled: a local, or a global in the program
Private Function NewTemp() As String
    tempCount_ = tempCount_ + 1
    NewTemp = "%t" & tempCount_
    If cx_.GetValue("kind") = "main" Then
        GlobalSlot NewTemp
    Else
        AddLocal cx_, NewTemp
    End If
End Function

' switch: the value once, then each case compared with = in order; the first match runs, a
' break leaves the switch, a continue does nothing (as in the tree VM), default when none matches
Private Sub EmitSwitch(node As ASF_Map)
    Dim value As String: value = NewTemp()
    EmitExpr node.GetValue("expr")
    EmitStoreName value
    Dim ends As Collection: Set ends = New Collection
    Dim cases As Collection: Set cases = node.GetValue("cases")
    Dim pair As Variant, nextCase As Long, caseValue As ASF_Map, caseBlock As Collection
    breaks_.Add ends
    conts_.Add Nothing
    If Not cases Is Nothing Then
        For Each pair In cases
            Set caseValue = pair(1)
            Set caseBlock = pair(2)
            EmitLoadName value
            EmitExpr caseValue
            Emit OP_JNSW, 0, 0, -2
            nextCase = nOp_
            EmitBlock caseBlock
            Emit OP_JMP, 0, 0, 0
            ends.Add nOp_
            Patch nextCase, nOp_ + 1
        Next pair
    End If
    If Not node.GetValue("default") Is Nothing Then EmitBlock node.GetValue("default")
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
    PatchAll ends, nOp_ + 1
End Sub

' try { } catch (e) { }: the try block is a region of the table; an error in it (a throw also
' from the functions it calls) clears the stack to its depth, pushes the value and jumps to the catch
Private Sub EmitTry(node As ASF_Map)
    Dim start As Long: start = nOp_ + 1
    Dim depth As Long: depth = depth_
    breaks_.Add Nothing
    conts_.Add Nothing
    EmitBlock node.GetValue("try")
    Emit OP_JMP, 0, 0, 0
    Dim skip As Long: skip = nOp_
    nTry_ = nTry_ + 1
    If nTry_ > UBound(tryStart_) Then
        ReDim Preserve tryStart_(1 To 2 * nTry_): ReDim Preserve tryEnd_(1 To 2 * nTry_)
        ReDim Preserve tryHandler_(1 To 2 * nTry_): ReDim Preserve tryDepth_(1 To 2 * nTry_)
    End If
    tryStart_(nTry_) = start: tryEnd_(nTry_) = nOp_: tryHandler_(nTry_) = nOp_ + 1: tryDepth_(nTry_) = depth
    depth_ = depth + 1
    If depth_ > maxDepth_ Then maxDepth_ = depth_
    If node.Exists("catchVar") Then
        EmitStoreName CStr(node.GetValue("catchVar"))
    Else
        Emit OP_POP, 0, 0, -1
    End If
    If Not node.GetValue("catch") Is Nothing Then EmitBlock node.GetValue("catch")
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
    Patch skip, nOp_ + 1
End Sub

' ---------------------------------------------------------------- modules and the host
' The names an import binds: the default, the namespace, then the named ones
Private Function ImportNames(node As ASF_Map) As Collection
    Dim names As Collection: Set names = New Collection
    If node.Exists("defaultImport") Then names.Add CStr(node.GetValue("defaultImport"))
    If node.Exists("namespaceImport") Then names.Add CStr(node.GetValue("namespaceImport"))
    If node.Exists("namedImports") Then
        Dim spec As Variant
        For Each spec In node.GetValue("namedImports")
            names.Add CStr(spec.GetValue("local"))
        Next spec
    End If
    Set ImportNames = names
End Function

' export: a module being loaded keeps the values by name (in a program, export does nothing but
' register a prototype, as in the tree VM); export fun/class exports the name when the module ends
Private Sub EmitExport(node As ASF_Map)
    Dim spec As Variant
    If node.GetValue("isDefault") Then
        EmitExpr node.GetValue("expression")
        Emit OP_EXPORT, Constant("__default"), 0, -1
    ElseIf node.Exists("namedExports") Then
        For Each spec In node.GetValue("namedExports")
            EmitLoadName CStr(spec.GetValue("local"))
            Emit OP_EXPORT, Constant(CStr(spec.GetValue("exported"))), 0, -1
        Next spec
    ElseIf node.Exists("declarationType") Then
        If isModule_ Then pendingExports_.Add CStr(node.GetValue("declarationName"))
    ElseIf node.Exists("exportType") Then
        If node.GetValue("exportType") = "prototype" Then EmitPrototype node.GetValue("prototypeNode"), CStr(node.GetValue("exportName"))
    End If
End Sub

' prototype.COM.Type name(...) { }: the function is compiled here and runs here when a host object
' of that type is asked for name
Private Sub EmitPrototype(node As ASF_Map, Optional ByVal exportName As String)
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "classType", CStr(node.GetValue("classType"))
    info.SetValue "methodName", CStr(node.GetValue("methodName"))
    info.SetValue "funcIndex", CLng(node.GetValue("funcIndex"))
    Dim params As Collection: Set params = New Collection
    Dim item As Variant
    If Not node.GetValue("params") Is Nothing Then
        For Each item In node.GetValue("params")
            params.Add CStr(item)
        Next item
    End If
    info.SetValue "params", params
    info.SetValue "internalFuncName", CStr(node.GetValue("internalFuncName"))
    Emit OP_MKFUNC, NamedFunction(CStr(node.GetValue("internalFuncName"))), 0, 1
    If Len(exportName) > 0 Then
        info.SetValue "exportName", exportName
        Emit OP_EXPORTPROTO, ObjectConstant(info), 0, -1
    Else
        Emit OP_PROTO, ObjectConstant(info), 0, -1
    End If
End Sub

' import: the VM loads the module (compiled and run once per engine) and pushes the values it binds,
' which go to the program's variables
Private Sub EmitImport(node As ASF_Map)
    If cx_.GetValue("kind") <> "main" Then NotSupported "an import inside a function"
    Dim names As Collection: Set names = ImportNames(node)
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "source", CStr(node.GetValue("source"))
    If node.Exists("defaultImport") Then info.SetValue "defaultImport", True
    If node.Exists("namespaceImport") Then info.SetValue "namespaceImport", True
    Dim named As Collection: Set named = New Collection
    Dim spec As Variant
    If node.Exists("namedImports") Then
        For Each spec In node.GetValue("namedImports")
            named.Add CStr(spec.GetValue("imported"))
        Next spec
    End If
    info.SetValue "named", named
    Emit OP_IMPORT, ObjectConstant(info), 0, names.Count
    Dim i As Long
    For i = names.Count To 1 Step -1
        EmitStoreName names(i)
    Next i
End Sub

' @(expr): the tree VM evaluates it with the variables in scope it names (globals first, so that
' locals win), as EvalVBExpressionWithScope injects them
Private Sub EmitVBAExpr(ByVal expr As String)
    Dim words As ASF_Map: Set words = New ASF_Map
    Dim i As Long, ch As String, word As String
    For i = 1 To Len(expr) + 1
        ch = Mid$(expr, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            word = word & ch
        Else
            If Len(word) > 0 Then
                If Not (Left$(word, 1) Like "[0-9]") Then words.SetValue LCase$(word), True
            End If
            word = vbNullString
        End If
    Next i
    Dim names As Collection: Set names = New Collection
    For i = 1 To assignedGlobals_.Count
        If words.Exists(LCase$(assignedGlobals_.key(i))) Then names.Add assignedGlobals_.key(i)
    Next i
    Dim locals As ASF_Map, free As ASF_Map
    If cx_.GetValue("kind") <> "main" Then
        Set free = cx_.GetValue("free")
        For i = 1 To free.Count
            If words.Exists(LCase$(free.key(i))) Then names.Add free.key(i)
        Next i
        Set locals = cx_.GetValue("locals")
        For i = 1 To locals.Count
            If words.Exists(LCase$(locals.key(i))) Then names.Add locals.key(i)
        Next i
    End If
    Dim joined As String
    For i = 1 To names.Count
        EmitLoadName names(i)
        If i > 1 Then joined = joined & vbLf
        joined = joined & names(i)
    Next i
    Emit OP_VBAEXPR, Constant(expr), Constant(joined), 1 - names.Count
End Sub

' [a, b, ...rest] = source
Private Sub EmitDestructuring(node As ASF_Map)
    Dim targets As Collection: Set targets = node.GetValue("targets")
    Dim hasRest As Long
    If node.Exists("restTarget") Then hasRest = 1
    EmitExpr node.GetValue("source")
    Emit OP_UNPACK, targets.Count, hasRest, targets.Count + hasRest - 1
    Dim i As Long
    For i = 1 To targets.Count
        EmitStoreName targets(i).GetValue("name")
    Next i
    If hasRest = 1 Then EmitStoreName node.GetValue("restTarget")
End Sub

Private Sub EmitAssign(node As ASF_Map)
    Dim target As ASF_Map: Set target = node.GetValue("left")
    Dim rhs As ASF_Map: Set rhs = node.GetValue("right")
    If target.GetValue("type") = "Variable" Then
        ' x = x + constant, on a plain local or a global
        Dim name As String: name = target.GetValue("name")
        Dim k As Variant, slot As Long, kind As String
        If rhs.GetValue("type") = "Binary" Then
            If rhs.GetValue("op") = "+" Then
                If IsVariableNamed(rhs.GetValue("left"), name) And FoldNumber(rhs.GetValue("right"), k) Then
                    kind = NameKind(name, slot)
                    If kind = "L" Then
                        Emit OP_INCL, slot, Constant(k), 0
                        Exit Sub
                    ElseIf kind = "G" Then
                        Emit OP_INCG, slot, Constant(k), 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    EmitExpr rhs
    EmitStoreTop target
End Sub

Private Function IsVariableNamed(node As ASF_Map, ByVal name As String) As Boolean
    If node Is Nothing Then Exit Function
    If node.GetValue("type") = "Variable" Then IsVariableNamed = (node.GetValue("name") = name)
End Function

Private Sub EmitIf(node As ASF_Map)
    Dim ends As Collection: Set ends = New Collection
    Dim falseJumps As Collection: Set falseJumps = New Collection
    EmitJumpIfFalse node.GetValue("cond"), falseJumps
    EmitBlock node.GetValue("then")
    Emit OP_JMP, 0, 0, 0: ends.Add nOp_
    PatchAll falseJumps, nOp_ + 1
    Dim conds As Collection: Set conds = node.GetValue("elseif_conds")
    Dim blocks As Collection: Set blocks = node.GetValue("elseif_blocks")
    Dim i As Long
    If Not conds Is Nothing Then
        For i = 1 To conds.Count
            Set falseJumps = New Collection
            EmitJumpIfFalse conds(i), falseJumps
            EmitBlock blocks(i)
            Emit OP_JMP, 0, 0, 0: ends.Add nOp_
            PatchAll falseJumps, nOp_ + 1
        Next i
    End If
    If node.GetValue("hasElse") Then EmitBlock node.GetValue("else")
    PatchAll ends, nOp_ + 1
End Sub

' A loop's lists of break and continue jumps
Private Sub OpenLoop()
    Dim jumps As Collection
    Set jumps = New Collection
    breaks_.Add jumps
    Set jumps = New Collection
    conts_.Add jumps
End Sub

Private Sub EmitWhile(node As ASF_Map)
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    EmitJumpIfFalse node.GetValue("cond"), exits
    OpenLoop
    EmitBlock node.GetValue("body")
    Emit OP_JMP, loopTop, 0, 0
    PatchAll conts_(conts_.Count), loopTop
    PatchAll exits, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

Private Sub EmitFor(node As ASF_Map)
    Dim kind As String: kind = "classic"
    If node.Exists("kind") Then kind = LCase$(CStr(node.GetValue("kind")))
    If kind <> "classic" Then
        EmitForEach node, kind
        Exit Sub
    End If
    If node.Exists("init") Then EmitLoopPart node.GetValue("init")
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    If node.Exists("cond") Then
        If Not node.GetValue("cond") Is Nothing Then EmitJumpIfFalse node.GetValue("cond"), exits
    End If
    OpenLoop
    EmitBlock node.GetValue("body")
    PatchAll conts_(conts_.Count), nOp_ + 1
    If node.Exists("step") Then EmitLoopPart node.GetValue("step")
    Emit OP_JMP, loopTop, 0, 0
    PatchAll exits, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

' for (x of iterable) and for (k in iterable): the values or keys are taken once, then visited
Private Sub EmitForEach(node As ASF_Map, ByVal kind As String)
    Dim list As String: list = NewTemp()
    Dim counter As String: counter = NewTemp()
    EmitExpr node.GetValue("iter")
    If kind = "in" Then Emit OP_ITERLIST, 1, 0, 0 Else Emit OP_ITERLIST, 0, 0, 0
    EmitStoreName list
    Emit OP_PUSHK, Constant(0#), 0, 1
    EmitStoreName counter
    Dim loopTop As Long: loopTop = nOp_ + 1
    EmitLoadName counter
    Emit OP_ADDK, Constant(1#), 0, 0
    EmitStoreName counter
    EmitLoadName list
    EmitLoadName counter
    Emit OP_FORNEXT, 0, 0, -1
    Dim exitJump As Long: exitJump = nOp_
    Dim target As ASF_Map
    If node.Exists("each") Then Set target = node.GetValue("each")
    If target Is Nothing Then
        Emit OP_POP, 0, 0, -1
    Else
        EmitStoreTop target
    End If
    OpenLoop
    EmitBlock node.GetValue("body")
    Emit OP_JMP, loopTop, 0, 0
    PatchAll conts_(conts_.Count), loopTop
    Patch exitJump, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

' The init or step of a for: a statement, or an expression whose value is dropped
Private Sub EmitLoopPart(part As ASF_Map)
    If part Is Nothing Then Exit Sub
    Select Case part.GetValue("type")
        Case "Assign", "ExprStmt", "Print", "If", "For", "While", "Return", "Break", "Continue"
            EmitStatement part
        Case Else
            EmitExpr part
            Emit OP_POP, 0, 0, -1
    End Select
End Sub

' ---------------------------------------------------------------- conditions
Private Function CompareJump(ByVal op As String) As Long
    Select Case op
        Case "<": CompareJump = OP_JNLT
        Case "<=": CompareJump = OP_JNLE
        Case ">": CompareJump = OP_JNGT
        Case ">=": CompareJump = OP_JNGE
        Case "==", "=": CompareJump = OP_JNEQ
        Case "!=": CompareJump = OP_JNNE
    End Select
End Function

' Jumps (added to jumps, patched by the caller) when cond is falsy; falls through otherwise
Private Sub EmitJumpIfFalse(cond As ASF_Map, jumps As Collection)
    Dim t As String: t = cond.GetValue("type")
    If t = "Binary" Then
        Dim op As String: op = cond.GetValue("op")
        If op = "&&" Then
            EmitJumpIfFalse cond.GetValue("left"), jumps
            EmitJumpIfFalse cond.GetValue("right"), jumps
            Exit Sub
        End If
        If op = "||" Then
            Dim trues As Collection: Set trues = New Collection
            EmitJumpIfTrue cond.GetValue("left"), trues
            EmitJumpIfFalse cond.GetValue("right"), jumps
            PatchAll trues, nOp_ + 1
            Exit Sub
        End If
        If CompareJump(op) <> 0 Then
            EmitExpr cond.GetValue("left")
            EmitExpr cond.GetValue("right")
            Emit CompareJump(op), 0, 0, -2
            jumps.Add nOp_
            Exit Sub
        End If
    ElseIf t = "Unary" Then
        If cond.GetValue("op") = "!" Then
            EmitJumpIfTrue cond.GetValue("expr"), jumps
            Exit Sub
        End If
    End If
    EmitExpr cond
    Emit OP_JMPF, 0, 0, -1
    jumps.Add nOp_
End Sub

Private Sub EmitJumpIfTrue(cond As ASF_Map, jumps As Collection)
    Dim t As String: t = cond.GetValue("type")
    If t = "Binary" Then
        Dim op As String: op = cond.GetValue("op")
        If op = "||" Then
            EmitJumpIfTrue cond.GetValue("left"), jumps
            EmitJumpIfTrue cond.GetValue("right"), jumps
            Exit Sub
        End If
        If op = "&&" Then
            Dim falses As Collection: Set falses = New Collection
            EmitJumpIfFalse cond.GetValue("left"), falses
            EmitJumpIfTrue cond.GetValue("right"), jumps
            PatchAll falses, nOp_ + 1
            Exit Sub
        End If
    ElseIf t = "Unary" Then
        If cond.GetValue("op") = "!" Then
            EmitJumpIfFalse cond.GetValue("expr"), jumps
            Exit Sub
        End If
    End If
    EmitExpr cond
    Emit OP_JMPT, 0, 0, -1
    jumps.Add nOp_
End Sub

' ---------------------------------------------------------------- expressions
Private Sub EmitExpr(node As ASF_Map)
    If node Is Nothing Then
        Emit OP_PUSHE, 0, 0, 1
        Exit Sub
    End If
    Dim k As Variant
    Select Case node.GetValue("type")
        Case "Variable": EmitLoadName node.GetValue("name")
        Case "Literal"
            k = node.GetValue("value")
            If IsObject(node.GetValue("value")) Then NotSupported "object literal value"
            Emit OP_PUSHK, Constant(k), 0, 1
        Case "Binary": EmitBinary node
        Case "Unary": EmitUnary node
        Case "Ternary"
            Dim falses As Collection: Set falses = New Collection
            EmitJumpIfFalse node.GetValue("cond"), falses
            EmitExpr node.GetValue("trueExpr")
            Emit OP_JMP, 0, 0, -1
            Dim endJump As Long: endJump = nOp_
            PatchAll falses, nOp_ + 1
            EmitExpr node.GetValue("falseExpr")
            Patch endJump, nOp_ + 1
        Case "Call": EmitCall node
        Case "Array": EmitArray node
        Case "Object": EmitObject node
        Case "Template": EmitTemplate node
        Case "Index"
            EmitExpr node.GetValue("base")
            EmitExpr node.GetValue("index")
            Emit OP_GETIDX, 0, 0, -1
        Case "Member"
            If Not node.Exists("prop") Then NotSupported "member without a name"
            EmitExpr node.GetValue("base")
            Emit OP_GETPROP, Constant(CStr(node.GetValue("prop"))), 0, 0
        Case "SuperCall": EmitSuperCall node
        Case "VBAexpr": EmitVBAExpr CStr(node.GetValue("expr"))
        Case "FuncLiteral"
            Dim ctx As ASF_Map: Set ctx = LiteralContext(node)
            If Not ctx.Exists("fn") Then
                ctx.SetValue "fn", NewFunction("@literal")
                pendingLits_.Add ctx
            End If
            Emit OP_MKCLOS, ctx.GetValue("fn"), 0, 1
        Case Else
            NotSupported "expression " & node.GetValue("type")
    End Select
End Sub

Private Sub EmitArray(node As ASF_Map)
    Dim items As Collection: Set items = node.GetValue("items")
    Dim i As Long, n As Long, flags As String, spread As Boolean
    If Not items Is Nothing Then n = items.Count
    For i = 1 To n
        If items(i).GetValue("type") = "SpreadElement" Then
            EmitExpr items(i).GetValue("argument")
            flags = flags & "1"
            spread = True
        Else
            EmitExpr items(i)
            flags = flags & "0"
        End If
    Next i
    If spread Then
        Emit OP_NEWARRS, n, Constant(flags), 1 - n
    Else
        Emit OP_NEWARR, n, 0, 1 - n
    End If
End Sub

' A literal part of a template with text in it (${null} is a literal too, without text)
Private Function IsTextPart(ByVal part As ASF_Map) As Boolean
    If part.GetValue("type") = "Literal" Then
        If Not IsObject(part.GetValue("value")) Then IsTextPart = (VarType(part.GetValue("value")) = vbString)
    End If
End Function

Private Sub EmitTemplate(node As ASF_Map)
    Dim parts As Collection: Set parts = node.GetValue("parts")
    Dim part As Variant, n As Long, text As String, allText As Boolean, partNode As ASF_Map
    allText = True
    For Each part In parts
        If Not part Is Nothing Then
            If IsTextPart(part) Then
                text = text & CStr(part.GetValue("value"))
            Else
                allText = False
            End If
        End If
    Next part
    If allText Then
        Emit OP_PUSHK, Constant(text), 0, 1
        Exit Sub
    End If
    For Each part In parts
        If Not part Is Nothing Then
            If IsTextPart(part) Then
                Emit OP_PUSHK, Constant(CStr(part.GetValue("value"))), 0, 1
            Else
                Set partNode = part
                EmitExpr partNode
            End If
            n = n + 1
        End If
    Next part
    Emit OP_TPL, n, 0, 1 - n
End Sub

Private Sub EmitObject(node As ASF_Map)
    Dim items As Collection: Set items = node.GetValue("items")
    Dim i As Long, n As Long, pair As Collection
    If Not items Is Nothing Then n = items.Count
    For i = 1 To n
        Set pair = items(i)
        Emit OP_PUSHK, Constant(CStr(pair(1))), 0, 1
        If CStr(pair(1)) = "__SPREAD__" Then
            EmitExpr pair(2).GetValue("argument")
        Else
            EmitExpr pair(2)
        End If
    Next i
    Emit OP_NEWOBJ, n, 0, 1 - 2 * n
End Sub

' Where a name lives for the function being compiled: "L" local, "C" local in a cell, "E" entry
' of the environment, "G" global, "F" named function; slot gets its number
Private Function NameKind(ByVal name As String, ByRef slot As Long) As String
    If cx_.GetValue("kind") <> "main" Then
        Dim locals As ASF_Map: Set locals = cx_.GetValue("locals")
        If locals.Exists(name) Then
            slot = locals.GetValue(name)
            Dim celled As ASF_Map: Set celled = cx_.GetValue("celled")
            If celled.Exists(name) Then NameKind = "C" Else NameKind = "L"
            Exit Function
        End If
        Dim free As ASF_Map: Set free = cx_.GetValue("free")
        If free.Exists(name) Then
            slot = free.GetValue(name)
            NameKind = "E"
            Exit Function
        End If
    End If
    If Not gSlot_.Exists(name) Then
        If GLOBALS_.gFuncTable.Exists(name) Then
            slot = NamedFunction(name)
            NameKind = "F"
            Exit Function
        End If
    End If
    slot = GlobalSlot(name)
    NameKind = "G"
End Function

Private Sub EmitLoadName(ByVal name As String)
    Dim slot As Long
    Select Case NameKind(name, slot)
        Case "L": Emit OP_LOADL, slot, 0, 1
        Case "C": Emit OP_LOADC, slot, 0, 1
        Case "E": Emit OP_LOADE, slot, 0, 1
        Case "F": Emit OP_MKFUNC, slot, 0, 1
        Case Else: Emit OP_LOADG, slot, 0, 1
    End Select
End Sub

' Stores the value on top of the stack in a variable
Private Sub EmitStoreName(ByVal name As String)
    Dim slot As Long
    Select Case NameKind(name, slot)
        Case "L": Emit OP_STOREL, slot, 0, -1
        Case "C": Emit OP_STOREC, slot, 0, -1
        Case "E": Emit OP_STOREE, slot, 0, -1
        Case "F"
            slot = GlobalSlot(name)
            Emit OP_STOREG, slot, 0, -1
        Case Else: Emit OP_STOREG, slot, 0, -1
    End Select
End Sub

' Stores the value on top of the stack in a target: a variable, obj.prop or c[i]; an array that
' changes is written back to where it came from, as arrays are values
Private Sub EmitStoreTop(target As ASF_Map)
    Select Case target.GetValue("type")
        Case "Variable"
            EmitStoreName target.GetValue("name")
        Case "Member"
            EmitExpr target.GetValue("base")
            Emit OP_SETPROPV, Constant(CStr(target.GetValue("prop"))), 0, -2
        Case "Index"
            If Not IsSimple(target.GetValue("index")) Then NotSupported "an index with side effects in an assignment"
            EmitContainerForWrite target.GetValue("base")
            EmitExpr target.GetValue("index")
            Emit OP_SETIDXR, 0, 0, -2
            EmitStoreTop target.GetValue("base")
        Case Else
            NotSupported "assignment to " & target.GetValue("type")
    End Select
End Sub

' The container an indexed assignment writes into: missing inner arrays are not errors
Private Sub EmitContainerForWrite(node As ASF_Map)
    If node.GetValue("type") = "Index" Then
        If Not IsSimple(node.GetValue("index")) Then NotSupported "an index with side effects in an assignment"
        EmitContainerForWrite node.GetValue("base")
        EmitExpr node.GetValue("index")
        Emit OP_GETIDXW, 0, 0, -1
    Else
        EmitExpr node
    End If
End Sub

' An expression without calls: evaluating it twice gives the same value
Private Function IsSimple(node As ASF_Map) As Boolean
    If node Is Nothing Then
        IsSimple = True
        Exit Function
    End If
    Select Case node.GetValue("type")
        Case "Literal", "Variable": IsSimple = True
        Case "Binary": IsSimple = IsSimple(node.GetValue("left")) And IsSimple(node.GetValue("right"))
        Case "Unary": IsSimple = IsSimple(node.GetValue("expr"))
        Case "Member": IsSimple = IsSimple(node.GetValue("base"))
        Case "Index": IsSimple = IsSimple(node.GetValue("base")) And IsSimple(node.GetValue("index"))
    End Select
End Function

Private Function BinaryOp(ByVal op As String) As Long
    Select Case op
        Case "+": BinaryOp = OP_ADD
        Case "-": BinaryOp = OP_SUB
        Case "*": BinaryOp = OP_MUL
        Case "/": BinaryOp = OP_DIV
        Case "%": BinaryOp = OP_MOD
        Case "^": BinaryOp = OP_POW
        Case "==", "=": BinaryOp = OP_EQ
        Case "!=": BinaryOp = OP_NE
        Case "<": BinaryOp = OP_LT
        Case "<=": BinaryOp = OP_LE
        Case ">": BinaryOp = OP_GT
        Case ">=": BinaryOp = OP_GE
        Case "&": BinaryOp = OP_CAT
        Case "<<": BinaryOp = OP_SHL
        Case ">>": BinaryOp = OP_SHR
    End Select
End Function

Private Sub EmitBinary(node As ASF_Map)
    Dim op As String: op = node.GetValue("op")
    Dim k As Variant, endJump As Long
    If FoldConstant(node, k) Then
        Emit OP_PUSHK, Constant(k), 0, 1
        Exit Sub
    End If
    Select Case op
        Case "&&", "||"
            EmitExpr node.GetValue("left")
            If op = "&&" Then Emit OP_ANDJ, 0, 0, -1 Else Emit OP_ORJ, 0, 0, -1
            endJump = nOp_
            EmitExpr node.GetValue("right")
            Emit OP_TOBOOL, 0, 0, 0
            Patch endJump, nOp_ + 1
            Exit Sub
        Case "??"
            EmitExpr node.GetValue("left")
            Emit OP_NULLJ, 0, 0, -1
            endJump = nOp_
            EmitExpr node.GetValue("right")
            Patch endJump, nOp_ + 1
            Exit Sub
    End Select
    Dim code As Long: code = BinaryOp(op)
    If code = 0 Then NotSupported "operator " & op
    EmitExpr node.GetValue("left")
    If code = OP_ADD Then
        If FoldNumber(node.GetValue("right"), k) Then
            Emit OP_ADDK, Constant(k), 0, 0
            Exit Sub
        End If
    End If
    EmitExpr node.GetValue("right")
    Emit code, 0, 0, -1
End Sub

Private Sub EmitUnary(node As ASF_Map)
    Dim k As Variant
    If FoldConstant(node, k) Then
        Emit OP_PUSHK, Constant(k), 0, 1
        Exit Sub
    End If
    EmitExpr node.GetValue("expr")
    Select Case node.GetValue("op")
        Case "!": Emit OP_NOT, 0, 0, 0
        Case "-": Emit OP_NEG, 0, 0, 0
        Case "typeof": Emit OP_TYPEOF, 0, 0, 0
        Case Else: NotSupported "unary " & node.GetValue("op")
    End Select
End Sub

Private Sub EmitCall(node As ASF_Map)
    If Not node.Exists("callee") Then NotSupported "call by name"
    Dim callee As ASF_Map: Set callee = node.GetValue("callee")
    If callee Is Nothing Then NotSupported "call without a callee"
    Dim args As Collection: Set args = node.GetValue("args")
    Dim n As Long, slot As Long, kind As String, name As String
    If Not args Is Nothing Then n = args.Count
    Select Case callee.GetValue("type")
        Case "Variable"
            name = callee.GetValue("name")
            kind = NameKind(name, slot)
            If kind = "F" Then
                ' a named function no variable hides: a direct call
                If EmitArgs(args) Then Emit OP_CALLFV, slot, 0, -n Else Emit OP_CALLF, slot, n, 1 - n
                Exit Sub
            End If
            If kind = "G" And Not GlobalAssigned(name) Then
                Select Case LCase$(name)
                    Case "range", "flatten", "clone", "isarray", "isnumeric", "foreach", "regex", "cwd", "scwd"
                        If EmitArgs(args) Then
                            Emit OP_CALLBV, Constant(name), 0, -n
                        Else
                            Emit OP_CALLB, Constant(name), n, 1 - n
                        End If
                        Exit Sub
                End Select
            End If
            EmitLoadName name
        Case "Member"
            EmitMethodCall callee, args
            Exit Sub
        Case "Super", "SuperCall"
            NotSupported "call of " & callee.GetValue("type")
        Case Else
            EmitExpr callee
    End Select
    If EmitArgs(args) Then Emit OP_CALLV, 0, 0, -n - 1 Else Emit OP_CALL, n, 0, -n
End Sub

' receiver.name(args): methods that change an array (push, sort...) write it back to the receiver
Private Sub EmitMethodCall(callee As ASF_Map, args As Collection)
    If Not callee.Exists("prop") Then NotSupported "method without a name"
    Dim recv As ASF_Map: Set recv = callee.GetValue("base")
    Dim name As String: name = callee.GetValue("prop")
    Dim n As Long
    If Not args Is Nothing Then n = args.Count
    Dim writes As Boolean
    Select Case LCase$(name)
        Case "push", "pop", "shift", "unshift", "splice", "copywithin", "reverse", "sort", "tosorted", "concat", "delete"
            writes = IsWritable(recv)
    End Select
    EmitExpr recv
    Dim spread As Boolean: spread = EmitArgs(args)
    ' b: the argument count, and the number of the method when the VM has it built in
    Dim id As Long: id = METHOD_SHIFT * MethodId(name)
    If writes Then
        If spread Then Emit OP_CALLMWV, Constant(name), id, -n Else Emit OP_CALLMW, Constant(name), n + id, 1 - n
        EmitStoreTop recv
    Else
        If spread Then Emit OP_CALLMV, Constant(name), id, -n - 1 Else Emit OP_CALLM, Constant(name), n + id, -n
    End If
End Sub

' The methods ASF_FastVM runs itself; 0 for the others, which run on the tree VM
Private Function MethodId(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "push": MethodId = 1
        Case "pop": MethodId = 2
        Case "map": MethodId = 3
        Case "filter": MethodId = 4
        Case "reduce": MethodId = 5
        Case "foreach": MethodId = 6
        Case "touppercase": MethodId = 10
        Case "tolowercase": MethodId = 11
        Case "trim": MethodId = 12
        Case "max": MethodId = 20
        Case "min": MethodId = 21
        Case "floor": MethodId = 22
        Case "ceil": MethodId = 23
        Case "abs": MethodId = 24
        Case "round": MethodId = 25
        Case "sqrt": MethodId = 26
        Case "pow": MethodId = 27
    End Select
End Function

' A target EmitStoreTop can write: a variable, a member, or an index without side effects
Private Function IsWritable(node As ASF_Map) As Boolean
    Select Case node.GetValue("type")
        Case "Variable": IsWritable = True
        Case "Member": IsWritable = node.Exists("prop")
        Case "Index": IsWritable = IsSimple(node.GetValue("index")) And IsWritable(node.GetValue("base"))
    End Select
End Function

' The arguments of a call; True when one is spread: the items and their count are then on the stack
Private Function EmitArgs(args As Collection) As Boolean
    Dim i As Long, flags As String
    If args Is Nothing Then Exit Function
    For i = 1 To args.Count
        If args(i) Is Nothing Then
            Emit OP_PUSHE, 0, 0, 1
            flags = flags & "0"
        ElseIf args(i).GetValue("type") = "SpreadElement" Then
            EmitExpr args(i).GetValue("argument")
            flags = flags & "1"
            EmitArgs = True
        Else
            EmitExpr args(i)
            flags = flags & "0"
        End If
    Next i
    If EmitArgs Then Emit OP_EXPAND, args.Count, Constant(flags), 1
End Function

' True when the program assigns name at its top level (a global it owns, not an injected one)
Private Function GlobalAssigned(ByVal name As String) As Boolean
    GlobalAssigned = assignedGlobals_.Exists(name)
End Function

' ---------------------------------------------------------------- constant folding
' A number literal, or an expression of number literals
Private Function FoldNumber(node As ASF_Map, ByRef k As Variant) As Boolean
    If node Is Nothing Then Exit Function
    If FoldConstant(node, k) Then FoldNumber = (VarType(k) = vbDouble)
End Function

Private Function FoldConstant(node As ASF_Map, ByRef k As Variant) As Boolean
    If node Is Nothing Then Exit Function
    Dim l As Variant, r As Variant
    Select Case node.GetValue("type")
        Case "Literal"
            If IsObject(node.GetValue("value")) Then Exit Function
            k = node.GetValue("value")
            FoldConstant = (VarType(k) = vbDouble Or VarType(k) = vbString)
        Case "Unary"
            If node.GetValue("op") = "-" Then
                If FoldConstant(node.GetValue("expr"), l) Then
                    If VarType(l) = vbDouble Then k = -l: FoldConstant = True
                End If
            End If
        Case "Binary"
            If Not FoldConstant(node.GetValue("left"), l) Then Exit Function
            If Not FoldConstant(node.GetValue("right"), r) Then Exit Function
            If VarType(l) = vbDouble And VarType(r) = vbDouble Then
                Select Case node.GetValue("op")
                    Case "+": k = l + r: FoldConstant = True
                    Case "-": k = l - r: FoldConstant = True
                    Case "*": k = l * r: FoldConstant = True
                    Case "/"
                        If r <> 0 Then k = l / r: FoldConstant = True
                End Select
            ElseIf VarType(l) = vbString And VarType(r) = vbString Then
                If node.GetValue("op") = "+" Then k = l & r: FoldConstant = True
            End If
    End Select
End Function
