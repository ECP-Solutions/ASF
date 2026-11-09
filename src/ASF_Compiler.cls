VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Compiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ASF_Compiler: compiles a program, and the named functions it calls, from its tokens to the
' bytecode that ASF_VM runs (see BYTECODE.md). A source with an error raises it.
Option Explicit

' opcodes: the numbers are the jump table of ASF_VM.Execute, keep both lists in step
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
Private Const OP_UNUSED83 As Long = 83
Private Const OP_VBAEXPR As Long = 84   ' a: the VBA expression, b: the names of the values under it, one per line
Private Const OP_PROTO As Long = 85     ' a: the prototype (classType, methodName, funcIndex, params); pops its function
Private Const OP_EXPORT As Long = 86    ' a: exported name; pops the value (a module being loaded keeps it)
Private Const OP_EXPORTPROTO As Long = 87 ' a: the prototype, with exportName; pops its function
Private Const OP_SUPER As Long = 88     ' a: 0 init, 1 constructor of the class under this and b arguments (-1: an array of them)
Private Const METHOD_SHIFT As Long = 1024 ' b of CALLM...: arguments + METHOD_SHIFT * method number

' compile errors: vbObjectError + these numbers, as ASF 3 numbered them
Private Const ERR_TERNARY As Long = 8001
Private Const ERR_OBJECT_COLON As Long = 8002
Private Const ERR_OBJECT_KEY As Long = 8003
Private Const ERR_OBJECT_SEMICOLON As Long = 8004
Private Const ERR_EXPORT As Long = 8013
Private Const ERR_REST_NAME As Long = 9001
Private Const ERR_REST_LAST As Long = 9002
Private Const ERR_REST_TWICE As Long = 9003
Private Const ERR_PARENTHESES As Long = 9024
Private Const ERR_SYNTAX As Long = 9025

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
Private clsInit_() As Long, clsCtor_() As Long          ' class of a module -> its fields and methods, its constructor alone
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

' the direct compiler: the session's tokens (tt_ kinds, tv_ values), the parser's position p_ and
' limit lim_, and the stack effect dl_ of each instruction (adj_: the catch value before a handler)
Private tt_() As String, tv_() As Variant, nT_ As Long
Private p_ As Long, lim_ As Long
Private mode_ As Long
Private scanUsed_ As ASF_Map, scanCtx_ As ASF_Map, scanLits_ As Boolean
Private raw_ As Boolean                 ' a class body: = in a for header is a comparison, as in ASF 3
Private dFuncs_ As ASF_Map              ' the named functions of the source being compiled
Private dFuncOrder_ As Collection
Private dParser_ As ASF_Parser
Private dPendingGlobals_ As Variant
Private directFn_ As Boolean
Private dl_() As Long, adj_() As Long, pendingAdj_ As Long
Private kOp_() As Long                  ' constant -> the instruction that added it
Private tplOp_ As Long                  ' the PUSHK of a template without ${...}
Private tplParts_ As ASF_Map
Private idxTemps_ As ASF_Map            ' the hidden variables of the indexes of the target being written
Private blockC_() As Variant            ' token -> Array(upTo, statements) of the block read from there
Private dNamed_ As Long                 ' the fun declarations of the source

Private Sub Class_Initialize()
    ReDim op_(1 To 256): ReDim a_(1 To 256): ReDim b_(1 To 256)
    ReDim dl_(1 To 256): ReDim adj_(1 To 256)
    ReDim k_(1 To 64): ReDim kOp_(1 To 64)
    ReDim tt_(1 To 256): ReDim tv_(1 To 256)
    ReDim blockC_(1 To 256)
    Set tplParts_ = New ASF_Map
    ReDim fnEntry_(1 To 16): ReDim fnNParams_(1 To 16): ReDim fnNLocals_(1 To 16)
    ReDim fnMaxStack_(1 To 16): ReDim fnName_(1 To 16)
    ReDim fnCellStart_(1 To 16): ReDim fnCellCount_(1 To 16)
    ReDim fnCapStart_(1 To 16): ReDim fnCapCount_(1 To 16)
    ReDim fnRest_(1 To 16): ReDim fnThis_(1 To 16): ReDim fnThisEnv_(1 To 16)
    ReDim cellSlot_(1 To 16): ReDim capKind_(1 To 16): ReDim capIdx_(1 To 16)
    ReDim fnKind_(1 To 16): ReDim tryStart_(1 To 4): ReDim tryEnd_(1 To 4): ReDim tryHandler_(1 To 4): ReDim tryDepth_(1 To 4): ReDim clsNew_(1 To 4): ReDim clsInit_(1 To 4): ReDim clsCtor_(1 To 4)
    Set kMap_ = New ASF_Map
    Set progFn_ = New ASF_Map
    Set progReason_ = New ASF_Map
    Set progGlobals_ = New ASF_Map
    Set progGBase_ = New ASF_Map
End Sub

Public Sub SetGlobals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Sub

' ---------------------------------------------------------------- what ASF_VM reads
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
' Class number -> the function that sets the fields and binds the methods of an instance (a class
' of a module, which a class of another script may extend), 0 for the others
Public Function ClassInits() As Variant
    ClassInits = clsInit_
End Function
' Class number -> its constructor alone (a class of a module), 0 for the others
Public Function ClassCtors() As Variant
    ClassCtors = clsCtor_
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
                  "NEWARRS JNSW ITERLIST FORNEXT UNPACK CLASS NEWINST BINDM CALLFT THROW IMPORT TREESTMT VBAEXPR PROTO EXPORT EXPORTPROTO SUPER", " ")
    Dim out As String, i As Long, fn As Long, k As Variant
    If Not Prepare(idx) Then
        Disassemble = "not taken: " & reason_
        Exit Function
    End If
    For fn = 1 To nFn_
        out = out & "fn " & fn & " " & fnName_(fn) & " entry " & fnEntry_(fn) & " params " & fnNParams_(fn) & _
              " locals " & fnNLocals_(fn) & " stack " & fnMaxStack_(fn) & " cells " & fnCellCount_(fn) & _
              " captures " & fnCapCount_(fn) & " rest " & fnRest_(fn) & " this " & fnThis_(fn) & _
              " thisenv " & fnThisEnv_(fn) & " kind " & fnKind_(fn) & vbLf
        For i = fnCellStart_(fn) To fnCellStart_(fn) + fnCellCount_(fn) - 1
            out = out & "  cell " & cellSlot_(i) & vbLf
        Next i
        For i = fnCapStart_(fn) To fnCapStart_(fn) + fnCapCount_(fn) - 1
            out = out & "  capture " & capKind_(i) & " " & capIdx_(i) & vbLf
        Next i
    Next fn
    For i = 1 To nTry_
        out = out & "try " & tryStart_(i) & " " & tryEnd_(i) & " " & tryHandler_(i) & " " & tryDepth_(i) & vbLf
    Next i
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
' True when program idx was compiled
Public Function Prepare(ByVal idx As Long) As Boolean
    Prepare = progFn_.Exists(CStr(idx))
    If Not Prepare Then reason_ = "no program " & idx
End Function

' A compile error: the source does not compile, and nothing of it is kept
Private Sub CompileError(ByVal message As String, Optional ByVal number As Long = ERR_SYNTAX)
    Err.Raise vbObjectError + number, "ASF_Compiler", message
End Sub

' A named function the code uses: from the tokens of the source being compiled, or else from the
' ones gPrograms keeps for it, which are added to the session's
Private Sub CompileNamedFunction(ByVal name As String)
    If Not dFuncs_ Is Nothing Then
        If dFuncs_.Exists(name) Then
            Dim info As ASF_Map: Set info = dFuncs_.GetValue(name)
            DCompileNamed name, info.GetValue("paramArray"), CStr(info.GetValue("rest")), info.GetValue("bodyS"), info.GetValue("bodyE")
            Exit Sub
        End If
    End If
    Dim p As Variant: p = GLOBALS_.gPrograms(GLOBALS_.gFuncTable.GetValue(name))
    Dim restName As String, params As Variant, from As Long, upTo As Long
    If Not IsEmpty(p(3)) Then restName = CStr(p(3))
    AppendStored p(4), p(5), from, upTo
    If Not GLOBALS_.gFuncParams.TryGetValue(name, params) Then params = Empty
    DCompileNamed name, params, restName, from, upTo
End Sub

' The function of a method declaration, as DEmitClassDecl made it
Private Function MethodFunctionOf(node As ASF_Map) As Long
    Dim pair As Variant
    For Each pair In methodFns_
        If pair(0) Is node Then
            MethodFunctionOf = pair(1)
            Exit Function
        End If
    Next pair
    CompileError "a method of a class not declared yet"
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
    directFn_ = False
    Set breaks_ = New Collection
    Set conts_ = New Collection
End Sub

' The frame of fn: its locals, the ones boxed in cells, and what its closures capture
Private Sub EndFunction(ByVal fn As Long, ByVal nParams As Long)
    fnNParams_(fn) = nParams
    fnNLocals_(fn) = cx_.GetValue("nL")
    If directFn_ Then fnMaxStack_(fn) = DMaxDepth(fnEntry_(fn)) + 2 Else fnMaxStack_(fn) = maxDepth_ + 2
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

' ---------------------------------------------------------------- emitting
Private Sub Emit(ByVal op As Long, ByVal a As Long, ByVal b As Long, ByVal stackDelta As Long)
    nOp_ = nOp_ + 1
    If nOp_ > UBound(op_) Then
        ReDim Preserve op_(1 To 2 * nOp_): ReDim Preserve a_(1 To 2 * nOp_): ReDim Preserve b_(1 To 2 * nOp_)
        ReDim Preserve dl_(1 To 2 * nOp_): ReDim Preserve adj_(1 To 2 * nOp_)
    End If
    op_(nOp_) = op: a_(nOp_) = a: b_(nOp_) = b
    dl_(nOp_) = stackDelta: adj_(nOp_) = pendingAdj_: pendingAdj_ = 0
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

' An object in the constant table (a declaration the VM makes values of)
Private Function ObjectConstant(ByVal obj As Object) As Long
    nK_ = nK_ + 1
    If nK_ > UBound(k_) Then ReDim Preserve k_(1 To 2 * nK_): ReDim Preserve kOp_(1 To 2 * nK_)
    Set k_(nK_) = obj
    kOp_(nK_) = nOp_ + 1
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
    If nK_ > UBound(k_) Then ReDim Preserve k_(1 To 2 * nK_): ReDim Preserve kOp_(1 To 2 * nK_)
    k_(nK_) = v
    kOp_(nK_) = nOp_ + 1
    kMap_.SetValue key, nK_
    Constant = nK_
End Function

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

' @(expr): ASF_Runtime evaluates it with the variables in scope it names (globals first, so that
' locals win)
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

' A loop's lists of break and continue jumps
Private Sub OpenLoop()
    Dim jumps As Collection
    Set jumps = New Collection
    breaks_.Add jumps
    Set jumps = New Collection
    conts_.Add jumps
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
        If GLOBALS_.gFuncTable.Exists(name) Or DirectFunction(name) Then
            slot = NamedFunction(name)
            NameKind = "F"
            Exit Function
        End If
    End If
    slot = GlobalSlot(name)
    NameKind = "G"
End Function

Private Function DirectFunction(ByVal name As String) As Boolean
    If Not dFuncs_ Is Nothing Then DirectFunction = dFuncs_.Exists(name)
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

' The methods ASF_VM runs itself; 0 for the others, which ASF_Runtime runs
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

' True when the program assigns name at its top level (a global it owns, not an injected one)
Private Function GlobalAssigned(ByVal name As String) As Boolean
    GlobalAssigned = assignedGlobals_.Exists(name)
End Function

' ================================================================ compiling a source
' A program is compiled from its tokens: statements become plans of token ranges
' (Array(kind, ...)); expressions are parsed and emitted in one pass. The same parser also scans
' (mode_ 1: the names and function literals of a body, for the scope analysis) and probes (mode_ 2:
' where an expression ends).

' Compiles src: its program index. A source with an error raises it, and leaves nothing behind
Public Function CompileSource(ByVal src As String, ByVal progName As String) As Long
    If GLOBALS_.ProgramCache.Exists(src) Then
        CompileSource = GLOBALS_.ProgramCache.GetValue(src)
        Exit Function
    End If
    If Not CheckParents(src) Then CompileError "Unbalanced opening or closing parentheses.", ERR_PARENTHESES
    Dim saveOp As Long, saveK As Long, saveFn As Long, saveCell As Long, saveCap As Long
    Dim saveCls As Long, saveTry As Long, saveG As Long
    saveOp = nOp_: saveK = nK_: saveFn = nFn_: saveCell = nCellSlot_: saveCap = nCap_
    saveCls = nCls_: saveTry = nTry_: saveG = nGTotal_
    On Error GoTo Failed
    If dParser_ Is Nothing Then
        Set dParser_ = New ASF_Parser
        dParser_.SetGlobals GLOBALS_
    End If
    ReDim tt_(1 To 256): ReDim tv_(1 To 256): nT_ = 0
    AppendTokenList dParser_.Tokenize(src)
    Set dFuncs_ = New ASF_Map
    Set dFuncOrder_ = New Collection
    Dim mainFn As Long: mainFn = DCompileProgram(progName)
    ' named functions first, then the program
    Dim i As Long, info As ASF_Map, idx As Long
    For i = 1 To dFuncOrder_.Count
        Set info = dFuncs_.GetValue(CStr(dFuncOrder_(i)))
        GLOBALS_.gPrograms.Add Array(CStr(dFuncOrder_(i)), Nothing, Nothing, info.GetValue("rest"), info.GetValue("types"), info.GetValue("vals"))
        GLOBALS_.gFuncTable.Add CStr(dFuncOrder_(i)), GLOBALS_.gPrograms.Count
        GLOBALS_.gFuncParams.Add CStr(dFuncOrder_(i)), info.GetValue("paramArray")
    Next i
    GLOBALS_.gPrograms.Add Array(progName, Nothing, Nothing, 1, src)
    idx = GLOBALS_.gPrograms.Count
    GLOBALS_.ProgramCache.Add src, idx
    progFn_.SetValue CStr(idx), mainFn
    progGlobals_.SetValue CStr(idx), dPendingGlobals_
    progGBase_.SetValue CStr(idx), gBase_
    reason_ = vbNullString
    version_ = version_ + 1
    CompileSource = idx
    Set dFuncs_ = Nothing
    Exit Function
Failed:
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    nOp_ = saveOp: nFn_ = saveFn: nCellSlot_ = saveCell: nCap_ = saveCap
    nCls_ = saveCls: nTry_ = saveTry: nGTotal_ = saveG
    RestoreConstants saveK
    Set dFuncs_ = Nothing
    mode_ = 0: raw_ = False
    reason_ = errDesc
    On Error GoTo 0
    Err.Raise errNum, errSrc, errDesc
End Function

' Drops the constants added after saveK
Private Sub RestoreConstants(ByVal saveK As Long)
    If nK_ = saveK Then Exit Sub
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
End Sub

' ---------------------------------------------------------------- tokens
' Brackets of the code: those in strings and comments are text, as in return 'a)';
Private Function CheckParents(src As String) As Boolean
    Dim i As Long, n As Long
    Dim openPar As Long
    Dim ch As String, quote As String
    n = Len(src)
    i = 1
    Do While i <= n
        ch = mid$(src, i, 1)
        If Len(quote) > 0 Then
            If ch = "\" Then
                i = i + 1
            ElseIf ch = quote Then
                quote = vbNullString
            End If
        Else
            Select Case ch
                Case "'", """", "`"
                    quote = ch
                Case "#"
                    i = LineEnd(src, i)
                Case "/"
                    If mid$(src, i + 1, 1) = "/" Then
                        i = LineEnd(src, i)
                    ElseIf mid$(src, i + 1, 1) = "*" Then
                        i = InStr(i + 2, src, "*/")
                        If i = 0 Then Exit Do
                        i = i + 1
                    End If
                Case "[", "{", "("
                    openPar = openPar + 1
                Case "]", "}", ")"
                    openPar = openPar - 1
            End Select
        End If
        i = i + 1
    Loop
    CheckParents = Not CBool(openPar And 1)
End Function
' Position of the last character of the line that holds position i
Private Function LineEnd(src As String, ByVal i As Long) As Long
    Do While i < Len(src)
        Select Case mid$(src, i + 1, 1)
            Case vbCr, vbLf: Exit Do
        End Select
        i = i + 1
    Loop
    LineEnd = i
End Function
' Input: tokenStr is the raw token text starting with backtick, e.g. "`/${regobj.escape(name)}/g`"
' Output: Collection of Array("LITERAL", text) or Array("EXPR", text)
Private Function ParseTemplatePartsFromString(ByVal tokenStr As String) As Collection
    Dim parts As New Collection
    If Len(tokenStr) = 0 Then
        Set ParseTemplatePartsFromString = parts
        Exit Function
    End If
    If mid$(tokenStr, 1, 1) <> "`" Then
        parts.Add Array("LITERAL", tokenStr)
        Set ParseTemplatePartsFromString = parts: Exit Function
    End If

    Dim i As Long: i = 2 ' start after opening backtick
    Dim Ls As Long: Ls = Len(tokenStr)
    Dim curLit As String: curLit = vbNullString
    Dim ch As String
    Dim nx As String
    Dim curExpr As String
    Dim depth As Long
    Dim c2 As String
    Dim escapedChar As String
    Dim braces As Long, quote As String
    
    Do While i <= Ls
        ch = mid$(tokenStr, i, 1)

        ' Backslash escapes in template literal (outside ${}): unescape `, /, \, $, {, }
        If ch = "\" Then
            nx = mid$(tokenStr, i + 1, 1)
            Select Case nx
                Case "`", "/", "\", "$", "{", "}"
                    curLit = curLit & nx
                    i = i + 2
                Case Else
                    ' keep the backslash as-is if it doesn't precede a recognized escape
                    curLit = curLit & ch
                    i = i + 1
            End Select
            GoTo ContinueMainLoop
        End If

        ' Placeholder start ${ ... }
        If ch = "$" And i < Ls And mid$(tokenStr, i + 1, 1) = "{" Then
            ' flush current literal (outside placeholders)
            If Len(curLit) > 0 Then
                parts.Add Array("LITERAL", curLit)
                curLit = vbNullString
            End If

            ' consume "${"
            i = i + 2

            ' Collect placeholder content with escaping and nested-${} support.
            ' semantics:
            '  - backslash inside placeholder escapes the next char if equal to '{' or '$' or '}'
            '  - unescaped "${" starts a nested split: flush current expr (if any) and start collecting nested
            '  - unescaped "}" closes the *current* placeholder chunk and the collected text is added as an EXPR
            curExpr = vbNullString
            depth = 1
            braces = 0: quote = vbNullString

            Do While i <= Ls And depth > 0
                c2 = mid$(tokenStr, i, 1)

                ' a string inside the placeholder is taken whole: ${'}'}
                If Len(quote) > 0 Then
                    If c2 = "\" And i < Ls Then
                        curExpr = curExpr & mid$(tokenStr, i, 2)
                        i = i + 2
                    Else
                        If c2 = quote Then quote = vbNullString
                        curExpr = curExpr & c2
                        i = i + 1
                    End If
                    GoTo ContinuePlaceholderLoop
                ElseIf c2 = "'" Or c2 = """" Then
                    quote = c2
                    curExpr = curExpr & c2
                    i = i + 1
                    GoTo ContinuePlaceholderLoop
                End If

                ' escaped char inside placeholder -> take next char literally
                If c2 = "\" Then
                    If i < Ls Then
                        escapedChar = mid$(tokenStr, i + 1, 1)
                        ' append the escaped character verbatim into the current expr buffer
                        Select Case escapedChar
                            Case "$", "{", "}"
                                curExpr = curExpr & escapedChar
                                i = i + 2
                            Case Else
                                'literal "\"
                                curExpr = curExpr & c2
                                i = i + 1
                        End Select
                    Else
                        ' trailing backslash => keep it
                        curExpr = curExpr & c2
                        i = i + 1
                    End If
                    GoTo ContinuePlaceholderLoop
                End If

                ' nested unescaped ${ -> split: flush current expr (if any) and continue with a fresh expr for nested part
                If c2 = "$" And i < Ls And mid$(tokenStr, i + 1, 1) = "{" Then
                    If Len(curExpr) > 0 Then
                        parts.Add Array("EXPR", curExpr)
                        curExpr = vbNullString
                    End If
                    ' consume nested "${" and increase depth
                    depth = depth + 1
                    i = i + 2
                    GoTo ContinuePlaceholderLoop
                End If

                ' the braces of the expression itself: ${{a: 1}}
                If c2 = "{" Then
                    braces = braces + 1
                ElseIf c2 = "}" And braces > 0 Then
                    braces = braces - 1
                    curExpr = curExpr & c2
                    i = i + 1
                    GoTo ContinuePlaceholderLoop
                End If

                ' unescaped '}' closes current placeholder level
                If c2 = "}" Then
                    ' consume closing brace
                    i = i + 1
                    depth = depth - 1
                    ' only add non-empty expression fragments
                    If Len(curExpr) > 0 Then
                        parts.Add Array("EXPR", curExpr)
                        curExpr = vbNullString
                    End If
                    ' if depth > 0 we continue collecting outer-level remainder into a new expr (will be added later or at another close)
                    GoTo ContinuePlaceholderLoop
                End If

                ' any other char inside placeholder: append literally
                curExpr = curExpr & c2
                i = i + 1

ContinuePlaceholderLoop:
            Loop

            ' If the placeholder wasn't properly closed (unterminated), treat whatever collected as an EXPR
            If Len(curExpr) > 0 Then
                parts.Add Array("EXPR", curExpr)
                curExpr = vbNullString
            End If

            ' continue scanning main template loop
            GoTo ContinueMainLoop
        End If

        ' closing backtick ends template (outside any placeholder)
        If ch = "`" Then
            If Len(curLit) > 0 Then parts.Add Array("LITERAL", curLit): curLit = vbNullString
            i = i + 1 ' consume closing backtick
            Exit Do
        End If

        ' normal literal char (outside placeholder)
        curLit = curLit & ch
        i = i + 1

ContinueMainLoop:
    Loop

    ' leftover (unterminated) literal
    If Len(curLit) > 0 Then parts.Add Array("LITERAL", curLit)

    Set ParseTemplatePartsFromString = parts
End Function

' Adds tokens after the session's: comments go, and a collapsed identifier (a[1].b) becomes its parts
Private Sub AppendTokenList(toks As Collection)
    Dim t As Variant
    For Each t In toks
        If t(0) <> "COMMENT" Then
            If t(0) = "IDENT" And (InStr(t(1), "[") > 0 Or InStr(t(1), ".") > 0) Then
                AppendCollapsed CStr(t(1))
            Else
                PutToken CStr(t(0)), t(1)
            End If
        End If
    Next t
End Sub

Private Sub PutToken(ByVal kind As String, ByVal value As Variant)
    nT_ = nT_ + 1
    ' one more slot: the end of the tokens (upTo) is read as a token that matches nothing
    If nT_ + 1 > UBound(tt_) Then
        ReDim Preserve tt_(1 To 2 * nT_): ReDim Preserve tv_(1 To 2 * nT_)
    End If
    tt_(nT_ + 1) = vbNullString: tv_(nT_ + 1) = Empty
    tt_(nT_) = kind
    tv_(nT_) = value
End Sub

' a[1][2].b, one identifier for the tokenizer: a, then [index] and .name parts
Private Sub AppendCollapsed(ByVal s As String)
    Dim pos As Long: pos = 1
    Dim ident As String
    Do While pos <= Len(s)
        If Mid$(s, pos, 1) = "." Or Mid$(s, pos, 1) = "[" Then Exit Do
        ident = ident & Mid$(s, pos, 1)
        pos = pos + 1
    Loop
    PutToken "IDENT", ident
    Dim name As String, d As Long, k As Long
    Do While pos <= Len(s)
        If Mid$(s, pos, 1) = "." Then
            pos = pos + 1
            name = vbNullString
            Do While pos <= Len(s)
                If Mid$(s, pos, 1) = "." Or Mid$(s, pos, 1) = "[" Then Exit Do
                name = name & Mid$(s, pos, 1)
                pos = pos + 1
            Loop
            PutToken "SYM", "."
            PutToken "IDENT", name
        ElseIf Mid$(s, pos, 1) = "[" Then
            d = 0
            For k = pos To Len(s)
                If Mid$(s, k, 1) = "[" Then d = d + 1
                If Mid$(s, k, 1) = "]" Then
                    d = d - 1
                    If d = 0 Then Exit For
                End If
            Next k
            If k > Len(s) Then CompileError "Invalid identifier: " & s
            PutToken "PAREN", "["
            AppendTokenList dParser_.Tokenize(Mid$(s, pos + 1, k - pos - 1))
            DropTrailingSemicolon
            PutToken "PAREN", "]"
            pos = k + 1
        Else
            CompileError "Invalid identifier: " & s
        End If
    Loop
End Sub

Private Sub DropTrailingSemicolon()
    Do While nT_ > 0
        If Not (tt_(nT_) = "SEP" And tv_(nT_) = ";") Then Exit Do
        nT_ = nT_ - 1
    Loop
End Sub

' The tokens of text after the session's (a part of a template): from and upTo give their range
Private Sub AppendText(ByVal text As String, ByRef from As Long, ByRef upTo As Long)
    from = nT_ + 1
    AppendTokenList dParser_.Tokenize(text)
    DropTrailingSemicolon
    upTo = nT_ + 1
End Sub

' Stored tokens (of a named function) after the session's
Private Sub AppendStored(ByRef types As Variant, ByRef vals As Variant, ByRef from As Long, ByRef upTo As Long)
    from = nT_ + 1
    Dim i As Long
    If IsArray(types) Then
        For i = LBound(types) To UBound(types)
            PutToken CStr(types(i)), vals(i)
        Next i
    End If
    upTo = nT_ + 1
End Sub

Private Function IsTok(ByVal i As Long, ByVal kind As String, ByVal value As String) As Boolean
    If i >= 1 And i <= nT_ Then IsTok = (tt_(i) = kind And tv_(i) = value)
End Function

Private Function IsWord(ByVal i As Long, ByVal word As String) As Boolean
    If i >= 1 And i <= nT_ Then
        If tt_(i) = "IDENT" Or tt_(i) = "KEYWORD" Then IsWord = (LCase$(tv_(i)) = word)
    End If
End Function

' The token that closes the bracket at i counting only its own kind, as the statement parsers do
Private Function CloseKind(ByVal i As Long, ByVal upTo As Long) As Long
    Dim openB As String: openB = tv_(i)
    Dim closeB As String
    Select Case openB
        Case "(": closeB = ")"
        Case "[": closeB = "]"
        Case "{": closeB = "}"
    End Select
    Dim d As Long
    i = i + 1
    Do While i < upTo
        If tt_(i) = "PAREN" Then
            If tv_(i) = openB Then
                d = d + 1
            ElseIf tv_(i) = closeB Then
                If d = 0 Then CloseKind = i: Exit Function
                d = d - 1
            End If
        End If
        i = i + 1
    Loop
    CloseKind = upTo
End Function

' The first token from i of kind and value, or upTo
Private Function FindTok(ByVal i As Long, ByVal upTo As Long, ByVal kind As String, ByVal value As String) As Long
    Do While i < upTo
        If tt_(i) = kind And tv_(i) = value Then Exit Do
        i = i + 1
    Loop
    FindTok = i
End Function

' ---------------------------------------------------------------- statements as plans
' The statements of the top level; named functions are
' registered (in dFuncs_), not statements
Private Function DTopLevel(ByVal from As Long, ByVal upTo As Long) As Collection
    Dim stmts As Collection: Set stmts = New Collection
    Dim i As Long, s As Long, plan As Variant, bD As Long, pD As Long, kD As Long
    i = from
    Do While i < upTo
        If tt_(i) = "KEYWORD" And tv_(i) = "import" Then
            stmts.Add DImport(i, upTo)
        ElseIf tt_(i) = "KEYWORD" And tv_(i) = "export" Then
            stmts.Add DExport(i, upTo)
        ElseIf tt_(i) = "IDENT" And tv_(i) = "fun" Then
            DNamedFunction i, upTo
        ElseIf tt_(i) = "IDENT" And tv_(i) = "class" Then
            stmts.Add Array("class", DClassDecl(i, upTo))
        ElseIf tt_(i) = "IDENT" And tv_(i) = "prototype" Then
            stmts.Add DPrototype(i, upTo, vbNullString)
        Else
            s = i: bD = 0: pD = 0: kD = 0
            Do While i < upTo
                If tt_(i) = "PAREN" Then
                    Select Case tv_(i)
                        Case "{": bD = bD + 1
                        Case "}": If bD > 0 Then bD = bD - 1
                        Case "(": pD = pD + 1
                        Case ")": If pD > 0 Then pD = pD - 1
                        Case "[": kD = kD + 1
                        Case "]": If kD > 0 Then kD = kD - 1
                    End Select
                End If
                If tt_(i) = "SEP" And tv_(i) = ";" And bD = 0 And pD = 0 And kD = 0 Then Exit Do
                i = i + 1
            Loop
            If i > s Then
                plan = DPlan(s, i)
                If plan(0) <> "option" Then stmts.Add plan
            End If
            i = i + 1
        End If
    Loop
    Set DTopLevel = stmts
End Function

' The statements of a block: a ; at depth 0 ends one,
' and so does the } that closes a block, unless elseif, else or catch follows
Private Function DBlock(ByVal from As Long, ByVal upTo As Long) As Collection
    Dim stmts As Collection
    Dim i As Long, s As Long, bD As Long, pD As Long, kD As Long, closed As Boolean, ended As Boolean
    Dim plan As Variant, hit As Variant
    ' the passes over a body (assignments, names, code) read its blocks once
    If from <= UBound(blockC_) Then
        If IsArray(blockC_(from)) Then
            hit = blockC_(from)
            If hit(0) = upTo Then
                Set DBlock = hit(1)
                Exit Function
            End If
        End If
    End If
    Set stmts = New Collection
    i = from
    Do While i < upTo
        s = i: bD = 0: pD = 0: kD = 0: closed = False: ended = False
        Do While i < upTo
            If tt_(i) = "PAREN" Then
                Select Case tv_(i)
                    Case "{": bD = bD + 1
                    Case "}"
                        If bD > 0 Then bD = bD - 1
                        If bD = 0 Then
                            i = i + 1
                            If Not (IsWord(i, "elseif") Or IsWord(i, "else") Or IsWord(i, "catch")) Or i >= upTo Then
                                closed = True
                                Exit Do
                            End If
                            GoTo NextToken
                        End If
                    Case "(": pD = pD + 1
                    Case ")": If pD > 0 Then pD = pD - 1
                    Case "[": kD = kD + 1
                    Case "]": If kD > 0 Then kD = kD - 1
                End Select
            End If
            If tt_(i) = "SEP" And tv_(i) = ";" And bD = 0 And pD = 0 And kD = 0 Then
                ended = True
                Exit Do
            End If
            i = i + 1
NextToken:
        Loop
        If i > s Then
            plan = DPlan(s, i)
            If plan(0) <> "option" Then stmts.Add plan
        End If
        If ended Then
            i = i + 1
        ElseIf closed Then
            If i < upTo Then
                If IsTok(i, "SEP", ";") Then i = i + 1
            End If
        End If
    Loop
    If from > UBound(blockC_) Then ReDim Preserve blockC_(1 To from + nT_)
    blockC_(from) = Array(upTo, stmts)
    Set DBlock = stmts
End Function

' One statement (tokens s..e-1)
Private Function DPlan(ByVal s As Long, ByVal e As Long) As Variant
    Dim k As Long, c As Long
    If e - s >= 2 And tt_(s) = "IDENT" And LCase$(tv_(s)) = "print" Then
        k = FindTok(s + 1, e, "PAREN", "(")
        If k < e Then
            c = CloseKind(k, e)
            DPlan = Array("print", DSplitArgs(k + 1, c, False))
            Exit Function
        End If
    End If
    If tt_(s) = "IDENT" Then
        Select Case LCase$(tv_(s))
            Case "if": DPlan = DPlanIf(s, e): Exit Function
            Case "for": DPlan = DPlanFor(s, e): Exit Function
            Case "while": DPlan = DPlanWhile(s, e): Exit Function
            Case "try": DPlan = DPlanTry(s, e): Exit Function
            Case "throw": DPlan = Array("throw", s + 1, e): Exit Function
            Case "switch": DPlan = DPlanSwitch(s, e): Exit Function
            Case "break": DPlan = Array("break"): Exit Function
            Case "continue": DPlan = Array("continue"): Exit Function
            Case "return": DPlan = Array("return", s + 1, e): Exit Function
            Case "option"
                If e - s >= 4 Then
                    If LCase$(tv_(s + 1)) = "base" And tt_(s + 2) = "NUMBER" Then
                        If CLng(tv_(s + 2)) = 0 Or CLng(tv_(s + 2)) = 1 Then
                            DPlan = Array("option")
                            Exit Function
                        End If
                    End If
                End If
        End Select
    End If
    ' an assignment: the first = or op= at depth 0
    Dim bD As Long, pD As Long, kD As Long, a As Long
    For k = s To e - 1
        If tt_(k) = "PAREN" Then
            Select Case tv_(k)
                Case "{": bD = bD + 1
                Case "}": If bD > 0 Then bD = bD - 1
                Case "(": pD = pD + 1
                Case ")": If pD > 0 Then pD = pD - 1
                Case "[": kD = kD + 1
                Case "]": If kD > 0 Then kD = kD - 1
            End Select
        End If
        If bD = 0 And pD = 0 And kD = 0 And tt_(k) = "OP" Then
            Select Case tv_(k)
                Case "=", "+=", "-=", "*=", "/=", "%=", "^=", "&=", "|=", "<<=", ">>="
                    a = k
                    Exit For
            End Select
        End If
    Next k
    If a > 0 Then
        If a - s >= 3 And IsTok(s, "PAREN", "[") And IsTok(a - 1, "PAREN", "]") Then
            DPlan = DPlanDestructuring(s, a, e)
        Else
            DPlan = Array("assign", s, a, CStr(tv_(a)), a + 1, e)
        End If
        Exit Function
    End If
    DPlan = Array("expr", s, e)
End Function

' Arguments (or print's): ranges split by , at depth 0; empty ones dropped (print), or Empty (calls)
Private Function DSplitArgs(ByVal from As Long, ByVal upTo As Long, ByVal keepEmpty As Boolean) As Collection
    Dim out As Collection: Set out = New Collection
    Dim i As Long, s As Long, d As Long
    s = from
    For i = from To upTo
        If i = upTo Then
            If i > s Or keepEmpty Then out.Add Array(s, i)
        ElseIf tt_(i) = "PAREN" Then
            Select Case tv_(i)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": If d > 0 Then d = d - 1
            End Select
        ElseIf tt_(i) = "SEP" And tv_(i) = "," And d = 0 Then
            If i > s Or keepEmpty Then out.Add Array(s, i)
            s = i + 1
        End If
    Next i
    Set DSplitArgs = out
End Function

Private Function DPlanDestructuring(ByVal s As Long, ByVal a As Long, ByVal e As Long) As Variant
    Dim names As Collection: Set names = New Collection
    Dim rest As String, hasRest As Boolean, k As Long
    k = s + 1
    Do While k < a - 1
        If tt_(k) = "SPREAD" Or (tt_(k) = "OP" And tv_(k) = "...") Then
            If hasRest Or k + 1 >= a - 1 Then CompileError "Invalid destructuring: a rest target takes a name and goes last"
            If tt_(k + 1) <> "IDENT" Then CompileError "Invalid destructuring: a rest target takes a name and goes last"
            rest = tv_(k + 1): hasRest = True
            k = k + 2
        ElseIf tt_(k) = "IDENT" Then
            If hasRest Then CompileError "Invalid destructuring: a rest target takes a name and goes last"
            names.Add CStr(tv_(k))
            k = k + 1
        Else
            k = k + 1
        End If
    Loop
    DPlanDestructuring = Array("destr", names, hasRest, rest, a + 1, e)
End Function

' if (c) { } elseif (c) { } else { }: parentheses and braces counted by kind
Private Function DPlanIf(ByVal s As Long, ByVal e As Long) As Variant
    Dim c1 As Long, c2 As Long, b1 As Long, b2 As Long, pos As Long
    Dim conds As Collection, blocks As Collection, hasElse As Boolean, eS As Long, eE As Long
    Set conds = New Collection: Set blocks = New Collection
    c1 = FindTok(s + 1, e, "PAREN", "(")
    If c1 >= e Then CompileError "Expected (condition) { ... } after if"
    c2 = CloseKind(c1, e)
    b1 = FindTok(c2, e, "PAREN", "{")
    If b1 >= e Then CompileError "Expected (condition) { ... } after if"
    b2 = CloseKind(b1, e)
    Dim cS As Long, cE As Long, tS As Long, tE As Long
    cS = c1 + 1: cE = c2: tS = b1 + 1: tE = b2
    pos = b2 + 1
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Do While pos < e
        If IsWord(pos, "elseif") And tt_(pos) = "IDENT" Then
            x1 = FindTok(pos + 1, e, "PAREN", "(")
            If x1 >= e Then Exit Do
            x2 = CloseKind(x1, e)
            y1 = FindTok(x2, e, "PAREN", "{")
            If y1 >= e Then Exit Do
            y2 = CloseKind(y1, e)
            conds.Add Array(x1 + 1, x2)
            blocks.Add Array(y1 + 1, y2)
            pos = y2 + 1
        ElseIf IsWord(pos, "else") And tt_(pos) = "IDENT" Then
            y1 = FindTok(pos + 1, e, "PAREN", "{")
            If y1 >= e Then Exit Do
            y2 = CloseKind(y1, e)
            hasElse = True: eS = y1 + 1: eE = y2
            Exit Do
        Else
            Exit Do
        End If
    Loop
    DPlanIf = Array("if", cS, cE, tS, tE, conds, blocks, hasElse, eS, eE)
End Function

' The end of a part of a for header: , or ; at depth 0, or the ) that closes the header
Private Function DHeaderPart(ByVal i As Long, ByVal e As Long, ByRef sepFound As Boolean) As Long
    Dim pD As Long, kD As Long, bD As Long
    sepFound = False
    Do While i < e
        If tt_(i) = "PAREN" Then
            Select Case tv_(i)
                Case "("
                    pD = pD + 1
                Case ")"
                    If pD > 0 Then pD = pD - 1 Else Exit Do
                Case "[": kD = kD + 1
                Case "]": If kD > 0 Then kD = kD - 1
                Case "{": bD = bD + 1
                Case "}": If bD > 0 Then bD = bD - 1
            End Select
        End If
        If pD = 0 And kD = 0 And bD = 0 And tt_(i) = "SEP" And (tv_(i) = "," Or tv_(i) = ";") Then
            sepFound = True
            Exit Do
        End If
        i = i + 1
    Loop
    DHeaderPart = i
End Function

Private Function DPlanFor(ByVal s As Long, ByVal e As Long) As Variant
    Dim i As Long, initS As Long, initE As Long, sep As Boolean
    i = FindTok(s + 1, e, "PAREN", "(")
    If i >= e Then CompileError "Expected (...) { ... } after for"
    i = i + 1
    initS = i
    initE = DHeaderPart(i, e, sep)
    i = initE
    If sep Then i = i + 1
    ' for (x in o) / for (x of a): in or of at depth 0 of the first part
    Dim k As Long, d As Long, found As Long, kind As String
    For k = initS To initE - 1
        If tt_(k) = "PAREN" Then
            Select Case tv_(k)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": If d > 0 Then d = d - 1
            End Select
        End If
        If d = 0 And tt_(k) = "IDENT" Then
            If LCase$(tv_(k)) = "in" Or LCase$(tv_(k)) = "of" Then
                found = k: kind = LCase$(tv_(k))
                Exit For
            End If
        End If
    Next k
    Dim b1 As Long, b2 As Long
    If found > 0 Then
        If IsTok(i, "PAREN", ")") Then i = i + 1
        b1 = FindTok(i, e, "PAREN", "{")
        If b1 >= e Then CompileError "Expected (...) { ... } after for"
        b2 = CloseKind(b1, e)
        DPlanFor = Array("foreach", kind, initS, found, found + 1, initE, b1 + 1, b2)
        Exit Function
    End If
    Dim condS As Long, condE As Long, stepS As Long, stepE As Long
    condS = i
    condE = DHeaderPart(i, e, sep)
    i = condE
    If sep Then i = i + 1
    stepS = i
    stepE = FindTok(i, e, "PAREN", ")")
    b1 = FindTok(stepE, e, "PAREN", "{")
    If b1 >= e Then CompileError "Expected (...) { ... } after for"
    b2 = CloseKind(b1, e)
    DPlanFor = Array("for", initS, initE, condS, condE, stepS, stepE, b1 + 1, b2)
End Function

Private Function DPlanWhile(ByVal s As Long, ByVal e As Long) As Variant
    Dim c1 As Long, c2 As Long, b1 As Long, b2 As Long
    c1 = FindTok(s + 1, e, "PAREN", "(")
    If c1 >= e Then CompileError "Expected (condition) { ... } after while"
    c2 = CloseKind(c1, e)
    b1 = FindTok(c2, e, "PAREN", "{")
    If b1 >= e Then CompileError "Expected (condition) { ... } after while"
    b2 = CloseKind(b1, e)
    DPlanWhile = Array("while", c1 + 1, c2, b1 + 1, b2)
End Function

Private Function DPlanTry(ByVal s As Long, ByVal e As Long) As Variant
    Dim b1 As Long, b2 As Long, k As Long, catchVar As String, c1 As Long, c2 As Long
    b1 = FindTok(s + 1, e, "PAREN", "{")
    If b1 >= e Then CompileError "Expected { ... } after try"
    b2 = CloseKind(b1, e)
    c1 = e: c2 = e
    For k = b2 + 1 To e - 1
        If tt_(k) = "IDENT" And LCase$(tv_(k)) = "catch" Then
            If IsTok(k + 1, "PAREN", "(") And k + 2 < e Then
                If tt_(k + 2) = "IDENT" Then catchVar = tv_(k + 2)
            End If
            c1 = FindTok(k + 1, e, "PAREN", "{")
            If c1 < e Then c2 = CloseKind(c1, e)
            Exit For
        End If
    Next k
    If c1 < e Then
        DPlanTry = Array("try", b1 + 1, b2, catchVar, c1 + 1, c2)
    Else
        DPlanTry = Array("try", b1 + 1, b2, catchVar, e, e)
    End If
End Function

Private Function DPlanSwitch(ByVal s As Long, ByVal e As Long) As Variant
    Dim i As Long, xS As Long, xE As Long
    i = FindTok(s + 1, e, "PAREN", "(")
    If i >= e Then CompileError "Expected (value) { ... } after switch"
    xS = i + 1
    xE = FindTok(xS, e, "PAREN", ")")
    i = FindTok(xE, e, "PAREN", "{")
    If i >= e Then CompileError "Expected (value) { ... } after switch"
    i = i + 1
    Dim cases As Collection: Set cases = New Collection
    Dim hasDefault As Boolean, dS As Long, dE As Long, vS As Long, pD As Long, kD As Long, b2 As Long
    Do While i < e
        If IsTok(i, "PAREN", "}") Then Exit Do
        If tt_(i) = "IDENT" And LCase$(tv_(i)) = "case" Then
            i = i + 1
            vS = i: pD = 0: kD = 0
            Do While i < e
                If tt_(i) = "PAREN" Then
                    Select Case tv_(i)
                        Case "(": pD = pD + 1
                        Case ")": If pD > 0 Then pD = pD - 1
                        Case "[": kD = kD + 1
                        Case "]": If kD > 0 Then kD = kD - 1
                        Case "{": If pD = 0 And kD = 0 Then Exit Do
                    End Select
                End If
                i = i + 1
            Loop
            If i >= e Then CompileError "switch: every case takes a block, as in case 2 { ... } and default { ... }; there is no case:, and cases do not fall through"
            If DColonCase(vS, i) Then CompileError "switch: every case takes a block, as in case 2 { ... } and default { ... }; there is no case:, and cases do not fall through"
            b2 = CloseKind(i, e)
            cases.Add Array(vS, i, i + 1, b2)
            i = b2 + 1
        ElseIf (tt_(i) = "IDENT" Or tt_(i) = "KEYWORD") And LCase$(tv_(i)) = "default" Then
            i = i + 1
            If IsTok(i, "OP", ":") Then CompileError "switch: every case takes a block, as in case 2 { ... } and default { ... }; there is no case:, and cases do not fall through"
            i = FindTok(i, e, "PAREN", "{")
            If i < e Then
                b2 = CloseKind(i, e)
                hasDefault = True: dS = i + 1: dE = b2
                i = b2 + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    DPlanSwitch = Array("switch", xS, xE, cases, hasDefault, dS, dE)
End Function

Private Function DColonCase(ByVal s As Long, ByVal e As Long) As Boolean
    Dim k As Long, d As Long, q As Long
    For k = s To e - 1
        If tt_(k) = "PAREN" Then
            Select Case tv_(k)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf tv_(k) = "?" And d = 0 Then
            q = q + 1
        ElseIf tv_(k) = ":" And d = 0 Then
            If q = 0 Then DColonCase = True: Exit Function
            q = q - 1
        End If
    Next k
End Function

' ---------------------------------------------------------------- declarations
' fun name(params) { body } at the top level: registered when the program compiles
Private Sub DNamedFunction(ByRef i As Long, ByVal upTo As Long)
    i = i + 1
    If i >= upTo Then CompileError "Expected function name and body after fun"
    Dim name As String: name = tv_(i)
    i = i + 1
    Dim params As Collection, rest As String
    DParams i, upTo, params, rest
    If Not IsTok(i, "PAREN", "{") Then CompileError "Expected function body"
    Dim b2 As Long: b2 = CloseKind(i, upTo)
    If GLOBALS_.gFuncTable.Exists(name) Or dFuncs_.Exists(name) Then CompileError "Function '" & name & "' is already defined"
    DRegisterFunction name, params, rest, i + 1, b2
    dNamed_ = dNamed_ + 1
    i = b2 + 1
End Sub

' (a, b, ...r): the parameters from i (at the parenthesis) and the rest one; i ends after the )
Private Sub DParams(ByRef i As Long, ByVal upTo As Long, ByRef params As Collection, ByRef rest As String)
    Set params = New Collection
    rest = vbNullString
    If Not IsTok(i, "PAREN", "(") Then Exit Sub
    i = i + 1
    Do While i < upTo
        If IsTok(i, "PAREN", ")") Then Exit Do
        If tt_(i) = "SPREAD" Or (tt_(i) = "OP" And tv_(i) = "...") Then
            If Len(rest) > 0 Then CompileError "Multiple rest parameters not allowed", ERR_REST_TWICE
            If i + 1 >= upTo Then CompileError "Expected parameter name after ...", ERR_REST_NAME
            If tt_(i + 1) <> "IDENT" Then CompileError "Expected parameter name after ...", ERR_REST_NAME
            rest = tv_(i + 1)
            i = i + 2
            If Not IsTok(i, "PAREN", ")") Then CompileError "Rest parameter must be last parameter", ERR_REST_LAST
        ElseIf tt_(i) = "IDENT" Then
            If Len(rest) > 0 Then CompileError "Rest parameter must be last parameter", ERR_REST_LAST
            params.Add CStr(tv_(i))
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    If IsTok(i, "PAREN", ")") Then i = i + 1
End Sub

' A named function this source defines: its tokens are kept, to compile it wherever it is used
Private Sub DRegisterFunction(ByVal name As String, params As Collection, ByVal rest As String, ByVal bodyS As Long, ByVal bodyE As Long)
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "params", params
    Dim pa() As String, j As Long
    If params.Count > 0 Then
        ReDim pa(0 To params.Count - 1)
        For j = 1 To params.Count
            pa(j - 1) = params(j)
        Next j
    Else
        ReDim pa(0 To 0)
    End If
    info.SetValue "paramArray", pa
    If Len(rest) > 0 Then info.SetValue "rest", rest Else info.SetValue "rest", Empty
    Dim types() As String, vals() As Variant
    If bodyE > bodyS Then
        ReDim types(0 To bodyE - bodyS - 1): ReDim vals(0 To bodyE - bodyS - 1)
        For j = bodyS To bodyE - 1
            types(j - bodyS) = tt_(j)
            vals(j - bodyS) = tv_(j)
        Next j
        info.SetValue "types", types
        info.SetValue "vals", vals
    Else
        info.SetValue "types", Empty
        info.SetValue "vals", Empty
    End If
    info.SetValue "bodyS", bodyS
    info.SetValue "bodyE", bodyE
    info.SetValue "index", GLOBALS_.gPrograms.Count + dFuncOrder_.Count + 1
    dFuncs_.SetValue name, info
    dFuncOrder_.Add name
End Sub

' class Name [extends Parent] { field a = x, b; constructor(...) { } [static] m(...) { } }
Private Function DClassDecl(ByRef i As Long, ByVal upTo As Long) As ASF_Map
    Dim decl As ASF_Map: Set decl = New ASF_Map
    decl.SetValue "direct", True
    i = i + 1
    If i >= upTo Then CompileError "Expected class name and body"
    decl.SetValue "name", CStr(tv_(i))
    i = i + 1
    If i < upTo Then
        If LCase$(tv_(i)) = "extends" Then
            i = i + 1
            If i >= upTo Then CompileError "Expected class name and body"
            decl.SetValue "superClass", CStr(tv_(i))
            i = i + 1
        End If
    End If
    If Not IsTok(i, "PAREN", "{") Then CompileError "Expected class body"
    i = i + 1
    Dim fields As Collection: Set fields = New Collection
    Dim methods As Collection: Set methods = New Collection
    Dim d As Long, member As String, isStatic As Boolean, fname As String, dS As Long, vD As Long
    Do While i < upTo
        If tt_(i) = "PAREN" And tv_(i) = "{" Then
            d = d + 1: i = i + 1
        ElseIf tt_(i) = "PAREN" And tv_(i) = "}" Then
            If d = 0 Then Exit Do
            d = d - 1: i = i + 1
        ElseIf tt_(i) = "IDENT" Then
            member = LCase$(tv_(i))
            isStatic = False
            If member = "static" Then
                isStatic = True
                i = i + 1
                If i >= upTo Then Exit Do
                member = LCase$(tv_(i))
            End If
            Select Case member
                Case "field"
                    If isStatic Then CompileError "Fields cannot be static"
                    i = i + 1
                    Do While i < upTo
                        If IsTok(i, "SEP", ";") Then
                            i = i + 1
                            Exit Do
                        End If
                        If tt_(i) = "IDENT" Then
                            fname = tv_(i)
                            i = i + 1
                            dS = 0
                            If IsTok(i, "OP", "=") Then
                                i = i + 1
                                dS = i: vD = 0
                                Do While i < upTo
                                    If tt_(i) = "PAREN" Then
                                        Select Case tv_(i)
                                            Case "(", "[", "{": vD = vD + 1
                                            Case ")", "]", "}": If vD > 0 Then vD = vD - 1
                                        End Select
                                    End If
                                    If vD = 0 And tt_(i) = "SEP" And (tv_(i) = "," Or tv_(i) = ";") Then Exit Do
                                    i = i + 1
                                Loop
                                If i = dS Then dS = 0
                            End If
                            fields.Add Array(fname, dS, i)
                            If IsTok(i, "SEP", ",") Then i = i + 1
                        Else
                            CompileError "Expected a field name"
                        End If
                    Loop
                Case "constructor"
                    If isStatic Then CompileError "Constructors cannot be static"
                    decl.SetValue "constructor", DMethodDecl(i, upTo, True, False)
                Case Else
                    methods.Add DMethodDecl(i, upTo, False, isStatic)
            End Select
        Else
            i = i + 1
        End If
    Loop
    If IsTok(i, "PAREN", "}") Then i = i + 1
    decl.SetValue "fields", fields
    decl.SetValue "classMethods", methods
    Set DClassDecl = decl
End Function

Private Function DMethodDecl(ByRef i As Long, ByVal upTo As Long, ByVal isCtor As Boolean, ByVal isStatic As Boolean) As ASF_Map
    Dim m As ASF_Map: Set m = New ASF_Map
    m.SetValue "name", CStr(tv_(i))
    m.SetValue "isConstructor", isCtor
    m.SetValue "isStatic", isStatic
    i = i + 1
    Dim params As Collection, rest As String
    DParams i, upTo, params, rest
    m.SetValue "params", params
    m.SetValue "rest", rest
    If Not IsTok(i, "PAREN", "{") Then CompileError "Expected method body"
    Dim b2 As Long: b2 = CloseKind(i, upTo)
    m.SetValue "bodyS", i + 1
    m.SetValue "bodyE", b2
    i = b2 + 1
    Set DMethodDecl = m
End Function

' prototype.COM.Type name(params) { body }: its function is registered as a named function; the
' parser goes on at the closing brace, as there
Private Function DPrototype(ByRef i As Long, ByVal upTo As Long, ByVal exportName As String) As Variant
    If Not (IsTok(i + 1, "SYM", ".") And IsTok(i + 2, "IDENT", "COM") And IsTok(i + 3, "SYM", ".")) Then CompileError "Expected prototype.COM.Type method(...) { ... }"
    If i + 5 >= upTo Then CompileError "Expected prototype.COM.Type method(...) { ... }"
    If tt_(i + 4) <> "IDENT" Or tt_(i + 5) <> "IDENT" Then CompileError "Expected prototype.COM.Type method(...) { ... }"
    Dim classType As String: classType = tv_(i + 4)
    Dim methodName As String: methodName = tv_(i + 5)
    i = i + 6
    Dim params As Collection, rest As String
    DParams i, upTo, params, rest
    If Not IsTok(i, "PAREN", "{") Then CompileError "Expected method body"
    Dim b2 As Long: b2 = CloseKind(i, upTo)
    Dim internal As String: internal = "__PROTOTYPE_" & UCase$(classType) & "_" & UCase$(methodName)
    If GLOBALS_.gFuncTable.Exists(internal) Or dFuncs_.Exists(internal) Then CompileError "Prototype " & classType & "." & methodName & " is already defined"
    DRegisterFunction internal, params, rest, i + 1, b2
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "classType", classType
    info.SetValue "methodName", methodName
    info.SetValue "funcIndex", dFuncs_.GetValue(internal).GetValue("index")
    info.SetValue "params", params
    info.SetValue "internalFuncName", internal
    i = b2
    If Len(exportName) > 0 Then info.SetValue "exportName", exportName
    DPrototype = Array("proto", info)
End Function

' import ... from 'module': the names it binds
Private Function DImport(ByRef i As Long, ByVal upTo As Long) As Variant
    Dim info As ASF_Map: Set info = New ASF_Map
    Dim named As Collection: Set named = New Collection
    info.SetValue "named", named
    i = i + 1
    If i >= upTo Then CompileError "Invalid import syntax"
    If tt_(i) = "STRING" Then
        info.SetValue "source", CStr(tv_(i))
        i = i + 1
        DImport = Array("import", info)
        Exit Function
    End If
    If IsTok(i, "OP", "*") Then
        If Not IsTok(i + 1, "KEYWORD", "as") Then CompileError "Invalid import syntax"
        If i + 2 >= upTo Then CompileError "Invalid import syntax"
        If tt_(i + 2) <> "IDENT" Then CompileError "Invalid import syntax"
        info.SetValue "namespaceImport", CStr(tv_(i + 2))
        i = i + 3
    ElseIf IsTok(i, "PAREN", "{") Then
        DImportList i, upTo, named
    ElseIf tt_(i) = "IDENT" Then
        info.SetValue "defaultImport", CStr(tv_(i))
        i = i + 1
        If IsTok(i, "SEP", ",") Then
            i = i + 1
            If IsTok(i, "PAREN", "{") Then DImportList i, upTo, named
        End If
    Else
        CompileError "Invalid import syntax"
    End If
    If Not IsTok(i, "IDENT", "from") Then CompileError "Invalid import syntax"
    i = i + 1
    If i >= upTo Then CompileError "Invalid import syntax"
    If tt_(i) <> "STRING" Then CompileError "Invalid import syntax"
    info.SetValue "source", CStr(tv_(i))
    i = i + 1
    DImport = Array("import", info)
End Function

' { a, b as c }: Array(imported, local) items; i ends after the }
Private Sub DImportList(ByRef i As Long, ByVal upTo As Long, named As Collection)
    Dim imported As String, localName As String
    i = i + 1
    Do While i < upTo
        If IsTok(i, "PAREN", "}") Then Exit Do
        If tt_(i) = "IDENT" Then
            imported = tv_(i): localName = imported
            i = i + 1
            If IsTok(i, "KEYWORD", "as") Then
                If i + 1 >= upTo Then CompileError "Invalid import syntax"
                If tt_(i + 1) <> "IDENT" Then CompileError "Invalid import syntax"
                localName = tv_(i + 1)
                i = i + 2
            End If
            named.Add Array(imported, localName)
            If IsTok(i, "SEP", ",") Then i = i + 1
        Else
            i = i + 1
        End If
    Loop
    If IsTok(i, "PAREN", "}") Then i = i + 1
End Sub

' export { a, b as c } / export default expression / export fun ... / export prototype...
Private Function DExport(ByRef i As Long, ByVal upTo As Long) As Variant
    Dim info As ASF_Map: Set info = New ASF_Map
    i = i + 1
    If i >= upTo Then CompileError "Invalid export syntax", ERR_EXPORT
    If IsTok(i, "KEYWORD", "default") Then
        i = i + 1
        Dim s As Long, d As Long
        s = i
        Do While i < upTo
            If tt_(i) = "PAREN" Then
                Select Case tv_(i)
                    Case "{", "(", "[": d = d + 1
                    Case "}", ")", "]": If d > 0 Then d = d - 1
                End Select
            End If
            If d = 0 And tt_(i) = "SEP" And tv_(i) = ";" Then Exit Do
            i = i + 1
        Loop
        info.SetValue "default", Array(s, i)
        If IsTok(i, "SEP", ";") Then i = i + 1
        DExport = Array("export", info)
        Exit Function
    End If
    If IsTok(i, "PAREN", "{") Then
        Dim named As Collection: Set named = New Collection
        DImportList i, upTo, named
        info.SetValue "named", named
        DExport = Array("export", info)
        Exit Function
    End If
    If tt_(i) = "IDENT" And tv_(i) = "prototype" Then
        Dim protoPlan As Variant
        Dim save As Long: save = i
        ' the export name: Type_method
        If i + 5 >= upTo Then CompileError "Invalid export syntax", ERR_EXPORT
        protoPlan = DPrototype(i, upTo, CStr(tv_(save + 4)) & "_" & CStr(tv_(save + 5)))
        ' export goes on after the closing brace
        i = i + 1
        DExport = protoPlan
        Exit Function
    End If
    If tt_(i) = "IDENT" And LCase$(tv_(i)) = "fun" Then
        ' the function is declared next; a module exports it when it ends
        If i + 1 >= upTo Then CompileError "Expected function name after 'fun'", ERR_EXPORT - 1
        If tt_(i + 1) <> "IDENT" Then CompileError "Expected function name after 'fun'", ERR_EXPORT - 1
        info.SetValue "declaration", CStr(tv_(i + 1))
        DExport = Array("export", info)
        Exit Function
    End If
    CompileError "Invalid export syntax", ERR_EXPORT
End Function

' ---------------------------------------------------------------- expressions
' mode_: 0 emits; 1 scans (names into scanUsed_, function literals analysed under scanCtx_ when
' scanLits_); 2 probes (only parses, to find where an expression ends)

' The expression of tokens s..e-1; tokens after it are ignored
Private Sub DExpr(ByVal s As Long, ByVal e As Long)
    Dim saveP As Long, saveL As Long
    saveP = p_: saveL = lim_
    p_ = s: lim_ = e
    DTernary
    p_ = saveP: lim_ = saveL
End Sub

Private Function PeekOp(ByVal op As String) As Boolean
    If p_ < lim_ Then PeekOp = (tt_(p_) = "OP" And tv_(p_) = op)
End Function

' cond ? a : b: the condition as a jump
Private Sub DTernary()
    Dim start As Long: start = p_
    If mode_ = 0 Then
        If TopQuestion(start, lim_) Then
            mode_ = 2
            DBinary 1
            mode_ = 0
            If PeekOp("?") Then
                Dim q As Long: q = p_
                Dim falses As Collection: Set falses = New Collection
                Dim saveL As Long: saveL = lim_
                DJumpIfFalse start, q, falses
                lim_ = saveL
                p_ = q + 1
                DTernary
                If Not PeekOp(":") Then CompileError "Expected ':' in ternary expression", ERR_TERNARY
                p_ = p_ + 1
                Emit OP_JMP, 0, 0, -1
                Dim endJump As Long: endJump = nOp_
                PatchAll falses, nOp_ + 1
                DTernary
                Patch endJump, nOp_ + 1
                Exit Sub
            End If
            p_ = start
        End If
        DBinary 1
        Exit Sub
    End If
    DBinary 1
    If PeekOp("?") Then
        p_ = p_ + 1
        DTernary
        If Not PeekOp(":") Then CompileError "Expected ':' in ternary expression", ERR_TERNARY
        p_ = p_ + 1
        DTernary
    End If
End Sub

' A ? at depth 0 from s on
Private Function TopQuestion(ByVal s As Long, ByVal e As Long) As Boolean
    Dim d As Long, k As Long
    For k = s To e - 1
        If tt_(k) = "PAREN" Then
            Select Case tv_(k)
                Case "(", "[", "{": d = d + 1
                Case ")", "]", "}": d = d - 1
            End Select
        ElseIf d = 0 And tt_(k) = "OP" Then
            If tv_(k) = "?" Then TopQuestion = True: Exit Function
        End If
    Next k
End Function

Private Function DPrec(ByVal op As String) As Long
    Select Case op
        Case "??": DPrec = 1
        Case "||": DPrec = 2
        Case "&&": DPrec = 3
        Case "==", "=", "!=": DPrec = 4
        Case "<", ">", "<=", ">=": DPrec = 5
        Case "<<", ">>": DPrec = 6
        Case "+", "-", "&": DPrec = 7
        Case "*", "/", "%": DPrec = 8
        Case "^": DPrec = 9
    End Select
End Function

' Precedence climbing; operands are emitted as they are parsed
Private Sub DBinary(ByVal minPrec As Long)
    Dim lStart As Long, rStart As Long, op As String, prec As Long, endJump As Long
    lStart = nOp_ + 1
    DUnary
    Do While p_ < lim_
        If tt_(p_) <> "OP" Then Exit Do
        op = tv_(p_)
        prec = DPrec(op)
        If prec = 0 Or prec < minPrec Then Exit Do
        p_ = p_ + 1
        If mode_ <> 0 Then
            If prec = 1 Or prec = 9 Then DBinary prec Else DBinary prec + 1
        Else
            Select Case op
                Case "&&", "||"
                    If op = "&&" Then Emit OP_ANDJ, 0, 0, -1 Else Emit OP_ORJ, 0, 0, -1
                    endJump = nOp_
                    DBinary prec + 1
                    Emit OP_TOBOOL, 0, 0, 0
                    Patch endJump, nOp_ + 1
                Case "??"
                    Emit OP_NULLJ, 0, 0, -1
                    endJump = nOp_
                    DBinary prec
                    Patch endJump, nOp_ + 1
                Case Else
                    If BinaryOp(op) = 0 Then CompileError "Unknown operator " & op
                    rStart = nOp_ + 1
                    If prec = 9 Then DBinary prec Else DBinary prec + 1
                    DBinaryCode op, lStart, rStart - 1
            End Select
        End If
    Loop
End Sub

' The operator over the left operand (lStart..lEnd) and the right one (lEnd+1..nOp_): two constants
' fold into one, x + constant is ADDK
Private Sub DBinaryCode(ByVal op As String, ByVal lStart As Long, ByVal lEnd As Long)
    Dim l As Variant, r As Variant, k As Variant, folded As Boolean
    If lEnd = lStart And nOp_ = lEnd + 1 And tplOp_ <> lStart And tplOp_ <> nOp_ Then
        If op_(lStart) = OP_PUSHK And op_(nOp_) = OP_PUSHK Then
            If Not IsObject(k_(a_(lStart))) And Not IsObject(k_(a_(nOp_))) Then
                l = k_(a_(lStart)): r = k_(a_(nOp_))
                If VarType(l) = vbDouble And VarType(r) = vbDouble Then
                    Select Case op
                        Case "+": k = l + r: folded = True
                        Case "-": k = l - r: folded = True
                        Case "*": k = l * r: folded = True
                        Case "/"
                            If r <> 0 Then k = l / r: folded = True
                    End Select
                ElseIf VarType(l) = vbString And VarType(r) = vbString Then
                    If op = "+" Then k = l & r: folded = True
                End If
            End If
            If folded Then
                DDropPush
                DDropPush
                Emit OP_PUSHK, Constant(k), 0, 1
                Exit Sub
            End If
        End If
    End If
    Dim code As Long: code = BinaryOp(op)
    If code = OP_ADD And nOp_ = lEnd + 1 Then
        If op_(nOp_) = OP_PUSHK Then
            If Not IsObject(k_(a_(nOp_))) Then
                If VarType(k_(a_(nOp_))) = vbDouble Then
                    op_(nOp_) = OP_ADDK
                    dl_(nOp_) = 0
                    depth_ = depth_ - 1
                    Exit Sub
                End If
            End If
        End If
    End If
    Emit code, 0, 0, -1
End Sub

' Drops the last instruction, a PUSHK, and its constant when that PUSHK added it
Private Sub DDropPush()
    Dim k As Long: k = a_(nOp_)
    If tplOp_ >= nOp_ Then tplOp_ = 0
    If k = nK_ And kOp_(k) = nOp_ Then
        kMap_.Remove ConstKey(k_(nK_))
        k_(nK_) = Empty
        nK_ = nK_ - 1
    End If
    nOp_ = nOp_ - 1
    depth_ = depth_ - 1
End Sub

Private Sub DUnary()
    If p_ < lim_ Then
        If tt_(p_) = "OP" And (tv_(p_) = "!" Or tv_(p_) = "-") Then
            Dim op As String: op = tv_(p_)
            p_ = p_ + 1
            Dim start As Long: start = nOp_ + 1
            DUnary
            If mode_ <> 0 Then Exit Sub
            If op = "-" And nOp_ = start And op_(nOp_) = OP_PUSHK Then
                If Not IsObject(k_(a_(nOp_))) Then
                    If VarType(k_(a_(nOp_))) = vbDouble Then
                        Dim v As Variant: v = -k_(a_(nOp_))
                        DDropPush
                        Emit OP_PUSHK, Constant(v), 0, 1
                        Exit Sub
                    End If
                End If
            End If
            If op = "!" Then Emit OP_NOT, 0, 0, 0 Else Emit OP_NEG, 0, 0, 0
            Exit Sub
        End If
    End If
    DPrimary
End Sub

Private Sub DPushK(ByRef v As Variant)
    If mode_ = 0 Then Emit OP_PUSHK, Constant(v), 0, 1
End Sub

Private Sub DPrimary()
    If p_ >= lim_ Then
        DPushK Empty
        Exit Sub
    End If
    Dim start As Long: start = p_
    Dim c As Long
    Select Case tt_(p_)
        Case "NUMBER"
            DPushK CDbl(tv_(p_))
            p_ = p_ + 1
        Case "STRING"
            If Len(tv_(p_)) >= 2 And Left$(tv_(p_), 1) = "`" And Right$(tv_(p_), 1) = "`" Then
                DTemplate p_
            Else
                DPushK tv_(p_)
            End If
            p_ = p_ + 1
            DPostfix start
        Case "VBEXPR"
            If mode_ = 0 Then EmitVBAExpr CStr(tv_(p_))
            p_ = p_ + 1
        Case "IDENT"
            DPrimaryIdent
        Case "PAREN"
            Select Case tv_(p_)
                Case "("
                    c = CloseKind(p_, lim_)
                    DExpr p_ + 1, c
                    p_ = c + 1
                    If p_ > lim_ Then p_ = lim_
                    DPostfix start
                Case "["
                    DArrayLiteral
                    DPostfix start
                Case "{"
                    DObjectLiteral
                    DPostfix start
                Case Else
                    DPushK Empty
            End Select
        Case Else
            DPushK Empty
    End Select
End Sub

Private Sub DPrimaryIdent()
    Dim name As String: name = tv_(p_)
    Select Case LCase$(name)
        Case "fun"
            DFunLiteral
            Exit Sub
        Case "typeof"
            p_ = p_ + 1
            DUnary
            If mode_ = 0 Then Emit OP_TYPEOF, 0, 0, 0
            Exit Sub
        Case "new"
            p_ = p_ + 1
            If p_ >= lim_ Then CompileError "Expected a class name after new"
            If tt_(p_) <> "IDENT" Then CompileError "Expected a class name after new"
            DPrimaryIdent
            Exit Sub
        Case "super"
            DSuper
            Exit Sub
        Case "true": DPushK True: p_ = p_ + 1: Exit Sub
        Case "false": DPushK False: p_ = p_ + 1: Exit Sub
        Case "null": DPushK Null: p_ = p_ + 1: Exit Sub
        Case "undefined": DPushK Empty: p_ = p_ + 1: Exit Sub
    End Select
    Dim start As Long: start = p_
    p_ = p_ + 1
    If mode_ = 1 Then scanUsed_.SetValue name, True
    If mode_ = 0 Then
        If p_ < lim_ Then
            If tt_(p_) = "PAREN" And tv_(p_) = "(" Then
                DCallName name
                DPostfix start
                Exit Sub
            End If
        End If
        EmitLoadName name
    End If
    DPostfix start
End Sub

' name(args): a named function, a built-in one or a value
Private Sub DCallName(ByVal name As String)
    Dim slot As Long, kind As String, n As Long, spread As Boolean
    kind = NameKind(name, slot)
    If kind = "F" Then
        DArgs n, spread
        If spread Then Emit OP_CALLFV, slot, 0, -n Else Emit OP_CALLF, slot, n, 1 - n
        Exit Sub
    End If
    If kind = "G" And Not GlobalAssigned(name) Then
        Select Case LCase$(name)
            Case "range", "flatten", "clone", "isarray", "isnumeric", "foreach", "regex", "cwd", "scwd"
                DArgs n, spread
                If spread Then Emit OP_CALLBV, Constant(name), 0, -n Else Emit OP_CALLB, Constant(name), n, 1 - n
                Exit Sub
        End Select
    End If
    EmitLoadName name
    DArgs n, spread
    If spread Then Emit OP_CALLV, 0, 0, -n - 1 Else Emit OP_CALL, n, 0, -n
End Sub

' The arguments of a call, from the ( at p_: n and whether one is
' spread; an empty argument is undefined
Private Sub DArgs(ByRef n As Long, ByRef spread As Boolean)
    Dim flags As String, isSpread As Boolean, s As Long, d1 As Long, d2 As Long, d3 As Long
    n = 0: spread = False
    p_ = p_ + 1
    Do
        isSpread = False
        If p_ < lim_ Then
            If tt_(p_) = "SPREAD" Or (tt_(p_) = "OP" And tv_(p_) = "...") Then
                isSpread = True
                p_ = p_ + 1
            End If
        End If
        s = p_: d1 = 0: d2 = 0: d3 = 0
        Do While p_ < lim_
            If tt_(p_) = "PAREN" Then
                Select Case tv_(p_)
                    Case "(": d1 = d1 + 1
                    Case ")": If d1 = 0 Then Exit Do Else d1 = d1 - 1
                    Case "[": d2 = d2 + 1
                    Case "]": If d2 = 0 Then Exit Do Else d2 = d2 - 1
                    Case "{": d3 = d3 + 1
                    Case "}": If d3 = 0 Then Exit Do Else d3 = d3 - 1
                End Select
            ElseIf tt_(p_) = "SEP" And tv_(p_) = "," And d1 = 0 And d2 = 0 And d3 = 0 Then
                Exit Do
            End If
            p_ = p_ + 1
        Loop
        If p_ > s Then
            DExpr s, p_
            n = n + 1
            If isSpread Then
                flags = flags & "1"
                spread = True
            Else
                flags = flags & "0"
            End If
        ElseIf Not isSpread Then
            DPushK Empty
            n = n + 1
            flags = flags & "0"
        End If
        If p_ < lim_ Then
            If tt_(p_) = "SEP" And tv_(p_) = "," Then
                p_ = p_ + 1
            ElseIf tt_(p_) = "PAREN" And tv_(p_) = ")" Then
                p_ = p_ + 1
                Exit Do
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    If spread And mode_ = 0 Then Emit OP_EXPAND, n, Constant(flags), 1
End Sub

' .prop, [index] and (args) after a primary that starts at token chainStart
Private Sub DPostfix(ByVal chainStart As Long)
    Dim c As Long, n As Long, spread As Boolean, prop As String, dotPos As Long
    Do While p_ < lim_
        If tt_(p_) = "SYM" And tv_(p_) = "." Then
            If p_ + 1 >= lim_ Then
                p_ = p_ + 1
                Exit Do
            End If
            If tt_(p_ + 1) <> "IDENT" Then
                p_ = p_ + 1
                Exit Do
            End If
            dotPos = p_
            prop = tv_(p_ + 1)
            p_ = p_ + 2
            If p_ < lim_ Then
                If tt_(p_) = "PAREN" And tv_(p_) = "(" Then
                    DMethodCall chainStart, dotPos, prop
                    GoTo NextPart
                End If
            End If
            If mode_ = 0 Then Emit OP_GETPROP, Constant(prop), 0, 0
        ElseIf tt_(p_) = "PAREN" And tv_(p_) = "[" Then
            c = CloseKind(p_, lim_)
            DExpr p_ + 1, c
            If mode_ = 0 Then Emit OP_GETIDX, 0, 0, -1
            p_ = c + 1
            If p_ > lim_ Then p_ = lim_
        ElseIf tt_(p_) = "PAREN" And tv_(p_) = "(" Then
            DArgs n, spread
            If mode_ = 0 Then
                If spread Then Emit OP_CALLV, 0, 0, -n - 1 Else Emit OP_CALL, n, 0, -n
            End If
        Else
            Exit Do
        End If
NextPart:
    Loop
End Sub

' receiver.name(args), the receiver being tokens recvS..recvE-1, already emitted
Private Sub DMethodCall(ByVal recvS As Long, ByVal recvE As Long, ByVal name As String)
    Dim n As Long, spread As Boolean, writes As Boolean
    If mode_ = 0 Then
        Select Case LCase$(name)
            Case "push", "pop", "shift", "unshift", "splice", "copywithin", "reverse", "sort", "tosorted", "concat", "delete"
                writes = DIsWritable(recvS, recvE)
        End Select
    End If
    DArgs n, spread
    If mode_ <> 0 Then Exit Sub
    Dim id As Long: id = METHOD_SHIFT * MethodId(name)
    If writes Then
        If spread Then Emit OP_CALLMWV, Constant(name), id, -n Else Emit OP_CALLMW, Constant(name), n + id, 1 - n
        DStoreTarget recvS, recvE
    Else
        If spread Then Emit OP_CALLMV, Constant(name), id, -n - 1 Else Emit OP_CALLM, Constant(name), n + id, -n
    End If
End Sub

Private Sub DArrayLiteral()
    Dim c As Long, s As Long, d As Long, n As Long, isSpread As Boolean, flags As String, spread As Boolean
    p_ = p_ + 1
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "]" Then
            p_ = p_ + 1
            If mode_ = 0 Then Emit OP_NEWARR, 0, 0, 1
            Exit Sub
        End If
    End If
    Do While p_ < lim_
        If tt_(p_) = "PAREN" And tv_(p_) = "]" Then Exit Do
        isSpread = False
        If tt_(p_) = "SPREAD" Or (tt_(p_) = "OP" And tv_(p_) = "...") Then
            isSpread = True
            p_ = p_ + 1
        End If
        s = p_: d = 0
        Do While p_ < lim_
            If d = 0 Then
                If (tt_(p_) = "SEP" And tv_(p_) = ",") Or (tt_(p_) = "PAREN" And tv_(p_) = "]") Then Exit Do
            End If
            If tt_(p_) = "PAREN" Then
                Select Case tv_(p_)
                    Case "[", "(", "{": d = d + 1
                    Case "]", ")", "}": If d > 0 Then d = d - 1
                End Select
            End If
            p_ = p_ + 1
        Loop
        If p_ > s Then
            DExpr s, p_
            n = n + 1
            If isSpread Then flags = flags & "1": spread = True Else flags = flags & "0"
        End If
        If p_ < lim_ Then
            If tt_(p_) = "SEP" And tv_(p_) = "," Then p_ = p_ + 1
        End If
    Loop
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "]" Then p_ = p_ + 1
    End If
    If mode_ <> 0 Then Exit Sub
    If spread Then
        Emit OP_NEWARRS, n, Constant(flags), 1 - n
    Else
        Emit OP_NEWARR, n, 0, 1 - n
    End If
End Sub

Private Sub DObjectLiteral()
    Dim s As Long, d As Long, n As Long, key As String
    p_ = p_ + 1
    Do While p_ < lim_
        If tt_(p_) = "PAREN" And tv_(p_) = "}" Then Exit Do
        If tt_(p_) = "SPREAD" Or (tt_(p_) = "OP" And tv_(p_) = "...") Then
            p_ = p_ + 1
            key = "__SPREAD__"
        ElseIf tt_(p_) = "IDENT" Or tt_(p_) = "STRING" Then
            key = tv_(p_)
            p_ = p_ + 1
            If Not PeekOp(":") Then CompileError "Expected ':' after object key", ERR_OBJECT_COLON
            p_ = p_ + 1
        Else
            CompileError "Invalid object key", ERR_OBJECT_KEY
        End If
        s = p_: d = 0
        Do While p_ < lim_
            If d = 0 Then
                If (tt_(p_) = "SEP" And tv_(p_) = ",") Or (tt_(p_) = "PAREN" And tv_(p_) = "}") Then Exit Do
                If tt_(p_) = "SEP" And tv_(p_) = ";" Then CompileError "Expected ',' between object properties, found ';'", ERR_OBJECT_SEMICOLON
            End If
            If tt_(p_) = "PAREN" Then
                Select Case tv_(p_)
                    Case "[", "(", "{": d = d + 1
                    Case "]", ")", "}": If d > 0 Then d = d - 1
                End Select
            End If
            p_ = p_ + 1
        Loop
        If key = "__SPREAD__" And p_ = s Then
            ' an empty spread adds nothing
        Else
            If mode_ = 0 Then Emit OP_PUSHK, Constant(key), 0, 1
            DExpr s, p_
            n = n + 1
        End If
        If p_ < lim_ Then
            If tt_(p_) = "SEP" And tv_(p_) = "," Then p_ = p_ + 1
        End If
    Loop
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "}" Then p_ = p_ + 1
    End If
    If mode_ = 0 Then Emit OP_NEWOBJ, n, 0, 1 - 2 * n
End Sub

' `...${x}...`: its parts are tokenized once (kept by the token's position), then emitted as
' one text when every part is text, else the parts and TPL
Private Sub DTemplate(ByVal at As Long)
    Dim parts As Collection
    If tplParts_.Exists(CStr(at)) Then
        Set parts = tplParts_.GetValue(CStr(at))
    Else
        Set parts = New Collection
        Dim raw As Variant, from As Long, upTo As Long
        For Each raw In ParseTemplatePartsFromString(CStr(tv_(at)))
            If raw(0) = "LITERAL" Then
                parts.Add Array(True, CStr(raw(1)), 0, 0)
            Else
                AppendText CStr(raw(1)), from, upTo
                If upTo - from = 1 Then
                    If tt_(from) = "STRING" And Left$(tv_(from), 1) <> "`" Then
                        parts.Add Array(True, CStr(tv_(from)), 0, 0)
                        GoTo NextPart
                    End If
                End If
                parts.Add Array(False, vbNullString, from, upTo)
            End If
NextPart:
        Next raw
        tplParts_.SetValue CStr(at), parts
    End If
    Dim tp As Variant, allText As Boolean, text As String, n As Long
    Dim saveP As Long, saveL As Long
    saveP = p_: saveL = lim_
    If mode_ <> 0 Then
        For Each tp In parts
            If Not tp(0) Then DExpr tp(2), tp(3)
        Next tp
        p_ = saveP: lim_ = saveL
        Exit Sub
    End If
    allText = True
    For Each tp In parts
        If tp(0) Then text = text & tp(1) Else allText = False
    Next tp
    If allText Then
        ' a template is not a literal: an operator does not fold it
        Emit OP_PUSHK, Constant(text), 0, 1
        tplOp_ = nOp_
        Exit Sub
    End If
    For Each tp In parts
        If tp(0) Then
            Emit OP_PUSHK, Constant(CStr(tp(1))), 0, 1
        Else
            DExpr tp(2), tp(3)
        End If
        n = n + 1
    Next tp
    p_ = saveP: lim_ = saveL
    Emit OP_TPL, n, 0, 1 - n
End Sub

' fun (params) { body }: a closure (emit), or a literal to analyse (scan)
Private Sub DFunLiteral()
    Dim litStart As Long: litStart = p_
    p_ = p_ + 1
    Dim params As Collection: Set params = New Collection
    Dim rest As String
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "(" Then
            p_ = p_ + 1
            Do While p_ < lim_
                If tt_(p_) = "PAREN" And tv_(p_) = ")" Then Exit Do
                If tt_(p_) = "SPREAD" Or (tt_(p_) = "OP" And tv_(p_) = "...") Then
                    p_ = p_ + 1
                    If p_ >= lim_ Then CompileError "Expected parameter name after ...", ERR_REST_NAME
                    If Len(rest) > 0 Then CompileError "Multiple rest parameters not allowed", ERR_REST_TWICE
                    If tt_(p_) <> "IDENT" Then CompileError "Expected parameter name after ...", ERR_REST_NAME
                    rest = tv_(p_)
                    p_ = p_ + 1
                    If PeekSep(",") Then CompileError "Rest parameter must be last parameter", ERR_REST_LAST
                ElseIf tt_(p_) = "IDENT" Then
                    If Len(rest) > 0 Then CompileError "Rest parameter must be last parameter", ERR_REST_LAST
                    params.Add CStr(tv_(p_))
                    p_ = p_ + 1
                    If PeekSep(",") Then p_ = p_ + 1
                Else
                    p_ = p_ + 1
                End If
            Loop
            If p_ < lim_ Then p_ = p_ + 1
        End If
    End If
    Dim bodyS As Long, bodyE As Long, b2 As Long
    bodyS = p_: bodyE = p_
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "{" Then
            b2 = CloseKind(p_, lim_)
            bodyS = p_ + 1: bodyE = b2
            p_ = b2 + 1
            If p_ > lim_ Then p_ = lim_
        End If
    End If
    Select Case mode_
        Case 0
            Dim ctx As ASF_Map: Set ctx = DLiteralContext(litStart)
            If Not ctx.Exists("fn") Then
                ctx.SetValue "fn", NewFunction("@literal")
                pendingLits_.Add ctx
            End If
            Emit OP_MKCLOS, ctx.GetValue("fn"), 0, 1
        Case 1
            If scanLits_ Then DAnalyzeLiteral litStart, params, rest, bodyS, bodyE, scanCtx_
    End Select
End Sub

Private Function PeekSep(ByVal v As String) As Boolean
    If p_ < lim_ Then PeekSep = (tt_(p_) = "SEP" And tv_(p_) = v)
End Function

' super(args): the parent's constructor; the arguments are its non-empty parts
Private Sub DSuper()
    p_ = p_ + 1
    Dim args As Collection: Set args = New Collection
    Dim s As Long, d As Long
    If p_ < lim_ Then
        If tt_(p_) = "PAREN" And tv_(p_) = "(" Then
            p_ = p_ + 1
            Do While p_ < lim_
                If tt_(p_) = "PAREN" And tv_(p_) = ")" Then Exit Do
                s = p_: d = 0
                Do While p_ < lim_
                    If tt_(p_) = "PAREN" Then
                        If tv_(p_) = "(" Then
                            d = d + 1
                        ElseIf tv_(p_) = ")" Then
                            If d = 0 Then Exit Do
                            d = d - 1
                        End If
                    ElseIf tt_(p_) = "SEP" And tv_(p_) = "," And d = 0 Then
                        Exit Do
                    End If
                    p_ = p_ + 1
                Loop
                If p_ > s Then args.Add Array(s, p_)
                If PeekSep(",") Then p_ = p_ + 1
            Loop
            If p_ < lim_ Then p_ = p_ + 1
        End If
    End If
    Dim a As Variant
    If mode_ <> 0 Then
        For Each a In args
            DExpr a(0), a(1)
        Next a
        Exit Sub
    End If
    DEmitSuperCall args
End Sub

' ---------------------------------------------------------------- conditions
' The expression of tokens s..e-1 as a condition: jumps get the jumps taken when
' it is false
Private Sub DJumpIfFalse(ByVal s As Long, ByVal e As Long, jumps As Collection)
    Dim at As Long, op As String
    DUnwrap s, e
    op = DTopOperator(s, e, at)
    Select Case op
        Case "&&"
            DJumpIfFalse s, at, jumps
            DJumpIfFalse at + 1, e, jumps
        Case "||"
            Dim trues As Collection: Set trues = New Collection
            DJumpIfTrue s, at, trues
            DJumpIfFalse at + 1, e, jumps
            PatchAll trues, nOp_ + 1
        Case "!"
            DJumpIfTrue s + 1, e, jumps
        Case "<", "<=", ">", ">=", "==", "=", "!="
            DExpr s, e
            op_(nOp_) = CompareJump(op)
            dl_(nOp_) = -2
            depth_ = depth_ - 1
            jumps.Add nOp_
        Case Else
            DExpr s, e
            Emit OP_JMPF, 0, 0, -1
            jumps.Add nOp_
    End Select
End Sub

Private Sub DJumpIfTrue(ByVal s As Long, ByVal e As Long, jumps As Collection)
    Dim at As Long, op As String
    DUnwrap s, e
    op = DTopOperator(s, e, at)
    Select Case op
        Case "||"
            DJumpIfTrue s, at, jumps
            DJumpIfTrue at + 1, e, jumps
        Case "&&"
            Dim falses As Collection: Set falses = New Collection
            DJumpIfFalse s, at, falses
            DJumpIfTrue at + 1, e, jumps
            PatchAll falses, nOp_ + 1
        Case "!"
            DJumpIfFalse s + 1, e, jumps
        Case Else
            DExpr s, e
            Emit OP_JMPT, 0, 0, -1
            jumps.Add nOp_
    End Select
End Sub

' (x): parentheses only group
Private Sub DUnwrap(ByRef s As Long, ByRef e As Long)
    Do While e - s >= 2
        If Not (tt_(s) = "PAREN" And tv_(s) = "(") Then Exit Do
        If CloseKind(s, e) <> e - 1 Then Exit Do
        s = s + 1: e = e - 1
    Loop
End Sub

' The root of the expression in tokens s..e-1, as the parser builds it: "?" for a conditional, the
' binary operator of lowest precedence at depth 0 (the last of them, or the first for ?? and ^,
' which group to the right), "!" for a ! over the whole expression, "" otherwise
Private Function DTopOperator(ByVal s As Long, ByVal e As Long, ByRef at As Long) As String
    Dim j As Long, d As Long, best As Long, bestPrec As Long, prec As Long, prevOperand As Boolean
    bestPrec = 99
    For j = s To e - 1
        Select Case tt_(j)
            Case "PAREN"
                Select Case tv_(j)
                    Case "(", "[", "{": d = d + 1
                    Case ")", "]", "}": d = d - 1
                End Select
                prevOperand = True
            Case "OP"
                If d = 0 Then
                    If tv_(j) = "?" Then
                        DTopOperator = "?"
                        Exit Function
                    End If
                    prec = 0
                    If prevOperand Then prec = DPrec(CStr(tv_(j)))
                    If prec > 0 Then
                        If prec < bestPrec Or (prec = bestPrec And prec <> 1 And prec <> 9) Then
                            bestPrec = prec
                            best = j
                        End If
                    End If
                End If
                prevOperand = False
            Case "SYM", "SEP", "SPREAD"
                prevOperand = False
            Case "IDENT"
                Select Case LCase$(tv_(j))
                    Case "typeof", "new": prevOperand = False
                    Case Else: prevOperand = True
                End Select
            Case Else
                prevOperand = True
        End Select
    Next j
    If best > 0 Then
        at = best
        DTopOperator = tv_(best)
    ElseIf e <= s Then
        Exit Function
    ElseIf tt_(s) = "OP" And tv_(s) = "!" Then
        DTopOperator = "!"
    End If
End Function

' ---------------------------------------------------------------- targets
' A chain: a root (a name, or anything else) and its parts .p, [i], (args) up to e; False when the
' tokens are not one
Private Function DChain(ByVal s As Long, ByVal e As Long, ByRef rootName As String, ops As Collection) As Boolean
    Set ops = New Collection
    rootName = vbNullString
    If s >= e Then Exit Function
    Dim r As Long, c As Long
    Select Case tt_(s)
        Case "IDENT"
            Select Case LCase$(tv_(s))
                Case "fun", "typeof", "new", "super", "true", "false", "null", "undefined": Exit Function
            End Select
            rootName = tv_(s)
            r = s + 1
        Case "STRING"
            r = s + 1
        Case "PAREN"
            Select Case tv_(s)
                Case "(", "[", "{"
                    c = CloseKind(s, e)
                    If c >= e Then Exit Function
                    If tv_(s) = "(" Then
                        Dim innerOps As Collection, innerName As String
                        If DChain(s + 1, c, innerName, innerOps) Then
                            If innerOps.Count = 0 Then rootName = innerName
                        End If
                    End If
                    r = c + 1
                Case Else
                    Exit Function
            End Select
        Case Else
            Exit Function
    End Select
    Do While r < e
        If tt_(r) = "SYM" And tv_(r) = "." Then
            If r + 1 >= e Then Exit Function
            If tt_(r + 1) <> "IDENT" Then Exit Function
            ops.Add Array("P", r, CStr(tv_(r + 1)), 0)
            r = r + 2
        ElseIf tt_(r) = "PAREN" And tv_(r) = "[" Then
            c = CloseKind(r, e)
            If c >= e Then Exit Function
            ops.Add Array("I", r, r + 1, c)
            r = c + 1
        ElseIf tt_(r) = "PAREN" And tv_(r) = "(" Then
            c = CloseKind(r, e)
            If c >= e Then Exit Function
            ops.Add Array("C", r, r + 1, c)
            r = c + 1
        Else
            Exit Function
        End If
    Loop
    DChain = True
End Function

' A variable, a property, or an index without side effects into a writable container
Private Function DIsWritable(ByVal s As Long, ByVal e As Long) As Boolean
    DUnwrap s, e
    Dim root As String, ops As Collection
    If Not DChain(s, e, root, ops) Then Exit Function
    If ops.Count = 0 Then
        DIsWritable = (Len(root) > 0)
        Exit Function
    End If
    Dim last As Variant: last = ops(ops.Count)
    Select Case last(0)
        Case "P": DIsWritable = True
        Case "I": DIsWritable = DIsSimple(last(2), last(3)) And DIsWritable(s, last(1))
    End Select
End Function

' No calls, literals of arrays, objects or templates, conditionals or functions
Private Function DIsSimple(ByVal s As Long, ByVal e As Long) As Boolean
    Dim j As Long, prevOperand As Boolean
    For j = s To e - 1
        Select Case tt_(j)
            Case "PAREN"
                Select Case tv_(j)
                    Case "(": If prevOperand Then Exit Function
                    Case "[": If Not prevOperand Then Exit Function
                    Case "{": Exit Function
                End Select
                prevOperand = (tv_(j) = ")" Or tv_(j) = "]")
            Case "OP"
                If tv_(j) = "?" Or tv_(j) = "..." Then Exit Function
                prevOperand = False
            Case "STRING"
                If Left$(tv_(j), 1) = "`" Then Exit Function
                prevOperand = True
            Case "VBEXPR", "SPREAD": Exit Function
            Case "IDENT"
                Select Case LCase$(tv_(j))
                    Case "fun", "new", "super": Exit Function
                    Case "typeof": prevOperand = False
                    Case Else: prevOperand = True
                End Select
            Case "SYM": prevOperand = False
            Case Else: prevOperand = True
        End Select
    Next j
    DIsSimple = True
End Function

' Stores the value on top of the stack in the target of tokens s..e-1
Private Sub DStoreTarget(ByVal s As Long, ByVal e As Long)
    DUnwrap s, e
    Dim root As String, ops As Collection
    If Not DChain(s, e, root, ops) Then CompileError "Invalid assignment target"
    If ops.Count = 0 Then
        If Len(root) = 0 Then CompileError "Invalid assignment target"
        EmitStoreName root
        Exit Sub
    End If
    Dim last As Variant: last = ops(ops.Count)
    Select Case last(0)
        Case "P"
            DExpr s, last(1)
            Emit OP_SETPROPV, Constant(CStr(last(2))), 0, -2
        Case "I"
            DContainerForWrite s, last(1)
            DIndex last(2), last(3)
            Emit OP_SETIDXR, 0, 0, -2
            DStoreTarget s, last(1)
        Case Else
            CompileError "Invalid assignment target: a call"
    End Select
End Sub

' a[i]... = v: an index with side effects (a call) is evaluated once, into a hidden variable, as the
' target is read and written back through it
Private Sub DPrepareIndexes(ByVal s As Long, ByVal e As Long)
    DUnwrap s, e
    Dim root As String, ops As Collection, op As Variant, t As String
    If Not DChain(s, e, root, ops) Then Exit Sub
    For Each op In ops
        If op(0) = "I" Then
            If Not DIsSimple(op(2), op(3)) Then
                t = NewTemp()
                DExpr op(2), op(3)
                EmitStoreName t
                If idxTemps_ Is Nothing Then Set idxTemps_ = New ASF_Map
                idxTemps_.SetValue CStr(op(2)) & ":" & CStr(op(3)), t
            End If
        End If
    Next op
End Sub

' The index of tokens s..e-1 of a target: its hidden variable, if DPrepareIndexes made one
Private Sub DIndex(ByVal s As Long, ByVal e As Long)
    Dim t As Variant
    If Not idxTemps_ Is Nothing Then
        If idxTemps_.TryGetValue(CStr(s) & ":" & CStr(e), t) Then
            EmitLoadName CStr(t)
            Exit Sub
        End If
    End If
    DExpr s, e
End Sub

' a value on the stack goes to the target of tokens s..e-1
Private Sub DAssignTo(ByVal s As Long, ByVal e As Long)
    DPrepareIndexes s, e
    DStoreTarget s, e
    Set idxTemps_ = Nothing
End Sub

' The container of a target: an indexed one is read with GETIDXW
Private Sub DContainerForWrite(ByVal s As Long, ByVal e As Long)
    Dim s2 As Long, e2 As Long
    s2 = s: e2 = e
    DUnwrap s2, e2
    Dim root As String, ops As Collection
    If DChain(s2, e2, root, ops) Then
        If ops.Count > 0 Then
            Dim last As Variant: last = ops(ops.Count)
            If last(0) = "I" Then
                DContainerForWrite s2, last(1)
                DIndex last(2), last(3)
                Emit OP_GETIDXW, 0, 0, -1
                Exit Sub
            End If
        End If
    End If
    DExpr s, e
End Sub

' The variable an assignment to tokens s..e-1 assigns: through its indexes
Private Function DTargetRoot(ByVal s As Long, ByVal e As Long) As String
    DUnwrap s, e
    Dim root As String, ops As Collection
    If Not DChain(s, e, root, ops) Then Exit Function
    Dim k As Long, op As Variant: k = ops.Count
    Do While k > 0
        op = ops(k)
        If op(0) <> "I" Then Exit Function
        k = k - 1
    Loop
    If ops.Count = 0 Then
        DTargetRoot = root
    Else
        op = ops(1)
        DTargetRoot = DTargetRoot(s, op(1))
    End If
End Function

' ---------------------------------------------------------------- analysis
' The names plans assign, in the order of their statements
Private Sub DCollectAssigned(plans As Collection, found As ASF_Map)
    Dim plan As Variant, nm As String, item As Variant, info As ASF_Map
    For Each plan In plans
        Select Case plan(0)
            Case "assign"
                nm = DTargetRoot(plan(1), plan(2))
                If Len(nm) > 0 Then found.SetValue nm, True
            Case "if"
                DCollectAssigned DBlock(plan(3), plan(4)), found
                For Each item In plan(6)
                    DCollectAssigned DBlock(item(0), item(1)), found
                Next item
                If plan(7) Then DCollectAssigned DBlock(plan(8), plan(9)), found
            Case "for"
                nm = DInitTarget(plan(1), plan(2))
                If Len(nm) > 0 Then found.SetValue nm, True
                If plan(6) > plan(5) Then
                    Dim stepPlan As Variant: stepPlan = DPlan(plan(5), plan(6))
                    If stepPlan(0) = "assign" Then
                        nm = DTargetRoot(stepPlan(1), stepPlan(2))
                        If Len(nm) > 0 Then found.SetValue nm, True
                    End If
                End If
                DCollectAssigned DBlock(plan(7), plan(8)), found
            Case "foreach"
                nm = DEachName(plan(2), plan(3))
                If Len(nm) > 0 Then found.SetValue nm, True
                DCollectAssigned DBlock(plan(6), plan(7)), found
            Case "while"
                DCollectAssigned DBlock(plan(3), plan(4)), found
            Case "switch"
                For Each item In plan(3)
                    DCollectAssigned DBlock(item(2), item(3)), found
                Next item
                If plan(4) Then DCollectAssigned DBlock(plan(5), plan(6)), found
            Case "try"
                DCollectAssigned DBlock(plan(1), plan(2)), found
                If Len(plan(3)) > 0 Then found.SetValue CStr(plan(3)), True
                DCollectAssigned DBlock(plan(4), plan(5)), found
            Case "destr"
                For Each item In plan(1)
                    found.SetValue CStr(item), True
                Next item
                If plan(2) Then found.SetValue CStr(plan(3)), True
            Case "class"
                Set info = plan(1)
                found.SetValue CStr(info.GetValue("name")), True
            Case "import"
                For Each item In DImportNames(plan(1))
                    found.SetValue CStr(item), True
                Next item
        End Select
    Next plan
End Sub

' The variable a for's first part assigns: x = v, outside a class body
Private Function DInitTarget(ByVal s As Long, ByVal e As Long) As String
    If raw_ Then Exit Function
    Dim at As Long, s2 As Long, e2 As Long
    s2 = s: e2 = e
    DUnwrap s2, e2
    If DTopOperator(s2, e2, at) <> "=" Then Exit Function
    If Not DAssignable(s2, at) Then Exit Function
    DInitTarget = DTargetRoot(s2, at)
End Function

' A variable, an index or a property
Private Function DAssignable(ByVal s As Long, ByVal e As Long) As Boolean
    DUnwrap s, e
    Dim root As String, ops As Collection
    If Not DChain(s, e, root, ops) Then Exit Function
    If ops.Count = 0 Then
        DAssignable = (Len(root) > 0)
    Else
        Dim last As Variant: last = ops(ops.Count)
        DAssignable = (last(0) <> "C")
    End If
End Function

' The variable of for (x of ...): a plain name
Private Function DEachName(ByVal s As Long, ByVal e As Long) As String
    DUnwrap s, e
    Dim root As String, ops As Collection
    If DChain(s, e, root, ops) Then
        If ops.Count = 0 Then DEachName = root
    End If
End Function

' The names an import binds: the default, the namespace, then the named ones (ImportNames)
Private Function DImportNames(ByVal info As ASF_Map) As Collection
    Dim names As Collection: Set names = New Collection
    If info.Exists("defaultImport") Then names.Add CStr(info.GetValue("defaultImport"))
    If info.Exists("namespaceImport") Then names.Add CStr(info.GetValue("namespaceImport"))
    Dim item As Variant
    For Each item In info.GetValue("named")
        names.Add CStr(item(1))
    Next item
    Set DImportNames = names
End Function

' Walks plans in the order of their parts: names go to used, function
' literals are analysed under ctx when withLiterals
Private Sub DScan(plans As Collection, used As ASF_Map, ctx As ASF_Map, ByVal withLiterals As Boolean)
    Dim saveMode As Long, saveUsed As ASF_Map, saveCtx As ASF_Map, saveLits As Boolean
    saveMode = mode_: Set saveUsed = scanUsed_: Set saveCtx = scanCtx_: saveLits = scanLits_
    mode_ = 1: Set scanUsed_ = used: Set scanCtx_ = ctx: scanLits_ = withLiterals
    DScanPlans plans
    mode_ = saveMode: Set scanUsed_ = saveUsed: Set scanCtx_ = saveCtx: scanLits_ = saveLits
End Sub

Private Sub DScanPlans(plans As Collection)
    Dim plan As Variant, item As Variant, info As ASF_Map, bounds As Variant
    For Each plan In plans
        Select Case plan(0)
            Case "expr", "throw", "return": DExpr plan(1), plan(2)
            Case "assign"
                DExpr plan(1), plan(2)
                If plan(3) <> "=" Then DExpr plan(1), plan(2)
                DExpr plan(4), plan(5)
            Case "destr"
                For Each item In plan(1)
                    scanUsed_.SetValue CStr(item), True
                Next item
                DExpr plan(4), plan(5)
            Case "print"
                For Each item In plan(1)
                    DExpr item(0), item(1)
                Next item
            Case "if"
                DExpr plan(1), plan(2)
                DScanPlans DBlock(plan(3), plan(4))
                If raw_ And plan(7) Then DScanPlans DBlock(plan(8), plan(9))
                For Each item In plan(5)
                    DExpr item(0), item(1)
                Next item
                For Each item In plan(6)
                    DScanPlans DBlock(item(0), item(1))
                Next item
                If Not raw_ And plan(7) Then DScanPlans DBlock(plan(8), plan(9))
            Case "for"
                DExpr plan(1), plan(2)
                DExpr plan(3), plan(4)
                If plan(6) > plan(5) Then
                    Dim stepPlans As Collection: Set stepPlans = New Collection
                    stepPlans.Add DPlan(plan(5), plan(6))
                    DScanPlans stepPlans
                End If
                DScanPlans DBlock(plan(7), plan(8))
            Case "foreach"
                DExpr plan(2), plan(3)
                DExpr plan(4), plan(5)
                DScanPlans DBlock(plan(6), plan(7))
            Case "while"
                DExpr plan(1), plan(2)
                DScanPlans DBlock(plan(3), plan(4))
            Case "try"
                DScanPlans DBlock(plan(1), plan(2))
                DScanPlans DBlock(plan(4), plan(5))
            Case "switch"
                DExpr plan(1), plan(2)
                For Each item In plan(3)
                    DExpr item(0), item(1)
                    DScanPlans DBlock(item(2), item(3))
                Next item
                If plan(4) Then DScanPlans DBlock(plan(5), plan(6))
            Case "export"
                Set info = plan(1)
                If info.Exists("default") Then
                    bounds = info.GetValue("default")
                    DExpr bounds(0), bounds(1)
                End If
        End Select
    Next plan
End Sub

' The scope of a function literal at token litStart: its locals, and what it takes from enclosing functions
Private Sub DAnalyzeLiteral(ByVal litStart As Long, params As Collection, ByVal rest As String, _
                            ByVal bodyS As Long, ByVal bodyE As Long, parent As ASF_Map)
    Dim ctx As ASF_Map: Set ctx = NewContext("literal", parent)
    ctx.SetValue "tok", litStart
    ctx.SetValue "bodyS", bodyS
    ctx.SetValue "bodyE", bodyE
    Dim i As Long
    For i = 1 To params.Count
        If Len(CStr(params(i))) > 0 Then AddLocal ctx, CStr(params(i))
    Next i
    ctx.SetValue "nParams", ctx.GetValue("nL")
    If Len(rest) > 0 Then ctx.SetValue "restSlot", AddLocal(ctx, rest)
    Dim saveRaw As Boolean: saveRaw = raw_
    raw_ = False
    Dim plans As Collection: Set plans = DBlock(bodyS, bodyE)
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    DCollectAssigned plans, assigned
    For i = 1 To assigned.Count
        If Not EnclosingHas(parent, assigned.key(i)) And Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
    Next i
    literals_.Add ctx
    Dim used As ASF_Map: Set used = New ASF_Map
    DScan plans, used, ctx, True
    raw_ = saveRaw
    Dim locals As ASF_Map: Set locals = ctx.GetValue("locals")
    If used.Exists("this") Then
        AddLocal ctx, "this"
        ResolveFreeAs ctx, "this", "%outerthis"
    End If
    For i = 1 To used.Count
        If Not locals.Exists(used.key(i)) Then ResolveFree ctx, used.key(i)
    Next i
End Sub

Private Function DLiteralContext(ByVal litStart As Long) As ASF_Map
    Dim ctx As Variant
    For Each ctx In literals_
        If ctx.Exists("tok") Then
            If ctx.GetValue("tok") = litStart And ctx.GetValue("parent") Is cx_ Then
                Set DLiteralContext = ctx
                Exit Function
            End If
        End If
    Next ctx
    CompileError "a function literal the analysis did not see"
End Function

' ---------------------------------------------------------------- statements
Private Sub DEmitPlans(plans As Collection)
    Dim plan As Variant
    For Each plan In plans
        DEmitPlan plan
    Next plan
End Sub

Private Sub DEmitPlan(plan As Variant)
    Dim item As Variant, n As Long
    Select Case plan(0)
        Case "assign": DEmitAssign plan(1), plan(2), CStr(plan(3)), plan(4), plan(5)
        Case "expr"
            Dim s As Long, e As Long, at As Long
            s = plan(1): e = plan(2)
            DUnwrap s, e
            ' (x = v) is an assignment too
            If Not raw_ Then
                If DTopOperator(s, e, at) = "=" Then
                    If DAssignable(s, at) Then
                        DEmitAssign s, at, "=", at + 1, e
                        Exit Sub
                    End If
                End If
            End If
            DExpr plan(1), plan(2)
            Emit OP_POP, 0, 0, -1
        Case "if": DEmitIf plan
        Case "for": DEmitFor plan
        Case "foreach": DEmitForEach plan
        Case "while": DEmitWhile plan
        Case "return"
            If cx_.Exists("ctor") Then
                If plan(2) > plan(1) Then
                    DExpr plan(1), plan(2)
                    Emit OP_POP, 0, 0, -1
                End If
                EmitLoadName "this"
                Emit OP_RET, 0, 0, -1
            ElseIf plan(2) <= plan(1) Then
                Emit OP_RET0, 0, 0, 0
            Else
                DExpr plan(1), plan(2)
                Emit OP_RET, 0, 0, -1
            End If
        Case "print"
            For Each item In plan(1)
                DExpr item(0), item(1)
                n = n + 1
            Next item
            Emit OP_PRINT, n, 0, -n
        Case "break"
            ' outside a loop, as in a try or a catch, it does nothing
            If breaks_.Count = 0 Then Exit Sub
            If Not breaks_(breaks_.Count) Is Nothing Then
                Emit OP_JMP, 0, 0, 0
                breaks_(breaks_.Count).Add nOp_
            End If
        Case "continue"
            If conts_.Count = 0 Then Exit Sub
            If Not conts_(conts_.Count) Is Nothing Then
                Emit OP_JMP, 0, 0, 0
                conts_(conts_.Count).Add nOp_
            End If
        Case "switch": DEmitSwitch plan
        Case "destr"
            DExpr plan(4), plan(5)
            Dim hasRest As Long
            If plan(2) Then hasRest = 1
            Emit OP_UNPACK, plan(1).Count, hasRest, plan(1).Count + hasRest - 1
            For Each item In plan(1)
                EmitStoreName CStr(item)
            Next item
            If hasRest = 1 Then EmitStoreName CStr(plan(3))
        Case "class": DEmitClassDecl plan(1)
        Case "try": DEmitTry plan
        Case "import": DEmitImport plan(1)
        Case "proto": DEmitPrototype plan(1)
        Case "export": DEmitExport plan(1)
        Case "throw"
            If plan(2) > plan(1) Then DExpr plan(1), plan(2) Else Emit OP_PUSHE, 0, 0, 1
            Emit OP_THROW, 0, 0, -1
        Case Else
            CompileError "statement " & plan(0)
    End Select
End Sub

' target op value: x op= v is x = x op v; x = x + constant is one instruction
Private Sub DEmitAssign(ByVal ts As Long, ByVal te As Long, ByVal op As String, ByVal vs As Long, ByVal ve As Long)
    Dim start As Long: start = nOp_
    If op = "=" Then
        DExpr vs, ve
    Else
        Dim lStart As Long, lEnd As Long, inner As String
        inner = Left$(op, Len(op) - 1)
        lStart = nOp_ + 1
        DExpr ts, te
        lEnd = nOp_
        If DPrec(inner) = 0 Or BinaryOp(inner) = 0 Then CompileError "Unknown operator " & inner
        DExpr vs, ve
        DBinaryCode inner, lStart, lEnd
    End If
    If nOp_ - start <> 2 Then
        DAssignTo ts, te
        Exit Sub
    End If
    If op_(start + 2) = OP_ADDK Then
        Dim s2 As Long, e2 As Long, root As String, ops As Collection, slot As Long, kind As String
        s2 = ts: e2 = te
        DUnwrap s2, e2
        If DChain(s2, e2, root, ops) Then
            If ops.Count = 0 And Len(root) > 0 Then
                kind = NameKind(root, slot)
                If (kind = "L" And op_(start + 1) = OP_LOADL And a_(start + 1) = slot) Or _
                   (kind = "G" And op_(start + 1) = OP_LOADG And a_(start + 1) = slot) Then
                    If kind = "L" Then op_(start + 1) = OP_INCL Else op_(start + 1) = OP_INCG
                    b_(start + 1) = a_(start + 2)
                    dl_(start + 1) = 0
                    nOp_ = start + 1
                    depth_ = depth_ - 1
                    Exit Sub
                End If
            End If
        End If
    End If
    DAssignTo ts, te
End Sub

Private Sub DEmitIf(plan As Variant)
    Dim ends As Collection: Set ends = New Collection
    Dim falses As Collection: Set falses = New Collection
    Dim i As Long, conds As Collection, blocks As Collection, c As Variant, b As Variant
    Set conds = plan(5): Set blocks = plan(6)
    DJumpIfFalse plan(1), plan(2), falses
    DEmitPlans DBlock(plan(3), plan(4))
    Emit OP_JMP, 0, 0, 0: ends.Add nOp_
    PatchAll falses, nOp_ + 1
    For i = 1 To conds.Count
        Set falses = New Collection
        c = conds(i): b = blocks(i)
        DJumpIfFalse c(0), c(1), falses
        DEmitPlans DBlock(b(0), b(1))
        Emit OP_JMP, 0, 0, 0: ends.Add nOp_
        PatchAll falses, nOp_ + 1
    Next i
    If plan(7) Then DEmitPlans DBlock(plan(8), plan(9))
    PatchAll ends, nOp_ + 1
End Sub

Private Sub DEmitWhile(plan As Variant)
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    DJumpIfFalse plan(1), plan(2), exits
    OpenLoop
    DEmitPlans DBlock(plan(3), plan(4))
    Emit OP_JMP, loopTop, 0, 0
    PatchAll conts_(conts_.Count), loopTop
    PatchAll exits, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

' for (init, cond, step) { }: the step after the body
Private Sub DEmitFor(plan As Variant)
    Dim s As Long, e As Long, at As Long
    s = plan(1): e = plan(2)
    DUnwrap s, e
    Dim isAssign As Boolean
    If Not raw_ And e > s Then
        If DTopOperator(s, e, at) = "=" Then isAssign = DAssignable(s, at)
    End If
    If isAssign Then
        DEmitAssign s, at, "=", at + 1, e
    Else
        DExpr plan(1), plan(2)
        Emit OP_POP, 0, 0, -1
    End If
    Dim loopTop As Long: loopTop = nOp_ + 1
    Dim exits As Collection: Set exits = New Collection
    DJumpIfFalse plan(3), plan(4), exits
    OpenLoop
    DEmitPlans DBlock(plan(7), plan(8))
    PatchAll conts_(conts_.Count), nOp_ + 1
    If plan(6) > plan(5) Then
        Dim stepPlan As Variant: stepPlan = DPlan(plan(5), plan(6))
        Select Case stepPlan(0)
            Case "assign", "expr", "print", "if", "for", "while", "return", "break", "continue"
                DEmitPlan stepPlan
            Case Else
                DExpr plan(5), plan(6)
                Emit OP_POP, 0, 0, -1
        End Select
    End If
    Emit OP_JMP, loopTop, 0, 0
    PatchAll exits, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

Private Sub DEmitForEach(plan As Variant)
    Dim list As String: list = NewTemp()
    Dim counter As String: counter = NewTemp()
    DExpr plan(4), plan(5)
    If plan(1) = "in" Then Emit OP_ITERLIST, 1, 0, 0 Else Emit OP_ITERLIST, 0, 0, 0
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
    If plan(3) > plan(2) Then DAssignTo plan(2), plan(3) Else Emit OP_POP, 0, 0, -1
    OpenLoop
    DEmitPlans DBlock(plan(6), plan(7))
    Emit OP_JMP, loopTop, 0, 0
    PatchAll conts_(conts_.Count), loopTop
    Patch exitJump, nOp_ + 1
    PatchAll breaks_(breaks_.Count), nOp_ + 1
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
End Sub

Private Sub DEmitSwitch(plan As Variant)
    Dim value As String: value = NewTemp()
    DExpr plan(1), plan(2)
    EmitStoreName value
    Dim ends As Collection: Set ends = New Collection
    Dim item As Variant, nextCase As Long
    breaks_.Add ends
    conts_.Add Nothing
    For Each item In plan(3)
        EmitLoadName value
        DExpr item(0), item(1)
        Emit OP_JNSW, 0, 0, -2
        nextCase = nOp_
        DEmitPlans DBlock(item(2), item(3))
        Emit OP_JMP, 0, 0, 0
        ends.Add nOp_
        Patch nextCase, nOp_ + 1
    Next item
    If plan(4) Then DEmitPlans DBlock(plan(5), plan(6))
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
    PatchAll ends, nOp_ + 1
End Sub

Private Sub DEmitTry(plan As Variant)
    Dim start As Long: start = nOp_ + 1
    Dim depth As Long: depth = depth_
    breaks_.Add Nothing
    conts_.Add Nothing
    DEmitPlans DBlock(plan(1), plan(2))
    Emit OP_JMP, 0, 0, 0
    Dim skip As Long: skip = nOp_
    nTry_ = nTry_ + 1
    If nTry_ > UBound(tryStart_) Then
        ReDim Preserve tryStart_(1 To 2 * nTry_): ReDim Preserve tryEnd_(1 To 2 * nTry_)
        ReDim Preserve tryHandler_(1 To 2 * nTry_): ReDim Preserve tryDepth_(1 To 2 * nTry_)
    End If
    tryStart_(nTry_) = start: tryEnd_(nTry_) = nOp_: tryHandler_(nTry_) = nOp_ + 1: tryDepth_(nTry_) = depth
    pendingAdj_ = depth + 1 - depth_
    depth_ = depth + 1
    If depth_ > maxDepth_ Then maxDepth_ = depth_
    If Len(plan(3)) > 0 Then
        EmitStoreName CStr(plan(3))
    Else
        Emit OP_POP, 0, 0, -1
    End If
    If plan(5) > plan(4) Then DEmitPlans DBlock(plan(4), plan(5))
    breaks_.Remove breaks_.Count
    conts_.Remove conts_.Count
    Patch skip, nOp_ + 1
End Sub

' ---------------------------------------------------------------- modules and classes
Private Sub DEmitImport(ByVal parsed As ASF_Map)
    If cx_.GetValue("kind") <> "main" Then CompileError "import goes at the top level of a script"
    Dim names As Collection: Set names = DImportNames(parsed)
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "source", CStr(parsed.GetValue("source"))
    If parsed.Exists("defaultImport") Then info.SetValue "defaultImport", True
    If parsed.Exists("namespaceImport") Then info.SetValue "namespaceImport", True
    Dim named As Collection: Set named = New Collection
    Dim item As Variant
    For Each item In parsed.GetValue("named")
        named.Add CStr(item(0))
    Next item
    info.SetValue "named", named
    Emit OP_IMPORT, ObjectConstant(info), 0, names.Count
    Dim i As Long
    For i = names.Count To 1 Step -1
        EmitStoreName names(i)
    Next i
End Sub

Private Sub DEmitExport(ByVal parsed As ASF_Map)
    Dim item As Variant
    If parsed.Exists("default") Then
        item = parsed.GetValue("default")
        DExpr item(0), item(1)
        Emit OP_EXPORT, Constant("__default"), 0, -1
    ElseIf parsed.Exists("named") Then
        For Each item In parsed.GetValue("named")
            EmitLoadName CStr(item(0))
            Emit OP_EXPORT, Constant(CStr(item(1))), 0, -1
        Next item
    ElseIf parsed.Exists("declaration") Then
        If isModule_ Then pendingExports_.Add CStr(parsed.GetValue("declaration"))
    End If
End Sub

Private Sub DEmitPrototype(ByVal parsed As ASF_Map)
    Dim info As ASF_Map: Set info = New ASF_Map
    info.SetValue "classType", parsed.GetValue("classType")
    info.SetValue "methodName", parsed.GetValue("methodName")
    info.SetValue "funcIndex", CLng(parsed.GetValue("funcIndex"))
    Dim params As Collection: Set params = New Collection
    Dim item As Variant
    For Each item In parsed.GetValue("params")
        params.Add CStr(item)
    Next item
    info.SetValue "params", params
    info.SetValue "internalFuncName", parsed.GetValue("internalFuncName")
    Emit OP_MKFUNC, NamedFunction(CStr(parsed.GetValue("internalFuncName"))), 0, 1
    If parsed.Exists("exportName") Then
        info.SetValue "exportName", parsed.GetValue("exportName")
        Emit OP_EXPORTPROTO, ObjectConstant(info), 0, -1
    Else
        Emit OP_PROTO, ObjectConstant(info), 0, -1
    End If
End Sub

' class Name { }: the class value, and functions for its constructor and methods
Private Sub DEmitClassDecl(ByVal decl As ASF_Map)
    If cx_.GetValue("kind") <> "main" Then CompileError "A class is declared at the top level of a script"
    Dim name As String: name = decl.GetValue("name")
    If Not classes_.Exists(name) Then CompileError "A class is declared at the top level of a script"
    If Not classes_.GetValue(name) Is decl Then CompileError "A class is declared at the top level of a script"
    ' a parent this script does not declare is a class value at run time: one a module made
    If decl.Exists("superClass") Then
        If Not classes_.Exists(CStr(decl.GetValue("superClass"))) Then decl.SetValue "externalParent", True
    End If
    nCls_ = nCls_ + 1
    If nCls_ > UBound(clsNew_) Then
        ReDim Preserve clsNew_(1 To 2 * nCls_): ReDim Preserve clsInit_(1 To 2 * nCls_): ReDim Preserve clsCtor_(1 To 2 * nCls_)
    End If
    clsNew_(nCls_) = NewFunction("@new " & name)
    pendingCls_.Add Array("new", decl, decl, clsNew_(nCls_))
    clsInit_(nCls_) = 0: clsCtor_(nCls_) = 0
    If isModule_ Then
        ' a class of another script may extend it: its parts run on its instances
        clsInit_(nCls_) = NewFunction("@init " & name)
        pendingCls_.Add Array("init", decl, decl, clsInit_(nCls_))
        If decl.Exists("constructor") Then clsCtor_(nCls_) = DConstructorFunction(decl)
    End If
    Dim cls As ASF_Map: Set cls = New ASF_Map
    cls.Add "type", "Class"
    cls.SetValue "name", name
    Dim fields As Collection: Set fields = New Collection
    Dim item As Variant
    For Each item In decl.GetValue("fields")
        fields.Add CStr(item(0))
    Next item
    cls.SetValue "fields", fields
    Dim descs As Collection: Set descs = New Collection
    Dim m As Variant, fn As Long
    For Each m In decl.GetValue("classMethods")
        fn = NewFunction(name & "." & m.GetValue("name"))
        methodFns_.Add Array(m, fn)
        pendingCls_.Add Array("method", decl, m, fn)
        descs.Add DMethodDescriptor(m, fn)
    Next m
    cls.SetValue "classMethods", descs
    If decl.Exists("constructor") Then cls.SetValue "constructor", DMethodDescriptor(decl.GetValue("constructor"), 0)
    If decl.Exists("superClass") Then cls.SetValue "superClass", decl.GetValue("superClass")
    Emit OP_CLASS, ObjectConstant(cls), nCls_, 1
    EmitStoreName name
End Sub

Private Function DMethodDescriptor(ByVal m As ASF_Map, ByVal fn As Long) As ASF_Map
    Dim d As ASF_Map: Set d = New ASF_Map
    d.Add "type", "Method"
    d.SetValue "name", CStr(m.GetValue("name"))
    Dim params As Collection: Set params = New Collection
    Dim item As Variant
    For Each item In m.GetValue("params")
        params.Add CStr(item)
    Next item
    d.SetValue "params", params
    d.SetValue "isStatic", CBool(m.GetValue("isStatic"))
    If fn > 0 Then d.SetValue "fn", fn
    Set DMethodDescriptor = d
End Function

Private Function DConstructorFunction(ByVal cls As ASF_Map) As Long
    Dim name As String: name = cls.GetValue("name")
    If ctorFn_.Exists(name) Then
        DConstructorFunction = ctorFn_.GetValue(name)
    Else
        DConstructorFunction = NewFunction("@ctor " & name)
        ctorFn_.SetValue name, DConstructorFunction
        pendingCls_.Add Array("ctor", cls, cls.GetValue("constructor"), DConstructorFunction)
    End If
End Function

' super(args): the parent's constructor on the same this
Private Sub DEmitSuperCall(args As Collection)
    Dim n As Long, cls As ASF_Map, parent As ASF_Map, a As Variant
    n = args.Count
    If cx_.Exists("class") Then Set cls = cx_.GetValue("class")
    If Not cls Is Nothing Then
        If cls.Exists("externalParent") Then
            ' the parent is a class of a module: its constructor is found at run time
            EmitLoadName CStr(cls.GetValue("superClass"))
            EmitLoadName "this"
            For Each a In args
                DExpr a(0), a(1)
            Next a
            Emit OP_SUPER, 1, n, -n - 1
            Emit OP_POP, 0, 0, -1
            Emit OP_PUSHE, 0, 0, 1
            Exit Sub
        End If
        Set parent = DParentOf(cls)
    End If
    If Not parent Is Nothing Then
        If Not parent.Exists("constructor") Then Set parent = Nothing
    End If
    If parent Is Nothing Then
        For Each a In args
            DExpr a(0), a(1)
            Emit OP_POP, 0, 0, -1
        Next a
    Else
        EmitLoadName "this"
        For Each a In args
            DExpr a(0), a(1)
        Next a
        Emit OP_CALLFT, DConstructorFunction(parent), n, -n
        Emit OP_POP, 0, 0, -1
    End If
    Emit OP_PUSHE, 0, 0, 1
End Sub

' ---------------------------------------------------------------- functions
Private Function DCompileProgram(ByVal progName As String) As Long
    Set gSlot_ = New ASF_Map
    ReDim gNames_(1 To 16): nG_ = 0: gBase_ = nGTotal_
    isModule_ = (Left$(progName, 8) = "@module:")
    Set pendingExports_ = New Collection
    Set fnOf_ = New ASF_Map
    Set pending_ = New Collection
    Set classes_ = New ASF_Map
    Set ctorFn_ = New ASF_Map
    Set pendingCls_ = New Collection
    Set methodFns_ = New Collection
    Set tplParts_ = New ASF_Map
    ReDim blockC_(1 To nT_ + 1)
    raw_ = False: mode_ = 0: dNamed_ = 0
    Dim stmts As Collection: Set stmts = DTopLevel(1, nT_ + 1)
    Dim plan As Variant, decl As ASF_Map
    For Each plan In stmts
        If plan(0) = "class" Then
            Set decl = plan(1)
            If classes_.Exists(CStr(decl.GetValue("name"))) Then CompileError "Class '" & decl.GetValue("name") & "' is already declared"
            classes_.SetValue CStr(decl.GetValue("name")), decl
        End If
    Next plan
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    DCollectAssigned stmts, assigned
    Dim i As Long
    For i = 1 To assigned.Count
        GlobalSlot assigned.key(i)
    Next i
    Set assignedGlobals_ = assigned
    Set literals_ = New Collection
    Set pendingLits_ = New Collection
    Dim mainFn As Long: mainFn = NewFunction("@main")
    Dim ctx As ASF_Map: Set ctx = NewContext("main", Nothing)
    If DLiterals(1, nT_ + 1) > dNamed_ Then DScan stmts, New ASF_Map, ctx, True
    BeginFunction ctx
    directFn_ = True
    fnEntry_(mainFn) = nOp_ + 1
    DEmitPlans stmts
    Dim nm As Variant
    For Each nm In pendingExports_
        EmitLoadName CStr(nm)
        Emit OP_EXPORT, Constant(CStr(nm)), 0, -1
    Next nm
    Emit OP_HALT, 0, 0, 0
    EndFunction mainFn, 0
    DrainFunctions
    DCheckFunctions
    Dim names() As String
    ReDim names(0 To nG_)
    For i = 1 To nG_
        names(i) = gNames_(i)
    Next i
    dPendingGlobals_ = names
    nGTotal_ = gBase_ + nG_
    DCompileProgram = mainFn
End Function

' A named function is compiled when code uses it, but every one must compile: the bodies of
' the others are read now, for the errors it would find
Private Sub DCheckFunctions()
    Dim i As Long, info As ASF_Map
    For i = 1 To dFuncOrder_.Count
        If Not fnOf_.Exists(CStr(dFuncOrder_(i))) Then
            Set info = dFuncs_.GetValue(CStr(dFuncOrder_(i)))
            Set literals_ = New Collection
            DScan DBlock(info.GetValue("bodyS"), info.GetValue("bodyE")), New ASF_Map, NewContext("named", Nothing), True
        End If
    Next i
End Sub

' The fun tokens of tokens s..e-1, and more when a template may hold one
Private Function DLiterals(ByVal s As Long, ByVal e As Long) As Long
    Dim j As Long
    For j = s To e - 1
        If tt_(j) = "IDENT" Then
            If LCase$(tv_(j)) = "fun" Then DLiterals = DLiterals + 1
        ElseIf tt_(j) = "STRING" Then
            If InStr(tv_(j), "${") > 0 Then DLiterals = DLiterals + nT_
        End If
    Next j
End Function

' True when tokens s..e-1 may name word (a template may)
Private Function DHasWord(ByVal s As Long, ByVal e As Long, ByVal word As String) As Boolean
    Dim j As Long
    For j = s To e - 1
        If tt_(j) = "IDENT" Then
            If tv_(j) = word Then
                DHasWord = True
                Exit Function
            End If
        ElseIf tt_(j) = "STRING" Then
            If InStr(tv_(j), "${") > 0 Then
                DHasWord = True
                Exit Function
            End If
        End If
    Next j
End Function

' The functions the program's code refers to, compiled in the order CompileProgram compiles them
Private Sub DrainFunctions()
    Dim name As String, work As Variant, ctx As ASF_Map
    Do While pending_.Count > 0 Or pendingLits_.Count > 0 Or pendingCls_.Count > 0
        If pending_.Count > 0 Then
            name = pending_(1)
            pending_.Remove 1
            CompileNamedFunction name
        ElseIf pendingCls_.Count > 0 Then
            work = pendingCls_(1)
            pendingCls_.Remove 1
            DCompileClassFunction work
        Else
            Set ctx = pendingLits_(1)
            pendingLits_.Remove 1
            DCompileLiteral ctx
        End If
    Loop
End Sub

' A named function compiled from its tokens
Private Sub DCompileNamed(ByVal name As String, ByVal params As Variant, ByVal rest As String, ByVal bodyS As Long, ByVal bodyE As Long)
    Dim fn As Long: fn = fnOf_.GetValue(name)
    Dim ctx As ASF_Map: Set ctx = NewContext("named", Nothing)
    Dim i As Long
    If IsArray(params) Then
        For i = LBound(params) To UBound(params)
            If Len(CStr(params(i))) > 0 Then AddLocal ctx, CStr(params(i))
        Next i
    End If
    Dim nParams As Long: nParams = ctx.GetValue("nL")
    If Len(rest) > 0 Then fnRest_(fn) = AddLocal(ctx, rest)
    Dim saveRaw As Boolean: saveRaw = raw_
    raw_ = False
    Dim plans As Collection: Set plans = DBlock(bodyS, bodyE)
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    DCollectAssigned plans, assigned
    For i = 1 To assigned.Count
        If Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
    Next i
    Dim used As ASF_Map: Set used = New ASF_Map
    If DHasWord(bodyS, bodyE, "this") Then DScan plans, used, ctx, False
    If used.Exists("this") Then fnThis_(fn) = AddLocal(ctx, "this")
    If DLiterals(bodyS, bodyE) > 0 Then DScan plans, New ASF_Map, ctx, True
    BeginFunction ctx
    directFn_ = True
    fnEntry_(fn) = nOp_ + 1
    DEmitPlans plans
    Emit OP_RET0, 0, 0, 0
    EndFunction fn, nParams
    raw_ = saveRaw
End Sub

Private Sub DCompileLiteral(ctx As ASF_Map)
    Dim fn As Long: fn = ctx.GetValue("fn")
    Dim locals As ASF_Map: Set locals = ctx.GetValue("locals")
    If ctx.Exists("restSlot") Then fnRest_(fn) = ctx.GetValue("restSlot")
    If locals.Exists("this") Then fnThis_(fn) = locals.GetValue("this")
    Dim free As ASF_Map: Set free = ctx.GetValue("free")
    If free.Exists("%outerthis") Then fnThisEnv_(fn) = free.GetValue("%outerthis")
    Dim saveRaw As Boolean: saveRaw = raw_
    raw_ = False
    BeginFunction ctx
    directFn_ = True
    fnEntry_(fn) = nOp_ + 1
    DEmitPlans DBlock(ctx.GetValue("bodyS"), ctx.GetValue("bodyE"))
    Emit OP_RET0, 0, 0, 0
    EndFunction fn, ctx.GetValue("nParams")
    raw_ = saveRaw
End Sub

' A function of a class: "new" (instance, fields, methods, constructor), "ctor" or "method"
Private Sub DCompileClassFunction(ByVal work As Variant)
    Dim cls As ASF_Map: Set cls = work(1)
    Dim fn As Long: fn = work(3)
    Dim ctx As ASF_Map: Set ctx = NewContext("named", Nothing)
    ctx.SetValue "class", cls
    Dim parent As ASF_Map, ctor As ASF_Map, external As Boolean, setup As Boolean
    Set parent = DParentOf(cls)
    external = cls.Exists("externalParent")
    ' "new" and "init" set the fields and bind the methods; "init" runs no constructor
    setup = (work(0) = "new" Or work(0) = "init")
    If work(0) = "new" Then
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
    ElseIf work(0) = "init" Then
        fnKind_(fn) = 1
    Else
        Set ctor = work(2)
        fnKind_(fn) = 1
    End If
    Dim saveRaw As Boolean: saveRaw = raw_
    raw_ = True
    Dim plans As Collection: Set plans = New Collection
    If Not ctor Is Nothing Then Set plans = DBlock(ctor.GetValue("bodyS"), ctor.GetValue("bodyE"))
    Dim i As Long, item As Variant
    If Not ctor Is Nothing Then
        For Each item In ctor.GetValue("params")
            AddLocal ctx, CStr(item)
        Next item
    End If
    Dim nParams As Long: nParams = ctx.GetValue("nL")
    ' new of a class whose parent is a module's, without its own constructor: the arguments go to
    ' the parent's constructor
    Dim restSlot As Long: restSlot = -1
    If work(0) = "new" And external And ctor Is Nothing Then
        restSlot = AddLocal(ctx, "%args")
        fnRest_(fn) = restSlot
    End If
    fnThis_(fn) = AddLocal(ctx, "this")
    Dim assigned As ASF_Map: Set assigned = New ASF_Map
    DCollectAssigned plans, assigned
    For i = 1 To assigned.Count
        If Not gSlot_.Exists(assigned.key(i)) Then AddLocal ctx, assigned.key(i)
    Next i
    If setup Then
        If Not parent Is Nothing Then DScanFields parent, ctx
        DScanFields cls, ctx
    End If
    DScan plans, New ASF_Map, ctx, True
    BeginFunction ctx
    directFn_ = True
    fnEntry_(fn) = nOp_ + 1
    If work(0) = "new" Then
        Emit OP_NEWINST, Constant(CStr(cls.GetValue("name"))), 0, 1
        Emit OP_STOREL, fnThis_(fn), 0, -1
    End If
    If setup Then
        If external Then
            ' the fields and methods of the module's class first
            EmitLoadName CStr(cls.GetValue("superClass"))
            Emit OP_LOADL, fnThis_(fn), 0, 1
            Emit OP_SUPER, 0, 0, -1
            Emit OP_POP, 0, 0, -1
        ElseIf Not parent Is Nothing Then
            DEmitFields parent, fnThis_(fn)
            DEmitBindMethods parent.GetValue("classMethods"), fnThis_(fn), True
        End If
        DEmitFields cls, fnThis_(fn)
        DEmitBindMethods cls.GetValue("classMethods"), fnThis_(fn), (work(0) = "init")
    End If
    DEmitPlans plans
    If restSlot >= 0 Then
        EmitLoadName CStr(cls.GetValue("superClass"))
        Emit OP_LOADL, fnThis_(fn), 0, 1
        Emit OP_LOADL, restSlot, 0, 1
        Emit OP_SUPER, 1, -1, -2
        Emit OP_POP, 0, 0, -1
    End If
    If setup Then
        Emit OP_LOADL, fnThis_(fn), 0, 1
        Emit OP_RET, 0, 0, -1
    Else
        Emit OP_RET0, 0, 0, 0
    End If
    EndFunction fn, nParams
    raw_ = saveRaw
End Sub

' The parent a class declared in this script extends, if this script declares it too
Private Function DParentOf(ByVal cls As ASF_Map) As ASF_Map
    If Not cls.Exists("superClass") Then Exit Function
    If cls.Exists("externalParent") Then Exit Function
    If classes_.Exists(CStr(cls.GetValue("superClass"))) Then Set DParentOf = classes_.GetValue(CStr(cls.GetValue("superClass")))
End Function

' The function literals of a class's field defaults
Private Sub DScanFields(ByVal cls As ASF_Map, ctx As ASF_Map)
    Dim f As Variant, saveMode As Long, saveUsed As ASF_Map, saveCtx As ASF_Map, saveLits As Boolean
    saveMode = mode_: Set saveUsed = scanUsed_: Set saveCtx = scanCtx_: saveLits = scanLits_
    mode_ = 1: Set scanUsed_ = New ASF_Map: Set scanCtx_ = ctx: scanLits_ = True
    For Each f In cls.GetValue("fields")
        If f(1) > 0 Then DExpr f(1), f(2)
    Next f
    mode_ = saveMode: Set scanUsed_ = saveUsed: Set scanCtx_ = saveCtx: scanLits_ = saveLits
End Sub

Private Sub DEmitFields(ByVal cls As ASF_Map, ByVal thisSlot As Long)
    Dim f As Variant
    For Each f In cls.GetValue("fields")
        If f(1) > 0 Then
            DExpr f(1), f(2)
        Else
            Emit OP_PUSHE, 0, 0, 1
        End If
        Emit OP_LOADL, thisSlot, 0, 1
        Emit OP_SETPROPV, Constant(CStr(f(0))), 0, -2
    Next f
End Sub

Private Sub DEmitBindMethods(methods As Collection, ByVal thisSlot As Long, ByVal withStatic As Boolean)
    Dim m As Variant, decl As ASF_Map
    For Each m In methods
        Set decl = m
        If withStatic Or Not decl.GetValue("isStatic") Then
            Emit OP_LOADL, thisSlot, 0, 1
            Emit OP_BINDM, Constant(CStr(decl.GetValue("name"))), MethodFunctionOf(decl), -1
        End If
    Next m
End Sub

' The maximum depth of the code of a function from entry on: the stack effect of each instruction,
' and the catch value pushed before a handler
Private Function DMaxDepth(ByVal entry As Long) As Long
    Dim i As Long, d As Long, m As Long
    For i = entry To nOp_
        d = d + adj_(i)
        If d > m Then m = d
        d = d + dl_(i)
        If d > m Then m = d
    Next i
    DMaxDepth = m
End Function
