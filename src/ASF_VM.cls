VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_VM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ASF_VM: runs the bytecode of ASF_Compiler in one dispatch loop (see BYTECODE.md). Calls are
' frames on arrays, not VBA calls; values live in one preallocated Variant stack.
Option Explicit

' opcodes: the numbers of ASF_Compiler, the order of the jump table in Execute
Private Const OP_ADD As Long = 8
Private Const OP_SHR As Long = 22
Private Const OP_JNLT As Long = 40
Private Const OP_JNNE As Long = 45

Private GLOBALS_ As ASF_Globals
Private RT_ As ASF_Runtime              ' the library of values: methods, printing, typeof, host objects
Private BC_ As ASF_Compiler
Private version_ As Long
Private trace_ As Boolean               ' EnableCallTrace: calls, returns and errors go to the trace

Private op_() As Long, oa_() As Long, ob_() As Long
Private k_() As Variant
Private hint_() As Long                 ' GETPROP's inline cache: where the key was last found
Private fnEntry_() As Long, fnNParams_() As Long, fnNLocals_() As Long, fnMaxStack_() As Long
Private fnCellStart_() As Long, fnCellCount_() As Long, cellSlot_() As Long
Private fnCapStart_() As Long, fnCapCount_() As Long, capKind_() As Long, capIdx_() As Long
Private fnRest_() As Long, fnThis_() As Long, fnThisEnv_() As Long
Private fnKind_() As Long               ' 1: a method, 2: a constructor
Private tryStart_() As Long, tryEnd_() As Long, tryHandler_() As Long, tryDepth_() As Long, nTry_ As Long
Private throwing_ As Boolean            ' a throw goes up the calls; LAST_ERROR_ holds its value
Private mainFn_ As Long
Private clsNew_() As Long, clsMap_() As Variant, nCls_ As Long ' class number -> constructor function, class value
Private clsKeep_() As Boolean, clsCap_ As Long ' the classes of modules stay after a run

' modules: run once per engine, then kept with their globals, classes and the cells made while loading
Private modules_ As ASF_Map             ' path -> {path, exports, defaultExport}
Private loading_ As Collection          ' paths being loaded, for circular imports
Private curExports_ As ASF_Map          ' the exports of the module being loaded
Private globKeep_() As Boolean, globCap_ As Long ' the global slots of modules stay after a run
Private keepCell_ As Long, keepEnv_ As Long     ' cells and environments below these stay too

' the arena of a run: cells (variables closures share) and environments (cell numbers)
Private cell_() As Variant, nCell_ As Long
Private env_() As Long, nEnv_ As Long
Private Const ENV_SHIFT As Double = 1048576     ' a closure's "fn" holds function + environment * ENV_SHIFT
Private Const METHOD_SHIFT As Long = 1024       ' b of CALLM...: arguments + METHOD_SHIFT * method number

Private stk_() As Variant, stkCap_ As Long, stkHigh_ As Long
Private glob_() As Variant
Private frRet_() As Long, frFp_() As Long, frRes_() As Long, frFn_() As Long, frEnv_() As Long, frCap_ As Long
Private frKind_() As Long, frRecv_() As Variant ' kind 1: a method call whose receiver is written back after it
Private math_ As Object
Private topSp_ As Long, topFt_ As Long  ' the stack and frames in use when ASF_Runtime is called

Private OUTPUT__ As Variant
Private LAST_ERROR_ As Variant
Private MathObj_ As Collection

Private Sub Class_Initialize()
    stkCap_ = 1024
    ReDim stk_(1 To stkCap_)
    frCap_ = 64
    ReDim frRet_(1 To frCap_): ReDim frFp_(1 To frCap_): ReDim frRes_(1 To frCap_): ReDim frFn_(1 To frCap_)
    ReDim frEnv_(1 To frCap_): ReDim frKind_(1 To frCap_): ReDim frRecv_(1 To frCap_)
    ReDim cell_(1 To 64): ReDim env_(1 To 64)
    Set MathObj_ = New Collection
    Set modules_ = New ASF_Map
    Set loading_ = New Collection
End Sub

' ClearModuleCache: modules load again, and their values go
Public Sub ClearModules()
    Set modules_ = New ASF_Map
    Dim i As Long
    For i = 1 To globCap_
        globKeep_(i) = False
        glob_(i) = Empty
    Next i
    For i = 1 To clsCap_
        clsKeep_(i) = False
        clsMap_(i) = Empty
    Next i
    For i = 1 To keepCell_
        cell_(i) = Empty
    Next i
    keepCell_ = 0: keepEnv_ = 0
End Sub

Public Sub SetHost(aGlobals As ASF_Globals, aRuntime As ASF_Runtime, aCompiler As ASF_Compiler)
    Set GLOBALS_ = aGlobals
    Set RT_ = aRuntime
    Set BC_ = aCompiler
End Sub

Public Property Get OUTPUT_() As Variant
    If IsObject(OUTPUT__) Then Set OUTPUT_ = OUTPUT__ Else OUTPUT_ = OUTPUT__
End Property

Public Property Get LastError() As Variant
    If IsObject(LAST_ERROR_) Then Set LastError = LAST_ERROR_ Else LastError = LAST_ERROR_
End Property

' ---------------------------------------------------------------- running a program
Public Sub Run(ByVal idx As Long, params As Variant)
    If version_ <> BC_.Version Then LoadCode
    LAST_ERROR_ = Empty
    throwing_ = False
    trace_ = RT_.EnableCallTrace
    BindGlobals idx, params, False
    Dim mainFn As Long: mainFn = BC_.ProgramFunction(idx)
    mainFn_ = mainFn
    If fnMaxStack_(mainFn) >= stkCap_ Then GrowStack fnMaxStack_(mainFn)
    stkHigh_ = fnMaxStack_(mainFn)
    RT_.SetHost Me
    Dim none As Variant
    Execute fnEntry_(mainFn), 0, 1, 0, 0, -1, 0, 0, 0, False, none
    RT_.SetHost Nothing
    ReleaseValues
End Sub

Private Sub LoadCode()
    op_ = BC_.Ops(): oa_ = BC_.ArgsA(): ob_ = BC_.ArgsB()
    k_ = BC_.Constants()
    fnEntry_ = BC_.FunctionEntries(): fnNParams_ = BC_.FunctionParams()
    fnNLocals_ = BC_.FunctionLocals(): fnMaxStack_ = BC_.FunctionStacks()
    fnCellStart_ = BC_.FunctionCellStarts(): fnCellCount_ = BC_.FunctionCellCounts(): cellSlot_ = BC_.CellSlots()
    fnCapStart_ = BC_.FunctionCaptureStarts(): fnCapCount_ = BC_.FunctionCaptureCounts()
    capKind_ = BC_.CaptureKinds(): capIdx_ = BC_.CaptureIndexes()
    fnRest_ = BC_.FunctionRests(): fnThis_ = BC_.FunctionThis(): fnThisEnv_ = BC_.FunctionThisEnv()
    fnKind_ = BC_.FunctionKinds()
    tryStart_ = BC_.TryStarts(): tryEnd_ = BC_.TryEnds(): tryHandler_ = BC_.TryHandlers(): tryDepth_ = BC_.TryDepths()
    nTry_ = BC_.TryCount()
    clsNew_ = BC_.ClassConstructors(): nCls_ = BC_.ClassCount()
    ' a module compiled while a program runs adds code: what the program holds stays
    If clsCap_ = 0 Then
        ReDim clsMap_(1 To UBound(clsNew_)): ReDim clsKeep_(1 To UBound(clsNew_))
    ElseIf UBound(clsNew_) > clsCap_ Then
        ReDim Preserve clsMap_(1 To UBound(clsNew_)): ReDim Preserve clsKeep_(1 To UBound(clsNew_))
    End If
    clsCap_ = UBound(clsNew_)
    If globCap_ = 0 Then
        ReDim glob_(0 To BC_.GlobalCount + 16): ReDim globKeep_(0 To BC_.GlobalCount + 16)
        globCap_ = BC_.GlobalCount + 16
    ElseIf BC_.GlobalCount > globCap_ Then
        ReDim Preserve glob_(0 To 2 * BC_.GlobalCount): ReDim Preserve globKeep_(0 To 2 * BC_.GlobalCount)
        globCap_ = 2 * BC_.GlobalCount
    End If
    ReDim hint_(1 To UBound(op_))
    version_ = BC_.Version
End Sub

' The global slots of program idx: Math, $1, $2... and the injected variables by name; a module
' has Math only
Private Sub BindGlobals(ByVal idx As Long, params As Variant, ByVal isModule As Boolean)
    Dim names As Variant: names = BC_.ProgramGlobals(idx)
    Dim gBase As Long: gBase = BC_.ProgramGlobalBase(idx)
    Dim n As Long: n = UBound(names)
    Dim injected As Collection
    If Not isModule Then Set injected = RT_.InjectedVariables
    Dim i As Long, k As Long, item As Variant, nm As String
    For i = 1 To n
        glob_(gBase + i) = Empty
        globKeep_(gBase + i) = isModule
        nm = names(i)
        If nm = "Math" Then
            Set math_ = RT_.MathObject
            Set glob_(gBase + i) = math_
        ElseIf Left$(nm, 1) = "$" And Not isModule Then
            If IsNumeric(Mid$(nm, 2)) And IsArray(params) Then
                k = CLng(Mid$(nm, 2))
                If k >= 1 And k <= UBound(params) - LBound(params) + 1 Then
                    SetItem glob_(gBase + i), RT_.PrepareParam(params(LBound(params) + k - 1))
                End If
            End If
        End If
        If Not injected Is Nothing Then
            For Each item In injected
                If CStr(item(LBound(item))) = nm Then SetItem glob_(gBase + i), item(UBound(item))
            Next item
        End If
    Next i
End Sub

' Drops the values of the run, so that nothing it built stays alive: the cells go with the arena,
' which breaks the cycles between closures and the variables they share
Private Sub ReleaseValues()
    Dim i As Long
    For i = 1 To stkHigh_
        stk_(i) = Empty
    Next i
    stkHigh_ = 0
    For i = 1 To globCap_
        If Not globKeep_(i) Then glob_(i) = Empty
    Next i
    For i = keepCell_ + 1 To nCell_
        cell_(i) = Empty
    Next i
    nCell_ = keepCell_
    nEnv_ = keepEnv_
    For i = 1 To clsCap_
        If Not clsKeep_(i) Then clsMap_(i) = Empty
    Next i
    Set math_ = Nothing
End Sub

Private Sub SetItem(ByRef dst As Variant, ByRef v As Variant)
    If IsObject(v) Then Set dst = v Else dst = v
End Sub

' ---------------------------------------------------------------- the dispatch loop
' Runs from pc until the program ends, or, for a call from ASF_Runtime (stopFt >= 0), until the
' frame it enters (enterFn with enterN arguments on the stack) returns: its value is the result
Private Function Execute(ByVal pc As Long, ByVal sp As Long, ByVal fp As Long, ByVal ft As Long, ByVal env As Long, _
                         ByVal stopFt As Long, ByVal enterFn As Long, ByVal enterN As Long, ByVal enterEnv As Long, _
                         ByVal hasThis As Boolean, ByRef thisV As Variant) As Variant
    Dim t As Boolean, vt As Long, fn As Long, n As Long, i As Long, newEnv As Long, res As Long, h As Long
    Dim callKind As Long, wrote As Boolean, m As ASF_Map, m2 As ASF_Map
    Dim rv As Variant, rv2 As Variant, thisVal As Variant, args As Variant, errDesc As String, errNum As Long
    Dim result As Variant
    On Error GoTo Failed
    If enterFn > 0 Then
        fn = enterFn: n = enterN: newEnv = enterEnv: res = sp - n
        SetItem thisVal, thisV
        GoTo CallSetup
    End If
Dispatch:
    On op_(pc) GoTo X1, X2, X3, X4, X5, X6, X7, X8, X9, X10, X11, X12, X13, X14, X15, X16, X17, X18, X19, X20, _
        X21, X22, X23, X24, X25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, X36, X37, X38, X39, X40, _
        X41, X42, X43, X44, X45, X46, X47, X48, X49, X50, X51, X52, X53, X54, X55, X56, X57, X58, X59, X60, X61, _
        X62, X63, X64, X65, X66, X67, X68, X69, X70, X71, X72, X73, X74, X75, X76, X77, X78, X79, X80, X81, X82, X83, X84, X85, X86, X87
X1: ' HALT
    GoTo Finished
X2: ' PUSHK
    sp = sp + 1
    stk_(sp) = k_(oa_(pc))
    pc = pc + 1: GoTo Dispatch
X3: ' PUSHE
    sp = sp + 1
    stk_(sp) = Empty
    pc = pc + 1: GoTo Dispatch
X4: ' LOADL
    sp = sp + 1
    If IsObject(stk_(fp + oa_(pc))) Then Set stk_(sp) = stk_(fp + oa_(pc)) Else stk_(sp) = stk_(fp + oa_(pc))
    pc = pc + 1: GoTo Dispatch
X5: ' STOREL
    If IsObject(stk_(sp)) Then Set stk_(fp + oa_(pc)) = stk_(sp) Else stk_(fp + oa_(pc)) = stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X6: ' LOADG
    sp = sp + 1
    If IsObject(glob_(oa_(pc))) Then Set stk_(sp) = glob_(oa_(pc)) Else stk_(sp) = glob_(oa_(pc))
    pc = pc + 1: GoTo Dispatch
X7: ' STOREG
    If IsObject(stk_(sp)) Then Set glob_(oa_(pc)) = stk_(sp) Else glob_(oa_(pc)) = stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X8: ' ADD
    If VarType(stk_(sp)) = vbDouble Then
        If VarType(stk_(sp - 1)) = vbDouble Then
            stk_(sp - 1) = stk_(sp - 1) + stk_(sp)
            sp = sp - 1
            pc = pc + 1: GoTo Dispatch
        End If
    End If
    rv = AddValues(stk_(sp - 1), stk_(sp))
    stk_(sp - 1) = rv
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X9: ' SUB
    stk_(sp - 1) = stk_(sp - 1) - stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X10: ' MUL
    stk_(sp - 1) = stk_(sp - 1) * stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X11: ' DIV
    If stk_(sp) = 0 Then Err.Raise vbObjectError + 2001, "VM.EvalExprNode", "Division by zero"
    stk_(sp - 1) = stk_(sp - 1) / stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X12: ' MOD
    stk_(sp - 1) = stk_(sp - 1) Mod stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X13: ' POW
    stk_(sp - 1) = stk_(sp - 1) ^ stk_(sp)
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X14: ' EQ
    ' two numbers (Integer, Long, Single, Double: % gives a Long) compare directly
    If VarType(stk_(sp)) >= vbInteger And VarType(stk_(sp)) <= vbDouble And VarType(stk_(sp - 1)) >= vbInteger And VarType(stk_(sp - 1)) <= vbDouble Then
        t = (stk_(sp - 1) = stk_(sp))
    Else
        t = SameValues(stk_(sp - 1), stk_(sp))
    End If
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X15: ' NE
    ' two numbers (Integer, Long, Single, Double: % gives a Long) compare directly
    If VarType(stk_(sp)) >= vbInteger And VarType(stk_(sp)) <= vbDouble And VarType(stk_(sp - 1)) >= vbInteger And VarType(stk_(sp - 1)) <= vbDouble Then
        t = (stk_(sp - 1) <> stk_(sp))
    Else
        t = Not SameValues(stk_(sp - 1), stk_(sp))
    End If
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X16: ' LT
    t = (stk_(sp - 1) < stk_(sp))
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X17: ' LE
    t = (stk_(sp - 1) <= stk_(sp))
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X18: ' GT
    t = (stk_(sp - 1) > stk_(sp))
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X19: ' GE
    t = (stk_(sp - 1) >= stk_(sp))
    stk_(sp - 1) = t
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X20: ' CAT
    rv = TextOf(stk_(sp - 1)) & TextOf(stk_(sp))
    stk_(sp - 1) = rv
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X21: ' SHL
    rv = ShiftValues(stk_(sp - 1), stk_(sp), True)
    stk_(sp - 1) = rv
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X22: ' SHR
    rv = ShiftValues(stk_(sp - 1), stk_(sp), False)
    stk_(sp - 1) = rv
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X23: ' NOT
    t = Not Truthy(stk_(sp))
    stk_(sp) = t
    pc = pc + 1: GoTo Dispatch
X24: ' NEG
    stk_(sp) = -CDbl(stk_(sp))
    pc = pc + 1: GoTo Dispatch
X25: ' TYPEOF
    rv = RT_.TypeOfValue(stk_(sp))
    stk_(sp) = rv
    pc = pc + 1: GoTo Dispatch
X26: ' JMP
    pc = oa_(pc): GoTo Dispatch
X27: ' JMPF
    vt = VarType(stk_(sp))
    If vt = vbBoolean Then
        t = stk_(sp)
    Else
        t = Truthy(stk_(sp))
    End If
    sp = sp - 1
    If t Then pc = pc + 1 Else pc = oa_(pc)
    GoTo Dispatch
X28: ' JMPT
    vt = VarType(stk_(sp))
    If vt = vbBoolean Then
        t = stk_(sp)
    Else
        t = Truthy(stk_(sp))
    End If
    sp = sp - 1
    If t Then pc = oa_(pc) Else pc = pc + 1
    GoTo Dispatch
X29: ' ANDJ
    If Truthy(stk_(sp)) Then
        sp = sp - 1
        pc = pc + 1
    Else
        stk_(sp) = False
        pc = oa_(pc)
    End If
    GoTo Dispatch
X30: ' ORJ
    If Truthy(stk_(sp)) Then
        stk_(sp) = True
        pc = oa_(pc)
    Else
        sp = sp - 1
        pc = pc + 1
    End If
    GoTo Dispatch
X31: ' TOBOOL
    t = Truthy(stk_(sp))
    stk_(sp) = t
    pc = pc + 1: GoTo Dispatch
X32: ' NULLJ
    If IsNull(stk_(sp)) Or IsEmpty(stk_(sp)) Then
        sp = sp - 1
        pc = pc + 1
    Else
        pc = oa_(pc)
    End If
    GoTo Dispatch
X33: ' POP
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X34: ' CALLF a: function, b: arguments
    fn = oa_(pc): n = ob_(pc): res = sp - n + 1: newEnv = 0
    GoTo CallSetup
X35: ' RET
    If ft = 0 Then
        SetItem OUTPUT__, stk_(sp)
        GoTo Finished
    End If
    If IsObject(stk_(sp)) Then Set rv = stk_(sp) Else rv = stk_(sp)
    GoTo Returning
X36: ' RET0
    If ft = 0 Then
        OUTPUT__ = Empty
        GoTo Finished
    End If
    rv = Empty
    GoTo Returning
X37: ' PRINT a: arguments
    PrintValues sp - oa_(pc) + 1, sp
    For i = sp - oa_(pc) + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - oa_(pc)
    pc = pc + 1: GoTo Dispatch
X38: ' INCL local a += constant b
    If VarType(stk_(fp + oa_(pc))) = vbDouble Then
        stk_(fp + oa_(pc)) = stk_(fp + oa_(pc)) + k_(ob_(pc))
    Else
        rv = AddValues(stk_(fp + oa_(pc)), k_(ob_(pc)))
        stk_(fp + oa_(pc)) = rv
    End If
    pc = pc + 1: GoTo Dispatch
X39: ' INCG global a += constant b
    If VarType(glob_(oa_(pc))) = vbDouble Then
        glob_(oa_(pc)) = glob_(oa_(pc)) + k_(ob_(pc))
    Else
        rv = AddValues(glob_(oa_(pc)), k_(ob_(pc)))
        glob_(oa_(pc)) = rv
    End If
    pc = pc + 1: GoTo Dispatch
X40: ' JNLT
    t = (stk_(sp - 1) < stk_(sp))
    GoTo CompareJump
X41: ' JNLE
    t = (stk_(sp - 1) <= stk_(sp))
    GoTo CompareJump
X42: ' JNGT
    t = (stk_(sp - 1) > stk_(sp))
    GoTo CompareJump
X43: ' JNGE
    t = (stk_(sp - 1) >= stk_(sp))
    GoTo CompareJump
X44: ' JNEQ
    ' two numbers (Integer, Long, Single, Double: % gives a Long) compare directly
    If VarType(stk_(sp)) >= vbInteger And VarType(stk_(sp)) <= vbDouble And VarType(stk_(sp - 1)) >= vbInteger And VarType(stk_(sp - 1)) <= vbDouble Then
        t = (stk_(sp - 1) = stk_(sp))
    Else
        t = SameValues(stk_(sp - 1), stk_(sp))
    End If
    GoTo CompareJump
X45: ' JNNE
    ' two numbers (Integer, Long, Single, Double: % gives a Long) compare directly
    If VarType(stk_(sp)) >= vbInteger And VarType(stk_(sp)) <= vbDouble And VarType(stk_(sp - 1)) >= vbInteger And VarType(stk_(sp - 1)) <= vbDouble Then
        t = (stk_(sp - 1) <> stk_(sp))
    Else
        t = Not SameValues(stk_(sp - 1), stk_(sp))
    End If
    GoTo CompareJump
X46: ' ADDK top + constant a
    If VarType(stk_(sp)) = vbDouble Then
        stk_(sp) = stk_(sp) + k_(oa_(pc))
    Else
        rv = AddValues(stk_(sp), k_(oa_(pc)))
        stk_(sp) = rv
    End If
    pc = pc + 1: GoTo Dispatch
X47: ' DUP
    sp = sp + 1
    If IsObject(stk_(sp - 1)) Then Set stk_(sp) = stk_(sp - 1) Else stk_(sp) = stk_(sp - 1)
    pc = pc + 1: GoTo Dispatch
X48: ' LOADC a: local holding a cell
    sp = sp + 1
    i = stk_(fp + oa_(pc))
    If IsObject(cell_(i)) Then Set stk_(sp) = cell_(i) Else stk_(sp) = cell_(i)
    pc = pc + 1: GoTo Dispatch
X49: ' STOREC
    i = stk_(fp + oa_(pc))
    If IsObject(stk_(sp)) Then Set cell_(i) = stk_(sp) Else cell_(i) = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X50: ' LOADE a: entry of the environment
    sp = sp + 1
    i = env_(env + 1 + oa_(pc))
    If IsObject(cell_(i)) Then Set stk_(sp) = cell_(i) Else stk_(sp) = cell_(i)
    pc = pc + 1: GoTo Dispatch
X51: ' STOREE
    i = env_(env + 1 + oa_(pc))
    If IsObject(stk_(sp)) Then Set cell_(i) = stk_(sp) Else cell_(i) = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X52: ' MKCLOS a: function literal
    sp = sp + 1
    Set stk_(sp) = MakeClosure(oa_(pc), fp, env)
    pc = pc + 1: GoTo Dispatch
X53: ' MKFUNC a: named function as a value
    sp = sp + 1
    Set stk_(sp) = MakeClosure(oa_(pc), fp, env)
    pc = pc + 1: GoTo Dispatch
X54: ' CALL a: arguments, under them the callee
    n = oa_(pc)
    GoTo DynamicCall
X55: ' NEWARR a: items
    n = oa_(pc)
    rv = NewArray(sp - n + 1, n)
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n + 1
    stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X56: ' NEWOBJ a: key and value pairs
    n = oa_(pc)
    Set rv = NewObject(sp - 2 * n + 1, n)
    For i = sp - 2 * n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - 2 * n + 1
    Set stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X57: ' GETIDX
    SetItem rv, IndexValue(stk_(sp - 1), stk_(sp))
    stk_(sp) = Empty
    sp = sp - 1
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X58: ' GETIDXW
    SetItem rv, IndexForWrite(stk_(sp - 1), stk_(sp))
    stk_(sp) = Empty
    sp = sp - 1
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X59: ' SETIDXR value, container, index -> container
    SetItem rv, WithIndex(stk_(sp - 1), stk_(sp), stk_(sp - 2))
    stk_(sp) = Empty
    stk_(sp - 1) = Empty
    sp = sp - 2
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X60: ' GETPROP a: key
    SetItem rv, PropertyValue(stk_(sp), k_(oa_(pc)), hint_(pc))
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X61: ' SETPROPV value, object; a: key
    SetProperty stk_(sp), k_(oa_(pc)), stk_(sp - 1)
    stk_(sp) = Empty
    stk_(sp - 1) = Empty
    sp = sp - 2
    pc = pc + 1: GoTo Dispatch
X62: ' CALLM a: method, b: arguments + METHOD_SHIFT * built-in method number
    n = ob_(pc) Mod METHOD_SHIFT: callKind = 0
    GoTo MethodCall
X63: ' CALLMW
    n = ob_(pc) Mod METHOD_SHIFT: callKind = 1
    GoTo MethodCall
X64: ' EXPAND a: arguments, b: flags of the spread ones -> the items and their count
    n = ExpandItems(sp - oa_(pc) + 1, oa_(pc), CStr(k_(ob_(pc))))
    sp = sp - oa_(pc) + n + 1
    stk_(sp) = CDbl(n)
    pc = pc + 1: GoTo Dispatch
X65: ' CALLFV a: function; the argument count on top
    n = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    fn = oa_(pc): res = sp - n + 1: newEnv = 0
    GoTo CallSetup
X66: ' CALLV
    n = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    GoTo DynamicCall
X67: ' CALLMV
    n = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    callKind = 0
    GoTo MethodCall
X68: ' CALLMWV
    n = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    callKind = 1
    GoTo MethodCall
X69: ' CALLB a: built-in, b: arguments
    n = ob_(pc)
    GoTo BuiltinCall
X70: ' CALLBV
    n = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    GoTo BuiltinCall
X71: ' TPL a: parts
    n = oa_(pc)
    rv = TemplateText(sp - n + 1, n)
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n + 1
    stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X72: ' NEWARRS a: items, b: flags of the spread ones
    n = oa_(pc)
    rv = NewArraySpread(sp - n + 1, n, CStr(k_(ob_(pc))))
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n + 1
    stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X73: ' JNSW a: the case value against the switch value, with =
    t = (stk_(sp) = stk_(sp - 1))
    GoTo CompareJump
X74: ' ITERLIST a: 0 values, 1 keys
    rv = IterList(stk_(sp), oa_(pc))
    stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X75: ' FORNEXT list, counter -> the item, or a jump to a after the last one
    i = stk_(sp)
    If i > ListLength(stk_(sp - 1)) Then
        stk_(sp) = Empty
        stk_(sp - 1) = Empty
        sp = sp - 2
        pc = oa_(pc): GoTo Dispatch
    End If
    SetItem rv, ListItem(stk_(sp - 1), i)
    stk_(sp) = Empty
    sp = sp - 1
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X76: ' UNPACK a: targets, b: 1 with a rest target
    rv = stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    If sp + oa_(pc) + 2 >= stkCap_ Then GrowStack sp + oa_(pc) + 2
    Unpack rv, oa_(pc), ob_(pc) = 1, sp
    sp = sp + oa_(pc) + ob_(pc)
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X77: ' CLASS a: declaration, b: class number
    sp = sp + 1
    Set stk_(sp) = NewClass(k_(oa_(pc)))
    Set clsMap_(ob_(pc)) = stk_(sp)
    clsKeep_(ob_(pc)) = Not curExports_ Is Nothing
    pc = pc + 1: GoTo Dispatch
X78: ' NEWINST a: class name
    Set m = New ASF_Map
    m.Add "type", "Instance"
    m.SetValue "className", k_(oa_(pc))
    sp = sp + 1
    Set stk_(sp) = m
    Set m = Nothing
    pc = pc + 1: GoTo Dispatch
X79: ' BINDM a: method name, b: its function; pops the instance
    Set m = stk_(sp)
    Set m2 = New ASF_Map
    m2.Add "type", "BoundMethod"
    m2.SetValue "name", k_(oa_(pc))
    m2.SetValue "fn", ob_(pc)
    m2.SetValue "instance", m
    m.SetValue CStr(k_(oa_(pc))), m2
    Set m = Nothing: Set m2 = Nothing
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X80: ' CALLFT a: function, b: arguments, under them this
    fn = oa_(pc): n = ob_(pc): res = sp - n: newEnv = 0
    hasThis = True
    SetItem thisVal, stk_(sp - n)
    GoTo CallSetup
X81: ' THROW
    SetItem LAST_ERROR_, stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    throwing_ = True
    If IsObject(LAST_ERROR_) Then errDesc = "throw" Else errDesc = CStr(LAST_ERROR_)
    If RT_.verbose Then GLOBALS_.gRuntimeLog.Add "VM statement error: " & errDesc
    GoTo Unwind
X82: ' IMPORT a: the import -> the values it binds
    topSp_ = sp: topFt_ = ft
    rv = ImportValues(k_(oa_(pc)))
    If sp + UBound(rv) + 2 >= stkCap_ Then GrowStack sp + UBound(rv) + 2
    For i = LBound(rv) To UBound(rv)
        sp = sp + 1
        SetItem stk_(sp), rv(i)
    Next i
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X83: ' (unused)
    pc = pc + 1: GoTo Dispatch
X84: ' VBAEXPR a: the expression, b: the names of the values on the stack
    If Len(k_(ob_(pc))) = 0 Then n = 0 Else n = UBound(Split(CStr(k_(ob_(pc))), vbLf)) + 1
    args = ArgsArray(sp - n + 1, n)
    topSp_ = sp: topFt_ = ft
    SetItem rv, RT_.VBAExpression(CStr(k_(oa_(pc))), Split(CStr(k_(ob_(pc))), vbLf), args)
    args = Empty
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n + 1
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch
X85: ' PROTO a: the prototype; pops its function
    topSp_ = sp: topFt_ = ft
    Set m = k_(oa_(pc))
    RT_.RegisterCOMPrototype CStr(m.GetValue("classType")), CStr(m.GetValue("methodName")), CLng(m.GetValue("funcIndex")), m.GetValue("params"), stk_(sp)
    Set m = Nothing
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X86: ' EXPORT a: name; pops the value, which the module being loaded keeps
    If Not curExports_ Is Nothing Then curExports_.SetValue CStr(k_(oa_(pc))), stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch
X87: ' EXPORTPROTO a: the prototype; pops its function
    topSp_ = sp: topFt_ = ft
    ExportPrototype k_(oa_(pc)), stk_(sp)
    stk_(sp) = Empty
    sp = sp - 1
    pc = pc + 1: GoTo Dispatch

MethodCall:
    ' receiver.name(n arguments): an object's own function is called here with this; anything
    ' else (built-in methods, host objects) goes to ASF_Runtime
    If IsObject(stk_(sp - n)) Then
        If TypeOf stk_(sp - n) Is ASF_Map Then
            Set m = stk_(sp - n)
            If m.TryGetValue(CStr(k_(oa_(pc))), rv) Then
                If ClosureCode(rv, fn, newEnv) Then
                    rv = Empty
                    hasThis = True
                    Set thisVal = m
                    res = sp - n
                    Set m = Nothing
                    GoTo CallSetup
                End If
                If ClassCode(rv, pc, fn, hasThis, thisVal) Then
                    rv = Empty
                    newEnv = 0
                    res = sp - n
                    Set m = Nothing
                    GoTo CallSetup
                End If
            ElseIf StaticMethod(m, CStr(k_(oa_(pc))), pc, fn) Then
                ' Class.method(): a method of the class, called without this
                newEnv = 0
                res = sp - n
                Set m = Nothing
                GoTo CallSetup
            End If
            rv = Empty
            Set m = Nothing
        End If
    End If
    topSp_ = sp: topFt_ = ft
    wrote = False
    If ob_(pc) >= METHOD_SHIFT Then
        t = NativeMethod(ob_(pc) \ METHOD_SHIFT, sp - n, n, rv, wrote, rv2)
    Else
        t = False
    End If
    If Not t Then
        args = ArgsArray(sp - n + 1, n)
        SetItem rv, RT_.CallMemberOn(stk_(sp - n), CStr(k_(oa_(pc))), args, wrote, rv2)
    End If
    args = Empty
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n
    If callKind = 1 Then
        If Not wrote Then SetItem rv2, stk_(sp)
        If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
        sp = sp + 1
        If IsObject(rv2) Then Set stk_(sp) = rv2 Else stk_(sp) = rv2
        rv2 = Empty
        callKind = 0
    Else
        If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    End If
    rv = Empty
    pc = pc + 1: GoTo Dispatch

BuiltinCall:
    args = ArgsArray(sp - n + 1, n)
    topSp_ = sp: topFt_ = ft
    SetItem rv, RT_.CallBuiltin(CStr(k_(oa_(pc))), args)
    args = Empty
    For i = sp - n + 1 To sp
        stk_(i) = Empty
    Next i
    sp = sp - n + 1
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    rv = Empty
    pc = pc + 1: GoTo Dispatch

DynamicCall:
    ' n arguments, under them the callee: a closure, or anything else, which gives undefined
    If Not ClosureCode(stk_(sp - n), fn, newEnv) Then
        ' a bound method, a method of a class, or a class: new Class(...)
        If Not ClassCode(stk_(sp - n), pc, fn, hasThis, thisVal) Then
            ' anything else gives undefined
            For i = sp - n To sp
                stk_(i) = Empty
            Next i
            sp = sp - n
            If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
            rv = Empty
            pc = pc + 1: GoTo Dispatch
        End If
        newEnv = 0
    End If
    res = sp - n
    GoTo CallSetup

CallSetup:
    ' fn with n arguments on top of the stack; res: where its value goes; newEnv: its environment;
    ' hasThis and thisVal: the receiver of a method call; callKind 1: write the receiver back after
    If sp + n + fnNLocals_(fn) + fnMaxStack_(fn) >= stkCap_ Then GrowStack sp + n + fnNLocals_(fn) + fnMaxStack_(fn)
    If fnRest_(fn) >= 0 Then
        rv = RestArgs(sp - n + 1 + fnNParams_(fn), n - fnNParams_(fn))
    End If
    Do While n > fnNParams_(fn)
        stk_(sp) = Empty
        sp = sp - 1
        n = n - 1
    Loop
    Do While n < fnNParams_(fn)
        sp = sp + 1
        stk_(sp) = Empty
        n = n + 1
    Loop
    If fnRest_(fn) >= 0 Then
        sp = sp + 1
        stk_(sp) = rv
        rv = Empty
        n = n + 1
    End If
    ft = ft + 1
    If ft > frCap_ Then GrowFrames
    frRet_(ft) = pc + 1: frFp_(ft) = fp: frFn_(ft) = fn: frEnv_(ft) = env: frRes_(ft) = res
    frKind_(ft) = callKind
    If callKind = 1 Then SetItem frRecv_(ft), thisVal
    callKind = 0
    fp = sp - n + 1
    env = newEnv
    ' the other locals start undefined, this is the receiver or the enclosing this, shared ones go in cells
    For i = n + 1 To fnNLocals_(fn)
        sp = sp + 1
        stk_(sp) = Empty
    Next i
    If fnThis_(fn) >= 0 Then
        If hasThis Then
            SetItem stk_(fp + fnThis_(fn)), thisVal
        ElseIf fnThisEnv_(fn) >= 0 Then
            SetItem stk_(fp + fnThis_(fn)), cell_(env_(env + 1 + fnThisEnv_(fn)))
        End If
    End If
    hasThis = False
    thisVal = Empty
    For i = fnCellStart_(fn) To fnCellStart_(fn) + fnCellCount_(fn) - 1
        stk_(fp + cellSlot_(i)) = NewCell(stk_(fp + cellSlot_(i)))
    Next i
    If sp > stkHigh_ Then stkHigh_ = sp
    If trace_ Then TraceCall fn, fp, n
    pc = fnEntry_(fn)
    GoTo Dispatch

Returning:
    ' rv: the value of the function in the frame ft
    If trace_ Then RT_.PopCallTrace TraceName(frFn_(ft)), rv
    For i = frRes_(ft) + 1 To sp
        stk_(i) = Empty
    Next i
    sp = frRes_(ft)
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    If frKind_(ft) = 1 Then
        sp = sp + 1
        SetItem stk_(sp), frRecv_(ft)
        frRecv_(ft) = Empty
    End If
    pc = frRet_(ft): fp = frFp_(ft): env = frEnv_(ft)
    ft = ft - 1
    If ft = stopFt Then
        SetItem result, rv
        rv = Empty
        GoTo Finished
    End If
    rv = Empty
    GoTo Dispatch

CompareJump:
    stk_(sp) = Empty
    stk_(sp - 1) = Empty
    sp = sp - 2
    If t Then pc = pc + 1 Else pc = oa_(pc)
    GoTo Dispatch

Failed:
    errNum = Err.Number: errDesc = Err.Description
    Resume Recover
Recover:
    callKind = 0: hasThis = False
    If Not throwing_ Then
        errDesc = OperatorMessage(op_(pc), sp, errNum, errDesc)
        LAST_ERROR_ = errDesc
    End If
    If RT_.verbose Then GLOBALS_.gRuntimeLog.Add "VM statement error: " & errDesc
Unwind:
    ' the innermost try around pc takes the error: the stack back to its depth, the value for the catch
    h = TryAt(pc)
    If h > 0 Then
        If ft = 0 Then fn = mainFn_ Else fn = frFn_(ft)
        n = fp - 1 + fnNLocals_(fn) + tryDepth_(h)
        For i = n + 1 To sp
            stk_(i) = Empty
        Next i
        sp = n + 1
        SetItem stk_(sp), LAST_ERROR_
        throwing_ = False
        pc = tryHandler_(h)
        GoTo Dispatch
    End If
    ' else an error ends the program at the top level, and makes a function return undefined
    If ft = 0 Then GoTo Finished
    If trace_ Then RT_.TraceError TraceName(frFn_(ft)), errDesc
    If RT_.verbose Then GLOBALS_.gRuntimeLog.Add "InvokeClosure error: Error during closure execution"
    ' a constructor that fails still gives its instance, as its statements stop there;
    ' methods and constructors keep a throw to themselves
    rv = Empty
    If fnKind_(frFn_(ft)) = 2 Then SetItem rv, stk_(fp + fnThis_(frFn_(ft)))
    If fnKind_(frFn_(ft)) > 0 Then throwing_ = False
    For i = frRes_(ft) To sp
        stk_(i) = Empty
    Next i
    sp = frRes_(ft)
    If IsObject(rv) Then Set stk_(sp) = rv Else stk_(sp) = rv
    If frKind_(ft) = 1 Then
        sp = sp + 1
        SetItem stk_(sp), frRecv_(ft)
        frRecv_(ft) = Empty
    End If
    pc = frRet_(ft): fp = frFp_(ft): env = frEnv_(ft)
    ft = ft - 1
    If ft = stopFt Then
        ' a call from a built-in method gives undefined, the throw stops there
        throwing_ = False
        SetItem result, rv
        rv = Empty
        GoTo Finished
    End If
    rv = Empty
    If throwing_ Then
        ' a throw goes on from the call in the caller
        pc = pc - 1
        GoTo Unwind
    End If
    GoTo Dispatch
Finished:
    If sp > stkHigh_ Then stkHigh_ = sp
    If IsObject(result) Then Set Execute = result Else Execute = result
End Function

Private Sub GrowStack(ByVal need As Long)
    Do While stkCap_ <= need
        stkCap_ = stkCap_ * 2
    Loop
    ReDim Preserve stk_(1 To stkCap_)
End Sub

Private Sub GrowFrames()
    frCap_ = frCap_ * 2
    ReDim Preserve frRet_(1 To frCap_): ReDim Preserve frFp_(1 To frCap_)
    ReDim Preserve frRes_(1 To frCap_): ReDim Preserve frFn_(1 To frCap_)
    ReDim Preserve frEnv_(1 To frCap_): ReDim Preserve frKind_(1 To frCap_): ReDim Preserve frRecv_(1 To frCap_)
End Sub

' A new cell holding v
Private Function NewCell(ByRef v As Variant) As Long
    nCell_ = nCell_ + 1
    If nCell_ > UBound(cell_) Then
        ReDim Preserve cell_(1 To 2 * nCell_)
    End If
    If IsObject(v) Then Set cell_(nCell_) = v Else cell_(nCell_) = v
    NewCell = nCell_
End Function

' A closure of function fn: its environment copies the cells it captures from the creating frame
Private Function MakeClosure(ByVal fn As Long, ByVal fp As Long, ByVal env As Long) As ASF_Map
    Dim envBase As Long, j As Long, n As Long
    n = fnCapCount_(fn)
    envBase = nEnv_
    If nEnv_ + n > UBound(env_) Then
        ReDim Preserve env_(1 To 2 * (nEnv_ + n))
    End If
    For j = 0 To n - 1
        If capKind_(fnCapStart_(fn) + j) = 0 Then
            env_(envBase + 1 + j) = stk_(fp + capIdx_(fnCapStart_(fn) + j))
        Else
            env_(envBase + 1 + j) = env_(env + 1 + capIdx_(fnCapStart_(fn) + j))
        End If
    Next j
    nEnv_ = nEnv_ + n
    Dim m As ASF_Map: Set m = New ASF_Map
    m.SetValue "type", "Closure"
    m.SetValue "fn", CDbl(fn) + CDbl(envBase) * ENV_SHIFT
    Set MakeClosure = m
End Function

' The function and environment of a closure made by MakeClosure; False for anything else
Private Function ClosureCode(ByRef v As Variant, ByRef fn As Long, ByRef env As Long) As Boolean
    If Not IsObject(v) Then Exit Function
    If Not TypeOf v Is ASF_Map Then Exit Function
    Dim m As ASF_Map: Set m = v
    Dim code As Variant
    If Not m.TryGetValue("fn", code) Then Exit Function
    If VarType(code) <> vbDouble Then Exit Function
    env = Int(code / ENV_SHIFT)
    fn = code - CDbl(env) * ENV_SHIFT
    ClosureCode = True
End Function

' The innermost try region around pc, 0 when none
Private Function TryAt(ByVal pc As Long) As Long
    Dim i As Long, best As Long
    For i = 1 To nTry_
        If tryStart_(i) <= pc And pc <= tryEnd_(i) Then
            If best = 0 Then
                best = i
            ElseIf tryStart_(i) > tryStart_(best) Then
                best = i
            End If
        End If
    Next i
    TryAt = best
End Function

' ---------------------------------------------------------------- modules
' import: the values it binds, in the order of the compiler's names (default, namespace, named)
Private Function ImportValues(ByVal info As ASF_Map) As Variant
    Dim path As String: path = RT_.ResolveModule(CStr(info.GetValue("source")))
    Dim moduleObj As ASF_Map: Set moduleObj = LoadModule(path)
    Dim exportsMap As ASF_Map: Set exportsMap = moduleObj.GetValue("exports")
    RegisterPrototypes exportsMap
    Dim named As Collection: Set named = info.GetValue("named")
    Dim out() As Variant, n As Long, i As Long
    ReDim out(0 To 2 + named.Count)
    If info.Exists("defaultImport") Then
        If moduleObj.Exists("defaultExport") Then SetItem out(n), moduleObj.GetValue("defaultExport")
        n = n + 1
    End If
    If info.Exists("namespaceImport") Then
        Dim ns As ASF_Map: Set ns = New ASF_Map
        For i = 1 To exportsMap.Count
            ns.SetValue exportsMap.key(i), exportsMap.GetValue(exportsMap.key(i))
        Next i
        Set out(n) = ns
        n = n + 1
    End If
    For i = 1 To named.Count
        If Not exportsMap.Exists(CStr(named(i))) Then
            Err.Raise vbObjectError + 9001, "VM.ExecuteStmtNode", "Export '" & named(i) & "' not found in module '" & path & "'"
        End If
        SetItem out(n), exportsMap.GetValue(CStr(named(i)))
        n = n + 1
    Next i
    If n = 0 Then
        ImportValues = Array()
    Else
        ReDim Preserve out(0 To n - 1)
        ImportValues = out
    End If
End Function

' The module at path, compiled and run on its first import
Private Function LoadModule(ByVal path As String) As ASF_Map
    If modules_.Exists(path) Then
        Set LoadModule = modules_.GetValue(path)
        Exit Function
    End If
    Dim k As Long
    For k = 1 To loading_.Count
        If loading_(k) = path Then Err.Raise vbObjectError + 9010, "VM.LoadModule", "Circular dependency detected: " & path
    Next k
    Dim src As String
    On Error Resume Next
    src = RT_.ReadModuleSource(path)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 9011, "VM.LoadModule", "Failed to load module: " & path
    End If
    On Error GoTo 0
    Dim idx As Long: idx = BC_.CompileSource(src, "@module:" & path)
    If version_ <> BC_.Version Then LoadCode
    loading_.Add path
    Dim exportsMap As ASF_Map: Set exportsMap = New ASF_Map
    GLOBALS_.gModuleExports.SetValue path, exportsMap
    Dim previousPath As String: previousPath = GLOBALS_.CURRENT_MODULE_PATH
    GLOBALS_.CURRENT_MODULE_PATH = path
    Dim previousExports As ASF_Map: Set previousExports = curExports_
    Set curExports_ = exportsMap
    BindGlobals idx, Empty, True
    ' the module's code runs as a call: an error or a throw in it stops it
    Dim savedSp As Long, savedFt As Long, first As Long, none As Variant, mainFn As Long
    mainFn = BC_.ProgramFunction(idx)
    savedSp = topSp_: savedFt = topFt_
    first = topSp_ + 1
    If first + fnMaxStack_(mainFn) + 8 >= stkCap_ Then GrowStack first + fnMaxStack_(mainFn) + 8
    stk_(first) = Empty
    Execute 0, first, 1, topFt_, 0, topFt_, mainFn, 0, 0, False, none
    stk_(first) = Empty
    topSp_ = savedSp: topFt_ = savedFt
    Set curExports_ = previousExports
    GLOBALS_.CURRENT_MODULE_PATH = previousPath
    loading_.Remove loading_.Count
    If nCell_ > keepCell_ Then keepCell_ = nCell_
    If nEnv_ > keepEnv_ Then keepEnv_ = nEnv_
    Dim moduleObj As ASF_Map: Set moduleObj = New ASF_Map
    moduleObj.SetValue "path", path
    moduleObj.SetValue "exports", exportsMap
    If exportsMap.Exists("__default") Then
        moduleObj.SetValue "defaultExport", exportsMap.GetValue("__default")
        exportsMap.Remove "__default"
    End If
    modules_.SetValue path, moduleObj
    Set LoadModule = moduleObj
End Function

' export prototype...: registered now, and kept as a descriptor that an import registers again
Private Sub ExportPrototype(ByVal info As ASF_Map, ByRef closure As Variant)
    RT_.RegisterCOMPrototype CStr(info.GetValue("classType")), CStr(info.GetValue("methodName")), CLng(info.GetValue("funcIndex")), info.GetValue("params"), closure
    If curExports_ Is Nothing Then Exit Sub
    Dim desc As ASF_Map: Set desc = New ASF_Map
    desc.SetValue "type", "PrototypeDescriptor"
    desc.SetValue "classType", info.GetValue("classType")
    desc.SetValue "methodName", info.GetValue("methodName")
    desc.SetValue "funcIndex", info.GetValue("funcIndex")
    desc.SetValue "internalFuncName", info.GetValue("internalFuncName")
    desc.SetValue "params", info.GetValue("params")
    desc.SetValue "closure", closure
    curExports_.SetValue CStr(info.GetValue("exportName")), desc
End Sub

' The prototypes a module exports go live on every import of it
Private Sub RegisterPrototypes(ByVal exportsMap As ASF_Map)
    Dim i As Long, v As Variant, desc As ASF_Map, closure As Variant
    For i = 1 To exportsMap.Count
        SetItem v, exportsMap.GetValue(exportsMap.key(i))
        If IsObject(v) Then
            If TypeOf v Is ASF_Map Then
                Set desc = v
                If desc.Exists("type") Then
                    If desc.GetValue("type") = "PrototypeDescriptor" Then
                        closure = Empty
                        If desc.Exists("closure") Then Set closure = desc.GetValue("closure")
                        RT_.RegisterCOMPrototype CStr(desc.GetValue("classType")), CStr(desc.GetValue("methodName")), CLng(desc.GetValue("funcIndex")), desc.GetValue("params"), closure
                    End If
                End If
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------- classes
' A class value: a copy of the one the compiler made, whose descriptors it shares
Private Function NewClass(ByVal template As ASF_Map) As ASF_Map
    Dim klass As ASF_Map: Set klass = New ASF_Map
    Dim i As Long
    For i = 1 To template.Count
        klass.SetValue template.key(i), template.GetValue(template.key(i))
    Next i
    Set NewClass = klass
End Function

' The function of a bound method (and its instance), a method of a class or a class (its constructor)
Private Function ClassCode(ByRef v As Variant, ByVal pc As Long, ByRef fn As Long, ByRef hasThis As Boolean, ByRef thisVal As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    If Not TypeOf v Is ASF_Map Then Exit Function
    Dim m As ASF_Map: Set m = v
    Dim kind As Variant, code As Variant, i As Long
    If Not m.TryGetValue("type", kind) Then Exit Function
    Select Case kind
        Case "BoundMethod", "Method"
            If Not m.TryGetValue("fn", code) Then Exit Function
            If VarType(code) <> vbLong Then Exit Function
            fn = code
            If kind = "BoundMethod" Then
                hasThis = True
                Set thisVal = m.GetValue("instance")
            End If
        Case "Class"
            For i = 1 To nCls_
                If IsObject(clsMap_(i)) Then
                    If clsMap_(i) Is m Then
                        fn = clsNew_(i)
                        ClassCode = True
                        Exit Function
                    End If
                End If
            Next i
            Exit Function
        Case Else
            Exit Function
    End Select
    ClassCode = True
End Function

' Class.name(): the class's method of that name, as EvalMemberNode finds it
Private Function StaticMethod(cls As ASF_Map, ByVal name As String, ByVal pc As Long, ByRef fn As Long) As Boolean
    Dim methods As Variant, item As Variant, code As Variant
    If Not cls.TryGetValue("classMethods", methods) Then Exit Function
    If Not IsObject(methods) Then Exit Function
    For Each item In methods
        If item.GetValue("name") = name Then
            If item.TryGetValue("fn", code) Then
                If VarType(code) = vbLong Then
                    fn = code
                    StaticMethod = True
                End If
            End If
            Exit Function
        End If
    Next item
End Function

' ---------------------------------------------------------------- built-in methods
' The methods this VM runs itself, as ASF_Runtime does them: id from ASF_Compiler.MethodId,
' the receiver at recvAt with n arguments after it. False leaves the call to ASF_Runtime.
Private Function NativeMethod(ByVal id As Long, ByVal recvAt As Long, ByVal n As Long, ByRef result As Variant, _
                              ByRef wrote As Boolean, ByRef written As Variant) As Boolean
    Dim recv As Variant
    SetItem recv, stk_(recvAt)
    Select Case id
        Case 1 To 9
            If Not IsArray(recv) Then Exit Function
            Select Case id
                Case 1: result = ArrayPush(recv, recvAt + 1, n, written): wrote = True
                Case 2
                    If UBound(recv) < LBound(recv) Then
                        result = Empty
                    Else
                        SetItem result, recv(UBound(recv))
                        If UBound(recv) = LBound(recv) Then
                            Dim one() As Variant
                            ReDim one(LBound(recv) To LBound(recv))
                            written = one
                        Else
                            ReDim Preserve recv(LBound(recv) To UBound(recv) - 1)
                            written = recv
                        End If
                        wrote = True
                    End If
                Case 3 To 6
                    Dim fn As Long, env As Long
                    If n < 1 Then
                        result = Empty
                    ElseIf Not ClosureCode(stk_(recvAt + 1), fn, env) Then
                        result = Empty
                    Else
                        Select Case id
                            Case 3: result = ArrayMap(recv, fn, env)
                            Case 4: result = ArrayFilter(recv, fn, env)
                            Case 5
                                If n >= 2 Then
                                    SetItem result, ArrayReduce(recv, fn, env, True, stk_(recvAt + 2))
                                Else
                                    SetItem result, ArrayReduce(recv, fn, env, False, Empty)
                                End If
                            Case 6
                                ArrayForEach recv, fn, env
                                result = recv
                        End Select
                    End If
                Case Else: Exit Function
            End Select
        Case 10 To 19
            If VarType(recv) <> vbString Then Exit Function
            Select Case id
                Case 10: result = UCase$(recv)
                Case 11: result = LCase$(recv)
                Case 12: result = Trim$(recv)
                Case Else: Exit Function
            End Select
        Case 20 To 29
            If Not IsObject(recv) Then Exit Function
            If Not recv Is math_ Then Exit Function
            If n = 0 And id <> 20 And id <> 21 Then Exit Function
            Dim i As Long, x As Variant
            Select Case id
                Case 20, 21
                    result = Empty
                    If n > 0 Then
                        result = stk_(recvAt + 1)
                        For i = 1 To n
                            If id = 20 Then
                                If stk_(recvAt + i) > result Then result = stk_(recvAt + i)
                            Else
                                If stk_(recvAt + i) < result Then result = stk_(recvAt + i)
                            End If
                        Next i
                    End If
                Case 22
                    x = stk_(recvAt + 1)
                    result = Fix(x)
                    result = result + ((x <> result) And (x < 0))
                Case 23
                    x = stk_(recvAt + 1)
                    result = Fix(x)
                    result = result - ((x <> result) And (x > 0))
                Case 24: result = Abs(stk_(recvAt + 1))
                Case 25: result = Round(stk_(recvAt + 1))
                Case 26: result = Sqr(stk_(recvAt + 1))
                Case 27
                    If n < 2 Then Exit Function
                    result = stk_(recvAt + 1) ^ stk_(recvAt + 2)
                Case Else: Exit Function
            End Select
        Case Else
            Exit Function
    End Select
    NativeMethod = True
End Function

' push: a new array with the items added, and its length
Private Function ArrayPush(ByRef arr As Variant, ByVal first As Long, ByVal n As Long, ByRef newArr As Variant) As Long
    Dim offs As Long, lb As Long, oldLen As Long, i As Long, out() As Variant
    If UBound(arr) = -1 Then
        lb = 1: oldLen = 1: offs = 1
    Else
        lb = LBound(arr): oldLen = UBound(arr) - LBound(arr) + 1
    End If
    ReDim out(lb To lb + oldLen + n - 1 - offs)
    If offs = 0 Then
        For i = lb To UBound(arr)
            If IsObject(arr(i)) Then Set out(i) = arr(i) Else out(i) = arr(i)
        Next i
    End If
    For i = 1 To n
        If IsObject(stk_(first + i - 1)) Then Set out(lb + oldLen + i - 1 - offs) = stk_(first + i - 1) Else out(lb + oldLen + i - 1 - offs) = stk_(first + i - 1)
    Next i
    newArr = out
    ArrayPush = oldLen + n
End Function

' f(args) for a closure of this VM, called by a built-in method; this is the array it works on
Private Function Callback(ByVal fn As Long, ByVal env As Long, ByRef thisV As Variant, ByVal n As Long, _
                        ByRef a1 As Variant, ByRef a2 As Variant, ByRef a3 As Variant) As Variant
    Dim savedSp As Long, savedFt As Long, first As Long
    savedSp = topSp_: savedFt = topFt_
    first = topSp_ + 1
    If first + n + fnNLocals_(fn) + fnMaxStack_(fn) + 8 >= stkCap_ Then GrowStack first + n + fnNLocals_(fn) + fnMaxStack_(fn) + 8
    stk_(first) = Empty
    If n >= 1 Then SetItem stk_(first + 1), a1
    If n >= 2 Then SetItem stk_(first + 2), a2
    If n >= 3 Then SetItem stk_(first + 3), a3
    Dim r As Variant
    SetItem r, Execute(0, first + n, 1, topFt_, 0, topFt_, fn, n, env, True, thisV)
    stk_(first) = Empty
    topSp_ = savedSp: topFt_ = savedFt
    If IsObject(r) Then Set Callback = r Else Callback = r
End Function

' map: each item (inside nested arrays too) through the function
Private Function ArrayMap(ByRef arr As Variant, ByVal fn As Long, ByVal env As Long) As Variant
    Dim cnt As Long, i As Long, out() As Variant
    cnt = UBound(arr) - LBound(arr) + 1
    If cnt < 0 Then cnt = 0
    If cnt = 0 Then
        ReDim out(1 To 1)
    Else
        ReDim out(1 To cnt)
    End If
    For i = 0 To cnt - 1
        If IsArray(arr(LBound(arr) + i)) Then
            out(1 + i) = ArrayMap(arr(LBound(arr) + i), fn, env)
        Else
            SetItem out(1 + i), Callback(fn, env, arr, 1, arr(LBound(arr) + i), Empty, Empty)
        End If
    Next i
    ArrayMap = out
End Function

Private Function ArrayFilter(ByRef arr As Variant, ByVal fn As Long, ByVal env As Long) As Variant
    Dim kept As Collection: Set kept = New Collection
    Dim i As Long, out() As Variant
    For i = LBound(arr) To UBound(arr)
        If Truthy(Callback(fn, env, arr, 1, arr(i), Empty, Empty)) Then kept.Add arr(i)
    Next i
    If kept.Count = 0 Then
        ReDim out(1 To 1)
    Else
        ReDim out(1 To kept.Count)
        For i = 1 To kept.Count
            If IsObject(kept(i)) Then Set out(i) = kept(i) Else out(i) = kept(i)
        Next i
    End If
    ArrayFilter = out
End Function

Private Function ArrayReduce(ByRef arr As Variant, ByVal fn As Long, ByVal env As Long, ByVal hasInit As Boolean, ByRef init As Variant) As Variant
    Dim acc As Variant, i As Long, start As Long
    If hasInit Then
        SetItem acc, init
        start = LBound(arr)
    Else
        SetItem acc, arr(LBound(arr))
        start = LBound(arr) + 1
    End If
    For i = start To UBound(arr)
        SetItem acc, Callback(fn, env, arr, 2, acc, arr(i), Empty)
    Next i
    If IsObject(acc) Then Set ArrayReduce = acc Else ArrayReduce = acc
End Function

Private Sub ArrayForEach(ByRef arr As Variant, ByVal fn As Long, ByVal env As Long)
    Dim i As Long, r As Variant
    For i = LBound(arr) To UBound(arr)
        SetItem r, Callback(fn, env, arr, 3, arr(i), CLng(1 + i - LBound(arr)), arr)
    Next i
End Sub

' ---------------------------------------------------------------- call trace (EnableCallTrace)
Private Sub TraceCall(ByVal fn As Long, ByVal first As Long, ByVal n As Long)
    Dim args As Collection: Set args = New Collection
    Dim i As Long
    For i = first To first + n - 1
        args.Add stk_(i)
    Next i
    RT_.PushCallTrace TraceName(fn), args
End Sub

' The name a function has in the trace: a function literal is <anonymous>, a constructor its class
Private Function TraceName(ByVal fn As Long) As String
    TraceName = BC_.FunctionName(fn)
    If TraceName = "@literal" Then
        TraceName = "<anonymous>"
    ElseIf Left$(TraceName, 5) = "@new " Then
        TraceName = Mid$(TraceName, 6)
    ElseIf Left$(TraceName, 6) = "@ctor " Then
        TraceName = Mid$(TraceName, 7)
    End If
End Function

' ---------------------------------------------------------------- calls from the runtime
' closure(args) for ASF_Runtime: a built-in method it runs (map, sort...) calls back a closure of
' this VM; the frame goes above the values in use, and the call ends when that frame returns
Public Function CallFromRuntime(closure As ASF_Map, args As Collection, ByVal hasThis As Boolean, ByRef thisVal As Variant) As Variant
    Dim fn As Long, env As Long, i As Long, n As Long, first As Long
    If Not ClosureCode(closure, fn, env) Then Exit Function
    Dim savedSp As Long, savedFt As Long
    savedSp = topSp_: savedFt = topFt_
    first = topSp_ + 1
    n = args.Count
    If first + n + fnNLocals_(fn) + fnMaxStack_(fn) + 8 >= stkCap_ Then GrowStack first + n + fnNLocals_(fn) + fnMaxStack_(fn) + 8
    stk_(first) = Empty
    For i = 1 To n
        SetItem stk_(first + i), args(i)
    Next i
    Dim r As Variant
    SetItem r, Execute(0, first + n, 1, topFt_, 0, topFt_, fn, n, env, hasThis, thisVal)
    stk_(first) = Empty
    topSp_ = savedSp: topFt_ = savedFt
    If IsObject(r) Then Set CallFromRuntime = r Else CallFromRuntime = r
End Function

Private Function ArgsArray(ByVal first As Long, ByVal n As Long) As Variant
    Dim a() As Variant, i As Long
    If n = 0 Then
        ArgsArray = Array()
        Exit Function
    End If
    ReDim a(0 To n - 1)
    For i = 0 To n - 1
        If IsObject(stk_(first + i)) Then Set a(i) = stk_(first + i) Else a(i) = stk_(first + i)
    Next i
    ArgsArray = a
End Function

' The extra arguments of a function with a rest parameter
Private Function RestArgs(ByVal first As Long, ByVal count As Long) As Variant
    Dim a() As Variant, i As Long
    If count <= 0 Then
        ReDim a(-1 To -1)
    ElseIf count = 1 And IsEmpty(stk_(first)) Then
        ReDim a(-1 To -1)
    Else
        ReDim a(1 To count)
        For i = 1 To count
            If IsObject(stk_(first + i - 1)) Then Set a(i) = stk_(first + i - 1) Else a(i) = stk_(first + i - 1)
        Next i
    End If
    RestArgs = a
End Function

' Spread arguments in place: strings give their characters, arrays their items (as ASF 3's
' calls do); returns the count of items now from first
Private Function ExpandItems(ByVal first As Long, ByVal n As Long, ByVal flags As String) As Long
    Dim items As Collection: Set items = New Collection
    Dim i As Long, j As Long, v As Variant
    For i = 1 To n
        SetItem v, stk_(first + i - 1)
        If Mid$(flags, i, 1) = "1" And Not IsObject(v) Then
            If VarType(v) = vbString Then
                For j = 1 To Len(v)
                    items.Add Mid$(v, j, 1)
                Next j
            ElseIf IsArray(v) Then
                For j = LBound(v) To UBound(v)
                    items.Add v(j)
                Next j
            Else
                items.Add v
            End If
        Else
            items.Add v
        End If
        v = Empty
        stk_(first + i - 1) = Empty
    Next i
    If first + items.Count + 2 >= stkCap_ Then GrowStack first + items.Count + 2
    For i = 1 To items.Count
        If IsObject(items(i)) Then Set stk_(first + i - 1) = items(i) Else stk_(first + i - 1) = items(i)
    Next i
    If first + items.Count > stkHigh_ Then stkHigh_ = first + items.Count
    ExpandItems = items.Count
End Function

' A template: arrays and objects printed, null as null, undefined as nothing
Private Function TemplateText(ByVal first As Long, ByVal n As Long) As String
    Dim i As Long, out As String
    For i = first To first + n - 1
        If IsArray(stk_(i)) Or IsMapValue(stk_(i)) Then
            out = out & RT_.FormatValue(stk_(i))
        ElseIf IsNull(stk_(i)) Then
            out = out & "null"
        ElseIf Not IsEmpty(stk_(i)) Then
            out = out & CStr(stk_(i))
        End If
    Next i
    TemplateText = out
End Function

' An array literal with spread items
Private Function NewArraySpread(ByVal first As Long, ByVal n As Long, ByVal flags As String) As Variant
    Dim items As Collection: Set items = New Collection
    Dim i As Long, j As Long, v As Variant
    For i = 1 To n
        SetItem v, stk_(first + i - 1)
        If Mid$(flags, i, 1) = "1" Then
            If IsObject(v) Then
                items.Add v
            ElseIf VarType(v) = vbString Then
                For j = 1 To Len(v)
                    items.Add Mid$(v, j, 1)
                Next j
            ElseIf IsArray(v) Then
                If UBound(v) <> -1 Then
                    For j = LBound(v) To UBound(v)
                        items.Add v(j)
                    Next j
                End If
            ElseIf Not IsEmpty(v) Then
                items.Add v
            End If
        Else
            items.Add v
        End If
        v = Empty
    Next i
    Dim arr() As Variant
    If items.Count = 0 Then
        ReDim arr(1 To 1)
    Else
        ReDim arr(1 To items.Count)
        For i = 1 To items.Count
            If IsObject(items(i)) Then Set arr(i) = items(i) Else arr(i) = items(i)
        Next i
    End If
    NewArraySpread = arr
End Function

' What for-of (values) or for-in (keys) visits: the items of arrays and strings, the keys of objects
Private Function IterList(ByRef v As Variant, ByVal keys As Long) As Variant
    Dim items As Collection: Set items = New Collection
    Dim i As Long, k As Variant
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If keys = 1 Then items.Add CLng(1 + i - LBound(v)) Else items.Add v(i)
        Next i
    ElseIf VarType(v) = vbString Then
        For i = 1 To Len(v)
            If keys = 1 Then items.Add i Else items.Add Mid$(v, i, 1)
        Next i
    ElseIf IsMapValue(v) Then
        Dim m As ASF_Map: Set m = v
        For Each k In m.keys()
            If keys = 1 Then items.Add k Else items.Add m.GetValue(CStr(k))
        Next k
    End If
    Dim arr() As Variant
    If items.Count = 0 Then
        ReDim arr(-1 To -1)
    Else
        ReDim arr(1 To items.Count)
        For i = 1 To items.Count
            If IsObject(items(i)) Then Set arr(i) = items(i) Else arr(i) = items(i)
        Next i
    End If
    IterList = arr
End Function

Private Function ListLength(ByRef list As Variant) As Long
    If UBound(list) = -1 Then Exit Function
    ListLength = UBound(list) - LBound(list) + 1
End Function

Private Function ListItem(ByRef list As Variant, ByVal i As Long) As Variant
    If IsObject(list(LBound(list) + i - 1)) Then Set ListItem = list(LBound(list) + i - 1) Else ListItem = list(LBound(list) + i - 1)
End Function

' [a, b, ...rest] = src: the values of the targets above sp, the first on top, the rest first
Private Sub Unpack(ByRef src As Variant, ByVal n As Long, ByVal hasRest As Boolean, ByVal sp As Long)
    If Not IsArray(src) Then Err.Raise vbObjectError + 6001, "VM.ExecuteStmtNode", "Cannot destructure non-array value"
    Dim lb As Long, ub As Long, i As Long, at As Long
    lb = LBound(src): ub = UBound(src)
    at = sp + 1
    If hasRest Then
        Dim remaining As Long: remaining = (ub - lb + 1) - n
        Dim restArr() As Variant
        If remaining > 0 Then
            ReDim restArr(1 To remaining)
            For i = 1 To remaining
                If IsObject(src(lb + n + i - 1)) Then Set restArr(i) = src(lb + n + i - 1) Else restArr(i) = src(lb + n + i - 1)
            Next i
        Else
            ReDim restArr(-1 To -1)
        End If
        stk_(at) = restArr
        at = at + 1
    End If
    For i = n To 1 Step -1
        If lb + i - 1 <= ub Then
            SetItem stk_(at), src(lb + i - 1)
        Else
            stk_(at) = Empty
        End If
        at = at + 1
    Next i
    If at > stkHigh_ Then stkHigh_ = at
End Sub

' ---------------------------------------------------------------- arrays and objects
Private Function NewArray(ByVal first As Long, ByVal n As Long) As Variant
    Dim arr() As Variant, i As Long
    If n = 0 Then
        ReDim arr(-1 To -1)
    Else
        ReDim arr(1 To n)
        For i = 1 To n
            If IsObject(stk_(first + i - 1)) Then Set arr(i) = stk_(first + i - 1) Else arr(i) = stk_(first + i - 1)
        Next i
    End If
    NewArray = arr
End Function

Private Function NewObject(ByVal first As Long, ByVal n As Long) As ASF_Map
    Dim m As ASF_Map: Set m = New ASF_Map
    Dim i As Long, j As Long, k As Variant, src As Variant
    For i = 0 To n - 1
        If CStr(stk_(first + 2 * i)) = "__SPREAD__" Then
            ' {...source}: an object's keys, an array's items under 1, 2..., a string's characters
            SetItem src, stk_(first + 2 * i + 1)
            If IsMapValue(src) Then
                Dim sm As ASF_Map: Set sm = src
                For Each k In sm.keys()
                    m.SetValue CStr(k), sm.GetValue(CStr(k))
                Next k
            ElseIf IsArray(src) Then
                For j = LBound(src) To UBound(src)
                    m.SetValue CStr(j - LBound(src) + 1), src(j)
                Next j
            ElseIf VarType(src) = vbString Then
                For j = 1 To Len(src)
                    m.SetValue CStr(j), Mid$(src, j, 1)
                Next j
            End If
            src = Empty
        Else
            m.SetValue CStr(stk_(first + 2 * i)), stk_(first + 2 * i + 1)
        End If
    Next i
    Set NewObject = m
End Function

' c[k]: indexes from 1, objects by key
Private Function IndexValue(ByRef c As Variant, ByRef k As Variant) As Variant
    Dim pos As Long: pos = CLng(k)
    If pos < 1 Then Err.Raise vbObjectError + 5001, "VM.EvalIndexNode", "Invalid index (must >= base)"
    If IsArray(c) Then
        If LBound(c) + pos - 1 > UBound(c) Then Err.Raise vbObjectError + 5005, "VM.EvalIndexNode", "Index out of bounds"
        If IsObject(c(LBound(c) + pos - 1)) Then Set IndexValue = c(LBound(c) + pos - 1) Else IndexValue = c(LBound(c) + pos - 1)
    ElseIf IsMapValue(c) Then
        Dim m As ASF_Map: Set m = c
        If IsObject(m.GetValue(CStr(pos))) Then Set IndexValue = m.GetValue(CStr(pos)) Else IndexValue = m.GetValue(CStr(pos))
    Else
        Err.Raise vbObjectError + 5006, "VM.EvalIndexNode", "Cannot index non-array/non-map"
    End If
End Function

' c[k] of a container about to be written: undefined where the write creates it
Private Function IndexForWrite(ByRef c As Variant, ByRef k As Variant) As Variant
    If IsArray(c) Then
        Dim pos As Long: pos = CLng(k)
        If pos < 1 Then Err.Raise vbObjectError + 5001, "VM.HandleAssignment", "Invalid index"
        If UBound(c) = -1 Then Exit Function
        If LBound(c) + pos - 1 > UBound(c) Then Exit Function
        If IsObject(c(LBound(c) + pos - 1)) Then Set IndexForWrite = c(LBound(c) + pos - 1) Else IndexForWrite = c(LBound(c) + pos - 1)
    ElseIf IsMapValue(c) Then
        If IsObject(c) Then Set IndexForWrite = IndexValue(c, k) Else IndexForWrite = IndexValue(c, k)
    End If
End Function

' c with c[k] = v: arrays grow as needed and are copies, objects are changed in place
Private Function WithIndex(ByRef c As Variant, ByRef k As Variant, ByRef v As Variant) As Variant
    If IsMapValue(c) Then
        Dim m As ASF_Map: Set m = c
        m.SetValue CStr(k), v
        Set WithIndex = m
        Exit Function
    End If
    Dim pos As Long: pos = CLng(k)
    If pos < 1 Then Err.Raise vbObjectError + 5001, "VM.HandleAssignment", "Invalid index"
    Dim arr As Variant
    If IsArray(c) Then
        arr = c
        If UBound(arr) = -1 Then
            Dim fresh() As Variant
            ReDim fresh(1 To pos)
            arr = fresh
        ElseIf LBound(arr) + pos - 1 > UBound(arr) Then
            ReDim Preserve arr(LBound(arr) To LBound(arr) + pos - 1)
        End If
    Else
        Err.Raise vbObjectError + 5002, "VM.HandleAssignment", "Cannot index non-array/non-map for set"
    End If
    If IsObject(v) Then Set arr(LBound(arr) + pos - 1) = v Else arr(LBound(arr) + pos - 1) = v
    WithIndex = arr
End Function

' obj.key; hint is the instruction's cache
Private Function PropertyValue(ByRef obj As Variant, ByVal key As String, ByRef hint As Long) As Variant
    If IsObject(obj) Then
        If TypeOf obj Is ASF_Map Then
            Dim m As ASF_Map: Set m = obj
            Dim v As Variant
            If m.GetHinted(key, hint, v) Then
                If IsObject(v) Then Set PropertyValue = v Else PropertyValue = v
                Exit Function
            End If
            ' a class keeps its static methods aside; a plain object without the key gives undefined
            If Not m.Exists("classMethods") Then Exit Function
        End If
    ElseIf key = "length" Then
        If IsArray(obj) Then
            If UBound(obj) = -1 Then PropertyValue = 0 Else PropertyValue = UBound(obj) - LBound(obj) + 1
            Exit Function
        End If
        If VarType(obj) = vbString Then
            PropertyValue = Len(obj)
            Exit Function
        End If
    End If
    Dim r As Variant
    SetItem r, RT_.MemberOf(obj, key)
    If IsObject(r) Then Set PropertyValue = r Else PropertyValue = r
End Function

Private Sub SetProperty(ByRef obj As Variant, ByVal key As String, ByRef v As Variant)
    If IsMapValue(obj) Then
        Dim m As ASF_Map: Set m = obj
        m.SetValue key, v
    Else
        RT_.SetMemberOf obj, key, v
    End If
End Sub

' ---------------------------------------------------------------- values
' Truthiness: numbers and numeric strings by value, 'true' and 'false' by name
Private Function Truthy(ByRef v As Variant) As Boolean
    If IsObject(v) Then Truthy = Not v Is Nothing: Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Truthy = v: Exit Function
    If IsNumeric(v) Then Truthy = (CDbl(v) <> 0): Exit Function
    If VarType(v) = vbString Then
        Select Case LCase$(v)
            Case "true": Truthy = True
            Case "false": Truthy = False
            Case Else: Truthy = (Len(v) > 0)
        End Select
        Exit Function
    End If
    Truthy = True
End Function

' +: numbers add, a string on either side concatenates
Private Function AddValues(ByRef a As Variant, ByRef b As Variant) As Variant
    If Not IsObject(a) And Not IsObject(b) Then
        If VarType(a) <> vbString And VarType(b) <> vbString Then
            If VarType(a) < vbArray And VarType(b) < vbArray Then AddValues = a + b: Exit Function
            RaiseOperandError "+", a, b
        End If
    ElseIf TypeName(a) <> "String" And TypeName(b) <> "String" Then
        If IsMapValue(a) Or IsMapValue(b) Then RaiseOperandError "+", a, b
        AddValues = a + b: Exit Function
    End If
    AddValues = TextOf(a) & TextOf(b)
End Function

' A side of + or & with a string: arrays and objects as a template prints them, null as nothing
Private Function TextOf(ByRef v As Variant) As String
    If IsArray(v) Or IsMapValue(v) Then
        TextOf = RT_.FormatValue(v)
    ElseIf Not IsNull(v) Then
        TextOf = CStr(v)
    End If
End Function

Private Function ShiftValues(ByRef a As Variant, ByRef b As Variant, ByVal toLeft As Boolean) As Variant
    If IsNumeric(a) And IsNumeric(b) Then
        If toLeft Then
            ShiftValues = Floor(Floor(a) * (2 ^ Abs(Floor(b))))
        Else
            ShiftValues = Floor(Floor(a) / (2 ^ Abs(Floor(b))))
        End If
    ElseIf IsArray(a) Or IsArray(b) Or IsObject(a) Or IsObject(b) Then
        If toLeft Then RaiseOperandError "<<", a, b Else RaiseOperandError ">>", a, b
    End If
End Function

Private Function Floor(ByRef value As Variant) As Long
    Dim tmp As Double
    tmp = CDbl(Fix(value))
    Floor = CLng(tmp + ((CDbl(value) <> tmp) And (CDbl(value) < 0)))
End Function

Private Function IsMapValue(ByRef v As Variant) As Boolean
    If IsObject(v) Then IsMapValue = TypeOf v Is ASF_Map
End Function

' ==: same objects, arrays with the same items, null as undefined
Private Function IsNumberType(ByRef v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte: IsNumberType = True
    End Select
End Function

Private Function SameValues(ByRef a As Variant, ByRef b As Variant) As Boolean
    If VarType(a) = VarType(b) Then
        Select Case VarType(a)
            Case vbString, vbDate, vbBoolean
                SameValues = (a = b)
                Exit Function
        End Select
    End If
    On Error GoTo NotEqual
    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then SameValues = (a Is b)
        Exit Function
    End If
    If IsArray(a) Or IsArray(b) Then
        If IsArray(a) And IsArray(b) Then SameValues = SameArrays(a, b)
        Exit Function
    End If
    ' a numeric text equals its number ('1' == 1), as LibreOffice Basic compares them; VBA would
    ' put every number before every text
    If VarType(a) = vbString And IsNumberType(b) Then
        If IsNumeric(a) Then SameValues = (CDbl(a) = CDbl(b))
        Exit Function
    End If
    If VarType(b) = vbString And IsNumberType(a) Then
        If IsNumeric(b) Then SameValues = (CDbl(a) = CDbl(b))
        Exit Function
    End If
    SameValues = (IIf(IsNull(a), Empty, a) = IIf(IsNull(b), Empty, b))
    Exit Function
NotEqual:
    SameValues = False
End Function

Private Function SameArrays(ByRef a As Variant, ByRef b As Variant) As Boolean
    If UBound(a) - LBound(a) <> UBound(b) - LBound(b) Then Exit Function
    Dim i As Long
    For i = 0 To UBound(a) - LBound(a)
        If Not SameValues(a(LBound(a) + i), b(LBound(b) + i)) Then Exit Function
    Next i
    SameArrays = True
End Function

Private Sub RaiseOperandError(ByVal op As String, ByRef a As Variant, ByRef b As Variant)
    Err.Raise vbObjectError + 5013, "VM.EvalBinary", "Operator " & op & " cannot be applied to " & OperandText(a) & " and " & OperandText(b)
End Sub

Private Function OperandText(ByRef v As Variant) As String
    If IsArray(v) Then
        OperandText = "an array"
    ElseIf IsMapValue(v) Then
        OperandText = "an object"
    ElseIf IsObject(v) Then
        If v Is Nothing Then OperandText = "null" Else OperandText = "a " & TypeName(v) & " object"
    ElseIf IsNull(v) Then
        OperandText = "null"
    ElseIf IsEmpty(v) Then
        OperandText = "undefined"
    ElseIf VarType(v) = vbString Then
        If Len(v) > 20 Then OperandText = "'" & Left$(v, 17) & "...'" Else OperandText = "'" & v & "'"
    Else
        OperandText = CStr(v)
    End If
End Function

' The message of an error of a binary operator
Private Function OperatorMessage(ByVal code As Long, ByVal sp As Long, ByVal errNum As Long, ByVal errDesc As String) As String
    OperatorMessage = errDesc
    If errNum = vbObjectError + 5013 Then Exit Function
    Dim sym As String
    Select Case code
        Case OP_ADD To OP_SHR: sym = Choose(code - OP_ADD + 1, "+", "-", "*", "/", "%", "^", "==", "!=", "<", "<=", ">", ">=", "&", "<<", ">>")
        Case OP_JNLT To OP_JNNE: sym = Choose(code - OP_JNLT + 1, "<", "<=", ">", ">=", "==", "!=")
        Case Else: Exit Function
    End Select
    If sp < 2 Then Exit Function
    Dim a As Variant, b As Variant
    SetItem a, stk_(sp - 1)
    SetItem b, stk_(sp)
    If errNum = 13 Or IsArray(a) Or IsArray(b) Or IsObject(a) Or IsObject(b) Then
        OperatorMessage = "Operator " & sym & " cannot be applied to " & OperandText(a) & " and " & OperandText(b)
    End If
End Function

Private Sub PrintValues(ByVal first As Long, ByVal last As Long)
    Dim i As Long, sb As String
    For i = first To last
        If i > first Then sb = sb & ", "
        sb = sb & RT_.FormatValue(stk_(i))
    Next i
    If RT_.verbose Then GLOBALS_.gRuntimeLog.Add "PRINT:" & sb
    Debug.Print sb
End Sub
