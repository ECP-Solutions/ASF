VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GLOBALS_ As ASF_Globals
Private Compiler_ As ASF_Compiler
Private VM_ As ASF_VM
Private OUTPUT__ As Variant
Private EVALUATOR_ As VBAexpressions
' the bytecode backend (src/ASF bytecode) is compiled only with the argument ASF_BYTECODE = 1
Private FASTVM_ON_ As Boolean
Private BACKEND_ As String
Private FALLBACK_ As String
#If ASF_BYTECODE Then
Private BC_ As ASF_Bytecode
Private FVM_ As ASF_FastVM
#End If

' True runs programs on the bytecode backend when it takes them, and on the tree VM otherwise
Public Property Get FastVM() As Boolean
    FastVM = FASTVM_ON_
End Property
Public Property Let FastVM(aValue As Boolean)
    FASTVM_ON_ = aValue
End Property
' The backend that ran the last program: "bytecode" or "tree"
Public Property Get Backend() As String
    Backend = BACKEND_
End Property
' Why the last program ran on the tree VM with FastVM on
Public Property Get FallbackReason() As String
    FallbackReason = FALLBACK_
End Property

Public Property Get verbose() As Boolean
    verbose = VM_.verbose
End Property
Public Property Let verbose(aValue As Boolean)
    VM_.verbose = aValue
End Property
Public Property Get EnableCallTrace() As Boolean
    EnableCallTrace = VM_.EnableCallTrace
End Property
Public Property Let EnableCallTrace(aValue As Boolean)
    VM_.EnableCallTrace = aValue
End Property
Public Property Get GetCallStackTrace() As String
    GetCallStackTrace = VM_.GetCallStackTrace
End Property
Public Property Get GetCallStack() As Collection
    Set GetCallStack = VM_.GetCallStack
End Property
Public Property Let WorkingDir(baseDir As String)
    GLOBALS_.CURRENT_MODULE_PATH = baseDir
End Property
Public Property Get WorkingDir() As String
    WorkingDir = GLOBALS_.CURRENT_MODULE_PATH
End Property
Public Property Get Evaluator() As VBAexpressions
    Set Evaluator = EVALUATOR_
End Property
Public Property Get CastInjectedVars() As Boolean
    CastInjectedVars = VM_.CastInjectedVars
End Property
Public Property Let CastInjectedVars(aValue As Boolean)
    VM_.CastInjectedVars = aValue
End Property
Public Property Get AppAccess() As Boolean
    AppAccess = VM_.AppAccess
End Property
Public Property Let AppAccess(aValue As Boolean)
    VM_.AppAccess = aValue
End Property
Public Property Get OverrideCollMethods() As Boolean
    OverrideCollMethods = VM_.OverrideCollMethods
End Property
Public Property Let OverrideCollMethods(aValue As Boolean)
    VM_.OverrideCollMethods = aValue
End Property
Private Sub Class_Initialize()
    INIT_
End Sub
Public Function Compile(src As String, Optional progName As String = "@anon") As Long
    If GLOBALS_ Is Nothing Then
        INIT_
    End If
    With Compiler_
        .SetGlobals GLOBALS_
        Compile = .CompileProgram(src, progName)
    End With
End Function
Public Function CreateEvaluator(expression As String) As VBAexpressions
    Set EVALUATOR_ = New VBAexpressions
    EVALUATOR_.Create expression
    Set CreateEvaluator = EVALUATOR_
End Function
Public Function GetGlobals()
    Set GetGlobals = GLOBALS_
End Function
Public Function SetGlobals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Function
Public Sub InjectVariable(varName As String, VarValue As Variant)
    VM_.InjectVariable varName, VarValue
End Sub
#If ASF_BYTECODE Then
' Runs program idx on the bytecode backend; False when the compiler does not take it
Private Function RunFast(idx As Long, params As Variant) As Boolean
    If BC_ Is Nothing Then
        Set BC_ = New ASF_Bytecode
        BC_.SetGlobals GLOBALS_
        Set FVM_ = New ASF_FastVM
        FVM_.SetHost GLOBALS_, VM_, BC_
    End If
    If VM_.EnableCallTrace Then
        FALLBACK_ = "call trace"
        Exit Function
    End If
    If Not BC_.Prepare(idx) Then
        FALLBACK_ = BC_.Reason
        Exit Function
    End If
    FVM_.Run idx, params
    vAssignment OUTPUT__, FVM_.OUTPUT_
    BACKEND_ = "bytecode"
    RunFast = True
End Function
#End If
' Runs program idx on the backend that takes it
Private Sub RunIndex(idx As Long, params As Variant)
    FALLBACK_ = vbNullString
    Dim done As Boolean
#If ASF_BYTECODE Then
    If FASTVM_ON_ Then done = RunFast(idx, params)
#Else
    If FASTVM_ON_ Then FALLBACK_ = "the bytecode backend is not compiled in (ASF_BYTECODE)"
#End If
    If Not done Then
        BACKEND_ = "tree"
        Set VM_.Globals = GLOBALS_
        VM_.RunProgramByIndex idx, params
        vAssignment OUTPUT__, VM_.OUTPUT_
    End If
    If VM_.verbose Then
        Debug.Print "=== Runtime Log ==="
        Dim elm As Variant
        For Each elm In GLOBALS_.gRuntimeLog
            Debug.Print elm
        Next elm
    End If
End Sub
Private Sub INIT_()
    Set GLOBALS_ = New ASF_Globals
    GLOBALS_.ASF_InitGlobals
    Set Compiler_ = New ASF_Compiler
    Set VM_ = New ASF_VM
    VM_.verbose = False
    VM_.EnableCallTrace = False
    VM_.AppAccess = False
    VM_.CastInjectedVars = True
    VM_.OverrideCollMethods = False
    Set VM_.Globals = GLOBALS_
End Sub
' The value of the last throw that no catch took, or the message of the last error, of the last Run
Public Property Get LastError() As Variant
#If ASF_BYTECODE Then
    If BACKEND_ = "bytecode" Then
        vAssignment LastError, FVM_.LastError
        Exit Property
    End If
#End If
    vAssignment LastError, VM_.LastError
End Property
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Function Execute(scriptPath As String, ParamArray params() As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        Dim params_ As Variant
        If Not IsMissing(params) Then params_ = params
        RunIndex Compile(ReadTextFile(scriptPath)), params_
        vAssignment Execute, OUTPUT__
    End If
End Function
Public Function Eval(Optional strValues As String) As String
    With EVALUATOR_
        Eval = .Eval(strValues)
    End With
End Function
Public Function Run(idx As Long, ParamArray params() As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        Dim params_ As Variant
        If Not IsMissing(params) Then params_ = params
        RunIndex idx, params_
        vAssignment Run, OUTPUT__
    End If
End Function
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
Public Sub ClearCallStack()
    VM_.ClearCallStack
End Sub
Public Function ReadTextFile(filePath As String) As String
    ReadTextFile = VM_.ReadTextFile(filePath)
End Function
' The separator of folders: the one path uses, or else the one of the current folder; "\" on
' Windows, "/" on macOS and Linux, ":" in the paths of Office 2011 for Mac. Application.PathSeparator
' is not in every host (PowerPoint, Access, Outlook and LibreOffice lack it)
Public Function PathSeparator(Optional ByVal path As String) As String
    Dim sep As String: sep = SeparatorOf(path)
    If LenB(sep) = 0 Then sep = SeparatorOf(CurDir$)
    If LenB(sep) = 0 Then sep = "\"
    PathSeparator = sep
End Function
' folder and name with the separator of folder between them (one, if folder ends with it)
Public Function JoinPath(ByVal folder As String, ByVal name As String) As String
    Dim sep As String: sep = PathSeparator(folder)
    If Right$(folder, 1) = sep Or LenB(folder) = 0 Then JoinPath = folder & name Else JoinPath = folder & sep & name
End Function
Private Function SeparatorOf(ByVal path As String) As String
    If InStr(path, "\") > 0 Then
        SeparatorOf = "\"
    ElseIf InStr(path, "/") > 0 Then
        SeparatorOf = "/"
    ElseIf InStr(path, ":") > 0 And Not (path Like "[A-Za-z]:*") Then
        SeparatorOf = ":"
    End If
End Function
Public Sub ClearModuleCache()
    If Not GLOBALS_ Is Nothing Then
        Set GLOBALS_.gModuleRegistry = New ASF_Map
        Set GLOBALS_.gModuleExports = New ASF_Map
        Set GLOBALS_.gLoadingModules = New Collection
        GLOBALS_.CURRENT_MODULE_PATH = vbNullString
    End If
End Sub
Public Function VarToASF(VarValue) As Variant
    VarToASF = VM_.VarToASF(VarValue)
End Function
