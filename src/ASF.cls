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
Private CAST_INJECTED_VARIABLES_ As Boolean
Private APPLICATION_ACCESS_ As Boolean

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
    CastInjectedVars = CAST_INJECTED_VARIABLES_
End Property
Public Property Let CastInjectedVars(aValue As Boolean)
    CAST_INJECTED_VARIABLES_ = aValue
End Property
Public Property Get AppAccess() As Boolean
    AppAccess = APPLICATION_ACCESS_
End Property
Public Property Let AppAccess(aValue As Boolean)
    APPLICATION_ACCESS_ = aValue
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
Private Sub INIT_()
    Set GLOBALS_ = New ASF_Globals
    GLOBALS_.ASF_InitGlobals
    Set Compiler_ = New ASF_Compiler
    Set VM_ = New ASF_VM
    VM_.verbose = False
    VM_.EnableCallTrace = False
    CAST_INJECTED_VARIABLES_ = True
    VM_.AppAccess = APPLICATION_ACCESS_
    VM_.CastInjectedVars = CAST_INJECTED_VARIABLES_
End Sub
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Function Execute(scriptPath As String) As Variant
    vAssignment Execute, Run(Compile(ReadTextFile(scriptPath)))
End Function
Public Function Eval(Optional strValues As String) As String
    With EVALUATOR_
        Eval = .Eval(strValues)
    End With
End Function
Public Function Run(idx As Long, ParamArray params() As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        With VM_
            Set .Globals = GLOBALS_
            .CastInjectedVars = CAST_INJECTED_VARIABLES_
            .AppAccess = APPLICATION_ACCESS_
            If Not IsMissing(params) Then
                Dim params_ As Variant: params_ = params
                .RunProgramByIndex idx, params_
            Else
                .RunProgramByIndex idx
            End If
            vAssignment OUTPUT__, .OUTPUT_
            vAssignment Run, .OUTPUT_
                If .verbose Then
                    Debug.Print "=== Runtime Log ==="
                    Dim elm As Variant
                    For Each elm In GLOBALS_.gRuntimeLog
                        Debug.Print elm
                    Next elm
                End If
        End With
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
