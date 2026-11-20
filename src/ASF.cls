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
' ASF: compiles scripts to bytecode (ASF_Compiler) and runs them (ASF_VM, with the library of
' values of ASF_Runtime). The classes of ASF 3, which ran the AST, are in src/old with the suffix _OLD.
Private Const ASF_VERSION As String = "4.0.2"
Private GLOBALS_ As ASF_Globals
Private RT_ As ASF_Runtime
Private CP_ As ASF_Compiler
Private VM_ As ASF_VM
Private CP_GLOBALS_ As ASF_Globals      ' the globals the compiler and the VM belong to
Private MAX_STEPS_ As Long              ' MaxSteps, kept for a VM made later
Private OUTPUT__ As Variant
Private EVALUATOR_ As VBAexpressions

Public Property Get Version() As String
    Version = ASF_VERSION
End Property
Public Property Get verbose() As Boolean
    verbose = RT_.verbose
End Property
Public Property Let verbose(aValue As Boolean)
    RT_.verbose = aValue
End Property
Public Property Get EnableCallTrace() As Boolean
    EnableCallTrace = RT_.EnableCallTrace
End Property
Public Property Let EnableCallTrace(aValue As Boolean)
    RT_.EnableCallTrace = aValue
End Property
Public Property Get GetCallStackTrace() As String
    GetCallStackTrace = RT_.GetCallStackTrace
End Property
Public Property Get GetCallStack() As Collection
    Set GetCallStack = RT_.GetCallStack
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
    CastInjectedVars = RT_.CastInjectedVars
End Property
Public Property Let CastInjectedVars(aValue As Boolean)
    RT_.CastInjectedVars = aValue
End Property
Public Property Get AppAccess() As Boolean
    AppAccess = RT_.AppAccess
End Property
Public Property Let AppAccess(aValue As Boolean)
    RT_.AppAccess = aValue
End Property
' The folder imports are limited to: an import that resolves outside it raises. Empty (the default):
' imports may read any path
Public Property Get ImportRoot() As String
    ImportRoot = RT_.ImportRoot
End Property
Public Property Let ImportRoot(aValue As String)
    RT_.ImportRoot = aValue
End Property
' The loop iterations and calls a run may make before it stops with "Step limit reached" in
' LastError, which no catch takes. 0 (the default): no limit
Public Property Get MaxSteps() As Long
    MaxSteps = MAX_STEPS_
End Property
Public Property Let MaxSteps(aValue As Long)
    If aValue < 0 Then aValue = 0
    MAX_STEPS_ = aValue
    If Not VM_ Is Nothing Then VM_.MaxSteps = aValue
End Property
Public Property Get OverrideCollMethods() As Boolean
    OverrideCollMethods = RT_.OverrideCollMethods
End Property
Public Property Let OverrideCollMethods(aValue As Boolean)
    RT_.OverrideCollMethods = aValue
End Property
Private Sub Class_Initialize()
    INIT_
End Sub
' The program index of src; an error in it raises a compile error
Public Function Compile(src As String, Optional progName As String = "@anon") As Long
    If GLOBALS_ Is Nothing Then INIT_
    PrepareCode
    Compile = CP_.CompileSource(src, progName)
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
    RT_.InjectVariable varName, VarValue
End Sub
' An object the host trusts, which stays an object while AppAccess is off (InjectVariable makes it
' its type name then): the host decides what scripts may reach, as the Python add-in's py
Public Sub InjectHostObject(varName As String, obj As Object)
    RT_.InjectHostObject varName, obj
End Sub
' The compiler and the VM of the current globals: compiled programs belong to the globals they
' were compiled with (SetGlobals changes them)
Private Sub PrepareCode()
    If Not CP_ Is Nothing Then
        If Not CP_GLOBALS_ Is GLOBALS_ Then Set CP_ = Nothing
    End If
    If CP_ Is Nothing Then
        Set CP_ = New ASF_Compiler
        CP_.SetGlobals GLOBALS_
        Set CP_GLOBALS_ = GLOBALS_
        Set VM_ = New ASF_VM
        VM_.SetHost GLOBALS_, RT_, CP_
        VM_.MaxSteps = MAX_STEPS_
    End If
    Set RT_.Globals = GLOBALS_
End Sub
Private Sub RunIndex(idx As Long, params As Variant)
    PrepareCode
    If Not VM_.Run(idx, params) Then Err.Raise vbObjectError + 9026, "ASF.Run", "No program " & idx & " was compiled with these globals"
    vAssignment OUTPUT__, VM_.RunOutput()
    If RT_.verbose Then
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
    Set RT_ = New ASF_Runtime
    RT_.verbose = False
    RT_.EnableCallTrace = False
    RT_.AppAccess = False
    RT_.CastInjectedVars = True
    RT_.OverrideCollMethods = False
    Set RT_.Globals = GLOBALS_
End Sub
' The value of the last throw that no catch took, or the message of the last error, of the last Run
' LastError and OUTPUT_ are functions: LibreOffice keeps the value of a Property Get between calls,
' and fails to replace an array there
Public Function LastError() As Variant
    If VM_ Is Nothing Then Exit Function
    vAssignment LastError, VM_.LastError
End Function
Public Function OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Function
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
' Run with the arguments in an array ($1 is its first item), for hosts that pass a varying number of
' them; Empty for none
Public Function RunArgs(idx As Long, args As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        RunIndex idx, args
        vAssignment RunArgs, OUTPUT__
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
    RT_.ClearCallStack
End Sub
Public Function ReadTextFile(filePath As String) As String
    ReadTextFile = RT_.ReadTextFile(filePath)
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
    If Not VM_ Is Nothing Then VM_.ClearModules
End Sub
Public Function VarToASF(VarValue) As Variant
    VarToASF = RT_.VarToASF(VarValue)
End Function
