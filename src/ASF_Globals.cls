VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Globals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module: Globals
Option Explicit

Public gPrograms As Collection          ' each item: Variant array (name, stmts Collection, scope raw Collection)
Public gFuncTable As ASF_Map               ' name -> program index
Public gFuncParams As ASF_Map              ' name -> parameter name array
Public ExpressionCache As ASF_Map          ' expr string -> compiled RPN or cached info (optional)
Public ProgramCache As ASF_Map             ' source string -> program index
Public gModuleRegistry As ASF_Map          ' modulePath -> Module object
Public gModuleExports As ASF_Map           ' modulePath -> exports map
Public gLoadingModules As Collection       ' Circular dependency detection
Public CURRENT_MODULE_PATH As String       ' Currently executing module
Public gRuntimeLog As Collection
Public DEBUG_TRACE As Boolean
Public STRICT_STRINGS As Boolean
Public gCallStack As Collection
Public gExprEvaluator As VBAexpressions
Public re As ASF_RegexEngine
Private MAX_CALL_DEPTH_ As Long
Public builtinArrayMethods As Variant

Public Function ASF_InitGlobals() As ASF_Globals
    If gPrograms Is Nothing Then Set gPrograms = New Collection
    If gFuncTable Is Nothing Then Set gFuncTable = New ASF_Map
    If gFuncParams Is Nothing Then Set gFuncParams = New ASF_Map
    If ExpressionCache Is Nothing Then Set ExpressionCache = New ASF_Map
    If ProgramCache Is Nothing Then Set ProgramCache = New ASF_Map
    If gRuntimeLog Is Nothing Then Set gRuntimeLog = New Collection
    If gCallStack Is Nothing Then Set gCallStack = New Collection
    If gModuleRegistry Is Nothing Then Set gModuleRegistry = New ASF_Map
    If gModuleExports Is Nothing Then Set gModuleExports = New ASF_Map
    If gLoadingModules Is Nothing Then Set gLoadingModules = New Collection
    If gExprEvaluator Is Nothing Then Set gExprEvaluator = New VBAexpressions
    If re Is Nothing Then Set re = New ASF_RegexEngine
    DEBUG_TRACE = True
    STRICT_STRINGS = True
    MAX_CALL_DEPTH_ = 512
    CURRENT_MODULE_PATH = vbNullString
    If IsEmpty(builtinArrayMethods) Then
        builtinArrayMethods = Array("push", "pop", "foreach", "unique", "concat", _
                                    "tostring", "join", "shift", "unshift", "delete", _
                                    "splice", "tospliced", "at", "copywithin", "entries", _
                                    "every", "find", "findindex", "findlast", "findlastindex", _
                                    "from", "includes", "indexof", "lastindexof", "of", _
                                    "reverse", "toreversed", "some", "with", "sort", _
                                    "tosorted")
    End If
    Set ASF_InitGlobals = Me
End Function
Public Property Get MAX_CALL_DEPTH() As Long
    MAX_CALL_DEPTH = MAX_CALL_DEPTH_
End Property
Public Property Let MAX_CALL_DEPTH(aValue As Long)
    MAX_CALL_DEPTH_ = aValue
End Property
