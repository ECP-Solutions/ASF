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
Public knownMSOcollections  As Variant
Private Const MSOColStr As String = "ActionSettings,Actions,AddIns,AddressEntries,AddressLists,Adjustments,AllDatabaseDiagrams,AllForms,AllFunctions," & _
                                    "AllMacros,AllModules,AllQueries,AllReports,AllStoredProcedures,AllTables,AllViews,AllowEditRanges,AnimationBehaviors," & _
                                    "AnimationPoints,Areas,Axes,Bookmarks,Borders,Breaks,CalculatedFields,CalculatedItems,CalculatedMembers,Cells," & _
                                    "Characters,ChartGroups,ChartObjects,Charts,Collection,Collections,Columns,Comments,Containers,Controls,CubeFields," & _
                                    "CustomLabels,CustomProperties,CustomViews,DataLabels,Departments,Dialogs,Dictionaries,Dictionary,DocumentWindows," & _
                                    "Documents,Editors,EmailSignatureEntries,Endnotes,FieldLinks,Fields,FileConverters,Filters,FirstLetterExceptions," & _
                                    "FontNames,Fonts,Footnotes,FormFields,Frames,GroupShapes,HeadersFooters,HeadingStyles,Hyperlinks,Indexes,InlineShapes," & _
                                    "ItemProperties,JPEGFiles,KeyBindings,KeysBoundTo,Languages,LegendEntries,Lines,ListColumns,ListEntries,ListItems," & _
                                    "ListLevels,ListObjects,ListParagraphs,ListRows,ListTemplates,Lists,MailMergeDataFields,MailMergeFieldNames,MailMergeFields," & _
                                    "MappedDataFields,NameSpaces,NamedSlideShows,Names,ODBCErrors,OLEDBErrors,OLEObjects,OtherCorrectionsExceptions," & _
                                    "OutlookBarGroups,OutlookBarShortcuts,Pages,Panes,Paragraphs,Parameters,Phonetics,PivotCaches,PivotFields,PivotFormulas," & _
                                    "PivotItemList,PivotItems,PivotTables,Placeholders,Points,PrintRanges,ProofreadingErrors,PublishObjects,QueryTables," & _
                                    "Recipients,Rectangles,Regions,Relationships,Reminders,Reports,Results,Revisions,Rules,Scenes,Sections,Sequence,Sequences," & _
                                    "SeriesCollection,ShapeNodes,Shapes,Slicers,SlideRange,SlideShowWindows,Slides,SmartTagActions,SmartTagRecognizers," & _
                                    "SmartTagTypes,SmartTags,StyleSheets,Styles,Subdocuments,SyncObjects,TabStops,Tabs,TaskPanes,Tasks,Templates,TextColumns," & _
                                    "TextStyleLevels,TextStyles,Trendlines,Trends,UsedObjects,VPageBreaks,Variables,Views,Visualizations,Watches,Windows,Words," & _
                                    "XMLChildNodeSuggestions,XMLNamespaces,XMLNodes,XMLSchemaReferences,XSLTransforms,XmlMaps,Zooms."
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
        builtinArrayMethods = Array("at", "concat", "copywithin", "delete", "entries", "every", "filter", "find", "findindex", _
                                    "findlast", "findlastindex", "foreach", "from", "includes", "indexof", "join", "lastindexof", _
                                    "length", "map", "of", "pop", "push", "reduce", "reverse", "shift", "slice", "some", "sort", _
                                    "splice", "toreversed", "tosorted", "tospliced", "tostring", "unique", "unshift", "with")
    End If
    If IsEmpty(knownMSOcollections) Then
        knownMSOcollections = Split(MSOColStr, ",")
    End If
    Set ASF_InitGlobals = Me
End Function
Public Property Get MAX_CALL_DEPTH() As Long
    MAX_CALL_DEPTH = MAX_CALL_DEPTH_
End Property
Public Property Let MAX_CALL_DEPTH(aValue As Long)
    MAX_CALL_DEPTH_ = aValue
End Property
