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
Private Const MSOColStr As String = "Accessories,Accounts,ActionSettings,Actions,AddIns,AddressBooks,AddressEntries,AddressLists,Adjustments,AllDatabaseDiagrams," & _
                                    "AllForms,AllFunctions,AllMacros,AllModules,AllQueries,AllReports,AllStoredProcedures,AllTables,AllViews,Animations,AnimationBehaviors," & _
                                    "AnimationPoints,Applications,Areas,Attachments,Axes,Bookmarks,Borders,Breaks,BuildingBlocks,BuildingBlockTypes,CalculatedFields," & _
                                    "CalculatedItems,CalculatedMembers,Categories,Cells,Characters,ChartColors,ChartGroups,ChartObjects,Charts,CodeData,CodeProject," & _
                                    "Collection,Collections,ColorSchemes,Columns,CommandBars,CommandBarControls,Comments,Connections,Containers,ContentControls,Controls," & _
                                    "ConversationHeader,CubeFields,CurrentData,CurrentProject,CustomLabels,CustomLayouts,CustomProperties,CustomTaskPanes,CustomViews," & _
                                    "CustomXMLParts,DataAccessPages,DataLabels,Databases,DataSources,DBEngine,Departments,Dialogs,Dictionaries,Dictionary,DocumentLibraryVersions," & _
                                    "Documents,DocumentWindows,Editors,EmailSignatureEntries,Endnotes,Errors,ExchangeDistributionList,ExchangeUser,Explorers,ExtraColors,FieldLinks," & _
                                    "Fields,FileConverters,FileDialogFilters,FileDialogSelectedItems,Filters,FirstLetterExceptions,Folders,FontNames,Fonts,Footnotes,FormatConditions," & _
                                    "FormFields,Forms,Frames,GroupShapes,Groups,HeadersFooters,HeadingStyles,HPageBreaks,Hyperlinks,Icons,Indexes,InlineShapes,Inspectors,Interior," & _
                                    "ItemProperties,Items,JPEGFiles,KeyBindings,KeysBoundTo,Languages,LanguageSettings,LegendEntries,Lines,ListColumns,ListEntries,ListItems," & _
                                    "ListLevels,ListObjects,ListParagraphs,ListRows,ListTemplates,Lists,MailMergeDataFields,MailMergeFieldNames,MailMergeFields,MappedDataFields," & _
                                    "ModelChanges,ModelConnection,ModelMeasures,ModelRelationships,ModelTables,Modules,MotionEffects,MsoDebugOptions,Names,NamedSlideShows," & _
                                    "NameSpaces,NewFile,ODBCErrors,OLEDBErrors,OLEObjects,OMathFunctions,OtherCorrectionsExceptions,Outline,OutlookBarGroups,OutlookBarShortcuts," & _
                                    "Pages,Panes,Paragraphs,Parameters,Phonetics,PickerResults,PivotCaches,PivotFields,PivotFormulas,PivotItemList,PivotItems,PivotLines,PivotTables," & _
                                    "Placeholders,Points,Presentations,PrintRanges,ProofreadingErrors,Properties,PropertyPages,Protection,PublishObjects,QueryDefs,QueryTables,Recipients," & _
                                    "RecentFiles,Rectangles,References,Regions,Relations,Relationships,Reminders,Reports,Research,Results,Revisions,Rules,Scenes,Sections,Selection," & _
                                    "Sentences,Sequence,Sequences,SeriesCollection,SeriesLines,ShapeNodes,Shapes,Sheets,Signatures,Slicers,SlideRange,Slides,SlideShowSettings," & _
                                    "SlideShowWindows,SlideSizes,SmartArtColors,SmartArtLayouts,SmartArtNodes,SmartArtQuickStyles,SmartTagActions,SmartTagRecognizers,SmartTagTypes," & _
                                    "SmartTags,Sort,SortFields,SparklineGroups,Stores,StyleSheets,Styles,Subdocuments,SyncObjects,TableDefs,TableOfAuthorities,TableOfContents," & _
                                    "TableOfFigures,Tables,TableStyles,TabStops,Tabs,TaskPanes,Tasks,Templates,TextColumns,TextInput,TextStyleLevels,TextStyles,Timing,Trendlines," & _
                                    "Trends,UsedObjects,Users,Validation,Variables,Views,Visualizations,VPageBreaks,Watches,WebPageFonts,Windows,Workbooks,Worksheets,Workspaces,Words," & _
                                    "XMLChildNodeSuggestions,XMLMapping,XMLNamespaces,XMLNodes,XMLSchemaReferences,XmlDataBinding,XmlMaps,XSLTransforms,Zooms"

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

