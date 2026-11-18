Attribute VB_Name = "ASFSettings"
' The settings of the ASF add-in, kept with SaveSetting: on Windows in the registry, under
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ASF\AddIn (where an administrator can also
' set them, by group policy); on Mac in a preferences file.
'   LibraryFolder   the folder of the scripts ASF_CALL runs and imports read
'   RunEnabled      "0" turns ASF_RUN off: only the scripts of the library run
'   StepLimit       the loop iterations and calls a script may make (MaxSteps; 0: no limit)
' Change them from the Immediate window or a macro, for example:
'   Application.Run "ASF.xlam!ASFSettings.ChangeSetting", "RunEnabled", "0"
Option Explicit
Option Private Module

Private Const APP_NAME As String = "ASF"
Private Const SECTION_NAME As String = "AddIn"
Public Const DEFAULT_STEP_LIMIT As Long = 10000000

' The folder of the library: the setting, or by default ASF\scripts in the user's application data
' (Windows), or ASF/scripts in the folder Office shares on Mac, which Excel can read without asking
Public Function LibraryFolder() As String
    LibraryFolder = Setting("LibraryFolder", DefaultLibraryFolder())
End Function

Public Function RunEnabled() As Boolean
    RunEnabled = (Setting("RunEnabled", "1") <> "0")
End Function

Public Function StepLimit() As Long
    Dim v As String: v = Setting("StepLimit", CStr(DEFAULT_STEP_LIMIT))
    If IsNumeric(v) Then StepLimit = CLng(v) Else StepLimit = DEFAULT_STEP_LIMIT
    If StepLimit < 0 Then StepLimit = 0
End Function

' Saves a setting; the worksheet functions use it from their next call
Public Sub ChangeSetting(ByVal name As String, ByVal value As String)
    SaveSetting APP_NAME, SECTION_NAME, name, value
    ASFEngine.ResetEngine
End Sub

Private Function Setting(ByVal name As String, ByVal default As String) As String
    On Error GoTo Fallback
    Setting = GetSetting(APP_NAME, SECTION_NAME, name, default)
    If LenB(Setting) = 0 Then Setting = default
    Exit Function
Fallback:
    Setting = default
End Function

Private Function DefaultLibraryFolder() As String
#If Mac Then
    ' HOME is the app's sandbox (~/Library/Containers/com.microsoft.Excel/Data): the user's folder is
    ' the part before /Library/Containers
    Dim home As String: home = Environ$("HOME")
    Dim p As Long: p = InStr(home, "/Library/Containers/")
    If p > 0 Then home = Left$(home, p - 1)
    DefaultLibraryFolder = home & "/Library/Group Containers/UBF8T346G9.Office/ASF/scripts"
#Else
    DefaultLibraryFolder = Environ$("APPDATA") & "\ASF\scripts"
#End If
End Function
