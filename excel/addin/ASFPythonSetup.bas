Attribute VB_Name = "ASFPythonSetup"
' The Python buttons of the ASF tab, after the ribbon (ASFRibbon): Python studio, Python console,
' Python settings, Python updates and About Python. The Python add-in (asfpy64.xll, asfpy32.xll) opens the studio
' and the console through its commands AsfPyStudio and AsfPyConsole; the settings are those it
' reads when Excel loads it (ASFSettings). Python runs on Excel for Windows only.
Option Explicit
Option Private Module

Private Const TITLE As String = "ASF Python"

' Python studio: the add-in's command AsfPyStudio opens it, or its window when it runs already
Public Sub Studio()
    Dim py As Object
    If Not Ready(py) Then Exit Sub
    On Error GoTo Failed
    Application.Run "AsfPyStudio"
    Exit Sub
Failed:
    MsgBox "Python studio failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python console: the add-in's Python at its prompt, in a console window of its own (its command
' AsfPyConsole), with pip("...") and the Python folder importable
Public Sub Console()
    Dim py As Object
    If Not Ready(py) Then Exit Sub
    On Error GoTo Failed
    Application.Run "AsfPyConsole"
    Exit Sub
Failed:
    MsgBox "Python console failed: " & Err.Description, vbExclamation, TITLE
End Sub

' Python settings: which Python, the Python folder, PythonAccess and PythonUpdates, each shown with
' its value; the add-in reads the first two when Excel loads it
Public Sub Settings()
    Dim choice As String, restart As Boolean
#If Mac Then
    MsgBox "Python needs the Python add-in, which runs on Excel for Windows.", vbInformation, TITLE
    Exit Sub
#End If
    Do
        choice = InputBox("1. Python: " & PythonInUse() & vbLf & _
                          "2. Python folder: " & ASFSettings.PythonFolder() & vbLf & _
                          "3. Python for library scripts (PythonAccess): " & ASFSettings.PythonAccess() & vbLf & _
                          "4. Updates of the Python add-in (PythonUpdates): " & ASFSettings.PythonUpdates() & vbLf & vbLf & _
                          "Number of the setting to change (Cancel when done):", TITLE & ": settings")
        Select Case Trim$(choice)
            Case "": Exit Do
            Case "1": If ChoosePython() Then restart = True
            Case "2": If ChooseFolder() Then restart = True
            Case "3": ChooseAccess
            Case "4": ChooseUpdates
        End Select
    Loop
    If restart Then MsgBox "Restart Excel to use the new settings: the Python add-in reads the Python and the Python folder when Excel loads it.", vbInformation, TITLE
End Sub

' About Python: the add-in, its version and Python, the settings and what the Python folder gave;
' offers to check for a newer add-in (its command AsfPyUpdate) or to open the log
Public Sub About()
    Dim py As Object, message As String, text As String, info As String
    text = "Add-in: " & AddinPath() & vbLf & _
           "Python: " & PythonInUse() & vbLf & _
           "Python folder: " & ASFSettings.PythonFolder() & vbLf & _
           "PythonAccess: " & ASFSettings.PythonAccess() & vbLf & _
           "PythonUpdates: " & ASFSettings.PythonUpdates() & vbLf & _
           "Log: " & LogPath() & vbLf & vbLf
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        text = text & message
    Else
        On Error Resume Next
        info = CStr(CallByName(py, "startInfo", VbGet))
        On Error GoTo 0
        If Len(info) > 600 Then info = Left$(info, 600) & "..."
        text = text & info
    End If
    Select Case MsgBox(text & vbLf & vbLf & "Yes: check for a newer add-in (Python updates). No: open the log.", _
                       vbYesNoCancel + vbInformation + vbDefaultButton3, TITLE & ": about")
        Case vbYes: CheckForUpdates
        Case vbNo: OpenText LogPath()
    End Select
End Sub

' Python updates: the Add-in page of Python studio, with the version, Check for updates, Install,
' and when Excel looks; with another Python than the add-in's own (PythonVenv), which has no
' studio, the same in message boxes
Public Sub Updates()
    Dim version As String
#If Mac Then
    MsgBox "Python needs the Python add-in, which runs on Excel for Windows.", vbInformation, TITLE
    Exit Sub
#End If
    If LenB(ASFSettings.PythonVenv()) = 0 Then
        On Error Resume Next
        Application.Run "AsfPyStudio", "addin"
        If Err.Number = 0 Then Exit Sub
        On Error GoTo 0
    End If
    version = AddinVersion()
    Select Case MsgBox(IIf(LenB(version) > 0, "ASF Python " & version, "The Python add-in: " & AddinPath()) & vbLf & _
                       "When Excel starts (PythonUpdates): " & ASFSettings.PythonUpdates() & vbLf & vbLf & _
                       "Yes: check for a newer add-in now." & vbLf & "No: change when Excel looks for one.", _
                       vbYesNoCancel + vbInformation, TITLE & ": updates")
        Case vbYes: CheckForUpdates
        Case vbNo: ChooseUpdates
    End Select
End Sub

' The add-in's command AsfPyUpdate: it asks the project's releases, offers a newer add-in, and says
' what came of it in a message box of its own, apart from Excel
Public Sub CheckForUpdates()
    On Error GoTo Failed
    Application.Run "AsfPyUpdate"
    Exit Sub
Failed:
    MsgBox "The Python add-in isn't loaded, or is older than its updates: the releases are at" & vbLf & _
           "https://github.com/ECP-Solutions/ASF/releases", vbExclamation, TITLE
End Sub

' --- helpers ---------------------------------------------------------------------------------------

' the add-in loaded and Python running; else a message and False
Private Function Ready(ByRef py As Object) As Boolean
    Dim message As String
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then
        MsgBox message & vbLf & vbLf & "Python settings and About Python on the ASF tab help set it up.", vbExclamation, TITLE
    Else
        Ready = True
    End If
End Function

' the version of the loaded add-in (py.startInfo starts with ASF Python x.y.z), or ""
Private Function AddinVersion() As String
    Dim py As Object, message As String, info As String
    Set py = ASFPython.Root(1, message)
    If py Is Nothing Then Exit Function
    On Error Resume Next
    info = CStr(CallByName(py, "startInfo", VbGet))
    On Error GoTo 0
    If Left$(info, 11) = "ASF Python " Then AddinVersion = Split(Mid$(info, 12), vbLf)(0)
End Function

Private Function PythonInUse() As String
    Dim venv As String: venv = ASFSettings.PythonVenv()
    If LenB(venv) = 0 Then PythonInUse = "the add-in's own" Else PythonInUse = venv
End Function

Private Function LogPath() As String
    LogPath = Environ$("TEMP") & "\asfpy.log"
End Function

' the loaded asfpy*.xll, or where to load it from
Private Function AddinPath() As String
    Dim ai As Object
    On Error Resume Next
    For Each ai In Application.AddIns
        If LCase$(ai.Name) Like "asfpy*.xll" And ai.Installed Then
            AddinPath = ai.FullName
            Exit Function
        End If
    Next ai
    AddinPath = "not loaded (File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse)"
End Function

' Which Python: the add-in's own, or another (a venv or an install), picked here
Private Function ChoosePython() As Boolean
    Dim answer As Long, folder As String
    answer = MsgBox("Which Python should the Python add-in run?" & vbLf & vbLf & _
                    "Yes: the Python the add-in carries (recommended). Packages in Python studio adds packages to it." & vbLf & vbLf & _
                    "No: another Python of Excel's bitness, 3.11 or later: a venv or a python.org install, picked next.", _
                    vbYesNoCancel + vbQuestion, TITLE & ": Python")
    If answer = vbYes Then
        If LenB(ASFSettings.PythonVenv()) > 0 Then
            ASFSettings.ClearSetting "PythonVenv"
            ChoosePython = True
        End If
    ElseIf answer = vbNo Then
        folder = PickFolder("Folder of the venv or of the Python install", ASFSettings.PythonVenv())
        If LenB(folder) = 0 Then Exit Function
        If LenB(Dir(folder & "\pyvenv.cfg")) = 0 And LenB(Dir(folder & "\python3*.dll")) = 0 Then
            MsgBox folder & vbLf & "has neither pyvenv.cfg (a venv) nor python3XX.dll (an install): it isn't a Python.", vbExclamation, TITLE
            Exit Function
        End If
        ASFSettings.ChangeSetting "PythonVenv", folder
        ChoosePython = True
    End If
End Function

Private Function ChooseFolder() As Boolean
    Dim folder As String
    folder = PickFolder("Python folder: the modules whose functions Excel gets", ASFSettings.PythonFolder())
    If LenB(folder) = 0 Then Exit Function
    ASFSettings.ChangeSetting "PythonFolder", folder
    ChooseFolder = True
End Function

' PythonAccess applies at once: library scripts read it each time they run
Private Sub ChooseAccess()
    Dim choice As String, current As Long
    current = InStr("offcallmodules", ASFSettings.PythonAccess())
    current = IIf(current = 1, 1, IIf(current = 4, 2, 3))
    choice = InputBox("What library scripts whose header asks for Python may do:" & vbLf & vbLf & _
                      "1. off: no Python (the default)" & vbLf & _
                      "2. call: py.call of the functions exported with @export" & vbLf & _
                      "3. modules: py.module too, which imports any module: everything Python can do" & vbLf & vbLf & _
                      "Formulas (ASF_RUN) never get Python, whatever this says.", TITLE & ": PythonAccess", CStr(current))
    Select Case Trim$(choice)
        Case "1": ASFSettings.ChangeSetting "PythonAccess", "off"
        Case "2": ASFSettings.ChangeSetting "PythonAccess", "call"
        Case "3": ASFSettings.ChangeSetting "PythonAccess", "modules"
    End Select
End Sub

' PythonUpdates applies the next time Excel loads the add-in
Private Sub ChooseUpdates()
    Dim choice As String, current As String
    current = IIf(ASFSettings.PythonUpdates() = "check", "1", IIf(ASFSettings.PythonUpdates() = "auto", "2", "3"))
    choice = InputBox("When Excel loads the Python add-in, at most once a day:" & vbLf & vbLf & _
                      "1. check: offer a newer add-in from the project's releases (the default)" & vbLf & _
                      "2. auto: download and install it, then say so" & vbLf & _
                      "3. off: don't look; About Python still checks when asked" & vbLf & vbLf & _
                      "An update is checked (SHA-256, bitness, version) before it replaces the add-in, and Excel uses it after a restart.", _
                      TITLE & ": PythonUpdates", current)
    Select Case Trim$(choice)
        Case "1": ASFSettings.ChangeSetting "PythonUpdates", "check"
        Case "2": ASFSettings.ChangeSetting "PythonUpdates", "auto"
        Case "3": ASFSettings.ChangeSetting "PythonUpdates", "off"
    End Select
End Sub

Private Function PickFolder(ByVal caption As String, ByVal start As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)          ' msoFileDialogFolderPicker
    fd.Title = TITLE & ": " & caption
    If LenB(start) > 0 Then fd.InitialFileName = start & "\"
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
End Function

Private Sub OpenText(ByVal path As String)
    On Error Resume Next
    Shell "notepad.exe """ & path & """", vbNormalFocus
End Sub
