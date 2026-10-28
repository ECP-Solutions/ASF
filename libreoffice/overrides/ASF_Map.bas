' ASF_Map for LibreOffice Basic: same public members as src/ASF_Map.cls.
' The VBA class keeps keys and values in a user-defined type whose array fields it
' resizes; LibreOffice cannot ReDim a field of a Type, so this one uses plain arrays.
' Keep it in step with src/ASF_Map.cls. Clone is named LOCloneMap here because
' LibreOffice will not call a method named Clone; tools/build.py renames the callers.
Option Explicit
Private kCap As Long, kIdx As Long
Private kSt() As Variant
Private vSt() As Variant
Private Sub Class_Initialize()
    BufferInit
End Sub
Private Sub BufferInit()
    kCap = 64: kIdx = -1
    ReDim kSt(0 To kCap - 1)
    ReDim vSt(0 To kCap - 1)
End Sub
Public Function Exists(key As String) As Boolean
    Dim i As Long, r As Boolean
    Do
        r = (kSt(i) = key)
        i = i + 1
    Loop While i <= kIdx And Not r
    Exists = r
End Function
Private Function GetKeyIndex(key As String) As Long
    Dim i As Long, r As Boolean
    Do
        r = (kSt(i) = key)
        i = i + 1
    Loop While i <= kIdx And Not r
    If r Then GetKeyIndex = i - 1 Else GetKeyIndex = -1
End Function
Public Sub Add(key As String, value As Variant)
    If Exists(key) Then Err.Raise -2147221504 + 10000, "Map.Add", "Key already exists: " & key
    kIdx = kIdx + 1
    If kIdx > kCap - 1 Then
        kCap = 2 * (kCap + 1)
        ReDim Preserve kSt(0 To kCap - 1)
        ReDim Preserve vSt(0 To kCap - 1)
    End If
    kSt(kIdx) = key
    If IsObject(value) Then Set vSt(kIdx) = value Else vSt(kIdx) = value
End Sub
Public Sub SetValue(key As String, value As Variant)
    If Exists(key) Then
        Dim idx As Long: idx = GetKeyIndex(key)
        If IsObject(value) Then Set vSt(idx) = value Else vSt(idx) = value
    Else
        Add key, value
    End If
End Sub
Public Function GetValue(key As String) As Variant
    Dim idx As Long: idx = GetKeyIndex(key)
    If idx > -1 Then
        If IsObject(vSt(idx)) Then Set GetValue = vSt(idx) Else GetValue = vSt(idx)
    Else
        Set GetValue = Nothing
    End If
End Function
Public Sub Remove(key As String)
    Dim index As Long: index = GetKeyIndex(key)
    If index > -1 Then
        Dim i As Long
        For i = index To kIdx - 1
            kSt(i) = kSt(i + 1)
            If IsObject(vSt(i + 1)) Then Set vSt(i) = vSt(i + 1) Else vSt(i) = vSt(i + 1)
        Next i
        kSt(kIdx) = Empty: vSt(kIdx) = Empty
        kIdx = kIdx - 1
    End If
End Sub
Public Function key(index As Long) As String
    key = kSt(index - 1)
End Function
Public Function keys() As Collection
    Dim out As Collection: Set out = New Collection
    Dim i As Long
    For i = 0 To kIdx
        out.Add kSt(i)
    Next i
    Set keys = out
End Function
Public Function items() As Collection
    Dim out As Collection: Set out = New Collection
    Dim i As Long
    For i = 0 To kIdx
        out.Add vSt(i)
    Next i
    Set items = out
End Function
Public Function Count() As Long
    Count = kIdx + 1
End Function
Public Sub Clear()
    BufferInit
End Sub
Public Function LOCloneMap() As ASF_Map
    Dim m As Object: Set m = NewASFMap()
    Dim k As Long
    For k = 0 To kIdx
        If LOTypeName(vSt(k)) = "ASF_Map" Then
            m.Add CStr(kSt(k)), vSt(k).LOCloneMap()
        Else
            m.Add CStr(kSt(k)), vSt(k)
        End If
    Next k
    Set LOCloneMap = m
End Function
