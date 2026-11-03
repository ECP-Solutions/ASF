VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Map
Option Explicit

Private Type Buffer
    Capacity As Long
    index As Long
    Storage() As Variant
End Type
Private pItems As Buffer   ' values, in the order their keys were added
' keys in a String array: comparing two Strings is cheaper than a Variant with a String
Private pKeys() As String
Private emptyVariant As Variant

Private Sub Class_Initialize()
    BufferInit
End Sub
 
' AST nodes and most script objects hold a handful of keys; the buffers double as needed
Private Sub BufferInit()
    pItems.Capacity = 8
    pItems.index = -1
    ReDim pItems.Storage(0 To pItems.Capacity - 1)
    ReDim pKeys(0 To pItems.Capacity - 1)
End Sub
Public Function Exists(key As String) As Boolean
    Exists = (GetKeyIndex(key) > -1)
End Function

Private Function GetKeyIndex(key As String) As Long
    Dim i As Long
    For i = 0 To pItems.index
        If pKeys(i) = key Then
            GetKeyIndex = i
            Exit Function
        End If
    Next i
    GetKeyIndex = -1
End Function

Private Sub ExpandBuffer()
    pItems.Capacity = 2 * (pItems.Capacity + 1)
    ReDim Preserve pItems.Storage(0 To pItems.Capacity - 1)
    ReDim Preserve pKeys(0 To pItems.Capacity - 1)
End Sub

' Add: raise error if key exists (Dictionary-like Add)
Public Sub Add(key As String, value As Variant)
    If GetKeyIndex(key) > -1 Then
        err.Raise vbObjectError + 10000, "Map.Add", "Key already exists: " & key
    End If
    AppendBufferValue key, value
End Sub
Private Sub AppendBufferValue(key As String, value As Variant)
    pItems.index = pItems.index + 1
    If pItems.index >= pItems.Capacity Then ExpandBuffer
    pKeys(pItems.index) = key
    vAssignment pItems.Storage(pItems.index), value
End Sub

' SetValue: set or replace value for key
Public Sub SetValue(key As String, value As Variant)
    Dim idx As Long
    idx = GetKeyIndex(key)
    If idx > -1 Then 'Update
        vAssignment pItems.Storage(idx), value
    Else 'Append
        AppendBufferValue key, value
    End If
End Sub

' ReplaceValue: set the value of an existing key; False, and nothing changed, when key is missing
Public Function ReplaceValue(key As String, value As Variant) As Boolean
    Dim idx As Long
    idx = GetKeyIndex(key)
    If idx > -1 Then
        vAssignment pItems.Storage(idx), value
        ReplaceValue = True
    End If
End Function

' GetValue (returns Nothing if not exists); the VM's most frequent call, so the lookup is inline
Public Property Get GetValue(key As String) As Variant
Attribute GetValue.VB_UserMemId = 0
    Dim i As Long
    For i = 0 To pItems.index
        If pKeys(i) = key Then
            If IsObject(pItems.Storage(i)) Then Set GetValue = pItems.Storage(i) Else GetValue = pItems.Storage(i)
            Exit Property
        End If
    Next i
    Set GetValue = Nothing
End Property

' TryGetValue: one lookup for Exists and GetValue; value is left alone when key is missing
Public Function TryGetValue(key As String, ByRef value As Variant) As Boolean
    Dim i As Long
    For i = 0 To pItems.index
        If pKeys(i) = key Then
            If IsObject(pItems.Storage(i)) Then Set value = pItems.Storage(i) Else value = pItems.Storage(i)
            TryGetValue = True
            Exit Function
        End If
    Next i
End Function

Public Sub Remove(key As String)
    Dim index As Long
    Dim i As Long
    index = GetKeyIndex(key)
    If index > -1 Then
        For i = index To pItems.index - 1
            pKeys(i) = pKeys(i + 1)
            vAssignment pItems.Storage(i), pItems.Storage(i + 1)
        Next i
        pKeys(pItems.index) = vbNullString
        pItems.Storage(pItems.index) = emptyVariant
        pItems.index = pItems.index - 1
    End If
End Sub
' Keys: return a key given an index (read-only copy)
Public Function key(index As Long) As String
    key = pKeys(index - 1)
End Function
' Keys: return a Collection of keys (read-only copy)
Public Function keys() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pItems.index
        out.Add pKeys(i)
    Next i
    Set keys = out
End Function
' Items: return a Collection of values (in same order as keys)
Public Function items() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pItems.index
        out.Add pItems.Storage(i)
    Next i
    Set items = out
End Function
' Count property
Public Property Get Count() As Long
    Count = pItems.index + 1
End Property
' Clear
Public Sub Clear()
    BufferInit
End Sub
' Clone (deep for ASF_Map values, other values are copied as Variants)
Public Function Clone() As ASF_Map
    Dim m As New ASF_Map
    Dim k As Long
    Dim child As ASF_Map
    For k = 0 To pItems.index
        If TypeName(pItems.Storage(k)) = "ASF_Map" Then
            Set child = pItems.Storage(k)
            m.AppendUnchecked pKeys(k), child.Clone
        Else
            m.AppendUnchecked pKeys(k), pItems.Storage(k)
        End If
    Next k
    Set Clone = m
End Function
' Append for Clone: the source keys are already distinct
Friend Sub AppendUnchecked(key As String, value As Variant)
    AppendBufferValue key, value
End Sub
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
