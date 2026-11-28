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
' The keys of an object in the order they came, and their values. A small object is searched
' key by key, which is the fastest way for a handful of keys. From INDEX_FROM keys on, the
' object also keeps a hash table: an object used as a dictionary then finds a key in a step or
' two whatever its size, where a search key by key made building it cost the square of its size.
' Each key's hash is computed once and kept, so the table grows without hashing the keys again;
' a lookup compares hashes, whole numbers, before it compares the texts.
' Plain arrays only, no arrays in a Type: LibreOffice Basic runs this class too.
Option Explicit

Private Const INDEX_FROM As Long = 16   ' keys from which the object keeps a hash table

Private pCap As Long            ' the arrays' size
Private pLast As Long           ' position of the last key, -1 when there is none
Private pKeys() As String       ' keys in a String array: comparing two Strings is cheaper than a Variant with a String
Private pVals() As Variant      ' values, in the order their keys were added
Private pHash() As Long         ' each key's hash (HashOf), or 0 while it isn't needed yet
Private pSlots() As Long        ' the hash table: 0 a free slot, else a key's position + 1
Private pMask As Long           ' the table's size - 1, a power of two - 1; -1 while there is no table
Private emptyVariant As Variant

Private Sub Class_Initialize()
    BufferInit
End Sub

' AST nodes and most script objects hold a handful of keys; the buffers double as needed
Private Sub BufferInit()
    pCap = 8
    pLast = -1
    pMask = -1
    ReDim pKeys(0 To pCap - 1)
    ReDim pVals(0 To pCap - 1)
    ReDim pHash(0 To pCap - 1)
    Erase pSlots
End Sub

' The hash of a key: a whole number from 2^23 to 2^24 - 1, never 0. Keys up to 24 characters
' long are hashed whole, longer ones by their length and their first and last 12 characters.
' The compiled program keeps the hashes of the keys it names (ASF_VM), with this function.
Public Function HashOf(ByRef key As String) As Long
    Dim h As Long, n As Long, i As Long
    n = Len(key)
    h = n And &H7FFFFF          ' kept below 2^23, so h * 31 never overflows a Long
    If n <= 24 Then
        For i = 1 To n
            h = ((h * 31) + AscW(Mid$(key, i, 1))) And &H7FFFFF
        Next i
    Else
        For i = 1 To 12
            h = ((h * 31) + AscW(Mid$(key, i, 1))) And &H7FFFFF
        Next i
        For i = n - 11 To n
            h = ((h * 31) + AscW(Mid$(key, i, 1))) And &H7FFFFF
        Next i
    End If
    HashOf = h Or &H800000
End Function

' The position of key, or -1; hash is its HashOf, or 0 to have it computed when it is needed
Private Function FindPos(ByRef key As String, ByVal hash As Long) As Long
    Dim i As Long, s As Long, p As Long
    If pMask < 0 Then
        If hash = 0 Then
            For i = 0 To pLast
                If pKeys(i) = key Then
                    FindPos = i
                    Exit Function
                End If
            Next i
        Else
            ' a hash known already: whole numbers compared first, texts only on a match
            For i = 0 To pLast
                If pHash(i) = hash Then
                    If pKeys(i) = key Then
                        FindPos = i
                        Exit Function
                    End If
                ElseIf pHash(i) = 0 Then
                    If pKeys(i) = key Then
                        pHash(i) = hash
                        FindPos = i
                        Exit Function
                    End If
                End If
            Next i
        End If
        FindPos = -1
        Exit Function
    End If
    If hash = 0 Then hash = HashOf(key)
    s = hash And pMask
    Do
        p = pSlots(s)
        If p = 0 Then
            FindPos = -1
            Exit Function
        End If
        If pHash(p - 1) = hash Then
            If pKeys(p - 1) = key Then
                FindPos = p - 1
                Exit Function
            End If
        End If
        s = (s + 1) And pMask
    Loop
End Function

' The table for the keys there are, at most half full; the kept hashes place them
Private Sub BuildIndex()
    Dim size As Long, i As Long, s As Long
    size = 64
    Do While size < 2 * (pLast + 1)
        size = size * 2
    Loop
    ReDim pSlots(0 To size - 1)
    pMask = size - 1
    For i = 0 To pLast
        If pHash(i) = 0 Then pHash(i) = HashOf(pKeys(i))
        s = pHash(i) And pMask
        Do While pSlots(s) <> 0
            s = (s + 1) And pMask
        Loop
        pSlots(s) = i + 1
    Next i
End Sub

Public Function Exists(key As String) As Boolean
    Exists = (FindPos(key, 0) > -1)
End Function

Private Sub ExpandBuffer()
    pCap = 2 * (pCap + 1)
    ReDim Preserve pKeys(0 To pCap - 1)
    ReDim Preserve pVals(0 To pCap - 1)
    ReDim Preserve pHash(0 To pCap - 1)
End Sub

' Add: raise error if key exists (Dictionary-like Add)
Public Sub Add(key As String, value As Variant)
    If FindPos(key, 0) > -1 Then
        Err.Raise vbObjectError + 10000, "Map.Add", "Key already exists: " & key
    End If
    AppendBufferValue key, value, 0
End Sub

Private Sub AppendBufferValue(ByRef key As String, ByRef value As Variant, ByVal hash As Long)
    Dim s As Long
    pLast = pLast + 1
    If pLast >= pCap Then ExpandBuffer
    pKeys(pLast) = key
    pHash(pLast) = hash
    If IsObject(value) Then Set pVals(pLast) = value Else pVals(pLast) = value
    If pMask < 0 Then
        If pLast + 1 >= INDEX_FROM Then BuildIndex
    ElseIf 2 * (pLast + 1) > pMask + 1 Then
        BuildIndex
    Else
        If hash = 0 Then
            hash = HashOf(key)
            pHash(pLast) = hash
        End If
        s = hash And pMask
        Do While pSlots(s) <> 0
            s = (s + 1) And pMask
        Loop
        pSlots(s) = pLast + 1
    End If
End Sub

' SetValue: set or replace value for key
Public Sub SetValue(key As String, value As Variant)
    Dim idx As Long
    idx = FindPos(key, 0)
    If idx > -1 Then 'Update
        If IsObject(value) Then Set pVals(idx) = value Else pVals(idx) = value
    Else 'Append
        AppendBufferValue key, value, 0
    End If
End Sub

' SetHashed: SetValue with the key's HashOf known already (a key the program names)
Public Sub SetHashed(key As String, ByVal hash As Long, value As Variant)
    Dim idx As Long
    idx = FindPos(key, hash)
    If idx > -1 Then
        If IsObject(value) Then Set pVals(idx) = value Else pVals(idx) = value
    Else
        AppendBufferValue key, value, hash
    End If
End Sub

' ReplaceValue: set the value of an existing key; False, and nothing changed, when key is missing
Public Function ReplaceValue(key As String, value As Variant) As Boolean
    Dim idx As Long
    idx = FindPos(key, 0)
    If idx > -1 Then
        If IsObject(value) Then Set pVals(idx) = value Else pVals(idx) = value
        ReplaceValue = True
    End If
End Function

' GetValue (returns Nothing if not exists); the VM's most frequent call, so a small object's
' search is inline. A Function, not a Property Get: LibreOffice gets a Property Get's value wrong
' here
Public Function GetValue(key As String) As Variant
Attribute GetValue.VB_UserMemId = 0
    Dim i As Long
    If pMask < 0 Then
        For i = 0 To pLast
            If pKeys(i) = key Then
                If IsObject(pVals(i)) Then Set GetValue = pVals(i) Else GetValue = pVals(i)
                Exit Function
            End If
        Next i
    Else
        i = FindPos(key, 0)
        If i > -1 Then
            If IsObject(pVals(i)) Then Set GetValue = pVals(i) Else GetValue = pVals(i)
            Exit Function
        End If
    End If
    Set GetValue = Nothing
End Function

' TryGetValue: one lookup for Exists and GetValue; value is left alone when key is missing
Public Function TryGetValue(key As String, ByRef value As Variant) As Boolean
    Dim i As Long
    If pMask < 0 Then           ' a small object, inline: the VM reads each closure's code with it
        For i = 0 To pLast
            If pKeys(i) = key Then
                If IsObject(pVals(i)) Then Set value = pVals(i) Else value = pVals(i)
                TryGetValue = True
                Exit Function
            End If
        Next i
        Exit Function
    End If
    i = FindPos(key, 0)
    If i > -1 Then
        If IsObject(pVals(i)) Then Set value = pVals(i) Else value = pVals(i)
        TryGetValue = True
    End If
End Function

' TryGetValue that tries position hint first and leaves there the position it found the key at:
' the inline cache of the bytecode VM, one comparison when the object has the same shape
Public Function GetHinted(key As String, ByRef hint As Long, ByRef value As Variant) As Boolean
    Dim i As Long
    If hint >= 0 And hint <= pLast Then
        If pKeys(hint) = key Then
            If IsObject(pVals(hint)) Then Set value = pVals(hint) Else value = pVals(hint)
            GetHinted = True
            Exit Function
        End If
    End If
    i = FindPos(key, 0)
    If i > -1 Then
        If IsObject(pVals(i)) Then Set value = pVals(i) Else value = pVals(i)
        hint = i
        GetHinted = True
    End If
End Function

' GetHinted with the key's HashOf known already (a key the program names)
Public Function GetHashed(key As String, ByVal hash As Long, ByRef hint As Long, ByRef value As Variant) As Boolean
    Dim i As Long
    If hint >= 0 And hint <= pLast Then
        If pKeys(hint) = key Then
            If IsObject(pVals(hint)) Then Set value = pVals(hint) Else value = pVals(hint)
            GetHashed = True
            Exit Function
        End If
    End If
    i = FindPos(key, hash)
    If i > -1 Then
        If IsObject(pVals(i)) Then Set value = pVals(i) Else value = pVals(i)
        hint = i
        GetHashed = True
    End If
End Function

Public Sub Remove(key As String)
    Dim index As Long
    Dim i As Long
    index = FindPos(key, 0)
    If index > -1 Then
        For i = index To pLast - 1
            pKeys(i) = pKeys(i + 1)
            pHash(i) = pHash(i + 1)
            If IsObject(pVals(i + 1)) Then Set pVals(i) = pVals(i + 1) Else pVals(i) = pVals(i + 1)
        Next i
        pKeys(pLast) = vbNullString
        pHash(pLast) = 0
        pVals(pLast) = emptyVariant
        pLast = pLast - 1
        ' the positions after it moved: the table is made again, or dropped for a small object
        If pMask >= 0 Then
            If pLast + 1 < INDEX_FROM \ 2 Then
                pMask = -1
                Erase pSlots
            Else
                BuildIndex
            End If
        End If
    End If
End Sub
' Keys: return a key given an index (read-only copy)
Public Function key(index As Long) As String
    key = pKeys(index - 1)
End Function
' ValueAt: the value at a position, from 1 in the order the keys came; key(index) is its key.
' Walking an object by position costs one step per key, where key(i) then GetValue(key(i))
' searches for each key again
Public Function ValueAt(ByVal index As Long) As Variant
    If IsObject(pVals(index - 1)) Then Set ValueAt = pVals(index - 1) Else ValueAt = pVals(index - 1)
End Function
' KeysCopy: the keys as they are now, from 0 (Count of them), for a walk that runs script code:
' the code may add or delete keys meanwhile
Public Function KeysCopy() As Variant
    KeysCopy = pKeys
End Function
' Keys: return a Collection of keys (read-only copy)
Public Function keys() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pLast
        out.Add pKeys(i)
    Next i
    Set keys = out
End Function
' Items: return a Collection of values (in same order as keys)
Public Function items() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To pLast
        out.Add pVals(i)
    Next i
    Set items = out
End Function
' Count property
Public Property Get Count() As Long
    Count = pLast + 1
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
    For k = 0 To pLast
        If TypeName(pVals(k)) = "ASF_Map" Then
            Set child = pVals(k)
            m.AppendUnchecked pKeys(k), child.Clone, pHash(k)
        Else
            m.AppendUnchecked pKeys(k), pVals(k), pHash(k)
        End If
    Next k
    Set Clone = m
End Function
' Append for Clone: the source keys are already distinct, and hash is the source's
Public Sub AppendUnchecked(key As String, value As Variant, Optional ByVal hash As Long = 0)
    AppendBufferValue key, value, hash
End Sub
