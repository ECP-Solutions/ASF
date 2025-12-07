VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module: StackFast
Option Explicit

Private m_arr() As Variant
Private m_top As Long
Private m_capacity As Long

Private Sub Class_Initialize()
    m_capacity = 32            ' initial capacity (power-of-two)
    ReDim m_arr(1 To m_capacity)
    m_top = 0
End Sub

Public Sub Clear()
    m_capacity = 16
    ReDim m_arr(1 To m_capacity)
    m_top = 0
End Sub

Public Property Get Count() As Long
    Count = m_top
End Property

Public Sub Push(ByVal v As Variant)
    ' Grow if needed (double strategy)
    If m_top >= m_capacity Then
        Dim newCap As Long
        newCap = m_capacity * 2
        ReDim Preserve m_arr(1 To newCap)
        m_capacity = newCap
    End If
    m_top = m_top + 1
    vAssignment m_arr(m_top), v
End Sub

Public Function Pop() As Variant
    If m_top = 0 Then
        Pop = Empty
        Exit Function
    End If
    vAssignment Pop, m_arr(m_top)
    m_arr(m_top) = Empty
    m_top = m_top - 1
End Function

Public Property Get Item(ByVal idx As Long) As Variant
    If idx >= 1 And idx <= m_top Then
        vAssignment Item, m_arr(idx)
    Else
        Item = Empty
    End If
End Property

Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
