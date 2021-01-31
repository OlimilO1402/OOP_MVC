VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcModels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As Object 'CalcModel
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Sub Push(aObj As Object)
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2 - 1)
    Set m_Arr(m_Count) = aObj
    m_Count = m_Count + 1
End Sub

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get Item(ByVal Index As Long) As CalcModel
    If m_Count <= Index Then Exit Property
    Set Item = m_Arr(Index)
End Property

Public Function Pop() As CalcModel
    m_Count = m_Count - 1
    Set Pop = m_Arr(m_Count)
    Set m_Arr(m_Count) = Nothing
End Function
