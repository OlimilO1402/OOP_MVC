VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_src   As Object
Private m_id    As Long
Private m_cmd   As String
Private m_when  As Long
Private m_modif As Long

Friend Sub New_(Src As Object, Optional ByVal aID As Long, Optional ByVal aCmd As String, Optional ByVal lwhen As Long, Optional ByVal modifi As Long)
    Set m_src = Src
    m_id = aID
    m_cmd = aCmd
    m_when = lwhen
    m_modif = modifi
End Sub

Public Property Get Source() As Object
    Set Source = m_src
End Property

Public Property Get ID() As Long
    ID = m_id
End Property

Public Property Get Cmd() As String
    Cmd = m_cmd
End Property

Public Property Get WWhen() As Long
    WWhen = m_when
End Property

Public Property Get Modifiers() As Long
    Modifiers = m_modif
End Property

