VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcModelDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CalcModel
Private m_Div1 As Double
Private m_Div2 As Double
Private m_calcValue As Double

Public Sub Div(Div1 As Double, Div2 As Double)
    m_Div1 = Div1
    m_Div2 = Div2
    m_calcValue = m_Div1 / m_Div2
End Sub

Public Property Get CalcValue() As Double
    CalcValue = m_calcValue
End Property

' --------========######## '     CalcModel    ' ########========-------- '
Private Property Get CalcModel_Param1() As Double
    CalcModel_Param1 = m_Div1
End Property
Private Property Get CalcModel_Param2() As Double
    CalcModel_Param2 = m_Div2
End Property
Private Sub CalcModel_Calc(ByVal Param1 As Double, Optional ByVal Param2 As Double)
    Me.Div Param1, Param2
End Sub
Private Property Get CalcModel_CalcValue() As Double
    CalcModel_CalcValue = m_calcValue
End Property
