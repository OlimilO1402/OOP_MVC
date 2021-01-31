VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcModelMul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CalcModel
Private m_Prd1 As Double
Private m_Prd2 As Double
Private m_calcValue As Double

Public Sub Mul(Prd1 As Double, Prd2 As Double)
    m_Prd1 = Prd1
    m_Prd2 = Prd2
    m_calcValue = m_Prd1 * m_Prd2
End Sub

Public Property Get CalcValue() As Double
    CalcValue = m_calcValue
End Property

' --------========######## '     CalcModel    ' ########========-------- '
Private Property Get CalcModel_Param1() As Double
    CalcModel_Param1 = m_Prd1
End Property
Private Property Get CalcModel_Param2() As Double
    CalcModel_Param2 = m_Prd2
End Property
Private Sub CalcModel_Calc(ByVal Param1 As Double, Optional ByVal Param2 As Double)
    Me.Mul Param1, Param2
End Sub
Private Property Get CalcModel_CalcValue() As Double
    CalcModel_CalcValue = m_calcValue
End Property
