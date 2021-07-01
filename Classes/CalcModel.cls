VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Param1() As Double
    
End Property

Public Property Get Param2() As Double
    
End Property

Public Sub Calc(ByVal aParam1 As Double, Optional ByVal aParam2 As Double)
    '
End Sub

Public Property Get CalcValue() As Double
    '
End Property


'// The Model performs all the calculations needed
'// and that is it. It doesn't know the View exists
'public class CalculatorModel {
'
'    // Holds the value of the sum of the numbers
'    // entered in the view
'
'    private int calculationValue;
'
'    public void addTwoNumbers(int firstNumber, int secondNumber){
'
'        calculationValue = firstNumber + secondNumber;
'
'    }
'
'    public int getCalculationValue(){
'
'        return calculationValue;
'
'    }
'}

