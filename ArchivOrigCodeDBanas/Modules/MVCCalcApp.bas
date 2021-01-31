Attribute VB_Name = "MVCCalcApp"
Option Explicit

'Private m_CalcController As CalcController

Public Sub Main()
    
	dim m_CalcController As CalcController
    Dim TheModel As New CalcModel ': Set TheModel = New CalcModel
    
    'we do not have to instantiate CalcView, VB alread did this for us
    Dim TheView As New CalcView ':   Set TheView = New CalcView
    
    Set m_CalcController = MNew.CalcController(TheView, TheModel)
    
    TheView.Show
    
End Sub

'public class MVCCalculator {
'
'    public static void main(String[] args) {
'
'        CalculatorView theView = new CalculatorView();
'
'        CalculatorModel theModel = new CalculatorModel();
'
'        CalculatorController theController = new CalculatorController(theView,theModel);
'
'        theView.setVisible(true);
'
'    }
'}
