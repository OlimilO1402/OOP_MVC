VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ActionListener
Private TheView  As CalcView
Private TheModel As CalcModel

Friend Sub New_(aView As CalcView, aModel As CalcModel)
    Set TheView = aView
    Set TheModel = aModel
    TheView.addCalculateListener Me
End Sub

Private Sub ActionListener_ActionPerformed(e As ActionEvent)
Try: On Error GoTo Catch
    Dim cx As Control: Set cx = e.Source
    If cx.Name = "BtnCalculate" Then
        Dim Num1 As Double: Num1 = TheView.FirstNum
        Dim Num2 As Double: Num2 = TheView.SecndNum
        TheModel.Calc Num1, Num2
        TheView.UpdateView TheModel
    ElseIf cx.Name = "CmbOperator" Then
        Select Case e.ID
        Case 0: Set TheModel = New CalcModelAdd
        Case 1: Set TheModel = New CalcModelSub
        Case 2: Set TheModel = New CalcModelMul
        Case 3: Set TheModel = New CalcModelDiv
        End Select
    End If
    Exit Sub
Catch:
    TheView.DisplayErrorMessage "Please enter 2 numbers" & vbCrLf & Err.Description
End Sub



'import java.awt.event.ActionEvent;
'import java.awt.event.ActionListener;
'
'// The Controller coordinates interactions
'// between the View and Model
'
'public class CalculatorController {
'
'    private CalculatorView theView;
'    private CalculatorModel theModel;
'
'    public CalculatorController(CalculatorView theView, CalculatorModel theModel) {
'        this.theView = theView;
'        this.theModel = theModel;
'
'        // Tell the View that when ever the calculate button
'        // is clicked to execute the actionPerformed method
'        // in the CalculateListener inner class
'
'        this.theView.addCalculateListener(new CalculateListener());
'    }
'
'    class CalculateListener implements ActionListener{
'
'        public void actionPerformed(ActionEvent e) {
'
'            int firstNumber, secondNumber = 0;
'
'            // Surround interactions with the view with
'            // a try block in case numbers weren't
'            // properly entered
'
'            try{
'
'                firstNumber = theView.getFirstNumber();
'                secondNumber = theView.getSecondNumber();
'
'                theModel.addTwoNumbers(firstNumber, secondNumber);
'
'                theView.setCalcSolution(theModel.getCalculationValue());
'
'            }
'
'            catch(NumberFormatException ex){
'
'                System.out.println(ex);
'
'                theView.displayErrorMessage("You Need to Enter 2 Integers");
'
'            }
'
'        }
'
'    }
'
'}

