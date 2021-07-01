VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'originally this is a internal class inside CalcController, in Java the CalcListener can access the members of the class CalcController
'we do not have this in VB so we must emulate this
Implements ActionListener

Private This As CalcController

Friend Sub New_(aCC As CalcController)
    Set This = aCC
End Sub

Private Sub ActionListener_ActionPerformed()
    Dim firstNumber As Double, secondNumber As Double
Try: On Error GoTo Catch
    
    firstNumber = This.TheView.FirstNum
    secondNumber = This.TheView.SecndNum
    
    This.TheModel.Add firstNumber, secondNumber
    
    This.TheView.CalcValue = This.TheModel.CalcValue
    
    Exit Sub
Catch:
    
    This.TheView.DisplayErrorMessage "You need to enter 2 numbers" & vbCrLf & Err.Description
End Sub

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

