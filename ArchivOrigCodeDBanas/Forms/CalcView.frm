VERSION 5.00
Begin VB.Form CalcView 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtCalcSolution 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnCalculate 
      Caption         =   "="
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtSecndNum 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TxtFirstNum 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "CalcView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Derek Banas MVC Java Tutorial on YT:
'https://www.youtube.com/watch?v=dTVVa2gfht8

'// This is the View
'// Its only job is to display what the user sees
'// It performs no calculations, but instead passes
'// information entered by the user to whomever needs
'// it.

'belongs to the button BtnCalculate
Private m_ActionListeners As New Collection

'setting up the components is done in the visual Form-editor

Public Property Get FirstNum() As Double
    
    If Double_TryParse(TxtFirstNum.Text, FirstNum) Then Exit Property
    Err.Description = "Please give a valid numeric value: """ & TxtFirstNum.Text & """ is not a number."
    Err.Raise 0
    
End Property

Public Property Get SecndNum() As Double
    
    If Double_TryParse(TxtSecndNum.Text, SecndNum) Then Exit Property
    Err.Description "Please give a valid numeric value: """ & TxtSecndNum.Text & """ is not a number."
    Err.Raise 0
    
End Property

Public Property Let CalcValue(Value As Double)
    
    TxtCalcSolution.Text = Value
    
End Property

Public Sub DisplayErrorMessage(errMess As String)
    
    MsgBox errMess
    
End Sub

'// If the calculateButton is clicked execute a method
'// in the Controller named actionPerformed
Public Sub addCalculateListener(listenForCalcButton As ActionListener)
    
    m_ActionListeners.Add listenForCalcButton
    
End Sub

Private Sub BtnCalculate_Click()
    
    Dim ial As ActionListener ': Set ial = m_ActionListeners.Item(1)
    For Each ial In m_ActionListeners: ial.ActionPerformed: Next
    
End Sub

'// This is the View
'// Its only job is to display what the user sees
'// It performs no calculations, but instead passes
'// information entered by the user to whomever needs
'// it.
'
'import java.awt.event.ActionListener;
'
'import javax.swing.*;
'
'public class CalculatorView extends JFrame{
'
'    private JTextField firstNumber  = new JTextField(10);
'    private JLabel additionLabel = new JLabel("+");
'    private JTextField secondNumber = new JTextField(10);
'    private JButton calculateButton = new JButton("Calculate");
'    private JTextField calcSolution = new JTextField(10);
'
'    CalculatorView(){
'
'        // Sets up the view and adds the components
'
'        JPanel calcPanel = new JPanel();
'
'        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
'        this.setSize(600, 200);
'
'        calcPanel.add(firstNumber);
'        calcPanel.add(additionLabel);
'        calcPanel.add(secondNumber);
'        calcPanel.add(calculateButton);
'        calcPanel.add(calcSolution);
'
'        this.add(calcPanel);
'
'        // End of setting up the components --------
'
'    }
'
'    public int getFirstNumber(){
'
'        return Integer.parseInt(firstNumber.getText());
'
'    }
'
'    public int getSecondNumber(){
'
'        return Integer.parseInt(secondNumber.getText());
'
'    }
'
'    public int getCalcSolution(){
'
'        return Integer.parseInt(calcSolution.getText());
'
'    }
'
'    public void setCalcSolution(int solution){
'
'        calcSolution.setText(Integer.toString(solution));
'
'    }
'
'    // If the calculateButton is clicked execute a method
'    // in the Controller named actionPerformed
'
'    void addCalculateListener(ActionListener listenForCalcButton){
'
'        calculateButton.addActionListener(listenForCalcButton);
'
'    }
'
'    // Open a popup that contains the error message passed
'
'    void displayErrorMessage(String errorMessage){
'
'        JOptionPane.showMessageDialog(this, errorMessage);
'
'    }
'
'}
