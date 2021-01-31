Attribute VB_Name = "MVCCalcApp"
Option Explicit

Public Sub Main()
    
    Dim TheModel As CalcModel: Set TheModel = New CalcModelAdd
    Dim TheView  As CalcView:  Set TheView = CalcView
    Dim CalcController As CalcController
    Set CalcController = MNew.CalcController(TheView, TheModel)
    
    TheView.Show
    
End Sub
