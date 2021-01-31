Attribute VB_Name = "MNew"
Option Explicit

Public Function CalcController(aView As CalcView, aModel As CalcModel) As CalcController
    
    Set CalcController = New CalcController: CalcController.New_ aView, aModel
    
End Function

Public Function CalcListener(aCC As CalcController) As CalcListener
    
    Set CalcListener = New CalcListener: CalcListener.New_ aCC
    
End Function

Public Function Double_TryParse(ByVal s As String, ByRef out_Val As Double) As Boolean
Try: On Error GoTo Catch
    If Not IsNumeric(s) Then Exit Function
    s = Replace(s, ",", ".")
    out_Val = Val(s)
    Double_TryParse = True
    Exit Function
Catch:
End Function

