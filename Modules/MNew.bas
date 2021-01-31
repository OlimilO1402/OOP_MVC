Attribute VB_Name = "MNew"
Option Explicit

Public Function CalcController(aView As CalcView, aModel As CalcModel) As CalcController
    Set CalcController = New CalcController: CalcController.New_ aView, aModel
End Function

Public Function ActionEvent(Src As Object, Optional ByVal aID As Long, Optional ByVal aCmd As String, Optional ByVal lwhen As Long, Optional ByVal modifi As Long) As ActionEvent
    Set ActionEvent = New ActionEvent: ActionEvent.New_ Src, aID, aCmd, lwhen, modifi
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

