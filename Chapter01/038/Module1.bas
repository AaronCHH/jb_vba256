Attribute VB_Name = "Module1"
Option Explicit

Sub ByValTest(ByVal xString As String)
    xString = "Window Vista"
End Sub

Sub ByRefTest(ByRef yString As String)
    yString = "Window Vista"
End Sub

Sub Test()
    Dim Hensu As String
    Hensu = "ExcelVBA"
    Call ByValTest(Hensu)
    MsgBox "�ǭȩI�s�����G:" & Hensu
    Hensu = "ExcelVBA"
    Call ByRefTest(Hensu)
    MsgBox "�ǧ}�I�s�����G:" & Hensu
End Sub

