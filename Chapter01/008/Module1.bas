Attribute VB_Name = "Module1"
Option Explicit

Sub TeisuTest()
    Dim myWeight As Double
    
    Const ToPound As Double = 2.20462
    
    myWeight = Val(InputBox("�п�J�魫!! (kg���)"))
    MsgBox "��" & Int(myWeight * ToPound) & "�^��!!"
    
End Sub


