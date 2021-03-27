Attribute VB_Name = "Module1"
Option Explicit

Sub TeisuTest()
    Dim myWeight As Double
    
    Const ToPound As Double = 2.20462
    
    myWeight = Val(InputBox("請輸入體重!! (kg單位)"))
    MsgBox "約" & Int(myWeight * ToPound) & "英鎊!!"
    
End Sub


