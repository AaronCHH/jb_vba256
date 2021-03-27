Attribute VB_Name = "Module1"
Option Explicit

Sub กำ???4()
    With Range("B5")
    Select Case Range("B4").Value - Range("B1").Value
        Case Is >= 5
            .Value = "????"
            .Font.ColorIndex = 3
        Case Is > 0
            .Value = "??"
            .Font.ColorIndex = 45
        Case 0
            .Value = "--"
            .Font.ColorIndex = 0
        Case Is <= -5
            .Value = "????"
            .Font.ColorIndex = 5
        Case Is < 0
            .Value = "??"
            .Font.ColorIndex = 43
    End Select
    End With
End Sub

