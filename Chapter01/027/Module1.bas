Attribute VB_Name = "Module1"
Option Explicit

Sub ±ø¥óif2()
    If Range("B4").Value > Range("B1").Value Then
       Range("B5").Value = "¡¶"
       Range("B5").Font.ColorIndex = 3
    Else
       Range("B5").Value = "¡¿"
       Range("B5").Font.ColorIndex = 5
    End If
End Sub

