Attribute VB_Name = "Module1"
Option Explicit

Sub ±ø¥óif3()
    If Range("B4").Value >= Range("B1").Value + 5 Then
       Range("B5").Value = "¡¶¡¶"
       Range("B5").Font.ColorIndex = 3
    ElseIf Range("B4").Value >= Range("B1") Then
       Range("B5").Value = "¡¶"
       Range("B5").Font.ColorIndex = 45
    Else
       Range("B5").Value = "¡¿"
       Range("B5").Font.ColorIndex = 5
    End If
End Sub


