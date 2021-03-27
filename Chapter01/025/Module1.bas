Attribute VB_Name = "Module1"
Option Explicit

Sub 物件處理()
    With Range("A3:D10")
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.ThemeColor = xlThemeColorAccent6
        .Interior.TintAndShade = 0.8
    End With
End Sub

