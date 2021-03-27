Attribute VB_Name = "Module1"
Option Explicit

Sub ¶É±×()
    With Range("B4:B6").Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
        .Gradient.ColorStops.Add(0).Color = RGB(255, 255, 255)
        .Gradient.ColorStops.Add(1).Color = Range("B3").Interior.Color
    End With
End Sub

