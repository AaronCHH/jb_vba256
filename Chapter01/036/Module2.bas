Attribute VB_Name = "Module2"
Option Explicit

Sub 雷達圖()
    Worksheets("成績表").Shapes.AddChart(xlRadarMarkers).Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=ActiveCell.Value
    Worksheets("成績表").Activate
End Sub

Sub 做成成績圖表()
    Worksheets("成績表").Range("A3:F4").Select
    Range("A4").Activate
    雷達圖
    Worksheets("成績表").Range("A3:F3, A5:F5").Select
    Range("A5").Activate
    雷達圖
End Sub

