Attribute VB_Name = "Module2"
Option Explicit

Sub 雷達圖(gData As Range, gName As Range)
    Dim gRange As Range
    Set gRange = Application.Union(Range("A3:F3"), gData)
    
    gRange.Select
    ActiveSheet.Shapes.AddChart(xlRadarMarkers).Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=gName.Value
    Worksheets("成績表").Activate
End Sub

Sub 做成成績圖表()
    雷達圖 Range("A4:F4"), Range("A4")
    雷達圖 Range("A5:F5"), Range("A5")
End Sub

