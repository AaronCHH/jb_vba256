Attribute VB_Name = "Module2"
Option Explicit

Sub �p�F��(gData As Range, gName As Range)
    Dim gRange As Range
    Set gRange = Application.Union(Range("A3:F3"), gData)
    
    gRange.Select
    ActiveSheet.Shapes.AddChart(xlRadarMarkers).Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=gName.Value
    Worksheets("���Z��").Activate
End Sub

Sub �������Z�Ϫ�()
    �p�F�� Range("A4:F4"), Range("A4")
    �p�F�� Range("A5:F5"), Range("A5")
End Sub

