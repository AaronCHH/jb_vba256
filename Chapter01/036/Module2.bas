Attribute VB_Name = "Module2"
Option Explicit

Sub �p�F��()
    Worksheets("���Z��").Shapes.AddChart(xlRadarMarkers).Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=ActiveCell.Value
    Worksheets("���Z��").Activate
End Sub

Sub �������Z�Ϫ�()
    Worksheets("���Z��").Range("A3:F4").Select
    Range("A4").Activate
    �p�F��
    Worksheets("���Z��").Range("A3:F3, A5:F5").Select
    Range("A5").Activate
    �p�F��
End Sub

