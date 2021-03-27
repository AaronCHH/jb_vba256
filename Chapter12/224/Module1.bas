Attribute VB_Name = "Module1"
Option Explicit

Sub 將崁入圖表移動到圖表工作表()
    ActiveSheet.ChartObjects("1").Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="業績圖表"
End Sub

Sub 將崁入圖表移動到其他工作表()
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet2"
End Sub


