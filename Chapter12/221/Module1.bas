Attribute VB_Name = "Module1"
Option Explicit

Sub 設定軸標籤()
    ActiveSheet.ChartObjects("數學G").Select
    With ActiveChart.Axes(Type:=xlValue)
        .HasTitle = True
        .AxisTitle.Text = "分數"
        .AxisTitle.Orientation = xlVertical
    End With
End Sub


