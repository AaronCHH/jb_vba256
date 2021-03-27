Attribute VB_Name = "Module1"
Option Explicit

Sub 設定軸刻度標籤()
    ActiveSheet.ChartObjects("數學G").Select
    ActiveChart.Axes(Type:=xlValue).TickLabels.NumberFormat = "0分"
    ActiveChart.Axes(Type:=xlCategory).TickLabels.Orientation = xlVertical
End Sub

Sub 設定版面配置()
    With ActiveChart
        .ApplyLayout (5)
        .HasTitle = False
        .Axes(Type:=xlValue).AxisTitle.Text = "版面配置 1"
    End With
End Sub

Sub 設定圖表的樣式()
    ActiveChart.ChartStyle = 29
End Sub
