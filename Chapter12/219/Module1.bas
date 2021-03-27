Attribute VB_Name = "Module1"
Option Explicit

Sub 變更圖表種類()
    ActiveSheet.ChartObjects("3教科G").Chart.ChartType = xl3DBarClustered
End Sub

