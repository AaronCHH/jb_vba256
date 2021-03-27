Attribute VB_Name = "Module1"
Option Explicit

Sub 設定圖例()
    ActiveSheet.ChartObjects("販賣G").Select
    With ActiveChart
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 12
    End With
End Sub

Sub 做成崁入圖表()
    Dim gr As Range
    Set gr = Worksheets("販賣業績").Range("H1:M15")
    With ActiveSheet.ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Name = "販賣G"
        .Chart.SetSourceData Range("A1:E5")
    End With
    Set gr = Nothing
End Sub

Sub 設定圖表中的各元素()
    ActiveChart.SetElement (msoElementLegendBottom)
End Sub



