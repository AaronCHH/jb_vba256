Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�Ϩ�()
    ActiveSheet.ChartObjects("�c��G").Select
    With ActiveChart
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 12
    End With
End Sub

Sub �����r�J�Ϫ�()
    Dim gr As Range
    Set gr = Worksheets("�c��~�Z").Range("H1:M15")
    With ActiveSheet.ChartObjects.Add(gr.Left, gr.Top, gr.Width, gr.Height)
        .Name = "�c��G"
        .Chart.SetSourceData Range("A1:E5")
    End With
    Set gr = Nothing
End Sub

Sub �]�w�Ϫ����U����()
    ActiveChart.SetElement (msoElementLegendBottom)
End Sub



