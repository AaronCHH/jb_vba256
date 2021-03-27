Attribute VB_Name = "Module1"
Sub 建立嵌入圖表()
    Dim gr As Range
    Set gr = Worksheets("3 教科").Range("H1:M15")
    With ActiveSheet.ChartObjects.Add( _
        gr.Left, gr.Top, gr.Width, gr.Height)
        .Name = "3 教科 G"
        .Chart.SetSourceData Range("B3:E13")
    End With
    Set gr = Nothing
End Sub


