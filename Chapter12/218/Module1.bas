Attribute VB_Name = "Module1"
Sub �إߴO�J�Ϫ�()
    Dim gr As Range
    Set gr = Worksheets("3 �Ь�").Range("H1:M15")
    With ActiveSheet.ChartObjects.Add( _
        gr.Left, gr.Top, gr.Width, gr.Height)
        .Name = "3 �Ь� G"
        .Chart.SetSourceData Range("B3:E13")
    End With
    Set gr = Nothing
End Sub


