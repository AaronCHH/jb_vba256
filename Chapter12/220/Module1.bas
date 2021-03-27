Attribute VB_Name = "Module1"
Option Explicit

Sub 設定圖表的標題()
    Dim a As String
    With ActiveSheet.ChartObjects("3教科G").Chart
        .HasTitle = True
        .ChartTitle.Text = Worksheets("綜合").Range("A1")
        .ChartTitle.Font.Size = 18
    End With
End Sub



