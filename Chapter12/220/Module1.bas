Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�Ϫ����D()
    Dim a As String
    With ActiveSheet.ChartObjects("3�Ь�G").Chart
        .HasTitle = True
        .ChartTitle.Text = Worksheets("��X").Range("A1")
        .ChartTitle.Font.Size = 18
    End With
End Sub



