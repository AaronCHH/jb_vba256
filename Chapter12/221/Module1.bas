Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�b����()
    ActiveSheet.ChartObjects("�ƾ�G").Select
    With ActiveChart.Axes(Type:=xlValue)
        .HasTitle = True
        .AxisTitle.Text = "����"
        .AxisTitle.Orientation = xlVertical
    End With
End Sub


