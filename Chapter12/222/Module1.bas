Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�b��׼���()
    ActiveSheet.ChartObjects("�ƾ�G").Select
    ActiveChart.Axes(Type:=xlValue).TickLabels.NumberFormat = "0��"
    ActiveChart.Axes(Type:=xlCategory).TickLabels.Orientation = xlVertical
End Sub

Sub �]�w�����t�m()
    With ActiveChart
        .ApplyLayout (5)
        .HasTitle = False
        .Axes(Type:=xlValue).AxisTitle.Text = "�����t�m 1"
    End With
End Sub

Sub �]�w�Ϫ��˦�()
    ActiveChart.ChartStyle = 29
End Sub
