Attribute VB_Name = "Module1"
Option Explicit

Sub �N�r�J�Ϫ��ʨ�Ϫ�u�@��()
    ActiveSheet.ChartObjects("1").Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="�~�Z�Ϫ�"
End Sub

Sub �N�r�J�Ϫ��ʨ��L�u�@��()
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet2"
End Sub


