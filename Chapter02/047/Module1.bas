# encoding=big5
Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷ��x�s��d��()
    Range("A3", "G7").Font.Italic = True
    Range("A3:G7").HorizontalAlignment = xlCenter
    Range("��").Borders.LineStyle = xlContinuous
    Range("A1").Font.Bold = True
    Range("A3:G3,A7").Style = "�n"
End Sub

