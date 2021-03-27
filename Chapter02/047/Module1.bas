# encoding=big5
Attribute VB_Name = "Module1"
Option Explicit

Sub 參照儲存格範圍()
    Range("A3", "G7").Font.Italic = True
    Range("A3:G7").HorizontalAlignment = xlCenter
    Range("表").Borders.LineStyle = xlContinuous
    Range("A1").Font.Bold = True
    Range("A3:G3,A7").Style = "好"
End Sub

