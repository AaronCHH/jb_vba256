Attribute VB_Name = "Module1"
Option Explicit

Sub ­«½Æ1()
    Dim i As Integer
    For i = 2 To 5
        Worksheets(i).Name = Cells(i, 1).Value
    Next i
End Sub

Sub KuriKaeshi1()
    Dim i As Integer
    For i = 5 To 2 Step -2
        Range(Cells(i, 1), Cells(i, 4)).Interior.ColorIndex = 40
    Next i
End Sub

