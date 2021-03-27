Attribute VB_Name = "Module1"
Option Explicit

Sub 日期時間加算()
    Dim i As Integer
    For i = 0 To 11
        Cells(i + 2, 2).Value = DateAdd("ww", 4 * i, Date)
    Next
End Sub

