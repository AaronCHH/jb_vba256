Attribute VB_Name = "Module1"
Option Explicit

Sub 取得終端儲存格()
    Dim sRange As Range, eRange As Range
    Dim i As Integer
    For i = 3 To Range("A2").End(xlDown).Row
        Set sRange = Cells(i, 1)
        Set eRange = sRange.End(xlToRight)
        If i Mod 2 = 0 Then
           Range(sRange, eRange).Interior.ColorIndex = 34
        End If
    Next
    Set sRange = Nothing
    Set eRange = Nothing
End Sub

