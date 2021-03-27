Attribute VB_Name = "Module1"
Option Explicit

Sub ÃC¦â¿@²H()
    Dim myRange As Range, i As Integer
    Set myRange = Range("A4:C10")
    For i = 1 To myRange.Rows.Count
        myRange.Rows(i).Interior.ThemeColor = xlThemeColorAccent6
        Select Case myRange.Cells(i, 3).Value
            Case 10: myRange.Rows(i).Interior.TintAndShade = 0.8
            Case 20: myRange.Rows(i).Interior.TintAndShade = 0.6
            Case 30: myRange.Rows(i).Interior.TintAndShade = 0
            Case 40: myRange.Rows(i).Interior.TintAndShade = -0.25
        End Select
    Next i
    Set myRange = Nothing
End Sub

    
