Attribute VB_Name = "Module1"
Option Explicit

Sub ¨ú±o¦CÄæ½s¸¹()
    Dim myRow As Long, myColumn As Long
    Dim myRange As Range, i As Integer
    
    myRow = Range("A3").End(xlDown).Row
    myColumn = Range("A3").End(xlToRight).Column
    For i = 4 To myRow
        Set myRange = Cells(i, myColumn)
        If myRange.Value > 1 Then
           myRange.Style = "Ãa"
        End If
    Next i
    Set myRange = Nothing
End Sub

