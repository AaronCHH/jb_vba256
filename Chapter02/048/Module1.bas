Attribute VB_Name = "Module1"
Option Explicit

Sub 參照儲存格()
    Dim i As Integer
    Cells(1, 7).Value = Date
    For i = 2 To 6
        If Cells(7, i).Value >= Cells(7, "G").Value Then
           Cells(7, i).Interior.ColorIndex = 3
        End If
    Next i
    Cells.Font.Size = 12
End Sub


