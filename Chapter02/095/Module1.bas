Attribute VB_Name = "Module1"
Option Explicit

Sub �x�s�洡�J�R��()
    Dim i As Integer, xRange As Range
    Set xRange = Range("A4").CurrentRegion
    For i = xRange.Rows.Count To 2 Step -1
        If xRange.Cells(i, 7) = 0 Then
           xRange.Rows(i).Delete shift:=xlShiftUp
        End If
    Next
    Set xRange = Nothing
End Sub

Sub �C�洡�J�R��()
    Range("A3").EntireRow.Insert
    Range("A3").EntireColumn.Delete
End Sub

