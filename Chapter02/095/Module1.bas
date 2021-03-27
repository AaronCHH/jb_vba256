Attribute VB_Name = "Module1"
Option Explicit

Sub 儲存格插入刪除()
    Dim i As Integer, xRange As Range
    Set xRange = Range("A4").CurrentRegion
    For i = xRange.Rows.Count To 2 Step -1
        If xRange.Cells(i, 7) = 0 Then
           xRange.Rows(i).Delete shift:=xlShiftUp
        End If
    Next
    Set xRange = Nothing
End Sub

Sub 列欄插入刪除()
    Range("A3").EntireRow.Insert
    Range("A3").EntireColumn.Delete
End Sub

