Attribute VB_Name = "Module1"
Option Explicit

Sub ���J��ӦC()
    Dim i As Integer, rcnt As Integer
    
    rcnt = Cells(Rows.Count, 1).End(xlUp).Row
    For i = rcnt To 3 Step -1
        If Cells(i, 1).Value Like "*��" Then
           Cells(i, 1).EntireRow.Insert
           Cells(i, 1).EntireRow.Interior.ColorIndex = 0
        End If
    Next i
End Sub

