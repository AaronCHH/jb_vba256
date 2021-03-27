Attribute VB_Name = "Module2"
Option Explicit

Sub 儲存格的合併和刪除()
    Dim i As Integer
    
    If Range("A3").MergeCells = True Then
       Range("A3").MergeCells = False
    End If
    Application.DisplayAlerts = False
    For i = 3 To 13 Step 2
        Range(Cells(i, 2), Cells(i, 2).Offset(1)).MergeCells = True
    Next i
    Application.DisplayAlerts = True
End Sub

