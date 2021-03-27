Attribute VB_Name = "Module1"
Option Explicit

Sub ±ø¥ó§PÂ_¨ç¼Æ()
    Dim i As Integer, cr As Range
    Set cr = Range("A3").CurrentRegion
    For i = 2 To cr.Rows.Count - 1
        cr.Rows(i).Interior.ColorIndex = _
            IIf(cr.Cells(i, 6).Value >= 520, 38, 34)
    Next
    Set cr = Nothing
End Sub



