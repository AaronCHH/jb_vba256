Attribute VB_Name = "Module1"
Option Explicit

Sub ¦r¦êÂà´«()
    Dim i As Integer
    
    For i = 2 To 5
        Cells(i, 2) = StrConv(Cells(i, 1), vbProperCase)
        Cells(i, 3) = StrConv(Cells(i, 1), vbUpperCase + vbNarrow)
    Next
End Sub



