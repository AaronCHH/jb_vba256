Attribute VB_Name = "Module1"
Option Explicit

Sub �r�ƽT�{()
    Dim i As Integer
    For i = 2 To 4
        Cells(i, 2) = Len(Cells(i, 1))
        Cells(i, 3) = LenB(StrConv(Cells(i, 1), vbFromUnicode))
    Next
End Sub

Sub �r�ƽT�{2()
    Dim i As Integer
    For i = 2 To 4
        Cells(i, 2) = LenB(Cells(i, 1))
    Next
End Sub

