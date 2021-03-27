Attribute VB_Name = "Module1"
Option Explicit

Sub 字數確認()
    Dim i As Integer
    For i = 2 To 4
        Cells(i, 2) = Len(Cells(i, 1))
        Cells(i, 3) = LenB(StrConv(Cells(i, 1), vbFromUnicode))
    Next
End Sub

Sub 字數確認2()
    Dim i As Integer
    For i = 2 To 4
        Cells(i, 2) = LenB(Cells(i, 1))
    Next
End Sub

